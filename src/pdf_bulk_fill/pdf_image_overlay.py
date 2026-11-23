"""Generic DocType-matched data-driven and constant image overlays via PyMuPDF + Pillow."""

from __future__ import annotations

import io
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence

import pymupdf
from PIL import Image, ImageOps, UnidentifiedImageError

from pdf_bulk_fill.errors import ImageOverlayError
from pdf_bulk_fill.pdf_text_overlay import measure_to_points
from pdf_bulk_fill.template_parser import parse_kv_pairs

logger = logging.getLogger(__name__)

SUPPORTED_IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".webp")


@dataclass(frozen=True)
class ImageAsset:
    """Properties of an input image asset (from data workbook or constant template row)."""

    file_name: str
    doc_type: str | None = None
    location: str = ""
    rotation: float = 0.0
    keep_aspect: bool = True


def _parse_bool(value: Any, *, default: bool = True) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    raw = str(value).strip().lower()
    if raw in ("true", "1", "yes", "y", "on"):
        return True
    if raw in ("false", "0", "no", "n", "off"):
        return False
    raise ImageOverlayError(f"Invalid boolean value for KeepAspect: {value!r}")


def parse_image_asset(raw: Any) -> ImageAsset | None:
    """Parse an ImageAsset from an ImageAsset instance, a dict, or a ``<key=value>`` string.

    Returns ``None`` when the input is empty/``None`` (allowing the caller to skip cleanly).
    """
    if raw is None:
        return None
    if isinstance(raw, ImageAsset):
        return raw

    if isinstance(raw, str):
        stripped = raw.strip()
        if not stripped:
            return None
        parsed = parse_kv_pairs(stripped)
        if not parsed:
            raise ImageOverlayError(
                f"Malformed image metadata {raw!r}; expected '<DocType=...><FileName=...>...'."
            )
        data: Mapping[str, Any] = parsed
    elif isinstance(raw, Mapping):
        data = raw
    else:
        raise ImageOverlayError(f"Unsupported image metadata type: {type(raw).__name__}")

    raw_file = data.get("FileName")
    file_name = "" if raw_file is None else str(raw_file).strip()
    if not file_name:
        raise ImageOverlayError("Image metadata is missing required 'FileName' property.")

    raw_doc_type = data.get("DocType")
    doc_type = str(raw_doc_type).strip() if raw_doc_type is not None and str(raw_doc_type).strip() else None

    raw_location = data.get("Location")
    location = "" if raw_location is None else str(raw_location).strip()

    raw_rotation = data.get("Rotation", 0.0)
    try:
        rotation = float(raw_rotation) if raw_rotation is not None and str(raw_rotation).strip() != "" else 0.0
    except ValueError as exc:
        raise ImageOverlayError(f"Invalid Rotation value: {raw_rotation!r}") from exc

    keep_aspect = _parse_bool(data.get("KeepAspect"), default=True)

    return ImageAsset(
        file_name=file_name,
        doc_type=doc_type,
        location=location,
        rotation=rotation,
        keep_aspect=keep_aspect,
    )


def fit_rect_preserving_aspect(
    target_rect: pymupdf.Rect,
    img_width: int,
    img_height: int,
) -> pymupdf.Rect:
    """Return the largest centered rectangle inside ``target_rect`` matching image aspect ratio."""
    if img_width <= 0 or img_height <= 0:
        raise ImageOverlayError(f"Invalid image dimensions: {img_width}x{img_height}")
    if target_rect.width <= 0 or target_rect.height <= 0:
        raise ImageOverlayError(f"Invalid target rectangle: {target_rect}")

    scale = min(target_rect.width / img_width, target_rect.height / img_height)
    fitted_w = img_width * scale
    fitted_h = img_height * scale
    x0 = target_rect.x0 + (target_rect.width - fitted_w) / 2.0
    y0 = target_rect.y0 + (target_rect.height - fitted_h) / 2.0
    return pymupdf.Rect(x0, y0, x0 + fitted_w, y0 + fitted_h)


def _probe_existing_file(candidate: Path) -> Path | None:
    if candidate.exists() and candidate.is_file():
        return candidate
    if not candidate.suffix:
        for ext in SUPPORTED_IMAGE_EXTENSIONS:
            with_ext = candidate.with_suffix(ext)
            if with_ext.exists() and with_ext.is_file():
                return with_ext
    return None


def resolve_image_file(
    file_name: str | Path,
    *,
    location: str | Path = "",
    base_dir: str | Path | None = None,
) -> Path:
    """Resolve an image path from ``file_name``, optional ``location``, and ``base_dir``.

    Probes common image extensions when ``file_name`` has no suffix.
    """
    name_str = str(file_name).strip()
    if not name_str:
        raise ImageOverlayError("Empty image file name.")

    name_path = Path(name_str)
    loc_str = str(location).strip() if location is not None else ""
    loc_path = Path(loc_str) if loc_str else None
    base_path = Path(base_dir) if base_dir is not None else None

    candidates: list[Path] = []
    if name_path.is_absolute():
        candidates.append(name_path)
    elif loc_path is not None:
        if loc_path.is_absolute():
            candidates.append(loc_path / name_path)
        elif base_path is not None:
            resolved_base = base_path.resolve()
            primary = (base_path / loc_path / name_path).resolve()
            if primary.is_relative_to(resolved_base):
                candidates.append(base_path / loc_path / name_path)
            if base_path.name == loc_path.name:
                dedup = (base_path / name_path).resolve()
                if dedup.is_relative_to(resolved_base):
                    candidates.append(base_path / name_path)
            if not candidates:
                raise ImageOverlayError(
                    f"Image location '{location}' escapes configured image directory '{base_path}'."
                )
        else:
            candidates.append(loc_path / name_path)
    elif base_path is not None:
        resolved_base = base_path.resolve()
        candidate = (base_path / name_path).resolve()
        if not candidate.is_relative_to(resolved_base):
            raise ImageOverlayError(
                f"Image path '{file_name}' escapes configured image directory '{base_path}'."
            )
        candidates.append(base_path / name_path)
    else:
        candidates.append(name_path)

    for candidate in candidates:
        found = _probe_existing_file(candidate)
        if found is not None:
            return found

    raise ImageOverlayError(f"Image file not found: {candidates[0]}")


def resolve_target_rect(
    params: Mapping[str, Any],
    *,
    page_height: float,
) -> pymupdf.Rect:
    """Compute a PyMuPDF target Rect from explicit ``X, Y, Width, Height`` (and optional ``Origin``)."""
    required = ("X", "Y", "Width", "Height")
    missing = [k for k in required if k not in params]
    if missing:
        raise ImageOverlayError(
            f"Image overlay requires all of {required}; missing: {', '.join(missing)}."
        )

    x = measure_to_points(params["X"])
    y_raw = measure_to_points(params["Y"])
    width = measure_to_points(params["Width"])
    height = measure_to_points(params["Height"])

    if width <= 0 or height <= 0:
        raise ImageOverlayError(f"Image Width and Height must be positive ({width}x{height}).")

    origin = str(params.get("Origin", "top-left")).strip().lower()
    if origin == "bottom-left":
        y0 = page_height - y_raw - height
    else:
        y0 = y_raw

    return pymupdf.Rect(x, y0, x + width, y0 + height)


def _load_normalized_image_bytes(
    image_path: Path,
    *,
    rotation: float = 0.0,
) -> tuple[bytes, int, int]:
    """Load an image with Pillow, apply EXIF transpose + clockwise rotation, and return PNG bytes."""
    try:
        with Image.open(image_path) as img:
            normalized = ImageOps.exif_transpose(img)
            if normalized.mode not in ("RGB", "RGBA"):
                normalized = normalized.convert("RGB")
            if rotation % 360 != 0:
                # Pillow rotates counter-clockwise; negate for intuitive clockwise degrees
                normalized = normalized.rotate(-rotation, expand=True)
            width, height = normalized.size
            buf = io.BytesIO()
            normalized.save(buf, format="PNG")
            return buf.getvalue(), width, height
    except (UnidentifiedImageError, OSError) as exc:
        raise ImageOverlayError(f"Invalid or corrupt image file '{image_path}': {exc}") from exc


def apply_image_overlays_to_page(
    page: pymupdf.Page,
    image_overlays: Sequence[Mapping[str, Any]],
    *,
    base_dir: str | Path | None = None,
    skip_missing: bool = False,
    warnings_out: list[str] | None = None,
) -> list[str]:
    """Process and insert image overlays onto an open PyMuPDF page strictly line-by-line.

    Each entry expects:
        ``{"name": ..., "asset": ImageAsset | dict | str | None, "param": {"X": ..., "Y": ..., "Width": ..., "Height": ..., "Origin": ..., "DocType": ...}}``

    When ``param["DocType"]`` is set, the row is placed only if ``asset.doc_type`` matches
    (case-insensitive); otherwise a warning is logged and the row is skipped.

    Returns:
        List of overlay names inserted onto the page.
    """
    inserted: list[str] = []
    page_height = page.rect.height

    for entry in image_overlays:
        name = str(entry.get("name", "image"))
        params = entry.get("param")
        if not isinstance(params, Mapping):
            raise ImageOverlayError(f"Image overlay '{name}' is missing required 'param' mapping.")

        asset = parse_image_asset(entry.get("asset"))
        if asset is None:
            msg = f"Skipping image overlay '{name}': no image metadata provided."
            logger.warning(msg)
            if warnings_out is not None:
                warnings_out.append(msg)
            continue

        expected_doc_type = params.get("DocType")
        if expected_doc_type is not None and str(expected_doc_type).strip():
            expected_norm = str(expected_doc_type).strip().upper()
            actual_norm = (asset.doc_type or "").strip().upper()
            if actual_norm != expected_norm:
                msg = (
                    f"Skipping image overlay '{name}': image DocType "
                    f"'{asset.doc_type or ''}' does not match slot DocType '{expected_doc_type}'."
                )
                logger.warning(msg)
                if warnings_out is not None:
                    warnings_out.append(msg)
                continue

        try:
            resolved_path = resolve_image_file(
                asset.file_name,
                location=asset.location,
                base_dir=base_dir,
            )
        except ImageOverlayError as exc:
            if skip_missing:
                msg = f"Skipping missing image overlay '{name}': {exc}"
                logger.warning(msg)
                if warnings_out is not None:
                    warnings_out.append(msg)
                continue
            raise

        target_rect = resolve_target_rect(params, page_height=page_height)
        img_bytes, img_w, img_h = _load_normalized_image_bytes(
            resolved_path, rotation=asset.rotation
        )

        dest_rect = (
            fit_rect_preserving_aspect(target_rect, img_w, img_h)
            if asset.keep_aspect
            else target_rect
        )
        page.insert_image(dest_rect, stream=img_bytes, keep_proportion=asset.keep_aspect)
        inserted.append(name)

    return inserted


def add_image_overlays_to_pdf(
    template_pdf: str | Path,
    output_pdf: str | Path,
    image_overlays: Sequence[Mapping[str, Any]],
    *,
    page_index: int = 0,
    base_dir: str | Path | None = None,
    skip_missing: bool = False,
    warnings_out: list[str] | None = None,
) -> list[str]:
    """Open ``template_pdf``, place image overlays on ``page_index``, and save to ``output_pdf``."""
    src_path = Path(template_pdf)
    if not src_path.exists():
        raise ImageOverlayError(f"PDF template not found: {src_path}")

    doc = pymupdf.open(src_path)
    try:
        if page_index < 0 or page_index >= len(doc):
            raise ImageOverlayError(
                f"Page index {page_index} out of range for '{src_path.name}' ({len(doc)} pages)."
            )
        inserted = apply_image_overlays_to_page(
            doc[page_index],
            image_overlays,
            base_dir=base_dir,
            skip_missing=skip_missing,
            warnings_out=warnings_out,
        )
        dst_path = Path(output_pdf)
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(dst_path, deflate=True)
        return inserted
    finally:
        doc.close()
