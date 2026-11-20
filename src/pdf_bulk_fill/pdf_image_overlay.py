"""Employee ID card (front & back) and general image overlay support via PyMuPDF + Pillow."""

from __future__ import annotations

import io
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence

import pymupdf
from PIL import Image, ImageOps, UnidentifiedImageError

from pdf_bulk_fill.errors import ImageOverlayError
from pdf_bulk_fill.pdf_text_overlay import measure_to_points

SUPPORTED_IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".webp")

# Preset bounding boxes (top-left coordinates in PDF points) for RR-Form.pdf (A4: 595.2 x 841.8)
# Upper blank area spans roughly y=85..485, x=40..555.
RR_FORM_IMAGE_SLOTS: dict[str, tuple[float, float, float, float]] = {
    "ID_FRONT": (120.0, 95.0, 475.0, 275.0),
    "ID_BACK": (120.0, 295.0, 475.0, 475.0),
    "ID_LEFT": (45.0, 140.0, 290.0, 390.0),
    "ID_RIGHT": (305.0, 140.0, 550.0, 390.0),
}


@dataclass(frozen=True)
class ImageOverlaySpec:
    """Resolved image overlay instruction."""

    image_path: Path
    target_rect: pymupdf.Rect
    keep_aspect: bool = True


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


def resolve_image_path(
    raw_path: str | Path,
    *,
    base_dir: str | Path | None = None,
    record_key: str | None = None,
) -> Path:
    """Resolve an image path from an absolute/relative path or ``{key}`` template pattern.

    If the resolved path has no extension and does not exist directly, probes common
    image extensions (``.jpg``, ``.jpeg``, ``.png``, ``.bmp``, ``.webp``).
    """
    path_str = str(raw_path).strip()
    if not path_str:
        raise ImageOverlayError("Empty image path.")

    if record_key is not None and "{key}" in path_str:
        path_str = path_str.replace("{key}", str(record_key))

    candidate = Path(path_str)
    if not candidate.is_absolute() and base_dir is not None:
        candidate = Path(base_dir) / candidate

    if candidate.exists() and candidate.is_file():
        return candidate

    if not candidate.suffix:
        for ext in SUPPORTED_IMAGE_EXTENSIONS:
            with_ext = candidate.with_suffix(ext)
            if with_ext.exists() and with_ext.is_file():
                return with_ext

    raise ImageOverlayError(f"Image file not found: {candidate}")


def resolve_target_rect(
    params: Mapping[str, Any],
    *,
    page_height: float,
    slots: Mapping[str, tuple[float, float, float, float]] = RR_FORM_IMAGE_SLOTS,
) -> pymupdf.Rect:
    """Compute a PyMuPDF target Rect from either a named ``Slot`` or explicit coordinates."""
    slot_name = params.get("Slot")
    if isinstance(slot_name, str) and slot_name.strip():
        key = slot_name.strip().upper()
        if key not in slots:
            raise ImageOverlayError(
                f"Unknown image slot '{slot_name}'. Valid slots: {', '.join(sorted(slots))}."
            )
        return pymupdf.Rect(*slots[key])

    required = ("X", "Y", "Width", "Height")
    missing = [k for k in required if k not in params]
    if missing:
        raise ImageOverlayError(
            f"Image overlay requires 'Slot' or all of {required}; missing: {', '.join(missing)}."
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


def _load_normalized_image_bytes(image_path: Path) -> tuple[bytes, int, int]:
    """Load an image with Pillow, apply EXIF orientation, and return PNG bytes + dimensions."""
    try:
        with Image.open(image_path) as img:
            normalized = ImageOps.exif_transpose(img)
            if normalized.mode not in ("RGB", "RGBA"):
                normalized = normalized.convert("RGB")
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
    record_key: str | None = None,
    skip_missing: bool = False,
) -> list[str]:
    """Insert one or more image overlays onto an open PyMuPDF page.

    Each entry expects:
        ``{"name": ..., "path": ..., "param": {"Slot": "ID_FRONT"} or {"X": ..., "Y": ..., "Width": ..., "Height": ...}}``

    Returns:
        List of overlay names successfully inserted.
    """
    inserted: list[str] = []
    page_height = page.rect.height

    for entry in image_overlays:
        name = str(entry.get("name", "image"))
        raw_path = entry.get("path", entry.get("string"))
        params = entry.get("param")
        if not isinstance(params, Mapping):
            raise ImageOverlayError(f"Image overlay '{name}' is missing required 'param' mapping.")

        try:
            resolved_path = resolve_image_path(
                "" if raw_path is None else raw_path,
                base_dir=base_dir,
                record_key=record_key,
            )
        except ImageOverlayError:
            if skip_missing:
                continue
            raise

        target_rect = resolve_target_rect(params, page_height=page_height)
        img_bytes, img_w, img_h = _load_normalized_image_bytes(resolved_path)

        keep_aspect = params.get("KeepAspect", True)
        if isinstance(keep_aspect, str):
            keep_aspect = keep_aspect.strip().lower() not in ("false", "0", "no")

        dest_rect = (
            fit_rect_preserving_aspect(target_rect, img_w, img_h)
            if keep_aspect
            else target_rect
        )
        page.insert_image(dest_rect, stream=img_bytes, keep_proportion=bool(keep_aspect))
        inserted.append(name)

    return inserted


def add_image_overlays_to_pdf(
    template_pdf: str | Path,
    output_pdf: str | Path,
    image_overlays: Sequence[Mapping[str, Any]],
    *,
    page_index: int = 0,
    base_dir: str | Path | None = None,
    record_key: str | None = None,
    skip_missing: bool = False,
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
            record_key=record_key,
            skip_missing=skip_missing,
        )
        dst_path = Path(output_pdf)
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(dst_path, deflate=True)
        return inserted
    finally:
        doc.close()
