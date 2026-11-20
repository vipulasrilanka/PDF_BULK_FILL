"""Coordinate-based PDF text overlays and legacy SrinkToFit support via PyMuPDF."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence

import pymupdf

from pdf_bulk_fill.errors import PdfRenderError

PDF_DPI = 72.0
MM_PER_INCH = 25.4
PDF_DPMM = PDF_DPI / MM_PER_INCH

DEFAULT_FONT = "Helvetica"
DEFAULT_FONT_SIZE = 12.0
DEFAULT_LINE_SPACE_FACTOR = 1.2
MIN_SHRINK_FONT_SIZE = 6.0

_FONT_MAP: dict[str, str] = {
    "helvetica": "helv",
    "helv": "helv",
    "helvetica-bold": "hebo",
    "hebo": "hebo",
    "helvetica-oblique": "heit",
    "helvetica-italic": "heit",
    "heit": "heit",
    "helvetica-boldoblique": "hebi",
    "hebi": "hebi",
    "times-roman": "tiro",
    "times": "tiro",
    "tiro": "tiro",
    "times-bold": "tibo",
    "tibo": "tibo",
    "times-italic": "tiit",
    "tiit": "tiit",
    "times-bolditalic": "tibi",
    "tibi": "tibi",
    "courier": "cour",
    "cour": "cour",
    "courier-bold": "cobo",
    "cobo": "cobo",
    "courier-oblique": "coit",
    "coit": "coit",
    "courier-boldoblique": "cobi",
    "cobi": "cobi",
}


@dataclass(frozen=True)
class RenderedLine:
    """Single positioned line of text ready to draw on a page."""

    text: str
    x_offset: float
    font_size: float


def measure_to_points(measure: Any) -> float:
    """Convert a length in points (int/float/pt), millimeters (mm), or inches (in) to PDF points."""
    if isinstance(measure, bool) or measure is None:
        raise PdfRenderError(f"Invalid coordinate/measure value: {measure!r}")
    if isinstance(measure, (int, float)):
        return float(measure)
    if isinstance(measure, str):
        raw = measure.strip().lower()
        if not raw:
            raise PdfRenderError("Empty coordinate/measure string.")
        try:
            if raw.endswith("mm"):
                return float(raw[:-2].strip()) * PDF_DPMM
            if raw.endswith("in"):
                return float(raw[:-2].strip()) * PDF_DPI
            if raw.endswith("pt"):
                return float(raw[:-2].strip())
            return float(raw)
        except ValueError as exc:
            raise PdfRenderError(f"Could not parse measure {measure!r}: {exc}") from exc
    raise PdfRenderError(f"Unsupported measure type: {type(measure).__name__}")


def resolve_font_name(font: str | None) -> str:
    """Map standard PDF font names (e.g., Helvetica) to PyMuPDF Base-14 font codes."""
    if not font:
        return "helv"
    return _FONT_MAP.get(font.strip().lower(), "helv")


def resolve_line_height(font_size: float, line_space: Any) -> float:
    """Compute line spacing in points from a numeric value or multiplier string (e.g. '1.2X')."""
    if line_space is None or line_space == "":
        return font_size * DEFAULT_LINE_SPACE_FACTOR
    if isinstance(line_space, (int, float)) and not isinstance(line_space, bool):
        return float(line_space)
    if isinstance(line_space, str):
        raw = line_space.strip()
        if raw.lower().endswith("x"):
            try:
                return float(raw[:-1]) * font_size
            except ValueError as exc:
                raise PdfRenderError(f"Invalid LineSpace multiplier: {line_space!r}") from exc
        return measure_to_points(raw)
    return font_size * DEFAULT_LINE_SPACE_FACTOR


def wrap_shrink_to_fit(
    text: str,
    *,
    font_name: str,
    font_size: float,
    max_width: float,
    max_lines: int,
    first_line_indent: float = 0.0,
    min_font_size: float = MIN_SHRINK_FONT_SIZE,
) -> list[RenderedLine]:
    """Word-wrap text within ``max_width`` and ``max_lines``, shrinking font size if needed.

    Supports legacy ``SrinkToFit(width, maxLines[, indent])`` behavior without global state.
    """
    if max_width <= 0 or max_lines <= 0:
        raise PdfRenderError("SrinkToFit requires positive max_width and max_lines.")
    if first_line_indent < 0 or first_line_indent >= max_width:
        raise PdfRenderError(
            f"Invalid first_line_indent ({first_line_indent}) for max_width ({max_width})."
        )

    words = str(text).split(" ")
    current_size = float(font_size)

    while current_size >= min_font_size:
        lines: list[RenderedLine] = []
        current_line = ""
        indent = first_line_indent
        available_width = max_width - indent

        for word in words:
            candidate = f"{current_line} {word}" if current_line else word
            width = pymupdf.get_text_length(
                candidate, fontname=font_name, fontsize=current_size
            )
            if width <= available_width:
                current_line = candidate
            else:
                if current_line:
                    lines.append(
                        RenderedLine(text=current_line, x_offset=indent, font_size=current_size)
                    )
                    indent = 0.0
                    available_width = max_width
                    current_line = word
                else:
                    # Single word itself is wider than available_width -> force shrink
                    lines.append(
                        RenderedLine(text=word, x_offset=indent, font_size=current_size)
                    )
                    lines.append(RenderedLine(text="", x_offset=0.0, font_size=current_size))
                    current_line = ""
                    break

        if current_line:
            lines.append(
                RenderedLine(text=current_line, x_offset=indent, font_size=current_size)
            )

        # Ensure every emitted line actually fits its width
        fits_width = all(
            pymupdf.get_text_length(line.text, fontname=font_name, fontsize=current_size)
            <= (max_width - line.x_offset)
            for line in lines
        )
        if len(lines) <= max_lines and fits_width:
            return lines

        current_size -= 1.0

    raise PdfRenderError(
        f"Text is too long to fit within {max_width:.1f}pt x {max_lines} lines "
        f"(min font size {min_font_size}pt)."
    )


def apply_text_overlays_to_page(
    page: pymupdf.Page,
    overlays: Sequence[Mapping[str, Any]],
) -> int:
    """Draw coordinate-based text overlays onto an open PyMuPDF page.

    Each overlay dict expects:
        ``{"name": ..., "string": ..., "param": {"X": ..., "Y": ..., ...}}``
    Returns the number of overlay entries drawn.
    """
    page_height = page.rect.height
    drawn_count = 0

    for overlay in overlays:
        params = overlay.get("param")
        if not isinstance(params, Mapping):
            raise PdfRenderError(
                f"Overlay {overlay.get('name')!r} is missing required 'param' dictionary."
            )
        if "X" not in params or "Y" not in params:
            raise PdfRenderError(
                f"Overlay {overlay.get('name')!r} must specify both 'X' and 'Y' coordinates."
            )

        x_pt = measure_to_points(params["X"])
        y_raw = measure_to_points(params["Y"])
        origin = str(params.get("Origin", "bottom-left")).strip().lower()
        y_pt = y_raw if origin == "top-left" else (page_height - y_raw)

        font_code = resolve_font_name(params.get("Font", DEFAULT_FONT))
        font_size = float(params.get("FontSize", DEFAULT_FONT_SIZE))
        line_space_spec = params.get("LineSpace")
        func = params.get("Function")

        text = "" if overlay.get("string") is None else str(overlay.get("string"))

        if isinstance(func, Mapping) and func.get("name") in ("SrinkToFit", "ShrinkToFit"):
            max_width = measure_to_points(func.get("param1"))
            max_lines = int(func.get("param2", 1))
            raw_indent = func.get("param3")
            indent = measure_to_points(raw_indent) if raw_indent is not None else 0.0
            lines = wrap_shrink_to_fit(
                text,
                font_name=font_code,
                font_size=font_size,
                max_width=max_width,
                max_lines=max_lines,
                first_line_indent=indent,
            )
        else:
            lines = [RenderedLine(text=text, x_offset=0.0, font_size=font_size)]

        cursor_y = y_pt
        for idx, line in enumerate(lines):
            if idx > 0:
                cursor_y += resolve_line_height(line.font_size, line_space_spec)
            page.insert_text(
                pymupdf.Point(x_pt + line.x_offset, cursor_y),
                line.text,
                fontname=font_code,
                fontsize=line.font_size,
                color=(0, 0, 0),
            )

        drawn_count += 1

    return drawn_count


def add_text_overlays_to_pdf(
    template_pdf: str | Path,
    output_pdf: str | Path,
    overlays: Sequence[Mapping[str, Any]],
    *,
    page_index: int = 0,
) -> int:
    """Open ``template_pdf``, apply coordinate-based text overlays, and save to ``output_pdf``."""
    src_path = Path(template_pdf)
    if not src_path.exists():
        raise PdfRenderError(f"PDF template not found: {src_path}")

    doc = pymupdf.open(src_path)
    try:
        if page_index < 0 or page_index >= len(doc):
            raise PdfRenderError(
                f"Page index {page_index} out of range for '{src_path.name}' ({len(doc)} pages)."
            )
        drawn = apply_text_overlays_to_page(doc[page_index], overlays)
        dst_path = Path(output_pdf)
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(dst_path, deflate=True)
        return drawn
    finally:
        doc.close()
