"""Unit & integration tests for coordinate-based text overlays and SrinkToFit."""

from __future__ import annotations

from pathlib import Path

import pymupdf
import pytest

from pdf_bulk_fill import (
    PdfRenderError,
    add_text_overlays_to_pdf,
    measure_to_points,
    resolve_font_name,
    resolve_line_height,
    wrap_shrink_to_fit,
)


def _create_blank_pdf(path: Path, width: float = 595.0, height: float = 842.0) -> Path:
    doc = pymupdf.open()
    doc.new_page(width=width, height=height)
    doc.save(path)
    doc.close()
    return path


def test_measure_to_points() -> None:
    assert measure_to_points(72) == 72.0
    assert measure_to_points("1in") == pytest.approx(72.0)
    assert measure_to_points("25.4mm") == pytest.approx(72.0)
    assert measure_to_points("36pt") == pytest.approx(36.0)
    assert measure_to_points("18.5") == pytest.approx(18.5)

    for bad in (None, True, "", "abcmm", [10]):
        with pytest.raises(PdfRenderError):
            measure_to_points(bad)


def test_resolve_font_name_and_line_height() -> None:
    assert resolve_font_name(None) == "helv"
    assert resolve_font_name("Helvetica-Bold") == "hebo"
    assert resolve_font_name("Times-Roman") == "tiro"
    assert resolve_font_name("UnknownFont") == "helv"

    assert resolve_line_height(12.0, None) == pytest.approx(14.4)
    assert resolve_line_height(12.0, 20) == pytest.approx(20.0)
    assert resolve_line_height(10.0, "1.5X") == pytest.approx(15.0)
    assert resolve_line_height(10.0, "10mm") == pytest.approx(10 * 72 / 25.4)

    with pytest.raises(PdfRenderError, match="Invalid LineSpace"):
        resolve_line_height(10.0, "badX")


def test_wrap_shrink_to_fit() -> None:
    # Fits without shrinking
    lines = wrap_shrink_to_fit(
        "Hello World",
        font_name="helv",
        font_size=12,
        max_width=200,
        max_lines=2,
        first_line_indent=20,
    )
    assert len(lines) == 1
    assert lines[0].text == "Hello World"
    assert lines[0].x_offset == 20.0
    assert lines[0].font_size == 12.0

    # Forces multi-line wrap and font shrink
    long_text = "One hundred and twenty five thousand four hundred and fifty rupees only"
    shrunk = wrap_shrink_to_fit(
        long_text,
        font_name="helv",
        font_size=14,
        max_width=150,
        max_lines=2,
        first_line_indent=10,
    )
    assert len(shrunk) <= 2
    assert shrunk[0].font_size < 14.0
    assert shrunk[0].x_offset == 10.0
    if len(shrunk) > 1:
        assert shrunk[1].x_offset == 0.0

    # Impossible to fit -> raises PdfRenderError
    with pytest.raises(PdfRenderError, match="too long to fit"):
        wrap_shrink_to_fit(
            "SUPERLONGTEXT " * 50,
            font_name="helv",
            font_size=12,
            max_width=40,
            max_lines=1,
        )

    # Invalid arguments
    with pytest.raises(PdfRenderError):
        wrap_shrink_to_fit("x", font_name="helv", font_size=12, max_width=0, max_lines=1)
    with pytest.raises(PdfRenderError):
        wrap_shrink_to_fit(
            "x", font_name="helv", font_size=12, max_width=50, max_lines=1, first_line_indent=60
        )


def test_add_text_overlays_to_pdf_a4_and_letter(tmp_path: Path) -> None:
    tpl = _create_blank_pdf(tmp_path / "a4_template.pdf", width=595.0, height=842.0)
    out = tmp_path / "out" / "overlaid.pdf"

    overlays = [
        {
            "name": "Date",
            "string": "2026/11/21",
            "param": {"X": "20mm", "Y": "25mm", "Font": "Helvetica-Bold", "FontSize": 11},
        },
        {
            "name": "HeaderTopLeft",
            "string": "TOP HEADER",
            "param": {"X": 40, "Y": 40, "Origin": "top-left", "Font": "Courier", "FontSize": 10},
        },
        {
            "name": "AmountInWords",
            "string": "Two hundred and fifty thousand rupees only",
            "param": {
                "X": "20mm",
                "Y": "100mm",
                "Font": "Helvetica",
                "FontSize": 12,
                "LineSpace": "1.3X",
                "Function": {"name": "SrinkToFit", "param1": "60mm", "param2": 2, "param3": "5mm"},
            },
        },
    ]

    drawn = add_text_overlays_to_pdf(tpl, out, overlays)
    assert drawn == 3
    assert out.exists()

    # Verify caller params were not mutated
    assert overlays[0]["param"]["X"] == "20mm"

    doc = pymupdf.open(out)
    try:
        text = doc[0].get_text()
        assert "2026/11/21" in text
        assert "TOP HEADER" in text
        assert "rupees only" in text
    finally:
        doc.close()


def test_add_text_overlays_error_cases(tmp_path: Path) -> None:
    tpl = _create_blank_pdf(tmp_path / "tpl.pdf")

    with pytest.raises(PdfRenderError, match="not found"):
        add_text_overlays_to_pdf(tmp_path / "missing.pdf", tmp_path / "out.pdf", [])

    with pytest.raises(PdfRenderError, match="out of range"):
        add_text_overlays_to_pdf(tpl, tmp_path / "out.pdf", [], page_index=2)

    with pytest.raises(PdfRenderError, match="missing required 'param'"):
        add_text_overlays_to_pdf(tpl, tmp_path / "out.pdf", [{"name": "Bad", "string": "x"}])

    with pytest.raises(PdfRenderError, match="both 'X' and 'Y'"):
        add_text_overlays_to_pdf(
            tpl, tmp_path / "out.pdf", [{"name": "Bad", "string": "x", "param": {"X": 10}}]
        )
