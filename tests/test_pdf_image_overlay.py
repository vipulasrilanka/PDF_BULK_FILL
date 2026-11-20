"""Unit & integration tests for employee ID card (front & back) image overlays."""

from __future__ import annotations

from pathlib import Path

import pymupdf
import pytest
from PIL import Image, ImageDraw

from pdf_bulk_fill import (
    ImageOverlayError,
    add_image_overlays_to_pdf,
    fit_rect_preserving_aspect,
    resolve_image_path,
    resolve_target_rect,
)

RR_FORM_PATH = Path(__file__).resolve().parent.parent / "templates" / "pdf" / "RR-Form.pdf"


def _make_id_image(path: Path, size: tuple[int, int], color: tuple[int, int, int], label: str) -> Path:
    img = Image.new("RGB", size, color=color)
    draw = ImageDraw.Draw(img)
    draw.rectangle([10, 10, size[0] - 10, size[1] - 10], outline=(20, 40, 80), width=4)
    draw.text((30, size[1] // 2 - 10), label, fill=(20, 20, 20))
    img.save(path)
    return path


def test_fit_rect_preserving_aspect() -> None:
    target = pymupdf.Rect(100, 100, 400, 300)  # 300 x 200
    # Landscape 600 x 300 (2:1) -> fits width 300, height 150, centered vertically
    fitted = fit_rect_preserving_aspect(target, 600, 300)
    assert fitted.width == pytest.approx(300.0)
    assert fitted.height == pytest.approx(150.0)
    assert fitted.y0 == pytest.approx(125.0)
    assert fitted.x0 == pytest.approx(100.0)

    # Portrait 100 x 200 (1:2) -> fits height 200, width 100, centered horizontally
    fitted_p = fit_rect_preserving_aspect(target, 100, 200)
    assert fitted_p.width == pytest.approx(100.0)
    assert fitted_p.height == pytest.approx(200.0)
    assert fitted_p.x0 == pytest.approx(200.0)

    with pytest.raises(ImageOverlayError):
        fit_rect_preserving_aspect(target, 0, 100)
    with pytest.raises(ImageOverlayError):
        fit_rect_preserving_aspect(pymupdf.Rect(0, 0, 0, 100), 100, 100)


def test_resolve_image_path_direct_pattern_and_extension_probe(tmp_path: Path) -> None:
    front = _make_id_image(tmp_path / "1001_front.png", (400, 250), (220, 235, 255), "FRONT")
    back = _make_id_image(tmp_path / "1001_back.jpg", (400, 250), (235, 255, 220), "BACK")

    assert resolve_image_path(front) == front
    assert resolve_image_path("{key}_front.png", base_dir=tmp_path, record_key="1001") == front
    # Probe extension automatically when omitted
    assert resolve_image_path("{key}_back", base_dir=tmp_path, record_key="1001") == back

    with pytest.raises(ImageOverlayError, match="Empty image path"):
        resolve_image_path("   ")
    with pytest.raises(ImageOverlayError, match="Image file not found"):
        resolve_image_path("missing_image.png", base_dir=tmp_path)


def test_resolve_target_rect_slots_and_custom_coords() -> None:
    rect = resolve_target_rect({"Slot": "id_front"}, page_height=841.8)
    assert rect == pymupdf.Rect(120.0, 95.0, 475.0, 275.0)

    custom_tl = resolve_target_rect(
        {"X": "10mm", "Y": "20mm", "Width": "50mm", "Height": "30mm"},
        page_height=842.0,
    )
    assert custom_tl.width == pytest.approx(50 * 72 / 25.4)
    assert custom_tl.height == pytest.approx(30 * 72 / 25.4)

    custom_bl = resolve_target_rect(
        {"X": 50, "Y": 100, "Width": 200, "Height": 100, "Origin": "bottom-left"},
        page_height=800.0,
    )
    assert custom_bl == pymupdf.Rect(50, 600, 250, 700)

    with pytest.raises(ImageOverlayError, match="Unknown image slot"):
        resolve_target_rect({"Slot": "NO_SUCH_SLOT"}, page_height=842.0)
    with pytest.raises(ImageOverlayError, match="missing"):
        resolve_target_rect({"X": 10, "Y": 10}, page_height=842.0)
    with pytest.raises(ImageOverlayError, match="must be positive"):
        resolve_target_rect({"X": 10, "Y": 10, "Width": 0, "Height": 20}, page_height=842.0)


def test_add_front_and_back_id_images_to_rr_form(tmp_path: Path) -> None:
    front = _make_id_image(tmp_path / "EMP01_front.png", (640, 400), (215, 232, 255), "NIC FRONT")
    back = _make_id_image(tmp_path / "EMP01_back.jpg", (640, 400), (225, 245, 220), "NIC BACK")
    out_pdf = tmp_path / "out" / "rr_with_ids.pdf"

    inserted = add_image_overlays_to_pdf(
        RR_FORM_PATH,
        out_pdf,
        [
            {"name": "IDFront", "path": "{key}_front", "param": {"Slot": "ID_FRONT"}},
            {"name": "IDBack", "path": str(back), "param": {"Slot": "ID_BACK"}},
        ],
        base_dir=tmp_path,
        record_key="EMP01",
    )
    assert inserted == ["IDFront", "IDBack"]
    assert out_pdf.exists()

    # Verify two additional images were embedded on page 0
    orig_doc = pymupdf.open(RR_FORM_PATH)
    new_doc = pymupdf.open(out_pdf)
    try:
        orig_imgs = len(orig_doc[0].get_images())
        new_imgs = len(new_doc[0].get_images())
        assert new_imgs >= orig_imgs + 2
    finally:
        orig_doc.close()
        new_doc.close()


def test_add_image_overlays_skip_missing_and_errors(tmp_path: Path) -> None:
    out_pdf = tmp_path / "out.pdf"

    # skip_missing=True ignores missing image without error
    inserted = add_image_overlays_to_pdf(
        RR_FORM_PATH,
        out_pdf,
        [{"name": "Missing", "path": "nope.png", "param": {"Slot": "ID_FRONT"}}],
        base_dir=tmp_path,
        skip_missing=True,
    )
    assert inserted == []

    # Corrupt image raises ImageOverlayError
    corrupt = tmp_path / "corrupt.jpg"
    corrupt.write_bytes(b"not-a-jpeg")
    with pytest.raises(ImageOverlayError, match="Invalid or corrupt image"):
        add_image_overlays_to_pdf(
            RR_FORM_PATH,
            out_pdf,
            [{"name": "Corrupt", "path": str(corrupt), "param": {"Slot": "ID_FRONT"}}],
        )

    # Missing template PDF / out-of-range page / missing param
    with pytest.raises(ImageOverlayError, match="PDF template not found"):
        add_image_overlays_to_pdf(tmp_path / "missing.pdf", out_pdf, [])
    with pytest.raises(ImageOverlayError, match="out of range"):
        add_image_overlays_to_pdf(RR_FORM_PATH, out_pdf, [], page_index=4)
    with pytest.raises(ImageOverlayError, match="missing required 'param'"):
        add_image_overlays_to_pdf(RR_FORM_PATH, out_pdf, [{"name": "NoParam", "path": "x.png"}])
