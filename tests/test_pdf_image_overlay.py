"""Unit & integration tests for generic DocType-matched & constant image overlays."""

from __future__ import annotations

from pathlib import Path

import pymupdf
import pytest
from PIL import Image, ImageDraw

from pdf_bulk_fill import (
    ImageAsset,
    ImageOverlayError,
    add_image_overlays_to_pdf,
    fit_rect_preserving_aspect,
    parse_image_asset,
    resolve_image_file,
    resolve_target_rect,
)

RR_FORM_PATH = Path(__file__).resolve().parent.parent / "templates" / "pdf" / "RR-Form.pdf"


def _make_id_image(
    path: Path, size: tuple[int, int], color: tuple[int, int, int], label: str
) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGB", size, color=color)
    draw = ImageDraw.Draw(img)
    draw.rectangle([10, 10, size[0] - 10, size[1] - 10], outline=(20, 40, 80), width=4)
    draw.text((20, size[1] // 2 - 10), label, fill=(20, 20, 20))
    img.save(path)
    return path


def test_parse_image_asset_kv_dict_and_defaults() -> None:
    assert parse_image_asset(None) is None
    assert parse_image_asset("   ") is None

    asset = parse_image_asset(
        "<DocType=NEW><FileName=1001_front.png><Location=ids><Rotation=90><KeepAspect=False>"
    )
    assert asset == ImageAsset(
        file_name="1001_front.png",
        doc_type="NEW",
        location="ids",
        rotation=90.0,
        keep_aspect=False,
    )

    # Defaults when optional keys omitted
    minimal = parse_image_asset("<DocType=OLD><FileName=1002_front>")
    assert minimal == ImageAsset(
        file_name="1002_front",
        doc_type="OLD",
        location="",
        rotation=0.0,
        keep_aspect=True,
    )

    # Constant image dict (no DocType)
    const_asset = parse_image_asset({"FileName": "logo.png", "KeepAspect": True})
    assert const_asset == ImageAsset(
        file_name="logo.png", doc_type=None, location="", rotation=0.0, keep_aspect=True
    )

    # Pass-through existing ImageAsset
    assert parse_image_asset(const_asset) is const_asset

    with pytest.raises(ImageOverlayError, match="Malformed image metadata"):
        parse_image_asset("plain_string_without_brackets")
    with pytest.raises(ImageOverlayError, match="Unsupported image metadata type"):
        parse_image_asset(12345)
    with pytest.raises(ImageOverlayError, match="missing required 'FileName'"):
        parse_image_asset("<DocType=NEW>")
    with pytest.raises(ImageOverlayError, match="Invalid Rotation"):
        parse_image_asset("<FileName=a.png><Rotation=bad>")
    with pytest.raises(ImageOverlayError, match="Invalid boolean"):
        parse_image_asset("<FileName=a.png><KeepAspect=maybe>")


def test_fit_rect_preserving_aspect() -> None:
    target = pymupdf.Rect(100, 100, 400, 300)  # 300 x 200
    fitted = fit_rect_preserving_aspect(target, 600, 300)
    assert fitted.width == pytest.approx(300.0)
    assert fitted.height == pytest.approx(150.0)
    assert fitted.y0 == pytest.approx(125.0)
    assert fitted.x0 == pytest.approx(100.0)

    fitted_p = fit_rect_preserving_aspect(target, 100, 200)
    assert fitted_p.width == pytest.approx(100.0)
    assert fitted_p.height == pytest.approx(200.0)
    assert fitted_p.x0 == pytest.approx(200.0)

    with pytest.raises(ImageOverlayError):
        fit_rect_preserving_aspect(target, 0, 100)
    with pytest.raises(ImageOverlayError):
        fit_rect_preserving_aspect(pymupdf.Rect(0, 0, 0, 100), 100, 100)


def test_resolve_image_file_locations_and_extension_probe(tmp_path: Path) -> None:
    ids_dir = tmp_path / "ids"
    front = _make_id_image(ids_dir / "1001_front.png", (400, 250), (220, 235, 255), "FRONT")
    back = _make_id_image(ids_dir / "1001_back.jpg", (400, 250), (235, 255, 220), "BACK")

    # Direct absolute path
    assert resolve_image_file(front) == front
    # Relative location joined with base_dir
    assert resolve_image_file("1001_front.png", location="ids", base_dir=tmp_path) == front
    # Deduplicated location when base_dir already ends in 'ids'
    assert resolve_image_file("1001_front.png", location="ids", base_dir=ids_dir) == front
    # Absolute location + extension probing
    assert resolve_image_file("1001_back", location=ids_dir) == back
    # Empty location -> base_dir
    assert resolve_image_file("1001_back", base_dir=ids_dir) == back

    with pytest.raises(ImageOverlayError, match="Empty image file name"):
        resolve_image_file("   ")
    with pytest.raises(ImageOverlayError, match="Image file not found"):
        resolve_image_file("missing.png", location="ids", base_dir=tmp_path)


def test_resolve_target_rect_explicit_coords() -> None:
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

    with pytest.raises(ImageOverlayError, match="missing"):
        resolve_target_rect({"X": 10, "Y": 10}, page_height=842.0)
    with pytest.raises(ImageOverlayError, match="must be positive"):
        resolve_target_rect({"X": 10, "Y": 10, "Width": 0, "Height": 20}, page_height=842.0)


def test_doctype_matching_skip_warning_rotation_and_constant_logo(tmp_path: Path) -> None:
    ids_dir = tmp_path / "ids"
    _make_id_image(ids_dir / "1001_front.png", (640, 400), (215, 232, 255), "NEW FRONT")
    _make_id_image(ids_dir / "1001_back.jpg", (640, 400), (225, 245, 220), "NEW BACK")
    _make_id_image(tmp_path / "logo.png", (200, 80), (40, 80, 160), "LOGO")
    out_pdf = tmp_path / "out" / "rr_matched.pdf"

    warnings: list[str] = []
    inserted = add_image_overlays_to_pdf(
        RR_FORM_PATH,
        out_pdf,
        [
            # OLD slots -> should warn & skip because employee asset DocType is NEW
            {
                "name": "IDFront_old",
                "asset": "<DocType=NEW><FileName=1001_front.png><Location=ids>",
                "param": {"X": 45, "Y": 140, "Width": 245, "Height": 250, "DocType": "OLD"},
            },
            # Empty cell -> should warn & skip
            {
                "name": "IDOptional",
                "asset": "",
                "param": {"X": 45, "Y": 140, "Width": 100, "Height": 100, "DocType": "OLD"},
            },
            # NEW slots -> should match and insert (one with 90° rotation, one with KeepAspect=False)
            {
                "name": "IDFront_new",
                "asset": "<DocType=new><FileName=1001_front.png><Location=ids><Rotation=90>",
                "param": {"X": 120, "Y": 95, "Width": 355, "Height": 180, "DocType": "NEW"},
            },
            {
                "name": "IDBack_new",
                "asset": "<DocType=NEW><FileName=1001_back><Location=ids><KeepAspect=False>",
                "param": {"X": 120, "Y": 295, "Width": 355, "Height": 180, "DocType": "NEW"},
            },
            # Constant company logo (no DocType) -> always inserted
            {
                "name": "CompanyLogo",
                "asset": {"FileName": "logo.png", "Rotation": 0, "KeepAspect": True},
                "param": {"X": 30, "Y": 20, "Width": 60, "Height": 25},
            },
        ],
        base_dir=tmp_path,
        warnings_out=warnings,
    )

    assert inserted == ["IDFront_new", "IDBack_new", "CompanyLogo"]
    assert len(warnings) == 2
    assert "does not match slot DocType 'OLD'" in warnings[0]
    assert "no image metadata" in warnings[1]
    assert out_pdf.exists()

    orig_doc = pymupdf.open(RR_FORM_PATH)
    new_doc = pymupdf.open(out_pdf)
    try:
        assert len(new_doc[0].get_images()) >= len(orig_doc[0].get_images()) + 3
    finally:
        orig_doc.close()
        new_doc.close()


def test_add_image_overlays_skip_missing_and_errors(tmp_path: Path) -> None:
    out_pdf = tmp_path / "out.pdf"
    warnings: list[str] = []

    inserted = add_image_overlays_to_pdf(
        RR_FORM_PATH,
        out_pdf,
        [
            {
                "name": "Missing",
                "asset": "<FileName=nope.png>",
                "param": {"X": 120, "Y": 95, "Width": 355, "Height": 180},
            }
        ],
        base_dir=tmp_path,
        skip_missing=True,
        warnings_out=warnings,
    )
    assert inserted == []
    assert len(warnings) == 1

    corrupt = tmp_path / "corrupt.jpg"
    corrupt.write_bytes(b"not-a-jpeg")
    with pytest.raises(ImageOverlayError, match="Invalid or corrupt image"):
        add_image_overlays_to_pdf(
            RR_FORM_PATH,
            out_pdf,
            [
                {
                    "name": "Corrupt",
                    "asset": {"FileName": str(corrupt)},
                    "param": {"X": 120, "Y": 95, "Width": 355, "Height": 180},
                }
            ],
        )

    with pytest.raises(ImageOverlayError, match="PDF template not found"):
        add_image_overlays_to_pdf(tmp_path / "missing.pdf", out_pdf, [])
    with pytest.raises(ImageOverlayError, match="out of range"):
        add_image_overlays_to_pdf(RR_FORM_PATH, out_pdf, [], page_index=4)
    with pytest.raises(ImageOverlayError, match="missing required 'param'"):
        add_image_overlays_to_pdf(
            RR_FORM_PATH, out_pdf, [{"name": "NoParam", "asset": "<FileName=x.png>"}]
        )
