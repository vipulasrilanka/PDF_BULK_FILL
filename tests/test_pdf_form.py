"""Unit & integration tests for AcroForm comb-field and checkbox filling on RR-Form.pdf."""

from __future__ import annotations

from pathlib import Path

import pymupdf
import pytest

from pdf_bulk_fill import (
    PdfRenderError,
    expand_comb_groups,
    fill_acroform,
    is_truthy_checkbox,
    list_form_fields,
)

RR_FORM_PATH = Path(__file__).resolve().parent.parent / "templates" / "pdf" / "RR-Form.pdf"


def test_is_truthy_checkbox() -> None:
    for val in (True, 1, "1", "Yes", "yes", "Y", "true", "ON", "X", "✓", "√"):
        assert is_truthy_checkbox(val) is True
    for val in (False, None, 0, "", "0", "No", "n", "false", "Off"):
        assert is_truthy_checkbox(val) is False

    with pytest.raises(PdfRenderError, match="Invalid checkbox value"):
        is_truthy_checkbox("maybe")


def test_list_form_fields_on_rr_form() -> None:
    fields = list_form_fields(RR_FORM_PATH)
    assert len(fields) == 14
    assert fields["Zone"]["max_len"] == 1
    assert fields["Zone"]["is_comb"] is True
    assert fields["NIC"]["max_len"] == 12
    assert fields["Name 1"]["max_len"] == 22
    assert fields["Name 2"]["max_len"] == 31
    assert fields["Name 3"]["max_len"] == 31
    assert fields["Name INI1"]["max_len"] == 20
    assert fields["Name INI 2"]["max_len"] == 31
    assert fields["Mobile"]["max_len"] == 10
    assert fields["Email"]["max_len"] == 27
    assert fields["Internet"]["type"] == "CheckBox"
    assert fields["SMS"]["type"] == "CheckBox"


def test_expand_comb_groups() -> None:
    expanded = expand_comb_groups(
        {
            "FullName": "don kamal bandara wickramasinghe mudiyanselage perera",
            "NameWithInitials": "D. K. B. Perera",
            "NIC": "901234567V",
        }
    )
    assert expanded["Name 1"] == "DON KAMAL BANDARA"
    assert expanded["Name 2"] == "WICKRAMASINGHE MUDIYANSELAGE"
    assert expanded["Name 3"] == "PERERA"
    assert expanded["Name INI1"] == "D. K. B. PERERA"
    assert expanded["Name INI 2"] == ""
    assert expanded["NIC"] == "901234567V"


def test_fill_rr_form_editable_mode(tmp_path: Path) -> None:
    out_pdf = tmp_path / "out" / "filled_editable.pdf"
    payload = {
        "Serial No": "0042",
        "Zone": "A",
        "Emp No": "123456",
        "Mem No": "000789",
        "NIC": "199012345678",
        "FullName": "KAMAL BANDARA WICKRAMASINGHE PERERA",
        "NameWithInitials": "K. B. W. PERERA",
        "Mobile": "0771234567",
        "Email": "KAMAL@EXAMPLE.COM",
        "Internet": "Yes",
        "SMS": False,
    }

    applied = fill_acroform(RR_FORM_PATH, out_pdf, payload, flatten=False)
    assert out_pdf.exists()
    assert applied["Zone"] == "A"
    assert applied["NIC"] == "199012345678"
    assert applied["Internet"] is True
    assert applied["SMS"] is False

    # Re-open filled PDF and verify widgets still exist with expected values
    doc = pymupdf.open(out_pdf)
    try:
        widgets = {w.field_name: w for w in doc[0].widgets()}
        assert widgets["Serial No"].field_value == "0042"
        assert widgets["Zone"].field_value == "A"
        assert widgets["Emp No"].field_value == "123456"
        assert widgets["Mem No"].field_value == "000789"
        assert widgets["NIC"].field_value == "199012345678"
        assert widgets["Name 1"].field_value == "KAMAL BANDARA"
        assert widgets["Name 2"].field_value == "WICKRAMASINGHE PERERA"
        assert widgets["Name 3"].field_value == ""
        assert widgets["Name INI1"].field_value == "K. B. W. PERERA"
        assert widgets["Mobile"].field_value == "0771234567"
        assert widgets["Email"].field_value == "KAMAL@EXAMPLE.COM"
        assert widgets["Internet"].field_value == "Yes"
        assert widgets["SMS"].field_value == "Off"
    finally:
        doc.close()


def test_fill_rr_form_flattened_mode(tmp_path: Path) -> None:
    out_pdf = tmp_path / "filled_flat.pdf"
    fill_acroform(
        RR_FORM_PATH,
        out_pdf,
        {
            "Zone": "B",
            "Emp No": "654321",
            "NIC": "901234567V",
            "Mobile": "0719876543",
            "Internet": True,
            "SMS": True,
        },
        flatten=True,
    )
    assert out_pdf.exists()

    doc = pymupdf.open(out_pdf)
    try:
        # Flattened PDF has no remaining interactive widgets
        assert list(doc[0].widgets() or []) == []
        # And the rendered page text contains the filled values
        page_text = doc[0].get_text()
        assert "9" in page_text and "V" in page_text
    finally:
        doc.close()


def test_fill_acroform_maxlen_clamping_and_strict_mode(tmp_path: Path) -> None:
    out_pdf = tmp_path / "clamped.pdf"
    applied = fill_acroform(
        RR_FORM_PATH,
        out_pdf,
        {"Zone": "ABC", "Mobile": "0771234567899"},
        strict_maxlen=False,
    )
    assert applied["Zone"] == "A"
    assert applied["Mobile"] == "0771234567"

    with pytest.raises(PdfRenderError, match="exceeds MaxLen"):
        fill_acroform(
            RR_FORM_PATH,
            tmp_path / "strict.pdf",
            {"Zone": "AB"},
            strict_maxlen=True,
        )


def test_fill_acroform_error_cases(tmp_path: Path) -> None:
    with pytest.raises(PdfRenderError, match="not found"):
        list_form_fields(tmp_path / "missing.pdf")

    with pytest.raises(PdfRenderError, match="out of range"):
        list_form_fields(RR_FORM_PATH, page_index=5)

    with pytest.raises(PdfRenderError, match="not found"):
        fill_acroform(tmp_path / "missing.pdf", tmp_path / "out.pdf", {})

    with pytest.raises(PdfRenderError, match="out of range"):
        fill_acroform(RR_FORM_PATH, tmp_path / "out.pdf", {}, page_index=3)

    with pytest.raises(PdfRenderError, match="Unknown PDF form field"):
        fill_acroform(
            RR_FORM_PATH,
            tmp_path / "out.pdf",
            {"DoesNotExist": "123"},
            strict_unknown_fields=True,
        )
