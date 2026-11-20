"""Unit tests for TEMPLATE.xlsx parsing (Data & Overlay sheets)."""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from pdf_bulk_fill import (
    OverlayDefinition,
    RecordEntry,
    TemplateParseError,
    apply_concat,
    get_required_source_files,
    load_overlay_definitions,
    load_record_list,
    parse_kv_pairs,
)


def _create_template_workbook(path: Path) -> Path:
    wb = openpyxl.Workbook()

    data_ws = wb.active
    data_ws.title = "Data"
    data_ws.append(["Index", "PrimaryKey", "ColC", "ColD", "Identifier"])
    data_ws.append([1, 1001, None, None, "Kamal"])
    data_ws.append([2, "EMP-002", None, None, "Nimal"])
    data_ws.append([3, 1003, None, None, None])  # fallback identifier to key
    data_ws.append([None, 1004, None, None, "Stopped"])

    overlay_ws = wb.create_sheet("Overlay")
    overlay_ws.append(["Index", "Name", "Content", "Param", "PreProcess"])
    overlay_ws.append([
        1,
        "EmpName",
        "<Type=File><File=EMP01.xlsx><Sheet=Employees><PrimeryKey=A><Value=C>",
        "<X=10mm><Y=20mm><Font=Helvetica><FontSize=12><Function=SrinkToFit(80mm, 2, 5mm)>",
        "<Function=changeTextCase(text, UPPER)>",
    ])
    overlay_ws.append([
        2,
        "!<CONCAT><EmpName>",
        "<Type=Text><Text= - Verified>",
        None,
        None,
    ])
    overlay_ws.append([
        3,
        "NIC",
        "<Type=FormField><File=EMP01.xlsx><Sheet=Employees><PrimaryKey=A><Value=B>",
        "<Field=NIC>",
        "<Function=changeTextCase(text, UPPER)>",
    ])
    overlay_ws.append([
        4,
        "IDFront",
        "<Type=Image><File=IMAGES.xlsx><Sheet=Photos><PrimaryKey=A><Value=B>",
        "<Slot=ID_FRONT>",
        None,
    ])
    overlay_ws.append([5, "StoppedRow", None, None, None])  # user-terminated

    wb.save(path)
    wb.close()
    return path


def test_parse_kv_pairs_scalars_and_functions() -> None:
    assert parse_kv_pairs(None) is None
    assert parse_kv_pairs("") is None
    assert parse_kv_pairs("None") is None
    assert parse_kv_pairs("no brackets here") is None

    parsed = parse_kv_pairs(
        "<X=100><Y=12.5><Negative=-4><Font=Helvetica><Function=FormatNumber(text, 2, Rs., /=)>"
    )
    assert parsed == {
        "X": 100,
        "Y": 12.5,
        "Negative": -4,
        "Font": "Helvetica",
        "Function": {
            "name": "FormatNumber",
            "param1": "text",
            "param2": 2,
            "param3": "Rs.",
            "param4": "/=",
        },
    }


def test_load_record_list_numeric_and_alphanumeric(tmp_path: Path) -> None:
    tpl = _create_template_workbook(tmp_path / "TEMPLATE.xlsx")
    records = load_record_list(tpl)
    assert records == [
        RecordEntry(key="1001", identifier="Kamal"),
        RecordEntry(key="EMP-002", identifier="Nimal"),
        RecordEntry(key="1003", identifier="1003"),
    ]


def test_load_record_list_stops_on_empty_key(tmp_path: Path) -> None:
    tpl = tmp_path / "stop.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Index", "Key", "", "", "Name"])
    ws.append([1, "A1", "", "", "First"])
    ws.append([2, None, "", "", "Second"])
    ws.append([3, "A3", "", "", "Third"])
    wb.save(tpl)
    wb.close()

    assert load_record_list(tpl) == [RecordEntry(key="A1", identifier="First")]


def test_load_overlay_definitions_and_source_files(tmp_path: Path) -> None:
    tpl = _create_template_workbook(tmp_path / "TEMPLATE.xlsx")
    overlays = load_overlay_definitions(tpl)

    assert len(overlays) == 4
    assert overlays[0].name == "EmpName"
    assert overlays[0].content_type == "File"
    assert overlays[0].content["PrimaryKey"] == "A"
    assert overlays[0].params["Function"] == {
        "name": "SrinkToFit",
        "param1": "80mm",
        "param2": 2,
        "param3": "5mm",
    }
    assert overlays[1].is_concat is True
    assert overlays[2].content_type == "FormField"
    assert overlays[3].content_type == "Image"

    assert get_required_source_files(overlays) == ["EMP01.xlsx", "IMAGES.xlsx"]


def test_load_template_error_cases(tmp_path: Path) -> None:
    with pytest.raises(TemplateParseError, match="not found"):
        load_record_list(tmp_path / "missing.xlsx")

    with pytest.raises(TemplateParseError, match="not found"):
        load_overlay_definitions(tmp_path / "missing.xlsx")

    # Missing sheets
    empty_wb_path = tmp_path / "empty.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Other"
    wb.save(empty_wb_path)
    wb.close()

    with pytest.raises(TemplateParseError, match="Sheet 'Data' not found"):
        load_record_list(empty_wb_path)
    with pytest.raises(TemplateParseError, match="Sheet 'Overlay' not found"):
        load_overlay_definitions(empty_wb_path)


@pytest.mark.parametrize(
    "row_data,match",
    [
        ([1, "BadContent", "not_kv_format", None, None], "Malformed Content"),
        ([1, "BadType", "<Type=Unknown><Text=x>", None, None], "Unsupported overlay Type"),
        ([1, None, "<Type=Text><Text=x>", None, None], "Missing overlay name"),
    ],
)
def test_load_overlay_malformed_rows(tmp_path: Path, row_data: list, match: str) -> None:
    tpl = tmp_path / "bad_overlay.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Overlay"
    ws.append(["Index", "Name", "Content", "Param", "PreProcess"])
    ws.append(row_data)
    wb.save(tpl)
    wb.close()

    with pytest.raises(TemplateParseError, match=match):
        load_overlay_definitions(tpl)


def test_apply_concat() -> None:
    resolved = [{"name": "Address", "string": "123 Main St"}]
    apply_concat(resolved, "!<CONCAT><Address>", ", Colombo")
    assert resolved[0]["string"] == "123 Main St, Colombo"

    with pytest.raises(TemplateParseError, match="Invalid concatenation directive"):
        apply_concat(resolved, "!<CONCAT>BadFormat", "x")

    with pytest.raises(TemplateParseError, match="not found"):
        apply_concat(resolved, "!<CONCAT><MissingField>", "x")
