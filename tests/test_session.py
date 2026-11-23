"""Unit tests for session configuration persistence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pdf_bulk_fill import (
    SessionConfig,
    SessionError,
    SourceFileEntry,
    load_session,
    resolve_source_path,
    save_session,
)


def test_save_and_load_session_roundtrip(tmp_path: Path) -> None:
    session_file = tmp_path / "nested" / "session.json"
    original = SessionConfig(
        pdf_file="/templates/RR-Form.pdf",
        template_file="/templates/TEMPLATE.xlsx",
        source_files=[
            SourceFileEntry(name="employees.xlsx", path="/data/input"),
            SourceFileEntry(name="payroll.xlsx", path="/data/payroll"),
        ],
        output_dir="/data/output",
    )

    written = save_session(session_file, original)
    assert written == session_file
    assert session_file.exists()

    loaded = load_session(session_file)
    assert loaded == original


def test_load_legacy_camel_case_session(tmp_path: Path) -> None:
    session_file = tmp_path / "session.json"
    legacy_payload = {
        "pdfFileName": "/path/to/Letterhead.pdf",
        "templateFileName": "/path/to/TEMPLATE.xlsx",
        "sourceFiles": [
            {"name": "2023_2024_WHT_Working.xlsx", "path": "/path/to/WHT"}
        ],
    }
    session_file.write_text(json.dumps(legacy_payload), encoding="utf-8")

    loaded = load_session(session_file)
    assert loaded.pdf_file == "/path/to/Letterhead.pdf"
    assert loaded.template_file == "/path/to/TEMPLATE.xlsx"
    assert loaded.source_files == [
        SourceFileEntry(name="2023_2024_WHT_Working.xlsx", path="/path/to/WHT")
    ]
    assert loaded.output_dir is None


def test_load_snake_case_session(tmp_path: Path) -> None:
    session_file = tmp_path / "session.json"
    payload = {
        "pdf_file": "templates/pdf/RR-Form.pdf",
        "template_file": "templates/excel/TEMPLATE.xlsx",
        "source_files": [],
        "output_dir": "data/output",
    }
    session_file.write_text(json.dumps(payload), encoding="utf-8")

    loaded = load_session(session_file)
    assert loaded.pdf_file == "templates/pdf/RR-Form.pdf"
    assert loaded.template_file == "templates/excel/TEMPLATE.xlsx"
    assert loaded.source_files == []
    assert loaded.output_dir == "data/output"


def test_load_session_missing_file(tmp_path: Path) -> None:
    with pytest.raises(SessionError, match="not found"):
        load_session(tmp_path / "does_not_exist.json")


def test_load_session_malformed_json(tmp_path: Path) -> None:
    bad_file = tmp_path / "bad.json"
    bad_file.write_text("{not valid json", encoding="utf-8")

    with pytest.raises(SessionError, match="Malformed JSON"):
        load_session(bad_file)


def test_load_session_non_object_json(tmp_path: Path) -> None:
    bad_file = tmp_path / "list.json"
    bad_file.write_text("[]", encoding="utf-8")

    with pytest.raises(SessionError, match="JSON object"):
        load_session(bad_file)


@pytest.mark.parametrize(
    "payload,match",
    [
        ({"templateFileName": "t.xlsx"}, "pdfFileName"),
        ({"pdfFileName": "f.pdf"}, "templateFileName"),
        ({"pdfFileName": "f.pdf", "templateFileName": "t.xlsx", "sourceFiles": "bad"}, "must be a list"),
        ({"pdfFileName": "f.pdf", "templateFileName": "t.xlsx", "sourceFiles": ["bad"]}, "must be an object"),
        ({"pdfFileName": "f.pdf", "templateFileName": "t.xlsx", "sourceFiles": [{"name": "", "path": "/a"}]}, "name"),
        ({"pdfFileName": "f.pdf", "templateFileName": "t.xlsx", "sourceFiles": [{"name": "a.xlsx", "path": ""}]}, "path"),
        ({"pdfFileName": "f.pdf", "templateFileName": "t.xlsx", "outputDir": 123}, "outputDir"),
        ({"pdfFileName": "f.pdf", "templateFileName": "t.xlsx", "flatten": "yes"}, "flatten"),
        ({"pdfFileName": "f.pdf", "templateFileName": "t.xlsx", "pageIndex": True}, "pageIndex"),
        ({"pdfFileName": "f.pdf", "templateFileName": "t.xlsx", "sourceFiles": [{"name": "a.xlsx", "path": "/a", "filePath": ""}]}, "filePath"),
    ],
)
def test_load_session_validation_errors(tmp_path: Path, payload: dict, match: str) -> None:
    session_file = tmp_path / "invalid.json"
    session_file.write_text(json.dumps(payload), encoding="utf-8")

    with pytest.raises(SessionError, match=match):
        load_session(session_file)


def test_session_full_options_roundtrip(tmp_path: Path) -> None:
    session_file = tmp_path / "full_session.json"
    original = SessionConfig(
        pdf_file="/tpl/RR-Form.pdf",
        template_file="/tpl/TEMPLATE.xlsx",
        source_files=[
            SourceFileEntry(
                name="employees.xlsx",
                path="/data/alt",
                file_path="/data/alt/staff_2026.xlsx",
            )
        ],
        output_dir="/data/out",
        source_dir="/data/src",
        image_dir="/data/imgs",
        page_index=1,
        flatten=True,
        skip_missing_images=True,
        allow_truncation=True,
    )
    save_session(session_file, original)
    loaded = load_session(session_file)
    assert loaded == original
    assert resolve_source_path("employees.xlsx", loaded, "/default") == Path(
        "/data/alt/staff_2026.xlsx"
    )


def test_resolve_source_path() -> None:
    config = SessionConfig(
        pdf_file="form.pdf",
        template_file="template.xlsx",
        source_files=[SourceFileEntry(name="emp.xlsx", path="/custom/dir")],
    )

    assert resolve_source_path("emp.xlsx", config, "/default/dir") == Path("/custom/dir/emp.xlsx")
    assert resolve_source_path("other.xlsx", config, "/default/dir") == Path("/default/dir/other.xlsx")
