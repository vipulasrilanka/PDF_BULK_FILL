"""End-to-end integration tests for batch runner, CLI, and GUI dialogs."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl
import pymupdf
import pytest
from PIL import Image, ImageDraw

from pdf_bulk_fill import (
    BatchSummary,
    PdfRenderError,
    SessionConfig,
    load_session,
    render_record_pdf,
    run_batch,
    sanitize_filename_part,
    save_session,
)
from pdf_bulk_fill.cli import main as cli_main
from pdf_bulk_fill.gui import pick_excel_file, pick_pdf_file, prompt_password

RR_FORM_PATH = Path(__file__).resolve().parent.parent / "templates" / "pdf" / "RR-Form.pdf"


def _make_id_image(path: Path, color: tuple[int, int, int], label: str) -> Path:
    img = Image.new("RGB", (400, 250), color=color)
    draw = ImageDraw.Draw(img)
    draw.rectangle([8, 8, 392, 242], outline=(20, 40, 80), width=4)
    draw.text((120, 115), label, fill=(10, 10, 10))
    img.save(path)
    return path


def _build_rr_form_fixture(base_dir: Path) -> tuple[Path, Path, Path]:
    """Create a complete employees.xlsx, TEMPLATE.xlsx, and front/back ID images."""
    img_dir = base_dir / "ids"
    img_dir.mkdir(parents=True, exist_ok=True)
    _make_id_image(img_dir / "1001_front.png", (215, 232, 255), "1001 FRONT")
    _make_id_image(img_dir / "1001_back.png", (225, 245, 220), "1001 BACK")
    _make_id_image(img_dir / "1002_front.jpg", (255, 235, 215), "1002 FRONT")
    _make_id_image(img_dir / "1002_back.jpg", (245, 225, 255), "1002 BACK")

    # Source employee workbook
    emp_xlsx = base_dir / "employees.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Employees"
    ws.append([
        "EmpID", "SerialNo", "Zone", "EmpNo", "MemNo", "NIC",
        "FullName", "NameINI", "Mobile", "Email", "Internet", "SMS", "CertDate",
    ])
    ws.append([
        1001, "0001", "A", "123456", "000789", "199012345678",
        "don kamal bandara wickramasinghe mudiyanselage perera",
        "d. k. b. w. m. perera",
        "0771234567", "kamal@example.com", "Yes", "Yes", "2026-11-21 00:00:00",
    ])
    ws.append([
        1002, "0002", "B", "654321", "000321", "921234567V",
        "nimal silva", "n. silva",
        "0719876543", "nimal@example.com", "Yes", "No", "2026-11-21 00:00:00",
    ])
    wb.save(emp_xlsx)
    wb.close()

    # TEMPLATE.xlsx
    tpl_xlsx = base_dir / "RR_TEMPLATE.xlsx"
    twb = openpyxl.Workbook()
    data_ws = twb.active
    data_ws.title = "Data"
    data_ws.append(["Index", "PrimaryKey", "C", "D", "Identifier"])
    data_ws.append([1, 1001, None, None, "Kamal Perera"])
    data_ws.append([2, 1002, None, None, "Nimal Silva"])

    ov_ws = twb.create_sheet("Overlay")
    ov_ws.append(["Index", "Name", "Content", "Param", "PreProcess"])
    rows = [
        (1, "SerialNo", "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=B>", "<Field=Serial No>", None),
        (2, "Zone", "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=C>", "<Field=Zone>", "<Function=changeTextCase(text, UPPER)>"),
        (3, "EmpNo", "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=D>", "<Field=Emp No>", None),
        (4, "MemNo", "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=E>", "<Field=Mem No>", None),
        (5, "NIC", "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=F>", "<Field=NIC>", "<Function=changeTextCase(text, UPPER)>"),
        (6, "FullName", "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=G>", "<Field=FullName>", "<Function=changeTextCase(text, UPPER)>"),
        (7, "NameINI", "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=H>", "<Field=NameWithInitials>", "<Function=changeTextCase(text, UPPER)>"),
        (8, "Mobile", "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=I>", "<Field=Mobile>", None),
        (9, "Email", "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=J>", "<Field=Email>", "<Function=changeTextCase(text, UPPER)>"),
        (10, "Internet", "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=K>", "<Field=Internet>", None),
        (11, "SMS", "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=L>", "<Field=SMS>", None),
        (12, "IDFront", "<Type=Image><Path={key}_front>", "<Slot=ID_FRONT>", None),
        (13, "IDBack", "<Type=Image><Path={key}_back>", "<Slot=ID_BACK>", None),
        (14, "CertDate", "<Type=File><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=M>", "<X=65><Y=68><Font=Helvetica><FontSize=10>", "<Function=FormatDate(text, %Y/%m/%d)>"),
    ]
    for r in rows:
        ov_ws.append(list(r))
    twb.save(tpl_xlsx)
    twb.close()

    return tpl_xlsx, emp_xlsx, img_dir


def test_sanitize_filename_part() -> None:
    assert sanitize_filename_part("Kamal/Perera:01") == "Kamal_Perera_01"
    assert sanitize_filename_part("   ") == "record"


def test_end_to_end_rr_form_batch_editable_and_flattened(tmp_path: Path) -> None:
    tpl_xlsx, _, img_dir = _build_rr_form_fixture(tmp_path)
    out_dir = tmp_path / "out_editable"
    session_out = tmp_path / "saved_session.json"

    summary = run_batch(
        RR_FORM_PATH,
        tpl_xlsx,
        out_dir,
        source_dir=tmp_path,
        image_dir=img_dir,
        flatten=False,
        save_session_path=session_out,
    )
    assert isinstance(summary, BatchSummary)
    assert summary.total == 2
    assert summary.succeeded == 2
    assert summary.failed == 0
    assert session_out.exists()

    loaded_session = load_session(session_out)
    assert loaded_session.source_files[0].name == "employees.xlsx"

    pdf1 = out_dir / "1001-Kamal Perera.pdf"
    pdf2 = out_dir / "1002-Nimal Silva.pdf"
    assert pdf1.exists()
    assert pdf2.exists()

    # Verify comb fields, checkboxes, date overlay, and embedded ID images in PDF 1
    orig_doc = pymupdf.open(RR_FORM_PATH)
    orig_img_count = len(orig_doc[0].get_images())
    orig_doc.close()

    doc1 = pymupdf.open(pdf1)
    try:
        widgets = {w.field_name: w.field_value for w in doc1[0].widgets()}
        assert widgets["Zone"] == "A"
        assert widgets["Emp No"] == "123456"
        assert widgets["NIC"] == "199012345678"
        assert widgets["Name 1"] == "DON KAMAL BANDARA"
        assert widgets["Name 2"] == "WICKRAMASINGHE MUDIYANSELAGE"
        assert widgets["Name 3"] == "PERERA"
        assert widgets["Internet"] == "Yes"
        assert widgets["SMS"] == "Yes"
        assert "2026/11/21" in doc1[0].get_text()
        assert len(doc1[0].get_images()) >= orig_img_count + 2
    finally:
        doc1.close()

    # Verify flattened run removes widgets
    flat_dir = tmp_path / "out_flat"
    flat_summary = run_batch(
        RR_FORM_PATH,
        tpl_xlsx,
        flat_dir,
        source_dir=tmp_path,
        image_dir=img_dir,
        flatten=True,
    )
    assert flat_summary.succeeded == 2
    doc_flat = pymupdf.open(flat_dir / "1002-Nimal Silva.pdf")
    try:
        assert list(doc_flat[0].widgets() or []) == []
        assert "2026/11/21" in doc_flat[0].get_text()
    finally:
        doc_flat.close()


def test_batch_per_record_error_isolation_and_concat(tmp_path: Path) -> None:
    tpl_xlsx, _, img_dir = _build_rr_form_fixture(tmp_path)
    # Add a 3rd record that does not exist in employees.xlsx + a !<CONCAT> row
    wb = openpyxl.load_workbook(tpl_xlsx)
    wb["Data"].append([3, 9999, None, None, "MissingUser"])
    wb["Overlay"].append([15, "!<CONCAT><CertDate>", "<Type=Text><Text= (Verified)>", None, None])
    wb.save(tpl_xlsx)
    wb.close()

    out_dir = tmp_path / "out"
    summary = run_batch(
        RR_FORM_PATH,
        tpl_xlsx,
        out_dir,
        source_dir=tmp_path,
        image_dir=img_dir,
    )
    assert summary.total == 3
    assert summary.succeeded == 2
    assert summary.failed == 1
    assert summary.results[2].success is False
    assert "9999" in (summary.results[2].error or "")

    # Verify !<CONCAT> appended to CertDate text overlay on valid PDF
    doc = pymupdf.open(out_dir / "1001-Kamal Perera.pdf")
    try:
        assert "2026/11/21(Verified)" in doc[0].get_text()
    finally:
        doc.close()


def test_render_record_pdf_error_cases(tmp_path: Path) -> None:
    with pytest.raises(PdfRenderError, match="not found"):
        render_record_pdf(tmp_path / "missing.pdf", tmp_path / "out.pdf")

    with pytest.raises(PdfRenderError, match="out of range"):
        render_record_pdf(RR_FORM_PATH, tmp_path / "out.pdf", page_index=5)


def test_cli_headless_and_session_and_errors(tmp_path: Path, capsys) -> None:
    tpl_xlsx, _, img_dir = _build_rr_form_fixture(tmp_path)
    out_dir = tmp_path / "cli_out"
    session_file = tmp_path / "session.json"

    rc = cli_main([
        "--pdf", str(RR_FORM_PATH),
        "--template", str(tpl_xlsx),
        "--source-dir", str(tmp_path),
        "--image-dir", str(img_dir),
        "--output-dir", str(out_dir),
        "--save-session", str(session_file),
        "--flatten",
    ])
    assert rc == 0
    assert (out_dir / "1001-Kamal Perera.pdf").exists()
    assert session_file.exists()

    # Run again using saved --session
    out_dir2 = tmp_path / "cli_out2"
    rc2 = cli_main([
        "--session", str(session_file),
        "--image-dir", str(img_dir),
        "--output-dir", str(out_dir2),
    ])
    assert rc2 == 0
    assert (out_dir2 / "1002-Nimal Silva.pdf").exists()

    # Error when session file does not exist
    rc_err = cli_main(["--session", str(tmp_path / "nope.json")])
    assert rc_err == 1

    # Error when GUI picker returns empty strings
    with patch("pdf_bulk_fill.cli.pick_pdf_file", return_value=""), patch(
        "pdf_bulk_fill.cli.pick_excel_file", return_value=""
    ):
        rc_empty = cli_main(["--gui"])
        assert rc_empty == 1


def test_gui_dialogs_mocked() -> None:
    mock_tk = MagicMock()
    mock_root = MagicMock()
    mock_tk.Tk.return_value = mock_root
    mock_filedialog = MagicMock()
    mock_filedialog.askopenfilename.side_effect = ["/tmp/sample.pdf", "/tmp/data.xlsx"]
    mock_simpledialog = MagicMock()
    mock_simpledialog.askstring.return_value = "secret"

    with patch(
        "pdf_bulk_fill.gui._load_tk",
        return_value=(mock_tk, mock_filedialog, mock_simpledialog),
    ):
        assert pick_pdf_file("Title", "/tmp") == "/tmp/sample.pdf"
        assert pick_excel_file("Title", "/tmp") == "/tmp/data.xlsx"
        assert prompt_password("locked.xlsx") == "secret"
        assert mock_root.withdraw.call_count == 3
        assert mock_root.destroy.call_count == 3
