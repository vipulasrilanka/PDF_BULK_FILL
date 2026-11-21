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
    load_session,
    render_record_pdf,
    run_batch,
    sanitize_filename_part,
)
from pdf_bulk_fill.cli import main as cli_main
from pdf_bulk_fill.gui import pick_excel_file, pick_pdf_file, prompt_password

RR_FORM_PATH = Path(__file__).resolve().parent.parent / "templates" / "pdf" / "RR-Form.pdf"


def _make_id_image(
    path: Path, size: tuple[int, int], color: tuple[int, int, int], label: str
) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGB", size, color=color)
    draw = ImageDraw.Draw(img)
    draw.rectangle([8, 8, size[0] - 8, size[1] - 8], outline=(20, 40, 80), width=4)
    draw.text((20, size[1] // 2 - 10), label, fill=(10, 10, 10))
    img.save(path)
    return path


def _build_rr_form_fixture(base_dir: Path) -> tuple[Path, Path, Path]:
    """Create employees.xlsx (mixed NEW + OLD NICs), unified TEMPLATE.xlsx, ID images & logo."""
    img_dir = base_dir / "ids"
    _make_id_image(img_dir / "1001_front.png", (640, 400), (215, 232, 255), "1001 NEW FRONT")
    _make_id_image(img_dir / "1001_back.png", (640, 400), (225, 245, 220), "1001 NEW BACK")
    _make_id_image(img_dir / "1002_front.jpg", (400, 620), (250, 241, 208), "1002 OLD FRONT")
    _make_id_image(img_dir / "1002_back.jpg", (400, 620), (250, 241, 208), "1002 OLD BACK")
    _make_id_image(base_dir / "logo.png", (180, 60), (30, 64, 175), "LOGO")

    # Source employee workbook with FrontImage (Col N) and BackImage (Col O) metadata
    emp_xlsx = base_dir / "employees.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Employees"
    ws.append([
        "EmpID", "SerialNo", "Zone", "EmpNo", "MemNo", "NIC",
        "FullName", "NameINI", "Mobile", "Email", "Internet", "SMS",
        "CertDate", "FrontImage", "BackImage",
    ])
    ws.append([
        1001, "0001", "A", "123456", "000789", "199012345678",
        "don kamal bandara wickramasinghe mudiyanselage perera",
        "d. k. b. w. m. perera",
        "0771234567", "kamal@example.com", "Yes", "Yes", "2026-11-21 00:00:00",
        "<DocType=NEW><FileName=1001_front.png><Location=ids><Rotation=0><KeepAspect=True>",
        "<DocType=NEW><FileName=1001_back.png><Location=ids><Rotation=0><KeepAspect=True>",
    ])
    ws.append([
        1002, "0002", "B", "654321", "000321", "921234567V",
        "nimal silva", "n. silva",
        "0719876543", "nimal@example.com", "Yes", "No", "2026-11-21 00:00:00",
        "<DocType=OLD><FileName=1002_front.jpg><Location=ids><Rotation=0><KeepAspect=True>",
        "<DocType=OLD><FileName=1002_back.jpg><Location=ids><Rotation=0><KeepAspect=True>",
    ])
    wb.save(emp_xlsx)
    wb.close()

    # Unified TEMPLATE.xlsx covering BOTH OLD (side-by-side) and NEW (top-bottom) + constant logo
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
        # 4 DocType-matched NIC slots (2 OLD side-by-side + 2 NEW top-bottom)
        (12, "IDFront_old", "<Type=Image><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=N>", "<X=45><Y=140><Width=245><Height=250><Origin=top-left><DocType=OLD>", None),
        (13, "IDBack_old", "<Type=Image><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=O>", "<X=305><Y=140><Width=245><Height=250><Origin=top-left><DocType=OLD>", None),
        (14, "IDFront_new", "<Type=Image><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=N>", "<X=120><Y=95><Width=355><Height=180><Origin=top-left><DocType=NEW>", None),
        (15, "IDBack_new", "<Type=Image><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=O>", "<X=120><Y=295><Width=355><Height=180><Origin=top-left><DocType=NEW>", None),
        # Constant static company logo (no Excel data file)
        (16, "CompanyLogo", "<Type=Image><FileName=logo.png><Rotation=0><KeepAspect=True>", "<X=35><Y=15><Width=50><Height=20><Origin=top-left>", None),
        (17, "CertDate", "<Type=File><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=M>", "<X=65><Y=68><Font=Helvetica><FontSize=10>", "<Function=FormatDate(text, %Y/%m/%d)>"),
    ]
    for r in rows:
        ov_ws.append(list(r))
    twb.save(tpl_xlsx)
    twb.close()

    return tpl_xlsx, emp_xlsx, base_dir


def test_sanitize_filename_part() -> None:
    assert sanitize_filename_part("Kamal/Perera:01") == "Kamal_Perera_01"
    assert sanitize_filename_part("   ") == "record"


def test_end_to_end_rr_form_batch_mixed_old_and_new_nics(tmp_path: Path) -> None:
    tpl_xlsx, _, img_base_dir = _build_rr_form_fixture(tmp_path)
    out_dir = tmp_path / "out_editable"
    session_out = tmp_path / "saved_session.json"

    summary = run_batch(
        RR_FORM_PATH,
        tpl_xlsx,
        out_dir,
        source_dir=tmp_path,
        image_dir=img_base_dir,
        flatten=False,
        save_session_path=session_out,
    )
    assert isinstance(summary, BatchSummary)
    assert summary.total == 2
    assert summary.succeeded == 2
    assert summary.failed == 0
    assert session_out.exists()

    # Each record skipped 2 non-matching NIC slots with warnings
    assert len(summary.results[0].warnings) == 2
    assert len(summary.results[1].warnings) == 2

    loaded_session = load_session(session_out)
    assert loaded_session.source_files[0].name == "employees.xlsx"

    pdf1 = out_dir / "1001-Kamal Perera.pdf"
    pdf2 = out_dir / "1002-Nimal Silva.pdf"
    assert pdf1.exists()
    assert pdf2.exists()

    # Verify comb fields, checkboxes, date overlay, and 3 embedded images (2 NIC + 1 logo)
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
        assert len(doc1[0].get_images()) >= orig_img_count + 3
    finally:
        doc1.close()

    # Verify flattened run removes widgets on OLD NIC record too
    flat_dir = tmp_path / "out_flat"
    flat_summary = run_batch(
        RR_FORM_PATH,
        tpl_xlsx,
        flat_dir,
        source_dir=tmp_path,
        image_dir=img_base_dir,
        flatten=True,
    )
    assert flat_summary.succeeded == 2
    doc_flat = pymupdf.open(flat_dir / "1002-Nimal Silva.pdf")
    try:
        assert list(doc_flat[0].widgets() or []) == []
        assert "2026/11/21" in doc_flat[0].get_text()
        assert len(doc_flat[0].get_images()) >= orig_img_count + 3
    finally:
        doc_flat.close()


def test_batch_per_record_error_isolation_and_concat(tmp_path: Path) -> None:
    tpl_xlsx, _, img_base_dir = _build_rr_form_fixture(tmp_path)
    wb = openpyxl.load_workbook(tpl_xlsx)
    wb["Data"].append([3, 9999, None, None, "MissingUser"])
    wb["Overlay"].append([18, "!<CONCAT><CertDate>", "<Type=Text><Text=(Verified)>", None, None])
    wb.save(tpl_xlsx)
    wb.close()

    out_dir = tmp_path / "out"
    summary = run_batch(
        RR_FORM_PATH,
        tpl_xlsx,
        out_dir,
        source_dir=tmp_path,
        image_dir=img_base_dir,
    )
    assert summary.total == 3
    assert summary.succeeded == 2
    assert summary.failed == 1
    assert summary.results[2].success is False
    assert "9999" in (summary.results[2].error or "")

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


def test_cli_headless_and_session_and_errors(tmp_path: Path) -> None:
    tpl_xlsx, _, img_base_dir = _build_rr_form_fixture(tmp_path)
    out_dir = tmp_path / "cli_out"
    session_file = tmp_path / "session.json"

    rc = cli_main([
        "--pdf", str(RR_FORM_PATH),
        "--template", str(tpl_xlsx),
        "--source-dir", str(tmp_path),
        "--image-dir", str(img_base_dir),
        "--output-dir", str(out_dir),
        "--save-session", str(session_file),
        "--flatten",
    ])
    assert rc == 0
    assert (out_dir / "1001-Kamal Perera.pdf").exists()
    assert session_file.exists()

    out_dir2 = tmp_path / "cli_out2"
    rc2 = cli_main([
        "--session", str(session_file),
        "--image-dir", str(img_base_dir),
        "--output-dir", str(out_dir2),
    ])
    assert rc2 == 0
    assert (out_dir2 / "1002-Nimal Silva.pdf").exists()

    assert cli_main(["--session", str(tmp_path / "nope.json")]) == 1

    with patch("pdf_bulk_fill.cli.pick_pdf_file", return_value=""), patch(
        "pdf_bulk_fill.cli.pick_excel_file", return_value=""
    ):
        assert cli_main(["--gui"]) == 1


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
