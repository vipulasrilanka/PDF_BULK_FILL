"""Comprehensive TDD regression tests for issues #26 and #27."""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pymupdf
import pytest

from pdf_bulk_fill import (
    PdfBulkFillError,
    PreprocessError,
    apply_preprocess,
    extract_number,
    format_currency,
    format_number,
    run_batch,
)
from pdf_bulk_fill.cli import main as cli_main

RR_FORM_PATH = Path(__file__).resolve().parent.parent / "templates" / "pdf" / "RR-Form.pdf"


def _make_blank_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    doc.new_page(width=595, height=842)
    doc.save(path)
    doc.close()
    return path


def _write_workbook(path: Path, sheets: dict[str, list[list]]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    first = True
    for name, rows in sheets.items():
        ws = wb.active if first else wb.create_sheet(name)
        ws.title = name
        first = False
        for row in rows:
            ws.append(row)
    wb.save(path)
    wb.close()
    return path


# =============================================================================
# Issue #26: Session replay ignores an explicit --source-dir override
# =============================================================================


class TestIssue26SessionSourceDirOverride:
    """Explicit --source-dir during --session replay must take precedence over saved source paths."""

    def test_cli_session_replay_respects_source_dir_override_and_rejects_missing(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        job_a = tmp_path / "job-a"
        job_b = tmp_path / "job-b"
        job_empty = tmp_path / "job-empty"
        job_empty.mkdir(parents=True)

        _write_workbook(
            job_a / "employees.xlsx",
            {"Employees": [["EmpID", "Serial"], [1001, "FIRST"]]},
        )
        _write_workbook(
            job_b / "employees.xlsx",
            {"Employees": [["EmpID", "Serial"], [1001, "SECOND"]]},
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [["Index", "Key", "C", "D", "ID"], [1, 1001, None, None, "Emp1"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "Serial",
                        "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=B>",
                        "<Field=Serial No>",
                        None,
                    ],
                ],
            },
        )

        out_a = tmp_path / "out-a"
        out_replay = tmp_path / "out-replay"
        out_b = tmp_path / "out-b"
        out_empty = tmp_path / "out-empty"
        session_file = tmp_path / "run.json"

        # 1. Initial run with job-a, saving session
        rc1 = cli_main([
            "--pdf", str(RR_FORM_PATH),
            "--template", str(tpl_xlsx),
            "--source-dir", str(job_a),
            "--output-dir", str(out_a),
            "--save-session", str(session_file),
        ])
        assert rc1 == 0
        doc_a = pymupdf.open(out_a / "1001-Emp1.pdf")
        try:
            widgets = {w.field_name: w.field_value for w in doc_a[0].widgets()}
            assert widgets["Serial No"] == "FIRST"
        finally:
            doc_a.close()

        # 2. Session-only replay -> still uses job-a (FIRST)
        rc_replay = cli_main([
            "--session", str(session_file),
            "--output-dir", str(out_replay),
        ])
        assert rc_replay == 0
        doc_replay = pymupdf.open(out_replay / "1001-Emp1.pdf")
        try:
            widgets = {w.field_name: w.field_value for w in doc_replay[0].widgets()}
            assert widgets["Serial No"] == "FIRST"
        finally:
            doc_replay.close()

        # 3. Replay with explicit --source-dir job-b -> MUST use job-b (SECOND)
        rc_b = cli_main([
            "--session", str(session_file),
            "--source-dir", str(job_b),
            "--output-dir", str(out_b),
        ])
        assert rc_b == 0
        doc_b = pymupdf.open(out_b / "1001-Emp1.pdf")
        try:
            widgets = {w.field_name: w.field_value for w in doc_b[0].widgets()}
            assert widgets["Serial No"] == "SECOND"
        finally:
            doc_b.close()

        # 4. Replay with explicit --source-dir job-empty -> MUST fail naming workbook & dir
        capsys.readouterr()
        rc_empty = cli_main([
            "--session", str(session_file),
            "--source-dir", str(job_empty),
            "--output-dir", str(out_empty),
        ])
        captured = capsys.readouterr()
        assert rc_empty == 1
        assert "employees.xlsx" in captured.err
        assert str(job_empty) in captured.err
        assert not (out_empty / "1001-Emp1.pdf").exists()

    def test_session_only_replay_still_supports_renamed_workbook(
        self, tmp_path: Path
    ) -> None:
        actual_source = tmp_path / "picked" / "custom_staff_file.xlsx"
        _write_workbook(
            actual_source,
            {"Employees": [["EmpID", "Serial"], [1001, "RENAMED"]]},
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [["Index", "Key", "C", "D", "ID"], [1, 1001, None, None, "Emp1"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "Serial",
                        "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=B>",
                        "<Field=Serial No>",
                        None,
                    ],
                ],
            },
        )
        out_dir = tmp_path / "out"
        session_file = tmp_path / "renamed_session.json"

        summary = run_batch(
            RR_FORM_PATH,
            tpl_xlsx,
            out_dir,
            file_locator=lambda fname: actual_source if fname == "employees.xlsx" else None,
            save_session_path=session_file,
        )
        assert summary.succeeded == 1
        (out_dir / "1001-Emp1.pdf").unlink()

        # Session-only replay without --source-dir must reopen custom_staff_file.xlsx
        rc = cli_main(["--session", str(session_file)])
        assert rc == 0
        doc = pymupdf.open(out_dir / "1001-Emp1.pdf")
        try:
            widgets = {w.field_name: w.field_value for w in doc[0].widgets()}
            assert widgets["Serial No"] == "RENAMED"
        finally:
            doc.close()


# =============================================================================
# Issue #27: Incomplete exponent such as 1e+ accepted as valid currency amount
# =============================================================================


class TestIssue27IncompleteExponentRejection:
    """Incomplete exponents (1e, 1e+, 1e-) must be rejected while valid sci notation works."""

    @pytest.mark.parametrize(
        "bad_numeric",
        [
            "1e",
            "1e+",
            "1e-",
            "1E",
            "1E+",
            "1E-",
            "-1e+",
            "+1e-",
            "12.5e+",
            ".5e-",
            "LKR 1e+",
            "Rs. 1e-",
            "1e+abc",
            "1e+2e3",
            "100.abc",
            "++100",
            "..50",
        ],
    )
    def test_incomplete_exponent_rejected_by_formatters(self, bad_numeric: str) -> None:
        with pytest.raises(PreprocessError):
            extract_number(bad_numeric)
        with pytest.raises(PreprocessError):
            format_currency(bad_numeric, decimal_points=2, currency="LKR")
        with pytest.raises(PreprocessError):
            format_number(bad_numeric, decimal_points=2)

    @pytest.mark.parametrize(
        "valid_val,expected_currency",
        [
            ("1e+2", "LKR 100.00"),
            ("1E+2", "LKR 100.00"),
            ("1e2", "LKR 100.00"),
            ("-2.5e+3", "LKR -2,500.00"),
            ("1.5e-1", "LKR 0.15"),
            ("Rs. 1e+2", "LKR 100.00"),
            ("1,234.50", "LKR 1,234.50"),
            ("Rs. 1,234.50", "LKR 1,234.50"),
            ("100/=", "LKR 100.00"),
        ],
    )
    def test_valid_scientific_notation_and_currency_still_work(
        self, valid_val: str, expected_currency: str
    ) -> None:
        assert format_currency(valid_val, decimal_points=2, currency="LKR") == expected_currency
        assert (
            apply_preprocess(
                valid_val,
                {"Function": {"name": "NumberToCurrency", "param2": 2, "param3": "LKR"}},
            )
            == expected_currency
        )

    def test_cli_middle_record_incomplete_exponent_fails_with_record_and_mapping_details(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        tpl_pdf = _make_blank_pdf(tmp_path / "blank.pdf")
        _write_workbook(
            tmp_path / "pay.xlsx",
            {
                "Salary": [
                    ["EmpID", "Amount"],
                    [1001, "1e+2"],
                    [1002, "1e+"],
                    [1003, "2500.50"],
                ]
            },
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [
                    ["Index", "Key", "C", "D", "Name"],
                    [1, 1001, None, None, "First"],
                    [2, 1002, None, None, "BadExp"],
                    [3, 1003, None, None, "Third"],
                ],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "SalaryAmount",
                        "<Type=File><File=pay.xlsx><Sheet=Salary><PrimaryKey=A><Value=B>",
                        "<X=50><Y=700><Font=Helvetica><FontSize=12>",
                        "<Function=NumberToCurrency(text, 2, LKR)>",
                    ],
                ],
            },
        )

        out_dir = tmp_path / "out"
        rc = cli_main([
            "--pdf", str(tpl_pdf),
            "--template", str(tpl_xlsx),
            "--source-dir", str(tmp_path),
            "--output-dir", str(out_dir),
        ])
        captured = capsys.readouterr()

        assert rc == 1
        assert "Completed: 2/3 succeeded, 1 failed." in captured.out
        assert "1002" in captured.err
        assert "SalaryAmount" in captured.err
        assert (out_dir / "1001-First.pdf").exists()
        assert not (out_dir / "1002-BadExp.pdf").exists()
        assert (out_dir / "1003-Third.pdf").exists()

        doc1 = pymupdf.open(out_dir / "1001-First.pdf")
        doc3 = pymupdf.open(out_dir / "1003-Third.pdf")
        try:
            assert "LKR 100.00" in doc1[0].get_text()
            assert "LKR 2,500.50" in doc3[0].get_text()
        finally:
            doc1.close()
            doc3.close()
