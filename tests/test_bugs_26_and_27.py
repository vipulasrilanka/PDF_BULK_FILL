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


class TestIssue27BabelCurrencyAndLocale:
    """Acceptance tests for CLDR/Babel currency markers, locales, mismatch & ambiguity checks."""

    @pytest.mark.parametrize(
        "raw_input,expected_currency,input_locale,expected_decimal_str",
        [
            # USD / en_US
            ("USD100", "USD", "en_US", "100"),
            ("USD 100", "USD", "en_US", "100"),
            ("100USD", "USD", "en_US", "100"),
            ("100 USD", "USD", "en_US", "100"),
            ("$100", "USD", "en_US", "100"),
            ("100 $", "USD", "en_US", "100"),
            ("USD\u00a0100", "USD", "en_US", "100"),
            ("100\u202f$", "USD", "en_US", "100"),
            ("-$100", "USD", "en_US", "-100"),
            ("$-100", "USD", "en_US", "-100"),
            ("USD1e+2", "USD", "en_US", "100"),
            ("1e+2 USD", "USD", "en_US", "100"),
            ("1,234", "USD", "en_US", "1234"),
            # LKR / en_LK
            ("LKR100", "LKR", "en_LK", "100"),
            ("LKR 100", "LKR", "en_LK", "100"),
            ("100LKR", "LKR", "en_LK", "100"),
            ("100 LKR", "LKR", "en_LK", "100"),
            ("Rs. 100", "LKR", "en_LK", "100"),
            ("Rs.100", "LKR", "en_LK", "100"),
            ("Rs. 1,234.50", "LKR", "en_LK", "1234.50"),
            ("100/=", "LKR", "en_LK", "100"),
            ("Rs. 100/=", "LKR", "en_LK", "100"),
            # EUR / de_DE
            ("1.234,56 €", "EUR", "de_DE", "1234.56"),
            ("€1.234,56", "EUR", "de_DE", "1234.56"),
            ("EUR1.234,56", "EUR", "de_DE", "1234.56"),
            ("1,234", "EUR", "de_DE", "1.234"),
            # EUR / fr_FR (narrow non-breaking space U+202F)
            ("1\u202f234,56 €", "EUR", "fr_FR", "1234.56"),
            ("1\u202f234,56\u202f€", "EUR", "fr_FR", "1234.56"),
            # CAD / fr_CA
            ("100,50 $", "CAD", "fr_CA", "100.50"),
            ("1\u00a0234,50 $", "CAD", "fr_CA", "1234.50"),
            # GBP / INR
            ("£250.75", "GBP", "en_GB", "250.75"),
            ("250.75GBP", "GBP", "en_GB", "250.75"),
            ("₹1,00,000.50", "INR", "en_IN", "100000.50"),
            ("INR500", "INR", "en_IN", "500"),
        ],
    )
    def test_parse_currency_amount_acceptance_matrix(
        self,
        raw_input: str,
        expected_currency: str,
        input_locale: str,
        expected_decimal_str: str,
    ) -> None:
        from decimal import Decimal
        from pdf_bulk_fill import parse_currency_amount

        parsed = parse_currency_amount(
            raw_input,
            expected_currency=expected_currency,
            input_locale=input_locale,
        )
        assert parsed.amount == Decimal(expected_decimal_str)
        assert parsed.currency == expected_currency

    @pytest.mark.parametrize(
        "bad_input,expected_currency,input_locale,match",
        [
            # Incomplete exponents with currency markers
            ("1e", "USD", "en_US", "scientific|exponent|malformed"),
            ("1e+", "USD", "en_US", "scientific|exponent|malformed"),
            ("1e-", "USD", "en_US", "scientific|exponent|malformed"),
            ("USD1e+", "USD", "en_US", "scientific|exponent|malformed"),
            ("1e+ USD", "USD", "en_US", "scientific|exponent|malformed"),
            # Currency mismatch
            ("USD100", "LKR", "en_LK", "mismatch"),
            ("100EUR", "USD", "en_US", "mismatch"),
            ("£100", "USD", "en_US", "mismatch"),
            ("USD 100 EUR", "USD", "en_US", "currency|mismatch|multiple"),
            ("USD100", "Rs.", "en_LK", "mismatch"),
            ("USD100", "Rs.", None, "ambiguous"),
            # Ambiguous currency symbol with no currency or locale configured
            ("$100", None, None, "ambiguous"),
            ("100 $", None, None, "ambiguous"),
            ("Rs. 100", None, None, "ambiguous"),
            # Arbitrary text / invalid multi-token numbers / duplicate signs / bare markers
            ("invoice100", "USD", "en_US", "valid|malformed"),
            ("100abc", "USD", "en_US", "valid|malformed"),
            ("USD 10 20", "USD", "en_US", "valid|malformed"),
            ("1,23", "USD", "en_US", "valid|grouping|malformed"),
            ("1,,234", "USD", "en_US", "valid|grouping|malformed"),
            ("1 20", "EUR", "fr_FR", "valid|grouping|malformed"),
            ("-$ -100", "USD", "en_US", "duplicate sign"),
            ("USD", "USD", "en_US", "missing numeric"),
            ("/=", "LKR", "en_LK", "invalid numeric"),
            ("100", "USD", "zz_INVALID", "locale"),
            ("100", "USD", "en_ZZ", "territory|locale"),
            ("100", "USD", "123", "locale"),
        ],
    )
    def test_parse_currency_amount_rejections(
        self,
        bad_input: str,
        expected_currency: str | None,
        input_locale: str | None,
        match: str,
    ) -> None:
        from pdf_bulk_fill import parse_currency_amount

        with pytest.raises(PreprocessError, match=f"(?i){match}"):
            parse_currency_amount(
                bad_input,
                expected_currency=expected_currency,
                input_locale=input_locale,
            )

    def test_template_settings_sheet_and_per_field_locale_override_end_to_end(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        tpl_pdf = _make_blank_pdf(tmp_path / "blank.pdf")
        _write_workbook(
            tmp_path / "amounts.xlsx",
            {
                "Pay": [
                    ["ID", "LocalAmt", "EuroAmt"],
                    [1, "LKR100", "1\u202f234,56 €"],
                    [2, "100LKR", "1.234,56 €"],
                    [3, "USD100", "1\u202f234,56 €"],  # Mismatch: USD in LKR column!
                    [4, "Rs. 1,234.50", "100,50 $"],  # Valid LKR + CAD override below
                ]
            },
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Settings": [
                    ["Key", "Value"],
                    ["Currency", "LKR"],
                    ["InputLocale", "en_LK"],
                ],
                "Data": [
                    ["Index", "Key", "C", "D", "Name"],
                    [1, 1, None, None, "Rec1"],
                    [2, 2, None, None, "Rec2"],
                    [3, 3, None, None, "Mismatch"],
                    [4, 4, None, None, "Rec4"],
                ],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "LocalPay",
                        "<Type=File><File=amounts.xlsx><Sheet=Pay><PrimaryKey=A><Value=B>",
                        "<X=50><Y=700><Font=Helvetica><FontSize=12>",
                        "<Function=NumberToCurrency(text, 2, LKR)>",
                    ],
                    [
                        2,
                        "FrenchEuro",
                        "<Type=File><File=amounts.xlsx><Sheet=Pay><PrimaryKey=A><Value=C>",
                        "<X=50><Y=650><Font=Helvetica><FontSize=12>",
                        "<Function=NumberToCurrency(text, 2, EUR, fr_FR)>",
                    ],
                ],
            },
        )

        # Verify Rec1 succeeds (LKR100 + fr_FR narrow NBSP euro),
        # Rec2 fails because '1.234,56 €' uses de_DE '.' grouping under fr_FR,
        # Rec3 fails with currency mismatch USD vs LKR on LocalPay,
        # Rec4 fails on FrenchEuro because '$' is CAD/USD not EUR.
        out_dir = tmp_path / "out"
        rc = cli_main([
            "--pdf", str(tpl_pdf),
            "--template", str(tpl_xlsx),
            "--source-dir", str(tmp_path),
            "--output-dir", str(out_dir),
        ])
        captured = capsys.readouterr()
        assert rc == 1
        assert "Completed: 1/4 succeeded, 3 failed." in captured.out
        assert "LocalPay" in captured.err and "mismatch" in captured.err.lower()

        doc1 = pymupdf.open(out_dir / "1-Rec1.pdf")
        try:
            text1 = doc1[0].get_text()
            assert "LKR 100.00" in text1
            assert "EUR 1,234.56" in text1
        finally:
            doc1.close()

    def test_ambiguous_dollar_in_batch_without_currency_or_locale_fails_record(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        tpl_pdf = _make_blank_pdf(tmp_path / "blank.pdf")
        _write_workbook(
            tmp_path / "pay.xlsx",
            {
                "Pay": [
                    ["ID", "Amt"],
                    [1, "$100"],
                    [2, "250.50"],
                ]
            },
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [
                    ["Index", "Key", "C", "D", "Name"],
                    [1, 1, None, None, "Ambiguous"],
                    [2, 2, None, None, "PlainValid"],
                ],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "AmountField",
                        "<Type=File><File=pay.xlsx><Sheet=Pay><PrimaryKey=A><Value=B>",
                        "<X=50><Y=700><Font=Helvetica><FontSize=12>",
                        "<Function=FormatNumber(text, 2, , )>",
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
        assert "Completed: 1/2 succeeded, 1 failed." in captured.out
        assert "AmountField" in captured.err
        assert "ambiguous" in captured.err.lower()
        assert not (out_dir / "1-Ambiguous.pdf").exists()
        assert (out_dir / "2-PlainValid.pdf").exists()

    def test_locale_only_disambiguation_decimal_input_and_cli_flags(
        self, tmp_path: Path
    ) -> None:
        from decimal import Decimal
        from pdf_bulk_fill import parse_currency_amount

        # Locale alone disambiguates '$' and 'Rs.'
        assert parse_currency_amount("$100", input_locale="en_US") == (
            parse_currency_amount("USD100", expected_currency="USD", input_locale="en_US")
        )
        assert parse_currency_amount("100,50 $", input_locale="fr_CA").currency == "CAD"
        assert parse_currency_amount("Rs. 100", input_locale="en_LK").currency == "LKR"

        # Unambiguous markers without expected_currency/locale
        assert parse_currency_amount("€50").currency == "EUR"
        assert parse_currency_amount("US$75").currency == "USD"

        # Legacy symbol expected_currency ('Rs.') requires a resolving locale
        assert parse_currency_amount("Rs. 100", expected_currency="Rs.", input_locale="en_LK").currency == "LKR"
        with pytest.raises(PreprocessError, match="(?i)ambiguous"):
            parse_currency_amount("Rs. 100", expected_currency="Rs.")

        # Direct Decimal input and non-finite rejection
        assert extract_number(Decimal("12.34"), Decimal) == Decimal("12.34")
        with pytest.raises(PreprocessError):
            parse_currency_amount(Decimal("NaN"))

        # CLI --currency and --input-locale flags resolve '$100' under CAD/fr_CA
        tpl_pdf = _make_blank_pdf(tmp_path / "blank.pdf")
        _write_workbook(tmp_path / "cad.xlsx", {"Pay": [["ID", "Amt"], [1, "100,50 $"]]})
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [["Index", "Key", "C", "D", "Name"], [1, 1, None, None, "Canada"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "CADPay",
                        "<Type=File><File=cad.xlsx><Sheet=Pay><PrimaryKey=A><Value=B>",
                        "<X=50><Y=700><Font=Helvetica><FontSize=12>",
                        "<Function=NumberToCurrency(text, 2, CAD)>",
                    ],
                ],
            },
        )
        out_dir = tmp_path / "out_cad"
        assert cli_main([
            "--pdf", str(tpl_pdf),
            "--template", str(tpl_xlsx),
            "--source-dir", str(tmp_path),
            "--output-dir", str(out_dir),
            "--currency", "CAD",
            "--input-locale", "fr_CA",
        ]) == 0
        doc = pymupdf.open(out_dir / "1-Canada.pdf")
        try:
            assert "CAD 100.50" in doc[0].get_text()
        finally:
            doc.close()


class TestIssue27FollowUpFiveGaps:
    """TDD regression tests for the 5 verification gaps reported on commit 9e3c051."""

    def test_gap1_numeric_vs_text_excel_cells_under_de_DE_reopened_pdf(
        self, tmp_path: Path
    ) -> None:
        tpl_pdf = _make_blank_pdf(tmp_path / "blank.pdf")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Pay"
        ws.append(["ID", "Amount"])
        ws.append([1, 1.234])       # Numeric Excel cell -> EUR 1.23
        ws.append([2, 1234.56])     # Numeric Excel cell -> EUR 1,234.56
        ws.append([3, "1.234"])     # Text Excel cell under de_DE -> EUR 1,234.00
        ws["B4"].number_format = "@"  # Explicitly text
        wb.save(tmp_path / "pay.xlsx")
        wb.close()

        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Settings": [
                    ["Key", "Value"],
                    ["Currency", "EUR"],
                    ["InputLocale", "de_DE"],
                ],
                "Data": [
                    ["Index", "Key", "C", "D", "Name"],
                    [1, 1, None, None, "NumSmall"],
                    [2, 2, None, None, "NumLarge"],
                    [3, 3, None, None, "TextGrouped"],
                ],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "EuroAmount",
                        "<Type=File><File=pay.xlsx><Sheet=Pay><PrimaryKey=A><Value=B>",
                        "<X=50><Y=700><Font=Helvetica><FontSize=12>",
                        "<Function=NumberToCurrency(text, 2, EUR)>",
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
        assert rc == 0

        for fname, expected in [
            ("1-NumSmall.pdf", "EUR 1.23"),
            ("2-NumLarge.pdf", "EUR 1,234.56"),
            ("3-TextGrouped.pdf", "EUR 1,234.00"),
        ]:
            doc = pymupdf.open(out_dir / fname)
            try:
                text = doc[0].get_text()
                assert expected in text, f"{fname}: expected {expected!r} in {text!r}"
                if fname == "1-NumSmall.pdf":
                    assert "EUR 1,234.00" not in text
            finally:
                doc.close()

    def test_gap2_large_template_literal_and_source_cell_preserve_decimal_precision(
        self, tmp_path: Path
    ) -> None:
        tpl_pdf = _make_blank_pdf(tmp_path / "blank.pdf")
        _write_workbook(
            tmp_path / "big.xlsx",
            {"Pay": [["ID", "Amt"], [1, "9007199254740993.01"]]},
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Settings": [
                    ["Key", "Value"],
                    ["Currency", "USD"],
                    ["InputLocale", "en_US"],
                ],
                "Data": [["Index", "Key", "C", "D", "Name"], [1, 1, None, None, "BigNum"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "LiteralAmt",
                        "<Type=Text><Text=9007199254740993.01>",
                        "<X=50><Y=720><Font=Helvetica><FontSize=12>",
                        "<Function=NumberToCurrency(text, 2, USD)>",
                    ],
                    [
                        2,
                        "SourceAmt",
                        "<Type=File><File=big.xlsx><Sheet=Pay><PrimaryKey=A><Value=B>",
                        "<X=50><Y=680><Font=Helvetica><FontSize=12>",
                        "<Function=NumberToCurrency(text, 2, USD)>",
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
        assert rc == 0

        doc = pymupdf.open(out_dir / "1-BigNum.pdf")
        try:
            text = doc[0].get_text()
            assert text.count("USD 9,007,199,254,740,993.01") == 2
            assert "9,007,199,254,740,994.00" not in text
        finally:
            doc.close()

    def test_gap3_unknown_or_ambiguous_configured_currency_rejected_before_writing(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        from pdf_bulk_fill import parse_currency_amount

        # Direct parser checks
        with pytest.raises(PreprocessError, match="(?i)unknown.*currency"):
            parse_currency_amount("USD100", expected_currency="USX")
        with pytest.raises(PreprocessError, match="(?i)ambiguous"):
            parse_currency_amount("$100", expected_currency="$")
        with pytest.raises(PreprocessError, match="(?i)ambiguous"):
            parse_currency_amount("100", expected_currency="$")

        # Batch with Currency=USX in Settings must fail BEFORE writing any PDF
        tpl_pdf = _make_blank_pdf(tmp_path / "blank.pdf")
        _write_workbook(tmp_path / "pay.xlsx", {"Pay": [["ID", "Amt"], [1, "USD100"]]})
        tpl_usx = _write_workbook(
            tmp_path / "TEMPLATE_USX.xlsx",
            {
                "Settings": [["Key", "Value"], ["Currency", "USX"]],
                "Data": [["Index", "Key", "C", "D", "Name"], [1, 1, None, None, "Rec1"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "PayAmt",
                        "<Type=File><File=pay.xlsx><Sheet=Pay><PrimaryKey=A><Value=B>",
                        "<X=50><Y=700><Font=Helvetica><FontSize=12>",
                        "<Function=NumberToCurrency(text, 2)>",
                    ],
                ],
            },
        )
        out_usx = tmp_path / "out_usx"
        rc_usx = cli_main([
            "--pdf", str(tpl_pdf),
            "--template", str(tpl_usx),
            "--source-dir", str(tmp_path),
            "--output-dir", str(out_usx),
        ])
        err_usx = capsys.readouterr().err
        assert rc_usx == 1
        assert "USX" in err_usx
        assert not (out_usx / "1-Rec1.pdf").exists()

        # Batch with Currency=$ and no resolving locale must reject and name the mapping
        _write_workbook(tmp_path / "pay_dollar.xlsx", {"Pay": [["ID", "Amt"], [1, "$100"]]})
        tpl_dollar = _write_workbook(
            tmp_path / "TEMPLATE_DOLLAR.xlsx",
            {
                "Data": [["Index", "Key", "C", "D", "Name"], [1, 1, None, None, "Rec1"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "DollarMapping",
                        "<Type=File><File=pay_dollar.xlsx><Sheet=Pay><PrimaryKey=A><Value=B>",
                        "<X=50><Y=700><Font=Helvetica><FontSize=12>",
                        "<Function=NumberToCurrency(text, 2, $)>",
                    ],
                ],
            },
        )
        out_dollar = tmp_path / "out_dollar"
        rc_dollar = cli_main([
            "--pdf", str(tpl_pdf),
            "--template", str(tpl_dollar),
            "--source-dir", str(tmp_path),
            "--output-dir", str(out_dollar),
        ])
        err_dollar = capsys.readouterr().err
        assert rc_dollar == 1
        assert "DollarMapping" in err_dollar
        assert "ambiguous" in err_dollar.lower()
        assert not (out_dollar / "1-Rec1.pdf").exists()

    @pytest.mark.parametrize(
        "currency,locale_id,amount_str",
        [
            ("BRL", "pt_BR", "1234.56"),
            ("BRL", "pt_BR", "-1234.56"),
            ("SEK", "sv_SE", "1234.56"),
            ("SEK", "sv_SE", "-1234.56"),
            ("PLN", "pl_PL", "1234.56"),
            ("PLN", "pl_PL", "-1234.56"),
            ("KRW", "ko_KR", "1235"),
            ("KRW", "ko_KR", "-1235"),
        ],
    )
    def test_gap4_babel_formatted_currency_roundtrip(
        self, currency: str, locale_id: str, amount_str: str
    ) -> None:
        from decimal import Decimal
        from babel.numbers import format_currency as babel_format_currency
        from pdf_bulk_fill import parse_currency_amount

        expected_amount = Decimal(amount_str)
        formatted = babel_format_currency(expected_amount, currency, locale=locale_id)
        parsed = parse_currency_amount(
            formatted,
            expected_currency=currency,
            input_locale=locale_id,
        )
        assert parsed.amount == expected_amount
        assert parsed.currency == currency

    def test_gap5_locale_specific_grouping_en_IN_vs_en_US(self) -> None:
        from decimal import Decimal
        from pdf_bulk_fill import parse_currency_amount

        # Accepted under en_IN (CLDR grouping (3, 2))
        parsed_in = parse_currency_amount(
            "1,00,000.50", expected_currency="INR", input_locale="en_IN"
        )
        assert parsed_in.amount == Decimal("100000.50")
        assert parsed_in.currency == "INR"

        # Rejected under en_US (CLDR grouping (3, 3)) and default locale
        with pytest.raises(PreprocessError, match="(?i)grouping|malformed"):
            parse_currency_amount(
                "1,00,000.50", expected_currency="USD", input_locale="en_US"
            )
        with pytest.raises(PreprocessError, match="(?i)grouping|malformed"):
            parse_currency_amount("1,00,000.50")
