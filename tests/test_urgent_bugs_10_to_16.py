"""Comprehensive regression tests for urgent issues #10 through #16.

Written test-first (TDD) to verify failure on the unfixed codebase and confirm
all edge cases and data combinations pass once fixed.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pymupdf
import pytest
from PIL import Image, ImageDraw

from pdf_bulk_fill import (
    ImageOverlayError,
    PdfBulkFillError,
    PdfRenderError,
    apply_preprocess,
    extract_number,
    fill_acroform,
    format_currency,
    format_number,
    resolve_image_file,
    run_batch,
)
from pdf_bulk_fill.cli import main as cli_main

RR_FORM_PATH = Path(__file__).resolve().parent.parent / "templates" / "pdf" / "RR-Form.pdf"


def _make_png(path: Path, color: tuple[int, int, int], label: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGB", (320, 200), color=color)
    draw = ImageDraw.Draw(img)
    draw.rectangle([6, 6, 314, 194], outline=(10, 10, 10), width=4)
    draw.text((20, 90), label, fill=(10, 10, 10))
    img.save(path)
    return path


def _make_blank_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    doc.new_page(width=595, height=842)
    doc.save(path)
    doc.close()
    return path


def _write_workbook(path: Path, sheets: dict[str, list[list]]) -> Path:
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
# Issue #10: Distinct record keys can overwrite the same output PDF
# =============================================================================


class TestIssue10OutputFilenameCollisions:
    """Keys that sanitize to the same output PDF must never silently overwrite."""

    @pytest.mark.parametrize(
        "key_a,id_a,key_b,id_b",
        [
            ("A/B", "Test", "A:B", "Test"),
            ("A*B", "User", "A?B", "User"),
            ("A\\B", "User", "A|B", "User"),
            ("A<B>", "User", "A\"B\"", "User"),
            ("ab", "Test", "AB", "Test"),  # Case-insensitive filesystem collision
            ("1001", "A/B", "1001", "A:B"),  # Identifier collision
            ("Caf\u00e9", "Test", "Cafe\u0301", "Test"),  # NFC vs NFD (macOS APFS)
            ("CAF\u00c9", "Test", "cafe\u0301", "Test"),  # Mixed case + NFC vs NFD
            ("1001", "Caf\u00e9", "1001", "Cafe\u0301"),  # Identifier NFC vs NFD
            ("o\u0302\u0323", "Test", "\u1ed9", "Test"),  # Multi-combining canonical equiv
            ("Stra\u00dfe", "Test", "STRASSE", "Test"),  # Unicode casefold (ß == SS on APFS)
            ("\u212b", "Test", "\u00c5", "Test"),  # Angstrom sign vs Latin A-ring
            ("\u2126", "Test", "\u03a9", "Test"),  # Ohm sign vs Greek Omega
            ("Rec", "Test.", "Rec", "Test"),  # Trailing dot collision
        ],
    )
    def test_colliding_keys_do_not_overwrite_or_report_false_success(
        self,
        tmp_path: Path,
        key_a: str,
        id_a: str,
        key_b: str,
        id_b: str,
    ) -> None:
        tpl_pdf = _make_blank_pdf(tmp_path / "blank.pdf")
        _write_workbook(
            tmp_path / "src.xlsx",
            {
                "Data": [
                    ["Key", "Val"],
                    [key_a, "FIRST_RECORD_CONTENT"],
                    [key_b, "SECOND_RECORD_CONTENT"],
                    ["SAFE_KEY", "SAFE_RECORD_CONTENT"],
                ]
            },
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [
                    ["Index", "Key", "C", "D", "Identifier"],
                    [1, key_a, None, None, id_a],
                    [2, key_b, None, None, id_b],
                    [3, "SAFE_KEY", None, None, "Unique"],
                ],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "Body",
                        "<Type=File><File=src.xlsx><Sheet=Data><PrimaryKey=A><Value=B>",
                        "<X=50><Y=700><Font=Helvetica><FontSize=12>",
                        None,
                    ],
                ],
            },
        )

        out_dir = tmp_path / "out"
        summary = run_batch(tpl_pdf, tpl_xlsx, out_dir, source_dir=tmp_path)

        # Invariant: every succeeded record must map to a distinct surviving PDF on disk
        succeeded_paths = [
            r.output_path for r in summary.results if r.success and r.output_path is not None
        ]
        assert len(succeeded_paths) == summary.succeeded
        assert len({p.name.lower() for p in succeeded_paths}) == len(succeeded_paths)
        assert all(p.exists() for p in succeeded_paths)

        # Colliding records must fail cleanly with a collision message; SAFE_KEY must succeed
        assert summary.succeeded == 1
        assert summary.failed == 2
        assert summary.results[0].success is False
        assert summary.results[1].success is False
        assert "collision" in (summary.results[0].error or "").lower()
        assert "collision" in (summary.results[1].error or "").lower()
        assert summary.results[2].success is True
        assert (out_dir / "SAFE_KEY-Unique.pdf").exists()
        # Only the unrelated SAFE_KEY PDF should exist in out_dir (no partial/overwritten PDF)
        assert sorted(p.name for p in out_dir.iterdir()) == ["SAFE_KEY-Unique.pdf"]

    def test_samefile_filesystem_guard_blocks_unexpected_collision(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Layer 2 defense: if OS resolves two distinct names to samefile, refuse overwrite."""
        tpl_pdf = _make_blank_pdf(tmp_path / "blank.pdf")
        _write_workbook(
            tmp_path / "src.xlsx",
            {"Data": [["Key", "Val"], ["K1", "FIRST"], ["K2", "SECOND"]]},
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [
                    ["Index", "Key", "C", "D", "Identifier"],
                    [1, "K1", None, None, "One"],
                    [2, "K2", None, None, "Two"],
                ],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "Body",
                        "<Type=File><File=src.xlsx><Sheet=Data><PrimaryKey=A><Value=B>",
                        "<X=50><Y=700><Font=Helvetica><FontSize=12>",
                        None,
                    ],
                ],
            },
        )
        out_dir = tmp_path / "out"
        # Pre-create a hardlink/alias simulation by monkeypatching exists+samefile for K2
        orig_exists = Path.exists

        def fake_exists(self: Path) -> bool:
            if self.name == "K2-Two.pdf":
                return True
            return orig_exists(self)

        monkeypatch.setattr(Path, "exists", fake_exists)
        monkeypatch.setattr("os.path.samefile", lambda a, b: True)

        summary = run_batch(tpl_pdf, tpl_xlsx, out_dir, source_dir=tmp_path)
        assert summary.succeeded == 1
        assert summary.failed == 1
        assert summary.results[0].success is True
        assert summary.results[1].success is False
        assert "collision" in (summary.results[1].error or "").lower()


# =============================================================================
# Issue #11: Missing ID image silently replaced with image from another directory
# =============================================================================


class TestIssue11ImageDirectoryIsolation:
    """Configured --image-dir / base_dir must never fall back to CWD."""

    def test_missing_image_in_base_dir_does_not_fall_back_to_cwd(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        cwd_dir = tmp_path / "cwd_workspace"
        configured_img_dir = tmp_path / "configured_images"
        configured_img_dir.mkdir(parents=True)

        # Place a decoy image ONLY in CWD/ids/1001_front.png
        _make_png(cwd_dir / "ids" / "1001_front.png", (255, 0, 0), "WRONG CWD IMAGE")
        monkeypatch.chdir(cwd_dir)

        with pytest.raises(ImageOverlayError, match="Image file not found"):
            resolve_image_file(
                "1001_front.png",
                location="ids",
                base_dir=configured_img_dir,
            )

    def test_batch_fails_instead_of_embedding_cwd_decoy_image(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        cwd_dir = tmp_path / "cwd_workspace"
        configured_img_dir = tmp_path / "configured_images"
        configured_img_dir.mkdir(parents=True)

        # Decoy image exists only in CWD/ids, NOT in configured_img_dir/ids
        _make_png(cwd_dir / "ids" / "1001_front.png", (255, 0, 0), "DECOY")
        monkeypatch.chdir(cwd_dir)

        _write_workbook(
            tmp_path / "employees.xlsx",
            {
                "Employees": [
                    ["EmpID", "FrontImage"],
                    [1001, "<DocType=NEW><FileName=1001_front.png><Location=ids>"],
                ]
            },
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [["Index", "Key", "C", "D", "ID"], [1, 1001, None, None, "Emp1"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "IDFront",
                        "<Type=Image><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=B>",
                        "<X=120><Y=95><Width=355><Height=180><Origin=top-left><DocType=NEW>",
                        None,
                    ],
                ],
            },
        )

        out_dir = tmp_path / "out"
        summary = run_batch(
            RR_FORM_PATH,
            tpl_xlsx,
            out_dir,
            source_dir=tmp_path,
            image_dir=configured_img_dir,
        )
        assert summary.succeeded == 0
        assert summary.failed == 1
        assert "not found" in (summary.results[0].error or "").lower()
        assert not (out_dir / "1001-Emp1.pdf").exists()

    def test_skip_missing_images_still_does_not_pick_cwd_decoy(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        cwd_dir = tmp_path / "cwd_workspace"
        configured_img_dir = tmp_path / "configured_images"
        configured_img_dir.mkdir(parents=True)
        _make_png(cwd_dir / "ids" / "1001_front.png", (255, 0, 0), "DECOY")
        monkeypatch.chdir(cwd_dir)

        _write_workbook(
            tmp_path / "employees.xlsx",
            {
                "Employees": [
                    ["EmpID", "FrontImage"],
                    [1001, "<DocType=NEW><FileName=1001_front.png><Location=ids>"],
                ]
            },
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [["Index", "Key", "C", "D", "ID"], [1, 1001, None, None, "Emp1"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "IDFront",
                        "<Type=Image><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=B>",
                        "<X=120><Y=95><Width=355><Height=180><Origin=top-left><DocType=NEW>",
                        None,
                    ],
                ],
            },
        )

        out_dir = tmp_path / "out"
        summary = run_batch(
            RR_FORM_PATH,
            tpl_xlsx,
            out_dir,
            source_dir=tmp_path,
            image_dir=configured_img_dir,
            skip_missing_images=True,
        )
        assert summary.succeeded == 1
        assert len(summary.results[0].warnings) == 1

        orig_doc = pymupdf.open(RR_FORM_PATH)
        out_doc = pymupdf.open(out_dir / "1001-Emp1.pdf")
        try:
            # Must NOT have embedded the CWD decoy image
            assert len(out_doc[0].get_images()) == len(orig_doc[0].get_images())
        finally:
            orig_doc.close()
            out_doc.close()


# =============================================================================
# Issue #12: FormatNumber changes valid decimals during PDF generation
# =============================================================================


class TestIssue12DecimalAndScientificNumberFormatting:
    """Small decimals, scientific notation, and leading-dot strings must format accurately."""

    @pytest.mark.parametrize(
        "raw_value,decimals,expected",
        [
            (0.00001, 5, "0.00001"),
            ("1e-05", 5, "0.00001"),
            ("1E-5", 5, "0.00001"),
            ("-1e-05", 5, "-0.00001"),
            ("2.5e3", 2, "2,500.00"),
            (".50", 5, "0.50000"),
            (".5", 2, "0.50"),
            ("-.50", 2, "-0.50"),
            ("+.25", 2, "0.25"),
            ("USD .50", 2, "0.50"),
            ("1,234.50", 2, "1,234.50"),
            ("0.000005", 6, "0.000005"),
        ],
    )
    def test_format_number_preserves_valid_decimals(
        self, raw_value, decimals: int, expected: str
    ) -> None:
        assert format_number(raw_value, decimal_points=decimals) == expected

    def test_format_number_end_to_end_from_excel_cells(self, tmp_path: Path) -> None:
        tpl_pdf = _make_blank_pdf(tmp_path / "blank.pdf")
        _write_workbook(
            tmp_path / "nums.xlsx",
            {
                "Rates": [
                    ["ID", "Rate"],
                    [1, 0.00001],  # openpyxl numeric -> str() gives '1e-05'
                    [2, ".50"],    # text leading-decimal
                ]
            },
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [
                    ["Index", "Key", "C", "D", "Name"],
                    [1, 1, None, None, "TinyFloat"],
                    [2, 2, None, None, "LeadingDot"],
                ],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "Rate",
                        "<Type=File><File=nums.xlsx><Sheet=Rates><PrimaryKey=A><Value=B>",
                        "<X=50><Y=700><Font=Helvetica><FontSize=12>",
                        "<Function=FormatNumber(text, 5, , )>",
                    ],
                ],
            },
        )
        out_dir = tmp_path / "out"
        summary = run_batch(tpl_pdf, tpl_xlsx, out_dir, source_dir=tmp_path)
        assert summary.succeeded == 2

        doc1 = pymupdf.open(out_dir / "1-TinyFloat.pdf")
        doc2 = pymupdf.open(out_dir / "2-LeadingDot.pdf")
        try:
            assert "0.00001" in doc1[0].get_text()
            assert "1.00000" not in doc1[0].get_text()
            assert "0.50000" in doc2[0].get_text()
            assert "50.00000" not in doc2[0].get_text()
        finally:
            doc1.close()
            doc2.close()


# =============================================================================
# Issue #13: Session replay loses flattening, image paths, and source selections
# =============================================================================


class TestIssue13SessionReplayFidelity:
    """Replaying with only --session must preserve flatten, image_dir, and renamed source files."""

    def test_session_replay_preserves_flatten_image_dir_and_renamed_source(
        self, tmp_path: Path
    ) -> None:
        img_root = tmp_path / "custom_images_root"
        _make_png(img_root / "ids" / "1001_front.png", (100, 180, 255), "FRONT")

        # Source workbook lives under a DIFFERENT filename than template references ('employees.xlsx')
        actual_source = tmp_path / "alt_sources" / "employees_nov_2026.xlsx"
        actual_source.parent.mkdir(parents=True)
        _write_workbook(
            actual_source,
            {
                "Employees": [
                    ["EmpID", "Zone", "FrontImage"],
                    [1001, "A", "<DocType=NEW><FileName=1001_front.png><Location=ids>"],
                ]
            },
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [["Index", "Key", "C", "D", "ID"], [1, 1001, None, None, "Kamal"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "Zone",
                        "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=B>",
                        "<Field=Zone>",
                        None,
                    ],
                    [
                        2,
                        "IDFront",
                        "<Type=Image><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=C>",
                        "<X=120><Y=95><Width=355><Height=180><Origin=top-left><DocType=NEW>",
                        None,
                    ],
                ],
            },
        )

        out_dir = tmp_path / "out"
        session_file = tmp_path / "session.json"

        # Initial run: uses file_locator to pick renamed file + flatten=True + custom image_dir
        summary1 = run_batch(
            RR_FORM_PATH,
            tpl_xlsx,
            out_dir,
            image_dir=img_root,
            flatten=True,
            file_locator=lambda fname: actual_source if fname == "employees.xlsx" else None,
            save_session_path=session_file,
        )
        assert summary1.succeeded == 1
        assert session_file.exists()

        # Delete output PDF to ensure replay truly regenerates it
        out_pdf = out_dir / "1001-Kamal.pdf"
        out_pdf.unlink()

        # Replay via CLI passing ONLY --session
        rc = cli_main(["--session", str(session_file)])
        assert rc == 0
        assert out_pdf.exists()

        orig_doc = pymupdf.open(RR_FORM_PATH)
        replayed_doc = pymupdf.open(out_pdf)
        try:
            # Must still be flattened (0 interactive widgets)
            assert list(replayed_doc[0].widgets() or []) == []
            # Must still have embedded the ID image from saved image_dir
            assert len(replayed_doc[0].get_images()) >= len(orig_doc[0].get_images()) + 1
        finally:
            orig_doc.close()
            replayed_doc.close()


# =============================================================================
# Issue #14: One malformed currency value aborts the batch before later records
# =============================================================================


class TestIssue14MalformedNumericIsolation:
    """Malformed numeric/currency values must fail only the affected record, never abort the batch."""

    @pytest.mark.parametrize(
        "bad_val",
        [
            "pending, review",
            ",",
            ",,,",
            "-",
            ".",
            "N/A",
            "TBD",
            "",
        ],
    )
    def test_malformed_currency_raises_typed_error_not_raw_valueerror(
        self, bad_val: str
    ) -> None:
        with pytest.raises(PdfBulkFillError):
            format_currency(bad_val, decimal_points=2, currency="Rs.")
        with pytest.raises(PdfBulkFillError):
            format_number(bad_val, decimal_points=2)
        with pytest.raises(PdfBulkFillError):
            extract_number(bad_val)

    def test_middle_record_malformed_currency_does_not_abort_cli_batch(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        tpl_pdf = _make_blank_pdf(tmp_path / "blank.pdf")
        _write_workbook(
            tmp_path / "pay.xlsx",
            {
                "Salary": [
                    ["EmpID", "Amount"],
                    [1001, "15000.50"],
                    [1002, "pending, review"],
                    [1003, "25000.00"],
                ]
            },
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [
                    ["Index", "Key", "C", "D", "Name"],
                    [1, 1001, None, None, "First"],
                    [2, 1002, None, None, "BadMiddle"],
                    [3, 1003, None, None, "ThirdValid"],
                ],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "Salary",
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
        # Non-zero exit because 1 record failed, BUT summary is printed & record 3 succeeded!
        assert rc == 1
        assert "Completed: 2/3 succeeded, 1 failed." in captured.out
        assert (out_dir / "1001-First.pdf").exists()
        assert not (out_dir / "1002-BadMiddle.pdf").exists()
        assert (out_dir / "1003-ThirdValid.pdf").exists()


# =============================================================================
# Issue #15: Unknown AcroForm field name is silently ignored
# =============================================================================


class TestIssue15UnknownAcroFormFields:
    """Typo or unknown PDF form field names must fail the record and name the typo."""

    @pytest.mark.parametrize(
        "bad_field_name",
        [
            "Mobil",       # Typo for Mobile
            "mobile",      # Wrong casing
            "Name INI2",   # Missing space vs 'Name INI 2'
            "UnknownBox",  # Completely nonexistent
        ],
    )
    def test_unknown_acroform_field_fails_batch_record(
        self, tmp_path: Path, bad_field_name: str
    ) -> None:
        _write_workbook(
            tmp_path / "emp.xlsx",
            {"Employees": [["EmpID", "Mobile"], [1001, "0771234567"]]},
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [["Index", "Key", "C", "D", "ID"], [1, 1001, None, None, "Kamal"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "MobileField",
                        "<Type=FormField><File=emp.xlsx><Sheet=Employees><PrimaryKey=A><Value=B>",
                        f"<Field={bad_field_name}>",
                        None,
                    ],
                ],
            },
        )

        out_dir = tmp_path / "out"
        summary = run_batch(RR_FORM_PATH, tpl_xlsx, out_dir, source_dir=tmp_path)
        assert summary.succeeded == 0
        assert summary.failed == 1
        assert bad_field_name in (summary.results[0].error or "")
        assert not (out_dir / "1001-Kamal.pdf").exists()


# =============================================================================
# Issue #16: Overlength AcroForm values are truncated without warning
# =============================================================================


class TestIssue16OverlengthAcroFormValues:
    """Overlength AcroForm fields must not be silently truncated without error or warning."""

    @pytest.mark.parametrize(
        "field_name,value,max_len",
        [
            ("Email", "verylongemployeename123456@example.com", 27),  # Cuts off right at '@'
            ("Zone", "AB", 1),
            ("Emp No", "1234567", 6),
            ("NIC", "1990123456789", 12),
            ("Mobile", "07712345678", 10),
        ],
    )
    def test_overlength_field_fails_by_default_in_batch(
        self, tmp_path: Path, field_name: str, value: str, max_len: int
    ) -> None:
        _write_workbook(
            tmp_path / "emp.xlsx",
            {"Employees": [["EmpID", "Val"], [1001, value]]},
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [["Index", "Key", "C", "D", "ID"], [1, 1001, None, None, "Kamal"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "Target",
                        "<Type=FormField><File=emp.xlsx><Sheet=Employees><PrimaryKey=A><Value=B>",
                        f"<Field={field_name}>",
                        None,
                    ],
                ],
            },
        )

        out_dir = tmp_path / "out"
        summary = run_batch(RR_FORM_PATH, tpl_xlsx, out_dir, source_dir=tmp_path)
        assert summary.succeeded == 0
        assert summary.failed == 1
        err = summary.results[0].error or ""
        assert field_name in err
        assert str(max_len) in err
        assert not (out_dir / "1001-Kamal.pdf").exists()

    def test_overlength_comb_group_fails_by_default_in_batch(self, tmp_path: Path) -> None:
        # FullName max capacity across Name 1 (22) + Name 2 (31) + Name 3 (31) = 84 chars
        huge_name = " ".join(["WICKRAMASINGHE"] * 8)
        _write_workbook(
            tmp_path / "emp.xlsx",
            {"Employees": [["EmpID", "FullName"], [1001, huge_name]]},
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [["Index", "Key", "C", "D", "ID"], [1, 1001, None, None, "Kamal"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "FullName",
                        "<Type=FormField><File=emp.xlsx><Sheet=Employees><PrimaryKey=A><Value=B>",
                        "<Field=FullName>",
                        None,
                    ],
                ],
            },
        )

        out_dir = tmp_path / "out"
        summary = run_batch(RR_FORM_PATH, tpl_xlsx, out_dir, source_dir=tmp_path)
        assert summary.succeeded == 0
        assert summary.failed == 1
        assert "exceeds" in (summary.results[0].error or "").lower()

    def test_allow_truncation_emits_explicit_warning_in_batch_and_cli(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        email = "verylongemployeename123456@example.com"
        _write_workbook(
            tmp_path / "emp.xlsx",
            {"Employees": [["EmpID", "Email"], [1001, email]]},
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [["Index", "Key", "C", "D", "ID"], [1, 1001, None, None, "Kamal"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [
                        1,
                        "Email",
                        "<Type=FormField><File=emp.xlsx><Sheet=Employees><PrimaryKey=A><Value=B>",
                        "<Field=Email>",
                        None,
                    ],
                ],
            },
        )

        out_dir = tmp_path / "out"
        rc = cli_main([
            "--pdf", str(RR_FORM_PATH),
            "--template", str(tpl_xlsx),
            "--source-dir", str(tmp_path),
            "--output-dir", str(out_dir),
            "--allow-truncation",
        ])
        captured = capsys.readouterr()
        assert rc == 0
        assert (out_dir / "1001-Kamal.pdf").exists()
        assert "Email" in captured.err and "27" in captured.err
        assert "truncat" in captured.err.lower()

    def test_fill_acroform_non_strict_records_warning(self, tmp_path: Path) -> None:
        warnings: list[str] = []
        fill_acroform(
            RR_FORM_PATH,
            tmp_path / "warn.pdf",
            {
                "Email": "verylongemployeename123456@example.com",
                "NameWithInitials": " ".join(["WICKRAMASINGHE"] * 5),
            },
            strict_maxlen=False,
            warnings_out=warnings,
        )
        assert len(warnings) == 2
        assert any("Email" in w and "27" in w for w in warnings)
        assert any("NameWithInitials" in w for w in warnings)

    def test_exact_boundary_lengths_succeed_with_zero_warnings(self, tmp_path: Path) -> None:
        warnings: list[str] = []
        fill_acroform(
            RR_FORM_PATH,
            tmp_path / "exact.pdf",
            {
                "Zone": "A",                           # 1/1
                "Emp No": "123456",                    # 6/6
                "Mem No": "654321",                    # 6/6
                "NIC": "199012345678",                 # 12/12
                "Mobile": "0771234567",                # 10/10
                "Email": "123456789012345@example.com",  # 27/27
            },
            strict_maxlen=True,
            warnings_out=warnings,
        )
        assert warnings == []


class TestAdditionalEdgeCombinations:
    """Extra edge combinations across #10-#16."""

    def test_issue11_empty_location_and_path_traversal_blocked(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        cwd_dir = tmp_path / "cwd"
        img_dir = tmp_path / "img_root"
        img_dir.mkdir(parents=True)
        _make_png(cwd_dir / "1001_front.png", (255, 0, 0), "CWD")
        _make_png(tmp_path / "secret.png", (255, 0, 0), "OUTSIDE")
        monkeypatch.chdir(cwd_dir)

        with pytest.raises(ImageOverlayError, match="Image file not found"):
            resolve_image_file("1001_front.png", location="", base_dir=img_dir)

        with pytest.raises(ImageOverlayError, match="escapes configured image directory"):
            resolve_image_file("secret.png", location="..", base_dir=img_dir)

        with pytest.raises(ImageOverlayError, match="escapes configured image directory"):
            resolve_image_file("../secret.png", location="", base_dir=img_dir)

    def test_issue12_and_14_currency_text_and_invalid_decimals(self) -> None:
        assert format_currency("1e-05", 5, "LKR") == "LKR 0.00001"
        assert format_currency(".50", 2, "Rs.") == "Rs. 0.50"
        assert (
            apply_preprocess(
                "2.5e3",
                {"Function": {"name": "NumberToText", "param2": "Integer"}},
            )
            == "Two thousand, five hundred"
        )
        for bad in (True, None, float("nan"), float("inf"), "12.34.56", "1,2,,3", "100,"):
            with pytest.raises(PdfBulkFillError):
                extract_number(bad)
        with pytest.raises(PdfBulkFillError, match="decimal places"):
            format_number("100", decimal_points="bad")  # type: ignore[arg-type]
        with pytest.raises(PdfBulkFillError, match="non-negative"):
            format_number("100", decimal_points=-1)

    def test_issue13_cli_replay_flags_can_be_overridden(self, tmp_path: Path) -> None:
        img_root = tmp_path / "imgs"
        _make_png(img_root / "ids" / "1001_front.png", (100, 180, 255), "FRONT")
        _write_workbook(
            tmp_path / "employees.xlsx",
            {
                "Employees": [
                    ["EmpID", "Zone", "FrontImage"],
                    [1001, "A", "<DocType=NEW><FileName=1001_front.png><Location=ids>"],
                ]
            },
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [["Index", "Key", "C", "D", "ID"], [1, 1001, None, None, "Kamal"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [1, "Zone", "<Type=FormField><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=B>", "<Field=Zone>", None],
                    [2, "IDFront", "<Type=Image><File=employees.xlsx><Sheet=Employees><PrimaryKey=A><Value=C>", "<X=120><Y=95><Width=355><Height=180><Origin=top-left><DocType=NEW>", None],
                ],
            },
        )
        out_dir = tmp_path / "out"
        session_file = tmp_path / "session.json"

        assert cli_main([
            "--pdf", str(RR_FORM_PATH),
            "--template", str(tpl_xlsx),
            "--source-dir", str(tmp_path),
            "--image-dir", str(img_root),
            "--output-dir", str(out_dir),
            "--flatten",
            "--save-session", str(session_file),
        ]) == 0

        # Replay with --no-flatten override -> PDF must become editable again
        out_dir2 = tmp_path / "out_editable"
        assert cli_main([
            "--session", str(session_file),
            "--output-dir", str(out_dir2),
            "--no-flatten",
        ]) == 0
        doc = pymupdf.open(out_dir2 / "1001-Kamal.pdf")
        try:
            assert len(list(doc[0].widgets() or [])) == 14
        finally:
            doc.close()

    def test_issue15_multiple_unknown_fields_and_comb_group_on_blank_pdf(
        self, tmp_path: Path
    ) -> None:
        blank_pdf = _make_blank_pdf(tmp_path / "blank.pdf")
        _write_workbook(
            tmp_path / "emp.xlsx",
            {"Employees": [["EmpID", "Name", "Mob"], [1001, "KAMAL", "077"]]},
        )
        tpl_xlsx = _write_workbook(
            tmp_path / "TEMPLATE.xlsx",
            {
                "Data": [["Index", "Key", "C", "D", "ID"], [1, 1001, None, None, "Kamal"]],
                "Overlay": [
                    ["Index", "Name", "Content", "Param", "PreProcess"],
                    [1, "FullName", "<Type=FormField><File=emp.xlsx><Sheet=Employees><PrimaryKey=A><Value=B>", "<Field=FullName>", None],
                    [2, "Mobil", "<Type=FormField><File=emp.xlsx><Sheet=Employees><PrimaryKey=A><Value=C>", "<Field=Mobil>", None],
                ],
            },
        )
        summary = run_batch(blank_pdf, tpl_xlsx, tmp_path / "out", source_dir=tmp_path)
        assert summary.failed == 1
        err = summary.results[0].error or ""
        assert "Mobil" in err
        assert "Name 1" in err

