"""Command-line interface (with optional Tkinter GUI pickers) for PDF_BULK_FILL."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

from pdf_bulk_fill.errors import PdfBulkFillError
from pdf_bulk_fill.gui import pick_excel_file, pick_pdf_file, prompt_password
from pdf_bulk_fill.runner import run_batch
from pdf_bulk_fill.session import SessionConfig, load_session


def build_parser() -> argparse.ArgumentParser:
    """Create the CLI argument parser."""
    parser = argparse.ArgumentParser(
        prog="pdf-bulk-fill",
        description=(
            "Bulk-fill PDF forms (AcroForm comb fields, checkboxes, employee ID images, "
            "and coordinate text overlays) from Excel workbooks."
        ),
    )
    parser.add_argument(
        "--session",
        type=str,
        default=None,
        help="Path to a session.json file specifying PDF, template, and source file paths.",
    )
    parser.add_argument("--pdf", type=str, default=None, help="Path to PDF template.")
    parser.add_argument("--template", type=str, default=None, help="Path to TEMPLATE.xlsx.")
    parser.add_argument(
        "--output-dir",
        type=str,
        default=None,
        help="Output directory for generated PDFs (default: data/output).",
    )
    parser.add_argument(
        "--source-dir",
        type=str,
        default=None,
        help="Directory containing source Excel workbooks.",
    )
    parser.add_argument(
        "--image-dir",
        type=str,
        default=None,
        help="Base directory for resolving relative employee ID card images.",
    )
    parser.add_argument(
        "--page",
        type=int,
        default=0,
        help="Zero-based page index in the PDF template (default: 0).",
    )
    parser.add_argument(
        "--flatten",
        action="store_true",
        help="Flatten AcroForm widgets into static print-ready content.",
    )
    parser.add_argument(
        "--skip-missing-images",
        action="store_true",
        help="Skip missing image overlays instead of failing the record.",
    )
    parser.add_argument(
        "--save-session",
        type=str,
        default=None,
        help="Optional path to save resolved session paths as JSON.",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Use Tkinter file/password dialogs to select missing inputs.",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """Run the pdf-bulk-fill CLI and return a process exit code."""
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)

    session: SessionConfig | None = None
    try:
        if args.session:
            session = load_session(args.session)

        pdf_file = args.pdf or (session.pdf_file if session else None)
        template_file = args.template or (session.template_file if session else None)
        output_dir = (
            args.output_dir
            or (session.output_dir if session and session.output_dir else None)
            or "data/output"
        )

        use_gui = args.gui or (not args.session and not pdf_file and not template_file)
        if use_gui:
            if not pdf_file:
                pdf_file = pick_pdf_file("Select PDF Template", Path.cwd())
            if not template_file:
                template_file = pick_excel_file("Select TEMPLATE.xlsx", Path.cwd())

        if not pdf_file or not template_file:
            print(
                "Error: Both --pdf and --template (or --session) are required.",
                file=sys.stderr,
            )
            return 1

        summary = run_batch(
            pdf_file,
            template_file,
            output_dir,
            session=session,
            source_dir=args.source_dir,
            image_dir=args.image_dir,
            page_index=args.page,
            flatten=args.flatten,
            skip_missing_images=args.skip_missing_images,
            password_provider=prompt_password if use_gui else None,
            file_locator=(
                (lambda fname: pick_excel_file(f"Locate {fname}", Path.cwd()))
                if use_gui
                else None
            ),
            save_session_path=args.save_session,
        )

        print(
            f"Completed: {summary.succeeded}/{summary.total} succeeded, {summary.failed} failed."
        )
        for item in summary.results:
            if item.success:
                print(f"  [OK]   {item.key} -> {item.output_path}")
            else:
                print(f"  [FAIL] {item.key} ({item.identifier}): {item.error}", file=sys.stderr)

        return 0 if summary.failed == 0 else 1
    except PdfBulkFillError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
