"""Optional Tkinter file and password dialogs for interactive mode.

Tkinter is imported lazily so headless CLI and library usage work even on
Python installations without Tk support.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from pdf_bulk_fill.errors import PdfBulkFillError


def _load_tk() -> tuple[Any, Any, Any]:
    """Lazily import tkinter, filedialog, and simpledialog."""
    try:
        import tkinter as tk
        from tkinter import filedialog, simpledialog
    except ImportError as exc:
        raise PdfBulkFillError(
            "Tkinter is not available in this Python environment. "
            "Use headless CLI flags (--pdf, --template, --session) or install python-tk."
        ) from exc
    return tk, filedialog, simpledialog


def pick_pdf_file(title: str = "Select PDF template", init_dir: str | Path = ".") -> str:
    """Open a file picker dialog to select a PDF file."""
    tk, filedialog, _ = _load_tk()
    root = tk.Tk()
    root.withdraw()
    try:
        return filedialog.askopenfilename(
            title=title,
            filetypes=[("PDF file", "*.pdf")],
            initialdir=str(init_dir),
        )
    finally:
        root.destroy()


def pick_excel_file(title: str = "Select Excel file", init_dir: str | Path = ".") -> str:
    """Open a file picker dialog to select an Excel (.xlsx) file."""
    tk, filedialog, _ = _load_tk()
    root = tk.Tk()
    root.withdraw()
    try:
        return filedialog.askopenfilename(
            title=title,
            filetypes=[("Microsoft Excel file", "*.xlsx")],
            initialdir=str(init_dir),
        )
    finally:
        root.destroy()


def prompt_password(file_name: str) -> str | None:
    """Prompt the user for a password to unlock an encrypted Excel workbook."""
    tk, _, simpledialog = _load_tk()
    root = tk.Tk()
    root.withdraw()
    try:
        return simpledialog.askstring(
            title="Password Required",
            prompt=f"Enter password for {file_name}:",
            show="*",
        )
    finally:
        root.destroy()
