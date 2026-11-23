"""Excel workbook loading, encrypted .xlsx decryption, and indexed lookups."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Any, Callable

import msoffcrypto
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.workbook import Workbook

from pdf_bulk_fill.errors import EncryptedFileError, ExcelLoadError


def normalize_key(value: Any) -> str:
    """Normalize a primary key (numeric or alphanumeric) to a stripped string."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _is_encrypted_office_file(path: Path) -> bool:
    """Check whether a file is an encrypted MS Office document."""
    try:
        with path.open("rb") as fh:
            office_file = msoffcrypto.OfficeFile(fh)
            return bool(office_file.is_encrypted())
    except Exception:
        return False


def _decrypt_and_load_workbook(path: Path, password: str) -> Workbook:
    """Decrypt an encrypted .xlsx file via a temporary file with guaranteed cleanup."""
    tmp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix="~temp~", suffix=".xlsx", delete=False
        ) as tmp_file:
            tmp_path = Path(tmp_file.name)

        with path.open("rb") as src_fh:
            office_file = msoffcrypto.OfficeFile(src_fh)
            office_file.load_key(password=password)
            with tmp_path.open("wb") as dst_fh:
                office_file.decrypt(dst_fh)

        return openpyxl.load_workbook(tmp_path, data_only=True)
    except Exception as exc:
        raise EncryptedFileError(
            f"Failed to decrypt workbook '{path.name}': {exc}"
        ) from exc
    finally:
        if tmp_path is not None and tmp_path.exists():
            os.unlink(tmp_path)


def open_workbook(
    file_path: str | Path,
    password: str | None = None,
    password_provider: Callable[[str], str | None] | None = None,
) -> Workbook:
    """Open an .xlsx workbook, decrypting it if password-protected.

    Args:
        file_path: Path to the .xlsx workbook.
        password: Optional password to use if the workbook is encrypted.
        password_provider: Optional callback receiving the filename and returning
            a password if the workbook is encrypted and ``password`` was not given.

    Raises:
        ExcelLoadError: If the file does not exist or cannot be opened.
        EncryptedFileError: If the file is encrypted and cannot be decrypted.
    """
    path = Path(file_path)
    if not path.exists():
        raise ExcelLoadError(f"Excel file not found: {path}")

    try:
        return openpyxl.load_workbook(path, data_only=True)
    except Exception as open_exc:
        if not _is_encrypted_office_file(path):
            raise ExcelLoadError(f"Could not open Excel file '{path}': {open_exc}") from open_exc

        resolved_password = password
        if resolved_password is None and password_provider is not None:
            resolved_password = password_provider(path.name)

        if not resolved_password:
            raise EncryptedFileError(
                f"Excel file '{path.name}' is encrypted and requires a password."
            ) from open_exc

        return _decrypt_and_load_workbook(path, resolved_password)


class WorkbookStore:
    """Manages loaded Excel workbooks with cached O(1) primary-key sheet indexes."""

    def __init__(self) -> None:
        self._workbooks: dict[str, Workbook] = {}
        self._paths: dict[str, Path] = {}
        # Cache key: (file_name, sheet_name, key_col_index) -> {normalized_key: row_values}
        self._indexes: dict[tuple[str, str, int], dict[str, tuple[Any, ...]]] = {}

    def load(
        self,
        file_path: str | Path,
        *,
        alias: str | None = None,
        password: str | None = None,
        password_provider: Callable[[str], str | None] | None = None,
    ) -> str:
        """Load a workbook into the store and return its registered name."""
        path = Path(file_path)
        name = alias or path.name
        workbook = open_workbook(
            path, password=password, password_provider=password_provider
        )
        self._workbooks[name] = workbook
        self._paths[name] = path.resolve()
        # Invalidate any previously cached indexes for this workbook name
        self._indexes = {k: v for k, v in self._indexes.items() if k[0] != name}
        return name

    @property
    def loaded_files(self) -> list[str]:
        """Return list of registered workbook names."""
        return list(self._workbooks.keys())

    def get_workbook(self, file_name: str) -> Workbook:
        """Return a loaded openpyxl Workbook by its registered name."""
        if file_name not in self._workbooks:
            raise ExcelLoadError(f"Workbook '{file_name}' is not loaded.")
        return self._workbooks[file_name]

    def get_file_path(self, file_name: str) -> Path:
        """Return the resolved Path for a loaded workbook."""
        if file_name not in self._paths:
            raise ExcelLoadError(f"Workbook '{file_name}' is not loaded.")
        return self._paths[file_name]

    def _get_or_build_index(
        self, file_name: str, sheet_name: str, key_col_idx: int
    ) -> dict[str, tuple[Any, ...]]:
        cache_key = (file_name, sheet_name, key_col_idx)
        if cache_key in self._indexes:
            return self._indexes[cache_key]

        workbook = self.get_workbook(file_name)
        if sheet_name not in workbook.sheetnames:
            raise ExcelLoadError(
                f"Sheet '{sheet_name}' not found in workbook '{file_name}'."
            )

        sheet = workbook[sheet_name]
        index: dict[str, tuple[Any, ...]] = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if key_col_idx - 1 >= len(row):
                continue
            key_str = normalize_key(row[key_col_idx - 1])
            if key_str and key_str not in index:
                index[key_str] = row

        self._indexes[cache_key] = index
        return index

    def lookup_raw(
        self,
        file_name: str,
        sheet_name: str,
        primary_key: Any,
        key_col: str,
        value_col: str,
    ) -> tuple[bool, Any]:
        """Look up a raw native Excel cell value by primary key using a cached sheet index.

        Returns:
            ``(True, raw_value)`` when ``primary_key`` is found (preserving native
            Python ``int``, ``float``, ``Decimal``, ``datetime``, ``str``, or ``""``
            for an empty cell), or ``(False, None)`` when ``primary_key`` is absent.
        """
        try:
            key_col_idx = column_index_from_string(key_col.strip().upper())
            value_col_idx = column_index_from_string(value_col.strip().upper())
        except ValueError as exc:
            raise ExcelLoadError(
                f"Invalid column identifier (key_col={key_col!r}, value_col={value_col!r}): {exc}"
            ) from exc

        index = self._get_or_build_index(file_name, sheet_name, key_col_idx)
        row = index.get(normalize_key(primary_key))
        if row is None:
            return False, None

        if value_col_idx - 1 >= len(row):
            return True, ""
        cell_val = row[value_col_idx - 1]
        if cell_val is None:
            return True, ""
        return True, cell_val

    def lookup(
        self,
        file_name: str,
        sheet_name: str,
        primary_key: Any,
        key_col: str,
        value_col: str,
    ) -> str | None:
        """Look up a cell value as a string by primary key using a cached sheet index.

        Returns:
            The string value from ``value_col`` (``""`` if the cell is empty),
            or ``None`` if ``primary_key`` is not found in ``key_col``.
        """
        found, raw_val = self.lookup_raw(
            file_name, sheet_name, primary_key, key_col, value_col
        )
        if not found:
            return None
        return str(raw_val)

    def close(self) -> None:
        """Close all loaded workbooks and clear cached indexes."""
        for workbook in self._workbooks.values():
            workbook.close()
        self._workbooks.clear()
        self._paths.clear()
        self._indexes.clear()

    def __enter__(self) -> WorkbookStore:
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        self.close()
