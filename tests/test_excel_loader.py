"""Unit tests for Excel workbook loading, decryption, and indexed lookups."""

from __future__ import annotations

import glob
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl
import pytest

from pdf_bulk_fill import (
    EncryptedFileError,
    ExcelLoadError,
    WorkbookStore,
    normalize_key,
    open_workbook,
)


def _create_sample_workbook(path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Employees"
    ws.append(["EmpID", "NIC", "FullName", "Email"])
    ws.append([1001, "901234567V", "KAMAL PERERA", "kamal@example.com"])
    ws.append([1002.0, "199512345678", "NIMAL SILVA", None])
    ws.append(["EMP-003", "881112223X", "SUNIL FERNANDO", "sunil@example.com"])
    wb.save(path)
    wb.close()
    return path


def test_normalize_key() -> None:
    assert normalize_key(None) == ""
    assert normalize_key(1001) == "1001"
    assert normalize_key(1002.0) == "1002"
    assert normalize_key(10.5) == "10.5"
    assert normalize_key("  EMP-003  ") == "EMP-003"


def test_open_standard_workbook(tmp_path: Path) -> None:
    xlsx = _create_sample_workbook(tmp_path / "emp.xlsx")
    wb = open_workbook(xlsx)
    try:
        assert "Employees" in wb.sheetnames
        assert wb["Employees"]["C2"].value == "KAMAL PERERA"
    finally:
        wb.close()


def test_open_workbook_missing_file(tmp_path: Path) -> None:
    with pytest.raises(ExcelLoadError, match="not found"):
        open_workbook(tmp_path / "missing.xlsx")


def test_open_workbook_corrupt_unencrypted_file(tmp_path: Path) -> None:
    bad_file = tmp_path / "corrupt.xlsx"
    bad_file.write_bytes(b"not a valid zip or office file")

    with pytest.raises(ExcelLoadError, match="Could not open Excel file"):
        open_workbook(bad_file)


def test_open_encrypted_workbook_requires_password(tmp_path: Path) -> None:
    fake_enc = tmp_path / "locked.xlsx"
    fake_enc.write_bytes(b"encrypted-placeholder")

    with patch("pdf_bulk_fill.excel_loader._is_encrypted_office_file", return_value=True):
        with pytest.raises(EncryptedFileError, match="requires a password"):
            open_workbook(fake_enc)


def test_open_encrypted_workbook_with_password_provider_and_temp_cleanup(
    tmp_path: Path,
) -> None:
    plain = _create_sample_workbook(tmp_path / "plain.xlsx")
    plain_bytes = plain.read_bytes()
    locked = tmp_path / "locked.xlsx"
    locked.write_bytes(b"encrypted-content")

    mock_office = MagicMock()
    mock_office.is_encrypted.return_value = True

    def fake_decrypt(dst_fh) -> None:
        dst_fh.write(plain_bytes)

    mock_office.decrypt.side_effect = fake_decrypt

    before_temps = set(glob.glob(str(Path(tempfile.gettempdir()) / "~temp~*.xlsx")))

    with patch("pdf_bulk_fill.excel_loader.msoffcrypto.OfficeFile", return_value=mock_office):
        wb = open_workbook(locked, password_provider=lambda name: "secret123")
        try:
            assert wb["Employees"]["C2"].value == "KAMAL PERERA"
        finally:
            wb.close()

    mock_office.load_key.assert_called_once_with(password="secret123")
    after_temps = set(glob.glob(str(Path(tempfile.gettempdir()) / "~temp~*.xlsx")))
    assert after_temps == before_temps


def test_open_encrypted_workbook_wrong_password_cleans_up_temp(tmp_path: Path) -> None:
    locked = tmp_path / "locked.xlsx"
    locked.write_bytes(b"encrypted-content")

    mock_office = MagicMock()
    mock_office.is_encrypted.return_value = True
    mock_office.decrypt.side_effect = ValueError("Invalid key")

    before_temps = set(glob.glob(str(Path(tempfile.gettempdir()) / "~temp~*.xlsx")))

    with patch("pdf_bulk_fill.excel_loader.msoffcrypto.OfficeFile", return_value=mock_office):
        with pytest.raises(EncryptedFileError, match="Failed to decrypt"):
            open_workbook(locked, password="wrong")

    after_temps = set(glob.glob(str(Path(tempfile.gettempdir()) / "~temp~*.xlsx")))
    assert after_temps == before_temps


def test_workbook_store_indexed_lookup_numeric_and_alphanumeric(tmp_path: Path) -> None:
    xlsx = _create_sample_workbook(tmp_path / "employees.xlsx")

    with WorkbookStore() as store:
        name = store.load(xlsx)
        assert name == "employees.xlsx"
        assert store.loaded_files == ["employees.xlsx"]
        assert store.get_file_path("employees.xlsx") == xlsx.resolve()

        # Numeric key lookup (int and str)
        assert store.lookup("employees.xlsx", "Employees", 1001, "A", "C") == "KAMAL PERERA"
        assert store.lookup("employees.xlsx", "Employees", "1001", "A", "D") == "kamal@example.com"

        # Float-stored integer key with empty cell value
        assert store.lookup("employees.xlsx", "Employees", 1002, "A", "C") == "NIMAL SILVA"
        assert store.lookup("employees.xlsx", "Employees", 1002, "A", "D") == ""

        # Alphanumeric key lookup
        assert store.lookup("employees.xlsx", "Employees", "EMP-003", "A", "B") == "881112223X"

        # Lookup by NIC column (different key column)
        assert store.lookup("employees.xlsx", "Employees", "901234567V", "B", "C") == "KAMAL PERERA"

        # Out-of-bounds value column returns empty string
        assert store.lookup("employees.xlsx", "Employees", 1001, "A", "Z") == ""

        # Missing primary key returns None
        assert store.lookup("employees.xlsx", "Employees", 9999, "A", "C") is None

    assert store.loaded_files == []


def test_workbook_store_error_cases(tmp_path: Path) -> None:
    xlsx = _create_sample_workbook(tmp_path / "employees.xlsx")

    with WorkbookStore() as store:
        with pytest.raises(ExcelLoadError, match="not loaded"):
            store.get_workbook("missing.xlsx")

        with pytest.raises(ExcelLoadError, match="not loaded"):
            store.get_file_path("missing.xlsx")

        store.load(xlsx)

        with pytest.raises(ExcelLoadError, match="Sheet 'BadSheet' not found"):
            store.lookup("employees.xlsx", "BadSheet", 1001, "A", "C")

        with pytest.raises(ExcelLoadError, match="Invalid column identifier"):
            store.lookup("employees.xlsx", "Employees", 1001, "123", "C")
