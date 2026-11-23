"""Parser for TEMPLATE.xlsx (Data and Overlay sheets).

Preserves backward compatibility with legacy <key=value> template definitions,
!<CONCAT><TargetName> concatenation, and function syntax FuncName(a, b, ...),
while supporting alphanumeric primary keys and FormField / Image overlay types.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl

from pdf_bulk_fill.errors import TemplateParseError
from pdf_bulk_fill.excel_loader import normalize_key

DEFAULT_OVERLAY_SHEET = "Overlay"
DEFAULT_RECORD_SHEET = "Data"
DEFAULT_SETTINGS_SHEET = "Settings"

# Column indices in Overlay sheet
COL_OVERLAY_INDEX = 0
COL_OVERLAY_NAME = 1
COL_OVERLAY_CONTENT = 2
COL_OVERLAY_PARAM = 3
COL_OVERLAY_PREPROC = 4

# Column indices in Data sheet
COL_RECORD_INDEX = 0
COL_RECORD_KEY = 1
COL_RECORD_IDENTIFIER = 4

SUPPORTED_CONTENT_TYPES = frozenset({"File", "Text", "FormField", "Image"})

# Keys whose values must remain literal strings (preserving precision, leading zeros, etc.)
_LITERAL_STRING_KEYS = frozenset(
    {
        "Text",
        "FileName",
        "DocType",
        "Field",
        "PrimaryKey",
        "PrimeryKey",
        "Sheet",
        "File",
        "Value",
        "Location",
        "Currency",
        "Locale",
        "InputLocale",
    }
)

_KV_PATTERN = re.compile(r"<([^=<>]+)=([^<>]*)>")
_FUNC_PATTERN = re.compile(r"^(\w+)\((.*)\)$")
_CONCAT_PATTERN = re.compile(r"^!<CONCAT><(.+)>$")


@dataclass(frozen=True)
class TemplateSettings:
    """Optional template-level defaults read from the Settings sheet of TEMPLATE.xlsx."""

    currency: str | None = None
    input_locale: str | None = None


@dataclass(frozen=True)
class RecordEntry:
    """Single record to process from the Data sheet."""

    key: str
    identifier: str


@dataclass(frozen=True)
class OverlayDefinition:
    """Parsed row from the Overlay sheet."""

    index: int
    name: str
    content: dict[str, Any]
    params: dict[str, Any] | None = None
    preprocess: dict[str, Any] | None = None

    @property
    def is_concat(self) -> bool:
        return self.name.startswith("!<CONCAT>")

    @property
    def content_type(self) -> str:
        return str(self.content.get("Type", ""))


def extract_scalar_value(raw: str) -> int | float | str:
    """Convert a string to int or float when applicable; otherwise return stripped str."""
    value = raw.strip()
    if not value:
        return ""
    if value.isdigit() or (value.startswith("-") and value[1:].isdigit()):
        return int(value)
    try:
        return float(value)
    except ValueError:
        return value


def parse_function_call(value: Any) -> Any:
    """Parse ``FuncName(arg1, arg2, ...)`` into a structured dict, or return value unchanged."""
    if not isinstance(value, str):
        return value
    match = _FUNC_PATTERN.match(value.strip())
    if not match:
        return value
    func_name = match.group(1)
    raw_args = match.group(2).strip()
    result: dict[str, Any] = {"name": func_name}
    if raw_args:
        for idx, arg in enumerate(raw_args.split(","), start=1):
            result[f"param{idx}"] = extract_scalar_value(arg)
    return result


def parse_kv_pairs(raw: Any) -> dict[str, Any] | None:
    """Parse ``<key=value>`` pairs from a template cell into a dictionary.

    Returns ``None`` if the cell is empty or contains no ``<key=value>`` pairs.
    """
    if raw is None:
        return None
    text = str(raw).strip()
    if not text or text == "None":
        return None

    matches = _KV_PATTERN.findall(text)
    if not matches:
        return None

    parsed: dict[str, Any] = {}
    for key, raw_val in matches:
        clean_key = key.strip()
        if clean_key in _LITERAL_STRING_KEYS:
            parsed[clean_key] = raw_val.strip()
        else:
            scalar = extract_scalar_value(raw_val)
            parsed[clean_key] = parse_function_call(scalar)
    return parsed


def load_record_list(
    template_path: str | Path,
    sheet_name: str = DEFAULT_RECORD_SHEET,
) -> list[RecordEntry]:
    """Load the ordered list of records from the Data sheet of TEMPLATE.xlsx."""
    path = Path(template_path)
    if not path.exists():
        raise TemplateParseError(f"Template file not found: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise TemplateParseError(
                f"Sheet '{sheet_name}' not found in template '{path.name}'."
            )
        sheet = wb[sheet_name]
        records: list[RecordEntry] = []

        for row in sheet.iter_rows(values_only=True):
            if not row or row[COL_RECORD_INDEX] is None:
                break
            row_index = str(row[COL_RECORD_INDEX]).strip()
            if not row_index:
                break
            if not row_index.isdigit():
                continue

            raw_key = row[COL_RECORD_KEY] if len(row) > COL_RECORD_KEY else None
            key = normalize_key(raw_key)
            if not key:
                break

            raw_id = (
                row[COL_RECORD_IDENTIFIER]
                if len(row) > COL_RECORD_IDENTIFIER
                else None
            )
            identifier = str(raw_id).strip() if raw_id is not None else key
            if not identifier:
                identifier = key

            records.append(RecordEntry(key=key, identifier=identifier))
        return records
    finally:
        wb.close()


def load_overlay_definitions(
    template_path: str | Path,
    sheet_name: str = DEFAULT_OVERLAY_SHEET,
) -> list[OverlayDefinition]:
    """Load overlay/field definitions from the Overlay sheet of TEMPLATE.xlsx."""
    path = Path(template_path)
    if not path.exists():
        raise TemplateParseError(f"Template file not found: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise TemplateParseError(
                f"Sheet '{sheet_name}' not found in template '{path.name}'."
            )
        sheet = wb[sheet_name]
        overlays: list[OverlayDefinition] = []

        for row in sheet.iter_rows(values_only=True):
            if not row or row[COL_OVERLAY_INDEX] is None:
                break
            row_index = str(row[COL_OVERLAY_INDEX]).strip()
            if not row_index:
                break
            if not row_index.isdigit():
                continue

            raw_content = row[COL_OVERLAY_CONTENT] if len(row) > COL_OVERLAY_CONTENT else None
            if raw_content is None or not str(raw_content).strip():
                break

            content_str = str(raw_content).strip()
            content = parse_kv_pairs(content_str)
            if not content or "Type" not in content:
                raise TemplateParseError(
                    f"Malformed Content definition at Overlay index {row_index}: {content_str!r}"
                )

            content_type = str(content["Type"])
            if content_type not in SUPPORTED_CONTENT_TYPES:
                raise TemplateParseError(
                    f"Unsupported overlay Type '{content_type}' at Overlay index {row_index}."
                )

            # Support both legacy 'PrimeryKey' and standard 'PrimaryKey'
            if "PrimeryKey" in content and "PrimaryKey" not in content:
                content["PrimaryKey"] = content["PrimeryKey"]

            raw_name = row[COL_OVERLAY_NAME] if len(row) > COL_OVERLAY_NAME else None
            name = str(raw_name).strip() if raw_name is not None else ""
            if not name:
                raise TemplateParseError(
                    f"Missing overlay name at Overlay index {row_index}."
                )

            raw_param = row[COL_OVERLAY_PARAM] if len(row) > COL_OVERLAY_PARAM else None
            raw_preproc = row[COL_OVERLAY_PREPROC] if len(row) > COL_OVERLAY_PREPROC else None

            overlays.append(
                OverlayDefinition(
                    index=int(row_index),
                    name=name,
                    content=content,
                    params=parse_kv_pairs(raw_param),
                    preprocess=parse_kv_pairs(raw_preproc),
                )
            )
        return overlays
    finally:
        wb.close()


def load_template_settings(
    template_path: str | Path,
    sheet_name: str = DEFAULT_SETTINGS_SHEET,
) -> TemplateSettings:
    """Load optional template-level settings (Currency, InputLocale) from TEMPLATE.xlsx.

    Returns empty defaults when the optional ``Settings`` sheet is absent.
    """
    path = Path(template_path)
    if not path.exists():
        raise TemplateParseError(f"Template file not found: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return TemplateSettings()

        sheet = wb[sheet_name]
        currency: str | None = None
        input_locale: str | None = None
        for row in sheet.iter_rows(values_only=True):
            if not row or len(row) < 2 or row[0] is None:
                continue
            key = str(row[0]).strip().lower()
            val = "" if row[1] is None else str(row[1]).strip()
            if not val:
                continue
            if key in ("currency", "defaultcurrency", "default_currency"):
                currency = val
            elif key in ("inputlocale", "input_locale", "locale"):
                input_locale = val
        return TemplateSettings(currency=currency, input_locale=input_locale)
    finally:
        wb.close()


def get_required_source_files(overlays: list[OverlayDefinition]) -> list[str]:
    """Return unique external Excel source filenames referenced by overlays, preserving order."""
    files: list[str] = []
    for overlay in overlays:
        file_name = overlay.content.get("File")
        if isinstance(file_name, str) and file_name.strip() and file_name not in files:
            files.append(file_name)
    return files


def apply_concat(
    resolved_overlays: list[dict[str, Any]],
    concat_name: str,
    value: str,
) -> None:
    """Append ``value`` to the target overlay referenced by ``!<CONCAT><TargetName>``."""
    match = _CONCAT_PATTERN.match(concat_name)
    if not match:
        raise TemplateParseError(
            f"Invalid concatenation directive {concat_name!r}; expected '!<CONCAT><TargetName>'."
        )
    target_name = match.group(1)
    for item in resolved_overlays:
        if item.get("name") == target_name:
            item["string"] = str(item.get("string", "")) + str(value)
            return
    raise TemplateParseError(
        f"Concatenation target '{target_name}' not found in previously resolved overlays."
    )
