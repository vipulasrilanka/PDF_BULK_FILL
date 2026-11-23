"""AcroForm comb-field (1 character per box) and checkbox filler using PyMuPDF."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Mapping, Sequence

import pymupdf

from pdf_bulk_fill.errors import PdfRenderError
from pdf_bulk_fill.preprocess import split_for_comb

logger = logging.getLogger(__name__)

# Multi-line comb groups for RR-Form.pdf
RR_FORM_COMB_GROUPS: dict[str, tuple[list[str], list[int]]] = {
    "FullName": (["Name 1", "Name 2", "Name 3"], [22, 31, 31]),
    "NameWithInitials": (["Name INI1", "Name INI 2"], [20, 31]),
}

_TRUTHY_STRINGS = frozenset({"1", "yes", "y", "true", "on", "x", "✓", "√"})
_FALSY_STRINGS = frozenset({"", "0", "no", "n", "false", "off"})


def is_truthy_checkbox(value: Any) -> bool:
    """Interpret common Excel/Python values as checkbox on/off booleans."""
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return value != 0
    normalized = str(value).strip().lower()
    if normalized in _TRUTHY_STRINGS:
        return True
    if normalized in _FALSY_STRINGS:
        return False
    raise PdfRenderError(f"Invalid checkbox value: {value!r}")


def expand_comb_groups(
    field_values: Mapping[str, Any],
    comb_groups: Mapping[str, tuple[Sequence[str], Sequence[int]]] = RR_FORM_COMB_GROUPS,
    *,
    uppercase: bool = True,
    strict: bool = False,
    warnings_out: list[str] | None = None,
) -> dict[str, Any]:
    """Expand high-level multi-line keys (e.g. ``FullName``) into individual comb fields."""
    expanded: dict[str, Any] = {}
    for key, value in field_values.items():
        if key in comb_groups:
            target_fields, box_lengths = comb_groups[key]
            text_val = "" if value is None else str(value)
            try:
                lines = split_for_comb(
                    text_val,
                    box_lengths,
                    uppercase=uppercase,
                    strict=True,
                )
            except ValueError as exc:
                if strict:
                    raise PdfRenderError(
                        f"Value for field '{key}' exceeds comb box capacity "
                        f"{list(box_lengths)} ({sum(box_lengths)} max chars): {exc}"
                    ) from exc
                lines = split_for_comb(
                    text_val,
                    box_lengths,
                    uppercase=uppercase,
                    strict=False,
                )
                msg = (
                    f"Truncated comb group '{key}' to fit capacity "
                    f"{list(box_lengths)} ({sum(box_lengths)} max chars): {exc}"
                )
                logger.warning(msg)
                if warnings_out is not None:
                    warnings_out.append(msg)
            for field_name, line_val in zip(target_fields, lines):
                expanded[field_name] = line_val
        else:
            expanded[key] = value
    return expanded


def list_form_fields(
    pdf_path: str | Path,
    page_index: int = 0,
) -> dict[str, dict[str, Any]]:
    """Inspect and return metadata for all AcroForm widgets on a PDF page."""
    path = Path(pdf_path)
    if not path.exists():
        raise PdfRenderError(f"PDF template not found: {path}")

    doc = pymupdf.open(path)
    try:
        if page_index < 0 or page_index >= len(doc):
            raise PdfRenderError(
                f"Page index {page_index} out of range for '{path.name}' ({len(doc)} pages)."
            )
        page = doc[page_index]
        fields: dict[str, dict[str, Any]] = {}
        for widget in page.widgets() or []:
            is_comb = bool(widget.field_flags & pymupdf.PDF_TX_FIELD_IS_COMB)
            on_state = (
                widget.on_state()
                if widget.field_type == pymupdf.PDF_WIDGET_TYPE_CHECKBOX
                else None
            )
            fields[widget.field_name] = {
                "name": widget.field_name,
                "type": widget.field_type_string,
                "max_len": widget.text_maxlen,
                "is_comb": is_comb,
                "rect": tuple(widget.rect),
                "on_state": on_state,
            }
        return fields
    finally:
        doc.close()


def apply_form_fields_to_page(
    page: pymupdf.Page,
    field_values: Mapping[str, Any],
    *,
    strict_maxlen: bool = False,
    strict_unknown_fields: bool = False,
    warnings_out: list[str] | None = None,
) -> dict[str, Any]:
    """Populate AcroForm widgets on an open PyMuPDF page in place.

    Returns a dictionary of ``{field_name: applied_value}``.
    """
    widgets = list(page.widgets() or [])
    known_names = {w.field_name for w in widgets}

    if strict_unknown_fields:
        unknown = sorted(set(field_values.keys()) - known_names)
        if unknown:
            available = ", ".join(sorted(known_names)) if known_names else "(none)"
            raise PdfRenderError(
                f"Unknown PDF form field(s): {', '.join(unknown)}. "
                f"Available fields: {available}."
            )

    applied: dict[str, Any] = {}
    for widget in widgets:
        name = widget.field_name
        if name not in field_values:
            continue

        raw_val = field_values[name]

        if widget.field_type == pymupdf.PDF_WIDGET_TYPE_CHECKBOX:
            checked = is_truthy_checkbox(raw_val)
            state = widget.on_state() if checked else "Off"
            widget.field_value = state
            widget.update()
            applied[name] = checked
        else:
            text_val = "" if raw_val is None else str(raw_val)
            max_len = widget.text_maxlen
            if max_len and len(text_val) > max_len:
                if strict_maxlen:
                    raise PdfRenderError(
                        f"Value for field '{name}' ({len(text_val)} chars) exceeds "
                        f"MaxLen of {max_len}."
                    )
                original_len = len(text_val)
                text_val = text_val[:max_len]
                msg = (
                    f"Truncated field '{name}' from {original_len} to {max_len} chars "
                    f"(saved as {text_val!r})."
                )
                logger.warning(msg)
                if warnings_out is not None:
                    warnings_out.append(msg)
            widget.field_value = text_val
            widget.update()
            applied[name] = text_val

    return applied


def fill_acroform(
    template_pdf: str | Path,
    output_pdf: str | Path,
    field_values: Mapping[str, Any],
    *,
    page_index: int = 0,
    flatten: bool = False,
    strict_maxlen: bool = False,
    strict_unknown_fields: bool = False,
    expand_groups: bool = True,
    warnings_out: list[str] | None = None,
) -> dict[str, Any]:
    """Fill AcroForm comb text fields and checkboxes in a PDF and save to ``output_pdf``."""
    src_path = Path(template_pdf)
    if not src_path.exists():
        raise PdfRenderError(f"PDF template not found: {src_path}")

    resolved_values = (
        expand_comb_groups(
            field_values, strict=strict_maxlen, warnings_out=warnings_out
        )
        if expand_groups
        else dict(field_values)
    )

    doc = pymupdf.open(src_path)
    try:
        if page_index < 0 or page_index >= len(doc):
            raise PdfRenderError(
                f"Page index {page_index} out of range for '{src_path.name}' ({len(doc)} pages)."
            )

        page = doc[page_index]
        applied = apply_form_fields_to_page(
            page,
            resolved_values,
            strict_maxlen=strict_maxlen,
            strict_unknown_fields=strict_unknown_fields,
            warnings_out=warnings_out,
        )

        if flatten:
            doc.bake(annots=False, widgets=True)

        dst_path = Path(output_pdf)
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(dst_path, deflate=True)
        return applied
    finally:
        doc.close()
