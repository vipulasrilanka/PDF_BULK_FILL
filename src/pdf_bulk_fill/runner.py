"""End-to-end batch runner combining Excel lookups, preprocessing, AcroForm, text & image overlays."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence

import pymupdf

from pdf_bulk_fill.errors import ExcelLoadError, PdfBulkFillError, PdfRenderError
from pdf_bulk_fill.excel_loader import WorkbookStore
from pdf_bulk_fill.pdf_form import apply_form_fields_to_page, expand_comb_groups
from pdf_bulk_fill.pdf_image_overlay import apply_image_overlays_to_page
from pdf_bulk_fill.pdf_text_overlay import apply_text_overlays_to_page
from pdf_bulk_fill.preprocess import apply_preprocess
from pdf_bulk_fill.session import (
    SessionConfig,
    SourceFileEntry,
    resolve_source_path,
    save_session,
)
from pdf_bulk_fill.template_parser import (
    DEFAULT_OVERLAY_SHEET,
    DEFAULT_RECORD_SHEET,
    OverlayDefinition,
    RecordEntry,
    apply_concat,
    get_required_source_files,
    load_overlay_definitions,
    load_record_list,
)

_UNSAFE_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*]+')


@dataclass(frozen=True)
class RecordResult:
    """Outcome of generating a single PDF record."""

    key: str
    identifier: str
    output_path: Path | None
    success: bool
    error: str | None = None


@dataclass
class BatchSummary:
    """Aggregated outcome of a bulk PDF fill run."""

    total: int
    succeeded: int
    failed: int
    results: list[RecordResult] = field(default_factory=list)


def sanitize_filename_part(text: str) -> str:
    """Remove filesystem-unsafe characters from a filename segment."""
    cleaned = _UNSAFE_FILENAME_CHARS.sub("_", str(text).strip())
    return cleaned or "record"


def _resolve_raw_value(
    overlay: OverlayDefinition,
    record: RecordEntry,
    store: WorkbookStore,
) -> str:
    """Fetch the raw string value for a single OverlayDefinition and RecordEntry."""
    content = overlay.content
    content_type = overlay.content_type

    if "File" in content:
        file_name = str(content["File"])
        sheet_name = str(content["Sheet"])
        key_col = str(content.get("PrimaryKey", content.get("PrimeryKey", "")))
        value_col = str(content["Value"])
        value = store.lookup(file_name, sheet_name, record.key, key_col, value_col)
        if value is None:
            raise ExcelLoadError(
                f"Primary key '{record.key}' not found in {file_name}[{sheet_name}:{key_col}]."
            )
        return value

    if content_type == "Image":
        return str(content.get("Path", content.get("Text", "")))

    return str(content.get("Text", ""))


def resolve_record_payload(
    record: RecordEntry,
    overlays: Sequence[OverlayDefinition],
    store: WorkbookStore,
) -> tuple[dict[str, Any], list[dict[str, Any]], list[dict[str, Any]]]:
    """Resolve all overlays for one record into (form_fields, text_overlays, image_overlays)."""
    form_fields: dict[str, Any] = {}
    text_overlays: list[dict[str, Any]] = []
    image_overlays: list[dict[str, Any]] = []

    # Unified list so !<CONCAT><TargetName> can append to either text overlays or form fields
    concat_targets: list[dict[str, Any]] = []

    for overlay in overlays:
        raw_val = _resolve_raw_value(overlay, record, store)
        processed_val = apply_preprocess(raw_val, overlay.preprocess)

        if overlay.is_concat:
            apply_concat(concat_targets, overlay.name, processed_val)
            continue

        if overlay.content_type == "Image":
            image_overlays.append(
                {
                    "name": overlay.name,
                    "path": processed_val,
                    "param": overlay.params or {},
                }
            )
        elif overlay.content_type == "FormField":
            params = overlay.params or {}
            field_name = str(params.get("Field", overlay.name))
            entry = {"name": overlay.name, "field": field_name, "string": processed_val}
            concat_targets.append(entry)
        else:
            entry = {
                "name": overlay.name,
                "string": processed_val,
                "param": overlay.params or {},
            }
            text_overlays.append(entry)
            concat_targets.append(entry)

    # Extract final form_fields values (after any !<CONCAT> mutations)
    for target in concat_targets:
        if "field" in target:
            form_fields[target["field"]] = target["string"]

    return form_fields, text_overlays, image_overlays


def render_record_pdf(
    template_pdf: str | Path,
    output_pdf: str | Path,
    *,
    form_fields: Mapping[str, Any] | None = None,
    text_overlays: Sequence[Mapping[str, Any]] | None = None,
    image_overlays: Sequence[Mapping[str, Any]] | None = None,
    page_index: int = 0,
    flatten: bool = False,
    image_base_dir: str | Path | None = None,
    record_key: str | None = None,
    skip_missing_images: bool = False,
) -> Path:
    """Render a single output PDF combining images, coordinate text, and AcroForm fields."""
    src_path = Path(template_pdf)
    if not src_path.exists():
        raise PdfRenderError(f"PDF template not found: {src_path}")

    doc = pymupdf.open(src_path)
    try:
        if page_index < 0 or page_index >= len(doc):
            raise PdfRenderError(
                f"Page index {page_index} out of range for '{src_path.name}' ({len(doc)} pages)."
            )
        page = doc[page_index]

        if image_overlays:
            apply_image_overlays_to_page(
                page,
                image_overlays,
                base_dir=image_base_dir,
                record_key=record_key,
                skip_missing=skip_missing_images,
            )

        if text_overlays:
            apply_text_overlays_to_page(page, text_overlays)

        if form_fields:
            expanded = expand_comb_groups(form_fields)
            apply_form_fields_to_page(page, expanded)

        if flatten:
            doc.bake(annots=False, widgets=True)

        dst_path = Path(output_pdf)
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(dst_path, deflate=True)
        return dst_path
    finally:
        doc.close()


def run_batch(
    pdf_template: str | Path,
    template_excel: str | Path,
    output_dir: str | Path,
    *,
    session: SessionConfig | None = None,
    source_dir: str | Path | None = None,
    image_dir: str | Path | None = None,
    page_index: int = 0,
    flatten: bool = False,
    skip_missing_images: bool = False,
    record_sheet: str = DEFAULT_RECORD_SHEET,
    overlay_sheet: str = DEFAULT_OVERLAY_SHEET,
    password_provider: Callable[[str], str | None] | None = None,
    file_locator: Callable[[str], str | Path | None] | None = None,
    save_session_path: str | Path | None = None,
) -> BatchSummary:
    """Execute a full batch PDF fill run across all records in ``template_excel``."""
    pdf_path = Path(pdf_template)
    tpl_path = Path(template_excel)
    out_dir = Path(output_dir)
    default_src_dir = Path(source_dir) if source_dir is not None else tpl_path.parent
    img_base_dir = Path(image_dir) if image_dir is not None else default_src_dir

    records = load_record_list(tpl_path, sheet_name=record_sheet)
    overlays = load_overlay_definitions(tpl_path, sheet_name=overlay_sheet)
    required_files = get_required_source_files(overlays)

    active_session = session or SessionConfig(
        pdf_file=str(pdf_path),
        template_file=str(tpl_path),
        source_files=[],
        output_dir=str(out_dir),
    )

    results: list[RecordResult] = []
    succeeded = 0
    failed = 0

    with WorkbookStore() as store:
        resolved_entries: list[SourceFileEntry] = []
        for file_name in required_files:
            candidate = resolve_source_path(file_name, active_session, default_src_dir)
            if not candidate.exists() and file_locator is not None:
                located = file_locator(file_name)
                if located:
                    candidate = Path(located)
            store.load(
                candidate,
                alias=file_name,
                password_provider=password_provider,
            )
            resolved_entries.append(
                SourceFileEntry(name=file_name, path=str(candidate.resolve().parent))
            )

        if save_session_path is not None:
            updated_session = SessionConfig(
                pdf_file=str(pdf_path.resolve()),
                template_file=str(tpl_path.resolve()),
                source_files=resolved_entries,
                output_dir=str(out_dir.resolve()),
            )
            save_session(save_session_path, updated_session)

        out_dir.mkdir(parents=True, exist_ok=True)
        for record in records:
            filename = (
                f"{sanitize_filename_part(record.key)}"
                f"-{sanitize_filename_part(record.identifier)}.pdf"
            )
            target_pdf = out_dir / filename
            try:
                form_fields, text_overlays, image_overlays = resolve_record_payload(
                    record, overlays, store
                )
                rendered_path = render_record_pdf(
                    pdf_path,
                    target_pdf,
                    form_fields=form_fields,
                    text_overlays=text_overlays,
                    image_overlays=image_overlays,
                    page_index=page_index,
                    flatten=flatten,
                    image_base_dir=img_base_dir,
                    record_key=record.key,
                    skip_missing_images=skip_missing_images,
                )
                results.append(
                    RecordResult(
                        key=record.key,
                        identifier=record.identifier,
                        output_path=rendered_path,
                        success=True,
                    )
                )
                succeeded += 1
            except PdfBulkFillError as exc:
                results.append(
                    RecordResult(
                        key=record.key,
                        identifier=record.identifier,
                        output_path=None,
                        success=False,
                        error=str(exc),
                    )
                )
                failed += 1

    return BatchSummary(
        total=len(records),
        succeeded=succeeded,
        failed=failed,
        results=results,
    )
