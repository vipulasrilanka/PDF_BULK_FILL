"""End-to-end batch runner combining Excel lookups, preprocessing, AcroForm, text & image overlays."""

from __future__ import annotations

import os
import re
import unicodedata
from collections import defaultdict
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

_UNSAFE_FILENAME_CHARS = re.compile(r'[\x00-\x1f<>:"/\\|?*]+')


@dataclass(frozen=True)
class RecordResult:
    """Outcome of generating a single PDF record."""

    key: str
    identifier: str
    output_path: Path | None
    success: bool
    error: str | None = None
    warnings: tuple[str, ...] = ()


@dataclass
class BatchSummary:
    """Aggregated outcome of a bulk PDF fill run."""

    total: int
    succeeded: int
    failed: int
    results: list[RecordResult] = field(default_factory=list)


def sanitize_filename_part(text: str) -> str:
    """Normalize Unicode (NFC) and remove filesystem-unsafe characters from a filename segment."""
    normalized = unicodedata.normalize("NFC", str(text))
    cleaned = _UNSAFE_FILENAME_CHARS.sub("_", normalized).strip(" .")
    return cleaned or "record"


def _canonical_filename_key(filename: str) -> str:
    """Compute a filesystem-safe canonical caseless key (Unicode 3.13: NFD(casefold(NFD(s)))).

    Detects collisions across NFC vs NFD (macOS APFS/HFS+), casefolding (e.g. ß vs SS),
    combining-mark reordering, and ASCII case differences.
    """
    nfd = unicodedata.normalize("NFD", filename)
    return unicodedata.normalize("NFD", nfd.casefold())


def _build_record_filename(record: RecordEntry) -> str:
    return (
        f"{sanitize_filename_part(record.key)}"
        f"-{sanitize_filename_part(record.identifier)}.pdf"
    )


def _resolve_raw_value(
    overlay: OverlayDefinition,
    record: RecordEntry,
    store: WorkbookStore,
) -> str:
    """Fetch the raw string value for a single OverlayDefinition and RecordEntry."""
    content = overlay.content

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
        if overlay.content_type == "Image":
            if "File" in overlay.content:
                # Data-driven image: cell in source workbook contains <DocType=...><FileName=...>
                asset_spec: Any = _resolve_raw_value(overlay, record, store)
            else:
                # Constant/static image: properties defined directly in TEMPLATE.xlsx Content
                asset_spec = overlay.content
            image_overlays.append(
                {
                    "name": overlay.name,
                    "asset": asset_spec,
                    "param": overlay.params or {},
                }
            )
            continue

        raw_val = _resolve_raw_value(overlay, record, store)
        processed_val = apply_preprocess(raw_val, overlay.preprocess)

        if overlay.is_concat:
            apply_concat(concat_targets, overlay.name, processed_val)
            continue

        if overlay.content_type == "FormField":
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
    skip_missing_images: bool = False,
    strict_unknown_fields: bool = True,
    strict_maxlen: bool = True,
    warnings_out: list[str] | None = None,
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
                skip_missing=skip_missing_images,
                warnings_out=warnings_out,
            )

        if text_overlays:
            apply_text_overlays_to_page(page, text_overlays)

        if form_fields:
            expanded = expand_comb_groups(
                form_fields,
                strict=strict_maxlen,
                warnings_out=warnings_out,
            )
            apply_form_fields_to_page(
                page,
                expanded,
                strict_maxlen=strict_maxlen,
                strict_unknown_fields=strict_unknown_fields,
                warnings_out=warnings_out,
            )

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
    strict_unknown_fields: bool = True,
    allow_truncation: bool = False,
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

    effective_source_dir = (
        Path(source_dir)
        if source_dir is not None
        else (Path(session.source_dir) if session and session.source_dir else tpl_path.parent)
    )
    effective_image_dir = (
        Path(image_dir)
        if image_dir is not None
        else (Path(session.image_dir) if session and session.image_dir else effective_source_dir)
    )

    records = load_record_list(tpl_path, sheet_name=record_sheet)
    overlays = load_overlay_definitions(tpl_path, sheet_name=overlay_sheet)
    required_files = get_required_source_files(overlays)

    active_session = session or SessionConfig(
        pdf_file=str(pdf_path),
        template_file=str(tpl_path),
        source_files=[],
        output_dir=str(out_dir),
    )

    # Pre-detect canonical caseless output filename collisions across the batch
    planned_filenames = [_build_record_filename(r) for r in records]
    collision_groups: dict[str, list[RecordEntry]] = defaultdict(list)
    for rec, fname in zip(records, planned_filenames):
        collision_groups[_canonical_filename_key(fname)].append(rec)

    results: list[RecordResult] = []
    succeeded = 0
    failed = 0
    batch_written_paths: list[Path] = []

    with WorkbookStore() as store:
        resolved_entries: list[SourceFileEntry] = []
        for file_name in required_files:
            candidate = resolve_source_path(file_name, active_session, effective_source_dir)
            if not candidate.exists() and file_locator is not None:
                located = file_locator(file_name)
                if located:
                    candidate = Path(located)
            store.load(
                candidate,
                alias=file_name,
                password_provider=password_provider,
            )
            resolved_path = candidate.resolve()
            resolved_entries.append(
                SourceFileEntry(
                    name=file_name,
                    path=str(resolved_path.parent),
                    file_path=str(resolved_path),
                )
            )

        if save_session_path is not None:
            updated_session = SessionConfig(
                pdf_file=str(pdf_path.resolve()),
                template_file=str(tpl_path.resolve()),
                source_files=resolved_entries,
                output_dir=str(out_dir.resolve()),
                source_dir=str(effective_source_dir.resolve()),
                image_dir=str(effective_image_dir.resolve()),
                page_index=page_index,
                flatten=flatten,
                skip_missing_images=skip_missing_images,
                allow_truncation=allow_truncation,
            )
            save_session(save_session_path, updated_session)

        out_dir.mkdir(parents=True, exist_ok=True)
        for record, filename in zip(records, planned_filenames):
            colliding = collision_groups[_canonical_filename_key(filename)]
            if len(colliding) > 1:
                conflict_desc = ", ".join(
                    f"({r.key!r}, {r.identifier!r})" for r in colliding
                )
                results.append(
                    RecordResult(
                        key=record.key,
                        identifier=record.identifier,
                        output_path=None,
                        success=False,
                        error=(
                            f"Output filename collision: '{filename}' is targeted by "
                            f"multiple records: {conflict_desc}."
                        ),
                    )
                )
                failed += 1
                continue

            target_pdf = out_dir / filename
            if target_pdf.exists() and any(
                os.path.samefile(target_pdf, prev) for prev in batch_written_paths
            ):
                results.append(
                    RecordResult(
                        key=record.key,
                        identifier=record.identifier,
                        output_path=None,
                        success=False,
                        error=(
                            f"Output filename collision: '{filename}' resolves to an "
                            "output file already written in this batch."
                        ),
                    )
                )
                failed += 1
                continue

            record_warnings: list[str] = []
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
                    image_base_dir=effective_image_dir,
                    skip_missing_images=skip_missing_images,
                    strict_unknown_fields=strict_unknown_fields,
                    strict_maxlen=not allow_truncation,
                    warnings_out=record_warnings,
                )
                batch_written_paths.append(rendered_path)
                results.append(
                    RecordResult(
                        key=record.key,
                        identifier=record.identifier,
                        output_path=rendered_path,
                        success=True,
                        warnings=tuple(record_warnings),
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
                        warnings=tuple(record_warnings),
                    )
                )
                failed += 1

    return BatchSummary(
        total=len(records),
        succeeded=succeeded,
        failed=failed,
        results=results,
    )
