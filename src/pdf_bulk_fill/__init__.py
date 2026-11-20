"""PDF_BULK_FILL package."""

from pdf_bulk_fill.errors import (
    EncryptedFileError,
    ExcelLoadError,
    ImageOverlayError,
    PdfBulkFillError,
    PdfRenderError,
    SessionError,
    TemplateParseError,
)
from pdf_bulk_fill.excel_loader import WorkbookStore, normalize_key, open_workbook
from pdf_bulk_fill.preprocess import (
    apply_preprocess,
    extract_number,
    format_currency,
    format_date,
    format_number,
    split_for_comb,
    validate_field_value,
)
from pdf_bulk_fill.session import (
    SessionConfig,
    SourceFileEntry,
    load_session,
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
    parse_kv_pairs,
)

__version__ = "0.1.0"

__all__ = [
    "__version__",
    "PdfBulkFillError",
    "SessionError",
    "ExcelLoadError",
    "EncryptedFileError",
    "TemplateParseError",
    "PdfRenderError",
    "ImageOverlayError",
    "SourceFileEntry",
    "SessionConfig",
    "load_session",
    "save_session",
    "resolve_source_path",
    "WorkbookStore",
    "open_workbook",
    "normalize_key",
    "DEFAULT_OVERLAY_SHEET",
    "DEFAULT_RECORD_SHEET",
    "RecordEntry",
    "OverlayDefinition",
    "load_record_list",
    "load_overlay_definitions",
    "get_required_source_files",
    "apply_concat",
    "parse_kv_pairs",
    "apply_preprocess",
    "extract_number",
    "format_date",
    "format_number",
    "format_currency",
    "split_for_comb",
    "validate_field_value",
]
