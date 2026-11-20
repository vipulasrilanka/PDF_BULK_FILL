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
from pdf_bulk_fill.session import (
    SessionConfig,
    SourceFileEntry,
    load_session,
    resolve_source_path,
    save_session,
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
]
