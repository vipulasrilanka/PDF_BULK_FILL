"""Custom exception hierarchy for PDF_BULK_FILL."""


class PdfBulkFillError(Exception):
    """Base exception for all PDF_BULK_FILL errors."""


class SessionError(PdfBulkFillError):
    """Raised when loading or saving session configuration fails."""


class ExcelLoadError(PdfBulkFillError):
    """Raised when opening or reading an Excel workbook fails."""


class EncryptedFileError(ExcelLoadError):
    """Raised when an Excel workbook is encrypted and requires a password."""


class TemplateParseError(PdfBulkFillError):
    """Raised when parsing TEMPLATE.xlsx fails."""


class PdfRenderError(PdfBulkFillError):
    """Raised when filling or rendering a PDF fails."""


class ImageOverlayError(PdfRenderError):
    """Raised when loading or placing an image overlay fails."""
