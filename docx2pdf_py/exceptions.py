"""Public exception hierarchy for conversion failures."""


class Docx2PdfError(Exception):
    """Base class for package-specific failures."""


class InvalidDocumentError(Docx2PdfError, ValueError):
    """Raised when the input is not a usable OOXML Word document."""


class EngineUnavailableError(Docx2PdfError, RuntimeError):
    """Raised when an explicitly requested conversion engine is unavailable."""


class ConversionError(Docx2PdfError, RuntimeError):
    """Raised when a conversion engine fails to produce a valid PDF."""


class ConversionTimeoutError(ConversionError, TimeoutError):
    """Raised when a conversion exceeds its configured timeout."""
