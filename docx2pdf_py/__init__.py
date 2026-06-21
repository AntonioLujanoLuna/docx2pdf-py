"""docx2pdf_py — conversión fiel de .docx a PDF usando solo librerías de Python.

Lee el OOXML del documento (estilos reales: fuentes, colores, bordes, sombreados,
tablas, imágenes, cabecera/pie) y lo recrea como HTML que WeasyPrint pagina a PDF.

Uso:
    from docx2pdf_py import convert
    convert("entrada.docx", "salida.pdf")
"""
from importlib.metadata import PackageNotFoundError, version

from .api import convert, convert_batch, convert_detailed
from .converter import Converter
from .engine_protocol import ConversionEngine
from .engines import default_engine, find_libreoffice, word_available
from .exceptions import (
    ConversionError,
    ConversionTimeoutError,
    Docx2PdfError,
    EngineUnavailableError,
    InvalidDocumentError,
)
from .models import (
    BatchItemResult,
    ConversionAttempt,
    ConversionOptions,
    ConversionResult,
)

try:
    __version__ = version("docx2pdf-py")
except PackageNotFoundError:
    __version__ = "0+unknown"
__all__ = [
    "convert", "convert_detailed", "convert_batch", "Converter", "ConversionOptions",
    "ConversionResult", "ConversionAttempt", "BatchItemResult", "ConversionEngine",
    "Docx2PdfError", "InvalidDocumentError", "EngineUnavailableError",
    "ConversionError", "ConversionTimeoutError", "__version__",
    "default_engine", "find_libreoffice", "word_available",
]
