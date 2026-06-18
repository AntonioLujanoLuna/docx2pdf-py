#!/usr/bin/env python3
"""Punto de entrada: convierte un .docx a PDF (solo Python)."""
from docx2pdf_py import convert

# === Configura aquí el documento a convertir ===
RUTA_DOCX = "documento.docx"
RUTA_PDF = "output.pdf"


if __name__ == "__main__":
    convert(RUTA_DOCX, RUTA_PDF)
    print(f"✅ {RUTA_DOCX}  ->  {RUTA_PDF}")
