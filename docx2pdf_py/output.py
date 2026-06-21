"""Validation and atomic publication of generated PDF files."""

from __future__ import annotations

import os
import re
import shutil
import tempfile
from pathlib import Path
from typing import Union

from .exceptions import ConversionError

Pathish = Union[str, os.PathLike[str]]

_PAGE_OBJECT = re.compile(rb"/Type\s*/Page(?!s)\b")
MAX_PDF_BYTES = 1024 * 1024 * 1024


def validate_pdf(path: Pathish) -> int:
    """Validate basic PDF structure and return its page-object count."""
    candidate = Path(path)
    try:
        with candidate.open("rb") as stream:
            header = stream.read(5)
            stream.seek(0, os.SEEK_END)
            size = stream.tell()
            stream.seek(max(0, size - 65536))
            tail = stream.read()
    except OSError as exc:
        raise ConversionError(f"conversion did not produce a readable PDF: {candidate}") from exc
    if header != b"%PDF-" or size < 32:
        raise ConversionError(f"conversion produced an invalid PDF: {candidate}")
    if size > MAX_PDF_BYTES:
        raise ConversionError(f"conversion produced an oversized PDF: {candidate}")
    if b"%%EOF" not in tail or b"startxref" not in tail:
        raise ConversionError(f"conversion produced an incomplete PDF: {candidate}")
    pages = len(_PAGE_OBJECT.findall(tail))
    if pages == 0:
        # Page objects can precede the tail in larger PDFs. Scan in bounded chunks.
        pages = 0
        with candidate.open("rb") as stream:
            while chunk := stream.read(1024 * 1024):
                pages += len(_PAGE_OBJECT.findall(chunk))
    if pages == 0:
        raise ConversionError(f"conversion produced a PDF with no pages: {candidate}")
    return pages


def publish_pdf(source: Pathish, destination: Pathish) -> str:
    """Validate and atomically replace ``destination`` with ``source``."""
    source_path = Path(source)
    requested_destination = os.fspath(destination)
    destination_path = Path(destination).absolute()
    validate_pdf(source_path)
    destination_path.parent.mkdir(parents=True, exist_ok=True)
    fd, staging_name = tempfile.mkstemp(
        prefix=f".{destination_path.name}.", suffix=".tmp", dir=destination_path.parent
    )
    os.close(fd)
    staging = Path(staging_name)
    try:
        shutil.copyfile(source_path, staging)
        validate_pdf(staging)
        os.replace(staging, destination_path)
    finally:
        staging.unlink(missing_ok=True)
    return requested_destination
