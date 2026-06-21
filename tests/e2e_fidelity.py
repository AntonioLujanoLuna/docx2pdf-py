#!/usr/bin/env python3
"""Generated cross-engine fidelity corpus with per-page text and raster checks."""

import os
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path

import fitz
from pypdf import PdfReader

from docx2pdf_py import convert
from docx2pdf_py.engines import find_libreoffice
from tests.e2e_smoke import CONTENT_TYPES, DECL, DOCUMENT_RELS, ROOT_RELS, STYLES, W


@dataclass(frozen=True)
class Case:
    name: str
    body: str
    page_text: tuple[str, ...]


CASES = (
    Case(
        "paragraphs",
        "<w:p><w:r><w:t>Alpha paragraph</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Beta paragraph</w:t></w:r></w:p>",
        ("Alpha paragraph\nBeta paragraph",),
    ),
    Case(
        "page-break",
        "<w:p><w:r><w:t>First page</w:t></w:r></w:p>"
        '<w:p><w:r><w:br w:type="page"/><w:t>Second page</w:t></w:r></w:p>',
        ("First page", "Second page"),
    ),
    Case(
        "table",
        "<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/>"
        "</w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell A</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>Cell B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        ("Cell A\nCell B",),
    ),
)


def make_docx(path: Path, body: str) -> None:
    document = (
        DECL
        + f'<w:document xmlns:w="{W}"><w:body>'
        + body
        + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'
        + "</w:body></w:document>"
    )
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS)
        archive.writestr("word/document.xml", document)
        archive.writestr("word/styles.xml", STYLES)


def assert_visible_ink(path: Path) -> None:
    document = fitz.open(path)
    for page in document:
        pixels = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5), colorspace=fitz.csGRAY)
        non_white = sum(value < 245 for value in pixels.samples)
        assert non_white / len(pixels.samples) > 0.0005, "rendered page appears blank"


def main() -> int:
    engine = os.environ.get("E2E_ENGINE", "libreoffice")
    if engine == "libreoffice" and not find_libreoffice():
        print("LibreOffice unavailable; fidelity corpus skipped.")
        return 0
    with tempfile.TemporaryDirectory() as tmp:
        directory = Path(tmp)
        for case in CASES:
            source = directory / f"{case.name}.docx"
            output = directory / f"{case.name}.pdf"
            make_docx(source, case.body)
            convert(source, output, engine=engine)
            reader = PdfReader(output)
            assert len(reader.pages) == len(case.page_text)
            for page, expected in zip(reader.pages, case.page_text):
                actual = "\n".join(line.strip() for line in page.extract_text().splitlines())
                for line in expected.splitlines():
                    assert line in actual
            assert_visible_ink(output)
            print(f"OK {engine}: {case.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
