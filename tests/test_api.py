"""Tests for the detailed API, options, timeouts, and atomic output."""

import subprocess
import zipfile

import pytest
from lxml import etree

from docx2pdf_py import ConversionOptions, convert_detailed
from docx2pdf_py import converter as C
from docx2pdf_py import engines as E
from docx2pdf_py.exceptions import ConversionError, ConversionTimeoutError
from docx2pdf_py.output import publish_pdf, validate_pdf
from tests.conftest import FAKE_PDF, document


def _write_pdf(path):
    path = str(path)
    with open(path, "wb") as stream:
        stream.write(FAKE_PDF)
    return path


def test_detailed_result_reports_fallback_engine(make_docx, monkeypatch):
    source = make_docx(document("<w:p/>"))
    monkeypatch.setattr(E, "word_available", lambda: True)
    monkeypatch.setattr(E, "convert_word",
                        lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("boom")))
    monkeypatch.setattr(E, "find_libreoffice", lambda: None)
    monkeypatch.setattr(C, "_convert_weasyprint",
                        lambda i, o, options=None: _write_pdf(o))

    result = convert_detailed(source, "output.pdf")

    assert result.engine == "weasyprint"
    assert result.path == "output.pdf"
    assert result.warnings and "word" in result.warnings[0]


def test_public_api_accepts_path_objects(make_docx, tmp_path, monkeypatch):
    source = make_docx(document("<w:p/>"))
    destination = tmp_path / "output.pdf"
    monkeypatch.setattr(E, "word_available", lambda: False)
    monkeypatch.setattr(E, "find_libreoffice", lambda: None)
    monkeypatch.setattr(C, "_convert_weasyprint",
                        lambda i, o, options=None: _write_pdf(o))

    result = convert_detailed(source, destination)

    assert result.path == str(destination)


def test_weasyprint_timeout_is_terminable(make_docx, tmp_path, monkeypatch):
    source = make_docx(document("<w:p/>"))

    def timeout(*args, **kwargs):
        raise subprocess.TimeoutExpired(args[0], kwargs["timeout"])

    monkeypatch.setattr(C, "run_process", timeout)
    with pytest.raises(ConversionTimeoutError, match="timed out after 1s"):
        C._convert_weasyprint(
            source,
            str(tmp_path / "output.pdf"),
            options=ConversionOptions(weasyprint_timeout=1),
        )


def test_atomic_publish_preserves_existing_output_on_invalid_pdf(tmp_path):
    source = tmp_path / "invalid.pdf"
    destination = tmp_path / "output.pdf"
    source.write_bytes(b"not a pdf")
    destination.write_bytes(b"%PDF-1.4 existing")

    with pytest.raises(ConversionError, match="invalid PDF"):
        publish_pdf(source, destination)

    assert destination.read_bytes() == b"%PDF-1.4 existing"


def test_pdf_validation_checks_structure_and_counts_pages(tmp_path):
    valid = tmp_path / "valid.pdf"
    valid.write_bytes(FAKE_PDF)
    assert validate_pdf(valid) == 1

    incomplete = tmp_path / "incomplete.pdf"
    incomplete.write_bytes(b"%PDF-1.4\n1 0 obj <</Type /Page>> endobj\n")
    with pytest.raises(ConversionError, match="incomplete PDF"):
        validate_pdf(incomplete)


def test_pdf_validation_accepts_structural_pdf_without_detectable_page_objects(tmp_path):
    compressed = tmp_path / "compressed.pdf"
    compressed.write_bytes(
        b"".join(
            (
                b"%PDF-1.5\n",
                b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
                b"2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n",
                (
            b"3 0 obj\n<< /Length 14 /Filter /FlateDecode >>\nstream\n"
            b"x\x9c+\xe4\x02\x00\x00\xee\x00y\nendstream\nendobj\n"
                ),
                b"xref\n0 4\n0000000000 65535 f \n",
                b"trailer\n<< /Root 1 0 R /Size 4 >>\nstartxref\n9\n%%EOF\n",
            )
        )
    )
    assert validate_pdf(compressed) is None


def test_converter_closes_zip_when_initialization_fails(tmp_path, monkeypatch):
    source = tmp_path / "bad.docx"
    with zipfile.ZipFile(source, "w") as archive:
        archive.writestr("word/document.xml", "<broken")
        archive.writestr("word/styles.xml", "<styles/>")

    closed = []
    original_close = zipfile.ZipFile.close

    def tracking_close(self):
        closed.append(True)
        original_close(self)

    monkeypatch.setattr(zipfile.ZipFile, "close", tracking_close)
    with pytest.raises(etree.XMLSyntaxError):
        C.Converter(str(source))
    assert closed
