"""Installed-package metadata and typing marker checks."""

from importlib.metadata import version
from importlib.resources import files

import docx2pdf_py


def test_runtime_version_comes_from_distribution_metadata():
    assert docx2pdf_py.__version__ == version("docx2pdf-py")


def test_typing_marker_is_packaged():
    assert files("docx2pdf_py").joinpath("py.typed").is_file()
