"""Secure OOXML package reading, namespaces, and unit helpers."""

from __future__ import annotations

import html
import posixpath
import re
import zipfile
from typing import Any

from lxml import etree

from .exceptions import InvalidDocumentError
from .output import Pathish

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
DC = "http://purl.org/dc/elements/1.1/"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
V = "urn:schemas-microsoft-com:vml"

_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
_EMU_PER_PT = 12700.0
_TWIP_PER_PT = 20.0
_TWIP_PER_CM = 566.929


def parse_xml(data: bytes, max_elements: int) -> Any:
    root = etree.fromstring(data, _PARSER)
    if sum(1 for _ in root.iter()) > max_elements:
        raise InvalidDocumentError("OOXML part exceeds the maximum XML element count")
    return root


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


def emu_pt(emu: object) -> float:
    return float(emu) / _EMU_PER_PT  # type: ignore[arg-type]


def first(element: Any, tag: str) -> Any:
    return None if element is None else element.find(w(tag))


def attr(element: Any, name: str) -> Any:
    return None if element is None else element.get(w(name))


def on(element: Any) -> bool | None:
    if element is None:
        return None
    return element.get(w("val")) not in ("0", "false", "off")


def tw_pt(twips: object) -> float:
    return float(twips) / _TWIP_PER_PT  # type: ignore[arg-type]


def tw_cm(twips: object) -> float:
    return float(twips) / _TWIP_PER_CM  # type: ignore[arg-type]


def esc(value: str) -> str:
    return html.escape(value, quote=False)


def keep_spaces(value: str) -> str:
    value = esc(value)
    value = re.sub(r"  +", lambda match: " " + " " * (len(match.group(0)) - 1), value)
    if value.startswith(" "):
        value = " " + value[1:]
    return value


class OOXMLPackage:
    """Bounded reader for potentially untrusted OOXML ZIP packages."""

    def __init__(
        self,
        path: Pathish,
        *,
        max_member_bytes: int,
        max_total_bytes: int,
        max_xml_elements: int,
    ) -> None:
        self.z: zipfile.ZipFile | None = zipfile.ZipFile(path)
        self._read_bytes = 0
        self._max_member_bytes = max_member_bytes
        self._max_total_bytes = max_total_bytes
        self._max_xml_elements = max_xml_elements

    def close(self) -> None:
        if self.z is not None:
            self.z.close()
            self.z = None

    def _read(self, name: str) -> bytes:
        if self.z is None:
            raise RuntimeError("OOXML package is already closed")
        info = self.z.getinfo(name)
        if info.file_size > self._max_member_bytes:
            raise InvalidDocumentError(f"oversized member in .docx: {name}")
        self._read_bytes += info.file_size
        if self._read_bytes > self._max_total_bytes:
            raise InvalidDocumentError("uncompressed .docx exceeds the maximum allowed size")
        return self.z.read(name)

    @staticmethod
    def _resolve_part(base: str, target: str) -> str:
        """Resolve a relationship target without allowing package traversal."""
        if target.startswith("/"):
            resolved = posixpath.normpath(target.lstrip("/"))
        else:
            resolved = posixpath.normpath(posixpath.join(base, target))
        if resolved == ".." or resolved.startswith("../"):
            raise InvalidDocumentError(f"relationship target escapes OOXML package: {target}")
        return resolved

    def _xml_part(self, name: str) -> Any:
        return parse_xml(self._read(name), self._max_xml_elements)

    def _require_xml_part(self, name: str) -> Any:
        try:
            return self._xml_part(name)
        except KeyError as exc:
            raise InvalidDocumentError(
                f"required OOXML part missing from .docx: {name}"
            ) from exc

    def _opt(self, name: str) -> Any:
        try:
            return self._xml_part(name)
        except KeyError:
            return None
