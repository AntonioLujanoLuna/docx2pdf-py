"""Backend contract used by conversion dispatch and third-party engines."""

from __future__ import annotations

from typing import Protocol

from .models import ConversionOptions
from .output import Pathish


class ConversionEngine(Protocol):
    """A discoverable DOCX-to-PDF conversion backend."""

    name: str

    def available(self) -> bool:
        """Return whether the backend can run in the current environment."""

    def convert(
        self, input_path: Pathish, output_path: Pathish, options: ConversionOptions
    ) -> str:
        """Convert a document and return the requested output path."""
