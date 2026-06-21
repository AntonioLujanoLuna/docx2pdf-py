"""Built-in engine adapters implementing :class:`ConversionEngine`."""

from __future__ import annotations

from dataclasses import dataclass
from typing import cast

from . import engines
from .engine_protocol import ConversionEngine
from .models import ConversionOptions, ResolvedEngine
from .output import Pathish


@dataclass(frozen=True)
class WordEngine:
    name: ResolvedEngine = "word"

    def available(self) -> bool:
        return engines.word_available()

    def convert(
        self, input_path: Pathish, output_path: Pathish, options: ConversionOptions
    ) -> str:
        return engines.convert_word(
            input_path, output_path, timeout=options.native_engine_timeout
        )


@dataclass(frozen=True)
class LibreOfficeEngine:
    name: ResolvedEngine = "libreoffice"

    def available(self) -> bool:
        return bool(engines.find_libreoffice())

    def convert(
        self, input_path: Pathish, output_path: Pathish, options: ConversionOptions
    ) -> str:
        return engines.convert_libreoffice(
            input_path, output_path, timeout=options.native_engine_timeout
        )


@dataclass(frozen=True)
class WeasyPrintEngine:
    name: ResolvedEngine = "weasyprint"

    def available(self) -> bool:
        return True

    def convert(
        self, input_path: Pathish, output_path: Pathish, options: ConversionOptions
    ) -> str:
        # Deferred to avoid importing the large OOXML renderer during discovery.
        from .converter import _convert_weasyprint

        return _convert_weasyprint(input_path, output_path, options=options)


BUILTIN_ENGINES: tuple[ConversionEngine, ...] = (
    cast(ConversionEngine, WordEngine()),
    cast(ConversionEngine, LibreOfficeEngine()),
    cast(ConversionEngine, WeasyPrintEngine()),
)
