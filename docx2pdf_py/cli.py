"""Command-line interface: docx2pdf-py input.docx [output.pdf]"""
from __future__ import annotations

import argparse
import glob
import os
import sys
from collections.abc import Sequence
from dataclasses import replace

from . import __version__
from .converter import convert_detailed
from .exceptions import Docx2PdfError
from .models import ConversionOptions


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="docx2pdf-py",
        description="Convert a .docx file to PDF using pure Python libraries.",
    )
    parser.add_argument(
        "input",
        nargs="?",
        help="path to the .docx file (default: first .docx found in the current directory)",
    )
    parser.add_argument(
        "output",
        nargs="?",
        default="output.pdf",
        help="path for the output PDF (default: output.pdf)",
    )
    parser.add_argument(
        "-f", "--force",
        action="store_true",
        help="overwrite the output file if it already exists",
    )
    parser.add_argument(
        "-e", "--engine",
        default="auto",
        choices=["auto", "word", "libreoffice", "weasyprint"],
        help=(
            "layout engine to use (default: auto). "
            "'word' and 'libreoffice' produce faithful pagination; "
            "'weasyprint' uses the built-in Python flow (approximate)."
        ),
    )
    parser.add_argument(
        "-q", "--quiet",
        action="store_true",
        help="suppress all output on success",
    )
    parser.add_argument(
        "--fallback",
        choices=["always", "unavailable-only", "never"],
        help="fallback policy for auto engine selection",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="show additional details during conversion",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args(argv)

    src = args.input
    if src is None:
        cands = sorted(glob.glob("*.docx"))
        if not cands:
            parser.error("no input file given and no .docx found in the current directory")
        src = cands[0]
        if not args.quiet:
            print(f"[docx2pdf-py] auto-selected input: {src}", file=sys.stderr)

    if not os.path.exists(src):
        parser.error(f"file not found: {src}")
    if os.path.exists(args.output) and not args.force:
        parser.error(f"output already exists: {args.output} (use -f to overwrite)")

    if args.verbose and not args.quiet:
        print(f"Input:  {src}")
        print(f"Output: {args.output}")
        print(f"Requested engine: {args.engine}")

    try:
        options = ConversionOptions.from_environment()
        if args.fallback:
            options = replace(options, fallback=args.fallback)
        result = convert_detailed(src, args.output, engine=args.engine, options=options)
    except (Docx2PdfError, OSError) as exc:
        parser.error(str(exc))

    if not args.quiet:
        for warning in result.warnings:
            print(f"[docx2pdf-py] {warning}", file=sys.stderr)
        print(f"OK {src} -> {args.output}  [engine: {result.engine}]")
        if args.verbose:
            print(
                f"Elapsed: {result.elapsed_seconds:.3f}s | "
                f"Pages: {result.page_count or 'unknown'} | "
                f"Output: {result.output_bytes} bytes"
            )
            for attempt in result.attempts:
                status = attempt.error or ("available" if attempt.available else "unavailable")
                print(
                    f"Attempt: {attempt.engine} | {attempt.elapsed_seconds:.3f}s | {status}"
                )

    return 0


if __name__ == "__main__":
    sys.exit(main())
