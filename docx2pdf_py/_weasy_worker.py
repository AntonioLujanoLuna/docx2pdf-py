"""Subprocess entry point used to make WeasyPrint rendering terminable."""

import sys
from pathlib import Path


def main() -> int:
    from weasyprint import HTML

    html_path, output_path = map(Path, sys.argv[1:3])
    HTML(filename=str(html_path)).write_pdf(str(output_path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
