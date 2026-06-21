"""Isolated Windows Word automation entry point."""

import sys


def main() -> int:
    from .engines import _convert_word_windows

    _convert_word_windows(sys.argv[1], sys.argv[2])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
