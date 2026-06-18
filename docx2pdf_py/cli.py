"""Interfaz de línea de comandos: docx2pdf-py entrada.docx [salida.pdf]."""
import argparse
import glob
import os
import sys

from .converter import convert


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="docx2pdf-py",
        description="Convierte un .docx a PDF usando solo librerías de Python.",
    )
    parser.add_argument("entrada", nargs="?",
                        help="ruta del .docx (por defecto, el primero del directorio)")
    parser.add_argument("salida", nargs="?", default="output.pdf",
                        help="ruta del PDF de salida (por defecto: output.pdf)")
    args = parser.parse_args(argv)

    src = args.entrada
    if src is None:
        cands = sorted(glob.glob("*.docx"))
        if not cands:
            parser.error("no se indicó entrada y no hay ningún .docx en el directorio")
        src = cands[0]
    if not os.path.exists(src):
        parser.error(f"no existe el archivo: {src}")

    convert(src, args.salida)
    print(f"✅ {src}  ->  {args.salida}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
