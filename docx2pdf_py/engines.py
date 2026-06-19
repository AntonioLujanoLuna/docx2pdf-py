"""Backends de conversión .docx -> PDF basados en motores de maquetación reales.

Un .docx no guarda páginas fijas: las calcula el motor de maquetación al
renderizar. Por eso el flujo propio (lxml + WeasyPrint) solo puede *aproximar*
la paginación de Word. Si en el sistema hay un motor real, lo usamos y el PDF
tiene **el mismo contenido por página** que el documento:

- **Microsoft Word** (Windows vía COM, macOS vía AppleScript): paginación
  idéntica a la de Word.
- **LibreOffice** (headless, multiplataforma): paginación fiel a cómo LibreOffice
  renderiza el documento (motor distinto al de Word, muy parecido pero no
  garantizado idéntico).

Si no hay ninguno disponible, quien llama recurre al flujo lxml + WeasyPrint.
"""
import os
import platform
import shutil
import subprocess
import tempfile

# Ejecutables de LibreOffice a buscar en el PATH, y rutas habituales de
# instalación en macOS/Windows donde el binario no suele estar en el PATH.
_SOFFICE_NAMES = ("soffice", "libreoffice")
_SOFFICE_PATHS = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)


# -- LibreOffice ------------------------------------------------------------
def find_libreoffice():
    """Ruta al ejecutable de LibreOffice (``soffice``) o ``None`` si no está.

    Se puede forzar con la variable de entorno ``SOFFICE_BIN``.
    """
    override = os.environ.get("SOFFICE_BIN")
    if override and os.path.exists(override):
        return override
    for name in _SOFFICE_NAMES:
        path = shutil.which(name)
        if path:
            return path
    for path in _SOFFICE_PATHS:
        if os.path.exists(path):
            return path
    return None


def convert_libreoffice(in_path, out_path, soffice=None, timeout=120):
    """Convierte ``in_path`` -> ``out_path`` con LibreOffice headless."""
    soffice = soffice or find_libreoffice()
    if not soffice:
        raise RuntimeError("LibreOffice (soffice) no está disponible")
    in_path = os.path.abspath(in_path)
    out_path = os.path.abspath(out_path)
    with tempfile.TemporaryDirectory() as tmp:
        # Perfil de usuario aislado: permite ejecuciones concurrentes y evita
        # chocar con una instancia de LibreOffice ya abierta por el usuario.
        profile = os.path.join(tmp, "profile")
        cmd = [
            soffice, "--headless", "--norestore",
            "-env:UserInstallation=file://" + profile,
            "--convert-to", "pdf", "--outdir", tmp, in_path,
        ]
        proc = subprocess.run(cmd, capture_output=True, timeout=timeout)
        produced = os.path.join(
            tmp, os.path.splitext(os.path.basename(in_path))[0] + ".pdf"
        )
        if proc.returncode != 0 or not os.path.exists(produced):
            detail = (proc.stderr or proc.stdout or b"").decode(errors="replace").strip()
            raise RuntimeError("LibreOffice no pudo convertir el documento" +
                               (f": {detail}" if detail else ""))
        os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
        shutil.copyfile(produced, out_path)
    return out_path


# -- Microsoft Word ---------------------------------------------------------
def word_available():
    """¿Hay Microsoft Word automatizable en este sistema?"""
    system = platform.system()
    if system == "Windows":
        for mod in ("win32com.client", "comtypes.client"):
            try:
                __import__(mod)
                return True
            except Exception:
                continue
        return False
    if system == "Darwin":
        return os.path.exists("/Applications/Microsoft Word.app")
    return False


def convert_word(in_path, out_path):
    """Convierte ``in_path`` -> ``out_path`` automatizando Microsoft Word."""
    in_path = os.path.abspath(in_path)
    out_path = os.path.abspath(out_path)
    system = platform.system()
    if system == "Windows":
        return _convert_word_windows(in_path, out_path)
    if system == "Darwin":
        return _convert_word_macos(in_path, out_path)
    raise RuntimeError("Microsoft Word solo se puede automatizar en Windows o macOS")


def _convert_word_windows(in_path, out_path):
    try:
        import win32com.client as client
        app = client.Dispatch("Word.Application")
    except Exception:
        import comtypes.client as client  # type: ignore
        app = client.CreateObject("Word.Application")
    app.Visible = False
    doc = None
    try:
        doc = app.Documents.Open(in_path, ReadOnly=True)
        doc.SaveAs(out_path, FileFormat=17)  # 17 = wdFormatPDF
    finally:
        if doc is not None:
            doc.Close(False)
        app.Quit()
    return out_path


def _convert_word_macos(in_path, out_path):
    # AppleScript: abre el documento en Word y lo guarda como PDF.
    script = (
        'tell application "Microsoft Word"\n'
        f'  set theDoc to open file name (POSIX file "{in_path}" as string)\n'
        f'  save as theDoc file name "{out_path}" file format format PDF\n'
        '  close theDoc saving no\n'
        'end tell'
    )
    proc = subprocess.run(["osascript", "-e", script], capture_output=True, timeout=120)
    if proc.returncode != 0 or not os.path.exists(out_path):
        detail = proc.stderr.decode(errors="replace").strip()
        raise RuntimeError("Word (macOS) no pudo convertir el documento" +
                           (f": {detail}" if detail else ""))
    return out_path


# -- selección de motor -----------------------------------------------------
def default_engine():
    """Motor que usaría el modo ``auto`` en este sistema, sin convertir nada."""
    if word_available():
        return "word"
    if find_libreoffice():
        return "libreoffice"
    return "weasyprint"
