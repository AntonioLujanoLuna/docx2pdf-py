#!/usr/bin/env python3
"""docx -> PDF con SOLO librerías de Python, fiel al original.

Lee el OOXML del .docx (estilos reales: fuentes, colores, bordes, sombreados,
tablas, cabecera/pie) y lo recrea como HTML, que WeasyPrint maqueta y pagina a
PDF. Las fuentes Calibri/Georgia se mapean a sus equivalentes métricos libres
Carlito/Gelasio.

Uso:
    from docx2pdf_py import convert
    convert("entrada.docx", "salida.pdf")
"""
import zipfile
import html as _html
import re
from lxml import etree
from weasyprint import HTML

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def w(tag):
    return f"{{{W}}}{tag}"


def emu_pt(emu):
    return float(emu) / 12700.0


def first(el, tag):
    if el is None:
        return None
    return el.find(w(tag))


def attr(el, name):
    if el is None:
        return None
    return el.get(w(name))


def on(el):
    """Un elemento booleano OOXML (w:b, w:i, ...) está activo salvo val=0/false."""
    if el is None:
        return None
    v = el.get(w("val"))
    return v not in ("0", "false", "off")


def tw_pt(twips):
    return float(twips) / 20.0


def tw_cm(twips):
    return float(twips) / 566.929


def esc(s):
    return _html.escape(s, quote=False)


def keep_spaces(s):
    """Conserva espacios múltiples/iniciales (HTML los colapsaría)."""
    s = esc(s)
    s = re.sub(r"  +", lambda m: " " + " " * (len(m.group(0)) - 1), s)
    if s.startswith(" "):
        s = " " + s[1:]
    return s


FONT_MAP = {
    "Calibri": "Carlito, Calibri, sans-serif",
    "Georgia": "Gelasio, Georgia, serif",
}

# Interlineado por defecto (ajustable para casar con el motor de referencia).
import os
BODY_LINE_HEIGHT = float(os.environ.get("BODY_LH", "1.0"))
CELL_LINE_HEIGHT = float(os.environ.get("CELL_LH", "1.16"))


def font_stack(name):
    if not name:
        return None
    return FONT_MAP.get(name, f"'{name}', sans-serif")


# ----------------------------------------------------------------------------
# Resolución de formato de "run" (carácter)
# ----------------------------------------------------------------------------
def rpr_dict(rpr):
    """Extrae propiedades de carácter de un <w:rPr>."""
    d = {}
    if rpr is None:
        return d
    fonts = first(rpr, "rFonts")
    if fonts is not None and attr(fonts, "ascii"):
        d["font"] = attr(fonts, "ascii")
    b = on(first(rpr, "b"))
    if b is not None:
        d["bold"] = b
    i = on(first(rpr, "i"))
    if i is not None:
        d["italic"] = i
    strike = on(first(rpr, "strike"))
    if strike is not None:
        d["strike"] = strike
    u = first(rpr, "u")
    if u is not None:
        d["underline"] = attr(u, "val") not in (None, "none")
    color = first(rpr, "color")
    if color is not None:
        v = attr(color, "val")
        if v and v != "auto":
            d["color"] = "#" + v
    sz = first(rpr, "sz")
    if sz is not None:
        d["size"] = float(attr(sz, "val")) / 2.0
    va = first(rpr, "vertAlign")
    if va is not None:
        d["va"] = attr(va, "val")
    return d


def run_css(d):
    css = []
    if d.get("font"):
        css.append(f"font-family:{font_stack(d['font'])}")
    if "bold" in d:
        css.append("font-weight:" + ("bold" if d["bold"] else "normal"))
    if "italic" in d:
        css.append("font-style:" + ("italic" if d["italic"] else "normal"))
    deco = []
    if d.get("underline"):
        deco.append("underline")
    if d.get("strike"):
        deco.append("line-through")
    if deco:
        css.append("text-decoration:" + " ".join(deco))
    if d.get("color"):
        css.append("color:" + d["color"])
    size = d.get("size")
    va = d.get("va")
    if va in ("superscript", "subscript"):
        css.append("vertical-align:" + ("super" if va == "superscript" else "sub"))
        if size:
            size = size * 0.7
    if size:
        css.append(f"font-size:{size:.1f}pt")
    return ";".join(css)


def border_css(b):
    """CSS de un borde OOXML (<w:top>/<w:bottom>/...)."""
    if b is None:
        return None
    val = attr(b, "val")
    if val in (None, "nil", "none"):
        return "none"
    sz = attr(b, "sz")
    width = max(float(sz) / 8.0, 0.5) if sz else 0.5
    color = attr(b, "color") or "000000"
    if color == "auto":
        color = "000000"
    return f"{width:.2f}pt solid #{color}"


class Converter:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.doc = etree.fromstring(self.z.read("word/document.xml"))
        self.styles = etree.fromstring(self.z.read("word/styles.xml"))
        self.rels = self._index_rels()
        self.style_rpr = self._index_styles()
        self.default = {"font": "Calibri", "color": "#4a4a4a", "size": 10.0}
        self._img_cache = {}
        self._pending_floats = []  # imágenes flotantes a emitir tras el bloque

        # cabecera/pie de la sección por defecto (según sectPr)
        sect = self.doc.find(w("body")).find(w("sectPr"))
        h = self._ref_part(sect, "headerReference")
        f = self._ref_part(sect, "footerReference")
        self.header_xml = h if h is not None else self._opt("word/header1.xml")
        self.footer_xml = f if f is not None else self._opt("word/footer1.xml")

    def _opt(self, name):
        try:
            return etree.fromstring(self.z.read(name))
        except KeyError:
            return None

    def _index_rels(self):
        try:
            root = etree.fromstring(self.z.read("word/_rels/document.xml.rels"))
        except KeyError:
            return {}
        return {r.get("Id"): r.get("Target") for r in root}

    def _ref_part(self, sect, tag):
        """Carga la parte (header/footer) referenciada como type='default'."""
        if sect is None:
            return None
        for ref in sect.findall(w(tag)):
            if ref.get(w("type")) == "default":
                rid = ref.get(f"{{{R}}}id")
                target = self.rels.get(rid)
                if target:
                    return self._opt("word/" + target)
        return None

    def _index_styles(self):
        out = {}
        for st in self.styles.findall(w("style")):
            sid = attr(st, "styleId")
            out[sid] = rpr_dict(first(st, "rPr"))
        return out

    # -- runs --------------------------------------------------------------
    def render_runs(self, p, base):
        """HTML de los runs de un párrafo, heredando 'base' (rPr de su estilo).

        Ignora los campos (fldChar/instrText) y su valor cacheado: p.ej. el
        campo PAGE del pie guarda un número que no debe imprimirse tal cual.
        """
        parts = []
        in_field = False
        for child in p:
            tag = etree.QName(child).localname
            if tag == "hyperlink":
                inner = self.render_runs(child, base)
                parts.append(f'<a style="color:inherit;text-decoration:underline">{inner}</a>')
            elif tag == "r":
                types = [fc.get(w("fldCharType")) for fc in child.findall(w("fldChar"))]
                if "begin" in types:
                    in_field = True
                skip = in_field or child.find(w("instrText")) is not None
                if "end" in types:
                    in_field = False
                if not skip:
                    parts.append(self._render_run(child, base))
        return "".join(parts)

    def _render_run(self, r, base):
        d = dict(base)
        d.update(rpr_dict(first(r, "rPr")))
        chunks = []
        images = []
        for child in r:
            tag = etree.QName(child).localname
            if tag == "drawing":
                # solo las imágenes EN LÍNEA van aquí; las flotantes (wp:anchor)
                # las emite el párrafo como bloque aparte (no dentro de su caja)
                if child.find(f"{{{WP}}}inline") is not None:
                    img = self._render_drawing(child)
                    if img:
                        images.append(img)
            elif tag == "t":
                chunks.append(keep_spaces(child.text or ""))
            elif tag == "tab":
                chunks.append("    ")
            elif tag == "cr":
                chunks.append("<br>")
            elif tag == "br":
                if child.get(w("type")) != "page":  # el salto de página
                    chunks.append("<br>")            # se gestiona en el párrafo
        text = "".join(chunks)
        out = ""
        if text:
            css = run_css(d)
            out = f'<span style="{css}">{text}</span>' if css else text
        return out + "".join(images)

    def _data_uri(self, target):
        if target not in self._img_cache:
            import base64
            ext = target.rsplit(".", 1)[-1].lower()
            mime = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png", "gif": "gif",
                    "bmp": "bmp", "svg": "svg+xml"}.get(ext, ext)
            data = base64.b64encode(self.z.read("word/" + target)).decode()
            self._img_cache[target] = f"data:image/{mime};base64,{data}"
        return self._img_cache[target]

    def _render_drawing(self, drawing):
        blip = drawing.find(".//" + f"{{{A}}}blip")
        if blip is None:
            return ""
        target = self.rels.get(blip.get(f"{{{R}}}embed"))
        if not target:
            return ""
        ext = drawing.find(".//" + f"{{{WP}}}extent")  # tamaño en EMU -> pt
        dims = ""
        if ext is not None and ext.get("cx") and ext.get("cy"):
            dims = (f"width:{emu_pt(ext.get('cx')):.1f}pt;"
                    f"height:{emu_pt(ext.get('cy')):.1f}pt;")
        return (f'<img src="{self._data_uri(target)}" '
                f'style="display:block;margin:6pt auto;max-width:100%;{dims}">')

    # -- párrafos ----------------------------------------------------------
    def render_paragraph(self, p, in_cell=False):
        ppr = first(p, "pPr")
        style_id = None
        base = dict(self.default)
        if ppr is not None:
            ps = first(ppr, "pStyle")
            style_id = attr(ps, "val") if ps is not None else None
            if style_id and style_id in self.style_rpr:
                base.update(self.style_rpr[style_id])

        css = []
        is_list = False
        if ppr is not None:
            jc = first(ppr, "jc")
            if jc is not None:
                m = {"both": "justify", "center": "center", "right": "right", "left": "left"}
                css.append("text-align:" + m.get(attr(jc, "val"), "left"))
            sp = first(ppr, "spacing")
            if sp is not None:
                if attr(sp, "before") is not None:
                    css.append(f"margin-top:{tw_pt(attr(sp,'before')):.1f}pt")
                if attr(sp, "after") is not None:
                    css.append(f"margin-bottom:{tw_pt(attr(sp,'after')):.1f}pt")
                line = attr(sp, "line")
                if line is not None and attr(sp, "lineRule") in (None, "auto"):
                    css.append(f"line-height:{float(line)/240.0:.2f}")
            ind = first(ppr, "ind")
            if ind is not None:
                if attr(ind, "left"):
                    css.append(f"margin-left:{tw_pt(attr(ind,'left')):.1f}pt")
                if attr(ind, "right"):
                    css.append(f"margin-right:{tw_pt(attr(ind,'right')):.1f}pt")
                if attr(ind, "hanging"):
                    css.append(f"text-indent:-{tw_pt(attr(ind,'hanging')):.1f}pt")
                elif attr(ind, "firstLine"):
                    css.append(f"text-indent:{tw_pt(attr(ind,'firstLine')):.1f}pt")
            pbdr = first(ppr, "pBdr")
            if pbdr is not None:
                for side in ("top", "bottom", "left", "right"):
                    bc = border_css(first(pbdr, side))
                    if bc and bc != "none":
                        css.append(f"border-{side}:{bc}")
                        sp_attr = first(pbdr, side)
                        if attr(sp_attr, "space"):
                            css.append(f"padding-{side}:{float(attr(sp_attr,'space')):.0f}pt")
            is_list = first(ppr, "numPr") is not None

        # tamaño/fuente por defecto del párrafo (para que también afecte a
        # bullets y a la altura de líneas vacías)
        css.append(f"font-family:{font_stack(base.get('font'))}")
        css.append(f"font-size:{base.get('size',10.0):.1f}pt")
        if base.get("color"):
            css.append("color:" + base["color"])
        if base.get("bold"):
            css.append("font-weight:bold")
        css.append("margin-top:0" if False else "")
        css = [c for c in css if c]

        # salto de página explícito (<w:br w:type="page"/>) dentro del párrafo
        if p.find(".//" + w("br") + "[@" + w("type") + "='page']") is not None:
            css.append("break-after:page")

        inner = self.render_runs(p, base)
        if is_list:
            inner = "– " + inner  # viñeta "–"
        if not inner:
            inner = " "
        para = f'<p style="{";".join(css)}">{inner}</p>'

        # imágenes flotantes (wp:anchor): se difieren para emitirlas como
        # bloque tras el bloque de nivel superior (párrafo o tabla), fuera de
        # cualquier caja/celda — como hace Word, que las saca del contenedor
        for dr in p.iter(w("drawing")):
            if dr.find(f"{{{WP}}}anchor") is not None:
                self._pending_floats.append(self._render_drawing(dr))
        return para

    # -- tablas ------------------------------------------------------------
    def render_table(self, tbl):
        tblpr = first(tbl, "tblPr")
        tblw = first(tblpr, "tblW") if tblpr is not None else None
        style = ["border-collapse:collapse", "table-layout:fixed"]
        if tblw is not None and attr(tblw, "type") == "dxa":
            style.append(f"width:{tw_pt(attr(tblw,'w')):.1f}pt")
        jc = first(tblpr, "jc") if tblpr is not None else None
        if jc is not None and attr(jc, "val") == "center":
            style.append("margin-left:auto")
            style.append("margin-right:auto")
        tblbdr = first(tblpr, "tblBorders") if tblpr is not None else None

        # anchos de columna (layout fijo)
        cols = ""
        grid = first(tbl, "tblGrid")
        if grid is not None:
            cols = "<colgroup>" + "".join(
                f'<col style="width:{tw_pt(attr(gc,"w")):.1f}pt">'
                for gc in grid.findall(w("gridCol"))
            ) + "</colgroup>"

        rows = []
        for tr in tbl.findall(w("tr")):
            cells = []
            for tc in tr.findall(w("tc")):
                cells.append(self._render_cell(tc, tblbdr))
            rows.append("<tr>" + "".join(cells) + "</tr>")
        return f'<table style="{";".join(style)}">{cols}{"".join(rows)}</table>'

    def _render_cell(self, tc, tblbdr):
        tcpr = first(tc, "tcPr")
        css = ["vertical-align:top"]
        colspan = ""
        tcbdr = first(tcpr, "tcBorders") if tcpr is not None else None
        for side in ("top", "bottom", "left", "right"):
            b = first(tcbdr, side) if tcbdr is not None else None
            if b is None and tblbdr is not None:
                b = first(tblbdr, side)
            bc = border_css(b)
            css.append(f"border-{side}:{bc if bc else 'none'}")
        if tcpr is not None:
            shd = first(tcpr, "shd")
            if shd is not None:
                fill = attr(shd, "fill")
                if fill and fill != "auto":
                    css.append(f"background-color:#{fill}")
            mar = first(tcpr, "tcMar")
            if mar is not None:
                for side in ("top", "bottom", "left", "right"):
                    m = first(mar, side)
                    if m is not None and attr(m, "w"):
                        css.append(f"padding-{side}:{tw_pt(attr(m,'w')):.1f}pt")
            else:
                css.append("padding:4pt 6pt")
            va = first(tcpr, "vAlign")
            if va is not None:
                vm = {"center": "middle", "bottom": "bottom"}.get(attr(va, "val"))
                if vm:
                    css[0] = "vertical-align:" + vm
            gs = first(tcpr, "gridSpan")
            if gs is not None:
                colspan = f' colspan="{attr(gs,"val")}"'
        else:
            css.append("padding:4pt 6pt")
        inner = "".join(self.render_paragraph(p, in_cell=True) for p in tc.findall(w("p")))
        return f'<td{colspan} style="{";".join(css)}">{inner}</td>'

    # -- cabecera / pie ----------------------------------------------------
    def _hf_paragraph(self, root, width_cm, is_footer):
        if root is None:
            return ""
        p = root.find(w("p"))
        if p is None:
            return ""
        ppr = first(p, "pPr")
        border = ""
        if ppr is not None:
            pbdr = first(ppr, "pBdr")
            if pbdr is not None:
                side = "top" if is_footer else "bottom"
                bc = border_css(first(pbdr, side))
                if bc and bc != "none":
                    border = f"border-{side}:{bc};padding-{side}:3pt;"
        base = dict(self.default)
        inner = self.render_runs(p, base)
        elem = "ftr" if is_footer else "hdr"
        pagenum = ""
        if is_footer:
            # campo PAGE -> contador de página alineado a la derecha
            inner = re.sub(r" {2,}", " ", inner)
            pagenum = '<span class="pageno" style="float:right"></span>'
        pos = "running(ftr)" if is_footer else "running(hdr)"
        style = (
            f"position:{pos};width:{width_cm:.2f}cm;{border}"
            f"font-family:{font_stack('Calibri')};color:#4a4a4a;"
        )
        return f'<div id="{elem}" style="{style}">{pagenum}{inner}</div>'

    # -- documento completo ------------------------------------------------
    def build_html(self):
        body = self.doc.find(w("body"))
        sect = body.find(w("sectPr"))
        pgsz = first(sect, "pgSz")
        pgmar = first(sect, "pgMar")
        pw = float(attr(pgsz, "w")) if pgsz is not None else 11906
        mt = float(attr(pgmar, "top")) if pgmar is not None else 1440
        mb = float(attr(pgmar, "bottom")) if pgmar is not None else 1440
        ml = float(attr(pgmar, "left")) if pgmar is not None else 1200
        mr = float(attr(pgmar, "right")) if pgmar is not None else 1200
        content_cm = tw_cm(pw - ml - mr)

        blocks = []
        for child in body:
            tag = etree.QName(child).localname
            if tag == "p":
                blocks.append(self.render_paragraph(child))
            elif tag == "tbl":
                blocks.append(self.render_table(child))
            else:
                continue
            if self._pending_floats:  # imágenes flotantes tras el bloque
                blocks.extend(self._pending_floats)
                self._pending_floats = []

        header = self._hf_paragraph(self.header_xml, content_cm, is_footer=False)
        footer = self._hf_paragraph(self.footer_xml, content_cm, is_footer=True)

        page_css = f"""
        @page {{
            size: A4;
            margin: {tw_cm(mt):.2f}cm {tw_cm(mr):.2f}cm {tw_cm(mb):.2f}cm {tw_cm(ml):.2f}cm;
            @top-center {{ content: element(hdr); }}
            @bottom-center {{ content: element(ftr); }}
        }}
        html {{ font-family: Carlito, Calibri, sans-serif; font-size: 10pt;
                color: #4a4a4a; }}
        body {{ margin: 0; }}
        p {{ margin: 0; line-height: {BODY_LINE_HEIGHT}; }}
        table {{ margin: 6pt 0; font-size: 10pt; }}
        td p {{ margin: 0; line-height: {CELL_LINE_HEIGHT}; }}
        .pageno::after {{ content: counter(page); }}
        """
        return (
            "<!DOCTYPE html><html><head><meta charset='utf-8'><style>"
            + page_css + "</style></head><body>"
            + header + footer + "".join(blocks)
            + "</body></html>"
        )


def convert(in_path, out_path):
    """Convierte ``in_path`` (.docx) a ``out_path`` (.pdf). Devuelve out_path."""
    conv = Converter(in_path)
    html = conv.build_html()
    HTML(string=html).write_pdf(out_path)
    return out_path
