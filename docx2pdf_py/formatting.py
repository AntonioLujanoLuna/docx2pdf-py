"""Pure numbering, font, run-property, and CSS formatting helpers."""

from __future__ import annotations

from typing import Any

from .ooxml import attr, first, on

FONT_MAP = {
    "Calibri": "Carlito, Calibri, sans-serif",
    "Georgia": "Gelasio, Georgia, serif",
}

GENERIC_FAMILY = {
    "Times New Roman": "serif",
    "Cambria": "serif",
    "Garamond": "serif",
    "Book Antiqua": "serif",
    "Palatino Linotype": "serif",
    "Courier New": "monospace",
    "Consolas": "monospace",
    "Lucida Console": "monospace",
}

HIGHLIGHT_COLORS = {
    "black": "#000000", "blue": "#0000FF", "cyan": "#00FFFF",
    "darkBlue": "#000080", "darkCyan": "#008080", "darkGray": "#808080",
    "darkGreen": "#008000", "darkMagenta": "#800080", "darkRed": "#800000",
    "darkYellow": "#808000", "green": "#00FF00", "lightGray": "#C0C0C0",
    "magenta": "#FF00FF", "red": "#FF0000", "white": "#FFFFFF",
    "yellow": "#FFFF00",
}

BULLET_GLYPHS = {
    "": "•", "": "▪", "": "✓", "": "➢",
    "o": "o", "•": "•", "▪": "▪", "·": "·",
    "–": "–", "−": "–", "*": "•",
}


def _to_letter(n: int) -> str:
    value = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        value = chr(65 + remainder) + value
    return value or "A"


def _to_roman(n: int) -> str:
    if n <= 0:
        return str(n)
    table = (
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"),
        (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"),
        (5, "V"), (4, "IV"), (1, "I"),
    )
    output = ""
    for value, symbol in table:
        while n >= value:
            output += symbol
            n -= value
    return output


def _format_num(n: int, fmt: str) -> str:
    if fmt == "lowerLetter":
        return _to_letter(n).lower()
    if fmt == "upperLetter":
        return _to_letter(n).upper()
    if fmt == "lowerRoman":
        return _to_roman(n).lower()
    if fmt == "upperRoman":
        return _to_roman(n).upper()
    if fmt == "decimalZero":
        return f"{n:02d}"
    return str(n)


def font_stack(name: str | None) -> str | None:
    if not name:
        return None
    if name in FONT_MAP:
        return FONT_MAP[name]
    return f"'{name}', {GENERIC_FAMILY.get(name, 'sans-serif')}"


def rpr_dict(rpr: Any) -> dict[str, Any]:
    properties: dict[str, Any] = {}
    if rpr is None:
        return properties
    fonts = first(rpr, "rFonts")
    if fonts is not None:
        if attr(fonts, "ascii"):
            properties["font"] = attr(fonts, "ascii")
        elif attr(fonts, "asciiTheme"):
            properties["font_theme"] = attr(fonts, "asciiTheme")
            properties["font"] = None
    for key, tag in (("bold", "b"), ("italic", "i"), ("strike", "strike")):
        value = on(first(rpr, tag))
        if value is not None:
            properties[key] = value
    underline = first(rpr, "u")
    if underline is not None:
        properties["underline"] = attr(underline, "val") not in (None, "none")
    color = first(rpr, "color")
    if color is not None and (value := attr(color, "val")) and value != "auto":
        properties["color"] = "#" + value
    size = first(rpr, "sz")
    if size is not None:
        properties["size"] = float(attr(size, "val")) / 2.0
    vertical = first(rpr, "vertAlign")
    if vertical is not None:
        properties["va"] = attr(vertical, "val")
    highlight = first(rpr, "highlight")
    if highlight is not None and (value := attr(highlight, "val")) and value != "none":
        properties["highlight"] = HIGHLIGHT_COLORS.get(value, value)
    for key, tag in (("caps", "caps"), ("smallcaps", "smallCaps")):
        value = on(first(rpr, tag))
        if value is not None:
            properties[key] = value
    return properties


def run_css(properties: dict[str, Any]) -> str:
    css = []
    if properties.get("font"):
        css.append(f"font-family:{font_stack(properties['font'])}")
    if "bold" in properties:
        css.append("font-weight:" + ("bold" if properties["bold"] else "normal"))
    if "italic" in properties:
        css.append("font-style:" + ("italic" if properties["italic"] else "normal"))
    decorations = []
    if properties.get("underline"):
        decorations.append("underline")
    if properties.get("strike"):
        decorations.append("line-through")
    if decorations:
        css.append("text-decoration:" + " ".join(decorations))
    if properties.get("color"):
        css.append("color:" + properties["color"])
    if properties.get("highlight"):
        css.append("background-color:" + properties["highlight"])
    if properties.get("caps"):
        css.append("text-transform:uppercase")
    elif properties.get("smallcaps"):
        css.append("font-variant:small-caps")
    size = properties.get("size")
    vertical = properties.get("va")
    if vertical in ("superscript", "subscript"):
        css.append("vertical-align:" + ("super" if vertical == "superscript" else "sub"))
        if size:
            size *= 0.7
    if size:
        css.append(f"font-size:{size:.1f}pt")
    return ";".join(css)


def border_css(border: Any) -> str | None:
    if border is None:
        return None
    value = attr(border, "val")
    if value in (None, "nil", "none"):
        return "none"
    size = attr(border, "sz")
    width = max(float(size) / 8.0, 0.5) if size else 0.5
    color = attr(border, "color") or "000000"
    if color == "auto":
        color = "000000"
    return f"{width:.2f}pt solid #{color}"
