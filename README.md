# docx2pdf-py

Conversión **fiel** de `.docx` a PDF usando **solo librerías de Python** (sin
LibreOffice ni Word). Lee el OOXML del documento (estilos reales: fuentes,
colores, bordes, sombreados, tablas, imágenes, cabecera/pie) y lo recrea como
HTML que **WeasyPrint** pagina a PDF.

```
.docx ──► leer OOXML (lxml) ──► HTML+CSS ──► WeasyPrint ──► PDF
```

## Instalación

```bash
pip install -e .            # desde el repo (modo desarrollo)
# o, una vez publicado:
# pip install docx2pdf-py
```

Dependencias: `weasyprint` y `lxml` (se instalan solas).

### Fuentes (importante para la fidelidad)

Si el documento usa **Calibri**/**Georgia** (no libres), instala sus equivalentes
métricamente compatibles **Carlito** y **Gelasio** en el sistema; WeasyPrint los
descubre vía fontconfig y los **incrusta** en el PDF. Instrucciones en
`requirements.txt`. Otras fuentes se usan si están instaladas.

## Uso

Como librería:

```python
from docx2pdf_py import convert

convert("entrada.docx", "salida.pdf")
```

Como comando:

```bash
docx2pdf-py entrada.docx salida.pdf
docx2pdf-py                      # usa el primer .docx del directorio -> output.pdf
```

## Qué reproduce

Portada, cabecera/pie (la referenciada como `default` en el `sectPr`) con número
de página, encabezados, párrafos con fuente/color/negrita/cursiva/alineación,
listas, tablas (bordes, sombreados, celdas combinadas), saltos de página
explícitos e **imágenes** (inline y flotantes, incrustadas en base64). Los campos
de Word (p. ej. `PAGE`) se interpretan, no se vuelca su valor cacheado.

## Limitaciones (conversor ligero, no un motor Word completo)

- **Listas**: pinta viñeta `–`; no reproduce numeración (`1.`, `a)`, …).
- **Fuentes**: mapea Calibri→Carlito y Georgia→Gelasio; otras caen a `sans-serif`.
- **Tamaño por defecto** 10 pt e **interlineado** ajustados a estilo "ofimático"
  común (configurables vía variables de entorno `BODY_LH` / `CELL_LH`).
- **Imágenes flotantes** se colocan como bloque (no solapan el texto como Word).
- **Cabecera/pie**: solo la `default` (ignora primera página / pares distintos).
- Fidelidad **visual alta**, no *pixel-perfect* (eso exigiría la fuente real y el
  motor de maquetación de Word).

## Estructura

```
docx2pdf_py/
  __init__.py     → expone convert(), Converter
  converter.py    → conversor OOXML -> HTML -> PDF
  cli.py          → comando docx2pdf-py
pyproject.toml    → metadatos y dependencias
main.py           → script de ejemplo (edita la ruta y ejecuta)
```

## Licencia

MIT — ver [LICENSE](LICENSE).
