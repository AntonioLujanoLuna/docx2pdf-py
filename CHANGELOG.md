# Changelog

Todas las novedades destacables de este proyecto se documentan aquí.
El formato sigue, a grandes rasgos, [Keep a Changelog](https://keepachangelog.com/).

## [Sin publicar]

### Añadido
- Protocolo extensible de motores, políticas explícitas de fallback y API de
  conversión por lotes con concurrencia acotada, cancelación y nombres seguros.
- Diagnósticos por conversión: intentos, errores, tiempos, tamaños y páginas.
- Soporte del flujo propio para overrides de numeración, notas al pie/finales,
  revisiones, cuadros de texto, ecuaciones, columnas y secciones con geometría
  propia, cabeceras de tabla repetibles y filas no divisibles.
- Corpus E2E generado con texto esperado por página y comprobación raster.
- Límites de elementos XML, validación estructural de PDF y terminación de
  árboles de procesos.
- API `convert_detailed()` con el motor realmente usado, avisos de fallback y
  opciones tipadas por conversión mediante `ConversionOptions`.
- Jerarquía pública de excepciones para documentos inválidos, motores no
  disponibles, errores de conversión y timeouts.
- Publicación atómica y validación de PDFs para todos los motores.
- Procesos terminables para WeasyPrint y la automatización de Word en Windows.
- Tests de CLI, empaquetado, timeouts, fallback y preservación de salidas.
- Automatización de dependencias, validación de wheel y publicación a PyPI con
  identidad federada en CI.
- **Resaltado de texto** (`w:highlight`): se reproduce con `background-color`,
  mapeando los colores con nombre de Word (yellow, green, cyan…).
- **Mayúsculas y versalitas** (`w:caps` / `w:smallCaps`) → `text-transform` /
  `font-variant`.
- **Glifos de viñeta**: las listas con viñeta usan el carácter del nivel
  (`lvlText`) mapeado a su equivalente Unicode (Wingdings/Symbol → `•`, `▪`, `✓`…)
  en lugar de un guion fijo.
- **Estilo de párrafo por defecto**: los párrafos sin `pStyle` explícito heredan
  el estilo marcado `w:default="1"` (normalmente *Normal*), como hace Word.
- **Metadatos del PDF**: título, autor, asunto y palabras clave se leen de
  `docProps/core.xml` y se trasladan a los metadatos del PDF.
- **Validación de entrada** en `convert()`: error claro si el archivo no existe
  o no es un ZIP/OOXML válido.
- **Marcador de tipos** `py.typed` para consumidores con *type checkers*.
- **CI**: trabajo de *lint* con ruff y *smoke test* de extremo a extremo
  (`tests/e2e_smoke.py`) que convierte un `.docx` real a PDF con LibreOffice.

### Cambiado
- El renderizado WeasyPrint extrae imágenes a recursos temporales locales para
  evitar inflar el HTML y duplicar memoria mediante base64.
- El conversor monolítico se divide en módulos de API, backends, OOXML,
  formato, procesos y publicación de salida.
- `pyproject.toml`: configuración de ruff y `ruff` añadido a las dependencias de
  desarrollo; metadatos de autoría corregidos.

## [0.1.0]

- Versión inicial: conversión `.docx` → PDF solo con Python (lxml + WeasyPrint),
  con dispatch opcional a Word/LibreOffice para paginación fiel.
