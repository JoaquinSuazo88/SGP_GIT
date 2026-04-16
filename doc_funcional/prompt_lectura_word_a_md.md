# Prompt — Conversión de Documento Word a Markdown

> **Cómo usar:** modifica los valores de la sección **Parámetros** y entrega el documento completo como prompt al agente.

---

## ▶ PARÁMETROS — modificar antes de usar

```
RUTA_WORD      = C:\Users\Joaquin.Suazo\Documents\SGP-Producción\Doc_Funcionales\Words\DRF_Inventario.docx
RUTA_SALIDA    = C:\Users\Joaquin.Suazo\Documents\SGP-Producción\Doc_Funcionales\MD
NOMBRE_CARPETA = DRF_Inventario
NOMBRE_MD      = DRF_Inventario.md
CARPETA_IMG    = imagenes
```

**Resultado esperado:**
```
{{RUTA_SALIDA}}\{{NOMBRE_CARPETA}}\
    {{NOMBRE_MD}}
    {{CARPETA_IMG}}\
        imagen_01.jpg
        imagen_02.jpg
        ...
```

---

## PROMPT

Convierte el documento Word ubicado en:

```
{{RUTA_WORD}}
```

a Markdown, respetando su estructura original. Guarda el resultado en:

```
{{RUTA_SALIDA}}\{{NOMBRE_CARPETA}}\{{NOMBRE_MD}}
```

---

### Paso 1 — Crear la estructura de directorios

Crea las carpetas necesarias si no existen:

```bash
mkdir -p "{{RUTA_SALIDA}}/{{NOMBRE_CARPETA}}/{{CARPETA_IMG}}"
```

---

### Paso 2 — Extraer las imágenes del documento Word

Un archivo `.docx` es un ZIP. Las imágenes se encuentran en la carpeta `word/media/` dentro del ZIP.

```bash
# Extraer todas las imágenes del DOCX al directorio temporal
unzip -o "{{RUTA_WORD}}" "word/media/*" -d /tmp/docx_extract/
```

Luego convierte cada imagen extraída a formato JPG y nómbralas correlativamente (`imagen_01.jpg`, `imagen_02.jpg`, etc.). Usa Python con Pillow:

```python
import os
import glob
from PIL import Image

media_dir = "/tmp/docx_extract/word/media/"
out_dir   = "{{RUTA_SALIDA}}/{{NOMBRE_CARPETA}}/{{CARPETA_IMG}}/"

archivos = sorted(glob.glob(os.path.join(media_dir, "*")))
for i, src in enumerate(archivos, start=1):
    try:
        img = Image.open(src).convert("RGB")
        dst = os.path.join(out_dir, f"imagen_{i:02d}.jpg")
        img.save(dst, "JPEG", quality=90)
        print(f"Guardada: {dst}")
    except Exception as e:
        print(f"No se pudo convertir {src}: {e}")
```

Si Pillow no está instalado: `pip install pillow`

Si `unzip` no está disponible en Windows, usa Python directamente:

```python
import zipfile, os, glob
from PIL import Image

docx_path = r"{{RUTA_WORD}}"
out_dir   = r"{{RUTA_SALIDA}}\{{NOMBRE_CARPETA}}\{{CARPETA_IMG}}"
tmp_dir   = r"/tmp/docx_media"
os.makedirs(tmp_dir, exist_ok=True)
os.makedirs(out_dir, exist_ok=True)

with zipfile.ZipFile(docx_path, "r") as z:
    media_files = [f for f in z.namelist() if f.startswith("word/media/")]
    for f in media_files:
        z.extract(f, tmp_dir)

archivos = sorted(glob.glob(os.path.join(tmp_dir, "word", "media", "*")))
for i, src in enumerate(archivos, start=1):
    try:
        img = Image.open(src).convert("RGB")
        dst = os.path.join(out_dir, f"imagen_{i:02d}.jpg")
        img.save(dst, "JPEG", quality=90)
        print(f"Guardada: {dst}")
    except Exception as e:
        print(f"No se pudo convertir {src}: {e}")
```

Anota la correspondencia entre el nombre original de cada archivo en `word/media/` y el nombre JPG asignado (`imagen_01.jpg`, etc.). La necesitarás en el Paso 4 para insertar las imágenes en el lugar correcto del Markdown.

---

### Paso 3 — Extraer el contenido del documento

Usa Python con `python-docx` para leer el contenido completo:

```python
from docx import Document

doc = Document(r"{{RUTA_WORD}}")

# Recorrer todos los elementos del cuerpo del documento
for bloque in doc.element.body:
    print(bloque.tag, bloque.text if hasattr(bloque, 'text') else '')
```

Si `python-docx` no está instalado: `pip install python-docx`

**Antes de extraer el contenido, identifica y omite:**
- La **portada** (primera página): todo el contenido hasta el primer salto de sección (`<w:sectPr>` embebido en un `<w:pPr>`), o hasta el primer encabezado de contenido real si no hay salto de sección explícito.
- La **tabla de contenido automática de Word** (TOC): bloques `<w:sdt>` que contienen un campo `TOC`, o párrafos con estilos `TOC 1`, `TOC 2`, `TOC 3`, etc.

Extrae los siguientes elementos **en el orden exacto en que aparecen en el documento**:

#### 3a. Párrafos y encabezados

```python
from docx import Document
from docx.oxml.ns import qn

doc = Document(r"{{RUTA_WORD}}")
for para in doc.paragraphs:
    estilo = para.style.name   # Heading 1, Heading 2, Normal, List Bullet, etc.
    texto  = para.text.strip()
    print(f"[{estilo}] {texto}")
```

Mapea los estilos Word a Markdown:

| Estilo Word | Markdown |
|---|---|
| `Heading 1` | `# Texto` |
| `Heading 2` | `## Texto` |
| `Heading 3` | `### Texto` |
| `Heading 4` | `#### Texto` |
| `Heading 5` | `##### Texto` |
| `Normal` | Párrafo normal |
| `List Bullet` / `List Bullet 2` | `- Texto` / `  - Texto` |
| `List Number` / `List Number 2` | `1. Texto` / `   1. Texto` |
| `Quote` / `Block Text` | `> Texto` |
| `Code` / `HTML Code` | `` `Texto` `` o bloque de código |

Formatos de texto en línea:

| Formato Word (en `run`) | Markdown |
|---|---|
| Bold (`run.bold`) | `**texto**` |
| Italic (`run.italic`) | `*texto*` |
| Bold + Italic | `***texto***` |
| Underline (`run.underline`) | `<u>texto</u>` |
| Strikethrough | `~~texto~~` |
| Hyperlink | `[texto](url)` |

#### 3b. Tablas

```python
for tabla in doc.tables:
    for i, fila in enumerate(tabla.rows):
        celdas = [c.text.strip() for c in fila.cells]
        print(" | ".join(celdas))
        if i == 0:
            print(" | ".join(["---"] * len(celdas)))  # separador de encabezado
```

#### 3c. Imágenes (posición en el documento)

Las imágenes en el XML del documento se referencian mediante el atributo `r:embed` que apunta a un `rId` en el archivo de relaciones (`word/_rels/document.xml.rels`). Para identificar la posición de cada imagen en el flujo del documento:

```python
import zipfile, xml.etree.ElementTree as ET

nsmap = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

with zipfile.ZipFile(r"{{RUTA_WORD}}", "r") as z:
    rels_xml = z.read("word/_rels/document.xml.rels")
    doc_xml  = z.read("word/document.xml")

rels_root = ET.fromstring(rels_xml)
rels = {}
for rel in rels_root:
    rels[rel.attrib["Id"]] = rel.attrib.get("Target", "")

doc_root = ET.fromstring(doc_xml)
for drawing in doc_root.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing"):
    blip = drawing.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
    if blip is not None:
        rid    = blip.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        target = rels.get(rid, "")
        nombre_original = os.path.basename(target)
        print(f"Imagen encontrada: rId={rid}, archivo={nombre_original}")
```

Con esta información, cruza `nombre_original` con la correspondencia generada en el Paso 2 para saber qué `imagen_XX.jpg` insertar en cada posición.

#### 3d. Comentarios del documento

Los comentarios de revisión se almacenan en `word/comments.xml` dentro del DOCX:

```python
with zipfile.ZipFile(r"{{RUTA_WORD}}", "r") as z:
    if "word/comments.xml" in z.namelist():
        comments_xml = z.read("word/comments.xml")
        root = ET.fromstring(comments_xml)
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        for comment in root.findall(f"{{{ns}}}comment"):
            cid    = comment.attrib.get(f"{{{ns}}}id")
            autor  = comment.attrib.get(f"{{{ns}}}author", "")
            fecha  = comment.attrib.get(f"{{{ns}}}date", "")
            texto  = " ".join(p.text or "" for p in comment.iter(f"{{{ns}}}t"))
            print(f"[Comentario #{cid}] {autor} ({fecha}): {texto}")
```

Para obtener el texto al que se refiere cada comentario, busca en `word/document.xml` las etiquetas `<w:commentRangeStart w:id="N"/>` y `<w:commentRangeEnd w:id="N"/>` que delimitan el texto comentado.

#### 3e. Elementos a excluir

##### Portada (primera página)

En Word, la portada ocupa una sección propia que termina con `<w:sectPr>` embebido dentro del `<w:pPr>` de su último párrafo. Al iterar `doc.element.body`, marca como "portada excluida" todo lo que aparece antes de ese salto de sección:

```python
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def tiene_section_break(para_elem):
    """Detecta el salto de sección que cierra la portada."""
    pPr = para_elem.find(f'{{{W_NS}}}pPr')
    if pPr is not None:
        if pPr.find(f'{{{W_NS}}}sectPr') is not None:
            return True
    return False

# Al iterar el body:
en_portada = True
for bloque in doc.element.body:
    if en_portada:
        if bloque.tag.endswith('}p') and tiene_section_break(bloque):
            en_portada = False  # a partir del SIGUIENTE bloque, es contenido real
        continue  # saltar todo lo que está en la portada
    # ... procesar bloque normalmente
```

Si el documento no tiene sección de portada explícita, usa como heurística saltar hasta el primer párrafo con estilo `Heading 1` o `Heading 2` que contenga texto.

##### Tabla de contenido automática de Word (TOC)

La TOC de Word se almacena como un `<w:sdt>` (Structured Document Tag) o como párrafos con estilos `TOC 1`, `TOC 2`, `TOC 3`. Omítelos:

```python
def es_toc(bloque):
    """Detecta si un bloque es la tabla de contenido de Word."""
    # Caso 1: bloque w:sdt con campo TOC
    if bloque.tag.endswith('}sdt'):
        for instr in bloque.iter(f'{{{W_NS}}}instrText'):
            if 'TOC' in (instr.text or ''):
                return True
        sdtPr = bloque.find(f'{{{W_NS}}}sdtPr')
        if sdtPr is not None:
            tag_elem = sdtPr.find(f'{{{W_NS}}}tag')
            if tag_elem is not None:
                val = tag_elem.get(f'{{{W_NS}}}val', '').lower()
                if 'toc' in val:
                    return True
    # Caso 2: párrafo con estilo TOC N
    if bloque.tag.endswith('}p'):
        pPr = bloque.find(f'{{{W_NS}}}pPr')
        if pPr is not None:
            pStyle = pPr.find(f'{{{W_NS}}}pStyle')
            if pStyle is not None:
                val = pStyle.get(f'{{{W_NS}}}val', '').lower()
                if val.startswith('toc') or val == 'tableofcontents':
                    return True
    return False
```

---

### Paso 4 — Generar el Markdown

Recorre el cuerpo del documento **en orden de aparición** (párrafos, tablas e imágenes intercaladas) y construye el Markdown de la siguiente forma:

#### Reglas de construcción

1. **Encabezados:** usa el nivel correspondiente según la tabla del Paso 3a.

2. **Párrafos normales:** escribe el texto con sus formatos en línea (bold, italic, etc.). Separa párrafos consecutivos con una línea en blanco.

3. **Listas:** respeta la jerarquía (bullet anidado = `  -`). Si hay una lista numerada, usa `1.`, `2.`, etc.

4. **Tablas:** genera la tabla Markdown con separador `---` en la segunda fila (encabezado). Si las celdas tienen texto con formato, aplica el formato en línea correspondiente.

5. **Imágenes:** en el punto exacto donde aparece la imagen en el documento, inserta:
   ```
   ![Imagen N]({{CARPETA_IMG}}/imagen_NN.jpg)
   ```
   Reemplaza `N` por el número correlativo. Si la imagen tiene un título o leyenda en el documento (párrafo inmediatamente siguiente con estilo `Caption`), úsalo como texto alternativo:
   ```
   ![Leyenda de la imagen]({{CARPETA_IMG}}/imagen_NN.jpg)
   *Leyenda de la imagen*
   ```

6. **Comentarios:** los comentarios de revisión se insertan en el Markdown como citas de bloque (`>`) inmediatamente después del párrafo al que pertenecen, con este formato:
   ```
   > 💬 **Comentario — <Autor> (<fecha>):** <texto del comentario>
   ```
   Si el comentario está asociado a un fragmento de texto específico (no a un párrafo completo), inserta el bloque después del párrafo que contiene ese fragmento.

7. **Líneas horizontales:** si el documento tiene separadores de sección (párrafos con borde inferior o estilo `Horizontal Line`), usa `---`.

8. **Bloques de código:** si hay párrafos con estilo `Code` o texto con fuente monoespaciada, envuélvelos en triple backtick.

9. **Numeración automática en encabezados:** muchos documentos Word usan numeración automática vía `numPr` (listas numeradas aplicadas a Headings). En ese caso el número **no está en el texto** del párrafo — lo genera Word desde `word/numbering.xml`. Detéctalo y resuélvelo:

   ```python
   import zipfile, xml.etree.ElementTree as ET
   from collections import defaultdict

   W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

   def obtener_numpr(para_elem):
       """Retorna (numId, ilvl) si el párrafo tiene numeración automática, o (None, None)."""
       pPr = para_elem.find(f'{{{W_NS}}}pPr')
       if pPr is None:
           return None, None
       numPr = pPr.find(f'{{{W_NS}}}numPr')
       if numPr is None:
           return None, None
       ilvl_e  = numPr.find(f'{{{W_NS}}}ilvl')
       numId_e = numPr.find(f'{{{W_NS}}}numId')
       ilvl  = int(ilvl_e.get(f'{{{W_NS}}}val',  0)) if ilvl_e  is not None else 0
       numId = int(numId_e.get(f'{{{W_NS}}}val', 0)) if numId_e is not None else 0
       return numId, ilvl

   # Contadores por (numId, ilvl)
   contadores = defaultdict(int)
   contadores_hijos = defaultdict(lambda: defaultdict(int))

   def generar_numero(numId, ilvl):
       """Genera el número correlativo para el nivel dado y reinicia niveles hijos."""
       contadores[(numId, ilvl)] += 1
       # Reiniciar niveles inferiores
       for lvl in range(ilvl + 1, 10):
           contadores[(numId, lvl)] = 0
       # Construir string: "1.", "1.1.", "1.1.1."
       partes = [str(contadores[(numId, l)]) for l in range(ilvl + 1) if contadores[(numId, l)] > 0]
       return ".".join(partes) + "."
   ```

   Al procesar un encabezado:
   - Si el texto ya comienza con un número (ej. `"1. Título"`), **no agregues otro** — el redactor lo escribió manualmente.
   - Si el texto no tiene número pero el párrafo tiene `numPr`, genera el número con `generar_numero()` y preponlo: `"1. " + texto`.
   - Usa ese mismo texto (con número) tanto en el encabezado del MD como en la entrada del índice.

   > ⚠️ **OBLIGATORIO:** Si el documento tiene encabezados con `numPr`, TODOS deben aparecer numerados en el MD y en el índice. Verifica antes de escribir el archivo que los primeros 5 encabezados del MD incluyen su número. Si no lo tienen, el proceso está incompleto.

#### Estructura del archivo generado

```markdown
# Título principal (del documento Word)

<!-- Metadatos opcionales si el documento los tiene -->
<!-- Autor: ... | Fecha: ... | Versión: ... -->

---

[contenido del documento en orden]
```

No agregar secciones, títulos ni texto que no estén en el documento original. El Markdown debe ser una representación fiel del Word, no una reinterpretación.

---

### Paso 5 — Generar e insertar el índice de navegación

Una vez generado el Markdown completo, construye un índice con todos los encabezados del documento e insértalo después de la portada y antes de la primera sección de contenido.

#### 5a. Identificar la portada y el punto de inserción

La portada es el bloque inicial del MD que contiene el título principal (`#`), subtítulos (`##`) y metadatos (autor, fecha, versión) **sin número de sección**. El índice se inserta justo antes del primer encabezado numerado o del primer encabezado de contenido real (por ejemplo, `# 1. Resumen Ejecutivo` o `# Introducción`).

#### 5b. Construir el índice

Recorre todos los encabezados del MD generado y construye la lista de navegación. Los anchors se generan así:
1. Tomar el texto del encabezado (sin los `#`).
2. Convertir a minúsculas.
3. Reemplazar espacios por `-`.
4. Eliminar caracteres que no sean letras, números, `-` ni letras acentuadas (eliminar `.`, `,`, `(`, `)`, `¿`, `?`, `!`, etc.).

```python
import re

def texto_a_anchor(texto):
    texto = texto.lower().strip()
    # Eliminar caracteres no permitidos (conservar letras, números, espacios, guiones, acentos)
    texto = re.sub(r"[^\w\s\-áéíóúüñàèìòùâêîôû]", "", texto, flags=re.UNICODE)
    # Reemplazar espacios por guiones
    texto = re.sub(r"\s+", "-", texto)
    return texto

# Ejemplo:
# "3.1 Componente PLANIFICACIÓN REAL" -> "31-componente-planificación-real"
# "2.3 Problemas Identificados"       -> "23-problemas-identificados"
```

El nivel de indentación en el índice sigue el nivel del encabezado:
- `#` (nivel 1) → sin sangría: `- [Texto](#anchor)`
- `##` (nivel 2) → 2 espacios: `  - [Texto](#anchor)`
- `###` (nivel 3) → 4 espacios: `    - [Texto](#anchor)`
- `####` (nivel 4) → 6 espacios: `      - [Texto](#anchor)`

**Excluir del índice** los encabezados de portada (los que no tienen número de sección y están antes del primer encabezado numerado) y el propio `## Índice`.

#### 5c. Insertar el índice en el MD

El bloque del índice tiene este formato:

```markdown
---

## Índice

- [Sección 1](#anchor-1)
  - [Subsección 1.1](#anchor-11)
  - ...
- [Sección 2](#anchor-2)
  ...

---
```

Se inserta entre la portada y el primer encabezado de contenido. Usa la herramienta Edit para reemplazar la transición exacta entre el último elemento de la portada y el primer encabezado de contenido.

---

### Paso 6 — Escribir el archivo final

El archivo debe comenzar con el título del documento como encabezado `#` de primer nivel. Usa el nombre del MD sin extensión, aplicando estas dos transformaciones en orden: primero reemplaza `-` por espacio, luego reemplaza `_` por ` - ` (espacio-guion-espacio) (por ejemplo, `DRF_Minutas-y-Recetas.md` → `DRF_Minutas y Recetas` → `# DRF - Minutas y Recetas`). A continuación viene el separador `---`, luego el índice y luego el contenido.

Escribe el Markdown generado (con título e índice incluidos) en:

```
{{RUTA_SALIDA}}\{{NOMBRE_CARPETA}}\{{NOMBRE_MD}}
```

Verifica al final:
- [ ] El número de imágenes en `{{CARPETA_IMG}}/` coincide con el número de imágenes en el documento Word.
- [ ] Cada imagen está referenciada en el MD con la ruta relativa correcta (`{{CARPETA_IMG}}/imagen_NN.jpg`).
- [ ] Los comentarios del documento están incluidos como bloques `>` en el MD.
- [ ] La jerarquía de encabezados del MD refleja fielmente la del Word.
- [ ] Las tablas tienen el separador de encabezado `---` en la segunda fila.
- [ ] El archivo inicia con `# <nombre del documento>` como título de primer nivel.
- [ ] El índice está presente después del título y antes del primer encabezado de contenido.
- [ ] Cada enlace del índice tiene el anchor correcto (verificar al menos los primeros 5).

---

### Notas sobre compatibilidad

- **Imágenes EMF/WMF** (metafiles de Windows): Pillow puede no soportarlas directamente. En ese caso, omite la conversión e indica en el MD `![Imagen no convertida — formato EMF/WMF]({{CARPETA_IMG}}/imagen_NN_original.emf)` y copia el archivo original sin convertir.
- **Objetos OLE embebidos** (Excel, Visio, etc.): no son imágenes estándar. Documenta su presencia con `> ⚠️ Objeto embebido no exportable (tipo: OLE)` en la posición correspondiente.
- **Texto en cuadros de texto / formas** (`<w:txbxContent>`): extraer su contenido y agregarlo al flujo del MD en la posición donde aparece el cuadro, precedido por una línea `---` si el cuadro tiene borde visible. **Importante:** python-docx puede exponer el mismo texto tanto al iterar el body (en la forma/cuadro) como en `doc.paragraphs` — lo que produce duplicados. Para evitarlo, registra los textos ya emitidos desde cuadros y omite cualquier párrafo posterior con el mismo texto exacto:

  ```python
  textos_cuadros = set()

  # Al emitir texto de un cuadro de texto:
  texto_cuadro = "...texto del cuadro..."
  textos_cuadros.add(texto_cuadro.strip())
  md_lines.append(texto_cuadro)

  # Al procesar párrafos normales, saltear si ya fue emitido como cuadro:
  if para.text.strip() in textos_cuadros:
      continue
  ```
- **Encabezados y pies de página**: si el documento tiene encabezado o pie relevante (ej. número de versión, nombre del documento), agrégalos al inicio del MD como comentario HTML: `<!-- Encabezado: ... -->`.
