# Prompt — Documentación de Reportes SGP Admin (Orientada a Lectura de Usuario)

> **Cómo usar:** modifica los valores de la sección **Parámetros** y entrega el documento completo como prompt al agente.
>
> **Diferencia con el prompt base:** este prompt genera un MD estructurado según la secuencia lógica en que un usuario opera la pantalla — de lo general a lo concreto, de la pregunta "¿para qué sirve?" hasta "¿qué obtengo?". Los pasos de lectura y extracción del código son idénticos al prompt base; solo cambia la estructura y el orden de las secciones del MD resultante.

---

## ▶ PARÁMETROS — modificar antes de usar

```
BASE_PROYECTO  = C:\Users\Joaquin.Suazo\Documents\SGP-Producción
RUTA_FUENTE    = {{BASE_PROYECTO}}\codigo_fuente\SGP_Admin
RUTA_SQL       = {{BASE_PROYECTO}}\base_de_datos\SGP_Admin.sql
RUTA_DOC       = {{BASE_PROYECTO}}\doc_funcional

FORMULARIO     = E_ExcelVarios.frm
NOMBRE_MD      = Exportacion_Excel_Varios.md
```

---

## PROMPT

Analiza el formulario VB6 ubicado en:

```
{{RUTA_FUENTE}}\{{FORMULARIO}}
```

y genera un documento Markdown funcional. Escríbelo en:

```
{{RUTA_DOC}}\md_pantallas\SGP_Admin\{{NOMBRE_MD}}
```

---

### Contexto del sistema

El sistema **SGP Admin** es el módulo administrativo centralizado de SGP, que gestiona la configuración, reportería y supervisión de los casinos Sodexo Chile. Está desarrollado en **Visual Basic 6** con base de datos **SQL Server**. Los formularios de tipo reporte (`I_`, `L_`, `R_`) permiten consultar, exportar e imprimir información operativa consolidada de uno o varios casinos.

Los documentos Markdown son de uso funcional: los leen analistas, jefes de casino, coordinadores de zona y administradores corporativos que no saben programación. El documento debe estar ordenado según la secuencia lógica en que el usuario interactúa con la pantalla: primero entiende para qué sirve, luego qué necesita preparar, luego cómo opera, luego qué restricciones existen y finalmente qué obtiene. Los detalles técnicos van al final como referencia.

---

### Paso 1 — Lee el formulario VB6

Lee el archivo completo. Extrae exclusivamente lo que está en el código — no inferir propósito, audiencia ni contexto de negocio más allá de lo que el formulario explicita.

- **Caption del formulario:** el título exacto que aparece en la barra del formulario (`Me.Caption` o la propiedad `Caption` del `.frm`).
- **Tablas principales:** las que aparecen en FROM, INTO, UPDATE, DELETE (nombre exacto).
- **SPs o funciones SQL:** referencias a `.Execute`, llamadas con prefijo `sgp_`, `sgpadm_` o `fg_Cal`. Para formularios con múltiples tipos de informe, anota qué SP se llama en cada `Case`.
- **Controles de la pantalla:** lista todos los controles visibles con su `Caption` o etiqueta real tal como aparece en el código:
  - Campos de filtro (textboxes, campos numéricos, campos de fecha) con su etiqueta Label.
  - Listas desplegables (combo) con sus `AddItem` — anota cada opción con su texto exacto y código interno si lo tiene.
  - Checkboxes y opciones (OptionButton) con su `Caption`.
  - Frames con su `Caption` y qué controles contienen.
  - Grillas (vaSpread / MSFlexGrid) y TreeView con su propósito observable (qué datos cargan).
  - Botones del Toolbar con su `ToolTipText` exacto.
  - Barra de progreso (ProgressBar) si existe.
- **Comportamiento condicional de controles:** qué paneles o botones se habilitan/deshabilitan según la opción seleccionada en el combo (busca los bloques `Select Case` en el evento del combo).
- **Formato de salida por tipo:** para cada tipo de informe (o para el único si no hay selector), determina si genera Excel (`CreateObject("Excel.Application")`), RTF con vista previa (`VSPrinter` / `Preview.VSPrinter`), RTF directo, o impresora. Anota la orientación del documento RTF si está presente (`.Orientation = orPortrait` / `orLandscape`).
- **Funciones de exportación:** si la generación del informe se delega a una función externa (en otro `.bas` o módulo), anota el nombre de esa función y en qué archivo está definida.
- **Validaciones:** bloques `If` con `MsgBox`, condiciones antes de ejecutar (fechas vacías, rangos excesivos, casinos sin datos, etc.). Anota el texto exacto del mensaje.
- **Flujo principal:** pasos que realiza el usuario desde que abre el formulario hasta obtener el resultado.

---

### Paso 2 — Busca SPs y funciones en el SQL

Si encontraste SPs o funciones en el paso anterior, conviértelos y búscalos:

```bash
iconv -f UTF-16 -t UTF-8 "{{RUTA_SQL}}" > /tmp/SGP_Admin_utf8.sql

grep -n "nombre_sp_o_funcion" /tmp/SGP_Admin_utf8.sql
```

Por cada SP o función encontrado, lee el bloque completo y extrae:
- Qué tablas consulta.
- Qué parámetros recibe y cuáles son opcionales (valor 0 = sin filtro).
- Qué lógica aplica (filtros, agrupaciones, cálculos).
- Qué columnas devuelve y qué representa cada una en el negocio.

Si no hay SPs (operaciones SQL inline), documenta igualmente las consultas principales encontradas en el VB6.

---

### Paso 3 — Redacta el MD con esta estructura orientada al usuario

La estructura del documento sigue el orden natural en que un usuario se aproxima a la pantalla: primero entiende de qué se trata, luego qué debe preparar, luego cómo opera, luego qué restricciones hay, luego qué obtiene, y al final consulta los detalles técnicos si los necesita.

---

#### Encabezado

```
# <Nombre funcional del reporte>

**Formulario:** `<nombre>.frm`
**Tabla(s) principal(es):** `<tabla1>` (<descripción breve en español>)
**Consulta principal:** `<nombre_sp>` — o bien — Sin procedimiento almacenado: consulta directa al servidor

---
```

---

#### Índice de navegación

Inmediatamente después del encabezado y antes de la Sección 1, genera un índice con enlaces internos a todas las secciones y subsecciones del documento. Usa listas anidadas con este formato:

```
## Índice

- [1 — ¿Para qué sirve esta pantalla?](#1--para-qué-sirve-esta-pantalla)
- [2 — ¿Qué necesito para usarla?](#2--qué-necesito-para-usarla)
- [3 — ¿Cómo se usa?](#3--cómo-se-usa)
  - [3.1 Flujo paso a paso](#31-flujo-paso-a-paso)
  - [3.2 Controles y acciones disponibles](#32-controles-y-acciones-disponibles)
- [4 — ¿Qué restricciones debo conocer?](#4--qué-restricciones-debo-conocer)
  - [4.1 Validaciones del sistema](#41-validaciones-del-sistema)
  - [4.2 Reglas de cálculo](#42-reglas-de-cálculo)
- [5 — ¿Qué obtengo?](#5--qué-obtengo)
  - [Resumen de tipos disponibles](#resumen-de-tipos-disponibles)
  - [(<código>) <Nombre>](#<anchor-del-tipo>)
  - …
- [6 — Referencia técnica](#6--referencia-técnica)
  - [Tablas que intervienen](#tablas-que-intervienen)
  - [Relación con otros módulos](#relación-con-otros-módulos)
```

**Reglas para generar los anchors de los tipos de informe:**

Los anchors de los subtítulos `### (<código>) <Nombre> (\`<función>\`)` se generan aplicando las siguientes transformaciones al texto completo del encabezado:
1. Convertir todo a minúsculas.
2. Eliminar los caracteres que no sean letras, números, espacios o guiones (eliminar paréntesis, backticks, `/`, `?`, `¿`, `—`, etc.).
3. Reemplazar cada espacio (o grupo de espacios consecutivos) por un guion `-`.

Ejemplo: `### (06) Menú Mensual (Formato Comercial) (\`ExportarExcelMenuMensualMKT\`)`
→ texto limpio: `06 menú mensual formato comercial exportarexcelmenumensualmkt`
→ anchor: `#06-menú-mensual-formato-comercial-exportarexcelmenumensualmkt`

Si el tipo tiene dos funciones separadas por ` / `, el `/` y los espacios alrededor producen un doble guion `--` en el anchor.

Ejemplo: `### (01) Menú Mensual (\`FuncA\` / \`FuncB\`)`
→ anchor: `#01-menú-mensual-funca--funcb`

Verifica que cada anchor del índice coincida exactamente con el generado por el encabezado correspondiente antes de escribir el MD.

**Links de retorno al índice:**

Inmediatamente después de cada título que aparece en el índice (secciones principales y subsecciones), agrega en la línea siguiente el link:

```
[↑ Volver al índice](#índice)
```

Esto aplica a todos los niveles: `##`, `###`. No agregarlo en el propio `## Índice` ni en el `# <título principal>` del documento.

---

#### Sección 1 — ¿Para qué sirve esta pantalla?

Escribe 2–3 párrafos que respondan exactamente esa pregunta, basándose solo en lo que el código muestra:
- Qué información entrega o qué acción permite realizar.
- Cómo se organiza visualmente la pantalla (panel de filtros, selector de tipo, árbol de servicios, barra de progreso, etc.).
- Si consolida datos de un casino o de múltiples, y si hay diferencia entre las opciones disponibles.

No inferir audiencia ni propósito de negocio más allá de lo que el formulario muestra.

---

#### Sección 2 — ¿Qué necesito para usarla?

Tabla con los filtros o parámetros que el usuario debe completar antes de poder ejecutar el reporte. Incluye todos los campos de cabecera y el selector de tipo si existe:

```
| Campo | Descripción | Obligatorio |
|---|---|---|
| <nombre del campo> | <qué representa y para qué sirve al operar la pantalla> | Sí / No |
```

Si algún campo tiene un buscador o abre un formulario auxiliar, indícalo en la columna Descripción.
Si el formulario carga datos automáticamente al abrirse sin requerir nada del usuario, indícalo aquí.

---

#### Sección 3 — ¿Cómo se usa?

##### 3.1 Flujo paso a paso

Genera el diagrama usando **Mermaid** (`flowchart TD`). El diagrama debe mostrar el proceso completo desde la perspectiva del usuario: qué ingresa, qué selecciona, qué ejecuta, qué valida el sistema y cómo obtiene el resultado.

Usa estas convenciones de forma:

| Tipo de nodo | Forma Mermaid | Cuándo usarla |
|---|---|---|
| Acción del usuario | `[Texto]` — rectángulo | El usuario selecciona un filtro, presiona un botón o ingresa datos |
| Acción del sistema | `(Texto)` — rectángulo redondeado | El sistema carga datos, valida, calcula o genera el documento |
| Decisión / validación | `{Texto}` — rombo | El sistema evalúa una condición antes de continuar |
| Consulta a base de datos | `[(Texto)]` — cilindro | Se ejecuta una consulta o procedimiento almacenado |
| Inicio / Fin | `([Texto])` — estadio | Punto de entrada o salida del flujo |

**Reglas para el diagrama:**
- Incluir siempre las ramas de validación con su rombo y mensajes reales.
- Los nodos de base de datos deben nombrar la tabla o SP real.
- Si hay múltiples formatos de salida (Excel, RTF, imprimir), mostrar cada uno como rama separada.
- No incluir nombres de funciones VB6 ni eventos internos.
- Si el formulario tiene tipos de informe con flujos distintos (ej. RTF vs Excel), mostrar ambas ramas desde el punto de decisión.

##### 3.2 Controles y acciones disponibles

Lista cada acción disponible en el formulario:

```
| Control / Acción | Descripción |
|---|---|
| **<Nombre exacto del botón o panel>** | <Qué hace cuando el usuario lo usa. Indicar qué habilita, qué carga o qué genera.> |
```

Ordena la tabla según el orden en que el usuario normalmente los utiliza: primero los filtros de búsqueda/carga, luego los selectores de configuración, luego los botones de ejecución y exportación, al final navegación y cierre.

---

#### Sección 4 — ¿Qué restricciones debo conocer?

##### 4.1 Validaciones del sistema

Cada restricción que el sistema impone, en el orden en que el usuario las encontraría:

```
| # | Cuándo aparece | Qué verifica el sistema | Qué ve o experimenta el usuario |
|---|---|---|---|
| 1 | Al ingresar el CECO | <condición> | <mensaje exacto o comportamiento> |
```

##### 4.2 Reglas de cálculo del formulario

Esta subsección existe **solo si hay cálculos que ocurren a nivel del formulario principal**, independientemente del tipo de informe seleccionado (por ejemplo: valores que se calculan al cargar los filtros, al seleccionar un servicio, o que se comparten entre todos los tipos). Si no existen ese tipo de cálculos, omite esta subsección completamente.

Los cálculos propios de cada tipo de informe se documentan dentro del subtítulo de ese tipo en la Sección 5, no aquí.

Para cada cálculo del formulario principal, usa este formato:

```
**<Nombre funcional del valor calculado>**

<Explicación en lenguaje simple: qué representa este valor y por qué el sistema lo calcula en lugar de leerlo directamente.>

**Fórmula o lógica:**
<Nombre resultado> = <Componente A> × <Componente B> …

| Componente | Qué representa | De dónde viene |
|---|---|---|
| <Componente A> | <descripción> | <tabla y campo, o "ingresado por el usuario"> |

> Ejemplo: <valores ficticios que ilustren el cálculo paso a paso>
```

---

#### Sección 5 — ¿Qué obtengo?

Esta sección describe exactamente qué información entrega el reporte al usuario. Si el formulario tiene un selector de tipo de informe, **cada tipo es un subtítulo propio**. Si es un reporte único, documenta directamente su contenido.

##### Paso previo — Identificar los tipos de informe

Lee el bloque `Combo1.AddItem` en el formulario. Cada línea tiene este patrón:

```
Combo1(0).AddItem "<Nombre visible>" & Space(150) & "(<código>)"
```

El texto antes del `Space(150)` es el nombre exacto que ve el usuario en el selector. **Usa esos nombres y códigos tal como aparecen** — no los renombres.

##### Resumen de tipos disponibles

Empieza esta sección con una tabla que lista todos los tipos disponibles, indicando el formato de salida de cada uno:

```
| Código | Nombre en el selector | Formato de salida | Procedimiento almacenado principal |
|---|---|---|---|
| (<código>) | <Nombre exacto del combo> | Excel / RTF | `<nombre_sp>` |
```

##### Subtítulo por tipo de informe

Crea un subtítulo `###` por cada tipo con el formato:

```
### (<código>) <Nombre exacto del combo> (`<nombre_función>`)
```

El nombre de la función se obtiene del bloque `Select Case` del formulario: es la función que se llama en el `Case` correspondiente al código del tipo. Si el tipo llama funciones distintas según una condición (por ejemplo, si una casilla está marcada), incluye ambas separadas por `/`.

Dentro de cada subtítulo documenta en este orden:

1. **Qué muestra:** 1–3 oraciones describiendo qué información contiene el archivo o documento generado, en lenguaje de usuario.

2. **Restricciones propias del tipo** (solo si las hay): condiciones específicas de fechas, volumen de datos u otras que apliquen únicamente a este tipo.

3. **Cómo se seleccionan los servicios:** indica si usa la grilla de servicios (casillas por servicio) o el árbol jerárquico (servicio → estructura de servicio), ya que determina el nivel de detalle disponible.

4. **Opciones de configuración disponibles:** lista los parámetros que el usuario puede ajustar antes de exportar. Solo las que apliquen a este tipo.

```
**Opciones de configuración disponibles:**
- **<Nombre de la opción>:** <qué controla y qué valores tiene>.
```

5. **Estructura de datos del informe:** tabla con cada campo o columna que el informe entrega al usuario. Para obtenerla, lee el código de construcción del Excel o del RTF (bucles de escritura de celdas, encabezados de columna, campos del Recordset) y el SP correspondiente (columnas del SELECT final).

```
**Estructura de datos del informe:**

| Campo / Columna | Descripción | Calculado |
|---|---|---|
| <nombre visible al usuario> | <qué representa en el proceso operativo> | Sí / No |
```

   - **Calculado = Sí** cuando el valor no proviene directamente de un campo almacenado, sino de: una operación aritmética, un SP, una función SQL o VB6, o una subconsulta.
   - Para cada campo con **Calculado = Sí**, agrega inmediatamente después de la tabla una subsección con este formato. Estos cálculos van aquí, dentro del tipo de informe — **no en la sección 4.2**, salvo que el mismo cálculo aplique de forma transversal a todos los tipos.

```
**Cálculo — <Nombre del campo>**

<Explicación en lenguaje simple: qué representa y por qué se calcula en lugar de almacenarse.>

**Fórmula o lógica:**
<Nombre resultado> = <Componente A> × <Componente B> …
— o bien —
<Descripción paso a paso si no es una fórmula simple>

| Componente | Qué representa | De dónde viene |
|---|---|---|
| <Componente A> | <descripción> | <tabla.campo>, SP `<nombre>`, o "ingresado por el usuario" |

> Ejemplo: <valores ficticios que ilustren el cálculo paso a paso>
```

6. **Estructura del archivo generado:** cómo está organizado el Excel o documento RTF que recibe el usuario.

```
**Formato de salida:** <Excel / Documento RTF>. <Una hoja por servicio / una única hoja / el usuario elige la ruta con cuadro de diálogo>. <Orientación retrato/paisaje si es RTF>. Encabezado en filas <N–M> con <qué contiene>. Datos desde fila <N>. <Otras características relevantes para el usuario: agrupaciones, subtotales, columnas fijas, etc.>
```

##### Cuando NO hay selector de tipos

Si el formulario genera un único reporte, omite la tabla de resumen y el subtítulo por tipo. Documenta directamente:
- Qué información muestra o entrega.
- Opciones de configuración disponibles.
- Formato de salida (pantalla / Excel / RTF / impresora).

---

#### Sección 6 — Referencia técnica

Esta sección es de consulta para quien necesite conocer las tablas o procedimientos involucrados. No es de lectura obligatoria para operar la pantalla.

##### Tablas que intervienen

```
| Tabla | Para qué se usa en este reporte | Campos clave |
|---|---|---|
| `<nombre_tabla>` | <rol: fuente principal, catálogo de referencia, tabla temporal, etc.> | `<campo1>`, `<campo2>` |
```

##### Relación con otros módulos

```
| Módulo | Relación |
|---|---|
| **<Nombre>** | <De dónde vienen los datos que este reporte consume / qué proceso los genera.> |
```

---

#### Pie del MD

```
---

*Fuentes: `<formulario>.frm`, SP `<nombre>` en `SGP_Admin.sql`, tabla(s) `<tabla>` en `SGP_Admin.sql`*
```

---

### Reglas de redacción — obligatorias en todo el documento

Todo texto explicativo del MD (párrafos, descripciones de secciones, entradas de tablas funcionales) debe escribirse en **lenguaje funcional de usuario**: claro, orientado a quien opera el sistema, sin exponer implementación interna. Los nombres técnicos pertenecen a tablas de referencia y bloques de código — no al texto que explica.

| ❌ Evitar en el texto explicativo | ✅ Usar en su lugar |
|---|---|
| Nombres de componentes VB6 como término principal (`vaSpread1`, `fpDateTime1`, `Combo1`) | "grilla de resultados", "campo de fecha", "lista desplegable" |
| Nombres de métodos internos como verbo principal (`GrabaRegistro()`, `LlenaDatos()`) | "el sistema guarda", "el sistema carga el catálogo" |
| Nombres de funciones de exportación de módulos `.bas` (`I_MenuPlanMecanoBloque`, `ExportarExcelMenuMensualMKT`) | Omitir del texto; solo incluirlos en el subtítulo del tipo entre paréntesis |
| Objetos de automatización o diálogos de sistema (`VSPrinter`, `Preview.VSPrinter`, `CommonDialog.ShowSave`, `CreateObject("Excel.Application")`) | "ventana de Vista Previa del sistema", "cuadro de diálogo de guardado donde el usuario elige el nombre y la carpeta del archivo", "genera un archivo Excel" |
| Nombres internos de formularios secundarios (`B_HistPm`, `B_MTaEst`) | "formulario de histórico de minutas en bloque", "selector de servicios", "selector de nutrientes" |
| Parámetros y constantes internas (`EstadoPresentacion = "1"`, `@@spid`, `paso_servicio`) | Omitir del texto; si es imprescindible mencionar la tabla temporal, describirla como "tabla temporal que aísla los datos de cada usuario conectado simultáneamente" |
| Campos de control de estado como condición visible (`red_IndentificadorIngSumaTablaGramaje = '1'`) | "marcados como referencia de gramaje en la receta" |
| Códigos de error numéricos (`-2147467259`) | "el registro tiene datos asociados en otra tabla" |
| Nombres de eventos VB6 (`Form_Load`, `ButtonClicked`) | "al abrir el formulario", "al hacer clic en el botón" |
| Variables globales como sujeto (`vg_tipbase`, `vg_pais`) | "según la configuración del sistema", "el país configurado en sesión" |
| Jerga SQL en texto libre ("JOIN", "WHERE", "NULL", "FK") | Reservar esos términos solo para bloques de código; en texto explicar en palabras |
| Nombres de tablas sin contexto | Siempre acompañar con la descripción funcional entre paréntesis la primera vez que aparecen |

---

*Referencia de estilo: `{{RUTA_DOC}}\md_pantallas\SGP_Admin\Informe_Planificación.md` — última actualización: 2026-03-23*
