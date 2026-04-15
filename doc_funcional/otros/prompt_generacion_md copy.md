# Prompt de Generación — Documentación Funcional de Formularios SGP

Este documento describe el proceso paso a paso para generar un MD funcional de cualquier formulario del sistema SGP Producción, siguiendo el estándar establecido.

---

## Contexto del proyecto

El sistema **SGP Local** es una aplicación de gestión de producción y servicios de alimentación para casinos Sodexo Chile. Está desarrollado en **Visual Basic 6** con base de datos **SQL Server** (y compatibilidad Access legacy). Los formularios `.frm` contienen la lógica de interfaz y negocio; los procedimientos almacenados y funciones están en el archivo SQL.

Los documentos Markdown generados son de uso funcional: los leen analistas, jefes de casino y coordinadores que no necesariamente conocen programación, por lo que el lenguaje debe ser claro, orientado al usuario y libre de jerga técnica interna.

---

## Archivos de referencia

| Recurso | Ruta |
|---|---|
| Formularios VB6 | `C:\Users\Joaquin.Suazo\Documents\SGP-Producción\codigo_fuente\<nombre>.frm` |
| Base de datos SQL (UTF-16 LE) | `C:\Users\Joaquin.Suazo\Documents\SGP-Producción\base_de_datos\SGP_Local.sql` |
| Carpeta de salida MDs | `C:\Users\Joaquin.Suazo\Documents\SGP-Producción\doc_funcional\` |
| MD de referencia de estilo | `doc_funcional\ListaPrecioCafeteria.md` |

**Conversión del archivo SQL antes de buscar en él:**
```bash
iconv -f UTF-16 -t UTF-8 \
  "C:/Users/Joaquin.Suazo/Documents/SGP-Producción/base_de_datos/SGP_Local.sql" \
  > /tmp/SGP_Local_utf8.sql
```

---

## Paso 1 — Leer el formulario VB6

Leer el archivo `.frm` completo. Identificar:

- **Nombre y propósito general** del formulario (¿qué tarea del usuario resuelve?)
- **Tablas principales** que se leen o escriben (buscar `FROM`, `INTO`, `UPDATE`, `DELETE` en el código)
- **Stored Procedures o funciones SQL** referenciados (buscar `.Execute`, `sgp_`, `fg_Cal`, `RutinaLectura.`)
- **Operaciones disponibles** (botones de toolbar o botones Command: Agregar, Modificar, Eliminar, Grabar, Cancelar, Refrescar, Imprimir, Cerrar)
- **Validaciones** presentes (bloques `If`, mensajes `MsgBox`, condiciones antes de grabar)
- **Estructura de la grilla** (columnas del Spread o MSFlexGrid y su origen de datos)
- **Campos de encabezado** que el usuario debe completar antes de operar (combos, textboxes de filtro)
- **Flujo principal** de uso (qué hace el usuario paso a paso)

---

## Paso 2 — Buscar SPs y funciones en el SQL

Por cada SP o función identificada en el paso anterior, buscar en `/tmp/SGP_Local_utf8.sql`:

```bash
grep -n "nombre_sp_o_funcion" /tmp/SGP_Local_utf8.sql
```

Luego leer el bloque completo del SP/función para entender:
- Qué tablas consulta o modifica
- Qué parámetros recibe
- Qué lógica aplica
- Qué devuelve

Si no existe SP (operaciones SQL inline), documentar igualmente las consultas encontradas en el VB6.

---

## Paso 3 — Redactar el MD

Crear el archivo en `doc_funcional\<NombreFuncional>.md` siguiendo exactamente esta estructura de secciones:

---

### Encabezado del MD

```markdown
# <Nombre funcional del formulario>

**Formulario VB6:** `<nombre>.frm`
**Tabla(s) principal(es):** `<tabla1>` (<descripción>), `<tabla2>` (<descripción>)
**SP principal de lectura / grabado:** `<nombre_sp>` — o — Sin Stored Procedures: todas las operaciones se realizan con SQL directo

---
```

---

### Sección: Contexto

Describir en 2-3 párrafos:
- Para qué sirve el formulario en el proceso operativo del casino
- A qué etapa del flujo de producción pertenece (planificación, salida, ventas, mermas, cierre)
- Si depende de fechas, periodos, minuta u otros prerequisitos
- Cómo se organiza visualmente (pestañas, paneles, etc.)

**Regla:** no mencionar nombres de variables internas, eventos VB6 ni métodos. Hablar siempre en términos del usuario y del proceso.

---

### Sección: Parámetros de Entrada

Tabla con los campos del encabezado que el usuario debe completar antes de operar:

```markdown
| Campo | Descripción | Obligatorio |
|---|---|---|
| <campo> | <qué representa y para qué sirve> | Sí / No |
```

Si no hay encabezado, indicar que el formulario carga automáticamente al abrirse.

---

### Sección: Estructura de la Grilla

Una subsección por cada grilla o pestaña que tenga el formulario:

```markdown
| Col | Nombre | Origen | Editable | Observaciones |
|---|---|---|---|---|
| 1 | <nombre visible> | `<tabla.campo>` | Sí / No | <condiciones o reglas> |
```

Incluir notas al pie si hay columnas ocultas con uso interno importante (por ejemplo, claves de referencia).

---

### Sección: Operaciones Disponibles

Tabla con cada botón/acción del formulario:

```markdown
| Botón | Acción |
|---|---|
| **Agregar** | <descripción funcional de qué hace> |
| **Modificar** | <descripción funcional> |
| **Eliminar** | <condiciones y resultado> |
| **Grabar** | <qué persiste y dónde> |
| **Cancelar** | <qué descarta y cómo restaura> |
| **Refrescar** | <qué recarga> |
| **Imprimir** | <qué informe genera> |
| **Cerrar** | Cierra el formulario. |
```

Si aplica por pestaña, agregar columna "Pestaña".

---

### Sección: Validaciones

Agrupar por contexto (pestaña o momento). Cada validación en una fila:

```markdown
| # | Momento | Condición | Resultado |
|---|---|---|---|
| 1 | Al grabar | <qué se verifica> | <mensaje o acción que toma el sistema> |
```

**Regla:** describir las condiciones en lenguaje de negocio, no en código. En vez de `precio = 0`, escribir "precio igual a cero". En vez de `-2147467259`, escribir "el registro tiene datos asociados en otra tabla".

---

### Sección: Flujo de Datos

Diagrama de texto que muestra los pasos del proceso desde la perspectiva del usuario:

```
1. Usuario ingresa: <campos requeridos>
        │
        ▼
2. Sistema carga: <qué datos muestra>
        │
        ▼
3. Usuario selecciona operación:

   [AGREGAR]          [MODIFICAR]        [ELIMINAR]
       │                   │                  │
   <pasos>            <pasos>            <pasos>
       │                   │                  │
       ▼                   ▼                  ▼
4. [GRABAR] → <descripción de lo que se persiste y en qué tabla>
```

**Regla:** no incluir nombres de funciones VB6 (`GrabaRegistro`, `MoverDatosGrillas`, etc.) ni transacciones internas. Describir solo lo que el usuario percibe o lo que el sistema hace visiblemente.

---

### Sección: Dónde se Almacena

Una subsección por tabla principal, con la descripción funcional de cada campo:

```markdown
### <Nombre descriptivo de la tabla> (`<nombre_tabla>`)

| Campo | Descripción |
|---|---|
| `campo_codigo` | <para qué sirve en el proceso> |
| `campo_fecha` | <cuándo se graba y qué representa> |
```

Al final de cada tabla: indicar la **clave primaria** y su significado funcional (qué combinación de valores identifica unívocamente un registro).

---

### Sección: Consultas de Lectura (si no hay SP)

Si el formulario usa SQL directo en vez de SPs, documentar cada consulta con:

1. **Título descriptivo** (qué obtiene)
2. **Párrafo explicativo** en lenguaje simple:
   - Qué información trae
   - Cuándo se ejecuta
   - Qué campos retorna y para qué se usan
   - Si cruza varias tablas, explicar por qué en palabras simples
3. **Bloque SQL** a continuación del párrafo

```markdown
**<Título de la consulta>**

> <Explicación en lenguaje simple: qué obtiene, cuándo se ejecuta, qué campos trae y para qué sirven. Si cruza tablas, explicar por qué sin usar jerga técnica.>

```sql
<consulta sql>
```
```

---

### Sección: SP / Funciones Referenciados (si existen)

Si el formulario llama a SPs o funciones SQL, documentar cada uno:

```markdown
### `nombre_sp` — <descripción en una línea>

**Parámetros de entrada:**

| Parámetro | Descripción |
|---|---|
| `:param` | <qué representa> |

**Lógica principal:**
<descripción en lenguaje funcional de qué hace el SP paso a paso>

**Tablas que modifica:** `<tabla1>`, `<tabla2>`
```

---

### Sección: Relación con Otros Módulos

Tabla de relaciones con otros formularios o procesos:

```markdown
| Módulo | Relación |
|---|---|
| **<Nombre del módulo>** | <cómo se relaciona: prerequisito, destino de los datos, usa los mismos datos, etc.> |
```

---

### Pie del MD

```markdown
---

*Fuentes: `<formulario>.frm`, `<otros archivos consultados>`, tabla(s) `<tabla1>`, `<tabla2>` en `SGP_Local.sql`*
```

---

## Reglas generales de redacción

| Evitar | Usar en su lugar |
|---|---|
| Nombres de variables VB6 (`vaSpread1`, `fpText1`, `modo`) | "grilla de artículos", "campo de búsqueda", "al agregar / al modificar" |
| Nombres de métodos internos (`GrabaRegistro()`, `MoverDatosGrillas()`) | "el sistema guarda el registro", "el sistema recarga la grilla" |
| Códigos de error numéricos (`-2147467259`, `3034`) | "el registro tiene datos asociados en otra tabla" |
| Nombres de eventos VB6 (`Form_Load`, `LeaveCell`, `ButtonClicked`) | "al abrir el formulario", "al salir de la fila", "al hacer clic en el botón" |
| Propiedades de formulario (`TabEnabled`, `Cancel = True`) | "la pestaña queda deshabilitada", "el cursor permanece en el campo" |
| Variables globales (`vg_tipbase`, `vg_pais`, `MuestraCasino`) | "según el motor de base de datos activo", "según el país configurado", "el casino activo en sesión" |
| Jerga SQL en texto libre ("JOIN", "WHERE", "NULL") | Reservar para bloques de código; en texto explicar con palabras |

---

## Ejemplo de invocación

Para generar el MD de un nuevo formulario, usar el siguiente prompt:

```
Analiza el formulario VB6 `C:\Users\Joaquin.Suazo\Documents\SGP-Producción\codigo_fuente\<NombreForm>.frm`
y genera un documento Markdown funcional siguiendo el estándar definido en
`doc_funcional\prompt_generacion_md.md`.

Si el formulario referencia Stored Procedures o funciones SQL, búscalos en el archivo SQL:
  iconv -f UTF-16 -t UTF-8 "...SGP_Local.sql" > /tmp/SGP_Local_utf8.sql

Escribe el resultado en:
  C:\Users\Joaquin.Suazo\Documents\SGP-Producción\doc_funcional\<NombreFuncional>.md

Usa como referencia de estilo el archivo:
  doc_funcional\ListaPrecioCafeteria.md
```

---

*Última actualización: 2026-03-13 — basado en `ListaPrecioCafeteria.md` como documento de referencia*
