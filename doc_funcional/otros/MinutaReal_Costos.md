# Planificación Real — Cálculo de Costos

**Formulario VB6:** `M_MinRea.frm`
**Sub principal de cálculo:** `CargarCosto()` / `MostrarCosto(Col)`
**Módulo de consultas:** `RutinaLectura.cls`
**SP de grabado:** `sgp_Ins_XmlMinutaReal`

---

## Contexto

El formulario de Planificación Real muestra un panel de costos (visible al presionar el botón **"Visualizar Costo"** o al intentar grabar). Este panel se organiza en **tres cuadros** que reflejan distintos horizontes de análisis: el mes completo, el día seleccionado, y el acumulado hasta ese día. Los valores se calculan enteramente en memoria dentro de la función `CargarCosto()` y se actualizan al navegar entre días con `MostrarCosto()`.

---

## Vector Central: VecCos

Todo el cálculo de costos se almacena en un vector bidimensional `VecCos(día, componente)`, donde el índice de fila corresponde al día del mes (1 a N) y el índice de columna a un tipo de dato:

| Índice | Contenido | Origen |
|---|---|---|
| `VecCos(d, 1)` | Costo Mat. Prima planificado del día | Σ (costo_receta × raciones) en la grilla |
| `VecCos(d, 2)` | Costo Estructura Fija del día | Consulta a `b_minutafijadia` o `b_minutafija` × PMP |
| `VecCos(d, 3)` | Costo Salida a Producción real del día | Consulta a `b_totventas` / `b_detventas` (doc SP − DP) |
| `VecCos(d, 4)` | Raciones planificadas del día | Fila de totales de la grilla (última fila) |
| `VecCos(d, 5)` | Raciones producidas del día | `b_minutaraciones` donde cliente = `'PRODUCIDAS'` |

---

## Fuentes de Datos por Componente

### ① Costo Materia Prima Planificado (`VecCos(d, 1)`)

Se recorre la grilla receta a receta para el día `d`:

```
Costo Mat.Prima día = Σ por cada receta del día:
    Costo Receta × Raciones Planificadas
```

| Variable | Campo | Tabla |
|---|---|---|
| Costo Receta | `mid_cosrec + mid_cosdes` | `b_minutadet` |
| Raciones Planificadas | `mid_numrac` | `b_minutadet` |

> El costo receta es el valor **congelado al grabar la planificación** (no el precio actual de bodega). Se puede actualizar manualmente con el botón **"Actualizar Costo Recetas"**, que recalcula usando `fg_CalCtoRecInv()` y escribe nuevamente en `b_minutadet`.

---

### ② Costo Estructura Fija (`VecCos(d, 2)`)

Representa los costos fijos del servicio (personal, energía, etc.) que no dependen de la receta planificada. Tiene dos fuentes según disponibilidad:

**Caso A — Existe estructura fija cargada por día (`b_minutafijadia`):**

```
Costo EF día = Σ (cantidad_producto × costo_producto)
```

| Campo | Tabla | Descripción |
|---|---|---|
| `mfd_canpro` | `b_minutafijadia` | Cantidad del ítem de estructura fija |
| `mfd_cospro` | `b_minutafijadia` | Costo unitario del ítem |

Consulta: `RutinaLectura.MinutaFijaDia(2, ...)` → `SUM(mfd_canpro × mfd_cospro)`

**Caso B — No existe por día, pero existe estructura fija genérica (`b_minutafija`):**

```
Costo EF día = Σ (PMP_producto × cantidad_receta_EF)
```

| Campo | Tabla | Descripción |
|---|---|---|
| `ppd_propon` | `b_productospmpdia` | Precio promedio ponderado del producto en esa fecha |
| `mif_canpro` | `b_minutafija` | Cantidad del ítem en la estructura fija |

Consulta: `RutinaLectura.MinutaFija(2 o 3, ...)` → `SUM(ppd_propon × mif_canpro)`

> La tabla `b_productospmpdia` se carga en una tabla temporal (`{usuario}_tmp_ProductoPMPCargarCostoRea`) con el PMP del día anterior al cierre diario. Si el motor de base de datos es Access se usa opción 2 (con tabla temporal); si es SQL Server se usa opción 3 (con `b_productospmpdia` directamente).

---

### ③ Costo Salida a Producción Real (`VecCos(d, 3)`)

Representa el costo **efectivamente ejecutado**: lo que salió de bodega para producción menos las devoluciones.

```
Costo Salida día = Σ (totales documentos SP) − Σ (totales documentos DP)
```

| Documento | Tipo | Efecto |
|---|---|---|
| Salida a Producción | `tov_tipdoc = 'SP'` | Suma (`+dev_ptotal`) |
| Devolución Producción | `tov_tipdoc = 'DP'` | Resta (`−dev_ptotal`) |

**Filtros aplicados:**
- Régimen y Servicio coinciden con los del formulario
- Solo productos de cuenta contable de insumo (parámetro `ctainsumo`)
- Documento no anulado ni pendiente (`tov_estdoc ≠ 'A'` y ≠ `'P'`)
- Bodega del contrato (`tov_codbod`)

Tablas: `b_totventas`, `b_detventas`, `b_productos`

---

### ④ Raciones Planificadas (`VecCos(d, 4)`)

Se lee directamente de la última fila de la grilla (fila de totales), columna de raciones del día `d`.

---

### ⑤ Raciones Producidas (`VecCos(d, 5)`)

```sql
SELECT SUM(mir_nrorac) FROM b_minutaraciones
WHERE mir_rutcli = 'PRODUCIDAS'
  AND mir_cencos = :contrato
  AND mir_codreg = :regimen
  AND mir_codser = :servicio
  AND mir_fecmin = :fecha_dia
```

Consulta: `RutinaLectura.MinutaRaciones(1, ..., 'PRODUCIDAS', fecha)`

---

## Los Tres Cuadros de Costos

### Cuadro 1 — "Total Mes" (Frame2 índice 1)

Muestra el **acumulado de todos los días del mes**, independientemente del día seleccionado en la grilla.

| Fila | Etiqueta | Valor | Fórmula |
|---|---|---|---|
| Mat.Prima | Planificado | `Label1(7)` | Σ VecCos(d,1) — todos los días |
| Est.Fija | Planificado | `Label1(11)` | Σ VecCos(d,2) — todos los días |
| Total | Planificado | `Label1(12)` | Label1(7) + Label1(11) |
| Rac. | — | `Label1(48)` | Σ VecCos(d,4) — todos los días |
| Cto.Band. (Mat.Prima) | Planificado | `Label1(40)` | Label1(7) / Label1(48) |
| Cto.Band. (Est.Fija) | Planificado | `Label1(41)` | Label1(11) / Label1(48) |
| **Cto.Band. Total** | **Planificado** | **`Label1(8)`** | **(Label1(7) + Label1(11)) / Label1(48)** |

> `Label1(8)` es el **Costo Bandeja total del mes** y se usa en la validación contra el Costo Techo antes de grabar.

---

### Cuadro 2 — "Día DD/MM/AAAA" (Frame2 índice 2)

Muestra los costos del **día actualmente seleccionado** en la grilla. El título del frame cambia dinámicamente con la fecha del día. Tiene columnas **Planificado** y **Realizado**.

| Fila | Etiqueta | Planificado | Realizado |
|---|---|---|---|
| Mat.Prima | — | `Label1(20)` = VecCos(d,1) | — |
| Est.Fija | — | `Label1(21)` = VecCos(d,2) | — |
| Cto.Total | — | `Label1(22)` = VecCos(d,1) + VecCos(d,2) | `Label1(23)` = VecCos(d,3) (Salida Producción) |
| Rac. | — | `Label1(44)` = VecCos(d,4) | `Label1(46)` = VecCos(d,5) |
| **Cto.Band.** | — | **`Label1(45)`** = (VecCos(d,1)+VecCos(d,2)) / VecCos(d,4) | **`Label1(47)`** = VecCos(d,3) / VecCos(d,5) |

> - **Cto.Band. Planificado**: costo por ración según planificación de recetas.
> - **Cto.Band. Realizado**: costo por ración según lo que efectivamente salió de bodega dividido por raciones producidas.

---

### Cuadro 3 — "Acumulado hasta DD/MM/AAAA" (Frame2 índice 3)

Muestra el **acumulado desde el día 1 del mes hasta el día seleccionado**. El título cambia dinámicamente. También tiene columnas Planificado y Realizado.

| Fila | Etiqueta | Planificado | Realizado |
|---|---|---|---|
| Mat.Prima | — | `Label1(31)` = Σ VecCos(1..d, 1) | — |
| Est.Fija | — | `Label1(32)` = Σ VecCos(1..d, 2) | — |
| Cto.Total | — | `Label1(33)` = Label1(31) + Label1(32) | `Label1(36)` = Σ VecCos(1..d, 3) |
| Rac. | — | `Label1(34)` = Σ VecCos(1..d, 4) | `Label1(37)` = Σ VecCos(1..d, 5) |
| **Cto.Band.** | — | **`Label1(35)`** = Label1(33) / Label1(34) | **`Label1(38)`** = Label1(36) / Label1(37) |

---

## Cuadros de Referencia: Costo Patrón (Techo y Piso)

Antes de mostrar los cuadros de costo, el sistema consulta `b_costopatron` para obtener los valores de referencia del mes:

| Parámetro | Descripción | Campo |
|---|---|---|
| **TECHO** | Costo máximo permitido por bandeja | `cpa_valor` donde `cpa_descripcion = 'TECHO'` |
| **PISO** | Costo mínimo de referencia por bandeja | `cpa_valor` donde `cpa_descripcion = 'PISO'` |

Consulta: `RutinaLectura.CostoPatron(1, regimen, servicio, anomes)` → `SELECT cpa_descripcion, cpa_valor FROM b_costopatron WHERE ...`

Estos valores se muestran como **filas de referencia en la grilla** (encabezado superior) si están definidos:
- Si existe TECHO → fila `SpreadHeader`: "Costo Patrón Techo"
- "Costo Minuta Día" siempre se muestra
- Si existe PISO → fila siguiente: "Costo Patrón Piso"

---

## Validación de Costo Techo al Grabar

Antes de ejecutar el grabado (`sgp_Ins_XmlMinutaReal`), el sistema verifica el costo contra el techo:

```
Si vCtoTec > 0 Y CostoBandeja > 0:
    Si CostoBandeja > (vCtoTec × 1.05):
        → Advertencia: "Costo minuta día (X) es mayor costo techo (Y)"
        → Recorre día a día e identifica cuáles superan vCtoTec
        → Muestra lista de días infractores (no bloquea, es informativo)
```

Donde `CostoBandeja` = `Label1(8)` = Costo Bandeja Total Mes.

---

## Actualización de Costos de Recetas ("Actualizar Costo Recetas")

Opción independiente que recalcula los costos de receta con los precios actuales de bodega:

```
1. Para cada receta planificada en el mes:
   cosali = fg_CalCtoRecInv(cod_receta, tip_receta, cuentas_insumo)
   cosdes = fg_CalCtoRecInv(cod_receta, tip_receta, cuentas_desechable)

2. UPDATE b_minutadet
   SET mid_cosrec = cosali, mid_cosdes = cosdes
   WHERE ... AND mid_fecmin = fecha AND mid_codrec = cod AND mid_tipmin = '2'
```

> Después de esta actualización los costos en la grilla se refrescan y los cuadros de costo se recalculan. Si hay cambios pendientes sin grabar, el sistema advierte que se graben primero.

---

## Resumen de Consultas y Tablas

| Componente | Consulta/Fuente | Tabla(s) |
|---|---|---|
| Costo Mat.Prima | Grilla en memoria (mid_cosrec + mid_cosdes) × mid_numrac | `b_minutadet`, `b_minuta` |
| Costo Estructura Fija (por día) | `RutinaLectura.MinutaFijaDia(2, ...)` | `b_minutafijadia`, `b_productos` |
| Costo Estructura Fija (genérico) | `RutinaLectura.MinutaFija(2 o 3, ...)` | `b_minutafija`, `b_productospmpdia` |
| Costo Salida Producción real | SQL inline (SP − DP) | `b_totventas`, `b_detventas`, `b_productos` |
| Raciones Planificadas | Grilla en memoria (fila totales) | — |
| Raciones Producidas | `RutinaLectura.MinutaRaciones(1, ..., 'PRODUCIDAS', ...)` | `b_minutaraciones` |
| Costo Patrón Techo/Piso | `RutinaLectura.CostoPatron(1, ...)` | `b_costopatron` |
| PMP para Estructura Fija | Tabla temporal en memoria | `b_productospmpdia` |
| Grabado planificación | `sgp_Ins_XmlMinutaReal` (XML) | `b_minutadet`, `b_minuta` |

---

## Diagrama Conceptual

```
                     CargarCosto()
                          │
          ┌───────────────┼────────────────┐
          │               │                │
    Grilla (VB6)   b_minutafijadia   b_totventas
    cosrec×numrac  o b_minutafija    (SP − DP)
    ═══ VecCos(d,1)  ══ VecCos(d,2)  ══ VecCos(d,3)
          │               │                │
          │          b_minutaraciones       │
          │          'PRODUCIDAS'           │
          │          ══ VecCos(d,5)         │
          │               │                │
          └───────────────┼────────────────┘
                          │
                    MostrarCosto(Col)
                          │
          ┌───────────────┼──────────────────┐
          │               │                  │
    Total Mes          Día d           Acumulado 1..d
   (todos los días)  (día activo)    (hasta día activo)
   Label1(7..8,11,   Label1(20..23,  Label1(31..38)
   12,40,41,48)      44..47)
```

---

*Fuentes: `M_MinRea.frm` (subs `CargarCosto` y `MostrarCosto`), `RutinaLectura.cls` (funciones `CostoPatron`, `MinutaFija`, `MinutaFijaDia`, `MinutaRaciones`), `SGP_Local.sql`*
