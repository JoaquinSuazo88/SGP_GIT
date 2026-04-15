# Fórmulas y Lógica — Raciones no Vendidas / Merma por Preparación

**Formulario VB6:** `M_MerPre.frm`
**SP principal de lectura:** `sgp_Sel_MermaPorPreparacion`
**SP de grabado:** `sgp_Upd_XmlMermaPreparacion`
**Función auxiliar:** `SGP_FN_RNVCantidadesReceta`

---

## Contexto

Este formulario registra, para un día y servicio determinado, las raciones de cada receta que **no se vendieron** y su equivalente en kilogramos. Se accede desde el módulo de **Cierre Diario** y requiere que exista previamente una minuta real (`mid_tipmin = '2'`) con raciones planificadas (`mid_numrac > 0`) y que se haya realizado la salida a producción.

---

## Parámetros de Entrada

| Campo | Descripción |
|---|---|
| Contrato (CeCo) | Identificador del casino |
| Régimen | Código del régimen alimenticio |
| Servicio | Código del servicio (desayuno, almuerzo, cena, etc.) |
| Fecha | Día a registrar |

---

## Estructura de la Grilla

Cada fila corresponde a una receta planificada en la minuta real del día.

| Col | Nombre | Origen | Editable |
|---|---|---|---|
| 1 | Código Receta | `b_receta.rec_codigo` | No |
| 2 | Nombre Receta | `b_receta.rec_nombre` | No |
| 3 | Raciones Planificadas | `b_minutadet.mid_numrac` | No |
| 4 | Costo Receta | `mid_cosrec + mid_cosdes` | No |
| 5 | Costo Total Planificado | Costo Receta × Raciones Planificadas | No |
| 6 | **Merma x Ración** | `mid_nummer` | **Sí** |
| 7 | **Merma x Kilo (Servido)** | `mid_mermaxcantservida` | **Sí** |
| 8 | Costo Merma | Costo Receta × Merma x Ración | No (calculado) |
| 9 | Nº Línea | `mid_numlin` | No (interno) |
| 10 | Kilo Bruto | `mid_mermaxkilo` | No (calculado) |

---

## Nivel 1 — Merma por Receta

### ① Merma x Ración

**Qué es:** Número de raciones de esa receta que no se vendieron en el día.

**Ingreso:** Manual por el usuario.

**Validación crítica:**
```
Merma x Ración  ≤  Raciones Planificadas
```
> El sistema rechaza cualquier valor que supere las raciones planificadas.
> Los campos Merma x Ración y Merma x Kilo van **obligatoriamente en par**: si se ingresa uno, debe ingresarse el otro.

---

### ② Merma x Kilo — Cantidad Servida

**Qué es:** Equivalente en kilogramos de las raciones no vendidas, calculado con el **peso neto servido** al comensal (después de preparación y cocción).

**Fórmula (calculada automáticamente al ingresar Merma x Ración):**

```
Merma x Kilo (Servido) = (Gramaje Servido por Ración × Raciones No Vendidas) / Factor Conversión
```

**Componentes:**

| Componente | Descripción | Origen |
|---|---|---|
| Gramaje Servido por Ración | Peso neto por porción (ver detalle abajo) | `SGP_FN_RNVCantidadesReceta(..., 'S')` |
| Raciones No Vendidas | Valor ingresado en Merma x Ración | Campo Col 6 |
| Factor de Conversión | Gramos por kilogramo según parámetro del contrato | `a_param.pargrarnve` (típicamente 1.000) |

**Cálculo del Gramaje Servido por Ración** (función `SGP_FN_RNVCantidadesReceta` con tipo `'S'`):

```
Gramaje Servido = Σ por cada ingrediente de la receta:
    Cantidad Efectiva × (% Aprovechamiento / 100) × (% Cocción / 100)
```

Donde **Cantidad Efectiva** depende del tipo de unidad del ingrediente:

| Condición | Cantidad Efectiva |
|---|---|
| Caso general | `red_canpro` (cantidad declarada en receta) |
| Und + C/u + `ing_facnut > 0` | `ROUND((100 / ing_facnut) × pro_facing, 0) × red_canpro` |

**Conversión para ingredientes en unidad "Und"**

Se activa cuando se cumplen las **tres condiciones** simultáneamente:
- La unidad del producto en bodega es `"Und"` (`a_unidad.uni_nomcor`)
- La unidad del ingrediente en la receta es `"C/u"` (`a_unidadmed.unm_nomcor`)
- El ingrediente tiene factor nutricional mayor a cero (`b_ingrediente.ing_facnut > 0`)

```
Cantidad Efectiva = ROUND( (100 / Factor Nutricional) × Facing del Producto, 0 ) × Cantidad Receta
```

| Variable | Campo | Significado |
|---|---|---|
| Factor Nutricional | `b_ingrediente.ing_facnut` | Número de unidades equivalentes a 100 g |
| Facing del Producto | `b_productos.pro_facing` | Peso en gramos de una unidad del producto |
| Cantidad Receta | `b_recetadet.red_canpro` | Cantidad indicada en la receta |

> **Ejemplo:** Huevo con `ing_facnut = 4` y `pro_facing = 250 g`:
> → `(100 / 4) × 250 = 6.250 g` como gramaje efectivo por unidad recetada.
> El `pro_facing` convierte "unidad" a gramos; el `ing_facnut` ajusta según la porción nutricional de referencia.

**Entrada inversa (edición directa del kilo):**
Si el usuario ingresa directamente los kilos, el sistema calcula las raciones implicadas:
```
Raciones Implicadas = (Kilos Ingresados / Gramaje Servido por Ración) × Factor Conversión
```
Si ese resultado supera las raciones planificadas, el valor es rechazado.

---

### ③ Kilo Bruto

**Qué es:** Equivalente en kilogramos usando el **peso bruto** de los ingredientes, es decir, antes de aplicar cocción o aprovechamiento.

**Fórmula:**

```
Kilo Bruto = (Gramaje Bruto por Ración × Raciones No Vendidas) / Factor Conversión
```

**Gramaje Bruto por Ración** (función `SGP_FN_RNVCantidadesReceta` con tipo `'B'`):

```
Gramaje Bruto = Σ por cada ingrediente de la receta:
    Cantidad Efectiva (sin ajustes de aprovechamiento ni cocción)
```

Aplica la misma **conversión para ingredientes en unidad "Und"** descrita en ②: cuando se cumplen las tres condiciones (Und + C/u + `ing_facnut > 0`), se usa `ROUND((100 / ing_facnut) × pro_facing, 0) × red_canpro` en lugar de `red_canpro` directamente.

---

### ④ Costo de la Merma

**Fórmula:**

```
Costo Merma = (Costo Alimentación + Costo Desechable) × Raciones No Vendidas
```

> Los costos utilizados son los **congelados al momento de grabar la planificación** (`mid_cosrec`, `mid_cosdes`), no el precio actual del producto en bodega.

---

### ⑤ Totales de la Grilla

| Total | Fórmula |
|---|---|
| **Total Costo Producción** | Σ (Costo Receta × Raciones Planificadas) para todas las recetas del día |
| **Total Costo Merma** | Σ (Costo Receta × Raciones No Vendidas) para todas las recetas del día |

---

## Nivel 2 — Mermas Kilos Globales del Día

Estos tres campos son **independientes de las recetas**. Se ingresan como totales del día completo y se registran en la tabla `b_mermadesconche`.

| Campo | Qué registra |
|---|---|
| **Merma Producción** | Kilos de alimento producidos que no llegaron a servirse (merma de cocina) |
| **Desconche** | Kilos de comida que el comensal dejó en el plato |
| **Pan** | Kilos de pan sobrante del día |

---

## Opción "No Considera Mermas"

Cuando se activa el checkbox **"No considera Mermas"**:

- Todos los valores de la grilla (merma x ración, merma x kilo) se ponen en cero y se bloquean.
- Los campos Producción, Desconche y Pan se deshabilitan.
- Se graba el indicador `Considera_Merma = 0` en `b_mermadesconche`.
- El proceso de **Cierre Diario** lee este indicador y omite las mermas de ese día en sus cálculos.

---

## Relación entre Merma x Ración y Merma x Kilo

```
                    RECETA
                       │
          ┌────────────┴────────────┐
          │                         │
   Merma x Ración            Merma x Kilo
  (raciones no vendidas)    (kilos no vendidos)
          │                         │
          │   Son equivalentes:     │
          └────────────┬────────────┘
                       │
        se convierten usando el Gramaje Servido
        de la receta y el factor de conversión
```

Ambos campos deben ingresarse siempre en conjunto. El sistema puede calcular uno a partir del otro, pero requiere que al grabar ambos sean coherentes.

---

## Dónde se Almacena

| Dato | Tabla | Campo |
|---|---|---|
| Raciones no vendidas por receta | `b_minutadet` | `mid_nummer` |
| Kilo Servido por receta | `b_minutadet` | `mid_mermaxcantservida` |
| Kilo Bruto por receta | `b_minutadet` | `mid_mermaxkilo` |
| Merma Producción del día | `b_mermadesconche` | `Merma_Produccion` |
| Merma Desconche del día | `b_mermadesconche` | `Merma_Desconche` |
| Merma Pan del día | `b_mermadesconche` | `Merma_Pan` |
| Indicador "sin mermas" | `b_mermadesconche` | `Considera_Merma` |

---

## Flujo del Proceso

```
1. Usuario ingresa: Contrato + Régimen + Servicio + Fecha
        │
        ▼
2. Sistema valida que exista minuta real para ese día
        │
        ▼
3. Se carga la grilla con las recetas planificadas
   y sus valores ya ingresados (si existen)
        │
        ▼
4. Usuario ingresa Merma x Ración (raciones no vendidas)
        │
        ▼
5. Sistema calcula automáticamente:
   - Merma x Kilo Servido = GramajeServido × Raciones / Factor
   - Kilo Bruto = GramajeBruto × Raciones / Factor
   - Costo Merma = CostoReceta × Raciones
        │
        ▼
6. Usuario ingresa Merma Producción, Desconche y Pan (kilos globales)
        │
        ▼
7. Al grabar: se actualiza b_minutadet y b_mermadesconche
```

---

*Fuentes: `M_MerPre.frm`, `sgp_Sel_MermaPorPreparacion`, `sgp_Upd_XmlMermaPreparacion`, `SGP_FN_RNVCantidadesReceta`*
