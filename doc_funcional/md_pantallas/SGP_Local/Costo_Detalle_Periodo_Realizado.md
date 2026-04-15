# Costo Detalle Periodo Realizado

**Formulario:** `I_FCost.frm` (modo `CosPer`)
**Función principal:** `I_CostoDetPeriodoRealizado` en `Informes.bas`
**Tablas principales:** `b_totventas` (cabecera de salidas de producción), `b_detventas` (líneas de producto por salida), `b_minutaraciones` (raciones producidas por día)
**Consulta principal:** Consulta directa SQL (sin Stored Procedures)

---

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
- [6 — Referencia técnica](#6--referencia-técnica)
  - [Tablas que intervienen](#tablas-que-intervienen)
  - [Relación con otros módulos](#relación-con-otros-módulos)

---

## 1 — ¿Para qué sirve esta pantalla?

[↑ Volver al índice](#índice)

El informe **Costo Detalle Periodo Realizado** muestra el costo real de producción de un casino durante un período de tiempo, desglosado día a día y por sector de servicio (por ejemplo, almuerzo de régimen normal, cena de régimen vegetariano).

Para cada día y combinación de régimen y servicio, el informe lista todos los ingredientes y productos que efectivamente salieron de bodega para producción, con su costo unitario, la cantidad consumida y el costo total. Al final de cada sector calcula el **costo por sector** dividiendo el total entre las raciones producidas, y al final del día entrega el **costo total del día** y el **costo por ración del día**.

El informe refleja lo que realmente ocurrió (salidas tipo `SP` ya procesadas), no una estimación planificada. Sirve para que el responsable del casino compare el gasto real de materia prima contra las raciones servidas y detecte desviaciones de costo.

---

## 2 — ¿Qué necesito para usarla?

[↑ Volver al índice](#índice)

| Requisito | Detalle |
|-----------|---------|
| **Contrato** | El contrato (centro de costo) debe existir en la base de datos y tener salidas de producción (`SP`) registradas en el período. |
| **Período** | Fecha inicial y fecha final deben estar dentro del **mismo mes y año**. El sistema no permite períodos que crucen meses ni años. |
| **Salidas procesadas** | Deben existir salidas de bodega de tipo `SP` (Salida Producción) para el contrato, en el período y bodega activa. Las salidas no deben estar en estado `A` (Anulado) ni `P` (Pendiente). |
| **Raciones producidas** | Para que el cálculo de costo por ración sea significativo, deben existir raciones registradas con tipo `PRODUCIDAS` en la tabla de raciones de minuta. Si no existen, el sistema mostrará el costo total pero no calculará el costo unitario por ración. |
| **Régimen y servicio** | Se debe seleccionar al menos un régimen y un servicio, ya sea "Todos" o una lista específica. |

---

## 3 — ¿Cómo se usa?

[↑ Volver al índice](#índice)

### 3.1 Flujo paso a paso

[↑ Volver al índice](#índice)

```mermaid
flowchart TD
    A([Abrir formulario\nI_FCost en modo CosPer]) --> B[Ingresar código de contrato\no buscar con ícono]
    B --> C[Verificar nombre de contrato\nautocompletado]
    C --> D[Ajustar Fecha Inicial\ndd/mm/yyyy]
    D --> E[Ajustar Fecha Final\ndd/mm/yyyy]
    E --> F{¿Filtrar regímenes?}
    F -- Todos --> G[Dejar seleccionado Todos]
    F -- Lista -- > H[Seleccionar Lista\ny elegir regímenes]
    G --> I{¿Filtrar servicios?}
    H --> I
    I -- Todos --> J[Dejar seleccionado Todos]
    I -- Lista --> K[Seleccionar Lista\ny elegir servicios]
    J --> L[Clic en Vista Previa]
    K --> L
    L --> M{Validaciones\ndel sistema}
    M -- Error --> N[Mensaje de error\nCorregir y reintentar]
    M -- OK --> O[Sistema consulta\nb_totventas / b_detventas\nb_minutaraciones\nb_minuta / b_minutadet\nb_receta / b_recetadet]
    O --> P{¿Hay datos?}
    P -- No --> Q[Mensaje: No existe información\no salidas sin sector indicado]
    P -- Sí --> R[Genera reporte RTF\npor día, régimen y servicio]
    R --> S([Vista previa en pantalla\nExportable a RTF])
```

### 3.2 Controles y acciones disponibles

[↑ Volver al índice](#índice)

| Control | Descripción |
|---------|-------------|
| **Campo Contrato** | Código del contrato (centro de costo). Se puede digitar directamente o buscar haciendo clic en el ícono de búsqueda. Al ingresar un contrato válido, el nombre se completa automáticamente. |
| **Fecha Inicial** | Fecha de inicio del período a informar, en formato `dd/mm/yyyy`. Se inicializa con la fecha del día. |
| **Fecha Final** | Fecha de término del período a informar, en formato `dd/mm/yyyy`. Se inicializa con la fecha del día. Para ver un mes completo, ajustar al primer y último día del mes. |
| **Marco Régimen — Todos** | (Opción por defecto.) Incluye todos los regímenes disponibles para el contrato. |
| **Marco Régimen — Lista** | Permite seleccionar uno o más regímenes específicos mediante un buscador. |
| **Marco Servicio — Todos** | (Opción por defecto.) Incluye todos los servicios disponibles. |
| **Marco Servicio — Lista** | Permite seleccionar uno o más servicios específicos mediante un buscador. |
| **Botón Vista Previa** | Ejecuta las validaciones y, si son correctas, genera y muestra el informe en pantalla. El archivo también se exporta en formato RTF a la ruta configurada en `vg_reporte`. |
| **Botón Histórico Planificación Teórica** | Acceso a otro informe relacionado (planificación teórica), disponible en la misma barra de herramientas. |
| **Botón Salir** | Cierra el formulario. |

> **Nota:** En este modo (`CosPer`) las opciones de tipo de informe (Planif. Teórico, Planif. Real, Salida Prod.) están ocultas. El informe trabaja exclusivamente con salidas de producción reales.

---

## 4 — ¿Qué restricciones debo conocer?

[↑ Volver al índice](#índice)

### 4.1 Validaciones del sistema

[↑ Volver al índice](#índice)

Las siguientes validaciones se ejecutan al presionar **Vista Previa**, en el orden indicado:

| N° | Mensaje del sistema | Condición que lo genera | Cómo resolverlo |
|----|--------------------|-----------------------|-----------------|
| 1 | `No existe contrato` | El código de contrato ingresado no existe en la base de datos. | Verificar el código o buscarlo con el ícono. |
| 2 | `Fecha origen Mayor destino` | La Fecha Inicial es posterior a la Fecha Final. | Corregir las fechas para que el período sea válido. |
| 3 | `Mes origen mayor destino` | La Fecha Inicial y la Fecha Final pertenecen a meses distintos. | El informe solo cubre un mes completo o parcial; ajustar ambas fechas al mismo mes. |
| 4 | `Año origen mayor destino` | La Fecha Inicial y la Fecha Final pertenecen a años distintos. | Ajustar ambas fechas al mismo año. |
| 5 | `Regimen debe ser informado` | Se seleccionó la opción "Lista" para régimen pero no se eligió ninguno. | Seleccionar al menos un régimen o cambiar a "Todos". |
| 6 | `Servicio debe ser informado` | Se seleccionó la opción "Lista" para servicio pero no se eligió ninguno. | Seleccionar al menos un servicio o cambiar a "Todos". |
| 7 | `No existe información ó bien las salidas no tiene indicada la opción x sector` | No hay salidas de producción (`SP`) para los filtros seleccionados, o existen pero sin sector asignado en las líneas de detalle (`dev_codsec = 0`). | Verificar que el período tenga salidas procesadas y que cada línea tenga sector indicado. |

### 4.2 Reglas de cálculo

[↑ Volver al índice](#índice)

| Regla | Descripción |
|-------|-------------|
| **Selección de salidas** | Solo se consideran documentos de tipo `SP` (Salida Producción) que no estén en estado `A` (Anulado) ni `P` (Pendiente), y que pertenezcan a la bodega activa (`vg_codbod`). |
| **Raciones producidas** | Se obtienen de `b_minutaraciones` filtrando por `mir_rutcli = 'PRODUCIDAS'`. Si el valor es nulo o cero, los cálculos de costo por ración no se muestran (denominador en cero se evita con `IIf(NumRac > 0, ...)`). |
| **Costo total por sector** | Suma de `dev_ptotal` (importe total de cada línea) para todas las líneas del sector, considerando solo las que tienen `dev_canmer <> 0` (cantidad de merma distinta de cero). |
| **Costo por sector / ración** | `Σ dev_ptotal del sector ÷ NumRac` (raciones producidas del día). Se muestra solo si `NumRac > 0`. |
| **Total día** | Suma acumulada de `dev_ptotal` de todos los sectores del día. |
| **Costo día / ración** | `Total día ÷ NumRac`. Se muestra solo si `NumRac > 0`. |
| **Estructura Fija** | Los ingredientes sin código (`dev_coding` vacío o nulo) se agrupan bajo el concepto "Estructura Fija" y se muestran al final, con `sec_orden = 999999999` para que queden siempre últimos. |
| **Recetas del sector** | Para cada sector se listan las recetas planificadas asociadas (de `b_minuta`, `b_minutadet`, `b_receta`), obtenidas de la minuta real (`mid_tipmin = '2'`) con raciones planificadas mayores a cero. Esto permite comparar lo planificado con lo consumido. |
| **Orden de presentación** | Las líneas se ordenan por `sec_orden` (orden del sector) y `dev_numlin` (número de línea del documento), garantizando consistencia entre salidas. |

---

## 5 — ¿Qué obtengo?

[↑ Volver al índice](#índice)

El informe genera un **documento RTF** (orientación vertical, tamaño carta) con una página por cada combinación de régimen, servicio y fecha de producción. Cada página tiene la siguiente estructura:

**Encabezado de página:**
- Título: "Costos Detalle Período Realizado"
- Folio del documento de salida, contrato y nombre del casino
- Fecha de emisión y fecha de producción
- Bodega
- Régimen y servicio
- Raciones producidas del día

**Encabezado de columnas:**

| Columna | Descripción | Calculado |
|---------|-------------|-----------|
| Código | Código del producto consumido (`pro_codigo`) | No |
| Descripción | Nombre del producto (`pro_nombre`) | No |
| UN | Unidad de medida abreviada (`uni_nomcor`) | No |
| Costo Unit. | Precio unitario del producto en el documento de salida (`dev_predoc`) | No |
| Cantidad | Cantidad total consumida (`dev_canmer`, suma de líneas agrupadas) | Sí (SUM) |
| Costo Total | Importe total de la línea (`dev_ptotal`, suma de líneas agrupadas) | Sí (SUM) |

**Subtotales por sector:**

| Fila de subtotal | Valor |
|-----------------|-------|
| Nombre del sector (ej. "Almuerzo") | Σ Costo Total de las líneas del sector |
| Costo x Sector | Σ Costo sector ÷ Raciones producidas |

**Totales del día:**

| Fila de total | Valor |
|--------------|-------|
| Total Día | Σ todos los costos del día |
| Costo Día | Total Día ÷ Raciones producidas |

**Contexto de recetas:** Para cada sector se imprimen los nombres de las recetas planificadas en la minuta real, permitiendo saber qué preparaciones correspondían a los insumos listados.

**Formato de salida:** RTF exportado a `vg_reporte` (ruta configurada globalmente) y también visualizado en pantalla mediante el visor integrado `VSPrinter`.

---

## 6 — Referencia técnica

[↑ Volver al índice](#índice)

### Tablas que intervienen

[↑ Volver al índice](#índice)

| Tabla | Rol en el informe | Campos clave usados |
|-------|------------------|---------------------|
| `b_totventas` | Cabecera de los documentos de salida de producción. Cada fila representa un documento `SP` (Salida Producción) por contrato, bodega, régimen, servicio y fecha. | `tov_rutcli`, `tov_tipdoc='SP'`, `tov_numdoc`, `tov_codbod`, `tov_fecemi`, `tov_fecpro`, `tov_codreg`, `tov_codser`, `tov_estdoc` |
| `b_detventas` | Líneas de detalle del documento de salida. Cada fila es un producto/ingrediente entregado. | `dev_rutcli`, `dev_tipdoc`, `dev_numdoc`, `dev_numlin`, `dev_coding` (ingrediente), `dev_codmer` (producto), `dev_canmin`, `dev_canmer`, `dev_predoc`, `dev_ptotal`, `dev_codsec` |
| `b_minutaraciones` | Registro de raciones por día, régimen, servicio y tipo de comensal. La fila con `mir_rutcli = 'PRODUCIDAS'` contiene las raciones efectivamente producidas. | `mir_cencos`, `mir_codreg`, `mir_codser`, `mir_fecmin`, `mir_rutcli`, `mir_nrorac` |
| `b_minuta` | Cabecera de la minuta de planificación. Vincula un período con un contrato, régimen y servicio. | `min_codigo`, `min_cencos`, `min_codreg`, `min_codser`, `min_fecmin` |
| `b_minutadet` | Detalle de la minuta: recetas planificadas para cada día y servicio. | `mid_codigo`, `mid_codrec`, `mid_tiprec`, `mid_tipmin='2'` (minuta real), `mid_numrac`, `mid_estser` |
| `b_receta` | Maestro de recetas del casino. | `rec_codigo`, `rec_nombre` |
| `b_recetadet` | Detalle de la receta: relación receta–ingrediente y tipo. | `red_codigo`, `red_tiprec`, `red_cencos`, `red_canpro` |
| `a_regimen` | Maestro de regímenes alimenticios (ej. Normal, Vegetariano, Hiposódico). | `reg_codigo`, `reg_nombre` |
| `a_servicio` | Maestro de servicios de comida (ej. Desayuno, Almuerzo, Cena). | `ser_codigo`, `ser_nombre` |
| `a_estservicio` | Relación entre servicio, contrato y sector. Vincula un servicio con el sector del casino al que corresponde. | `ess_cencos`, `ess_codigo`, `ess_codsec` |
| `a_sector` | Maestro de sectores del casino (agrupaciones físicas o funcionales de los servicios). | `sec_codigo`, `sec_nombre`, `sec_orden` |
| `b_ingrediente` | Maestro de ingredientes. Cada ingrediente es la especificación técnica (ej. "Pollo entero"). | `ing_codigo`, `ing_nombre`, `ing_unimed` |
| `b_productos` | Maestro de productos de bodega. Cada producto es el artículo físico almacenado (ej. "Pollo entero congelado 1kg"). | `pro_codigo`, `pro_nombre`, `pro_coduni`, `pro_facing` |
| `a_unidadmed` | Unidades de medida del ingrediente (ej. kg, lt, unidad). | `unm_codigo`, `unm_nomcor` |
| `a_unidad` | Unidades de compra/bodega del producto (ej. caja, saco, bolsa). | `uni_codigo`, `uni_nomcor` |
| `[usuario]_tmp_CostoxPeriodo` | Tabla temporal creada en tiempo de ejecución con los documentos `SP` del período. Se utiliza para obtener la lista de folios únicos a procesar, luego se itera sobre ellos. Se elimina al inicio si ya existe. | `reg_codigo`, `reg_nombre`, `ser_codigo`, `ser_nombre`, `tov_fecemi`, `tov_fecpro`, `tov_numdoc`, `dev_codsec` |

### Relación con otros módulos

[↑ Volver al índice](#índice)

| Módulo relacionado | Tipo de relación |
|-------------------|-----------------|
| **Salida Bodega Producción** | Es el módulo que genera los documentos `SP` en `b_totventas` y `b_detventas` que este informe consulta. Sin salidas registradas no hay datos que mostrar. |
| **Planificación / Minuta** | Proporciona el contexto de las recetas planificadas (`b_minuta`, `b_minutadet`, `b_receta`, `b_recetadet`) que se imprimen junto al detalle de costos para facilitar la comparación. |
| **Raciones Producidas** | La fila `PRODUCIDAS` en `b_minutaraciones` es el único denominador para el costo por ración. Es gestionada por el formulario de cierre diario o de producción; su ausencia deja los campos de costo unitario en blanco. |
| **Mantenedores externos** | Los maestros de régimen, servicio, sector, ingrediente y producto son gestionados desde los módulos de Contrato y Régimen; el módulo de Producción solo los consulta. |
| **Informe Costo x Sector** | Informe hermano del mismo formulario `I_FCost.frm` (modo `CosSec`). Mientras este informe desglosa por período y folio, el de Costo x Sector agrupa por sector de forma resumida. |
| **Histórico Planificación Teórica** | Accesible desde la misma barra de herramientas. Permite comparar el costo teórico planificado con el costo real que muestra este informe. |

---

*Fuentes: `I_FCost.frm`, función `I_CostoDetPeriodoRealizado` en `Informes.bas`, tablas `b_totventas`, `b_detventas`, `b_minutaraciones`, `b_minuta`, `b_minutadet`, `b_receta`, `b_recetadet`, `a_regimen`, `a_servicio`, `a_sector`, `a_estservicio`, `b_ingrediente`, `b_productos`, `a_unidadmed`, `a_unidad` en `SGP_Local.sql`*
