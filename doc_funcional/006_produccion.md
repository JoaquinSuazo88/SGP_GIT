# Módulo: Producción

> **Estado de documentación:** `Borrador automático`  
> **Fuente:** Generado desde análisis de 7 sesiones de reunión  
> **Última actualización:** 2026-02-27  
> **Pendiente de validación por:** Responsable técnico | Responsable de negocio

---

## 1. Descripción General

## Descripción General del Módulo "Producción"

El módulo **Producción** gestiona el registro y control de las actividades productivas diarias en sitios de alimentación (casinos), abarcando la planificación de minutas, el ingreso de cantidades producidas, el control de raciones, mermas y salidas de producción, así como la generación de reportes comparativos entre lo planificado y lo realizado (reporte de aportes). Es utilizado principalmente por **operadores de sitio** para el registro diario, y por **administradores centrales** para configuración de parámetros, desbloqueo de registros y supervisión de reportes. El módulo se posiciona en el **núcleo del flujo operativo diario**: actúa como requisito previo para el cierre diario, condiciona la impresión de requisiciones y alimenta los reportes de costos y diferencias entre producción y venta.

> *Nota: No se dispone de información sobre elementos de base de datos ni sobre roles intermedios (supervisores de zona), lo que limita la descripción de la arquitectura de acceso por perfil.*

---

## 2. Reglas de Negocio

> Las reglas marcadas con 🔴 son explícitas (declaradas directamente en la reunión).
> Las marcadas con 🟡 son inferidas (deducidas del contexto).

**Total:** 122 reglas (105 explícitas · 17 inferidas)

### Cierre diario y bloqueos operativos

| # | Regla | Certeza | Fuente |
|---|-------|:-------:|--------|
| 1 | El parámetro 'servicio principal' permite seleccionar un servicio (ej. almuerzo) y marcarlo como obligatorio para el ingreso de cantidades producidas (Q producidas) en el sitio. | 🔴 Explícita | [Sesión 03a — 00:04:14] |
| 2 | Si un servicio está marcado como obligatorio para Q producidas y el sitio no las ingresa en el día, el sistema bloquea el cierre diario impidiendo continuar con otras tareas. | 🔴 Explícita | [Sesión 03a — 00:04:14] |
| 3 | Las cantidades producidas (Q producidas) se bloquean en la planificación después de 72 horas, impidiendo su edición posterior. | 🔴 Explícita | [Sesión 03a — 00:04:14] |
| 4 | Si las Q producidas quedan en blanco y se superan las 72 horas de bloqueo, el sitio no puede retroactivamente ingresarlas; quedan en cero y solo el administrador central puede desbloquear desde el mantenedor. | 🔴 Explícita | [Sesión 03a — 00:06:12] |
| 5 | Las actividades diarias configurables como obligatorias incluyen: salida de producción, control de raciones, mermas y raciones no vendidas. | 🔴 Explícita | [Sesión 03a — 00:06:12] |
| 6 | Algunas actividades diarias son bloqueantes para el cierre diario y otras solo generan un mensaje de advertencia permitiendo continuar. | 🔴 Explícita | [Sesión 03a — 00:06:12] |
| 7 | La versión 2.27 del sistema incorpora la restricción de que si el sitio no ha ingresado las Q producidas, no puede imprimir las requisiciones. | 🔴 Explícita | [Sesión 03a — 00:08:09] |
| 8 | Las Q producidas afectan los reportes de planificación y los cálculos de diferencia entre producido y vendido; si se quedan en cero los reportes se ven impactados. | 🔴 Explícita | [Sesión 03a — 00:08:09] |
| 9 | La versión 2.27 con la restricción de impresión de requisiciones requiere que el parámetro de servicio principal esté configurado para que la lógica funcione correctamente. | 🔴 Explícita | [Sesión 03a — 00:08:09] |
| 10 | La versión 2.27 fue aprobada y piloteada pero no llegó a ponerse en producción para todos los sitios antes del proceso de migración. | 🔴 Explícita | [Sesión 03a — 00:08:09] |
| 11 | El cierre diario es un paso manual que realiza el operador del sitio posicionándose sobre el día en el mantenedor correspondiente; al cerrarlo el estado cambia de color. | 🔴 Explícita | [Sesión 03a — 02:45:23] |

### Planificación y cantidades producidas

| # | Regla | Certeza | Fuente |
|---|-------|:-------:|--------|
| 1 | Los sitios tienen habilitada la opción de agregar una receta en líneas en blanco de la planificación, haciendo doble clic para abrir el maestro de recetas y asignar cantidad. | 🔴 Explícita | [Sesión 03a — 00:38:27] |
| 2 | Para reemplazar una receta, el usuario del sitio debe poner en cero la receta original y agregar la nueva en una línea en blanco; no existe una función de reemplazo directo. | 🔴 Explícita | [Sesión 03a — 00:38:27] |
| 3 | El número de raciones planificadas y el número de raciones realmente cocinadas pueden diferir, y esta diferencia debería reflejarse en el campo Q del reporte comparativo. | 🔴 Explícita | [Sesión 03a — 00:38:27] |
| 4 | En las salidas de producción, el sistema muestra por defecto la cantidad planificada para cada producto con una aproximación de ±5 unidades respecto al valor exacto. | 🔴 Explícita | [Sesión 03a — 01:40:30] |
| 5 | El bodeguero puede sobrescribir la cantidad por defecto (planificada) e ingresar la cantidad real entregada al momento de registrar la salida. | 🔴 Explícita | [Sesión 03a — 01:40:30] |
| 6 | Si la planificación y los pasos previos están correctos, las salidas de producción deberían confirmarse por defecto sin necesidad de ajustes significativos. | 🔴 Explícita | [Sesión 03a — 01:40:30] |
| 7 | El campo 'producidas' en el control de raciones proviene de la planificación y no es editable por el usuario. | 🔴 Explícita | [Sesión 03a — 02:13:36] |
| 8 | El campo 'programado' corresponde a la cantidad de raciones planificadas por el chef, que puede modificar lo que viene desde los planificadores; a este dato modificado se le denomina 'real'. | 🔴 Explícita | [Sesión 03a — 02:36:52] |
| 9 | El costo tiene cuatro estados secuenciales según el flujo de producción: costo planificado (administrador/SGP), costo del sitio (costo teórico), costo real y costo realizado. | 🔴 Explícita | [Sesión 04 — 03:32:36] |
| 10 | El costo planificado corresponde al nivel del administrador/SGP. | 🔴 Explícita | [Sesión 04 — 03:32:36] |
| 11 | El costo teórico corresponde al costo del sitio, generado cuando la minuta del administrador se pasa a la minuta del sitio. | 🔴 Explícita | [Sesión 04 — 03:32:36] |
| 12 | El costo real corresponde a lo que el sitio declara que va a sacar a producción. | 🔴 Explícita | [Sesión 04 — 03:32:36] |
| 13 | El costo realizado corresponde a lo que el sitio efectivamente produjo, considerando salidas adicionales y devoluciones. | 🔴 Explícita | [Sesión 04 — 03:32:36] |
| 14 | Los costos relevantes para mostrar en el sistema son el costo planificado (SGP administrador) y el costo del sitio (costo local). | 🔴 Explícita | [Sesión 04 — 03:32:36] |
| 15 | La estructura de la minuta se mantiene igual al pasar del administrador al sitio, solo cambia la denominación del costo. | 🔴 Explícita | [Sesión 04 — 03:32:36] |
| 16 | El sistema registra tres métricas de raciones: planificada, producida y vendida, permitiendo comparar brechas entre producción y entrega real. | 🔴 Explícita | [Sesión 04 — 03:43:32] |
| 17 | El sistema compara lo producido contra lo efectivamente entregado (ya sea personal u otra categoría) y esa diferencia es la brecha que el planificador ajusta mes a mes. | 🔴 Explícita | [Sesión 04 — 03:43:32] |

### Salidas de producción y requisiciones

| # | Regla | Certeza | Fuente |
|---|-------|:-------:|--------|
| 1 | La vista detallada de la requisición sirve para el área de producción/chef, mientras que la resumida sirve para bodega para preparar el despacho de productos. | 🔴 Explícita | [Sesión 03a — 00:57:07] |
| 2 | Los adicionales (consumos no planificados) actualmente no generan una requisición formal; se registran en papeles o formularios impresos informalmente fuera del sistema. | 🔴 Explícita | [Sesión 03a — 00:57:07] |
| 3 | Los adicionales deben generar una salida de producción y requieren ser registrados en el sistema para no quedar fuera del control de consumo. | 🔴 Explícita | [Sesión 03a — 00:57:07] |
| 4 | Los adicionales deben registrarse como una salida identificada y separada de la planificación, con trazabilidad para poder justificar consumos no planificados. | 🔴 Explícita | [Sesión 03a — 00:57:07] |
| 5 | Los adicionales deben estar vinculados obligatoriamente a un servicio y régimen específico para garantizar trazabilidad y correcta imputación. | 🔴 Explícita | [Sesión 03a — 00:59:07] |
| 6 | El sistema actualmente exige seleccionar régimen y servicio al registrar una salida extra o adicional. | 🔴 Explícita | [Sesión 03a — 00:59:07] |
| 7 | Cuando los usuarios registran adicionales en un servicio incorrecto (por conveniencia o error), los reportes de desviación muestran datos erróneos atribuidos al servicio equivocado. | 🔴 Explícita | [Sesión 03a — 00:59:07] |
| 8 | La estructura fija de servicio es una funcionalidad que anteriormente se usaba para ingresar desechables y alcuzas como salida aparte, pero ya no se utiliza. | 🔴 Explícita | [Sesión 03a — 00:59:07] |
| 9 | Las salidas de producción se cargan seleccionando régimen y servicio, y el sistema muestra por defecto los productos y cantidades que provienen de la planificación. | 🔴 Explícita | [Sesión 03a — 01:38:32] |
| 10 | Las salidas de producción pueden visualizarse en modo resumido o detallado; el modo detallado presenta los productos agrupados por sector según la estructura definida. | 🔴 Explícita | [Sesión 03a — 01:38:32] |
| 11 | Las salidas por sector permiten generar informes de costos desagregados por cada sector de la minuta (sopa, ensalada, plato de fondo, acompañamiento, postre). | 🔴 Explícita | [Sesión 03a — 01:38:32] |
| 12 | Si una estructura de minuta no está sectorizada, el sistema no permite realizar las salidas por sector. | 🔴 Explícita | [Sesión 03a — 01:38:32] |
| 13 | El registro de mermas y salidas se realiza día a día y servicio por servicio, guardándose de forma incremental. | 🔴 Explícita | [Sesión 03a — 02:36:52] |
| 14 | Las salidas generadas por eventos especiales se suman actualmente a las salidas de producción regular, sin diferenciación en los reportes de costo. | 🔴 Explícita | [Sesión 03a — 02:45:23] |
| 15 | Se propone que las salidas de eventos especiales se identifiquen como una categoría separada dentro de las salidas de producción para facilitar el análisis de costos. | 🟡 Inferida | [Sesión 03a — 02:45:23] |
| 16 | El reporte 'Insumos no planificados en salida a bodega' muestra por servicio el total de la salida y los insumos utilizados que no estaban planificados, permitiendo identificar cambios de sabor o producto. | 🔴 Explícita | [Sesión 03b — 00:39:04] |
| 17 | El reporte presenta un defecto: cuando se usa más de un producto (planificado y no planificado), suma el total como 'no planificado', en lugar de mostrar solo la diferencia (el extra utilizado). | 🔴 Explícita | [Sesión 03b — 00:39:04] |
| 18 | El reporte muestra tanto lo que se sacó sin estar planificado como lo que estaba planificado y no fue utilizado, requiriendo cruces de análisis para interpretarlo correctamente. | 🔴 Explícita | [Sesión 03b — 00:39:04] |
| 19 | El reporte permite filtrar por servicio del régimen y muestra el costo total de la salida, indicando cuánto correspondía a insumos planificados o no planificados. | 🔴 Explícita | [Sesión 03b — 00:39:04] |

### Control de raciones y venta

| # | Regla | Certeza | Fuente |
|---|-------|:-------:|--------|
| 1 | El control de raciones permite registrar diariamente la cantidad vendida por cliente; actualmente se digita manualmente pero se integrará desde el sistema SPRS. | 🔴 Explícita | [Sesión 03a — 02:13:36] |
| 2 | El maestro de clientes está abierto en el sitio permitiendo crear clientes genéricos o ficticios sin validación, lo que genera datos no controlados. | 🔴 Explícita | [Sesión 03a — 02:13:36] |
| 3 | Las muestras de referencia en el control de raciones equivalen aproximadamente a tres bandejas. | 🔴 Explícita | [Sesión 03a — 02:13:36] |
| 4 | El total de raciones vendidas por cliente más el personal más las muestras debe ser cercano o igual a las raciones producidas planificadas. | 🔴 Explícita | [Sesión 03a — 02:13:36] |
| 5 | La suma de raciones por cliente más personal más muestras debe cuadrar con el total de raciones producidas planificadas. | 🔴 Explícita | [Sesión 03a — 02:15:36] |
| 6 | Con la integración con SPRS, los clientes del control de raciones provendrán del sistema externo y no podrán crearse clientes o RUT ficticios. | 🔴 Explícita | [Sesión 03a — 02:15:36] |
| 7 | Todo cliente que haya sido facturado en la compañía está registrado en SPRS como venta. | 🔴 Explícita | [Sesión 03a — 02:15:36] |
| 8 | El cliente en el control de raciones se identifica mediante un RUT (root/mandante); con la integración SPRS, este dato vendría validado del sistema externo. | 🟡 Inferida | [Sesión 03a — 02:15:36] |
| 9 | En el control de raciones, el mandante aparece primero y luego se listan los contratistas asociados a ese mandante. | 🔴 Explícita | [Sesión 03a — 02:17:36] |
| 10 | Solo los contratistas a los que se les vendió a través del sistema de venta de contratista deben aparecer como clientes en el listado. | 🔴 Explícita | [Sesión 03a — 02:17:36] |
| 11 | La cantidad de raciones corresponde a los vales quemados (consumidos) contados para el mandante y cada contratista. | 🔴 Explícita | [Sesión 03a — 02:17:36] |
| 12 | El módulo de control de raciones permanece siempre abierto (no se bloquea) hasta el cierre de mes, permitiendo correcciones posteriores. | 🔴 Explícita | [Sesión 03a — 02:17:36] |
| 13 | El usuario puede ingresar una venta teórica (estimado) de raciones y luego corregirla tras validación con el cliente antes de facturar. | 🔴 Explícita | [Sesión 03a — 02:17:36] |
| 14 | Se requiere una integración entre el SGP y el sistema de venta de contratista para obtener los vales quemados. | 🔴 Explícita | [Sesión 03a — 02:17:36] |
| 15 | La cantidad de raciones final a facturar (después de negociación con el cliente) se registra en el sistema SPRS, no en el SGP. | 🔴 Explícita | [Sesión 03a — 02:19:49] |
| 16 | El módulo 'Venta Contado' fue creado originalmente para registrar ventas directas sin facturación a clientes, ingresando montos totales en pesos. | 🔴 Explícita | [Sesión 03a — 02:19:49] |
| 17 | El ingreso de venta contado como monto total sin desglose de raciones provoca pérdida de visibilidad de las raciones. | 🔴 Explícita | [Sesión 03a — 02:19:49] |
| 18 | Existen registros de venta contado creados sin precio de venta asociado, lo que se considera un problema del estado actual del sistema. | 🔴 Explícita | [Sesión 03a — 02:21:51] |
| 19 | Para los servicios de pago contado de alimentación se debe obligar a ingresar tanto la cantidad (Q) como el precio (PE) en el módulo correspondiente. | 🔴 Explícita | [Sesión 03a — 02:21:51] |
| 20 | El control de raciones solo debe registrar cantidades (Q), y la venta contado debe manejarse en un módulo separado que capture cantidad y precio. | 🟡 Inferida | [Sesión 03a — 02:21:51] |
| 21 | La venta anticipada de contratista debería reducir significativamente el uso del módulo de venta contado. | 🟡 Inferida | [Sesión 03a — 02:21:51] |
| 22 | El formulario de control de raciones debe registrar únicamente raciones (Q), y la venta contado debería alimentarlo desde un módulo separado. | 🟡 Inferida | [Sesión 03a — 02:23:48] |
| 23 | Para los servicios tipo 'estar médico' (salud), no es posible ingresar un precio por ración (P/Q) porque se venden productos individuales de precio variable. | 🔴 Explícita | [Sesión 03a — 02:23:48] |
| 24 | Se propone parametrizar un precio estándar por ración para un sitio, permitiendo que el operador lo modifique si la venta fue a un precio distinto. | 🟡 Inferida | [Sesión 03a — 02:23:48] |
| 25 | La salida de bodega incluye detalle de productos a precio costo, y se propone vincularlo al ingreso de venta aplicándole precio de venta. | 🟡 Inferida | [Sesión 03a — 02:23:48] |
| 26 | En servicios de salud (estares médicos), los movimientos de productos se registran a través de la salida de producción, pero la venta no puede ingresarse como P/Q. | 🔴 Explícita | [Sesión 03a — 02:25:48] |
| 27 | Anteriormente se creaban clientes ficticios (productos) para detallar la venta de productos en estares médicos, lo que generó una proliferación de clientes incorrectos en el sistema. | 🔴 Explícita | [Sesión 03a — 02:25:48] |
| 28 | Actualmente, la única forma de ingresar la venta de un estar médico es registrar el total vendido del día para ese servicio, sin detalle por producto. | 🔴 Explícita | [Sesión 03a — 02:25:48] |
| 29 | Se propone vincular la salida de bodega al ingreso de venta para los casos de estares médicos, trayendo el detalle de productos y asignando precio de venta. | 🟡 Inferida | [Sesión 03a — 02:25:48] |
| 30 | En algunos sitios del segmento salud, la venta de productos en estares médicos puede representar entre el 20% y 30% del total de ventas del sitio. | 🔴 Explícita | [Sesión 03a — 02:25:48] |
| 31 | Los productos vendidos en estares médicos no son 100% planificables porque dependen de la demanda diaria variable del cliente. | 🔴 Explícita | [Sesión 03a — 02:25:48] |
| 32 | Las raciones pueden clasificarse como 'personal' o 'vendida', y también existe la categoría 'muestra' o 'referencia'. | 🟡 Inferida | [Sesión 04 — 03:43:32] |
| 33 | Los vales de venta generan salidas de producción, pero los ingresos asociados se registran previamente; el control requiere el vale quemado para efectos de trazabilidad. | 🔴 Explícita | [Sesión 04 — 03:43:32] |
| 34 | El cliente en el sistema de ventas debería representarse con un RUT ficticio denominado 'vales' para consolidar todos los consumos bajo ese concepto, sin necesidad de identificar al consumidor individual. | 🟡 Inferida | [Sesión 04 — 03:43:32] |

### Mermas, desconche y raciones no vendidas

| # | Regla | Certeza | Fuente |
|---|-------|:-------:|--------|
| 1 | Las raciones no vendidas se registran como mermas de línea o mermas de preparación dentro del módulo de producción. | 🔴 Explícita | [Sesión 03a — 02:27:59] |
| 2 | Las mermas de producción y las mermas de desconcha se registran también en el mismo módulo de producción. | 🔴 Explícita | [Sesión 03a — 02:27:59] |
| 3 | La merma de producción se registra como un monto total en kilos, sin desglose por tipo de producto. | 🔴 Explícita | [Sesión 03a — 02:27:59] |
| 4 | Para registrar merma de producción detallada por producto, sería necesario crear un producto específico por cada tipo de merma en el sistema. | 🟡 Inferida | [Sesión 03a — 02:27:59] |
| 5 | Se propone agrupar las mermas de producción por familias de productos (ej. proteínas, verduras) en lugar de producto individual, para facilitar el ingreso. | 🟡 Inferida | [Sesión 03a — 02:30:04] |
| 6 | Se identifican al menos dos categorías de merma de producción: merma natural de materia prima (limpieza de carnes y verduras) y merma por mala cocción o mala manipulación. | 🔴 Explícita | [Sesión 03a — 02:30:04] |
| 7 | Se requeriría crear un mantenedor en el administrador para gestionar los motivos o grupos de merma de producción. | 🟡 Inferida | [Sesión 03a — 02:30:04] |
| 8 | La merma de desconcha se registra en kilos como valor total, sin distinción de tipo de residuo. | 🔴 Explícita | [Sesión 03a — 02:32:07] |
| 9 | El sistema anterior (FLMS) separaba la merma de desconcha en biológica y no biológica. | 🔴 Explícita | [Sesión 03a — 02:32:07] |
| 10 | La capacidad de separar residuos biológicos y no biológicos en el desconche depende de si el sitio tiene autodesconche (separación realizada por el propio cliente). | 🔴 Explícita | [Sesión 03a — 02:32:07] |
| 11 | Se propone registrar tres valores en la merma de desconcha: biológico, no biológico y pan; el campo no biológico no sería obligatorio. | 🟡 Inferida | [Sesión 03a — 02:32:07] |
| 12 | La separación del pan en la merma de desconcha es opcional según la capacidad operativa del sitio. | 🔴 Explícita | [Sesión 03a — 02:32:07] |
| 13 | El pan debería registrarse separado de las raciones no vendidas en la merma de desconcha. | 🔴 Explícita | [Sesión 03a — 02:34:04] |
| 14 | La merma (de desconcha, producción y raciones no vendidas) siempre se registra por servicio, no como merma general del día. | 🔴 Explícita | [Sesión 03a — 02:34:04] |
| 15 | Se propone usar un mantenedor aparte para configurar si los tres valores de desconcha (biológico, no biológico, pan) van juntos o separados según el sitio. | 🟡 Inferida | [Sesión 03a — 02:34:04] |
| 16 | Las mermas de producción, desconcha y raciones no vendidas se registran por servicio; la única merma no registrada por servicio es la de bodega. | 🔴 Explícita | [Sesión 03a — 02:34:55] |
| 17 | La merma de desconcha (EFAN) también se registra por servicio. | 🔴 Explícita | [Sesión 03a — 02:34:55] |
| 18 | La pantalla de raciones no vendidas muestra las recetas planificadas, la cantidad planificada, el costo unitario y el costo total de lo planificado. | 🔴 Explícita | [Sesión 03a — 02:34:55] |
| 19 | Las mermas de raciones no vendidas se pueden ingresar por raciones o por kilo según lo que sea más práctico para el sitio o el producto. | 🔴 Explícita | [Sesión 03a — 02:34:55] |
| 20 | Al ingresar la merma en kilos, el sistema calcula automáticamente el costo de la merma y la conversión a raciones equivalentes. | 🔴 Explícita | [Sesión 03a — 02:34:55] |
| 21 | Al ingresar la merma por unidad, el sistema acepta directamente el valor en unidades sin conversión adicional. | 🔴 Explícita | [Sesión 03a — 02:34:55] |
| 22 | La unidad de medida utilizada en el sistema es siempre el kilogramo (kilo), sin admitir otras unidades, ya que todos los reportes e informes se expresan en esa unidad. | 🔴 Explícita | [Sesión 03a — 02:36:52] |
| 23 | Las mermas de producción y desconches se registran solo en kilos por los sitios en campos específicos del módulo de mermas de línea | 🔴 Explícita | [Sesión 04 — 03:38:38] |
| 24 | Actualmente solo se utiliza la información de mermas para verificar cobertura; no se realiza análisis de volumen (toneladas) con dichos datos | 🔴 Explícita | [Sesión 04 — 03:38:38] |
| 25 | El reporte de estaciones no vendidas incluye los campos: seco, régimen, servicio, descripción, fecha de minuta, periodo, descripción de receta, programado, costo, costo total, merma, gramos frutos y cantidad servida | 🔴 Explícita | [Sesión 04 — 03:38:38] |

### Servicios especiales y adicionales

| # | Regla | Certeza | Fuente |
|---|-------|:-------:|--------|
| 1 | El módulo de ventas/servicios especiales fue creado para registrar salidas de eventos no planificados, resolviendo el problema de sumar esas salidas a servicios regulares como el almuerzo. | 🔴 Explícita | [Sesión 03a — 02:42:58] |
| 2 | En el módulo de servicios especiales se puede registrar el evento con número de raciones o con monto total de venta; ambas formas son válidas. | 🔴 Explícita | [Sesión 03a — 02:42:58] |
| 3 | El cierre del evento especial se realiza mediante un candado en la interfaz; solo al cerrarlo se generan los movimientos de stock en bodega, el costo asociado y la venta. | 🔴 Explícita | [Sesión 03a — 02:42:58] |
| 4 | En el estado de resultados, los eventos especiales del mes se muestran como un total sumado, no detallados por evento individual. | 🔴 Explícita | [Sesión 03a — 02:42:58] |
| 5 | Existe un reporte de detalle (footcos) donde se puede ver el desglose de cada evento especial individualmente. | 🔴 Explícita | [Sesión 03a — 02:42:58] |
| 6 | El módulo de servicios especiales permite registrar devoluciones de productos de un evento ya cerrado, revirtiendo parcialmente el movimiento de stock. | 🔴 Explícita | [Sesión 03a — 02:45:23] |

### Reportes de aportes y costos

| # | Regla | Certeza | Fuente |
|---|-------|:-------:|--------|
| 1 | El reporte de aportes compara lo planificado (raciones y costo estimado) versus lo realizado (lo que realmente se sacó de bodega) para cada día del mes, mostrando el costo por bandeja planificado y el real. | 🔴 Explícita | [Sesión 03a — 00:36:09] |
| 2 | El acumulado del reporte de aportes se calcula dinámicamente según el día seleccionado, sumando todos los días hasta ese punto del mes. | 🔴 Explícita | [Sesión 03a — 00:36:09] |
| 3 | El reporte de aportes no puede ser editado desde la vista del administrador; la línea de datos corresponde a lo registrado en el sitio. | 🔴 Explícita | [Sesión 03a — 00:36:09] |
| 4 | Si hay cambios frecuentes en las minutas por parte de los sitios, se debe verificar si dichos cambios afectan la frecuencia de uso de recetas o cárnicos. | 🟡 Inferida | [Sesión 03a — 00:36:09] |
| 5 | En el reporte de aportes, el costo bandeja acumulado se calcula como promedio ponderado (total costo / total raciones acumuladas), mientras que materia prima, costo total y raciones son sumatoria directa de los días. | 🔴 Explícita | [Sesión 03a — 00:38:27] |
| 6 | El separador decimal del sistema debe estandarizarse a coma (,) y el separador de miles a punto (.), conforme al estándar chileno. | 🔴 Explícita | [Sesión 03a — 01:40:30] |
| 7 | Los mantenedores del módulo de producción comparten los mismos iconos de acción (imprimir, visualizar, incluir) que el módulo administrador. | 🔴 Explícita | [Sesión 03a — 02:36:52] |
| 8 | La mayoría de los mantenedores del sistema cuentan con un botón de imprimir, incluyendo el de mermas. | 🔴 Explícita | [Sesión 03a — 02:38:49] |
| 9 | Los informes del módulo de reportes muestran de forma separada todo lo registrado en los mantenedores de producción (mermas, salidas, etc.). | 🔴 Explícita | [Sesión 03a — 02:38:49] |
| 10 | Se requiere trazabilidad y aprobación para las salidas y mermas significativas, de modo que no puedan realizarse sin respaldo ni validación de un supervisor. | 🟡 Inferida | [Sesión 03a — 02:38:49] |


---

## 3. Pantallas y Formularios

### ControlDeRaciones

**Tipo:** formulario  
**Descripción:** Formulario para registrar diariamente las raciones vendidas por cliente, personal y muestras, y comparar con las producidas planificadas  
**Mencionado en:** [Sesión 03a — 02:13:36], [Sesión 03a — 02:17:36]

![ControlDeRaciones](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b061_01_ControlDeRaciones_021736.jpg)

*Capturas adicionales:*
- [sesion_03a — 02:13:36](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b059_01_ControlDeRaciones_021336.jpg)

### EstacionesNoVendidas

**Tipo:** reporte  
**Descripción:** Reporte que muestra estaciones sin venta con campos de seco, régimen, servicio, fecha minuta, periodo, receta, programado, costo, merma, gramos y cantidad servida  
**Mencionado en:** [Sesión 04 — 03:38:38]

![EstacionesNoVendidas](../../capturas/sesion_04/enfoque_a/sesion_04_ui_s04_b108_01_EstacionesNoVendidas_033838.jpg)

### EstructuraFijaServicio

**Tipo:** formulario  
**Descripción:** Módulo/pantalla que se usaba anteriormente para registrar salidas de desechables y alcuzas como planificación aparte; actualmente en desuso.  
**Mencionado en:** [Sesión 03a — 00:59:07]

![EstructuraFijaServicio](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b029_02_EstructuraFijaServicio_005907.jpg)

### FormularioAdicionales

**Tipo:** formulario  
**Descripción:** Formulario impreso (informal, fuera del sistema) donde algunos sitios anotan los consumos adicionales no planificados durante el servicio.  
**Mencionado en:** [Sesión 03a — 00:57:07]

![FormularioAdicionales](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b028_01_FormularioAdicionales_005707.jpg)

### FormularioDevolucionServiciosEspeciales

**Tipo:** formulario  
**Descripción:** Subformulario dentro del módulo de servicios especiales que permite registrar devoluciones de productos de un evento ya cerrado.  
**Mencionado en:** [Sesión 03a — 02:45:23]

![FormularioDevolucionServiciosEspeciales](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b075_01_FormularioDevolucionServiciosEspeciales_024523.jpg)

### FormularioIngresRacionesMerma

**Tipo:** formulario  
**Descripción:** Pantalla donde se registran las raciones no vendidas por preparación, indicando cantidad en kilos y total; se ingresa día a día y servicio por servicio.  
**Mencionado en:** [Sesión 03a — 02:36:52]

![FormularioIngresRacionesMerma](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b071_01_FormularioIngresRacionesMerma_023652.jpg)

### FormularioServiciosEspeciales

**Tipo:** formulario  
**Descripción:** Módulo para registrar salidas y ventas de eventos no planificados; permite ingresar productos usados, raciones o monto total; se cierra con un candado para generar movimientos de stock.  
**Mencionado en:** [Sesión 03a — 02:42:58]

![FormularioServiciosEspeciales](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b074_01_FormularioServiciosEspeciales_024258.jpg)

### ImpresionRequisiciones

**Tipo:** reporte  
**Descripción:** Impresión de requisiciones de insumos para bodega; en versión 2.27 se bloquea su impresión si no se han ingresado las Q producidas.  
**Mencionado en:** [Sesión 03a — 00:08:09]

![ImpresionRequisiciones](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b005_01_ImpresionRequisiciones_000809.jpg)

### IngresoMermasRacionesNoVendidas

**Tipo:** formulario  
**Descripción:** Formulario para ingresar mermas de raciones no vendidas por servicio, mostrando recetas planificadas, costo y permitiendo ingreso en raciones o kilos con conversión automática.  
**Mencionado en:** [Sesión 03a — 02:34:55]

![IngresoMermasRacionesNoVendidas](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b070_01_IngresoMermasRacionesNoVendidas_023455.jpg)

### MaestroClientes

**Tipo:** formulario  
**Descripción:** Mantenedor de clientes usado en el control de raciones; actualmente permite crear clientes y RUT ficticios sin validación  
**Mencionado en:** [Sesión 03a — 02:15:36]

![MaestroClientes](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b060_01_MaestroClientes_021536.jpg)

### MantenedorActividadesDiarias

**Tipo:** formulario  
**Descripción:** Sección dentro del mantenedor de casino que permite marcar como obligatorias actividades como salida de producción, control de raciones, mermas y raciones no vendidas para el cierre diario.  
**Mencionado en:** [Sesión 03a — 00:06:12]

![MantenedorActividadesDiarias](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b004_01_MantenedorActividadesDiarias_000612.jpg)

### MantenedorCasino

**Tipo:** formulario  
**Descripción:** Mantenedor general del casino donde se configuran parámetros del sitio, incluyendo actividades diarias obligatorias.  
**Mencionado en:** [Sesión 03a — 00:04:14]

![MantenedorCasino](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b003_02_MantenedorCasino_000414.jpg)

### MantenedorCierreDiario

**Tipo:** formulario  
**Descripción:** Pantalla donde el operador selecciona el día y ejecuta el cierre diario del sitio; al cerrarse, el registro cambia de color como indicador visual.  
**Mencionado en:** [Sesión 03a — 02:45:23]

![MantenedorCierreDiario](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b075_02_MantenedorCierreDiario_024523.jpg)

### MantenedorMermaProduccion

**Tipo:** formulario  
**Descripción:** Mantenedor propuesto para definir grupos o motivos de merma de producción (ej. proteínas, verduras, preparación fallida).  
**Mencionado en:** [Sesión 03a — 02:30:04]

![MantenedorMermaProduccion](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b067_01_MantenedorMermaProduccion_023004.jpg)

### MantenedorServicioPrincipal

**Tipo:** formulario  
**Descripción:** Permite seleccionar un servicio (ej. almuerzo) y configurarlo como obligatorio para el ingreso de Q producidas en el cierre diario del sitio.  
**Mencionado en:** [Sesión 03a — 00:04:14]

![MantenedorServicioPrincipal](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b003_01_MantenedorServicioPrincipal_000414.jpg)

### PopupMaestroReceta

**Tipo:** popup  
**Descripción:** Ventana emergente que se abre al hacer doble clic en una línea en blanco de la planificación, permite seleccionar una receta y asignarle cantidad.  
**Mencionado en:** [Sesión 03a — 00:38:27]

![PopupMaestroReceta](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b019_01_PopupMaestroReceta_003827.jpg)

### RacionesNoVendidas

**Tipo:** formulario  
**Descripción:** Pantalla para ingresar mermas de línea, de preparación, de producción, de desconcha y de pan dentro del módulo de producción.  
**Mencionado en:** [Sesión 03a — 02:27:59]

![RacionesNoVendidas](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b066_01_RacionesNoVendidas_022759.jpg)

### ReporteAportesDiarios

**Tipo:** reporte  
**Descripción:** Muestra comparativo día a día entre lo planificado y lo realizado (salidas de bodega), con costo bandeja y acumulado mensual. Visible tanto en el administrador como en los sitios.  
**Mencionado en:** [Sesión 03a — 00:36:09]

![ReporteAportesDiarios](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b018_01_ReporteAportesDiarios_003609.jpg)

### ReporteCostoPlanificado

**Tipo:** reporte  
**Descripción:** Reporte que muestra el costo planificado a nivel del administrador/SGP.  
**Mencionado en:** [Sesión 04 — 03:32:36]

![ReporteCostoPlanificado](../../capturas/sesion_04/enfoque_a/sesion_04_ui_s04_b105_01_ReporteCostoPlanificado_033236.jpg)

### ReporteCostoReal

**Tipo:** reporte  
**Descripción:** Reporte que muestra lo que el sitio declara que va a sacar a producción.  
**Mencionado en:** [Sesión 04 — 03:32:36]

![ReporteCostoReal](../../capturas/sesion_04/enfoque_a/sesion_04_ui_s04_b105_03_ReporteCostoReal_033236.jpg)

### ReporteCostoRealizado

**Tipo:** reporte  
**Descripción:** Reporte que muestra lo que el sitio efectivamente produjo, incluyendo salidas adicionales y devoluciones.  
**Mencionado en:** [Sesión 04 — 03:32:36]

![ReporteCostoRealizado](../../capturas/sesion_04/enfoque_a/sesion_04_ui_s04_b105_04_ReporteCostoRealizado_033236.jpg)

### ReporteCostoTeorico

**Tipo:** reporte  
**Descripción:** Reporte que muestra el costo del sitio (costo teórico), generado al pasar la minuta del administrador al sitio.  
**Mencionado en:** [Sesión 04 — 03:32:36]

![ReporteCostoTeorico](../../capturas/sesion_04/enfoque_a/sesion_04_ui_s04_b105_02_ReporteCostoTeorico_033236.jpg)

### ReporteDetalleEventosEspeciales

**Tipo:** reporte  
**Descripción:** Reporte que muestra el detalle de cada evento especial registrado en el módulo de servicios especiales.  
**Mencionado en:** [Sesión 03a — 02:42:58]

![ReporteDetalleEventosEspeciales](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b074_02_ReporteDetalleEventosEspeciales_024258.jpg)

### ReporteEstadoResultados

**Tipo:** reporte  
**Descripción:** Informe que muestra el resultado operacional incluyendo el total de ventas especiales del mes.  
**Mencionado en:** [Sesión 03a — 02:42:58]

![ReporteEstadoResultados](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b074_03_ReporteEstadoResultados_024258.jpg)

### ReporteInsumosNoPlanificadosSalidaBodega

**Tipo:** reporte  
**Descripción:** Muestra por servicio del régimen los insumos utilizados no planificados y su costo, así como los planificados no utilizados. Permite filtrar por servicio.  
**Mencionado en:** [Sesión 03b — 00:39:04]

![ReporteInsumosNoPlanificadosSalidaBodega](../../capturas/sesion_03b/enfoque_a/sesion_03b_ui_s03b_b020_01_ReporteInsumosNoPlanificadosSalidaBodega_003904.jpg)

### ReporteNoPlaneado

**Tipo:** reporte  
**Descripción:** Reporte o vista que muestra las salidas de productos no planificadas (adicionales) por servicio, usado para analizar desviaciones de consumo.  
**Mencionado en:** [Sesión 03a — 00:59:07]

![ReporteNoPlaneado](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b029_01_ReporteNoPlaneado_005907.jpg)

### ReporteOperacional

**Tipo:** reporte  
**Descripción:** Reporte del sitio que muestra venta, costos, merma y resumen operacional general  
**Mencionado en:** [Sesión 04 — 03:38:38]

![ReporteOperacional](../../capturas/sesion_04/enfoque_a/sesion_04_ui_s04_b108_02_ReporteOperacional_033838.jpg)

### ReportePlanificadaProducidaVendida

**Tipo:** reporte  
**Descripción:** Muestra las líneas de raciones planificadas, producidas y vendidas con información del proveedor/cliente para análisis de brechas.  
**Mencionado en:** [Sesión 04 — 03:43:32]

![ReportePlanificadaProducidaVendida](../../capturas/sesion_04/enfoque_a/sesion_04_ui_s04_b111_01_ReportePlanificadaProducidaVendida_034332.jpg)

### SalidaDeProduccion

**Tipo:** formulario  
**Descripción:** Formulario para registrar la salida de productos de un servicio, usado en estares médicos para los movimientos de productos.  
**Mencionado en:** [Sesión 03a — 02:25:48]

![SalidaDeProduccion](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b065_01_SalidaDeProduccion_022548.jpg)

### SalidasDeProduccion

**Tipo:** formulario  
**Descripción:** Pantalla donde se registran las salidas de bodega a producción por régimen y servicio, con cantidades planificadas y reales, organizable por sector  
**Mencionado en:** [Sesión 03a — 01:38:32]

![SalidasDeProduccion](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b048_01_SalidasDeProduccion_013832.jpg)

### SalidasProduccionPorSector

**Tipo:** formulario  
**Descripción:** Vista de salidas de producción organizada por sector (sopa, ensalada, etc.) con cantidades planificadas y campo para ingresar cantidad real entregada  
**Mencionado en:** [Sesión 03a — 01:40:30]

![SalidasProduccionPorSector](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b049_01_SalidasProduccionPorSector_014030.jpg)

### VentaContado

**Tipo:** formulario  
**Descripción:** Módulo para registrar ventas al contado sin facturación a cliente, originalmente diseñado para ingresar monto total de venta diaria.  
**Mencionado en:** [Sesión 03a — 02:19:49]

![VentaContado](../../capturas/sesion_03a/enfoque_a/sesion_03a_ui_s03a_b062_01_VentaContado_021949.jpg)


---

## 4. Flujos Identificados

## Flujos de uso del módulo "Producción"

### Flujo 1: Ingreso de cantidades producidas y cierre diario

1. El sitio accede a la planificación del día y registra las Q producidas para el servicio principal configurado (ej. almuerzo).
2. El sistema valida que las Q producidas estén ingresadas antes de permitir el cierre diario.
3. Si el sitio no las ingresa y transcurren 72 horas, el campo se bloquea automáticamente y queda en cero.
4. El administrador central puede desbloquear el campo desde el mantenedor si es necesario corregir retroactivamente.
5. Con las Q producidas registradas, el sitio puede continuar con las actividades obligatorias del cierre (salida de producción, control de raciones, mermas, raciones no vendidas).
6. Las actividades bloqueantes impiden el cierre; las no bloqueantes generan advertencia y permiten continuar.

### Flujo 2: Registro de salidas de producción

1. El sitio selecciona régimen y servicio para la salida.
2. El sistema carga por defecto los productos y cantidades provenientes de la planificación.
3. El usuario puede visualizar la salida en modo resumido (para bodega) o detallado (para producción/chef, agrupado por sector).
4. Si existen consumos no planificados (adicionales), el sitio los registra como salida separada, vinculada obligatoriamente a un servicio y régimen.

### Flujo 3: Impresión de requisiciones

1. El sitio verifica que las Q producidas estén ingresadas (requisito en versión 2.27).
2. El sistema habilita la impresión de la requisición en formato resumido (bodega) o detallado (producción).

### Flujo 4: Consulta del reporte de aportes

1. El administrador selecciona un día del mes; el sistema calcula acumulados dinámicamente hasta ese día.
2. El reporte compara raciones y costos planificados versus realizados. Los datos provienen del registro del sitio y no son editables desde la vista administrador.

---
> **Nota:** No hay información suficiente para reconstruir el flujo específico de mermas y raciones no vendidas.

---

## 5. Elementos de Base de Datos

> ⚠️ Sin información disponible en las sesiones analizadas.

---

## 6. Dudas Abiertas

> Estas dudas deben ser resueltas antes de que el documento sea considerado validado.

### 🔴 Prioridad Alta

| # | Duda | Tipo | Fuente |
|---|------|------|--------|
| 1 | Se debe definir qué actividades diarias son bloqueantes y cuáles son solo alertas en la nueva versión del sistema, dado que el mecanismo actual genera trabajo operativo significativo sin lograr el objetivo de obligar el ingreso de datos. | pendiente_definicion | [Sesión 03a — 00:06:12] |
| 2 | Se debe confirmar si la validación de servicios principales y actividades diarias se mantendrá en el nuevo sistema tal como está o si se rediseñará el mecanismo de obligatoriedad. | pendiente_definicion | [Sesión 03a — 00:08:09] |
| 3 | La versión 2.27 está aprobada y piloteada pero no desplegada para todos los sitios; se debe definir si se incorpora en la migración o se replantea la lógica. | pendiente_definicion | [Sesión 03a — 00:08:09] |
| 4 | ¿Se debe agregar un segundo campo Q en el reporte de aportes para registrar las raciones realmente cocinadas, diferenciando del Q planificado? | pendiente_definicion | [Sesión 03a — 00:38:27] |
| 5 | ¿Cómo se debería implementar el registro de adicionales en el sistema para evitar el uso de papeles y garantizar su trazabilidad? | pendiente_definicion | [Sesión 03a — 00:57:07] |
| 6 | ¿El personal en servicio puede acceder al sistema para registrar adicionales en tiempo real, o es inviable operacionalmente? | pendiente_definicion | [Sesión 03a — 00:57:07] |
| 7 | ¿Cómo evitar que los usuarios registren adicionales en un servicio incorrecto sin restringir excesivamente el sistema (ej: minería puede servir choclo en desayuno)? | pendiente_definicion | [Sesión 03a — 00:59:07] |
| 8 | ¿El cambio de configuración de separador decimal/miles de punto a coma generará problemas en datos existentes o en integraciones con otros módulos? | requiere_validacion | [Sesión 03a — 01:40:30] |
| 9 | Queda pendiente definir el mecanismo exacto de integración del control de raciones con el sistema SPRS para evitar la digitación manual. | pendiente_definicion | [Sesión 03a — 02:13:36] |
| 10 | Queda pendiente definir si los contratistas vendrán del sistema externo donde están cargados (distinto a SPRS) o si se integrarán directamente desde SPRS al control de raciones. | pendiente_definicion | [Sesión 03a — 02:15:36] |
| 11 | Pendiente definir cómo se implementará la integración entre SGP y el sistema de venta contratista para alimentar automáticamente los vales quemados. | pendiente_definicion | [Sesión 03a — 02:17:36] |
| 12 | No está claro si el ingreso final de raciones a facturar se realiza dentro del SPRS o dentro del SGP. | sin_respuesta | [Sesión 03a — 02:19:49] |
| 13 | Pendiente definir cómo manejar la venta contado en relación al módulo de control de raciones y el SPRS. | pendiente_definicion | [Sesión 03a — 02:21:51] |
| 14 | Pendiente definir cómo vincular la salida de bodega (con detalle de productos) al ingreso de venta contado para los casos de clínicas/estares médicos. | pendiente_definicion | [Sesión 03a — 02:23:48] |
| 15 | Pendiente definir cómo se vincula la salida de bodega al ingreso de venta para los casos particulares de estares médicos/clínicas, y cómo se maneja en el SPRS. | pendiente_definicion | [Sesión 03a — 02:25:48] |
| 16 | No está claro si el sistema SPRS registra el detalle de venta de estares médicos o solo un total, lo que afecta la integración. | sin_respuesta | [Sesión 03a — 02:25:48] |
| 17 | Pendiente decidir si se mantiene la separación biológico/no biológico/pan en la merma de desconcha del nuevo sistema, dado que el sistema anterior FLMS la tenía. | pendiente_definicion | [Sesión 03a — 02:32:07] |
| 18 | Se mencionan 'ventas especiales' como algo faltante en la revisión; queda pendiente confirmar si este submódulo fue cubierto completamente. | pendiente_definicion | [Sesión 03a — 02:36:52] |
| 19 | Queda pendiente definir si se mantendrá la impresión en papel de los registros diarios o se reemplazará por un flujo digital con aprobación. | pendiente_definicion | [Sesión 03a — 02:38:49] |
| 20 | Se propone un flujo de cierre diario con resumen de grandes cifras (merma, diferencia producido vs vendido) pero no está definido cómo implementarlo. | pendiente_definicion | [Sesión 03a — 02:38:49] |
| 21 | Queda pendiente definir si el cierre diario se mantendrá como paso manual del sitio o se modificará su flujo. | pendiente_definicion | [Sesión 03a — 02:45:23] |
| 22 | El reporte 'Insumos no planificados' tiene un defecto de cálculo: muestra el total del producto (planificado + no planificado) como no planificado en lugar de mostrar solo el exceso. Se debe definir si se corregirá este comportamiento. | pendiente_definicion | [Sesión 03b — 00:39:04] |
| 23 | Se debe confirmar cuáles son exactamente los costos que el sistema mostrará: solo el planificado (SGP administrador) y el local (costo del sitio), o también otros estados. | pendiente_definicion | [Sesión 04 — 03:32:36] |
| 24 | Solo el 40% de los sitios está registrando mermas de producción y desconches; no está definido qué acción se tomará con los sitios que no lo hacen | pendiente_definicion | [Sesión 04 — 03:38:38] |
| 25 | Se debe definir cómo administrar los vales en el sistema: la integración con el sistema de proveedores (SPRS u otro) para registrar la venta efectiva aún no está resuelta. | pendiente_definicion | [Sesión 04 — 03:43:32] |

### 🟡 Prioridad Media

| # | Duda | Tipo | Fuente |
|---|------|------|--------|
| 1 | ¿El reporte de aportes debe mantenerse igual en el nuevo sistema o requiere mejoras de usabilidad para que los sitios lo usen más activamente? | pendiente_definicion | [Sesión 03a — 00:36:09] |
| 2 | ¿La estructura fija de servicio debe eliminarse del sistema o simplemente deshabilitarse? | pendiente_definicion | [Sesión 03a — 00:59:07] |
| 3 | ¿El nuevo esquema de separación por tipo de producto (congelados vs otros) en la requisición reemplazaría o complementaría el orden lógico por sector en salidas de producción? | pendiente_definicion | [Sesión 03a — 01:38:32] |
| 4 | No queda claro por qué el total de cliente en el control de raciones no suma correctamente en el ejemplo mostrado. | sin_respuesta | [Sesión 03a — 02:13:36] |
| 5 | No queda definido si el campo identificador del cliente en la integración SPRS será el RUT (root/mandante) u otro identificador. | pendiente_definicion | [Sesión 03a — 02:15:36] |
| 6 | Se requiere definir si es deseable implementar trazabilidad que muestre el valor original de raciones ingresado y la corrección posterior negociada con el cliente. | pendiente_definicion | [Sesión 03a — 02:19:49] |
| 7 | Pendiente definir el mecanismo concreto para manejar precio por ración en el control de raciones (precio parametrizado vs. ingreso manual). | pendiente_definicion | [Sesión 03a — 02:23:48] |
| 8 | Se debate si es válido o factible crear productos individuales en el sistema para registrar cada tipo de merma de producción de forma detallada. | pendiente_definicion | [Sesión 03a — 02:27:59] |
| 9 | No está claro cómo se pesan físicamente las mermas de producción en los sitios (¿bolsa conjunta o separada por tipo?). | sin_respuesta | [Sesión 03a — 02:27:59] |
| 10 | Pendiente definir con los sitios si es factible y adecuado agrupar las mermas de producción por familias de productos. | pendiente_definicion | [Sesión 03a — 02:30:04] |
| 11 | No está definida la granularidad final de los grupos de merma de producción (¿proteínas, verduras, preparación fallida?). | pendiente_definicion | [Sesión 03a — 02:30:04] |
| 12 | No está definido si el campo 'no biológico' será obligatorio o no en el registro de merma de desconcha. | pendiente_definicion | [Sesión 03a — 02:32:07] |
| 13 | Pendiente definir dónde se configurará si los campos de desconcha (biológico, no biológico, pan) van agrupados o separados por sitio. | pendiente_definicion | [Sesión 03a — 02:34:04] |
| 14 | En sitios pequeños con desayuno y almuerzo, el desconche puede realizarse al final del día acumulando bolsas de distintos servicios, lo que dificulta el registro por servicio. | pendiente_definicion | [Sesión 03a — 02:34:55] |
| 15 | El término utilizado para el dato real modificado por el chef fue mencionado indistintamente como 'real' y 'curial'; se requiere confirmar cuál es el término oficial del sistema. | inconsistencia | [Sesión 03a — 02:36:52] |
| 16 | Se menciona revisar si se quiere ver el detalle de los eventos especiales en el estado de resultados o mantenerlo como total; queda pendiente. | pendiente_definicion | [Sesión 03a — 02:42:58] |
| 17 | Queda pendiente definir si las salidas de eventos especiales se separarán de las salidas de producción regular en los reportes de costo. | pendiente_definicion | [Sesión 03a — 02:45:23] |
| 18 | Se consulta qué gestión se realiza con los resultados del reporte de insumos no planificados: si se hacen ajustes de inventario o es solo control. | sin_respuesta | [Sesión 03b — 00:39:04] |
| 19 | Existe desorden en la gestión de clientes del sistema; se debe definir si se usará RUT ficticio 'vales' o si se mantendrá la estructura actual de clientes. | pendiente_definicion | [Sesión 04 — 03:43:32] |

### ⚪ Prioridad Baja

| # | Duda | Tipo | Fuente |
|---|------|------|--------|
| 1 | ¿La regla de aproximación ±5 es fija en el sistema o es configurable por parámetro? | requiere_validacion | [Sesión 03a — 01:40:30] |
| 2 | No queda claro si los informes del módulo 13 serán revisados en la misma sesión o en una posterior. | sin_respuesta | [Sesión 03a — 02:45:23] |


---

## 7. Observaciones y Comentarios

> Esta sección es para uso colaborativo. Los responsables técnicos y de negocio pueden agregar comentarios, correcciones o contexto adicional aquí.

_Sin observaciones registradas aún._

---

## 8. Tabla de Fuentes

> Todos los bloques de transcripción que aportaron información a este módulo, en orden cronológico.

| Sesión | Timestamp | Resumen del bloque |
|--------|-----------|-------------------|
| Sesión 03a | 00:04:14 | ¿Porque como vamos a hacer la migración de Ashur para que no estuviera bloqueado y por algún motivo si tuviera un proble... |
| Sesión 03a | 00:06:12 | Si les quedó en blanco por cualquier motivo se le olvidó poner las Q, no pueden volver a ponerlas. Ya siempre van a qued... |
| Sesión 03a | 00:08:09 | No, pregunta, Cecilia, disculpa, esta validación anterior de los servicios principales y esta la vamos a mantener en el ... |
| Sesión 03a | 00:36:09 | centralizadamente. Lo único sería que si ellos cambiaran mucho las minutas, acá las preparaciones habría que ver si esos... |
| Sesión 03a | 00:38:27 | En este caso, acá el planificado y realizado, aquí es donde podríamos tener el otro Q, si es que cambia. Sí, planifiqué ... |
| Sesión 03a | 00:57:07 | entregar 30 kg de pollo, independiente que vaya para la sopa o Claro, por eso lo que entiendo, ay, disculpa, Sí, bueno. ... |
| Sesión 03a | 00:59:07 | ir a poner la El choclo la decir 10 kg de choclo y listo y me va a parecer como que Te amo. está planificado pero igual ... |
| Sesión 03a | 01:38:32 | Voy a las salidas de producción. Acá lo que hace el sistema les voy a mostrar el mismo primero. Yo cargo el régimen. Yo ... |
| Sesión 03a | 01:40:30 | Lo puedo hacer por sector, ya acá si te fijas acá tengo la sopa, está planificada la crema espárrago, por lo tanto en la... |
| Sesión 03a | 02:13:36 | Esos son con el ingreso de el stock, los movimientos que tenga stock durante el día. e y la venta eso sería el control d... |
| Sesión 03a | 02:15:36 | La suma del total cliente debería darme la suma de todo esto que tengo acá abajo. O sea, sí, o sea, debería cuadrarme co... |
| Sesión 03a | 02:17:36 | Debería aparecer el mandante y luego abajo deberían aparecer los contratistas asociados a ese mandante, que es lo que es... |
| Sesión 03a | 02:19:49 | ¿Pregunto, es deseable que yo, por ejemplo, si conté bandeja y conté 100 y después el cliente me dijo que eran 99, tener... |
| Sesión 03a | 02:21:51 | Ya porque no lo facturo, sí, pero ¿qué pasa con esos eso que ingreso como venta contado? Pierdo la visibilidad de las ra... |
| Sesión 03a | 02:23:48 | Entonces, a lo mejor, porque aquí tiene que ir un Q, lo que debería hacer es tal vez en uno de los módulos que deshabili... |
| Sesión 03a | 02:25:48 | Ya y después para hacer los movimientos sí puedo sacar los productos a través de la salida de producción de ese servicio... |
| Sesión 03a | 02:27:59 | Me puede pedir producto un montón de ventas. Sí, Ya entendí. por eso ahora está ingresado como un total. Ya, okay. OKOK.... |
| Sesión 03a | 02:30:04 | funciona? Puede ser, sí, puede ser así, pero a mí me consta que en la en las presentaciones y todos los jefes y todo bie... |
| Sesión 03a | 02:32:07 | Eso, pero importante es lo que puso la grises, Yeah. porque uno es el de el la misma producción natural que tengo que de... |
| Sesión 03a | 02:34:04 | ya y la Pau ahí se adelantó pero el pan estaba separado el pan en el en el ideal de que el sitio lo pueda separar y pued... |
| Sesión 03a | 02:34:55 | Yeah. Yeah, okay. Completo. Okay, yeah. Pero sí es por servicio, yo tengo que ingresarlo por servicio. Debería ingresarl... |
| Sesión 03a | 02:36:52 | Aquí jalea de jalea me quedaron 3 raciones sin vender. Me dice cuántos kilos son y cuántas es la misma total. Y esto es ... |
| Sesión 03a | 02:38:49 | Necesitamos un Necesitamos un impreso para este tipo de formulario, pregunta. Para que las metas. Para la mayoría de los... |
| Sesión 03a | 02:42:58 | ¿Por qué? ¿Por qué era un dolor para la planificación? Porque si yo tenía un evento especial, como no tenía dónde ingres... |
| Sesión 03a | 02:45:23 | Y las salidas se suman a las salidas de producción. No está detallada como eventos especiales. A lo mejor eso es algo qu... |
| Sesión 03b | 00:39:04 | Aquí me muestra por servicio, yo puedo filtrar los servicios del régimen y me dice. Por ejemplo, para el desayuno hice u... |
| Sesión 04 | 03:32:36 | A ver, este sí lo voy a abrir porque tiene varios. Para información. Ahí está puesto el pipe, voy a tratar de guardarme ... |
| Sesión 04 | 03:38:38 | No sé si se acuerdan cuando la Ceci mostró donde se ingresaban las mermas de línea que habían unos campos abajo para ing... |
| Sesión 04 | 03:43:32 | Mira, esto es lo que trae planificada producida. Hmm. T. Planificada, producida y vendida. Esas son los las líneas que t... |

---

## 9. Historial de Cambios

| Fecha | Autor | Tipo | Descripción |
|-------|-------|------|-------------|
| 2026-02-27 | Claude Code (automático) | Creación | Documento generado automáticamente desde análisis de transcripciones |
