# Precio Venta Cliente

**Formulario VB6:** `M_PVtaCl.frm`
**Tabla principal:** `b_preciovta`
**Módulo de consulta:** `RutinaLectura.PrecioVta` (query builder en `RutinaLectura.cls`)

---

## Contexto

Este formulario permite registrar y mantener el **precio de venta individual por cliente** para una combinación de Contrato + Régimen + Servicio + Fecha de vigencia. Es el punto de partida para el control de raciones: el precio aquí registrado queda asociado a las raciones planificadas en `b_minutaraciones` a partir de esa fecha.

---

## Parámetros de Entrada (Encabezado)

| Campo | Descripción | Obligatorio |
|---|---|---|
| Contrato (CeCo) | Identificador del casino | Sí |
| Régimen | Código del régimen alimenticio | Sí |
| Servicio | Código del servicio (desayuno, almuerzo, etc.) | Sí |
| Inicio de Validez | Fecha desde la cual rige el precio (formato dd/mm/yyyy) | Sí |

> Los cuatro campos deben estar completos antes de poder Agregar, Modificar o Eliminar. El sistema rechaza la operación con mensaje de advertencia si alguno falta.

---

## Estructura de la Grilla

| Col | Nombre | Origen | Editable |
|---|---|---|---|
| 1 | RUT Cliente | `b_clientes.cli_codigo` (formateado con dígito verificador) | Sí (en modo A) |
| 2 | Nombre Cliente | `b_clientes.cli_nombre` | No (cargado automáticamente) |
| 3 | **Precio de Venta** | `b_preciovta.prv_preven` | **Sí** |
| 4 | RUT interno | Copia del RUT sin formatear (uso interno) | No |

---

## Operaciones Disponibles

| Botón | Acción |
|---|---|
| **Agregar** | Habilita ingreso de una nueva fila. Requiere encabezado completo. |
| **Modificar** | Habilita edición del precio en la fila seleccionada. Requiere encabezado completo. |
| **Eliminar** | Borra el registro del cliente seleccionado. Verifica si tiene raciones asociadas antes de confirmar. |
| **Grabar** | Persiste el registro (INSERT o UPDATE) en `b_preciovta`. |
| **Cancelar** | Descarta el cambio pendiente y restaura el valor anterior desde la base de datos. |
| **Refrescar** | Recarga la grilla con los datos actuales de la base de datos. |
| **Imprimir** | Genera el informe de precios de venta para el encabezado seleccionado. |
| **Cerrar** | Cierra el formulario. |

---

## Validaciones

### 1. Encabezado completo
Antes de Agregar, Modificar o Eliminar, el sistema verifica que los cuatro campos del encabezado (Contrato, Régimen, Servicio y Fecha) estén ingresados. Si alguno falta, se muestra el mensaje _"Falta información en encabezado..."_ y la operación no continúa.

---

### 2. Fecha válida
Al cambiar la fecha de inicio de validez, el sistema verifica que el valor ingresado sea una fecha válida. Si no lo es, la grilla se limpia y no se realiza ninguna consulta.

---

### 3. Cliente debe existir y ser de tipo externo
Al ingresar o salir del campo RUT en la grilla, el sistema consulta `b_clientes` y valida:
- El RUT ingresado debe existir en la tabla de clientes.
- El cliente debe ser de **tipo 1** (cliente externo; excluye personal interno u otras categorías).
- El cliente debe estar **activo** (`cli_activo = '1'`).

Si no se cumple alguna de estas condiciones, el campo RUT queda en blanco y el cursor vuelve a esa celda.

---

### 4. RUT con dígito verificador
Al abandonar el campo RUT, el sistema aplica la función de validación de RUT chileno. Si el dígito verificador es incorrecto, el valor no se acepta.

---

### 5. Cliente duplicado en la grilla
Al ingresar un RUT, el sistema recorre todas las filas existentes en la grilla y verifica que ese RUT no esté ya ingresado. Si se detecta duplicado, se muestra el mensaje _"Cliente existe"_ y se limpia la celda.

---

### 6. Cliente duplicado en la base de datos (solo al Agregar)
Antes de insertar un nuevo registro, el sistema consulta `b_preciovta` con los mismos parámetros (Contrato + Régimen + Servicio + Fecha + RUT). Si ya existe un precio registrado para ese cliente en ese período, se muestra el mensaje _"Cliente existe"_ y se cancela la inserción.

---

### 7. RUT requerido para grabar
Al intentar grabar, el sistema verifica que el RUT del cliente no esté vacío. Si falta, se muestra el mensaje _"Falta información..."_ y el foco vuelve a la celda del RUT.

> **Nota:** La validación de precio mínimo (`precio ≥ 1`) está presente en el código pero comentada. El sistema acepta precio igual a cero.

---

### 8. Eliminación con raciones asociadas
Antes de eliminar un precio, el sistema verifica si ese cliente tiene raciones registradas en `b_minutaraciones` con fecha igual o posterior a la fecha de vigencia del precio. Existen dos escenarios:

| Situación | Mensaje mostrado |
|---|---|
| Tiene raciones asociadas | _"Elimina registro que está asociado control raciones..."_ (requiere confirmación explícita) |
| Sin raciones asociadas | _"Elimina registro..."_ (confirmación estándar) |

En ambos casos se requiere confirmación del usuario (Sí / No).

---

## Flujo de Datos

```
1. Usuario ingresa: Contrato + Régimen + Servicio + Fecha
        │
        ▼
2. Sistema carga la grilla con los precios registrados
   para esa combinación (b_preciovta ⟶ b_clientes)
        │
        ▼
3. Usuario selecciona operación:

   [AGREGAR]                 [MODIFICAR]            [ELIMINAR]
       │                         │                       │
   Ingresa RUT cliente       Edita precio           Verifica raciones
       │                         │                  asociadas en b_minutaraciones
   Sistema valida RUT            │                       │
   (existencia, tipo, activo)    │                  Confirma y borra
       │                         │                  de b_preciovta
   Ingresa precio                │
       │                         │
       ▼                         ▼
4. [GRABAR] → INSERT o UPDATE en b_preciovta
```

---

## Dónde se Almacena

| Dato | Tabla | Campo |
|---|---|---|
| Centro de costo | `b_preciovta` | `prv_cencos` |
| Régimen | `b_preciovta` | `prv_codreg` |
| Servicio | `b_preciovta` | `prv_codser` |
| Fecha de vigencia | `b_preciovta` | `prv_fecvig` |
| RUT cliente | `b_preciovta` | `prv_rutcli` |
| Precio de venta | `b_preciovta` | `prv_preven` |
| Indicador integración SPRS | `b_preciovta` | `prv_SPRS` |

**Clave primaria:** combinación de los cinco primeros campos (`prv_cencos` + `prv_codreg` + `prv_codser` + `prv_fecvig` + `prv_rutcli`). No puede existir más de un precio para el mismo cliente en el mismo período.

---

## Consulta de Lectura (`RutinaLectura.PrecioVta`)

El formulario no llama a un Stored Procedure dedicado. La lectura se realiza mediante una query construida dinámicamente:

```sql
SELECT a.cli_codigo,
       ISNULL(a.cli_nombre,'')   AS cli_nombre,
       ISNULL(b.prv_preven, 0)   AS prv_preven,
       ISNULL(b.prv_SPRS, '')    AS prv_SPRS
FROM   b_clientes   a,
       b_preciovta  b
WHERE  b.prv_rutcli = a.cli_codigo
AND    b.prv_cencos = :cencos
AND    b.prv_codreg = :codreg
AND    b.prv_codser = :codser
AND    b.prv_fecvig = :fecha          -- formato YYYYMMDD
AND   (b.prv_rutcli = :codcli OR :codcli = '')
AND    a.cli_tipo   = 1
```

> Cuando `codcli` está vacío, retorna todos los clientes del período. Cuando tiene valor, filtra por cliente específico (usado para validar duplicados al agregar).

---

## Relación con Otros Módulos

| Módulo | Relación |
|---|---|
| **Control de Raciones** (`b_minutaraciones`) | El precio vigente se aplica a las raciones planificadas a partir de `prv_fecvig`. Al eliminar un precio, el sistema alerta si hay raciones que ya lo están usando. |
| **Historial de precios** (`B_HistPm`) | Desde el botón de fecha (Image1 Index=3), se puede consultar el historial de precios anteriores para el mismo cliente/régimen/servicio. |
| **Limpieza de datos históricos** (`sgp_Del_Limpiadatos`) | SP de mantenimiento que borra registros de `b_preciovta` con fecha de vigencia anterior a un corte definido. No es invocado desde este formulario. |
| **Integración SPRS** | La columna `prv_SPRS` indica si el precio proviene del sistema SPRS de Sodexo. La validación que bloqueaba edición de registros SPRS está comentada en el código (inactiva). |

---

*Fuentes: `M_PVtaCl.frm`, `RutinaLectura.cls`, tabla `b_preciovta` en `SGP_Local.sql`*
