# Prompt: Relevamiento del Módulo de Producción – Sistema SGP

## Contexto

Estás analizando el sistema actual **SGP (Sistema de Gestión...)** con el objetivo de generar documentación técnica y funcional completa del **módulo de Producción**. Esta documentación será entregada al proveedor que desarrollará la nueva versión del sistema, por lo que debe ser exhaustiva, precisa y autocontenida.

El sistema está desarrollado en **Visual Basic 6** con base de datos **SQL Server**. Tiene una arquitectura de múltiples capas que debés analizar **todas sin excepción**:

- **Frontend:** formularios y lógica de presentación en Visual Basic 6 (`.frm`, `.bas`, `.cls`)
- **Backend / lógica de negocio:** módulos y clases VB6
- **Base de datos:** Stored Procedures, Triggers, Funciones (UDF) y Vistas (Views)

Tienes acceso a los siguientes recursos en la carpeta del proyecto:

- `codigo_fuente/` – Código fuente VB6 de la aplicación actual
- `manual_usuario/` – Manual de usuario del módulo de Producción
- `Documentos/` – Transcripciones de reuniones con usuarios del sistema
- `base_de_datos/` – Scripts SQL (tablas, stored procedures, triggers, funciones y vistas)

### Formularios VB6 identificados del módulo de Producción

Analizá **todos** los archivos del proyecto, pero prestá especial atención a los siguientes formularios que corresponden al módulo de Producción:

| Archivo | Funcionalidad |
|---|---|
| `M_Plami1.frm` | Búsqueda Planificación Real |
| `M_MinRea.frm` | Detalle Planificación Real |
| `I_SalBod.frm` | Requisición |
| `M_MerPre.frm` | Raciones no Vendidas |
| `M_ConRac.frm` | Control de Raciones |
| `M_VtaCon.frm` | Venta Servicio Contado |
| `M_RCDiar.frm` | Lista Precio Cafetería |
| `M_RCDiar.frm` | Precio Venta Cliente |
| `M_RCDiar.frm` | Cierre diario |

> ⚠️ Esta lista puede estar incompleta. Si durante el análisis encontrás formularios adicionales relacionados con Producción que no figuran aquí, incluirlos en la documentación e indicá que no estaban en la lista original.

---

## Tarea

Analiza **todos los archivos disponibles** de forma cruzada y genera un documento Markdown con la documentación completa del módulo de Producción. El documento debe cubrir los siguientes apartados:

---

## Estructura del documento a generar

### 1. Introducción al módulo
- Propósito y alcance del módulo de Producción dentro del SGP
- Actores/usuarios que interactúan con el módulo y sus roles
- Relación con otros módulos del sistema (si aplica)

### 2. Entidades y modelo de datos
- Tablas principales involucradas en el módulo (extraídas del script de BD)
- Descripción de cada campo relevante (tipo, restricciones, valores posibles)
- Relaciones entre tablas (claves foráneas, cardinalidades)
- Incluir un **diagrama entidad-relación en Mermaid**

### 3. Funcionalidades por submódulo
Para cada funcionalidad identificada (planificación, requisición, mermas, etc.) documentar:

- **Descripción** de la funcionalidad
- **Pantallas / formularios** involucrados (campos, controles, filtros)
- **Flujo de usuario** paso a paso
- **Diagrama de flujo en Mermaid** para los procesos más relevantes
- **Validaciones** de frontend (VB6) y backend/BD
- **Reglas de negocio** aplicadas
- **Resultados / efectos** en la base de datos u otros módulos

### 4. Reglas de negocio consolidadas
Listar todas las reglas de negocio identificadas en el módulo, numeradas y con su fuente de referencia (código / manual / reunión). Ejemplo:

> **RN-PROD-001** – 2.	Solo se podrán sumar una cantidad determinada de preparaciones por día. Esta cantidad es configurable a nivel del sitio.  
> *Fuente: reuniones/reunion_02.txt, codigo_fuente/producción/frmStock.frm*

### 5. Validaciones del sistema
Listar todas las validaciones identificadas, separando:
- Validaciones de formato/obligatoriedad de campos (VB6)
- Validaciones de integridad de datos (constraints, triggers)
- Validaciones de negocio (condiciones complejas en SP o código VB6)

### 6. Objetos de base de datos – Análisis detallado

Esta sección es crítica. Los stored procedures, triggers, funciones y vistas suelen contener reglas de negocio que **no están documentadas en ningún otro lugar**. Analizá cada objeto con el siguiente criterio:

#### 6.1 Stored Procedures
Para cada SP identificado documentar:
- **Nombre y propósito**
- **Parámetros de entrada y salida** (nombre, tipo, si es opcional)
- **Lógica principal:** qué tablas lee/modifica y en qué orden
- **Reglas de negocio implícitas:** condiciones `IF`, `CASE`, umbrales, cálculos
- **Manejo de errores y transacciones:** uso de `BEGIN TRAN`, `ROLLBACK`, `RAISERROR`
- **Desde dónde se llama** (formulario VB6 o desde otro SP)
- **Efecto en otras tablas o módulos**

Prestar especial atención a SP relacionados con:
- Planificación real.
- Generación de requisición.
- Mermas de Producción, Desconche y Pan.
- Raciones no vendidas.

#### 6.2 Triggers
Para cada trigger documentar:
- **Tabla sobre la que actúa** y evento (`INSERT`, `UPDATE`, `DELETE`)
- **Momento de ejecución** (`AFTER` / `INSTEAD OF`)
- **Lógica que implementa:** qué hace exactamente y por qué
- **Tablas adicionales que afecta**
- **Reglas de negocio que implementa** (muchas veces son validaciones o propagaciones automáticas que el usuario ni sabe que existen)
- **Interacción con otros triggers** (orden de ejecución si hay múltiples sobre la misma tabla)

Prestar especial atención a triggers que:
- Actualicen datos de la planificación.
- Registro de mermas diferentes a las de inventario.

#### 6.3 Funciones (UDF)
Para cada función documentar:
- **Nombre, tipo** (escalar o tabla) **y propósito**
- **Parámetros de entrada**
- **Lógica de cálculo** (especialmente fórmulas de valorización, costo promedio, FIFO, etc.)
- **Desde dónde se utiliza** (SP, vistas, consultas VB6)

#### 6.4 Vistas (Views)
Para cada vista documentar:
- **Nombre y propósito**
- **Tablas que consolida**
- **Campos calculados o derivados**
- **Desde dónde se consume** (reportes, formularios VB6, otros SP)

### 7. Integraciones y dependencias
- Otros módulos que consumen o alimentan datos de Producción
- Procesos automáticos, jobs del SQL Agent o tareas programadas
- Describir el flujo de datos entre módulos con un **diagrama de secuencia Mermaid**

### 8. Trazabilidad y auditoría
- Tablas o mecanismos de log de operaciones
- Qué eventos quedan registrados (quién, cuándo, qué cambió)
- Cómo se consulta el historial desde la interfaz

### 9. Valorización y costos
- Cómo se calculan los costos
- Impacto de ajustes de raciones en el costeo del servicio
- SP o funciones involucradas en estos cálculos

### 10. Reportes y consultas
- Listado de reportes disponibles en el módulo
- Filtros disponibles, campos mostrados y lógica de cálculo si aplica
- Vistas o SP que los alimentan

### 11. Casos especiales y excepciones
- Comportamientos no estándar detectados en el código o mencionados en reuniones
- Workarounds conocidos por los usuarios
- Funcionalidades incompletas o con deuda técnica identificada

### 12. Preguntas abiertas
Listar todos los puntos que no pudieron documentarse con certeza y requieren validación con el usuario o el equipo técnico. Formato sugerido:

> ❓ **Pregunta-PROD-001** – El SP `sp_AjusteStock` tiene una condición para depósitos con código > 900 que no está explicada en ningún otro documento. ¿Corresponde a depósitos virtuales o transitorios?

### 13. Glosario
- Términos del dominio utilizados en el módulo con su definición

---

## Instrucciones de análisis

1. **Prioriza las reuniones más recientes** si hay varias transcripciones: la lógica de negocio puede haber evolucionado y versiones anteriores pueden estar desactualizadas. Verificá la fecha de cada archivo antes de comenzar.

2. **No omitas ninguna capa tecnológica.** El sistema tiene lógica distribuida entre VB6 y SQL Server. Es común en sistemas de esta arquitectura que reglas críticas estén únicamente en triggers o stored procedures, sin reflejo en el código VB6 ni en el manual.

3. **Cruza siempre las cuatro fuentes** (código VB6 + scripts SQL + manual + reuniones) para detectar discrepancias entre lo documentado, lo implementado y lo que el usuario realmente usa.

4. Si encontrás **contradicciones entre fuentes**, documentalas explícitamente con una nota:  
   > ⚠️ *Discrepancia detectada: el manual indica X, pero el trigger `trg_MovStock` implementa Y. Según reunión del DD/MM, los usuarios experimentan el comportamiento Y.*

5. Para cada **diagrama de flujo**, usá sintaxis Mermaid válida (`flowchart TD` o `sequenceDiagram` según corresponda).

6. Si una funcionalidad **no está suficientemente documentada** en ninguna fuente, marcala con:  
   > 🔍 *Requiere validación con el usuario – información insuficiente para documentar con certeza.*

7. Incluir **referencias de archivo** en cada sección para que el nuevo proveedor pueda rastrear el origen de cada decisión.

8. Para los objetos SQL, prestá especial atención al **manejo de transacciones**: en sistemas VB6 + SQL Server es frecuente que parte de la transacción se controle desde el código VB6 (con `BeginTrans`) y otra parte desde los SP, lo que puede generar comportamientos sutiles que deben quedar documentados.

---

## Formato de salida

- Documento **Markdown** con tabla de contenidos al inicio (TOC con links)
- Usar encabezados jerárquicos (`##`, `###`, `####`)
- Diagramas embebidos en bloques de código Mermaid (` ```mermaid `)
- Tablas Markdown para campos de formularios, parámetros de SP y columnas de BD
- Notas especiales con blockquotes (`>`) y los emojis ⚠️, 🔍 y ❓ según corresponda

### Imágenes de formularios y pantallas

Las imágenes se agregarán manualmente luego de generado el documento. La estructura de carpetas será:

```
doc_funcional/
├── produccion.md
└── imagenes/
    ├── frm_alta_articulo.png
    ├── frm_movimientos.png
    └── ...
```

En cada sección donde se documente un formulario o pantalla, incluí un placeholder con la siguiente sintaxis:

```markdown
![Alta de Artículo](imagenes/frm_alta_articulo.png)
> 📸 *Pendiente: insertar captura de pantalla del formulario frmAltaArticulo*
```

Criterios para los placeholders:
- Usá como nombre de archivo el nombre real del formulario VB6 (`.frm`) en minúsculas, reemplazando espacios por guiones bajos
- Incluí un placeholder por cada formulario principal del módulo
- Si un proceso tiene pantallas intermedias relevantes (confirmaciones, grillas de selección, mensajes de error importantes), agregá un placeholder adicional para cada una
- Para reportes, incluí también un placeholder con el nombre del reporte

---

## Sugerencias para este relevamiento

### Antes de ejecutar el prompt

- **Organizá los archivos con nombres descriptivos** antes de pasárselos a Claude Code (ej: `reunion_01_operadores_2024-11.txt`). Facilita las referencias cruzadas y la identificación de la reunión más reciente.
- Si el script de base de datos es un único archivo grande, verificá que incluya todos los objetos: tablas, SP, triggers, UDF y vistas. Si están separados por archivos, indicáselo explícitamente a Claude Code al inicio.
- Indicale a Claude Code la estructura de carpetas antes de empezar para que no omita archivos.

### Durante el análisis

- Ejecutá el análisis **submódulo por submódulo** si Producción es grande. Mejor varios documentos enfocados que uno superficial.
- Pedile a Claude Code que genere primero un **índice de funcionalidades y objetos SQL detectados** para que lo valides antes de que desarrolle toda la documentación. Así detectás omisiones temprano.
- Prestá especial atención a los **triggers**: en sistemas VB6 es habitual que los usuarios no sepan que existen y que contengan validaciones críticas que dieron origen a bugs históricos.

### Validación del documento generado

- Una vez generado, **compartí el documento con los usuarios clave** del módulo para una revisión de exactitud antes de entregarlo al proveedor.
- Revisá especialmente la sección de **Preguntas abiertas**: cada ítem sin respuesta es un riesgo potencial en el nuevo desarrollo.
- Solicitá al proveedor que firme una **conformidad de comprensión** del documento antes de iniciar el desarrollo, para reducir riesgos de malentendidos.