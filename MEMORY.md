# Memoria del Proyecto SGP-Producción

## Proyecto
Repositorio de documentación para el Módulo de Producción/Inventario de SGP LOCAL.
Sistema: Visual Basic 6 + SQL Server. Casinos Sodexo Chile.

## Entregables GENERADOS
- `doc_funcional/Documentacion_SGP_Produccion.md` — (2026-02-27) 2090 líneas, 79KB, 13 secciones, 57 RN, 14 preguntas abiertas, glosario 40 términos
- `doc_funcional/flujos_produccion.html` — ELIMINADO por solicitud del usuario
- `doc_funcional/produccion_sgp.html` — (2026-03-02) HTML interactivo completo, 1614 líneas. Ver detalles en sección HTML abajo.

## Estructura de carpetas
- `codigo_fuente/` — 553 archivos VB6 (.frm/.bas/.cls)
- `base_de_datos/SGP_Local.sql` — 39,776 líneas UTF-16 LE
  - Convertir: `iconv -f UTF-16 -t UTF-8 "...SGP_Local.sql" > /tmp/SGP_Local_utf8.sql`
- `manual_usuario/MANUAL DE PROCEDIMIENTOS SGP v.226.docx`
- `Documentos/` — 14 archivos .docx de sesiones

## Referencia técnica detallada
Ver `memory/detail_forms_sps.md` — 12 formularios VB6 analizados, SPs por módulo (planificación, raciones, mermas, requisición, salida bodega, cierre, FLMS), UDFs, vistas y tablas principales.

## REGLAS DE NEGOCIO (17 confirmadas)
1. PC 'SvrAppCont' único autorizado para cierre diario
2. Bloqueo automático planificación 72h antes del cierre (min_indblo=11)
3. Merma × Preparación ≤ raciones planificadas (mid_numrac)
4. Cambio receta con mermas existentes rechazado
5. Máx 5 recetas adicionales/día/servicio ('addreceta'), color amarillo
6. Recetas ≥10000 = centralizadas AMD, solo lectura en regímenes 5-etapas
7. 3 tipos receta: Patrón(tiprec=0,cencos='0'), Local(tiprec=-1), xRégimen(tiprec>0)
8. PMP = ((PMP_ant×CantBod)+(Precio×CantIng))/(CantBod+CantIng)
9. Fila PRODUCIDAS requiere password (a_param 'parcomdia')
10. Al marcar facturado=1: DELETE raciones clientes, preserva PRODUCIDAS/PERSONAL/MERMAS
11. Requisición: solo días abiertos, solo minuta real (mid_tipmin='2')
12. Requisición cantidad = (raciones × gramaje) / pro_facing
13. Cierre requiere 14+ validaciones previas (CierrePeriodo)
14. Días feriados (b_Fecha_Inhabiles) → recálculo PMP especial
15. Costo receta se congela al grabar (mid_cosrec, mid_cosdes)
16. Producción NO tiene mantenedores propios (depende de Contrato/Régimen/Servicio externos)
17. Modificar receta = DELETE completo + INSERT (no edición incremental)

## HALLAZGOS DE REUNIONES
- Sesión 1 (17 Dic 2025): Flujo chef-bodeguero en papel. Adicionales: papel, luego se digita todo junto desfasado
- Sesión 2 (23 Dic 2025): GAP: merma producción NO registrada. GAP: faltan raciones PRODUCIDAS (solo hay plan+vendidas). Problema conectividad
- Sesión 3 (24 Dic 2025): Producción NO tiene mantenedores propios
- SGP Upgrade 1 (29 Ene 2026): Flujo inventario completo: proveedor→guías→traspasos→stock→mermas→salidas→dev
- SGP Upgrade 2 (9 Feb 2026): Formato requisición detallado preferido. Chef/bodeguero negocian redondeamiento. Límite 1 mes
- Sesión 11 Feb 2026: Proceso sin digital mientras cocina. Contratos que obligan mismas alternativas vs flexibles
- Sesión 12 Feb 2026: Servicios especiales precio por comensal O por total. Salidas bloquean stock hasta cierre

## PATRONES ARQUITECTURA CONFIRMADOS
- Transacciones: BeginTrans/CommitTrans desde VB6 Y BEGIN TRAN en SPs (ambos niveles)
- XML: sgp_Ins_XmlMinutaReal, sgp_Upd_XmlMermaPreparacion, sgp_DelIns_formatorequesicionestdetallado
- Correlativos: b_parametros par_tipdoc (SP/DP/ME/SE/DE)
- CierrePeriodo(): función VB6 con 14+ índices de validación, usada en todos los forms
- Fecha cifrada: a_param 'ciediario' desencriptado con sgp_p_desencripta()
- NO hay triggers en el archivo SQL (confirmado)
- vg_tipbase: '1'=Access (legacy), '2'=SQL Server

## PREGUNTAS ABIERTAS
- ¿Timeout o consistencia si se interrumpe el cierre diario?
- ¿Modo offline para conectividad deficiente?
- ¿Por qué servicios 11056 y 11057 excluidos de validación raciones producidas?
- ¿Qué pasa con adicionales que no están en maestro de productos?
- ¿'parcomdia' tiene renovación periódica de password?
- ¿Integración FLMS es tiempo real o batch?
- ¿M_SalidaServicioEspeciales tiene mismo flujo de cierre que salidas normales?

## HTML produccion_sgp.html — Estructura y patrones (2026-03-02)
Formato idéntico a `inventario_sgp.html` (referencia de estilo del proyecto).
- CSS: variables --primary #1e3a5f, --accent #e85d04, sidebar 270px fijo, mermaid@10 CDN
- Clases clave: `.card`, `.rule-item`, `.sp-card/.sp-header/.sp-body` (collapsibles), `.db-tabs/.db-section`, `.question`, `.glos-item`, `.process-card`, `#flow-panel`
- JS: `showSection(id)`, `selectProcess(id)` (12 submódulos con Mermaid interactivo), `showDbTab(tab)`
- **Zoom interactivo**: clase `zoomable-mermaid` + `style="cursor:zoom-in;"` en el div `.mermaid`. Modal `#mermaid-zoom-modal` con SVG clonado escalado. Click fuera cierra. Escala: 2.8× flowcharts, 2.2× sequenceDiagrams.
  - Diagramas con zoom: Arquitectura, ER, Integraciones General, FLMS→SGP, SGP→SAP, Flujo de Costos
- **ER Diagram**: además tiene `style="zoom:1.4"` + wrapper `overflow-x:auto` para +40% tamaño base
- **Preguntas abiertas**: formato `❓ Pregunta-PROD-001` a `❓ Pregunta-PROD-014` (no Q-XXX)
- Secciones (13): sec-intro, sec-datos, sec-flujos, sec-reglas, sec-validaciones, sec-bd, sec-integraciones, sec-trazabilidad, sec-valorizacion, sec-reportes, sec-casos, sec-preguntas, sec-glosario
- Estadísticas: 9 diagramas Mermaid, 38 rule-items, 15 sp-cards, 14 preguntas, 38 glos-items

## Estrategia de generación HTML (para futuros HTMLs grandes)
Usar 4 agentes secuenciales/paralelos para evitar límite de tokens:
1. **Agente 1** (foreground): Crea esqueleto con placeholders `<!-- PLACEHOLDER_XXX -->` en cada sección
2. **Agente 2** (background): Rellena sec-intro, sec-datos, sec-flujos
3. **Agente 3** (background): Rellena sec-reglas, sec-validaciones, sec-bd
4. **Agente 4** (background): Rellena sec-integraciones…sec-glosario
Agentes 2/3/4 corren en paralelo después de que Agente 1 termina. Usar Edit (no Write) para reemplazar placeholders.

## Preferencias del usuario (Joaquin)
- Responder siempre en español
- Generar documentación exhaustiva cruzando las 4 fuentes (VB6, SQL, manual, reuniones)
- Entregables: Markdown técnico (13 secciones) + HTML interactivo con mismo formato que inventario_sgp.html
- Confirmar plan antes de ejecutar tareas grandes (usuario lo solicitó explícitamente)
