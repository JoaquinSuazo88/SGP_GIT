# SGP Producción — Detalle Formularios VB6, SPs, UDFs, Vistas y Tablas

## FORMULARIOS VB6 ANALIZADOS

### M_Plami1.frm (1246 líneas) — Selector/Lanzador Planificación
- Selector modal: abre M_MinTeo (PlaTeo) o M_MinRea (PlaRea)
- Controles: fpText(RUT contrato), fpLongInteger1(1)(régimen), fpLongInteger1(2)(servicio), fpDateTime1(mes/año), vaSpread1(calendario)
- NO usa SPs directos. Usa RutinaLectura.Cliente/Regimen/Servicio/EstServicio/Minutas
- Asigna vars globales: vg_codcasino, vg_codregimen, vg_codservicio, vg_fecha, Vg_FechaDesde, Vg_FechaHasta
- Valida: período no bloqueado (ValidarAccesoMinutaBloqueyBloqueo), min_indblo IN (2,11)
- NO usa BeginTrans VB6

### M_MinRea.frm (5112 líneas) — Editor Planificación Real
- Grilla vaSpread1: 5 cols/día (CodEstructura, NomReceta, NºRaciones, Costo, CodReceta oculto)
- Verde=local/patrón, Amarillo=centralizada 5-etapas (regímenes>9999, solo lectura)
- Parámetros a_param: 'addreceta'(máx recetas adicionales/día, típico=5), '5etapas'
- SP: sgp_Ins_XmlMinutaReal(XMLMinuta, Ceco, CodRegimen, CodServicio, FechaDia INT, Raciones INT, Color VARCHAR(1), Usuario)
  - Valida mermas vs raciones → rechaza si raciones < mermas existentes
  - Valida cambio receta con mermas → rechaza
  - Color=0: preserva costos; Color=1: actualiza b_minutaraciones (PRODUCIDAS)
  - BEGIN TRAN/COMMIT/ROLLBACK internos. Retorna @Num_error, @error, @cRows
  - Tablas: b_minuta, b_minutadet (mid_tipmin='2' DELETE+INSERT), b_minutaraciones
- SP: sgp_Sel_ValidarMinBloque(Ceco, FechaMinuta): valida min_indblo IN (2,11)
- fn_sgp_p_CalculaCosaliCosdes: costo alimento(1) o desechable(2) por receta
- VectorCol(): mapa posición días en grilla. DetallePlantillaMinuta(): carga grilla inicial
- NO usa BeginTrans VB6 (transacciones en SP)

### M_ConRac.frm (2930 líneas) — Control de Raciones
- Grilla clientes×días. Fila 'PRODUCIDAS' protegida con password (a_param 'parcomdia' desencriptado)
- SPs: sgp_Sel_DetalleLecturaxPeriodo(Ceco, YYYYMM), sgp_Del_MinutaRaciones, sgp_Ins_MinutaRaciones
- SPs: sgp_Del_MinutaRacionFacturable(rango), sgp_Ins_MinutaRacionFacturable(Ceco,Reg,Ser,Fecha,Fac)
  - Si Fac=1: DELETE raciones clientes (preserva PRODUCIDAS/PERSONAL/MERMAS) de b_minutaraciones
- SPs: sgp_Sel_MinutaRacionesFacturable, sgp_Sel_MinutaconcomensalesCeroConRac
- b_minutaraciones: mir_rutcli puede ser RUT, 'PRODUCIDAS', 'PERSONAL', 'MERMAS'
- BeginTrans/CommitTrans VB6 en líneas 1908, 1981, 2059, 2068

### M_MerPre.frm (80KB) — Mermas por Preparación
- Grilla 10 cols: CodRec, NomRec, RacionesPlan, CostoUnit, CostoTotal, Merma×Raciones(edit), Merma×KilosServida(edit), CostoMerma(calc), NumLin, MermaBruta×Kilos(edit)
- Controles extra: Desconche/Pan/Produccion (fpDoubleSingle), ChcMerma("No considera Mermas")
- SPs: sgp_Sel_ValidarMinutaMermaPorPreparacion, sgp_Sel_MermaPorPreparacion(cencos,codreg,codser,fecha,codbod)
- SP: sgp_Upd_XmlMermaPreparacion(XML, cencos, codreg, codser, fecha, ConsideraMerma, Desconche FLOAT, Pan FLOAT, Produccion FLOAT, usuario)
  - XML: `<GrabaMerma><Merma CR="codreceta" NM="nummerma" MO="mermaxkilo" MS="mermaservida" NL="numlin"/></GrabaMerma>`
  - UPDATE b_minutadet (mid_nummer, mid_mermaxkilo, mid_mermaxcantservida)
  - INSERT/UPDATE b_mermadesconche (Considera_Merma, Merma_Desconche, Merma_Pan, Merma_Produccion)
- Validación: Col5 Merma×Raciones ≤ mid_numrac. Días < vg_ciedia → bloqueados (rojo)
- NO usa BeginTrans VB6

### I_SalBod.frm (41KB) — Requisición Salida Bodega
- 7 tipos informe: 0=Resumido, 1=xSector, 2=xEstructura Det, 3=xEstructura Res, 4=Resumen, 5=Devolución, 6=MenosDev
- Tipo 2 genera y guarda en BD para SAP. Resto = Crystal Reports
- SP: sgp_DelIns_formatorequesicionestdetallado(XMLServicio, XMLRegimen, cencos, codbod, yyyymmdd_ini, yyyymmdd_fin, usuario)
  - Cálculo: (mid_numrac × red_canpro) / pro_facing. Solo días abiertos. mid_tipmin='2'
- SP: sgp_Upd_ValidarProductoVigente(cencos, codbod)
- XML: `<Servicio><Ser Ser="id"/></Servicio>` y `<Regimen><Reg Reg="id"/></Regimen>`
- Exporta Excel automático si servicios sin comensales (warning) y para tipo 2

### M_RCDiar.frm (59KB) — Cierre Diario
- Calendario mes. Colores: cyan=habilitado, azul=cerrado no enviado, verde=cerrado enviado
- Solo PC en a_param 'SvrAppCont' puede cerrar (valida GetComputerName)
- 14+ validaciones CierrePeriodo antes de cerrar
- b_casinotipoactividades: tipos 1-10 (Proveedores, SalidaProd, Devoluciones, Mermas, RNV, CtrlRaciones, Cafetería, VentaServ, VentaDir, Inventario)
- SP: sgp_Upd_ReabrirCierreDiario(cencos, yyyymmdd): UPDATE ppd_saldo=0, DELETE+INSERT b_productospmpdia
- Cálculo PMP: vg_tipbase='2'→CalcularPMPDiaSql() o CalcularPMPDiaSqlPEL() (si reproceso SAP)
- Si inventario rotativo: UPDATE b_productospmpdia SET ppd_saldo = tin_stofis
- Días feriados b_Fecha_Inhabiles → recalcula PMP por cada feriado

### M_VtaCon.frm (81KB) — Venta Servicio Contado
- Grilla calendario monto×día. Tab2=Detalle Centro Costo (si cliente es de CoCo en b_clientecencos)
- 5 formas pago: Contado(0), Cheque(1), Cheque Restaurant(2), TarjetaCrédito(3), Vale(4)
- Tablas: b_ventacontado(vtc_codigo, vtc_forpag, vtc_totmon, vtc_opccli), b_ventacontadodet
- NO usa SPs (SQL inline). BeginTrans/CommitTrans VB6 en Borrar y Confirmar

### M_SalBod.frm (3564 líneas) — Salida Bodega a Producción
- Tipo documento: SP. Correlativo en b_parametros par_tipdoc='SP'
- Lee minuta real (b_minutadet mid_tipmin='2') O estructura fija (b_minutafijadia)
- Calcula composición: raciones × red_canpro / rec_basrac. Busca PMP en b_productospmpdia
- Tablas: b_totventas(tov_tipdoc='SP'), b_detventas, b_bodegas(UPDATE -canmer)
- SP: sgp_Sel_ValidarDevolucionProduccion(Ceco, codreg, codser, codbod, Fecha, numdoc)
- BeginTrans/CommitTrans VB6. 2 modos: A=INSERT, M=DELETE+INSERT. Control stock negativo=ROLLBACK

### M_VenCaf.frm (2089 líneas) — Venta Cafetería
- SSTab: Tab1=Venta Cafetería, Tab2=Inventario Producto
- Tipos pago: CR=Crédito(requiere cliente), CU=Cuenta(requiere cliente), CA=Contado
- Tablas: b_totventascaf(tvc_estado=''/'C'), b_detventascaf, b_detventascafpro, b_bodegas
- Op=1 (Cerrar): UPDATE tvc_estado='C' + decrementa b_bodegas. Op=2 (Reabrir): revierte stock
- BeginTrans/CommitTrans VB6

### M_VenDir.frm (1447 líneas) — Venta Directa
- Tablas: b_totventas, b_detventas, b_bodegas. BeginTrans/CommitTrans VB6
- Control stock visual: color azul=sobrepasa stock

### M_Produ1.frm (2741 líneas) — Árbol Ingrediente (SOLO LECTURA)
- Tablas lectura: b_ingrediente, b_productos, a_unidad, b_productospmpdia

### M_TabGra.frm (1782 líneas) — Tabla Gramaje
- TreeView zona→sub-segmento→ing.receta→régimen + grilla
- SP: sgpadm_s_zona(6, 0, '')

---

## STORED PROCEDURES DOCUMENTADOS

### Planificación Real
- sgp_Ins_XmlMinutaReal: XML→b_minuta+b_minutadet+b_minutaraciones. BEGIN/COMMIT/ROLLBACK
- sgp_Ins_XmlMinutaTeorica: mid_tipmin='1', sin validaciones merma
- sgp_Sel_ValidarMinBloque: SELECT b_minuta WHERE min_indblo IN (2,11)
- sgp_Sel_ValidarMinutaBloqueModificado / sgp_Sel_ValidarMinutaBloqueNuevo

### Control de Raciones
- sgp_Del_MinutaRaciones / sgp_Ins_MinutaRaciones: DELETE/INSERT simple b_minutaraciones
- sgp_Del_MinutaRacionFacturable: DELETE rango b_minutaracionfacturable
- sgp_Ins_MinutaRacionFacturable: Si Fac=1 → DELETE raciones clientes (preserva PRODUCIDAS/PERSONAL/MERMAS)
- sgp_Sel_DetalleLecturaxPeriodo: GROUP BY b_detallelectura por período
- sgp_Sel_MinutaRacionesFacturable, sgp_Sel_MinutaconcomensalesCeroConRac
- sgp_Sel_ValidarRacionesProducidas: min_racrea=0, excluye ser 11056/11057, excluye facturadas
- sgp_Sel_ValidarClienteSinraciones: clientes en b_preciovta sin raciones en el período

### Mermas
- sgp_Upd_XmlMermaPreparacion: UPDATE b_minutadet + INSERT/UPDATE b_mermadesconche
- sgp_Sel_MermaPorPreparacion: CantBruta = CASE mid_nummer>0 AND mid_mermaxkilo=0 THEN SGP_FN_RNVCantidadesReceta*mid_nummer ELSE mid_mermaxkilo
- sgp_Sel_ValidarMinutaMermaPorPreparacion, sgp_Sel_DetalleMermasCierreDiario, sgp_Sel_MermaDesconcheCierreDiario

### Requisición
- sgp_DelIns_formatorequesicionestdetallado: calcula (mid_numrac×red_canpro)/pro_facing
- sgp_Sel_formatorequesicionestdetallado, sgp_Sel_ValidarServicioComensalesCeroSalProduccion

### Salida Bodega
- sgp_Sel_ValidarDevolucionProduccion: verifica DP para SP
- sgp_Sel_CalcularSalidaProduccionMinutaTeorica: compara teórico vs realizado
- FLMS_SGP_ValidaCargaSalidasDeBodega: Estado=1 OK, Estado=3 error
- FLMS_SGP_ValidaCargaDevolucionesDeBodega, FLMS_SGP_ValidaCargaMermasDeBodega
- FLMS_SGP_ReprocesaCargaSalidasDeBodega: reprocesa Estado=3

### Cierre Diario
- sgp_Upd_ReabrirCierreDiario: UPDATE ppd_saldo=0, DELETE+INSERT b_productospmpdia
- sgp_Sel_CierrePeriodo, sgp_Sel_TraerCostoFoodCostMinutaCierreDiario (#TempPaso teórico+real+vendido)
- sgp_Sel_ValidarInventarioCalendarizado, sgp_Upd_ValidarInventarioCalendarizado
- FLMS_SGP_ValidaRacionesNoVendidas: CURSOR b_minuta_SGP_FLMS, estado=1/3

### Integración FLMS
- ProcesoDeIntegracion_FLMStoSGP: proceso principal integración
- FLMS_SGP_IU_LogIntegraMermaMin_Out, FLMS_SGP_IU_LogIntegraOut

---

## FUNCIONES SQL (UDFs)
- fn_sgp_p_CalculaCosaliCosdes(Op, CenCos, CodRec, TipRec, TipMin, Fecha): costo alimento(1)/desechable(2). ctainsumo/ctalimdes. Receta patrón: cencos='0'
- fn_sgp_Pro_TraerDiaSeguridad(codigo, Ceco): días seguridad, árbol a_tipopro→b_paramdesp
- FLMS_SGP_IdentificaErrores: 9 validaciones (1=ceco, 2=tipdoc, 4=régimen, 5=servicio, 6=total≠suma, 7=ingrediente, 8=producto, 9=cantidad<0)
- sgp_p_desencripta(@psw_encripta): CHAR(ASCII(char) - 73 - posición). Desencripta ciediario y parcomdia

## VISTAS SQL
- Sel_FechaUltimoCierre: FechaCierre (ayer), FechaProceso (hoy), desencripta 'ciediario'
- Sel_Minuta_Planificada: minutas reales con ponderación ROUND((mid_numrac/min_racrea)*100,0)
- Sel_Precio_Stock_SGP: precios+stocks diarios (ppd_fecdia=último cierre, bod_canmer>0)
- Sel_Producto_Ingrediente_SGP: producto→ingrediente con PMP, JOIN b_productosing+b_contlistpreing+b_productospmpdia

---

## TABLAS PRINCIPALES
- b_minuta: cencos, fecmin(INT YYYYMMDD), codreg, codser, indblo(0/2/11), racrea, racteo, tipmin('1'/'2')
- b_minutadet: FK→b_minuta, codrec, numlin, numrac, cosrec, cosdes, tipmin, nummer, mermaxkilo, mermaxcantservida, tiprec, estser
- b_minutaraciones: cencos, codreg, codser, rutcli(RUT/'PRODUCIDAS'/'PERSONAL'/'MERMAS'), fecmin, nrorac
- b_minutaracionfacturable: cencos, codreg, codser, fecmin, facturado(0/1)
- b_mermadesconche: IdCeco, IdRegimen, IdServicio, Fecha_Merma(INT), Considera_Merma, Merma_Desconche, Merma_Pan, Merma_Produccion, Usuario
- b_formatorequesicionestdetallado: requisición guardada para SAP
- b_ventacontado: vtc_codigo, cencos, codreg, codser, fecvta, forpag, totmon, opccli
- b_ventacontadodet: detalle por centro costo
- b_cierreperiodo: cencos, periodo, estado(1=abierto, 0=cerrado), fecter
- b_parametros: par_correlativo por tipo/bodega (SP/DP/ME/SE/DE)
- a_param: cencos, codigo, valor (cifrado para 'ciediario' y 'parcomdia', 'SvrAppCont', 'addreceta', '5etapas', 'pargrarnve', 'ctainsumo', 'ctalimdes')
- log_enviocierrediario, log_cierrediario, b_Fecha_Inhabiles
- b_productospmpdia: cencos, fecdia, codpro, propon(PMP), saldo
- b_casinotipoactividades: tipo 1-10 (validaciones cierre diario)
- b_casinoservicioprincipales: IdCeco, IdRegimen, IdServicio, Activo, Preferido
- b_detallelectura: lecturas vales por punto
- b_totventas_SGP_FLMS: tabla interfaz FLMS, Estado NULL/1/3
