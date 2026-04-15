# DRF - KPI


---

## Índice

- [Historial de Versiones](#historial-de-versiones)
- [Confidencialidad](#confidencialidad)
- [Información del Proyecto](#información-del-proyecto)
- [Responsables](#responsables)
- [Situación Actual](#situación-actual)
- [Propósito del proyecto](#propósito-del-proyecto)
- [Alcance del proyecto](#alcance-del-proyecto)
- [Reporte A13](#reporte-a13)
- [Costo Planificado/teórico/Planificado Real/ Realizado Alimentación y Desechables](#costo-planificadoteóricoplanificado-real-realizado-alimentación-y-desechables)
- [Comensales Minuta Bloque](#comensales-minuta-bloque)
- [Ajuste de Inventario](#ajuste-de-inventario)
- [Compras por Periodo](#compras-por-periodo)
- [Consumo Proyectado Real](#consumo-proyectado-real)
- [Costo Bandeja Planificado](#costo-bandeja-planificado)
- [Detalle Inventario](#detalle-inventario)
- [Jerarquías Waste Watch](#jerarquías-waste-watch)
- [Merma Bodega](#merma-bodega)
- [Merma Desconche Producción](#merma-desconche-producción)
- [Raciones No Vendidas (Merma Línea)](#raciones-no-vendidas-merma-línea)
- [Reporte Food Cost](#reporte-food-cost)
- [Reporte WW Global](#reporte-ww-global)
- [Respaldo2021- 2022](#respaldo2021--2022)
- [Traspasos desde la CD](#traspasos-desde-la-cd)
- [Último Cierre](#último-cierre)

---


# Historial de Versiones


| Versión | Fecha | Autor | Descripción |
| --- | --- | --- | --- |
| **1.0** | 02-03-2026 | Dania Contreras | Primera Versión |
| **2.0** | 11-03-2026 | Dania Contreras | Segunda Versión |


# Confidencialidad


La información de este documento y documentos anexos es propiedad de **SODEXO CHILE** y de carácter confidencial, por lo cual el proveedor debe mantener la información en reserva y usarla sólo para el propósito de prestar los servicios solicitados.


El proveedor se obliga además a tomar las medidas para que quienes tengan acceso a la Información, guarden bajo estricta reserva, protejan y no revelen a terceros dicha Información, siendo responsabilidad del proveedor velar por el cumplimiento de esta obligación.


En caso de avanzar con el proyecto, el proveedor deberá firmar un documento de Confidencialidad de la Información (NDA Sodexo), donde se describe con mayor detalle estas obligaciones.


Toda la información entregada por el proveedor para la evaluación de un servicio, sistema y/o solución informática será propiedad de **SODEXO CHILE**, sin que esto signifique un costo o genere algún tipo de cargo para la empresa.


# Información del Proyecto


| Estructura | Descripción |
| --- | --- |
| Segmento | Servicios de alimentación / operación de Casinos |
| Área | Área reportería |
| Sección | Levantamiento KPI |
| Proyecto | SGP Upgrade |


# Responsables


| ROL | Nombre | Correo Electrónico |
| --- | --- | --- |
| Sponsor | Francisco González | [francisco.gonzalez@sodexo.com](mailto:francisco.gonzalez@sodexo.com) |
| Líder Proyecto | Claudia Muñoz | [Claudia.munoz@sodexo.com](mailto:Claudia.munoz@sodexo.com) |
| Key User | Andrés Jimenes | [Andres.jimenez@sodexo.com](mailto:Andres.jimenez@sodexo.com) |
| Líder TI | Dania Contreras | [marcelo.gonzalez@sodexo.com](mailto:marcelo.gonzalez@sodexo.com) |


# Situación Actual


Actualmente, la información de los KPIs operativos se obtiene a través de procedimientos almacenados en SQL Server que son ejecutados diariamente mediante paquetes SSIS. Estos procedimientos extraen, transforman y consolidan datos desde múltiples tablas del sistema SGP Administrador y local, generando archivos de salida en formato delimitado por pipe (|) que son consumidos por Power BI para su visualización en dashboards. Estos archivos son disponibilizados en un sharepoint.


Los documentos en sharepoint son un repositorio de información, por lo que las fórmulas o los cálculos deben ver en el documento respectivo.


# Propósito del proyecto


El propósito de este documento es realizar un levantamiento detallado de la información contenida en cada uno de los reportes Excel generados por el sistema SGP Administrador y Local. Este levantamiento consiste en revisar reporte por reporte para identificar y documentar:

- Qué información presenta cada reporte y cuál es su significado de negocio.
- Qué parámetros controlan el comportamiento de cada reporte.
- Qué reglas de negocio están implícitas en la lógica de los procedimientos almacenados.

# Alcance del proyecto


El proyecto abarca el levantamiento de todos los reportes Excel generados por el sistema SGP Administrador y Local que actualmente alimentan los dashboards de Power BI.


Por la naturaleza de la información que contiene el SharePoint es importante considerar accesos diferenciados por usuario.


# Reporte A13


## Propósito Funcional


El reporte A13 es el estado de resultados operativo de cada punto de servicio de Sodexo. Funciona como una radiografía financiera mensual que reúne en un solo lugar toda la información económica del sitio.


## Tabla


| **Campo Destino (CSV)** | **Descripción** | **Origen** | **OK/NO OK** |
| --- | --- | --- | --- |
| Profit Center | Código del Centro de Beneficio (Profit Center) en el sistema SAP | Mantenedor de Centros de Costos |  |
| Ceco | Código del Centro de Beneficio (Profit Center) en el sistema SAP | Mantenedor de Centros de Costos |  |
| Descrip. Ceco | Nombre descriptivo del Centro de Costo | Mantenedor de Centros de Costos |  |
| Periodo | Período contable al que corresponden los datos, expresado como un número entero en formato Año+Mes | Cas_B_A13InsumosFCost |  |
| Fecha Cierre | Fecha en que se cerró/consolidó el período contable para ese CECO. | Cas_B_A13InsumosFCost |  |
| Glosa | Comentario, observación o descripción adicional sobre los datos del período |  |  |
| Alimentos | Costo total de alimentos e insumos consumidos en el período para ese CECO. |  |  |
| Cant./Lim_Desc | Cantidad de raciones. Contar cuántos servicios se vendieron. |  |  |
| P. Vta./Total | Campo va a depender de la glosa |  |  |
| Total/% | Campo va a depender de la glosa |  |  |


## Glosa


Cada línea tiene un propósito diferente. Los mismos campos (columnas) se reutilizan para distintos tipos de información según qué tipo de registro sea (identificado por la Glosa).


| **Glosa** | **Alimentos** | **Cant./Lim_Desc** | **P. Vta./Total** | **Total/%** | **OK/NO OK** |
| --- | --- | --- | --- | --- | --- |
| Limpieza y Desechables |  |  |  | $Total Limpieza y Desechables |  |
| Venta Contado |  |  |  | $Total Ventas Contado |  |
| Venta Servicio Especial |  |  |  | $Total Ventas Servicios Especiales |  |
| Tipo de Servicio | - | # Cantidad de raciones vendidas. | $Precio por raciones vendidas. | Cantidad de raciones vendidas* Precio por raciones vendidas |  |
| Total | - |  | - | Ventas Totales |  |
| Inventario Inicial | $ Valor Alimentos | $Valor otros Insumos | $Valor Alimentos + $Valor otros insumos | ($Valor Alimentos + $Valor otros insumos)/ Ventas Totales |  |
| Centralización Compras | $Compras de Alimentos | $Compras de otros insumos | $Compras de Alimentos + $Compras de otros insumos | ($Compras de Alimentos + $Compras de otros insumos)/ Ventas Totales |  |
| Compras Fofis <br> --- <br> 💬 **Comentario — Contreras Dania (2026-01-19):** Compras fuera de stock. Compras de emergencia. | $Compras de Alimentos Fofis | $Compras de otros insumos Fofis | $Compras de Alimentos Fofis + $Compras de otros insumos Fofis | ($Compras de Alimentos Fofis + $Compras de otros insumos Fofis)/ Ventas Totales |  |
| Compras No Estoqueable <br> --- <br> 💬 **Comentario — Contreras Dania (2026-03-02):** Productos que no controlan stock. | $Compras de Alimentos No estoqueable | $Compras de otros insumos No estoqueable | $Compras de Alimentos No estoqueable + $Compras de otros insumos No estoqueable | ($Compras de Alimentos No estoqueable + $Compras de otros insumos No estoqueable)/ Ventas Totales |  |
| Traspasos Recibido | $Alimentos Traspasados | $Otros insumos Traspasados | $Alimentos Traspasados + $Otros insumos Traspasados | ($Alimentos Traspasados + $Otros insumos Traspasados) / Ventas Totales |  |
| Costo Logistico | $Costo de transporte |  | $Costo de transporte | $Costo de transporte/ Ventas Totales |  |
| Traspasos Emitidos | - $Alimentos Traspasados emitidos | - $Otros insumos Traspasados emitidos | - $Alimentos Traspasados emitidos - $Otros insumos Traspasados emitidos | (- $Alimentos Traspasados emitidos - $Otros insumos Traspasados emitidos)/ Ventas Totales |  |
| Traspaso Prod. Term. | - $Alimentos Traspasados Producto Terminado | - $Otros insumos Traspasados Producto Terminado | - $Alimentos Traspasados Producto Terminado - $Otros insumos Traspasados Producto Terminado | (- $Alimentos Traspasados Producto Terminado - $Otros insumos Traspasados Producto Terminado)/ Ventas Totales |  |
| Mermas Bodega | - $Merma de Alimentos | - $Merma de Otros Insumos | - $Merma de Alimentos - $Merma de Otros Insumos | (- $Merma de Alimentos - $Merma de Otros Insumos) / Ventas Totales |  |
| Salida Producción <br> --- <br> 💬 **Comentario — Contreras Dania (2026-01-21):** Es la sumatoria de todo lo entregado y digitado de bodega a producción. | - $Alimentos utilizados en producción | - $Otros insumos utilizados en producción | - $Salida Producción Total = - $Alimentos utilizados en producción - $Otros insumos utilizados en producción | (- $Alimentos utilizados en producción - $Otros insumos utilizados en producción) / Ventas Totales |  |
| Devolución Producción | $Alimentos Devueltos | $Otros insumos Devueltos | $Devolución Producción Total = $Alimentos Devueltos + $Otros insumos Devueltos | ($Alimentos Devueltos + $Otros insumos Devueltos)/ Ventas Totales |  |
| Ajuste Inventario | $Alimentos Devueltos Ajuste de Inventario | $Otros insumos Ajuste de Inventario | $Alimentos Devueltos Ajuste de Inventario + $Otros insumos Ajuste de Inventario | ($Alimentos Devueltos Ajuste de Inventario + $Otros insumos Ajuste de Inventario)/ Ventas Totales |  |
| Toma Inventario | $Alimentos Toma Inventario | $Otros insumos Toma Inventario | $Alimentos Toma Inventario + $Otros insumos Toma Inventario | ($Alimentos Toma Inventario + $Otros insumos Toma Inventario)/  Ventas Totales |  |
| F.Cost (Sal. & Dev.) | (- $Salida de Producción Alimentos + $Devolución Alimentos)/Ventas Totales | (- $Salida de Producción Otros Insumos + $Devolución Otros Insumos)/ Ventas Totales | - $Salida Producción Total + $Devolución Producción Total | - $Salida Producción Total + $Devolución Producción Total Ventas Totales |  |
| F.Cost(Sal. & Dev. Vta.Ser.Esp) | (- $Salida de Producción Alimentos Venta Servicios Especiales + $Devolución Alimentos Venta Servicios Especiales)/Ventas Totales | (- $Salida de Producción Otros Insumos Venta Servicios Especiales + $Devolución Otros Insumos Venta Servicios Especiales)/Ventas Totales | -$Salida de Producción Venta Servicios Especiales Total + $Devolución Alimentos Ventas Servicios Especiales Total | -$Salida de Producción Venta Servicios Especiales Total + $Devolución Alimentos Ventas Servicios Especiales Total/ Ventas Totales |  |
| F.Cost (Merm. & Aju.) | - $Merma de Alimentos + $Alimentos Devueltos Ajuste de Inventario/Ventas Totales | - $Merma de Otros Insumos + $Otros Insumos Devueltos Ajuste de Inventario/Ventas Totales | - $Merma Total + $Otros Total Ajuste de Inventario | - $Merma Total + $Otros Total Ajuste de Inventario/ Ventas Totales |  |
| F.Cost Tras.Prod.Term. | -Traspaso Prod. Term./ Ventas Totales |  | -Traspaso Prod. Term | -Traspaso Prod. Term./ Ventas Totales |  |
| F.Cost C. No Estoq. | Compras No Estoqueable./ Ventas Totales |  | Compras No Estoqueable. | Compras No Estoqueable./ Ventas Totales |  |
| F.Cost Flete Insumo |  |  |  |  |  |
| F.Cost Logistico | $Costo de transporte/ Ventas Totales |  | $Costo de transporte | $Costo de transporte/ Ventas Totales |  |
| F.Cost Total | $Food Cost Total Alimentos/ Ventas Totales | $Food Cost Total Otros Insumos/ Ventas Totales | $Food Cost Total Alimentos + $Food Cost Total Otros Insumos | $Food Cost Total Alimentos + $Food Cost Total Otros Insumos/ Ventas Totales |  |
| Totales Nominales | $Food Cost Total Alimentos | $Food Cost Total Otros Insumos | $Food Cost Total Alimentos + $Food Cost Total Otros Insumos | $Food Cost Total Alimentos + $Food Cost Total Otros Insumos/ Ventas Totales |  |
| Nº de días Trabajados |  |  |  | Cantidad de N° de días Trabajador |  |
| Nº de días de Stock |  |  |  | Cantidad de Nº de días de Stock |  |
| Raciones no Vendidas |  |  | $Raciones no Vendidas | $Raciones no Vendidas/ Ventas Totales |  |
| AGUA |  |  |  | Total Gastos AGUA |  |
| ASESORIAS - 410028 |  |  |  | Total Gastos ASESORIAS - 410028 |  |
| ANALISIS DE LABORATORIO - 410037 |  |  |  | Total Gastos ANALISIS DE LABORATORIO - 410037 |  |
| ALIMENTO DEL PERSONA - 410002 |  |  |  | Total Gastos ALIMENTO DEL PERSONA - 410002 |  |
| ARQUITECTURA/OBRAS CIVILES - 410143 |  |  |  | Total Gastos ARQUITECTURA/OBRAS CIVILES - 410143 |  |
| ARRENDAMIENTOS OFICINAS - 410065 |  |  |  | Total Gastos ARRENDAMIENTOS OFICINAS - 410065 |  |
| ARRIENDO EQ.COMPUTAC - 410071 |  |  |  | Total Gastos ARRIENDO EQ.COMPUTAC - 410071 |  |
| ARRIENDO VARIOS - 410146 |  |  |  | Total Gastos ARRIENDO VARIOS - 410146 |  |
| ARRIENDO EQUIPOS MENORES - 410145 |  |  |  | Total Gastos ARRIENDO EQUIPOS MENORES - 410145 |  |
| ARRIENDOS DE BIENES - 410066 |  |  |  | Total Gastos ARRIENDOS DE BIENES - 410066 |  |
| ARRIENDOS DEVEHICULOS Y MAQUINARIAS - 410067 |  |  |  | Total Gastos ARRIENDOS DEVEHICULOS Y MAQUINARIAS - 410067 |  |
| BAM Y CONECTIVIDAD - 410166 |  |  |  | Total Gastos BAM Y CONECTIVIDAD - 410166 |  |
| BIENESTAR - 410020 |  |  |  | Total Gastos BIENESTAR - 410020 |  |
| COMUNIC.INTERNAS/VALIJA - 410041 |  |  |  | Total Gastos COMUNIC.INTERNAS/VALIJA - 410041 |  |
| CONT.DE SOP.Y MANT. - 410077 |  |  |  | Total Gastos CONT.DE SOP.Y MANT. - 410077 |  |
| CONTROL DE PLAGAS - 410144 |  |  |  | Total Gastos CONTROL DE PLAGAS - 410144 |  |
| CORRESPONDENCIA / VALIJA / VALORES - 410053 |  |  |  | Total Gastos CORRESPONDENCIA / VALIJA / VALORES - 410053 |  |
| ENLACES - 410076 |  |  |  | Total Gastos ENLACES - 410076 |  |
| EXAMENES AL PERSONAL Y GTOS POLICLINICO - 410141 |  |  |  | Total Gastos EXAMENES AL PERSONAL Y GTOS POLICLINICO - 410141 |  |
| FLETES - 410036 |  |  |  | Total Gastos FLETES - 410036 |  |
| FLETES NO INSUMOS - 410155 |  |  |  | Total Gastos FLETES NO INSUMOS - 410155 |  |
| FOTOCOPIAS - 410032 |  |  |  | Total Gastos FOTOCOPIAS - 410032 |  |
| GTOS.ANIMACION - 410051 |  |  |  | Total Gastos GTOS.ANIMACION - 410051 |  |
| GASTOS BANCARIOS - 410058 |  |  |  | Total Gastos GASTOS BANCARIOS - 410058 |  |
| GASTOS DE IMPORTACION - 410142 |  |  |  | Total Gastos GASTOS DE IMPORTACION - 410142 |  |
| GASTOS COMPUTACIONAL - 410027 |  |  |  | Total Gastos GASTOS COMPUTACIONAL - 410027 |  |
| GTO. PUBLICIDAD CAS - 410049 |  |  |  | Total Gastos GTO. PUBLICIDAD CAS - 410049 |  |
| GTOS.CAPAC.PERS. - 410022 |  |  |  | Total Gastos GTO. PUBLICIDAD CAS - 410049 |  |
| GTOS.REPRESENTACION - 410050 |  |  |  | Total Gastos GTOS.REPRESENTACION - 410050 |  |
| GASTOS NOTARIALES - 410059 |  |  |  | Total Gastos GASTOS NOTARIALES - 410059 |  |
| GTOS.VARIOS - 410099 |  |  |  | Total Gastos GTOS.VARIOS - 410099 |  |
| HONOR. PERS. - 410019 |  |  |  | Total Gastos HONOR. PERS. - 410019 |  |
| GTOS.VEHIC.ACEPTADOS - 410046 |  |  |  | Total Gastos GTOS.VEHIC.ACEPTADOS - 410046 |  |
| GTOS.VEHIC.RECHAZADO - 410047 |  |  |  | Total Gastos GTOS.VEHIC.RECHAZADO - 410047 |  |
| IMP.Y UTIL.ESCRIT. - 410026 |  |  |  | Total Gastos IMP.Y UTIL.ESCRIT. - 410026 |  |
| INSUMOS SERV.GLOBAL - 410003 |  |  |  | Total Gastos INSUMOS SERV.GLOBAL - 410003 |  |
| LAVANDERIA - 410034 |  |  |  | Total Gastos LAVANDERIA - 410034 |  |
| MOVILIZACION LOCAL - 410154 |  |  |  | Total Gastos MOVILIZACION LOCAL - 410154 |  |
| MOVILIZACION - 410042 |  |  |  | Total Gastos MOVILIZACION - 410042 |  |
| PATENTES MUNICIP / C - 410033 |  |  |  | Total Gastos PATENTES MUNICIP / C - 410033 |  |
| PEQ. MATERIAL - 410024 |  |  |  | Total Gastos PEQ. MATERIAL - 410024 |  |
| RELACIONES PUBLICAS Y O. - 410052 |  |  |  | Total Gastos RELACIONES PUBLICAS Y O. - 410052 |  |
| REP.Y MANT.EQ.COMP. - 410072 |  |  |  | Total Gastos REP.Y MANT.EQ.COMP. - 410072 |  |
| REPARAC.Y MANTEN. - 410044 |  |  |  | Total Gastos REPARAC.Y MANTEN. - 410044 |  |
| SEGUROS - 410030 |  |  |  | Total Gastos SEGUROS - 410030 |  |
| SERVICIO PAGO REMUNERACIONES - 410031 |  |  |  | Total Gastos SERVICIO PAGO REMUNERACIONES - 410031 |  |
| SUB-CONTRATOS DE SERVICIOS - 410126 |  |  |  | Total Gastos SUB-CONTRATOS DE SERVICIOS - 410126 |  |
| TELECOMUNICACIONES - 410054 |  |  |  | Total Gastos TELECOMUNICACIONES - 410054 |  |
| TELEFONIA FIJA - 410148 |  |  |  | Total Gastos TELEFONIA FIJA - 410148 |  |
| TELEFONIA MOVIL - 410165 |  |  |  | Total Gastos TELEFONIA MOVIL - 410165 |  |
| UNIFORMES - 410025 |  |  |  | Total Gastos UNIFORMES - 410025 |  |
| VIAJES/ESTADIAS INTERNAC - 410040 |  |  |  | Total Gastos VIAJES/ESTADIAS INTERNAC - 410040 |  |
| VIAJES/ESTADIAS NACIONAL - 410039 |  |  |  | Total Gastos VIAJES/ESTADIAS NACIONAL - 410039 |  |
| VAJILLA - 410029 |  |  |  | Total Gastos VAJILLA - 410029 |  |
| Total Gastos Generales |  |  |  | Total Gastos Generales |  |
| 1 Depreciación |  |  | $Depreciación | $Depreciación/ Ventas Totales |  |
| 2 Gestión Personal |  |  | $Gestión Personal |  |  |
| 3 Cuota Negociación |  |  | $Cuota Negociación |  |  |
| 4 Bono vac. (Bienestar) |  |  | $Bono vac. (Bienestar) |  |  |
| 5 Cuota Dirigente Sindical <br> --- <br> 💬 **Comentario — Contreras Dania (2026-02-02):** Todos los sitios están en 0. Mantener? |  |  | $Cuota Dirigente Sindical |  |  |
| Total Costo Personal |  |  | Total Costo Personal | Total Costo Personal/ Ventas Totales |  |
| 6 Nº Horas Extra |  |  |  | # Nº Horas Extra |  |
| 7 Nº Dias Trabajados |  |  |  | $ Nº Dias Trabajados |  |
| 8 Nº Dias Ausencia |  |  |  | # Nº Dias Ausencia |  |
| TOTAL DE GASTOS |  |  | - $Salida Producción Total + $Devolución Producción Total + Mermas Totales + Ajuste de Inventario | TOTAL DE GASTOS |  |
| UTILIDAD - OPERACIONAL |  |  |  | UTILIDAD - OPERACIONAL |  |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio activos y operativos. Únicamente se incluyen centros de costo que están habilitados, dados de alta en el sistema organizacional (PEL) y que no hayan sido marcados como desactivados. Esto asegura que el reporte refleje solo operaciones vigentes.
- Estructura por glosas (tipo de registro). Cada línea del reporte tiene un significado diferente según su glosa. Las mismas columnas (Alimentos, Limpieza/Desechables, Total, Porcentaje) cambian de significado dependiendo del tipo de registro. Por ejemplo, en la glosa "Inventario Inicial" las columnas muestran valores de inventario, pero en "Mermas" muestran costos de perdidas.
- Reutilización de columnas El reporte usa un formato compacto donde las 4 columnas de valores (Alimentos, Cant./Lim_Desc, P.Vta./Total, Total/%) se reutilizan para distintos propósitos según la glosa. Esto permite consolidar muchos indicadores en un solo dataset.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.

# Costo Planificado/teórico/Planificado Real/ Realizado Alimentación y Desechables


## Propósito Funcional


El reporte de Costo Planificado/Teorico/Real/Realizado es una herramienta de control de desviaciones que compara lo que se planeó gastar contra lo que realmente se gastó en la producción de cada punto de servicio.


El reporte compara tres tipos de costos por centro de costo (CECO), régimen y servicio para análisis de desviaciones en Power BI:


| **Tipo de Costo** | **Descripción** |
| --- | --- |
| Teórico | Costo planificado al momento de que se planifica la minuta. |
| Real | Costo de la minuta real. |
| Realizado | Costo realmente producido. (Salida y devolución de producción) |


Al comparar estos tres valores, el reporte permite detectar dónde se está ahorrando y dónde se está gastando de más. Por ejemplo, si el costo realizado es mayor que el teórico, significa que ese centro de costo está gastando más de lo planificado, lo cual es una señal de alerta para la gestión.


Del proceso se obtienen 2 archivos por separados, alimentación y desechables.


## Tabla


| **Campo Destino (CSV)** | **Descripción** | **OK/NO OK** |
| --- | --- | --- |
| Profit Center | Código del Profit Center en AX |  |
| Ceco | Código del Centro de Costo SAP |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |
| Regimen | Código del tipo de régimen alimenticio |  |
| Descripcion Regimen | Descripción del régimen |  |
| Servicio | Código del servicio |  |
| Descripcion Servicio | Descripción del servicio |  |
| Fecha | Fecha Minuta YYYYMMDD |  |
| Periodo | Período contable en formato YYYYMM |  |
| Costo Bandeja Teórico | Costo planificado por bandeja. |  |
| Nro. Rac. Teórica | Cantidad de raciones que se planificó producir. |  |
| Costo Total Teórico | Costo total planificado para todas las raciones teóricas. (Costo Bandeja teórico x Nro. Rac. teórica) |  |
| Costo Bandeja Real | Costo bandeja de lo que se va a producir. |  |
| Nro. Rac. Real | Cantidad de raciones que se van a vender. |  |
| Costo Total Real | Costo bandeja real. (Costo Bandeja Real x Nro. Rac. Real) |  |
| Desviacion C. Ban. Plan. | Diferencia entre el Costo Bandeja Real y el Costo Bandeja Teórico por bandeja. Mide el impacto de cambio de precio de la bandeja. |  |
| Costo Bandeja Realizado | Costo producido por bandeja. |  |
| Nro. Rac. Realizado | Cantidad de raciones realmente producidas. |  |
| Costo Total Realizado | Costo Total Realizado. (Costo Bandeja Realizado x Nro. Rac. Realizado) |  |
| Desviacion C. Ban. Realizado | Diferencia entre el Costo Bandeja Realizado y el Costo Bandeja Real por bandeja. Mide el impacto de cambio de precio de la bandeja. |  |


## Reglas del negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio activos y operativos. Se incluyen únicamente centros de costo habilitados, dados de alta en el sistema organizacional y no eliminados.
- Separación entre alimentos y desechables. El reporte se ejecuta por separado para alimentos (insumos de comida) y desechables (articulos de limpieza y descartables), permitiendo analizar cada tipo de costo de forma independiente.
- Días sin actividad aparecen en cero. Los fines de semana, feriados o días sin producción aparecen con valores en cero. No se eliminan del reporte para mantener la continuidad del calendario.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.

# Comensales Minuta Bloque


## Propósito Funcional


El reporte de Comensales Minuta Bloque muestra la planificación de comensales de cada punto de servicio. Refleja cuantas raciones se programaron preparar cada día según la minuta establecida, desglosado por tipo de régimen alimenticio y servicio.


## Tabla


| **Campo Destino** | **Descripción** | **Tabla** | **Origen** | **OK/NO OK** |
| --- | --- | --- | --- | --- |
| Profit Center | Código del Profit Center en AX |  |  |  |
| Ceco | Código del Centro de Costo SAP |  |  |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |  |  |
| Regimen | Código del tipo de régimen alimenticio |  |  |  |
| Descripcion Regimen | Descripción del régimen |  |  |  |
| Servicio | Código del servicio |  |  |  |
| Descripcion Servicio | Descripción del servicio |  |  |  |
| Fecha Minuta | Fecha de la minuta planificada |  |  |  |
| Periodo | Período contable en formato YYYYMM |  |  |  |
| Comensales | Cantidad de comensales planificados | Documento Minuta |  | Viene |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio reales y activos. Se excluyen los centros de costo que sean de prueba, propuestas comerciales o diseños. Solo se incluyen sitios que están efectivamente operando y dados de alta en el sistema organizacional.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”. se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.

# Ajuste de Inventario


## Propósito Funcional


El reporte de Ajuste de Inventario registra las diferencias entre el inventario esperado y el inventario real de cada punto de servicio. Cuando al hacer un conteo físico se encuentra más o menos producto del que debería haber, se genera un ajuste que queda documentado en este reporte.


## Tabla


| **Campo Destino** | **Descripción** | **Origen** | **OK/NO OK** |
| --- | --- | --- | --- |
| Profit Center | Código del Profit Center en AX |  |  |
| Ceco | Código del Centro de Costo SAP |  |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |  |
| Periodo | Período contable en formato YYYYMM |  |  |
| F_Movimiento | Fecha de movimiento. |  |  |
| Familia | jerarquía completa de la familia del producto. Muestra el árbol de clasificación desde la categoría raíz hasta la subfamilia, separado por delimitadores. Ejemplo: Alimentos > Carnes > Res. | Maestro de Producto |  |
| Codigo Producto | Código producto SGP. |  |  |
| Nombre Producto | Nombre descriptivo del producto. |  |  |
| Precio | Precio de costo unitario del producto. |  |  |
| Unidad | Nombre de la unidad de medida del producto. (Unidad/Litro) |  |  |
| Cantidad | Cantidad total de mercadería ajustada. |  |  |
|  | Monto total valorizado del ajuste. Se calcula como la suma de precio * cantidad aplicando un signo según el tipo de ajuste: positivo para sobrantes y negativo para faltantes. |  |  |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio activos. El reporte solo considera centros de costo (CECOs) que estén activos y dados de alta correctamente en el sistema. Los inactivos se excluyen.
- Solo documentos validos Se toman únicamente los ajustes de inventario que estén en estado procesado. Se excluyen los anulados.
- Solo insumos y alimentos/desechables. No se incluyen todos los productos. Solo entran aquellos cuya cuenta contable corresponda a insumos o alimentos y desechables, según la configuración del sistema.
- Signo del monto:
- Si sobra producto, el monto es positivo (ganancia).
- Si falta producto, el monto es negativo (perdida).
- Clasificación por familia de producto. Cada producto se muestra con su jerarquía de familia completa (por ejemplo: Alimentos > Carnes > Res), lo que permite analizar los ajustes por categoría.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.

# Compras por Periodo


## Propósito Funcional


El reporte de Compras Proveedor PAP muestra el detalle de todas las compras realizadas a proveedores por cada punto de servicio. Permite conocer que se compró, a quien, en que cantidad, a qué precio y en qué fecha, documento por documento.


## Tabla


| **Campo Destino** | **Descripción** | **OK/NO OK** |
| --- | --- | --- |
| Profit Center | Código del Profit Center en AX. |  |
| Id | 0 = Indica nombre proveedor y 1 indica el detalle de las compras al proveedor. |  |
| Ceco | Código del Centro de Costo SAP |  |
| RutProveedor | RUT del proveedor que suministro el producto. |  |
| Tipo Documento | Tipo de documento. |  |
| Nro. Documento | En registros de detalle: número del documento de compra. |  |
| Fecha | Fecha de recepción de la mercadería en formato dd/mm/yyyy. |  |
| Periodo | Período contable en formato YYYYMM |  |
| Codigo | Código producto SGP. |  |
| Descripcion | Nombre descriptivo del producto. |  |
| Cta. SAP | Cuenta contable a la que se imputa el producto. |  |
| Precio Unitario | Precio unitario del producto con impuestos incluidos. |  |
| Unidad | Nombre de la unidad de medida del producto. (Und/Kg/Lt) |  |
| Cantidad | Cantidad de unidades recibidas del producto. Solo se incluyen líneas con cantidad recibida mayor a cero. |  |
| Total | Monto total de la línea: precio unitario multiplicado por la cantidad recibida. Para notas de crédito y devoluciones el valor es negativo, representando una reducción en el gasto. |  |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto) El periodo inicia en septiembre y cierra en agosto del siguiente año. Esto responde a la estructura corporativa de Sodexo.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- No filtra por CECOs activos: A diferencia de otros reportes, este procedimiento no genera una lista de CECOs validos previamente. Trae todas las compras que existan en el rango de fechas, independientemente del estado del CECO.
- Solo compras efectivamente recibidas. Únicamente se incluyen compras donde la mercadería fue físicamente recibida (cantidad recibida mayor a cero). Las ordenes pendientes de recepción no aparecen.
- Notas de crédito y devoluciones restan Cuando hay una nota de crédito (NC) o un contradocumento (CE), los montos se registran con signo negativo. Esto significa que el proveedor devolvió dinero o se anuló parcialmente una compra, reduciendo el gasto total.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.
- Tipo de Documento:

| **Sigla** | **Significado** | **Efecto en el monto** | **Descripción** |
| --- | --- | --- | --- |
| **FE** | Factura Electrónica | Positivo (suma) | Documento tributario electrónico emitido por el proveedor. |
| **FA** | Factura | Positivo (suma) | Factura tradicional (no electrónica) emitida por el proveedor. |
| **FP** | Factura Provisionadas | Positivo (suma) | Factura que llega incompleta se crea automáticamente un FP a espera de la NC. |
| **GE** | Guia Electrónica | Positivo (suma) | Documento que registra la recepción física de mercadería en el punto de servicio. Se usa cuando la mercadería llega antes que la factura. |
| **GD** | Guia de Despacho | Positivo (suma) | Documento que acompaña el traslado de mercadería desde el proveedor al punto de servicio. |
| **BO** | Boleta | Positivo (suma) | Comprobante de compra menor, generalmente para adquisiciones al contado o de bajo monto. |
| **SN** | Solicitud Nota de Crédito | Positivo (suma) | Solicitud que aumenta el monto adeudado al proveedor (ajuste de precio al alza, cobro adicional). |
| **ND** | Nota de Debito | Positivo (suma) | Documento que aumenta el monto adeudado al proveedor (ajuste de precio al alza, cobro adicional). |
| **NC** | Nota de Crédito | **Negativo (resta)** | Documento que reduce el monto adeudado al proveedor. Se genera por devoluciones, descuentos o correcciones de facturación. Los montos se multiplican por -1. |
| **CE** | Nota de Crédito Electrónica | **Negativo (resta)** | Anulación o devolución de una entrada de mercadería. Revierte una recepción previa. Los montos se multiplican por -1. |


# Consumo Proyectado Real


## Propósito Funcional


El reporte de Consumo Proyectado y Real compara lo que se planeó consumir de cada producto contra lo que realmente se consumió en la producción de cada punto de servicio. Permite hacer seguimiento producto por producto, día a día, para detectar desviaciones entre la planificación y la ejecución.


Se tiene tres niveles de cantidad comparados:

- **Cantidad teórica:** Lo que dice la receta que se debe usar por ración.
- **Cantidad planificada**: Luego de cambios de los sitios es la cantidad que se va a usar, que puede diferir de la teórica por ajustes operativos realizados en la minuta.
- **Cantidad realizada:** Lo que efectivamente salió de bodega y se usó en la producción. Es el consumo real.

## Tabla


| **Campo Destino** | **Descripción** | **OK/NO OK** |
| --- | --- | --- |
| Profit Center | Código del Profit Center en AX. |  |
| Ceco | Código del Centro de Costo SAP |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |
| Regimen | Código del tipo de régimen alimenticio |  |
| Descripcion Regimen | Descripción del régimen |  |
| Servicio | Código del servicio |  |
| Descripcion Servicio | Descripción del servicio |  |
| Periodo | Período contable en formato YYYYMM |  |
| Fecha | Período consumo DDYYYYMM |  |
| Codigo Producto | Código producto SGP |  |
| Descripcion Producto | Nombre descriptivo del producto. |  |
| Unidad | Nombre de la unidad de medida del producto. (Und/Kg/Lt) |  |
| Cantidad Teorica | Cantidad teórica. |  |
| Cantidad Planificada | Cantidad planificada ajustada. |  |
| Cantidad Realizada | Cantidad efectivamente consumida. |  |
| Cantidad Devolucion | Cantidad devuelta sin usar. |  |
| PMP | Precio medio ponderado. |  |
| Racion_Teorica | Raciones planificadas. |  |
| Usuario_Mod_Racion_Real | Usuario que modifico raciones reales. |  |
| Racion_Real | Raciones efectivamente producidas. |  |
| Fecha_Mod_Racion_Real | Fecha de modificación de raciones reales. |  |
| Usuario_Salida_Produccion | Usuario que registro la salida a producción. |  |
| Racion_Salida_Produccion | Raciones en la salida a producción. |  |
| Fecha_Mod_Salida_Produccion | Fecha de registro de salida a producción. |  |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio reales y activos. Se excluyen los centros de costo que sean de prueba, propuestas comerciales o diseños. Solo se incluyen sitios que están efectivamente operando y dados de alta en el sistema organizacional.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.
- Trazabilidad de modificaciones. El reporte registra quien y cuando hizo cambios críticos.

# Costo Bandeja Planificado


## Propósito Funcional


El reporte de Costo Bandeja Planificado es una herramienta de seguimiento del costo unitario planificado por ración en cada punto de servicio de Sodexo. Muestra cuanto está presupuestado gastar por bandeja en cada tipo de servicio y régimen, junto con la cantidad de comensales que se espera atender. Este costo es mensual.


## Tabla


| **Campo Destino** | **Descripción** | **OK/NO OK** |
| --- | --- | --- |
| Profit Center | Código del Profit Center en AX. |  |
| Org. Compra | Organización de compra asociada al CECO. Define la región del ceco (Por ejemplo: CL14 Región Metropolitana, CL16: Sexta Región). |  |
| Ceco | Código del Centro de Costo |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |
| Regimen | Código del tipo de régimen alimenticio |  |
| Descripcion Regimen | Descripción del régimen |  |
| Servicio | Código del servicio |  |
| Descripcion Servicio | Descripción del servicio |  |
| Periodo | Período contable en formato YYYYMM |  |
| Costo Plato | Costo unitario planificado por bandeja. |  |
| Comensales Totales | Cantidad de comensales planificados |  |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio reales y activos. Se excluyen los centros de costo que sean de prueba, propuestas comerciales o diseños. Solo se incluyen sitios que están efectivamente operando y dados de alta en el sistema organizacional.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.

# Detalle Inventario


## Propósito Funcional


El reporte de Detalle Inventario muestra el stock físico valorizado de cada producto al momento de la toma de inventario en cada punto de servicio de Sodexo. Entrega una fotografía del inventario real a nivel de producto, con su precio unitario, cantidad contada y valor total.


## Tabla


| **Campo Destino** | **Descripción** | **OK/NO OK** |
| --- | --- | --- |
| Profit Center | Código del Profit Center en AX. |  |
| Ceco | Código del Centro de Costo |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |
| Fecha Inventario | Fecha de la toma de inventario DDMMYYYY. |  |
| Periodo | Período contable en formato YYYYMM |  |
| Codigo Producto | Código producto SGP. |  |
| Descripcion Producto | Nombre descriptivo del producto. |  |
| Cta. SAP | Cuenta contable a la que se imputa el producto. |  |
| Precio | Precio unitario del producto al momento de la toma. |  |
| Unidad | Nombre de la unidad de medida del producto. |  |
| Cantidad | Stock fisico contado en la toma de inventario |  |
| Total | Valor total del producto (Precio x Cantidad) |  |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio reales y activos. Se excluyen los centros de costo que sean de prueba, propuestas comerciales o diseños. Solo se incluyen sitios que están efectivamente operando y dados de alta en el sistema organizacional.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.
- El reporte excluye automáticamente los productos que figuran en la toma con cantidad cero. Solo se reportan los artículos que efectivamente tenían existencia al momento del conteo.
- Valoración al precio de la toma.  El precio unitario es el precio de costo del producto vigente al momento de registrar la toma. Este precio puede variar entre periodos según el precio de compra más reciente.

# Jerarquías Waste Watch


## Propósito Funcional


Los reportes de Waste Watch en Power BI utilizan dos archivos externos que se cargan manualmente como fuente de datos estos son mantenidos y actualizados de forma independiente. Su función principal es proveer el detalle organizacional y el estado de los CECOs que los reportes necesitan para filtrar, agrupar y navegar la información jerárquicamente.


**1. BD Sitios**


Contiene la jerarquía organizacional de cada CECO. Se usa en Power BI para habilitar los filtros y agrupaciones por nivel de gestión.


| **Campo** | **Descripción** |
| --- | --- |
| Profit Center | Código del Profit Center en AX. |
| Ceco | Código del Centro de Costo SAP. |
| Descripción Ceco | Nombre del punto de servicio. |
| Planificador | Responsable de planificación del sitio. |
| Segmento | Segmento de negocio (Mining, Corporate Services, Defense, etc.) |
| ADC | Contacto del sitio. |
| N+1 | Primer nivel de jefatura directa |
| N+2 | Segundo nivel de jefatura |
| N+3 | Tercer nivel de jefatura |
| N+4 | Cuarto nivel de jefatura |
| N+5 | Quinto nivel de jefatura |


**2. Estatus PF Base Unificada**


Contiene el estado operacional de cada CECO por mes y año fiscal. Permite identificar que sitios están activos, su clasificación de plataforma y su cadena de responsables.


| **Campo** | **Descripci****ó****n** |
| --- | --- |
| PC | Codigo del Profit Center |
| Ceco | Codigo del Centro de Costo |
| Ceco Descripción | Nombre del punto de servicio |
| Segmento | Segmento de negocio |
| PLATAFORMA/NO PLATAFORMA | Clasificacion del sitio segun si opera bajo plataforma Sodexo |
| STATUS Mes | Estado del sitio en el mes reportado |
| Status | Estado general del sitio. (Abierto, Cerrado, No Aplica) |
| Planificador | Responsable de planificación. |
| ADC | Contacto del sitio. |
| DIRECTOR | Director responsable del sitio. |
| DIRECTOR REGIONAL | Director regional. |
| GERENTE | Gerente del sitio. |
| Org. Compras | Organización de compra asociada. |
| MES | Mes al que corresponde el registro. |
| FY | Ano fiscal (Fiscal Year). |


# Merma Bodega


## Propósito Funcional


El reporte de Merma Bodega muestra el detalle de productos dados de baja en bodega por concepto de merma, indicando el motivo, la cantidad, el precio y el valor total de cada producto mermado. Permite controlar las pérdidas de insumos antes de que lleguen a producción. Puede ser fecha de vencimiento, mal estado producto y mala manipulación.


## Tabla


| **Campo Destino** | **Descripción** | **OK/NO OK** |
| --- | --- | --- |
| Profit Center | Código del Profit Center en AX |  |
| Ceco | Código del Centro de Costo (punto de servicio). |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |
| Tipo Documento | Tipo de documento. Siempre es 'ME' (Merma de Bodega), ya que el reporte filtra exclusivamente por este tipo. |  |
| Numero Documento | Numero correlativo del documento de merma emitido en el sitio. Permite agrupar todos los productos de una misma merma. |  |
| Tipo de Merma | Motivo por el que se dio de baja el producto. Ejemplos observados: fecha de vencimiento, mal estado producto y mala manipulación. |  |
| Fecha Emision | Fecha en que se registró el documento de merma en el sistema, en formato DD-MM-YYYY. |  |
| Periodo | Período contable en formato YYYYMM |  |
| Codigo Producto | Código producto SGP. |  |
| Descripcion Producto | Nombre descriptivo del producto. |  |
| Unidad | Unidad de medida del producto (Kilo, Unidad, Litro). |  |
| Precio | Precio unitario de costo del producto al momento de la merma. |  |
| Cantidad | Cantidad del producto dado de baja por merma, expresada en la unidad del producto. |  |
| Equivalencia | Cantidad mermada convertida a gramos (para Kg/Lt/Und con factor nutricional). Permite comparar mermas de productos con distintas unidades en una misma escala. |  |
| Total | Valor total de la merma del producto, calculado como Precio x Cantidad. |  |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio reales y activos. Se excluyen los centros de costo que sean de prueba, propuestas comerciales o diseños. Solo se incluyen sitios que están efectivamente operando y dados de alta en el sistema organizacional.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.
- Solo documentos de merma vigentes Se excluyen documentos anulados y pendientes. Solo se incluyen mermas confirmadas y procesadas.
- Motivo de merma obligatorio. Cada documento de merma debe tener asociado un tipo de merma (motivo). Los más comunes son fecha de vencimiento y mal estado del producto.
- Numero de documento como agrupador Un mismo documento de merma puede contener múltiples productos. El Numero Documento permite identificar y agrupar todos los productos de un mismo acto de merma.

# Merma Desconche Producción


## Propósito Funcional


El reporte de Merma Desconche producción muestra las pérdidas de insumos que ocurren durante el proceso de producción en cocina y la merma que los comensales dejan en la bandeja, clasificadas en tres tipos: desconche general, desconche de pan y desconche de producción. A diferencia de la Merma Bodega (que registra productos dados de baja antes de producción), este reporte captura lo que se pierde al momento de preparar los alimentos.


## Tabla


| **Campo Destino** | **Descripción** | **OK/NO OK** |
| --- | --- | --- |
| Profit Center | Código del Profit Center en AX |  |
| Ceco | Código del Centro de Costo (punto de servicio). |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |
| Regimen | Código del tipo de régimen alimenticio |  |
| Descripcion Regimen | Descripción del régimen |  |
| Servicio | Código del servicio |  |
| Descripcion Servicio | Descripción del servicio |  |
| Periodo | Período contable en formato YYYYMM |  |
| Fecha Minuta | Fecha en que se registró la merma de producción, en formato DD/MM/YYYY. |  |
| Descripcion Merma | Tipo de merma de producción. Siempre es uno de tres valores fijos. Desconche General, Desconche Pan, Desconche Producción. |  |
| Kilo | Kilos mermados para ese tipo de desconche en la fecha. Puede ser 0 si no hubo merma ese día. |  |
| Costo | Costo unitario por kilo configurado para ese tipo de desconche, según el segmento del CECO y el servicio. Puede ser 0 si no hay costo configurado en b_CostoMermas. |  |
| Costo Total | Valor total de la merma (Kilo x Costo). Es 0 si los Kilos o el Costo son 0. |  |


## Tipos de Desconche


La equivalencia convierte la cantidad mermada a una unidad común (gramos o CC) usando el factor de conversión del producto (tabla pro_facing) y el factor nutricional del ingrediente principal (tabla ing_facnut). La lógica es la siguiente:


| **Descripcion Merma** | **Campo de Kilos** | **Campo de Costo** | **Descripcion** |
| --- | --- | --- | --- |
| Desconche General | Merma_Desconche | Costo_Desconche | Perdida generada en el proceso de preparación general de los alimentos (pelado, corte, limpieza, etc.) |
| Desconche Pan | Merma_Pan | Costo_Pan | Perdida ocurrida por el comensal que deja en la bandeja, específicamente del pan. |
| Desconche Produccion | Merma_Produccion | Costo_Produccion | Perdida ocurrida por el comensal que deja en la bandeja. |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio reales y activos. Se excluyen los centros de costo que sean de prueba, propuestas comerciales o diseños. Solo se incluyen sitios que están efectivamente operando y dados de alta en el sistema organizacional.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.
- Tres tipos de desconche siempre presentes. Por cada día registrado aparecen tres filas (una por tipo). Si un tipo no tuvo merma, aparece con Kilo = 0. Si no hay tarifa configurada, aparece con Costo = 0 y Costo Total = 0.
- Costo configurado por tarifa, no calculado en tiempo real. El costo por kg no se obtiene del precio del insumo sino de una tabla de tarifas (b_CostoMermas) preconfigurada por segmento y servicio. Esto implica que el costo puede estar desactualizado si la tarifa no se revisa periódicamente.

# Raciones No Vendidas (Merma Línea)


## Propósito Funcional


El reporte de Raciones No Vendidas (también llamado Merma de Línea) muestra las raciones que fueron producidas, pero no consumidas ni vendidas, cuantificando el costo de esa pérdida por receta, CECO, régimen y servicio. Representa el desperdicio que ocurre al final de la línea de servicio: la comida que se preparó, pero no llego a ninguna bandeja.


## Tabla


| **Campo Destino** | **Descripción** | **OK/NO OK** |
| --- | --- | --- |
| Profit Center | Código del Profit Center en AX |  |
| Ceco | Código del Centro de Costo (punto de servicio). |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |
| Regimen | Código del tipo de régimen alimenticio |  |
| Descripcion Regimen | Descripción del régimen |  |
| Servicio | Código del servicio |  |
| Descripcion Servicio | Descripción del servicio |  |
| Fecha Minuta | Fecha en que se registró la merma de producción, en formato DD/MM/YYYY. |  |
| Periodo | Período contable en formato YYYYMM |  |
| Receta | Código de Receta |  |
| Descripción Receta | Nombre de la receta |  |
| Programado | Raciones producidas/programadas |  |
| Costo | Costo unitario de la receta por ración, considerando solo el componente de alimentos (CostoRecetaAlimento) |  |
| Total Costo | Costo total de las raciones programadas (Programado x Costo). Representa el costo de producción de esa receta ese día |  |
| Merma | Cantidad de raciones que no fueron vendidas ni consumidas (CantidadMerma). Es la merma de línea propiamente tal. |  |
| Total Merma | Costo total de las raciones mermadas, incluyendo alimentos y desechables: CantidadMerma * (CostoRecetaAlimento + CostoRecetaDesechable) |  |
| Gramos Brutos | Equivalencia en gramos brutos de las raciones mermadas |  |
| Cantidad Servida | Porciones servidas equivalentes de las raciones mermadas. |  |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio reales y activos. Se excluyen los centros de costo que sean de prueba, propuestas comerciales o diseños. Solo se incluyen sitios que están efectivamente operando y dados de alta en el sistema organizacional.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.
- El costo de la merma incluye alimentos y desechables El Total Merma considera tanto el costo del alimento (CostoRecetaAlimento) como el costo de los desechables (CostoRecetaDesechable) asociados a cada ración mermada, ya que ambos se prepararon y no se utilizaron.
- El Total Costo solo incluye alimentos El campo Total Costo (costo de las raciones programadas) usa solo CostoRecetaAlimento, sin incluir desechables.

# Reporte Food Cost


## Propósito Funcional


El reporte Food Cost mide la eficiencia en el uso de insumos de cada punto de servicio. Responde a la pregunta fundamental de la operación de food service: ¿de cada peso que se vende, cuanto se destina a cubrir el costo de los ingredientes y materiales?


## Tabla


| **Campo Destino** | **Descripción** | **OK/NO OK** |
| --- | --- | --- |
| Profit Center | Código del Profit Center en AX |  |
| Ceco | Código del Centro de Costo (punto de servicio). |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |
| Regimen | Código del tipo de régimen alimenticio |  |
| Descripcion Regimen | Descripción del régimen |  |
| Servicio | Código del servicio |  |
| Descripcion Servicio | Descripción del servicio |  |
| Fecha | Fecha del registro en formato YYYYMMDD |  |
| Periodo | Período contable en formato YYYYMM |  |
| Servicio Raciones Vendida | Cantidad de raciones vendidas al cliente en el día. Es 0 para registros de Ventas Especiales |  |
| Venta Dia | Monto total de ventas del día (venta crédito + venta contada) |  |
| Valor Bandeja | Precio promedio cobrado por ración vendida: (Venta_Dia + Venta_Contado) / Raciones_Vendidas. Es 0 si no hay venta o raciones |  |
| Raciones Producidas | Cantidad de raciones efectivamente preparadas en cocina |  |
| Costo Dia | Costo total de insumos del día. Para Op=1 usa Costo_Realizado_Alim; para Op=2 usa Costo_Realizado_Desec. |  |
| Costo Bandeja | Costo promedio por racion producida: Costo_Dia / Raciones_Producidas. Es 0 si no hay costo o raciones. |  |
| Costo Bandeja Vendido | Costo promedio por ración vendida: Costo_Dia / Raciones_Vendidas. Permite evaluar el costo por ración entregada al cliente. |  |
| Food Cost | Porcentaje del costo sobre la venta: (Costo_Dia / Venta_Dia) * 100. Indicador clave de eficiencia: de cada $100 vendidos, cuanto se destinó a insumos. |  |
| Orden |  |  |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio reales y activos. Se excluyen los centros de costo que sean de prueba, propuestas comerciales o diseños. Solo se incluyen sitios que están efectivamente operando y dados de alta en el sistema organizacional.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.
- Ventas especiales separadas. Las ventas de servicios especiales (eventos, catering extraordinario) se reportan aparte del servicio regular, para no mezclar los indicadores de la operación diaria.

# Reporte Q Detallado


## Propósito Funcional


El reporte Q Detallado muestra el desglose diario de raciones por punto de servicio, comparando tres tipos de raciones: las planificadas en la minuta, las producidas en cocina y las vendidas por cliente. Permite identificar diferencias entre lo que se planifico, lo que se produjo y lo que finalmente se vendió cada día.


## Tabla


| **Campo Destino** | **Descripción** | **OK/NO OK** |
| --- | --- | --- |
| Profit Center | Código del Profit Center en AX. |  |
| Ceco | Código del Centro de Costo |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |
| Regimen | Código del tipo de régimen alimenticio |  |
| Descripcion Regimen | Descripción del régimen |  |
| Servicio | Código del servicio |  |
| Descripcion Servicio | Descripción del servicio |  |
| Fecha | Fecha de la minuta en formato DD/MM/YYYY. |  |
| Periodo | Período contable en formato YYYYMM |  |
| Descripcion Q | Tipo de racion: Planificada, Producida o Vendida. |  |
| Codigo Cliente | RUT o codigo del cliente (solo para Vendida). |  |
| Cliente | Nombre del cliente (solo para Vendida). |  |
| Raciones | Cantidad de raciones. |  |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre. Después de agosto se otorgan días adicionales para completar registros antes del cierre definitivo.
- Solo puntos de servicio reales y activos. Se excluyen los centros de costo que sean de prueba, propuestas comerciales o diseños. Solo se incluyen sitios que están efectivamente operando y dados de alta en el sistema organizacional.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.
- El reporte excluye los siguientes códigos de servicio: 10940, 10897, 11056 y 11057.
- Tres tipos de raciones por día Cada día puede mostrar hasta tres tipos de filas por CECO/régimen/servicio:
- Planificada: Lo que se programó producir según la minuta.
- Producida: Lo que efectivamente salió de la cocina.
- Vendida: Lo que se entregó a cada cliente o consumidor.
- Clientes internos con códigos textuales. Algunos consumidores no son clientes externos con RUT, sino categorías internas como PERSONAL (consumo del personal del sitio) o MUESTRA R (raciones de prueba o muestra). Estos aparecen sin nombre en el campo Cliente.

# Reporte Q Resumido


## Propósito Funcional


El reporte de Raciones Q Resumido muestra el resumen mensual de raciones por punto de servicio, comparando tres totales: las raciones planificadas en la minuta, las producidas en cocina y las vendidas a clientes. A diferencia del Q Detallado (que muestra cada día con detalle por cliente), este reporte consolida los totales del mes en una sola fila por CECO, régimen y servicio.


## Tabla


| **Campo Destino** | **Descripción** | **OK/NO OK** |
| --- | --- | --- |
| Profit Center | Código del Profit Center en AX. |  |
| Ceco | Código del Centro de Costo |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |
| Regimen | Código del tipo de régimen alimenticio |  |
| Descripcion Regimen | Descripción del régimen |  |
| Servicio | Código del servicio |  |
| Descripcion Servicio | Descripción del servicio |  |
| Fecha Inicio | Primer día del mes procesado |  |
| Fecha Fin | Ultimo día con datos del mes |  |
| Periodo | Periodo contable en formato YYYYMM |  |
| Planificadas | Total raciones planificadas en la minuta del mes |  |
| Producidas | Total raciones producidas en el me |  |
| Vendidas | Total raciones vendidas a clientes en el mes |  |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio reales y activos. Se excluyen los centros de costo que sean de prueba, propuestas comerciales o diseños. Solo se incluyen sitios que están efectivamente operando y dados de alta en el sistema organizacional.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.
- Limite al ultimo cierre del sitio La Fecha Fin del mes queda limitada al último día cerrado en cas_log_envio. Esto garantiza que los datos sean definitivos y reflejen solo periodos ya cerrados. Si un CECO no tiene cierre, no aparece en el reporte.
- Exclusion de servicios especiales Los servicios de Cafeteria y Eventos Especiales se excluyen del reporte porque tienen una lógica de raciones diferente al servicio regular.
- Tres tipos de raciones por día Cada día puede mostrar hasta tres tipos de filas por CECO/régimen/servicio:
- Planificada: Lo que se programó producir según la minuta.
- Producida: Lo que efectivamente salió de la cocina.
- Vendida: Lo que se entregó a cada cliente o consumidor.

# Reporte WW Global


La carpeta debe estar disponible para que el proveedor externo suba información.


# Respaldo2021- 2022


Se guarda el archivo de A13, Comensales Minuta Bloque y Compras por Periodo (Compras Proveedores). se mantienen 5 años móviles


# Traspasos desde la CD


## Propósito Funcional


El reporte de Traspasos CD muestra el detalle de productos traspasados desde el Centro de Distribucion (CD) hacia los puntos de servicio (CECOs), indicando el número de guía de despacho, el producto, el precio, la cantidad y el valor total de cada movimiento. Permite controlar el flujo de insumos que llegan desde el centro de distribución a cada casino.


## Tabla


| **Campo Destino** | **Descripción** | **OK/NO OK** |
| --- | --- | --- |
| Profit Center | Código del Profit Center en AX. |  |
| Ceco | Código del Centro de Costo |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |
| Origen | Origen del traspaso. Siempre es 'CD' (Centro de Distribucion). |  |
| Nro. Guia Despacho |  |  |
| Codigo Producto | Código producto SGP. |  |
| Nombre Producto | Nombre descriptivo del producto. |  |
| Precio | Precio unitario de costo del producto al momento del traspaso. |  |
| Unidad | Unidad de medida del producto (Kilo, Unidad, Litro, etc.). |  |
| Cantidad | Cantidad del producto traspasada desde el CD al CECO, expresada en la unidad del producto. |  |
| Fecha Movimiento | Fecha en que se realizó el traspaso, en formato DD/MM/YYYY. |  |
| Periodo | Periodo contable en formato YYYYMM derivado de la Fecha Movimiento. |  |
| Total | Valor económico total del producto traspasado: Precio x Cantidad,. |  |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio reales y activos. Se excluyen los centros de costo que sean de prueba, propuestas comerciales o diseños. Solo se incluyen sitios que están efectivamente operando y dados de alta en el sistema organizacional.
- Se ejecuta de forma diaria (acumulando datos del periodo fiscal en curso) y también de forma anual (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles. El proceso se ejecuta a 3:00 AM para la extracción de datos y luego a las 7:00 AM se dejan los archivos en el sharepoint.
- Solo traspasos vigentes desde el CD Se incluyen exclusivamente documentos de tipo 'TR' (Traspaso), emitidos por el Centro de Distribución (tov_codser = 1), que no estén anulados ni pendientes.
- Solo productos contabilizables. Se filtran únicamente los productos cuya cuenta contable corresponde a insumos o alimentos/desechables (parametros ctainsumo y ctalimdes). Esto excluye productos administrativos o de otro tipo que no impactan el food cost.
- Solo movimientos que afectan inventario El campo dev_mueinv = 'S' asegura que solo se incluyan los productos que generan movimiento en el inventario del CECO receptor.
- Origen siempre CD El campo Origen siempre muestra 'CD' porque el filtro tov_codser = 1 restringe los resultados al Centro de Distribución. No hay traspasos de otro origen en este reporte.
- Numero de guía como agrupador. Un mismo documento de traspaso (guía de despacho) puede contener múltiples productos. El Nro. Guia Despacho permite identificar y agrupar todos los productos de un mismo envío.

# Último Cierre


## Propósito Funcional


El reporte de Ultimo Cierre muestra el estado de cierre de cada punto de servicio por mes, indicando la fecha del último envío de datos registrado, cuantos días han transcurrido desde ese cierre y cuantos días de atraso acumula el sitio. Es el indicador de cumplimiento operacional: permite saber que sitios están al día con el registro de su información y cuales están atrasados.


## Tabla


| **Campo Destino** | **Descripción** | **OK/NO OK** |
| --- | --- | --- |
| Profit Center | Código del Profit Center en AX. |  |
| Ceco | Código del Centro de Costo |  |
| Descripcion Ceco | Descripción del Centro de Costo |  |
| UltimoCierre | Fecha del último cierre registrado por el CECO en ese mes. |  |
| Estado al | Fecha de referencia del mes consultado. Para meses pasados es el último día del mes a las 23:59. Para el mes actual es la fecha del día de ejecución |  |
| Periodo | Periodo contable en formato YYYYMM |  |
| Diferencia | Días transcurridos entre el UltimoCierre y el Estado al. Un valor de 0 significa que el CECO cerro en la fecha de referencia. Un valor mayor indica días de rezago. |  |
| Dias Inhabiles | Días no hábiles (festivos u otros días configurados) para ese CECO entre el UltimoCierre y el Estado al, según la tabla cas_b_fecha_inhabiles. |  |
| Dias Atraso | Días efectivos de atraso del CECO en el mes, calculado como la suma de (diferencia - días inhábiles) por cada cierre del mes. El mínimo es 0 (nunca negativo). |  |
| Promedio Dias de Atraso | Promedio de días de atraso por cierre en el mes: SUM(Días Atraso) / COUNT(cierres del mes). Permite evaluar el comportamiento histórico de cumplimiento del sitio. |  |


## Reglas del Negocio

- Año fiscal francés (septiembre - agosto). El reporte sigue el calendario corporativo de Sodexo, que inicia en septiembre y cierra en agosto del siguiente año. El reporte muestra desde septiembre al mes actual.
- Días de gracia para el cierre anual. Después de terminar agosto, se otorga un margen de 5 días adicionales para que los sitios completen el registro de sus movimientos antes de consolidar el cierre del periodo.
- Solo puntos de servicio reales y activos. Se excluyen los centros de costo que sean de prueba, propuestas comerciales o diseños. Solo se incluyen sitios que están efectivamente operando y dados de alta en el sistema organizacional.
- Se ejecuta de forma **diaria** (acumulando datos del periodo fiscal en curso) y también de forma **anual** (para el cierre del año fiscal en agosto). Cuanto termina el periodo, el archivo se guarda en otra carpeta “Histórico”, se mantienen 5 años móviles.


