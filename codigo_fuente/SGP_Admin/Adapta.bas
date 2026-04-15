Attribute VB_Name = "Adapta"
Function ActVersion()
Dim nVer As Long, aVer As Long
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset

nVer = CLng(App.Major & App.Minor & App.Revision)
aVer = TipoDato(GetParametro("version"), 0)
If nVer > aVer And aVer = 0 Then
    vg_db.Execute "insert into a_param values ('version', 'Versión del Sistema', 'N', '101')"
    aVer = 101
End If
If nVer > aVer And aVer = 101 Then
    vg_db.Execute "alter table b_totventas add column tov_numinf long"
    vg_db.Execute "update b_totventas set tov_numinf=0"
    vg_db.Execute "update a_param set par_valor='102' where par_codigo='version'"
    aVer = 102
End If
If nVer > aVer And aVer = 102 Then
    vg_db.Execute "insert into a_opcsistema values (2090000, 'Control Traspasos entre Casinos')"
    vg_db.Execute "insert into a_infcfcfofi values ('T', 1, 0, Null)"
    vg_db.Execute "update a_param set par_valor='103' where par_codigo='version'"
    aVer = 103
End If
If nVer > aVer And aVer = 103 Then
    vg_db.Execute "drop table b_minutaraciones"
    vg_db.Execute "create table b_minutaraciones (mir_cencos char(10), mir_codreg int, mir_codser int, mir_fecmin int, mir_rutcli char(10), mir_nrorac int, Constraint b_minutaraciones_pk Primary Key (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli))"
    vg_db.Execute "update a_param set par_valor='104' where par_codigo='version'"
    aVer = 104
End If
If nVer > aVer And aVer = 104 Then
    vg_db.Execute "create table b_preciovta (prv_cencos char(10), prv_codreg int, prv_codser int, prv_fecvig int, prv_rutcli char(10), prv_preven double, Constraint b_preciovta_pk Primary Key (prv_cencos, prv_codreg, prv_codser, prv_fecvig, prv_rutcli))"
    vg_db.Execute "insert into a_tipoajuste values (3, 'Inventario inicial', 1, 'A')"
    vg_db.Execute "update a_param set par_valor='105' where par_codigo='version'"
    aVer = 105
End If
If nVer > aVer And aVer = 105 Then
    vg_db.Execute "insert into a_opcsistema values (2075000, 'Precio de Venta Cliente')"
    vg_db.Execute "update a_param set par_valor='106' where par_codigo='version'"
    aVer = 106
End If
If nVer > aVer And aVer = 106 Then
    vg_db.Execute "insert into a_opcsistema values (3080000, 'Informe de Mermas por Período')"
    vg_db.Execute "insert into a_opcsistema values (3090000, 'Informe de Ventas Directas por Período')"
    vg_db.Execute "update a_param set par_valor='107' where par_codigo='version'"
    aVer = 107
End If
If nVer > aVer And aVer = 107 Then
    vg_db.Execute "update a_param set par_valor='108' where par_codigo='version'"
    aVer = 108
End If
If nVer > aVer And aVer = 108 Then
    vg_db.Execute "alter table b_totcompras add column toc_ordcom char(10)"
    vg_db.Execute "update b_totcompras set toc_ordcom=''"
    vg_db.Execute "update a_param set par_valor='109' where par_codigo='version'"
    aVer = 109
End If
If nVer > aVer And aVer = 109 Then
    vg_db.Execute "insert into a_opcsistema values (2100000, 'Control Fondo Fijo (FOFI)')"
    vg_db.Execute "insert into a_opcsistema values (2110000, 'Resultado Operacional Mensual (A13)')"
    vg_db.Execute "alter table b_tomainv add column tin_ciemes long"
    vg_db.Execute "update b_tomainv set tin_ciemes=0"
    vg_db.Execute "insert into a_param values ('diasstock', 'Dias de Stock (A13)', 'N', '30')"
    vg_db.Execute "insert into a_param values ('ctamovil', 'Movilizacion', 'C', '410005')"
    vg_db.Execute "update a_param set par_valor='110' where par_codigo='version'"
    aVer = 110
End If
If nVer > aVer And aVer = 110 Then
    vg_db.Execute "alter table b_productos add column pro_ctrsto int"
    vg_db.Execute "update b_productos set pro_ctrsto=iif(pro_ctacon='410001' or pro_ctacon='410004',1,0)"
    vg_db.Execute "update a_param set par_valor='111' where par_codigo='version'"
    aVer = 111
End If
If nVer > aVer And aVer = 111 Then
    vg_db.Execute "alter table a_unidad add column uni_codunm int, uni_valuni double"
    vg_db.Execute "update a_unidad set uni_codunm=1 where uni_codigo=1"
    vg_db.Execute "update a_unidad set uni_codunm=2 where uni_codigo=2"
    vg_db.Execute "update a_unidad set uni_codunm=3 where uni_codigo=3"
    vg_db.Execute "create table b_productocompra (pco_codpro char(20), pco_codigo char(20), pco_nombre char(100), pco_undemb double, pco_fecven int, Constraint b_produtocompra_pk Primary Key (pco_codpro, pco_codigo))"
    vg_db.Execute "update a_param set par_valor='112' where par_codigo='version'"
    aVer = 112
End If
If nVer > aVer And aVer = 112 Then
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (2020000, 'Consumo Ingrediente')"
   vg_db.Execute "INSERT INTO a_param VALUES ('ctagastos2', 'Cuentas de Gastos Generales', 'C', ' ')"
   vg_db.Execute "ALTER TABLE b_clientes ADD cli_nomcontable NVARCHAR(50), cli_emailcontable NVARCHAR(50)"
   vg_db.Execute "ALTER TABLE a_impuesto ADD imp_codsap NVARCHAR(20)"
   vg_db.Execute "UPDATE a_impuesto SET imp_codsap='123060' WHERE imp_codigo=1"
   vg_db.Execute "UPDATE a_impuesto SET imp_codsap='123070' WHERE imp_codigo=3"
   vg_db.Execute "update a_param set par_valor='113' where par_codigo='version'"
   aVer = 113
End If
If nVer > aVer And aVer = 113 Then
    vg_db.Execute "update a_param set par_valor='114' where par_codigo='version'"
    aVer = 114
End If
If nVer > aVer And aVer = 114 Then
   vg_db.Execute "DROP TABLE b_minutacasino"
   vg_db.Execute "CREATE TABLE b_minutacasino (mic_cencos char(10), mic_codreg int, mic_codser int, mic_fecmin int, mic_fecenv int, Constraint b_minutacasino_pk Primary Key (mic_cencos, mic_codreg, mic_codser, mic_fecmin))"
   vg_db.Execute "ALTER TABLE b_receta ADD rec_fecvig int"
   vg_db.Execute "UPDATE b_receta set rec_fecvig=0"
   vg_db.Execute "update a_param set par_valor='115' where par_codigo='version'"
   aVer = 115
End If
If nVer > aVer And aVer = 115 Then
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1080000, 'Parametrizar Nº Recetas 5 Etapas')"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1090000, 'Parametrizar Costo Patrón 5 Etapas')"
   vg_db.Execute "update a_param set par_valor='116' where par_codigo='version'"
   aVer = 116
End If
If nVer > aVer And aVer = 116 Then
   'Crear campo % de precio en tabla param, para validar precio
    vg_db.Execute "INSERT INTO a_param VALUES ('porprepro', 'validación % precio producto', 'N', '20')"
   'Actualizar campo ind modificado
   vg_db.Execute "UPDATE a_impuesto SET imp_indmod='N'"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1100000, 'Gramo Familia Producto 5 Etapas')"
   vg_db.Execute "update a_param set par_valor='117' where par_codigo='version'"
   aVer = 117
End If
If nVer > aVer And aVer = 117 Then
    vg_db.Execute "INSERT INTO a_opcsistema VALUES (1092000, 'Parametrizar Costo Patrón Techo 5 Etapas')"
    vg_db.Execute "UPDATE a_opcsistema SET opc_nombre='Parametrizar Costo Patrón Piso 5 Etapas' WHERE opc_codigo=1090000"
    vg_db.Execute "ALTER TABLE a_estservicio ADD ess_racmin FLOAT"
    vg_db.Execute "UPDATE a_estservicio SET ess_racmin=0"
    vg_db.Execute "update a_param set par_valor='118' where par_codigo='version'"
    aVer = 118
End If
If nVer > aVer And aVer = 118 Then
    vg_db.Execute "update a_param set par_valor='119' where par_codigo='version'"
    aVer = 119
End If
If nVer > aVer And aVer = 119 Then
   'Insertar campo grupo vulnerable a tabla b_receta
   vg_db.Execute "ALTER TABLE b_receta ADD rec_gruvul NTEXT"
   'Insertar campo grupo vulnerable a tabla cliente, para ser enviado el concepto mencionado
   vg_db.Execute "ALTER TABLE b_clientes ADD cli_gruvul nvarchar(1)"
   vg_db.Execute "UPDATE b_clientes SET cli_gruvul='N'"
   vg_db.Execute "UPDATE a_param SET par_valor='120' WHERE par_codigo='version'"
   aVer = 120
End If
If nVer > aVer And aVer = 120 Then
    vg_db.Execute "UPDATE a_param SET par_valor='121' WHERE par_codigo='version'"
    aVer = 121
End If
If nVer > aVer And aVer = 121 Then
   'Insertar datos a tabla zona y contratos
'    vg_db.Execute "INSERT INTO a_zona VALUES (1, 'Region Metropolitana')"
'    vg_db.Execute "UPDATE b_clientes SET cli_codzon=1"
   'Modificar opción de sistema y incluir nueva opción
    vg_db.Execute "UPDATE a_opcsistema SET opc_nombre='Zona' WHERE opc_codigo=4150000"
    vg_db.Execute "INSERT INTO a_opcsistema VALUES (4160000, 'Parámetros Generales')"
    'Actualizar versión
    vg_db.Execute "UPDATE a_param SET par_valor='122' WHERE par_codigo='version'"
    aVer = 122
End If
If nVer > aVer And aVer = 122 Then
   '------- Insertar campo activa modulo pacientes
   vg_db.Execute "ALTER TABLE b_clientes ADD cli_modpac nvarchar(1)"
   vg_db.Execute "UPDATE b_clientes SET cli_modpac='N'"
   '------- Isertar campo maestro productos
   vg_db.Execute "ALTER TABLE b_productos ADD pro_maepro INT"
   vg_db.Execute "UPDATE b_productos SET pro_maepro=1"
   vg_db.Execute "UPDATE b_productos SET pro_maepro=0 WHERE pro_ctacon NOT IN ('410001','410004')"
   
   '------- Insertar campo a la tabla cliente tipo servicio y segmento
   vg_db.Execute "ALTER TABLE b_clientes ADD cli_codtis int, cli_codseg int"
   vg_db.Execute "UPDATE b_clientes SET cli_codtis=1, cli_codseg=0"
   '------- Incluir opción tipo de servicio
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4170000, 'Tipo de Servicio')"
   '------- Incluir opción segmento
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4180000, 'Segmento')"
   vg_db.Execute "UPDATE a_param SET par_valor='123' WHERE par_codigo='version'"
   aVer = 123
End If
If nVer > aVer And aVer = 123 Then
   '-------> Actualizar opción sistema parametros contrato x parametro web service
   vg_db.Execute "UPDATE a_opcsistema SET opc_nombre='Parámetro Web Service' WHERE opc_codigo=4160000"
   '------- Incluir campo tabla b_minutadet costo desechable
   vg_db.Execute "ALTER TABLE b_minutadet ADD mid_cosdes float"
   vg_db.Execute "UPDATE b_minutadet SET mid_cosdes=0"
   '------- Insertar campo código sap y facturable tabla servicio
   vg_db.Execute "ALTER TABLE a_servicio ADD ser_codsap varchar(20), ser_facturable varchar(1)"
   '------- Insertar campo impuesto adicional
   vg_db.Execute "ALTER TABLE a_impuesto ADD imp_adicional int"
   vg_db.Execute "UPDATE a_impuesto SET imp_adicional='0' WHERE imp_codigo IN (1,11)"
   vg_db.Execute "UPDATE a_impuesto SET imp_adicional='1' WHERE imp_codigo NOT IN (1,11)"
   '------- Insertar campo sociedad sap
   vg_db.Execute "ALTER TABLE b_clientes ADD cli_socsap varchar(4)"
   vg_db.Execute "UPDATE b_clientes SET cli_socsap='SDXO'"
   '------- Insertar campo activo contrato
   vg_db.Execute "ALTER TABLE b_clientes ADD cli_activo varchar(1)"
   vg_db.Execute "UPDATE b_clientes SET cli_activo='1'"
   '------- Insertar campo activo envio SAP
   vg_db.Execute "ALTER TABLE b_clientes ADD cli_envsap varchar(1)"
   vg_db.Execute "UPDATE b_clientes SET cli_envsap='0'"
   '------- Incluir opción tipo documento
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4062000, 'Tipo Documento')"
   vg_db.Execute "UPDATE a_param SET par_valor='124' WHERE par_codigo='version'"
   aVer = 124
End If
If nVer > aVer And aVer = 124 Then
    vg_db.Execute "UPDATE a_param SET par_valor='125' WHERE par_codigo='version'"
    aVer = 125
End If
If nVer > aVer And aVer = 125 Then
   '-------> Incluir opción Lista precio
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1015000, 'Lista de Precio')"
   '-------> Incluir opción importa desde excel Lista precio
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1016000, 'Importar Lista de Precio Desde Excel')"
   '-------> Incluir opción Asociar lista precio
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1017000, 'Asociar Lista de Precio')"
   '-------> Incluir opción Actualizar lista de precio planificaión
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1018000, 'Actualizar Lista de Precio Planificación')"
   '-------> Incluir opción Grupo de cambio ingrediente en receta SGP
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4011000, 'Grupo de Cambio Ingrediente en Receta SGP')"
   '-------> Incluir opción Habilitar cambio ingrediente en receta SGP
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4012000, 'Habilitar Cambio Ingrediente en Receta SGP')"
   '------- Insertar campo codigo producto a la tabla ingrediente, para relacionar el costo del producto
   vg_db.Execute "ALTER TABLE b_ingrediente ADD ing_codpro varchar(20)"
   '------- Insertar campo codigo lista de precio a la tabla encabezado minuta
   vg_db.Execute "ALTER TABLE b_minuta ADD min_codlpr int"
   '------- Insertar campo activo tabla regimen
   vg_db.Execute "ALTER TABLE a_regimen ADD reg_activo varchar(1)"
   vg_db.Execute "UPDATE a_regimen SET reg_activo='1'"
   '------- Insertar campo lista precio tabla zona
   vg_db.Execute "ALTER TABLE a_zona ADD zon_codlpr int"
   '-------> Incluir opción Parametro de Recetas
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4190000, 'Parametros Recetas')"
   '-------> Incluir parametro lista precio en recetas
   vg_db.Execute "INSERT INTO a_param VALUES ('parlprrec', 'Parametro Lista Precio Recetas', 'N', '0')"
   '-------> Incluir opción Parametro de costo planificación minutas
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (2012000, 'Costo Minutas')"
   
   vg_db.Execute "UPDATE a_param SET par_valor='126' WHERE par_codigo='version'"
   aVer = 126
End If
If nVer > aVer And aVer = 126 Then
   RS1.Open "SELECT * FROM b_clientes WHERE cli_envsap='1'", vg_db, adOpenForwardOnly
   Do While Not RS1.EOF
      vg_db.Execute "INSERT INTO b_casinointerfaz VALUES ('" & RS1!cli_codigo & "', 1)"
      RS1.MoveNext
   Loop
   RS1.Close: Set RS1 = Nothing
   vg_db.Execute "ALTER TABLE b_clientes DROP COLUMN cli_envsap"
   '-------> Actualizar opciones de sistemas
   '-------> Incluir opción
   
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4000000 WHERE dpe_codopc=4010000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4230000 WHERE dpe_codopc=4810000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4220000 WHERE dpe_codopc=4800000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4210000 WHERE dpe_codopc=4190000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4200000 WHERE dpe_codopc=4160000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4180000 WHERE dpe_codopc=4180000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4170000 WHERE dpe_codopc=4170000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4160000 WHERE dpe_codopc=4150000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4150000 WHERE dpe_codopc=4140000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4140000 WHERE dpe_codopc=4100000"
'   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4130000 WHERE dpe_codopc=4120000"
'   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4120000 WHERE dpe_codopc=4110000"
'   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4110000 WHERE dpe_codopc=4080000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4100000 WHERE dpe_codopc=4070000"
'   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4090000 WHERE dpe_codopc=4062000"
'   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4080000 WHERE dpe_codopc=4060000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4070000 WHERE dpe_codopc=4050000"
'   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4060000 WHERE dpe_codopc=4040000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4050000 WHERE dpe_codopc=4030000"
'   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4040000 WHERE dpe_codopc=4020000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4030000 WHERE dpe_codopc=4015000"
'   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4020000 WHERE dpe_codopc=4011000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=4010000 WHERE dpe_codopc=4012000"
   
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4000000 WHERE opc_codigo=4010000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4230000 WHERE opc_codigo=4810000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4220000 WHERE opc_codigo=4800000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4210000 WHERE opc_codigo=4190000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4200000 WHERE opc_codigo=4160000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4180000 WHERE opc_codigo=4180000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4170000 WHERE opc_codigo=4170000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4160000 WHERE opc_codigo=4150000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4150000 WHERE opc_codigo=4140000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4140000 WHERE opc_codigo=4100000"
'   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4130000 WHERE opc_codigo=4120000"
'   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4120000 WHERE opc_codigo=4110000"
'   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4110000 WHERE opc_codigo=4080000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4100000 WHERE opc_codigo=4070000"
'   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4090000 WHERE opc_codigo=4062000"
'   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4080000 WHERE opc_codigo=4060000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4070000 WHERE opc_codigo=4050000"
'   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4060000 WHERE opc_codigo=4040000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4050000 WHERE opc_codigo=4030000"
'   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4040000 WHERE opc_codigo=4020000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4030000 WHERE opc_codigo=4015000"
'   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4020000 WHERE opc_codigo=4011000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=4010000 WHERE opc_codigo=4012000"
   
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4190000, 'Tipo Envío SAP')"
'   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4130000, 'Sub-Segmento')"
'   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4120000, 'Casino')"
'   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4140000, 'Servicio')"
'   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4100000, 'Tipo de Plato')"
'   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4090000, 'Tipo de Documento')"
'   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4080000, 'Cuenta Contable')"
'   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4060000, 'Nutrientes')"
'   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4040000, 'Unidad de Stock')"
'   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4020000, 'Habilitar Cambio Ingrediente en Receta SGP')"
   '------- Insertar campo codigo producto tabla producto ingrediente, para relacionar el costo del producto
   vg_db.Execute "ALTER TABLE b_productosing ADD pri_propre int"
   vg_db.Execute "UPDATE b_productosing SET pri_propre=0"
   RS1.Open "SELECT ing_codpro, ing_codigo FROM b_ingrediente WHERE ing_codpro <> '' OR (ing_codpro) IS NOT NULL ", vg_db, adOpenForwardOnly
   Do While Not RS1.EOF
'      DoEvents
      vg_db.Execute "UPDATE b_productosing SET pri_propre = 1 WHERE pri_coding ='" & RS1!ing_codigo & "' AND pri_codpro = '" & RS1!ing_codpro & "'"
      RS1.MoveNext
   Loop
   RS1.Close: Set RS1 = Nothing
   '------- Eliminar campo codigo producto a la tabla ingrediente
   vg_db.Execute "ALTER TABLE b_ingrediente DROP COLUMN ing_codpro"
   
   vg_db.Execute "UPDATE a_param SET par_valor='127' WHERE par_codigo='version'"
   aVer = 127
End If
If nVer > aVer And aVer = 127 Then
   '------- Insertar campo activo y ultima fecha modificación tabla proveedor
   vg_db.Execute "ALTER TABLE b_proveedor ADD prv_activo varchar(1), prv_fecumo datetime, prv_origen varchar(1)"
   vg_db.Execute "UPDATE b_proveedor SET prv_activo='0', prv_origen='1'"
   
   '-------> Incluir opción Aosciar productos SAC & SGP - Proveedores
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1011000, 'Asociar Productos SAC vs SGP', null, null)"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1012000, 'Proveedores', null, null)"
   vg_db.Execute "UPDATE a_param SET par_valor='128' WHERE par_codigo='version'"
   aVer = 128
End If
If nVer > aVer And aVer = 128 Then
   Dim CodOpc As Long
   '-------> Incluir opción de envio listas de precios sac
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1032000, 'Generación Archivos Planos Lista Precios', null, null)"
   '-------> Insertar campo que indique estado de proceso
   vg_db.Execute "ALTER TABLE b_minuta ADD min_estact varchar(1)"
   vg_db.Execute "UPDATE b_minuta SET min_estact='0'"
   '-------> Insertar campo a la tabla lista precio
   vg_db.Execute "ALTER TABLE b_listaprecio ADD lpr_codcec varchar(4), lpr_codcco varchar(10), lpr_activo varchar(1)"
   vg_db.Execute "UPDATE b_listaprecio SET lpr_activo='1'"
   vg_db.Execute "ALTER TABLE b_detlistaprecio ADD dlp_codcec varchar(4), dlp_codcco varchar(10), dlp_dtsac varchar(6), dlp_nrosem int"
'   vg_db.BeginTrans
'   RS1.Open "SELECT * FROM a_derechosperfil WHERE SUBSTRING(CONVERT(CHAR(20),dpe_codopc),1,1)= '1'", vg_db, adOpenStatic
'   If Not RS1.EOF Then
'      vg_db.Execute "DELETE a_derechosperfil FROM a_derechosperfil WHERE SUBSTRING(CONVERT(CHAR(20),dpe_codopc),1,1)= '1'"
'      vg_db.Execute "DELETE a_opcsistema FROM a_opcsistema WHERE SUBSTRING(CONVERT(CHAR(20),opc_codigo),1,1)= '1'"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1000000, 'Productos', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1010000, 'Asociar Productos SAC vs SGP', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1020000, 'Proveedores', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1030000, 'Generación Archivos Planos Productos', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1040000, 'Lista de Precio', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1050000, 'Importar Lista de Precio Desde SAC', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1060000, 'Importar Lista de Precio Desde Excel', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1070000, 'Asociar Lista de Precio', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1080000, 'Actualizar Lista de Precio Planificación', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1090000, 'Recetas', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1100000, 'Generación Archivos Planos Recetas', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1110000, 'Planificación Minutas', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1120000, 'Tabla Gramaje', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1130000, 'Generación Archivos Planod Planificación', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1140000, 'Parametrizar NºRecetas 5 Etapas', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1150000, 'Parametrizar Costo Patrón Piso 5 Etapas', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1160000, 'Parametrizar Costo Patrón Techo 5 Etapas', null, null)"
'      vg_db.Execute "INSERT INTO a_opcsistema VALUES (1170000, 'Gramaje Familia Producto 5 Etapas', null, null)"
'      Do While Not RS1.EOF
'         CodOpc = 0
'         If RS1!dpe_codopc = 1010000 Then
'            CodOpc = 1000000
'         ElseIf RS1!dpe_codopc = 1011000 Then
'            CodOpc = 1010000
'         ElseIf RS1!dpe_codopc = 1012000 Then
'            CodOpc = 1020000
'         ElseIf RS1!dpe_codopc = 1020000 Then
'            CodOpc = 1030000
'         ElseIf RS1!dpe_codopc = 1015000 Then
'            CodOpc = 1040000
'         ElseIf RS1!dpe_codopc = 1016000 Then
'            CodOpc = 1060000
'         ElseIf RS1!dpe_codopc = 1017000 Then
'            CodOpc = 1070000
'         ElseIf RS1!dpe_codopc = 1018000 Then
'            CodOpc = 1080000
'         ElseIf RS1!dpe_codopc = 1030000 Then
'            CodOpc = 1090000
'         ElseIf RS1!dpe_codopc = 1040000 Then
'            CodOpc = 1100000
'         ElseIf RS1!dpe_codopc = 1050000 Then
'            CodOpc = 1110000
'         ElseIf RS1!dpe_codopc = 1060000 Then
'            CodOpc = 1120000
'         ElseIf RS1!dpe_codopc = 1070000 Then
'            CodOpc = 1130000
'         ElseIf RS1!dpe_codopc = 1080000 Then
'            CodOpc = 1140000
'         ElseIf RS1!dpe_codopc = 1090000 Then
'            CodOpc = 1150000
'         ElseIf RS1!dpe_codopc = 1092000 Then
'            CodOpc = 1160000
'         ElseIf RS1!dpe_codopc = 1100000 Then
'            CodOpc = 1170000
'         End If
'         If CodOpc > 0 Then
'            vg_db.Execute "INSERT INTO a_derechosperfil VALUES (" & RS1!dpe_codper & ", " & CodOpc & ", " & RS1!dpe_deracc & ", " & RS1!dpe_deragr & ", " & RS1!dpe_dermod & ", " & RS1!dpe_dereli & ", " & RS1!dpe_derimp & ")"
'         End If
'         RS1.MoveNext
'      Loop
'   End If
'   RS1.Close: Set RS1 = Nothing
'   vg_db.CommitTrans
   vg_db.Execute "UPDATE a_param SET par_valor='129' WHERE par_codigo='version'"
   aVer = 129
End If
If nVer > aVer And aVer = 129 Then
   '-------> Insertar campo tabla clientes con la opción sobreescribe receta si es igual a 0 = solo fijos, 1 = todos y 2 = ninguno
   vg_db.Execute "ALTER TABLE b_clientes ADD cli_sobrec varchar(1)"
   vg_db.Execute "UPDATE b_clientes SET cli_sobrec='0'"
   
   vg_db.Execute "UPDATE a_param SET par_valor='130' WHERE par_codigo='version'"
   aVer = 130
End If
If nVer > aVer And aVer = 130 Then
   vg_db.Execute "UPDATE a_param SET par_valor = '131' WHERE par_codigo = 'version'"
   aVer = 131
End If
If nVer > aVer And aVer = 131 Then
'   vg_db.Execute "UPDATE a_param SET par_valor = '132' WHERE par_codigo = 'version'"
   vg_db.Execute "sgpadm_iu_param 'M', 'version', '', '', '132'"
   aVer = 132
End If
If nVer > aVer And aVer = 132 Then
   vg_db.Execute "sgpadm_iu_param 'M', 'version', '', '', '133'"
   aVer = 133
End If
If nVer > aVer And aVer = 133 Then
   vg_db.Execute "sgpadm_iu_param 'M', 'version', '', '', '134'"
   aVer = 134
End If
End Function
