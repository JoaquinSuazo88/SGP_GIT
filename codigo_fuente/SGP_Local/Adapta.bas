Attribute VB_Name = "Adapta"
Dim nVer         As Long
Dim aVer         As Long
Dim cencos       As String
Dim codbod       As Long
Dim codigo       As Long
Dim Sql          As String
Dim sql1         As String
Dim sql2         As String
Dim RS           As New ADODB.Recordset
Dim RS1          As New ADODB.Recordset
Dim RS2          As New ADODB.Recordset
Dim RS3          As New ADODB.Recordset
Dim vg_dbsubesql As ADODB.Connection
Dim BaseDatos    As String
Dim fso          As Object
Dim procCompleto As String
   
Function ActVersion()

'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
Set fso = CreateObject("Scripting.FileSystemObject")

'On Error GoTo Man_Error
nVer = CLng(App.Major & App.Minor & App.Revision)
aVer = TipoDato(GetParametro("version"), 0)

If nVer < aVer Then
   
'   MsgBox "Versión del sistema no corresponde v" & aVer & VgLinea & "Realice la actualiazción en su PC o bien comunicase con la mesa de ayuda 8100555" & VgLinea & "        Proceso cancelado ...", vbCritical + vbOKOnly, "SGP"
   MsgBox "Versión del sistema no corresponde v" & nVer & ", " & VgLinea & _
          "actualmente esta con la versión " & aVer & ". " & VgLinea & _
          "Se bloquearan los acceso al sistema. " & VgLinea & _
          "Realice la actualización en su PC, desde menú principal del SGP." & VgLinea, vbCritical + vbOKOnly, "SGP"

'   End

End If

If nVer > aVer And aVer = 0 And nVer = 0 Then
    
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
    
    vg_db.Execute "insert into a_opcsistema values (2090000, 'Control Traspasos entre Contratos')"
    vg_db.Execute "insert into a_infcfcfofi values ('T', 1, 0, Null)"
    vg_db.Execute "update a_param set par_valor='103' where par_codigo='version'"
    aVer = 103

End If

If nVer > aVer And aVer = 103 Then
    
    vg_db.Execute "DROP TABLE b_minutaraciones"
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

'    vg_db.Execute "insert into a_opcsistema values (3043000, 'Informe Facturación Clientes')"
'    vg_db.Execute "insert into a_opcsistema values (3045000, 'Informe Stock')"
    vg_db.Execute "update a_param set par_valor='111' where par_codigo='version'"
    aVer = 111

End If

If nVer > aVer And aVer = 111 Then
    
    vg_db.Execute "update a_param set par_valor='112' where par_codigo='version'"
    aVer = 112

End If

If nVer > aVer And aVer = 112 Then
    
    vg_db.Execute "alter table b_totcompras add column toc_fledoc double"
    vg_db.Execute "update b_totcompras set toc_fledoc=0"
    vg_db.Execute "alter table b_detcompras add column dec_prefle double"
    vg_db.Execute "update b_detcompras set dec_prefle=0"
    vg_db.Execute "insert into a_opcsistema values (2120000, 'Cartola Inventario')"
    vg_db.Execute "insert into a_param values ('fvcon', 'Familia Verduras Congeladas', 'C', '98')"
    vg_db.Execute "update a_param set par_valor='35' where par_codigo='carne1'"
    vg_db.Execute "update a_param set par_valor='33' where par_codigo='carne2'"
    vg_db.Execute "update a_param set par_valor='113' where par_codigo='version'"
    aVer = 113

End If

If nVer > aVer And aVer = 113 Then
    
    vg_db.Execute "update a_param set par_valor='114' where par_codigo='version'"
    aVer = 114

End If

If nVer > aVer And aVer = 114 Then
    
    vg_db.Execute "update a_param set par_valor='115' where par_codigo='version'"
    aVer = 115

End If

If nVer > aVer And aVer = 115 Then
    
    vg_db.Execute "alter table b_minutapedido add column ped_stoact double"
    vg_db.Execute "alter table b_minutapedido add column ped_proped double"
    vg_db.Execute "update b_minutapedido set ped_stoact=0, ped_proped=0"
    vg_db.Execute "update a_param set par_valor='116' where par_codigo='version'"
    aVer = 116

End If

If nVer > aVer And aVer = 116 Then
    
    vg_db.Execute "update a_param set par_valor='117' where par_codigo='version'"
    aVer = 117

End If

If nVer > aVer And aVer = 117 Then
    
    vg_db.Execute "create table b_ventacontado (vtc_fecvta int, vtc_rutcli char(10), vtc_codreg int, vtc_codser int, vtc_forpag int, vtc_monvta double, Constraint b_ventacontado_pk Primary Key (vtc_fecvta, vtc_rutcli, vtc_codreg, vtc_codser, vtc_forpag))"

    vg_db.Execute "insert into a_opcsistema values (2077000, 'Venta Servicio Contado')"
    vg_db.Execute "update a_param set par_valor='118' where par_codigo='version'"
    aVer = 118

End If

If nVer > aVer And aVer = 118 Then
    
    vg_db.Execute "insert into a_opcsistema values (3100000, 'Informe Consulta Salida ó Devolución a Bodega')"
    vg_db.Execute "update a_param set par_valor='119' where par_codigo='version'"
    aVer = 119

End If

If nVer > aVer And aVer = 119 Then
    
    vg_db.Execute "alter table b_totcompras alter column toc_docaso longtext"
    vg_db.Execute "insert into a_opcsistema values (3047000, 'Ficha Stock')"
'    vg_db.Execute "insert into a_opcsistema values (2078000, 'Recalculo Precio Prom. Ponderado')"
    vg_db.Execute "alter table b_productos add column pro_ctrsto int"
    vg_db.Execute "update a_opcsistema set opc_codigo=2079000 where opc_codigo=2060000"
    vg_db.Execute "update a_derechosperfil set dpe_codopc=2079000 where dpe_codopc=2060000"
    vg_db.Execute "update b_productos set pro_ctrsto=iif(pro_ctacon='410001' or pro_ctacon='410004',1,0)"
    vg_db.Execute "update a_param set par_valor='120' where par_codigo='version'"
    aVer = 120

End If

If nVer > aVer And aVer = 120 Then
    
    vg_db.Execute "create table b_cierreperiodo (cie_periodo int, cie_fecini int, cie_fecter int, cie_estado int, Constraint b_cierreperiodo_pk Primary Key (cie_periodo))"
    '------- Traer cierres periodo
    Dim est As Boolean
    Dim fecini As Long, fecfin As Long, diatop As Long
    fecini = 0: fecfin = 0: diatop = 0: est = False
    RS1.Open "select distinct tin_fectom, tin_ciemes from b_tomainv where tin_ciemes<>0 order by tin_fectom", vg_db, adOpenStatic
    
    If Not RS1.EOF Then
       
       Do While Not RS1.EOF
          
          fecini = Format(dBoM(Mid(RS1!tin_fectom, 7, 2) & "/" & Mid(RS1!tin_fectom, 5, 2) & "/" & Mid(RS1!tin_fectom, 1, 4)), "yyyymmdd")
          fecfin = Format((Mid(RS1!tin_fectom, 7, 2) & "/" & Mid(RS1!tin_fectom, 5, 2) & "/" & Mid(RS1!tin_fectom, 1, 4)), "yyyymmdd")
          vg_db.Execute "insert into b_cierreperiodo values (" & Mid(fecini, 1, 6) & ", " & IIf(Mid(RS1!tin_fectom, 7, 2) > 27, fecini, (Mid(fecini, 1, 6) & fg_pone_cero(Str((diatop + 1)), 2))) & ", " & fecfin & ", 0)"
          diatop = Val(Mid(fecfin, 7, 2))
          RS1.MoveNext
       
       Loop
       If Not CierrePeriodo(fecfin, vg_codbod, 2) Then est = True: vg_db.Execute "UPDATE b_cierreperiodo SET cie_estado=1 WHERE cie_cencos='" & MuestraCasino(1) & "' AND cie_periodo=" & Val(Mid(fecini, 1, 6)) & ""
    
    End If
    RS1.Close: Set RS1 = Nothing
    
    Do While Mid(fecini, 1, 4) <> 2020
       
       If fecini = 0 And fecfin = 0 Then
          
          fecini = Format(dBoM(Date), "yyyymmdd"): fecfin = Format(dEoM(Date), "yyyymmdd")
          diatop = Val(Mid(fecfin, 7, 2))
          vg_db.Execute "insert into b_cierreperiodo values (" & Mid(fecini, 1, 6) & ", " & IIf(Mid(fecfin, 7, 2) > 27, fecini, (Mid(fecini, 1, 6) & (diatop + 1))) & ", " & fecfin & ", 2)"
       
       Else
          If (fecfin + 1) > Format(dEoM(Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymmdd") Then
             
             fecini = Format(dBoM(BEoM(Mid(fecini, 7, 2) & "/" & Mid(fecini, 5, 2) & "/" & Mid(fecini, 1, 4))), "yyyymmdd")
          
          Else
             
             fecini = Format(dEoM(Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymm") & fg_pone_cero(Str(Val(diatop + 1)), 2) 'fg_pone_cero(Str(Val(Mid(fecini, 7, 2))), 2)
          
          End If
          fecfin = Format(BEoM(Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymm") & fg_pone_cero(Str(Val(diatop)), 2) 'Mid(fecfin, 7, 2)
          
          If fecfin + 1 > Format(dEoM("01/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymmdd") Then
             
             fecfin = Format(dEoM("01/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymmdd")
          
          End If
          If (fecfin + 1) > Format(dEoM(Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymmdd") Then
             
             vg_db.Execute "insert into b_cierreperiodo values (" & IIf(diatop > 30, Mid(fecini, 1, 6), Mid(fecfin, 1, 6)) & ", " & IIf(Mid(fecfin, 7, 2) > 27, fecini, Mid(fecini, 1, 4) & Mid(fecini, 5, 2) & fg_pone_cero(Str(Val(diatop)), 2)) & ", " & Mid(fecfin, 1, 4) & Mid(fecfin, 5, 2) & Mid(fecfin, 7, 2) & ", 2)"
          
          Else
             
             vg_db.Execute "insert into b_cierreperiodo values (" & IIf(diatop > 30, Mid(fecini, 1, 6), Mid(fecfin, 1, 6)) & ", " & IIf(Mid(fecfin, 7, 2) > 27, fecini, Mid(fecini, 1, 4) & Mid(fecini, 5, 2) & fg_pone_cero(Str(Val(diatop + 1)), 2)) & ", " & Mid(fecfin, 1, 4) & Mid(fecfin, 5, 2) & Mid(fecfin, 7, 2) & ", 2)"
          
          End If
       
       End If
       If est = False Then est = True: vg_db.Execute "update b_cierreperiodo set cie_estado=1 where cie_periodo=" & Val(Mid(fecini, 1, 6)) & ""
    
    Loop
    '------- Fin cierres periodo
    vg_db.Execute "insert into a_opcsistema values (4160000, 'Calendario de Cierres de Mes')"
    vg_db.Execute "update a_param set par_valor='121' where par_codigo='version'"
    aVer = 121

End If

If nVer > aVer And aVer = 121 Then
    
    vg_db.Execute "alter table b_totcompras add column toc_docsnc char(255)"
    vg_db.Execute "update b_detcompras set dec_prefle=0 where dec_prefle is null"
    vg_db.Execute "update b_totcompras set toc_fledoc=0 where toc_fledoc is null"
    vg_db.Execute "update b_totcompras set toc_docsnc = NULL"
    '------- Traer Solicitud pendientes y actualizar el nuevo campo toc_docscn
    RS1.Open "SELECT b_totcompras.* FROM b_totcompras " & _
             "where VAL(b_totcompras.toc_docaso) IN (SELECT b_totcompras.toc_numdoc FROM b_totcompras WHERE b_totcompras.toc_tipdoc='NC' AND (NOT ISNULL (b_totcompras.toc_docaso) or b_totcompras.toc_docaso<>'')) " & _
             "AND   b_totcompras.toc_tipdoc = 'SN' AND (b_totcompras.toc_docsnc = '' OR ISNULL(b_totcompras.toc_docsnc))", vg_db, adOpenStatic
    
    If Not RS1.EOF Then
       
       Do While Not RS1.EOF
          
          vg_db.Execute "update b_totcompras set toc_docsnc = '" & RS1!toc_docaso & "' where toc_numdoc = " & RS1!toc_numdoc & " and toc_rutpro='" & RS1!toc_rutpro & "' and toc_tipdoc='" & RS1!toc_tipdoc & "' and cdate(toc_fecemi)='" & RS1!toc_fecemi & "'"
          RS1.MoveNext
       
       Loop
    
    End If
    RS1.Close: Set RS1 = Nothing
    '------- Fin traer Solicitud pendientes y actualizar el nuevo campo toc_docscn
    
    '------- Actualizar Solicitud pendientes agregando impuesto
    vg_db.Execute "insert into b_detcomprasimp (imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp)  " & _
                  "select distinct a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, b.ipr_codimp, c.imp_pctimp, (( ((a.dec_canmer-a.dec_canrec)*a.dec_prerec) - ((a.dec_canmer-a.dec_canrec)*a.dec_prerec)*(a.dec_pctdes / 100))*(c.imp_pctimp/100)) " & _
                  "from b_detcompras a, b_productosimp b, a_impuesto c where a.dec_codmer=b.ipr_codpro and b.ipr_codimp=c.imp_codigo and a.dec_tipdoc='SN'"
    '------- Fin actualizar Solicitud pendientes agregando impuesto
    vg_db.Execute "ALTER TABLE b_cierreperiodo ADD COLUMN cie_proantali double, cie_gdpenmesali double, cie_gdpenmesantali double, cie_sncpenmesali double, cie_sncpenmesantali double, cie_proantgrl double, cie_gdpenmesgrl double, cie_gdpenmesantgrl double, cie_sncpenmesgrl double, cie_sncpenmesantgrl double, cie_proantdes double, cie_gdpenmesdes double, cie_gdpenmesantdes double, cie_sncpenmesdes double, cie_sncpenmesantdes double"
    vg_contra = "23380"
    vg_codbod = 1
    vg_db.Execute "UPDATE b_cierreperiodo SET cie_proantali =0, cie_gdpenmesali =0, cie_gdpenmesantali =0, cie_sncpenmesali =0, cie_sncpenmesantali =0, cie_proantgrl =0, cie_gdpenmesgrl =0, cie_gdpenmesantgrl =0, cie_sncpenmesgrl =0, cie_sncpenmesantgrl =0, cie_proantdes =0, cie_gdpenmesdes =0, cie_gdpenmesantdes =0, cie_sncpenmesdes =0, cie_sncpenmesantdes =0 WHERE cie_cencos='" & MuestraCasino(1) & "'"
    RS1.Open "SELECT * FROM b_cierreperiodo WHERE cie_cencos='" & MuestraCasino(1) & "' AND cie_estado=0", vg_db, adOpenStatic
    If Not RS1.EOF Then
       
       Do While Not RS1.EOF
          
          CalcularProvisiones MuestraCasino(1), RS1!cie_periodo, RS1!cie_fecini, RS1!cie_fecter, True
          RS1.MoveNext
       
       Loop
    
    End If
    RS1.Close: Set RS1 = Nothing
    
    vg_db.Execute "insert into a_opcsistema values (2130000, 'Control Facturas Compras(Cierre de Mes)')"
    vg_db.Execute "update a_param set par_valor='123' where par_codigo='version'"
    aVer = 123

End If

If nVer > aVer And aVer = 123 Then
    
    '------- Borrar campo tiprec de texto y transformar a campo entero, del encabezado receta
    vg_db.Execute "alter table b_receta drop column rec_tiprec"
    vg_db.Execute "alter table b_receta add column rec_tiprec int"
    vg_db.Execute "update b_receta set rec_tiprec=1"
    
    '------- Borrar clave principal y incluir un nuevo campo del tipo receta, volver a crear la clave principal incluyendo el nuevo campo,
    '------- luego crear el concepto de receta patrón y local, del detalle receta
    vg_db.Execute "drop index PrimaryKey on b_recetadet"
    vg_db.Execute "alter table b_recetadet add column red_tiprec int"
    vg_db.Execute "update b_recetadet set red_tiprec='1'"
    vg_db.Execute "alter table b_recetadet add Constraint b_recetadet_pk Primary Key (red_codigo, red_nroite, red_tiprec)"
    RS1.Open "select * from b_receta ", vg_db, adOpenStatic
    
    If Not RS1.EOF Then
       
       Do While Not RS1.EOF
          
          vg_db.Execute "insert into b_recetadet (red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec) select red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, 0 from b_recetadet where red_codigo=" & RS1!rec_codigo & ""
          RS1.MoveNext
       
       Loop
    
    End If
    RS1.Close: Set RS1 = Nothing
    
    '------- Crear nuevo campo en encabezado minutas comensales teorico y reales, luego mover raciones desde raciones vendidas
    vg_db.Execute "alter table b_minuta add column min_racteo int, min_racrea int"
    vg_db.Execute "update b_minuta set min_racteo=0, min_racrea=0"
    RS1.Open " select mir_cencos, mir_codreg, mir_codser, mir_fecmin, sum(mir_nrorac) as nrorac from b_minutaraciones group by mir_cencos, mir_codreg, mir_codser, mir_fecmin order by mir_fecmin", vg_db, adOpenStatic
    If Not RS1.EOF Then
       
       Do While Not RS1.EOF
          
          vg_db.Execute "update b_minuta set min_racteo=" & RS1!nrorac & ", min_racrea=" & RS1!nrorac & " where min_cencos='" & RS1!mir_cencos & "' and min_codreg=" & RS1!mir_codreg & " and min_codser=" & RS1!mir_codser & " and min_fecmin=" & RS1!mir_fecmin & ""
          RS1.MoveNext
       
       Loop
    
    End If
    RS1.Close: Set RS1 = Nothing
    '------- Crear nuevo campo en detalle minutas tipo receta
    vg_db.Execute "alter table b_minutadet add column mid_tiprec int"
    vg_db.Execute "update b_minutadet set mid_tiprec=1"
    
    '------- Crear nueva tabla comensales estimados
    vg_db.Execute "create table a_serviciorac (sra_codser int, sra_coditem int, sra_serdia int, sra_raciones int, Constraint a_serviciorac_pk Primary Key (sra_codser, sra_coditem, sra_serdia))"
    vg_db.Execute "update a_param set par_valor='124' where par_codigo='version'"
    aVer = 124

End If

If nVer > aVer And aVer = 124 Then
    
    vg_db.Execute "update a_param set par_valor='125' where par_codigo='version'"
    aVer = 125

End If

If nVer > aVer And aVer = 125 Then
    
    '------- Crear nuevas opciones
    vg_db.Execute "insert into a_opcsistema values (2077400, 'Lista de Precio Cafetería')"
    vg_db.Execute "insert into a_opcsistema values (2077600, 'Registro de Venta Cafetería')"

    '------- Crear tablas para Lista de Precio Cafetería
    vg_db.Execute "create table b_totpreciocaf (tpc_codigo char(20), tpc_nombre char(50), tpc_precio double, Constraint b_totpreciocaf_pk Primary Key (tpc_codigo))"
    vg_db.Execute "create table b_detpreciocaf (dpc_codigo char(20), dpc_codmer char(20), dpc_cantidad double, Constraint b_detpreciocaf_pk Primary Key (dpc_codigo, dpc_codmer))"
    vg_db.Execute "ALTER TABLE b_detpreciocaf ADD CONSTRAINT FK_b_detpreciocaf_b_totpreciocaf FOREIGN KEY (dpc_codigo) REFERENCES b_totpreciocaf (tpc_codigo)"
    vg_db.Execute "ALTER TABLE b_detpreciocaf ADD CONSTRAINT FK_b_detpreciocaf_b_productos FOREIGN KEY (dpc_codmer) REFERENCES b_productos (pro_codigo)"

    '------- Crear tablas para Venta Cafetería
    vg_db.Execute "create table b_totventascaf (tvc_cencos char(10), tvc_fecing date, tvc_codbod int, tvc_estado char(1), Constraint b_totventascaf_pk Primary Key (tvc_cencos, tvc_fecing))"
    vg_db.Execute "create table b_detventascaf (dvc_cencos char(10), dvc_fecing date, dvc_numlin int, dvc_articulo char(20), dvc_canart double, dvc_precio double, dvc_tippag char(2), dvc_rutcli char(10), dvc_cencli char(10), dvc_tipdoc char(2), dvc_numdoc int, dvc_fecdoc date, Constraint b_detventascaf_pk Primary Key (dvc_cencos, dvc_fecing, dvc_numlin))"
    vg_db.Execute "create table b_detventascafpro (dvp_cencos char(10), dvp_fecing date, dvp_codmer char(20), dvp_cancal double, dvp_candig double, dvp_precos double, Constraint b_detventascafpro_pk Primary Key (dvp_cencos, dvp_fecing, dvp_codmer))"
    
    vg_db.Execute "ALTER TABLE b_detventascaf ADD CONSTRAINT FK_b_detventascaf_b_totventascaf FOREIGN KEY (dvc_cencos, dvc_fecing) REFERENCES b_totventascaf (tvc_cencos, tvc_fecing)"
    vg_db.Execute "ALTER TABLE b_detventascafpro ADD CONSTRAINT FK_b_detventascafpro_b_totventascaf FOREIGN KEY (dvp_cencos, dvp_fecing) REFERENCES b_totventascaf (tvc_cencos, tvc_fecing)"
    
    vg_db.Execute "update a_param set par_valor='126' where par_codigo='version'"
    aVer = 126

End If

If nVer > aVer And aVer = 126 Then
    
    vg_db.Execute "ALTER TABLE b_detventascaf ADD CONSTRAINT FK_b_detventascaf_b_totpreciocaf FOREIGN KEY (dvc_articulo) REFERENCES b_totpreciocaf (tpc_codigo)"
    vg_db.Execute "ALTER TABLE b_detventascafpro ADD CONSTRAINT FK_b_detventascafpro_b_productos FOREIGN KEY (dvp_codmer) REFERENCES b_productos (pro_codigo)"
    vg_db.Execute "insert into a_opcsistema values (3110000, 'Ventas por artículos de cafetería ')"
    vg_db.Execute "insert into a_opcsistema values (3120000, 'Ventas cafetería por cliente-centro costo')"
    vg_db.Execute "insert into a_opcsistema values (3130000, 'Ventas cafetería por cliente-centro costo detalle')"
    vg_db.Execute "insert into a_opcsistema values (3140000, 'Salida de bodega por ventas cafetería')"
    vg_db.Execute "update a_param set par_valor='127' where par_codigo='version'"
    aVer = 127

End If

If nVer > aVer And aVer = 127 Then
    
    vg_db.Execute "update a_param set par_valor='128' where par_codigo='version'"
    aVer = 128

End If

If nVer > aVer And aVer = 128 Then
   
   '------- actualizar encabezado y detalle receta los campo rec_tiprec-red_tiprec, receta local=1 reemplazara -1
   vg_db.Execute "update b_receta set rec_tiprec=-1 where rec_tiprec=1"
   vg_db.Execute "update b_recetadet set red_tiprec=-1 where red_tiprec=1"
   '------- actualizar detalle planificación el campo mid_tiprec, receta local=1 reemplazara -1
   vg_db.Execute "update b_minutadet set mid_tiprec=-1 where mid_tiprec=1"
   vg_db.Execute "update a_param set par_valor='129' where par_codigo='version'"
   vg_db.Execute "insert into a_param values ('5etapas', 'Contratos 5 Etapas', 'C', 'N')"
   aVer = 129

End If

If nVer > aVer And aVer = 129 Then
   
   '------- Actualizar Cuenta contable Movilización
   vg_db.Execute "UPDATE a_param SET par_valor='410042' WHERE par_codigo='ctamovil'"
   '------- Incluir un concepto en tabla de impuesto, llamado código sap
   vg_db.Execute "ALTER TABLE a_impuesto ADD COLUMN imp_codsap CHAR(20)"
   vg_db.Execute "UPDATE a_impuesto SET imp_codsap='123060' WHERE imp_codigo=1"
   vg_db.Execute "UPDATE a_impuesto SET imp_codsap='123070' WHERE imp_codigo=3"
   vg_db.Execute "UPDATE a_impuesto SET imp_codsap='123050' WHERE imp_codigo=8"
   '------- Anexar datos a la tabla a_param para armar plano cfc sap, para armar encabezado y detalle
   vg_db.Execute "INSERT INTO a_param VALUES ('claencsap', 'Clave Contabilización Sap Encabezado', 'N', '31')"
   vg_db.Execute "INSERT INTO a_param VALUES ('cladetsap', 'Clave Contabilización Sap Detalle', 'N', '40')"
   '------- Anexar datos a la tabla a_param para armar plano cfc sap impuesto
   vg_db.Execute "INSERT INTO a_param VALUES ('docexento', 'Indicador Impuesto Sap Exento', 'C', 'C0')"
   vg_db.Execute "INSERT INTO a_param VALUES ('docafecto', 'Indicador Impuesto Sap Efecto', 'C', 'C1')"
   RS1.Open "SELECT * FROM a_param WHERE par_codigo='ctagastos2'", vg_db, adOpenStatic
   If RS1.EOF Then vg_db.Execute "INSERT INTO a_param VALUES ('ctagastos2', 'Cuentas de Gastos Generales', 'C', ' ')"
   RS1.Close: Set RS1 = Nothing
   '------- Anexar datos a la tabla a_param prar armar cuenta contable
   RS1.Open "SELECT * FROM a_param WHERE par_codigo='datcont'", vg_db, adOpenStatic
   If RS1.EOF Then vg_db.Execute "INSERT INTO a_param VALUES ('datcont', 'Jorge PAz', 'C', 'jpaz@sodexho.cl')"
   RS1.Close: Set RS1 = Nothing
   '------- anexar concepto en tabla estructura servicio, llamado sector del menú
   vg_db.Execute "ALTER TABLE a_estservicio ADD COLUMN ess_codsec INT"
   vg_db.Execute "UPDATE a_estservicio SET ess_codsec=1"
   '------- Agregar tabla x sector
   vg_db.Execute "CREATE TABLE a_sector (sec_codigo INT, sec_nombre CHAR(50), sec_orden INT, CONSTRAINT a_sector_pk PRIMARY KEY (sec_codigo))"
   vg_db.Execute "INSERT INTO a_sector VALUES (1,'Sector No Definido',1)"
   vg_db.Execute "ALTER TABLE a_estservicio ADD CONSTRAINT PK_a_estservicio_a_sector FOREIGN KEY (ess_codsec) REFERENCES a_sector (sec_codigo)"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4117000, 'Sector')"
   '------- Anexar concepto en tabla detalle ventas llamado sector
   vg_db.Execute "ALTER TABLE b_detventas ADD COLUMN dev_codsec INT"
   '------- Incluir parametros para vizualizar ingredientes en salida producción, devolución producción, pedido mensual, pedido adicional
   vg_db.Execute "INSERT INTO a_param VALUES ('ingsalpro','Ocultar Ing. Salida 0=Muestra:1=Oculta','N', '0')"
   vg_db.Execute "INSERT INTO a_param VALUES ('ingdevpro','Ocultar Ing. Dev. 0=Muestra:1=Oculta','N', 0)"
   vg_db.Execute "INSERT INTO a_param VALUES ('ingpedmen','Ocultar Ing. Ped.M. 0=Muestra:1=Oculta','N', 0)"
   vg_db.Execute "INSERT INTO a_param VALUES ('ingpedadi','Ocultar Ing. Ped.Ad. 0=Muestra:1=Oculta','N', 0)"
   '------- Incluir parametro vizualizar salida producción
   vg_db.Execute "INSERT INTO a_param VALUES ('salressec','Sal. resumido-Sector 0=Resumido:1=Sector','N', 0)"
   vg_db.Execute "UPDATE a_param SET par_valor='130' WHERE par_codigo='version'"
   aVer = 130

End If
'If nVer > aVer And aVer = 130 Then
'   '------- Agregar tabla x control raciones vigatec
'   vg_db.Execute "CREATE TABLE b_alumnosraciones (alr_cencos char(10), alr_rutalumno char(9), alr_fecha int, alr_nombrealumno CHAR(50), alr_curso char(20), alr_nombreapoderado char(50), alr_servicio char(20), alr_raciones int, alr_codregimen int, alr_codservicio int, CONSTRAINT b_alumnosraciones_pk PRIMARY KEY (alr_cencos, alr_rutalumno, alr_fecha))"
'   vg_db.Execute "CREATE TABLE b_pagoalumnos (pal_cencos char(10), pal_rutalumno char(12), pal_fecha int, pal_nombrealumno CHAR(100), pal_curso char(50), pal_nombreapoderado char(100), pal_servicio char(50), pal_raciones int, pal_montopago double, pal_codregimen int, pal_codservicio int, CONSTRAINT b_controlraciones_pk PRIMARY KEY (pal_cencos, pal_rutalumno, pal_fecha))"
'   vg_db.Execute "UPDATE a_param SET par_valor='131' WHERE par_codigo='version'"
'   aVer = 131
'End If
If nVer > aVer And aVer = 130 Then
   
   vg_db.Execute "UPDATE a_param SET par_valor='131' WHERE par_codigo='version'"
   aVer = 131

End If

If nVer > aVer And aVer = 131 Then
   
   vg_db.Execute "INSERT INTO a_param VALUES ('carne3', 'Familia Carnes Ave', 'C', '32')"
   vg_db.Execute "UPDATE a_param SET par_nombre='Familia Carnes Vacuno' WHERE par_codigo='carne1'"
   vg_db.Execute "UPDATE a_param SET par_nombre='Familia Carnes Cerdo' WHERE par_codigo='carne2'"
   vg_db.Execute "UPDATE a_param SET par_valor='132' WHERE par_codigo='version'"
   aVer = 132

End If

If nVer > aVer And aVer = 132 Then
   
   vg_db.Execute "UPDATE a_param SET par_valor='133' WHERE par_codigo='version'"
   aVer = 133

End If

If nVer > aVer And aVer = 133 Then
   
   '------- Actualizar ingrediente que tengan precio negativo
    vg_db.Execute "UPDATE b_ingrediente a INNER JOIN b_productos b ON (a.ing_codped=b.pro_codigo) AND (a.ing_codcom=b.pro_codigo) SET a.ing_precos = iif(b.pro_propon<0,0,b.pro_propon/b.pro_facing)" & _
                  "WHERE a.ing_precos< 0"
   '------- Mover zero al stock negativo
   vg_db.Execute "UPDATE b_bodegas set bod_canmer=0 WHERE bod_codbod=" & vg_codbod & " AND bod_canmer<0"
   '------- Actualizar costo ingrediente
   vg_db.Execute "ALTER TABLE b_receta ADD rec_fecvig int"
   vg_db.Execute "UPDATE b_receta set rec_fecvig=0"
   vg_db.Execute "UPDATE a_param SET par_valor='134' WHERE par_codigo='version'"
   aVer = 134

End If

If nVer > aVer And aVer = 134 Then
   
   '------- Crear tablas para Presupuesto y Proyección
   vg_db.Execute "CREATE TABLE b_presupuestoproyeccion (ppr_cencos char(10),ppr_anomes int, ppr_tipo char(1), ppr_codigo int, ppr_descripcion char(50), ppr_valor double, Constraint b_presupuestoproyeccion_pk Primary Key (ppr_cencos, ppr_anomes, ppr_tipo, ppr_codigo))"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (2077800, 'Presupuesto y Proyección')"
   '------- Insertar campo total cantidad recibidad en tabla detalle compras
   vg_db.Execute "ALTER TABLE b_detcompras ADD dec_ptotrec double"
   vg_db.Execute "UPDATE b_detcompras set dec_ptotrec=iif(dec_canmer=dec_canrec AND dec_precom=dec_prerec,dec_ptotal,(dec_canrec*dec_prerec))"
   
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3032000, 'Food Cost Diario')"
   vg_db.Execute "UPDATE a_param SET par_valor='135' WHERE par_codigo='version'"
   aVer = 135

End If

If nVer > aVer And aVer = 135 Then
   
   '------- Incluir parametros para vizualizar familia productos en toma inventario
   vg_db.Execute "INSERT INTO a_param VALUES ('opfampro','Ocultar fam. prod. toma inv.','N', '0')"
   vg_db.Execute "UPDATE a_param SET par_valor='136' WHERE par_codigo='version'"
   aVer = 136

End If

If nVer > aVer And aVer = 136 Then
   '------- Actualizar ptotal de salida producción
   RS1.Open "SELECT * FROM b_totventas WHERE tov_fecpro>=cdate('01/06/2006') AND tov_tipdoc='SP'", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         vg_db.Execute "UPDATE b_detventas SET dev_ptotal=(dev_precos*dev_canmer) WHERE dev_rutcli='" & RS1!tov_rutcli & "' AND dev_numdoc=" & RS1!tov_numdoc & " AND dev_tipdoc='" & RS1!tov_tipdoc & "'"
         RS2.Open "SELECT dev_tipdoc, ROUND(SUM(dev_ptotal),0) AS ptotal FROM b_detventas WHERE dev_rutcli='" & RS1!tov_rutcli & "' AND dev_numdoc=" & RS1!tov_numdoc & " AND dev_tipdoc='" & RS1!tov_tipdoc & "' GROUP BY dev_tipdoc", vg_db, adOpenStatic
         If Not RS2.EOF Then
            vg_db.Execute "UPDATE b_totventas SET tov_totdoc=" & RS2!ptotal & " WHERE tov_rutcli='" & RS1!tov_rutcli & "' AND tov_numdoc=" & RS1!tov_numdoc & " AND tov_tipdoc='" & RS1!tov_tipdoc & "'"
         End If
         RS2.Close: Set RS2 = Nothing
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   vg_db.Execute "UPDATE a_param SET par_valor='137' WHERE par_codigo='version'"
   aVer = 137
End If
If nVer > aVer And aVer = 137 Then
   '------- Incluir nuevo concepto en detalle receta cantidad x merma
   vg_db.Execute "ALTER TABLE b_minutadet ADD COLUMN mid_nummer DOUBLE"
   '------- Insertar mantenedor y listador raciones no vendidas
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (2052000, 'Merma x Preparación')"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3082000, 'Merma x Preparación')"
   '------- Tabla parametro x despacho
   vg_db.Execute "CREATE TABLE b_paramdesp (pad_codigo INT, pad_tipo CHAR(1), CONSTRAINT b_paramdesp_pk PRIMARY KEY (pad_codigo))"
   RS1.Open "SELECT * FROM a_tipopro WHERE tip_previo=0", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         vg_db.Execute "INSERT INTO b_paramdesp (pad_codigo, pad_tipo) SELECT tip_codigo, 'S' FROM a_tipopro WHERE tip_previo=" & RS1!tip_codigo & ""
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4012000, 'Parametro Despacho')"
   '------- incluir concepto raciones producidas
   RS1.Open "SELECT mir_cencos, mir_codreg, mir_codser, mir_fecmin, SUM(mir_nrorac) AS nrorac FROM b_minutaraciones " & _
            "WHERE mir_rutcli<>'PERSONAL' GROUP BY mir_cencos, mir_codreg, mir_codser, mir_fecmin", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         vg_db.Execute "INSERT INTO b_minutaraciones VALUES ('" & RS1!mir_cencos & "', " & RS1!mir_codreg & ", " & RS1!mir_codser & ", " & RS1!mir_fecmin & ", 'PRODUCIDAS', " & RS1!nrorac & ")"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   'Actualizar ingrediente que tengan precio negativo
    vg_db.Execute "UPDATE b_ingrediente a INNER JOIN b_productos b ON (a.ing_codped=b.pro_codigo) AND (a.ing_codcom=b.pro_codigo) SET a.ing_precos = iif(b.pro_propon<0,0,b.pro_propon/b.pro_facing)" & _
                  "WHERE a.ing_precos< 0"
   '------- Mover zero al stock negativo
   vg_db.Execute "UPDATE b_bodegas set bod_canmer=0 WHERE bod_codbod=" & vg_codbod & " AND bod_canmer<0"
    vg_db.Execute "UPDATE a_param SET par_valor='138' WHERE par_codigo='version'"
   aVer = 138
End If
If nVer > aVer And aVer = 138 Then
'   'Crear tabla tabla paso receta 5 etapas
'   RS1.Open "SELECT DISTINCT red_codigo INTO paso FROM b_recetadet WHERE red_tiprec>=10000", vg_db, adOpenStatic
'   Set RS1 = Nothing
'   'Insertar recetas 5 etapas
'   RS1.Open "SELECT * FROM a_regimen WHERE reg_codigo>=10000", vg_db, adOpenStatic
'   Do While Not RS1.EOF
'      vg_db.Execute "INSERT INTO b_recetadet (red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec) SELECT DISTINCT red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, " & RS1!reg_codigo & " FROM b_recetadet WHERE red_tiprec=0 AND red_codigo NOT IN (SELECT red_codigo FROM paso)"
'      RS1.MoveNext
'   Loop
'   RS1.Close: Set RS1 = Nothing
'   vg_db.Execute "DROP TABLE PASO"
   'Crear tabla costo patron
   vg_db.Execute "CREATE TABLE b_costopatron (cpa_cencos char(10), cpa_codreg int, cpa_codser int, cpa_anomes int, cpa_descripcion char(10), cpa_valor double, Constraint b_costopatron_pk Primary Key (cpa_cencos, cpa_codreg, cpa_codser, cpa_anomes, cpa_descripcion))"
   '------- Incluir nuevo concepto en detalle receta cantidad x merma
   vg_db.Execute "ALTER TABLE b_minutadet ADD COLUMN mid_rec5eta CHAR(1)"
   vg_db.Execute "UPDATE b_minuta INNER JOIN b_minutadet ON b_minuta.min_codigo=b_minutadet.mid_codigo SET b_minutadet.mid_rec5eta=IIf(b_minuta.min_codreg>=10000 AND b_minuta.min_codser>=10000,'1','0')"
   vg_db.Execute "UPDATE a_param SET par_valor='139' WHERE par_codigo='version'"
   aVer = 139
End If
If nVer > aVer And aVer = 139 Then
   'Modificar concepto 5 etapas tabla servicio y regimen
   vg_db.Execute "UPDATE a_regimen SET reg_nombre=reg_nombre & '(5)' WHERE reg_codigo>=10000 AND len(trim(reg_nombre))<26"
   vg_db.Execute "UPDATE a_servicio SET ser_nombre=ser_nombre & '(5)' WHERE ser_codigo>=10000 AND len(trim(ser_nombre))<26"
   'Actualizar detalle receta campo 5 etapas
   vg_db.Execute "UPDATE b_minuta INNER JOIN b_minutadet ON b_minuta.min_codigo=b_minutadet.mid_codigo SET b_minutadet.mid_rec5eta=IIf(b_minuta.min_codreg>=10000 AND b_minuta.min_codser>=10000,'1','0')"
   'Actualizar ingrediente que tengan precio negativo
   vg_db.Execute "UPDATE b_ingrediente a INNER JOIN b_productos b ON (a.ing_codped=b.pro_codigo) AND (a.ing_codcom=b.pro_codigo) SET a.ing_precos=iif(b.pro_propon<0,0,b.pro_propon/b.pro_facing)" & _
                 "WHERE a.ing_precos< 0"
   'Mover zero al stock negativo
   vg_db.Execute "UPDATE b_bodegas set bod_canmer=0 WHERE bod_codbod=" & vg_codbod & " AND bod_canmer<0"
   'Crear campo % de precio en tabla param, para validar precio
   RS1.Open "SELECT * FROM a_param WHERE par_codigo='porprepro'", vg_db, adOpenStatic
   If RS1.EOF Then vg_db.Execute "INSERT INTO a_param VALUES ('porprepro', 'validación % precio producto', 'N', '20')"
   RS1.Close: Set RS1 = Nothing
   'Crear campo que indica acepta diferencia precio tabla detalle compra y detalle ventas
   vg_db.Execute "ALTER TABLE b_detcompras ADD dec_acepre char(1)"
   vg_db.Execute "UPDATE b_detcompras SET dec_acepre='N'"
   vg_db.Execute "ALTER TABLE b_detventas ADD dev_acepre char(1)"
   vg_db.Execute "UPDATE b_detventas SET dev_acepre='N'"
   'Crear tabla gramo familia producto
   vg_db.Execute "CREATE TABLE b_gramofamproducto (gfp_cencos char(10), gfp_codtip int, gfp_graini double, gfp_grafin double, Constraint b_gramofamproducto_pk Primary Key (gfp_cencos, gfp_codtip))"
   'Agregar concepto Impuesto a modificar y actualizar datos ya creado
   vg_db.Execute "ALTER TABLE a_impuesto ADD COLUMN imp_indmod CHAR(1)"
   vg_db.Execute "UPDATE a_impuesto SET imp_indmod='N'"
   vg_db.Execute "UPDATE a_param SET par_valor='140' WHERE par_codigo='version'"
   aVer = 140
End If
If nVer > aVer And aVer = 140 Then
   'Incluir opción de salida producción preparada
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (2032000, 'Salida Producción Cerrada')"
   vg_db.Execute "UPDATE a_opcsistema SET opc_nombre='Salida Producción Preparada' WHERE opc_codigo=2030000"
   '------- Insertar listador costo x sector
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3034000, 'Costo x Sector')"
   'Incluir concepto raciones minima, en estructura servicio
    vg_db.Execute "ALTER TABLE a_estservicio ADD ess_racmin FLOAT"
    vg_db.Execute "UPDATE a_estservicio SET ess_racmin=0"
'    vg_db.Execute "UPDATE a_estservicio SET ess_codsec=1 WHERE ess_codsec IS NULL OR ess_codsec=0"
   'Eliminar tabla gramosfamproductos
   vg_db.Execute "DROP TABLE b_gramofamproducto"
   '------- Crear tabla gramo familia producto
   vg_db.Execute "CREATE TABLE b_gramofamproducto (gfp_cencos char(10), gfp_catdie int, gfp_tiprec int, gfp_fampro int, gfp_graini double, gfp_grafin double, Constraint b_gramofamproducto_pk Primary Key (gfp_cencos, gfp_catdie, gfp_tiprec, gfp_fampro))"
   vg_db.Execute "UPDATE a_param SET par_valor='141' WHERE par_codigo='version'"
   aVer = 141
End If
If nVer > aVer And aVer = 141 Then
   vg_db.Execute "UPDATE a_param SET par_valor='142' WHERE par_codigo='version'"
   aVer = 142
End If
If nVer > aVer And aVer = 142 Then
   RS1.Open "SELECT * INTO paso_gramofamproducto FROM b_gramofamproducto", vg_db, adOpenStatic
   Set RS1 = Nothing
   '------- Eliminar tabla gramosfamproductos
   vg_db.Execute "DROP TABLE b_gramofamproducto"
   '------- Crear tabla gramo familia producto
   vg_db.Execute "CREATE TABLE b_gramofamproducto (gfp_cencos char(10), gfp_codreg int, gfp_catdie int, gfp_tiprec int, gfp_fampro int, gfp_graini double, gfp_grafin double, Constraint b_gramofamproducto_pk Primary Key (gfp_cencos, gfp_codreg, gfp_catdie, gfp_tiprec, gfp_fampro))"
   RS1.Open "SELECT * FROM a_regimen WHERE reg_codigo>=10000", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         vg_db.Execute "INSERT INTO b_gramofamproducto (gfp_cencos, gfp_codreg, gfp_catdie, gfp_tiprec, gfp_fampro, gfp_graini, gfp_grafin) SELECT gfp_cencos, " & RS1!reg_codigo & ", gfp_catdie, gfp_tiprec, gfp_fampro, gfp_graini, gfp_grafin FROM b_gramofamproducto WHERE gfp_cencos<>''"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   '------- Eliminar tabla temporal gramosfamproductos
   vg_db.Execute "DROP TABLE paso_gramofamproducto"
   vg_db.Execute "UPDATE a_param SET par_valor='143' WHERE par_codigo='version'"
   aVer = 143
End If
If nVer > aVer And aVer = 143 Then
   '------- Incluir opción de informe Insumos no Planificados en Salida Bodega
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3112000, 'Insumos no Planificados en Salida Bodega')"
   '------- Incluir opción de informe costo totales del período
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3022000, 'Costos Totales del Período')"
   '------- Modificar concepto food cost diario x food cost
   vg_db.Execute "UPDATE a_opcsistema SET opc_nombre='Food Cost' WHERE opc_codigo=3032000"
   '------- Modificar concepto costo teórico - real - food cost x costo teórico - real - realizado
   vg_db.Execute "UPDATE a_opcsistema SET opc_nombre='Costo Plan. Teórico - Plan. Real - Realizado' WHERE opc_codigo=3030000"
   vg_db.Execute "UPDATE a_param SET par_valor='144' WHERE par_codigo='version'"
   aVer = 144
End If
If nVer > aVer And aVer = 144 Then
   '------- Crear tabla minuta fija x día
   vg_db.Execute "CREATE TABLE b_minutafijadia (mfd_cencos char(10), mfd_codreg int, mfd_codser int, mfd_fecha int, mfd_codpro char(20), mfd_tipmin char(1), mfd_canpro double, mfd_cospro double, Constraint b_minutafijadia_pk Primary Key (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin))"
   '------- Crear campo grupo vulnerabilidad en tabla receta
   vg_db.Execute "ALTER TABLE b_receta ADD rec_gruvul MEMO"
   '------- Borrar de la tabla minutacosto lo referido a item tipo minuta de la estructura fija
   vg_db.Execute "DELETE b_minutacosto FROM b_minutacosto WHERE mic_tipmin='3'"
   '------- Incluir opción informe detalle ajuste inventario
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3114000, 'Ajuste Invetario')"
   '-------Actualizar versión
   vg_db.Execute "UPDATE a_param SET par_valor='145' WHERE par_codigo='version'"
   aVer = 145
End If
If nVer > aVer And aVer = 145 Then
   '------- Actualizar minuta 5 etapas
   RS1.Open "SELECT DISTINCT b.mid_codigo, b.mid_tipmin, b.mid_numlin, b.mid_codrec, b.mid_rec5eta FROM b_minuta a, b_minutadet b WHERE a.min_codigo=b.mid_codigo AND a.min_fecmin>=20070701 AND a.min_fecmin<=20070731 AND b.mid_tipmin='1'", vg_db, adOpenStatic
   Do While Not RS1.EOF
      vg_db.Execute "UPDATE b_minutadet SET mid_rec5eta=" & IIf(IsNull(RS1!mid_rec5eta), "Null", RS1!mid_rec5eta) & " WHERE mid_tipmin='2' AND mid_codigo=" & RS1!mid_codigo & " AND mid_numlin=" & RS1!mid_numlin & " AND mid_codrec=" & RS1!mid_codrec & " AND mid_rec5eta IS NULL"
      RS1.MoveNext
   Loop
   RS1.Close: Set RS1 = Nothing
   vg_db.Execute "UPDATE a_param SET par_valor='146' WHERE par_codigo='version'"
   aVer = 146
End If
If nVer > aVer And aVer = 146 Then
   '------- Crear tabla cliente centro costo
   vg_db.Execute "CREATE TABLE b_clientecencos (clc_codigo char(10), clc_codcli char(10), clc_nombre char(50), Constraint b_clientecencos_pk Primary Key (clc_codigo, clc_codcli))"
   vg_db.Execute "ALTER TABLE b_clientecencos ADD CONSTRAINT FK_b_clientecencos_b_clientes FOREIGN KEY (clc_codcli) REFERENCES b_clientes (cli_codigo)"
   '------- Actualizar tabla b_ventacontado y crear una tabla detalle
   RS1.Open "SELECT * INTO PASO FROM b_ventacontado", vg_db, adOpenStatic
   Set RS1 = Nothing
   vg_db.Execute "DROP TABLE b_ventacontado"
   vg_db.Execute "CREATE TABLE b_ventacontado (vtc_codigo int, vtc_cencos char(10), vtc_codreg int, vtc_codser int, vtc_fecvta int, vtc_forpag int, vtc_totmon double, vtc_opccli char(1), Constraint b_ventacontado_pk Primary Key (vtc_codigo))"
   codigo = 0: cencos = ""
   RS1.Open "SELECT * FROM a_param WHERE par_codigo='casino'", vg_db, adOpenStatic
   If Not RS1.EOF Then cencos = Trim(RS1!par_valor)
   RS1.Close: Set RS1 = Nothing
   RS1.Open "SELECT * FROM paso", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         RS2.Open "SELECT vtc_codigo FROM b_ventacontado ORDER BY vtc_codigo DESC", vg_db, adOpenStatic
         If Not RS2.EOF Then RS2.MoveFirst: codigo = RS2!vtc_codigo + 1 Else codigo = 1
         RS2.Close: Set RS2 = Nothing
         vg_db.Execute "INSERT INTO b_ventacontado VALUES (" & codigo & ", '" & cencos & "', " & RS1!vtc_codreg & ", " & RS1!vtc_codser & ", " & RS1!vtc_fecvta & ", " & RS1!vtc_forpag & ", " & RS1!vtc_monvta & ", '0')"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   vg_db.Execute "CREATE TABLE b_ventacontadodet (vtd_codigo int, vtd_numlin int, vtd_codcli char(10), vtd_codcco char(10), vtd_descripcion char(50), vtd_detmon double, Constraint b_ventacontadodet_pk Primary Key (vtd_codigo, vtd_numlin))"
   vg_db.Execute "ALTER TABLE b_ventacontadodet ADD CONSTRAINT FK_b_ventacontadodet_b_ventacontado FOREIGN KEY (vtd_codigo) REFERENCES b_ventacontado (vtc_codigo)"
   'Relacion tabla detalle venta contado con cliente centro de costo
   vg_db.Execute "ALTER TABLE b_ventacontadodet ADD CONSTRAINT FK_b_ventacontadodet_b_clientecencos FOREIGN KEY (vtd_codcco, vtd_codcli) REFERENCES b_clientecencos (clc_codigo, clc_codcli)"
   vg_db.Execute "DROP TABLE paso"
   'Incluir opción de exportación y importación planificación
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1032000, 'Exportar Planificación Minuta')"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1034000, 'Importar Planificación Minuta')"
   '------- Actualizar versión
   vg_db.Execute "UPDATE a_param SET par_valor='147' WHERE par_codigo='version'"
   aVer = 147
End If
If nVer > aVer And aVer = 147 Then
   '------- Crear tabla grupo paciente
   vg_db.Execute "CREATE TABLE a_grupopaciente (grp_codigo int, grp_nombre char(50), grp_othervalue char(100), grp_estado char(1), Constraint b_grupopaciente_pk Primary Key (grp_codigo))"
   '------- Crear tabla pacientes
   vg_db.Execute "CREATE TABLE b_pacientes (pac_codigo char(10), pac_nombre char(30), pac_appaterno char(20), pac_apmaterno char(20), pac_sexo char(1), pac_codgrp int, pac_nrocam char(10), pac_codreg int, pac_presdiet char(255), pac_comentario char(255), pac_fecing date, pac_fecalt date, pac_origen char(20), pac_estado char(1), Constraint b_pacientes_pk Primary Key (pac_codigo))"
   '------- Crear relación regimen con pacientes y departamento
   vg_db.Execute "ALTER TABLE b_pacientes ADD CONSTRAINT FK_b_pacientes_b_regimen FOREIGN KEY (pac_codreg) REFERENCES a_regimen (reg_codigo)"
   vg_db.Execute "ALTER TABLE b_pacientes ADD CONSTRAINT FK_b_pacientes_b_grupopaciente FOREIGN KEY (pac_codgrp) REFERENCES a_grupopaciente (grp_codigo)"
   '------- Crear tabla usuario vs grupo paciente
   vg_db.Execute "CREATE TABLE b_usuariogrupopac (ugp_codgrp int, ugp_codusu char(20), Constraint b_usuariogrupopac_pk Primary Key (ugp_codgrp, ugp_codusu))"
   '------- Crear relación regimen con pacientes y departamento
   vg_db.Execute "ALTER TABLE b_usuariogrupopac ADD CONSTRAINT FK_b_usuariogrupopac_b_usuario FOREIGN KEY (ugp_codusu) REFERENCES a_usuarios (usu_codigo)"
   vg_db.Execute "ALTER TABLE b_usuariogrupopac ADD CONSTRAINT FK_b_usuariogrupopac_a_grupopaciente FOREIGN KEY (ugp_codgrp) REFERENCES a_grupopaciente (grp_codigo)"
   '------- Agregar campo a la tabla servicio hora tope cobro, hora entrega, hora modif. PDA
   vg_db.Execute "ALTER TABLE a_servicio ADD ser_horcob date, ser_horent date, ser_horpda date" 'Hora Corta'" 'stampdate('HH:Nn')"
   '------- Crear tabla encabezado y detalle toma pedidos
   vg_db.Execute "CREATE TABLE b_tomapedido (top_codigo int, top_cencos char(10), top_codreg int, top_fecped date, top_codpac char(10), top_tipmen char(1), top_codusu char(20), Constraint b_tomapedido_pk Primary Key (top_codigo))"
   vg_db.Execute "CREATE TABLE b_tomapedidodet (tpd_codigo int, tpd_numlin int, tpd_codreg int, tpd_codser int, tpd_codmin int, tpd_estser int, tpd_codrec int, tpd_tiprec int, tpd_cansel double, tpd_canser double, tpd_caning double, tpd_prorec char(1), Constraint b_tomapedidodet_pk Primary Key (tpd_codigo, tpd_numlin))"
   vg_db.Execute "CREATE TABLE b_tomapedidodetrec (tdr_codigo int, tdr_numlin int, tdr_codrec int, tdr_nroite int, tdr_coding char(20), tdr_canpro double, tdr_cospro double, tdr_pctapr double, tdr_pctcoc double, tdr_pctnut double, Constraint b_tomapedidodetrec_pk Primary Key (tdr_codigo, tdr_numlin, tdr_codrec, tdr_nroite))"
   '------- Crear relación encabezado y detalle toma pedido
   vg_db.Execute "ALTER TABLE b_tomapedidodet ADD CONSTRAINT FK_b_tomapedidodet_b_tomapedido FOREIGN KEY (tpd_codigo) REFERENCES b_tomapedido (top_codigo)"
   vg_db.Execute "ALTER TABLE b_tomapedidodetrec ADD CONSTRAINT FK_b_tomapedidodetrec_b_tomapedidodet FOREIGN KEY (tdr_codigo, tdr_numlin) REFERENCES b_tomapedidodet (tpd_codigo, tpd_numlin)"
   '------- Crear relación encabezado toma pedido vs regimen
   vg_db.Execute "ALTER TABLE b_tomapedido ADD CONSTRAINT FK_b_tomapedido_b_regimen FOREIGN KEY (top_codreg) REFERENCES a_regimen (reg_codigo)"
   '------- Crear relación detalle toma pedido vs regimen
   vg_db.Execute "ALTER TABLE b_tomapedidodet ADD CONSTRAINT FK_b_tomapedidodet_b_regimen FOREIGN KEY (tpd_codreg) REFERENCES a_regimen (reg_codigo)"
   '------- Crear relación detalle toma pedido vs servicio
   vg_db.Execute "ALTER TABLE b_tomapedidodet ADD CONSTRAINT FK_b_tomapedidodet_b_servicio FOREIGN KEY (tpd_codser) REFERENCES a_servicio (ser_codigo)"
   '------- Crear relación encabezado toma pedido vs servicio
   vg_db.Execute "ALTER TABLE b_tomapedido ADD CONSTRAINT FK_b_tomapedidodet_a_usuarios FOREIGN KEY (top_codusu) REFERENCES a_usuarios (usu_codigo)"
   '------- Crear relación encabezado toma pedido vs pacientes
   vg_db.Execute "ALTER TABLE b_tomapedido ADD CONSTRAINT FK_b_tomapedido_b_pacientes FOREIGN KEY (top_codpac) REFERENCES b_pacientes (pac_codigo)"
   'Modificación a informe A13 dejando
   vg_db.Execute "UPDATE a_param SET par_valor='148' WHERE par_codigo='version'"
   aVer = 148
End If
If nVer > aVer And aVer = 148 Then
   '------- Icluir opciones paciente
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (5010000, 'Paciente - Horario Servicio')"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (5020000, 'Paciente - Grupo Paciente')"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (5030000, 'Paciente - Usuario Grupo Paciente')"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (5040000, 'Paciente - Pacientes')"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (5050000, 'Paciente- Toma Pedido')"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (5060000, 'Paciente- Control de Ingesta')"

   vg_db.Execute "INSERT INTO a_opcsistema VALUES (5070000, 'Paciente - Informe de Aporte Nutricional')"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (5080000, 'Paciente - Informe de Producción')"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (5090000, 'Paciente - Informe Detalle de Consumo')"
   '------- Agregar campo a la tabla detalle receta toma pedido
   vg_db.Execute "ALTER TABLE b_tomapedidodetrec ADD tdr_cannco double, tdr_cancon double"
   vg_db.Execute "UPDATE a_param SET par_valor='149' WHERE par_codigo='version'"
   aVer = 149
End If
If nVer > aVer And aVer = 149 Then
   '------- Consultar si existe codigo bodega
   RS1.Open "SELECT DISTINCT * FROM a_bodega ORDER BY bod_codigo", vg_db, adOpenStatic
   If RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "No existe bodega, proceso cancelado " & VgLinea & "Comunicase con departamento de informatica" & VgLinea & "        Proceso cancelado ...", vbCritical + vbOKOnly, "SGP": Exit Function
   codbod = RS1!bod_codigo
   RS1.Close: Set RS1 = Nothing
   '------- Crear tabla usuario vs contratos
   vg_db.Execute "CREATE TABLE b_usuariocontratos (uco_codusu char(20), uco_codcon char(20), Constraint b_usuariocontratos_pk Primary Key (uco_codusu, uco_codcon))"
   '------- Crear relación tabla usuario contratos
   vg_db.Execute "ALTER TABLE b_usuariocontratos ADD CONSTRAINT FK_b_usuariocontratos_a_usuarios FOREIGN KEY (uco_codusu) REFERENCES a_usuarios (usu_codigo)"
   vg_db.Execute "ALTER TABLE b_usuariocontratos ADD CONSTRAINT FK_b_usuariocontratos_b_clientes FOREIGN KEY (uco_codcon) REFERENCES b_clientes (cli_codigo)"
   '------- Mover datos a tabla b_usuariocontratos
   RS1.Open "SELECT * FROM a_usuarios", vg_db, adOpenStatic
   Do While Not RS1.EOF
      vg_db.Execute "INSERT INTO b_usuariocontratos (uco_codusu, uco_codcon) SELECT DISTINCT '" & RS1!usu_codigo & "', a.par_valor FROM a_param a, b_clientes b WHERE a.par_valor=b.cli_codigo AND b.cli_tipo=0 AND a.par_codigo='casino'"
      RS1.MoveNext
   Loop
   RS1.Close: Set RS1 = Nothing
   '------- Modificar concepto contrato x contrato
   vg_db.Execute "UPDATE a_opcsistema SET opc_nombre='Contratos' WHERE opc_codigo=4110000"
   '------- Agregar campo a la tabla clientes el concepto bodega
   vg_db.Execute "ALTER TABLE b_clientes ADD cli_codbod int, cli_codtis int, cli_codseg int"
   vg_db.Execute "UPDATE b_clientes SET cli_codbod=0"
   '------- Crear relación tabla cliente vs bodega
   
'   vg_db.Execute "ALTER TABLE a_bodega ADD CONSTRAINT FK_a_bodega_b_clientes FOREIGN KEY (bod_codigo) REFERENCES a_bodega (cli_codbod)"
   '------- Mover datos al campo codbod de tabla cliente
   RS1.Open "SELECT bod_codigo FROM a_bodega", vg_db, adOpenStatic
   If Not RS1.EOF Then
      vg_db.Execute "UPDATE b_clientes INNER JOIN b_usuariocontratos ON b_clientes.cli_codigo=b_usuariocontratos.uco_codcon SET b_clientes.cli_codbod =" & RS1!bod_codigo & ""
   End If
   RS1.Close: Set RS1 = Nothing
   
   '-------- Agregar campo contrato a tabla cfc-fofi
   RS1.Open "SELECT * FROM a_param WHERE par_codigo='casino'", vg_db, adOpenStatic
   If Not RS1.EOF Then cencos = RS1!par_valor Else RS1.Close: Set RS1 = Nothing: MsgBox "No existe parametrizaicón de contrato en sistema " & VgLinea & "Comunicase con departamento de informatica" & VgLinea & "        Proceso cancelado ...", vbCritical + vbOKOnly, "SGP": Exit Function
   RS1.Close: Set RS1 = Nothing
   RS1.Open "SELECT * INTO PASO FROM a_infcfcfofi", vg_db, adOpenStatic
   vg_db.Execute "DROP TABLE a_infcfcfofi"
   vg_db.Execute "CREATE TABLE a_infcfcfofi (inf_cencos char(10), inf_tipo char(1), inf_numero int, inf_feccie int, inf_usuario char(20), Constraint a_infcfcfofi_pk Primary Key (inf_cencos, inf_tipo, inf_numero))"
   '------- Crear relación tabla infcfcfofi con contrato
   vg_db.Execute "ALTER TABLE a_infcfcfofi ADD CONSTRAINT FK_a_infcfcfofi_b_clientes FOREIGN KEY (inf_cencos) REFERENCES b_clientes (cli_codigo)"
   '------- Insertar datos tabla infcfcfofi
   vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie, inf_usuario) SELECT '" & cencos & "', inf_tipo, inf_numero, inf_feccie, inf_usuario FROM PASO"
   vg_db.Execute "DROP TABLE PASO"
   
   '------- Modificar tabla calendario cierre de mes
   RS1.Open "SELECT * INTO PASO FROM b_cierreperiodo", vg_db, adOpenStatic
   Set RS1 = Nothing
   vg_db.Execute "DROP TABLE b_cierreperiodo"
   vg_db.Execute "CREATE TABLE b_cierreperiodo (cie_cencos char(10), cie_periodo int, cie_fecini int, cie_fecter int, cie_estado int, cie_proantali double, cie_gdpenmesali double, cie_gdpenmesantali double, cie_sncpenmesali double, cie_sncpenmesantali double, cie_proantgrl double, cie_gdpenmesgrl double, cie_gdpenmesantgrl double, cie_sncpenmesgrl double, cie_sncpenmesantgrl double, cie_proantdes double, cie_gdpenmesdes double, cie_gdpenmesantdes double, cie_sncpenmesdes double, cie_sncpenmesantdes double, Constraint a_infcfcfofi_pk Primary Key (cie_cencos, cie_periodo))"
   '------- Crear relación tabla calendario cierre de mes con contrato
   vg_db.Execute "ALTER TABLE b_cierreperiodo ADD CONSTRAINT FK_b_cierreperiodo_b_clientes FOREIGN KEY (cie_cencos) REFERENCES b_clientes (cli_codigo)"
   '------- Insertar datos tabla calendario cierre de mes
   vg_db.Execute "INSERT INTO b_cierreperiodo (cie_cencos, cie_periodo, cie_fecini, cie_fecter, cie_estado, cie_proantali, cie_gdpenmesali, cie_gdpenmesantali, cie_sncpenmesali, cie_sncpenmesantali, cie_proantgrl, cie_gdpenmesgrl, cie_gdpenmesantgrl, cie_sncpenmesgrl, cie_sncpenmesantgrl, cie_proantdes, cie_gdpenmesdes, cie_gdpenmesantdes, cie_sncpenmesdes, cie_sncpenmesantdes) SELECT '" & cencos & "', cie_periodo, cie_fecini, cie_fecter, cie_estado, cie_proantali, cie_gdpenmesali, cie_gdpenmesantali, cie_sncpenmesali, cie_sncpenmesantali, cie_proantgrl, cie_gdpenmesgrl, cie_gdpenmesantgrl, cie_sncpenmesgrl, cie_sncpenmesantgrl, cie_proantdes, cie_gdpenmesdes, cie_gdpenmesantdes, cie_sncpenmesdes, cie_sncpenmesantdes FROM PASO"
   vg_db.Execute "DROP TABLE PASO"

   '------- Modificar tabla minuta costo
   RS1.Open "SELECT * INTO PASO FROM b_minutacosto", vg_db, adOpenStatic
   Set RS1 = Nothing
   vg_db.Execute "DROP TABLE b_minutacosto"
   vg_db.Execute "CREATE TABLE b_minutacosto (mic_cencos char(10), mic_fecval int, mic_tipmin char(1), mic_codpro char(20), mic_cospro double, Constraint b_minutacosto_pk Primary Key (mic_cencos, mic_fecval, mic_tipmin, mic_codpro))"
   '------- Insertar datos tabla calendario cierre de mes
   vg_db.Execute "INSERT INTO b_minutacosto (mic_cencos, mic_fecval, mic_tipmin, mic_codpro, mic_cospro) SELECT '" & cencos & "', mic_fecval, mic_tipmin, mic_codpro, mic_cospro FROM PASO"
   vg_db.Execute "DROP TABLE PASO"
   
   '------- Crear tabla contrato lista precio producto
   vg_db.Execute "CREATE TABLE b_contlistprepro (cpp_cencos char(10), cpp_codpro char(20), cpp_upreco double, cpp_fecuco date, cpp_propon double, Constraint b_contlistprepro_pk Primary Key (cpp_cencos, cpp_codpro))"
   '------- Crear relación tabla lista precio producto vs ingredientes
   vg_db.Execute "ALTER TABLE b_contlistprepro ADD CONSTRAINT FK_b_contlistprepro_b_productos FOREIGN KEY (cpp_codpro) REFERENCES b_productos (pro_codigo)"
   '------- Crear relación tabla lista precio producto vs b_clientes
   vg_db.Execute "ALTER TABLE b_contlistprepro ADD CONSTRAINT FK_b_contlistprepro_b_clientes FOREIGN KEY (cpp_cencos) REFERENCES b_clientes (cli_codigo)"
   
   '------- Mover datos
   vg_db.Execute "INSERT INTO b_contlistprepro (cpp_cencos, cpp_codpro, cpp_upreco, cpp_fecuco, cpp_propon) SELECT '" & cencos & "', pro_codigo, pro_upreco, pro_fecuco, pro_propon FROM b_productos"
   
   '------- Crear tabla contrato lista precio ingrediente
   vg_db.Execute "CREATE TABLE b_contlistpreing (cpi_cencos char(10), cpi_coding char(20), cpi_precos double, cpi_feccos int, cpi_codcom char(20), cpi_codped char(20), Constraint b_contlistpreing_pk Primary Key (cpi_cencos, cpi_coding))"
   '------- Crear relación tabla lista precio ingrediente vs ingredientes
   vg_db.Execute "ALTER TABLE b_contlistpreing ADD CONSTRAINT FK_b_contlistpreing_b_ingrediente FOREIGN KEY (cpi_coding) REFERENCES b_ingrediente (ing_codigo)"
   '------- Crear relación tabla lista precio ingrediente vs b_clientes
   vg_db.Execute "ALTER TABLE b_contlistpreing ADD CONSTRAINT FK_b_contlistpreing_b_clientes FOREIGN KEY (cpi_cencos) REFERENCES b_clientes (cli_codigo)"
   '------- Mover datos
   vg_db.Execute "INSERT INTO b_contlistpreing (cpi_cencos, cpi_coding, cpi_precos, cpi_feccos, cpi_codcom, cpi_codped) SELECT '" & cencos & "', ing_codigo, ing_precos, ing_feccos, ing_codcom, ing_codped FROM b_ingrediente"
   
   '------- Generar tabla tipo servicio
   vg_db.Execute "CREATE TABLE a_tiposervicio (tis_codigo int, tis_nombre char(50), Constraint a_tiposervicio_pk Primary Key (tis_codigo))"
   '------- Generar tabla segmento
   vg_db.Execute "CREATE TABLE a_segmento (seg_codigo int, seg_nombre char(50), Constraint a_segmento_pk Primary Key (seg_codigo))"
   
   '------- Insertar opción tipo servicio
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4084000, 'Tipo de Servicio')"
   '------- Insertar opción segmento
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4086000, 'Segmento')"
   '------- Insertar opción cambio de contrato
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4011000, 'Cambio de Contrato')"
   
   '------- Borrar clave principal y incluir un nuevo campo contrato detalle receta, volver a crear la clave principal incluyendo el nuevo campo,
   vg_db.Execute "DROP INDEX b_recetadet_pk ON b_recetadet"
   vg_db.Execute "ALTER TABLE b_recetadet ADD COLUMN red_cencos char(10)"
   vg_db.Execute "UPDATE b_recetadet SET red_cencos=IIf(red_tiprec <> 0, '" & cencos & "', '0')"
   vg_db.Execute "ALTER TABLE b_recetadet ADD Constraint b_recetadet_pk Primary Key (red_codigo, red_nroite, red_tiprec, red_cencos)"


   '------- Borrar clave principal y incluir un nuevo campo contrato lista precio cafeteria
   vg_db.Execute "ALTER TABLE b_detpreciocaf DROP CONSTRAINT FK_b_detpreciocaf_b_totpreciocaf"
   vg_db.Execute "ALTER TABLE b_detpreciocaf DROP CONSTRAINT FK_b_detpreciocaf_b_productos"
   vg_db.Execute "ALTER TABLE b_detventascaf DROP CONSTRAINT FK_b_detventascaf_b_totpreciocaf"
   
   vg_db.Execute "DROP INDEX b_detpreciocaf_pk ON b_detpreciocaf"
   vg_db.Execute "ALTER TABLE b_detpreciocaf ADD COLUMN dpc_cencos char(10)"
   vg_db.Execute "UPDATE b_detpreciocaf SET dpc_cencos='" & cencos & "'"
   vg_db.Execute "ALTER TABLE b_detpreciocaf ADD Constraint b_detpreciocaf_pk Primary Key (dpc_cencos, dpc_codigo, dpc_codmer)"
   
   vg_db.Execute "DROP INDEX b_totpreciocaf_pk ON b_totpreciocaf"
   vg_db.Execute "ALTER TABLE b_totpreciocaf ADD COLUMN tpc_cencos char(10)"
   vg_db.Execute "UPDATE b_totpreciocaf SET tpc_cencos='" & cencos & "'"
   vg_db.Execute "ALTER TABLE b_totpreciocaf ADD Constraint b_totpreciocaf_pk Primary Key (tpc_cencos, tpc_codigo)"

   vg_db.Execute "ALTER TABLE b_detpreciocaf ADD CONSTRAINT FK_b_detpreciocaf_b_totpreciocaf FOREIGN KEY (dpc_cencos, dpc_codigo) REFERENCES b_totpreciocaf (tpc_cencos, tpc_codigo)"
   vg_db.Execute "ALTER TABLE b_detpreciocaf ADD CONSTRAINT FK_b_detpreciocaf_b_productos FOREIGN KEY (dpc_codmer) REFERENCES b_productos (pro_codigo)"
   vg_db.Execute "ALTER TABLE b_detventascaf ADD CONSTRAINT FK_b_detventascaf_b_totpreciocaf FOREIGN KEY (dvc_cencos, dvc_articulo) REFERENCES b_totpreciocaf (tpc_cencos, tpc_codigo)"
    
   
   '------- Borrar clave principal y incluir un nuevo campo contrato parametro
   vg_db.Execute "DROP INDEX PrimaryKey ON a_param"
   vg_db.Execute "ALTER TABLE a_param ADD COLUMN par_cencos char(10)"
   vg_db.Execute "UPDATE a_param SET par_cencos='" & cencos & "'"
   vg_db.Execute "ALTER TABLE a_param ADD Constraint a_param_pk Primary Key (par_cencos, par_codigo)"
   
   '------- Borrar clave principal y incluir un nuevo campo contrato parametro despacho
   vg_db.Execute "DROP INDEX b_paramdesp_pk ON b_paramdesp"
   vg_db.Execute "ALTER TABLE b_paramdesp ADD COLUMN pad_cencos char(10)"
   vg_db.Execute "UPDATE b_paramdesp SET pad_cencos='" & cencos & "'"
   vg_db.Execute "ALTER TABLE b_paramdesp ADD Constraint b_paramdesp_pk Primary Key (pad_cencos, pad_codigo)"

   '------- Borrar clave principal y incluir un nuevo campo contrato servicio raciones
   vg_db.Execute "DROP INDEX a_serviciorac_pk ON a_serviciorac"
   vg_db.Execute "ALTER TABLE a_serviciorac ADD COLUMN sra_cencos char(10)"
   vg_db.Execute "UPDATE a_serviciorac SET sra_cencos='" & cencos & "'"
   vg_db.Execute "ALTER TABLE a_serviciorac ADD Constraint a_serviciorac_pk Primary Key (sra_cencos, sra_codser, sra_coditem, sra_serdia)"

   '------- Borrar clave principal y incluir un nuevo campo contrato estructura servicio
   vg_db.Execute "DROP INDEX PrimaryKey ON a_estservicio"
   vg_db.Execute "ALTER TABLE a_estservicio ADD COLUMN ess_cencos char(10)"
   vg_db.Execute "UPDATE a_estservicio SET ess_cencos='" & cencos & "'"
   vg_db.Execute "ALTER TABLE a_estservicio ADD Constraint a_estservicio_pk Primary Key (ess_cencos, ess_codser, ess_codigo)"

   '------- Asignar relación a la tabla bodega vs ventas - compras - bodegas - venta cafeteria - toma inventario
   vg_db.Execute "ALTER TABLE b_totventas ADD CONSTRAINT FK_b_totventas_a_bodega FOREIGN KEY (tov_codbod) REFERENCES a_bodega (bod_codigo)"
   vg_db.Execute "ALTER TABLE b_totcompras ADD CONSTRAINT FK_b_totcompras_a_bodega FOREIGN KEY (toc_codbod) REFERENCES a_bodega (bod_codigo)"
   vg_db.Execute "ALTER TABLE b_bodegas ADD CONSTRAINT FK_b_bodegas_a_bodega FOREIGN KEY (bod_codbod) REFERENCES a_bodega (bod_codigo)"
   vg_db.Execute "ALTER TABLE b_totventascaf ADD CONSTRAINT FK_b_totventascaf_a_bodega FOREIGN KEY (tvc_codbod) REFERENCES a_bodega (bod_codigo)"
   vg_db.Execute "ALTER TABLE b_tomainv ADD CONSTRAINT FK_b_tomainv_a_bodega FOREIGN KEY (tin_codbod) REFERENCES a_bodega (bod_codigo)"
   
   vg_db.Execute "UPDATE a_param SET par_valor='150' WHERE par_codigo='version'"
   aVer = 150
End If
If nVer > aVer And aVer = 150 Then
   '------- Consultar si existe codigo bodega
   RS1.Open "SELECT DISTINCT * FROM a_bodega ORDER BY bod_codigo", vg_db, adOpenStatic
   If RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "No existe bodega, proceso cancelado " & VgLinea & "Comunicase con departamento de informatica" & VgLinea & "        Proceso cancelado ...", vbCritical + vbOKOnly, "SGP": Exit Function
   codbod = RS1!bod_codigo
   RS1.Close: Set RS1 = Nothing
   
   '------- Crear Tabla correlativo
   vg_db.Execute "CREATE TABLE b_parametros (par_tipdoc char(2), par_codbod int, par_correlativo int, Constraint b_parametros_pk Primary Key (par_tipdoc, par_codbod))"
   vg_db.Execute "INSERT INTO b_parametros VALUES ('AI', " & codbod & ",0)"
   vg_db.Execute "INSERT INTO b_parametros VALUES ('SP', " & codbod & ",0)"
   vg_db.Execute "INSERT INTO b_parametros VALUES ('DP', " & codbod & ",0)"
   vg_db.Execute "INSERT INTO b_parametros VALUES ('ME', " & codbod & ",0)"
   '------- Traer ultimo correlativo tipo documento salida producción
   RS1.Open "SELECT tov_numdoc FROM b_totventas WHERE tov_tipdoc='SP' AND tov_codbod=" & codbod & " ORDER BY tov_numdoc DESC", vg_db, adOpenStatic
   If Not RS1.EOF Then
      RS1.MoveFirst
      vg_db.Execute "UPDATE b_parametros SET par_correlativo=" & RS1!tov_numdoc & " WHERE par_codbod=" & codbod & " AND par_tipdoc='SP'"
   End If
   RS1.Close: Set RS1 = Nothing
   
   '------- Traer ultimo correlativo tipo documento ajuste inventario
   RS1.Open "SELECT tov_numdoc FROM b_totventas WHERE tov_tipdoc='AI' AND tov_codbod=" & codbod & " ORDER BY tov_numdoc DESC", vg_db, adOpenStatic
   If Not RS1.EOF Then
      RS1.MoveFirst
      vg_db.Execute "UPDATE b_parametros SET par_correlativo=" & RS1!tov_numdoc & " WHERE par_codbod=" & codbod & " AND par_tipdoc='AI'"
   End If
   RS1.Close: Set RS1 = Nothing
   
   '------- Traer ultimo correlativo tipo documento devolucion producción
   RS1.Open "SELECT tov_numdoc FROM b_totventas WHERE tov_tipdoc='DP' AND tov_codbod=" & codbod & " ORDER BY tov_numdoc DESC", vg_db, adOpenStatic
   If Not RS1.EOF Then
      RS1.MoveFirst
      vg_db.Execute "UPDATE b_parametros SET par_correlativo=" & RS1!tov_numdoc & " WHERE par_codbod=" & codbod & " AND par_tipdoc='DP'"
   End If
   RS1.Close: Set RS1 = Nothing
   
   '------- Traer ultimo correlativo tipo documento mermas
   RS1.Open "SELECT tov_numdoc FROM b_totventas WHERE tov_tipdoc='ME' AND tov_codbod=" & codbod & " ORDER BY tov_numdoc DESC", vg_db, adOpenStatic
   If Not RS1.EOF Then
      RS1.MoveFirst
      vg_db.Execute "UPDATE b_parametros SET par_correlativo=" & RS1!tov_numdoc & " WHERE par_codbod=" & codbod & " AND par_tipdoc='ME'"
   End If
   RS1.Close: Set RS1 = Nothing
   vg_db.Execute "UPDATE a_param SET par_valor='151' WHERE par_codigo='version'"
   aVer = 151
End If
If nVer > aVer And aVer = 151 Then
   '------- Insertar campo a la tabla productos maestro producto
   vg_db.Execute "ALTER TABLE b_productos ADD COLUMN pro_maepro int"
   vg_db.Execute "UPDATE b_productos SET pro_maepro=1"
   vg_db.Execute "UPDATE b_productos SET pro_maepro=0 WHERE pro_ctacon NOT IN ('410001','410004')"
   '------- Insertar dato a tabla tipo servicio
   vg_db.Execute "INSERT INTO a_tiposervicio VALUES (1, 'Alimentación')"
   vg_db.Execute "UPDATE b_clientes SET cli_codtis=1 WHERE cli_codbod>0 AND cli_tipo=0"
   vg_db.Execute "UPDATE a_param SET par_valor='152' WHERE par_codigo='version'"
   aVer = 152
End If
If nVer > aVer And aVer = 152 Then
   '-------> Incluir campo tabla b_totpreciocaf activo ó desactivo
   vg_db.Execute "ALTER TABLE b_totpreciocaf ADD COLUMN tpc_activo char(1)"
   vg_db.Execute "UPDATE b_totpreciocaf SET tpc_activo='1'"
   '-------> Incluir campo tabla b_minutadet costo desechable
   vg_db.Execute "ALTER TABLE b_minutadet ADD COLUMN mid_cosdes double"
   vg_db.Execute "UPDATE b_minutadet SET mid_cosdes=0"
   '-------> Incluir campo en tabla b_totcompras indicando si el documento fue enviado a sap
   vg_db.Execute "ALTER TABLE b_totcompras ADD COLUMN toc_envsap char(1)"
   vg_db.Execute "UPDATE b_totcompras SET toc_envsap='0'" '0= corresponde a no enviado, 1=corresponde a enviado
   '-------> Incluir campo en tabla b_tomainv indicando si el inventario fue enviado a sap
   vg_db.Execute "ALTER TABLE b_tomainv ADD COLUMN tin_envsap char(1)"
   vg_db.Execute "UPDATE b_tomainv SET tin_envsap='0'" '0= corresponde a no enviado, 1=corresponde a enviado
   '-------> Incluir tabla cfc - inventario - guia ventas
   vg_db.Execute "CREATE TABLE sap_cfc (cfc_codigo int, cfc_numlin int, cfc_nuedoc char(1), cfc_socied char(4), cfc_cladoc char(2), cfc_feccon char(8), cfc_fecdoc char(8), cfc_refere char(16), cfc_texcab char(25), cfc_mondoc char(5), cfc_clacon char(2), cfc_cueaux char(10), cfc_mtodoc char(11), cfc_asigna char(18), cfc_glosa char(50), cfc_ccosto char(10), cfc_codimp char(4), cfc_ctaimp char(10), cfc_monimp char(11), cfc_imprec char(2), cfc_otrimp char(2), Constraint sap_cfc_pk Primary Key (cfc_codigo, cfc_numlin))"
   vg_db.Execute "CREATE TABLE sap_inv (inv_codigo int, inv_numlin int, inv_nuedoc char(1), inv_socied char(4), inv_cladoc char(2), inv_feccon char(8), inv_fecdoc char(8), inv_refere char(16), inv_texcab char(25), inv_mondoc char(5), inv_clacon char(2), inv_cueaux char(10), inv_mtodoc char(11), inv_asigna char(18), inv_glosa char(50), inv_ccosto char(10), inv_codimp char(4), inv_ctaimp char(10), inv_monimp char(11), inv_imprec char(2), inv_otrimp char(2), Constraint sap_inv_pk Primary Key (inv_codigo, inv_numlin))"
   vg_db.Execute "CREATE TABLE sap_guiavta (gvt_codigo int, gvt_numlin int, gvt_pedvta char(10), gvt_codcli char(10), gvt_desmer char(10), gvt_censum char(10), gvt_fecent char(10), gvt_codmat char(18), gvt_desmat char(40), gvt_cantidad char(15), gvt_prevta char(13), gvt_tipmon char(5), gvt_glosa1 char(40), gvt_glosa2 char(40), gvt_glosa3 char(40), Constraint sap_guiavta_pk Primary Key (gvt_codigo, gvt_numlin))"
   '-------> Incluir opción de generar guía venta SAP
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (2140000, 'Generar Guía Venta SAP')"
   '-------> Crear nuevos campos a la tabla minutaraciones
   vg_db.Execute "ALTER TABLE b_minutaraciones ADD COLUMN mir_nroguia int, mir_codcli char(10)"
   '-------> Craer tabla guias de ventas sap
   vg_db.Execute "CREATE TABLE b_totguiavta (tgv_rutcli char(10), tgv_codsuc char(20), tgv_numdoc int, tgv_fecing date, tgv_glosa1 char(40), tgv_glosa2 char(40), tgv_glosa3 char(40), tgv_perfac int, tgv_fecini date, tgv_fecfin date, tgv_envsap char(1), Constraint b_totguiavta_pk Primary Key (tgv_rutcli, tgv_codsuc, tgv_numdoc))"
   vg_db.Execute "CREATE TABLE b_detguiavta (dgv_rutcli char(10), dgv_codsuc char(20), dgv_numdoc int, dgv_numlin int, dgv_codreg int, dgv_codser int, dgv_nomser char(40), dgv_desser char(40), dgv_codsap char(20), dgv_racsgp int, dgv_racguia int, dgv_presgp double, dgv_preguia double, dgv_codcli char(10), Constraint b_detguiavta_pk Primary Key (dgv_rutcli, dgv_codsuc, dgv_numdoc, dgv_numlin))"
   vg_db.Execute "ALTER TABLE b_detguiavta ADD CONSTRAINT FK_b_detguiavta_b_totguiavta FOREIGN KEY (dgv_rutcli, dgv_codsuc, dgv_numdoc) REFERENCES b_totguiavta (tgv_rutcli, tgv_codsuc, tgv_numdoc)"
   vg_db.Execute "ALTER TABLE b_detguiavta ADD CONSTRAINT FK_b_detguiavta_a_regimen FOREIGN KEY (dgv_codreg) REFERENCES a_regimen (reg_codigo)"
   vg_db.Execute "ALTER TABLE b_detguiavta ADD CONSTRAINT FK_b_detguiavta_a_servicio FOREIGN KEY (dgv_codser) REFERENCES a_servicio (ser_codigo)"
   '-------> Insertar campo código sap y facturable tabla servicio
   vg_db.Execute "ALTER TABLE b_detguiavta DROP CONSTRAINT FK_b_detguiavta_a_servicio"
   vg_db.Execute "ALTER TABLE a_estservicio DROP CONSTRAINT a_servicioa_estservicio"
   vg_db.Execute "ALTER TABLE b_minuta DROP CONSTRAINT a_serviciob_minuta"
   vg_db.Execute "ALTER TABLE b_tomapedidodet DROP CONSTRAINT FK_b_tomapedidodet_b_servicio"
   RS1.Open "SELECT * INTO PASO FROM a_servicio", vg_db, adOpenStatic
   Set RS1 = Nothing
'   vg_db.Execute "DROP TABLE a_servicio"
'   vg_db.Execute "CREATE TABLE a_servicio (ser_codigo int, ser_nombre char(30), ser_orden int, ser_codsap char(20), ser_facturable char(1), ser_horcob date, ser_horent date, ser_horpda date, Constraint a_servicio_pk Primary Key (ser_codigo))"
   
'   Set wksPredef = DBEngine.Workspaces(0)
'   Set dbsRubrica = wksPredef.OpenDatabase(dir_trabajo & BaseDeDato, , , dbLangGeneral)
'   Set tdfRubrica = dbsRubrica.TableDefs("a_servicio")
'   tdfRubrica.Fields("ser_codigo").DefaultValue = 0
'   tdfRubrica.Fields("ser_nombre").AllowZeroLength = False
'   dbsRubrica.Close: Set dbsRubrica = Nothing
'   wksPredef.Close: Set wksPredef = Nothing
'   Set tdfRubrica = Nothing
   vg_db.Execute "ALTER TABLE a_servicio DROP COLUMN ser_horcob, ser_horent, ser_horpda"
   vg_db.Execute "ALTER TABLE a_servicio ADD COLUMN ser_codsap char(20), ser_facturable char(1), ser_activo char(1), ser_horcob date, ser_horent date, ser_horpda date"
   vg_db.Execute "UPDATE a_servicio SET ser_activo='1'"
   RS1.Open "SELECT * FROM PASO", vg_db, adOpenStatic
   Do While Not RS1.EOF
      horcob = IIf(IsNull(RS1!ser_horcob), "00:00", RS1!ser_horcob)
      horent = IIf(IsNull(RS1!ser_horent), "00:00", RS1!ser_horent)
      horpda = IIf(IsNull(RS1!ser_horpda), "00:00", RS1!ser_horpda)
'      vg_db.Execute "INSERT INTO a_servicio (ser_codigo, ser_nombre, ser_orden, ser_codsap, ser_facturable, ser_horcob, ser_horent, ser_horpda) " & _
'                    "VALUES (" & RS1!ser_codigo & ", '" & RS1!ser_nombre & "', " & RS1!ser_orden & ", null, null, '" & horcob & "', '" & horent & "', '" & horpda & "')"
      If IsNull(RS1!ser_horcob) Or Trim(RS1!ser_horcob) = "" Then vg_db.Execute "UPDATE a_servicio SET ser_horcob=null WHERE ser_codigo=" & RS1!ser_codigo & ""
      If IsNull(RS1!ser_horent) Or Trim(RS1!ser_horent) = "" Then vg_db.Execute "UPDATE a_servicio SET ser_horent=null WHERE ser_codigo=" & RS1!ser_codigo & ""
      If IsNull(RS1!ser_horpda) Or Trim(RS1!ser_horpda) = "" Then vg_db.Execute "UPDATE a_servicio SET ser_horpda=null WHERE ser_codigo=" & RS1!ser_codigo & ""
      RS1.MoveNext
   Loop
   RS1.Close: Set RS1 = Nothing
   vg_db.Execute "DROP TABLE PASO"
   vg_db.Execute "ALTER TABLE a_estservicio ADD CONSTRAINT a_servicioa_estservicio FOREIGN KEY (ess_codser) REFERENCES a_servicio (ser_codigo)"
   vg_db.Execute "ALTER TABLE b_minuta ADD CONSTRAINT a_serviciob_minuta FOREIGN KEY (min_codser) REFERENCES a_servicio (ser_codigo)"
   vg_db.Execute "ALTER TABLE b_tomapedidodet ADD CONSTRAINT FK_b_tomapedidodet_a_servicio FOREIGN KEY (tpd_codser) REFERENCES a_servicio (ser_codigo)"
   vg_db.Execute "ALTER TABLE b_detguiavta ADD CONSTRAINT FK_b_detguiavta_a_servicio FOREIGN KEY (dgv_codser) REFERENCES a_servicio (ser_codigo)"
   '-------> Crear tabla sucursal sap
   vg_db.Execute "CREATE TABLE b_sucursalcliente (scl_codigo char(20), scl_codcli char(10), scl_direccion char(50), Constraint log_procesos_pk Primary Key (scl_codigo, scl_codcli))"
   vg_db.Execute "ALTER TABLE b_sucursalcliente ADD CONSTRAINT FK_b_sucursalcliente_b_clientes FOREIGN KEY (scl_codcli) REFERENCES b_clientes (cli_codigo)"
   '-------> Insertar campo a tabla cliente si corresponde cliente sap
   vg_db.Execute "ALTER TABLE b_clientes ADD cli_codcli char(10), cli_clisap char(1)"
   '-------> Crear log proceso
   vg_db.Execute "CREATE TABLE log_procesos (cencos char(10), numero int, fecha date, tipo_proceso char(1), rut char(10), tipo_documento char(2), num_documento char(10), num_cfc int, estado char(1), mensaje memo, envio int, anulado char(1), Constraint log_procesos_pk Primary Key (cencos, numero, fecha, tipo_proceso))"
   '-------> Insertar campo impuesto adicional
   vg_db.Execute "ALTER TABLE a_impuesto ADD imp_adicional int"
   vg_db.Execute "UPDATE a_impuesto SET imp_adicional='0' WHERE imp_codigo IN (1, 11)"
   vg_db.Execute "UPDATE a_impuesto SET imp_adicional='1' WHERE imp_codigo NOT IN (1, 11)"
   '-------> Modificar campos a tabla usuario contratos
   vg_db.Execute "ALTER TABLE b_usuariocontratos DROP CONSTRAINT FK_b_usuariocontratos_a_usuarios"
   vg_db.Execute "ALTER TABLE b_usuariocontratos DROP CONSTRAINT FK_b_usuariocontratos_b_clientes"
   vg_db.Execute "ALTER TABLE b_usuariocontratos ALTER COLUMN uco_codusu char(20)"
   vg_db.Execute "ALTER TABLE b_usuariocontratos ALTER COLUMN uco_codcon char(10)"
   vg_db.Execute "ALTER TABLE b_usuariocontratos ADD CONSTRAINT FK_b_usuariocontratos_a_usuarios FOREIGN KEY (uco_codusu) REFERENCES a_usuarios (usu_codigo)"
   vg_db.Execute "ALTER TABLE b_usuariocontratos ADD CONSTRAINT FK_b_usuariocontratos_b_clientes FOREIGN KEY (uco_codcon) REFERENCES b_clientes (cli_codigo)"
   '-------> Incluir opción de enviar toma inventario
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (2079100, 'Enviar Inventario SAP')"
   '-------> Incluir opción de autorización de ajuste
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (2079010, 'Autorización Ajuste')"
   '-------> Incluir campo en tabla b_tomainv indicando si esta autorizado ajuste
   vg_db.Execute "ALTER TABLE b_tomainv ADD COLUMN tin_autaju char(1)"
   vg_db.Execute "UPDATE b_tomainv SET tin_autaju='0'" '0= corresponde no autorizado ajuste, 1=corresponde autorizado ajuste
   '-------> Incluir concepto sociedad sap y envio sap documento a tabla centro de costo
   vg_db.Execute "ALTER TABLE b_clientes ADD COLUMN cli_socsap char(4), cli_envsap char(1)"
   vg_db.Execute "UPDATE b_clientes SET cli_envsap='0'" '0= corresponde no envío, 1=corresponde envío
   '-------> Crear tabla tipo documento
   vg_db.Execute "CREATE TABLE a_tipodocumento (tdo_codigo char(2), tdo_nombre char(50), tdo_cladoc char(2), tdo_orden int, Constraint b_tipodocumento_pk Primary Key (tdo_codigo))"
   '-------> Incluir campo tabla b_clientes concepto cierre de ventas y activo
   vg_db.Execute "ALTER TABLE b_clientes ADD COLUMN cli_cievta char(1), cli_ciedia int, cli_activo char(1)"
   vg_db.Execute "UPDATE b_clientes SET cli_cievta='1', cli_ciedia=0, cli_activo='1'"
   '-------> Ingresar dato tabla tipo documento
   vg_db.Execute "INSERT INTO a_tipodocumento VALUES ('BH', 'BOLETA DE HONORARIOS', '',9)"
   vg_db.Execute "INSERT INTO a_tipodocumento VALUES ('BO', 'BOLETA', '',8)"
   vg_db.Execute "INSERT INTO a_tipodocumento VALUES ('CE', 'NOTA DE CREDITO ELECTRONICA', 'CE', 5)"
   vg_db.Execute "INSERT INTO a_tipodocumento VALUES ('CG', 'COMPROBANTE DE GASTO', '', 10)"
   vg_db.Execute "INSERT INTO a_tipodocumento VALUES ('DE', 'NOTA DE DEBITO ELECTRONICA', 'DE', 7)"
   vg_db.Execute "INSERT INTO a_tipodocumento VALUES ('FA', 'FACTURA', 'FA', 1)"
   vg_db.Execute "INSERT INTO a_tipodocumento VALUES ('FE', 'FACTURA ELECTRONICA', 'FE', 2)"
   vg_db.Execute "INSERT INTO a_tipodocumento VALUES ('GD', 'GUIA DESPACHO', '', 3)"
   vg_db.Execute "INSERT INTO a_tipodocumento VALUES ('NC', 'NOTA DE CREDITO', 'NC', 4)"
   vg_db.Execute "INSERT INTO a_tipodocumento VALUES ('ND', 'NOTA DE DEBITO', 'ND', 6)"
   '-------> Mover enviado a tabla toma inventario
   vg_db.Execute "UPDATE b_tomainv INNER JOIN b_cierreperiodo ON val(mid(b_tomainv.tin_fectom,1,6))=b_cierreperiodo.cie_periodo SET b_tomainv.tin_envsap = '1',  b_tomainv.tin_autaju='1' WHERE b_cierreperiodo.cie_estado=0"
   RS1.Open "SELECT DISTINCT a.tin_fectom, a.tin_codbod FROM b_tomainv a, b_cierreperiodo b WHERE val(mid(a.tin_fectom,1,6))=b.cie_periodo AND b.cie_estado=1", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         If CierrePeriodo(RS1!tin_fectom, RS1!tin_codbod, 2) Then
            vg_db.Execute "UPDATE b_tomainv SET tin_envsap = '1', tin_autaju='1' WHERE tin_fectom=" & RS1!tin_fectom & ""
         End If
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   
   vg_db.Execute "UPDATE a_param SET par_valor='153' WHERE par_codigo='version'"
   aVer = 153
End If
If nVer > aVer And aVer = 153 Then
   '------- Incluir opción informe Costo Real del Periodo
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4152000, 'Limpiar Base de Dato')"
   RS1.Open "SELECT * FROM b_clientes WHERE cli_tipo=0 AND cli_codbod>0", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         '------- Incluir opción parametro login limpia base de dato
         vg_db.Execute "INSERT INTO a_param VALUES ('usulimbas', 'Login Limpia Base de Dato', 'C', 'MONITOR', '" & RS1!cli_codigo & "')"
         '------- Incluir opción parametro password limpia base de dato
         vg_db.Execute "INSERT INTO a_param VALUES ('paslimbas', 'Password Limpia Base de Dato', 'C', '" & fg_Encripta("sdxo2008*") & "', '" & RS1!cli_codigo & "')"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   
   vg_db.Execute "UPDATE a_param SET par_valor='154' WHERE par_codigo='version'"
   aVer = 154
End If
If nVer > aVer And aVer = 154 Then
   '------->  Generar tabla ABC
   vg_db.Execute "CREATE TABLE a_curvaabc (abc_codigo char(1), abc_nombre char(30), abc_porce int, Constraint a_curvaabc_pk Primary Key (abc_codigo))"
   vg_db.Execute "INSERT INTO a_curvaabc VALUES ('A', 'Curva A', 70)"
   vg_db.Execute "INSERT INTO a_curvaabc VALUES ('B', 'Curva B', 20)"
   vg_db.Execute "INSERT INTO a_curvaabc VALUES ('C', 'Curva C', 10)"
'   '-------> Generar tabla envio costo planif. teorico- planif. real-realizado
'   vg_db.Execute "CREATE TABLE p_costrr (trr_usuario char(20), trr_cencos char(10), trr_codreg int, trr_codser int, trr_fecmin date, trr_nomcen char(50), trr_nomreg char(50), trr_nomser char(50), trr_cospis double, trr_costec double, trr_tenura float, trr_tecoin float, trr_tecode float, trr_renura float, trr_recoin float, trr_recode float, trr_rdnura float, trr_rdcoin float, trr_rdcode float, Constraint p_costrr_pk Primary Key (trr_usuario, trr_cencos, trr_codreg, trr_codser, trr_fecmin))"
   '-------> Incluir opción informe Costo Real del Periodo
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3024000, 'Costo Detalle Periodo Realizado')"
   '-------> Incluir opción informe Curva ABC
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3026000, 'Curva ABC')"
   '-------> Incluir opción tabla Curva ABC
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (4142000, 'Curva ABC')"
   '-------> Incluir opción informe inflaciòn interna
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3028000, 'Inflación Interna')"
   '-------> Incluir opción informe producto sin movimiento
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3048000, 'Producto Sin Movimiento')"
   vg_db.Execute "UPDATE a_param SET par_valor='155' WHERE par_codigo='version'"
   aVer = 155
End If
If nVer > aVer And aVer = 155 Then
   '-------> Incluir opción informe inflaciòn interna
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3036000, 'Comparativo de Raciones')"
   '-------> Incluir parametro cierre solicitud nota de credito
   RS1.Open "SELECT * FROM b_clientes WHERE cli_tipo=0 AND cli_codbod>0", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         vg_db.Execute "INSERT INTO a_param VALUES ('parsn', 'Parametro Solicitud Nota Credito', 'N', '6', '" & RS1!cli_codigo & "')"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   
   vg_db.Execute "UPDATE a_param SET par_valor='156' WHERE par_codigo='version'"
   aVer = 156
End If
If nVer > aVer And aVer = 156 Then
   Dim hora As String
   '-------> Insertar campo a la tabla b_totcompras
   vg_db.Execute "ALTER TABLE b_totcompras ADD COLUMN toc_fecdig DATE, toc_fecper INT"
   RS1.Open "SELECT * FROM b_cierreperiodo WHERE cie_estado=0 OR cie_estado=1 ORDER BY cie_cencos, cie_periodo", vg_db, adOpenStatic
   hora = " 00:00:00"
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         vg_db.Execute "UPDATE b_totcompras SET toc_fecdig=toc_fecemi & '" & hora & "', toc_fecper=" & RS1!cie_periodo & " WHERE toc_fecemi>=CDATE('" & fg_Ctod1(RS1!cie_fecini) & "') AND toc_fecemi<=CDATE('" & fg_Ctod1(RS1!cie_fecter) & "') AND toc_codbod IN (SELECT DISTINCT cli_codbod FROM b_clientes WHERE cli_codigo='" & RS1!cie_cencos & "' AND cli_tipo=0)"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   '-------> Actualiza versión
   vg_db.Execute "UPDATE a_param SET par_valor='157' WHERE par_codigo='version'"
   aVer = 157
End If
If nVer > aVer And aVer = 157 Then
   V_Acceso.Label1(0).Visible = True: V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un momento Actua.:"
   V_Acceso.Label1(0).Caption = "  Versión 158"
   '-------> Create clave secundaria tabla encabezado minuta
   vg_db.Execute "CREATE INDEX SecondKey ON b_minuta (min_cencos, min_fecmin)"
   '-------> Create clave secundaria tabla encabezado documento proveedor
   vg_db.Execute "CREATE INDEX SecondKey ON b_totcompras (toc_fecper)"
   '-------> Create clave secundaria tabla detalle minuta
   vg_db.Execute "CREATE INDEX SecondKey ON b_minutadet (mid_codrec)"
   '-------> Create clave secundaria tabla productos ingredientes
   vg_db.Execute "CREATE INDEX SecondKey ON b_contlistpreing (cpi_codcom, cpi_codped)"
   '-------> Actualizar periodo
   RS1.Open "SELECT * FROM b_cierreperiodo WHERE cie_estado=0 OR cie_estado=1 ORDER BY cie_cencos, cie_periodo", vg_db, adOpenStatic
   hora = " 00:00:00"
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         vg_db.Execute "UPDATE b_totcompras SET toc_fecdig=toc_fecemi & '" & hora & "', toc_fecper=" & RS1!cie_periodo & " WHERE toc_fecemi>=CDATE('" & fg_Ctod1(RS1!cie_fecini) & "') AND toc_fecemi<=CDATE('" & fg_Ctod1(RS1!cie_fecter) & "') AND toc_codbod IN (SELECT DISTINCT cli_codbod FROM b_clientes WHERE cli_codigo='" & RS1!cie_cencos & "' AND cli_tipo=0) AND ((toc_fecper) IS NULL OR toc_fecper<0)"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   '-------> Actualizar opciones de sistemas
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2000000 WHERE dpe_codopc=2010000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2010000 WHERE dpe_codopc=2020000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2020000 WHERE dpe_codopc=2030000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2030000 WHERE dpe_codopc=2040000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2240000 WHERE dpe_codopc=2130000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2230000 WHERE dpe_codopc=2120000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2220000 WHERE dpe_codopc=2110000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2210000 WHERE dpe_codopc=2100000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2200000 WHERE dpe_codopc=2090000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2190000 WHERE dpe_codopc=2080000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2080000 WHERE dpe_codopc=2070000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2180000 WHERE dpe_codopc=2079000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2170000 WHERE dpe_codopc=2078000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2160000 WHERE dpe_codopc=2077800"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2140000 WHERE dpe_codopc=2140000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2130000 WHERE dpe_codopc=2077600"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2120000 WHERE dpe_codopc=2077400"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2110000 WHERE dpe_codopc=2077000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2100000 WHERE dpe_codopc=2076000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2090000 WHERE dpe_codopc=2075000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2070000 WHERE dpe_codopc=2055000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2060000 WHERE dpe_codopc=2052000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2050000 WHERE dpe_codopc=2050000"
   vg_db.Execute "UPDATE a_derechosperfil SET dpe_codopc=2040000 WHERE dpe_codopc=2045000"
   
   
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2000000 WHERE opc_codigo=2010000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2010000 WHERE opc_codigo=2020000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2020000 WHERE opc_codigo=2030000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2030000 WHERE opc_codigo=2040000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2240000 WHERE opc_codigo=2130000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2230000 WHERE opc_codigo=2120000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2220000 WHERE opc_codigo=2110000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2210000 WHERE opc_codigo=2100000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2200000 WHERE opc_codigo=2090000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2190000 WHERE opc_codigo=2080000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2080000 WHERE opc_codigo=2070000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2180000 WHERE opc_codigo=2079000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2170000 WHERE opc_codigo=2078000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2160000 WHERE opc_codigo=2077800"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2140000 WHERE opc_codigo=2140000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2130000 WHERE opc_codigo=2077600"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2120000 WHERE opc_codigo=2077400"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2110000 WHERE opc_codigo=2077000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2100000 WHERE opc_codigo=2076000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2090000 WHERE opc_codigo=2075000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2070000 WHERE opc_codigo=2055000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2060000 WHERE opc_codigo=2052000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2050000 WHERE opc_codigo=2050000"
   vg_db.Execute "UPDATE a_opcsistema SET opc_codigo=2040000 WHERE opc_codigo=2045000"
   '-------> Incluir opción inventario cierre diario
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (2150000, 'Cierre Diario')"
   '-------> Crear tabla productos pmp diario
   vg_db.Execute "CREATE TABLE b_productospmpdia (ppd_cencos char(10), ppd_codpro char(20), ppd_fecdia int, ppd_propon float, ppd_saldo float, Constraint b_productospmpdia_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia))"
   vg_db.Execute "ALTER TABLE b_productospmpdia ADD CONSTRAINT FK_b_productospmpdia_b_productos FOREIGN KEY (ppd_codpro) REFERENCES b_productos (pro_codigo)"
   vg_db.Execute "ALTER TABLE b_productospmpdia ADD CONSTRAINT FK_b_productospmpdia_b_clientes FOREIGN KEY (ppd_cencos) REFERENCES b_clientes (cli_codigo)"
   
   '------- Anexar datos a la tabla a_param para manejar cierre dieario
   RS1.Open "SELECT DISTINCT cie_cencos, cie_fecini, cie_fecter FROM b_cierreperiodo WHERE cie_estado=1", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         RS2.Open "SELECT MAX(a.tin_fectom) AS tin_fectom FROM b_tomainv a, b_clientes b WHERE a.tin_codbod=b.cli_codbod AND b.cli_codigo='" & RS1!cie_cencos & "'", vg_db, adOpenStatic
         If Not RS2.EOF And Not IsNull(RS2!tin_fectom) Then
            vg_db.Execute "INSERT INTO a_param  VALUES ('ciediario', 'Ultimo Cierre', 'C', '" & fg_Encripta(LimpiaDato(CDate(fg_Ctod1(CStr(RS2!tin_fectom))) + 1)) & "', '" & Trim(RS1!cie_cencos) & "')"
         Else
            vg_db.Execute "INSERT INTO a_param  VALUES ('ciediario', 'Ultimo Cierre', 'C', '" & fg_Encripta(LimpiaDato(CDate(fg_Ctod1(CStr(RS1!cie_fecter))) + 1)) & "', '" & Trim(RS1!cie_cencos) & "')"
            '-------> Insertar tabla b_productospmpdia
            vg_db.Execute "INSERT INTO b_productospmpdia(ppd_cencos, ppd_codpro, ppd_fecdia, ppd_propon, ppd_saldo) " & _
                          "SELECT DISTINCT '" & Trim(RS1!cie_cencos) & "', a.pro_codigo, " & RS1!cie_fecter & ", 0, 0 " & _
                          "FROM b_productos a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR a.pro_maepro < 1) AND c.cli_codigo = '" & Trim(RS1!cie_cencos) & "' AND (b.tis_codigo = a.pro_maepro OR a.pro_maepro < 1)"
         End If
         RS2.Close: Set RS2 = Nothing
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   
   '-------> Incluir campo tabla b_paramdesp despacho diarios
   vg_db.Execute "ALTER TABLE b_paramdesp ADD COLUMN pad_diario char(7)"
   vg_db.Execute "UPDATE b_paramdesp SET pad_diario = '0000000'"
   
   '-------> Crear tabla log cierrediario
   vg_db.Execute "CREATE TABLE log_cierrediario (fecha date, feccie date, usuario char(20), tipocierre char(255), Constraint log_cierrediario_pk Primary Key (fecha))"
   
   '-------> Incluir campo en tabla b_totcompras indicando si el documento fue enviado a sap
   vg_db.Execute "CREATE TABLE a_tipointerfaz (tii_codigo int, tii_nombre char(50), Constraint a_tipointerfaz_pk Primary Key (tii_codigo))"
   vg_db.Execute "INSERT INTO a_tipointerfaz VALUES (1, 'ENVIA CFC')"
   vg_db.Execute "INSERT INTO a_tipointerfaz VALUES (2, 'ENVIA INVENTARIO')"
   vg_db.Execute "INSERT INTO a_tipointerfaz VALUES (3, 'ENVIA GUIA VENTA')"
   vg_db.Execute "CREATE TABLE b_casinointerfaz (cai_cencos char(10), cai_codtii int, Constraint b_casinointerfaz_pk Primary Key (cai_cencos, cai_codtii))"
   RS1.Open "SELECT cli_codigo FROM b_clientes WHERE cli_tipo=0 AND cli_envsap='1'", vg_db, adOpenForwardOnly
   Do While Not RS1.EOF
      vg_db.Execute "INSERT INTO b_casinointerfaz VALUES ('" & RS1!cli_codigo & "', 1)"
      RS1.MoveNext
   Loop
   RS1.Close: Set RS1 = Nothing
   vg_db.Execute "ALTER TABLE b_clientes DROP COLUMN cli_envsap"
   vg_codbod = 0
   
   '-------> Actualiza versión
   vg_db.Execute "UPDATE a_param SET par_valor='158' WHERE par_codigo='version'"
   aVer = 158
   V_Acceso.Label1(0).Visible = False: V_Acceso.Label1(1).Visible = False
End If
If nVer > aVer And aVer = 158 Then
   Dim FecInv As Long, Fecha As Date
   '-------> Insertar campo a la tabla b_productospmpdia
   vg_db.Execute "ALTER TABLE b_productospmpdia ADD COLUMN ppd_upreco double, ppd_fecuco DATE"
   V_Acceso.Frame1(0).Enabled = False
   V_Acceso.Frame1(1).Enabled = False
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(0).Visible = True
   V_Acceso.Label1(1).Caption = "Procesando : "
   V_Acceso.Label1(0).Caption = ""
   RS1.Open "SELECT DISTINCT a.par_cencos, a.par_valor, b.cli_codbod " & _
            "FROM a_param a, b_clientes b " & _
            "WHERE par_codigo = 'ciediario' AND a.par_cencos = b.cli_codigo", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
'         vg_db.Execute "UPDATE (b_tomainv INNER JOIN b_clientes ON b_tomainv.tin_codbod = b_clientes.cli_codbod) INNER JOIN b_cierreperiodo ON b_clientes.cli_codigo = b_cierreperiodo.cie_cencos SET b_tomainv.tin_ciemes = 0 " & _
'                       "WHERE Val(Mid(b_tomainv.tin_fectom,1,6)) = b_cierreperiodo.cie_periodo AND b_cierreperiodo.cie_estado = 1 AND b_tomainv.tin_fectom<b_cierreperiodo.cie_fecter AND b_clientes.cli_codbod = " & RS1!cli_codbod & " AND b_clientes.cli_codigo = '" & RS1!par_cencos & "'"
         FecInv = 0
         RS2.Open "SELECT MAX(tin_fectom) AS tin_fectom " & _
                  "FROM b_tomainv WHERE tin_ciemes<>0", vg_db, adOpenStatic
         If Not RS2.EOF And Not IsNull(RS2!tin_fectom) Then FecInv = IIf(RS2!tin_fectom > 0, RS2!tin_fectom, 0)
         RS2.Close: Set RS2 = Nothing
         FecInv = 20100331
         If FecInv > 0 Then
            '-------> Borrar tabla b_productospmpdia
'            vg_db.Execute "DELETE b_productospmpdia FROM b_productospmpdia WHERE ppd_cencos = '" & Trim(RS1!par_cencos) & "' AND ppd_codpro IN (SELECT a.pro_codigo FROM b_productos a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR a.pro_maepro < 1) AND c.cli_codigo = '" & RS1!par_cencos & "' AND (b.tis_codigo = a.pro_maepro OR a.pro_maepro < 1))"
            vg_contra = Trim(RS1!par_cencos)
            vg_codbod = RS1!cli_codbod
            vg_DCa = IIf(IsNull(GetParametro("parcandec")), 3, GetParametro("parcandec"))
            If vg_DCa = 0 Then vg_DCa = 3
            Fecha = CDate(fg_Ctod1(FecInv)) '+ 1
            If CDate(fg_Desencripta(TipoDato((RS1!par_valor), ""))) - 1 < Fecha Then
               vg_ciedia = Fecha - 1
               V_Acceso.Label1(1).Visible = True
               V_Acceso.Label1(1).Caption = "Procesando : "
               V_Acceso.Label1(0).Caption = Mid(Fecha - 1, 1, 2) & "/" & Mid(CDate(fg_Desencripta(TipoDato((RS1!par_valor), ""))) - 1, 1, 2)
               V_Acceso.Combo1(0).Clear
               V_Acceso.Combo1(0).AddItem "Proc. Día : " & Fecha - 1
               V_Acceso.Combo1(0).ListIndex = 0
               If vg_tipbase = "1" Then
                  CalcularPMPDiaAccess V_Acceso, False, True
               Else
                  CalcularPMPDiaSql V_Acceso, False, True
               End If
            End If
            Do While Fecha <= CDate(fg_Desencripta(TipoDato((RS1!par_valor), ""))) - 1
               V_Acceso.Label1(1).Visible = True
               V_Acceso.Label1(1).Caption = "Procesando : "
               V_Acceso.Label1(0).Caption = Mid(Fecha, 1, 2) & "/" & Mid(CDate(fg_Desencripta(TipoDato((RS1!par_valor), ""))) - 1, 1, 2)
               V_Acceso.Combo1(0).Clear
               V_Acceso.Combo1(0).AddItem "Proc. Día : " & Fecha
               V_Acceso.Combo1(0).ListIndex = 0
               vg_ciedia = Fecha
               '-------> Mover Cantidad decimales
               vg_DCa = IIf(IsNull(GetParametro("parcandec")), 3, GetParametro("parcandec"))
               If vg_DCa = 0 Then vg_DCa = 3
               If vg_tipbase = "1" Then
                  CalcularPMPDiaAccess V_Acceso, False, True
               Else
                  CalcularPMPDiaSql V_Acceso, False, True
               End If
               '-------> Actualizar b_productospmpdia
               If vg_tipbase = "1" Then
                  vg_db.Execute "UPDATE b_productospmpdia INNER JOIN b_tomainv ON b_productospmpdia.ppd_codpro = b_tomainv.tin_codpro SET b_productospmpdia.ppd_saldo = b_tomainv.tin_stofis " & _
                                "WHERE b_productospmpdia.ppd_fecdia = " & Format(CDate(vg_ciedia), "yyyymmdd") & " AND b_productospmpdia.ppd_cencos = '" & vg_contra & "' AND b_tomainv.tin_codbod = " & vg_codbod & " AND b_tomainv.tin_fectom = " & Format(CDate(vg_ciedia), "yyyymmdd") & " AND b_tomainv.tin_ciemes = 0"
               Else
                  vg_db.Execute "UPDATE b_productospmpdia SET b_productospmpdia.ppd_saldo = b.tin_stofis FROM b_productospmpdia a, b_tomainv b WHERE a.ppd_codpro = b.tin_codpro " & _
                                "AND a.ppd_fecdia = " & Format(CDate(vg_ciedia), "yyyymmdd") & " AND a.ppd_cencos = '" & vg_contra & "' AND b.tin_codbod = " & vg_codbod & " AND b.tin_fectom = " & Format(CDate(vg_ciedia), "yyyymmdd") & " " 'AND b.tin_ciemes = 0"
               End If
               If vg_tipbase = "1" Then
                  '-------> Actualizar precio toma inventario
                  vg_db.Execute "UPDATE b_tomainv INNER JOIN b_productospmpdia ON (b_tomainv.tin_fectom = b_productospmpdia.ppd_fecdia) AND (b_tomainv.tin_codpro = b_productospmpdia.ppd_codpro) SET b_tomainv.tin_propon = b_productospmpdia.ppd_propon " & _
                                "WHERE b_tomainv.tin_ciemes = 0 AND Mid(b_tomainv.tin_fectom, 1, 6) = " & Format(CDate(fg_Desencripta(TipoDato((RS1!par_valor), ""))), "yyyymm") & " And b_tomainv.tin_codbod = " & vg_codbod & " AND b_productospmpdia.ppd_cencos = '" & vg_contra & "' AND b_productospmpdia.ppd_fecdia = " & Format(CDate(Fecha), "yyyymmdd") & ""
                  '-------> Actualizar precio en ajuste inventario
                  vg_db.Execute "UPDATE (b_totventas INNER JOIN b_detventas ON (b_totventas.tov_numdoc = b_detventas.dev_numdoc) AND (b_totventas.tov_tipdoc = b_detventas.dev_tipdoc) AND (b_totventas.tov_rutcli = b_detventas.dev_rutcli)) INNER JOIN b_productospmpdia ON (b_detventas.dev_codmer = b_productospmpdia.ppd_codpro) AND (b_totventas.tov_rutcli = b_productospmpdia.ppd_cencos) SET b_detventas.dev_precos = b_productospmpdia.ppd_propon, b_detventas.dev_predoc = b_productospmpdia.ppd_propon, b_detventas.dev_ptotal = (b_detventas.dev_canmer*b_productospmpdia.ppd_propon) " & _
                                "WHERE b_productospmpdia.ppd_cencos = '" & vg_contra & "' AND b_productospmpdia.ppd_fecdia = " & Format(CDate(Fecha), "yyyymmdd") & " AND b_totventas.tov_codbod= " & vg_codbod & "  AND b_totventas.tov_fecemi = CDate('" & Fecha & "') AND b_totventas.tov_estdoc <> 'A' AND b_totventas.tov_tipdoc = 'AI'"
               Else
                  '-------> Actualizar precio toma inventario
                  vg_db.Execute "UPDATE b_tomainv SET b_tomainv.tin_propon = b.ppd_propon FROM b_tomainv a, b_productospmpdia b WHERE a.tin_fectom = b.ppd_fecdia AND a.tin_codpro = b.ppd_codpro  " & _
                                "AND a.tin_ciemes = 0 AND convert(int,substring(convert(varchar(8),a.tin_fectom), 1, 6)) = " & Format(CDate(fg_Desencripta(TipoDato((RS1!par_valor), ""))), "yyyymm") & " And a.tin_codbod = " & vg_codbod & " AND b.ppd_cencos = '" & vg_contra & "' AND b.ppd_fecdia = " & Format(CDate(Fecha), "yyyymmdd") & ""
                  '-------> Actualizar precio en ajuste inventario
                  vg_db.Execute "UPDATE b_detventas SET b_detventas.dev_precos = c.ppd_propon, b_detventas.dev_predoc = c.ppd_propon, b_detventas.dev_ptotal = (b.dev_canmer*c.ppd_propon) FROM b_totventas a,  b_detventas b, b_productospmpdia c WHERE a.tov_numdoc = b.dev_numdoc AND a.tov_tipdoc = b.dev_tipdoc AND a.tov_rutcli = b.dev_rutcli AND b.dev_codmer = c.ppd_codpro AND a.tov_rutcli = c.ppd_cencos  " & _
                                "AND c.ppd_cencos = '" & vg_contra & "' AND c.ppd_fecdia = " & Format(CDate(Fecha), "yyyymmdd") & " AND a.tov_codbod= " & vg_codbod & "  AND a.tov_fecemi = '" & Format(Fecha, "yyyymmdd") & "' AND a.tov_estdoc <> 'A' AND a.tov_tipdoc = 'AI'"
               End If
               If vg_tipbase = "1" Then
                  RS2.Open "SELECT a.tov_numdoc, a.tov_tipdoc, a.tov_rutcli, a.tov_fecemi, Sum(b.dev_ptotal) AS dev_ptotal " & _
                           "FROM b_totventas a, b_detventas b " & _
                           "WHERE a.tov_numdoc = b.dev_numdoc " & _
                           "AND   a.tov_tipdoc = b.dev_tipdoc " & _
                           "AND   a.tov_rutcli =  b.dev_rutcli " & _
                           "AND   a.tov_fecemi = CDate('" & Fecha & "') " & _
                           "AND   a.tov_tipdoc = 'AI' " & _
                           "AND   a.tov_estdoc <> 'A' " & _
                           "AND   a.tov_codbod = " & vg_codbod & " " & _
                           "AND   a.tov_rutcli = '" & vg_contra & "' " & _
                           "GROUP BY  a.tov_numdoc, a.tov_tipdoc, a.tov_rutcli, a.tov_fecemi", vg_db, adOpenStatic
               Else
                  RS2.Open "SELECT a.tov_numdoc, a.tov_tipdoc, a.tov_rutcli, a.tov_fecemi, Sum(b.dev_ptotal) AS dev_ptotal " & _
                           "FROM b_totventas a, b_detventas b " & _
                           "WHERE a.tov_numdoc = b.dev_numdoc " & _
                           "AND   a.tov_tipdoc = b.dev_tipdoc " & _
                           "AND   a.tov_rutcli =  b.dev_rutcli " & _
                           "AND   a.tov_fecemi = '" & Format(Fecha, "yyyymmdd") & "' " & _
                           "AND   a.tov_tipdoc = 'AI' " & _
                           "AND   a.tov_estdoc <> 'A' " & _
                           "AND   a.tov_codbod = " & vg_codbod & " " & _
                           "AND   a.tov_rutcli = '" & vg_contra & "' " & _
                           "GROUP BY  a.tov_numdoc, a.tov_tipdoc, a.tov_rutcli, a.tov_fecemi", vg_db, adOpenStatic
               End If
               If Not RS2.EOF Then
                  Do While Not RS2.EOF
                    vg_db.Execute "UPDATE b_totventas SET tov_totdoc = " & RS2!dev_ptotal & " WHERE tov_numdoc = " & RS2!tov_numdoc & " AND tov_tipdoc = '" & RS2!tov_tipdoc & "' AND tov_rutcli = '" & RS2!tov_rutcli & "' AND tov_fecemi = " & RS2!tov_fecemi & " AND tov_estdoc <> 'A' AND tov_codbod = " & vg_codbod & ""
                     RS2.MoveNext
                  Loop
               End If
               RS2.Close: Set RS2 = Nothing
               Fecha = Fecha + 1
            Loop
         End If
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   V_Acceso.Combo1(0).Clear
   vg_contra = ""
   vg_codbod = 0
   vg_ciedia = ""
   V_Acceso.Frame1(0).Enabled = True
   V_Acceso.Frame1(1).Enabled = True
'   V_Acceso.Label2.Top = 360
'   V_Acceso.Label2.Caption = "Sodeho Chile S.A."
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Label1(0).Visible = False
   '-------> Borrar tabla b_contlistprepro
   vg_db.Execute "ALTER TABLE b_contlistprepro DROP CONSTRAINT FK_b_contlistprepro_b_productos"
   vg_db.Execute "ALTER TABLE b_contlistprepro DROP CONSTRAINT FK_b_contlistprepro_b_clientes"
   vg_db.Execute "DROP TABLE b_contlistprepro"
   vg_db.Execute "UPDATE a_param SET par_valor='159' WHERE par_codigo='version'"
   aVer = 159
End If
If nVer > aVer And aVer = 159 Then
   '-------> Insertar campo a la tabla b_proveedor
   vg_db.Execute "ALTER TABLE b_proveedor ADD COLUMN prv_activo char(1), prv_fecumo DATE, prv_origen char(1)"
   vg_db.Execute "UPDATE b_proveedor SET prv_activo='0', prv_fecumo=null, prv_origen='1'"
   '-------> Crear tabla actualiza datos
   vg_db.Execute "CREATE TABLE b_actuadatos (ada_nomtab char(255), ada_fecumo date, Constraint log_actuadatos_pk Primary Key (ada_nomtab))"
   vg_db.Execute "INSERT INTO b_actuadatos VALUES ('b_proveedor', null)"
   '-------> Actualiza tabla b_detventas 1 ajuste inventario - 2 reajuste precio - 3 ajuste inventario y reajuste precio
   vg_db.Execute "UPDATE b_detventas SET dev_acepre='1' WHERE dev_tipdoc = 'AI'"
   '-------> Actualiza versión
   vg_db.Execute "UPDATE a_param SET par_valor='160' WHERE par_codigo='version'"
   aVer = 160
End If
If nVer > aVer And aVer = 160 Then
   vg_db.Execute "UPDATE b_contlistpreing SET b_contlistpreing.cpi_codcom = b_contlistpreing.cpi_codped WHERE b_contlistpreing.cpi_codcom = '0'"
   vg_db.Execute "UPDATE a_param SET par_valor = '161' WHERE par_codigo='version'"
   aVer = 161
End If
If nVer > aVer And aVer = 161 Then
   RS1.Open "SELECT * FROM a_param WHERE par_codigo = 'modprove'", vg_db, adOpenStatic
   If RS1.EOF Then
      RS1.Close: Set RS1 = Nothing
      '-------> Bloquear proveedor
      RS1.Open "SELECT * FROM b_clientes WHERE cli_tipo=0 AND cli_codbod>0", vg_db, adOpenStatic
      Do While Not RS1.EOF
         vg_db.Execute "INSERT INTO a_param VALUES ('modprove', 'Parametro Modificar Proveedor', 'N', '1', '" & Trim(RS1!cli_codigo) & "')"
         RS1.MoveNext
      Loop
      RS1.Close: Set RS1 = Nothing
   Else
      RS1.Close: Set RS1 = Nothing
   End If
   '-------> Incluir opción comparativo curva abc y comparativo costo teorico v/s negociado
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3027000, 'Comparativo Curva ABC')"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3031000, 'Comparativo Costo Teórico vs Negociado')"
   
   '-------> Incluir lista precio sac
   vg_db.Execute "CREATE TABLE b_formatocompras (foc_codsac char(20), foc_codcat int, foc_nomsac char(100), foc_unisac char(20), foc_vigini datetime, foc_flexec int, foc_vigfin datetime, Constraint b_formatocompras_pk Primary Key (foc_codsac))"
   vg_db.Execute "CREATE TABLE b_formatocomprassgp (fcs_codsac char(20), fcs_codsgp char(20), fcs_sgppre int, Constraint b_formatocomprassgp_pk Primary Key (fcs_codsac, fcs_codsgp))"
   vg_db.Execute "ALTER TABLE b_formatocomprassgp ADD CONSTRAINT FK_b_formatocomprassgp_b_formatocompras FOREIGN KEY (fcs_codsac) REFERENCES b_formatocompras (foc_codsac)"
   vg_db.Execute "ALTER TABLE b_formatocomprassgp ADD CONSTRAINT FK_b_formatocomprassgp_b_productos FOREIGN KEY (fcs_codsgp) REFERENCES b_productos (pro_codigo)"
   
   '-------> Incluir lista precio sac
'   vg_db.Execute "CREATE TABLE b_sac_listaprecio (lps_cencos char(10), lps_fecini datetime, lps_fecfin datetime, lps_periodo char(06), lps_codsac char(20), lps_precio double, Constraint b_sac_listaprecio_pk Primary Key (lps_cencos, lps_fecini, lps_fecfin, lps_periodo, lps_codsac))"
   vg_db.Execute "CREATE TABLE b_sac_listaprecio (lps_cencos char(10), lps_periodo char(06), lps_codsac char(20), lps_precio double, Constraint b_sac_listaprecio_pk Primary Key (lps_cencos, lps_periodo, lps_codsac))"
'   vg_db.Execute "ALTER TABLE b_sac_listaprecio ADD CONSTRAINT FK_b_sac_listaprecio_b_formatocompras FOREIGN KEY (lps_codsac) REFERENCES b_formatocompras (foc_codsac)"
   vg_db.Execute "ALTER TABLE b_sac_listaprecio ADD CONSTRAINT FK_b_sac_listaprecio_b_clientes FOREIGN KEY (lps_cencos) REFERENCES b_clientes (cli_codigo)"
   
   vg_db.Execute "UPDATE a_param SET par_valor = '162' WHERE par_codigo='version'"
   aVer = 162
End If
If nVer > aVer And aVer = 162 Then
   '-------> Insertar campo tabla clientes con la opción sobreescribe receta si es igual a 0 = solo fijos, 1 = todos y 2 = ninguno
   vg_db.Execute "ALTER TABLE b_clientes ADD cli_sobrec varchar(1)"
   vg_db.Execute "UPDATE b_clientes SET cli_sobrec = '0' WHERE cli_tipo = 0"
   '-------> Actualiza versión
   vg_db.Execute "UPDATE a_param SET par_valor='163' WHERE par_codigo='version'"
   aVer = 163
End If
If nVer > aVer And aVer = 163 Then
   '-------> Borrar relación formato compras encabezado con el detalle
   vg_db.Execute "ALTER TABLE b_formatocomprassgp DROP CONSTRAINT FK_b_formatocomprassgp_b_formatocompras"
   '-------> Crear tabla detalle orden de compras recibido
   vg_db.Execute "CREATE TABLE log_enviocierrediario (cencos char(10), fecha datetime, estenv char(1), fecsub char(30), Constraint log_enviocierrediario_pk Primary Key (cencos, fecha))"
   '-------> Crear tabla detalle orden de compras recibido
   vg_db.Execute "CREATE TABLE b_ocsacrecibido (ocr_rutpro char(10), ocr_tipdoc char(2), ocr_numdoc int, ocr_numlin int, ocr_codprodsgp char(20), ocr_codprodsac char(20), ocr_cancom double, ocr_precom double, ocr_canrec double, ocr_fecoc datetime, ocr_canoc double, ocr_preoc double, Constraint b_doccompras_pk Primary Key (ocr_rutpro, ocr_tipdoc, ocr_numdoc, ocr_numlin, ocr_codprodsgp))"
   '-------> Relacionar tabla b_ocsacrecibido & totcompras
   'vg_db.Execute "ALTER TABLE b_ocsacrecibido  ADD  CONSTRAINT [FK_b_ocsacrecibido_b_totcompras] FOREIGN KEY(ocr_rutpro, ocr_tipdoc, ocr_numdoc) References b_totcompras(toc_rutpro, toc_tipdoc, toc_numdoc)"
   '-------> Crear tabla ordenes de compras sac
   vg_db.Execute "CREATE TABLE b_ocsac (cadfil_cdfil char(10), solite_dtent datetime, pedido_dtped datetime, tipsol_idsol int, pedido_cdped int, tabcen_cdcen char(4), cadfor_idfor int, cadfor_nrcgc char(20), pedido_dtref char(6), pedido_nrsem int, pedido_pcref double, pedido_flenc int, tabemp_idemp int, pedido_flexp int, solfil_idsol int, cpopro_cdpro char(20), pedite_qtcpa double, pedite_vlpco double, pedite_flafo int, pedite_flagr int, pedite_nrafo int, [timestamp] datetime, Constraint b_ocsac_pk Primary Key (cadfil_cdfil, solite_dtent, pedido_cdped, solfil_idsol, cpopro_cdpro))"
   '-------> proceso de borrado gestion.ini y crearlo nuevamente con accesso a Access o Sql Server Express
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   If Dir(dir_trabajo & "gestion.ini") <> "" Then
      Kill dir_trabajo & "gestion.ini"
      Open dir_trabajo & "Gestion.ini" For Output As #1
      Print #1, "[Parametros de Configuración]"
      Print #1, ""
      Print #1, "[Tipo Base Dato]"
      Print #1, "Base = " & """" & "1" & """"
      Print #1, ""
      Print #1, "[Path]"
      Print #1, "Ruta = " & """" & dir_trabajo & """"
      Print #1, ""
      Print #1, "[Base_de_datos]"
      Print #1, "Mdb =  " & """" & BaseDeDato & """"
      Print #1, ""
      Print #1, "[Gif]"
      Print #1, "Ruta = " & """" & dir_trabajo & """"
      Print #1, ""
      Print #1, "[Provider]"
      Print #1, "Jet = " & """" & "Microsoft.Jet.OLEDB.4.0" & """"
      Print #1, ""
      Print #1, "[SQL SERVER]"
      Print #1, "Servidor = " & """" & "104359CL200705\SQLEXPRESS" & """"
      Print #1, "Database = ®­³Á"
      Print #1, "Usuario = ½²¼"
      Print #1, "Password = ½²¼·¾½"
      Close #1
      '-------> mover parametros
      vg_tipbase = MiFunc("Tipo Base Dato", "Gestion.Ini", "Base")
      dir_trabajo = MiFunc("Path", "Gestion.Ini", "Ruta")
      BaseDeDato = MiFunc("Base_de_datos", "Gestion.Ini", "Mdb")
      Provider = MiFunc("Provider", "Gestion.Ini", "Jet")
      RutaGif = MiFunc("Gif", "Gestion.Ini", "Ruta")
      vg_DirLog = dir_trabajo & "LogoSdx.jpg"
   End If
   '-------> Actualiza versión
   vg_db.Execute "UPDATE a_param SET par_valor='164' WHERE par_codigo='version'"
   aVer = 164
End If
If nVer > aVer And aVer = 164 Then
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 165 ....."
   V_Acceso.Refresh
   RS1.Open "SELECT DISTINCT par_valor, par_cencos FROM a_param WHERE par_codigo = 'ciediario'", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         V_Acceso.Refresh
         vg_db.Execute "DELETE b_productospmpdia FROM b_productospmpdia WHERE ppd_cencos = '" & RS1!par_cencos & "' AND ppd_fecdia <> " & Format((CDate(fg_Desencripta(TipoDato((RS1!par_valor), ""))) - 1), "yyyymmdd") & " AND ppd_propon = 0 AND ppd_saldo = 0"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   vg_db.Execute "UPDATE a_param SET par_valor = '165' WHERE par_codigo = 'version'"
   aVer = 165
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 165 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 166 ....."
   V_Acceso.Refresh
   '-------> Agregar Campo a la tabla log_cierrediario
   If vg_tipbase = "1" Then
      vg_db.Execute "ALTER TABLE log_cierrediario ADD cencos varchar(10) NULL"
      vg_db.Execute "DROP INDEX log_cierrediario_pk ON log_cierrediario"
   Else
      vg_db.Execute "ALTER TABLE log_cierrediario ADD cencos varchar(10)"
      vg_db.Execute "ALTER TABLE log_cierrediario DROP PK_log_cierrediario"
   End If
   '-------> Agregar registo a_param que identifique la actualización de base dato.
   RS1.Open "SELECT * FROM b_clientes WHERE cli_tipo = 0 AND cli_codbod > 0", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         vg_db.Execute "INSERT INTO a_param VALUES ('paractbase', 'Parametro Actualización Base de Datos', 'C', ' ', '" & RS1!cli_codigo & "')"
         vg_db.Execute "UPDATE log_cierrediario SET cencos = '" & RS1!cli_codigo & "'"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   If vg_tipbase = "1" Then
      vg_db.Execute "ALTER TABLE log_cierrediario ADD Constraint log_cierrediario_pk Primary Key (fecha, cencos)"
   Else
     vg_db.Execute "ALTER TABLE log_cierrediario ALTER COLUMN cencos varchar(10) NOT NULL"
     vg_db.Execute "ALTER TABLE log_cierrediario ADD CONSTRAINT PK_log_cierrediario PRIMARY KEY  CLUSTERED " & _
                   "(fecha, cencos)  ON [PRIMARY]"
   End If
   '-------> Actualiza versión
   vg_db.Execute "UPDATE a_param SET par_valor = '166' WHERE par_codigo = 'version'"
   aVer = 166
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 166 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 167 ....."
   V_Acceso.Refresh
   vg_db.Execute "UPDATE a_param SET par_valor = '167' WHERE par_codigo = 'version'"
   aVer = 167
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 167 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 168 ....."
   V_Acceso.Refresh
   '-------> Crear Tabla Fecha Inhabiles , modificar incluir campoa a tabla b_paramdesp, b_casinoactividadesdiaria
   If vg_tipbase = "1" Then
      vg_db.Execute "CREATE TABLE a_tipoactividad (tia_codigo int, tia_nombre char(50), Constraint PK_b_tipoactividad Primary Key (tia_codigo))"
      vg_db.Execute "CREATE TABLE b_casinotipoactividades (cta_cencos char(10), cta_tipact int, Constraint PK_b_casinoactividadesdiaria Primary Key (cta_cencos, cta_tipact))"
      vg_db.Execute "ALTER TABLE b_casinotipoactividades ADD CONSTRAINT FK_b_casinotipoactividades_b_clientes FOREIGN KEY (cta_cencos) REFERENCES b_clientes (cli_codigo)"
      vg_db.Execute "ALTER TABLE b_casinotipoactividades ADD CONSTRAINT FK_b_casinotipoactividades_a_tipoactividad FOREIGN KEY (cta_tipact) REFERENCES a_tipoactividad (tia_codigo)"
      vg_db.Execute "CREATE TABLE b_Fecha_Inhabiles (CFI_CeCo char(10), CFI_Fecha datetime, CFI_Glosa char(100), Constraint PK_b_Fecha_Inhabiles Primary Key (CFI_CeCo, CFI_Fecha))"
      vg_db.Execute "ALTER TABLE b_Fecha_Inhabiles ADD CONSTRAINT FK_b_Fecha_Inhabiles_b_clientes FOREIGN KEY (CFI_CeCo) REFERENCES b_clientes (cli_codigo)"
      vg_db.Execute "CREATE TABLE b_casinoparametrostock (cps_cencos char(10), cps_invsto char(1), cps_reqmen char(1), cps_porinv double, cps_liscri char(1), cps_diario char(1), cps_ajuimp char(1), Constraint PK_b_casinoparametrostock Primary Key (cps_cencos))"
      vg_db.Execute "ALTER TABLE b_casinoparametrostock ADD CONSTRAINT FK_b_casinoparametrostock_b_clientes FOREIGN KEY (cps_cencos) REFERENCES b_clientes (cli_codigo)"
      vg_db.Execute "ALTER TABLE b_paramdesp ALTER COLUMN pad_tipo char(3)"
      vg_db.Execute "ALTER TABLE b_paramdesp ADD pad_diaseg int"
      vg_db.Execute "CREATE TABLE b_minutapedidos (ped_codcas char(10), ped_fecped int, ped_anomes int, ped_tipped int, ped_coding char(20), ped_codpro char(20), ped_codsac char(20), ped_canmin double, ped_canped double, ped_fecenv int, ped_stoact double, ped_proped double, ped_ordrec double, ped_conrea double, Constraint PK_b_minutapedidos Primary Key (ped_codcas, ped_fecped, ped_anomes, ped_tipped, ped_coding, ped_codpro))"
      vg_db.Execute "ALTER TABLE b_proveedor ADD prv_regimp char(1), prv_autret char(1), prv_cuohor char(1), prv_codmun int, prv_docele char(1)"
      vg_db.Execute "UPDATE b_proveedor SET prv_docele = 'N'"
      vg_db.Execute "ALTER TABLE b_clientes ADD cli_codmun int"
      vg_db.Execute "ALTER TABLE b_productos ADD pro_codref int, pro_codrei int, pro_cuohor char(1)"
      vg_db.Execute "CREATE TABLE a_municipio (mun_codigo int, mun_nombre char(100), mun_retobl char(1), Constraint PK_a_municipio Primary Key (mun_codigo))"
      vg_db.Execute "CREATE TABLE b_retencionfuente (ref_codigo int, ref_nombre char(100), ref_portar double, ref_codcta char(10), ref_tipret char(10), ref_indret char(10), Constraint PK_b_retencionfuente Primary Key (ref_codigo))"
      vg_db.Execute "CREATE TABLE b_retencionica (rei_codigo int, rei_nombre char(100), Constraint PK_b_retencionica Primary Key (rei_codigo))"
      vg_db.Execute "CREATE TABLE b_detretencionica (dri_codigo int, dri_codmun int, dri_portar double, dri_codcta char(10), dri_tipret char(10), dri_indret char(10), Constraint PK_b_detretencionica Primary Key (dri_codigo, dri_codmun ))"
      vg_db.Execute "ALTER TABLE b_detretencionica ADD CONSTRAINT FK_b_detretencionica_a_municipio FOREIGN KEY(dri_codmun) References a_municipio (mun_codigo)"
      vg_db.Execute "ALTER TABLE b_detretencionica ADD CONSTRAINT FK_b_detretencionica_b_retencionica FOREIGN KEY (dri_codigo) References b_retencionica (rei_codigo)"
      vg_db.Execute "CREATE TABLE a_pais (pai_codigo char(10), pai_nombre char(100), Constraint PK_a_pais Primary Key (pai_codigo))"
      vg_db.Execute "ALTER TABLE b_totcompras ADD toc_fecrem datetime"
      vg_db.Execute "UPDATE b_totcompras SET toc_fecrem = toc_fecemi"
      vg_db.Execute "ALTER TABLE b_gastosa13 ADD gas_valpro double"
      vg_db.Execute "ALTER TABLE a_tipopro ADD tip_activo char(1)"
      vg_db.Execute "UPDATE a_tipopro SET tip_activo = 'S'"
   Else
      vg_db.Execute "CREATE TABLE dbo.a_tipoactividad (tia_codigo int, tia_nombre varchar(50), Constraint PK_b_tipoactividad Primary Key (tia_codigo))"
      vg_db.Execute "CREATE TABLE dbo.b_casinotipoactividades (cta_cencos varchar(10), cta_tipact int, Constraint PK_b_casinotipoactividades Primary Key (cta_cencos, cta_tipact))"
      vg_db.Execute "ALTER TABLE dbo.b_casinotipoactividades ADD CONSTRAINT FK_b_casinotipoactividades_b_clientes FOREIGN KEY (cta_cencos) REFERENCES b_clientes (cli_codigo)"
      vg_db.Execute "ALTER TABLE dbo.b_casinotipoactividades ADD CONSTRAINT FK_b_casinotipoactividades_a_tipoactividad FOREIGN KEY (cta_tipact) REFERENCES a_tipoactividad (tia_codigo)"
      vg_db.Execute "CREATE TABLE dbo.b_Fecha_Inhabiles (CFI_CeCo varchar(10), CFI_Fecha datetime, CFI_Glosa varchar(100), Constraint PK_b_Fecha_Inhabiles Primary Key (CFI_CeCo, CFI_Fecha))"
      vg_db.Execute "ALTER TABLE dbo.b_Fecha_Inhabiles ADD CONSTRAINT FK_b_Fecha_Inhabiles_b_clientes FOREIGN KEY (CFI_CeCo) REFERENCES b_clientes (cli_codigo)"
      vg_db.Execute "CREATE TABLE dbo.b_casinoparametrostock (cps_cencos varchar(10), cps_invsto varchar(1), cps_reqmen varchar(1), cps_porinv float, cps_liscri varchar(1), cps_diario varchar(1), cps_ajuimp varchar(1), Constraint PK_b_casinoparametrostock Primary Key (cps_cencos))"
      vg_db.Execute "ALTER TABLE dbo.b_casinoparametrostock ADD CONSTRAINT FK_b_casinoparametrostock_b_clientes FOREIGN KEY (cps_cencos) REFERENCES b_clientes (cli_codigo)"
      vg_db.Execute "ALTER TABLE dbo.b_paramdesp ALTER COLUMN pad_tipo varchar(3)"
      vg_db.Execute "ALTER TABLE dbo.b_paramdesp ADD pad_diaseg int"
      vg_db.Execute "CREATE TABLE dbo.b_minutapedidos (ped_codcas varchar(10), ped_fecped int, ped_anomes int, ped_tipped int, ped_coding varchar(20), ped_codpro varchar(20), ped_codsac varchar(20), ped_canmin float, ped_canped float, ped_fecenv int, ped_stoact float, ped_proped float, ped_ordrec float, ped_conrea float, Constraint PK_b_minutapedidos Primary Key (ped_codcas, ped_fecped, ped_anomes, ped_tipped, ped_coding, ped_codpro))"
      vg_db.Execute "ALTER TABLE dbo.b_proveedor ADD prv_regimp varchar(1), prv_autret varchar(1), prv_cuohor varchar(1), prv_codmun int, prv_docele varchar(1)"
      vg_db.Execute "UPDATE dbo.b_proveedor SET prv_docele = 'N'"
      vg_db.Execute "ALTER TABLE dbo.b_clientes ADD cli_codmun int"
      vg_db.Execute "ALTER TABLE dbo.b_productos ADD pro_codref int, pro_codrei int, pro_cuohor varchar(1)"
      vg_db.Execute "CREATE TABLE dbo.a_municipio (mun_codigo int, mun_nombre varchar(100), mun_retobl varchar(1), Constraint PK_a_municipio Primary Key (mun_codigo))"
      vg_db.Execute "CREATE TABLE dbo.b_retencionfuente (ref_codigo int, ref_nombre varchar(100), ref_portar float, ref_codcta varchar(10), ref_tipret varchar(10), ref_indret varchar(10), Constraint PK_b_retencionfuente Primary Key (ref_codigo))"
      vg_db.Execute "CREATE TABLE dbo.b_retencionica (rei_codigo int, rei_nombre varchar(100), Constraint PK_b_retencionica Primary Key (rei_codigo))"
      vg_db.Execute "CREATE TABLE dbo.b_detretencionica (dri_codigo int, dri_codmun int, dri_portar float, dri_codcta varchar(10), dri_tipret varchar(10), dri_indret varchar(10), Constraint PK_b_detretencionica Primary Key (dri_codigo, dri_codmun ))"
      vg_db.Execute "ALTER TABLE dbo.b_detretencionica ADD CONSTRAINT FK_b_detretencionica_a_municipio FOREIGN KEY(dri_codmun) References DBO.a_municipio (mun_codigo)"
      vg_db.Execute "ALTER TABLE dbo.b_detretencionica ADD CONSTRAINT FK_b_detretencionica_b_retencionica FOREIGN KEY (dri_codigo) References DBO.b_retencionica (rei_codigo)"
      vg_db.Execute "CREATE TABLE dbo.a_pais (pai_codigo varchar(10), pai_nombre varchar(100), Constraint PK_a_pais Primary Key (pai_codigo))"
      vg_db.Execute "ALTER TABLE dbo.b_totcompras ADD toc_fecrem datetime"
      vg_db.Execute "UPDATE b_totcompras SET toc_fecrem = toc_fecemi"
      vg_db.Execute "ALTER TABLE dbo.b_gastosa13 ADD gas_valpro float"
      vg_db.Execute "ALTER TABLE dbo.a_tipopro ADD tip_activo varchar(1)"
      vg_db.Execute "UPDATE dbo.a_tipopro SET tip_activo = 'S'"
   End If
   vg_db.Execute "INSERT INTO a_tipoactividad VALUES (1, 'Facturas de Proveedores')"
   vg_db.Execute "INSERT INTO a_tipoactividad VALUES (2, 'Salidas a Producción')"
   vg_db.Execute "INSERT INTO a_tipoactividad VALUES (3, 'Devolución a Bodega')"
   vg_db.Execute "INSERT INTO a_tipoactividad VALUES (4, 'Mermas')"
   vg_db.Execute "INSERT INTO a_tipoactividad VALUES (5, 'Raciones no Vendidas')"
   vg_db.Execute "INSERT INTO a_tipoactividad VALUES (6, 'Control de Raciones')"
   vg_db.Execute "INSERT INTO a_tipoactividad VALUES (7, 'Registro de Venta Cafeteria')"
   vg_db.Execute "INSERT INTO a_tipoactividad VALUES (8, 'Venta de Servicios de Contado')"
   vg_db.Execute "INSERT INTO a_tipoactividad VALUES (9, 'Venta Directa')"
   vg_db.Execute "INSERT INTO a_tipoactividad VALUES (10, 'Inventario Rotativo')"
   '-------> Mover Datos a tabla pais
   vg_db.Execute "INSERT INTO a_pais VALUES ('CL', 'Chile')"
   vg_db.Execute "INSERT INTO a_pais VALUES ('CO', 'Colombia')"
   vg_db.Execute "INSERT INTO a_pais VALUES ('PE', 'Peru')"
   vg_db.Execute "INSERT INTO a_pais VALUES ('AR', 'Argentina')"
   '-------> Mover datos a tabla a_param por defecto chile
   RS1.Open "SELECT DISTINCT cli_codigo FROM b_clientes WHERE cli_tipo = 0 AND cli_codbod > 0", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         vg_db.Execute "INSERT INTO a_param VALUES ('parpais', 'Parametro Pais', 'C', 'CL', '" & RS1!cli_codigo & "')"
         vg_db.Execute "INSERT INTO a_param VALUES ('pariva', 'Parametro iva', 'C', '1', '" & RS1!cli_codigo & "')"
         vg_db.Execute "INSERT INTO a_param VALUES ('parivacig', 'Parametro iva Cigarrillo', 'C', '11', '" & RS1!cli_codigo & "')"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   
   vg_db.Execute "UPDATE b_paramdesp SET pad_tipo = 'Q1' WHERE pad_tipo = 'Q'"
   vg_db.Execute "UPDATE b_paramdesp SET pad_diario = '1000000' WHERE pad_tipo = 'S'"
   '-------> Incluir opción informe detalle cartola inventario
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3049000, 'Detalle Cartola Inventario')"
   vg_db.Execute "UPDATE a_param SET par_valor = '168' WHERE par_codigo = 'version'"
   aVer = 168
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 168 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 169 ....."
   V_Acceso.Refresh
   If vg_tipbase <> "1" Then
'      Dim BaseDatos  As String
      BaseDatos = "Actualizador.mdb"
      Set vg_dbsubesql = New ADODB.Connection
      vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
      vg_dbsubesql.ConnectionTimeout = 30
      vg_dbsubesql.CommandTimeout = 600
      vg_dbsubesql.Open
      RS1.Open "SELECT * FROM Procedimiento where Version = '169' order by id1", vg_dbsubesql, adOpenStatic
      If Not RS1.EOF Then
         Do While Not RS1.EOF
            vg_db.Execute ("" & RS1!Procedimiento & "")
            RS1.MoveNext
         Loop
      End If
      RS1.Close: Set RS1 = Nothing
      vg_dbsubesql.Close
   End If

   If vg_tipbase = "1" Then
      vg_db.Execute "ALTER TABLE b_productos ADD pro_tipord char(1)"
      vg_db.Execute "UPDATE b_productos SET pro_tipord = '2'"
      vg_db.Execute "ALTER TABLE b_tomainv ADD tin_tipinv char(1)"
      vg_db.Execute "UPDATE b_tomainv SET tin_tipinv = '2'"
      vg_db.Execute "ALTER TABLE b_clientes ADD cli_ccisac int, cli_cecsac char(4)"
      vg_db.Execute "ALTER TABLE b_minutapedidos ADD ped_persac char(6), ped_semsac int"
   Else
      vg_db.Execute "ALTER TABLE dbo.b_productos ADD pro_tipord varchar(1)"
      vg_db.Execute "UPDATE dbo.b_productos SET pro_tipord = '2'"
      vg_db.Execute "ALTER TABLE dbo.b_tomainv ADD tin_tipinv varchar(1)"
      vg_db.Execute "UPDATE dbo.b_tomainv SET tin_tipinv = '2'"
      vg_db.Execute "ALTER TABLE dbo.b_clientes ADD cli_ccisac int, cli_cecsac varchar(4)"
      vg_db.Execute "ALTER TABLE dbo.b_minutapedidos ADD ped_persac varchar(6), ped_semsac int"
   End If
   vg_db.Execute "UPDATE a_param SET par_valor = '169' WHERE par_codigo = 'version'"
   aVer = 169
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 169 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 170 ....."
   V_Acceso.Refresh
   If vg_tipbase = "2" Then
      BaseDatos = "Actualizador.mdb"
      Set vg_dbsubesql = New ADODB.Connection
      vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
      vg_dbsubesql.ConnectionTimeout = 30
      vg_dbsubesql.CommandTimeout = 600
      vg_dbsubesql.Open
      RS1.Open "SELECT * FROM Procedimiento WHERE Version = '170' order by id1", vg_dbsubesql, adOpenStatic
      If Not RS1.EOF Then
         Do While Not RS1.EOF
            vg_db.Execute ("" & RS1!Procedimiento & "")
            RS1.MoveNext
         Loop
      End If
      RS1.Close: Set RS1 = Nothing
      vg_dbsubesql.Close
   End If

   '-------> Actualiza versión
   vg_db.Execute "UPDATE a_param SET par_valor='170' WHERE par_codigo='version'"
   aVer = 170
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 170 Then
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 171 ....."
   V_Acceso.Refresh
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   '-------> Insertar campo tablas y creación
   If vg_tipbase = "1" Then
      vg_db.Execute "ALTER TABLE a_impuesto ADD imp_cimsap1 char(2), imp_cimsap2 char(2), imp_cimsap3 char(2), imp_cimsap4 char(4)"
      vg_db.Execute "ALTER TABLE b_clientes ADD cli_codreg int"
      vg_db.Execute "ALTER TABLE b_detcompras ADD dec_cmefac double, dec_pmefac double, dec_crefac double, dec_prefac double, dec_faccon double"
      vg_db.Execute "UPDATE b_detcompras SET dec_cmefac = dec_canmer, dec_pmefac = dec_precom, dec_crefac = dec_canrec, dec_prefac = dec_prerec, dec_faccon = 0"
      vg_db.Execute "ALTER TABLE b_formatocompras ADD foc_faccon double"
      vg_db.Execute "UPDATE b_formatocompras SET foc_faccon = 1"
      vg_db.Execute "ALTER TABLE b_formatocomprassgp ADD fcs_cenpre int"
      vg_db.Execute "UPDATE b_formatocomprassgp SET fcs_cenpre = 0"
      vg_db.Execute "ALTER TABLE b_productos ADD pro_tippro char(1)"
      vg_db.Execute "UPDATE b_productos SET pro_tippro = '0'"
      vg_db.Execute "CREATE TABLE a_region (reg_codigo int, reg_nombre char(50), Constraint PK_a_region Primary Key (reg_codigo))"
      vg_db.Execute "CREATE TABLE a_clasedocsap (cds_coddoc char(2), cds_codreg int, cds_cdosap char(2), Constraint PK_a_clasedocsap Primary Key (cds_coddoc, cds_codreg))"
   Else
      vg_db.Execute "ALTER TABLE dbo.a_impuesto ADD imp_cimsap1 varchar(2), imp_cimsap2 varchar(2), imp_cimsap3 varchar(2), imp_cimsap4 varchar(2)"
      vg_db.Execute "ALTER TABLE dbo.b_clientes ADD cli_codreg int"
      vg_db.Execute "ALTER TABLE dbo.b_detcompras ADD dec_cmefac float, dec_pmefac float, dec_crefac float, dec_prefac float, dec_faccon float"
      vg_db.Execute "UPDATE dbo.b_detcompras SET dec_cmefac = dec_canmer, dec_pmefac = dec_precom, dec_crefac = dec_canrec, dec_prefac = dec_prerec, dec_faccon = 0"
      vg_db.Execute "ALTER TABLE dbo.b_formatocompras ADD foc_faccon float"
      vg_db.Execute "UPDATE dbo.b_formatocompras SET foc_faccon = 1"
      vg_db.Execute "ALTER TABLE dbo.b_formatocomprassgp ADD fcs_cenpre int"
      vg_db.Execute "UPDATE dbo.b_formatocomprassgp SET fcs_cenpre = 0"
      vg_db.Execute "ALTER TABLE dbo.b_productos ADD pro_tippro varchar(1)"
      vg_db.Execute "UPDATE dbo.b_productos SET pro_tippro = '0'"
      vg_db.Execute "CREATE TABLE dbo.a_region (reg_codigo int, reg_nombre varchar(50), Constraint PK_a_region Primary Key (reg_codigo))"
      vg_db.Execute "CREATE TABLE dbo.a_clasedocsap (cds_coddoc varchar(2), cds_codreg int, cds_cdosap varchar(2), Constraint PK_a_clasedocsap Primary Key (cds_coddoc, cds_codreg))"
   End If
   RS1.Open "SELECT DISTINCT a.cli_codigo, b.par_valor FROM b_clientes a, a_param b WHERE a.cli_codigo = b.par_cencos AND b.par_codigo = 'parpais' AND a.cli_tipo = 0 AND a.cli_codbod > 0", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         If RS1!par_valor = "CL" Then
            vg_db.Execute "UPDATE a_impuesto SET imp_cimsap1 = 'C1' WHERE imp_adicional = 0"
            vg_db.Execute "UPDATE a_impuesto SET imp_cimsap1 = 'C0' WHERE imp_codigo NOT IN (3, 8) AND imp_adicional <> 0"
            vg_db.Execute "UPDATE a_impuesto SET imp_cimsap1 = '' WHERE imp_codigo IN (3, 8)"
            '-------> Insertar tipo moneda
            RS2.Open "SELECT par_valor FROM a_param WHERE par_codigo = 'tipmonsap' AND par_cencos = '" & RS1!cli_codigo & "'", vg_db, adOpenStatic
            If RS2.EOF Then
               vg_db.Execute "INSERT INTO a_param VALUES ('tipmonsap', 'Parametro Tipo Moneda SAP', 'C', 'CLP', '" & RS1!cli_codigo & "')"
            End If
            RS2.Close: Set RS2 = Nothing
            '-------> Insertar codigo sap exento1
            RS2.Open "SELECT par_valor FROM a_param WHERE par_codigo = 'codsapexe1' AND par_cencos = '" & RS1!cli_codigo & "'", vg_db, adOpenStatic
            If RS2.EOF Then
               vg_db.Execute "INSERT INTO a_param VALUES ('codsapexe1', 'Código sap Exento Nº1', 'C', 'C0', '" & RS1!cli_codigo & "')"
            End If
            RS2.Close: Set RS2 = Nothing
         ElseIf RS1!par_valor = "CO" Then
            RS2.Open "SELECT par_valor FROM a_param WHERE par_codigo = 'tipmonsap' AND par_cencos = '" & RS1!cli_codigo & "'", vg_db, adOpenStatic
            If RS2.EOF Then
               vg_db.Execute "INSERT INTO a_param VALUES ('tipmonsap', 'Parametro Tipo Moneda SAP', 'C', 'COP', '" & RS1!cli_codigo & "')"
            End If
            RS2.Close: Set RS2 = Nothing
         End If
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   '-------> Asignar nuevo correlativo de folio si existe documento electronicos
   Dim corre As Long
   Dim periodo As Long
   corre = 0
   periodo = 0
   RS1.Open "SELECT DISTINCT a.inf_cencos, a.inf_numero, b.cli_codbod FROM a_infcfcfofi a, b_clientes b WHERE a.inf_cencos = b.cli_codigo AND b.cli_codbod > 0 AND a.inf_tipo = 'C' AND (a.inf_feccie = 0 OR (a.inf_feccie) IS NULL) AND (a.inf_usuario = '' OR (a.inf_usuario) IS NULL)", vg_db, adOpenStatic
   Do While Not RS1.EOF
      sql1 = IIf(vg_tipbase = "1", " IIF(toc_tipdoc <> 'FE' OR toc_tipdoc <> 'DE' OR toc_tipdoc <> 'CE','FA', '') AS facnor, IIF(toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE','FE', '') AS facele, COUNT(IIF(toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE','FE', 'FA')) AS nreg ", " (CASE WHEN toc_tipdoc <> 'FE' OR toc_tipdoc <> 'DE' OR toc_tipdoc <> 'CE' THEN 'FA' END) facnor, (CASE WHEN toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE' THEN 'FE' END) AS facele, COUNT(CASE WHEN toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE' THEN 'FE' ELSE 'FA' END) AS nreg ")
      '-------> Traer Periodo
      RS2.Open "SELECT * FROM b_cierreperiodo WHERE cie_cencos ='" & RS1!inf_cencos & "' AND cie_estado = 1", vg_db, adOpenStatic
      periodo = 0
      If Not RS2.EOF Then
         periodo = IIf(IsNull(RS2!cie_periodo), 0, RS2!cie_periodo)
      End If
      RS2.Close: Set RS2 = Nothing
      
      If vg_tipbase = "1" Then
         RS2.Open "SELECT TOP 1 (SELECT DISTINCT 'FA' FROM  b_totcompras a WHERE a.toc_codbod = " & RS1!cli_codbod & "  AND a.toc_numinf = " & RS1!inf_numero & " AND a.toc_tipdoc NOT IN ('FE','DE','CE','SN') AND a.toc_envsap = '0' AND a.toc_fecper = " & periodo & ") AS facnor, " & _
                  "(SELECT DISTINCT 'FE' FROM b_totcompras b WHERE b.toc_codbod = " & RS1!cli_codbod & " AND b.toc_numinf = " & RS1!inf_numero & " AND b.toc_tipdoc IN ('FE','DE','CE') AND b.toc_envsap = '0' AND b.toc_fecper = " & periodo & ") AS facele, " & _
                  "COUNT(IIF(toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE','FE', 'FA')) AS nreg " & _
                  "FROM b_totcompras WHERE toc_codbod = " & RS1!cli_codbod & " AND toc_numinf = " & RS1!inf_numero & " AND toc_envsap = '0' AND toc_tipdoc not in ('SN') AND toc_fecper = " & periodo & "", vg_db, adOpenStatic
      Else
         RS2.Open "SELECT TOP 1 (SELECT TOP 1 'FA' FROM b_totcompras a WHERE a.toc_codbod = " & RS1!cli_codbod & "  AND a.toc_numinf = " & RS1!inf_numero & " AND a.toc_tipdoc NOT IN ('FE', 'CE','DE', 'SN') AND a.toc_envsap = '0' AND a.toc_fecper = " & periodo & ") facnor, " & _
                  "(SELECT TOP 1 'FE' FROM b_totcompras b WHERE b.toc_codbod = " & RS1!cli_codbod & " AND b.toc_numinf = " & RS1!inf_numero & " AND b.toc_tipdoc  IN ('FE', 'CE','DE') AND b.toc_envsap = '0' AND b.toc_fecper = " & periodo & ") AS facele, " & _
                  "COUNT(CASE WHEN toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE' THEN 'FE' ELSE 'FA' END) AS nreg " & _
                  "FROM b_totcompras WHERE toc_codbod = " & RS1!cli_codbod & " AND toc_numinf = " & RS1!inf_numero & " AND toc_envsap = '0' AND toc_tipdoc not in ('SN') AND toc_fecper = " & periodo & " GROUP BY toc_tipdoc, toc_tipdoc", vg_db, adOpenStatic
'         RS2.Open "SELECT DISTINCT " & sql1 & " FROM b_totcompras WHERE toc_codbod = " & RS1!cli_codbod & " AND toc_numinf = " & RS1!inf_numero & " GROUP BY toc_tipdoc", vg_db, adOpenStatic
      End If
      If Not RS2.EOF Then
         If RS2!nreg > 1 And (Not IsNull(RS2!facnor) And Not IsNull(RS2!facele) And Trim(RS2!facele) <> "") Then
            RS2.Close: Set RS2 = Nothing
            RS2.Open "SELECT MAX(inf_numero) AS Mayor FROM a_infcfcfofi WHERE inf_cencos = '" & Trim(RS1!inf_cencos) & "' AND inf_tipo = 'C'", vg_db, adOpenStatic
            corre = TipoDato(RS2!mayor, 0) + 1
            RS2.Close: Set RS2 = Nothing
            sql2 = IIf(vg_tipbase = "1", "  val(toc_docaso) ", " convert(int,toc_docaso) ")
            vg_db.Execute "INSERT INTO a_infcfcfofi VALUES ('" & Trim(RS1!inf_cencos) & "', 'C', " & corre & ", 0, NULL)"
            vg_db.Execute "UPDATE b_totcompras SET toc_numinf = " & corre & " WHERE toc_codbod = " & RS1!cli_codbod & " AND toc_numinf = " & RS1!inf_numero & " AND toc_tipdoc IN ('SN') AND " & sql2 & " IN (SELECT DISTINCT toc_numdoc FROM b_totcompras WHERE toc_codbod = " & RS1!cli_codbod & " AND toc_numinf = " & RS1!inf_numero & " AND toc_tipdoc IN ('FE','DE','CE')) AND toc_rutpro IN  (SELECT DISTINCT toc_rutpro FROM b_totcompras WHERE toc_codbod = " & RS1!cli_codbod & " AND toc_numinf = " & RS1!inf_numero & " AND toc_tipdoc IN ('FE','DE','CE'))"
            vg_db.Execute "UPDATE b_totcompras SET toc_numinf = " & corre & " WHERE toc_codbod = " & RS1!cli_codbod & " AND toc_numinf = " & RS1!inf_numero & " AND toc_tipdoc IN ('FE','DE','CE')"
         Else
            RS2.Close: Set RS2 = Nothing
         End If
      Else
         RS2.Close: Set RS2 = Nothing
      End If
      RS1.MoveNext
   Loop
   RS1.Close: Set RS1 = Nothing
   '-------> Agregar nueva opción al sistema
   vg_db.Execute "INSERT INTO a_opcsistema values (4061000, 'Tipo Documento')"
   If vg_tipbase = "2" Then
      BaseDatos = "Actualizador.mdb"
      Set vg_dbsubesql = New ADODB.Connection
      vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
      vg_dbsubesql.ConnectionTimeout = 30
      vg_dbsubesql.CommandTimeout = 600
      vg_dbsubesql.Open
      RS1.Open "SELECT * FROM Procedimiento WHERE Version = '171' order by id1", vg_dbsubesql, adOpenStatic
      If Not RS1.EOF Then
         Do While Not RS1.EOF
            vg_db.Execute ("" & RS1!Procedimiento & "")
            RS1.MoveNext
         Loop
      End If
      RS1.Close: Set RS1 = Nothing
      vg_dbsubesql.Close
   End If

   '-------> Actualiza versión
   vg_db.Execute "UPDATE a_param SET par_valor = '171' WHERE par_codigo = 'version'"
   aVer = 171
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 171 Then
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 172 ....."
   V_Acceso.Refresh
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   RS.Open "SELECT DISTINCT b.cli_codigo, b.cli_codbod  FROM a_param a, b_clientes b WHERE a.par_cencos = b.cli_codigo AND b.cli_codbod > 0 AND a.par_valor = 'CL'", vg_db, adOpenStatic
   Do While Not RS.EOF
        RS1.Open "SELECT * FROM b_cierreperiodo WHERE cie_cencos = '" & RS!cli_codigo & "' AND cie_estado = 0 AND cie_periodo > 201005", vg_db, adOpenStatic
        If Not RS1.EOF Then
           Do While Not RS1.EOF
              '-------> Solicitud nota credito
              sql2 = IIf(vg_tipbase = "1", " 'XXX' & '" & RS1!cie_periodo & "' ", " 'XXX' + '" & RS1!cie_periodo & "' ")
              RS2.Open "SELECT distinct * FROM b_totcompras WHERE toc_codbod = " & RS!cli_codbod & " AND toc_tipdoc = 'SN' AND toc_docsnc = " & sql2 & "", vg_db, adOpenStatic
              If Not RS2.EOF Then
                 Do While Not RS2.EOF
                    vg_db.Execute "UPDATE b_totcompras SET toc_docsnc = '' WHERE toc_rutpro = '" & RS2!toc_rutpro & "' AND toc_numdoc = " & RS2!toc_numdoc & " AND toc_codbod = " & RS2!toc_codbod & " AND toc_tipdoc = 'SN' AND toc_docsnc = " & sql2 & ""
                    RS2.MoveNext
                 Loop
                 RS2.MoveFirst
              End If
              '-------> Guia del mes
              RS3.Open "SELECT distinct toc_fecper, toc_docaso, toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod " & _
                       "From b_totcompras " & _
                       "WHERE toc_fecper > " & RS1!cie_periodo & " AND (toc_docaso) Is Not Null And (toc_docaso) <> '' AND toc_tipdoc IN ('FA','FE', 'CE', 'NC') AND toc_codbod = " & RS!cli_codbod & "", vg_db, adOpenStatic
              If Not RS3.EOF Then
                 Do While Not RS3.EOF
                    sql1 = "": sql1 = "(" & Mid(fg_CambiaChar(RS3!toc_docaso, ";", ","), 1, Len(RS3!toc_docaso) - 1) & ")"
                     vg_db.Execute "UPDATE b_totcompras SET toc_docaso = null WHERE toc_numdoc  In " & sql1 & " AND toc_rutpro = '" & RS3!toc_rutpro & "' AND (toc_tipdoc = 'GD' OR toc_tipdoc = 'SN') AND toc_codbod = " & RS!cli_codbod & ""
                    RS3.MoveNext
                 Loop
                 RS3.MoveFirst
              End If
              vg_codbod = RS!cli_codbod
              vg_contra = Trim(RS!cli_codigo)
              CalcularProvisiones Trim(RS!cli_codigo), RS1!cie_periodo, RS1!cie_fecini, RS1!cie_fecter, True
              '-------> Solicitud nota credito
              If Not RS2.EOF Then
                 RS2.MoveFirst
                 Do While Not RS2.EOF
                    sql2 = IIf(vg_tipbase = "1", " 'XXX' & '" & RS1!cie_periodo & "' ", " 'XXX' + '" & RS1!cie_periodo & "' ")
'                    vg_db.Execute "UPDATE b_totcompras SET toc_docsnc = " & sql2 & " WHERE toc_rutpro = '" & RS2!toc_rutpro & "' AND toc_numdoc = " & RS2!toc_numdoc & " AND toc_codbod = " & RS2!toc_codbod & " AND toc_tipdoc = 'SN' AND toc_docsnc = " & sql2 & ""
                    vg_db.Execute "UPDATE b_totcompras SET toc_docsnc = " & sql2 & " WHERE toc_rutpro = '" & RS2!toc_rutpro & "' AND toc_numdoc = " & RS2!toc_numdoc & " AND toc_codbod = " & RS2!toc_codbod & " AND toc_tipdoc = 'SN' AND " & sql2 & " <> ''"
                    RS2.MoveNext
                 Loop
              End If
              RS2.Close: Set RS2 = Nothing
              '-------> Guia despacho
              If Not RS3.EOF Then
                 RS3.MoveFirst
                 Do While Not RS3.EOF
                    sql1 = "": sql1 = "(" & Mid(fg_CambiaChar(RS3!toc_docaso, ";", ","), 1, Len(RS3!toc_docaso) - 1) & ")"
                    vg_db.Execute "UPDATE b_totcompras SET toc_docaso = '" & RS3!toc_numdoc & "' WHERE toc_numdoc  In " & sql1 & " AND toc_rutpro = '" & RS3!toc_rutpro & "' AND (toc_tipdoc = 'GD' OR toc_tipdoc = 'SN') AND toc_codbod = " & RS!cli_codbod & ""
                    RS3.MoveNext
                 Loop
              End If
              RS3.Close: Set RS3 = Nothing
              RS1.MoveNext
           Loop
        End If
        RS1.Close: Set RS1 = Nothing
        RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
   vg_codbod = 0
   vg_contra = ""
   If vg_tipbase = "2" Then
      
      BaseDatos = "Actualizador.mdb"
      Set vg_dbsubesql = New ADODB.Connection
      vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
      vg_dbsubesql.ConnectionTimeout = 30
      vg_dbsubesql.CommandTimeout = 600
      vg_dbsubesql.Open
      RS1.Open "SELECT * FROM Procedimiento WHERE Version = '172' order by id1", vg_dbsubesql, adOpenStatic
      If Not RS1.EOF Then
         
         Do While Not RS1.EOF
            
            vg_db.Execute ("" & RS1!Procedimiento & "")
            RS1.MoveNext
         
         Loop
      
      End If
      RS1.Close: Set RS1 = Nothing
      vg_dbsubesql.Close
   
   End If
  
   vg_db.Execute "UPDATE a_param SET par_valor = '172' WHERE par_codigo='version'"
   aVer = 172
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
ActVersionI
ActVersionII
Exit Function
Man_Error:
'If Err.Number = -2147467259 Then MsgBox Err & ":  " & "No existe base de datos... " & Chr(13) & "Comunicase con departamento de informatica" & Chr(13) & "Actualización cancelada", vbExclamation + vbOKOnly, "Mantención sistema SGP": End
Resume Next
End Function

Sub ExecProcedimientoObject(ByVal sqlstr As String, ByVal OBJECT_ID As String, Optional ByVal xtype As String = "U", Optional ByVal existe As String = "NOT")
    On Error GoTo e
    Dim sql_aux As String
    
    sql_aux = ""
      sql_aux = " IF " & existe & " EXISTS ( SELECT  1 From sys.objects WHERE   object_id = OBJECT_ID(N'" & OBJECT_ID & "') AND type =( N'" & xtype & "' ) ) "
      sql_aux = sql_aux & vbCrLf & "BEGIN "
      sql_aux = sql_aux & vbCrLf & sqlstr
      sql_aux = sql_aux & vbCrLf & "End "
      
      vg_db.Execute ("" & sql_aux & "")
      Exit Sub
e:
    MsgBox "Error ExecProcedimientoObject: " & Err.Description, vbCritical + vbInformation, "SGP"
    
End Sub

Sub ExecProcedimientoColumna(ByVal sqlstr As String, ByVal table_name As String, ByVal column_name As String, Optional ByVal existe As String = "NOT")
    On Error GoTo e
    Dim sql_aux As String
    
    sql_aux = ""
    sql_aux = " IF " & existe & " EXISTS(SELECT 1 FROM information_schema.[columns] WHERE table_name='" & table_name & "' AND column_name='" & column_name & "') "
    sql_aux = sql_aux & vbCrLf & "begin "
    sql_aux = sql_aux & vbCrLf & " "
    sql_aux = sql_aux & vbCrLf & sqlstr
    sql_aux = sql_aux & vbCrLf & " "
    sql_aux = sql_aux & vbCrLf & "end "
    
    vg_db.Execute ("" & sql_aux & "")
    Exit Sub
e:
    MsgBox "Error ExecProcedimientoColumna: " & Err.Description, vbCritical + vbInformation, "SGP"
    
End Sub

Sub ExecProcedimientoInsertarDatos(ByVal sqlstr As String, ByVal nomtabla As String, ByVal condicion As String)
    On Error GoTo e
    Dim sql_aux As String
    
    sql_aux = ""
    sql_aux = " if not exists(select 1 from " & nomtabla & " where " & condicion & ") "
    sql_aux = sql_aux & vbCrLf & "begin "
    sql_aux = sql_aux & vbCrLf & " "
    sql_aux = sql_aux & vbCrLf & sqlstr
    sql_aux = sql_aux & vbCrLf & " "
    sql_aux = sql_aux & vbCrLf & "end "
    
    vg_db.Execute ("" & sql_aux & "")
    Exit Sub
e:

    MsgBox "Error ExecProcedimientoInsertar: " & Err.Description, vbCritical + vbInformation, "SGP"
    
End Sub
Function ActVersionII()

On Error GoTo Man_Error

If nVer > aVer And aVer = 228 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.2.9....."
   V_Acceso.Refresh
        
   vg_db.Execute "UPDATE a_param SET par_valor = '229' WHERE par_codigo = 'version'"
   aVer = 229
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

'-------> Borrar Actualizador Mdb
If Dir(dir_trabajo & "Actualizador.mdb") <> "" Then Kill dir_trabajo & "Actualizador.mdb"

Set fso = Nothing

Exit Function
Man_Error:

Set fso = Nothing

'If Err.Number = -2147467259 Then
MsgBox "Error... " & Err.Description & Chr(13) & "Comunicase con departamento de informatica", vbExclamation + vbOKOnly, "SGP"
Resume Next
End Function

Function ActVersionI()
On Error GoTo Man_Error
'Dim RS As New ADODB.Recordset
'Dim RS1 As New ADODB.Recordset
'Dim nVer As Long, aVer As Long
'nVer = CLng(App.Major & App.Minor & App.Revision)
'aVer = TipoDato(GetParametro("version"), 0)

If nVer > aVer And aVer = 172 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 173 ....."
   V_Acceso.Refresh

   If vg_tipbase = "1" Then
        vg_db.Execute "CREATE TABLE Log_EnvioMinutasitioremoto (cencos char(10), fecpro DateTime, FecIni DateTime, FecFin DateTime, FecEnv DateTime, FecRec DateTime, Estado Char(1), Mensaje Char(255), RegistroEnvio Char(255), Constraint PK_Log_EnvioMinutasitioremoto Primary Key (cencos,fecpro))"
        vg_db.Execute "CREATE TABLE log_regenviominutasitioremoto (cencos char(10), codreg int, codser int, fecini int, fecfin int, Constraint PK_Log_regenviominutasitioremoto Primary Key (cencos, codreg, codser, fecini, fecfin))"
        vg_db.Execute "ALTER TABLE b_minutadet ADD mid_modmina char(1), mid_modminb char(1)"
        vg_db.Execute "UPDATE b_minutadet SET mid_modmina = '0', mid_modminb = '0'"
   Else
        vg_db.Execute "CREATE TABLE dbo.Log_EnvioMinutasitioremoto (cencos varchar(10), fecpro DateTime, FecIni DateTime, FecFin DateTime, FecEnv DateTime, FecRec DateTime, Estado VarChar(1), Mensaje VarChar(255), RegistroEnvio VarChar(255), Constraint PK_Log_EnvioMinutasitioremoto Primary Key (Cencos,fecpro))"
        vg_db.Execute "CREATE TABLE dbo.log_regenviominutasitioremoto (cencos varchar(10), codreg int, codser int, fecini int, fecfin int, Constraint PK_log_regenviominutasitioremoto Primary Key (cencos, codreg, codser, fecini, fecfin))"
        vg_db.Execute "ALTER TABLE dbo.b_minutadet ADD mid_modmina varchar(1), mid_modminb varchar(1)"
        vg_db.Execute "UPDATE dbo.b_minutadet SET mid_modmina = '0', mid_modminb = '0'"
   End If
     
   If vg_tipbase = "2" Then
      BaseDatos = "Actualizador.mdb"
      Set vg_dbsubesql = New ADODB.Connection
      vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
      vg_dbsubesql.ConnectionTimeout = 30
      vg_dbsubesql.CommandTimeout = 600
      vg_dbsubesql.Open
      RS1.Open "SELECT * FROM Procedimiento WHERE Version = '173' order by id1", vg_dbsubesql, adOpenStatic
      If Not RS1.EOF Then
         Do While Not RS1.EOF
            vg_db.Execute ("" & RS1!Procedimiento & "")
            RS1.MoveNext
         Loop
      End If
      RS1.Close: Set RS1 = Nothing
      vg_dbsubesql.Close
   End If
  
  '-------> Incluir opción informe detalle cartola inventario
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1022000, 'Bloque Minuta')"
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1080000, 'Envio Bloque Minuta')"
  '-------> Actualiza versión
   vg_db.Execute "UPDATE a_param SET par_valor='173' WHERE par_codigo='version'"
   aVer = 173
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 173 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 173.1 ....."
   V_Acceso.Refresh
   '-------> Respaldar tabla log_regenviominutasitioremoto x log_regenviominutasitioremoto_backup
   vg_db.Execute "select * into log_regenviominutasitioremoto_backup from log_regenviominutasitioremoto"
   '-------> crear nuevamente tabla log_regenviominutasitioremoto
   If vg_tipbase = "1" Then
      vg_db.Execute "DROP TABLE log_regenviominutasitioremoto"
      vg_db.Execute "CREATE TABLE log_regenviominutasitioremoto (id int, cencos char(10), codreg int, codser int, fecini int, fecfin int, fecpro datetime)"
   ElseIf vg_tipbase = "2" Then
      vg_db.Execute "DROP TABLE dbo.log_regenviominutasitioremoto"
      vg_db.Execute "CREATE TABLE dbo.log_regenviominutasitioremoto (id bigint NOT NULL, cencos varchar(10), codreg int, codser int, fecini int, fecfin int, fecpro datetime not null default getdate())"
   End If
   If vg_tipbase = "2" Then
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       RS1.Open "SELECT * FROM Procedimiento WHERE Version '1731' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
   End If
   
   '-------> Incluir opción informe Analisis de Consumos Precio Fijo
   vg_db.Execute "INSERT INTO a_opcsistema VALUES (3029000, 'Analisis de Consumos Precio Fijo')"
   '-------> Actualizar nueva versión
   vg_db.Execute "UPDATE a_param SET par_valor = '1731' WHERE par_codigo = 'version'"
   aVer = 1731
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 1731 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.73.2 ....."
   V_Acceso.Refresh

   vg_db.Execute "UPDATE a_param SET par_valor = '1732' WHERE par_codigo = 'version'"
   aVer = 1732
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If

If nVer > aVer And aVer = 1732 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      
      KillProcess ("sgpsdx.exe")
   
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.73.3 ....."
   V_Acceso.Refresh
      
   vg_db.Execute "UPDATE a_param SET par_valor = '1733' WHERE par_codigo = 'version'"
   aVer = 1733
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 1733 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.73.4 ....."
   V_Acceso.Refresh

'  '-------> Incluir opción informe detalle cartola inventario
'   vg_db.Execute "INSERT INTO a_opcsistema VALUES (1022000, 'Bloque Minuta')"
   
   vg_db.Execute "UPDATE a_param SET par_valor = '1734' WHERE par_codigo = 'version'"
   aVer = 1734
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If
If nVer > aVer And aVer = 1734 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      
      KillProcess ("sgpsdx.exe")
   
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.73.5 ....."
   V_Acceso.Refresh
   
   RS.Open "SELECT DISTINCT b.cli_codigo, b.cli_codbod  FROM a_param a, b_clientes b WHERE a.par_cencos = b.cli_codigo AND b.cli_codbod > 0 AND a.par_valor = 'CL'", vg_db, adOpenStatic
   Do While Not RS.EOF
      
      '-------> reprocesar días cierre
      Set RS1 = vg_db.Execute("select * from a_param where par_codigo = 'rprociedia' and par_cencos = '" & RS!cli_codigo & "'")
      If RS1.EOF Then
         vg_db.Execute "insert into a_param values ('rprociedia', 'Reprocesar Cierre Diario', 'C', 'S', '" & RS!cli_codigo & "')"
      End If
      RS1.Close: Set RS1 = Nothing

      '-------> fecha reproceso días cierre
      Set RS1 = vg_db.Execute("select * from a_param where par_codigo = 'fecrprodia' and par_cencos = '" & RS!cli_codigo & "'")
      If RS1.EOF Then
         vg_db.Execute "insert into a_param values ('fecrprodia', 'Fecha Reprocesar Cierre Diario', 'C', '" & fg_Encripta("01/09/2011") & "', '" & RS!cli_codigo & "')"
      End If
      RS1.Close: Set RS1 = Nothing
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   vg_db.Execute "UPDATE a_param SET par_valor = '1735' WHERE par_codigo = 'version'"
   aVer = 1735
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 1735 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      
      KillProcess ("sgpsdx.exe")
   
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.73.6 ....."
   V_Acceso.Refresh
   
   vg_db.Execute "UPDATE a_param SET par_valor = '1736' WHERE par_codigo = 'version'"
   aVer = 1736
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If
If nVer > aVer And aVer = 1736 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      
      KillProcess ("sgpsdx.exe")
   
   End If
   
   If Dir(dir_trabajo & "Actualizador.mdb") <> "" Then
        
        V_Acceso.Label1(1).Visible = True
        V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.73.7 ....."
        V_Acceso.Refresh
        
        If vg_tipbase = "1" Then
           
           vg_db.Execute "ALTER TABLE a_estservicio ADD ess_marcaplatos char(1)"
        
        Else
             Sql = ""
             Sql = "ALTER TABLE dbo.a_estservicio ADD ess_marcaplatos varchar(1)"
             Call ExecProcedimientoColumna(Sql, "a_estservicio", "ess_marcaplatos")
             'vg_db.Execute "ALTER TABLE dbo.a_estservicio ADD ess_marcaplatos varchar(1)"
        End If
        '-------> actualizar estado estructura 0 = no activa plato; 1 = activa plato
        vg_db.Execute "update a_estservicio set ess_marcaplatos = '0'"
        
        If vg_tipbase = "2" Then
            
            BaseDatos = "Actualizador.mdb"
            Set vg_dbsubesql = New ADODB.Connection
            vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
            vg_dbsubesql.ConnectionTimeout = 30
            vg_dbsubesql.CommandTimeout = 600
            vg_dbsubesql.Open
            RS1.Open "SELECT * FROM Procedimiento WHERE Version = '1737' order by id1", vg_dbsubesql, adOpenStatic
            If Not RS1.EOF Then
               
               Do While Not RS1.EOF
                  
                  vg_db.Execute ("" & RS1!Procedimiento & "")
                  RS1.MoveNext
               
               Loop
            
            End If
            RS1.Close: Set RS1 = Nothing
            vg_dbsubesql.Close
        
        End If
        
        vg_db.Execute "UPDATE a_param SET par_valor = '1737' WHERE par_codigo = 'version'"
        aVer = 1737
        V_Acceso.Label1(1).Visible = False
        V_Acceso.Refresh
   End If
End If
If nVer > aVer And aVer = 1737 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.74 ....."
   V_Acceso.Refresh
   'Crear tabla ventas vales
   If vg_tipbase = "1" Then
      vg_db.Execute "CREATE TABLE a_pto_atencion (ate_codatencion int, ate_descripcion Char(100), Constraint PK_a_pto_atencion Primary Key (ate_codatencion))"
      vg_db.Execute "CREATE TABLE b_persona (per_rut char(10), cli_codigo Char(100), per_nombre char(100), per_codbarra char(100), Constraint PK_b_persona Primary Key (per_rut))"
      vg_db.Execute "CREATE TABLE a_pto_lectura_vales (lec_codlecvales int, lec_nombrepc Char(100), lec_ubicacion char(100), lec_activo int, Constraint PK_a_pto_lectura_vales Primary Key (lec_codlecvales))"
      vg_db.Execute "CREATE TABLE a_pto_lectura_vales_servicio (lec_codlecvales int, reg_codigo int, ser_codigo int, ate_codatencion int, Constraint PK_a_pto_lectura_vales_servicio Primary Key (lec_codlecvales, reg_codigo, ser_codigo, ate_codatencion))"
      vg_db.Execute "CREATE TABLE a_pto_lectura_vales_pto_atencion (lec_codlecvales int, ate_codatencion int, Constraint PK_a_pto_lectura_vales_pto_atencion Primary Key (lec_codlecvales, ate_codatencion))"
      vg_db.Execute "CREATE TABLE b_detallelectura (cli_codigo char(10), cli_codigo_rutcliente char(10), reg_codigo int, ser_codigo int, ate_codatencion int, codigobarra char(100), fechahoraregistro datetime, fechahoravale datetime, Constraint PK_b_detallelectura Primary Key (cli_codigo, cli_codigo_rutcliente, reg_codigo, ser_codigo, ate_codatencion, codigobarra, fechahoravale))"
      vg_db.Execute "CREATE TABLE a_par_codigo_barra_cas (cbar_atributo char(50), cli_codigo char(10), cbar_posinicial int, cbar_largo int, Constraint PK_a_par_codigo_barra_cas Primary Key (cbar_atributo, cli_codigo))"
      vg_db.Execute "CREATE TABLE a_par_tipo_vales (ID_Tipo_Vale char(100), cli_codigo char(10), Nombre char(100), Fecha datetime, Descripcion char(100), Largo int, Posicion_inicial int, Constraint PK_a_par_tipo_vales Primary Key (ID_Tipo_Vale, cli_codigo))"
   
      '-------> Crear relaciones
      vg_db.Execute "ALTER TABLE a_pto_lectura_vales_pto_atencion ADD CONSTRAINT FK_a_pto_lectura_vales_pto_atencion_a_pto_atencion FOREIGN KEY(ate_codatencion) References a_pto_atencion (ate_codatencion)"
      vg_db.Execute "ALTER TABLE b_detallelectura ADD CONSTRAINT FK_b_detallelectura_a_pto_atencion FOREIGN KEY(ate_codatencion) References a_pto_atencion (ate_codatencion)"
      vg_db.Execute "ALTER TABLE b_detallelectura ADD CONSTRAINT FK_b_detallelectura_a_servicio FOREIGN KEY(ser_codigo) References a_servicio (ser_codigo)"
      vg_db.Execute "ALTER TABLE b_detallelectura ADD CONSTRAINT FK_b_detallelectura_a_regimen FOREIGN KEY(reg_codigo) References a_regimen (reg_codigo)"
      vg_db.Execute "ALTER TABLE b_detallelectura ADD CONSTRAINT FK_b_detallelectura_b_clientes FOREIGN KEY(cli_codigo) References b_clientes (cli_codigo)"
      vg_db.Execute "ALTER TABLE b_persona ADD CONSTRAINT FK_b_persona_b_clientes FOREIGN KEY(cli_codigo) References b_clientes (cli_codigo)"
      vg_db.Execute "ALTER TABLE a_par_codigo_barra_cas ADD CONSTRAINT FK_a_par_codigo_barra_cas_b_clientes FOREIGN KEY(cli_codigo) References b_clientes (cli_codigo)"
      vg_db.Execute "ALTER TABLE a_pto_lectura_vales_servicio ADD CONSTRAINT FK_a_pto_lectura_vales_servicio_a_pto_lectura_vales FOREIGN KEY(lec_codlecvales) References a_pto_lectura_vales (lec_codlecvales)"
      vg_db.Execute "ALTER TABLE a_pto_lectura_vales_pto_atencion ADD CONSTRAINT FK_a_pto_lectura_vales_pto_atencion_a_pto_lectura_vales FOREIGN KEY(lec_codlecvales) References a_pto_lectura_vales (lec_codlecvales)"
      vg_db.Execute "ALTER TABLE a_pto_lectura_vales_servicio ADD CONSTRAINT FK_a_pto_lectura_vales_servicio_a_servicio FOREIGN KEY(ser_codigo) References a_servicio (ser_codigo)"
      vg_db.Execute "ALTER TABLE a_pto_lectura_vales_servicio ADD CONSTRAINT FK_a_pto_lectura_vales_servicio_a_regimen FOREIGN KEY(reg_codigo) References a_regimen (reg_codigo)"
'      vg_db.Execute "ALTER TABLE a_par_codigo_barra_cas  ADD  CONSTRAINT FK_a_par_codigo_barra_cas_a_par_tipo_vales FOREIGN KEY(id_tipo_vale, cli_codigo) " & _
'                    "References a_par_tipo_vales(ID_Tipo_Vale, cli_codigo)"
      vg_db.Execute "ALTER TABLE a_pto_lectura_vales_servicio ADD CONSTRAINT FK_a_pto_lectura_vales_servicio_a_pto_atencion FOREIGN KEY(ate_codatencion) References a_pto_atencion (ate_codatencion)"
   Else
   
    Sql = "CREATE TABLE dbo.a_pto_atencion (ate_codatencion int, ate_descripcion varchar(100), Constraint PK_a_pto_atencion Primary Key (ate_codatencion)) "
    Call ExecProcedimientoObject(Sql, "a_pto_atencion")
    Sql = "CREATE TABLE dbo.b_persona (per_rut varchar(10), cli_codigo varchar(10), per_nombre varchar(100), per_codbarra varchar(100), Constraint PK_b_persona Primary Key (per_rut)) "
    Call ExecProcedimientoObject(Sql, "b_persona")
    Sql = "CREATE TABLE dbo.a_pto_lectura_vales (lec_codlecvales int, lec_nombrepc varchar(100), lec_ubicacion varchar(100), lec_activo bit, Constraint PK_a_pto_lectura_vales Primary Key (lec_codlecvales)) "
    Call ExecProcedimientoObject(Sql, "a_pto_lectura_vales")
    Sql = "CREATE TABLE dbo.a_pto_lectura_vales_servicio (lec_codlecvales int, reg_codigo int, ser_codigo int, ate_codatencion int, Constraint PK_a_pto_lectura_vales_servicio Primary Key (lec_codlecvales, reg_codigo, ser_codigo, ate_codatencion)) "
    Call ExecProcedimientoObject(Sql, "a_pto_lectura_vales_servicio")
    Sql = "CREATE TABLE dbo.a_pto_lectura_vales_pto_atencion (lec_codlecvales int, ate_codatencion int, Constraint PK_a_pto_lectura_vales_pto_atencion Primary Key (lec_codlecvales, ate_codatencion)) "
    Call ExecProcedimientoObject(Sql, "a_pto_lectura_vales_pto_atencion")
    Sql = "CREATE TABLE dbo.b_detallelectura (cli_codigo varchar(10), cli_codigo_rutcliente varchar(10), reg_codigo int, ser_codigo int, ate_codatencion int, codigobarra varchar(100), fechahoraregistro datetime not null default getdate(), fechahoravale datetime not null default getdate(), Constraint PK_b_detallelectura Primary Key (cli_codigo, cli_codigo_rutcliente, reg_codigo, ser_codigo, ate_codatencion, codigobarra, fechahoravale)) "
    Call ExecProcedimientoObject(Sql, "b_detallelectura")
    Sql = "CREATE TABLE dbo.a_par_codigo_barra_cas (cbar_atributo varchar(50), cli_codigo varchar(10), cbar_posinicial int, cbar_largo int, Constraint PK_a_par_codigo_barra_cas Primary Key (cbar_atributo, cli_codigo)) "
    Call ExecProcedimientoObject(Sql, "a_par_codigo_barra_cas")
    Sql = "CREATE TABLE dbo.a_par_tipo_vales (ID_Tipo_Vale varchar(100), cli_codigo varchar(10), Nombre varchar(100), Fecha datetime, Descripcion varchar(100), Largo int, Posicion_inicial int, Constraint PK_a_par_tipo_vales Primary Key (ID_Tipo_Vale, cli_codigo)) "
    Call ExecProcedimientoObject(Sql, "a_par_tipo_vales")
   
      'vg_db.Execute "CREATE TABLE dbo.a_pto_atencion (ate_codatencion int, ate_descripcion varchar(100), Constraint PK_a_pto_atencion Primary Key (ate_codatencion))"
      'vg_db.Execute "CREATE TABLE dbo.b_persona (per_rut varchar(10), cli_codigo varchar(10), per_nombre varchar(100), per_codbarra varchar(100), Constraint PK_b_persona Primary Key (per_rut))"
      'vg_db.Execute "CREATE TABLE dbo.a_pto_lectura_vales (lec_codlecvales int, lec_nombrepc varchar(100), lec_ubicacion varchar(100), lec_activo bit, Constraint PK_a_pto_lectura_vales Primary Key (lec_codlecvales))"
      'vg_db.Execute "CREATE TABLE dbo.a_pto_lectura_vales_servicio (lec_codlecvales int, reg_codigo int, ser_codigo int, ate_codatencion int, Constraint PK_a_pto_lectura_vales_servicio Primary Key (lec_codlecvales, reg_codigo, ser_codigo, ate_codatencion))"
      'vg_db.Execute "CREATE TABLE dbo.a_pto_lectura_vales_pto_atencion (lec_codlecvales int, ate_codatencion int, Constraint PK_a_pto_lectura_vales_pto_atencion Primary Key (lec_codlecvales, ate_codatencion))"
      'vg_db.Execute "CREATE TABLE dbo.b_detallelectura (cli_codigo varchar(10), cli_codigo_rutcliente varchar(10), reg_codigo int, ser_codigo int, ate_codatencion int, codigobarra varchar(100), fechahoraregistro datetime not null default getdate(), fechahoravale datetime not null default getdate(), Constraint PK_b_detallelectura Primary Key (cli_codigo, cli_codigo_rutcliente, reg_codigo, ser_codigo, ate_codatencion, codigobarra, fechahoravale))"
      'vg_db.Execute "CREATE TABLE dbo.a_par_codigo_barra_cas (cbar_atributo varchar(50), cli_codigo varchar(10), cbar_posinicial int, cbar_largo int, Constraint PK_a_par_codigo_barra_cas Primary Key (cbar_atributo, cli_codigo))"
      'vg_db.Execute "CREATE TABLE dbo.a_par_tipo_vales (ID_Tipo_Vale varchar(100), cli_codigo varchar(10), Nombre varchar(100), Fecha datetime, Descripcion varchar(100), Largo int, Posicion_inicial int, Constraint PK_a_par_tipo_vales Primary Key (ID_Tipo_Vale, cli_codigo))"
      
      '-------> Crear relaciones
    Sql = "ALTER TABLE dbo.a_pto_lectura_vales_pto_atencion ADD CONSTRAINT FK_a_pto_lectura_vales_pto_atencion_a_pto_atencion FOREIGN KEY(ate_codatencion) References dbo.a_pto_atencion (ate_codatencion) "
    Call ExecProcedimientoObject(Sql, "FK_a_pto_lectura_vales_pto_atencion_a_pto_atencion", "F")
    Sql = "ALTER TABLE dbo.b_detallelectura ADD CONSTRAINT FK_b_detallelectura_a_pto_atencion FOREIGN KEY(ate_codatencion) References dbo.a_pto_atencion (ate_codatencion) "
    Call ExecProcedimientoObject(Sql, "FK_b_detallelectura_a_pto_atencion", "F")
    Sql = "ALTER TABLE dbo.b_detallelectura ADD CONSTRAINT FK_b_detallelectura_a_servicio FOREIGN KEY(ser_codigo) References dbo.a_servicio (ser_codigo) "
    Call ExecProcedimientoObject(Sql, "FK_b_detallelectura_a_servicio", "F")
    Sql = "ALTER TABLE dbo.b_detallelectura ADD CONSTRAINT FK_b_detallelectura_a_regimen FOREIGN KEY(reg_codigo) References dbo.a_regimen (reg_codigo) "
    Call ExecProcedimientoObject(Sql, "FK_b_detallelectura_a_regimen", "F")
    Sql = "ALTER TABLE dbo.b_detallelectura ADD CONSTRAINT FK_b_detallelectura_b_clientes FOREIGN KEY(cli_codigo) References dbo.b_clientes (cli_codigo) "
    Call ExecProcedimientoObject(Sql, "FK_b_detallelectura_b_clientes", "F")
    Sql = "ALTER TABLE dbo.b_persona ADD CONSTRAINT FK_b_persona_b_clientes FOREIGN KEY(cli_codigo) References dbo.b_clientes (cli_codigo) "
    Call ExecProcedimientoObject(Sql, "FK_b_persona_b_clientes", "F")
    Sql = "ALTER TABLE dbo.a_par_codigo_barra_cas ADD CONSTRAINT FK_a_par_codigo_barra_cas_b_clientes FOREIGN KEY(cli_codigo) References dbo.b_clientes (cli_codigo) "
    Call ExecProcedimientoObject(Sql, "FK_a_par_codigo_barra_cas_b_clientes", "F")
    Sql = "ALTER TABLE dbo.a_pto_lectura_vales_servicio ADD CONSTRAINT FK_a_pto_lectura_vales_servicio_a_pto_lectura_vales FOREIGN KEY(lec_codlecvales) References dbo.a_pto_lectura_vales (lec_codlecvales) "
    Call ExecProcedimientoObject(Sql, "FK_a_pto_lectura_vales_servicio_a_pto_lectura_vales", "F")
    Sql = "ALTER TABLE dbo.a_pto_lectura_vales_pto_atencion ADD CONSTRAINT FK_a_pto_lectura_vales_pto_atencion_a_pto_lectura_vales FOREIGN KEY(lec_codlecvales) References dbo.a_pto_lectura_vales (lec_codlecvales) "
    Call ExecProcedimientoObject(Sql, "FK_a_pto_lectura_vales_pto_atencion_a_pto_lectura_vales", "F")
    Sql = "ALTER TABLE dbo.a_pto_lectura_vales_servicio ADD CONSTRAINT FK_a_pto_lectura_vales_servicio_a_servicio FOREIGN KEY(ser_codigo) References dbo.a_servicio (ser_codigo) "
    Call ExecProcedimientoObject(Sql, "FK_a_pto_lectura_vales_servicio_a_servicio", "F")
    Sql = "ALTER TABLE dbo.a_pto_lectura_vales_servicio ADD CONSTRAINT FK_a_pto_lectura_vales_servicio_a_regimen FOREIGN KEY(reg_codigo) References dbo.a_regimen (reg_codigo) "
    Call ExecProcedimientoObject(Sql, "FK_a_pto_lectura_vales_servicio_a_regimen", "F")
    Sql = "ALTER TABLE dbo.a_pto_lectura_vales_servicio ADD CONSTRAINT FK_a_pto_lectura_vales_servicio_a_pto_atencion FOREIGN KEY(ate_codatencion) References dbo.a_pto_atencion (ate_codatencion) "
    Call ExecProcedimientoObject(Sql, "FK_a_pto_lectura_vales_servicio_a_pto_atencion", "F")
      
      'vg_db.Execute "ALTER TABLE dbo.a_pto_lectura_vales_pto_atencion ADD CONSTRAINT FK_a_pto_lectura_vales_pto_atencion_a_pto_atencion FOREIGN KEY(ate_codatencion) References dbo.a_pto_atencion (ate_codatencion)"
      'vg_db.Execute "ALTER TABLE dbo.b_detallelectura ADD CONSTRAINT FK_b_detallelectura_a_pto_atencion FOREIGN KEY(ate_codatencion) References dbo.a_pto_atencion (ate_codatencion)"
      'vg_db.Execute "ALTER TABLE dbo.b_detallelectura ADD CONSTRAINT FK_b_detallelectura_a_servicio FOREIGN KEY(ser_codigo) References dbo.a_servicio (ser_codigo)"
      'vg_db.Execute "ALTER TABLE dbo.b_detallelectura ADD CONSTRAINT FK_b_detallelectura_a_regimen FOREIGN KEY(reg_codigo) References dbo.a_regimen (reg_codigo)"
      'vg_db.Execute "ALTER TABLE dbo.b_detallelectura ADD CONSTRAINT FK_b_detallelectura_b_clientes FOREIGN KEY(cli_codigo) References dbo.b_clientes (cli_codigo)"
      'vg_db.Execute "ALTER TABLE dbo.b_persona ADD CONSTRAINT FK_b_persona_b_clientes FOREIGN KEY(cli_codigo) References dbo.b_clientes (cli_codigo)"
      'vg_db.Execute "ALTER TABLE dbo.a_par_codigo_barra_cas ADD CONSTRAINT FK_a_par_codigo_barra_cas_b_clientes FOREIGN KEY(cli_codigo) References dbo.b_clientes (cli_codigo)"
      'vg_db.Execute "ALTER TABLE dbo.a_pto_lectura_vales_servicio ADD CONSTRAINT FK_a_pto_lectura_vales_servicio_a_pto_lectura_vales FOREIGN KEY(lec_codlecvales) References dbo.a_pto_lectura_vales (lec_codlecvales)"
      'vg_db.Execute "ALTER TABLE dbo.a_pto_lectura_vales_pto_atencion ADD CONSTRAINT FK_a_pto_lectura_vales_pto_atencion_a_pto_lectura_vales FOREIGN KEY(lec_codlecvales) References dbo.a_pto_lectura_vales (lec_codlecvales)"
      'vg_db.Execute "ALTER TABLE dbo.a_pto_lectura_vales_servicio ADD CONSTRAINT FK_a_pto_lectura_vales_servicio_a_servicio FOREIGN KEY(ser_codigo) References dbo.a_servicio (ser_codigo)"
      'vg_db.Execute "ALTER TABLE dbo.a_pto_lectura_vales_servicio ADD CONSTRAINT FK_a_pto_lectura_vales_servicio_a_regimen FOREIGN KEY(reg_codigo) References dbo.a_regimen (reg_codigo)"
'      vg_db.Execute "ALTER TABLE dbo.a_par_codigo_barra_cas  ADD  CONSTRAINT FK_a_par_codigo_barra_cas_a_par_tipo_vales FOREIGN KEY(id_tipo_vale, cli_codigo) " & _
'                    "References DBO.a_par_tipo_vales(ID_Tipo_Vale, cli_codigo)"
      'vg_db.Execute "ALTER TABLE dbo.a_pto_lectura_vales_servicio ADD CONSTRAINT FK_a_pto_lectura_vales_servicio_a_pto_atencion FOREIGN KEY(ate_codatencion) References dbo.a_pto_atencion (ate_codatencion)"
   End If
   
   '-------> Agregar nueva opción al sistema
    Sql = "INSERT INTO a_opcsistema values (6010000, 'Lectura Vale - Punto Atención') "
    Call ExecProcedimientoInsertarDatos(Sql, " a_opcsistema ", " opc_codigo = '6010000'")
    Sql = "INSERT INTO a_opcsistema values (6020000, 'Lectura Vale - Personal') "
    Call ExecProcedimientoInsertarDatos(Sql, " a_opcsistema ", " opc_codigo = '6020000'")
    Sql = "INSERT INTO a_opcsistema values (6030000, 'Lectura Vale - Punto Lectura de Vales') "
    Call ExecProcedimientoInsertarDatos(Sql, " a_opcsistema ", " opc_codigo = '6030000'")
    Sql = "INSERT INTO a_opcsistema values (6080000, 'Lectura Vale - Lectura de Vales') "
    Call ExecProcedimientoInsertarDatos(Sql, " a_opcsistema ", " opc_codigo = '6080000'")
    Sql = "INSERT INTO a_opcsistema values (6150000, 'Lectura Vale - Reporte Generico') "
    Call ExecProcedimientoInsertarDatos(Sql, " a_opcsistema ", " opc_codigo = '6150000'")
   
   'vg_db.Execute "INSERT INTO a_opcsistema values (6010000, 'Lectura Vale - Punto Atención')"
   'vg_db.Execute "INSERT INTO a_opcsistema values (6020000, 'Lectura Vale - Personal')"
   'vg_db.Execute "INSERT INTO a_opcsistema values (6030000, 'Lectura Vale - Punto Lectura de Vales')"
   'vg_db.Execute "INSERT INTO a_opcsistema values (6080000, 'Lectura Vale - Lectura de Vales')"
   'vg_db.Execute "INSERT INTO a_opcsistema values (6150000, 'Lectura Vale - Reporte Generico')"
   vg_db.Execute "UPDATE a_param SET par_valor = '174' WHERE par_codigo = 'version'"
   aVer = 174
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If

If nVer > aVer And aVer = 174 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.75 ....."
   V_Acceso.Refresh

   If vg_tipbase = "1" Then
      '-------> Borrar relación tabla a_par_codigo_barra_cas - a_par_tipo_vales - b_clientes
      vg_db.Execute "ALTER TABLE a_par_codigo_barra_cas DROP CONSTRAINT FK_a_par_codigo_barra_cas_a_par_tipo_vales"
      vg_db.Execute "ALTER TABLE a_par_codigo_barra_cas DROP CONSTRAINT FK_a_par_codigo_barra_cas_b_clientes"
      
      '-------> Borrar campo tabla b_clientes
      vg_db.Execute "alter table b_clientes drop column id_tipo_vale"
    
      '-------> Eliminar tabla a_par_tipo_vales - a_par_codigo_barra_cas
      vg_db.Execute "DROP TABLE a_par_tipo_vales"
      vg_db.Execute "DROP TABLE a_par_codigo_barra_cas"
      
      '-------> Crear nueva estructura tabla a_par_codigo_barra_cas
      vg_db.Execute "CREATE TABLE a_par_codigo_barra_cas (cbar_atributo char(50), cli_codigo char(10), cbar_posinicial int, cbar_largo int, Constraint PK_a_par_codigo_barra_cas Primary Key (cbar_atributo, cli_codigo))"
      
      '-------> Crear relación tabla a_par_codigo_barra_cas - b_clientes
      vg_db.Execute "ALTER TABLE a_par_codigo_barra_cas ADD CONSTRAINT FK_a_par_codigo_barra_cas_b_clientes FOREIGN KEY(cli_codigo) References b_clientes (cli_codigo)"
   Else
      '-------> Borrar relación tabla a_par_codigo_barra_cas - a_par_tipo_vales - b_clientes
      Sql = "ALTER TABLE dbo.a_par_codigo_barra_cas DROP CONSTRAINT FK_a_par_codigo_barra_cas_a_par_tipo_vales "
      Call ExecProcedimientoObject(Sql, "FK_a_par_codigo_barra_cas_a_par_tipo_vales", "F", "")
      Sql = "ALTER TABLE dbo.a_par_codigo_barra_cas DROP CONSTRAINT FK_a_par_codigo_barra_cas_b_clientes "
      Call ExecProcedimientoObject(Sql, "FK_a_par_codigo_barra_cas_b_clientes", "F", "")
      
      'vg_db.Execute "ALTER TABLE dbo.a_par_codigo_barra_cas DROP CONSTRAINT FK_a_par_codigo_barra_cas_a_par_tipo_vales"
'      vg_db.Execute "ALTER TABLE dbo.a_par_codigo_barra_cas DROP CONSTRAINT FK_a_par_codigo_barra_cas_b_clientes"
      
      '-------> Borrar campo tabla b_clientes
      Sql = "alter table dbo.b_clientes drop column id_tipo_vale "
      Call ExecProcedimientoColumna(Sql, "b_clientes", "id_tipo_vale", "")
      'vg_db.Execute "alter table dbo.b_clientes drop column id_tipo_vale"
      
      '-------> Eliminar tabla a_par_tipo_vales - a_par_codigo_barra_cas
      Sql = "DROP TABLE dbo.a_par_tipo_vales "
      Call ExecProcedimientoObject(Sql, "a_par_tipo_vales", "")
      Sql = "DROP TABLE dbo.a_par_codigo_barra_cas "
      Call ExecProcedimientoObject(Sql, "a_par_codigo_barra_cas", "")
      
      'vg_db.Execute "DROP TABLE dbo.a_par_tipo_vales"
      'vg_db.Execute "DROP TABLE dbo.a_par_codigo_barra_cas"
      
      '-------> Crear nueva estructura tabla a_par_codigo_barra_cas
      Sql = "CREATE TABLE dbo.a_par_codigo_barra_cas (cbar_atributo varchar(50), cli_codigo varchar(10), cbar_posinicial int, cbar_largo int, Constraint PK_a_par_codigo_barra_cas Primary Key (cbar_atributo, cli_codigo)) "
      Call ExecProcedimientoObject(Sql, "a_par_codigo_barra_cas")
      
      'vg_db.Execute "CREATE TABLE dbo.a_par_codigo_barra_cas (cbar_atributo varchar(50), cli_codigo varchar(10), cbar_posinicial int, cbar_largo int, Constraint PK_a_par_codigo_barra_cas Primary Key (cbar_atributo, cli_codigo))"
      
      '-------> Crear relación tabla a_par_codigo_barra_cas - b_clientes
      Sql = "ALTER TABLE dbo.a_par_codigo_barra_cas ADD CONSTRAINT FK_a_par_codigo_barra_cas_b_clientes FOREIGN KEY(cli_codigo) References dbo.b_clientes (cli_codigo) "
      Call ExecProcedimientoObject(Sql, "FK_a_par_codigo_barra_cas_b_clientes", "F")
      
      'vg_db.Execute "ALTER TABLE dbo.a_par_codigo_barra_cas ADD CONSTRAINT FK_a_par_codigo_barra_cas_b_clientes FOREIGN KEY(cli_codigo) References dbo.b_clientes (cli_codigo)"
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '175' WHERE par_codigo = 'version'"
   aVer = 175
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 175 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.76 ....."
   V_Acceso.Refresh
   If vg_tipbase = "2" Then
      
      Sql = ""
      Sql = " IF not EXISTS ( SELECT  * From sys.objects WHERE   object_id = OBJECT_ID(N'[dbo].[b_formatocompras_sap]') AND type =( N'U' ) ) "
      Sql = Sql & VgLinea & "BEGIN "
      Sql = Sql & VgLinea & "CREATE TABLE dbo.b_formatocompras_sap (fcs_CodMaterial varchar(20), fcs_DenMaterial varchar(100), fcs_CodGrpArt varchar(9), fcs_DenGrpArt varchar(60), fcs_CodUniMed varchar(3), fcs_DenUniMed varchar(30), fcs_FechaCreacion int, fcs_flexec varchar(2), fcs_vigfin int, fcs_faccon float, fcs_fecmodsap int, fcs_fecmodsgp datetime NOT NULL DEFAULT (getdate()), fcs_ctacon varchar(10), Constraint PK_b_formatocompras_sap Primary Key (fcs_CodMaterial)) "
      Sql = Sql & VgLinea & "End "
      vg_db.Execute ("" & Sql & "")
      
'      vg_db.Execute "CREATE TABLE dbo.b_formatocompras_sap (fcs_CodMaterial varchar(20), fcs_DenMaterial varchar(100), fcs_CodGrpArt varchar(9), fcs_DenGrpArt varchar(60), fcs_CodUniMed varchar(3), fcs_DenUniMed varchar(30), fcs_FechaCreacion int, fcs_flexec varchar(2), fcs_vigfin int, fcs_faccon float, fcs_fecmodsap int, fcs_fecmodsgp datetime NOT NULL DEFAULT (getdate()), fcs_ctacon varchar(10), Constraint PK_b_formatocompras_sap Primary Key (fcs_CodMaterial))"
      
      Sql = ""
      Sql = " IF not EXISTS ( SELECT  * From sys.objects WHERE   object_id = OBJECT_ID(N'[dbo].[b_formatocompras_sap_sgp]') AND type =( N'U' ) ) "
      Sql = Sql & VgLinea & "BEGIN "
      Sql = Sql & VgLinea & "CREATE TABLE dbo.b_formatocompras_sap_sgp (fss_CodMaterial varchar(20), fss_CodSgp varchar(20), fss_SgpPre int, fss_fecmod datetime NOT NULL DEFAULT (getdate()), Constraint PK_b_formatocompras_sap_sgp Primary Key (fss_CodMaterial, fss_CodSgp)) "
      Sql = Sql & VgLinea & "End "
      vg_db.Execute ("" & Sql & "")
      
'      vg_db.Execute "CREATE TABLE dbo.b_formatocompras_sap_sgp (fss_CodMaterial varchar(20), fss_CodSgp varchar(20), fss_SgpPre int, fss_fecmod datetime NOT NULL DEFAULT (getdate()), Constraint PK_b_formatocompras_sap_sgp Primary Key (fss_CodMaterial, fss_CodSgp))"
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '176' WHERE par_codigo = 'version'"
   aVer = 176
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 176 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.77 ....."
   V_Acceso.Refresh
   '------- Incluir opción parametro password limpia base de dato
   RS1.Open "SELECT * FROM b_clientes WHERE cli_tipo=0 AND cli_codbod>0", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         '------- Incluir opción parametro password limpia base de dato
         Sql = ""
         Sql = "INSERT INTO a_param VALUES ('pasminblo', 'Password Minuta Bloque', 'C', '" & fg_Encripta("sdxo2008*") & "', '" & RS1!cli_codigo & "') "
         Call ExecProcedimientoInsertarDatos(Sql, "a_param", "par_codigo = 'pasminblo' and par_cencos = '" & RS1!cli_codigo & "'")
         'vg_db.Execute "INSERT INTO a_param VALUES ('pasminblo', 'Password Minuta Bloque', 'C', '" & fg_Encripta("sdxo2008*") & "', '" & RS1!cli_codigo & "')"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
      
   '-------> Crear tabla
    Sql = ""
    Sql = " IF not EXISTS ( SELECT  * From sys.objects WHERE   object_id = OBJECT_ID(N'[dbo].[Carga_Ruta_Compra]') AND type =( N'U' ) ) "
    Sql = Sql & VgLinea & "BEGIN "
    Sql = Sql & VgLinea & "CREATE TABLE [dbo].[Carga_Ruta_Compra] ([id_carga] [bigint] IDENTITY(1,1) NOT NULL, [fecha] [datetime] not null default getdate(), [usuario] [varchar](100), Constraint PK_Carga_Ruta_Compra Primary Key (id_carga)) "
    Sql = Sql & VgLinea & "End "
    vg_db.Execute ("" & Sql & "")
   
'   vg_db.Execute "CREATE TABLE [dbo].[Carga_Ruta_Compra] ([id_carga] [bigint] IDENTITY(1,1) NOT NULL, [fecha] [datetime] not null default getdate(), [usuario] [varchar](100), Constraint PK_Carga_Ruta_Compra Primary Key (id_carga))"
    Sql = ""
    Sql = " IF not EXISTS ( SELECT  * From sys.objects WHERE   object_id = OBJECT_ID(N'[dbo].[ruta_compras]') AND type =( N'U' ) ) "
    Sql = Sql & VgLinea & "BEGIN "
    Sql = Sql & VgLinea & "CREATE TABLE [dbo].[ruta_compras](  [ID_ruta_compra] [int] IDENTITY(1,1) NOT NULL, [Fecha_despacho] [varchar](100) NULL, [ID_centro_de_costo] [varchar](100) NULL, [Familia_producto] [varchar](100) NULL, [ID_proveedor] [varchar](100) NULL, [Sucursal] [varchar](100) NULL, [Sigla_de_ruta] [varchar](100) NULL, [Descripcion_sigla] [varchar](100) NULL, [Observaciones] [varchar](100) NULL, [Ruta_archivo] [varchar](100) NULL , [id_carga] [bigint] NULL, Constraint PK_ruta_compras Primary Key (ID_ruta_compra)) "
    Sql = Sql & VgLinea & "End "
    vg_db.Execute ("" & Sql & "")
   
'   vg_db.Execute "CREATE TABLE [dbo].[ruta_compras](  [ID_ruta_compra] [int] IDENTITY(1,1) NOT NULL, [Fecha_despacho] [varchar](100) NULL, [ID_centro_de_costo] [varchar](100) NULL, [Familia_producto] [varchar](100) NULL, [ID_proveedor] [varchar](100) NULL, [Sucursal] [varchar](100) NULL, [Sigla_de_ruta] [varchar](100) NULL, [Descripcion_sigla] [varchar](100) NULL, [Observaciones] [varchar](100) NULL, [Ruta_archivo] [varchar](100) NULL , [id_carga] [bigint] NULL, Constraint PK_ruta_compras Primary Key (ID_ruta_compra))"
    Sql = ""
    Sql = " IF not EXISTS ( SELECT  * From sys.objects WHERE   object_id = OBJECT_ID(N'[dbo].[convenios_mvi]') AND type =( N'U' ) ) "
    Sql = Sql & VgLinea & "BEGIN "
    Sql = Sql & VgLinea & "CREATE TABLE [dbo].[convenios_mvi] ( [Reg_info] [varchar](100) NULL, [Proveedor] [varchar](100) NULL, [Denominacion_Proveedor] [varchar](100) NULL, [Material] [varchar](100) NULL, [Denominacion_Material] [varchar](100) NULL, [OrgC] [varchar](100) NULL, [Denominacion1] [varchar](100) NULL, [Ce] [varchar](100) NULL, [Denominacion_Centro] [varchar](100) NULL, [GCp] [varchar](100) NULL, [Denominacion2] [varchar](100) NULL, [Prec_neto] [varchar](100) NULL, [Mon] [varchar](100) NULL, [Ctd_mn] [varchar](100) NULL, [Tipo_de_Co] [varchar](100) NULL, [Perfil_de] [varchar](100) NULL, [PzE] [varchar](100) NULL, [Plazo_Anul] [varchar](100) NULL, [Valido_de] [varchar](100) NULL, [Validez_a] [varchar](100) NULL, [Importe] [varchar](100) NULL, [Un] [varchar](100) NULL, [por] [varchar](100) NULL, [UM] [varchar](100) NULL, [B] [varchar](100) NULL, Ruta_archivo [varchar](100) NULL) "
    Sql = Sql & VgLinea & "End "
    vg_db.Execute ("" & Sql & "")
   
'   vg_db.Execute "CREATE TABLE [dbo].[convenios_mvi] ( [Reg_info] [varchar](100) NULL, [Proveedor] [varchar](100) NULL, [Denominacion_Proveedor] [varchar](100) NULL, [Material] [varchar](100) NULL, [Denominacion_Material] [varchar](100) NULL, [OrgC] [varchar](100) NULL, [Denominacion1] [varchar](100) NULL, [Ce] [varchar](100) NULL, [Denominacion_Centro] [varchar](100) NULL, [GCp] [varchar](100) NULL, [Denominacion2] [varchar](100) NULL, [Prec_neto] [varchar](100) NULL, [Mon] [varchar](100) NULL, [Ctd_mn] [varchar](100) NULL, [Tipo_de_Co] [varchar](100) NULL, [Perfil_de] [varchar](100) NULL, [PzE] [varchar](100) NULL, [Plazo_Anul] [varchar](100) NULL, [Valido_de] [varchar](100) NULL, [Validez_a] [varchar](100) NULL, [Importe] [varchar](100) NULL, [Un] [varchar](100) NULL, [por] [varchar](100) NULL, [UM] [varchar](100) NULL, [B] [varchar](100) NULL, Ruta_archivo [varchar](100) NULL)"

   '-------> Crear relaciones
   Sql = ""
   Sql = " IF not EXISTS (SELECT 1 FROM sys.sysobjects AS S WHERE name ='FK_ruta_compras_Carga_Ruta_Compra') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "ALTER TABLE dbo.ruta_compras ADD CONSTRAINT FK_ruta_compras_Carga_Ruta_Compra FOREIGN KEY(id_carga) References dbo.Carga_Ruta_Compra (id_carga) "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "end "
   vg_db.Execute ("" & Sql & "")
   
'   vg_db.Execute "ALTER TABLE dbo.ruta_compras ADD CONSTRAINT FK_ruta_compras_Carga_Ruta_Compra FOREIGN KEY(id_carga) References dbo.Carga_Ruta_Compra (id_carga)"
      
   If vg_tipbase = "2" Then
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       RS1.Open "SELECT * FROM Procedimiento where Version = '177' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
   End If
   
   '-------> Actualizar versión
   vg_db.Execute "UPDATE a_param SET par_valor = '177' WHERE par_codigo = 'version'"
   aVer = 177
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 177 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.7.8 ....."
   V_Acceso.Refresh
   
   Sql = ""
   Sql = " IF NOT EXISTS(SELECT * FROM information_schema.[columns] WHERE table_name='a_estservicio' AND column_name='ess_marcaplatos') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & " ALTER TABLE a_estservicio ADD ess_marcaplatos varchar(1) "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & " UPDATE a_estservicio set ess_marcaplatos = '0'"
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "end "
   vg_db.Execute ("" & Sql & "")
   
   vg_db.Execute "UPDATE a_param SET par_valor = '178' WHERE par_codigo = 'version'"
   aVer = 178
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 178 Then
   
   If Dir(dir_trabajo & "Actualizador.mdb") = "" Then
      MsgBox "No existe Actualizador.mdb, para actualizar la nueva versión " & aVer & VgLinea & "Vuelva bajar la actualización y actualice o bien comuniquese con su monitor" & VgLinea & "        Proceso cancelado ...", vbCritical + vbOKOnly, "SGP"
      End
   End If
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.7.9 ....."
   V_Acceso.Refresh
   
    Sql = ""
    Sql = " IF EXISTS ( SELECT  * From sys.objects WHERE   object_id = OBJECT_ID(N'[dbo].[RUTA_PEDIDO_FINAL]') AND type IN ( N'U' ) ) "
    Sql = Sql & VgLinea & "BEGIN "
    Sql = Sql & VgLinea & "DROP TABLE dbo.RUTA_PEDIDO_FINAL "
    Sql = Sql & VgLinea & "End "
    
    vg_db.Execute ("" & Sql & "")
  
    Sql = ""
    Sql = " IF EXISTS ( SELECT  * From sys.objects WHERE   object_id = OBJECT_ID(N'[dbo].[PEDIDO_TEMP_FINAL ]') AND type IN ( N'U' ) ) "
    Sql = Sql & VgLinea & "BEGIN "
    Sql = Sql & VgLinea & "DROP TABLE dbo.PEDIDO_TEMP_FINAL "
    Sql = Sql & VgLinea & "End "
    
    vg_db.Execute ("" & Sql & "")
    
    Sql = ""
    Sql = " IF EXISTS ( SELECT  * From sys.objects WHERE   object_id = OBJECT_ID(N'[dbo].[ruta_temp_CALCULO]') AND type IN ( N'U' ) ) "
    Sql = Sql & VgLinea & "BEGIN "
    Sql = Sql & VgLinea & "DROP TABLE dbo.ruta_temp_CALCULO "
    Sql = Sql & VgLinea & "End "

    vg_db.Execute ("" & Sql & "")
   
   If vg_tipbase = "2" Then
      BaseDatos = "Actualizador.mdb"
      Set vg_dbsubesql = New ADODB.Connection
      vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
      vg_dbsubesql.ConnectionTimeout = 30
      vg_dbsubesql.CommandTimeout = 600
      vg_dbsubesql.Open
      RS1.Open "SELECT * FROM Procedimiento WHERE Version = '179' order by id1", vg_dbsubesql, adOpenStatic
      If Not RS1.EOF Then
         Do While Not RS1.EOF
            vg_db.Execute ("" & RS1!Procedimiento & "")
            RS1.MoveNext
         Loop
      End If
      RS1.Close: Set RS1 = Nothing
      vg_dbsubesql.Close
   End If

  
   Sql = ""
   Sql = " IF NOT EXISTS(SELECT * FROM information_schema.[columns] WHERE table_name='log_regenviominutasitioremoto' AND column_name='id') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & "DROP TABLE dbo.log_regenviominutasitioremoto "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "CREATE TABLE dbo.log_regenviominutasitioremoto (id bigint NOT NULL, cencos varchar(10), codreg int, codser int, fecini int, fecfin int, fecpro datetime not null default getdate()) "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "end "
   vg_db.Execute ("" & Sql & "")
   
   '-------> Crear nueva tabla minuta pedidos
    Sql = ""
    Sql = " IF not EXISTS ( SELECT  * From sys.objects WHERE   object_id = OBJECT_ID(N'[dbo].[b_minutaspedidoAut]') AND type =( N'U' ) ) "
    Sql = Sql & VgLinea & "BEGIN "
    Sql = Sql & VgLinea & "CREATE TABLE dbo.b_minutaspedidoAut (CECO varchar(10), Proveedor varchar(10), FechaRuta int, Periodo int, Cod_Ingrediente varchar(20), Cod_ProductoSAP varchar(20), Cod_SGP varchar(20), Cantidad_Despacho float, Id_Ruta bigint, Constraint PK_b_minutaspedidoAut Primary Key (CECO, Proveedor, FechaRuta, Periodo, Cod_Ingrediente, Cod_ProductoSAP)) "
    Sql = Sql & VgLinea & "End "

    vg_db.Execute ("" & Sql & "")
   
'   vg_db.Execute "CREATE TABLE dbo.b_minutaspedidoAut (CECO varchar(10), Proveedor varchar(10), FechaRuta int, Periodo int, Cod_Ingrediente varchar(20), Cod_ProductoSAP varchar(20), Cod_SGP varchar(20), Cantidad_Despacho float, Id_Ruta bigint, Constraint PK_b_minutaspedidoAut Primary Key (CECO, Proveedor, FechaRuta, Periodo, Cod_Ingrediente, Cod_ProductoSAP))"

   '-------> Crear nueva tabla log plataforma electronica
   'Log factura PEL
   '* grilla proveedor-numdoc-tipo documento- fecha - observación
   '* frame de fondo
   '* estado 2 y 3 reprocesa solo los estados 3
   '* 1 ingresada - 2 con problema - 3 reprocesar
   '* mover estado 3 a 1
   '* mostrar ultimo 100 registro en la grilla

   
    Sql = ""
    Sql = " IF not EXISTS ( SELECT  * From sys.objects WHERE   object_id = OBJECT_ID(N'[dbo].[Log_FacturaSAP]') AND type =( N'U' ) ) "
    Sql = Sql & VgLinea & "BEGIN "
    Sql = Sql & VgLinea & "CREATE TABLE dbo.Log_FacturaSAP (Ceco varchar(10) NOT NULL, Proveedor varchar(20) NOT NULL, NumeroFactura varchar(20) NOT NULL, Fecha datetime NULL, Estado int NULL, Observacion varchar(1000) NULL, TipoDocumento varchar(02) NOT NULL, Constraint PK_Log_FacturaSAP Primary Key (Ceco, Proveedor, NumeroFactura, TipoDocumento)) "
    Sql = Sql & VgLinea & "End "

    vg_db.Execute ("" & Sql & "")
   
'   vg_db.Execute "CREATE TABLE dbo.Log_FacturaSAP (Ceco varchar(10) NOT NULL, Proveedor varchar(20) NOT NULL, NumeroFactura varchar(20) NOT NULL, Fecha datetime NULL, Estado int NULL, Observacion varchar(1000) NULL, TipoDocumento varchar(02) NOT NULL, Constraint PK_Log_FacturaSAP Primary Key (Ceco, Proveedor, NumeroFactura, TipoDocumento))"
   
   '------- Incluir opción parametro Xml
   RS1.Open "SELECT * FROM dbo.b_clientes WHERE cli_tipo=0 AND cli_codbod>0", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         '------- Incluir opción parametro Xml
         Sql = ""
         Sql = "INSERT INTO a_param VALUES ('parxml', 'Parametro Xml', 'N', '" & 30 & "', '" & RS1!cli_codigo & "') "
         Call ExecProcedimientoInsertarDatos(Sql, "a_param", " par_codigo = 'parxml' and par_cencos = '" & RS1!cli_codigo & "'")
         'vg_db.Execute "INSERT INTO a_param VALUES ('parxml', 'Parametro Xml', 'N', '" & 30 & "', '" & RS1!cli_codigo & "')"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
      
   vg_db.Execute "UPDATE a_param SET par_valor = '179' WHERE par_codigo = 'version'"
   aVer = 179
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 179 Then
   If Dir(dir_trabajo & "Actualizador.mdb") = "" Then
      MsgBox "No existe Actualizador.mdb, para actualizar la nueva versión " & aVer & VgLinea & "Vuelva bajar la actualización y actualice o bien comuniquese con su monitor" & VgLinea & "        Proceso cancelado ...", vbCritical + vbOKOnly, "SGP"
      End
   End If
   
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.8.0 ....."
   V_Acceso.Refresh
   
   '-------> Crear nuevo campo tabla Carga_Ruta_Compra
   Sql = " ALTER TABLE dbo.Carga_Ruta_Compra ADD Periodo int"
   Call ExecProcedimientoColumna(Sql, "Carga_Ruta_Compra", "Periodo")
   'vg_db.Execute "ALTER TABLE dbo.Carga_Ruta_Compra ADD Periodo int"
   
   '-------> Crear nuevos datos a la tabla Carga_Ruta_Compra y ruta_compras
   Set RS1 = vg_db.Execute("SELECT  DISTINCT SUBSTRING(rc.Fecha_despacho, 1, 6) AS Periodo FROM dbo.ruta_compras AS rc ORDER BY Periodo ASC")
   Do While Not RS1.EOF
      '-------> Insertar encabezado ruta de compras
      Sql = " INSERT INTO dbo.Carga_Ruta_Compra  "
      Sql = Sql & " (usuario, Periodo) "
      Sql = Sql & " VALUES ('" & vg_NUsr & "', " & RS1!periodo & ")"
        
      vg_db.Execute (Sql)
        
      '-------> Traer id_carga de la tabla Carga_Ruta_Compra
      Sql = " SELECT isnull(max(convert(float,id_carga)),0) as id_carga FROM dbo.Carga_Ruta_Compra "
      
      Set RS = vg_db.Execute(Sql)
      
      If Not RS.EOF Then
         '-------> Insertar detalle ruta de compras
         vg_db.Execute ("insert into dbo.ruta_compras (Fecha_despacho, ID_centro_de_costo, Familia_producto , ID_proveedor, Sucursal, Sigla_de_ruta, Descripcion_sigla, Observaciones, Ruta_archivo, id_carga) " & _
                        "SELECT Fecha_despacho, ID_centro_de_costo, Familia_producto , ID_proveedor, Sucursal, Sigla_de_ruta, Descripcion_sigla, Observaciones, Ruta_archivo, " & RS(0) & " FROM dbo.ruta_compras AS rc Where SUBSTRING(rc.Fecha_despacho, 1, 6) = '" & RS1!periodo & "'")
      End If
      RS.Close: Set RS = Nothing
      
      RS1.MoveNext
   Loop
   RS1.Close: Set RS1 = Nothing
   
   '-------> Traer id_carga de la tabla Carga_Ruta_Compra
   Sql = " SELECT isnull(min(convert(float,id_carga)),0) as id_carga FROM dbo.Carga_Ruta_Compra "
     
   Set RS = vg_db.Execute(Sql)
      
   If Not RS.EOF Then
      '-------> Borrar detalle ruta de compras
      vg_db.Execute ("delete dbo.ruta_compras where id_carga = " & RS(0) & "")
      '-------> Borrar encabezado ruta de compras
      vg_db.Execute ("delete dbo.Carga_Ruta_Compra where id_carga = " & RS(0) & "")
   
   End If
   RS.Close: Set RS = Nothing
   
   '-------> Actualizar campo cli_activo que este null de la tabla b_clientes
   vg_db.Execute "UPDATE dbo.b_clientes SET cli_activo = '1' WHERE cli_activo IS NULL"
   
   '-------> Borrar relación tabla a_par_codigo_barra_cas - b_clientes
   Sql = ""
   Sql = " IF EXISTS (SELECT 1 FROM sys.sysobjects AS S WHERE name ='FK_a_par_codigo_barra_cas_b_clientes') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "ALTER TABLE dbo.a_par_codigo_barra_cas DROP CONSTRAINT FK_a_par_codigo_barra_cas_b_clientes "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "end "
   vg_db.Execute ("" & Sql & "")
   
'   vg_db.Execute "ALTER TABLE dbo.a_par_codigo_barra_cas DROP CONSTRAINT FK_a_par_codigo_barra_cas_b_clientes"
      
   '-------> Eliminar tabla a_par_codigo_barra_cas
   Sql = ""
   Sql = " IF EXISTS (SELECT 1 FROM sys.sysobjects AS S WHERE name ='a_par_codigo_barra_cas') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "DROP TABLE dbo.a_par_codigo_barra_cas "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "end "
   vg_db.Execute ("" & Sql & "")
   
'   vg_db.Execute "DROP TABLE dbo.a_par_codigo_barra_cas"
      
   '-------> Crear nueva estructura tabla a_par_codigo_barra_cas
   Sql = ""
   Sql = " IF not EXISTS (SELECT 1 FROM sys.sysobjects AS S WHERE name ='a_par_codigo_barra_cas') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "CREATE TABLE dbo.a_par_codigo_barra_cas (a_par_id_codigo int, atr_codigo_barra int, cli_codigo varchar(10), cbar_posinicial int, cbar_largo int, Constraint PK_a_par_codigo_barra_cas Primary Key (a_par_id_codigo, atr_codigo_barra, cli_codigo)) "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "end "
   vg_db.Execute ("" & Sql & "")
   
'   vg_db.Execute "CREATE TABLE dbo.a_par_codigo_barra_cas (a_par_id_codigo int, atr_codigo_barra int, cli_codigo varchar(10), cbar_posinicial int, cbar_largo int, Constraint PK_a_par_codigo_barra_cas Primary Key (a_par_id_codigo, atr_codigo_barra, cli_codigo))"
      
   '-------> Crear nuevo campo tabla b_clientes
   Sql = ""
   Sql = " IF NOT EXISTS(SELECT * FROM information_schema.[columns] WHERE table_name='b_clientes' AND column_name='cli_tipo_vale') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "ALTER TABLE dbo.b_clientes ADD cli_tipo_vale varchar(200) "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "end "
   vg_db.Execute ("" & Sql & "")
   
'   vg_db.Execute "ALTER TABLE dbo.b_clientes ADD cli_tipo_vale varchar(200)"
   
   '------- Incluir opción parametro Xml
   RS1.Open "SELECT * FROM dbo.b_clientes WHERE cli_tipo = 0 AND cli_codbod > 0", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         If RS1!cli_codigo = "23260" Then
            vg_db.Execute ("UPDATE b_clientes SET cli_tipo_vale = '00' WHERE cli_codigo = '816989000'")
            vg_db.Execute ("UPDATE b_clientes SET cli_tipo_vale = '01' WHERE cli_codigo = '81698'")
         End If
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   
   '-------> Crear relación tabla a_par_codigo_barra_cas - b_clientes
   Sql = ""
   Sql = " IF not EXISTS (SELECT 1 FROM sys.sysobjects AS S WHERE name ='FK_a_par_codigo_barra_cas_b_clientes') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "ALTER TABLE dbo.a_par_codigo_barra_cas ADD CONSTRAINT FK_a_par_codigo_barra_cas_b_clientes FOREIGN KEY(cli_codigo) References dbo.b_clientes (cli_codigo) "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "end "
   vg_db.Execute ("" & Sql & "")
   
'   vg_db.Execute "ALTER TABLE dbo.a_par_codigo_barra_cas ADD CONSTRAINT FK_a_par_codigo_barra_cas_b_clientes FOREIGN KEY(cli_codigo) References dbo.b_clientes (cli_codigo)"
   
    Sql = ""
    Sql = " IF EXISTS ( SELECT  * From sys.objects WHERE   object_id = OBJECT_ID(N'[dbo].[RUTA_PEDIDO_FINAL]') AND type IN ( N'U' ) ) "
    Sql = Sql & VgLinea & "BEGIN "
    Sql = Sql & VgLinea & "DROP TABLE dbo.RUTA_PEDIDO_FINAL "
    Sql = Sql & VgLinea & "End "
    
    vg_db.Execute ("" & Sql & "")
  
    Sql = ""
    Sql = " IF EXISTS ( SELECT  * From sys.objects WHERE   object_id = OBJECT_ID(N'[dbo].[PEDIDO_TEMP_FINAL ]') AND type IN ( N'U' ) ) "
    Sql = Sql & VgLinea & "BEGIN "
    Sql = Sql & VgLinea & "DROP TABLE dbo.PEDIDO_TEMP_FINAL "
    Sql = Sql & VgLinea & "End "
    
    vg_db.Execute ("" & Sql & "")
    
    Sql = ""
    Sql = " IF EXISTS ( SELECT  * From sys.objects WHERE   object_id = OBJECT_ID(N'[dbo].[ruta_temp_CALCULO]') AND type IN ( N'U' ) ) "
    Sql = Sql & VgLinea & "BEGIN "
    Sql = Sql & VgLinea & "DROP TABLE dbo.ruta_temp_CALCULO "
    Sql = Sql & VgLinea & "End "

    vg_db.Execute ("" & Sql & "")
   
   If vg_tipbase = "2" Then
      BaseDatos = "Actualizador.mdb"
      Set vg_dbsubesql = New ADODB.Connection
      vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
      vg_dbsubesql.ConnectionTimeout = 30
      vg_dbsubesql.CommandTimeout = 600
      vg_dbsubesql.Open
      RS1.Open "SELECT * FROM Procedimiento WHERE Version = '180' order by id1", vg_dbsubesql, adOpenStatic
      If Not RS1.EOF Then
         Do While Not RS1.EOF
            vg_db.Execute ("" & RS1!Procedimiento & "")
            RS1.MoveNext
         Loop
      End If
      RS1.Close: Set RS1 = Nothing
      vg_dbsubesql.Close
   End If
   
   vg_db.Execute "UPDATE dbo.a_param SET par_valor = '180' WHERE par_codigo = 'version'"
   aVer = 180
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 180 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.8.1 ....."
   V_Acceso.Refresh
      
   If vg_tipbase = "2" Then
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '181' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
   End If
   
   '-------> Actualizar versión
   vg_db.Execute "UPDATE dbo.a_param SET par_valor = '181' WHERE par_codigo = 'version'"
   aVer = 181
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If

If nVer > aVer And aVer = 181 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.8.2 ....."
   V_Acceso.Refresh
      
   '------> Actualizar tabla a_param cuenta contable ctagastos
   Sql = ""
   Sql = " UPDATE dbo.a_param "
   Sql = Sql & VgLinea & "SET par_valor = '410099;410077;410072;410071;410066;410056;410054;410053;410051;410049;410047;410046;410044;410042;410040;410039;410038;410037;410035;410034;410033;410032;410031;410030;410029;410028;410027;410026;410025;410024;410023;410022;410019;410003;410002' "
   Sql = Sql & VgLinea & "WHERE par_codigo = 'ctagastos'"

   vg_db.Execute ("" & Sql & "")
   
   '-------> Insertar concepto flete insumo a la tabla a_param
   
   RS1.Open "SELECT * FROM b_clientes WHERE cli_tipo=0 AND cli_codbod>0", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         '------- Incluir opción parametro Xml
         Sql = ""
         Sql = "INSERT INTO dbo.a_param (par_codigo, par_nombre, par_tipo, par_valor, par_cencos) VALUES ('ctafleins', 'Cuenta Flete Insumo', 'C', '410036', '" & RS1!cli_codigo & "')"
         Call ExecProcedimientoInsertarDatos(Sql, "a_param", "par_codigo = 'ctafleins' and par_cencos = '" & RS1!cli_codigo & "'")
         'vg_db.Execute "INSERT INTO dbo.a_param (par_codigo, par_nombre, par_tipo, par_valor, par_cencos) " & _
                       "VALUES ('ctafleins', 'Cuenta Flete Insumo', 'C', '410036', '" & RS1!cli_codigo & "')"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing

   '-------> Actualizar versión
   vg_db.Execute "UPDATE dbo.a_param SET par_valor = '182' WHERE par_codigo = 'version'"
   aVer = 182
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 182 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.8.2 ....."
   V_Acceso.Refresh
   '-------> agregar un nuevo campo restrición ingreso documento proveedor y traspasos
   Sql = ""
   Sql = " IF NOT EXISTS(SELECT * FROM information_schema.[columns] WHERE table_name='b_proveedor' AND column_name='prv_permiteingdoc') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & " ALTER TABLE b_proveedor ADD prv_permiteingdoc bit "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "end "
   vg_db.Execute ("" & Sql & "")
   
   '-------> agregar un nuevo campo a la tabla
   Sql = ""
   Sql = " IF EXISTS(SELECT * FROM information_schema.[columns] WHERE table_name='Log_FacturaSAP' AND column_name='NumeroFactura') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & " ALTER TABLE [dbo].[Log_FacturaSAP] DROP CONSTRAINT [PK_Log_FacturaSAP] "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & " ALTER TABLE dbo.Log_FacturaSAP ALTER COLUMN NumeroFactura VARCHAR(50) NOT NULL "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & " ALTER TABLE [dbo].[Log_FacturaSAP] Add CONSTRAINT [PK_Log_FacturaSAP] PRIMARY KEY CLUSTERED "
   Sql = Sql & VgLinea & " ( "
   Sql = Sql & VgLinea & " [Ceco] ASC, "
   Sql = Sql & VgLinea & " [Proveedor] ASC, "
   Sql = Sql & VgLinea & " [NumeroFactura] ASC, "
   Sql = Sql & VgLinea & " [TipoDocumento] Asc "
   Sql = Sql & VgLinea & " )WITH (PAD_INDEX = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF) ON [PRIMARY]"
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "end "
   vg_db.Execute ("" & Sql & "")
   
   vg_db.Execute "UPDATE a_param SET par_valor = '183' WHERE par_codigo = 'version'"
   aVer = 183
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If
If nVer > aVer And aVer = 183 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.8.4 ....."
   V_Acceso.Refresh
   '-------> agregar un nuevo campo tabla b_minuta ID_Bloque
   Sql = ""
   Sql = " IF NOT EXISTS ( SELECT  * "
   Sql = Sql & VgLinea & "             From information_schema.[Columns] "
   Sql = Sql & VgLinea & "             WHERE   table_name = 'b_minuta' "
   Sql = Sql & VgLinea & "                     AND column_name = 'ID_Bloque' ) "
   Sql = Sql & VgLinea & " BEGIN "
   Sql = Sql & VgLinea & "     ALTER TABLE b_minuta "
   Sql = Sql & VgLinea & "     ADD ID_Bloque INT NOT NULL  DEFAULT(0) "
   Sql = Sql & VgLinea & " End "
   vg_db.Execute ("" & Sql & "")
   
   '-------> agregar un nuevo campo a la tabla a_estservicio agrupacion estructura
   Sql = ""
   Sql = " IF NOT EXISTS ( SELECT  * "
   Sql = Sql & VgLinea & "             From information_schema.[Columns] "
   Sql = Sql & VgLinea & "             WHERE   table_name = 'a_estservicio' "
   Sql = Sql & VgLinea & "                     AND column_name = 'ess_agrupacionestructura' ) "
   Sql = Sql & VgLinea & " BEGIN "
   Sql = Sql & VgLinea & "     ALTER TABLE a_estservicio "
   Sql = Sql & VgLinea & "     ADD ess_agrupacionestructura INT NOT NULL  DEFAULT(0) "
   Sql = Sql & VgLinea & " End "
   vg_db.Execute ("" & Sql & "")
   
   '-------> Crear nueva tabla b_minutabloque
   Sql = ""
   Sql = " IF NOT EXISTS ( SELECT  1 "
   Sql = Sql & VgLinea & "             FROM    sys.sysobjects AS S "
   Sql = Sql & VgLinea & "             WHERE   name = 'b_MinutaBloque' ) "
   Sql = Sql & VgLinea & " BEGIN "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "     CREATE TABLE [dbo].[b_MinutaBloque] "
   Sql = Sql & VgLinea & "     ( [ID_Bloque]   [bigint]   NOT NULL , "
   Sql = Sql & VgLinea & "       [Ceco] [varchar](10) NOT NULL , "
   Sql = Sql & VgLinea & "       [Regimen] [int] NULL , "
   Sql = Sql & VgLinea & "       [Servicio] [int] NULL , "
   Sql = Sql & VgLinea & "       [FechaDesde] [datetime] NULL , "
   Sql = Sql & VgLinea & "       [FechaHasta] [datetime] NULL , "
   Sql = Sql & VgLinea & "       [IdEstadoMinuta] [int] NULL , "
   Sql = Sql & VgLinea & "       CONSTRAINT [PK_b_MinutaBloque] PRIMARY KEY CLUSTERED "
   Sql = Sql & VgLinea & "         ( [ID_Bloque] ASC, [Ceco] ASC ) "
   Sql = Sql & VgLinea & "         WITH ( PAD_INDEX = OFF, IGNORE_DUP_KEY = OFF ) ON [PRIMARY] "
   Sql = Sql & VgLinea & "     ) "
   Sql = Sql & VgLinea & "         ON "
   Sql = Sql & VgLinea & "     [Primary] "
   Sql = Sql & VgLinea & " End "
   vg_db.Execute ("" & Sql & "")
   
   '-------> Crear nueva tabla b_minutagrupoestructura
   Sql = ""
   Sql = "  IF NOT EXISTS ( SELECT  1 "
   Sql = Sql & VgLinea & "              FROM    sys.sysobjects AS S "
   Sql = Sql & VgLinea & "              WHERE   name = 'b_minutagrupoestructura' ) "
   Sql = Sql & VgLinea & "  BEGIN "
   Sql = Sql & VgLinea & "  "
   Sql = Sql & VgLinea & "      CREATE TABLE [dbo].[b_minutagrupoestructura] "
   Sql = Sql & VgLinea & "      ( [mge_id_bloque] [bigint] NOT NULL , "
   Sql = Sql & VgLinea & "        [mge_cencos] [VarChar](10) "
   Sql = Sql & VgLinea & "                                   NOT NULL , "
   Sql = Sql & VgLinea & "        [mge_grupoestructura] [int] NOT NULL , "
   Sql = Sql & VgLinea & "        [mge_ponderaciontotal] [float] NULL , "
   Sql = Sql & VgLinea & "        CONSTRAINT [PK_b_minutagrupoestructura] PRIMARY KEY CLUSTERED "
   Sql = Sql & VgLinea & "          ( [mge_id_bloque] ASC, [mge_cencos] ASC, [mge_grupoestructura] ASC) "
   Sql = Sql & VgLinea & "          WITH ( PAD_INDEX = OFF, IGNORE_DUP_KEY = OFF ) ON [PRIMARY] "
   Sql = Sql & VgLinea & "      ) "
   Sql = Sql & VgLinea & "          ON "
   Sql = Sql & VgLinea & "      [Primary] "
   Sql = Sql & VgLinea & "  End "
   vg_db.Execute ("" & Sql & "")
   
   '-------> Crear nueva tabla a_EstadoMinuta
   Sql = ""
   Sql = "  IF NOT EXISTS ( SELECT  1 "
   Sql = Sql & VgLinea & "             FROM    sys.sysobjects AS S "
   Sql = Sql & VgLinea & "             WHERE   name = 'a_EstadoMinuta' ) "
   Sql = Sql & VgLinea & " BEGIN "
   Sql = Sql & VgLinea & "  "
   Sql = Sql & VgLinea & "     CREATE TABLE [dbo].[a_EstadoMinuta] "
   Sql = Sql & VgLinea & "     ( [IdEstadoMinuta] [int] NOT NULL , "
   Sql = Sql & VgLinea & "       [DescripcionEstado] [varchar](100) NULL , "
   Sql = Sql & VgLinea & "       [Fecha] [DateTime] "
   Sql = Sql & VgLinea & "         NULL "
   Sql = Sql & VgLinea & "         CONSTRAINT [DF_a_EstadoMinuta_Fecha] DEFAULT ( GETDATE() ) , "
   Sql = Sql & VgLinea & "       CONSTRAINT [PK_a_EstadoMinuta] PRIMARY KEY CLUSTERED "
   Sql = Sql & VgLinea & "         ( [IdEstadoMinuta] ASC ) "
   Sql = Sql & VgLinea & "         WITH ( PAD_INDEX = OFF, IGNORE_DUP_KEY = OFF ) ON [PRIMARY] "
   Sql = Sql & VgLinea & "     ) "
   Sql = Sql & VgLinea & "         ON "
   Sql = Sql & VgLinea & "     [Primary] "
   Sql = Sql & VgLinea & " End "
   vg_db.Execute ("" & Sql & "")
   
   '-------> agregar un nuevo campo tabla b_minutadet
   Sql = ""
   Sql = " IF NOT EXISTS(SELECT * FROM information_schema.[columns] WHERE table_name='b_minutadet' AND column_name='mid_porrac') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & " ALTER TABLE b_minutadet ADD mid_porrac float "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "end "
   vg_db.Execute ("" & Sql & "")
   
   '-------> Mover sector cero a tabla a_sector
   Sql = ""
   Sql = " IF NOT EXISTS(SELECT * FROM information_schema.[columns] WHERE table_name='a_sector' AND column_name='sec_codigo') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "    IF NOT EXISTS (SELECT sec_codigo FROM dbo.a_sector AS as2 WHERE as2.sec_codigo = 0) "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "     begin "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "      INSERT INTO dbo.a_sector "
   Sql = Sql & VgLinea & "        ( sec_codigo, sec_nombre, sec_orden ) "
   Sql = Sql & VgLinea & "      Values "
   Sql = Sql & VgLinea & "        ( 0, 'Sector no Definido', 0 ) "
   Sql = Sql & VgLinea & "     End "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "  End "
   vg_db.Execute ("" & Sql & "")
   
   '-------> agregar un nuevo campo tabla b_clientes tipominuta
   Sql = ""
   Sql = " IF NOT EXISTS(SELECT * FROM information_schema.[columns] WHERE table_name='b_clientes' AND column_name='cli_tipominuta') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & " ALTER TABLE b_clientes ADD cli_tipominuta int"
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "end "
   vg_db.Execute ("" & Sql & "")
   
   '-------> agregar un nuevo campo tabla b_clientes tipoformatocompras
   Sql = ""
   Sql = " IF NOT EXISTS(SELECT * FROM information_schema.[columns] WHERE table_name='b_clientes' AND column_name='cli_tipoformatocompras') "
   Sql = Sql & VgLinea & "begin "
   Sql = Sql & VgLinea & " ALTER TABLE b_clientes ADD cli_tipoformatocompras int "
   Sql = Sql & VgLinea & " "
   Sql = Sql & VgLinea & "end "
   vg_db.Execute ("" & Sql & "")
   
   If vg_tipbase = "2" Then
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '184' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
   End If
   
   '-------> Ejecuta minuta bloque
   Sql = ""
   Sql = Sql & "Execute GeneraMinutaBloque"
   vg_db.Execute ("" & Sql & "")
   
   Sql = ""
   Sql = Sql & "Drop PROC [dbo].[GeneraMinutaBloque]"
   vg_db.Execute ("" & Sql & "")
   
   vg_db.Execute "UPDATE a_param SET par_valor = '184' WHERE par_codigo = 'version'"
   aVer = 184
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If

If nVer > aVer And aVer = 184 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.8.5 ....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '185' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '185' WHERE par_codigo = 'version'"
   aVer = 185
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If

If nVer > aVer And aVer = 185 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.8.6 ....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '186' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '186' WHERE par_codigo = 'version'"
   aVer = 186
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If

If nVer > aVer And aVer = 186 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.8.7 ....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '187' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '187' WHERE par_codigo = 'version'"
   aVer = 187
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If

If nVer > aVer And aVer = 187 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.8.8 ....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '188' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '188' WHERE par_codigo = 'version'"
   aVer = 188
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If

Dim AddReceta As Long
Dim FecCie As Long

If nVer > aVer And aVer = 188 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.8.9 ....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '189' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
   End If
  
   '-------> Insertar Numero de receta wen planificación
   AddReceta = 0
   FecCie = 0
   RS1.Open "SELECT * FROM b_clientes WHERE cli_tipo=0 AND cli_codbod>0", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         If RS1!cli_tipominuta = "3" Then
            
            '-------> Sacar Ultima fecha cierre
            FecCie = 0
            Set RS2 = vg_db.Execute("SELECT * FROM dbo.a_param AS ap WHERE ap.par_cencos = '" & RS1!cli_codigo & "' AND ap.par_codigo = 'ciediario'")
            If Not RS2.EOF Then
               FecCie = Format(CDate(fg_Desencripta(TipoDato(RS2!par_valor, ""))) - 1, "yyyymm")
            End If
            RS2.Close: Set RS2 = Nothing
            
            If FecCie > 0 Then
               
               '-------> Actualizar minutas simap no adicional
               vg_db.Execute "Update DBO.b_minutadet " & _
                             "Set mid_rec5eta = 1 " & _
                             "FROM dbo.b_minuta AS bm " & _
                             "INNER JOIN dbo.b_minutadet AS bm2 ON bm.min_codigo = bm2.mid_codigo " & _
                             "Where CONVERT(INT,SUBSTRING( CONVERT(VARCHAR(8),bm.min_fecmin),1,6)) >= " & FecCie & " " & _
                             "AND bm.min_codreg > 9999 " & _
                             "AND bm.min_codser > 9999 " & _
                             "AND bm.min_cencos = '" & RS1!cli_codigo & "'"
            
            End If
            
            AddReceta = 5
         
         Else
            AddReceta = 200
         End If
         '------- Incluir opción parametro Xml
         Sql = ""
         Sql = "INSERT INTO dbo.a_param (par_codigo, par_nombre, par_tipo, par_valor, par_cencos) VALUES ('addreceta', 'Numero receta', 'N', '" & AddReceta & "', '" & RS1!cli_codigo & "')"
         Call ExecProcedimientoInsertarDatos(Sql, "a_param", "par_codigo = 'addreceta' and par_cencos = '" & RS1!cli_codigo & "'")
         
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   
   vg_db.Execute "UPDATE a_param SET par_valor = '189' WHERE par_codigo = 'version'"
   aVer = 189
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
End If

If nVer > aVer And aVer = 189 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.9.0 ....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '190' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
   End If

   '-------> Insertar Numero de receta wen planificación
   AddReceta = 0
   FecCie = 0
   RS1.Open "SELECT * FROM b_clientes WHERE cli_tipo = 0 AND cli_codbod > 0", vg_db, adOpenStatic
   
   If Not RS1.EOF Then
      
      Do While Not RS1.EOF
         
         If RS1!cli_tipominuta = "3" Then
            
            '-------> Sacar Ultima fecha cierre
            FecCie = 0
            Set RS2 = vg_db.Execute("SELECT * FROM dbo.a_param AS ap WHERE ap.par_cencos = '" & RS1!cli_codigo & "' AND ap.par_codigo = 'ciediario'")
            If Not RS2.EOF Then
               FecCie = Format(CDate(fg_Desencripta(TipoDato(RS2!par_valor, ""))) - 1, "yyyymm")
            End If
            RS2.Close: Set RS2 = Nothing
            
            If FecCie > 0 Then
               
               '-------> Actualizar minutas simap no adicional
               vg_db.Execute "Update DBO.b_minutadet " & _
                             "Set mid_rec5eta = 1 " & _
                             "FROM dbo.b_minuta AS bm " & _
                             "INNER JOIN dbo.b_minutadet AS bm2 ON bm.min_codigo = bm2.mid_codigo " & _
                             "Where CONVERT(INT,SUBSTRING( CONVERT(VARCHAR(8),bm.min_fecmin),1,6)) >= " & FecCie & " " & _
                             "AND bm.min_codreg > 9999 " & _
                             "AND bm.min_codser > 9999 " & _
                             "AND bm.min_cencos = '" & RS1!cli_codigo & "'"
            
            End If
            
            AddReceta = 5
            
            Set RS2 = vg_db.Execute("SELECT * FROM dbo.a_param AS ap WHERE ap.par_cencos = '" & RS1!cli_codigo & "' AND ap.par_codigo = 'addreceta'")
            If Not RS2.EOF Then
               
               vg_db.Execute ("update a_param set par_valor = '" & AddReceta & "' where par_codigo = 'addreceta' and par_cencos = '" & RS1!cli_codigo & "'")
            
            End If
            RS2.Close: Set RS2 = Nothing
            
         
         Else
            AddReceta = 200
         End If
         '------- Incluir opción parametro Xml
         Sql = ""
         Sql = "INSERT INTO dbo.a_param (par_codigo, par_nombre, par_tipo, par_valor, par_cencos) VALUES ('addreceta', 'Numero receta', 'N', '" & AddReceta & "', '" & RS1!cli_codigo & "')"
         Call ExecProcedimientoInsertarDatos(Sql, "a_param", "par_codigo = 'addreceta' and par_cencos = '" & RS1!cli_codigo & "'")
         
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   
   vg_db.Execute "UPDATE a_param SET par_valor = '190' WHERE par_codigo = 'version'"
   aVer = 190
   
   '-------> Borrar concepto descarga
   vg_db.Execute ("DELETE  a_param WHERE par_codigo = 'Descarga'")
   
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 190 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.9.1 ....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
       
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '191' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
   
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '191' WHERE par_codigo = 'version'"
   aVer = 191
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 191 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.9.2 ....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
       
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '192' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
   
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '192' WHERE par_codigo = 'version'"
   aVer = 192
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 192 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.93 ....."
   V_Acceso.Refresh
   
   vg_db.Execute "UPDATE a_param SET par_valor = '193' WHERE par_codigo = 'version'"
   aVer = 193
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 193 Then
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.9.4 ....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
   
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '194' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          
          Do While Not RS1.EOF
             
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          
          Loop
       
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
       
   End If
   
   '-------> Actualizar minuta raciones cero producidas posteriores al periodo cierre de mes
   vg_db.Execute ("UPDATE dbo.b_minuta SET min_racrea = 0 FROM dbo.b_minuta AS bm WITH (NOLOCK) INNER JOIN dbo.b_cierreperiodo AS bc WITH (NOLOCK) ON bc.cie_cencos = bm.min_cencos AND bc.cie_estado = 1 AND min_fecmin > bc.cie_fecter")
   vg_db.Execute ("UPDATE dbo.b_minutaraciones SET mir_nrorac = 0 FROM dbo.b_minutaraciones AS bm WITH (NOLOCK ) INNER JOIN dbo.b_cierreperiodo AS bc WITH (NOLOCK ) ON bc.cie_cencos = bm.mir_cencos AND bc.cie_estado = 1 AND bm.mir_fecmin > bc.cie_fecter AND mir_rutcli = 'PRODUCIDAS'")

   vg_db.Execute ("sgp_Ins_TipoActividad 6")
   
   vg_db.Execute "UPDATE a_param SET par_valor = '194' WHERE par_codigo = 'version'"
   aVer = 194
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
   
End If

If nVer > aVer And aVer = 194 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.9.5....."
   V_Acceso.Refresh
   
   vg_db.Execute "UPDATE a_param SET par_valor = '195' WHERE par_codigo = 'version'"
   aVer = 195
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 195 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      
      KillProcess ("sgpsdx.exe")
   
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.9.6....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
   
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '196' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          
          Do While Not RS1.EOF
             
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          
          Loop
       
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
       
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '196' WHERE par_codigo = 'version'"
   aVer = 196
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh
   
End If

If nVer > aVer And aVer = 196 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.9.7....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
   
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '197' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          
          Do While Not RS1.EOF
             
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          
          Loop
       
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
       
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '197' WHERE par_codigo = 'version'"
   aVer = 197
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 197 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.9.8....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
   
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '198' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          
          Do While Not RS1.EOF
             
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          
          Loop
       
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
       
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '198' WHERE par_codigo = 'version'"
   aVer = 198
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 198 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      KillProcess ("sgpsdx.exe")
   End If
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 1.9.9....."
   V_Acceso.Refresh
   
   vg_db.Execute "UPDATE a_param SET par_valor = '199' WHERE par_codigo = 'version'"
   aVer = 199
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 199 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      
      KillProcess ("sgpsdx.exe")
   
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.0.0....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
   
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '200' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          
          Do While Not RS1.EOF
             
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          
          Loop
       
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
       
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '200' WHERE par_codigo = 'version'"
   aVer = 200
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 200 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      
      KillProcess ("sgpsdx.exe")
   
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.0.1....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
   
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '201' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          
          Do While Not RS1.EOF
             
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          
          Loop
       
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
       
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '201' WHERE par_codigo = 'version'"
   aVer = 201
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 201 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      
      KillProcess ("sgpsdx.exe")
   
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.0.2....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
   
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '202' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          
          Do While Not RS1.EOF
             
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          
          Loop
       
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
       
   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '202' WHERE par_codigo = 'version'"
   aVer = 202
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 202 Then
   
   If ConsultaProcess("sgpsdx.exe") Then
      
      KillProcess ("sgpsdx.exe")
   
   End If
   
   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.0.3....."
   V_Acceso.Refresh
   
   If vg_tipbase = "2" Then
   
       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open
       
       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '203' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then
          
          Do While Not RS1.EOF
             
             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext
          
          Loop
       
       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close
       
   End If
   
   RS.Open "SELECT * FROM b_clientes WHERE cli_tipo = 0 AND cli_codbod > 0", vg_db, adOpenStatic
   
   If Not RS.EOF Then
      
      Do While Not RS.EOF
         
'         If RS!cli_tipominuta = "3" Then
       
            '------- Incluir opción parametro Xml
            Sql = ""
            Sql = "INSERT INTO dbo.a_param (par_codigo, par_nombre, par_tipo, par_valor, par_cencos) VALUES ('Script01', 'sp_iu_minutaPedidoCentralizado', 'C', '', '" & RS!cli_codigo & "')"
            Call ExecProcedimientoInsertarDatos(Sql, "a_param", "par_codigo = 'Script01' and par_cencos = '" & RS!cli_codigo & "'")
         
           '------- Incluir opción parametro Xml
           Sql = ""
           Sql = "INSERT INTO dbo.a_param (par_codigo, par_nombre, par_tipo, par_valor, par_cencos) VALUES ('addreceta', 'Numero receta', 'N', '" & 5 & "', '" & RS!cli_codigo & "')"
           Call ExecProcedimientoInsertarDatos(Sql, "a_param", "par_codigo = 'addreceta' and par_cencos = '" & RS!cli_codigo & "'")
         
           
'         End If
            
         RS.MoveNext
          
      Loop
          
   End If
   RS.Close
   Set RS = Nothing
   
   vg_db.Execute "UPDATE a_param SET par_valor = '203' WHERE par_codigo = 'version'"
   aVer = 203
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 203 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.0.4....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '204' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If

    Dim est As Boolean
    Dim fecini As Long, fecfin As Long, diatop As Long
    fecini = 0
    fecfin = 0
    diatop = 0
    est = False
    RS1.Open "select cie_cencos,max(cie_fecter) as cie_fecter " & _
             "from b_cierreperiodo as a " & _
             "inner join b_clientes as b on a.cie_cencos = b.cli_codigo " & _
             "             and b.cli_activo = 1 " & _
             "             and b.cli_tipo = 0 " & _
             "group by cie_cencos ", vg_db, adOpenStatic
    
    If Not RS1.EOF Then
       
       Do While Not RS1.EOF
          
          fecini = 0
          fecfin = 0
          diatop = 0
          est = False
    
          fecini = Format(dBoM(Mid(RS1!cie_fecter, 7, 2) & "/" & Mid(RS1!cie_fecter, 5, 2) & "/" & Mid(RS1!cie_fecter, 1, 4)), "yyyymmdd")
          fecfin = Format((Mid(RS1!cie_fecter, 7, 2) & "/" & Mid(RS1!cie_fecter, 5, 2) & "/" & Mid(RS1!cie_fecter, 1, 4)), "yyyymmdd")
          diatop = Val(Mid(fecfin, 7, 2))
          
          Do While Mid(fecini, 1, 4) <> 2040
               
               If fecini = 0 And fecfin = 0 Then
                  
                  fecini = Format(dBoM(Date), "yyyymmdd"): fecfin = Format(dEoM(Date), "yyyymmdd")
                  diatop = Val(Mid(fecfin, 7, 2))
                  vg_db.Execute "insert into b_cierreperiodo (cie_cencos, cie_periodo, cie_fecini, cie_fecter, cie_estado) values ('" & RS1!cie_cencos & "', " & Mid(fecini, 1, 6) & ", " & IIf(Mid(fecfin, 7, 2) > 27, fecini, (Mid(fecini, 1, 6) & (diatop + 1))) & ", " & fecfin & ", 2)"
               
               Else
                  If (fecfin + 1) > Format(dEoM(Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymmdd") Then
                     
                     fecini = Format(dBoM(BEoM(Mid(fecini, 7, 2) & "/" & Mid(fecini, 5, 2) & "/" & Mid(fecini, 1, 4))), "yyyymmdd")
                  
                  Else
                     
                     fecini = Format(dEoM(Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymm") & fg_pone_cero(Str(Val(diatop + 1)), 2) 'fg_pone_cero(Str(Val(Mid(fecini, 7, 2))), 2)
                  
                  End If
                  fecfin = Format(BEoM(Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymm") & fg_pone_cero(Str(Val(diatop)), 2) 'Mid(fecfin, 7, 2)
                  
                  If fecfin + 1 > Format(dEoM("01/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymmdd") Then
                     
                     fecfin = Format(dEoM("01/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymmdd")
                  
                  End If
                  If (fecfin + 1) > Format(dEoM(Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymmdd") Then
                     
                     vg_db.Execute "insert into b_cierreperiodo (cie_cencos, cie_periodo, cie_fecini, cie_fecter, cie_estado) values ('" & RS1!cie_cencos & "', " & IIf(diatop > 30, Mid(fecini, 1, 6), Mid(fecfin, 1, 6)) & ", " & IIf(Mid(fecfin, 7, 2) > 27, fecini, Mid(fecini, 1, 4) & Mid(fecini, 5, 2) & fg_pone_cero(Str(Val(diatop)), 2)) & ", " & Mid(fecfin, 1, 4) & Mid(fecfin, 5, 2) & Mid(fecfin, 7, 2) & ", 2)"
                  
                  Else
                     
                     vg_db.Execute "insert into b_cierreperiodo (cie_cencos, cie_periodo, cie_fecini, cie_fecter, cie_estado) values ('" & RS1!cie_cencos & "', " & IIf(diatop > 30, Mid(fecini, 1, 6), Mid(fecfin, 1, 6)) & ", " & IIf(Mid(fecfin, 7, 2) > 27, fecini, Mid(fecini, 1, 4) & Mid(fecini, 5, 2) & fg_pone_cero(Str(Val(diatop + 1)), 2)) & ", " & Mid(fecfin, 1, 4) & Mid(fecfin, 5, 2) & Mid(fecfin, 7, 2) & ", 2)"
                  
                  End If
               
               End If
'               If est = False Then est = True: vg_db.Execute "update b_cierreperiodo set cie_estado=1 where cie_periodo=" & Val(Mid(fecini, 1, 6)) & ""
            
          Loop
          '------- Fin cierres periodo
          
          
          RS1.MoveNext
       
       Loop
    
    End If
    RS1.Close
    Set RS1 = Nothing
    
   vg_db.Execute "UPDATE a_param SET par_valor = '204' WHERE par_codigo = 'version'"
   aVer = 204
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 204 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.0.5....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
     
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '205' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If

   vg_db.Execute "UPDATE a_param SET par_valor = '205' WHERE par_codigo = 'version'"
   aVer = 205
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 205 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.0.6....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '206' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If

   vg_db.Execute "UPDATE a_param SET par_valor = '206' WHERE par_codigo = 'version'"
   aVer = 206
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If
   
If nVer > aVer And aVer = 206 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.0.7....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '207' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '207' WHERE par_codigo = 'version'"
   aVer = 207
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 207 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.0.8....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '208' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '208' WHERE par_codigo = 'version'"
   aVer = 208
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 208 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.0.9....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '209' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '209' WHERE par_codigo = 'version'"
   aVer = 209
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 209 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.1.0....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '210' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '210' WHERE par_codigo = 'version'"
   aVer = 210
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 210 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.1.1....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '211' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '211' WHERE par_codigo = 'version'"
   aVer = 211
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 211 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.1.2....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '212' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '212' WHERE par_codigo = 'version'"
   aVer = 212
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 212 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.1.3....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '213' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             vg_db.Execute ("" & RS1!Procedimiento & "")
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '213' WHERE par_codigo = 'version'"
   aVer = 213
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 213 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.1.4....."
   V_Acceso.Refresh
   
   vg_db.Execute "UPDATE a_param SET par_valor = '214' WHERE par_codigo = 'version'"
   aVer = 214
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If


If nVer > aVer And aVer = 214 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.1.5....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '215' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             If Trim(RS1!Procedimiento) <> "" Then
                
                vg_db.Execute ("" & RS1!Procedimiento & "")
                
             End If
             
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '215' WHERE par_codigo = 'version'"
   aVer = 215
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 215 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.1.6....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '216' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             If Trim(RS1!Procedimiento) <> "" Then
                
                vg_db.Execute ("" & RS1!Procedimiento & "")
                
             End If
             
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '216' WHERE par_codigo = 'version'"
   aVer = 216
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 216 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.1.7....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '217' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             If Trim(RS1!Procedimiento) <> "" Then
                
                vg_db.Execute ("" & RS1!Procedimiento & "")
                
             End If
             
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '217' WHERE par_codigo = 'version'"
   aVer = 217
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 217 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.1.8....."
   V_Acceso.Refresh
  
   vg_db.Execute "UPDATE a_param SET par_valor = '218' WHERE par_codigo = 'version'"
   aVer = 218
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 218 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.1.9....."
   V_Acceso.Refresh
  
   vg_db.Execute "UPDATE a_param SET par_valor = '219' WHERE par_codigo = 'version'"
   aVer = 219
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 219 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.2.0....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '220' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             If Trim(RS1!Procedimiento) <> "" Then
                
                vg_db.Execute ("" & RS1!Procedimiento & "")
                
             End If
             
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '220' WHERE par_codigo = 'version'"
   aVer = 220
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 219 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.2.0....."
   V_Acceso.Refresh
  
   vg_db.Execute "UPDATE a_param SET par_valor = '220' WHERE par_codigo = 'version'"
   aVer = 220
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 220 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.2.1....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '221' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             If Trim(RS1!Procedimiento) <> "" Then
                
                vg_db.Execute ("" & RS1!Procedimiento & "")
                
             End If
             
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '221' WHERE par_codigo = 'version'"
   aVer = 221
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 221 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.2.2....."
   V_Acceso.Refresh
  
   vg_db.Execute "UPDATE a_param SET par_valor = '222' WHERE par_codigo = 'version'"
   aVer = 222
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 222 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.2.3....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '223' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             If Trim(RS1!Procedimiento) <> "" Then
                
                vg_db.Execute ("" & RS1!Procedimiento & "")
                
             End If
             
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '223' WHERE par_codigo = 'version'"
   aVer = 223
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 223 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.2.4....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '224' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             If Trim(RS1!Procedimiento) <> "" Then
                
                vg_db.Execute ("" & RS1!Procedimiento & "")
                
             End If
             
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '224' WHERE par_codigo = 'version'"
   aVer = 224
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 224 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.2.5....."
   V_Acceso.Refresh

   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '225' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             If Trim(RS1!Procedimiento) <> "" Then
                
                vg_db.Execute ("" & RS1!Procedimiento & "")
                
             End If
             
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
   
   vg_db.Execute "UPDATE a_param SET par_valor = '225' WHERE par_codigo = 'version'"
   aVer = 225
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 225 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.2.6....."
   V_Acceso.Refresh
    
   vg_db.Execute "UPDATE a_param SET par_valor = '226' WHERE par_codigo = 'version'"
   aVer = 226
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 226 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.2.7....."
   V_Acceso.Refresh
  
   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
       Set vg_dbsubesql = New ADODB.Connection
       vg_dbsubesql.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDatos + "' ;Persist Security Info=False"
       vg_dbsubesql.ConnectionTimeout = 30
       vg_dbsubesql.CommandTimeout = 600
       vg_dbsubesql.Open

       RS1.Open "SELECT * FROM Procedimiento WHERE Version = '227' order by id1", vg_dbsubesql, adOpenStatic
       If Not RS1.EOF Then

          Do While Not RS1.EOF

             If Trim(RS1!Procedimiento) <> "" Then
                
                vg_db.Execute ("" & RS1!Procedimiento & "")
                
             End If
             
             RS1.MoveNext

          Loop

       End If
       RS1.Close: Set RS1 = Nothing
       vg_dbsubesql.Close

   End If
      
   vg_db.Execute "UPDATE a_param SET par_valor = '227' WHERE par_codigo = 'version'"
   aVer = 227
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

If nVer > aVer And aVer = 227 Then

   If ConsultaProcess("sgpsdx.exe") Then

      KillProcess ("sgpsdx.exe")

   End If

   V_Acceso.Label1(1).Visible = True
   V_Acceso.Label1(1).Caption = "Un Momento Actualizando Versión 2.2.8....."
   V_Acceso.Refresh
  
   If vg_tipbase = "2" Then

       BaseDatos = "Actualizador.mdb"
       
       'validar si existe NombreArchivo
       If Not fso.FileExists(dir_trabajo & BaseDatos) Then
   
          Set fso = Nothing
          MsgBox "No existe archivo de actualización Versión " & nVer & "." & Chr(13) & "Comunicase con su monitor o bien mesa de ayuda." & Chr(13) & "            Proceso Cancelado ...", vbExclamation + vbOKOnly, "SGP"
          End
       
       End If
       
        Set vg_dbsubesql = New ADODB.Connection
        vg_dbsubesql.ConnectionString = "Provider=" & Trim(Provider) & ";Data Source=" & dir_trabajo & BaseDatos & ";Persist Security Info=False"
        vg_dbsubesql.ConnectionTimeout = 30
        vg_dbsubesql.CommandTimeout = 600
        vg_dbsubesql.Open

        RS1.Open "SELECT * FROM Procedimiento WHERE Version = '228' ORDER BY id1", vg_dbsubesql, adOpenStatic

        If Not RS1.EOF Then
            
            Do While Not RS1.EOF
                
                procCompleto = ""

                If Not IsNull(RS1!Procedimiento) Then
                    procCompleto = procCompleto & RS1!Procedimiento
                End If
                If Not IsNull(RS1!Procedimiento_II) Then
                    procCompleto = procCompleto & vbCrLf & RS1!Procedimiento_II
                End If

                If Trim(procCompleto) <> "" Then
                    vg_db.Execute procCompleto
                End If

                RS1.MoveNext
                Loop
        End If

        RS1.Close: Set RS1 = Nothing
        vg_dbsubesql.Close

   End If
      
   vg_db.Execute "UPDATE a_param SET par_valor = '228' WHERE par_codigo = 'version'"
   aVer = 228
   V_Acceso.Label1(1).Visible = False
   V_Acceso.Refresh

End If

Exit Function
Man_Error:

Set fso = Nothing

'If Err.Number = -2147467259 Then
MsgBox "Error... " & Err.Description & Chr(13) & "Comunicase con departamento de informatica", vbExclamation + vbOKOnly, "SGP"
Resume Next
End Function


