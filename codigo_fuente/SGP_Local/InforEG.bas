Attribute VB_Name = "InforEG"
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim ibusca As Long, i As Long
Dim itab As Integer, swvalidar As Integer, itexto As Integer, opboton As Integer
Dim cAccion As String, modo As String, codigo As String, incluir As String, alterar As String, eliminar As String, imprimir As String
Dim vecdatos(11) As String

Public Function I_FoFi(codcas As String, FechaFol As Long, Folio As Long)

Dim RS1 As New ADODB.Recordset, RS2 As New ADODB.Recordset  'Recordset para Consulta
Dim Fecha As Date, i As Double, tDesc As Double, tAlim As Double, tVarios As Double, cCta As String
Dim TotDesc As Double, TotAlim As Double, TotVarios As Double, TotEgre As Double, TotIva As Double, pctimp As Double
Dim tIVA As Double, tEgre As Double, NumDoc As Double, aAp As String, j As Long
Dim tMovil As Double, TotMovil As Double, crut As String

On Local Error GoTo Error_SalirFoFi

fg_carga ""
MsgTitulo = "Informes de Rendición de Gastos FOFI"
j = Len(Trim(Str(FechaFol)))

If j = 7 Then
    
    Fecha = CDate("0" + Mid(Trim(Str(FechaFol)), 1, 1) + "/" + Mid(Trim(Str(FechaFol)), 2, 2) + "/" + Mid(Trim(Str(FechaFol)), 4, 4))

Else
   
   Fecha = CDate(Mid(Trim(Str(FechaFol)), 1, 2) + "/" + Mid(Trim(Str(FechaFol)), 3, 2) + "/" + Mid(Trim(Str(FechaFol)), 5, 4))

End If

If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS2.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo='" & codcas & "' AND cli_tipo=0", vg_db, adOpenStatic
If RS2.EOF Then fg_descarga: RS2.Close: Set RS2 = Nothing: Exit Function
aAp = Trim(vg_NUsr) & "_tmp_InfFoFi"
'---Cheque la existencia de tablas temporales. Si existen las elimina---
fg_CheckTmp aAp

vg_db.BeginTrans

'vg_db.Execute "SELECT a.toc_rutpro, a.toc_tipdoc, a.toc_numdoc, a.toc_fecemi, a.toc_ivadoc, a.toc_fledoc, b.dec_codmer, b.dec_numlin, b.dec_ptotal, c.pro_ctacon, d.prv_nombre, a.toc_totdoc INTO " & aAp & _
'              " FROM b_totcompras a, b_detcompras b, b_productos c, b_proveedor d WHERE a.toc_codbo=" & vg_codbod & " AND a.toc_tipinf='F' AND a.toc_tipdoc<>'SN' AND a.toc_numinf=" & Folio & _
'              " AND a.toc_rutpro=b.dec_rutpro AND  a.toc_tipdoc=b.dec_tipdoc AND a.toc_numdoc=b.dec_numdoc AND  a.toc_rutpro=d.prv_codigo AND b.dec_codmer=c.pro_codigo " & _
'              " GROUP BY  a.toc_rutpro, a.toc_tipdoc, a.toc_numdoc, c.pro_ctacon, a.toc_fecemi, a.toc_fledoc, b.dec_codmer, b.dec_numlin, b.dec_ptotal, a.toc_totdoc, a.toc_ivadoc, d.prv_nombre"

vg_db.Execute "SELECT a.toc_rutpro, a.toc_tipdoc, a.toc_numdoc, a.toc_fecemi, a.toc_ivadoc, a.toc_fledoc, b.dec_codmer, b.dec_numlin, b.dec_ptotal, c.pro_ctacon, d.prv_nombre, a.toc_totdoc INTO " & aAp & _
              " FROM b_totcompras a, b_detcompras b, b_productos c, b_proveedor d WHERE a.toc_codbo=" & vg_codbod & " AND a.toc_tipinf='F' AND a.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo <>'SN') AND a.toc_numinf=" & Folio & _
              " AND a.toc_rutpro=b.dec_rutpro AND  a.toc_tipdoc=b.dec_tipdoc AND a.toc_numdoc=b.dec_numdoc AND  a.toc_rutpro=d.prv_codigo AND b.dec_codmer=c.pro_codigo " & _
              " GROUP BY  a.toc_rutpro, a.toc_tipdoc, a.toc_numdoc, c.pro_ctacon, a.toc_fecemi, a.toc_fledoc, b.dec_codmer, b.dec_numlin, b.dec_ptotal, a.toc_totdoc, a.toc_ivadoc, d.prv_nombre"

vg_db.CommitTrans

vg_db.BeginTrans

vg_db.Execute "ALTER TABLE " & aAp & " ADD COLUMN Cuenta CHAR(20)"
vg_db.Execute "UPDATE " & aAp & " SET Cuenta='Alimentos' WHERE pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
vg_db.Execute "UPDATE " & aAp & " SET Cuenta='Desechables' WHERE pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
vg_db.Execute "UPDATE " & aAp & " SET Cuenta='Movilizacion' WHERE pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctamovil"), ";", "','") & "')"
vg_db.Execute "UPDATE " & aAp & " SET Cuenta='Varios' WHERE cuenta = '' OR ISNULL(cuenta) "

vg_db.CommitTrans

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT * FROM " & aAp & " ", vg_db, adOpenStatic
'---------------------Fin Carga de Datos -----------------
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fg_descarga: MsgBox "No existen datos para consulta...", vbExclamation + vbOKOnly, MsgTitulo: Exit Function

Preview.Refresh
Preview.Cls

With Preview.VSPrinter
     
     vg_reporte = dir_trabajo_Inf & "FOFI" & codcas & Format(Date, "yyyymm") & ".rtf"
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & vg_reporte
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orLandscape
    .MarginLeft = 500
    .StartDoc
    .TextAlign = taLeftTop
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = Trim(fg_ponepiepagina) & " Contrato :  " & (RS2!cli_codigo) & "  " & Trim(RS2!cli_nombre) & "| VºBº ADC______________________________________________|Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 13500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = "Rendición Fondo Fijo FOFI"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, 1) = 13500: .TableCell(tcAlign, 1) = taLeftTop
    .TableCell(tcColWidth, 2) = 13500: .TableCell(tcAlign, 2) = taLeftTop
    .TableCell(tcColWidth, 3) = 13500: .TableCell(tcAlign, 3) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
'    .TableCell(tcText, 1) = "Contrato               " & " (" & Trim(CodCas) & " " & Trim(nomcas) & ")" & Space(200) & " " & "Folio Nº" & Str(Folio)
    .TableCell(tcText, 2) = "Período              " & " " & UCase(Meses(CStr(Fecha))) & "/" & Str(Year(Fecha))
    .TableCell(tcText, 3) = "Monto Asignado" & " " & "_______________"
    Print #1, .TableCell(tcText, 1, 1)
    Print #1, .TableCell(tcText, 2, 1)
    Print #1, .TableCell(tcText, 3, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 11: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3800: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1200: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1000: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcColWidth, , 10) = 1000: .TableCell(tcAlign, , 10) = taRightTop
    .TableCell(tcColWidth, , 11) = 1200: .TableCell(tcAlign, , 11) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Fecha"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "TD"
    .TableCell(tcText, 1, 4) = "ND"
    .TableCell(tcText, 1, 5) = "Total Egreso"
    .TableCell(tcText, 1, 6) = "Alim."
    .TableCell(tcText, 1, 7) = "Desech."
    .TableCell(tcText, 1, 8) = "Movil."
    .TableCell(tcText, 1, 9) = "Varios"
    .TableCell(tcText, 1, 10) = "Cod. Cta."
    .TableCell(tcText, 1, 11) = "I.V.A"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & _
              .TableCell(tcText, 1, 9) & "|" & .TableCell(tcText, 1, 10) & "|" & .TableCell(tcText, 1, 11)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 11: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3800: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1200: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1000: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcColWidth, , 10) = 1000: .TableCell(tcAlign, , 10) = taRightTop
    .TableCell(tcColWidth, , 11) = 1200: .TableCell(tcAlign, , 11) = taRightTop
    
    i = 1
    tDesc = 0
    tAlim = 0
    tVarios = 0
    cCta = ""
    NumDoc = 0
    TotAlim = 0
    TotDesc = 0
    TotVarios = 0
    
    Do While Not RS1.EOF
        
        tAlim = 0
        tDesc = 0
        tVarios = 0
        tMovil = 0
        NumDoc = RS1!toc_numdoc: crut = RS1!toc_rutpro
        
        Do While Not RS1.EOF And RS1!toc_numdoc = NumDoc And RS1!toc_rutpro = crut
          '------- Traer Impuesto adicionales
          pctimp = 0
          
          If RS3.State = 1 Then RS3.Close
          RS3.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient

          RS3.Open RutinaLectura.DocProImp(1, RS1!toc_rutpro, RS1!toc_tipdoc, RS1!toc_numdoc, RS1!dec_numlin, RS1!dec_codmer), vg_db, adOpenStatic
          If RS3.EOF Then RS3.Close: Set RS3 = Nothing Else pctimp = RS3!imd_monimp: RS3.Close: Set RS3 = Nothing
            
            cCta = RS1!pro_ctacon
            
            If Trim(RS1!cuenta) = "Alimentos" Then
                
                tAlim = tAlim + RS1!dec_ptotal + pctimp + RS1!toc_fledoc
                TotAlim = TotAlim + RS1!dec_ptotal + pctimp + RS1!toc_fledoc
            
            End If
            
            If Trim(RS1!cuenta) = "Desechables" Then
                
                tDesc = tDesc + RS1!dec_ptotal + pctimp + RS1!toc_fledoc
                TotDesc = TotDesc + RS1!dec_ptotal + pctimp + RS1!toc_fledoc
            
            End If
            
            If Trim(RS1!cuenta) = "Movilizacion" Then
                
                tMovil = tMovil + RS1!dec_ptotal + pctimp + RS1!toc_fledoc
                TotMovil = TotMovil + RS1!dec_ptotal + pctimp + RS1!toc_fledoc
            
            End If

            If Trim(RS1!cuenta) = "Varios" Then
                
                tVarios = tVarios + RS1!dec_ptotal + pctimp + RS1!toc_fledoc
                TotVarios = TotVarios + RS1!dec_ptotal + pctimp + RS1!toc_fledoc
                i = i + 1
            
            End If
            
            .TableCell(tcText, i, 1) = RS1!toc_fecemi
            .TableCell(tcText, i, 2) = "(" & fg_PintaRut(RS1!toc_rutpro) & ")" & " " & Mid(Trim(RS1!prv_nombre), 1, 30)
            .TableCell(tcText, i, 3) = Trim(RS1!toc_tipdoc)
            .TableCell(tcText, i, 4) = RS1!toc_numdoc
            
            If Trim(RS1!cuenta) = "Varios" Then
                
                .TableCell(tcText, i, 5) = 0
                .TableCell(tcText, i, 6) = 0
                .TableCell(tcText, i, 7) = 0
                .TableCell(tcText, i, 8) = 0
                .TableCell(tcText, i, 9) = Format(tVarios, fg_Pict(6, vg_DPr))
                .TableCell(tcText, i, 10) = RS1!pro_ctacon
                .TableCell(tcText, i, 11) = 0
            
            Else
                
                .TableCell(tcText, i, 5) = Format(RS1!toc_totdoc, fg_Pict(6, vg_DPr))
                .TableCell(tcText, i, 6) = Format(tAlim, fg_Pict(6, vg_DPr))
                .TableCell(tcText, i, 7) = Format(tDesc, fg_Pict(6, vg_DPr))
                .TableCell(tcText, i, 8) = Format(tMovil, fg_Pict(6, vg_DPr))
                .TableCell(tcText, i, 9) = 0
                .TableCell(tcText, i, 10) = " "
                .TableCell(tcText, i, 11) = Format(RS1!toc_ivadoc, fg_Pict(6, vg_DPr))
            
            End If
            
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                      .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & _
                      .TableCell(tcText, i, 9) & "|" & .TableCell(tcText, i, 10) & "|" & .TableCell(tcText, i, 11)
            
            tIVA = RS1!toc_ivadoc
            tEgre = RS1!toc_totdoc
            
            RS1.MoveNext
            If RS1.EOF Then Exit Do
        
        Loop
        
        i = i + 1
        TotIva = TotIva + tIVA
        TotEgre = TotEgre + tEgre
    
    Loop
    
    RS1.Close
    Set RS1 = Nothing
    
    RS2.Close
    Set RS2 = Nothing
    .TableCell(tcRows) = i - 1
    .TableBorder = tbAll
    .EndTable
    .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 11: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3800: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1200: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1000: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcColWidth, , 10) = 1000: .TableCell(tcAlign, , 10) = taRightTop
    .TableCell(tcColWidth, , 11) = 1200: .TableCell(tcAlign, , 11) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    '--- Redondea a 0 los valores (vg_dpr), sino ocupar variable (vg_dca)
    .TableCell(tcText, 1, 2) = "Totales"
    .TableCell(tcText, 1, 5) = Format(TotEgre, fg_Pict(6, vg_DPr))
    .TableCell(tcText, 1, 6) = Format(TotAlim, fg_Pict(6, vg_DPr))
    .TableCell(tcText, 1, 7) = Format(TotDesc, fg_Pict(6, vg_DPr))
    .TableCell(tcText, 1, 8) = Format(TotMovil, fg_Pict(6, vg_DPr))
    .TableCell(tcText, 1, 9) = Format(TotVarios, fg_Pict(6, vg_DPr))
    .TableCell(tcText, 1, 11) = Format(TotIva, fg_Pict(6, vg_DPr))
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & _
              .TableCell(tcText, 1, 9) & "|" & .TableCell(tcText, 1, 10) & "|" & .TableCell(tcText, 1, 11)
    .TableBorder = tbBox
    .EndTable
    .EndDoc

End With

Close #1
fg_descarga
Preview.Show 1

Exit Function
Error_SalirFoFi:
    fg_descarga
    MsgBox "Error:" & Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Resume Next

End Function

Sub I_VenDirect(casnom As String, fecini As String, fecter As String, codbod As Long, nombod As String, TipBus As Long)
'----------------------------------------------------
'---Nombre de Informe: Informe de Venta Directa por cliente
'---Creador: Miguel Solorza P.
'---Fecha de Crecación: 31-08-2004.
'----------------------------------------------------
Dim RS1 As New ADODB.Recordset
Dim cVen As String, tVen As Double, tTot As Double, cCas As String
On Local Error GoTo Error_SalirVen
fg_carga ""
MsgTitulo = "Informe Venta Directa"
Casino = Trim(Mid(casnom, 1, InStr(1, casnom, "|") - 1))
nomcas = Trim(Mid(casnom, InStr(1, casnom, "|") + 1, Len(casnom)))
sql1 = IIf(vg_tipbase = "1", " cdate('" & fecini & "') ", " '" & Format(fecini, "yyyymmdd") & "' ")
sql2 = IIf(vg_tipbase = "1", " cdate('" & fecter & "') ", " '" & Format(fecter, "yyyymmdd") & "' ")
If TipBus = 1 Then '--- Un cliente
    RS1.Open "SELECT a.tov_codcas,b.dev_codmer,a.tov_codbod,sum(b.dev_canmer) AS totalcantidad,sum(b.dev_ptotal) AS  ptotal, c.pro_nombre,d.uni_nomcor,e.cli_nombre " & _
             "FROM b_totventas a, b_detventas b, b_productos c, a_unidad d ,b_clientes e " & "WHERE (a.tov_fecemi >= " & sql1 & " AND  a.tov_fecemi <= " & sql2 & ") AND (a.tov_tipdoc = 'FA' OR a.tov_tipdoc = 'FE' or a.tov_tipdoc = 'GD') " & _
             "AND a.tov_numdoc = b.dev_numdoc AND a.tov_codbod = " & codbod & " AND a.tov_codcas = '" & fg_DespintaRut(I_VenDir.fpText1(1).text) & "' " & _
             "AND a.tov_tipdoc = b.dev_tipdoc AND c.pro_coduni = d.uni_codigo AND c.pro_codigo = b.dev_codmer AND a.tov_estdoc <> 'A' AND a.tov_estdoc <> 'P' " & _
             "AND a.tov_codcas = e.cli_codigo " & _
             "GROUP BY a.tov_codcas,b.dev_codmer,a.tov_codbod,b.dev_codmer,b.dev_ptotal,c.pro_nombre,d.uni_nomcor, e.cli_nombre", vg_db, adOpenStatic
ElseIf TipBus = 2 Then '--- Todos los clientes
    RS1.Open "SELECT a.tov_codcas,b.dev_codmer,a.tov_codbod,sum(b.dev_canmer) AS totalcantidad,sum(b.dev_ptotal) AS  ptotal, c.pro_nombre,d.uni_nomcor,e.cli_nombre " & _
             "FROM b_totventas a, b_detventas b, b_productos c, a_unidad d, b_clientes e " & "WHERE (a.tov_fecemi >= " & sql1 & " AND a.tov_fecemi <= " & sql2 & ") AND (a.tov_tipdoc = 'FA' OR a.tov_tipdoc = 'FE' or a.tov_tipdoc = 'GD') " & _
             "AND a.tov_numdoc = b.dev_numdoc AND a.tov_codbod = " & codbod & " AND a.tov_codcas <> '' " & _
             "AND a.tov_tipdoc = b.dev_tipdoc AND c.pro_coduni = d.uni_codigo AND c.pro_codigo = b.dev_codmer AND a.tov_estdoc <> 'A' AND a.tov_estdoc <> 'P' " & _
             "AND a.tov_codcas = e.cli_codigo " & _
             "GROUP BY a.tov_codcas,b.dev_codmer,a.tov_codbod,b.dev_codmer,b.dev_ptotal,c.pro_nombre,d.uni_nomcor,e.cli_nombre", vg_db, adOpenStatic
End If
If RS1.EOF Then fg_descarga: MsgBox "No existen datos para la consulta...", vbExclamation + vbOKOnly, MsgTitulo: RS1.Close: Set RS1 = Nothing: Exit Sub
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .TextAlign = taLeftTop
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = "Ventas por Período"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Contrato": .TableCell(tcText, 1, 2) = Casino & " " & nomcas
    .TableCell(tcText, 2, 1) = "Bodega": .TableCell(tcText, 2, 2) = Trim(nombod)
    .TableCell(tcText, 3, 1) = "Período": .TableCell(tcText, 3, 2) = fecini & " - " & fecter
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    .StartTable
    .TableCell(tcCols) = 5: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 2000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 2000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Cantidad"
    .TableCell(tcText, 1, 4) = "Unidad"
    .TableCell(tcText, 1, 5) = "Total"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 5: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 2000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 2000: .TableCell(tcAlign, , 5) = taRightTop
    i = 1: cVen = "": tVen = 0: tTot = 0: cCas = ""
    Do While Not RS1.EOF
        If RS1!cli_nombre <> cVen Then
            If cVen <> "" Then
                .TableCell(tcFontBold, i) = True
                .TableCell(tcText, i, 2) = "Total Cliente " & fg_PintaRut(cCas)
                .TableCell(tcText, i, 5) = Format(tVen, fg_Pict(8, vg_DPr)): tVen = 0:
                Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & .TableCell(tcText, i, 5)
                i = i + 1
            End If
            If cVen <> "" Then i = i + 1
            .TableCell(tcFontBold, i) = True: .TableCell(tcColSpan, i, 1) = 5
            .TableCell(tcText, i, 1) = fg_PintaRut(RS1!tov_codcas) & "  " & RS1!cli_nombre: cVen = RS1!cli_nombre: cCas = RS1!tov_codcas
            Print #1, .TableCell(tcText, i, 1)
             i = i + 1
        End If
        .TableCell(tcText, i, 1) = RS1!dev_codmer
        .TableCell(tcText, i, 2) = RS1!pro_nombre
        .TableCell(tcText, i, 3) = Format(RS1!totalcantidad, fg_Pict(6, vg_DCa))
        .TableCell(tcText, i, 4) = RS1!uni_nomcor
        .TableCell(tcText, i, 5) = Format(RS1!ptotal, fg_Pict(8, vg_DPr))
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5)
        tVen = tVen + RS1!ptotal
        tTot = tTot + RS1!ptotal
        RS1.MoveNext
        i = i + 1
    Loop
    If tTot <> 0 Then
        .TableCell(tcFontBold, i) = True
        .TableCell(tcText, i, 2) = "Total Cliente " & fg_PintaRut(cCas)
        .TableCell(tcText, i, 5) = Format(tVen, fg_Pict(8, vg_DPr))
        Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & .TableCell(tcText, i, 5)
        i = i + 1
        .TableCell(tcFontBold, i) = True
        .TableCell(tcText, i, 2) = "Total General"
        .TableCell(tcText, i, 5) = Format(tTot, fg_Pict(8, vg_DPr))
        Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & .TableCell(tcText, i, 5)
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
fg_descarga
Preview.Show 1
Exit Sub
Error_SalirVen:
    fg_descarga
    MsgBox "Error:" & Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Sub I_MerPer(casnom As String, fecini As String, fecter As String, codmer As Long, codbod As Long, TipMer As String)
'----------------------------------------------------
'---Nombre de Informe: Informe de mermas por período
'---Creador: Miguel Solorza P.
'---Fecha de Crecación: 30-08-2004.
'----------------------------------------------------
Dim RS1 As New ADODB.Recordset
Dim cAju As String, tAju As Double, tTot As Double
Dim cSql As String, sql1 As String, sql2 As String
On Local Error GoTo Error_SalirMer
fg_carga ""
MsgTitulo = "Informe de mermas por Período"
Casino = Trim(Mid(casnom, 1, InStr(1, casnom, "|") - 1))
nomcas = Trim(Mid(casnom, InStr(1, casnom, "|") + 1, Len(casnom)))
cSql = IIf(codmer = 0, " ", " and a.tov_codser = " & codmer & " ")            'si el filto de tipo de merma es todas
sql1 = IIf(vg_tipbase = "1", " cdate('" & fecini & "') ", " '" & Format(fecini, "yyyymmdd") & "' ")
sql2 = IIf(vg_tipbase = "1", " cdate('" & fecter & "') ", " '" & Format(fecter, "yyyymmdd") & "' ")
RS1.Open "SELECT a.tov_codser,b.dev_codmer,a.tov_codbod,sum(b.dev_canmer) AS totalcantidad,sum(b.dev_ptotal) AS  ptotal, c.pro_nombre,d.uni_nomcor, e.aju_nombre " & _
         "FROM b_totventas a, b_detventas b, b_productos c, a_unidad d, a_tipoajuste e " & "WHERE (a.tov_fecemi >= " & sql1 & " AND a.tov_fecemi <= " & sql2 & ") " & cSql & " AND a.tov_tipdoc = 'ME' " & _
         "AND a.tov_numdoc = b.dev_numdoc AND a.tov_codbod = " & codbod & " AND a.tov_rutcli = '" & Casino & "' " & _
         "AND a.tov_tipdoc = b.dev_tipdoc AND c.pro_coduni = d.uni_codigo AND (a.tov_codser = e.aju_codigo AND e.aju_tipaju = 0) AND c.pro_codigo = b.dev_codmer " & _
         "AND a.tov_estdoc <> 'A' AND a.tov_estdoc <> 'P' AND a.tov_rutcli = b.dev_rutcli " & _
         "GROUP BY a.tov_codser,b.dev_codmer,a.tov_codbod,b.dev_codmer,b.dev_ptotal,c.pro_nombre,d.uni_nomcor,e.aju_nombre", vg_db, adOpenStatic
If RS1.EOF Then fg_descarga: MsgBox "No existen datos para la consulta...", vbExclamation + vbOKOnly, MsgTitulo: RS1.Close: Set RS1 = Nothing: Exit Sub
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .TextAlign = taLeftTop
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = "Mermas por Período"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1200: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Contrato": .TableCell(tcText, 1, 2) = ": " & Casino & " " & nomcas
    .TableCell(tcText, 2, 1) = "Tipo Merma": .TableCell(tcText, 2, 2) = ": " & Trim(TipMer)
    .TableCell(tcText, 3, 1) = "Período": .TableCell(tcText, 3, 2) = ": " & fecini & " - " & fecter
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    .StartTable
    .TableCell(tcCols) = 5: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 2000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 2000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Cantidad"
    .TableCell(tcText, 1, 4) = "Unidad"
    .TableCell(tcText, 1, 5) = "Total"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & _
              .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 5: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 2000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 2000: .TableCell(tcAlign, , 5) = taRightTop
    i = 1: cAju = "": tAju = 0: tTot = 0
    Do While Not RS1.EOF
        If RS1!aju_nombre <> cAju Then
            If cAju <> "" Then
                .TableCell(tcFontBold, i) = True
                .TableCell(tcText, i, 2) = "Total " & cAju
                .TableCell(tcText, i, 5) = Format(tAju, fg_Pict(8, vg_DPr)): tAju = 0
                Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & .TableCell(tcText, i, 5)
                i = i + 1
            End If
            If cAju <> "" Then i = i + 1
            .TableCell(tcFontBold, i) = True: .TableCell(tcColSpan, i, 1) = 5
            .TableCell(tcText, i, 1) = RS1!aju_nombre: cAju = RS1!aju_nombre
            Print #1, .TableCell(tcText, i, 1)
            i = i + 1
        End If
        .TableCell(tcText, i, 1) = RS1!dev_codmer
        .TableCell(tcText, i, 2) = RS1!pro_nombre
        .TableCell(tcText, i, 3) = Format(RS1!totalcantidad, fg_Pict(6, vg_DCa))
        .TableCell(tcText, i, 4) = RS1!uni_nomcor
        .TableCell(tcText, i, 5) = Format(RS1!ptotal, fg_Pict(8, vg_DPr))
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & _
                  .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5)
        tAju = tAju + RS1!ptotal
        tTot = tTot + RS1!ptotal
        RS1.MoveNext
        i = i + 1
    Loop
    If tTot <> 0 Then
        .TableCell(tcFontBold, i) = True
        .TableCell(tcText, i, 2) = "Total " & cAju
        .TableCell(tcText, i, 5) = Format(tAju, fg_Pict(8, vg_DPr))
        Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & .TableCell(tcText, i, 5)
        i = i + 1
        .TableCell(tcFontBold, i) = True
        .TableCell(tcText, i, 2) = "Total General"
        .TableCell(tcText, i, 5) = Format(tTot, fg_Pict(8, vg_DPr))
        Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & .TableCell(tcText, i, 5)
        i = i + 1
        
    End If
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
fg_descarga
Preview.Show 1
Exit Sub
Error_SalirMer:
    fg_descarga
    MsgBox "Error:" & Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Sub I_SalidasDevolBod(cencos As String, codreg As String, codser As String, fecini As String, fecter As String)
Dim aAp As String, aAD As String, auxreg As Long, auxser As Long, nomcen As String
Dim StrImp As String, StrImpb As String, tipdoc As String, sql1 As String, sql2 As String
Dim cCta As String, tCta As Double, tSer As Double, tTot As Double
'----------------------------------------------------
'---Nombre de Informe: Informe de Salidas / Devoluciones a Producción
'---Creador: Miguel Solorza P.(Corrección Alexis Morgado)
'---Fecha de Crecación: 30-08-2004.
'----------------------------------------------------
fg_carga ""
MsgTitulo = "Informes de Producción"
On Local Error GoTo Error_Salir
'---------------------Carga de Datos ------------------------
aAp = Trim(vg_NUsr) & "_tmp_InfSalidas"
aAD = Trim(vg_NUsr) & "_tmp_InfDevol"
'---Cheque la existencia de tablas temporales. Si existen las elimina---
fg_CheckTmp aAp
fg_CheckTmp aAD
vg_db.BeginTrans
sql1 = IIf(vg_tipbase = "1", " cdate('" & fecini & "') ", " '" & Format(fecini, "yyyymmdd") & "' ")
sql2 = IIf(vg_tipbase = "1", " cdate('" & fecter & "') ", " '" & Format(fecter, "yyyymmdd") & "' ")
vg_db.Execute "SELECT d.dev_tipdoc, a.tov_codreg, a.tov_codser, d.dev_codmer, b.pro_nombre, b.pro_ctacon, c.uni_nomcor, SUM(d.dev_canmer) AS totalcantidad, SUM(d.dev_ptotal) AS ptotal INTO " & aAp & " " & _
              "FROM b_totventas a, b_detventas d, b_productos b, a_unidad c " & _
              "WHERE a.tov_rutcli = '" & cencos & "' " & _
              "AND   a.tov_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
              "AND   a.tov_tipdoc = '" & IIf(I_SalBod.Combo1(2).ListIndex = 5, "DP", "SP") & "' " & _
              "AND   a.tov_codbod = " & vg_codbod & " AND a.tov_estdoc <> 'A' AND a.tov_estdoc <> 'P' AND d.dev_canmer <> 0 AND (a.tov_fecpro >= " & sql1 & " AND a.tov_fecpro <= " & sql2 & ") " & " AND tov_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
              "AND   d.dev_rutcli = a.tov_rutcli AND d.dev_tipdoc = a.tov_tipdoc AND d.dev_numdoc = a.tov_numdoc AND d.dev_codmer = b.pro_codigo AND b.pro_coduni = c.uni_codigo " & _
              "GROUP BY d.dev_tipdoc, a.tov_codreg, a.tov_codser, d.dev_codmer, b.pro_nombre, b.pro_ctacon, c.uni_nomcor"
vg_db.Execute "SELECT d.dev_tipdoc, a.tov_codreg, a.tov_codser, d.dev_codmer, b.pro_nombre, b.pro_ctacon, c.uni_nomcor, SUM(d.dev_canmer) AS totalcantidad, SUM(d.dev_ptotal) AS ptotal INTO " & aAD & " " & _
              "FROM b_totventas a, b_detventas d, b_productos b, a_unidad c " & _
              "WHERE a.tov_rutcli = '" & cencos & "' AND a.tov_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND a.tov_tipdoc = 'DP' " & _
              "AND   a.tov_codbod = " & vg_codbod & " AND a.tov_estdoc <> 'A' AND a.tov_estdoc <> 'P' AND d.dev_canmer <> 0 AND (a.tov_fecpro >= " & sql1 & " AND a.tov_fecpro <= " & sql2 & ") " & " AND tov_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
              "AND d.dev_rutcli = a.tov_rutcli AND d.dev_tipdoc = a.tov_tipdoc AND d.dev_numdoc = a.tov_numdoc AND d.dev_codmer = b.pro_codigo AND b.pro_coduni = c.uni_codigo " & _
              "GROUP BY d.dev_tipdoc, a.tov_codreg, a.tov_codser, d.dev_codmer, b.pro_nombre, b.pro_ctacon, c.uni_nomcor"
vg_db.CommitTrans
vg_db.BeginTrans
If vg_tipbase = "1" Then
   vg_db.Execute "ALTER TABLE " & aAp & " ADD COLUMN Cuenta CHAR(20)"
Else
   vg_db.Execute "ALTER TABLE " & aAp & " ADD  Cuenta varchar(20)"
End If
vg_db.Execute "UPDATE " & aAp & " SET Cuenta='Alimentos' WHERE pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
vg_db.Execute "UPDATE " & aAp & " SET Cuenta='Desechables' WHERE pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
vg_db.Execute "UPDATE " & aAp & " SET Cuenta='Otros' WHERE (pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctagastos"), ";", "','") & "') OR pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctagastos2"), ";", "','") & "'))"
If I_SalBod.Combo1(2).ListIndex = 6 Then
    If vg_tipbase = "1" Then
       vg_db.Execute "UPDATE " & aAp & " INNER JOIN " & aAD & " b ON (" & aAp & ".dev_codmer=b.dev_codmer) AND (" & aAp & ".tov_codser=b.tov_codser) AND (" & aAp & ".tov_codreg=b.tov_codreg) SET " & aAp & ".totalcantidad= " & aAp & ".totalcantidad-b.totalcantidad," & aAp & ".ptotal=" & aAp & ".ptotal-b.ptotal"
    Else
       vg_db.Execute "UPDATE " & aAp & " SET " & aAp & ".totalcantidad = " & aAp & ".totalcantidad-b.totalcantidad, " & aAp & ".ptotal = " & aAp & ".ptotal-b.ptotal FROM  " & aAp & ", " & aAD & " b WHERE " & aAp & ".dev_codmer = b.dev_codmer AND " & aAp & ".tov_codser = b.tov_codser AND " & aAp & ".tov_codreg = b.tov_codreg "
    End If
End If
vg_db.CommitTrans
'---------------------Fin Carga de Datos -----------------
RS1.Open "SELECT a.*, b.*, c.* FROM " & aAp & " a, a_servicio b, a_regimen c WHERE a.tov_codreg=c.reg_codigo AND a.tov_codser=b.ser_codigo ORDER BY b.ser_orden, a.cuenta, a.dev_codmer", vg_db, adOpenStatic
If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: MsgBox "No existen datos para consulta...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .TextAlign = taLeftTop
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    If I_SalBod.Combo1(2).ListIndex = 4 Then
        .TableCell(tcText, 1, 1) = "Resumen de Salidas para Producción"
    ElseIf I_SalBod.Combo1(2).ListIndex = 5 Then
        .TableCell(tcText, 1, 1) = "Resumen de Devoluciónes de Producción"
    ElseIf I_SalBod.Combo1(2).ListIndex = 6 Then
        .TableCell(tcText, 1, 1) = "Resumen de Salidas Menos Devoluciones de Producción"
    End If
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Contrato"
    RS2.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & cencos & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If Not RS2.EOF Then .TableCell(tcFontBold, 1, 2) = True: .TableCell(tcText, 1, 2) = RS2!cli_codigo & " " & Trim(RS2!cli_nombre): nomcen = Trim(RS2!cli_nombre)
    RS2.Close: Set RS2 = Nothing
    .TableCell(tcText, 2, 1) = "Período": .TableCell(tcText, 2, 2) = fecini & " - " & fecter
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    .TableBorder = tbNone
    .EndTable
    .StartTable
    .TableCell(tcCols) = 5: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 2000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 2000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Cantidad"
    .TableCell(tcText, 1, 4) = "Unidad"
    .TableCell(tcText, 1, 5) = "Total"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 5: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 2000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 2000: .TableCell(tcAlign, , 5) = taRightTop
    i = 1: cCta = "": tSer = 0: tCta = 0: tTot = 0: auxreg = 0: auxser = 0
    Do While Not RS1.EOF
        If RS1!ser_codigo <> auxser Or RS1!reg_codigo <> auxreg Or RS1!cuenta <> cCta Then
           If cCta <> "" Then
              .TableCell(tcFontBold, i) = True
              .TableCell(tcText, i, 2) = "Total " & cCta
              .TableCell(tcText, i, 5) = Format(tCta, fg_Pict(6, vg_DPr))
              Print #1, "|" & .TableCell(tcText, i, 2) & "|||" & .TableCell(tcText, i, 5)
              tCta = 0: i = i + 1
           End If
           If RS1!ser_codigo <> auxser Or RS1!reg_codigo <> auxreg Then
              If auxser > 0 Then
                 .TableCell(tcFontBold, i) = True
                 .TableCell(tcText, i, 2) = "Total " & Trim(RS1!ser_nombre)
                 .TableCell(tcText, i, 5) = Format(tSer, fg_Pict(6, vg_DPr))
                 Print #1, "|" & .TableCell(tcText, i, 2) & "|||" & .TableCell(tcText, i, 5)
                 auxser = 0: i = i + 2
              End If
              If auxser > 0 Then i = i + 1
              .TableCell(tcFontBold, i) = True: .TableCell(tcColSpan, i, 1) = 5
              .TableCell(tcText, i, 1) = "(" & RS1!reg_codigo & ") " & Trim(RS1!reg_nombre) & " (" & RS1!ser_codigo & ") " & Trim(RS1!ser_nombre)
              Print #1, .TableCell(tcText, i, 1)
              auxreg = RS1!reg_codigo: auxser = RS1!ser_codigo: i = i + 1
           End If
           If cCta <> "" Then i = i + 1
           .TableCell(tcFontBold, i) = True: .TableCell(tcColSpan, i, 1) = 5
           .TableCell(tcText, i, 1) = IIf(IsNull(RS1!cuenta), "No esta bien definido cuenta contable", RS1!cuenta)
           Print #1, .TableCell(tcText, i, 1)
            cCta = IIf(IsNull(RS1!cuenta), "No esta bien definido cuenta contable", RS1!cuenta): i = i + 1
        End If
        .TableCell(tcText, i, 1) = RS1!dev_codmer
        .TableCell(tcText, i, 2) = RS1!pro_nombre
        .TableCell(tcText, i, 3) = Format(RS1!totalcantidad, fg_Pict(6, vg_DCa))
        .TableCell(tcText, i, 4) = RS1!uni_nomcor
        .TableCell(tcText, i, 5) = Format(RS1!ptotal, fg_Pict(6, vg_DPr))
        Print #1, .TableCell(tcText, i, 1) & "|"; .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5)
        tSer = tSer + RS1!ptotal
        tCta = tCta + RS1!ptotal
        tTot = tTot + RS1!ptotal
        RS1.MoveNext
        i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing
    If tTot <> 0 Then
        .TableCell(tcFontBold, i) = True
        .TableCell(tcText, i, 2) = "Total " & cCta
        .TableCell(tcText, i, 5) = Format(tCta, fg_Pict(6, vg_DPr))
        Print #1, "|" & .TableCell(tcText, i, 2) & "|||" & .TableCell(tcText, i, 5)
        Print #1, " "
        i = i + 1
        .TableCell(tcFontBold, i) = True
        .TableCell(tcText, i, 2) = "Total " & CSer
        .TableCell(tcText, i, 5) = Format(tSer, fg_Pict(6, vg_DPr))
        Print #1, "|" & .TableCell(tcText, i, 2) & "|||" & .TableCell(tcText, i, 5)
        Print #1, " "
        i = i + 1
        .TableCell(tcFontBold, i) = True
        .TableCell(tcText, i, 2) = "Total General"
        .TableCell(tcText, i, 5) = Format(tTot, fg_Pict(6, vg_DPr))
        Print #1, "|" & .TableCell(tcText, i, 2) & "|||" & .TableCell(tcText, i, 5)
        Print #1, " "
        i = i + 1
        Print #1, " "
    End If
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
    Close #1
End With
fg_descarga
Preview.Show 1
Exit Sub
Error_Salir:
    fg_descarga
    MsgBox "Error:" & Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Sub I_Pendocpro()

Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long, auxnumdoc As Long, pctdes As Double
Dim auxrutpro As String, auxnompro As String, auxtipdoc As String, nomcta As String, nomcli As String, sql1 As String, sql2 As String, sql3 As String
Dim canali As Double, candes As Double, cangrl As Double, canotros As Double, totali As Double, totdes As Double, tototros As Double, totgrl As Double, pctimp As Double
Dim tanali As Double, tandes As Double, tangrl As Double
Dim v_tipdoc As String, j As Long

On Local Error GoTo Error_Pendientes

fg_carga ""
MsgTitulo = "Informe Documentos Pendientes"
'----------------------------------------------------
'---Nombre de Informe: Informe Documentos pendientes Proveedor
'---Creador: Miguel Solorza P.
'---Fecha de Crecación: 23-08-2004.
'----------------------------------------------------
Preview.Refresh
With Preview.VSPrinter
    
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    .FontSize = 7
    LogoEmp
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    
    If I_DocPen.Combo1(0).ListIndex = 0 Then
        
        .TableCell(tcText, 1) = "Informe de Guías de Despacho Pendientes"
        v_tipdoc = "GD"
    
    ElseIf I_DocPen.Combo1(0).ListIndex = 1 Then
        
        .TableCell(tcText, 1) = "Informe de Solicitudes de Nota de Crédito Pendientes"
        v_tipdoc = "SN"
    
    End If
    
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    
    .StartTable
    .TableCell(tcCols) = 2
    .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 1000
    .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000
    .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcFontBold, 1, 1) = True
    .TableCell(tcText, 1, 1) = "Contrato"
    .TableCell(tcFontBold, 2, 1) = True
    .TableCell(tcText, 2, 1) = "Bodega"
     Print #1, .TableCell(tcText, 1, 1)
     Print #1, .TableCell(tcText, 2, 1)
     
     If RS1.State = 1 Then RS1.Close
     RS1.CursorLocation = adUseClient
     vg_db.CursorLocation = adUseClient
     RS1.Open RutinaLectura.Cliente(1, MuestraCasino(1), ""), vg_db, adOpenStatic
     If Not RS1.EOF Then .TableCell(tcFontBold, 1, 2) = True: .TableCell(tcText, 1, 2) = ": " & Trim(RS1!cli_nombre) & " (" & Trim(RS1!cli_codigo) & ")"
     RS1.Close: Set RS1 = Nothing
     
     If RS1.State = 1 Then RS1.Close
     RS1.CursorLocation = adUseClient
     vg_db.CursorLocation = adUseClient
     RS1.Open RutinaLectura.Bodega(3, vg_codbod, ""), vg_db, adOpenStatic
     If Not RS1.EOF Then .TableCell(tcFontBold, 2, 2) = True: .TableCell(tcText, 2, 2) = ": " & Trim(RS1!bod_nombre) & " (" & RS1!bod_codigo & ")"
     RS1.Close: Set RS1 = Nothing
          
     Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
     Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
     .TableBorder = tbNone
     .EndTable
     .text = Chr(13): .text = Chr(13)
    
    canali = 0
    candes = 0
    cangrl = 0
    canotros = 0
    tototros = 0
    tanali = 0
    tandes = 0
    tangrl = 0
    pctdes = 0
    auxtipdoc = ""
    auxnomdoc = ""
    auxnompro = ""
    
    .StartTable
    .TableCell(tcCols) = 10: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 900: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 900: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 400: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 1300: .TableCell(tcAlign, , 5) = taLeftTop
    .TableCell(tcColWidth, , 6) = 3000: .TableCell(tcAlign, , 6) = taLeftTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1000: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcColWidth, , 10) = 1000: .TableCell(tcAlign, , 10) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Fecha Emisión"
    .TableCell(tcText, 1, 2) = "ND"
    .TableCell(tcText, 1, 3) = IIf(v_tipdoc = "GD", "", "Nº. Factura")
    .TableCell(tcText, 1, 4) = "TD"
    .TableCell(tcText, 1, 5) = "R.U.T"
    .TableCell(tcText, 1, 6) = "Nombre"
    .TableCell(tcText, 1, 7) = "Alim"
    .TableCell(tcText, 1, 8) = "Desech"
    .TableCell(tcText, 1, 9) = "Otros"
    .TableCell(tcText, 1, 10) = "Monto"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & .TableCell(tcText, 1, 9) & "|" & .TableCell(tcText, 1, 10)
    .TableBorder = tbBoxRows
    .EndTable
    .StartTable
    .TableCell(tcCols) = 10: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 900: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 900: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 400: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 1300: .TableCell(tcAlign, , 5) = taLeftTop
    .TableCell(tcColWidth, , 6) = 3000: .TableCell(tcAlign, , 6) = taLeftTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1000: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcColWidth, , 10) = 1000: .TableCell(tcAlign, , 10) = taRightTop
    
'    If vg_tipbase = "1" Then
'
'       sql1 = "SELECT a.toc_rutpro, a.toc_tipdoc, a.toc_numdoc, a.toc_fecrem, a.toc_docaso, " & _
'              "b.dec_numlin, b.dec_codmer, b.dec_canmer, b.dec_precom, " & _
'              "b.dec_pctdes, c.pro_ctacon, b.dec_canrec, b.dec_prerec, " & _
'              "d.prv_nombre, d.prv_codigo, b.dec_ptotal, dec_ptotrec " & _
'              "FROM  b_totcompras a, b_detcompras b, b_productos c, b_proveedor d " & _
'              "WHERE (a.toc_fecrem >= cdate('" & I_DocPen.Fecha(0).text & "')" & " AND a.toc_fecrem <= cdate('" & I_DocPen.Fecha(1).text & "')) AND a.toc_rutpro=b.dec_rutpro " & _
'              "AND   a.toc_tipdoc = '" & v_tipdoc & "'" & _
'              "AND   a.toc_tipdoc = b.dec_tipdoc " & _
'              "AND   a.toc_numdoc = b.dec_numdoc " & _
'              "AND   b.dec_codmer = c.pro_codigo " & _
'              "AND   b.dec_rutpro = d.prv_codigo " & _
'              "AND   a.toc_codbod = " & vg_codbod & " "
'        sql2 = IIf(v_tipdoc = "GD", "AND (a.toc_docaso='' OR (a.toc_docaso) is null) ", "AND (trim(a.toc_docsnc) = '' OR (a.toc_docsnc) IS NULL) ")
'
'    Else
       
'       sql1 = "SELECT a.toc_rutpro, a.toc_tipdoc, a.toc_numdoc, a.toc_fecrem, a.toc_docaso, " & _
'              "b.dec_numlin, b.dec_codmer, b.dec_canmer, b.dec_precom, " & _
'              "b.dec_pctdes, c.pro_ctacon, b.dec_canrec, b.dec_prerec, " & _
'              "d.prv_nombre, d.prv_codigo, b.dec_ptotal, dec_ptotrec " & _
'              "FROM  b_totcompras a, b_detcompras b, b_productos c, b_proveedor d " & _
'              "WHERE (a.toc_fecrem >= '" & Format(I_DocPen.Fecha(0).text, "yyyymmdd") & "'" & " AND a.toc_fecrem <= '" & Format(I_DocPen.Fecha(1).text, "yyyymmdd") & "') AND a.toc_rutpro = b.dec_rutpro " & _
'              "AND   a.toc_tipdoc = '" & v_tipdoc & "'" & _
'              "AND   a.toc_tipdoc = b.dec_tipdoc " & _
'              "AND   a.toc_numdoc = b.dec_numdoc " & _
'              "AND   b.dec_codmer = c.pro_codigo " & _
'              "AND   b.dec_rutpro = d.prv_codigo " & _
'              "AND   a.toc_codbod = " & vg_codbod & " "
'        sql2 = IIf(v_tipdoc = "GD", "AND (a.toc_docaso = '' OR (a.toc_docaso) IS NULL) ", "AND (LTRIM(a.toc_docsnc) = '' OR (a.toc_docsnc) IS NULL) ")
'
''    End If
'    sql3 = "ORDER BY a.toc_fecrem, a.toc_rutpro, a.toc_tipdoc, a.toc_numdoc, b.dec_numlin, " & _
'          "b.dec_codmer, b.dec_canmer, b.dec_precom, b.dec_pctdes, c.pro_ctacon, " & _
'          "b.dec_canrec, b.dec_prerec, d.prv_nombre, d.prv_codigo"
'    RS1.Open sql1 & sql2 & sql3, vg_db, adOpenStatic
'
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS1 = vg_db.Execute("sgp_Sel_DocumentoPendientesProveedor '" & vg_codbod & "', '" & v_tipdoc & "','" & Format(I_DocPen.Fecha(0).text, "yyyymmdd") & "', '" & Format(I_DocPen.Fecha(1).text, "yyyymmdd") & "'")
    If RS1.EOF Then
        
        fg_descarga
        RS1.Close
        Set RS1 = Nothing
        .TableBorder = tbAll
        .EndDoc
        Close #1
        MsgBox "No existen datos...", vbInformation
        Exit Sub
    
    End If
    
    If Not RS1.EOF Then
       
       i = 1
       auxtipdoc = ""
       auxnomdoc = ""
       auxnompro = ""
       totali = 0
       totdes = 0
       totgrl = 0
       tototros = 0
       
       Do While Not RS1.EOF
            
            auxrutpro = RS1!toc_rutpro
            auxnumdoc = RS1!toc_numdoc
            auxtipdoc = RS1!toc_tipdoc
            canali = 0
            candes = 0
            canotros = 0
            
            Do While Not RS1.EOF And Trim(RS1!toc_rutpro) = Trim(auxrutpro) And RS1!toc_numdoc = auxnumdoc And auxtipdoc = RS1!toc_tipdoc
                  
                  .TableCell(tcText, i, 1) = RS1!toc_fecrem
                  .TableCell(tcText, i, 2) = RS1!toc_numdoc
                  .TableCell(tcText, i, 3) = IIf(v_tipdoc = "GD", "", RS1!toc_docaso)
                  .TableCell(tcText, i, 4) = (RS1!toc_tipdoc)
                  .TableCell(tcText, i, 5) = fg_PintaRut(RS1!prv_codigo)
                  .TableCell(tcText, i, 6) = RS1!prv_nombre
                  pctimp = 0
'                  RS2.Open RutinaLectura.DocProImp(1, RS1!toc_rutpro, RS1!toc_tipdoc, RS1!toc_numdoc, RS1!dec_numlin, RS1!dec_codmer), vg_db, adOpenStatic
'                  If RS2.EOF Then RS2.Close: Set RS2 = Nothing Else pctimp = RS2!imd_monimp: RS2.Close: Set RS2 = Nothing
                  pctimp = RS1!imd_monimp
                  pctdes = 0
                  
                  If RS1!dec_pctdes > 0 Then pctdes = RS1!dec_pctdes
                  
                  If v_tipdoc = "SN" Then
                      
                      If RS1!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
                          
                          canali = Round(canali + ((RS1!dec_ptotal - RS1!dec_ptotrec)) - (((RS1!dec_ptotal - RS1!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr)   'vg_DCa)
                          totali = Round(totali + ((RS1!dec_ptotal - RS1!dec_ptotrec)) - (((RS1!dec_ptotal - RS1!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr)   'vg_DCa)
                      
                      ElseIf RS1!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
                          
                          candes = Round(candes + ((RS1!dec_ptotal - RS1!dec_ptotrec)) - (((RS1!dec_ptotal - RS1!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr)   'vg_DCa)
                          totdes = Round(totdes + ((RS1!dec_ptotal - RS1!dec_ptotrec)) - (((RS1!dec_ptotal - RS1!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr)   'vg_DCa)
                      
                      Else
                         
                         canotros = Round(canotros + ((RS1!dec_ptotal - RS1!dec_ptotrec)) - (((RS1!dec_ptotal - RS1!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr)   'vg_DCa)
                         tototros = Round(tototros + ((RS1!dec_ptotal - RS1!dec_ptotrec)) - (((RS1!dec_ptotal - RS1!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr)   'vg_DCa)
                      
                      End If
                  
                  ElseIf v_tipdoc = "GD" Then
                      
                      If RS1!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
                         
                         canali = canali + ((RS1!dec_canrec * RS1!dec_precom) - (((RS1!dec_canmer) * RS1!dec_precom) * (pctdes / 100)))
                         totali = totali + ((RS1!dec_canrec * RS1!dec_precom) - (((RS1!dec_canmer) * RS1!dec_precom) * (pctdes / 100)))
                      
                      ElseIf RS1!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
                        
                        candes = candes + ((RS1!dec_canrec * RS1!dec_precom) - ((RS1!dec_canrec * RS1!dec_precom) * (pctdes / 100)))
                        totdes = totdes + ((RS1!dec_canrec * RS1!dec_precom) - ((RS1!dec_canrec * RS1!dec_precom) * (pctdes / 100)))
                      
                      Else
                         
                         canotros = canotros + ((RS1!dec_canrec * RS1!dec_precom) - ((RS1!dec_canrec * RS1!dec_precom) * (pctdes / 100)))
                         tototros = tototros + ((RS1!dec_canrec * RS1!dec_precom) - ((RS1!dec_canrec * RS1!dec_precom) * (pctdes / 100)))
                      
                      End If
                  
                  End If
                  
                  totgrl = totgrl + cangrl
                  
                  RS1.MoveNext
                 
                 .TableCell(tcText, i, 7) = Format(canali, fg_Pict(6, 0))
                 .TableCell(tcText, i, 8) = Format(candes, fg_Pict(6, 0))
                 .TableCell(tcText, i, 9) = Format(canotros, fg_Pict(6, 0))
                 .TableCell(tcText, i, 10) = Format(canali + candes + canotros, fg_Pict(6, 0))
                 
                 If RS1.EOF Then Exit Do
            
            Loop
            
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                      .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9) & "|" & .TableCell(tcText, i, 10)
            
            i = i + 1
       
       Loop
    
    Else
        
        i = i + 1
    
    End If
    i = i + 1
    .TableCell(tcFontBold, i) = True: .TableCell(tcRowHeight, i) = 230
    .TableCell(tcText, i, 1) = ""
    .TableCell(tcText, i, 2) = ""
    .TableCell(tcText, i, 3) = "Total"
    .TableCell(tcText, i, 4) = ""
    .TableCell(tcText, i, 5) = ""
    .TableCell(tcText, i, 6) = ""
    .TableCell(tcText, i, 7) = IIf(totali <> 0, Format(totali, fg_Pict(6, 0)), Format(0, fg_Pict(6, 0)))
    .TableCell(tcText, i, 8) = IIf(totdes <> 0, Format(totdes, fg_Pict(6, 0)), Format(0, fg_Pict(6, 0)))
    .TableCell(tcText, i, 9) = IIf(tototros <> 0, Format(tototros, fg_Pict(6, 0)), Format(0, fg_Pict(6, 0)))
    .TableCell(tcText, i, 10) = IIf(totali <> 0 Or totdes <> 0 Or tototros <> 0, Format((totali + totdes + tototros), fg_Pict(6, 0)), Format(0, fg_Pict(6, 0)))
    
    Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
               .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9) & "|" & .TableCell(tcText, i, 10)
    
    RS1.Close
    Set RS1 = Nothing
    Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc

End With
fg_descarga
Preview.Show 1

Exit Sub
Error_Pendientes:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Resume Next
    Close #1
    Exit Sub

End Sub

Sub I_ComprasPer()

Dim TExe As Double, TNet As Double, tIVA As Double, TOtr As Double, TFle As Double
Dim tTot As Double, v_rut As String
Dim TGen As Double, TGeniva As Double, TGenexec As Double, TGenOImp As Double, TGennet As Double, p As Double, TGenfle As Double
Dim fecini As String
Dim fecfin As String
Dim codbod As String
Dim codpro As String
Dim tipdoc As String
Dim op1    As String
Dim op2    As String
Dim op3    As String
Dim op4    As String
Dim RS1 As New ADODB.Recordset
fg_carga ""

On Local Error GoTo Error_Compras

MsgTitulo = "Informe de Compras por Período"
'----------------------------------------------------
'---Nombre de Informe: Informe de Compras por Período
'---Creador: Miguel Solorza P.
'---Fecha de Crecación: 09-08-2004.
'----------------------------------------------------
Preview.Refresh
With Preview.VSPrinter
    
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    LogoEmp
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .FontSize = 7
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .text = Chr(13): .text = Chr(13)
    .TableCell(tcText, 1, 1) = "Compras por Período "
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 1
    .TableCell(tcRows) = 4
    .TableCell(tcColWidth, , 1) = 10500
    .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcFontSize, 1) = 8
    .TableCell(tcFontBold, , 1) = True
    
    i = 1
    fecini = ""
    fecfin = ""
    codbod = ""
    codpro = ""
    tipdoc = ""
    op1 = "0"
    op2 = "0"
    op3 = "0"
    op4 = "0"
    
    If I_ComPer.Check1(0).Value = 1 Then
        
        i = i + 1
        .TableCell(tcText, i, 1) = "Compras entre " & I_ComPer.Fecha(0).text & " y el " & I_ComPer.Fecha(1).text
        Print #1, .TableCell(tcText, i, 1)
        op1 = 1
        fecini = Format(I_ComPer.Fecha(0).text, "yyyymmdd")
        fecfin = Format(I_ComPer.Fecha(1).text, "yyyymmdd")
        
    End If
    
    If I_ComPer.Check1(1).Value = 1 Then
        
        i = i + 1
        .TableCell(tcText, i, 1) = "Bodega : " & Trim(Mid(I_ComPer.Combo1(0).text, 1, 50))
        Print #1, .TableCell(tcText, i, 1)
        op2 = 1
        codbod = fg_codigocbo(I_ComPer.Combo1, 0, 10, 0)
        
    End If
    
    If I_ComPer.Check1(2).Value = 1 Then
        
       op3 = 1
       v_rut = fg_DespintaRut(I_ComPer.fpText1(0).text)
       
       codpro = v_rut
   
    End If
    
    If I_ComPer.Check1(3).Value = 1 Then
        
        i = i + 1
        .TableCell(tcText, i, 1) = "Tipo de Documento : (" & Trim(fg_codigocbo(I_ComPer.Combo1, 1, 2, "")) & ")" & " - " & Trim(Mid(I_ComPer.Combo1(1).text, 1, 50))
        Print #1, .TableCell(tcText, i, 1)
        op4 = 1
        tipdoc = Trim(fg_codigocbo(I_ComPer.Combo1, 1, 2, ""))
        
    End If
    .TableCell(tcRows) = i
    .TableBorder = tbBottom
    .EndTable
    
    vg_Consulta = ""
    v_rut = fg_DespintaRut(I_ComPer.fpText1(0).text)
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
'    RS1.Open "SELECT DISTINCT a.*, b.* FROM b_totcompras a, b_proveedor b " & vg_Consulta & " AND a.toc_rutpro = b.prv_codigo AND a.toc_codbod = " & vg_codbod & "" & " ORDER BY a.toc_rutpro, a.toc_fecrem, a.toc_tipdoc, a.toc_numdoc", vg_db, adOpenStatic
     Set RS1 = vg_db.Execute("sgp_Sel_ResumenComprasxPeriodo '" & op1 & "', '" & fecini & "', '" & fecfin & "', '" & op2 & "', '" & codbod & "', '" & op3 & "', '" & codpro & "', '" & op4 & "', '" & tipdoc & "'")
    If RS1.EOF Then fg_descarga: MsgBox "No existen datos para la consulta...", vbExclamation + vbOKOnly, MsgTitulo: Close #1: RS1.Close: Set RS1 = Nothing: Exit Sub
    
    .StartTable
    .TableCell(tcCols) = 10: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 400: .TableCell(tcAlign, , 1) = taCenterTop
    .TableCell(tcColWidth, , 2) = 800: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 3000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 900: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1200: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1000: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcColWidth, , 10) = 1200: .TableCell(tcAlign, , 10) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True:  .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "TD"
    .TableCell(tcText, 1, 2) = "Doc. Nº"
    .TableCell(tcText, 1, 3) = "Proveedor"
    .TableCell(tcText, 1, 4) = "F.Emisión"
    .TableCell(tcText, 1, 5) = "Exento"
    .TableCell(tcText, 1, 6) = "Neto"
    .TableCell(tcText, 1, 7) = "Flete"
    .TableCell(tcText, 1, 8) = "I.V.A"
    .TableCell(tcText, 1, 9) = "O.Imp."
    .TableCell(tcText, 1, 10) = "Total"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & .TableCell(tcText, 1, 9) & "|" & .TableCell(tcText, 1, 10)
    .TableBorder = tbBox
    .EndTable
    
    .StartTable
    .TableCell(tcCols) = 10: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 400: .TableCell(tcAlign, , 1) = taCenterTop
    .TableCell(tcColWidth, , 2) = 800: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 3000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 900: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1200: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1000: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcColWidth, , 10) = 1200: .TableCell(tcAlign, , 10) = taRightTop
    i = 1
    TGen = 0
    TGeniva = 0
    TGenexec = 0
    TGenOImp = 0
    TGennet = 0
    TGenfle = 0
    
    Do While Not RS1.EOF
        
        v_provee = RS1!toc_rutpro
        TExe = 0
        TNet = 0
        tIVA = 0
        TOtr = 0
        tTot = 0
        p = 0
        TFle = 0
        
        Do While Not RS1.EOF
            
            If RS1!toc_rutpro = v_provee Then
                
                .TableCell(tcText, i, 1) = RS1!toc_tipdoc
                .TableCell(tcText, i, 2) = RS1!toc_numdoc
                .TableCell(tcText, i, 3) = "(" & fg_PintaRut(RS1!toc_rutpro) & ") " & Mid$(RS1!prv_nombre, 1, 20)
                .TableCell(tcText, i, 4) = Format(RS1!toc_fecrem, "dd/mm/yyyy")
                .TableCell(tcText, i, 5) = IIf(Trim(RS1!toc_tipdoc) = "NC" Or Trim(RS1!toc_tipdoc) = "CE", "(" & Format(RS1!toc_exedoc, fg_Pict(8, vg_DPr)) & ")", Format(RS1!toc_exedoc, fg_Pict(8, vg_DPr)))
                .TableCell(tcText, i, 6) = IIf(Trim(RS1!toc_tipdoc) = "NC" Or Trim(RS1!toc_tipdoc) = "CE", "(" & Format(RS1!toc_netdoc, fg_Pict(8, vg_DPr)) & ")", Format(RS1!toc_netdoc, fg_Pict(8, vg_DPr))) 'Format(RS1!toc_netdoc, fg_Pict(8, vg_DPr))
                .TableCell(tcText, i, 7) = IIf(Trim(RS1!toc_tipdoc) = "NC" Or Trim(RS1!toc_tipdoc) = "CE", "(" & Format(RS1!toc_fledoc, fg_Pict(8, vg_DPr)) & ")", Format(RS1!toc_fledoc, fg_Pict(8, vg_DPr))) 'Format(IIf(IsNull(RS1!toc_fledoc), 0, RS1!toc_fledoc), fg_Pict(8, vg_DPr))
                .TableCell(tcText, i, 8) = IIf(Trim(RS1!toc_tipdoc) = "NC" Or Trim(RS1!toc_tipdoc) = "CE", "(" & Format(RS1!toc_ivadoc, fg_Pict(8, vg_DPr)) & ")", Format(RS1!toc_ivadoc, fg_Pict(8, vg_DPr))) 'Format(RS1!toc_ivadoc, fg_Pict(8, vg_DPr))
                .TableCell(tcText, i, 9) = IIf(Trim(RS1!toc_tipdoc) = "NC" Or Trim(RS1!toc_tipdoc) = "CE", "(" & Format(RS1!toc_otrimp, fg_Pict(8, vg_DPr)) & ")", Format(RS1!toc_otrimp, fg_Pict(8, vg_DPr))) 'Format(IIf(IsNull(RS1!toc_otrimp), 0, RS1!toc_otrimp), fg_Pict(8, vg_DPr))
                .TableCell(tcText, i, 10) = IIf(Trim(RS1!toc_tipdoc) = "NC" Or Trim(RS1!toc_tipdoc) = "CE", "(" & Format(RS1!toc_totdoc, fg_Pict(8, vg_DPr)) & ")", Format(RS1!toc_totdoc, fg_Pict(8, vg_DPr)))
                
                Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                          .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9) & "|" & .TableCell(tcText, i, 10)
                
                TExe = TExe + IIf(Trim(RS1!toc_tipdoc) = "NC" Or Trim(RS1!toc_tipdoc) = "CE", RS1!toc_exedoc * -1, RS1!toc_exedoc)
                TNet = TNet + IIf(Trim(RS1!toc_tipdoc) = "NC" Or Trim(RS1!toc_tipdoc) = "CE", RS1!toc_netdoc * -1, RS1!toc_netdoc)
                tIVA = tIVA + IIf(Trim(RS1!toc_tipdoc) = "NC" Or Trim(RS1!toc_tipdoc) = "CE", RS1!toc_ivadoc * -1, RS1!toc_ivadoc)
                TFle = TFle + IIf(Trim(RS1!toc_tipdoc) = "NC" Or Trim(RS1!toc_tipdoc) = "CE", IIf(IsNull(RS1!toc_fledoc), 0, RS1!toc_fledoc * -1), IIf(IsNull(RS1!toc_fledoc), 0, RS1!toc_fledoc))
                TOtr = TOtr + RS1!toc_otrimp
                tTot = tTot + IIf(Trim(RS1!toc_tipdoc) = "NC" Or Trim(RS1!toc_tipdoc) = "CE", RS1!toc_totdoc * -1, RS1!toc_totdoc):
                
                RS1.MoveNext
                i = i + 1
            
            Else
                
                Exit Do
            
            End If
        
        Loop
        
        If tTot <> 0 Then
            
            .TableCell(tcFontBold, i) = True
            .TableCell(tcText, i, 3) = "Total Proveedor"
            .TableCell(tcText, i, 5) = Format(TExe, fg_Pict(8, vg_DPr))
            .TableCell(tcText, i, 6) = Format(TNet, fg_Pict(8, vg_DPr))
            .TableCell(tcText, i, 7) = Format(TFle, fg_Pict(8, vg_DPr))
            .TableCell(tcText, i, 8) = Format(tIVA, fg_Pict(8, vg_DPr))
            .TableCell(tcText, i, 9) = Format(TOtr, fg_Pict(8, vg_DPr))
            .TableCell(tcText, i, 10) = Format(tTot, fg_Pict(8, vg_DPr))
            
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                      .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9) & "|" & .TableCell(tcText, i, 10)
            
            TGen = TGen + tTot
            TGeniva = TGeniva + tIVA
            TGenfle = TGenfle + TFle
            TGenexec = TGenexec + TExe
            TGenOImp = TGenOImp + TOtr
            TGennet = TGennet + TNet
            
            i = i + 2
            p = 0
            Print #1, " "
        
        End If
     
     Loop
     
     i = i + 1
    .TableCell(tcFontBold, i) = True
    .TableCell(tcText, i, 3) = "Total General"
    .TableCell(tcText, i, 5) = Format(TGenexec, fg_Pict(8, vg_DPr))
    .TableCell(tcText, i, 6) = Format(TGennet, fg_Pict(8, vg_DPr))
    .TableCell(tcText, i, 7) = Format(TGenfle, fg_Pict(8, vg_DPr))
    .TableCell(tcText, i, 8) = Format(TGeniva, fg_Pict(8, vg_DPr))
    .TableCell(tcText, i, 9) = Format(TGenOImp, fg_Pict(8, vg_DPr))
    .TableCell(tcText, i, 10) = Format(TGen, fg_Pict(8, vg_DPr))
    Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
              .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9) & "|" & .TableCell(tcText, i, 10)
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc

End With
fg_descarga
Preview.Show 1
Exit Sub
Error_Compras:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
    
End Sub

Sub I_DetalleCom()

Dim TExe As Double, TNet As Double, tIVA As Double, TOtr As Double
Dim tTot As Double, v_rut As String, v_Switch As Double
Dim TGen As Double, TGeniva As Double, TGenexec As Double, TGenOImp As Double, TGennet As Double
Dim RS1 As New ADODB.Recordset
Dim op1 As String
Dim op2 As String
Dim op3 As String
Dim op4 As String
Dim op5 As String
Dim op6 As String
Dim fi  As String
Dim ff  As String
Dim codbod As String
Dim fampro As String
Dim codpro As String
Dim codprov As String
Dim tipdoc  As String

On Local Error GoTo Error_DetalleCom

fg_carga ""
MsgTitulo = "Informe de Detalle Compras por Período"
'----------------------------------------------------
'---Nombre de Informe: Informe detalle Compras Proveedor
'---Creador: Miguel Solorza P.(Corrección Alexis Morgado)
'---Fecha de Crecación: 12-08-2004.
'----------------------------------------------------

Preview.Refresh
With Preview.VSPrinter
    
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    LogoEmp
    .FontSize = 7
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .text = Chr(13): .text = Chr(13)
    .TableCell(tcText, 1, 1) = "Detalle de Compras por Periodo"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 5
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftMiddle
    i = 1
    
    op1 = ""
    op2 = ""
    op3 = ""
    op4 = ""
    op5 = ""
    op6 = ""
    fi = ""
    ff = ""
    codbod = ""
    fampro = ""
    codpro = ""
    codprov = ""
    tipdoc = ""
    
    If I_DetCom.Check1(0).Value = 1 Then
        
        .TableCell(tcRows) = i: .TableCell(tcFontSize, i) = 8: .TableCell(tcFontBold, , 1) = True
        .TableCell(tcText, i, 1) = "Compras entre " & I_DetCom.Fecha(0).text & "   y el  " & I_DetCom.Fecha(1).text
        Print #1, .TableCell(tcText, i, 1)
        i = i + 1
        op1 = "1"
        fi = Format(I_DetCom.Fecha(0).text, "yyyymmdd")
        ff = Format(I_DetCom.Fecha(1).text, "yyyymmdd")
    
    End If
    
    If I_DetCom.Check1(1).Value = 1 Then
        
        .TableCell(tcRows) = i: .TableCell(tcFontSize, i) = 8: .TableCell(tcFontBold, i, 1) = True
        .TableCell(tcText, i, 1) = "Bodega: " & Trim(Mid(I_DetCom.Combo1(0).text, 1, 50))
        Print #1, .TableCell(tcText, i, 1)
        i = i + 1
        op2 = "1"
        codbod = fg_codigocbo(I_DetCom.Combo1, 0, 10, 0)
    
    End If
    
    If I_DetCom.Check1(2).Value = 1 Then
         
        .TableCell(tcRows) = i
        .TableCell(tcFontSize, i) = 8
        .TableCell(tcFontBold, i, 1) = True
        .TableCell(tcText, i, 1) = "Familia de Producto: " & Trim(Mid(I_DetCom.Combo1(1).text, 1, 50))
        Print #1, .TableCell(tcText, i, 1)
        i = i + 1
        op3 = "1"
        fampro = fg_codigocbo(I_DetCom.Combo1, 1, 10, 0)
    
    End If
    
    If I_DetCom.Check1(3).Value = 1 Then
        
        .TableCell(tcRows) = i: .TableCell(tcFontSize, i) = 8: .TableCell(tcFontBold, i, 1) = True
        .TableCell(tcText, i, 1) = "Producto: " & Trim(Mid(I_DetCom.Combo1(2).text, 1, 50))
        Print #1, .TableCell(tcText, i, 1)
        i = i + 1
        op4 = "1"
        codpro = fg_Quitachar(Trim(fg_codigocbo(I_DetCom.Combo1, 2, 20, 0)), "(")
    
    End If
    
    If I_DetCom.Check1(4).Value = 1 Then
        
       op5 = "1"
       v_rut = fg_DespintaRut(I_DetCom.fpText1(0).text)
       codprov = v_rut
    
    End If
    
    If I_DetCom.Check1(5).Value = 1 Then
        
        .TableCell(tcRows) = i: .TableCell(tcFontSize, i) = 8: .TableCell(tcFontBold, i, 1) = True
        .TableCell(tcText, i, 1) = "Tipo de Documento: (" & Trim(fg_codigocbo(I_DetCom.Combo1, 3, 2, "")) & ")" & " - " & Trim(Mid(I_DetCom.Combo1(3).text, 1, 50))
        Print #1, .TableCell(tcText, i, 1)
        i = i + 1
        op6 = "1"
        tipdoc = fg_codigocbo(I_DetCom.Combo1, 3, 2, "")
    
    End If
    
'    vg_Consulta = ""
'    v_rut = fg_DespintaRut(I_DetCom.fpText1(0).text)
'
'    If I_DetCom.Check1(0).Value = 1 Then
'
'        vg_Consulta = IIf(vg_tipbase = "1", "WHERE toc_fecrem >= cdate('" & Format(I_DetCom.Fecha(0).text, "dd/mm/yyyy") & "') AND toc_fecrem <= cdate('" & Format(I_DetCom.Fecha(1).text, "dd/mm/yyyy") & "')" & " AND toc_tipdoc <> 'SN'", "WHERE toc_fecrem >= '" & Format(I_DetCom.Fecha(0).text, "yyyymmdd") & "' AND toc_fecrem <= '" & Format(I_DetCom.Fecha(1).text, "yyyymmdd") & "'" & " AND toc_tipdoc <> 'SN'")
'
'    End If
'
'    If I_DetCom.Check1(1).Value = 1 Then
'
'        vg_Consulta = vg_Consulta & IIf(Trim(vg_Consulta) <> "", " AND ", "WHERE ") & "toc_codbod = " & fg_codigocbo(I_DetCom.Combo1, 0, 10, 0) & " AND toc_tipdoc <> 'SN'"
'
'    End If
'
'    If I_DetCom.Check1(2).Value = 1 Then
'
'        vg_Consulta = vg_Consulta & IIf(Trim(vg_Consulta) <> "", " AND ", "WHERE ") & "tip_codigo = " & fg_codigocbo(I_DetCom.Combo1, 1, 10, 0) & " AND toc_tipdoc <> 'SN'"
'
'    End If
'
'    If I_DetCom.Check1(3).Value = 1 Then
'
'        vg_Consulta = vg_Consulta & IIf(Trim(vg_Consulta) <> "", " AND ", "WHERE ") & " dec_codmer = '" & fg_Quitachar(Trim(fg_codigocbo(I_DetCom.Combo1, 2, 20, 0)), "(") & "'" & " AND toc_tipdoc <> 'SN'"
'
'    End If
'
'    If I_DetCom.Check1(4).Value = 1 Then
'
'        vg_Consulta = vg_Consulta & IIf(Trim(vg_Consulta) <> "", " AND ", "WHERE ") & "toc_rutpro = '" & v_rut & "'" & " AND toc_tipdoc <> 'SN'"
'
'    End If
'
'    If I_DetCom.Check1(5).Value = 1 Then
'
'        vg_Consulta = vg_Consulta & IIf(Trim(vg_Consulta) <> "", " AND ", "WHERE ") & "toc_tipdoc = '" & fg_codigocbo(I_DetCom.Combo1, 3, 2, "") & "'"
'
'    End If

    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
'    RS1.Open "SELECT DISTINCT * FROM a_unidad INNER JOIN ((b_proveedor INNER JOIN b_totcompras ON b_proveedor.prv_codigo = b_totcompras.toc_rutpro) INNER JOIN ((a_tipopro INNER JOIN b_productos ON a_tipopro.tip_codigo=b_productos.pro_codtip) INNER JOIN b_detcompras ON b_productos.pro_codigo = b_detcompras.dec_codmer) ON (b_totcompras.toc_numdoc = b_detcompras.dec_numdoc) AND (b_totcompras.toc_tipdoc = b_detcompras.dec_tipdoc) AND (b_totcompras.toc_rutpro = b_detcompras.dec_rutpro)) ON a_unidad.uni_codigo = b_productos.pro_coduni " & vg_Consulta & " AND toc_codbod = " & vg_codbod & " ORDER BY toc_rutpro, toc_fecrem, toc_tipdoc,toc_numdoc", vg_db, adOpenStatic
    Set RS1 = vg_db.Execute("sgp_Sel_DetalleComprasxPeriodo '" & op1 & "', '" & fi & "', '" & ff & "', '" & op2 & "', '" & codbod & "', '" & op3 & "', '" & fampro & "', '" & op4 & "', '" & codpro & "', '" & op5 & "', '" & codprov & "', '" & op6 & "', '" & tipdoc & "'")
    If RS1.EOF Then fg_descarga: MsgBox "No existen datos para la consulta...", vbExclamation + vbOKOnly, MsgTitulo: RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .TableBorder = tbBottom
    .EndTable
    
    .StartTable
    .TableCell(tcCols) = 5: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1500: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True:  .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Unidad"
    .TableCell(tcText, 1, 4) = "Cantidad"
    .TableCell(tcText, 1, 5) = "Total"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|"; .TableCell(tcText, 1, 5)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 5: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1500: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcFontBold, 1, 1, 10000, 6) = False:   .TableCell(tcRowHeight, 1) = 230
    
    i = 1
    TGen = 0
    TGeniva = 0
    TGenexec = 0
    TGenOImp = 0
    TGennet = 0
    
    Do While Not RS1.EOF
        
        If Not RS1.EOF Then
            
            i = i + 1
            .TableCell(tcColSpan, i, 1) = 5: .TableCell(tcFontBold, i, 1) = True:   .TableCell(tcRowHeight, 1) = 230
            .TableCell(tcText, i, 1) = "Proveedor: " & fg_PintaRut(RS1!toc_rutpro)
            Print #1, .TableCell(tcText, i, 1)
            i = i + 1
        
        End If
        
        .TableCell(tcFontBold, i, 2) = False:   .TableCell(tcRowHeight, 1) = 230
        v_provee = RS1!toc_rutpro
        tTot = 0
        
        Do While Not RS1.EOF
            
            If RS1!toc_rutpro = v_provee Then
                
                .TableCell(tcText, i, 1) = RS1!dec_codmer
                .TableCell(tcText, i, 2) = RS1!pro_nombre
                .TableCell(tcText, i, 3) = RS1!uni_nombre
                .TableCell(tcText, i, 4) = Format(RS1!dec_canmer, fg_Pict(8, vg_DCa))
                .TableCell(tcText, i, 5) = IIf(Trim(RS1!toc_tipdoc) = "NC" Or Trim(RS1!toc_tipdoc) = "CE", "(" & Format(RS1!dec_ptotal + RS1!dec_prefle, fg_Pict(8, vg_DPr)) & ")", Format(RS1!dec_ptotal + RS1!dec_prefle, fg_Pict(8, vg_DPr))) 'Format(Round(RS1!dec_ptotal + RS1!dec_prefle), fg_Pict(8, vg_DPr))
                
                Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|"; .TableCell(tcText, i, 5)
                TNet = TNet + IIf(Trim(RS1!toc_tipdoc) = "NC" Or Trim(RS1!toc_tipdoc) = "CE", (RS1!dec_ptotal + IIf(IsNull(RS1!dec_prefle), 0, RS1!dec_prefle)) * -1, RS1!dec_ptotal + IIf(IsNull(RS1!dec_prefle), 0, RS1!dec_prefle)) 'RS1!dec_ptotal + IIf(IsNull(RS1!dec_prefle), 0, RS1!dec_prefle)
                RS1.MoveNext: i = i + 1
            
            Else
                
                Exit Do
            
            End If
        
        Loop
        
        .TableCell(tcFontBold, i, 2) = True:   .TableCell(tcRowHeight, 1) = 230
        .TableCell(tcText, i, 2) = "Total Proveedor"
        .TableCell(tcFontBold, i, 5) = True:   .TableCell(tcRowHeight, 1) = 230
        .TableCell(tcText, i, 5) = Format(TNet, fg_Pict(8, vg_DPr))
        Print #1, vbTab & vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & .TableCell(tcText, i, 5)
        TGen = TGen + TNet
        TNet = 0
    
    Loop
    
    i = i + 2
    .TableCell(tcFontBold, i, 2) = True:   .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, i, 2) = "Total General"
    .TableCell(tcFontBold, i, 5) = True:   .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, i, 5) = Format(TGen, fg_Pict(8, vg_DPr))
    Print #1, vbTab & vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & .TableCell(tcText, i, 5)
    i = i + 1
    RS1.Close
    Set RS1 = Nothing
    Close #1
    
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc

End With
fg_descarga
Preview.Show 1
Exit Sub
Error_DetalleCom:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Sub i_Fampro()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Familia
fg_carga ""
MsgTitulo = "Informe Familias Productos"
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    LogoEmp
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Familia de Productos" & Chr(13)
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open "select * from a_tipopro where tip_previo=0 order by tip_nombre", vg_db, adOpenStatic
    If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Categoría"
    .TableCell(tcText, 1, 2) = "Subcategoria"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!tip_nombre
        .TableCell(tcText, i, 2) = ""
        RS2.Open "SELECT * FROM a_tipopro WHERE tip_previo = " & RS1!tip_codigo & " AND tip_activo in ('1','S') ORDER BY tip_nombre", vg_db, adOpenStatic
        If Not RS2.EOF Then
            Do While Not RS2.EOF
                .TableCell(tcText, i, 2) = RS2!tip_nombre
                .TableCell(tcText, i, 3) = ""
                RS3.Open "SELECT * FROM a_tipopro WHERE tip_previo = " & RS2!tip_codigo & " AND tip_activo in ('1','S') ORDER BY tip_nombre", vg_db, adOpenStatic
                If Not RS3.EOF Then
                   Do While Not RS3.EOF
                      .TableCell(tcText, i, 3) = RS3!tip_nombre
                      Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
                      RS3.MoveNext: i = i + 1
                   Loop
                Else
                   Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
                End If
                RS3.Close: Set RS3 = Nothing
                RS2.MoveNext: i = i + 1
            Loop
        Else
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2)
            i = i + 1
        End If
        RS2.Close: Set RS2 = Nothing
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
fg_descarga
Preview.Show 1
Exit Sub
Error_Familia:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Sub I_UniEnv()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_UniEnv
fg_carga ""
MsgTitulo = "Informe de Unidades de Envase"
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    LogoEmp
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #FreeFile
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Unidades de Envase"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.Unidad(2, 0, ""), vg_db, adOpenStatic
    If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Nombre"
    .TableCell(tcText, 1, 3) = "Nombre Corto"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!uni_codigo
        .TableCell(tcText, i, 2) = RS1!uni_nombre
        .TableCell(tcText, i, 3) = RS1!uni_nomcor
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2)
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
fg_descarga
Preview.Show 1
Exit Sub
Error_UniEnv:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Sub i_uniemb()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_UniEmb
fg_carga ""
MsgTitulo = "Informe de Unidades de Embalaje"
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    LogoEmp
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Unidades de Embalaje"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.Embalaje(1, 0, ""), vg_db, adOpenStatic
    If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Nombre"
    .TableCell(tcText, 1, 3) = "Nombre Corto"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!emb_codigo
        .TableCell(tcText, i, 2) = RS1!emb_nombre
        .TableCell(tcText, i, 3) = RS1!emb_nomcor
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2)
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
fg_descarga
Preview.Show 1
Exit Sub
Error_UniEmb:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Sub I_Nutri()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Nutriente
fg_carga ""
MsgTitulo = "Informe de Nutrientes"
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    LogoEmp
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Nutrientes"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.Nutriente(3, 0, ""), vg_db, adOpenStatic
    If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 2000: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Nombre"
    .TableCell(tcText, 1, 2) = "Nombre Corto"
    .TableCell(tcText, 1, 3) = "Principal"
    .TableCell(tcText, 1, 4) = "Ordenamiento"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 2000: .TableCell(tcAlign, , 4) = taCenterTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!nut_nombre
        .TableCell(tcText, i, 2) = RS1!nut_nomuni
        .TableCell(tcText, i, 3) = IIf(RS1!nut_indpri = 1, "SI", "NO")
        .TableCell(tcText, i, 4) = RS1!nut_secnro
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
fg_descarga
Preview.Show 1
Exit Sub
Error_Nutriente:
    fg_descarga
     MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
     Close #1
     Exit Sub
End Sub

Sub I_impues()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Impto
MsgTitulo = "Informe de Impuestos"
Preview.Show
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orLandscape '= orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .FontSize = 8
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 13500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Impuestos"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.Impuesto(3, 0, ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 11: .TableCell(tcRows) = 6
    .TableCell(tcColWidth, , 1) = 900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1600: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 1200: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taCenterTop
    .TableCell(tcColWidth, , 6) = 1300: .TableCell(tcAlign, , 6) = taCenterTop
    .TableCell(tcColWidth, , 7) = 1100: .TableCell(tcAlign, , 7) = taCenterTop
    .TableCell(tcColWidth, , 8) = 1100: .TableCell(tcAlign, , 8) = taCenterTop
    .TableCell(tcColWidth, , 9) = 1100: .TableCell(tcAlign, , 9) = taCenterTop
    .TableCell(tcColWidth, , 10) = 1100: .TableCell(tcAlign, , 10) = taCenterTop
    .TableCell(tcColWidth, , 11) = 1300: .TableCell(tcAlign, , 11) = taCenterTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcBackColor, 2) = vbYellow: .TableCell(tcFontBold, 2) = True: .TableCell(tcRowHeight, 2) = 230
    .TableCell(tcBackColor, 3) = vbYellow: .TableCell(tcFontBold, 3) = True: .TableCell(tcRowHeight, 3) = 230
    .TableCell(tcBackColor, 4) = vbYellow: .TableCell(tcFontBold, 4) = True: .TableCell(tcRowHeight, 4) = 230
    .TableCell(tcBackColor, 5) = vbYellow: .TableCell(tcFontBold, 5) = True: .TableCell(tcRowHeight, 5) = 230
    .TableCell(tcBackColor, 6) = vbYellow: .TableCell(tcFontBold, 6) = True: .TableCell(tcRowHeight, 6) = 230
    .TableCell(tcText, 6, 1) = "Código"
    .TableCell(tcText, 6, 2) = "Nombre"
    .TableCell(tcText, 5, 3) = "Valor"
    .TableCell(tcText, 6, 3) = "Impuesto"
    .TableCell(tcText, 4, 4) = "Impuesto"
    .TableCell(tcText, 5, 4) = "No"
    .TableCell(tcText, 6, 4) = "Recuperable"
    .TableCell(tcText, 5, 5) = "Modifica"
    .TableCell(tcText, 6, 5) = "Impuesto"
    .TableCell(tcText, 4, 6) = "Cuenta "
    .TableCell(tcText, 5, 6) = "Contable "
    .TableCell(tcText, 6, 6) = "SAP"
    .TableCell(tcText, 1, 7) = IIf(vg_pais = "CO", "Código", "")
    .TableCell(tcText, 2, 7) = IIf(vg_pais = "CO", "Impuesto", "")
    .TableCell(tcText, 3, 7) = IIf(vg_pais = "CO", "SAP", "")
    .TableCell(tcText, 4, 7) = IIf(vg_pais = "CO", "Insumos", "Código")
    .TableCell(tcText, 5, 7) = IIf(vg_pais = "CO", "Casinos", "Impuesto")
    .TableCell(tcText, 6, 7) = IIf(vg_pais = "CO", "Gravadas", "SAP")
    .TableCell(tcText, 1, 8) = IIf(vg_pais = "CO", "Código", "")
    .TableCell(tcText, 2, 8) = IIf(vg_pais = "CO", "Impuesto", "")
    .TableCell(tcText, 3, 8) = IIf(vg_pais = "CO", "SAP", "")
    .TableCell(tcText, 4, 8) = IIf(vg_pais = "CO", "Insumos", "")
    .TableCell(tcText, 5, 8) = IIf(vg_pais = "CO", "Casinos", "")
    .TableCell(tcText, 6, 8) = IIf(vg_pais = "CO", "No Gravada", "")
    .TableCell(tcText, 1, 9) = IIf(vg_pais = "CO", "Código", "")
    .TableCell(tcText, 2, 9) = IIf(vg_pais = "CO", "Impuesto", "")
    .TableCell(tcText, 3, 9) = IIf(vg_pais = "CO", "SAP", "")
    .TableCell(tcText, 4, 9) = IIf(vg_pais = "CO", "Servicios", "")
    .TableCell(tcText, 5, 9) = IIf(vg_pais = "CO", "Casinos", "")
    .TableCell(tcText, 6, 9) = IIf(vg_pais = "CO", "Gravada", "")
    .TableCell(tcText, 1, 10) = IIf(vg_pais = "CO", "Código", "")
    .TableCell(tcText, 2, 10) = IIf(vg_pais = "CO", "Impuesto", "")
    .TableCell(tcText, 3, 10) = IIf(vg_pais = "CO", "SAP", "")
    .TableCell(tcText, 4, 10) = IIf(vg_pais = "CO", "Servicios", "")
    .TableCell(tcText, 5, 10) = IIf(vg_pais = "CO", "Casinos", "")
    .TableCell(tcText, 6, 10) = IIf(vg_pais = "CO", "No Gravada", "")
    .TableCell(tcText, 5, 11) = "Impuesto"
    .TableCell(tcText, 6, 11) = "Adicional"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & .TableCell(tcText, 1, 9) & "|" & .TableCell(tcText, 1, 10) & "|" & .TableCell(tcText, 1, 11)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 11: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1600: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 1200: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taCenterTop
    .TableCell(tcColWidth, , 6) = 1300: .TableCell(tcAlign, , 6) = taCenterTop
    .TableCell(tcColWidth, , 7) = 1100: .TableCell(tcAlign, , 7) = taCenterTop
    .TableCell(tcColWidth, , 8) = 1100: .TableCell(tcAlign, , 8) = taCenterTop
    .TableCell(tcColWidth, , 9) = 1100: .TableCell(tcAlign, , 9) = taCenterTop
    .TableCell(tcColWidth, , 10) = 1100: .TableCell(tcAlign, , 10) = taCenterTop
    .TableCell(tcColWidth, , 11) = 1300: .TableCell(tcAlign, , 11) = taCenterTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = Trim(RS1!imp_codigo)
        .TableCell(tcText, i, 2) = Trim(RS1!imp_nombre)
        .TableCell(tcText, i, 3) = Format(RS1!imp_pctimp, fg_Pict(6, 2))
        .TableCell(tcText, i, 4) = IIf(RS1!imp_inccos = 1, "NO", "SI")
        .TableCell(tcText, i, 5) = IIf(RS1!imp_indmod = "S", "SI", "NO")
        .TableCell(tcText, i, 6) = IIf(IsNull(RS1!imp_codsap), "", Trim(RS1!imp_codsap))
        .TableCell(tcText, i, 7) = IIf(IsNull(RS1!imp_cimsap1), "", Trim(RS1!imp_cimsap1))
        .TableCell(tcText, i, 8) = IIf(IsNull(RS1!imp_cimsap2), "", Trim(RS1!imp_cimsap2))
        .TableCell(tcText, i, 9) = IIf(IsNull(RS1!imp_cimsap3), "", Trim(RS1!imp_cimsap3))
        .TableCell(tcText, i, 10) = IIf(IsNull(RS1!imp_cimsap4), "", Trim(RS1!imp_cimsap4))
        .TableCell(tcText, i, 11) = IIf(RS1!imp_adicional = 0, "NO", "SI")
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9) & "|" & .TableCell(tcText, i, 10) & "|" & .TableCell(tcText, i, 11)
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Sub
Error_Impto:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Sub I_Perfil()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Perfiles
fg_carga ""
MsgTitulo = "Informe de Pefiles"
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    LogoEmp
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Perfiles de Acceso"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open "select * from a_perfil order by per_nombre", vg_db, adOpenStatic
    If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Nombre"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!per_codigo
        .TableCell(tcText, i, 2) = RS1!per_nombre
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2)
         i = i + 1
         RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
fg_descarga
Preview.Show 1
Exit Sub
Error_Perfiles:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Sub I_acceso()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim codigo As Integer
On Local Error GoTo Error_Acceso
fg_carga ""
MsgTitulo = "Informe de Accesos"
Preview.Refresh
M_Perfil.vaSpread1.Row = M_Perfil.vaSpread1.ActiveRow
M_Perfil.vaSpread1.Col = 2
Nombre = M_Perfil.vaSpread1.text
M_Perfil.vaSpread1.Col = 1
codigo = M_Perfil.vaSpread1.text
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
     LogoEmp
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Perfil: " & Nombre
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    RS1.Open "SELECT a_derechosperfil.*, a_opcsistema.opc_nombre FROM a_opcsistema INNER JOIN a_derechosperfil ON a_opcsistema.opc_codigo = a_derechosperfil.dpe_codopc WHERE a_derechosperfil.dpe_codper=" & codigo, vg_db, adOpenStatic
    If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1000: .TableCell(tcAlign, , 2) = taCenterTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taCenterTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taCenterTop
    .TableCell(tcColWidth, , 7) = 2000: .TableCell(tcAlign, , 7) = taCenterTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Nombre"
    .TableCell(tcText, 1, 2) = "Acción"
    .TableCell(tcText, 1, 3) = "Agregar"
    .TableCell(tcText, 1, 4) = "Modificar"
    .TableCell(tcText, 1, 5) = "Eliminar"
    .TableCell(tcText, 1, 6) = "Imprimir"
    .TableCell(tcText, 1, 7) = ""
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & _
               .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7)
    
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1000: .TableCell(tcAlign, , 2) = taCenterTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taCenterTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taCenterTop
    .TableCell(tcColWidth, , 7) = 2000: .TableCell(tcAlign, , 7) = taCenterTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!opc_nombre
        .TableCell(tcText, i, 2) = IIf(RS1!dpe_deracc = 1, "SI", "NO")
        .TableCell(tcText, i, 3) = IIf(RS1!dpe_deragr = 1, "SI", "NO")
        .TableCell(tcText, i, 4) = IIf(RS1!dpe_dermod = 1, "SI", "NO")
        .TableCell(tcText, i, 5) = IIf(RS1!dpe_dereli = 1, "SI", "NO")
        .TableCell(tcText, i, 6) = IIf(RS1!dpe_derimp = 1, "SI", "NO")
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & _
                   .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
fg_descarga
Preview.Show 1
Exit Sub
Error_Acceso:
    fg_descarga
     MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Sub I_Usuari()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim codigo As Integer
On Local Error GoTo Error_Usuarios
fg_carga ""
MsgTitulo = "Informe de Usuarios"
Preview.Refresh
M_Perfil.vaSpread1.Row = M_Perfil.vaSpread1.ActiveRow
M_Perfil.vaSpread1.Col = 2
Nombre = M_Perfil.vaSpread1.text
M_Perfil.vaSpread1.Col = 1
codigo = M_Perfil.vaSpread1.text
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
     LogoEmp
     vg_Archxls = fg_ArchivoTxt
     Open vg_Archxls For Output As #1
    .FontSize = 8
    .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Listado de Usuarios "
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
     .text = Chr(13): .text = Chr(13)
    RS1.Open "SELECT a_usuarios.usu_nombre, a_usuarios.usu_password, a_usuarios.usu_telefono, a_usuarios.usu_email, a_usuarios.usu_oficina, a_usuarios.usu_depart, a_perfil.per_nombre" & _
            " FROM a_perfil INNER JOIN a_usuarios ON a_perfil.per_codigo = a_usuarios.usu_perfil", vg_db, adOpenStatic
    If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 6: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 1500: .TableCell(tcAlign, , 5) = taLeftTop
    .TableCell(tcColWidth, , 6) = 2000: .TableCell(tcAlign, , 6) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Usuario"
    .TableCell(tcText, 1, 2) = "Oficina"
    .TableCell(tcText, 1, 3) = "Departamento"
    .TableCell(tcText, 1, 4) = "Telefono"
    .TableCell(tcText, 1, 5) = "E-mail"
    .TableCell(tcText, 1, 6) = "Perfil"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & _
              .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 6: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 1500: .TableCell(tcAlign, , 5) = taLeftTop
    .TableCell(tcColWidth, , 6) = 2000: .TableCell(tcAlign, , 6) = taLeftTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!usu_Nombre
        .TableCell(tcText, i, 2) = RS1!usu_oficina
        .TableCell(tcText, i, 3) = RS1!usu_depart
        .TableCell(tcText, i, 4) = RS1!usu_telefono
        .TableCell(tcText, i, 5) = RS1!usu_email
        .TableCell(tcText, i, 6) = RS1!per_nombre
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & _
            .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6)
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
fg_descarga
Preview.Show 1
Exit Sub
Error_Usuarios:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Sub I_DocProvee()

Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim TExe As Double, TNet As Double, tIVA As Double, TOtr As Double, tTot As Double, TFle As Double

On Local Error GoTo Error_DocProvee

fg_carga ""
MsgTitulo = "Impresión Documentos Proveedor"
'----------------------------------------------------
'---Nombre de Informe: Informe documentos Proveedor (FA-GD,etc.)
'---Creador: Miguel Solorza P.(Corrección Alexis Morgado)
'---Fecha de Crecación: 20-07-2004.
'----------------------------------------------------
Preview.Refresh
With Preview.VSPrinter
    
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .FontSize = 8
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .text = Chr(13)
    .text = Chr(13)
    
    Select Case fg_TraerRelacionTipoDocumento(vg_TDC)
    
        Case "FA", "FE"
            
            .TableCell(tcText, 1, 1) = "Factura N° " & Trim(Str(vg_NDC))
        
        Case "GD"
            
            .TableCell(tcText, 1, 1) = "Guía de Despacho N° " & Trim(Str(vg_NDC))
        
        Case "NC", "CE"
            
            .TableCell(tcText, 1, 1) = "Nota de Crédito N° " & Trim(Str(vg_NDC))
        
        Case "ND", "DE"
            
            .TableCell(tcText, 1, 1) = "Nota de Debito N° " & Trim(Str(vg_NDC))
        
        Case "BO"
            
            .TableCell(tcText, 1, 1) = "Boleta N° " & Trim(Str(vg_NDC))
        
        Case "BH"
            
            .TableCell(tcText, 1, 1) = "Boleta de Honorarios N° " & Trim(Str(vg_NDC))
    
    End Select
    
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13)
    .text = Chr(13)
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
'    RS1.Open "SELECT a.*, b.*, c.*, d.*, e.uni_nomcor, f.bod_nombre FROM b_totcompras a, b_detcompras b, b_proveedor c, b_productos d, a_unidad e, a_bodega f WHERE a.toc_rutpro = b.dec_rutpro AND a.toc_tipdoc = b.dec_tipdoc AND a.toc_numdoc = b.dec_numdoc AND a.toc_rutpro = c.prv_codigo AND b.dec_codmer = d.pro_codigo AND d.pro_coduni = e.uni_codigo AND a.toc_codbod = f.bod_codigo AND a.toc_codbod = " & vg_codbod & " AND a.toc_rutpro = '" & vg_RDC & "' AND a.toc_tipdoc = '" & vg_TDC & "' AND a.toc_numdoc = " & vg_NDC, vg_db, adOpenStatic
    Set RS1 = vg_db.Execute("sgp_Sel_DocumentoProveedorCaratula " & vg_codbod & ", '" & vg_RDC & "', '" & vg_TDC & "', " & vg_NDC & "")
    If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 4
    .TableCell(tcRows) = 5
    .TableCell(tcColWidth, , 1) = 2000
    .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4000
    .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1500
    .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 3000
    .TableCell(tcAlign, , 4) = taLeftTop
    
    .TableCell(tcText, 1, 1) = "Nombre Proveedor"
    .TableCell(tcText, 1, 2) = IIf(IsNull(RS1!prv_nombre), "", RS1!prv_nombre)
    .TableCell(tcText, 1, 3) = "Rut Proveedor"
    .TableCell(tcText, 1, 4) = fg_PintaRut(RS1!prv_codigo)
    .TableCell(tcText, 2, 1) = "Dirección"
    .TableCell(tcText, 2, 2) = IIf(IsNull(RS1!prv_direccion), "", RS1!prv_direccion)
    .TableCell(tcText, 2, 3) = "Comuna"
    .TableCell(tcText, 2, 4) = IIf(IsNull(RS1!prv_comuna), "", RS1!prv_comuna)
    .TableCell(tcText, 3, 1) = "Fecha Emisión"
    .TableCell(tcText, 3, 2) = Format(RS1!toc_fecemi, "dd/MM/yyyy")
'    .TableCell(tcText, 3, 3) = "Fecha Vencimiento"
    .TableCell(tcText, 3, 3) = "Fecha Recep."
    .TableCell(tcText, 3, 4) = Format(RS1!toc_fecrem, "dd/MM/yyyy")
'    .TableCell(tcText, 3, 4) = Format(RS1!toc_fecven, "dd/MM/yyyy")
    .TableCell(tcText, 4, 1) = "Bodega"
    .TableCell(tcText, 4, 2) = RS1!bod_nombre
    
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    Print #1, .TableCell(tcText, 4, 1) & "|" & .TableCell(tcText, 4, 2)
    
    .TableBorder = tbNone
    .EndTable
    .FontSize = 7
    .StartTable
    .TableCell(tcCols) = 7
    .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1500
    .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4600
    .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 600
    .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 800
    .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1000
    .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000
    .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000
    .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow
    .TableCell(tcFontBold, 1) = True
    .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Unid."
    .TableCell(tcText, 1, 4) = "Cantidad"
    .TableCell(tcText, 1, 5) = "Precio"
    .TableCell(tcText, 1, 6) = "Descto."
    .TableCell(tcText, 1, 7) = "Total"
    
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7)
    
    .TableBorder = tbBox
    .EndTable
    TExe = RS1!toc_exedoc: TNet = RS1!toc_netdoc: TFle = RS1!toc_fledoc: tIVA = RS1!toc_ivadoc: TOtr = RS1!toc_otrimp: tTot = RS1!toc_totdoc
    
    .StartTable
    .TableCell(tcCols) = 7
    .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1500
    .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4600
    .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 600
    .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 800
    .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1000
    .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000
    .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000
    .TableCell(tcAlign, , 7) = taRightTop
    i = 1
    
    Do While Not RS1.EOF
        
        .TableCell(tcText, i, 1) = RS1!dec_codmer
        .TableCell(tcText, i, 2) = RS1!pro_nombre
        .TableCell(tcText, i, 3) = RS1!uni_nomcor
        .TableCell(tcText, i, 4) = Format(IIf(vg_pais <> "CO", RS1!dec_canmer, RS1!dec_cmefac), fg_Pict(8, vg_DCa))
        .TableCell(tcText, i, 5) = Format(IIf(vg_pais <> "CO", RS1!dec_precom, RS1!dec_pmefac), fg_Pict(8, 2)) 'vg_DCa))
        .TableCell(tcText, i, 6) = Format(RS1!dec_valdes, fg_Pict(8, 0)) 'vg_dca
        .TableCell(tcText, i, 7) = Format(RS1!dec_ptotal, fg_Pict(8, vg_DPr))
        
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                  .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)

        RS1.MoveNext
        i = i + 1
    
    Loop
    '.FontSize = 8
    i = i + 1
    .TableCell(tcText, i, 1) = ""
    i = i + 1
    .TableCell(tcFontBold, i) = True: .TableCell(tcText, i, 2) = "Total Exento": .TableCell(tcText, i, 7) = Format(TExe, fg_Pict(8, vg_DPr))
    Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & vbTab & vbTab & .TableCell(tcText, i, 7)
    i = i + 1
    .TableCell(tcFontBold, i) = True: .TableCell(tcText, i, 2) = "Total Neto": .TableCell(tcText, i, 7) = Format(TNet, fg_Pict(8, vg_DPr))
    Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & vbTab & vbTab & .TableCell(tcText, i, 7)
    i = i + 1
    .TableCell(tcFontBold, i) = True: .TableCell(tcText, i, 2) = "Flete": .TableCell(tcText, i, 7) = Format(TFle, fg_Pict(8, vg_DPr))
    Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & vbTab & vbTab & .TableCell(tcText, i, 7)
    i = i + 1
    .TableCell(tcFontBold, i) = True: .TableCell(tcText, i, 2) = "Total IVA": .TableCell(tcText, i, 7) = Format(tIVA, fg_Pict(8, vg_DPr))
    Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & vbTab & vbTab & .TableCell(tcText, i, 7)
    i = i + 1
    .TableCell(tcFontBold, i) = True: .TableCell(tcText, i, 2) = "Total Otr.Imp.": .TableCell(tcText, i, 7) = Format(TOtr, fg_Pict(8, vg_DPr))
    Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & vbTab & vbTab & .TableCell(tcText, i, 7)
    i = i + 1
    .TableCell(tcFontBold, i) = True: .TableCell(tcText, i, 2) = "Total": .TableCell(tcText, i, 7) = Format(tTot, fg_Pict(8, vg_DPr))
    Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & vbTab & vbTab & .TableCell(tcText, i, 7)
    i = i + 1
    RS1.Close
    Set RS1 = Nothing
    
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
'    .CurrentY = .CurrentY - 1250
'    .MarginRight = .MarginRight - 250
'    .Text = String(141, "_")
    '------ Busca solicitud---
    If fg_TraerRelacionTipoDocumento(fg_codigocbo(M_DocPro.Combo2, 0, 2, "")) <> "FA" And fg_TraerRelacionTipoDocumento(fg_codigocbo(M_DocPro.Combo2, 0, 2, "")) <> "FE" Then Close #1: GoTo Fin_Docto
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
'    RS1.Open "SELECT a.*, b.*, c.*, d.*, e.uni_nomcor, f.bod_nombre FROM b_totcompras a, b_detcompras b, b_proveedor c, b_productos d, a_unidad e, a_bodega f WHERE a.toc_rutpro = b.dec_rutpro AND a.toc_tipdoc = b.dec_tipdoc AND a.toc_numdoc = b.dec_numdoc AND a.toc_rutpro = c.prv_codigo AND b.dec_codmer = d.pro_codigo AND d.pro_coduni = e.uni_codigo AND a.toc_codbod = f.bod_codigo AND a.toc_codbod = " & vg_codbod & " AND a.toc_rutpro = '" & vg_RDC & "' AND a.toc_tipdoc = '" & "SN" & "' AND ltrim(a.toc_docaso) = '" & Trim(Str(M_DocPro.Double1(5).Value)) & "'", vg_db, adOpenStatic
    Set RS1 = vg_db.Execute("sgp_Sel_DocumentoProveedorCaratulaSN  " & vg_codbod & ", '" & vg_RDC & "', '" & Trim(Str(M_DocPro.Double1(5).Value)) & "'")
    If RS1.EOF Then
        
        RS1.Close
        Set RS1 = Nothing
        Close #1
        GoTo Fin_Docto 'Cierra Docto.
        
    End If
    
    .NewPage 'Abre nueva pagina
    .Orientation = orLandscape
    .MarginLeft = 500
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||" 'Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    LogoEmp
    .FontSize = 8
    '---MSP 08-09-2004--- Separa los dos documentos por medio de print
    Print #1, " "
    Print #1, " "
    Print #1, " "
    Print #1, " "
    Print #1, " "
    
    .StartTable
    .TableCell(tcCols) = 1
    .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 13500
    .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcFontSize, 1) = 14
    .TableCell(tcFontBold, 1) = True
    .text = Chr(13)
    .text = Chr(13)
    .TableCell(tcText, 1, 1) = "Solicitud Nota de Crédito N° " & Trim(RS1!toc_numdoc)
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13)
    .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 4
    .TableCell(tcRows) = 4
    .TableCell(tcColWidth, , 1) = 2000
    .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4000
    .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000
    .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 2500
    .TableCell(tcAlign, , 4) = taLeftTop
    
    .TableCell(tcText, 1, 1) = "Nombre Proveedor"
    .TableCell(tcText, 1, 2) = RS1!prv_nombre
    .TableCell(tcText, 1, 3) = "Rut Proveedor"
    .TableCell(tcText, 1, 4) = fg_PintaRut(RS1!prv_codigo)
    .TableCell(tcText, 2, 1) = "Dirección"
    .TableCell(tcText, 2, 2) = IIf(IsNull(RS1!prv_direccion), "", RS1!prv_direccion)
    .TableCell(tcText, 2, 3) = "Comuna"
    .TableCell(tcText, 2, 4) = IIf(IsNull(RS1!prv_comuna), "", RS1!prv_comuna)
    .TableCell(tcText, 3, 1) = "Fecha Emisión"
    .TableCell(tcText, 3, 2) = Format(RS1!toc_fecemi, "dd/MM/yyyy")
'    .TableCell(tcText, 3, 3) = "Fecha Vencimiento"
    .TableCell(tcText, 3, 3) = "Fecha Recep."
    .TableCell(tcText, 3, 4) = Format(RS1!toc_fecrem, "dd/MM/yyyy")
'    .TableCell(tcText, 3, 4) = Format(RS1!toc_fecven, "dd/MM/yyyy")
'    .TableCell(tcText, 3, 4) = Format(RS1!toc_fecemi, "dd/MM/yyyy")
    .TableCell(tcText, 4, 1) = "Bodega"
    .TableCell(tcText, 4, 2) = IIf(IsNull(RS1!bod_nombre), "", RS1!bod_nombre)
    
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    Print #1, .TableCell(tcText, 4, 1) & "|" & .TableCell(tcText, 4, 2)
    
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 10
    .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 900
    .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2500
    .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 800
    .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 900
    .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 900
    .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 900
    .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 900
    .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 900
    .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 900
    .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcColWidth, , 10) = 900
    .TableCell(tcAlign, , 10) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow
    .TableCell(tcFontBold, 1) = True
    .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Unidad"
    .TableCell(tcText, 1, 4) = "Cantidad"
    .TableCell(tcText, 1, 5) = "C.Rec."
    .TableCell(tcText, 1, 6) = "Dif. "
    .TableCell(tcText, 1, 7) = "P.Fact."
    .TableCell(tcText, 1, 8) = "P.Rec"
    .TableCell(tcText, 1, 9) = "Dif."
    .TableCell(tcText, 1, 10) = "Total"
    
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7)
    .TableBorder = tbBox
    .EndTable
    TExe = RS1!toc_exedoc: TNet = RS1!toc_netdoc: tIVA = RS1!toc_ivadoc: TOtr = RS1!toc_otrimp: tTot = RS1!toc_totdoc
    
    .StartTable
    .TableCell(tcCols) = 10: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 800: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 900: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 900: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 900: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 900: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 900: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 900: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcColWidth, , 10) = 900: .TableCell(tcAlign, , 10) = taRightTop
    i = 1
    '---Imprime Detalle de Docto. de Solicitud---
    Do While Not RS1.EOF
        
        .TableCell(tcText, i, 1) = RS1!dec_codmer
        .TableCell(tcText, i, 2) = RS1!pro_nombre
        .TableCell(tcText, i, 3) = RS1!uni_nomcor
        .TableCell(tcText, i, 4) = Format(IIf(vg_pais <> "CO", RS1!dec_canmer, RS1!dec_cmefac), fg_Pict(8, vg_DCa))
        .TableCell(tcText, i, 5) = Format(IIf(vg_pais <> "CO", RS1!dec_canrec, RS1!dec_crefac), fg_Pict(8, vg_DCa))
        .TableCell(tcText, i, 6) = Format(IIf(vg_pais <> "CO", (RS1!dec_canmer - RS1!dec_canrec), (RS1!dec_cmefac - RS1!dec_crefac)), fg_Pict(8, vg_DCa))
        .TableCell(tcText, i, 7) = Format(IIf(vg_pais <> "CO", RS1!dec_precom, RS1!dec_pmefac), fg_Pict(8, 2))
        .TableCell(tcText, i, 8) = Format(IIf(vg_pais <> "CO", RS1!dec_prerec, RS1!dec_prefac), fg_Pict(8, 2))
        .TableCell(tcText, i, 9) = Format(IIf(vg_pais <> "CO", (RS1!dec_precom - RS1!dec_prerec), (RS1!dec_pmefac - RS1!dec_prefac)), fg_Pict(8, 2)) 'vg_DCa))
        .TableCell(tcText, i, 10) = Format((RS1!dec_ptotal - RS1!dec_ptotrec), fg_Pict(8, vg_DPr)) 'Format((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_prerec, fg_Pict(8, vg_DPr))
        
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                  .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)
        
        RS1.MoveNext
        i = i + 1
    
    Loop
    
    '--- Fin imprime detalle
    '--- Trailer de Documento
    i = i + 1
    .TableCell(tcFontBold, i) = True: .TableCell(tcText, i, 2) = "Total Neto": .TableCell(tcText, i, 10) = Format(TNet, fg_Pict(8, vg_DPr))
    Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & vbTab & vbTab & .TableCell(tcText, i, 10)
    
    i = i + 1
    RS1.Close
    Set RS1 = Nothing
    Close #1
    
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    '--- Fin Trailer
    
Fin_Docto:
    .EndDoc
End With
fg_descarga
Preview.Show 1

Exit Sub
Error_DocProvee:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub

End Sub

Sub I_TipoServicio()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_UniEmb
fg_carga ""
MsgTitulo = "Informe de Tipo de Servicio"
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    LogoEmp
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Tipo de Servicio"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.TipoServicio(3, 0, ""), vg_db, adOpenStatic
    If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Nombre"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!tis_codigo
        .TableCell(tcText, i, 2) = RS1!tis_nombre
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2)
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
fg_descarga
Preview.Show 1
Exit Sub
Error_UniEmb:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Sub I_Segmento()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_UniEmb
fg_carga ""
MsgTitulo = "Informe de segmento"
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    LogoEmp
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Segmento"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.Segmento(2, 0, ""), vg_db, adOpenStatic
    If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Nombre"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!seg_codigo
        .TableCell(tcText, i, 2) = RS1!seg_nombre
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2)
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
fg_descarga
Preview.Show 1
Exit Sub
Error_UniEmb:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub
