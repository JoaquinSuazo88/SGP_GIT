Attribute VB_Name = "InforEG"
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim ibusca As Long, i As Long
Dim itab As Integer, swvalidar As Integer, itexto As Integer, opboton As Integer
Dim cAccion As String, modo As String, codigo As String, incluir As String, alterar As String, eliminar As String, imprimir As String
Dim vecdatos(11) As String

Sub I_FoFi(CodCas As String, FechaFol As Long, Folio As Long)
Dim RS1 As New ADODB.Recordset 'Recordset para Consulta
Dim fecha As Date, i As Double, tDesc As Double, tAlim As Double, tVarios As Double, cCta As String
Dim TotDesc As Double, TotAlim As Double, TotVarios As Double, TotEgre As Double, TotIva As Double
Dim tIVA As Double, tEgre As Double, numdoc As Double, aAp As String, J As Long
Dim tMovil As Double, TotMovil As Double, crut As String
On Local Error GoTo Error_SalirFoFi
Msgtitulo = "Informes de Rendición de Gastos FOFI"
J = Len(Trim(Str(FechaFol)))
If J = 7 Then
    fecha = CDate("0" + Mid(Trim(Str(FechaFol)), 1, 1) + "/" + Mid(Trim(Str(FechaFol)), 2, 2) + "/" + Mid(Trim(Str(FechaFol)), 4, 4))
Else
   fecha = CDate(Mid(Trim(Str(FechaFol)), 1, 2) + "/" + Mid(Trim(Str(FechaFol)), 3, 2) + "/" + Mid(Trim(Str(FechaFol)), 5, 4))
End If

aAp = Trim(vg_NUsr) & "_tmp_InfFoFi"
'---Cheque la existencia de tablas temporales. Si existen las elimina---
fg_CheckTmp aAp
vg_db.BeginTrans
vg_db.Execute "select a.toc_rutpro,a.toc_tipdoc,a.toc_numdoc,a.toc_fecemi,a.toc_ivadoc,b.dec_codmer,b.dec_ptotal,c.pro_ctacon, d.prv_nombre, a.toc_totdoc  into " & aAp & _
              " from b_totcompras a, b_detcompras b, b_productos c, b_proveedor  d where   a.toc_tipinf = 'F' and a.toc_numinf =" & Folio & _
              " and a.toc_rutpro = b.dec_rutpro and  a.toc_tipdoc = b.dec_tipdoc and a.toc_numdoc = b.dec_numdoc and  a.toc_rutpro = d.prv_codigo and b.dec_codmer = c.pro_codigo " & _
              " group by  a.toc_rutpro ,a.toc_tipdoc, a.toc_numdoc ,c.pro_ctacon,a.toc_fecemi ,b.dec_codmer ,b.dec_ptotal, a.toc_totdoc,a.toc_ivadoc,d.prv_nombre"
vg_db.CommitTrans
vg_db.BeginTrans
vg_db.Execute "alter table " & aAp & " add column Cuenta char(20)"
vg_db.Execute "update " & aAp & " set Cuenta='Alimentos' where pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
vg_db.Execute "update " & aAp & " set Cuenta='Desechables' where pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
vg_db.Execute "update " & aAp & " set Cuenta='Movilizacion' where pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctamovil"), ";", "','") & "')"
vg_db.Execute "update " & aAp & " set Cuenta='Varios' where cuenta = ''or isnull(cuenta) "
vg_db.CommitTrans
RS1.Open "select * from " & aAp & " ", vg_db, adOpenStatic
'---------------------Fin Carga de Datos -----------------
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "No existen datos para consulta...", vbExclamation + vbOKOnly, Msgtitulo:  Exit Sub
Preview.Show
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\FOFI.rtf"
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
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 13500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = "Rendición Fondo Fijo FOFI"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, 1) = 13500: .TableCell(tcAlign, 1) = taLeftTop
    .TableCell(tcColWidth, 2) = 13500: .TableCell(tcAlign, 2) = taLeftTop
    .TableCell(tcColWidth, 3) = 13500: .TableCell(tcAlign, 3) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1) = "Casino               " & " (" & Trim(CodCas) & " " & Trim(nomcas) & ")" & Space(200) & " " & "Folio Nº" & Str(Folio)
    .TableCell(tcText, 2) = "Período              " & " " & UCase(MonthName(Month(fecha))) & "/" & Str(Year(fecha))
    .TableCell(tcText, 3) = "Monto Asignado" & " " & "_______________"
    Print #1, .TableCell(tcText, 1, 1)
    Print #1, .TableCell(tcText, 2, 1)
    Print #1, .TableCell(tcText, 3, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13)
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
    i = 1: tDesc = 0: tAlim = 0: tVarios = 0: cCta = "": numdoc = 0
    Do While Not RS1.EOF
        tAlim = 0: tDesc = 0: tVarios = 0: tMovil = 0
        numdoc = RS1!toc_numdoc: crut = RS1!toc_rutpro
        Do While Not RS1.EOF And RS1!toc_numdoc = numdoc And RS1!toc_rutpro = crut
            cCta = RS1!pro_ctacon
            If Trim(RS1!cuenta) = "Alimentos" Then
                tAlim = tAlim + RS1!dec_ptotal
                TotAlim = TotAlim + tAlim
            End If
            If Trim(RS1!cuenta) = "Desechables" Then
                tDesc = tDesc + RS1!dec_ptotal
                TotDesc = TotDesc + tDesc
            End If
            If Trim(RS1!cuenta) = "Movilizacion" Then
                tMovil = tMovil + RS1!dec_ptotal
                TotMovil = TotMovil + tMovil
            End If

            If Trim(RS1!cuenta) = "Varios" Then
                tVarios = tVarios + RS1!dec_ptotal
                TotVarios = TotVarios + tVarios
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
                .TableCell(tcText, i, 9) = Round(tVarios, vg_DPr)
                .TableCell(tcText, i, 10) = RS1!pro_ctacon
                .TableCell(tcText, i, 11) = 0
            Else
                .TableCell(tcText, i, 5) = RS1!toc_totdoc
                .TableCell(tcText, i, 6) = Round(tAlim, vg_DPr)
                .TableCell(tcText, i, 7) = Round(tDesc, vg_DPr)
                .TableCell(tcText, i, 8) = Round(tMovil, vg_DPr)
                .TableCell(tcText, i, 9) = 0
                .TableCell(tcText, i, 10) = " "
                .TableCell(tcText, i, 11) = Round(RS1!toc_ivadoc, vg_DPr)
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
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcRows) = i - 1
    .TableBorder = tbAll
    .EndTable
    .Text = Chr(13)
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
    .TableCell(tcText, 1, 5) = Round(TotEgre, vg_DPr)
    .TableCell(tcText, 1, 6) = Round(TotAlim, vg_DPr)
    .TableCell(tcText, 1, 7) = Round(TotDesc, vg_DPr)
    .TableCell(tcText, 1, 8) = Round(TotMovil, vg_DPr)
    .TableCell(tcText, 1, 9) = Round(TotVarios, vg_DPr)
    .TableCell(tcText, 1, 11) = Round(TotIva, vg_DPr)
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & _
              .TableCell(tcText, 1, 9) & "|" & .TableCell(tcText, 1, 10) & "|" & .TableCell(tcText, 1, 11)
    .TableBorder = tbBox
    .EndTable
    .EndDoc
End With
fg_descarga
Close #1
Exit Sub
Error_SalirFoFi:
    MsgBox "Error:" & Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Resume Next
End Sub

Sub I_VenDirect(casnom As String, fecini As String, fecter As String, codbod As Long, nombod As String, TipBus As Long)
'----------------------------------------------------
'---Nombre de Informe: Informe de Venta Directa por cliente
'---Creador: Miguel Solorza P.
'---Fecha de Crecación: 31-08-2004.
'----------------------------------------------------
Dim RS1 As New ADODB.Recordset
Dim cVen As String, tVen As Double, tTot As Double, cCas As String
On Local Error GoTo Error_SalirVen
Msgtitulo = "Informe Venta Directa"
Casino = Trim(Mid(casnom, 1, InStr(1, casnom, "|") - 1))
nomcas = Trim(Mid(casnom, InStr(1, casnom, "|") + 1, Len(casnom)))
If TipBus = 1 Then '--- Un cliente
    RS1.Open "select a.tov_codcas,b.dev_codmer,a.tov_codbod,sum(b.dev_canmer) as totalcantidad,sum(b.dev_ptotal) as  ptotal, c.pro_nombre,d.uni_nomcor,e.cli_nombre " & _
             "from b_totventas a, b_detventas b, b_productos c, a_unidad d ,b_clientes e " & "where (a.tov_fecemi >= cdate('" & fecini & "') and  a.tov_fecemi <= cdate('" & fecter & "')) and (a.tov_tipdoc = 'FA' or a.tov_tipdoc = 'GD') " & _
             "and a.tov_numdoc = b.dev_numdoc and a.tov_codbod = " & codbod & " and a.tov_codcas = '" & fg_DespintaRut(I_VenDir.fpText1(1).Text) & "' " & _
             "and a.tov_tipdoc = b.dev_tipdoc and c.pro_coduni = d.uni_codigo and c.pro_codigo=b.dev_codmer  and a.tov_estdoc <> 'A' " & _
             "and a.tov_codcas = e.cli_codigo " & _
             "group by a.tov_codcas,b.dev_codmer,a.tov_codbod,b.dev_codmer,b.dev_ptotal,c.pro_nombre,d.uni_nomcor, e.cli_nombre", vg_db, adOpenStatic
ElseIf TipBus = 2 Then '--- Todos los clientes
    RS1.Open "select a.tov_codcas,b.dev_codmer,a.tov_codbod,sum(b.dev_canmer) as totalcantidad,sum(b.dev_ptotal) as  ptotal, c.pro_nombre,d.uni_nomcor,e.cli_nombre " & _
             "from b_totventas a, b_detventas b, b_productos c, a_unidad d, b_clientes e " & "where (a.tov_fecemi >= cdate('" & fecini & "') and  a.tov_fecemi <= cdate('" & fecter & "')) and (a.tov_tipdoc = 'FA' or a.tov_tipdoc = 'GD') " & _
             "and a.tov_numdoc = b.dev_numdoc and a.tov_codbod = " & codbod & " and a.tov_codcas <> '' " & _
             "and a.tov_tipdoc = b.dev_tipdoc and c.pro_coduni = d.uni_codigo and c.pro_codigo=b.dev_codmer  and a.tov_estdoc <> 'A' " & _
             "and a.tov_codcas = e.cli_codigo " & _
             "group by a.tov_codcas,b.dev_codmer,a.tov_codbod,b.dev_codmer,b.dev_ptotal,c.pro_nombre,d.uni_nomcor,e.cli_nombre", vg_db, adOpenStatic
End If
If RS1.EOF Then MsgBox "No existen datos para la consulta...", vbExclamation + vbOKOnly, Msgtitulo: RS1.Close: Set RS1 = Nothing: Exit Sub
Preview.Show
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Resumen.rtf"
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
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
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
    .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Casino": .TableCell(tcText, 1, 2) = Casino & " " & nomcas
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
Exit Sub
Error_SalirVen:
    MsgBox "Error:" & Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
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
Dim cSql As String
On Local Error GoTo Error_SalirMer
Msgtitulo = "Informe de mermas por Período"
Casino = Trim(Mid(casnom, 1, InStr(1, casnom, "|") - 1))
nomcas = Trim(Mid(casnom, InStr(1, casnom, "|") + 1, Len(casnom)))
cSql = IIf(codmer = 0, " ", " and a.tov_codser = " & codmer & " ")            'si el filto de tipo de merma es todas
RS1.Open "select a.tov_codser,b.dev_codmer,a.tov_codbod,sum(b.dev_canmer) as totalcantidad,sum(b.dev_ptotal) as  ptotal, c.pro_nombre,d.uni_nomcor, e.aju_nombre " & _
         "from b_totventas a, b_detventas b, b_productos c, a_unidad d, a_tipoajuste e " & "where (a.tov_fecemi >= cdate('" & fecini & "') and  a.tov_fecemi <= cdate('" & fecter & "')) " & cSql & " and a.tov_tipdoc = 'ME' " & _
         "and a.tov_numdoc = b.dev_numdoc and a.tov_codbod = " & codbod & " and a.tov_rutcli = '" & Casino & "' " & _
         "and a.tov_tipdoc = b.dev_tipdoc and c.pro_coduni = d.uni_codigo and (a.tov_codser = e.aju_codigo and e.aju_tipaju = 0) and c.pro_codigo=b.dev_codmer " & _
         "and a.tov_estdoc <> 'A' and a.tov_rutcli = b.dev_rutcli " & _
         "group by a.tov_codser,b.dev_codmer,a.tov_codbod,b.dev_codmer,b.dev_ptotal,c.pro_nombre,d.uni_nomcor,e.aju_nombre", vg_db, adOpenStatic
If RS1.EOF Then MsgBox "No existen datos para la consulta...", vbExclamation + vbOKOnly, Msgtitulo: RS1.Close: Set RS1 = Nothing: Exit Sub
Preview.Show
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Resumen.rtf"
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
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
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
    .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Casino": .TableCell(tcText, 1, 2) = Casino & " " & nomcas
    .TableCell(tcText, 2, 1) = "Tipo Merma": .TableCell(tcText, 2, 2) = Trim(TipMer)
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
Exit Sub
Error_SalirMer:
    MsgBox "Error:" & Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Close #1
    Exit Sub
End Sub

Sub I_SalidasDevolBod(casnom As String, regimen As String, codser As Long, fecini As String, fecter As String)
Dim aAp As String, aAD As String
Dim StrImp As String, StrImpb As String, TipDoc As String
Dim cCta As String, cSer As String, tCta As Double, tSer As Double, tTot As Double
'----------------------------------------------------
'---Nombre de Informe: Informe de Salidas / Devoluciones a Producción
'---Creador: Miguel Solorza P.(Corrección Alexis Morgado)
'---Fecha de Crecación: 30-08-2004.
'----------------------------------------------------
Msgtitulo = "Informes de Producción"

Casino = Trim(Mid(casnom, 1, InStr(1, casnom, "|") - 1))
nomcas = Trim(Mid(casnom, InStr(1, casnom, "|") + 1, Len(casnom)))
codreg = Val(Mid(regimen, 1, InStr(1, regimen, "|") - 1))
nomreg = Trim(Mid(regimen, InStr(1, regimen, "|") + 1, Len(regimen)))
On Local Error GoTo Error_Salir
'---------------------Carga de Datos ------------------------
aAp = Trim(vg_NUsr) & "_tmp_InfSalidas"
aAD = Trim(vg_NUsr) & "_tmp_InfDevol"
'---Cheque la existencia de tablas temporales. Si existen las elimina---
fg_CheckTmp aAp
fg_CheckTmp aAD
vg_db.BeginTrans
vg_db.Execute "select d.dev_tipdoc, a.tov_codreg, a.tov_codser, d.dev_codmer, b.pro_nombre, b.pro_ctacon, c.uni_nomcor, sum(d.dev_canmer) as totalcantidad, SUM(d.dev_ptotal) as ptotal into " & aAp & " " & _
              "from b_totventas a, b_detventas d, b_productos b, a_unidad c Where a.tov_rutcli='" & Casino & "' and (a.tov_codser=" & codser & " or " & codser & "=0) and a.tov_tipdoc='" & IIf(I_SalBod.Combo1(2).ListIndex = 2, "DP", "SP") & "' " & _
              "and a.tov_estdoc<>'A' and d.dev_canmer<>0 and (a.tov_fecpro>=cdate('" & fecini & "') and a.tov_fecpro<=cdate('" & fecter & "')) " & " and tov_codreg = " & codreg & _
              "and d.dev_rutcli=a.tov_rutcli and d.dev_tipdoc=a.tov_tipdoc and d.dev_numdoc=a.tov_numdoc and d.dev_codmer=b.pro_codigo and b.pro_coduni=c.uni_codigo " & _
              "GROUP BY d.dev_tipdoc, a.tov_codreg, a.tov_codser, d.dev_codmer, b.pro_nombre, b.pro_ctacon, c.uni_nomcor"
vg_db.Execute "select d.dev_tipdoc, a.tov_codreg, a.tov_codser, d.dev_codmer, b.pro_nombre, b.pro_ctacon, c.uni_nomcor, sum(d.dev_canmer) as totalcantidad, SUM(d.dev_ptotal) as ptotal into " & aAD & " " & _
              "from b_totventas a, b_detventas d, b_productos b, a_unidad c Where a.tov_rutcli='" & Casino & "' and (a.tov_codser=" & codser & " or " & codser & "=0) and a.tov_tipdoc='DP' " & _
              "and a.tov_estdoc<>'A' and d.dev_canmer<>0 and (a.tov_fecpro>=cdate('" & fecini & "') and a.tov_fecpro<=cdate('" & fecter & "')) " & " and tov_codreg = " & codreg & _
              "and d.dev_rutcli=a.tov_rutcli and d.dev_tipdoc=a.tov_tipdoc and d.dev_numdoc=a.tov_numdoc and d.dev_codmer=b.pro_codigo and b.pro_coduni=c.uni_codigo " & _
              "GROUP BY d.dev_tipdoc, a.tov_codreg, a.tov_codser, d.dev_codmer, b.pro_nombre, b.pro_ctacon, c.uni_nomcor"
vg_db.CommitTrans
vg_db.BeginTrans
vg_db.Execute "alter table " & aAp & " add column Cuenta char(20)"
vg_db.Execute "update " & aAp & " set Cuenta='Alimentos' where pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
vg_db.Execute "update " & aAp & " set Cuenta='Desechables' where pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
vg_db.Execute "update " & aAp & " set Cuenta='Otros' where pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctagastos"), ";", "','") & "')"
If I_SalBod.Combo1(2).ListIndex = 3 Then
    'vg_db.Execute "update  adm_tmp_InfSalidas inner join adm_tmp_InfDevol b on (adm_tmp_InfSalidas.dev_codmer=b.dev_codmer) AND (adm_tmp_InfSalidas.tov_codser=b.tov_codser) and (adm_tmp_InfSalidas.tov_codreg=b.tov_codreg) set adm_tmp_InfSalidas.totalcantidad=adm_tmp_InfSalidas.totalcantidad-b.totalcantidad, adm_tmp_InfSalidas.ptotal=adm_tmp_InfSalidas.ptotal-b.ptotal"
    vg_db.Execute "update " & aAp & " inner join " & aAD & " b on (" & aAp & ".dev_codmer=b.dev_codmer) AND (" & aAp & ".tov_codser=b.tov_codser) and (" & aAp & ".tov_codreg=b.tov_codreg) set " & aAp & ".totalcantidad= " & aAp & ".totalcantidad-b.totalcantidad," & aAp & ".ptotal=" & aAp & ".ptotal-b.ptotal"
End If
vg_db.CommitTrans
'---------------------Fin Carga de Datos -----------------
RS1.Open "select a.*, b.* from " & aAp & " a, a_servicio b where a.tov_codser=b.ser_codigo order by b.ser_orden, a.cuenta, a.dev_codmer", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "No existen datos para consulta...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
Casino = Trim(Mid(casnom, 1, InStr(1, casnom, "|") - 1))
nomcas = Trim(Mid(casnom, InStr(1, casnom, "|") + 1, Len(casnom)))
codreg = Val(Mid(regimen, 1, InStr(1, regimen, "|") - 1))
nomreg = Trim(Mid(regimen, InStr(1, regimen, "|") + 1, Len(regimen)))
Preview.Show
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Resumen.rtf"
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
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    If I_SalBod.Combo1(2).ListIndex = 1 Then
        .TableCell(tcText, 1, 1) = "Resumen de Salidas para Producción"
    ElseIf I_SalBod.Combo1(2).ListIndex = 2 Then
        .TableCell(tcText, 1, 1) = "Resumen de Devoluciónes de Producción"
    ElseIf I_SalBod.Combo1(2).ListIndex = 3 Then
        .TableCell(tcText, 1, 1) = "Resumen de Salidas Menos Devoluciones de Producción"
    End If
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Casino": .TableCell(tcText, 1, 2) = Casino & " " & nomcas
    .TableCell(tcText, 2, 1) = "Regimen": .TableCell(tcText, 2, 2) = codreg & " " & nomreg
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
    i = 1: cSer = "": cCta = "": tSer = 0: tCta = 0: tTot = 0
    Do While Not RS1.EOF
        If RS1!ser_nombre <> cSer Or RS1!cuenta <> cCta Then
            If cCta <> "" Then
                .TableCell(tcFontBold, i) = True
                .TableCell(tcText, i, 2) = "Total " & cCta
                .TableCell(tcText, i, 5) = Format(tCta, fg_Pict(6, vg_DPr))
                Print #1, "|" & .TableCell(tcText, i, 2) & "|||" & .TableCell(tcText, i, 5)
                tCta = 0: i = i + 1
            End If
            If RS1!ser_nombre <> cSer Then
                If cSer <> "" Then
                    .TableCell(tcFontBold, i) = True
                    .TableCell(tcText, i, 2) = "Total " & cSer
                    .TableCell(tcText, i, 5) = Format(tSer, fg_Pict(6, vg_DPr))
                    Print #1, "|" & .TableCell(tcText, i, 2) & "|||" & .TableCell(tcText, i, 5)
                     tSer = 0: i = i + 1
                End If
                If cSer <> "" Then i = i + 1
                .TableCell(tcFontBold, i) = True: .TableCell(tcColSpan, i, 1) = 5
                .TableCell(tcText, i, 1) = RS1!ser_nombre
                Print #1, .TableCell(tcText, i, 1)
                cSer = RS1!ser_nombre: i = i + 1
                
            End If
            If cCta <> "" Then i = i + 1
            .TableCell(tcFontBold, i) = True: .TableCell(tcColSpan, i, 1) = 5
            .TableCell(tcText, i, 1) = RS1!cuenta
            Print #1, .TableCell(tcText, i, 1)
             cCta = RS1!cuenta: i = i + 1
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
        .TableCell(tcText, i, 2) = "Total " & cSer
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
Exit Sub
Error_Salir:
    MsgBox "Error:" & Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Close #1
    Exit Sub
End Sub


Sub I_Pendocpro()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long, auxnumdoc As Long, pctdes As Double
Dim auxrutpro As String, auxnompro As String, auxtipdoc As String, nomcta As String, nomcli As String
Dim canali As Double, candes As Double, cangrl As Double, canotros As Double, totali As Double, totdes As Double, tototros As Double, totgrl As Double, pctimp As Double
Dim tanali As Double, tandes As Double, tangrl As Double
Dim v_tipdoc As String, J As Long
On Local Error GoTo Error_Pendientes
Msgtitulo = "Informe Documentos Pendientes"
'----------------------------------------------------
'---Nombre de Informe: Informe Documentos pendientes Proveedor
'---Creador: Miguel Solorza P.
'---Fecha de Crecación: 23-08-2004.
'----------------------------------------------------
Preview.Show
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "Página : %d"
    .FontSize = 8
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
    canali = 0: candes = 0: cangrl = 0: canotros = 0: tototros = 0: tanali = 0: tandes = 0: tangrl = 0: pctdes = 0
    auxtipdoc = "": auxnomdoc = "": auxnompro = ""
    .StartTable
    .TableCell(tcCols) = 8: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taCenterTop
    .TableCell(tcColWidth, , 2) = 1000: .TableCell(tcAlign, , 2) = taCenterTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 3000: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taCenterTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taCenterTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taCenterTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taCenterTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Fecha Emisión"
    .TableCell(tcText, 1, 2) = "ND"
    .TableCell(tcText, 1, 3) = "R.U.T"
    .TableCell(tcText, 1, 4) = "Nombre"
    .TableCell(tcText, 1, 5) = "Alim"
    .TableCell(tcText, 1, 6) = "Desech"
    .TableCell(tcText, 1, 7) = "Otros"
    .TableCell(tcText, 1, 8) = "Monto"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
               .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8)
    .TableBorder = tbBoxRows
    .EndTable
    .StartTable
    .TableCell(tcCols) = 8: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1000: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 3000: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    RS1.Open "select b_totcompras.toc_rutpro, b_totcompras.toc_tipdoc, b_totcompras.toc_numdoc,b_totcompras.toc_fecemi, b_detcompras.dec_numlin, " & _
          "b_detcompras.dec_codmer, b_detcompras.dec_canmer, b_detcompras.dec_precom, b_detcompras.dec_pctdes,b_productos.pro_ctacon, " & _
          "b_detcompras.dec_canrec, b_detcompras.dec_prerec,b_proveedor.prv_nombre,b_proveedor.prv_codigo " & _
          "from  b_totcompras, b_detcompras, b_productos, b_proveedor " & _
          "where (toc_fecemi >= cdate('" & I_DocPen.fecha(0).Text & "')" & " and toc_fecemi <= cdate('" & I_DocPen.fecha(1).Text & "')) and b_totcompras.toc_rutpro=b_detcompras.dec_rutpro " & _
          "and   b_totcompras.toc_tipdoc= '" & v_tipdoc & "'" & _
          "and   b_totcompras.toc_tipdoc=b_detcompras.dec_tipdoc " & _
          "and   b_totcompras.toc_numdoc=b_detcompras.dec_numdoc " & _
          "and   b_detcompras.dec_codmer=b_productos.pro_codigo " & _
          "and   b_detcompras.dec_rutpro=b_proveedor.prv_codigo " & _
          "and   (b_totcompras.toc_docaso='' or b_totcompras.toc_docaso is null)  " & _
          "order by b_totcompras.toc_rutpro, b_totcompras.toc_tipdoc, b_totcompras.toc_numdoc,b_totcompras.toc_fecemi, b_detcompras.dec_numlin, " & _
          "b_detcompras.dec_codmer, b_detcompras.dec_canmer, b_detcompras.dec_precom, b_detcompras.dec_pctdes,b_productos.pro_ctacon, " & _
          "b_detcompras.dec_canrec, b_detcompras.dec_prerec,b_proveedor.prv_nombre,b_proveedor.prv_codigo ", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: .TableBorder = tbAll: .EndDoc: Close #1: Exit Sub
    If Not RS1.EOF Then
       i = 1
       auxtipdoc = "": auxnomdoc = "": auxnompro = "": totali = 0: totdes = 0: totgrl = 0: tototros = 0
       Do While Not RS1.EOF
            auxrutpro = RS1!toc_rutpro: auxnumdoc = RS1!toc_numdoc
            canali = 0: candes = 0: canotros = 0
            Do While Not RS1.EOF And RS1!toc_rutpro = Trim(auxrutpro) And RS1!toc_numdoc = auxnumdoc
                  .TableCell(tcText, i, 1) = RS1!toc_fecemi
                  .TableCell(tcText, i, 2) = RS1!toc_numdoc
                  .TableCell(tcText, i, 3) = fg_PintaRut(RS1!prv_codigo)
                  .TableCell(tcText, i, 4) = RS1!prv_nombre
                  pctimp = 0
                  RS2.Open "select b_detcomprasimp.imd_monimp, a_impuesto.imp_pctimp, a_impuesto.imp_inccos " & _
                           "from  b_detcomprasimp, a_impuesto " & _
                           "where b_detcomprasimp.imd_rutdoc='" & RS1!toc_rutpro & "' " & _
                           "and   b_detcomprasimp.imd_tipdoc='" & RS1!toc_tipdoc & "' " & _
                           "and   b_detcomprasimp.imd_numdoc=" & RS1!toc_numdoc & " " & _
                           "and   b_detcomprasimp.imd_numlin=" & RS1!dec_numlin & " " & _
                           "and   b_detcomprasimp.imd_codpro='" & RS1!dec_codmer & "' " & _
                           "and   b_detcomprasimp.imd_codimp=a_impuesto.imp_codigo " & _
                           "and   a_impuesto.imp_inccos=1", vg_db, adOpenStatic
                  If RS2.EOF Then RS2.Close: Set RS2 = Nothing Else pctimp = RS2!imd_monimp: RS2.Close: Set RS2 = Nothing
                  pctdes = 0
                  If RS1!dec_pctdes > 0 Then pctdes = RS1!dec_pctdes
                  If v_tipdoc = "SN" Then
                      If RS1!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
                         canali = canali + (((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_prerec) - (((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_prerec) * (pctdes / 100)) + pctimp)
                         totali = totali + (((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_precom) - (((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_precom) * (pctdes / 100)) + pctimp)
                      ElseIf RS1!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
                         candes = candes + (((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_prerec) - (((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_prerec) * (pctdes / 100)) + pctimp)
                         totdes = totdes + (((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_precom) - (((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_precom) * (pctdes / 100)) + pctimp)
                      Else
                         canotros = canotros + (((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_prerec) - (((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_prerec) * (pctdes / 100)) + pctimp)
                         tototros = tototros + (((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_precom) - (((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_precom) * (pctdes / 100)) + pctimp)
                      End If
                  ElseIf v_tipdoc = "GD" Then
                      If RS1!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
                         canali = canali + ((RS1!dec_canrec * RS1!dec_precom) - (((RS1!dec_canmer) * RS1!dec_precom) * (pctdes / 100)))
                         totali = totali + ((RS1!dec_canmer * RS1!dec_precom) - (((RS1!dec_canmer) * RS1!dec_precom) * (pctdes / 100)))
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
                 .TableCell(tcText, i, 5) = Format(canali, fg_Pict(6, 0))
                 .TableCell(tcText, i, 6) = Format(candes, fg_Pict(6, 0))
                 .TableCell(tcText, i, 7) = Format(canotros, fg_Pict(6, 0))
                 .TableCell(tcText, i, 8) = Format(canali + candes + canotros, fg_Pict(6, 0))
                 If RS1.EOF Then Exit Do
            Loop
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                      .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8)
            i = i + 1
       Loop
    Else
        i = i + 1
    End If
    .TableCell(tcFontBold, i) = True: .TableCell(tcRowHeight, i) = 230
    .TableCell(tcText, i, 1) = ""
    .TableCell(tcText, i, 2) = "Total"
    .TableCell(tcText, i, 3) = ""
    .TableCell(tcText, i, 4) = ""
    .TableCell(tcText, i, 5) = IIf(totali <> 0, Format(totali, fg_Pict(6, 0)), Format(0, fg_Pict(6, 0)))
    .TableCell(tcText, i, 6) = IIf(totdes <> 0, Format(totdes, fg_Pict(6, 0)), Format(0, fg_Pict(6, 0)))
    .TableCell(tcText, i, 7) = IIf(tototros <> 0, Format(tototros, fg_Pict(6, 0)), Format(0, fg_Pict(6, 0)))
    .TableCell(tcText, i, 8) = IIf(totali <> 0 Or totdes <> 0 Or tototros <> 0, Format((totali + totdes + tototros), fg_Pict(6, 0)), Format(0, fg_Pict(6, 0)))
    Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
               .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8)
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Sub
Error_Pendientes:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Resume Next
    Close #1
    Exit Sub
End Sub

Sub I_ComprasPer()
Dim TExe As Double, TNet As Double, tIVA As Double, TOtr As Double
Dim tTot As Double, v_rut As String
Dim TGen As Double, TGeniva As Double, TGenexec As Double, TGenOImp As Double, TGennet As Double, P As Double
Dim RS1 As New ADODB.Recordset
'On Local Error GoTo Error_Compras
Msgtitulo = "Informe de Compras por Período"
'----------------------------------------------------
'---Nombre de Informe: Informe de Compras por Período
'---Creador: Miguel Solorza P.
'---Fecha de Crecación: 09-08-2004.
'----------------------------------------------------
Preview.Show
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\ComprasPer.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
    LogoEmp
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .FontSize = 7
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .Text = Chr(13): .Text = Chr(13)
    .TableCell(tcText, 1, 1) = "Compras por Período "
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 4
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcFontSize, 1) = 8: .TableCell(tcFontBold, , 1) = True
    i = 1
    If I_ComPer.Check1(0).Value = 1 Then
        i = i + 1
        .TableCell(tcText, i, 1) = "Compras entre " & I_ComPer.fecha(0).Text & " y el " & I_ComPer.fecha(1).Text
        Print #1, .TableCell(tcText, i, 1)
    End If
    If I_ComPer.Check1(1).Value = 1 Then
        i = i + 1
        .TableCell(tcText, i, 1) = "Bodega : " & Trim(Mid(I_ComPer.Combo1(0).Text, 1, 50))
        Print #1, .TableCell(tcText, i, 1)
    End If
    If I_ComPer.Check1(3).Value = 1 Then
        i = i + 1
        .TableCell(tcText, i, 1) = "Tipo de Documento : (" & Trim(fg_codigocbo(I_ComPer.Combo1, 1, 2, "")) & ")" & " - " & Trim(Mid(I_ComPer.Combo1(1).Text, 1, 50))
        Print #1, .TableCell(tcText, i, 1)
    End If
    .TableCell(tcRows) = i
    .TableBorder = tbBottom
    .EndTable
    
    vg_Consulta = ""
    v_rut = fg_DespintaRut(I_ComPer.fpText1(0).Text)
    If I_ComPer.Check1(0).Value = 1 Then
        vg_Consulta = "where toc_fecemi >= cdate('" & Format(I_ComPer.fecha(0).Text, "dd/mm/yyyy") & "') and toc_fecemi <= cdate('" & Format(I_ComPer.fecha(1).Text, "dd/mm/yyyy") & "')" & " and toc_tipdoc <> 'SN'"
    End If
    If I_ComPer.Check1(1).Value = 1 Then
        vg_Consulta = vg_Consulta & IIf(Trim(vg_Consulta) <> "", " and ", "where ") & "toc_codbod = " & fg_codigocbo(I_ComPer.Combo1, 0, 10, 0) & " and toc_tipdoc <> 'SN'"
    End If
    If I_ComPer.Check1(2).Value = 1 Then
        vg_Consulta = vg_Consulta & IIf(Trim(vg_Consulta) <> "", " and ", "where ") & "toc_rutpro = '" & v_rut & "'" & " and toc_tipdoc <> 'SN'"
    End If
    If I_ComPer.Check1(3).Value = 1 Then
        vg_Consulta = vg_Consulta & IIf(Trim(vg_Consulta) <> "", " and ", "where ") & "toc_tipdoc = '" & Trim(fg_codigocbo(I_ComPer.Combo1, 1, 2, "")) & "'"
    End If
    RS1.Open "select distinct a.*, b.* from b_totcompras a, b_proveedor b " & vg_Consulta & " and  a.toc_rutpro =b.prv_codigo" & " order by toc_rutpro, toc_fecemi, toc_tipdoc,toc_numdoc", vg_db, adOpenStatic
    If RS1.EOF Then MsgBox "No existen datos para la consulta...", vbExclamation + vbOKOnly, Msgtitulo: Close #1: RS1.Close: Set RS1 = Nothing: Exit Sub
    .StartTable
    .TableCell(tcCols) = 9: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 400: .TableCell(tcAlign, , 1) = taCenterTop
    .TableCell(tcColWidth, , 2) = 800: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 3000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 900: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1200: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1200: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1200: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True:  .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "TD"
    .TableCell(tcText, 1, 2) = "Doc. Nº"
    .TableCell(tcText, 1, 3) = "Proveedor"
    .TableCell(tcText, 1, 4) = "F.Emisión"
    .TableCell(tcText, 1, 5) = "Exento"
    .TableCell(tcText, 1, 6) = "Neto"
    .TableCell(tcText, 1, 7) = "I.V.A"
    .TableCell(tcText, 1, 8) = "O.Imp."
    .TableCell(tcText, 1, 9) = "Total"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & .TableCell(tcText, 1, 9)
    .TableBorder = tbBox
    .EndTable
    
    .StartTable
    .TableCell(tcCols) = 9: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 400: .TableCell(tcAlign, , 1) = taCenterTop
    .TableCell(tcColWidth, , 2) = 800: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 3000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 900: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1200: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1200: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1200: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcRowHeight, 1) = 230
    i = 1: TGen = 0: TGeniva = 0: TGenexec = 0: TGenOImp = 0: TGennet = 0
    Do While Not RS1.EOF
        'i = i + 1
        '.TableCell(tcColSpan, i, 1) = 8: .TableCell(tcFontBold, i) = True: .TableCell(tcAlign, i, 1) = taLeftTop
        '.TableCell(tcText, i, 1) = "Proveedor: " & fg_PintaRut(RS1!toc_rutpro)
        'Print #1, .TableCell(tcText, i, 1)
        'i = i + 1
        v_provee = RS1!toc_rutpro
        TExe = 0: TNet = 0: tIVA = 0: TOtr = 0: tTot = 0: P = 0
        Do While Not RS1.EOF
            If RS1!toc_rutpro = v_provee Then
                .TableCell(tcText, i, 1) = RS1!toc_tipdoc
                .TableCell(tcText, i, 2) = RS1!toc_numdoc
                If P = 0 Then
                    .TableCell(tcText, i, 3) = "(" & fg_PintaRut(RS1!toc_rutpro) & ") " & Mid$(RS1!prv_nombre, 1, 20)
                    P = 1
                Else
                    .TableCell(tcText, i, 3) = " "
                    
                End If
                .TableCell(tcText, i, 4) = Format(RS1!toc_fecemi, "dd/mm/yyyy")
                .TableCell(tcText, i, 5) = Format(RS1!toc_exedoc, fg_Pict(8, vg_DPr))
                .TableCell(tcText, i, 6) = Format(RS1!toc_netdoc, fg_Pict(8, vg_DPr))
                .TableCell(tcText, i, 7) = Format(RS1!toc_ivadoc, fg_Pict(8, vg_DPr))
                .TableCell(tcText, i, 8) = Format(RS1!toc_otrimp, fg_Pict(8, vg_DPr))
                .TableCell(tcText, i, 9) = Format(RS1!toc_totdoc, fg_Pict(8, vg_DPr))
                Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                          .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9)
                TExe = TExe + RS1!toc_exedoc: TNet = TNet + RS1!toc_netdoc: tIVA = tIVA + RS1!toc_ivadoc
                TOtr = TOtr + RS1!toc_otrimp: tTot = tTot + RS1!toc_totdoc:
                RS1.MoveNext: i = i + 1
            Else
                Exit Do
            End If
        Loop
        If tTot <> 0 Then
            .TableCell(tcFontBold, i) = True
            .TableCell(tcText, i, 3) = "Total Proveedor"
            .TableCell(tcText, i, 5) = Format(TExe, fg_Pict(8, vg_DPr))
            .TableCell(tcText, i, 6) = Format(TNet, fg_Pict(8, vg_DPr))
            .TableCell(tcText, i, 7) = Format(tIVA, fg_Pict(8, vg_DPr))
            .TableCell(tcText, i, 8) = Format(TOtr, fg_Pict(8, vg_DPr))
            .TableCell(tcText, i, 9) = Format(tTot, fg_Pict(8, vg_DPr))
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                      .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9)
            TGen = TGen + tTot: TGeniva = TGeniva + tIVA
            TGenexec = TGenexec + TExe: TGenOImp = TGenOImp + TOtr: TGennet = TGennet + TNet
            i = i + 2: P = 0
            Print #1, " "
'            .TableCell(tcFontBold, i) = True
'            .TableCell(tcText, i, 3) = "Total General"
'            .TableCell(tcText, i, 5) = Format(TGenexec, fg_Pict(8, vg_DPr))
'            .TableCell(tcText, i, 6) = Format(TGennet, fg_Pict(8, vg_DPr))
'            .TableCell(tcText, i, 7) = Format(TGeniva, fg_Pict(8, vg_DPr))
'            .TableCell(tcText, i, 8) = Format(TGenOImp, fg_Pict(8, vg_DPr))
'            .TableCell(tcText, i, 9) = Format(TGen, fg_Pict(8, vg_DPr))
'            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
'                      .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9)
'            i = i + 1
        End If
     Loop
     i = i + 1
    .TableCell(tcFontBold, i) = True
    .TableCell(tcText, i, 3) = "Total General"
    .TableCell(tcText, i, 5) = Format(TGenexec, fg_Pict(8, vg_DPr))
    .TableCell(tcText, i, 6) = Format(TGennet, fg_Pict(8, vg_DPr))
    .TableCell(tcText, i, 7) = Format(TGeniva, fg_Pict(8, vg_DPr))
    .TableCell(tcText, i, 8) = Format(TGenOImp, fg_Pict(8, vg_DPr))
    .TableCell(tcText, i, 9) = Format(TGen, fg_Pict(8, vg_DPr))
    Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
              .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9)
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Sub
Error_Compras:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Close #1
    Exit Sub
End Sub

Sub I_DetalleCom()
Dim TExe As Double, TNet As Double, tIVA As Double, TOtr As Double
Dim tTot As Double, v_rut As String, v_Switch As Double
Dim TGen As Double, TGeniva As Double, TGenexec As Double, TGenOImp As Double, TGennet As Double
Dim RS1 As New ADODB.Recordset
On Local Error GoTo Error_DetalleCom
Msgtitulo = "Informe de Detalle Compras por Período"
'----------------------------------------------------
'---Nombre de Informe: Informe detalle Compras Proveedor
'---Creador: Miguel Solorza P.(Corrección Alexis Morgado)
'---Fecha de Crecación: 12-08-2004.
'----------------------------------------------------
Preview.Show
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\DetalleCom.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
    LogoEmp
    .FontSize = 7
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .Text = Chr(13): .Text = Chr(13)
    .TableCell(tcText, 1, 1) = "Detalle de Compras por Periodo"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 5
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftMiddle
    i = 1
    If I_DetCom.Check1(0).Value = 1 Then
        .TableCell(tcRows) = i: .TableCell(tcFontSize, i) = 8: .TableCell(tcFontBold, , 1) = True
        .TableCell(tcText, i, 1) = "Compras entre " & I_DetCom.fecha(0).Text & "   y el  " & I_DetCom.fecha(1).Text
        Print #1, .TableCell(tcText, i, 1)
        i = i + 1
    End If
    If I_DetCom.Check1(1).Value = 1 Then
        .TableCell(tcRows) = i: .TableCell(tcFontSize, i) = 8: .TableCell(tcFontBold, i, 1) = True
        .TableCell(tcText, i, 1) = "Bodega: " & Trim(Mid(I_DetCom.Combo1(0).Text, 1, 50))
        Print #1, .TableCell(tcText, i, 1)
        i = i + 1
    End If
    If I_DetCom.Check1(2).Value = 1 Then
         .TableCell(tcRows) = i: .TableCell(tcFontSize, i) = 8: .TableCell(tcFontBold, i, 1) = True
        .TableCell(tcText, i, 1) = "Familia de Producto: " & Trim(Mid(I_DetCom.Combo1(1).Text, 1, 50))
        Print #1, .TableCell(tcText, i, 1)
        i = i + 1
    End If
    If I_DetCom.Check1(3).Value = 1 Then
        .TableCell(tcRows) = i: .TableCell(tcFontSize, i) = 8: .TableCell(tcFontBold, i, 1) = True
        .TableCell(tcText, i, 1) = "Producto: " & Trim(Mid(I_DetCom.Combo1(2).Text, 1, 50))
        Print #1, .TableCell(tcText, i, 1)
        i = i + 1
    End If
    If I_DetCom.Check1(5).Value = 1 Then
        .TableCell(tcRows) = i: .TableCell(tcFontSize, i) = 8: .TableCell(tcFontBold, i, 1) = True
        .TableCell(tcText, i, 1) = "Tipo de Documento: (" & Trim(fg_codigocbo(I_DetCom.Combo1, 3, 2, "")) & ")" & " - " & Trim(Mid(I_DetCom.Combo1(3).Text, 1, 50))
        Print #1, .TableCell(tcText, i, 1)
        i = i + 1
    End If
    vg_Consulta = ""
    v_rut = fg_DespintaRut(I_DetCom.fpText1(0).Text)
    If I_DetCom.Check1(0).Value = 1 Then
        vg_Consulta = "where toc_fecemi >= cdate('" & Format(I_DetCom.fecha(0).Text, "dd/mm/yyyy") & "') and toc_fecemi <= cdate('" & Format(I_DetCom.fecha(1).Text, "dd/mm/yyyy") & "')" & " and toc_tipdoc <> 'SN'"
    End If
    If I_DetCom.Check1(1).Value = 1 Then
        vg_Consulta = vg_Consulta & IIf(Trim(vg_Consulta) <> "", " and ", "where ") & "toc_codbod = " & fg_codigocbo(I_DetCom.Combo1, 0, 10, 0) & " and toc_tipdoc <> 'SN'"
    End If
    If I_DetCom.Check1(2).Value = 1 Then
        vg_Consulta = vg_Consulta & IIf(Trim(vg_Consulta) <> "", " and ", "where ") & "tip_codigo = " & fg_codigocbo(I_DetCom.Combo1, 1, 10, 0) & " and toc_tipdoc <> 'SN'"
    End If
    If I_DetCom.Check1(3).Value = 1 Then
        vg_Consulta = vg_Consulta & IIf(Trim(vg_Consulta) <> "", " and ", "where ") & " dec_codmer = '" & Trim(fg_codigocbo(I_DetCom.Combo1, 2, 20, 0)) & "'" & " and toc_tipdoc <> 'SN'"
    End If
    If I_DetCom.Check1(4).Value = 1 Then
        vg_Consulta = vg_Consulta & IIf(Trim(vg_Consulta) <> "", " and ", "where ") & "toc_rutpro = '" & v_rut & "'" & " and toc_tipdoc <> 'SN'"
    End If
    If I_DetCom.Check1(5).Value = 1 Then
        vg_Consulta = vg_Consulta & IIf(Trim(vg_Consulta) <> "", " and ", "where ") & "toc_tipdoc = '" & fg_codigocbo(I_DetCom.Combo1, 3, 2, "") & "'"
    End If
    RS1.Open "SELECT distinct* FROM a_unidad INNER JOIN ((b_proveedor INNER JOIN b_totcompras ON b_proveedor.prv_codigo = b_totcompras.toc_rutpro) INNER JOIN ((a_tipopro INNER JOIN b_productos ON a_tipopro.tip_codigo = b_productos.pro_codtip) INNER JOIN b_detcompras ON b_productos.pro_codigo = b_detcompras.dec_codmer) ON (b_totcompras.toc_numdoc = b_detcompras.dec_numdoc) AND (b_totcompras.toc_tipdoc = b_detcompras.dec_tipdoc) AND (b_totcompras.toc_rutpro = b_detcompras.dec_rutpro)) ON a_unidad.uni_codigo = b_productos.pro_coduni " & vg_Consulta & " order by toc_rutpro, toc_fecemi, toc_tipdoc,toc_numdoc", vg_db, adOpenStatic
    If RS1.EOF Then MsgBox "No existen datos para la consulta...", vbExclamation + vbOKOnly, Msgtitulo: RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
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
    i = 1: TGen = 0: TGeniva = 0: TGenexec = 0: TGenOImp = 0: TGennet = 0
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
                .TableCell(tcText, i, 5) = Format(RS1!dec_ptotal, fg_Pict(8, vg_DPr))
                Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|"; .TableCell(tcText, i, 5)
                TNet = TNet + RS1!dec_ptotal
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
    Loop
    i = i + 1
    .TableCell(tcFontBold, i, 2) = True:   .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, i, 2) = "Total General"
    .TableCell(tcFontBold, i, 5) = True:   .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, i, 5) = Format(TGen, fg_Pict(8, vg_DPr))
    Print #1, vbTab & vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & .TableCell(tcText, i, 5)
    i = i + 1
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Sub
Error_DetalleCom:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Close #1
    Exit Sub
End Sub

Sub i_Fampro()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Familia
Msgtitulo = "Informe Familias Productos"
Preview.Show
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
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
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_tipopro where tip_previo=0 order by tip_nombre", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
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
        RS2.Open "select * from a_tipopro where tip_previo=" & RS1!tip_codigo & " order by tip_nombre", vg_db, adOpenStatic
        If Not RS2.EOF Then
            Do While Not RS2.EOF
                .TableCell(tcText, i, 2) = RS2!tip_nombre
                Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2)
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
Exit Sub
Error_Familia:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Close #1
    Exit Sub
End Sub

Sub I_UniEnv()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_UniEnv
Msgtitulo = "Informe de Unidades de Envase"
Preview.Show
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
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
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_unidad order by uni_nombre", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Nombre"
    .TableCell(tcText, 1, 2) = "Nombre Corto"
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
        .TableCell(tcText, i, 1) = RS1!uni_nombre
        .TableCell(tcText, i, 2) = RS1!uni_nomcor
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2)
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Sub
Error_UniEnv:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Close #1
    Exit Sub
End Sub

Sub i_uniemb()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_UniEmb
Msgtitulo = "Informe de Unidades de Embalaje"
Preview.Show
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
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
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_embalaje order by emb_nombre", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Nombre"
    .TableCell(tcText, 1, 2) = "Nombre Corto"
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
        .TableCell(tcText, i, 1) = RS1!emb_nombre
        .TableCell(tcText, i, 2) = RS1!emb_nomcor
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2)
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Sub
Error_UniEmb:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Close #1
    Exit Sub
End Sub

Sub I_Nutri()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Nutriente
Msgtitulo = "Informe de Nutrientes"
Preview.Show
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
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
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_nutriente order by nut_secnro, nut_nombre", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
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
Exit Sub
Error_Nutriente:
     MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
     Close #1
     Exit Sub
End Sub

Sub I_impues()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Impto
Msgtitulo = "Informe de Impuestos"
Preview.Show
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .FontSize = 8
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Impuestos"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_impuesto order by imp_codigo", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Nombre"
    .TableCell(tcText, 1, 2) = "Impuesto"
    .TableCell(tcText, 1, 3) = "Incluye en Costo"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taCenterTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!imp_nombre
        .TableCell(tcText, i, 2) = Format(RS1!imp_pctimp, fg_Pict(6, 2))
        .TableCell(tcText, i, 3) = IIf(RS1!imp_inccos = 1, "SI", "NO")
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
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
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Close #1
    Exit Sub
End Sub

Sub I_Perfil()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Perfiles
Msgtitulo = "Informe de Pefiles"
Preview.Show
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
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
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_perfil order by per_nombre", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
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
Exit Sub
Error_Perfiles:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Close #1
    Exit Sub
End Sub

Sub I_acceso()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim codigo As Integer
On Local Error GoTo Error_Acceso
Msgtitulo = "Informe de Accesos"
Preview.Show
Preview.Refresh
M_Perfil.vaSpread1.Row = M_Perfil.vaSpread1.ActiveRow
M_Perfil.vaSpread1.Col = 2
Nombre = M_Perfil.vaSpread1.Text
M_Perfil.vaSpread1.Col = 1
codigo = M_Perfil.vaSpread1.Text
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
     LogoEmp
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Perfil: " & Nombre
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    RS1.Open "SELECT a_derechosperfil.*, a_opcsistema.opc_nombre FROM a_opcsistema INNER JOIN a_derechosperfil ON a_opcsistema.opc_codigo = a_derechosperfil.dpe_codopc WHERE a_derechosperfil.dpe_codper=" & codigo, vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
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
Exit Sub
Error_Acceso:
     MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Close #1
    Exit Sub
End Sub

Sub I_Usuari()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim codigo As Integer
On Local Error GoTo Error_Usuarios
Msgtitulo = "Informe de Usuarios"
Preview.Show
Preview.Refresh
M_Perfil.vaSpread1.Row = M_Perfil.vaSpread1.ActiveRow
M_Perfil.vaSpread1.Col = 2
Nombre = M_Perfil.vaSpread1.Text
M_Perfil.vaSpread1.Col = 1
codigo = M_Perfil.vaSpread1.Text
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
     LogoEmp
     vg_Archxls = fg_ArchivoTxt
     Open vg_Archxls For Output As #1
    .FontSize = 8
    .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Listado de Usuarios "
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
     .Text = Chr(13): .Text = Chr(13)
    RS1.Open "SELECT a_usuarios.usu_nombre, a_usuarios.usu_password, a_usuarios.usu_telefono, a_usuarios.usu_email, a_usuarios.usu_oficina, a_usuarios.usu_depart, a_perfil.per_nombre" & _
            " FROM a_perfil INNER JOIN a_usuarios ON a_perfil.per_codigo = a_usuarios.usu_perfil", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
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
Exit Sub
Error_Usuarios:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Close #1
    Exit Sub
End Sub

Sub I_DocProvee()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim TExe As Double, TNet As Double, tIVA As Double, TOtr As Double, tTot As Double
On Local Error GoTo Error_DocProvee
Msgtitulo = "Impresión Documentos Proveedor"
'----------------------------------------------------
'---Nombre de Informe: Informe documentos Proveedor (FA-GD,etc.)
'---Creador: Miguel Solorza P.(Corrección Alexis Morgado)
'---Fecha de Crecación: 20-07-2004.
'----------------------------------------------------
Preview.Show
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\DocProvee.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orPortrait
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .FontSize = 8
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .Text = Chr(13): .Text = Chr(13)
    Select Case vg_TDC
    Case "FA"
        .TableCell(tcText, 1, 1) = "Factura N° " & Trim(Str(vg_NDC))
    Case "GD"
        .TableCell(tcText, 1, 1) = "Guía de Despacho N° " & Trim(Str(vg_NDC))
    Case "NC"
        .TableCell(tcText, 1, 1) = "Nota de Crédito N° " & Trim(Str(vg_NDC))
    Case "ND"
        .TableCell(tcText, 1, 1) = "Nota de Debito N° " & Trim(Str(vg_NDC))
    Case "BO"
        .TableCell(tcText, 1, 1) = "Boleta N° " & Trim(Str(vg_NDC))
    Case "BH"
        .TableCell(tcText, 1, 1) = "Boleta de Honorarios N° " & Trim(Str(vg_NDC))
    End Select
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select a.*, b.*, c.*, d.*, e.uni_nomcor, f.bod_nombre from b_totcompras a, b_detcompras b, b_proveedor c, b_productos d, a_unidad e, a_bodega f where a.toc_rutpro=b.dec_rutpro and a.toc_tipdoc=b.dec_tipdoc and a.toc_numdoc=b.dec_numdoc and a.toc_rutpro=c.prv_codigo and b.dec_codmer=d.pro_codigo and d.pro_coduni=e.uni_codigo and a.toc_codbod=f.bod_codigo and a.toc_rutpro='" & vg_RDC & "' and a.toc_tipdoc='" & vg_TDC & "' and a.toc_numdoc=" & vg_NDC, vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 5
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcText, 1, 1) = "Nombre Proveedor"
    .TableCell(tcText, 1, 2) = RS1!prv_nombre
    .TableCell(tcText, 1, 3) = "Rut Proveedor"
    .TableCell(tcText, 1, 4) = fg_PintaRut(RS1!prv_codigo)
    .TableCell(tcText, 2, 1) = "Dirección"
    .TableCell(tcText, 2, 2) = RS1!prv_direccion
    .TableCell(tcText, 2, 3) = "Comuna"
    .TableCell(tcText, 2, 4) = RS1!prv_comuna
    .TableCell(tcText, 3, 1) = "Fecha Emisión"
    .TableCell(tcText, 3, 2) = Format(RS1!toc_fecemi, "dd/MM/yyyy")
    .TableCell(tcText, 3, 3) = "Fecha Vencimiento"
    .TableCell(tcText, 3, 4) = Format(RS1!toc_fecven, "dd/MM/yyyy")
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
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4600: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 600: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
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
    TExe = RS1!toc_exedoc: TNet = RS1!toc_netdoc: tIVA = RS1!toc_ivadoc: TOtr = RS1!toc_otrimp: tTot = RS1!toc_totdoc
    .StartTable
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4600: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 600: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!dec_codmer
        .TableCell(tcText, i, 2) = RS1!pro_nombre
        .TableCell(tcText, i, 3) = RS1!uni_nomcor
        .TableCell(tcText, i, 4) = Format(RS1!dec_canmer, fg_Pict(8, vg_DCa))
        .TableCell(tcText, i, 5) = Format(RS1!dec_precom, fg_Pict(8, vg_DPr))
        .TableCell(tcText, i, 6) = Format(RS1!dec_valdes, fg_Pict(8, vg_DPr))
        .TableCell(tcText, i, 7) = Format(RS1!dec_ptotal, fg_Pict(8, vg_DPr))
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                  .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)

        RS1.MoveNext: i = i + 1
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
    .TableCell(tcFontBold, i) = True: .TableCell(tcText, i, 2) = "Total IVA": .TableCell(tcText, i, 7) = Format(tIVA, fg_Pict(8, vg_DPr))
    Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & vbTab & vbTab & .TableCell(tcText, i, 7)
    i = i + 1
    .TableCell(tcFontBold, i) = True: .TableCell(tcText, i, 2) = "Total Otr.Imp.": .TableCell(tcText, i, 7) = Format(TOtr, fg_Pict(8, vg_DPr))
    Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & vbTab & vbTab & .TableCell(tcText, i, 7)
    i = i + 1
    .TableCell(tcFontBold, i) = True: .TableCell(tcText, i, 2) = "Total": .TableCell(tcText, i, 7) = Format(tTot, fg_Pict(8, vg_DPr))
    Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & vbTab & vbTab & .TableCell(tcText, i, 7)
    i = i + 1
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    '------ Busca solicitud---
    If fg_codigocbo(M_DocPro.Combo2, 0, 2, "") <> "FA" Then Close #1: GoTo Fin_Docto
    
    RS1.Open "select a.*, b.*, c.*, d.*, e.uni_nomcor, f.bod_nombre from b_totcompras a, b_detcompras b, b_proveedor c, b_productos d, a_unidad e, a_bodega f where a.toc_rutpro=b.dec_rutpro and a.toc_tipdoc=b.dec_tipdoc and a.toc_numdoc=b.dec_numdoc and a.toc_rutpro=c.prv_codigo and b.dec_codmer=d.pro_codigo and d.pro_coduni=e.uni_codigo and a.toc_codbod=f.bod_codigo and a.toc_rutpro='" & vg_RDC & "' and a.toc_tipdoc='" & "SN" & "' and trim(a.toc_docaso)='" & Trim(Str(M_DocPro.Double1(5).Value)) & "'", vg_db, adOpenStatic
    If RS1.EOF Then
        RS1.Close: Set RS1 = Nothing: Close #1
        GoTo Fin_Docto 'Cierra Docto.
    End If
    .NewPage 'Abre nueva pagina
    .Orientation = orLandscape
    .MarginLeft = 500
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
    LogoEmp
    .FontSize = 8
    '---MSP 08-09-2004--- Separa los dos documentos por medio de print
    Print #1, " ": Print #1, " ": Print #1, " ": Print #1, " ": Print #1, " "
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 13500: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .Text = Chr(13): .Text = Chr(13)
    .TableCell(tcText, 1, 1) = "Solicitud Nota de Crédito N° " & Trim(RS1!toc_numdoc)
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 4
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcText, 1, 1) = "Nombre Proveedor"
    .TableCell(tcText, 1, 2) = RS1!prv_nombre
    .TableCell(tcText, 1, 3) = "Rut Proveedor"
    .TableCell(tcText, 1, 4) = fg_PintaRut(RS1!prv_codigo)
    .TableCell(tcText, 2, 1) = "Dirección"
    .TableCell(tcText, 2, 2) = RS1!prv_direccion
    .TableCell(tcText, 2, 3) = "Comuna"
    .TableCell(tcText, 2, 4) = RS1!prv_comuna
    .TableCell(tcText, 3, 1) = "Fecha Emisión"
    .TableCell(tcText, 3, 2) = Format(RS1!toc_fecemi, "dd/MM/yyyy")
    .TableCell(tcText, 3, 3) = "Fecha Vencimiento"
    .TableCell(tcText, 3, 4) = Format(RS1!toc_fecven, "dd/MM/yyyy")
    .TableCell(tcText, 4, 1) = "Bodega"
    .TableCell(tcText, 4, 2) = RS1!bod_nombre
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    Print #1, .TableCell(tcText, 4, 1) & "|" & .TableCell(tcText, 4, 2)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 10: .TableCell(tcRows) = 1
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
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
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
        .TableCell(tcText, i, 4) = RS1!dec_canmer
        .TableCell(tcText, i, 5) = Format(RS1!dec_canrec, fg_Pict(8, vg_DCa))
        .TableCell(tcText, i, 6) = Format((RS1!dec_canmer - RS1!dec_canrec), fg_Pict(8, vg_DCa))
        .TableCell(tcText, i, 7) = Format(RS1!dec_precom, fg_Pict(8, vg_DPr))
        .TableCell(tcText, i, 8) = Format(RS1!dec_prerec, fg_Pict(8, vg_DPr))
        .TableCell(tcText, i, 9) = Format((RS1!dec_precom - RS1!dec_prerec), fg_Pict(8, vg_DPr))
        .TableCell(tcText, i, 10) = Format((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_prerec, fg_Pict(8, vg_DPr))
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                  .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)
        RS1.MoveNext: i = i + 1
    Loop
    '--- Fin imprime detalle
    '--- Trailer de Documento
    i = i + 1
    .TableCell(tcFontBold, i) = True: .TableCell(tcText, i, 2) = "Total Neto": .TableCell(tcText, i, 10) = Format(TNet, fg_Pict(8, vg_DPr))
     Print #1, vbTab & .TableCell(tcText, i, 2) & vbTab & vbTab & vbTab & vbTab & vbTab & .TableCell(tcText, i, 10)
     i = i + 1
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    '--- Fin Trailer
Fin_Docto:
    .EndDoc
End With
Exit Sub
Error_DocProvee:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Close #1
    Exit Sub
End Sub

