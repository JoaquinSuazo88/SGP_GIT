Attribute VB_Name = "InforMB"
'Dim RS1 As New ADODB.Recordset, RS2 As New ADODB.Recordset, RS3 As New ADODB.Recordset, RS4 As New ADODB.Recordset
'Dim inf_nreceta As String, cdetalle As String, opcionsalto As String
'Dim inf_opcion As Integer
'Dim porcentaje As Integer
'Dim j As Integer
'Dim k As Integer
'
'Public Function I_Inv_ABC_Stock(cForm As Form, codbod As Long)
'Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
'Dim auxpro As String, aAp As String, sql1 As String, sql2 As String, auxtip As Long
'
'On Local Error GoTo Error_Productos
'Tot_General = 0
'porcentaje = 10
'Fecha_Car = Mid(fecini, 7, 2)
'Fecha_Car = Fecha_Car + "/" + Mid(fecini, 5, 2)
'Fecha_Car = Fecha_Car + "/" + Mid(fecini, 1, 4)
'fg_carga ""
'mgstitulo = "Informe para Toma de Inventario"
'j = 2
'Preview.Refresh
'With Preview.VSPrinter
'    .Styles.Apply "Default"
'    .ExportFormat = vpxRTF
''    .ExportFile = App.Path & "\Reporte.rtf"
'    .ExportFile = vg_reporte
'    .Preview = True
'    .PreviewPage = 1
'    .Orientation = orLandscape 'orPortrait
'    .MarginLeft = 1500
'    .StartDoc
'    .PageBorder = 0
'    .HdrFontName = "Arial"
'    .HdrFontSize = 9
'    .HdrFontBold = False
'    .Header = "" & fg_poneencpagina & "||"
'    .Footer = "" & fg_ponepiepagina & "||Página : %d"
'    ExportHeaderFooter Preview.VSPrinter
'    .FontSize = 9
'    vg_Archxls = fg_ArchivoTxt
'    Open vg_Archxls For Output As #1
'    LogoEmp
'
'    '-------> Traer curva ABC
'    RS1.Open "SELECT * FROM a_curvaabc", vg_db, adOpenStatic
'    If Not RS1.EOF Then
'       Do While Not RS1.EOF
'          If RS1!abc_codigo = "A" Then curvaa = RS1!abc_porce
'          If RS1!abc_codigo = "B" Then curvab = RS1!abc_porce
'          If RS1!abc_codigo = "C" Then curvac = RS1!abc_porce
'          RS1.MoveNext
'       Loop
'    End If
'    RS1.Close: Set RS1 = Nothing
'
'    vg_db.Execute "DELETE FROM tem_curvaABC"
'
'    .StartTable
'    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
'    .TableCell(tcColWidth, , 1) = 10000: .TableCell(tcAlign, , 1) = taCenterMiddle
'    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
'    .TableCell(tcText, 1, 1) = "Informe para Toma de Inventario"
'    '.TableCell(tcText, 2, 1) = "Fecha Informe:  " + CDate(Date)
'    Print #1, .TableCell(tcText, 1, 1)
'    .TableBorder = tbNone
'    .EndTable
'    .text = Chr(13): .text = Chr(13)
'
'    aAp1 = Trim(vg_NUsr) & "_tmp_DetCurvaABC_X"
'    fg_CheckTmp aAp1
'
'    RS2.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, a.pro_coduni, b.bod_canmer " & _
'             "INTO " & aAp1 & " FROM b_productos a, b_bodegas b " & _
'             "WHERE a.pro_codigo = b.bod_codpro " & _
'             "AND b.bod_codbod = " & vg_codbod & " " & _
'             "AND b.bod_canmer <> 0 " & _
'             "ORDER BY b.bod_canmer DESC", vg_db, adOpenStatic
'
'    '-------> Consultar total general curva ABC
'    totgrl = 0
'    RS2.Open "SELECT SUM(bod_canmer) AS totgra FROM " & aAp1 & "", vg_db
'    If Not RS2.EOF Then totgrl = IIf(IsNull(RS2!totgra), 0, RS2!totgra)
'    RS2.Close: Set RS2 = Nothing
'
'    RS2.Open "SELECT a.pro_codigo, a.pro_nombre, b.uni_nomcor, round(a.bod_canmer,3) as cantidad " & _
'             "FROM " & aAp1 & " a, a_unidad b " & _
'             "WHERE a.pro_coduni = b.uni_codigo", vg_db, adOpenStatic
'
'    indcur = 1: curvaABC = curvaa: curva = 0: totcur = 0
'    If RS2.EOF Then MsgBox "No existe información para crear informe para toma de inventario": fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
'    .StartTable
'    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
'    .TableCell(tcColWidth, , 1) = 1200: .TableCell(tcAlign, , 1) = taLeftTop
'    .TableCell(tcColWidth, , 2) = 4800: .TableCell(tcAlign, , 2) = taLeftTop
'    .TableCell(tcColWidth, , 3) = 1350: .TableCell(tcAlign, , 3) = taLeftTop
'    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taLeftTop
'    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
'    .TableCell(tcText, 1, 1) = "Código Producto"
'    .TableCell(tcText, 1, 2) = "Descripción"
'    .TableCell(tcText, 1, 3) = "Unidad"
'    .TableCell(tcText, 1, 4) = "Cantidad"
'    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
'    .TableBorder = tbBox
'    .EndTable
'
'    .StartTable
'    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
'    .TableCell(tcColWidth, , 1) = 1200: .TableCell(tcAlign, , 1) = taLeftTop
'    .TableCell(tcColWidth, , 2) = 4800: .TableCell(tcAlign, , 2) = taLeftTop
'    .TableCell(tcColWidth, , 3) = 800: .TableCell(tcAlign, , 3) = taLeftTop
'    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taRightTop
'
'    i = 1
'    If Not RS2.EOF Then
'        i = i + 1
'        Do While Not RS2.EOF
'            If totgrl > 0 And Not IsNull(RS2!cantidad) Then
'               curva = curva + ((RS2!cantidad / totgrl) * 100)
'            Else
'               curva = curva + 0
'            End If
'
'            If curva > curvaABC And curvaABC <> curvac And curva <= 99 Then
'               Llenar_Informe k
'               curvaABC = IIf(indcur = 1, curvab, curvac): curva = 0
'               curva = curva + ((RS2!cantidad / totgrl) * 100)
'               .TableCell(tcColSpan, i, 5) = 2
'               indcur = indcur + 1: totcur = 0
'            End If
'
'            vg_db.Execute "INSERT INTO tem_curvaABC (tem_codigo, tem_nombre, tem_nomcor, tem_cantidad) " & _
'            "VALUES (" & Trim(RS2!pro_codigo) & ", '" & Trim(RS2!pro_nombre) & "', '" & Trim(RS2!uni_nomcor) & "', " & Trim(RS2!cantidad) & ")"
'
'           RS2.MoveNext
'        Loop
'    End If
'    Llenar_Informe k
'    .TableBorder = tbNone
'    .EndTable
'    RS2.Close: Set RS2 = Nothing
'    .EndDoc
'End With
'Preview.Show 1
'Exit Function
'
'Error_Productos:
'    fg_descarga
'    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
'    Close #1
'    Exit Function
'End Function
'
'Function Llenar_Informe(k)
'    With Preview.VSPrinter
'    RS3.Open "SELECT SUM(tem_cantidad) AS totgra FROM tem_curvaABC", vg_db
'    If Not RS3.EOF Then totgrl = IIf(IsNull(RS3!totgra), 0, RS3!totgra)
'    RS3.Close: Set RS3 = Nothing
'
'    RS3.Open "SELECT * FROM tem_curvaABC ORDER BY tem_cantidad DESC", vg_db, adOpenStatic
'    If Not RS3.EOF Then
'        Do While Not RS3.EOF
'            If totgrl > 0 And Not IsNull(RS3!tem_cantidad) Then
'               curva = curva + ((RS3!tem_cantidad / totgrl) * 100)
'            Else
'               curva = curva + 0
'            End If
'
'            If curva > 0 And curva <= 10 Then
'               .TableCell(tcText, j, 1) = Trim(RS3!tem_codigo)
'               .TableCell(tcText, j, 2) = Trim(RS3!tem_nombre)
'               .TableCell(tcText, j, 3) = Trim(RS3!tem_nomcor)
'               .TableCell(tcText, j, 4) = "__________________"
''              .TableCell(tcText, j, 4) = Trim(RS3!tem_cantidad)
'               j = j + 1: curva = 0
'               curva = curva + ((RS3!tem_cantidad / totgrl) * 100)
'            End If
'        RS3.MoveNext
'        Loop
'        k = j
'        vg_db.Execute "DELETE FROM tem_curvaABC"
'        RS3.Close: Set RS3 = Nothing
'    End If
'    End With
'End Function
'
'Public Function I_Inv_Gen_Stock(cForm As Form, codbod As Long, porinv As Long)
'Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
'Dim auxpro As String, aAp As String, sql1 As String, sql2 As String, auxtip As Long
'
'On Local Error GoTo Error_Productos
'Tot_General = 0
'porcentaje = 10
'Fecha_Car = Mid(fecini, 7, 2)
'Fecha_Car = Fecha_Car + "/" + Mid(fecini, 5, 2)
'Fecha_Car = Fecha_Car + "/" + Mid(fecini, 1, 4)
'fg_carga ""
'mgstitulo = "Informe para Toma de Inventario"
'j = 2
'Preview.Refresh
'With Preview.VSPrinter
'    .Styles.Apply "Default"
'    .ExportFormat = vpxRTF
''    .ExportFile = App.Path & "\Reporte.rtf"
'    .ExportFile = vg_reporte
'    .Preview = True
'    .PreviewPage = 1
'    .Orientation = orLandscape 'orPortrait
'    .MarginLeft = 1500
'    .StartDoc
'    .PageBorder = 0
'    .HdrFontName = "Arial"
'    .HdrFontSize = 9
'    .HdrFontBold = False
'    .Header = "" & fg_poneencpagina & "||"
'    .Footer = "" & fg_ponepiepagina & "||Página : %d"
'    ExportHeaderFooter Preview.VSPrinter
'    .FontSize = 9
'    vg_Archxls = fg_ArchivoTxt
'    Open vg_Archxls For Output As #1
'    LogoEmp
'
'    .StartTable
'    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
'    .TableCell(tcColWidth, , 1) = 10000: .TableCell(tcAlign, , 1) = taCenterMiddle
'    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
'    .TableCell(tcText, 1, 1) = "Informe para Toma de Inventario"
''    .TableCell(tcText, 2, 1) = "Fecha Informe:  " + Fecha_Car
'    Print #1, .TableCell(tcText, 1, 1)
'    .TableBorder = tbNone
'    .EndTable
'    .text = Chr(13): .text = Chr(13)
'
'    aAp1 = Trim(vg_NUsr) & "_tmp_DetCurvaABC_X"
'    fg_CheckTmp aAp1
'
'    RS2.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, a.pro_coduni, b.bod_canmer " & _
'             "INTO " & aAp1 & " FROM b_productos a, b_bodegas b " & _
'             "WHERE a.pro_codigo = b.bod_codpro " & _
'             "AND b.bod_codbod = " & vg_codbod & " " & _
'             "AND b.bod_canmer <> 0 " & _
'             "ORDER BY b.bod_canmer DESC", vg_db, adOpenStatic
'
'    totgrl = 0
'    RS2.Open "SELECT SUM(bod_canmer) AS totgra FROM " & aAp1 & "", vg_db
'    If Not RS2.EOF Then totgrl = IIf(IsNull(RS2!totgra), 0, RS2!totgra)
'    RS2.Close: Set RS2 = Nothing
'
'    RS2.Open "SELECT a.pro_codigo, a.pro_nombre, b.uni_nomcor, round(a.bod_canmer,3) as cantidad " & _
'             "FROM " & aAp1 & " a, a_unidad b " & _
'             "WHERE a.pro_coduni = b.uni_codigo", vg_db, adOpenStatic
'
'    indcur = 1: curvaABC = curvaa: curva = 0: totcur = 0
'    If RS2.EOF Then MsgBox "No existe información para crear informe para toma de inventario": fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
'    .StartTable
'    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
'    .TableCell(tcColWidth, , 1) = 1200: .TableCell(tcAlign, , 1) = taLeftTop
'    .TableCell(tcColWidth, , 2) = 4800: .TableCell(tcAlign, , 2) = taLeftTop
'    .TableCell(tcColWidth, , 3) = 2350: .TableCell(tcAlign, , 3) = taLeftTop
'    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taLeftTop
'    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
'    .TableCell(tcText, 1, 1) = "Código Producto"
'    .TableCell(tcText, 1, 2) = "Descripción"
'    .TableCell(tcText, 1, 3) = "Unidad"
'    .TableCell(tcText, 1, 4) = "Cantidad"
'    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
'    .TableBorder = tbBox
'    .EndTable
'
'    .StartTable
'    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
'    .TableCell(tcColWidth, , 1) = 1200: .TableCell(tcAlign, , 1) = taLeftTop
'    .TableCell(tcColWidth, , 2) = 4800: .TableCell(tcAlign, , 2) = taLeftTop
'    .TableCell(tcColWidth, , 3) = 800: .TableCell(tcAlign, , 3) = taLeftTop
'    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taRightTop
'
'    i = 1
'    If Not RS2.EOF Then
'        Do While Not RS2.EOF
'            If totgrl > 0 And Not IsNull(RS2!cantidad) Then
'               'curva = curva + ((RS2!xcostot / totgrl) * 100)
'               valor = valor + ((RS2!cantidad / totgrl) * 100)
'            Else
'               valor = valor + 0
'            End If
'
'            If valor > 0 And valor <= porinv Then
'               .TableCell(tcText, i, 1) = Trim(RS2!pro_codigo)
'               .TableCell(tcText, i, 2) = Trim(RS2!pro_nombre)
'               .TableCell(tcText, i, 3) = Trim(RS2!uni_nomcor)
''               .TableCell(tcText, i, 4) = "__________________"
'               .TableCell(tcText, i, 4) = Trim(RS2!cantidad)
'               i = i + 1
'               'valor = 0
'               'valor = valor + ((RS2!cantidad / totgrl) * 100)
'            End If
'        RS2.MoveNext
'        Loop
'    End If
'    .TableBorder = tbNone
'    .EndTable
'    RS2.Close: Set RS2 = Nothing
'    .EndDoc
'End With
'Preview.Show 1
'Exit Function
'
'Error_Productos:
'    fg_descarga
'    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
'    Close #1
'    Exit Function
'End Function
'
'Public Function I_Inv_ABC_Req(cForm As Form)
'Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
'Dim auxpro As String, aAp As String, sql1 As String, sql2 As String, auxtip As Long
'
'On Local Error GoTo Error_Productos
'Tot_General = 0
'porcentaje = 10
'Fecha_Car = Mid(fecini, 7, 2)
'Fecha_Car = Fecha_Car + "/" + Mid(fecini, 5, 2)
'Fecha_Car = Fecha_Car + "/" + Mid(fecini, 1, 4)
'fg_carga ""
'mgstitulo = "Informe para Toma de Inventario"
'j = 2
'Preview.Refresh
'With Preview.VSPrinter
'    .Styles.Apply "Default"
'    .ExportFormat = vpxRTF
''    .ExportFile = App.Path & "\Reporte.rtf"
'    .ExportFile = vg_reporte
'    .Preview = True
'    .PreviewPage = 1
'    .Orientation = orLandscape 'orPortrait
'    .MarginLeft = 1500
'    .StartDoc
'    .PageBorder = 0
'    .HdrFontName = "Arial"
'    .HdrFontSize = 9
'    .HdrFontBold = False
'    .Header = "" & fg_poneencpagina & "||"
'    .Footer = "" & fg_ponepiepagina & "||Página : %d"
'    ExportHeaderFooter Preview.VSPrinter
'    .FontSize = 9
'    vg_Archxls = fg_ArchivoTxt
'    Open vg_Archxls For Output As #1
'    LogoEmp
'
'    '-------> Traer curva ABC
'    RS1.Open "SELECT * FROM a_curvaabc", vg_db, adOpenStatic
'    If Not RS1.EOF Then
'       Do While Not RS1.EOF
'          If RS1!abc_codigo = "A" Then curvaa = RS1!abc_porce
'          If RS1!abc_codigo = "B" Then curvab = RS1!abc_porce
'          If RS1!abc_codigo = "C" Then curvac = RS1!abc_porce
'          RS1.MoveNext
'       Loop
'    End If
'    RS1.Close: Set RS1 = Nothing
'
'    vg_db.Execute "DELETE FROM tem_curvaABC"
'
'    .StartTable
'    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
'    .TableCell(tcColWidth, , 1) = 10000: .TableCell(tcAlign, , 1) = taCenterMiddle
'    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
'    .TableCell(tcText, 1, 1) = "Informe para Toma de Inventario"
'    '.TableCell(tcText, 2, 1) = "Fecha Informe:  " + CDate(Date)
'    Print #1, .TableCell(tcText, 1, 1)
'    .TableBorder = tbNone
'    .EndTable
'    .text = Chr(13): .text = Chr(13)
'
'    aAp1 = Trim(vg_NUsr) & "_tmp_DetCurvaABC_X"
'    fg_CheckTmp aAp1
'
'    RS2.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, a.pro_coduni, sum(b.dev_canmer) as cantidad " & _
'             "INTO " & aAp1 & " FROM b_productos a, b_detventas b " & _
'             "WHERE a.pro_codigo = b.dev_codmer " & _
'             "AND dev_rutcli = '" & MuestraCasino(1) & "' " & _
'             "AND b.dev_canmer <> 0 " & _
'             "GROUP BY a.pro_codigo, a.pro_nombre, a.pro_coduni", vg_db, adOpenStatic
'
'    '-------> Consultar total general curva ABC
'    totgrl = 0
'    RS2.Open "SELECT SUM(cantidad) AS totgra FROM " & aAp1 & "", vg_db
'    If Not RS2.EOF Then totgrl = IIf(IsNull(RS2!totgra), 0, RS2!totgra)
'    RS2.Close: Set RS2 = Nothing
'
'    RS2.Open "SELECT a.pro_codigo, a.pro_nombre, b.uni_nomcor, cantidad " & _
'             "FROM " & aAp1 & " a, a_unidad b " & _
'             "WHERE a.pro_coduni = b.uni_codigo " & _
'             "ORDER BY cantidad DESC", vg_db, adOpenStatic
'
'    indcur = 1: curvaABC = curvaa: curva = 0: totcur = 0
'    If RS2.EOF Then MsgBox "No existe información para crear informe para toma de inventario": fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
'    .StartTable
'    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
'    .TableCell(tcColWidth, , 1) = 1200: .TableCell(tcAlign, , 1) = taLeftTop
'    .TableCell(tcColWidth, , 2) = 4800: .TableCell(tcAlign, , 2) = taLeftTop
'    .TableCell(tcColWidth, , 3) = 1350: .TableCell(tcAlign, , 3) = taLeftTop
'    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taLeftTop
'    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
'    .TableCell(tcText, 1, 1) = "Código Producto"
'    .TableCell(tcText, 1, 2) = "Descripción"
'    .TableCell(tcText, 1, 3) = "Unidad"
'    .TableCell(tcText, 1, 4) = "Cantidad"
'    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
'    .TableBorder = tbBox
'    .EndTable
'
'    .StartTable
'    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
'    .TableCell(tcColWidth, , 1) = 1200: .TableCell(tcAlign, , 1) = taLeftTop
'    .TableCell(tcColWidth, , 2) = 4800: .TableCell(tcAlign, , 2) = taLeftTop
'    .TableCell(tcColWidth, , 3) = 800: .TableCell(tcAlign, , 3) = taLeftTop
'    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taRightTop
'
'    i = 1
'    If Not RS2.EOF Then
'        i = i + 1
'        Do While Not RS2.EOF
'            If totgrl > 0 And Not IsNull(RS2!cantidad) Then
'               curva = curva + ((RS2!cantidad / totgrl) * 100)
'            Else
'               curva = curva + 0
'            End If
'
'            If curva > curvaABC And curvaABC <> curvac And curva <= 99 Then
'               Llenar_Informe k
'               curvaABC = IIf(indcur = 1, curvab, curvac): curva = 0
'               curva = curva + ((RS2!cantidad / totgrl) * 100)
'               .TableCell(tcColSpan, i, 5) = 2
'               indcur = indcur + 1: totcur = 0
'            End If
'
'            vg_db.Execute "INSERT INTO tem_curvaABC (tem_codigo, tem_nombre, tem_nomcor, tem_cantidad) " & _
'            "VALUES (" & Trim(RS2!pro_codigo) & ", '" & Trim(RS2!pro_nombre) & "', '" & Trim(RS2!uni_nomcor) & "', " & Trim(RS2!cantidad) & ")"
'           RS2.MoveNext
'        Loop
'    End If
'    Llenar_Informe k
'    .TableBorder = tbNone
'    .EndTable
'    RS2.Close: Set RS2 = Nothing
'    .EndDoc
'End With
'Preview.Show 1
'Exit Function
'
'Error_Productos:
'    fg_descarga
'    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
'    Close #1
'    Exit Function
'End Function
'
'Public Function I_Inv_Gen_Req(cForm As Form, porinv As Long)
'Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
'Dim auxpro As String, aAp As String, sql1 As String, sql2 As String, auxtip As Long
'
'On Local Error GoTo Error_Productos
'Tot_General = 0
'porcentaje = 10
'Fecha_Car = Mid(fecini, 7, 2)
'Fecha_Car = Fecha_Car + "/" + Mid(fecini, 5, 2)
'Fecha_Car = Fecha_Car + "/" + Mid(fecini, 1, 4)
'fg_carga ""
'mgstitulo = "Informe para Toma de Inventario"
'j = 2
'Preview.Refresh
'With Preview.VSPrinter
'    .Styles.Apply "Default"
'    .ExportFormat = vpxRTF
''    .ExportFile = App.Path & "\Reporte.rtf"
'    .ExportFile = vg_reporte
'    .Preview = True
'    .PreviewPage = 1
'    .Orientation = orLandscape 'orPortrait
'    .MarginLeft = 1500
'    .StartDoc
'    .PageBorder = 0
'    .HdrFontName = "Arial"
'    .HdrFontSize = 9
'    .HdrFontBold = False
'    .Header = "" & fg_poneencpagina & "||"
'    .Footer = "" & fg_ponepiepagina & "||Página : %d"
'    ExportHeaderFooter Preview.VSPrinter
'    .FontSize = 9
'    vg_Archxls = fg_ArchivoTxt
'    Open vg_Archxls For Output As #1
'    LogoEmp
'
'    .StartTable
'    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
'    .TableCell(tcColWidth, , 1) = 10000: .TableCell(tcAlign, , 1) = taCenterMiddle
'    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
'    .TableCell(tcText, 1, 1) = "Informe para Toma de Inventario"
''    .TableCell(tcText, 2, 1) = "Fecha Informe:  " + Fecha_Car
'    Print #1, .TableCell(tcText, 1, 1)
'    .TableBorder = tbNone
'    .EndTable
'    .text = Chr(13): .text = Chr(13)
'
'    aAp1 = Trim(vg_NUsr) & "_tmp_DetCurvaABC_X"
'    fg_CheckTmp aAp1
'
'    RS2.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, a.pro_coduni, sum(b.dev_canmer) as cantidad " & _
'             "INTO " & aAp1 & " FROM b_productos a, b_detventas b " & _
'             "WHERE a.pro_codigo = b.dev_codmer " & _
'             "AND dev_rutcli = '" & MuestraCasino(1) & "' " & _
'             "AND b.dev_canmer <> 0 " & _
'             "GROUP BY a.pro_codigo, a.pro_nombre, a.pro_coduni", vg_db, adOpenStatic
'
'    totgrl = 0
'    RS2.Open "SELECT SUM(cantidad) AS totgra FROM " & aAp1 & "", vg_db
'    If Not RS2.EOF Then totgrl = IIf(IsNull(RS2!totgra), 0, RS2!totgra)
'    RS2.Close: Set RS2 = Nothing
'
'    RS2.Open "SELECT a.pro_codigo, a.pro_nombre, b.uni_nomcor, cantidad " & _
'             "FROM " & aAp1 & " a, a_unidad b " & _
'             "WHERE a.pro_coduni = b.uni_codigo " & _
'             "ORDER BY cantidad DESC", vg_db, adOpenStatic
'
'    indcur = 1: curvaABC = curvaa: curva = 0: totcur = 0
'    If RS2.EOF Then MsgBox "No existe información para crear informe para toma de inventario": fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
'    .StartTable
'    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
'    .TableCell(tcColWidth, , 1) = 1200: .TableCell(tcAlign, , 1) = taLeftTop
'    .TableCell(tcColWidth, , 2) = 4800: .TableCell(tcAlign, , 2) = taLeftTop
'    .TableCell(tcColWidth, , 3) = 2350: .TableCell(tcAlign, , 3) = taLeftTop
'    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taLeftTop
'    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
'    .TableCell(tcText, 1, 1) = "Código Producto"
'    .TableCell(tcText, 1, 2) = "Descripción"
'    .TableCell(tcText, 1, 3) = "Unidad"
'    .TableCell(tcText, 1, 4) = "Cantidad"
'    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
'    .TableBorder = tbBox
'    .EndTable
'
'    .StartTable
'    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
'    .TableCell(tcColWidth, , 1) = 1200: .TableCell(tcAlign, , 1) = taLeftTop
'    .TableCell(tcColWidth, , 2) = 4800: .TableCell(tcAlign, , 2) = taLeftTop
'    .TableCell(tcColWidth, , 3) = 800: .TableCell(tcAlign, , 3) = taLeftTop
'    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taRightTop
'
'    i = 1
'    If Not RS2.EOF Then
'        Do While Not RS2.EOF
'            If totgrl > 0 And Not IsNull(RS2!cantidad) Then
'               'curva = curva + ((RS2!xcostot / totgrl) * 100)
'               valor = valor + ((RS2!cantidad / totgrl) * 100)
'            Else
'               valor = valor + 0
'            End If
'
'            If valor > 0 And valor <= porinv Then
'               .TableCell(tcText, i, 1) = Trim(RS2!pro_codigo)
'               .TableCell(tcText, i, 2) = Trim(RS2!pro_nombre)
'               .TableCell(tcText, i, 3) = Trim(RS2!uni_nomcor)
''               .TableCell(tcText, i, 4) = "__________________"
'               .TableCell(tcText, i, 4) = Trim(RS2!cantidad)
'               i = i + 1
'               'valor = 0
'               'valor = valor + ((RS2!cantidad / totgrl) * 100)
'            End If
'        RS2.MoveNext
'        Loop
'    End If
'    .TableBorder = tbNone
'    .EndTable
'    RS2.Close: Set RS2 = Nothing
'    .EndDoc
'End With
'Preview.Show 1
'Exit Function
'
'Error_Productos:
'    fg_descarga
'    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
'    Close #1
'    Exit Function
'End Function
'
''***********************COPIA RESPALDO FUNCION*****************************
''Public Function I_Productos_Inv(cForm As Form, codbod As Long, codpro As String, codtip As Long, fecini As Long, fecfin As Long)
''Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
''Dim auxpro As String, aAp As String, sql1 As String, sql2 As String, auxtip As Long
''
''On Local Error GoTo Error_Productos
''Tot_General = 0
''Fecha_Car = Mid(fecini, 7, 2)
''Fecha_Car = Fecha_Car + "/" + Mid(fecini, 5, 2)
''Fecha_Car = Fecha_Car + "/" + Mid(fecini, 1, 4)
''fg_carga ""
''mgstitulo = "Informe para Toma de Inventario"
''Preview.Refresh
''With Preview.VSPrinter
''    .Styles.Apply "Default"
''    .ExportFormat = vpxRTF
'''    .ExportFile = App.Path & "\Reporte.rtf"
''    .ExportFile = vg_reporte
''    .Preview = True
''    .PreviewPage = 1
''    .Orientation = orLandscape 'orPortrait
''    .MarginLeft = 1500
''    .StartDoc
''    .PageBorder = 0
''    .HdrFontName = "Arial"
''    .HdrFontSize = 9
''    .HdrFontBold = False
''    .Header = "" & fg_poneencpagina & "||"
''    .Footer = "" & fg_ponepiepagina & "||Página : %d"
''    ExportHeaderFooter Preview.VSPrinter
''    .FontSize = 9
''    vg_Archxls = fg_ArchivoTxt
''    Open vg_Archxls For Output As #1
''    LogoEmp
''
''    '-------> Traer curva ABC
''    RS1.Open "SELECT * FROM a_curvaabc", vg_db, adOpenStatic
''    If Not RS1.EOF Then
''       Do While Not RS1.EOF
''          If RS1!abc_codigo = "A" Then curvaa = RS1!abc_porce
''          If RS1!abc_codigo = "B" Then curvab = RS1!abc_porce
''          If RS1!abc_codigo = "C" Then curvac = RS1!abc_porce
''          RS1.MoveNext
''       Loop
''    End If
''    RS1.Close: Set RS1 = Nothing
''
''    .StartTable
''    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
''    .TableCell(tcColWidth, , 1) = 10000: .TableCell(tcAlign, , 1) = taCenterMiddle
''    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
''    .TableCell(tcText, 1, 1) = "Informe para Toma de Inventario"
''    .TableCell(tcText, 2, 1) = "Fecha Informe:  " + Fecha_Car
''    Print #1, .TableCell(tcText, 1, 1)
''    .TableBorder = tbNone
''    .EndTable
''    .text = Chr(13): .text = Chr(13)
''
''    aAp1 = Trim(vg_NUsr) & "_tmp_DetCurvaABC_X"
''    fg_CheckTmp aAp1
''
''    RS2.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, a.pro_coduni, b.bod_canmer " & _
''             "INTO " & aAp1 & " FROM b_productos a, b_bodegas b " & _
''             "WHERE a.pro_codigo = b.bod_codpro " & _
''             "AND b.bod_codbod = " & vg_codbod & " " & _
''             "AND b.bod_canmer <> 0 " & _
''             "ORDER BY b.bod_canmer DESC", vg_db, adOpenStatic
''
''    '-------> Consultar total general curva ABC
''    totgrl = 0
''    RS2.Open "SELECT SUM(bod_canmer) AS totgra FROM " & aAp1 & "", vg_db
''    If Not RS2.EOF Then totgrl = IIf(IsNull(RS2!totgra), 0, RS2!totgra)
''    RS2.Close: Set RS2 = Nothing
''
''    RS2.Open "SELECT a.pro_codigo, a.pro_nombre, b.uni_nomcor, round(a.bod_canmer,3) as cantidad " & _
''             "FROM " & aAp1 & " a, a_unidad b " & _
''             "WHERE a.pro_coduni = b.uni_codigo", vg_db, adOpenStatic
''
''    indcur = 1: curvaABC = curvaa: curva = 0: totcur = 0
''    If RS2.EOF Then MsgBox "No existe información": fg_descarga: RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
''    .StartTable
''    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
''    .TableCell(tcColWidth, , 1) = 1200: .TableCell(tcAlign, , 1) = taLeftTop
''    .TableCell(tcColWidth, , 2) = 4800: .TableCell(tcAlign, , 2) = taLeftTop
''    .TableCell(tcColWidth, , 3) = 800: .TableCell(tcAlign, , 3) = taLeftTop
''    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taRightTop
''    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
''    .TableCell(tcText, 1, 1) = "Código Producto"
''    .TableCell(tcText, 1, 2) = "Descripción"
''    .TableCell(tcText, 1, 3) = "Unidad"
''    .TableCell(tcText, 1, 4) = "Cantidad"
''    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
''    .TableBorder = tbBox
''    .EndTable
''
''    .StartTable
''    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
''    .TableCell(tcColWidth, , 1) = 1200: .TableCell(tcAlign, , 1) = taLeftTop
''    .TableCell(tcColWidth, , 2) = 4800: .TableCell(tcAlign, , 2) = taLeftTop
''    .TableCell(tcColWidth, , 3) = 800: .TableCell(tcAlign, , 3) = taLeftTop
''    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taRightTop
''
''    i = 1
''    If Not RS2.EOF Then
''        '.TableCell(tcFontBold, i, 1) = True: .TableCell(tcText, i, 1) = "Curva A"
''        i = i + 1
''        Do While Not RS2.EOF
''            If totgrl > 0 And Not IsNull(RS2!cantidad) Then
''               curva = curva + ((RS2!cantidad / totgrl) * 100)
''            Else
''               curva = curva + 0
''            End If
''
''            'If curvaABC <> curvac And curva <= 99 Then
''            If curva > curvaABC And curvaABC <> curvac And curva <= 99 Then
''               curvaABC = IIf(indcur = 1, curvab, curvac): curva = 0
''               curva = curva + ((RS2!cantidad / totgrl) * 100)
''               '------- Imprimir total curva
''               i = i + 1
''               .TableCell(tcColSpan, i, 5) = 2
''               '.TableCell(tcFontBold, i, 2) = True: .TableCell(tcText, i, 2) = IIf(indcur = 1, "Total General Curva A ", IIf(indcur = 2, "Total General Curva B", "Total General Curva C"))
''               '.TableCell(tcFontBold, i, 4) = True: .TableCell(tcText, i, 4) = Format(totcur, fg_Pict(6, 2))
''               '.TableCell(tcFontBold, i, 4) = True: .TableCell(tcText, i, 4) = Format(0, fg_Pict(6, 2)) & " %"
''               If totgrl > 0 Then .TableCell(tcFontBold, i, 8) = True: .TableCell(tcText, i, 8) = Format((totcur / totgrl) * 100, fg_Pict(6, 2)) & " %"
''               indcur = indcur + 1: totcur = 0
''               i = i + 2
''               '.TableCell(tcFontBold, i, 1) = True: .TableCell(tcText, i, 1) = IIf(indcur = 2, "Curva B", "Curva C")
''               i = i + 1
''            End If
''
''           .TableCell(tcText, i, 1) = Trim(RS2!pro_codigo)
''           .TableCell(tcText, i, 2) = Trim(RS2!pro_nombre)
''           .TableCell(tcText, i, 3) = Trim(RS2!uni_nomcor)
''           .TableCell(tcText, i, 4) = "__________________"
''           .TableCell(tcText, i, 8) = Format(0, fg_Pict(6, 2)) & " %"
''           If totgrl > 0 Then .TableCell(tcText, i, 8) = Format(((RS2!cantidad / totgrl) * 100), fg_Pict(3, 2)) & " %"
''           totcur = totcur + IIf(IsNull(RS2!cantidad), 0, RS2!cantidad)
''           RS2.MoveNext: i = i + 1
''        Loop
''    End If
''
''    '-------> Imprimir total curva
''    i = i + 2
''    .TableCell(tcColSpan, i, 5) = 2
''    '.TableCell(tcFontBold, i, 2) = True: .TableCell(tcText, i, 2) = IIf(indcur = 1, "Total General Curva A ", IIf(indcur = 2, "Total General Curva B", "Total General Curva C"))
''    '.TableCell(tcFontBold, i, 4) = True: .TableCell(tcText, i, 4) = Format(totcur, fg_Pict(6, 2))
''    .TableCell(tcFontBold, i, 4) = True
''    .TableCell(tcText, i, 8) = Format(0, fg_Pict(6, 2)) & " %"
''    If totgrl > 0 Then .TableCell(tcText, i, 8) = Format((totcur / totgrl) * 100, fg_Pict(6, 2)) & " %"
''
''    '-------> Imprimir total general servicio
''    i = i + 2
''    .TableCell(tcColSpan, i, 5) = 2
''    '.TableCell(tcFontBold, i, 2) = True: .TableCell(tcText, i, 2) = "Total General Servicio "
''    '.TableCell(tcFontBold, i, 4) = True: .TableCell(tcText, i, 4) = Format(totgrl, fg_Pict(6, 2))
''    .TableCell(tcRows) = i
''    .TableBorder = tbNone
''    .EndTable
''    RS2.Close: Set RS2 = Nothing
''    .EndDoc
''End With
''Preview.Show 1
''Exit Function
''
''Error_Productos:
''    fg_descarga
''    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
''    Close #1
''    Exit Function
''End Function
''
