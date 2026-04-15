Attribute VB_Name = "InforAN"
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset
Dim RS5 As New ADODB.Recordset
Dim RS6 As New ADODB.Recordset
Global vg_lineas As Long
Dim numlin As Integer
Dim inf_ncasino As String, inf_nregimen As String, inf_detaporte As String, inf_ndia As String
Dim inf_nreceta As String, cdetalle As String, opcionsalto As String
Dim inf_opcion As Integer

Public Function I_CatDie()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_CatDie
MsgTitulo = "Informe de Categorias Dieteticas"
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Categoría Dietética"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_recetacatdie where car_previo=0 order by car_nombre", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Categoría"
    .TableCell(tcText, 1, 2) = "Sub Categoría"
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
        .TableCell(tcText, i, 1) = RS1!car_nombre
        .TableCell(tcText, i, 2) = ""
        RS2.Open "select * from a_recetacatdie where car_previo=" & RS1!car_codigo & " order by car_nombre", vg_db, adOpenStatic
        If Not RS2.EOF Then
            Print #1, .TableCell(tcText, i, 1)
            Do While Not RS2.EOF
                .TableCell(tcText, i, 2) = RS2!car_nombre
                Print #1, vbTab & .TableCell(tcText, i, 2)
                RS2.MoveNext: i = i + 1
            Loop
        Else
            Print #1, .TableCell(tcText, i, 1)
            i = i + 1
        End If
        RS2.Close: Set RS2 = Nothing
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Function
Error_CatDie:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_TipPla()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_TipoPla
MsgTitulo = "Informe de Tipos de Plato"
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Tipo de Plato"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_recetatippla where tip_previo=0 order by tip_nombre", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Categoría"
    .TableCell(tcText, 1, 2) = "Sub Categoría"
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
        Print #1, .TableCell(tcText, i, 1)
        RS2.Open "select * from a_recetatippla where tip_previo=" & RS1!tip_codigo & " order by tip_nombre", vg_db, adOpenStatic
        If Not RS2.EOF Then
            Do While Not RS2.EOF
                .TableCell(tcText, i, 2) = RS2!tip_nombre
                Print #1, vbTab & .TableCell(tcText, i, 2)
                RS2.MoveNext: i = i + 1
            Loop
        Else
            Print #1, .TableCell(tcText, i, 1)
            i = i + 1
        End If
        RS2.Close: Set RS2 = Nothing
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Function
Error_TipoPla:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Regime()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Regimen
MsgTitulo = "Informe de Regimen"
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Régimen"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_regimen order by reg_codigo", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 5500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taLeftTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = RS1!reg_codigo
            .TableCell(tcText, i, 2) = RS1!reg_nombre
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2)
            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Function
Error_Regimen:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Servic()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Servicios
MsgTitulo = "Informe de Servicios"
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Servicio"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_servicio order by ser_codigo", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Orden"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = RS1!ser_codigo
            .TableCell(tcText, i, 2) = RS1!ser_nombre
            .TableCell(tcText, i, 3) = RS1!ser_orden
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Function
Error_Servicios:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Bodega()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Bodega
MsgTitulo = "Informe de Bodegas"
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Bodegas"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_bodega order by bod_codigo", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Ubicación"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = RS1!bod_codigo
            .TableCell(tcText, i, 2) = RS1!bod_nombre
            .TableCell(tcText, i, 3) = RS1!bod_ubicac
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Function
Error_Bodega:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Provee()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Proveedor
MsgTitulo = "Informe de Proveedores"
Preview.Show
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orLandscape 'orPortrait
    .MarginLeft = 500
    .StartDoc
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
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Maestro Proveedores"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from b_proveedor order by prv_codigo", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
    .StartTable
    .TableCell(tcCols) = 11: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1400: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1800: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1800: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1300: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 1300: .TableCell(tcAlign, , 5) = taLeftTop
    .TableCell(tcColWidth, , 6) = 900: .TableCell(tcAlign, , 6) = taLeftTop
    .TableCell(tcColWidth, , 7) = 900: .TableCell(tcAlign, , 7) = taLeftTop
    .TableCell(tcColWidth, , 8) = 900: .TableCell(tcAlign, , 8) = taLeftTop
    .TableCell(tcColWidth, , 9) = 1500: .TableCell(tcAlign, , 9) = taLeftTop
    .TableCell(tcColWidth, , 10) = 1500: .TableCell(tcAlign, , 10) = taLeftTop
    .TableCell(tcColWidth, , 11) = 1500: .TableCell(tcAlign, , 11) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Dirección"
    .TableCell(tcText, 1, 4) = "Comuna"
    .TableCell(tcText, 1, 5) = "Ciudad"
    .TableCell(tcText, 1, 6) = "Fono 1"
    .TableCell(tcText, 1, 7) = "Fono 2"
    .TableCell(tcText, 1, 8) = "Fax"
    .TableCell(tcText, 1, 9) = "Contacto"
    .TableCell(tcText, 1, 10) = "Giro"
    .TableCell(tcText, 1, 11) = "E-Mail"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5) & "|" & _
              .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & .TableCell(tcText, 1, 9) & "|" & .TableCell(tcText, 1, 10) & "|" & .TableCell(tcText, 1, 11) & "|"
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 11: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1400: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1800: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1800: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1300: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 1300: .TableCell(tcAlign, , 5) = taLeftTop
    .TableCell(tcColWidth, , 6) = 900: .TableCell(tcAlign, , 6) = taLeftTop
    .TableCell(tcColWidth, , 7) = 900: .TableCell(tcAlign, , 7) = taLeftTop
    .TableCell(tcColWidth, , 8) = 900: .TableCell(tcAlign, , 8) = taLeftTop
    .TableCell(tcColWidth, , 9) = 1500: .TableCell(tcAlign, , 9) = taLeftTop
    .TableCell(tcColWidth, , 10) = 1500: .TableCell(tcAlign, , 10) = taLeftTop
    .TableCell(tcColWidth, , 11) = 1500: .TableCell(tcAlign, , 11) = taLeftTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 2) = RS1!prv_nombre
            .TableCell(tcText, i, 1) = fg_PintaRut(RS1!prv_codigo)
            .TableCell(tcText, i, 3) = RS1!prv_direccion
            .TableCell(tcText, i, 4) = RS1!prv_comuna
            .TableCell(tcText, i, 5) = RS1!prv_ciudad
            .TableCell(tcText, i, 6) = RS1!prv_fono1
            .TableCell(tcText, i, 7) = RS1!prv_fono2
            .TableCell(tcText, i, 8) = RS1!prv_fax
            .TableCell(tcText, i, 9) = RS1!prv_percon
            .TableCell(tcText, i, 10) = RS1!prv_giro
            .TableCell(tcText, i, 11) = RS1!prv_emapro
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5) & "|" & _
                      .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9) & "|" & .TableCell(tcText, i, 10) & "|" & .TableCell(tcText, i, 11) & "|"
            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcRows) = i
    .PenColor = &HC0C0C0
    .TableBorder = tbAll
    .EndTable
    .EndDoc
    Close #1
End With
Exit Function
Error_Proveedor:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_TarjetaRecetas(cuenta As Long)
Dim i As Long, fil As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim codrec As Integer, nomrec As String, NomFan As String, lc_codrec As Long
Dim canser As Double, cannet As Double, totcanser As Double, totcannet As Double, totcosto As Double, canpro As Double
On Local Error GoTo Error_Tarjeta
MsgTitulo = "Informe de Recetas"
fg_carga ""
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
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10000: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 10: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Tarjeta Recetas"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .FontBold = False
    .FontSize = 8
    LogoEmp
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "SELECT pro.pro_nombre, pro.pro_propon, red_codigo, red.red_pctnut, red.red_pctapr, red.red_pctcoc, " & _
             "red.red_canpro, red.red_cospro FROM b_productos pro ,b_recetadet red " & _
             "where pro.pro_codigo = red.red_codpro order by red.red_codigo, red.red_nroite", vg_db, adOpenStatic
    I_Receta.ProgressBar1.Scrolling = ccScrollingStandard
    I_Receta.ProgressBar1.Max = cuenta
    I_Receta.ProgressBar1.Visible = True
    I_Receta.ProgressBar1.Value = 0
    For i = 1 To I_Receta.vaSpread1.MaxRows
        I_Receta.vaSpread1.Col = 1: I_Receta.vaSpread1.Row = i
        If I_Receta.vaSpread1.Value = "1" Then
            I_Receta.vaSpread1.Col = 2: codrec = I_Receta.vaSpread1.Text
            If RS1.RecordCount > 0 Then RS1.MoveFirst
            RS1.Find "red_codigo=" & codrec, , adSearchForward
            If Not RS1.EOF Then
                lc_codrec = RS1!red_codigo
                .Text = Chr(13): .Text = Chr(13)
                I_Receta.vaSpread1.Col = 3: nomrec = I_Receta.vaSpread1.Text
                I_Receta.vaSpread1.Col = 4: NomFan = I_Receta.vaSpread1.Text
                'CATEGORIA DIETETICA,TIPO PLATO Y TIPO RACIONES
                .StartTable
                .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
                .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcText, 1, 1) = "Cat. Dietetica"
                .TableCell(tcText, 2, 1) = "Tipo Plato"
                .TableCell(tcText, 3, 1) = "Nro. Raciones"
                I_Receta.vaSpread1.Col = 5: .TableCell(tcText, 1, 2) = I_Receta.vaSpread1.Text
                I_Receta.vaSpread1.Col = 6: .TableCell(tcText, 2, 2) = I_Receta.vaSpread1.Text
                I_Receta.vaSpread1.Col = 7: .TableCell(tcText, 3, 2) = I_Receta.vaSpread1.Text
                'RS1.Open "SELECT car.car_nombre FROM a_recetacatdie car INNER JOIN b_receta rec ON " & _
                         "car.car_codigo = rec.rec_catdie where rec.rec_codigo=" & codrec, vg_db, adOpenStatic
                Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
                Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
                Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
                .TableBorder = tbNone
                .EndTable
                'NOMBRE RECETA
                .StartTable
                .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
                .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
                .TableCell(tcFontSize, 1) = 9: .TableCell(tcFontBold, 1) = True
                .TableCell(tcText, 1, 1) = "* " & nomrec & " *"
                .TableCell(tcText, 1, 1) = IIf(I_Receta.Option1(0).Value = True, "* " & nomrec & " *", "* " & NomFan & " *")
                Print #1, .TableCell(tcText, 1, 1)
                .TableBorder = tbNone
                .EndTable
                .Text = Chr(13)
                .StartTable
                .TableCell(tcCols) = 8: .TableCell(tcRows) = 1
                .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcColWidth, , 2) = 1000: .TableCell(tcAlign, , 2) = taRightTop
                .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taRightTop
                .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
                .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
                .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
                .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
                .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
                .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
                .TableCell(tcText, 1, 1) = "Nombre Producto"
                .TableCell(tcText, 1, 2) = "C.Bruta"
                .TableCell(tcText, 1, 3) = "%Aprov."
                .TableCell(tcText, 1, 4) = "%A.Coc."
                .TableCell(tcText, 1, 5) = "C.Servir"
                .TableCell(tcText, 1, 6) = "%P.Nut."
                .TableCell(tcText, 1, 7) = "C.Neta"
                .TableCell(tcText, 1, 8) = "Costo"
                Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
                           .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8)
                .TableBorder = tbAll
                .EndTable
                .StartTable
                .TableCell(tcCols) = 8: .TableCell(tcRows) = 150
                .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcColWidth, , 2) = 1000: .TableCell(tcAlign, , 2) = taRightTop
                .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taRightTop
                .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
                .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
                .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
                .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
                .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
                'RS1.Open "SELECT pro.pro_nombre, pro.pro_propon, red.red_pctnut, red.red_pctapr, red.red_pctcoc, red.red_canpro, red.red_cospro " & _
                         "FROM b_productos pro INNER JOIN b_recetadet red ON pro.pro_codigo = red.red_codpro where red.red_codigo=" & codrec & " order by red.red_nroite", vg_db, adOpenStatic
                fil = 0
                totcanser = 0
                totcannet = 0
                totcosto = 0
                Do While Not RS1.EOF And codrec = lc_codrec
                    If Not RS1.EOF Then lc_codrec = RS1!red_codigo Else lc_codrec = 0
                    If codrec = lc_codrec Then
                        fil = fil + 1
                        canpro = Format(RS1!red_canpro, fg_Pict(6, 5))
                        canser = Format((RS1!red_pctapr / 100) * canpro * (RS1!red_pctcoc / 100), fg_Pict(6, 5))
                        cannet = Format((RS1!red_pctnut / 100) * canpro, fg_Pict(6, 5))
                        .TableCell(tcText, fil, 1) = RS1!pro_nombre
                        .TableCell(tcText, fil, 2) = canpro
                        .TableCell(tcText, fil, 3) = RS1!red_pctapr
                        .TableCell(tcText, fil, 4) = RS1!red_pctcoc
                        .TableCell(tcText, fil, 5) = canser
                        .TableCell(tcText, fil, 6) = RS1!red_pctnut
                        .TableCell(tcText, fil, 7) = cannet
                        .TableCell(tcText, fil, 8) = Format(canpro * RS1!pro_propon, fg_Pict(6, vg_DCa))
                         Print #1, .TableCell(tcText, fil, 1) & "|" & .TableCell(tcText, fil, 2) & "|" & .TableCell(tcText, fil, 3) & "|" & .TableCell(tcText, fil, 4) & "|" & _
                                   .TableCell(tcText, fil, 5) & "|" & .TableCell(tcText, fil, 6) & "|" & .TableCell(tcText, fil, 7) & "|" & .TableCell(tcText, fil, 8)
                        totcanser = totcanser + canser
                        totcannet = totcannet + cannet
                        totcosto = Format(totcosto + canpro * RS1!pro_propon, fg_Pict(6, vg_DCa))
                    End If
                    RS1.MoveNext
                Loop
                .TableCell(tcRows) = fil
                .PenColor = &HC0C0C0
                .TableBorder = tbAll
                .EndTable
                .StartTable
                .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
                .TableCell(tcColWidth, , 1) = 6500
                .TableCell(tcColWidth, , 2) = 1000
                .TableCell(tcColWidth, , 3) = 2000
                .TableCell(tcColWidth, , 4) = 1000
                .TableCell(tcFontBold, 1) = True
                .TableCell(tcText, 1, 1) = "": .TableCell(tcAlign, , 1) = taRightTop
                .TableCell(tcText, 1, 2) = totcanser: .TableCell(tcAlign, , 2) = taRightTop
                .TableCell(tcText, 1, 3) = totcannet: .TableCell(tcAlign, , 3) = taRightTop
                .TableCell(tcText, 1, 4) = Format(totcosto, fg_Pict(6, vg_DCa)): .TableCell(tcAlign, , 4) = taRightTop
                Print #1, "|" & "|" & "|"; .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "||" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
                .TableBorder = tbNone
                .EndTable
                If vg_lineas > 200 And i <> I_Receta.vaSpread1.MaxRows Then
                    .NewPage
                End If
            End If
            I_Receta.ProgressBar1.Value = I_Receta.ProgressBar1.Value + 1
        End If
    Next i
    RS1.Close: Set RS1 = Nothing
    .EndDoc
    I_Receta.ProgressBar1.Value = I_Receta.ProgressBar1.Max
    I_Receta.ProgressBar1.Visible = False
    Close #1
End With
Preview.Show 1
fg_descarga
Exit Function
Error_Tarjeta:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_AporteRecetas(cuenta As Long)
Dim i As Long, X As Long, fil As Long, Col As Long, VecCol(100) As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim codrec As Integer, nomrec As String, NomFan As String, vector(100) As Double
Dim canser As Double, cannet As Double, totcanser As Double, totcannet As Double, totcosto As Double, canpro As Double
Dim alto As Long, aAp As String, lc_codrec As Long, cantidad As Double, lc_codpro As String
Dim cLin As String, J As Long, icol As Long, totLineas As Long
On Local Error GoTo Error_ApoRecet
MsgTitulo = "Informe de Recetas"
fg_carga ""
aAp = Trim(vg_NUsr) & "_tmp_imprec"
alto = 0
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    If I_Receta.List1.SelCount > 7 Then .Zoom = 125
    .PreviewPage = 1
    .Orientation = IIf(I_Receta.List1.SelCount > 7, orLandscape, orPortrait)
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 10
    .HdrFontBold = True
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = IIf(I_Receta.List1.SelCount > 7, 13500, 10500): .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 10: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Informe Recetas Aporte Nutricional"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .FontSize = 7
    .Footer = "||Página : %d"
    'LogoEmp
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "SELECT red.red_codigo, ing.ing_nombre, red.red_pctnut, red.red_pctapr, " & _
             "red.red_pctcoc, red.red_canpro, red.red_codpro, rec.rec_basrac, ing.ing_facnut " & _
             "FROM b_ingrediente ing, b_recetadet red, b_receta rec " & _
             "where ing.ing_codigo=red.red_codpro and rec.rec_codigo=red.red_codigo order by red.red_codigo, red.red_nroite", vg_db, adOpenStatic
    RS2.Open "select pnu_codpro, pnu_codapo, pnu_canapo from b_productonut", vg_db, adOpenStatic
    I_Receta.ProgressBar1.Scrolling = ccScrollingStandard
    I_Receta.ProgressBar1.Max = cuenta
    I_Receta.ProgressBar1.Visible = True
    I_Receta.ProgressBar1.Value = 0
    totLineas = 0
    For i = 1 To I_Receta.vaSpread1.MaxRows
        I_Receta.vaSpread1.Col = 1: I_Receta.vaSpread1.Row = i
        If I_Receta.vaSpread1.Value = "1" Then
            I_Receta.vaSpread1.Col = 2: codrec = I_Receta.vaSpread1.Text
            If RS1.RecordCount > 0 Then RS1.MoveFirst
            RS1.Find "red_codigo=" & codrec, , adSearchForward
            If Not RS1.EOF Then
                lc_codrec = RS1!red_codigo
                For X = 1 To 100
                    vector(X) = 0
                Next X
                .Text = Chr(13): .Text = Chr(13)
                I_Receta.vaSpread1.Col = 3: nomrec = I_Receta.vaSpread1.Text
                I_Receta.vaSpread1.Col = 4: NomFan = I_Receta.vaSpread1.Text
                'CATEGORIA DIETETICA,TIPO PLATO Y TIPO RACIONES
                .StartTable
                .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
                .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcText, 1, 1) = "Cat. Dietetica"
                .TableCell(tcText, 2, 1) = "Tipo Plato"
                .TableCell(tcText, 3, 1) = "Nro. Raciones"
                I_Receta.vaSpread1.Col = 5: .TableCell(tcText, 1, 2) = I_Receta.vaSpread1.Text
                I_Receta.vaSpread1.Col = 6: .TableCell(tcText, 2, 2) = I_Receta.vaSpread1.Text
                I_Receta.vaSpread1.Col = 7: .TableCell(tcText, 3, 2) = I_Receta.vaSpread1.Text
                Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
                Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
                Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
                .TableBorder = tbNone
                totLineas = totLineas + .TableCell(tcRows)
                .EndTable
                'NOMBRE RECETA
                .StartTable
                .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
                .TableCell(tcColWidth, , 1) = IIf(I_Receta.List1.SelCount > 7, 13500, 10500): .TableCell(tcAlign, , 1) = taCenterMiddle
                .TableCell(tcFontSize, 1) = 9: .TableCell(tcFontBold, 1) = True
                .TableCell(tcText, 1, 1) = IIf(I_Receta.Option1(0).Value = True, "* " & nomrec & " *", "* " & NomFan & " *")
                Print #1, .TableCell(tcText, 1, 1)
                .TableBorder = tbNone
                totLineas = totLineas + .TableCell(tcRows)
                .EndTable
                .Text = Chr(13)
                .StartTable
                .TableCell(tcCols) = 5 + I_Receta.List1.SelCount: .TableCell(tcRows) = 1
                .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
                .TableCell(tcColWidth, , 3) = 700: .TableCell(tcAlign, , 3) = taRightTop
                .TableCell(tcColWidth, , 4) = 700: .TableCell(tcAlign, , 4) = taRightTop
                .TableCell(tcColWidth, , 5) = 700: .TableCell(tcAlign, , 5) = taRightTop
                .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True:  .TableCell(tcRowHeight, 1) = 230
                .TableCell(tcText, 1, 1) = "Código"
                .TableCell(tcText, 1, 2) = "Nombre Ingrediente"
                .TableCell(tcText, 1, 3) = "C.Bruta"
                .TableCell(tcText, 1, 4) = "C.Servir"
                .TableCell(tcText, 1, 5) = "C.Neta"
                cLin = cLin & .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5)
                Col = 6
                'cLin = ""
                For X = 0 To I_Receta.List1.ListCount - 1
                    If I_Receta.List1.Selected(X) = True Then
                        .TableCell(tcColWidth, , Col) = 800: .TableCell(tcAlign, , Col) = taRightTop
                        .TableCell(tcText, 1, Col) = I_Receta.List1.List(X)
                        cLin = cLin & "|" & .TableCell(tcText, 1, Col)
                        VecCol(I_Receta.List1.ItemData(X)) = Col
                        Col = Col + 1
                    End If
                Next X
                icol = Col - 1
                Print #1, cLin
                cLin = ""
                .TableBorder = tbAll
                totLineas = totLineas + .TableCell(tcRows)
                .EndTable
                cLin = ""
                .StartTable
                .TableCell(tcCols) = 5 + I_Receta.List1.SelCount: .TableCell(tcRows) = 150
                .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
                .TableCell(tcColWidth, , 3) = 700: .TableCell(tcAlign, , 3) = taRightTop
                .TableCell(tcColWidth, , 4) = 700: .TableCell(tcAlign, , 4) = taRightTop
                .TableCell(tcColWidth, , 5) = 700: .TableCell(tcAlign, , 5) = taRightTop
                .TableCell(tcText, 1, 1, 1, 5 + I_Receta.List1.SelCount) = ""
                totcanser = 0
                totcannet = 0
                totcanpro = 0
                fil = 0
                Do While Not RS1.EOF And codrec = lc_codrec
                    If Not RS1.EOF Then lc_codrec = RS1!red_codigo Else lc_codrec = 0
                    If codrec = lc_codrec Then
                        fil = fil + 1
                        .TableCell(tcText, fil, 6, fil, I_Receta.List1.SelCount + 6) = Format(0, fg_Pict(9, vg_DCa))
                        canpro = Format(RS1!red_canpro, fg_Pict(6, 5))
                        canser = Format((RS1!red_pctapr / 100) * canpro * (RS1!red_pctcoc / 100), fg_Pict(6, 5))
                        cannet = Format((RS1!red_pctnut / 100) * canpro, fg_Pict(6, 5))
                        .TableCell(tcText, fil, 1) = RS1!red_codpro
                        .TableCell(tcText, fil, 2) = RS1!ing_nombre
                        .TableCell(tcText, fil, 3) = canpro
                        .TableCell(tcText, fil, 4) = canser
                        .TableCell(tcText, fil, 5) = cannet
                        totcanpro = totcanpro + canpro
                        totcanser = totcanser + canser
                        totcannet = totcannet + cannet
                        Col = 6
                        If RS2.RecordCount > 0 Then RS2.MoveFirst
                        RS2.Find "pnu_codpro='" & RS1!red_codpro & "'", , adSearchForward
                        If Not RS2.EOF Then
                            lc_codpro = RS2!pnu_codpro
                            Do While Not RS2.EOF And lc_codpro = RS1!red_codpro
                                If Not RS2.EOF Then lc_codpro = RS2!pnu_codpro Else lc_codpro = 0
                                If lc_codpro = RS1!red_codpro Then
                                    cantidad = (((RS1!red_pctnut / 100) * (RS2!pnu_canapo * (RS1!red_canpro / RS1!rec_basrac))) / RS1!ing_facnut)
                                    .TableCell(tcText, fil, VecCol(Val(RS2!pnu_codapo))) = Format(cantidad, fg_Pict(6, vg_DCa))
                                    vector(Val(RS2!pnu_codapo)) = vector(Val(RS2!pnu_codapo)) + Format(cantidad, fg_Pict(6, vg_DCa))
                                    'cLin = cLin & "|" & .TableCell(tcText, fil, VecCol(Val(RS2!pnu_codapo)))
                                    Col = Col + 1
                                End If
                                RS2.MoveNext
                            Loop
                        Else
                            vector(Col) = vector(Col) + Format(0, fg_Pict(6, vg_DCa))
                            Col = Col + 1
                        End If
                    End If
                    RS1.MoveNext
                Loop
                For H = 1 To fil
                    cLin = ""
                    For J = 1 To icol
                        cLin = cLin & .TableCell(tcText, H, J) & "|"
                    Next J
                    Print #1, cLin
                Next H
                cLin = ""
                .TableCell(tcColWidth, 1, 6, fil, 5 + I_Receta.List1.SelCount) = 800: .TableCell(tcAlign, 1, 6, fil, 5 + I_Receta.List1.SelCount) = taRightTop
                .TableCell(tcRows) = fil
                .PenColor = &HC0C0C0
                .TableBorder = tbAll
                totLineas = totLineas + .TableCell(tcRows)
                .EndTable
                cLin = ""
                .StartTable
                .TableCell(tcCols) = 4 + I_Receta.List1.SelCount: .TableCell(tcRows) = 1
                .TableCell(tcColWidth, , 1) = 3500
                .TableCell(tcColWidth, , 2) = 700
                .TableCell(tcColWidth, , 3) = 700
                .TableCell(tcColWidth, , 4) = 700
                .TableCell(tcFontBold, 1) = True
                .TableCell(tcText, 1, 1) = "Totales" & Space(29): .TableCell(tcAlign, , 1) = taRightTop
                .TableCell(tcText, 1, 2) = totcanpro: .TableCell(tcAlign, , 2) = taRightTop
                .TableCell(tcText, 1, 3) = totcanser: .TableCell(tcAlign, , 3) = taRightTop
                .TableCell(tcText, 1, 4) = totcannet: .TableCell(tcAlign, , 4) = taRightTop
                cLin = cLin & "|" & .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) '& "|" & .TableCell(tcText, 1, 5)
                Col = 5
                For X = 0 To I_Receta.List1.ListCount - 1
                    If I_Receta.List1.Selected(X) = True Then
                        .TableCell(tcColWidth, , Col) = 800: .TableCell(tcAlign, , Col) = taRightTop
                        .TableCell(tcText, 1, Col) = Format(vector(I_Receta.List1.ItemData(X)), fg_Pict(6, 2))
                        cLin = cLin & "|" & .TableCell(tcText, 1, Col)
                        Col = Col + 1
                    End If
                Next X
                Print #1, cLin
                Print #1, " "
                cLin = ""
                .TableBorder = tbNone
                totLineas = totLineas + .TableCell(tcRows)
                .EndTable
                Dim varx As Integer
                varx = IIf(I_Receta.List1.SelCount > 7, 30, 40)
                If totLineas > varx And i <> I_Receta.vaSpread1.MaxRows Then
                    totLineas = 0
                    .NewPage
                End If
            End If
            I_Receta.ProgressBar1.Value = I_Receta.ProgressBar1.Value + 1
        End If
    Next i
    RS1.Close: Set RS1 = Nothing
    RS2.Close: Set RS2 = Nothing
    .EndDoc
    I_Receta.ProgressBar1.Value = I_Receta.ProgressBar1.Max
    I_Receta.ProgressBar1.Visible = False
    Close #1
End With
Preview.Show 1
fg_descarga
Exit Function
Error_ApoRecet:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_NombreRecetas(cuenta As Long)
Dim i As Long, X As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim codrec As Integer, nomrec As String, NomFan As String, catdie As String, tippla As String
On Local Error GoTo Error_NomRec
MsgTitulo = "Informe de Recetas"
fg_carga ""
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
    .FontSize = 7
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Nombre Recetas"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Categoría Dietética"
    .TableCell(tcText, 1, 2) = Trim(M_Receta.Label2(8).Caption)
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 5000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Nombre"
    .TableCell(tcText, 1, 3) = "Tipo Plato"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    .TableBorder = tbBox
    .EndTable
    
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 5000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    X = 0
    I_Receta.ProgressBar1.Scrolling = ccScrollingSmooth
    I_Receta.ProgressBar1.Max = cuenta
    I_Receta.ProgressBar1.Visible = True
    I_Receta.ProgressBar1.Value = 0
    For i = 1 To I_Receta.vaSpread1.MaxRows
        I_Receta.vaSpread1.Row = i: I_Receta.vaSpread1.Col = 1
        If I_Receta.vaSpread1.Text = "1" Then
            X = X + 1
            I_Receta.vaSpread1.Col = 2: codrec = I_Receta.vaSpread1.Text
            I_Receta.vaSpread1.Col = 3: nomrec = I_Receta.vaSpread1.Text
            I_Receta.vaSpread1.Col = 4: NomFan = I_Receta.vaSpread1.Text
            I_Receta.vaSpread1.Col = 6: tippla = I_Receta.vaSpread1.Text
            .TableCell(tcText, X, 1) = codrec
            .TableCell(tcText, X, 2) = IIf(I_Receta.Option1(0).Value = True, nomrec, NomFan)
            .TableCell(tcText, X, 3) = tippla
            Print #1, .TableCell(tcText, X, 1) & "|" & .TableCell(tcText, X, 2) & "|" & .TableCell(tcText, X, 3)
            I_Receta.ProgressBar1.Value = I_Receta.ProgressBar1.Value + 1
        End If
    Next i
    .PenColor = &HC0C0C0
    .TableBorder = tbAll
    .TableCell(tcRows) = X
    .EndTable
    .EndDoc
    Close #1
    I_Receta.ProgressBar1.Value = I_Receta.ProgressBar1.Max
    I_Receta.ProgressBar1.Visible = False
End With
Preview.Show 1
fg_descarga
Exit Function
Error_NomRec:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Productos()
Dim i As Long, X As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long, codigo As String
On Local Error GoTo Error_Productos
MsgTitulo = "Informe de Productos"
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orLandscape
    .MarginLeft = 500
    .Zoom = 110
    .StartDoc
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
    .FontSize = 7
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 13500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Productos"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .TextAlign = taCenterTop
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 9: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4200: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2800: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 650: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 650: .TableCell(tcAlign, , 5) = taCenterTop
    .TableCell(tcColWidth, , 6) = 1050: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1050: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1050: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1050: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = IIf(I_Produc.Option1(0).Value = True, "Nombre", "Nombre Fantasía")
    .TableCell(tcText, 1, 3) = "Familia"
    .TableCell(tcText, 1, 4) = "Uni.Env"
    .TableCell(tcText, 1, 5) = "Uni.Emb"
    .TableCell(tcText, 1, 6) = "Cant.xUni."
    .TableCell(tcText, 1, 7) = "Ult.Precio"
    .TableCell(tcText, 1, 8) = "Fec.Ult.Comp."
    .TableCell(tcText, 1, 9) = "P.M.P."
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & .TableCell(tcText, 1, 9)
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 9: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4200: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2800: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 650: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 650: .TableCell(tcAlign, , 5) = taCenterTop
    .TableCell(tcColWidth, , 6) = 1050: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1050: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1050: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1050: .TableCell(tcAlign, , 9) = taRightTop
    X = 1
    For i = 1 To I_Produc.vaSpread1.MaxRows
        I_Produc.vaSpread1.Row = i: I_Produc.vaSpread1.Col = 1
        If I_Produc.vaSpread1.Text = "1" Then
            I_Produc.vaSpread1.Col = 2: codigo = I_Produc.vaSpread1.Text
            RS1.Open "SELECT tip.tip_codigo,tip.tip_nombre, uni.uni_nomcor, emb.emb_nomcor, pro.* FROM a_embalaje emb INNER JOIN (a_unidad uni INNER JOIN (a_tipopro tip INNER JOIN b_productos pro ON tip.tip_codigo=pro.pro_codtip) ON uni.uni_codigo=pro.pro_coduni) ON emb.emb_codigo=pro.pro_codemb where pro.pro_codigo='" & Trim(codigo) & "' order by pro.pro_codigo", vg_db, adOpenStatic
            If Not RS1.EOF Then
                Do While Not RS1.EOF
                    .TableCell(tcText, X, 1) = RS1!pro_codigo
                    .TableCell(tcText, X, 2) = RS1!pro_nombre
                    I_Produc.vaSpread1.Col = 4
                    .TableCell(tcText, X, 3) = I_Produc.vaSpread1.Value 'fg_BuscaenArbol(RS1!tip_codigo, "a_tipopro", "tip_codigo")    'M_Produc.fpayuda(0).Caption 'RS1!tip_nombre
                    .TableCell(tcText, X, 4) = RS1!uni_nomcor
                    .TableCell(tcText, X, 5) = RS1!emb_nomcor
                    .TableCell(tcText, X, 6) = RS1!pro_uniemb
                    .TableCell(tcText, X, 7) = IIf(IsNull(RS1!pro_upreco), 0, Format(RS1!pro_upreco, fg_Pict(9, vg_DPr)))
                    .TableCell(tcText, X, 8) = IIf(IsNull(RS1!pro_fecuco), "", RS1!pro_fecuco)
                    .TableCell(tcText, X, 9) = IIf(IsNull(RS1!pro_propon), 0, Format(RS1!pro_propon, fg_Pict(9, vg_DPr)))
                    Print #1, .TableCell(tcText, X, 1) & "|" & .TableCell(tcText, X, 2) & "|" & .TableCell(tcText, X, 3) & "|" & .TableCell(tcText, X, 4) & "|" & _
                              .TableCell(tcText, X, 5) & "|" & .TableCell(tcText, X, 6) & "|" & .TableCell(tcText, X, 7) & "|" & "'" & .TableCell(tcText, X, 8) & "|" & .TableCell(tcText, X, 9)
                    RS1.MoveNext: X = X + 1
                Loop
            Else
                X = X + 1
            End If
            RS1.Close: Set RS1 = Nothing
        End If
    Next i
    .TableCell(tcRows) = X - 1
    .PenColor = &HC0C0C0
    .TableBorder = tbAll
    .EndTable
    .EndDoc
End With
Preview.Show 1
Close #1
Exit Function
Error_Productos:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Ingrediente()
Dim i As Long, X As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long, codigo As String
On Local Error GoTo Error_Productos
MsgTitulo = "Informe de Ingredientes"
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = orLandscape
    .MarginLeft = 500
    .Zoom = 110
    .StartDoc
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 8
    .HdrFontBold = False
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .Footer = "||Página : %d"
    .FontSize = 7
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 13500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Ingredientes"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .TextAlign = taCenterTop
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 11: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4200: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 650: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 800: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 800: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 800: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 800: .TableCell(tcAlign, , 8) = taCenterTop
    .TableCell(tcColWidth, , 9) = 800: .TableCell(tcAlign, , 9) = taCenterTop
    .TableCell(tcColWidth, , 10) = 1100: .TableCell(tcAlign, , 10) = taRightTop
    .TableCell(tcColWidth, , 11) = 1100: .TableCell(tcAlign, , 11) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = IIf(I_Produc.Option1(0).Value = True, "Nombre", "Nombre Fantasía")
    .TableCell(tcText, 1, 3) = "Uni.Med."
    .TableCell(tcText, 1, 4) = "%Aprov."
    .TableCell(tcText, 1, 5) = "%Coc."
    .TableCell(tcText, 1, 6) = "%Aprov.Nut."
    .TableCell(tcText, 1, 7) = "Fac.Nut."
    .TableCell(tcText, 1, 8) = "P.A.V.B."
    .TableCell(tcText, 1, 9) = "I.Gr.Verd."
    .TableCell(tcText, 1, 10) = "Fec.Ult.Com."
    .TableCell(tcText, 1, 11) = "PMP"
    
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & _
              .TableCell(tcText, 1, 9) & "|" & .TableCell(tcText, 1, 10) & "|" & .TableCell(tcText, 1, 11)
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 11: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4200: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 650: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 800: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 800: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 800: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 800: .TableCell(tcAlign, , 8) = taCenterTop
    .TableCell(tcColWidth, , 9) = 800: .TableCell(tcAlign, , 9) = taCenterTop
    .TableCell(tcColWidth, , 10) = 1100: .TableCell(tcAlign, , 10) = taRightTop
    .TableCell(tcColWidth, , 11) = 1100: .TableCell(tcAlign, , 11) = taRightTop
    X = 1
    For i = 1 To I_Produc.vaSpread1.MaxRows
        I_Produc.vaSpread1.Row = i: I_Produc.vaSpread1.Col = 1
        If I_Produc.vaSpread1.Text = "1" Then
            I_Produc.vaSpread1.Col = 2: codigo = I_Produc.vaSpread1.Text
            RS1.Open "SELECT ing.*, unm.unm_nomcor from a_unidadmed unm, b_ingrediente ing " & _
                     "where unm.unm_codigo=ing.ing_unimed and ing.ing_codigo='" & Trim(codigo) & "' order by ing.ing_codigo", vg_db, adOpenStatic
            If Not RS1.EOF Then
                Do While Not RS1.EOF
                    .TableCell(tcText, X, 1) = RS1!ing_codigo
                    .TableCell(tcText, X, 2) = IIf(I_Produc.Option1(0).Value = True, RS1!ing_nombre, RS1!ing_nomfan)
                    .TableCell(tcText, X, 3) = RS1!unm_nomcor
                    .TableCell(tcText, X, 4) = RS1!ing_pctapr
                    .TableCell(tcText, X, 5) = RS1!ing_pctcoc
                    .TableCell(tcText, X, 6) = RS1!ing_pctnut
                    .TableCell(tcText, X, 7) = RS1!ing_facnut
                    .TableCell(tcText, X, 8) = IIf(RS1!ing_indpav = 0, "", "x")
                    .TableCell(tcText, X, 9) = IIf(RS1!ing_indgrv = 0, "", "x")
                    .TableCell(tcText, X, 10) = IIf(RS1!ing_precos = 0, "", RS1!ing_precos)
                    .TableCell(tcText, X, 11) = RS1!ing_feccos
                    Print #1, .TableCell(tcText, X, 1) & "|" & .TableCell(tcText, X, 2) & "|" & .TableCell(tcText, X, 3) & "|" & .TableCell(tcText, X, 4) & "|" & _
                              .TableCell(tcText, X, 5) & "|" & .TableCell(tcText, X, 6) & "|" & .TableCell(tcText, X, 7) & "|" & "'" & .TableCell(tcText, X, 8) & "|" & _
                              .TableCell(tcText, X, 9) & "|" & .TableCell(tcText, X, 10) & "|" & .TableCell(tcText, X, 11)
                    RS1.MoveNext: X = X + 1
                Loop
            Else
                X = X + 1
            End If
            RS1.Close: Set RS1 = Nothing
        End If
    Next i
    .TableCell(tcRows) = X - 1
    .PenColor = &HC0C0C0
    .TableBorder = tbAll
    .EndTable
    .EndDoc
End With
Preview.Show 1
Close #1
Exit Function
Error_Productos:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_AporteProductos()
Dim i As Long, X As Long, fil As Long, Col As Long, VecCol(100) As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim CodPro As String, NomPro As String, NomFan As String
Dim canser As Double, cannet As Double, totcanser As Double, totcannet As Double, totcosto As Double, canpro As Double
Dim alto As Long, cLin As String
On Local Error GoTo Error_AporProduc
MsgTitulo = "Informe de Aporte Ingredientes"
alto = 0
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    If I_Produc.List1.SelCount > 9 Then .Zoom = 125
    .Orientation = IIf(I_Produc.List1.SelCount > 9, orLandscape, orPortrait)
    .MarginLeft = 500
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 10
    .HdrFontBold = True
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = IIf(I_Produc.List1.SelCount > 9, 13500, 10500): .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 10: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Informe Ingredientes Aporte Nutricional /100gr"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .FontSize = 7
    .Footer = "||Página : %d"
    LogoEmp
    .TextAlign = taLeftTop
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2 + I_Produc.List1.SelCount: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    cLin = cLin & .TableCell(tcText, 1, 1)
    .TableCell(tcText, 1, 2) = IIf(I_Produc.Option1(0).Value = True, "Nombre", "Nombre Fantasía")
    cLin = cLin & "|" & .TableCell(tcText, 1, 2)
    Col = 3
    For X = 0 To I_Produc.List1.ListCount - 1
        If I_Produc.List1.Selected(X) = True Then
            .TableCell(tcColWidth, , Col) = 800: .TableCell(tcAlign, , Col) = taRightTop
            .TableCell(tcText, 1, Col) = I_Produc.List1.List(X)
            cLin = cLin & "|" & .TableCell(tcText, 1, Col)
            VecCol(I_Produc.List1.ItemData(X)) = Col
            Col = Col + 1
        End If
    Next X
    Print #1, cLin
    .TableBorder = tbAll
    .EndTable
    .StartTable
    aAp = Trim(vg_NUsr) & "_tmp_imppro"
    RS1.Open "select * from " & aAp & " where tem_codpat='0'", vg_db, adOpenStatic
    .TableCell(tcCols) = 2 + I_Produc.List1.SelCount: .TableCell(tcRows) = 10000
    .TableCell(tcText, 2, 1, 800, 2 + I_Produc.List1.SelCount) = "0.00"
    cLin = ""
    fil = 1
    Do While Not RS1.EOF
        fil = fil + 1
        .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
        .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
        .TableCell(tcText, fil, 1) = RS1!tem_codigo
        cLin = cLin & .TableCell(tcText, fil, 1)
        .TableCell(tcText, fil, 2) = IIf(I_Produc.Option1(0).Value = True, RS1!tem_nombre, RS1!tem_nomfan)
        cLin = cLin & "|" & .TableCell(tcText, fil, 2)
        RS2.Open "select * from " & aAp & " where tem_codpat='" & Trim(RS1!tem_codigo) & "'", vg_db, adOpenStatic
        Do While Not RS2.EOF
            .TableCell(tcText, fil, VecCol(Val(RS2!tem_codigo))) = Format(RS2!tem_nombre, fg_Pict(6, 2))
            RS2.MoveNext
        Loop
        For i = 3 To Col - 1
            cLin = cLin & "|" & .TableCell(tcText, fil, i)
        Next i
        Print #1, cLin
        cLin = ""
        RS2.Close: Set RS2 = Nothing
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcColWidth, 1, 3, fil, Col) = 800: .TableCell(tcAlign, 1, 3, fil, Col) = taRightTop
    .TableCell(tcRows) = fil
    .PenColor = &HC0C0C0
    .TableBorder = tbAll
    .EndTable
    .EndDoc
End With
Preview.Show 1
Exit Function
Error_AporProduc:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_ImpuestoProductos()
Dim i As Long, X As Long, fil As Long, Col As Long, VecCol(100) As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim CodPro As String, NomPro As String, NomFan As String, aAp As String
Dim canser As Double, cannet As Double, totcanser As Double, totcannet As Double, totcosto As Double, canpro As Double
Dim alto As Long, cLin As String
On Local Error GoTo Error_Impuestos
MsgTitulo = "Informe de Impuestos Productos"
alto = 0
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
    .Preview = True
    .PreviewPage = 1
    .Orientation = IIf(I_Produc.List1.SelCount > 7, orLandscape, orPortrait)
    .MarginLeft = 500
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    .StartDoc
    .PageBorder = 0
    .HdrFontName = "Arial"
    .HdrFontSize = 10
    .HdrFontBold = True
    .Header = "||Fecha : " & Format(Date, "dd/mm/yyyy")
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = IIf(I_Produc.List1.SelCount > 7, 13500, 10500): .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 10: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Informe Productos Impuestos"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .FontSize = 8
    .Footer = "||Página : %d"
    LogoEmp
    .Text = Chr(13): .Text = Chr(13)
    .TextAlign = taCenterTop
    .StartTable
    .TableCell(tcCols) = 2 + I_Produc.List1.SelCount: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    cLin = cLin & .TableCell(tcText, 1, 1)
    .TableCell(tcText, 1, 2) = IIf(I_Produc.Option1(0).Value = True, "Nombre", "Nombre Fantasía")
    cLin = cLin & "|" & .TableCell(tcText, 1, 2)
    Col = 3
    cLin1 = ""
    For X = 0 To I_Produc.List1.ListCount - 1
        If I_Produc.List1.Selected(X) = True Then
            .TableCell(tcColWidth, , Col) = 1000: .TableCell(tcAlign, , Col) = taCenterTop
            .TableCell(tcText, 1, Col) = I_Produc.List1.List(X)
            cLin = cLin & "|" & .TableCell(tcText, 1, Col)
            VecCol(I_Produc.List1.ItemData(X)) = Col
            Col = Col + 1
            cLin = cLin
        End If
    Next X
    Print #1, cLin
    .TableBorder = tbAll
    .EndTable
    .StartTable
    aAp = Trim(vg_NUsr) & "_tmp_imppro"
    RS1.Open "select * from " & aAp & " where tem_codpat='0'", vg_db, adOpenStatic
    .TableCell(tcCols) = 2 + I_Produc.List1.SelCount: .TableCell(tcRows) = 10000
    cLin = ""
    fil = 1
    Do While Not RS1.EOF
        fil = fil + 1
        .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
        .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
        .TableCell(tcText, fil, 1) = RS1!tem_codigo
        .TableCell(tcText, fil, 2) = IIf(I_Produc.Option1(0).Value = True, RS1!tem_nombre, RS1!tem_nomfan)
        cLin = cLin & .TableCell(tcText, fil, 1) & "|" & .TableCell(tcText, fil, 2)
        RS2.Open "select * from " & aAp & " where tem_codpat='" & Trim(RS1!tem_codigo) & "'", vg_db, adOpenStatic
        Do While Not RS2.EOF
            .TableCell(tcText, fil, VecCol(Val(RS2!tem_codigo))) = Format(RS2!tem_nombre, fg_Pict(6, 2))
            'cLin = cLin & "|" & .TableCell(tcText, fil, VecCol(Val(RS2!tem_codigo)))
            RS2.MoveNext
        Loop
        For i = 3 To Col - 1
            cLin = cLin & "|" & .TableCell(tcText, fil, i)
        Next
        Print #1, cLin
        cLin = ""
        RS2.Close: Set RS2 = Nothing
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcColWidth, 1, 3, fil, Col) = 1000: .TableCell(tcAlign, 1, 3, fil, Col) = taCenterTop
    .TableCell(tcRows) = fil
    .PenColor = &HC0C0C0
    .TableBorder = tbAll
    .EndTable
    .EndDoc
End With
Preview.Show 1
Exit Function
Error_Impuestos:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Resume Next
    Close #1
    Exit Function
End Function

Public Function I_CtaCon()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_CtaCon
MsgTitulo = "Informe de Cuentas Contables"
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Cuenta Contable"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    RS2.Open "select * from a_param where par_codigo in ('ctagastos', 'ctainsumo', 'ctalimdes')", vg_db, adOpenStatic
    RS1.Open "select * from a_ctacontable order by cta_codigo", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Cuenta Asignada"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
           .TableCell(tcText, i, 1) = RS1!cta_codigo
           .TableCell(tcText, i, 2) = RS1!cta_nombre
           cParam = "": encuentra = False
           cCta = RS1!cta_codigo
           Do While Not RS2.EOF
              v_inicio = 0: v_final = 0
              If Not IsNull(RS2!par_valor) And Trim(RS2!par_valor) <> "" Then v_inicio = Val(Mid(Trim(fg_Quitachar(RS2!par_valor, ";")), 1, 6))
              If Not IsNull(RS2!par_valor) And Trim(RS2!par_valor) <> "" Then v_final = Val(Mid(Trim(fg_Quitachar(RS2!par_valor, ";")), Len(Trim(fg_Quitachar(RS2!par_valor, ";"))) - 5, 6))
              If v_inicio < v_final Then
                 For J = v_inicio To v_final
                     If J = Val(Trim(cCta)) Then
                        .TableCell(tcText, i, 3) = RS2!par_nombre
                     End If
                 Next J
              Else
                 For J = v_final To v_inicio
                     If J = Val(Trim(cCta)) Then
                        .TableCell(tcText, i, 3) = RS2!par_nombre
                     End If
                 Next J
              End If
              RS2.MoveNext
           Loop
           RS2.MoveFirst
           Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
           RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing: RS2.Close: Set RS2 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Function
Error_CtaCon:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_SalDevBod(Form As Object, tipo As String)
Dim rutcli As String, numdoc As Long, i As Long, total As Double, aAp As String
Dim numlin As Long, codmer As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String
On Local Error GoTo Error_SalDevBod
If tipo = "SP" Then
    MsgTitulo = "Informe de Salida a Producción"
ElseIf tipo = "DP" Then
    MsgTitulo = "Informe de Devolución de Producción"
End If
Preview.Show
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
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
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = IIf(tipo = "SP", "Salida de Bodega a Producción", "Devolución de Producción a Bodega")
    .TableCell(tcText, 2, 1) = Form.Label1.Caption
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 2, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3800: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 4) = 3700: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 3) = True
    .TableCell(tcText, 1, 1) = "Folio"
    .TableCell(tcText, 1, 2) = Form.fpLongInteger1(0).Text
    .TableCell(tcText, 1, 3) = "Casino"
    .TableCell(tcText, 1, 4) = Trim(Form.fpText1(1).Text) & " - " & Trim(Form.fpayuda(1).Caption)
    .TableCell(tcText, 2, 1) = "F. Emisión"
    .TableCell(tcText, 2, 2) = Form.fpDateTime1(0)
    .TableCell(tcText, 2, 3) = "Bodega"
    .TableCell(tcText, 2, 4) = Trim(Left(Form.Combo1(1).List(Form.Combo1(1).ListIndex), 50))
    .TableCell(tcText, 3, 1) = "F. Producción"
    .TableCell(tcText, 3, 2) = Form.fpDateTime1(1)
    .TableCell(tcText, 3, 3) = "Servicios"
    .TableCell(tcText, 3, 4) = Trim(Left(Form.Combo1(0).List(Form.Combo1(0).ListIndex), 50))
    Print #1, .TableCell(tcText, 1, 1) & "|"; .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    Print #1, .TableCell(tcText, 2, 1) & "|"; .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4)
    Print #1, .TableCell(tcText, 3, 1) & "|"; .TableCell(tcText, 3, 2) & "|" & .TableCell(tcText, 3, 3) & "|" & .TableCell(tcText, 3, 4)
    .TableBorder = tbBoxRows
    rutcli = Trim(LimpiaDato(Form.fpText1(1).Text))
    numdoc = Form.fpLongInteger1(0).Text
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .FontSize = 7
    .StartTable
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4800: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 600: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 800: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Unid."
    .TableCell(tcText, 1, 4) = IIf(tipo = "SP", "Cant.Calculada", "Cant.Salida")
    .TableCell(tcText, 1, 5) = IIf(tipo = "SP", "Cant.Salida", "Cant.Devolver")
    .TableCell(tcText, 1, 6) = "P.M.P."
    .TableCell(tcText, 1, 7) = "Total"
    Print #1, .TableCell(tcText, 1, 1) & "|"; .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|"; .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4800: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 600: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 800: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    'RS3.Open "Select   ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, sum(dev.dev_canmin * pro.pro_facing) as canmin " & _
             "From     b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm " & _
             "Where    ing.ing_codigo=dev.dev_coding " & _
             "And      ing.ing_unimed=unm.unm_codigo " & _
             "And      dev.dev_codmer=pro.pro_codigo " & _
             "And      dev.dev_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).Text)) & "'" & _
             "And      dev.dev_tipdoc='" & tipo & "' " & _
             "And      dev.dev_numdoc=" & Val(Form.fpLongInteger1(0).Text) & " " & _
             "Group by ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor " & _
             "Order by max(dev.dev_numlin)", vg_db, adOpenStatic
    aAp = Trim(vg_NUsr) & "_tmp_SalBod"
    fg_CheckTmp aAp
    RS3.Open "Select   ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, sum(dev.dev_canmin * pro.pro_facing) as canmin, max(dev_numlin) as num " & _
             "Into " & aAp & " " & _
             "From     b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm " & _
             "Where    ing.ing_codigo=dev.dev_coding " & _
             "And      ing.ing_unimed=unm.unm_codigo " & _
             "And      dev.dev_codmer=pro.pro_codigo " & _
             "And      dev.dev_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).Text)) & "' " & _
             "And      dev.dev_tipdoc='" & tipo & "' " & _
             "And      dev.dev_numdoc=" & Val(Form.fpLongInteger1(0).Text) & " " & _
             "Group by ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor " & _
             "Order by max(dev.dev_numlin)", vg_db, adOpenStatic
    Set RS3 = Nothing
    vg_db.Execute "Insert Into " & aAp & " " & _
                  "Select   '' as ing_codigo, 'Estructura Fija' as ing_nombre, '' as unm_nomcor, 0 as canmin, max(dev_numlin) as num " & _
                  "From     b_detventas " & _
                  "Where    dev_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).Text)) & "' " & _
                  "And      dev_tipdoc='" & tipo & "' " & _
                  "And      dev_numdoc=" & Val(Form.fpLongInteger1(0).Text) & " " & _
                  "And      dev_coding=''"
    RS3.Open "Select ing_codigo, ing_nombre, unm_nomcor, canmin, num From " & aAp & " Order by num", vg_db, adOpenStatic
    i = 1: total = 0
    Do While Not RS3.EOF
        .TableCell(tcFontBold, i) = True
        .TableCell(tcText, i, 1) = RS3!ing_codigo
        .TableCell(tcText, i, 2) = RS3!ing_nombre
        .TableCell(tcText, i, 3) = RS3!unm_nomcor
        .TableCell(tcText, i, 4) = IIf(RS3!ing_codigo = "", "", Format(RS3!canmin, fg_Pict(9, vg_DCa)))
        Print #1, .TableCell(tcText, i, 1) & "|"; .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
        i = i + 1
        RS1.Open "Select   dev.dev_codmer, dev.dev_canmin, dev.dev_canmer, dev.dev_predoc, " & _
                 "         dev.dev_ptotal, dev.dev_descri, uni.uni_nomcor " & _
                 "From     b_detventas dev, b_productos pro, a_unidad uni " & _
                 "Where    dev.dev_codmer=pro.pro_codigo " & _
                 "And      pro.pro_coduni=uni.uni_codigo " & _
                 "And      dev.dev_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).Text)) & "'" & _
                 "And      dev.dev_tipdoc='" & tipo & "' " & _
                 "And      dev.dev_numdoc=" & Val(Form.fpLongInteger1(0).Text) & " " & _
                 "And      dev.dev_coding='" & RS3!ing_codigo & "' " & _
                 "Group by dev.dev_codmer, dev.dev_canmin, dev.dev_canmer, dev.dev_predoc, " & _
                 "         dev.dev_ptotal, dev.dev_descri, uni.uni_nomcor " & _
                 "Order by max(dev.dev_numlin)", vg_db, adOpenStatic
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = RS1!dev_codmer
            .TableCell(tcText, i, 2) = RS1!dev_descri
            .TableCell(tcText, i, 3) = RS1!uni_nomcor
            .TableCell(tcText, i, 4) = Format(RS1!dev_canmin, fg_Pict(9, vg_DCa))
            .TableCell(tcText, i, 5) = Format(RS1!dev_canmer, fg_Pict(9, vg_DCa))
            .TableCell(tcText, i, 6) = Format(RS1!dev_predoc, fg_Pict(9, vg_DPr))
            .TableCell(tcText, i, 7) = Format(RS1!dev_ptotal, fg_Pict(9, vg_DPr))
            Print #1, .TableCell(tcText, i, 1) & "|"; .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                      .TableCell(tcText, i, 5) & "|"; .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)
            total = total + Format(RS1!dev_ptotal, fg_Pict(9, 2))
            RS1.MoveNext: i = i + 1
        Loop
        RS3.MoveNext
        RS1.Close: Set RS1 = Nothing
        i = i + 1
        .TableCell(tcText, i, 1) = ""
    Loop
    RS3.Close: Set RS3 = Nothing
    .TableCell(tcRows) = i - 1
    .PenColor = &HC0C0C0
    .TableBorder = tbBottom
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcFontBold, 1, 1, 1, 3) = True
    .TableCell(tcColWidth, 1, 1) = 7500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, 1, 2) = 1500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, 1, 3) = 1500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcText, 1, 2) = "Total"
    .TableCell(tcText, 1, 3) = Format(total, fg_Pict(9, vg_DPr))
    Print #1, "|||||" & .TableCell(tcText, 1, 2) & "|"; .TableCell(tcText, 1, 3)
    .TableBorder = tbNone
    Print #1, " ": Print #1, "|||||" & "_____________________"
    Print #1, " "
    If tipo = "SP" Then
        Print #1, Space(100) & "Entregado conforme"
    Else
        Print #1, Space(100) & "Recibido conforme"
    End If
    .EndTable
    .FontBold = True
    .CurrentX = 8800
    .CurrentY = 14000
    .Text = IIf(tipo = "SP", "_____________________", "____________________")
    .CurrentX = 8950
    .CurrentY = 14200
    .Text = IIf(tipo = "SP", "Entregado conforme", "Recibido conforme")
    .EndDoc
    Close #1
End With
Exit Function
Error_SalDevBod:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Mermas(Form As Object)
Dim rutcli As String, numdoc As Long, i As Long, total As Double
Dim numlin As Long, codmer As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String
On Local Error GoTo Error_Mermas1
MsgTitulo = "Mermas"
Preview.Show
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
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
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = "Mermas"
    .TableCell(tcText, 2, 1) = Form.Label1.Caption
    Print #1, .TableCell(tcText, 1, 1)
    Print #1, .TableCell(tcText, 2, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3800: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 4) = 3700: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 3) = True
    .TableCell(tcText, 1, 1) = "Folio"
    .TableCell(tcText, 1, 2) = Form.fpLongInteger1(0).Text
    .TableCell(tcText, 1, 3) = "Bodega"
    .TableCell(tcText, 1, 4) = Trim(Left(Form.Combo1(1).List(Form.Combo1(1).ListIndex), 50))
    .TableCell(tcText, 2, 1) = "F. Emisión"
    .TableCell(tcText, 2, 2) = Form.fpDateTime1(0)
    .TableCell(tcText, 2, 3) = "Tipo de Merma"
    .TableCell(tcText, 2, 4) = Trim(Left(Form.Combo1(0).List(Form.Combo1(0).ListIndex), 50))
    .TableCell(tcText, 3, 1) = "Casino"
    .TableCell(tcText, 3, 2) = Trim(LimpiaDato(Form.fpText1(1).Text)) & " - " & Trim(Form.fpayuda(1).Caption)
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbBox
    rutcli = Trim(LimpiaDato(Form.fpText1(1).Text))
    numdoc = Form.fpLongInteger1(0).Text
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 6: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1500: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Unidad"
    .TableCell(tcText, 1, 4) = "Cantidad"
    .TableCell(tcText, 1, 5) = "P.M.P."
    .TableCell(tcText, 1, 6) = "Total"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6)
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 6: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1500: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
    RS1.Open "select dev.*, uni.uni_nombre from b_detventas dev, b_productos pro, a_unidad uni " & _
             "where dev.dev_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).Text)) & "' and dev.dev_tipdoc='ME' " & _
             "and dev.dev_numdoc=" & Val(Form.fpLongInteger1(0).Text) & " and dev.dev_codmer=pro.pro_codigo " & _
             "and pro.pro_coduni=uni.uni_codigo order by dev.dev_numlin", vg_db, adOpenStatic
    i = 1: total = 0
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!dev_codmer
        .TableCell(tcText, i, 2) = RS1!dev_descri
        .TableCell(tcText, i, 3) = RS1!uni_nombre
        .TableCell(tcText, i, 4) = Format(RS1!dev_canmer, fg_Pict(9, vg_DCa))
        .TableCell(tcText, i, 5) = Format(RS1!dev_predoc, fg_Pict(9, vg_DPr))
        .TableCell(tcText, i, 6) = Format(RS1!dev_ptotal, fg_Pict(9, vg_DPr))
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6)
        total = total + Format(RS1!dev_ptotal, fg_Pict(9, 2))
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcRows) = i - 1
    .PenColor = &HC0C0C0
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcFontBold, 1, 1, 1, 3) = True
    .TableCell(tcColWidth, 1, 1) = 7500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, 1, 2) = 1500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, 1, 3) = 1500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcText, 1, 2) = "Total"
    .TableCell(tcText, 1, 3) = Format(total, fg_Pict(9, vg_DPr))
    Print #1, "||||" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    Print #1, Chr(13) & Chr(13)
    Print #1, "||||" & "_____________________"
    Print #1, " "
    Print #1, "||||" & "Entregado conforme"
    .TableBorder = tbNone
    .EndTable
    .FontBold = True
    .CurrentX = 8800
    .CurrentY = 14000
    .Text = "_____________________"
    .CurrentX = 8950
    .CurrentY = 14200
    .Text = "Entregado conforme"
    .EndDoc
    Close #1
End With
Exit Function
Error_Mermas1:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Traspaso(Form As Object)
Dim rutcli As String, numdoc As Long, i As Long, total As Double
Dim numlin As Long, codmer As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String
Dim cLin As String
On Local Error GoTo Error_Traspaso
MsgTitulo = "Traspasos"
Preview.Show
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
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
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = "Traspaso entre Casinos"
    .TableCell(tcText, 2, 1) = Form.Label1.Caption
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 2, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3800: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 4) = 3700: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 3) = True
    .TableCell(tcText, 1, 1) = "Folio"
    .TableCell(tcText, 1, 2) = Form.fpLongInteger1(0).Text
    .TableCell(tcText, 1, 3) = "Bodega"
    .TableCell(tcText, 1, 4) = Trim(Left(Form.Combo1(1).List(Form.Combo1(1).ListIndex), 50))
    .TableCell(tcText, 2, 1) = "F. Emisión"
    .TableCell(tcText, 2, 2) = Form.fpDateTime1(0)
    .TableCell(tcText, 2, 3) = "Tipo Traspaso"
    .TableCell(tcText, 2, 4) = IIf(Form.Option1(1).Value = True, "Recibido", "Entregado")
    .TableCell(tcText, 3, 1) = "Casino"
    .TableCell(tcText, 3, 2) = Trim(LimpiaDato(Form.fpText1(0).Text)) & " - " & Trim(Form.fpayuda(0).Caption)
    .TableCell(tcText, 3, 3) = Trim(Form.Label3(0).Caption)
    .TableCell(tcText, 3, 4) = Trim(LimpiaDato(Form.fpText1(1).Text)) & " - " & Trim(Form.fpayuda(1).Caption)
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2) & "|" & .TableCell(tcText, 3, 3) & "|" & .TableCell(tcText, 3, 4)
    .TableBorder = tbBox
    rutcli = Trim(LimpiaDato(Form.fpText1(1).Text))
    numdoc = Form.fpLongInteger1(0).Text
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    cLin = ""
    .StartTable
    .TableCell(tcCols) = IIf(Form.Option1(1).Value = True, 7, 6): .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = IIf(Form.Option1(1).Value = True, 1500, 2000): .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = IIf(Form.Option1(1).Value = True, 3000, 3500): .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    
    If Form.Option1(1).Value = True Then .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , IIf(Form.Option1(1).Value = True, 6, 5)) = 1500: .TableCell(tcAlign, , IIf(Form.Option1(1).Value = True, 6, 5)) = taRightTop
    .TableCell(tcColWidth, , IIf(Form.Option1(1).Value = True, 7, 6)) = 1500: .TableCell(tcAlign, , IIf(Form.Option1(1).Value = True, 7, 6)) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Unidad"
    .TableCell(tcText, 1, 4) = "Cantidad"
    cLin = cLin & .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    If Form.Option1(1).Value = True Then .TableCell(tcText, 1, 5) = "Cant.Recibida": cLin = cLin & "|" & .TableCell(tcText, 1, 5)
    .TableCell(tcText, 1, IIf(Form.Option1(1).Value = True, 6, 5)) = IIf(Form.Option1(1).Value = True, "Precio", "P.M.P."): cLin = cLin & "|" & .TableCell(tcText, 1, IIf(Form.Option1(1).Value = True, 6, 5))
    .TableCell(tcText, 1, IIf(Form.Option1(1).Value = True, 7, 6)) = "Total": cLin = cLin & "|" & .TableCell(tcText, 1, IIf(Form.Option1(1).Value = True, 7, 6))
    Print #1, cLin
    .TableBorder = tbAll
    cLin = ""
    .EndTable
    .StartTable
    .TableCell(tcCols) = IIf(Form.Option1(1).Value = True, 7, 6): .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = IIf(Form.Option1(1).Value = True, 1500, 2000): .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = IIf(Form.Option1(1).Value = True, 3000, 3500): .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    If Form.Option1(1).Value = True Then .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , IIf(Form.Option1(1).Value = True, 6, 5)) = 1500: .TableCell(tcAlign, , IIf(Form.Option1(1).Value = True, 6, 5)) = taRightTop
    .TableCell(tcColWidth, , IIf(Form.Option1(1).Value = True, 7, 6)) = 1500: .TableCell(tcAlign, , IIf(Form.Option1(1).Value = True, 7, 6)) = taRightTop
    RS1.Open "select dev.*, uni.uni_nombre from b_detventas dev, b_productos pro, a_unidad uni " & _
             "where dev.dev_rutcli='" & Trim(LimpiaDato(Form.fpText1(0).Text)) & "' and dev.dev_tipdoc='TR' " & _
             "and dev.dev_numdoc=" & Val(Form.fpLongInteger1(0).Text) & " and dev.dev_codmer=pro.pro_codigo " & _
             "and pro.pro_coduni=uni.uni_codigo order by dev.dev_numlin", vg_db, adOpenStatic
    i = 1: total = 0
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!dev_codmer
        .TableCell(tcText, i, 2) = RS1!dev_descri
        .TableCell(tcText, i, 3) = RS1!uni_nombre
        .TableCell(tcText, i, 4) = Format(IIf(Form.Option1(1).Value = True, RS1!dev_canmin, RS1!dev_canmer), fg_Pict(9, vg_DCa))
        'cLin = cLin & .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
        If Form.Option1(1).Value = True Then .TableCell(tcText, i, 5) = Format(RS1!dev_canmer, fg_Pict(9, vg_DCa)): cLin = cLin & "|" & .TableCell(tcText, i, 5) & "|"
        .TableCell(tcText, i, IIf(Form.Option1(1).Value = True, 6, 5)) = Format(RS1!dev_predoc, fg_Pict(9, vg_DPr)): cLin = cLin & .TableCell(tcText, i, IIf(Form.Option1(1).Value = True, 6, 5))
        .TableCell(tcText, i, IIf(Form.Option1(1).Value = True, 7, 6)) = Format(RS1!dev_ptotal, fg_Pict(9, vg_DPr)): cLin = cLin & "|" & .TableCell(tcText, i, IIf(Form.Option1(1).Value = True, 7, 6))
        total = total + Format(RS1!dev_ptotal, fg_Pict(9, 2))
        'Print #1, cLin
        RS1.MoveNext: i = i + 1
    Loop
    For J = 1 To i - 1
        cLin = ""
        For k = 1 To .TableCell(tcCols)
            cLin = cLin & .TableCell(tcText, J, k) & "|"
        Next k
        Print #1, cLin
    Next J
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcRows) = i - 1
    .PenColor = &HC0C0C0
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcFontBold, 1, 1, 1, 3) = True
    .TableCell(tcColWidth, 1, 1) = 7500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, 1, 2) = 1500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, 1, 3) = 1500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcText, 1, 2) = "Total"
    .TableCell(tcText, 1, 3) = Format(total, fg_Pict(9, vg_DPr))
    Print #1, "|||||" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    Print #1, " ": Print #1, " "
    Print #1, "|||||" & "_____________________"
    Print #1, "|||||" & IIf(Form.Option1(1).Value = True, "Recibido conforme", "Entregado conforme")
    .TableBorder = tbNone
    .EndTable
    .FontBold = True
    .CurrentX = 8800
    .CurrentY = 14000
    .Text = IIf(Form.Option1(1).Value = True, "____________________", "_____________________")
    .CurrentX = 8950
    .CurrentY = 14200
    .Text = IIf(Form.Option1(1).Value = True, "Recibido conforme", "Entregado conforme")
    .EndDoc
    Close #1
End With
Exit Function
Error_Traspaso:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_VentaDir(Form As Object, tipo As String)
Dim rutcli As String, numdoc As Long, i As Long, total As Double
Dim numlin As Long, codmer As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String
On Local Error GoTo Error_VtaDir
MsgTitulo = "Venta Directa"
Preview.Show
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
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
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = IIf(tipo = "FA", "Factura", "Guia Despacho") & " - Venta Directa"
    .TableCell(tcText, 2, 1) = Form.Label1.Caption
    Print #1, .TableCell(tcText, 1, 1)
    Print #1, .TableCell(tcText, 2, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3800: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 4) = 3700: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 3) = True
    .TableCell(tcText, 1, 1) = "Documento"
    .TableCell(tcText, 1, 2) = Trim(Left(Form.Combo1(0).List(Form.Combo1(1).ListIndex), 50))
    .TableCell(tcText, 1, 3) = "Folio"
    .TableCell(tcText, 1, 4) = Form.fpLongInteger1(0).Text
    .TableCell(tcText, 2, 1) = "Bodega"
    .TableCell(tcText, 2, 2) = Trim(Left(Form.Combo1(1).List(Form.Combo1(1).ListIndex), 50))
    .TableCell(tcText, 2, 3) = "F. Emisión"
    .TableCell(tcText, 2, 4) = Form.fpDateTime1(0)
    .TableCell(tcText, 3, 1) = "Casino"
    .TableCell(tcText, 3, 2) = Trim(LimpiaDato(Form.fpText1(0).Text)) & " - " & Trim(Form.fpayuda(0).Caption)
    .TableCell(tcText, 3, 3) = "Cliente"
    .TableCell(tcText, 3, 4) = Trim(LimpiaDato(Form.fpText1(1).Text)) & " - " & Trim(Form.fpayuda(1).Caption)
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2) & "|" & .TableCell(tcText, 3, 3) & "|" & .TableCell(tcText, 3, 4)
    .TableBorder = tbBox
    rutcli = Trim(LimpiaDato(Form.fpText1(1).Text))
    numdoc = Form.fpLongInteger1(0).Text
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1700: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3200: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1500: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Unidad"
    .TableCell(tcText, 1, 4) = "Cantidad"
    .TableCell(tcText, 1, 5) = "%Sob.Costo"
    .TableCell(tcText, 1, 6) = "Precio"
    .TableCell(tcText, 1, 7) = "Total"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7)
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1700: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3200: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1500: .TableCell(tcAlign, , 7) = taRightTop
    RS1.Open "select dev.*, uni.uni_nombre from b_detventas dev, b_productos pro, a_unidad uni " & _
             "where dev.dev_rutcli='" & Trim(LimpiaDato(Form.fpText1(0).Text)) & "' and dev.dev_tipdoc='" & tipo & "' " & _
             "and dev.dev_numdoc=" & Val(Form.fpLongInteger1(0).Text) & " and dev.dev_codmer=pro.pro_codigo " & _
             "and pro.pro_coduni=uni.uni_codigo order by dev.dev_numlin", vg_db, adOpenStatic
    i = 1: total = 0
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!dev_codmer
        .TableCell(tcText, i, 2) = RS1!dev_descri
        .TableCell(tcText, i, 3) = RS1!uni_nombre
        .TableCell(tcText, i, 4) = Format(RS1!dev_canmer, fg_Pict(9, vg_DCa))
        .TableCell(tcText, i, 5) = Format(RS1!dev_porcen, fg_Pict(9, 2))
        .TableCell(tcText, i, 6) = Format(RS1!dev_predoc, fg_Pict(9, vg_DPr))
        .TableCell(tcText, i, 7) = Format(RS1!dev_ptotal, fg_Pict(9, vg_DPr))
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                  .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)
        total = total + Format(RS1!dev_ptotal, fg_Pict(9, 2))
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcRows) = i - 1
    .PenColor = &HC0C0C0
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcFontBold, 1, 1, 1, 3) = True
    .TableCell(tcColWidth, 1, 1) = 7900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, 1, 2) = 1500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, 1, 3) = 1500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcText, 1, 2) = "Total"
    .TableCell(tcText, 1, 3) = Format(total, fg_Pict(9, vg_DPr))
    Print #1, "|||||" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    Print #1, " ": Print #1, " "
    Print #1, "|||||_____________________"
    Print #1, " "
    Print #1, "|||||" & "Entregado conforme"
    .TableBorder = tbNone
    .EndTable
    .FontBold = True
    .CurrentX = 8800
    .CurrentY = 14000
    .Text = "_____________________"
    .CurrentX = 8950
    .CurrentY = 14200
    .Text = "Entregado conforme"
    .EndDoc
    Close #1
End With
Exit Function
Error_VtaDir:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_TipAju()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_TipAju
MsgTitulo = "Informe Tipo de Ajuste"
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Tipo de Ajuste"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_tipoajuste where aju_tipaju=1 order by aju_codigo", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 6000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Tipo"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 6000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taLeftTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 2) = RS1!aju_nombre
            .TableCell(tcText, i, 1) = RS1!aju_codigo
            .TableCell(tcText, i, 3) = IIf(RS1!aju_tipo = "D", "Descuenta", "Aumenta")
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Function
Error_TipAju:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_TipMer()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Mermas
MsgTitulo = "Informe Tipo de Mermas"
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Tipo de Merma"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_tipoajuste where aju_tipaju=0 order by aju_codigo", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 5000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 5000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = RS1!aju_codigo
            .TableCell(tcText, i, 2) = RS1!aju_nombre
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2)
            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Exit Function
Error_Mermas:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Toma1(SQL As String)
Dim i As Long
On Local Error GoTo Error_Toma1
MsgTitulo = "Toma de Inventario"
RS1.Open SQL, vg_db, adOpenStatic
If RS1.EOF Then
    RS1.Close: Set RS1 = Nothing
    MsgBox "No existe información...", vbExclamation + vbOKOnly, MsgTitulo
    Exit Function
End If
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Listado para toma de Inventario"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcFontSize, 1, 1, 3, 2) = 8: .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 2) = False
    .TableCell(tcText, 1, 1) = "Bodega": .TableCell(tcText, 1, 2) = Left(M_TomInv.Combo1(0).Text, 50)
    .TableCell(tcText, 2, 1) = "Toma de Inventario": .TableCell(tcText, 2, 2) = M_TomInv.Date1(0).Text
    .TableCell(tcText, 3, 1) = "Tipo de Producto": .TableCell(tcText, 3, 2) = IIf(I_TomInv.optTIPPRO(0).Value = True, Left(I_TomInv.Combo1(0).Text, 50), "Todos")
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 2000: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Unidad"
    .TableCell(tcText, 1, 4) = "Cantidad"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 2000: .TableCell(tcAlign, , 4) = taCenterTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i) = ""
            .TableCell(tcText, i + 1, 1) = RS1!tin_codpro
            .TableCell(tcText, i + 1, 2) = RS1!pro_nombre
            .TableCell(tcText, i + 1, 3) = RS1!uni_nomcor
            .TableCell(tcText, i + 1, 4) = "___________"
            Print #1, "|" & .TableCell(tcText, i + 1, 1) & "|" & .TableCell(tcText, i + 1, 2) & "|"; .TableCell(tcText, i + 1, 3) & .TableCell(tcText, i + 1, 4)
            RS1.MoveNext: i = i + 2
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
    Close #1
End With
Preview.Show 1
Exit Function
Error_Toma1:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Toma2(SQL As String)
Dim i As Long
On Local Error GoTo Error_Toma2
MsgTitulo = "Toma de Inventario"
RS1.Open SQL, vg_db, adOpenStatic
If RS1.EOF Then
    RS1.Close: Set RS1 = Nothing
    MsgBox "No existe información...", vbExclamation + vbOKOnly, MsgTitulo
    Exit Function
End If
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Diferencias en toma de Inventario"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcFontSize, 1, 1, 3, 2) = 8: .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 2) = False
    .TableCell(tcText, 1, 1) = "Bodega": .TableCell(tcText, 1, 2) = Left(M_TomInv.Combo1(0).Text, 50)
    .TableCell(tcText, 2, 1) = "Toma de Inventario": .TableCell(tcText, 2, 2) = M_TomInv.Date1(0).Text
    .TableCell(tcText, 3, 1) = "Tipo de Producto": .TableCell(tcText, 3, 2) = IIf(I_TomInv.optTIPPRO(0).Value = True, Left(I_TomInv.Combo1(0).Text, 50), "Todos")
    Print #1, .TableCell(tcText, 1, 1) & "|"; .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|"; .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|"; .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 6: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1500: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Unidad"
    .TableCell(tcText, 1, 4) = "Stock Físico"
    .TableCell(tcText, 1, 5) = "Stock Sistema"
    .TableCell(tcText, 1, 6) = "Diferencia"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & _
              .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 6: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1500: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = RS1!tin_codpro
            .TableCell(tcText, i, 2) = RS1!pro_nombre
            .TableCell(tcText, i, 3) = RS1!uni_nomcor
            .TableCell(tcText, i, 4) = Format(RS1!tin_stofis, fg_Pict(9, vg_DCa))
            .TableCell(tcText, i, 5) = Format(RS1!tin_stosis, fg_Pict(9, vg_DCa))
            .TableCell(tcText, i, 6) = Format(RS1!tin_stofis - RS1!tin_stosis, fg_Pict(9, vg_DCa))
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & _
                      .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6)
            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
    Close #1
End With
Preview.Show 1
Exit Function
Error_Toma2:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Toma3(SQL As String)
Dim i As Long, sumCuenta As Double, sumTipo As Double, sumCuentaLimDes As Double, sumCuentaAlimen As Double
On Local Error GoTo Error_Toma3
MsgTitulo = "Toma de Inventario"
RS3.Open SQL, vg_db, adOpenStatic
If RS3.EOF Then
    RS3.Close: Set RS3 = Nothing
    MsgBox "No existe información...", vbExclamation + vbOKOnly, MsgTitulo
    Exit Function
End If
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
    .FontSize = 7
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Inventario Físico Valorizado"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcFontSize, 1, 1, 3, 2) = 8: .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 2) = False
    .TableCell(tcText, 1, 1) = "Bodega": .TableCell(tcText, 1, 2) = Left(M_TomInv.Combo1(0).Text, 50)
    .TableCell(tcText, 2, 1) = "Toma de Inventario": .TableCell(tcText, 2, 2) = M_TomInv.Date1(0).Text
    .TableCell(tcText, 3, 1) = "Tipo de Producto": .TableCell(tcText, 3, 2) = IIf(I_TomInv.optTIPPRO(0).Value = True, Left(I_TomInv.Combo1(0).Text, 50), "Todos")
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 200: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 4600: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 500: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1500: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 200
    .TableCell(tcText, 1, 2) = "Código"
    .TableCell(tcText, 1, 3) = "Descripción"
    .TableCell(tcText, 1, 4) = "Unidad"
    .TableCell(tcText, 1, 5) = "Cantidad"
    .TableCell(tcText, 1, 6) = "Precio"
    .TableCell(tcText, 1, 7) = "Total"
    Print #1, .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
             .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7)
    .TableBorder = tbBox
    .EndTable
    sumCuentaLimDes = 0: sumCuentaAlimen = 0
    If Not RS3.EOF Then
        RS1.Open "select cta_codigo, cta_nombre from a_ctacontable order by cta_nombre", vg_db, adOpenStatic
        Do While Not RS1.EOF
            RS3.Filter = "pro_ctacon='" & RS1!cta_codigo & "'"
            If RS3.RecordCount > 0 Then RS3.Find "pro_ctacon='" & RS1!cta_codigo & "'", , adSearchForward
            If Not RS3.EOF Then
                sumCuenta = 0
                .StartTable
                .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
                .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftTop: .TableCell(tcFontBold, 2, 1) = True
                .TableCell(tcText, 2, 1) = RS1!cta_codigo & "  " & RS1!cta_nombre: .TableCell(tcFontUnderline) = True
                Print #1, .TableCell(tcText, 2, 1)
                .TableBorder = tbNone
                .EndTable
                RS2.Open "select tip_codigo, tip_nombre from a_tipopro order by tip_nombre", vg_db, adOpenStatic
                Do While Not RS2.EOF
                    RS3.Filter = "pro_ctacon='" & RS1!cta_codigo & "' and pro_codtip=" & RS2!tip_codigo
                    If RS3.RecordCount > 0 Then RS3.Find "pro_codtip=" & RS2!tip_codigo, , adSearchForward
                    If Not RS3.EOF Then
                        .StartTable
                        .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
                        .TableCell(tcFontBold, 2, 1) = True
                        .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftTop
                        .TableCell(tcText, 2, 1) = RS2!tip_codigo & "  " & RS2!tip_nombre
                        Print #1, .TableCell(tcText, 2, 1)
                        .TableBorder = tbNone
                        .EndTable
                        .StartTable
                        .TableCell(tcCols) = 7: .TableCell(tcRows) = 5000
                        .TableCell(tcColWidth, , 1) = 200: .TableCell(tcAlign, , 1) = taLeftTop
                        .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taLeftTop
                        .TableCell(tcColWidth, , 3) = 4600: .TableCell(tcAlign, , 3) = taLeftTop
                        .TableCell(tcColWidth, , 4) = 500: .TableCell(tcAlign, , 4) = taCenterTop
                        .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
                        .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
                        .TableCell(tcColWidth, , 7) = 1500: .TableCell(tcAlign, , 7) = taRightTop
                        i = 1
                        sumTipo = 0
                        Do While Not RS3.EOF
                            .TableCell(tcText, i, 2) = RS3!tin_codpro
                            .TableCell(tcText, i, 3) = RS3!pro_nombre
                            .TableCell(tcText, i, 4) = RS3!uni_nomcor
                            .TableCell(tcText, i, 5) = Format(RS3!tin_stofis, fg_Pict(9, vg_DCa))
                            .TableCell(tcText, i, 6) = Format(RS3!tin_propon, fg_Pict(9, vg_DPr))
                            .TableCell(tcText, i, 7) = Format(Format(RS3!tin_stofis, fg_Pict(9, vg_DCa)) * Format(RS3!tin_propon, fg_Pict(9, vg_DPr)), fg_Pict(9, vg_DPr))
                            Print #1, .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                                      .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)

                            sumTipo = sumTipo + Round(Format(RS3!tin_stofis, fg_Pict(9, vg_DCa)) * Format(RS3!tin_propon, fg_Pict(9, vg_DPr)), vg_DPr)
                            RS3.MoveNext: i = i + 1
                        Loop
                        sumCuenta = sumCuenta + sumTipo
                        .TableCell(tcText, i + 1, 6) = "Total Familia"
                        .TableCell(tcText, i + 1, 7) = Format(sumTipo, fg_Pict(9, vg_DPr))
                        Print #1, "|" & "|" & "|" & "|" & .TableCell(tcText, i + 1, 6) & "|" & .TableCell(tcText, i + 1, 7)
                        .TableCell(tcRows) = i + 2
                        .TableBorder = tbNone
                        .EndTable
                    End If
                    RS2.MoveNext
                Loop
                If RS1!cta_codigo = GetParametro("ctalimdes") Then sumCuentaLimDes = sumCuenta
                If RS1!cta_codigo = GetParametro("ctainsumo") Then sumCuentaAlimen = sumCuenta
                RS2.Close: Set RS2 = Nothing
                .StartTable
                .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
                .TableCell(tcFontBold) = True: .TableCell(tcAlign) = taRightTop
                .TableCell(tcColWidth, , 1) = 9300: .TableCell(tcColWidth, , 2) = 1500
                .TableCell(tcText, , 1) = "Total Cuenta"
                .TableCell(tcText, , 2) = Format(sumCuenta, fg_Pict(9, vg_DPr))
                Print #1, .TableCell(tcText, 1, 1)
                Print #1, .TableCell(tcText, 1, 2)
                .TableBorder = tbNone
                .EndTable
                .StartTable
                .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
                .TableCell(tcFontBold) = True
                .TableCell(tcColWidth) = 10900: .TableCell(tcAlign) = taRightTop
                .TableCell(tcText, 1, 1) = String(144, "_")
                Print #1, .TableCell(tcText, 1, 1)
                .TableBorder = tbNone
                .EndTable
            End If
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
    End If
    RS3.Close: Set RS3 = Nothing
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 6
    .TableCell(tcFontBold) = True
    .TableCell(tcColWidth, , 1) = 1900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 7500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcText, 2, 1) = "Alimentos & Bebidas":       .TableCell(tcText, 2, 2) = Format(sumCuentaAlimen, fg_Pict(9, vg_DPr))
    .TableCell(tcText, 3, 1) = "Limpieza  & Desechables":   .TableCell(tcText, 3, 2) = Format(sumCuentaLimDes, fg_Pict(9, vg_DPr))
    .TableCell(tcText, 4, 2) = String(15, "_")
    .TableCell(tcText, 5, 1) = "Totales Generales":         .TableCell(tcText, 5, 2) = Format(sumCuentaLimDes + sumCuentaAlimen, fg_Pict(9, vg_DPr))
    Print #1, " ": Print #1, " "
    Print #1, .TableCell(tcText, 2, 1) & "|"; .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|"; .TableCell(tcText, 3, 2)
    Print #1, .TableCell(tcText, 4, 2)
    Print #1, .TableCell(tcText, 5, 1) & "|"; .TableCell(tcText, 5, 2)
    .TableBorder = tbBottom
    .EndTable
    .EndDoc
    Close #1
End With
Preview.Show 1
Exit Function
Error_Toma3:
    MsgBox Err.Number & " " & Err.Description, vbExclamation
    Close #1
    Exit Function
End Function

Public Function I_Toma4(SQL As String)
Dim i As Long, sumCuenta As Double, sumTipo As Double, sumCuentaLimDes As Double, sumCuentaAlimen As Double
On Local Error GoTo Error_Toma4
MsgTitulo = "Toma de Inventario"
RS3.Open SQL, vg_db, adOpenStatic
If RS3.EOF Then
    RS3.Close: Set RS3 = Nothing
    MsgBox "No existe información...", vbExclamation + vbOKOnly, MsgTitulo
    Exit Function
End If
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
    .FontSize = 7
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Inventario Sistema Valorizado"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcFontSize, 1, 1, 3, 2) = 8: .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 2) = False
    .TableCell(tcText, 1, 1) = "Bodega": .TableCell(tcText, 1, 2) = Left(M_TomInv.Combo1(0).Text, 50)
    .TableCell(tcText, 2, 1) = "Toma de Inventario": .TableCell(tcText, 2, 2) = M_TomInv.Date1(0).Text
    .TableCell(tcText, 3, 1) = "Tipo de Producto": .TableCell(tcText, 3, 2) = IIf(I_TomInv.optTIPPRO(0).Value = True, Left(I_TomInv.Combo1(0).Text, 50), "Todos")
    Print #1, .TableCell(tcText, 1, 1) & "|"; .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|"; .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 1, 1) & "|"; .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 200: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 4600: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 500: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1500: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 200
    .TableCell(tcText, 1, 2) = "Código"
    .TableCell(tcText, 1, 3) = "Descripción"
    .TableCell(tcText, 1, 4) = "Unidad"
    .TableCell(tcText, 1, 5) = "Cantidad"
    .TableCell(tcText, 1, 6) = "Precio"
    .TableCell(tcText, 1, 7) = "Total"
     Print #1, .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
             .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7)
    .TableBorder = tbBox
    .EndTable
    sumCuentaLimDes = 0: sumCuentaAlimen = 0
    If Not RS3.EOF Then
        RS1.Open "select cta_codigo, cta_nombre from a_ctacontable order by cta_nombre", vg_db, adOpenStatic
        Do While Not RS1.EOF
            RS3.Filter = "pro_ctacon='" & RS1!cta_codigo & "'"
            If RS3.RecordCount > 0 Then RS3.Find "pro_ctacon='" & RS1!cta_codigo & "'", , adSearchForward
            If Not RS3.EOF Then
                sumCuenta = 0
                .StartTable
                .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
                .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftTop: .TableCell(tcFontBold, 2, 1) = True
                .TableCell(tcText, 2, 1) = RS1!cta_codigo & "  " & RS1!cta_nombre: .TableCell(tcFontUnderline) = True
                Print #1, .TableCell(tcText, 2, 1)
                .TableBorder = tbNone
                .EndTable
                RS2.Open "select tip_codigo, tip_nombre from a_tipopro order by tip_nombre", vg_db, adOpenStatic
                Do While Not RS2.EOF
                    RS3.Filter = "pro_ctacon='" & RS1!cta_codigo & "' and pro_codtip=" & RS2!tip_codigo
                    If RS3.RecordCount > 0 Then RS3.Find "pro_codtip=" & RS2!tip_codigo, , adSearchForward
                    If Not RS3.EOF Then
                        .StartTable
                        .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
                        .TableCell(tcFontBold, 2, 1) = True
                        .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftTop
                        .TableCell(tcText, 2, 1) = RS2!tip_codigo & "  " & RS2!tip_nombre
                        Print #1, .TableCell(tcText, 2, 1)
                        .TableBorder = tbNone
                        .EndTable
                        .StartTable
                        .TableCell(tcCols) = 7: .TableCell(tcRows) = 5000
                        .TableCell(tcColWidth, , 1) = 200: .TableCell(tcAlign, , 1) = taLeftTop
                        .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taLeftTop
                        .TableCell(tcColWidth, , 3) = 4600: .TableCell(tcAlign, , 3) = taLeftTop
                        .TableCell(tcColWidth, , 4) = 500: .TableCell(tcAlign, , 4) = taCenterTop
                        .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
                        .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
                        .TableCell(tcColWidth, , 7) = 1500: .TableCell(tcAlign, , 7) = taRightTop
                        i = 1
                        sumTipo = 0
                        Do While Not RS3.EOF
                            .TableCell(tcText, i, 2) = RS3!tin_codpro
                            .TableCell(tcText, i, 3) = RS3!pro_nombre
                            .TableCell(tcText, i, 4) = RS3!uni_nomcor
                            .TableCell(tcText, i, 5) = Format(RS3!tin_stosis, fg_Pict(9, vg_DCa))
                            .TableCell(tcText, i, 6) = Format(RS3!tin_propon, fg_Pict(9, vg_DPr))
                            .TableCell(tcText, i, 7) = Format(Format(RS3!tin_stosis, fg_Pict(9, vg_DCa)) * Format(RS3!tin_propon, fg_Pict(9, vg_DPr)), fg_Pict(9, vg_DPr))
                            Print #1, .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                                      .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)
                            sumTipo = sumTipo + Round(Format(RS3!tin_stosis, fg_Pict(9, vg_DCa)) * Format(RS3!tin_propon, fg_Pict(9, vg_DPr)), vg_DPr)
                            RS3.MoveNext: i = i + 1
                        Loop
                        sumCuenta = sumCuenta + sumTipo
                        .TableCell(tcText, i + 1, 6) = "Total Familia"
                        .TableCell(tcText, i + 1, 7) = Format(sumTipo, fg_Pict(9, vg_DPr))
                        Print #1, "|" & "|" & "|" & "|" & .TableCell(tcText, i + 1, 6) & "|" & .TableCell(tcText, i + 1, 7)
                        .TableCell(tcRows) = i + 2
                        .TableBorder = tbNone
                        .EndTable
                    End If
                    RS2.MoveNext
                Loop
                RS2.Close: Set RS2 = Nothing
                If RS1!cta_codigo = GetParametro("ctalimdes") Then sumCuentaLimDes = sumCuentaLimDes + sumCuenta
                If RS1!cta_codigo = GetParametro("ctainsumo") Then sumCuentaAlimen = sumCuentaAlimen + sumCuenta
                .StartTable
                .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
                .TableCell(tcFontBold) = True: .TableCell(tcAlign) = taRightTop
                .TableCell(tcColWidth, , 1) = 9300: .TableCell(tcColWidth, , 2) = 1500
                .TableCell(tcText, , 1) = "Total Cuenta"
                .TableCell(tcText, , 2) = Format(sumCuenta, fg_Pict(9, vg_DPr))
                Print #1, .TableCell(tcText, 1, 1)
                Print #1, .TableCell(tcText, 1, 2)
                .TableBorder = tbNone
                .EndTable
                .StartTable
                .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
                .TableCell(tcFontBold) = True
                .TableCell(tcColWidth) = 10900: .TableCell(tcAlign) = taRightTop
                .TableCell(tcText, 1, 1) = String(144, "_")
                Print #1, .TableCell(tcText, 1, 1)
                .TableBorder = tbNone
                .EndTable
            End If
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
    End If
    RS3.Close: Set RS3 = Nothing
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 6
    .TableCell(tcFontBold) = True
    .TableCell(tcColWidth, , 1) = 1900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 7500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcText, 2, 1) = "Alimentos & Bebidas":       .TableCell(tcText, 2, 2) = Format(sumCuentaAlimen, fg_Pict(9, vg_DPr))
    .TableCell(tcText, 3, 1) = "Limpieza  & Desechables":   .TableCell(tcText, 3, 2) = Format(sumCuentaLimDes, fg_Pict(9, vg_DPr))
    .TableCell(tcText, 4, 2) = String(15, "_")
    .TableCell(tcText, 5, 1) = "Totales Generales":         .TableCell(tcText, 5, 2) = Format(sumCuentaLimDes + sumCuentaAlimen, fg_Pict(9, vg_DPr))
    Print #1, " ": Print #1, " "
    Print #1, .TableCell(tcText, 2, 1) & "|"; .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|"; .TableCell(tcText, 3, 2)
    Print #1, .TableCell(tcText, 4, 2)
    Print #1, .TableCell(tcText, 5, 1) & "|"; .TableCell(tcText, 5, 2)
    .TableBorder = tbBottom
    .EndTable
    .EndDoc
    .EndDoc
    Close #1
End With
Preview.Show 1
Exit Function
Error_Toma4:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Toma5(SQL As String)
Dim i As Long
On Local Error GoTo Error_Toma5
MsgTitulo = "Toma de Inventario"
RS1.Open SQL, vg_db, adOpenStatic
If RS1.EOF Then
    RS1.Close: Set RS1 = Nothing
    MsgBox "No existe información...", vbExclamation + vbOKOnly, MsgTitulo
    Exit Function
End If
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
    .FontSize = 7
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Diferencias Físico v/s Sistema - Valorizado"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcFontSize, 1, 1, 3, 2) = 8: .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 2) = False
    .TableCell(tcText, 1, 1) = "Bodega": .TableCell(tcText, 1, 2) = Left(M_TomInv.Combo1(0).Text, 50)
    .TableCell(tcText, 2, 1) = "Toma de Inventario": .TableCell(tcText, 2, 2) = M_TomInv.Date1(0).Text
    .TableCell(tcText, 3, 1) = "Tipo de Producto": .TableCell(tcText, 3, 2) = IIf(I_TomInv.optTIPPRO(0).Value = True, Left(I_TomInv.Combo1(0).Text, 50), "Todos")
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 9: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 500: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1000: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Unidad"
    .TableCell(tcText, 1, 4) = "P.M.P."
    .TableCell(tcText, 1, 5) = "Stock Físico"
    .TableCell(tcText, 1, 6) = "Total Físico"
    .TableCell(tcText, 1, 7) = "Stock Sist."
    .TableCell(tcText, 1, 8) = "Total Sist."
    .TableCell(tcText, 1, 9) = "Diferencia"
    Print #1, .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8); "|" & .TableCell(tcText, 1, 9)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 9: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 500: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1000: .TableCell(tcAlign, , 9) = taRightTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = RS1!tin_codpro
            .TableCell(tcText, i, 2) = RS1!pro_nombre
            .TableCell(tcText, i, 3) = RS1!uni_nomcor
            .TableCell(tcText, i, 4) = Format(RS1!tin_propon, fg_Pict(9, vg_DPr))
            .TableCell(tcText, i, 5) = Format(RS1!tin_stofis, fg_Pict(9, vg_DCa))
            .TableCell(tcText, i, 6) = Format(Format(RS1!tin_stofis, fg_Pict(9, vg_DCa)) * Format(RS1!tin_propon, fg_Pict(9, vg_DPr)), fg_Pict(9, vg_DPr))
            .TableCell(tcText, i, 7) = Format(RS1!tin_stosis, fg_Pict(9, vg_DCa))
            .TableCell(tcText, i, 8) = Format(Format(RS1!tin_stosis, fg_Pict(9, vg_DCa)) * Format(RS1!tin_propon, fg_Pict(9, vg_DPr)), fg_Pict(9, vg_DPr))
            .TableCell(tcText, i, 9) = Format(RS1!tin_stofis - RS1!tin_stosis, fg_Pict(9, vg_DCa))
            Print #1, .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                      .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8); "|" & .TableCell(tcText, i, 9)

            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
    Close #1
End With
Preview.Show 1
Exit Function
Error_Toma5:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_SalBodega(casnom As String, regimen As String, codser As Long, strFecini As String, strFecter As String)
Dim i As Long, Nfecini As Date, Nfecter As Date, Casino As String, nomcas As String, codreg As Long, nomreg As String
Dim sqlSER As String, cantfija As Double, cantprodxdia As Double, fecha As Date, numFecha As Long
'On Local Error GoTo Error_SalBod
MsgTitulo = "Salida de Bodega"
fg_carga ""
Casino = Trim(Mid(casnom, 1, InStr(1, casnom, "|") - 1))
nomcas = Trim(Mid(casnom, InStr(1, casnom, "|") + 1, Len(casnom)))
codreg = Val(Mid(regimen, 1, InStr(1, regimen, "|") - 1))
nomreg = Trim(Mid(regimen, InStr(1, regimen, "|") + 1, Len(regimen)))
sqlSER = IIf(codser = 0, " ", " where ser_codigo=" & codser & " ")
Preview.Show
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
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
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = "Formato de Requisición"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Casino"
    .TableCell(tcText, 1, 2) = Casino & " " & nomcas
    .TableCell(tcText, 2, 1) = "Regimen"
    .TableCell(tcText, 2, 2) = codreg & " " & nomreg
    .TableCell(tcText, 3, 1) = "Rango Facha"
    .TableCell(tcText, 3, 2) = strFecini & " - " & strFecter
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    Nfecini = Format(strFecini, "dd/mm/yyyy")
    Nfecter = Format(strFecter, "dd/mm/yyyy")
    For fecha = Nfecini To Nfecter
        numFecha = Val(Format(fecha, "yyyy") & Right("0" & Format(fecha, "mm"), 2) & Right("0" & Format(fecha, "dd"), 2))
        Dim ndia As Long
        ndia = fg_NumDia(Trim(Left(fg_Fecha_Dia(Trim(Str(numFecha)), 2), Len(fg_Fecha_Dia(Trim(Str(numFecha)), 2)) - 2)))
        RS1.Open "Select ser_codigo, ser_nombre from a_servicio " & sqlSER & " order by ser_codigo", vg_db, adOpenStatic
        Do While Not RS1.EOF
            RS2.Open "Select distinct mi.min_fecmin, mi.min_codser, pro.pro_codigo, pro.pro_nombre, uni.uni_nombre, pro.pro_propon, " & _
                     "(red.red_canpro * mid.mid_numrac) AS cantprodxdia " & _
                     "From  b_productos pro, b_minuta mi, b_minutadet mid, b_receta rec, b_recetadet red, a_unidad uni " & _
                     "Where pro.pro_codigo=red.red_codpro And rec.rec_codigo=mid.mid_codrec " & _
                     "And rec.rec_codigo=red.red_codigo And mi.min_codigo=mid.mid_codigo " & _
                     "And pro.pro_coduni=uni.uni_codigo And mid.mid_tipmin='2' And mi.min_fecmin=" & numFecha & " " & _
                     "And mi.min_cencos='" & Trim(Casino) & "' And mi.min_codreg=" & codreg & " " & _
                     "And mi.min_codser=" & RS1!ser_codigo, vg_db, adOpenStatic
            RS3.Open "Select mif.mif_dianro, pro.pro_codigo, pro.pro_nombre, uni.uni_nombre, pro.pro_propon, mif.mif_canpro " & _
                     "From b_minutafija mif, b_productos pro, a_unidad uni " & _
                     "Where pro.pro_codigo=mif.mif_codpro And pro.pro_coduni=uni.uni_codigo " & _
                     "And mif.mif_cencos='" & Trim(Casino) & "' And mif.mif_codreg=" & codreg & " And mif.mif_codser=" & RS1!ser_codigo & " " & _
                     "And mif.mif_fecval=(select max(mif2.mif_fecval) From b_minutafija as mif2) And mif.mif_dianro=" & ndia & " " & _
                     "Group by mif.mif_dianro, pro.pro_codigo, pro.pro_nombre, uni.uni_nombre, pro.pro_propon, mif.mif_canpro", vg_db, adOpenStatic
            If Not RS2.EOF Or Not RS3.EOF Then
                .StartTable
                .TableCell(tcCols) = 4: .TableCell(tcRows) = 2
                .TableCell(tcFontBold, , 1, , 1) = True: .TableCell(tcFontBold, , 3, , 3) = True
                .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taLeftTop
                .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taLeftTop
                .TableCell(tcColWidth, , 4) = 4500: .TableCell(tcAlign, , 4) = taLeftTop
                .TableCell(tcBackColor, 2) = vbYellow:: .TableCell(tcRowHeight, 2) = 200
                .TableCell(tcText, 2, 1) = "Servicio"
                .TableCell(tcText, 2, 2) = RS1!ser_codigo & " " & RS1!ser_nombre
                .TableCell(tcText, 2, 3) = "Fecha"
                .TableCell(tcText, 2, 4) = fg_Ctod1(numFecha)
                Print #1, .TableCell(tcText, 2, 1) & " " & .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4)
                .TableBorder = tbNone
                .EndTable
                .StartTable
                .TableCell(tcCols) = 6: .TableCell(tcRows) = 2
                .TableCell(tcColWidth, , 1) = 150: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
                .TableCell(tcColWidth, , 3) = 5300: .TableCell(tcAlign, , 3) = taLeftTop
                .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
                .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
                .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taCenterTop
                .TableCell(tcText, 2, 2) = "Código"
                .TableCell(tcText, 2, 3) = "Descripción"
                .TableCell(tcText, 2, 4) = "Cantidad"
                .TableCell(tcText, 2, 5) = "Cant.Salida"
                .TableCell(tcText, 2, 6) = "Unidad"
                Print #1, .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4) & "|" & _
                          .TableCell(tcText, 2, 5) & "|" & .TableCell(tcText, 2, 6)
                .TableBorder = tbNone
                .EndTable
                .StartTable
                .TableCell(tcCols) = 6: .TableCell(tcRows) = 500
                .TableCell(tcColWidth, , 1) = 150: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
                .TableCell(tcColWidth, , 3) = 5300: .TableCell(tcAlign, , 3) = taLeftTop
                .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
                .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
                .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taCenterTop
                i = 1
                If Not RS2.EOF Then
                    .TableCell(tcText, 2, 2) = "Minuta Real": .TableCell(tcFontBold, 2, 2) = True
                    Print #1, .TableCell(tcText, 2, 2)
                    i = 3
                    Do While Not RS2.EOF
                        .TableCell(tcText, i, 2) = RS2!pro_codigo
                        .TableCell(tcText, i, 3) = RS2!pro_nombre
                        .TableCell(tcText, i, 4) = Format(Round(TipoDato(RS2!cantprodxdia, 0), vg_DCa), fg_Pict(9, vg_DCa))
                        .TableCell(tcText, i, 5) = "_________"
                        .TableCell(tcText, i, 6) = RS2!uni_nombre
                        Print #1, .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                                  .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6)
                        i = i + 1
                        RS2.MoveNext
                    Loop
                End If
                If Not RS3.EOF Then
                    i = i + 1
                    .TableCell(tcText, i, 2) = "Estructura Fija": .TableCell(tcFontBold, i, 2) = True
                    Print #1, .TableCell(tcText, i, 2)
                    i = i + 1
                    Do While Not RS3.EOF
                        .TableCell(tcText, i, 2) = RS3!pro_codigo
                        .TableCell(tcText, i, 3) = RS3!pro_nombre
                        .TableCell(tcText, i, 4) = Format(TipoDato(RS3!mif_canpro, 0), fg_Pict(9, vg_DCa))
                        .TableCell(tcText, i, 5) = "_________"
                        .TableCell(tcText, i, 6) = RS3!uni_nombre
                        Print #1, .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                                  .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6)
                        i = i + 1
                        RS3.MoveNext
                    Loop
                End If
                .TableCell(tcRows) = i
                .TableBorder = tbNone
                .EndTable
            End If
            RS2.Close: Set RS2 = Nothing
            RS3.Close: Set RS3 = Nothing
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
    Next fecha
     Close #1
    .EndDoc
End With
fg_descarga
Exit Function
Error_SalBod:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_SalBodega2(casnom As String, regimen As String, codser As Long, fecini As String, fecter As String)
Dim i As Long, Nfecini As Long, Nfecter As Long, Casino As String, nomcas As String, codreg As Long, nomreg As String
Dim sqlSER As String, cantfija As Double, cantprodxdia As Double, fecha As String
On Local Error GoTo Error_Requisicion
fg_carga ""
Casino = Trim(Mid(casnom, 1, InStr(1, casnom, "|") - 1))
nomcas = Trim(Mid(casnom, InStr(1, casnom, "|") + 1, Len(casnom)))
codreg = Val(Mid(regimen, 1, InStr(1, regimen, "|") - 1))
nomreg = Trim(Mid(regimen, InStr(1, regimen, "|") + 1, Len(regimen)))
sqlSER = IIf(codser = 0, " ", " where ser_codigo=" & codser & " ")
Preview.Show
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
    .ExportFile = App.Path & "\Reporte.rtf"
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
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = "Formato de Requisición"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Casino"
    .TableCell(tcText, 1, 2) = Casino & " " & nomcas
    .TableCell(tcText, 2, 1) = "Regimen"
    .TableCell(tcText, 2, 2) = codreg & " " & nomreg
    .TableCell(tcText, 3, 1) = "Rango Facha"
    .TableCell(tcText, 3, 2) = fecini & " - " & fecter
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    Nfecini = Val(Format(fecini, "yyyy") & Right("0" & Format(fecini, "mm"), 2) & Right("0" & Format(fecini, "dd"), 2))
    'Nfecter = Val(Format(fecter, "yyyy") & Right("0" & Format(fecter, "mm"), 2) & Right("0" & Format(fecter, "dd"), 2))
    ndia = 0
    'For Fecha = Nfecini To Nfecter
    fecha = fecini
     Do While CDate(fecha) <= CDate(fecter)
        Nfecini = Val(Format(fecha, "yyyy") & Right("0" & Format(fecha, "mm"), 2) & Right("0" & Format(fecha, "dd"), 2))
        'Dim ndia As Long
        'ndia = fg_NumDia(Trim(Left(fg_Fecha_Dia(Trim(Str(Fecha)), 2), Len(fg_Fecha_Dia(Trim(Str(Fecha)), 2)) - 2)))
        ndia = fg_NumDia(Trim(Left(fg_Fecha_Dia(Trim(Str(Nfecini)), 2), Len(fg_Fecha_Dia(Trim(Str(Nfecini)), 2)) - 2)))
        RS1.Open "Select ser_codigo, ser_nombre from a_servicio " & sqlSER & " order by ser_codigo", vg_db, adOpenStatic
        Do While Not RS1.EOF
            RS2.Open "Select distinct mi.min_fecmin, mi.min_codser, pro.pro_codigo, pro.pro_nombre, uni.uni_nombre, pro.pro_propon, " & _
                     "(red.red_canpro * mid.mid_numrac) AS cantprodxdia " & _
                     "From  b_productos pro, b_minuta mi, b_minutadet mid, b_receta rec, b_recetadet red, a_unidad uni " & _
                     "Where pro.pro_codigo=red.red_codpro And rec.rec_codigo=mid.mid_codrec " & _
                     "And rec.rec_codigo=red.red_codigo And mi.min_codigo=mid.mid_codigo " & _
                     "And pro.pro_coduni=uni.uni_codigo And mid.mid_tipmin='2' And mi.min_fecmin=" & fecha & " " & _
                     "And mi.min_cencos='" & Trim(Casino) & "' And mi.min_codreg=" & codreg & " " & _
                     "And mi.min_codser=" & RS1!ser_codigo, vg_db, adOpenStatic
            RS3.Open "Select mif.mif_dianro, pro.pro_codigo, pro.pro_nombre, uni.uni_nombre, pro.pro_propon, mif.mif_canpro " & _
                     "From b_minutafija mif, b_productos pro, a_unidad uni " & _
                     "Where pro.pro_codigo=mif.mif_codpro And pro.pro_coduni=uni.uni_codigo " & _
                     "And mif.mif_cencos='" & Trim(Casino) & "' And mif.mif_codreg=" & codreg & " And mif.mif_codser=" & RS1!ser_codigo & " " & _
                     "And mif.mif_fecval=(select max(mif2.mif_fecval) From b_minutafija as mif2) And mif.mif_dianro=" & ndia & " " & _
                     "Group by mif.mif_dianro, pro.pro_codigo, pro.pro_nombre, uni.uni_nombre, pro.pro_propon, mif.mif_canpro", vg_db, adOpenStatic
            If Not RS2.EOF Or Not RS3.EOF Then
                .StartTable
                .TableCell(tcCols) = 4: .TableCell(tcRows) = 2
                .TableCell(tcFontBold, , 1, , 1) = True: .TableCell(tcFontBold, , 3, , 3) = True
                .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taLeftTop
                .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taLeftTop
                .TableCell(tcColWidth, , 4) = 4500: .TableCell(tcAlign, , 4) = taLeftTop
                .TableCell(tcBackColor, 2) = vbYellow:: .TableCell(tcRowHeight, 2) = 200
                .TableCell(tcText, 2, 1) = "Servicio"
                .TableCell(tcText, 2, 2) = RS1!ser_codigo & " " & RS1!ser_nombre
                .TableCell(tcText, 2, 3) = "Fecha"
                .TableCell(tcText, 2, 4) = CDate(fecha)
                Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4)
                .TableBorder = tbNone
                .EndTable
                .StartTable
                .TableCell(tcCols) = 6: .TableCell(tcRows) = 2
                .TableCell(tcColWidth, , 1) = 150: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
                .TableCell(tcColWidth, , 3) = 5300: .TableCell(tcAlign, , 3) = taLeftTop
                .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
                .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
                .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taCenterTop
                .TableCell(tcText, 2, 2) = "Código"
                .TableCell(tcText, 2, 3) = "Descripción"
                .TableCell(tcText, 2, 4) = "Cantidad"
                .TableCell(tcText, 2, 5) = "Cant.Salida"
                .TableCell(tcText, 2, 6) = "Unidad"
                Print #1, .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4) & "|" & .TableCell(tcText, 2, 5) & "|" & .TableCell(tcText, 2, 6)
                .TableBorder = tbNone
                .EndTable
                .StartTable
                .TableCell(tcCols) = 6: .TableCell(tcRows) = 500
                .TableCell(tcColWidth, , 1) = 150: .TableCell(tcAlign, , 1) = taLeftTop
                .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
                .TableCell(tcColWidth, , 3) = 5300: .TableCell(tcAlign, , 3) = taLeftTop
                .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
                .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
                .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taCenterTop
                If Not RS2.EOF Then
                    .TableCell(tcText, 2, 2) = "Cantidad Calculada": .TableCell(tcFontBold, 2, 2) = True
                    Print #1, .TableCell(tcText, 2, 2)
                    i = 3
                    Do While Not RS2.EOF
                        .TableCell(tcText, i, 2) = RS2!pro_codigo
                        .TableCell(tcText, i, 3) = RS2!pro_nombre
                        .TableCell(tcText, i, 4) = Format(Round(TipoDato(RS2!cantprodxdia, 0), vg_DCa), fg_Pict(9, vg_DCa))
                        .TableCell(tcText, i, 5) = "_________"
                        .TableCell(tcText, i, 6) = RS2!uni_nombre
                        Print #1, .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|"; .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6)
                        i = i + 1
                        RS2.MoveNext
                    Loop
                End If
                If Not RS3.EOF Then
                    i = i + 1
                    .TableCell(tcText, i, 2) = "Cantidad Fija": .TableCell(tcFontBold, i, 2) = True
                    Print #1, .TableCell(tcText, i, 2)
                    i = i + 1
                    Do While Not RS3.EOF
                        .TableCell(tcText, i, 2) = RS3!pro_codigo
                        .TableCell(tcText, i, 3) = RS3!pro_nombre
                        .TableCell(tcText, i, 4) = Format(TipoDato(RS3!mif_canpro, 0), fg_Pict(9, vg_DCa))
                        .TableCell(tcText, i, 5) = "_________"
                        .TableCell(tcText, i, 6) = RS3!uni_nombre
                        Print #1, .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|"; .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6)
                        i = i + 1
                        RS3.MoveNext
                    Loop
                End If
                .TableCell(tcRows) = i
                .TableBorder = tbNone
                .EndTable
            End If
            RS2.Close: Set RS2 = Nothing
            RS3.Close: Set RS3 = Nothing
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing: Close #1
        fecha = Str(CDate(fecha) + 1)
    Loop
    .EndDoc
End With
fg_descarga
Exit Function
Error_Requisicion:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Ajuste(bodega As String, fecha As String)
Dim i As Long, codbod As Long, nombod As String
codbod = Trim(Mid(bodega, 1, InStr(1, bodega, "|") - 1))
nombod = Trim(Mid(bodega, InStr(1, bodega, "|") + 1, Len(bodega)))
On Local Error GoTo Error_Ajustes
MsgTitulo = "Informe de Ajustes"
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
    .FontSize = 8
    .PenColor = &HC0C0C0
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Ajuste de Inventario"
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcFontSize, 1, 1, 2, 2) = 8: .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 2) = False
    .TableCell(tcText, 1, 1) = "Bodega": .TableCell(tcText, 1, 2) = nombod
    .TableCell(tcText, 2, 1) = "Toma de Inventario": .TableCell(tcText, 2, 2) = fecha
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 5: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 5700: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 900: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1100: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Unidad"
    .TableCell(tcText, 1, 4) = "Diferencia"
    .TableCell(tcText, 1, 5) = "Precio"
    .TableBorder = tbNone
    .EndTable
    RS1.Open "select tov.tov_numdoc, aju.aju_nombre from b_totventas tov, a_tipoajuste aju " & _
             "where tov.tov_codser=aju.aju_codigo and tov.tov_fecemi=Cdate('" & fecha & "') and " & _
             "tov_codbod=" & codbod & " and tov.tov_tipdoc='AI' and tov.tov_estdoc<>'A' order by tov.tov_numdoc", vg_db, adOpenStatic
    Do While Not RS1.EOF
        .StartTable
        .TableCell(tcCols) = 4: .TableCell(tcRows) = 2
        .TableCell(tcColWidth, , 1) = 800: .TableCell(tcAlign, , 1) = taLeftTop
        .TableCell(tcColWidth, , 2) = 900: .TableCell(tcAlign, , 2) = taLeftTop
        .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taLeftTop
        .TableCell(tcColWidth, , 4) = 8000: .TableCell(tcAlign, , 4) = taLeftTop
        .TableCell(tcRowHeight, 2) = 230: .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 3) = True
        .TableCell(tcText, 2, 1) = "Folio"
        .TableCell(tcText, 2, 2) = RS1!tov_numdoc
        .TableCell(tcText, 2, 3) = "Concepto"
        .TableCell(tcText, 2, 4) = RS1!aju_nombre
        .TableBorder = tbBottom
        .EndTable
        RS2.Open "select dev.dev_codmer, pro.pro_nombre, dev.dev_precos, uni.uni_nombre, dev.dev_canmer " & _
                 "from b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
                 "where tov.tov_rutcli=dev.dev_rutcli and tov.tov_tipdoc=dev.dev_tipdoc and tov.tov_numdoc=dev.dev_numdoc " & _
                 "and pro.pro_codigo=dev.dev_codmer and uni.uni_codigo=pro.pro_coduni " & _
                 "and tov.tov_fecemi=Cdate('" & fecha & "') and tov_codbod=" & codbod & " and tov.tov_tipdoc='AI' " & _
                 "and tov.tov_estdoc<>'A' and tov.tov_numdoc=" & RS1!tov_numdoc & " order by dev.dev_numlin", vg_db, adOpenStatic
        .StartTable
        .TableCell(tcCols) = 5: .TableCell(tcRows) = 2000
        .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
        .TableCell(tcColWidth, , 2) = 5700: .TableCell(tcAlign, , 2) = taLeftTop
        .TableCell(tcColWidth, , 3) = 900: .TableCell(tcAlign, , 3) = taCenterTop
        .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
        .TableCell(tcColWidth, , 5) = 1100: .TableCell(tcAlign, , 5) = taRightTop
        i = 1
        Do While Not RS2.EOF
            .TableCell(tcText, i, 1) = RS2!dev_codmer
            .TableCell(tcText, i, 2) = RS2!pro_nombre
            .TableCell(tcText, i, 3) = RS2!uni_nombre
            .TableCell(tcText, i, 4) = Format(RS2!dev_canmer, fg_Pict(9, vg_DCa))
            .TableCell(tcText, i, 5) = Format(RS2!dev_precos, fg_Pict(9, vg_DPr))
            RS2.MoveNext: i = i + 1
        Loop
        RS2.Close: Set RS2 = Nothing
        .TableCell(tcRows) = i - 1
        '.PenColor = &HC0C0C0
        .TableBorder = tbNone
        .EndTable
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    .EndDoc
End With
Preview.Show 1
Exit Function
Error_Ajustes:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_StockxFecha(Formu As Form)
Dim i As Long, sumCuenta As Double, sumTipo As Double, sqlTP As String, sqlCU As String, v_codbod As Long
On Local Error GoTo Error_StockFecha
fg_carga ""
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
    .FontSize = 7
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Informe de Stock"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcFontSize, 1, 1, 1, 1) = 8: .TableCell(tcFontBold, , 1) = True ': .TableCell(tcFontBold, , 2) = False
    .TableCell(tcText, 1, 1) = "Bodega":           .TableCell(tcText, 1, 2) = Left(I_MovSto.Combo1(0).Text, 50)
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    .TableBorder = tbNone
    .EndTable
    v_codbod = fg_codigocbo(Formu.Combo1, 0, 10, 0)
    sqlTP = ""
    If Formu.optTIPPRO(0).Value = True Then sqlTP = "and pro.pro_codtip=" & Val(fg_codigocbo(Formu.Combo1, 1, 10, 0)) & " "
    sqlCU = ""
    If Formu.optCUENTA(0).Value = True Then sqlCU = "and pro.pro_ctacon='" & Trim(Mid(Trim(Formu.Combo1(2).List(Formu.Combo1(2).ListIndex)), Len(Trim(Formu.Combo1(2).List(Formu.Combo1(2).ListIndex))) - 10, 10)) & "' "
    RS1.Open "select cta.cta_codigo, cta.cta_nombre from a_ctacontable cta, b_bodegas bod, b_productos pro " & _
             "where cta.cta_codigo=pro.pro_ctacon and bod.bod_codpro=pro.pro_codigo " & sqlCU & _
             "and bod.bod_codbod=" & v_codbod & " group by cta.cta_codigo, cta.cta_nombre", vg_db, adOpenStatic
    If Not RS1.EOF Then
        .StartTable
        .TableCell(tcCols) = 6: .TableCell(tcRows) = 1
        .TableCell(tcColWidth, , 1) = 200: .TableCell(tcAlign, , 1) = taLeftTop
        .TableCell(tcColWidth, , 2) = 1700: .TableCell(tcAlign, , 2) = taLeftTop
        .TableCell(tcColWidth, , 3) = 5900: .TableCell(tcAlign, , 3) = taLeftTop
        .TableCell(tcColWidth, , 4) = 500: .TableCell(tcAlign, , 4) = taCenterTop
        .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
        .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
        .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 200
        .TableCell(tcText, 1, 2) = "Código"
        .TableCell(tcText, 1, 3) = "Descripción"
        .TableCell(tcText, 1, 4) = "Unidad"
        .TableCell(tcText, 1, 5) = "Stock"
        .TableCell(tcText, 1, 6) = "Precio"
        Print #1, .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6)
        .TableBorder = tbNone
        .EndTable
    End If
    Do While Not RS1.EOF
        sumCuenta = 0
        .StartTable
        .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
        .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftTop: .TableCell(tcFontBold, 2, 1) = True
        .TableCell(tcText, 2, 1) = RS1!cta_codigo & "  " & RS1!cta_nombre
        Print #1, .TableCell(tcText, 2, 1)
        .TableBorder = tbNone
        .EndTable
        RS2.Open "select tip.tip_codigo, tip.tip_nombre from a_tipopro tip, b_productos pro, b_bodegas bod " & _
                 "where pro.pro_codigo=bod.bod_codpro and tip.tip_codigo=pro.pro_codtip " & sqlTP & "and bod.bod_codbod=" & v_codbod & " and pro.pro_ctacon='" & Trim(RS1!cta_codigo) & "' group by tip.tip_codigo, tip.tip_nombre", vg_db, adOpenStatic
        Do While Not RS2.EOF
            .StartTable
            .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
            .TableCell(tcFontBold, 2, 1) = True
            .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftTop
            .TableCell(tcText, 2, 1) = RS2!tip_codigo & "  " & RS2!tip_nombre
            Print #1, .TableCell(tcText, 2, 1)
            .TableBorder = tbNone
            .EndTable
            .StartTable
            .TableCell(tcCols) = 6: .TableCell(tcRows) = 10000
            .TableCell(tcColWidth, , 1) = 200: .TableCell(tcAlign, , 1) = taLeftTop
            .TableCell(tcColWidth, , 2) = 1700: .TableCell(tcAlign, , 2) = taLeftTop
            .TableCell(tcColWidth, , 3) = 5900: .TableCell(tcAlign, , 3) = taLeftTop
            .TableCell(tcColWidth, , 4) = 500: .TableCell(tcAlign, , 4) = taCenterTop
            .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
            .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
            i = 1
            RS3.Open "select pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, bod.bod_canmer, pro.pro_propon " & _
                     "from b_productos pro, a_unidad uni, b_bodegas bod where pro.pro_codigo=bod.bod_codpro " & _
                     "and uni.uni_codigo=pro.pro_coduni and bod.bod_codbod=" & v_codbod & " and pro.pro_ctacon='" & Trim(RS1!cta_codigo) & "' " & _
                     "and pro.pro_codtip=" & RS2!tip_codigo & " order by pro_nombre", vg_db, adOpenStatic
            If Not RS3.EOF Then
                sumTipo = 0
                Do While Not RS3.EOF
                    .TableCell(tcText, i, 2) = RS3!pro_codigo
                    .TableCell(tcText, i, 3) = RS3!pro_nombre
                    .TableCell(tcText, i, 4) = RS3!uni_nomcor
                    .TableCell(tcText, i, 5) = Format(RS3!bod_canmer, fg_Pict(9, vg_DCa))
                    .TableCell(tcText, i, 6) = Format(RS3!pro_propon, fg_Pict(9, vg_DPr))
                    Print #1, .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6)
                    sumTipo = sumTipo + Round(RS3!pro_propon, vg_DPr)
                    RS3.MoveNext: i = i + 1
                Loop
                sumCuenta = sumCuenta + sumTipo
            Else
                i = i + 1
            End If
            RS3.Close: Set RS3 = Nothing
            RS2.MoveNext
            .TableCell(tcText, i + 1, 5) = "Total Familia"
            .TableCell(tcText, i + 1, 6) = Format(sumTipo, fg_Pict(9, vg_DPr))
            Print #1, .TableCell(tcText, i + 1, 5) & "|" & .TableCell(tcText, i + 1, 6)
            .TableCell(tcRows) = i + 2
            .TableBorder = tbNone
            .EndTable
        Loop
        .StartTable
        .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
        .TableCell(tcFontBold) = True: .TableCell(tcAlign) = taRightTop
        .TableCell(tcColWidth, , 1) = 9300: .TableCell(tcColWidth, , 2) = 1500
        .TableCell(tcText, 1, 1) = "Total Cuenta"
        .TableCell(tcText, 1, 2) = Format(sumCuenta, fg_Pict(9, vg_DPr))
        Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
        .TableBorder = tbNone
        .EndTable
        .StartTable
        .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
        .TableCell(tcFontBold) = True
        .TableCell(tcColWidth) = 10900: .TableCell(tcAlign) = taRightTop
        .TableCell(tcText, 1, 1) = String(144, "_")
        Print #1, .TableCell(tcText, 1, 1)
        .TableBorder = tbNone
        .EndTable
        RS1.MoveNext
        RS2.Close: Set RS2 = Nothing
    Loop
    RS1.Close: Set RS1 = Nothing
    .EndDoc
    Close #1
End With
Preview.Show 1
fg_descarga
Exit Function
Error_StockFecha:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Sub I_UniMed()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_UniEnv
MsgTitulo = "Informe de Unidades de Medida"
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
    .TableCell(tcText, 1, 1) = "Unidades de Medida"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    RS1.Open "select * from a_unidadmed order by unm_nombre", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Sub
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 4500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Nombre"
    .TableCell(tcText, 1, 3) = "Nombre Corto"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 4500: .TableCell(tcAlign, , 3) = taLeftTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!unm_codigo
        .TableCell(tcText, i, 2) = RS1!unm_nombre
        .TableCell(tcText, i, 3) = RS1!unm_nomcor
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
Error_UniEnv:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Public Function I_FacturaCli(Formu As I_FacCli, SQL As String)
Dim i As Long, rutcli As String, total As Double, totalgen As Double
On Local Error GoTo Error_Bodega
MsgTitulo = "Facturación Clientes"
fg_carga ""
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 13: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Facturación Clientes"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .Text = Chr(13): .Text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 4
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Casino"
    .TableCell(tcText, 1, 2) = Formu.fpText1(1).Text & " - " & Formu.fpayuda(1).Caption
    .TableCell(tcText, 2, 1) = "Regimen"
    .TableCell(tcText, 2, 2) = Trim(Str(Val(fg_codigocbo(Formu.Combo1, 0, 10, 0)))) & " - " & Trim(Left(Formu.Combo1(0).List(Formu.Combo1(0).ListIndex), 50))
    .TableCell(tcText, 3, 1) = "Rango Facha"
    .TableCell(tcText, 3, 2) = Formu.fpDateTime1(0).Text & " - " & Formu.fpDateTime1(1).Text
    '.TableCell(tcText, 4, 1) = "Servicio"
    '.TableCell(tcText, 4, 2) = IIf(Formu.optTIPSER(1) = True, "Todos", Trim(Left(Formu.Combo1(1).List(Formu.Combo1(1).ListIndex), 50)))
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    Print #1, .TableCell(tcText, 4, 1) & "|" & .TableCell(tcText, 4, 2)
    .TableBorder = tbNone
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 5500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = Space(10) & "Servicio"
    .TableCell(tcText, 1, 2) = "Nº Raciones"
    .TableCell(tcText, 1, 3) = "Precio"
    .TableCell(tcText, 1, 4) = "Total"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 5500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taRightTop
    i = 1
    rutcli = ""
    RS1.Open SQL, vg_db, adOpenStatic
    If RS1.EOF Then
        RS1.Close: Set RS1 = Nothing
        Close #1
        fg_descarga
        MsgBox "No existe información...", vbExclamation + vbOKOnly, MsgTitulo
        Exit Function
    End If
    If Not RS1.EOF Then
        total = 0
        totalgen = 0
        Do While Not RS1.EOF
            If rutcli <> RS1!cli_codigo Then
                i = i + 1
                .TableCell(tcText, i, 1) = RS1!cli_codigo & " - " & RS1!cli_nombre
                .TableCell(tcFontBold, i, 1) = True
                Print #1, .TableCell(tcText, i, 1)
                rutcli = RS1!cli_codigo
                i = i + 2
            End If
            .TableCell(tcText, i, 1) = Space(10) & RS1!ser_codigo & " - " & RS1!ser_nombre
            .TableCell(tcText, i, 2) = Space(10) & Format(RS1!cantidad, fg_Pict(9, 0))
            RS2.Open "Select   prv.prv_preven " & _
                     "From     b_preciovta prv " & _
                     "Where    prv.prv_cencos='" & Trim(Formu.fpText1(1).Text) & "' " & _
                     "And      prv.prv_codreg=" & Val(fg_codigocbo(Formu.Combo1, 0, 10, 0)) & " " & _
                     "And      prv.prv_codser=" & RS1!ser_codigo & " " & _
                     "And      prv.prv_fecvig=(select max(prv2.prv_fecvig) from b_preciovta prv2) " & _
                     "And      prv.prv_rutcli='" & RS1!cli_codigo & "'" & _
                     "" & _
                     "", vg_db, adOpenStatic
            If Not RS2.EOF Then
                .TableCell(tcText, i, 3) = Space(10) & Format(RS2!prv_preven, fg_Pict(9, vg_DPr))
                .TableCell(tcText, i, 4) = Space(10) & Format(RS2!prv_preven * RS1!cantidad, fg_Pict(9, 0))
                total = total + (RS2!prv_preven * RS1!cantidad)
            Else
                .TableCell(tcText, i, 3) = Space(10) & "0"
                .TableCell(tcText, i, 4) = Space(10) & "0"
            End If
            
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
            RS2.Close: Set RS2 = Nothing
            RS1.MoveNext
            If Not RS1.EOF Then rutcli2 = RS1!cli_codigo Else rutcli2 = ""
            If rutcli <> rutcli2 Or RS1.EOF Then
                .TableCell(tcFontUnderline, i, 4) = True
                i = i + 1
                '.TableCell(tcText, i, 1) = "Total"
                .TableCell(tcText, i, 4) = Format(total, fg_Pict(9, vg_DPr))
                .TableCell(tcFontBold, i) = True
                totalgen = totalgen + total
                total = 0
                Print #1, .TableCell(tcText, i, 4)
            End If
            i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcRows) = i
    .TableBorder = tbBottom
    .EndTable
    .StartTable
        .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 5500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 3) = "Total General"
    .TableCell(tcText, 1, 4) = Format(totalgen, fg_Pict(9, vg_DPr))
    Print #1, "||" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    .TableBorder = tbNone
    .EndTable
    Close #1
    .EndDoc
End With
fg_descarga
Preview.Show
Preview.Refresh
Exit Function
Error_Bodega:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Resume Next
    Close #1
    Exit Function
End Function

