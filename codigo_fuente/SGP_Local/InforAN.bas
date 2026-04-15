Attribute VB_Name = "InforAN"
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset
Dim RS5 As New ADODB.Recordset
Dim RS6 As New ADODB.Recordset
Dim numlin As Integer
Dim inf_ncasino As String, inf_nregimen As String, inf_detaporte As String, inf_ndia As String
Dim inf_nreceta As String, cdetalle As String, opcionsalto As String
Dim inf_opcion As Integer

Public Function I_CatDie()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_CatDie
fg_carga ""
MsgTitulo = "Informe de Categorias Dieteticas"
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
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.CategoriaDietetica(1, 0, ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
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
        RS2.Open RutinaLectura.CategoriaDietetica(6, RS1!car_codigo, ""), vg_db, adOpenStatic
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
Preview.Show 1
fg_descarga
Exit Function
Error_CatDie:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_TipPla()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_TipoPla
fg_carga ""
MsgTitulo = "Informe de Tipos de Plato"
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
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.RecetaTipoPlato(1, 0, ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
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
        RS2.Open RutinaLectura.RecetaTipoPlato(7, RS1!tip_codigo, ""), vg_db, adOpenStatic
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
fg_descarga
Preview.Show 1
Exit Function
Error_TipoPla:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Regime()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Regimen
fg_carga ""
MsgTitulo = "Informe de Regimen"
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
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.Regimen(1, 0, ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
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
Preview.Show 1
fg_descarga
Exit Function
Error_Regimen:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Servic(modpac As Boolean)
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Servicios
fg_carga ""
MsgTitulo = "Informe de Servicios"
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
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.Servicio(6, 0, ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
    .StartTable
    .TableCell(tcCols) = 9: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 800: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2400: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 700: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taLeftTop
    .TableCell(tcColWidth, , 6) = 700: .TableCell(tcAlign, , 6) = taLeftTop
    .TableCell(tcColWidth, , 7) = 1500: .TableCell(tcAlign, , 7) = taLeftTop
    .TableCell(tcColWidth, , 8) = 1500: .TableCell(tcAlign, , 8) = taLeftTop
    .TableCell(tcColWidth, , 9) = 1500: .TableCell(tcAlign, , 9) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Orden"
    .TableCell(tcText, 1, 4) = "Código SAP"
    .TableCell(tcText, 1, 5) = "Facturable"
    .TableCell(tcText, 1, 6) = "Activo"
    .TableCell(tcText, 1, 7) = IIf(modpac, "Hr. Tope Cobro", "")
    .TableCell(tcText, 1, 8) = IIf(modpac, "Hora Entrega", "")
    .TableCell(tcText, 1, 9) = IIf(modpac, "Hr. Ultima Modif. PDA", "")
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & .TableCell(tcText, 1, 9)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 9: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 800: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2400: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 700: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taLeftTop
    .TableCell(tcColWidth, , 6) = 700: .TableCell(tcAlign, , 6) = taLeftTop
    .TableCell(tcColWidth, , 7) = 1500: .TableCell(tcAlign, , 7) = taLeftTop
    .TableCell(tcColWidth, , 8) = 1500: .TableCell(tcAlign, , 8) = taLeftTop
    .TableCell(tcColWidth, , 9) = 1500: .TableCell(tcAlign, , 9) = taLeftTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = RS1!ser_codigo
            .TableCell(tcText, i, 2) = Trim(RS1!ser_nombre)
            .TableCell(tcText, i, 3) = RS1!ser_orden
            .TableCell(tcText, i, 4) = IIf(IsNull(RS1!ser_codsap), "", Trim(RS1!ser_codsap))
            .TableCell(tcText, i, 5) = IIf(IsNull(RS1!ser_facturable) Or Trim(RS1!ser_facturable = "") Or RS1!ser_facturable = "0", "NO", "SI")
            .TableCell(tcText, i, 6) = IIf(IsNull(RS1!ser_activo) Or Trim(RS1!ser_activo = "") Or RS1!ser_activo = "0", "NO", "SI")
            .TableCell(tcText, i, 7) = IIf(modpac, IIf(IsNull(RS1!ser_horcob) Or Trim(RS1!ser_horcob) = "", "", RS1!ser_horcob), "")
            .TableCell(tcText, i, 8) = IIf(modpac, IIf(IsNull(RS1!ser_horent) Or Trim(RS1!ser_horent) = "", "", RS1!ser_horent), "")
            .TableCell(tcText, i, 9) = IIf(modpac, IIf(IsNull(RS1!ser_horpda), "", RS1!ser_horpda), "")
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9)
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
Preview.Show 1
fg_descarga
Exit Function
Error_Servicios:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_EstructuraServicio(cencos As String, codser As Long)
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long

On Local Error GoTo Error_EstServicios

fg_carga ""
MsgTitulo = "Informe de Estructura Servicios"
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    '------- Leer servicio
    RS1.Open RutinaLectura.Servicio(8, codser, ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Estructura Servicio : " & RS1!ser_nombre & " (" & RS1!ser_codigo & ")"
    RS1.Close: Set RS1 = Nothing
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 8000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcFontBold, 1, 1) = True: .TableCell(tcText, 1, 1) = "Contrato"
    .TableCell(tcFontBold, 2, 1) = True: .TableCell(tcText, 2, 1) = ""
    RS1.Open RutinaLectura.Cliente(1, cencos, ""), vg_db, adOpenStatic
    If Not RS1.EOF Then .TableCell(tcFontBold, 1, 2) = True: .TableCell(tcText, 1, 2) = ": " & Trim(RS1!cli_codigo) & " " & Trim(RS1!cli_nombre)
    RS1.Close: Set RS1 = Nothing
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
'    RS1.Open "SELECT a.ess_codser, a.ess_codigo, a.ess_nombre, a.ess_orden, a.ess_codsec, b.sec_nombre FROM a_estservicio a LEFT JOIN a_sector b ON a.ess_codsec = b.sec_codigo WHERE a.ess_codser = " & codser & " AND a.ess_cencos = '" & cencos & "' ORDER BY a.ess_orden, a.ess_nombre", vg_db, adOpenStatic
     Set RS1 = vg_db.Execute("sgp_Sel_DetalleEstructuraServicio '" & cencos & "', " & codser & "")
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
    .StartTable
    .TableCell(tcCols) = 5: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 3500: .TableCell(tcAlign, , 5) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Orden"
    .TableCell(tcText, 1, 4) = "Sector"
    .TableCell(tcText, 1, 5) = "Descripción"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 5: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 3500: .TableCell(tcAlign, , 5) = taLeftTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = RS1!ess_codigo
            .TableCell(tcText, i, 2) = RS1!ess_nombre
            .TableCell(tcText, i, 3) = RS1!ess_orden
            .TableCell(tcText, i, 4) = IIf(IsNull(RS1!ess_codsec) Or RS1!ess_codsec = 0, "", RS1!ess_codsec)
            .TableCell(tcText, i, 5) = IIf(IsNull(RS1!sec_nombre), "", Trim(RS1!sec_nombre))
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5)
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
Preview.Show 1
fg_descarga
Exit Function
Error_EstServicios:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_LPCafeteria()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Servicios
fg_carga ""
MsgTitulo = "Lista de Precio Cafetería"
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
    .FontSize = 13
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 16: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Lista de precios cafetería"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13): .text = Chr(13)
    RS1.Open "SELECT * FROM b_totpreciocaf WHERE tpc_cencos = '" & MuestraCasino(1) & "' ORDER BY tpc_codigo, tpc_nombre", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
    .TextAlign = taCenterMiddle
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 7000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 320
    .TableCell(tcText, 1, 1) = "Artículo de cafetería"
    .TableCell(tcText, 1, 2) = "Precio"
    .TableCell(tcText, 1, 3) = "Activo"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 7000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taCenterTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = Trim(RS1!tpc_nombre)
            .TableCell(tcText, i, 2) = Format(RS1!tpc_precio, fg_Pict(9, vg_DPr))
            .TableCell(tcText, i, 3) = IIf(RS1!tpc_activo = "1", "SI", "NO")
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i - 1
    .PenColor = &H8000000C
    
    .TableBorder = tbAll
    .EndTable
    .EndDoc
End With
Preview.Show 1
fg_descarga
Exit Function
Error_Servicios:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_LPCafeteriaDet()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long, cCod2 As String
On Local Error GoTo Error_Servicios
fg_carga ""
MsgTitulo = "Lista de Precio Cafetería"
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
    .FontSize = 13
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 16: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Lista de precios cafetería con composición"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13): .text = Chr(13)
    RS1.Open "SELECT b_totpreciocaf.tpc_codigo, b_totpreciocaf.tpc_nombre, b_totpreciocaf.tpc_precio, b_totpreciocaf.tpc_activo, b_detpreciocaf.dpc_codmer, b_productos.pro_nombre, a_unidad.uni_nomcor, b_detpreciocaf.dpc_cantidad " & _
             "FROM a_unidad RIGHT JOIN (b_productos RIGHT JOIN (b_detpreciocaf RIGHT JOIN b_totpreciocaf ON b_detpreciocaf.dpc_cencos = b_totpreciocaf.tpc_cencos AND b_detpreciocaf.dpc_codigo = b_totpreciocaf.tpc_codigo) ON b_productos.pro_codigo = b_detpreciocaf.dpc_codmer) " & _
             "ON a_unidad.uni_codigo=b_productos.pro_coduni WHERE b_totpreciocaf.tpc_cencos = '" & MuestraCasino(1) & "' ORDER BY b_totpreciocaf.tpc_codigo, b_detpreciocaf.dpc_codmer", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
    .TextAlign = taCenterMiddle
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 8000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 320
    .TableCell(tcText, 1, 1) = "Artículo de cafetería"
    .TableCell(tcText, 1, 2) = "Precio"
    .TableCell(tcText, 1, 3) = "Activo"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 6: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 4500: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taRightMiddle
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taLeftMiddle
    .TableCell(tcColWidth, , 5) = 2000: .TableCell(tcAlign, , 5) = taRightMiddle
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taCenterMiddle
    i = 1: cCod2 = ""
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            If cCod2 <> Trim(RS1!tpc_codigo) Then
                .TableCell(tcColSpan, i, 1) = 4
                .TableCell(tcText, i, 1) = Trim(RS1!tpc_nombre)
                .TableCell(tcRowHeight, i, 1) = 450
                .TableCell(tcText, i, 5) = Format(RS1!tpc_precio, fg_Pict(9, vg_DPr))
                .TableCell(tcText, i, 6) = IIf(RS1!tpc_activo = "1", "SI", "NO")
                cCod2 = Trim(RS1!tpc_codigo)
                Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6)
                i = i + 1
            End If
            If Not IsNull(RS1!dpc_codmer) Then
                .TableCell(tcFontSize, i) = 8
                .TableCell(tcText, i, 1) = "      " & Trim(RS1!dpc_codmer)
                .TableCell(tcText, i, 2) = Trim(RS1!pro_nombre)
                .TableCell(tcText, i, 3) = Format(RS1!dpc_cantidad, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
                .TableCell(tcText, i, 4) = Trim(RS1!uni_nomcor)
                Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5)
                i = i + 1
            End If
            RS1.MoveNext
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i - 1
    .PenColor = &H8000000C
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Preview.Show 1
fg_descarga
Exit Function
Error_Servicios:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_CompLPCafeteria(codigo As Long)
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_EstServicios
fg_carga ""
MsgTitulo = "Lista de Precio Cafetería - Composición"
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    '------- Leer servicio
    RS1.Open "SELECT * FROM b_totpreciocaf WHERE tpc_cencos = '" & MuestraCasino(1) & "' AND tpc_codigo = '" & codigo & "'", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Composición artículo de cafetería : " & Trim(RS1!tpc_nombre) & " (" & Trim(RS1!tpc_codigo) & ")"
    RS1.Close: Set RS1 = Nothing
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open "SELECT dpc.*, pro.pro_nombre, uni.uni_nomcor FROM b_detpreciocaf dpc, b_productos pro, a_unidad uni WHERE pro.pro_codigo = dpc.dpc_codmer AND pro.pro_coduni = uni.uni_codigo AND dpc.dpc_cencos = '" & MuestraCasino(1) & "' AND dpc.dpc_codigo = '" & codigo & "'", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 2500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 5000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código producto"
    .TableCell(tcText, 1, 2) = "Nombre"
    .TableCell(tcText, 1, 3) = "Unidad"
    .TableCell(tcText, 1, 4) = "Cantidad"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 2500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 5000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taRightTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = Trim(RS1!dpc_codmer)
            .TableCell(tcText, i, 2) = Trim(RS1!pro_nombre)
            .TableCell(tcText, i, 3) = Trim(RS1!uni_nomcor)
            .TableCell(tcText, i, 4) = Format(RS1!dpc_cantidad, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
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
Preview.Show 1
fg_descarga
Exit Function
Error_EstServicios:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_VenCafArt(cForm As Form)
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long, sql1 As String, sql2 As String
Dim total As Double
On Local Error GoTo Error_Servicios
fg_carga ""
MsgTitulo = Trim(cForm.Caption)
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 13: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = Trim(cForm.Caption)
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    sql1 = IIf(vg_tipbase = "1", " cdate('" & Trim(cForm.fpDateTime1(0).text) & "') ", " '" & Format(Trim(cForm.fpDateTime1(0).text), "yyyymmdd") & "' ")
    sql2 = IIf(vg_tipbase = "1", " cdate('" & Trim(cForm.fpDateTime1(1).text) & "') ", " '" & Format(Trim(cForm.fpDateTime1(1).text), "yyyymmdd") & "' ")
    RS1.Open "SELECT    a.dvc_articulo, c.tpc_nombre, a.dvc_precio, SUM(a.dvc_canart) AS dvc_canart " & _
             "FROM      b_detventascaf a, b_totventascaf b, b_totpreciocaf c " & _
             "WHERE     b.tvc_cencos = a.dvc_cencos " & _
             "AND       b.tvc_fecing = a.dvc_fecing " & _
             "AND       b.tvc_cencos = c.tpc_cencos " & _
             "AND       a.dvc_articulo = c.tpc_codigo " & _
             "AND       b.tvc_estado = 'C' AND b.tvc_cencos = '" & Trim(cForm.fpText1(0).text) & "' " & _
             "AND      (a.dvc_fecing >= " & sql1 & " and a.dvc_fecing <= " & sql2 & ") " & _
             "AND       b.tvc_codbod = " & Val(fg_codigocbo(cForm.Combo1, 0, 10, "")) & " " & _
             "AND      (a.dvc_articulo = '" & Trim(cForm.fpText1(2).text) & "' OR '" & Trim(cForm.fpText1(2).text) & "' = '') " & _
             "GROUP BY  a.dvc_articulo, c.tpc_nombre, a.dvc_precio " & _
             "ORDER BY  a.dvc_articulo", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 4
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Contrato"
    .TableCell(tcText, 1, 2) = Trim(cForm.fpText1(0).text) & " - " & Trim(cForm.fpayuda(0).Caption)
    .TableCell(tcText, 2, 1) = "Bodega"
    .TableCell(tcText, 2, 2) = Trim(Left(cForm.Combo1(0).text, 50))
    .TableCell(tcText, 3, 1) = "Periodo"
    .TableCell(tcText, 3, 2) = Trim(cForm.fpDateTime1(0).text) & " - " & Trim(cForm.fpDateTime1(1).text)
    .TableCell(tcText, 4, 1) = "Articulo de cafetería"
    .TableCell(tcText, 4, 2) = IIf(cForm.OptTipCli(3).Value = True, "Todos", Trim(cForm.fpText1(2).text) & " " & Trim(cForm.fpayuda(2).Caption))
    
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    Print #1, .TableCell(tcText, 4, 1) & "|" & .TableCell(tcText, 4, 2)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 4000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 2500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 240
    .TableCell(tcText, 1, 1) = "Artículo de cafetería"
    .TableCell(tcText, 1, 2) = "Cantidad"
    .TableCell(tcText, 1, 3) = "Precio"
    .TableCell(tcText, 1, 4) = "Total"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 4000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 2500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taRightTop
    i = 1: total = 0
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = Trim(RS1!tpc_nombre)
            .TableCell(tcText, i, 2) = Format(RS1!dvc_canart, fg_Pict(9, vg_DCa))
            .TableCell(tcText, i, 3) = Format(RS1!dvc_precio, fg_Pict(9, vg_DPr))
            .TableCell(tcText, i, 4) = Format(RS1!dvc_canart * RS1!dvc_precio, fg_Pict(9, vg_DPr))
            total = total + Format(RS1!dvc_canart * RS1!dvc_precio, fg_Pict(9, vg_DPr))
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcFontBold, i, 3, i, 4) = True
    .TableCell(tcText, i, 3) = "Total"
    .TableCell(tcText, i, 4) = Format(total, fg_Pict(9, vg_DPr))
    Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
    .TableCell(tcRows) = i
    .PenColor = &H8000000C
    .TableBorder = tbAll
    .EndTable
    .EndDoc
    Close #1
End With
Preview.Show 1
fg_descarga
Exit Function
Error_Servicios:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_VenCafCli(cForm As Form)
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim total As Double, crut As String, sql1 As String, sql2 As String
On Local Error GoTo Error_Servicios
fg_carga ""
MsgTitulo = Trim(cForm.Caption)
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 13: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = Trim(cForm.Caption)
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    crut = cForm.fpText1(1).text
    If InStr(Trim(crut), "-") = 0 And (cForm.lc_Aux = "VenCaf2" Or cForm.lc_Aux = "VenCaf3") And cForm.OptTipCli(0).Value Then crut = fg_RutDig(Trim(crut))
    sql1 = IIf(vg_tipbase = "1", " cdate('" & Trim(cForm.fpDateTime1(0).text) & "') ", " '" & Format(Trim(cForm.fpDateTime1(0).text), "yyyymmdd") & "' ")
    sql2 = IIf(vg_tipbase = "1", " cdate('" & Trim(cForm.fpDateTime1(1).text) & "') ", " '" & Format(Trim(cForm.fpDateTime1(1).text), "yyyymmdd") & "' ")
    RS1.Open "select    dvc_rutcli, cli_nombre, dvc_cencli, sum(dvc_canart * dvc_precio) as total " & _
             "from      b_detventascaf, b_totventascaf, b_clientes " & _
             "where     tvc_cencos = dvc_cencos " & _
             "and       tvc_fecing = dvc_fecing " & _
             "and       dvc_rutcli = cli_codigo " & _
             "and       tvc_estado = 'C' and tvc_cencos = '" & Trim(cForm.fpText1(0).text) & "' " & _
             "and       (dvc_fecing >= " & sql1 & " and dvc_fecing <= " & sql2 & ") " & _
             "and       tvc_codbod = " & Val(fg_codigocbo(cForm.Combo1, 0, 10, "")) & " " & _
             "and       (dvc_rutcli = '" & fg_DespintaRut(Trim(crut)) & "' or '" & Trim(crut) & "' = '') " & _
             "group by  dvc_rutcli, cli_nombre, dvc_cencli " & _
             "order by  dvc_rutcli, cli_nombre, dvc_cencli", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
    
    'dvc_cencos  dvc_fecing  dvc_numlin  dvc_articulo    dvc_canart  dvc_precio  dvc_tippag  dvc_rutcli  dvc_cencli  dvc_tipdoc  dvc_numdoc  dvc_fecdoc
    '24570       03/08/2005  1   1                       10  550 CR  190                         0   12:00:00 AM
    
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 4
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Contrato"
    .TableCell(tcText, 1, 2) = Trim(cForm.fpText1(0).text) & " - " & Trim(cForm.fpayuda(0).Caption)
    .TableCell(tcText, 2, 1) = "Bodega"
    .TableCell(tcText, 2, 2) = Trim(Left(cForm.Combo1(0).text, 50))
    .TableCell(tcText, 3, 1) = "Periodo"
    .TableCell(tcText, 3, 2) = Trim(cForm.fpDateTime1(0).text) & " - " & Trim(cForm.fpDateTime1(1).text)
    .TableCell(tcText, 4, 1) = "Cliente"
    .TableCell(tcText, 4, 2) = IIf(cForm.OptTipCli(1).Value = True, "Todos", fg_PintaRut(Trim(crut)) & " " & Trim(cForm.fpayuda(1).Caption))
    '.TableCell(tcText, 5, 1) = "Articulo de cafetería"
    '.TableCell(tcText, 5, 2) = IIf(cForm.OptTipCli(3).Value = True, "Todos", Trim(cForm.fpText1(2).Text) & " " & Trim(cForm.fpayuda(2).Caption))
    
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    Print #1, .TableCell(tcText, 4, 1) & "|" & .TableCell(tcText, 4, 2)
    'Print #1, .TableCell(tcText, 5, 1) & "|" & .TableCell(tcText, 5, 2)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 5500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 240
    .TableCell(tcText, 1, 1) = "Cliente"
    .TableCell(tcText, 1, 2) = "Centro de costo"
    .TableCell(tcText, 1, 3) = "Precio"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 5500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 2000: .TableCell(tcAlign, , 3) = taRightTop
    i = 1: total = 0
    'dvc_rutcli, cli_nombre, dvc_cencli, sum(dvc_canart * dvc_precio) as total
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = fg_PintaRut(Trim(RS1!dvc_rutcli)) & "  " & Trim(RS1!cli_nombre)
            .TableCell(tcText, i, 2) = Trim(RS1!dvc_cencli)
            .TableCell(tcText, i, 3) = Format(RS1!total, fg_Pict(9, vg_DPr))
            total = total + Format(RS1!total, fg_Pict(9, vg_DPr))
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcFontBold, i, 2, i, 3) = True
    .TableCell(tcText, i, 2) = "Total"
    .TableCell(tcText, i, 3) = Format(total, fg_Pict(9, vg_DPr))
    Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
    .TableCell(tcRows) = i
    .PenColor = &H8000000C
    .TableBorder = tbAll
    .EndTable
    .EndDoc
    Close #1
End With
Preview.Show 1
fg_descarga
Exit Function
Error_Servicios:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_VenCafCliArt(cForm As Form)
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim total As Double, subtotal As Double, crut As String, cRut2 As String, cCen2 As String, sql1 As String, sql2 As String
On Local Error GoTo Error_Servicios
fg_carga ""
MsgTitulo = Trim(cForm.Caption)
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 13: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = Trim(cForm.Caption)
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    crut = cForm.fpText1(1).text
    If InStr(Trim(crut), "-") = 0 And (cForm.lc_Aux = "VenCaf2" Or cForm.lc_Aux = "VenCaf3") And cForm.OptTipCli(0).Value Then crut = fg_RutDig(Trim(crut))
    sql1 = IIf(vg_tipbase = "1", " cdate('" & Trim(cForm.fpDateTime1(0).text) & "') ", " '" & Format(Trim(cForm.fpDateTime1(0).text), "yyyymmdd") & "' ")
    sql2 = IIf(vg_tipbase = "1", " cdate('" & Trim(cForm.fpDateTime1(1).text) & "') ", " '" & Format(Trim(cForm.fpDateTime1(1).text), "yyyymmdd") & "' ")
    RS1.Open "SELECT    a.dvc_rutcli, c.cli_nombre, a.dvc_cencli, a.dvc_articulo, d.tpc_nombre, a.dvc_precio, SUM(a.dvc_canart) AS dvc_canart " & _
             "FROM      b_detventascaf a, b_totventascaf b, b_clientes c, b_totpreciocaf d " & _
             "WHERE     b.tvc_cencos = a.dvc_cencos " & _
             "AND       b.tvc_fecing = a.dvc_fecing " & _
             "AND       a.dvc_rutcli = c.cli_codigo " & _
             "AND       b.tvc_cencos = d.tpc_cencos " & _
             "AND       a.dvc_articulo = d.tpc_codigo " & _
             "AND       b.tvc_estado = 'C' AND b.tvc_cencos = '" & Trim(cForm.fpText1(0).text) & "' " & _
             "AND      (a.dvc_fecing >= " & sql1 & " AND a.dvc_fecing <= " & sql2 & ") " & _
             "AND       b.tvc_codbod = " & Val(fg_codigocbo(cForm.Combo1, 0, 10, "")) & " " & _
             "AND      (a.dvc_rutcli = '" & fg_DespintaRut(Trim(crut)) & "' OR '" & Trim(crut) & "' = '') " & _
             "AND      (a.dvc_articulo = '" & Trim(cForm.fpText1(2).text) & "' OR '" & Trim(cForm.fpText1(2).text) & "' = '') " & _
             "GROUP BY  a.dvc_rutcli, c.cli_nombre, a.dvc_cencli, a.dvc_articulo, d.tpc_nombre, a.dvc_precio " & _
             "ORDER BY  a.dvc_rutcli, c.cli_nombre, a.dvc_cencli, a.dvc_articulo, d.tpc_nombre", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
        
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 5
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Contrato"
    .TableCell(tcText, 1, 2) = Trim(cForm.fpText1(0).text) & " - " & Trim(cForm.fpayuda(0).Caption)
    .TableCell(tcText, 2, 1) = "Bodega"
    .TableCell(tcText, 2, 2) = Trim(Left(cForm.Combo1(0).text, 50))
    .TableCell(tcText, 3, 1) = "Periodo"
    .TableCell(tcText, 3, 2) = Trim(cForm.fpDateTime1(0).text) & " - " & Trim(cForm.fpDateTime1(1).text)
    .TableCell(tcText, 4, 1) = "Cliente"
    .TableCell(tcText, 4, 2) = IIf(cForm.OptTipCli(1).Value = True, "Todos", fg_PintaRut(Trim(crut)) & " " & Trim(cForm.fpayuda(1).Caption))
    .TableCell(tcText, 5, 1) = "Articulo de cafetería"
    .TableCell(tcText, 5, 2) = IIf(cForm.OptTipCli(3).Value = True, "Todos", Trim(cForm.fpText1(2).text) & " " & Trim(cForm.fpayuda(2).Caption))
    
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    Print #1, .TableCell(tcText, 4, 1) & "|" & .TableCell(tcText, 4, 2)
    Print #1, .TableCell(tcText, 5, 1) & "|" & .TableCell(tcText, 5, 2)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 4000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 2500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 240
    .TableCell(tcText, 1, 1) = "Artículo de cafetería"
    .TableCell(tcText, 1, 2) = "Cantidad"
    .TableCell(tcText, 1, 3) = "Precio"
    .TableCell(tcText, 1, 4) = "Total"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 4000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 2500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 2500: .TableCell(tcAlign, , 4) = taRightTop
    i = 1: total = 0
    'dvc_rutcli, cli_nombre, dvc_cencli, sum(dvc_canart * dvc_precio) as total
    If Not RS1.EOF Then
        cRut2 = "": cCen2 = "VALOR NULO"
        Do While Not RS1.EOF
            If Trim(cRut2) <> Trim(RS1!dvc_rutcli) Or cCen2 <> Trim(RS1!dvc_cencli) Then
                If cRut2 <> "" Then
                    .TableCell(tcFontBold, i, 3, i, 4) = True
                    .TableCell(tcText, i, 3) = "Total"
                    .TableCell(tcText, i, 4) = Format(totalsub, fg_Pict(9, vg_DPr))
                    Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
                    i = i + 2
                End If
                .TableCell(tcColSpan, i, 1) = 2
                .TableCell(tcFontBold, i, 1) = True
                .TableCell(tcAlign, i, 1) = taLeftMiddle
                .TableCell(tcText, i, 1) = "Cliente : " & fg_PintaRut(Trim(RS1!dvc_rutcli)) & "  " & Trim(RS1!cli_nombre)
                
                .TableCell(tcColSpan, i, 3) = 2
                .TableCell(tcFontBold, i, 3) = True
                .TableCell(tcAlign, i, 3) = taLeftMiddle
                .TableCell(tcText, i, 3) = "Centro de costo : " & Trim(RS1!dvc_cencli)

                i = i + 1
                cRut2 = Trim(RS1!dvc_rutcli)
                cCen2 = Trim(RS1!dvc_cencli)
                totalsub = 0
            End If
            
            .TableCell(tcText, i, 1) = Trim(RS1!tpc_nombre)
            .TableCell(tcText, i, 2) = Format(RS1!dvc_canart, fg_Pict(9, vg_DCa))
            .TableCell(tcText, i, 3) = Format(RS1!dvc_precio, fg_Pict(9, vg_DPr))
            .TableCell(tcText, i, 4) = Format(RS1!dvc_canart * RS1!dvc_precio, fg_Pict(9, vg_DPr))
            total = total + Format(RS1!dvc_canart * RS1!dvc_precio, fg_Pict(9, vg_DPr))
            totalsub = totalsub + Format(RS1!dvc_canart * RS1!dvc_precio, fg_Pict(9, vg_DPr))
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcFontBold, i, 3, i, 4) = True
    .TableCell(tcText, i, 3) = "Total"
    .TableCell(tcText, i, 4) = Format(totalsub, fg_Pict(9, vg_DPr))
    Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
    i = i + 2
    .TableCell(tcFontBold, i, 3, i, 4) = True
    .TableCell(tcText, i, 3) = "Total General"
    .TableCell(tcText, i, 4) = Format(total, fg_Pict(9, vg_DPr))
    Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
    .TableCell(tcRows) = i
    .PenColor = &H8000000C
    .TableBorder = tbNone
    .EndTable
    .EndDoc
    Close #1
End With
Preview.Show 1
fg_descarga
Exit Function
Error_Servicios:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_VenCafPro(cForm As Form)
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long, sql1 As String, sql2 As String
Dim total As Double, subtotal As Double
On Local Error GoTo Error_Servicios
fg_carga ""
MsgTitulo = Trim(cForm.Caption)
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 13: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = Trim(cForm.Caption)
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    sql1 = IIf(vg_tipbase = "1", " cdate('" & Trim(cForm.fpDateTime1(0).text) & "') ", " '" & Format(Trim(cForm.fpDateTime1(0).text), "yyyymmdd") & "' ")
    sql2 = IIf(vg_tipbase = "1", " cdate('" & Trim(cForm.fpDateTime1(1).text) & "') ", " '" & Format(Trim(cForm.fpDateTime1(1).text), "yyyymmdd") & "' ")
    RS1.Open "SELECT    pro_codigo, pro_nombre, uni_nomcor, SUM(dvp_candig) AS candig, SUM(dvp_candig*dvp_precos) AS total " & _
             "FROM      b_detventascafpro, b_totventascaf, b_productos, a_unidad " & _
             "WHERE     tvc_cencos = dvp_cencos " & _
             "AND       tvc_fecing = dvp_fecing " & _
             "AND       dvp_codmer = pro_codigo " & _
             "AND       pro_coduni = uni_codigo " & _
             "AND       tvc_estado = 'C' and tvc_cencos = '" & Trim(cForm.fpText1(0).text) & "' " & _
             "AND       (tvc_fecing >= " & sql1 & " AND tvc_fecing <= " & sql2 & ") " & _
             "AND       tvc_codbod = " & Val(fg_codigocbo(cForm.Combo1, 0, 10, "")) & " " & _
             "AND       (pro_codigo = '" & Trim(cForm.fpText1(1).text) & "' OR '" & Trim(cForm.fpText1(1).text) & "'='') " & _
             "AND       (pro_codtip = " & Val(cForm.fpText1(2).text) & " OR " & Val(cForm.fpText1(2).text) & "=0) " & _
             "GROUP BY  pro_codigo, pro_nombre, uni_nomcor " & _
             "ORDER BY  pro_codigo, pro_nombre", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
    
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 5
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Contrato"
    .TableCell(tcText, 1, 2) = Trim(cForm.fpText1(0).text) & " - " & Trim(cForm.fpayuda(0).Caption)
    .TableCell(tcText, 2, 1) = "Bodega"
    .TableCell(tcText, 2, 2) = Trim(Left(cForm.Combo1(0).text, 50))
    .TableCell(tcText, 3, 1) = "Periodo"
    .TableCell(tcText, 3, 2) = Trim(cForm.fpDateTime1(0).text) & " - " & Trim(cForm.fpDateTime1(1).text)
    .TableCell(tcText, 4, 1) = "Producto"
    .TableCell(tcText, 4, 2) = IIf(cForm.OptTipCli(1).Value = True, "Todos", Trim(cForm.fpText1(1).text) & " " & Trim(cForm.fpayuda(1).Caption))
    .TableCell(tcText, 5, 1) = "Familia"
    .TableCell(tcText, 5, 2) = IIf(cForm.OptTipCli(3).Value = True, "Todas", Trim(cForm.fpText1(2).text) & " " & Trim(cForm.fpayuda(2).Caption))
    
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    Print #1, .TableCell(tcText, 4, 1) & "|" & .TableCell(tcText, 4, 2)
    Print #1, .TableCell(tcText, 5, 1) & "|" & .TableCell(tcText, 5, 2)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 6: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 800: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1300: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1500: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 240
    .TableCell(tcText, 1, 1) = "Codigo"
    .TableCell(tcText, 1, 2) = "Producto"
    .TableCell(tcText, 1, 3) = "Unidad"
    .TableCell(tcText, 1, 4) = "Cantidad"
    .TableCell(tcText, 1, 5) = "Precio costo"
    .TableCell(tcText, 1, 6) = "Total"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5)
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 6: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 800: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1300: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1500: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
    i = 1: total = 0
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = Trim(RS1!pro_codigo)
            .TableCell(tcText, i, 2) = Trim(RS1!pro_nombre)
            .TableCell(tcText, i, 3) = Trim(RS1!uni_nomcor)
            .TableCell(tcText, i, 4) = Format(RS1!candig, fg_Pict(9, vg_DCa))
            .TableCell(tcText, i, 5) = Format(RS1!total / RS1!candig, fg_Pict(9, vg_DPr))
            .TableCell(tcText, i, 6) = Format(RS1!total, fg_Pict(9, vg_DPr))
            total = total + Format(RS1!total, fg_Pict(9, vg_DPr))
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5)
            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcFontBold, i, 5, i, 6) = True
    .TableCell(tcText, i, 5) = "Total"
    .TableCell(tcText, i, 6) = Format(total, fg_Pict(9, vg_DPr))
    Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5)
    .TableCell(tcRows) = i
    .PenColor = &H8000000C
    .TableBorder = tbAll
    .EndTable
    .EndDoc
    Close #1
End With
Preview.Show 1
fg_descarga
Exit Function
Error_Servicios:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function


Public Function I_ComensalesEstimados(cencos As String, codser As Long)
Dim i As Long, j As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long, coditem As Long
Dim vectot(7) As Double
On Local Error GoTo Error_ComensalesEstimados
fg_carga ""
MsgTitulo = "Informe de Comensales Estimados"
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    '------- Leer servicio
    RS1.Open RutinaLectura.Servicio(8, codser, ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Comensales Estimados Servicio : " & RS1!ser_nombre & " (" & RS1!ser_codigo & ")"
    RS1.Close: Set RS1 = Nothing
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 8000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcFontBold, 1, 1) = True: .TableCell(tcText, 1, 1) = "Contrato"
    .TableCell(tcFontBold, 2, 1) = True: .TableCell(tcText, 2, 1) = ""
    RS1.Open RutinaLectura.Cliente(1, cencos, ""), vg_db, adOpenStatic
    If Not RS1.EOF Then .TableCell(tcFontBold, 1, 2) = True: .TableCell(tcText, 1, 2) = ": " & Trim(RS1!cli_codigo) & " " & Trim(RS1!cli_nombre)
    RS1.Close: Set RS1 = Nothing
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.ServicioRaciones(2, codser), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
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
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Lunes"
    .TableCell(tcText, 1, 3) = "Martes"
    .TableCell(tcText, 1, 4) = "Miércoles"
    .TableCell(tcText, 1, 5) = "Jueves"
    .TableCell(tcText, 1, 6) = "Viernes"
    .TableCell(tcText, 1, 7) = "Sábado"
    .TableCell(tcText, 1, 8) = "Domingo"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 8: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1000: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    For i = 1 To 7
        vectot(i) = 0
    Next i
    i = 1: coditem = 0
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          If RS1!sra_coditem <> coditem Then
             If coditem > 0 Then i = i + 1
             If RS1!sra_coditem = 1 Then
                .TableCell(tcText, i, 1) = "Clientes"
             ElseIf RS1!sra_coditem = 2 Then
                .TableCell(tcText, i, 1) = "Comensales"
             ElseIf RS1!sra_coditem = 3 Then
                .TableCell(tcText, i, 1) = "Donación"
             End If
             coditem = RS1!sra_coditem
          End If
          .TableCell(tcText, i, (RS1!sra_serdia + 1)) = RS1!sra_raciones
          vectot(RS1!sra_serdia) = vectot(RS1!sra_serdia) + RS1!sra_raciones
          Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
          RS1.MoveNext
       Loop
       i = i + 2
       .TableCell(tcText, i, 1) = "Totales"
       For j = 1 To 7
          .TableCell(tcText, i, (j + 1)) = IIf(vectot(j) > 0, vectot(j), "")
       Next j
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Preview.Show 1
fg_descarga
Exit Function
Error_ComensalesEstimados:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Bodega()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Bodega
fg_carga ""
MsgTitulo = "Informe de Bodegas"
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
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.Bodega(3, 0, ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
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
Preview.Show 1
fg_descarga
Exit Function
Error_Bodega:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Provee()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Proveedor
fg_carga ""
MsgTitulo = "Informe de Proveedores"
Preview.Refresh
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
    .Preview = True
    .PreviewPage = 1
    .Orientation = orLandscape
    .MarginLeft = 500
    .StartDoc
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
    .TableCell(tcColWidth, , 1) = 13500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Maestro Proveedores"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.Proveedor(1, "", ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
    .StartTable
    .TableCell(tcCols) = 13: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1200: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1700: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1800: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1300: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 1300: .TableCell(tcAlign, , 5) = taLeftTop
    .TableCell(tcColWidth, , 6) = 900: .TableCell(tcAlign, , 6) = taLeftTop
    .TableCell(tcColWidth, , 7) = 900: .TableCell(tcAlign, , 7) = taLeftTop
    .TableCell(tcColWidth, , 8) = 900: .TableCell(tcAlign, , 8) = taLeftTop
    .TableCell(tcColWidth, , 9) = 1500: .TableCell(tcAlign, , 9) = taLeftTop
    .TableCell(tcColWidth, , 10) = 1200: .TableCell(tcAlign, , 10) = taLeftTop
    .TableCell(tcColWidth, , 11) = 1000: .TableCell(tcAlign, , 11) = taLeftTop
    .TableCell(tcColWidth, , 12) = 700: .TableCell(tcAlign, , 12) = taLeftTop
    .TableCell(tcColWidth, , 13) = 700: .TableCell(tcAlign, , 13) = taLeftTop
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
    .TableCell(tcText, 1, 12) = "Origen"
    .TableCell(tcText, 1, 13) = "Estado"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5) & "|" & _
              .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & .TableCell(tcText, 1, 9) & "|" & .TableCell(tcText, 1, 10) & "|" & .TableCell(tcText, 1, 11) & "|" & .TableCell(tcText, 1, 12) & "|" & "|" & .TableCell(tcText, 1, 13) & "|"
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 13: .TableCell(tcRows) = 20000
    .TableCell(tcColWidth, , 1) = 1200: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1700: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1800: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1300: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 1300: .TableCell(tcAlign, , 5) = taLeftTop
    .TableCell(tcColWidth, , 6) = 900: .TableCell(tcAlign, , 6) = taLeftTop
    .TableCell(tcColWidth, , 7) = 900: .TableCell(tcAlign, , 7) = taLeftTop
    .TableCell(tcColWidth, , 8) = 900: .TableCell(tcAlign, , 8) = taLeftTop
    .TableCell(tcColWidth, , 9) = 1500: .TableCell(tcAlign, , 9) = taLeftTop
    .TableCell(tcColWidth, , 10) = 1200: .TableCell(tcAlign, , 10) = taLeftTop
    .TableCell(tcColWidth, , 11) = 1000: .TableCell(tcAlign, , 11) = taLeftTop
    .TableCell(tcColWidth, , 12) = 700: .TableCell(tcAlign, , 12) = taLeftTop
    .TableCell(tcColWidth, , 13) = 700: .TableCell(tcAlign, , 13) = taCenterTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 2) = Trim(RS1!prv_nombre)
            .TableCell(tcText, i, 1) = Trim(fg_PintaRut(RS1!prv_codigo))
            .TableCell(tcText, i, 3) = IIf(IsNull(RS1!prv_direccion), "", Trim(RS1!prv_direccion))
            .TableCell(tcText, i, 4) = IIf(IsNull(RS1!prv_comuna), "", Trim(RS1!prv_comuna))
            .TableCell(tcText, i, 5) = IIf(IsNull(RS1!prv_ciudad), "", Trim(RS1!prv_ciudad))
            .TableCell(tcText, i, 6) = IIf(IsNull(RS1!prv_fono1), "", Trim(RS1!prv_fono1))
            .TableCell(tcText, i, 7) = IIf(IsNull(RS1!prv_fono2), "", Trim(RS1!prv_fono2))
            .TableCell(tcText, i, 8) = IIf(IsNull(RS1!prv_fax), "", Trim(RS1!prv_fax))
            .TableCell(tcText, i, 9) = IIf(IsNull(RS1!prv_percon), "", Trim(RS1!prv_percon))
            .TableCell(tcText, i, 10) = IIf(IsNull(RS1!prv_giro), "", Trim(RS1!prv_giro))
            .TableCell(tcText, i, 11) = IIf(IsNull(RS1!prv_emapro), "", Trim(RS1!prv_emapro))
            .TableCell(tcText, i, 12) = IIf(IsNull(RS1!prv_origen) Or Trim(RS1!prv_origen) = "0", "Local", "Centralizado")
            .TableCell(tcText, i, 13) = IIf(IsNull(RS1!prv_activo) Or Trim(RS1!prv_activo) = "1", "Inactivo", IIf(RS1!prv_activo = "2", "Elim.", "Activo"))
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & .TableCell(tcText, i, 5) & "|" & _
                      .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9) & "|" & .TableCell(tcText, i, 10) & "|" & .TableCell(tcText, i, 11) & "|" & .TableCell(tcText, i, 12) & "|" & "|" & .TableCell(tcText, i, 13) & "|"
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
Preview.Show 1
fg_descarga
Exit Function
Error_Proveedor:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_TarjetaRecetas(cuenta As Long, tiprec As Long)
Dim i As Long, X As Long, fil As Long
Dim CodRec As Integer, nomrec As String, NomFan As String, lc_codrec As Long, totLineas As Long
Dim canser As Double, cannet As Double, totcanser As Double, totcannet As Double, totcosto As Double, canpro As Double
Dim acorec As Variant
On Local Error GoTo Error_Tarjeta
MsgTitulo = "Informe de Recetas"
fg_carga ""
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
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 10: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Tarjeta Recetas " & IIf(tiprec = 0, "(Patrón)", IIf(tiprec = -1, "(Local)", "X Regimen " & tiprec & "  " & I_Receta.fpayuda(1).Caption & ""))
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .FontBold = False
    .FontSize = 8
    LogoEmp
    .text = Chr(13): .text = Chr(13)
    acorec = ""
    cuenta = 0
    For i = 1 To I_Receta.vaSpread1.MaxRows
        I_Receta.vaSpread1.Row = i: I_Receta.vaSpread1.Col = 1
        If I_Receta.vaSpread1.Value = "1" And I_Receta.vaSpread1.RowHidden = False Then
            I_Receta.vaSpread1.Col = 2: acorec = acorec & IIf(acorec <> "", ",", "") & I_Receta.vaSpread1.text
            cuenta = cuenta + 1
        End If
    Next i
    
    If I_Receta.Check1.Value = 1 Then
       RS1.Open "SELECT rec.rec_nombre, rec.rec_nomfan, ing.ing_nombre, d.cpi_precos, red_codigo, red.red_pctnut, red.red_pctapr, red.red_pctcoc, " & _
                "red.red_canpro, red.red_cospro, rec.rec_metpre FROM b_ingrediente ing ,b_recetadet red, b_receta rec, b_contlistpreing d " & _
                "WHERE rec.rec_codigo = red.red_codigo AND ing.ing_codigo = red.red_codpro AND ing.ing_codigo = d.cpi_coding AND rec.rec_codigo IN (" & acorec & ") " & _
                "AND   d.cpi_cencos = '" & MuestraCasino(1) & "' AND red.red_tiprec=" & tiprec & " AND ((red.red_tiprec<>0 AND red.red_cencos='" & MuestraCasino(1) & "') OR (red.red_tiprec=0 AND red.red_cencos='0')) ORDER BY rec.rec_nombre, red.red_codigo, red.red_nroite", vg_db, adOpenStatic
    Else
       RS1.Open "SELECT rec.rec_nombre, rec.rec_nomfan, ing.ing_nombre, c.cpi_precos, red_codigo, red.red_pctnut, red.red_pctapr, red.red_pctcoc, " & _
                "red.red_canpro, red.red_cospro FROM b_ingrediente ing ,b_recetadet red, b_contlistpreing c, b_receta rec " & _
                "WHERE rec.rec_codigo = red.red_codigo and ing.ing_codigo=red.red_codpro AND ing.ing_codigo=c.cpi_coding AND rec.rec_codigo IN (" & acorec & ") AND c.cpi_cencos='" & MuestraCasino(1) & "' AND red.red_tiprec=" & tiprec & " AND ((red.red_tiprec<>0 AND red.red_cencos='" & MuestraCasino(1) & "') OR (red.red_tiprec=0 AND red.red_cencos='0')) ORDER BY rec.rec_nombre, red.red_codigo, red.red_nroite", vg_db, adOpenStatic
    End If
    I_Receta.ProgressBar1.Scrolling = ccScrollingStandard
    I_Receta.ProgressBar1.max = cuenta
    I_Receta.ProgressBar1.Visible = True
    I_Receta.ProgressBar1.Value = 0
    totLineas = 0
    paso = 0
    If Not RS1.EOF Then
        Do While Not RS1.EOF
        CodRec = RS1!red_codigo
        If I_Receta.Check1.Value = 1 Then Preview.rtfPic.TextRTF = "": Preview.rtfPic.TextRTF = IIf(IsNull(RS1!rec_metpre), "", RS1!rec_metpre)
        lc_codrec = RS1!red_codigo
        .text = Chr(13): .text = Chr(13)
        nomrec = IIf(IsNull(RS1!rec_nombre), "", Trim(RS1!rec_nombre))
        NomFan = IIf(IsNull(RS1!rec_nomfan), "", Trim(RS1!rec_nomfan))
        'CATEGORIA DIETETICA,TIPO PLATO Y TIPO RACIONES
        .StartTable
        .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
        .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
        .TableCell(tcColWidth, , 2) = 7000: .TableCell(tcAlign, , 1) = taLeftTop
        .TableCell(tcText, 1, 1) = "Cat. Dietetica"
        .TableCell(tcText, 2, 1) = "Tipo Plato"
        .TableCell(tcText, 3, 1) = "Nro. Raciones"
        '***LOCALIZAR FILA
        I_Receta.vaSpread1.SetActiveCell 1, I_Receta.vaSpread1.SearchCol(2, 0, I_Receta.vaSpread1.MaxRows, Trim(Str(CodRec)), SearchFlagsNone) ', SearchFlagsCaseSensitive)
        I_Receta.vaSpread1.Row = I_Receta.vaSpread1.ActiveRow
        I_Receta.vaSpread1.Col = 5: .TableCell(tcText, 1, 2) = I_Receta.vaSpread1.text
        I_Receta.vaSpread1.Col = 6: .TableCell(tcText, 2, 2) = I_Receta.vaSpread1.text
        I_Receta.vaSpread1.Col = 7: .TableCell(tcText, 3, 2) = I_Receta.vaSpread1.text
        Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
        Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
        Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
        .TableBorder = tbNone
        totLineas = totLineas + .TableCell(tcRows)
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
        totLineas = totLineas + .TableCell(tcRows)
        .EndTable
        .text = Chr(13)
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
        .TableCell(tcText, 1, 1) = "Nombre Ingrediente"
        .TableCell(tcText, 1, 2) = "C.Bruta"
        .TableCell(tcText, 1, 3) = "%Aprov."
        .TableCell(tcText, 1, 4) = "%A.Coc."
        .TableCell(tcText, 1, 5) = "%P.Nut."
        .TableCell(tcText, 1, 6) = "C.Servir"
        .TableCell(tcText, 1, 7) = "C.Neta"
        .TableCell(tcText, 1, 8) = "Costo"
        Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
                   .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8)
        .TableBorder = tbAll
        totLineas = totLineas + .TableCell(tcRows)
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
        fil = 0
        totcanser = 0
        totcannet = 0
        totcosto = 0
        Do While Not RS1.EOF And CodRec = lc_codrec
            If Not RS1.EOF Then lc_codrec = RS1!red_codigo Else lc_codrec = 0
            If CodRec = lc_codrec Then
                fil = fil + 1
                canpro = Format(RS1!red_canpro, fg_Pict(6, vg_RDCa))
                canser = Format((RS1!red_pctapr / 100) * canpro * (RS1!red_pctcoc / 100), fg_Pict(6, vg_RDCa))
                cannet = Format((RS1!red_pctnut / 100) * canpro, fg_Pict(6, vg_RDCa))
                .TableCell(tcText, fil, 1) = RS1!ing_nombre
                .TableCell(tcText, fil, 2) = Format(canpro, fg_Pict(6, vg_RDCa))
                .TableCell(tcText, fil, 3) = RS1!red_pctapr
                .TableCell(tcText, fil, 4) = RS1!red_pctcoc
                .TableCell(tcText, fil, 5) = RS1!red_pctnut
                .TableCell(tcText, fil, 6) = Format(canser, fg_Pict(6, vg_RDCa))
                .TableCell(tcText, fil, 7) = Format(cannet, fg_Pict(6, vg_RDCa))
                .TableCell(tcText, fil, 8) = Format(canpro * RS1!cpi_precos, fg_Pict(6, vg_DCa))
                 Print #1, .TableCell(tcText, fil, 1) & "|" & .TableCell(tcText, fil, 2) & "|" & .TableCell(tcText, fil, 3) & "|" & .TableCell(tcText, fil, 4) & "|" & _
                           .TableCell(tcText, fil, 5) & "|" & .TableCell(tcText, fil, 6) & "|" & .TableCell(tcText, fil, 7) & "|" & .TableCell(tcText, fil, 8)
                totcanser = totcanser + canser
                totcannet = totcannet + cannet
                totcosto = Format(totcosto + canpro * RS1!cpi_precos, fg_Pict(6, vg_DCa))
            End If
            RS1.MoveNext
        Loop
        .TableCell(tcRows) = fil
        .PenColor = &HC0C0C0
        .TableBorder = tbAll
        totLineas = totLineas + .TableCell(tcRows)
        .EndTable
        .StartTable
        .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
        .TableCell(tcColWidth, , 1) = 7500
        .TableCell(tcColWidth, , 2) = 1000
        .TableCell(tcColWidth, , 3) = 1000
        .TableCell(tcColWidth, , 4) = 1000
        .TableCell(tcFontBold, 1) = True
        .TableCell(tcText, 1, 1) = "": .TableCell(tcAlign, , 1) = taRightTop
        .TableCell(tcText, 1, 2) = totcanser: .TableCell(tcAlign, , 2) = taRightTop
        .TableCell(tcText, 1, 3) = totcannet: .TableCell(tcAlign, , 3) = taRightTop
        .TableCell(tcText, 1, 4) = Format(totcosto, fg_Pict(6, vg_DCa)): .TableCell(tcAlign, , 4) = taRightTop
        Print #1, "|" & "|" & "|"; .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "||" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
        .TableBorder = tbNone
        totLineas = totLineas + .TableCell(tcRows)
        .EndTable
        If I_Receta.Check1.Value = 1 And Trim(Preview.rtfPic.text) <> "" Then
            .text = Chr(13): .text = Chr(13)
            .StartTable
            .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftTop 'taCenterMiddle
            .TableCell(tcFontSize, 1) = 9: .TableCell(tcFontBold, 1) = True
            .TableCell(tcText, 1, 1) = Chr(13) + Preview.rtfPic.text 'TextRTF
            Print #1, .TableCell(tcText, 1, 1)
            .TableBorder = tbBox
            totLineas = totLineas + .TableCell(tcRows)
            .EndTable
            For X = 1 To Len(Preview.rtfPic.text)
                If Asc(Mid(Preview.rtfPic.text, X, 1)) = 13 Then totLineas = totLineas + 1
            Next X
        End If
        If totLineas > 35 And .CurrentY < 14500 Then
            totLineas = 0: .NewPage
        End If
'            RS1.MoveNext
            I_Receta.ProgressBar1.Value = I_Receta.ProgressBar1.Value + 1
        Loop
    
    End If
'            I_Receta.ProgressBar1.Value = I_Receta.ProgressBar1.Value + 1
'        End If
'    Next i
    RS1.Close: Set RS1 = Nothing
    .EndDoc
    I_Receta.ProgressBar1.Value = I_Receta.ProgressBar1.max
    I_Receta.ProgressBar1.Visible = False
    Close #1
End With

Preview.Show 1
fg_descarga
Exit Function
Error_Tarjeta:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_AporteRecetas(cuenta As Long, tiprec As Long)
Dim i As Long, X As Long, fil As Long, Col As Long, VecCol(100) As Long
Dim CodRec As Integer, codrec2 As Integer, nomrec As String, NomFan As String, vector(100) As Double
Dim canser As Double, cannet As Double, totcanser As Double, totcannet As Double, totcosto As Double, canpro As Double
Dim alto As Long, aAp As String, Cantidad As Double, lc_codpro As String, horini As Date, horter As Date
Dim cLin As String, j As Long, icol As Long, totLineas As Long, acorec As String, avance As Long
Dim arr
On Local Error GoTo Error_ApoRecet
MsgTitulo = "Informe de Recetas " & IIf(tiprec = 0, "(Patrón)", IIf(tiprec = -1, "(Local)", "X Regimen " & tiprec & "  " & I_Receta.fpayuda(1).Caption & ""))
fg_carga ""
aAp = Trim(vg_NUsr) & "_tmp_imprec"
alto = 0
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
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
    .Header = "" & fg_poneencpagina & "||"
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = IIf(I_Receta.List1.SelCount > 7, 13500, 10500): .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 10: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Informe Recetas Aporte Nutricional " & IIf(tiprec = 0, "(Patrón)", IIf(tiprec = -1, "(Local)", "X Regimen " & tiprec & "  " & I_Receta.fpayuda(1).Caption & ""))
    .TableBorder = tbNone
    .EndTable
    .FontSize = 7
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    'LogoEmp
    .text = Chr(13): .text = Chr(13)
    acorec = ""
    For i = 1 To I_Receta.vaSpread1.MaxRows
        I_Receta.vaSpread1.Row = i: I_Receta.vaSpread1.Col = 1
        If I_Receta.vaSpread1.Value = "1" And I_Receta.vaSpread1.RowHidden = False Then
            I_Receta.vaSpread1.Col = 2: acorec = acorec & IIf(acorec <> "", ",", "") & I_Receta.vaSpread1.text
        End If
    Next i
    RS1.Open "SELECT count(red.red_codigo) nreg " & _
             "FROM b_ingrediente ing, b_recetadet red, b_receta rec " & _
             "WHERE ing.ing_codigo=red.red_codpro " & _
             "AND rec.rec_codigo=red.red_codigo " & _
             "AND rec.rec_codigo IN (" & acorec & ") " & _
             "AND red.red_tiprec=" & tiprec & " AND ((red.red_tiprec<>0 AND red.red_cencos='" & MuestraCasino(1) & "') OR (red.red_tiprec=0 AND red.red_cencos='0')) " & _
             "", vg_db, adOpenStatic
    If Not RS1.EOF Then avance = RS1!nreg
    RS1.Close: Set RS1 = Nothing
    RS1.Open "SELECT red.red_codigo, ing.ing_nombre, red.red_pctnut, red.red_pctapr, " & _
             "red.red_pctcoc, red.red_canpro, red.red_codpro, rec.rec_basrac, ing.ing_facnut " & _
             "FROM b_ingrediente ing, b_recetadet red, b_receta rec " & _
             "WHERE ing.ing_codigo=red.red_codpro " & _
             "AND rec.rec_codigo=red.red_codigo " & _
             "AND rec.rec_codigo IN (" & acorec & ") " & _
             "AND red.red_tiprec=" & tiprec & " AND ((red.red_tiprec<>0 AND red.red_cencos='" & MuestraCasino(1) & "') OR (red.red_tiprec=0 AND red.red_cencos='0')) " & _
             "ORDER BY rec.rec_nombre, red.red_nroite", vg_db, adOpenStatic
    If Not RS1.EOF Then
    RS2.Open "SELECT distinct pnu_codpro, pnu_codapo, pnu_canapo FROM b_productonut", vg_db, adOpenStatic
    If Not RS2.EOF Then
       arr = RS2.GetRows
       RS2.Close: Set RS2 = Nothing
       Preview.vaSpread1.MaxRows = 0
       Preview.vaSpread1.MaxCols = 3
       For i = 0 To UBound(arr, 2)
            Preview.vaSpread1.MaxRows = Preview.vaSpread1.MaxRows + 1
            Preview.vaSpread1.Row = Preview.vaSpread1.MaxRows
        
            Preview.vaSpread1.Col = 1
            Preview.vaSpread1.text = arr(0, i)
        
            Preview.vaSpread1.Col = 2
            Preview.vaSpread1.text = Trim(arr(1, i))
        
            Preview.vaSpread1.Col = 3
            Preview.vaSpread1.text = arr(2, i)
       Next i
    End If
    If RS2.State = 1 Then RS2.Close: Set RS2 = Nothing
    
    I_Receta.ProgressBar1.Scrolling = ccScrollingSmooth
    I_Receta.ProgressBar1.max = avance
    I_Receta.ProgressBar1.Visible = True
    I_Receta.ProgressBar1.Value = 0
    totLineas = 0
    CodRec = 0
    horini = Format(Now, "hh:mm:ss")
    Do While Not RS1.EOF
        If CodRec <> RS1!red_codigo Then
            '***LOCALIZAR FILA
            CodRec = RS1!red_codigo
            I_Receta.vaSpread1.SetActiveCell 1, I_Receta.vaSpread1.SearchCol(2, 0, I_Receta.vaSpread1.MaxRows, Trim(Str(CodRec)), SearchFlagsNone) ', SearchFlagsCaseSensitive)
            I_Receta.vaSpread1.Row = I_Receta.vaSpread1.ActiveRow
            For X = 1 To 100
                vector(X) = 0
            Next X
            .text = Chr(13): .text = Chr(13)
            I_Receta.vaSpread1.Col = 3: nomrec = I_Receta.vaSpread1.text
            I_Receta.vaSpread1.Col = 4: NomFan = I_Receta.vaSpread1.text
            'CATEGORIA DIETETICA,TIPO PLATO Y TIPO RACIONES
            .StartTable
            .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
            .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
            .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 1) = taLeftTop
            .TableCell(tcText, 1, 1) = "Cat. Dietetica"
            .TableCell(tcText, 2, 1) = "Tipo Plato"
            .TableCell(tcText, 3, 1) = "Nro. Raciones"
            I_Receta.vaSpread1.Col = 5: .TableCell(tcText, 1, 2) = I_Receta.vaSpread1.text
            I_Receta.vaSpread1.Col = 6: .TableCell(tcText, 2, 2) = I_Receta.vaSpread1.text
            I_Receta.vaSpread1.Col = 7: .TableCell(tcText, 3, 2) = I_Receta.vaSpread1.text
            .TableBorder = tbNone
            totLineas = totLineas + .TableCell(tcRows)
            .EndTable
            'NOMBRE RECETA
            .StartTable
            .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = IIf(I_Receta.List1.SelCount > 7, 13500, 10500): .TableCell(tcAlign, , 1) = taCenterMiddle
            .TableCell(tcFontSize, 1) = 9: .TableCell(tcFontBold, 1) = True
            .TableCell(tcText, 1, 1) = IIf(I_Receta.Option1(0).Value = True, "* " & nomrec & " *", "* " & NomFan & " *")
            .TableBorder = tbNone
            totLineas = totLineas + .TableCell(tcRows)
            .EndTable
            .text = Chr(13)
            .StartTable
            .TableCell(tcCols) = 5 + I_Receta.List1.SelCount: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = 1300: .TableCell(tcAlign, , 1) = taLeftTop
            .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
            .TableCell(tcColWidth, , 3) = 800: .TableCell(tcAlign, , 3) = taRightTop
            .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
            .TableCell(tcColWidth, , 5) = 800: .TableCell(tcAlign, , 5) = taRightTop
            .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True:  .TableCell(tcRowHeight, 1) = 230
            .TableCell(tcText, 1, 1) = "Código"
            .TableCell(tcText, 1, 2) = "Nombre Ingrediente"
            .TableCell(tcText, 1, 3) = "C.Bruta"
            .TableCell(tcText, 1, 4) = "C.Servir"
            .TableCell(tcText, 1, 5) = "C.Neta"
            Col = 6
            For X = 0 To I_Receta.List1.listcount - 1
                If I_Receta.List1.Selected(X) = True Then
                    .TableCell(tcColWidth, , Col) = 750: .TableCell(tcAlign, , Col) = taRightTop
                    .TableCell(tcText, 1, Col) = I_Receta.List1.List(X)
                    VecCol(I_Receta.List1.ItemData(X)) = Col
                    Col = Col + 1
                End If
            Next X
            icol = Col - 1
            .TableBorder = tbAll
            totLineas = totLineas + .TableCell(tcRows)
            .EndTable
        
            CodRec = RS1!red_codigo
                        
            .StartTable
            .TableCell(tcCols) = 5 + I_Receta.List1.SelCount: .TableCell(tcRows) = 150
            .TableCell(tcColWidth, , 1) = 1300: .TableCell(tcAlign, , 1) = taLeftTop
            .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
            .TableCell(tcColWidth, , 3) = 800: .TableCell(tcAlign, , 3) = taRightTop
            .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
            .TableCell(tcColWidth, , 5) = 800: .TableCell(tcAlign, , 5) = taRightTop
            .TableCell(tcText, 1, 1, 1, 5 + I_Receta.List1.SelCount) = ""
            totcanser = 0
            totcannet = 0
            totcanpro = 0
            fil = 0
            
        End If
        
        I_Receta.ProgressBar1.Value = I_Receta.ProgressBar1.Value + 1
        fil = fil + 1
        .TableCell(tcText, fil, 6, fil, I_Receta.List1.SelCount + 6) = Format(0, fg_Pict(9, vg_DCa))
        canpro = Format(RS1!red_canpro, fg_Pict(6, vg_RDCa))
        canser = Format((RS1!red_pctapr / 100) * canpro * (RS1!red_pctcoc / 100), fg_Pict(6, vg_RDCa))
        cannet = Format((RS1!red_pctnut / 100) * canpro, fg_Pict(6, vg_RDCa))
        .TableCell(tcText, fil, 1) = RS1!red_codpro
        .TableCell(tcText, fil, 2) = RS1!ing_nombre
        .TableCell(tcText, fil, 3) = Format(canpro, fg_Pict(6, vg_RDCa))
        .TableCell(tcText, fil, 4) = Format(canser, fg_Pict(6, vg_RDCa))
        .TableCell(tcText, fil, 5) = Format(cannet, fg_Pict(6, vg_RDCa))
        totcanpro = totcanpro + canpro
        totcanser = totcanser + canser
        totcannet = totcannet + cannet
        Col = 6
'        If RS2.RecordCount > 0 Then RS2.MoveFirst
'        RS2.Find "pnu_codpro='" & Trim(RS1!red_codpro) & "'", , adSearchForward
'        If Not RS2.EOF Then
'            lc_codpro = RS2!pnu_codpro
'            Do While Not RS2.EOF And lc_codpro = Trim(RS1!red_codpro)
'                If Not RS2.EOF Then lc_codpro = RS2!pnu_codpro Else lc_codpro = 0
'                If lc_codpro = Trim(RS1!red_codpro) Then
'                    cantidad = (((RS1!red_pctnut / 100) * (RS2!pnu_canapo * (RS1!red_canpro / RS1!rec_basrac))) / RS1!ing_facnut)
'                    .TableCell(tcText, fil, VecCol(Val(RS2!pnu_codapo))) = Format(cantidad, fg_Pict(6, vg_DCa))
'                    vector(Val(RS2!pnu_codapo)) = vector(Val(RS2!pnu_codapo)) + Format(cantidad, fg_Pict(6, vg_DCa))
'                    Col = Col + 1
'                End If
'                RS2.MoveNext
'            Loop
'        Else
'            vector(Col) = vector(Col) + Format(0, fg_Pict(6, vg_DCa))
'            Col = Col + 1
'        End If
        Trim (CStr(RS1!red_codpro))
        ind_ini = Preview.vaSpread1.SearchCol(1, -1, Preview.vaSpread1.MaxRows, Trim(CStr(RS1!red_codpro)), SearchFlagsEqual)
        codpro = ""
        For ind_par = ind_ini To Preview.vaSpread1.MaxRows
            Preview.vaSpread1.Row = ind_par
            Preview.vaSpread1.Col = 1
            If Preview.vaSpread1.text <> Trim(RS1!red_codpro) Then Exit For
            Preview.vaSpread1.Col = 2
            codapo = Preview.vaSpread1.text
            Preview.vaSpread1.Col = 3
            canapo = Preview.vaSpread1.text
            Cantidad = (((RS1!red_pctnut / 100) * (canapo * (RS1!red_canpro / RS1!rec_basrac))) / RS1!ing_facnut)
            .TableCell(tcText, fil, VecCol(Val(codapo))) = Format(Cantidad, fg_Pict(6, vg_DCa))
            vector(Val(codapo)) = vector(Val(codapo)) + Format(Cantidad, fg_Pict(6, vg_DCa))
            Col = Col + 1
'            For j = 1 To inut
'                If vecdie(j) = codapo Then
'                   .TableCell(tcAlign, X, Y + j) = taRightTop
'                   .TableCell(tcText, X, Y + j) = Format(((((RS1!red_pctnut / 100) * (canapo * (RS1!canpro))) / RS1!ing_facnut)), fg_Pict(6, 2))
'                   vecrec(j) = CCur(vecrec(j) + ((((RS1!red_pctnut / 100) * (canapo * (RS1!canpro))) / RS1!ing_facnut)))
'                   vecdia(j) = CCur(vecdia(j) + ((((RS1!red_pctnut / 100) * (canapo * (RS1!canpro))) / RS1!ing_facnut)))
'                   Exit For
'                End If
'            Next j
        Next ind_par
        
        RS1.MoveNext
        If Not RS1.EOF Then codrec2 = RS1!red_codigo Else codrec2 = 0
        If CodRec <> codrec2 Or RS1.EOF Then
            CodRec = 0
            .TableCell(tcColWidth, 1, 6, fil, 5 + I_Receta.List1.SelCount) = 750: .TableCell(tcAlign, 1, 6, fil, 5 + I_Receta.List1.SelCount) = taRightTop
            .TableCell(tcRows) = fil
            .PenColor = &HC0C0C0
            .TableBorder = tbAll
            totLineas = totLineas + .TableCell(tcRows)
            .EndTable
            .StartTable
            .TableCell(tcCols) = 4 + I_Receta.List1.SelCount: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = 3300
            .TableCell(tcColWidth, , 2) = 800
            .TableCell(tcColWidth, , 3) = 800
            .TableCell(tcColWidth, , 4) = 800
            .TableCell(tcFontBold, 1) = True
            .TableCell(tcText, 1, 1) = "Totales" & Space(36): .TableCell(tcAlign, , 1) = taRightTop
            .TableCell(tcText, 1, 2) = Format(totcanpro, fg_Pict(6, vg_RDCa)): .TableCell(tcAlign, , 2) = taRightTop
            .TableCell(tcText, 1, 3) = Format(totcanser, fg_Pict(6, vg_RDCa)): .TableCell(tcAlign, , 3) = taRightTop
            .TableCell(tcText, 1, 4) = Format(totcannet, fg_Pict(6, vg_RDCa)): .TableCell(tcAlign, , 4) = taRightTop
            
            Col = 5
            For X = 0 To I_Receta.List1.listcount - 1
                If I_Receta.List1.Selected(X) = True Then
                    .TableCell(tcColWidth, , Col) = 750: .TableCell(tcAlign, , Col) = taRightTop
                    .TableCell(tcText, 1, Col) = Format(vector(I_Receta.List1.ItemData(X)), fg_Pict(6, 2))
                    Col = Col + 1
                End If
            Next X
            .TableBorder = tbNone
            totLineas = totLineas + .TableCell(tcRows)
            .EndTable
            Dim varx As Integer
            varx = IIf(I_Receta.List1.SelCount > 7, 25, 40)
            If totLineas > varx And .CurrentY < IIf(I_Receta.List1.SelCount > 7, 11000, 14500) Then
                totLineas = 0: .NewPage
            End If
        End If
    Loop
'    RS2.Close: Set RS2 = Nothing
    End If
    RS1.Close: Set RS1 = Nothing
    horter = Format(Now, "hh:mm:ss")
    .EndDoc
    I_Receta.ProgressBar1.Visible = False
End With
'MsgBox horini & " - " & horter
'MsgBox CDate(horter - horini)
Preview.Show 1
fg_descarga
Exit Function
Error_ApoRecet:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Resume Next
    Exit Function
End Function

Public Function I_NombreRecetas(cuenta As Long, tiprec As Long)
Dim i As Long, X As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim CodRec As Long, nomrec As String, NomFan As String, catdie As String, tippla As String
Dim acorec  As Variant
On Local Error GoTo Error_NomRec
MsgTitulo = "Informe de Recetas"
fg_carga ""
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
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Nombre Recetas " & IIf(I_Receta.Option2(0).Value = True, "(Patrón)", IIf(I_Receta.Option2(1).Value = True, "(Local)", "(x Regimen) " & tiprec & "  " & I_Receta.fpayuda(1).Caption & ""))
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13)
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
    .text = Chr(13)
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
    I_Receta.ProgressBar1.max = cuenta
    I_Receta.ProgressBar1.Visible = True
    I_Receta.ProgressBar1.Value = 0
    acorec = ""
    For i = 1 To I_Receta.vaSpread1.MaxRows
        I_Receta.vaSpread1.Row = i: I_Receta.vaSpread1.Col = 1
        If I_Receta.vaSpread1.Value = "1" And I_Receta.vaSpread1.RowHidden = False Then
            I_Receta.vaSpread1.Col = 2: acorec = acorec & IIf(acorec <> "", ",", "") & I_Receta.vaSpread1.text
            cuenta = cuenta + 1
        End If
    Next i
    
    RS1.Open "SELECT DISTINCT a.rec_codigo, a.rec_nombre, a.rec_nomfan FROM b_receta a, b_recetadet b " & _
             "WHERE a.rec_codigo = b.red_codigo " & _
             "AND   a.rec_codigo In (" & acorec & ") " & _
             "AND   b.red_tiprec = " & tiprec & " " & _
             "AND ((b.red_tiprec<>0 AND b.red_cencos = '" & MuestraCasino(1) & "') OR (b.red_tiprec = 0 AND b.red_cencos = '0')) ORDER BY a.rec_nombre", vg_db, adOpenStatic
    X = 1
    Do While Not RS1.EOF
       '***LOCALIZAR FILA
       CodRec = RS1!rec_codigo
       I_Receta.vaSpread1.SetActiveCell 1, I_Receta.vaSpread1.SearchCol(2, 0, I_Receta.vaSpread1.MaxRows, Trim(Str(CodRec)), SearchFlagsNone) ', SearchFlagsCaseSensitive)
       I_Receta.vaSpread1.Row = I_Receta.vaSpread1.ActiveRow
       I_Receta.vaSpread1.Col = 6: tippla = I_Receta.vaSpread1.text
       .TableCell(tcText, X, 1) = RS1!rec_codigo
       .TableCell(tcText, X, 2) = IIf(I_Receta.Option1(0).Value = True, IIf(IsNull(RS1!rec_nombre), "", Trim(RS1!rec_nombre)), IIf(IsNull(RS1!rec_nomfan), "", Trim(RS1!rec_nomfan)))
       .TableCell(tcText, X, 3) = tippla
       Print #1, .TableCell(tcText, X, 1) & "|" & .TableCell(tcText, X, 2) & "|" & .TableCell(tcText, X, 3)
       I_Receta.ProgressBar1.Value = I_Receta.ProgressBar1.Value + 1
'       For i = 1 To 1000
'       Next i
       RS1.MoveNext: X = X + 1
    Loop
    RS1.Close: Set RS1 = Nothing
    .PenColor = &HC0C0C0
    .TableBorder = tbAll
    .TableCell(tcRows) = X
    .EndTable
    .EndDoc
    Close #1
    I_Receta.ProgressBar1.Value = I_Receta.ProgressBar1.max
    I_Receta.ProgressBar1.Visible = False
End With
Preview.Show 1
fg_descarga
Exit Function
Error_NomRec:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Productos()
Dim i As Long, X As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long, codigo As String
On Local Error GoTo Error_Productos
fg_carga ""
MsgTitulo = "Informe de Productos"
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
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
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
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
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 11: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 2600: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 650: .TableCell(tcAlign, , 5) = taCenterTop
    .TableCell(tcColWidth, , 6) = 650: .TableCell(tcAlign, , 6) = taCenterTop
    .TableCell(tcColWidth, , 7) = 900: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 900: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1050: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcColWidth, , 10) = 900: .TableCell(tcAlign, , 10) = taRightTop
    .TableCell(tcColWidth, , 11) = 750: .TableCell(tcAlign, , 11) = taCenterTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = IIf(I_Produc.Option1(0).Value = True, "Nombre", "Nombre Fantasía")
    .TableCell(tcText, 1, 3) = "Disp. Cont."
    .TableCell(tcText, 1, 4) = "Familia"
    .TableCell(tcText, 1, 5) = "Uni.Env"
    .TableCell(tcText, 1, 6) = "Uni.Emb"
    .TableCell(tcText, 1, 7) = "Cant.xUni."
    .TableCell(tcText, 1, 8) = "Ult.Precio"
    .TableCell(tcText, 1, 9) = "Fec.Ult.Comp."
    .TableCell(tcText, 1, 10) = "P.M.P."
    .TableCell(tcText, 1, 11) = "Ctr.Stock"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & .TableCell(tcText, 1, 9) & "|" & .TableCell(tcText, 1, 10) & "|" & .TableCell(tcText, 1, 11)
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 11: .TableCell(tcRows) = 60000
    .TableCell(tcColWidth, , 1) = 900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 2600: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 650: .TableCell(tcAlign, , 5) = taCenterTop
    .TableCell(tcColWidth, , 6) = 650: .TableCell(tcAlign, , 6) = taCenterTop
    .TableCell(tcColWidth, , 7) = 900: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 900: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1050: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcColWidth, , 10) = 900: .TableCell(tcAlign, , 10) = taRightTop
    .TableCell(tcColWidth, , 11) = 750: .TableCell(tcAlign, , 11) = taCenterTop
    X = 1
    If vg_tipbase = "1" Then
       Dim aAp As String
       '-------> Insert tabla productospmpdia
       aAp = Trim(vg_NUsr) & "_tmp_ProductoPMPInf"
       fg_CheckTmp aAp
       vg_db.Execute "SELECT TOP 1 ppd_cencos, ppd_codpro, 0 AS ppd_propon, 0 AS ppd_upreco, null AS ppd_fecuco, Max(ppd_fecdia) AS ppd_fecdia " & _
                     "INTO " & aAp & " " & _
                     "FROM b_productospmpdia " & _
                     "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                     "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(Date, "yyyymmdd") & " " & _
                     "AND   ppd_propon > 0 " & _
                     "GROUP BY ppd_cencos, ppd_codpro ORDER BY Max(ppd_fecdia) DESC"
       vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
       vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon, " & aAp & ".ppd_upreco=b_productospmpdia.ppd_upreco, " & aAp & ".ppd_fecuco=b_productospmpdia.ppd_fecuco"
       vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_upreco, ppd_fecuco, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_upreco, ppd_fecuco, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
    End If
    For i = 1 To I_Produc.vaSpread1.MaxRows
        I_Produc.vaSpread1.Row = i: I_Produc.vaSpread1.Col = 1
        If I_Produc.vaSpread1.text = "1" Then
            I_Produc.vaSpread1.Col = 2: codigo = I_Produc.vaSpread1.text
            If vg_tipbase = "1" Then
                RS1.Open "SELECT DISTINCT tip.tip_codigo,tip.tip_nombre, uni.uni_nomcor, emb.emb_nomcor, pro.*, (SELECT DISTINCT ppd_upreco FROM " & aAp & " WHERE pro.pro_codigo = ppd_codpro AND ppd_cencos = '" & MuestraCasino(1) & "') AS ppd_upreco, (SELECT DISTINCT ppd_fecuco FROM " & aAp & " WHERE pro.pro_codigo = ppd_codpro AND ppd_cencos = '" & MuestraCasino(1) & "') AS ppd_fecuco, (SELECT DISTINCT ppd_propon FROM " & aAp & " WHERE pro.pro_codigo = ppd_codpro AND ppd_cencos = '" & MuestraCasino(1) & "') AS ppd_propon, IIF(pro.pro_maepro = 0,'Ambos', b.tis_nombre) AS tis_nombre " & _
                         "FROM a_embalaje emb, a_unidad uni, a_tipopro tip, b_productos pro, a_tiposervicio b, b_clientes c " & _
                         "WHERE (b.tis_codigo = c.cli_codtis OR pro.pro_maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = pro.pro_maepro OR pro.pro_maepro < 1) AND tip.tip_codigo = pro.pro_codtip " & _
                         "AND   uni.uni_codigo = pro.pro_coduni " & _
                         "AND   emb.emb_codigo = pro.pro_codemb " & _
                         "AND   pro.pro_codigo = '" & Trim(codigo) & "' ORDER BY pro.pro_codigo", vg_db, adOpenStatic
            Else
               RS1.Open "SELECT DISTINCT tip.tip_codigo,tip.tip_nombre, uni.uni_nomcor, emb.emb_nomcor, pro.*, (SELECT TOP 1 ppd_upreco FROM b_productospmpdia e WHERE pro.pro_codigo = ppd_codpro AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia <= " & Format(Date, "yyyymmdd") & " ORDER BY ppd_fecdia DESC) AS ppd_upreco, (SELECT TOP 1 ppd_fecuco FROM b_productospmpdia WHERE pro.pro_codigo = ppd_codpro AND ppd_cencos = '" & MuestraCasino(1) & "'AND ppd_fecdia <= " & Format(Date, "yyyymmdd") & " ORDER BY ppd_fecdia DESC) AS ppd_fecuco,  (SELECT TOP 1 ppd_propon FROM b_productospmpdia WHERE pro.pro_codigo = ppd_codpro AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia <= " & Format(Date, "yyyymmdd") & " ORDER BY ppd_fecdia DESC) AS ppd_propon, CASE WHEN pro.pro_maepro=0 THEN 'Ambos' ELSE b.tis_nombre END AS tis_nombre " & _
                        "FROM a_embalaje emb, a_unidad uni, a_tipopro tip, b_productos pro, a_tiposervicio b, b_clientes c " & _
                        "WHERE (b.tis_codigo = c.cli_codtis OR pro.pro_maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = pro.pro_maepro OR pro.pro_maepro < 1) AND tip.tip_codigo = pro.pro_codtip " & _
                        "AND   uni.uni_codigo = pro.pro_coduni " & _
                        "AND   emb.emb_codigo = pro.pro_codemb " & _
                        "AND   pro.pro_codigo = '" & Trim(codigo) & "' ORDER BY pro.pro_codigo", vg_db, adOpenStatic
            End If
            If Not RS1.EOF Then
                Do While Not RS1.EOF
                    .TableCell(tcText, X, 1) = RS1!pro_codigo
                    .TableCell(tcText, X, 2) = Trim(RS1!pro_nombre)
                    .TableCell(tcText, X, 3) = IIf(RS1!pro_maepro = 0, "Ambos", Trim(RS1!tis_nombre))
                    I_Produc.vaSpread1.Col = 4
                    .TableCell(tcText, X, 4) = I_Produc.vaSpread1.Value
                    .TableCell(tcText, X, 5) = Trim(RS1!uni_nomcor)
                    .TableCell(tcText, X, 6) = Trim(RS1!emb_nomcor)
                    .TableCell(tcText, X, 7) = Trim(RS1!pro_uniemb)
                    .TableCell(tcText, X, 8) = IIf(IsNull(RS1!ppd_upreco), 0, Format(RS1!ppd_upreco, fg_Pict(9, vg_DPr)))
                    .TableCell(tcText, X, 9) = IIf(IsNull(RS1!ppd_fecuco), "", RS1!ppd_fecuco)
                    .TableCell(tcText, X, 10) = IIf(IsNull(RS1!ppd_propon), 0, Format(RS1!ppd_propon, fg_Pict(9, vg_DPr)))
                    .TableCell(tcText, X, 11) = IIf(IsNull(RS1!pro_ctrsto) Or RS1!pro_ctrsto = 0, "", "X")
                    Print #1, .TableCell(tcText, X, 1) & "|" & .TableCell(tcText, X, 2) & "|" & .TableCell(tcText, X, 3) & "|" & .TableCell(tcText, X, 4) & "|" & _
                              .TableCell(tcText, X, 5) & "|" & .TableCell(tcText, X, 6) & "|" & .TableCell(tcText, X, 7) & "|" & "'" & .TableCell(tcText, X, 8) & "|" & .TableCell(tcText, X, 9) & "|" & .TableCell(tcText, X, 10) & "|" & .TableCell(tcText, X, 11)
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
'-------> Borrar tablas temporales
If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
Close #1
fg_descarga
Exit Function
Error_Productos:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Ingrediente()
Dim i As Long, X As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long, codigo As String
On Local Error GoTo Error_Productos
fg_carga ""
MsgTitulo = "Informe de Ingredientes"
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
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
    .Header = "" & fg_poneencpagina & "||"
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
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
    .text = Chr(13): .text = Chr(13)
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
    .TableCell(tcCols) = 11: .TableCell(tcRows) = 60000
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
        If I_Produc.vaSpread1.text = "1" Then
            I_Produc.vaSpread1.Col = 2: codigo = I_Produc.vaSpread1.text
            RS1.Open RutinaLectura.Ingrediente(3, Trim(codigo), ""), vg_db, adOpenStatic
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
                    .TableCell(tcText, X, 11) = IIf(RS1!cpi_precos = 0, "", RS1!cpi_precos)
                    .TableCell(tcText, X, 10) = RS1!cpi_feccos
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
fg_descarga
Exit Function
Error_Productos:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_AporteProductos()
Dim i As Long, X As Long, fil As Long, Col As Long, VecCol(100) As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim codpro As String, NomPro As String, NomFan As String
Dim canser As Double, cannet As Double, totcanser As Double, totcannet As Double, totcosto As Double, canpro As Double
Dim alto As Long, cLin As String
On Local Error GoTo Error_AporProduc
fg_carga ""
MsgTitulo = "Informe de Aporte Ingredientes"
alto = 0
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
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
    .Header = "" & fg_poneencpagina & "||"
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = IIf(I_Produc.List1.SelCount > 9, 13500, 10500): .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 10: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Informe Ingredientes Aporte Nutricional /100gr"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .FontSize = 7
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    LogoEmp
    .TextAlign = taLeftTop
    .text = Chr(13): .text = Chr(13)
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
    For X = 0 To I_Produc.List1.listcount - 1
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
    RS1.Open "SELECT * FROM " & aAp & " WHERE tem_codpat='0'", vg_db, adOpenStatic
    .TableCell(tcCols) = 2 + I_Produc.List1.SelCount: .TableCell(tcRows) = 60000
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
        RS2.Open "SELECT * FROM " & aAp & " WHERE tem_codpat='" & Trim(RS1!tem_codigo) & "'", vg_db, adOpenStatic
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
fg_descarga
Exit Function
Error_AporProduc:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_ImpuestoProductos()
Dim i As Long, X As Long, fil As Long, Col As Long, VecCol(100) As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
Dim codpro As String, NomPro As String, NomFan As String, aAp As String
Dim canser As Double, cannet As Double, totcanser As Double, totcannet As Double, totcosto As Double, canpro As Double
Dim alto As Long, cLin As String
On Local Error GoTo Error_Impuestos
fg_carga ""
MsgTitulo = "Informe de Impuestos Productos"
alto = 0
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\Reporte.rtf"
    .ExportFile = vg_reporte
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
    .Header = "" & fg_poneencpagina & "||"
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = IIf(I_Produc.List1.SelCount > 7, 13500, 10500): .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 10: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Informe Productos Impuestos"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .FontSize = 8
    .Footer = "" & fg_ponepiepagina & "||Página : %d"
    ExportHeaderFooter Preview.VSPrinter
    LogoEmp
    .text = Chr(13): .text = Chr(13)
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
    For X = 0 To I_Produc.List1.listcount - 1
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
    RS1.Open "SELECT * FROM " & aAp & " WHERE tem_codpat = '0'", vg_db, adOpenStatic
    .TableCell(tcCols) = 2 + I_Produc.List1.SelCount: .TableCell(tcRows) = 60000
    cLin = ""
    fil = 1
    Do While Not RS1.EOF
        fil = fil + 1
        .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
        .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
        .TableCell(tcText, fil, 1) = RS1!tem_codigo
        .TableCell(tcText, fil, 2) = IIf(I_Produc.Option1(0).Value = True, RS1!tem_nombre, RS1!tem_nomfan)
        cLin = cLin & .TableCell(tcText, fil, 1) & "|" & .TableCell(tcText, fil, 2)
        RS2.Open "SELECT * FROM " & aAp & " WHERE tem_codpat = '" & Trim(RS1!tem_codigo) & "'", vg_db, adOpenStatic
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
fg_descarga
Exit Function
Error_Impuestos:
    fg_descarga
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
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.CuentaContable(1, "", ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: Exit Function
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 5500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
'    .TableCell(tcText, 1, 3) = "Cuenta Asignada"
    .TableCell(tcText, 1, 3) = ""
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 5500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1: .TableCell(tcAlign, , 3) = taLeftTop
    i = 1
    If Not RS1.EOF Then
        Do While Not RS1.EOF
           .TableCell(tcText, i, 1) = RS1!cta_codigo
           .TableCell(tcText, i, 2) = RS1!cta_nombre
           Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
           RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    RS1.Close: Set RS1 = Nothing: Close #1 'RS2.Close: Set RS2 = Nothing: Close #1
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

Public Function I_SalDevBod(Form As Object, Tipo As String)

Dim rutcli As String, NumDoc As Long, i As Long, j As Long, total As Double, aAp As String, titopc As String, codsec As String, coding As String
Dim numlin As Long, codmer As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String, codreg As Long, codser As Long
Dim tipsal As Boolean, totsec As Double, totmin As Double, totsmi As Double

On Local Error GoTo Error_SalDevBod

fg_carga ""
tipsal = False
If Tipo = "SP" Then
   
   MsgTitulo = "Informe de Salida a Producción"

ElseIf Tipo = "DP" Then
   
   MsgTitulo = "Informe de Devolución de Producción"

End If

'------- Consultar si salida es resumido ó sector
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT DISTINCT dev_codsec FROM b_detventas WHERE  dev_rutcli = '" & LimpiaDato(Trim(Form.fpText1(1).text)) & "' AND dev_tipdoc = '" & Tipo & "' AND dev_numdoc = " & Val(Form.fpLongInteger1(0).text) & "", vg_db, adOpenStatic
If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: MsgBox "No existe salida producción...", vbExclamation + vbOKOnly, MsgTitulo: Exit Function
MsgTitulo = MsgTitulo & IIf(IsNull(RS1!dev_codsec), " Resumido", " Sector")
titopc = IIf(IsNull(RS1!dev_codsec), " (Resumido)", " (Sector)")
tipsal = IIf(IsNull(RS1!dev_codsec), True, False)
RS1.Close: Set RS1 = Nothing
Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\" & vg_NUsr & "Reporte.rtf"
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
saldevbod1:
    vg_Archxls = ""
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = IIf(Tipo = "SP", "Salida de Bodega a Producción" & titopc, "Devolución de Producción" & titopc)
    .TableCell(tcText, 2, 1) = Form.Label1.Caption
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 2, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3300: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 4) = 4200: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 3) = True
    .TableCell(tcText, 1, 1) = "Folio"
    .TableCell(tcText, 1, 2) = Form.fpLongInteger1(0).text
    .TableCell(tcText, 1, 3) = "Contrato"
    .TableCell(tcText, 1, 4) = Trim(Form.fpText1(1).text) & " - " & Trim(Form.fpayuda(1).Caption)
    .TableCell(tcText, 2, 1) = "F. Emisión"
    .TableCell(tcText, 2, 2) = Form.fpDateTime1(0)
    .TableCell(tcText, 2, 3) = "Bodega"
    .TableCell(tcText, 2, 4) = Trim(Left(Form.Combo1(1).List(Form.Combo1(1).ListIndex), 50))
    .TableCell(tcText, 3, 1) = "F. Producción"
    .TableCell(tcText, 3, 2) = Form.fpDateTime1(1)
    .TableCell(tcText, 3, 3) = IIf(vg_tipser, "", "Servicios")
    .TableCell(tcText, 3, 4) = ""
    If Tipo = "DP" Then
'    If Not vg_tipser Then
       .TableCell(tcText, 3, 4) = Trim(Left(Form.Combo1(0).List(Form.Combo1(0).ListIndex), 50))
    Else
       .TableCell(tcText, 3, 4) = IIf(Not vg_tipser, Trim(Form.fpayuda(0).Caption) & " - " & Trim(Form.fpayuda(3).Caption), "")
    End If
'    .TableCell(tcText, 3, 4) = IIf(Not vg_tipser, Trim(Form.fpayuda(0).Caption) & " - " & Trim(Form.fpayuda(3).Caption), "")  'Trim(Left(Form.Combo1(0).List(Form.Combo1(0).ListIndex), 50)))
    Print #1, .TableCell(tcText, 1, 1) & "|"; .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    Print #1, .TableCell(tcText, 2, 1) & "|"; .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4)
    Print #1, .TableCell(tcText, 3, 1) & "|"; .TableCell(tcText, 3, 2) & "|" & .TableCell(tcText, 3, 3) & "|" & .TableCell(tcText, 3, 4)
    .TableBorder = tbBoxRows
    rutcli = Trim(LimpiaDato(Form.fpText1(1).text))
    NumDoc = Form.fpLongInteger1(0).text
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .FontSize = 7
    .StartTable
    .TableCell(tcCols) = 8: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 600: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 800: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcBackColor, 2) = vbYellow: .TableCell(tcFontBold, 2) = True: .TableCell(tcRowHeight, 2) = 230
    .TableCell(tcText, 1, 1) = ""
    .TableCell(tcText, 1, 2) = ""
    .TableCell(tcText, 1, 3) = ""
    .TableCell(tcText, 1, 4) = "Cantidad"
    .TableCell(tcText, 1, 5) = "Cantidad"
    .TableCell(tcText, 1, 6) = ""
    .TableCell(tcText, 1, 7) = "Total"
    .TableCell(tcText, 1, 8) = "Total"
    .TableCell(tcText, 2, 1) = "Código"
    .TableCell(tcText, 2, 2) = "Descripción"
    .TableCell(tcText, 2, 3) = "Unid."
    .TableCell(tcText, 2, 4) = IIf(Tipo = "SP", "Planif.", "Realizada")
    .TableCell(tcText, 2, 5) = IIf(Tipo = "SP", "Realizada", "Devolver")
    .TableCell(tcText, 2, 6) = "P.M.P."
    .TableCell(tcText, 2, 7) = "Planif."
    .TableCell(tcText, 2, 8) = "Realizada"
    Print #1, .TableCell(tcText, 1, 1) & "|"; .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|"; .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8)
    Print #1, .TableCell(tcText, 2, 1) & "|"; .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4) & "|" & _
              .TableCell(tcText, 2, 5) & "|"; .TableCell(tcText, 2, 6) & "|" & .TableCell(tcText, 2, 7) & "|" & .TableCell(tcText, 2, 8)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 8: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 600: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 800: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
saldevbod2:
    If tipsal = True Then
       
        If RS3.State = 1 Then RS3.Close
        RS3.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
       
       If vg_tipbase = "1" Then
          
          RS3.Open "SELECT ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 as sec_codigo, '' as sec_nombre, 0 as sec_orden, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
                   "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
                   "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_unidad uni " & _
                   "WHERE tov.tov_rutcli=dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
                   "AND   tov.tov_numdoc=dev.dev_numdoc AND dev.dev_coding=ing.ing_codigo " & _
                   "AND   ing.ing_unimed=unm.unm_codigo AND dev.dev_codmer=pro.pro_codigo " & _
                   "AND   pro.pro_coduni=uni.uni_codigo " & _
                   "AND   tov.tov_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).text)) & "' AND tov.tov_numdoc=" & Val(Form.fpLongInteger1(0).text) & " " & _
                   "AND   tov.tov_tipdoc='" & Tipo & "' AND tov.tov_codbod=" & vg_codbod & " " & _
                   "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin ORDER BY dev.dev_numlin " & _
                   "UNION ALL " & _
                   "SELECT 'estfij' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' AS unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, -1 AS sec_codigo, '' AS sec_nombre, 0 AS sec_orden, 0 AS canmin,  dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
                   "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
                   "WHERE tov.tov_rutcli=dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
                   "AND   tov.tov_numdoc=dev.dev_numdoc AND  dev.dev_codmer=pro.pro_codigo " & _
                   "AND   pro.pro_coduni=uni.uni_codigo AND  tov.tov_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).text)) & "' " & _
                   "AND   dev.dev_numdoc=" & Val(Form.fpLongInteger1(0).text) & " AND  tov.tov_tipdoc='" & Tipo & "' AND tov.tov_codbod=" & vg_codbod & " " & _
                   "AND  (dev.dev_coding='' OR ISNULL(dev.dev_coding) OR dev.dev_codsec = -1) ORDER BY dev.dev_numlin", vg_db, adOpenStatic
       Else
          
          RS3.Open "SELECT ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 as sec_codigo, '' as sec_nombre, 0 as sec_orden, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
                   "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
                   "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_unidad uni " & _
                   "WHERE tov.tov_rutcli=dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
                   "AND   tov.tov_numdoc=dev.dev_numdoc AND dev.dev_coding=ing.ing_codigo " & _
                   "AND   ing.ing_unimed=unm.unm_codigo AND dev.dev_codmer=pro.pro_codigo " & _
                   "AND   pro.pro_coduni=uni.uni_codigo " & _
                   "AND   tov.tov_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).text)) & "' AND tov.tov_numdoc=" & Val(Form.fpLongInteger1(0).text) & " " & _
                   "AND   tov.tov_tipdoc='" & Tipo & "' AND tov.tov_codbod=" & vg_codbod & " " & _
                   "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin " & _
                   "UNION ALL " & _
                   "SELECT 'estfij' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' AS unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, -1 AS sec_codigo, '' AS sec_nombre, 0 AS sec_orden, 0 AS canmin,  dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
                   "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
                   "WHERE tov.tov_rutcli=dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
                   "AND   tov.tov_numdoc=dev.dev_numdoc AND  dev.dev_codmer=pro.pro_codigo " & _
                   "AND   pro.pro_coduni=uni.uni_codigo AND  tov.tov_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).text)) & "' " & _
                   "AND   dev.dev_numdoc=" & Val(Form.fpLongInteger1(0).text) & " AND  tov.tov_tipdoc='" & Tipo & "' AND tov.tov_codbod=" & vg_codbod & " " & _
                   "AND  (dev.dev_coding='' OR (dev.dev_coding) IS NULL OR dev.dev_codsec = -1) ORDER BY dev.dev_numlin", vg_db, adOpenStatic
       End If
    ElseIf tipsal = False Then
       
        If RS3.State = 1 Then RS3.Close
        RS3.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
       
       If vg_tipbase = "1" Then
          RS3.Open "SELECT ing.ing_codigo, ing.ing_nombre,unm.unm_nomcor, sec.sec_codigo, sec.sec_nombre, sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
                   "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
                   "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_sector sec, a_unidad uni " & _
                   "WHERE tov.tov_rutcli=dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
                   "AND   tov.tov_numdoc=dev.dev_numdoc AND dev.dev_coding=ing.ing_codigo " & _
                   "AND   ing.ing_unimed=unm.unm_codigo AND dev.dev_codmer=pro.pro_codigo " & _
                   "AND   dev.dev_codsec=sec.sec_codigo AND pro.pro_coduni=uni.uni_codigo " & _
                   "AND   tov.tov_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).text)) & "' AND tov.tov_numdoc=" & Val(Form.fpLongInteger1(0).text) & " " & _
                   "AND   tov.tov_tipdoc='" & Tipo & "' AND tov.tov_codbod=" & vg_codbod & " " & _
                   "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor,  sec.sec_codigo, sec.sec_nombre,  sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin ORDER BY sec.sec_orden, dev.dev_numlin " & _
                   "UNION ALL " & _
                   "SELECT '' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' as unm_nomcor, -1 AS sec_codigo, 'Estructura Fija' AS sec_nombre, 999999999 AS sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 AS canmin,  dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
                   "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
                   "WHERE tov.tov_rutcli=dev.dev_rutcli AND  tov.tov_tipdoc=dev.dev_tipdoc " & _
                   "AND   tov.tov_numdoc=dev.dev_numdoc AND  dev.dev_codmer=pro.pro_codigo " & _
                   "AND   pro.pro_coduni=uni.uni_codigo AND  tov.tov_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).text)) & "' " & _
                   "AND   dev.dev_numdoc=" & Val(Form.fpLongInteger1(0).text) & " AND tov.tov_tipdoc='" & Tipo & "' AND tov.tov_codbod=" & vg_codbod & " " & _
                   "AND  (dev.dev_coding='' OR (dev.dev_coding) IS NULL OR dev.dev_codsec = -1) ORDER BY sec_orden, dev.dev_numlin", vg_db, adOpenStatic
       Else
          RS3.Open "SELECT ing.ing_codigo, ing.ing_nombre,unm.unm_nomcor, sec.sec_codigo, sec.sec_nombre, sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
                   "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
                   "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_sector sec, a_unidad uni " & _
                   "WHERE tov.tov_rutcli=dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
                   "AND   tov.tov_numdoc=dev.dev_numdoc AND dev.dev_coding=ing.ing_codigo " & _
                   "AND   ing.ing_unimed=unm.unm_codigo AND dev.dev_codmer=pro.pro_codigo " & _
                   "AND   dev.dev_codsec=sec.sec_codigo AND pro.pro_coduni=uni.uni_codigo " & _
                   "AND   tov.tov_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).text)) & "' AND tov.tov_numdoc=" & Val(Form.fpLongInteger1(0).text) & " " & _
                   "AND   tov.tov_tipdoc='" & Tipo & "' AND tov.tov_codbod=" & vg_codbod & " " & _
                   "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor,  sec.sec_codigo, sec.sec_nombre,  sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin " & _
                   "UNION ALL " & _
                   "SELECT '' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' as unm_nomcor, -1 AS sec_codigo, 'Estructura Fija' AS sec_nombre, 999999999 AS sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 AS canmin,  dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
                   "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
                   "WHERE tov.tov_rutcli=dev.dev_rutcli AND  tov.tov_tipdoc=dev.dev_tipdoc " & _
                   "AND   tov.tov_numdoc=dev.dev_numdoc AND  dev.dev_codmer=pro.pro_codigo " & _
                   "AND   pro.pro_coduni=uni.uni_codigo AND  tov.tov_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).text)) & "' " & _
                   "AND   dev.dev_numdoc=" & Val(Form.fpLongInteger1(0).text) & " AND tov.tov_tipdoc='" & Tipo & "' AND tov.tov_codbod=" & vg_codbod & " " & _
                   "AND  (dev.dev_coding='' OR (dev.dev_coding) IS NULL OR dev.dev_codsec = -1) ORDER BY sec_orden, dev.dev_numlin", vg_db, adOpenStatic
       End If
       If Tipo = "DP" Then
          codreg = Val(Mid(Form.Combo1(0), Len(Trim(Form.Combo1(0).text)) - 22, 10))
          codser = Val(Mid(Form.Combo1(0), Len(Trim(Form.Combo1(0).text)) - 10, 10))
       ElseIf Tipo = "SP" Then
          codreg = Val(Form.fpLongInteger1(1).Value) 'Val(Mid(Form.Combo1(0), Len(Trim(Form.Combo1(0).text)) - 22, 10))
          codser = Val(Form.fpLongInteger1(2).Value) 'Val(Mid(Form.Combo1(0), Len(Trim(Form.Combo1(0).text)) - 10, 10))
       End If
       
        If RS4.State = 1 Then RS4.Close
        RS4.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
       
       RS4.Open "SELECT DISTINCT rec.rec_codigo, rec.rec_nombre, sec.sec_codigo, sec.sec_orden, sec.sec_nombre, mid.mid_numrac " & _
                "FROM b_minuta mi, b_minutadet mid, b_receta rec, b_recetadet red, a_servicio ser, a_estservicio ess, a_sector sec " & _
                "WHERE rec.rec_codigo=mid.mid_codrec AND rec.rec_codigo=red.red_codigo AND red.red_tiprec=mid.mid_tiprec AND ((red.red_tiprec<>0 AND red.red_cencos='" & MuestraCasino(1) & "') OR (red.red_tiprec=0 AND red.red_cencos='0')) " & _
                "AND   mi.min_codigo=mid.mid_codigo AND mi.min_codser=ser.ser_codigo AND mi.min_codser=" & codser & " AND mi.min_cencos=ess.ess_cencos AND mid.mid_estser=ess.ess_codigo AND mi.min_codser=ess.ess_codser AND ess.ess_codsec=sec.sec_codigo " & _
                "AND   mi.min_cencos='" & Trim(LimpiaDato(Form.fpText1(1).text)) & "' AND mi.min_codreg=" & codreg & " AND mi.min_fecmin=" & Format(Form.fpDateTime1(1).text, "yyyymmdd") & " AND mid.mid_tipmin='2' AND mid.mid_numrac>0 AND red.red_canpro>0 " & _
                "ORDER BY sec.sec_orden", vg_db, adOpenForwardOnly
    
'                "AND   mi.min_cencos='" & Trim(LimpiaDato(Form.fpText1(1).Text)) & "' AND mi.min_codreg=" & codreg & " AND mi.min_fecmin=" & Format(Form.fpDateTime1(1).Text, "yyyymmdd") & " AND mid.mid_tipmin='2' AND red.red_canpro>0 " &

    End If
    i = 1: j = 1: total = 0: totsec = 0: totsmi = 0: totmin = 0
    Do While Not RS3.EOF
       If Not tipsal And codsac <> RS3!sec_codigo Then
          If Not tipsal And i > 1 Then
             i = i + 1
             .TableCell(tcFontSize, j, 7) = 8: .TableCell(tcText, j, 7) = Format(totsmi, fg_Pict(9, vg_DPr)): totsmi = 0
             .TableCell(tcFontSize, j, 8) = 8: .TableCell(tcText, j, 8) = Format(totsec, fg_Pict(9, vg_DPr)): j = i: totsec = 0
          End If
          .TableCell(tcFontBold, i) = True
          .TableCell(tcFontSize, i, 1) = 8: .TableCell(tcText, i, 1) = IIf(RS3!sec_codigo = -1, "", RS3!sec_codigo)
          .TableCell(tcFontSize, i, 2) = 8: .TableCell(tcText, i, 2) = RS3!sec_nombre
          .TableCell(tcText, i, 3) = ""
          .TableCell(tcText, i, 4) = ""
          Print #1, .TableCell(tcText, i, 1) & "|"; .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
          codsac = RS3!sec_codigo
          coding = ""
          If Not RS4.EOF Then
             i = i + 1
             Do While Not RS4.EOF
                If RS4!sec_codigo = RS3!sec_codigo Then
                   .TableCell(tcColSpan, i, 2) = 8
                   .TableCell(tcFontBold, i, 2) = True: .TableCell(tcText, i, 2) = "(" & Trim(RS4!rec_codigo) & ") " & Trim(RS4!rec_nombre) & " [Nº.Rac. " & RS4!mid_numrac & "]"
                   i = i + 1
                End If
                RS4.MoveNext
             Loop
             RS4.MoveFirst
          End If
          i = i + IIf(i > 1, 2, 1)
       End If
       '------- Ingrediente
        If coding <> RS3!ing_codigo Then
           If Form.Check1(0).Value = 0 Then
              .TableCell(tcFontBold, i) = True
              .TableCell(tcText, i, 1) = Trim(RS3!ing_codigo)
              .TableCell(tcText, i, 2) = Trim(RS3!ing_nombre)
              .TableCell(tcText, i, 3) = Trim(RS3!unm_nomcor)
              .TableCell(tcText, i, 4) = IIf(RS3!ing_codigo = "", "", Format(RS3!canmin, fg_Pict(9, vg_DCa)))
              Print #1, .TableCell(tcText, i, 1) & "|"; .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3)
              i = i + 1
           End If
           coding = RS3!ing_codigo
        End If
        '------- Producto
        .TableCell(tcText, i, 1) = Trim(RS3!pro_codigo)
        .TableCell(tcText, i, 2) = Trim(RS3!pro_nombre)
        .TableCell(tcText, i, 3) = Trim(RS3!uni_nomcor)
        .TableCell(tcText, i, 4) = Format(RS3!dev_canmin, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa))) 'vg_DCa))
        .TableCell(tcText, i, 5) = Format(RS3!dev_canmer, fg_Pict(9, vg_DCa))
        .TableCell(tcText, i, 6) = Format(RS3!dev_predoc, fg_Pict(9, vg_DPr))
        .TableCell(tcText, i, 7) = Format((RS3!dev_canmin * RS3!dev_predoc), fg_Pict(9, vg_DPr))
        .TableCell(tcText, i, 8) = Format(RS3!dev_ptotal, fg_Pict(9, vg_DPr))
        Print #1, .TableCell(tcText, i, 1) & "|"; .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                  .TableCell(tcText, i, 5) & "|"; .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8)
        If RS3!dev_canmer <> 0 Then total = total + Format(RS3!dev_ptotal, fg_Pict(9, 2))
        totmin = Round(totmin + Format((RS3!dev_canmin * RS3!dev_predoc), fg_Pict(9, vg_DPr))) '(RS3!dev_canmin * RS3!dev_predoc)
        If Not tipsal Then
           If RS3!dev_canmer <> 0 Then totsec = (totsec + RS3!dev_ptotal)
           totsmi = Round(totsmi + Format((RS3!dev_canmin * RS3!dev_predoc), fg_Pict(9, vg_DPr)))
        End If
        RS3.MoveNext: i = i + 1
        .TableCell(tcText, i, 1) = ""
    Loop
    RS3.Close: Set RS3 = Nothing
    If Not tipsal Then RS4.Close: Set RS4 = Nothing
    If Not tipsal Then
       .TableCell(tcFontSize, j, 7) = 8
       .TableCell(tcText, j, 7) = Format(totsmi, fg_Pict(9, vg_DPr))
       .TableCell(tcFontSize, j, 8) = 8
       .TableCell(tcText, j, 8) = Format(totsec, fg_Pict(9, vg_DPr))
       j = i: totsec = 0: totsmi = 0
    End If
    .TableCell(tcRows) = i - 1
    .PenColor = &HC0C0C0
    .TableBorder = tbBottom
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcFontSize, 1, 1) = 8: .TableCell(tcFontBold, 1, 1, 1, 4) = True
    .TableCell(tcColWidth, 1, 1) = 7600: .TableCell(tcAlign, , 1) = taRightTop
    .TableCell(tcColWidth, 1, 2) = 1000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, 1, 3) = 1000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, 1, 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcText, 1, 1) = "Totales "
    .TableCell(tcFontSize, 1, 3) = 8: .TableCell(tcText, 1, 3) = Format(totmin, fg_Pict(9, vg_DPr))
    .TableCell(tcFontSize, 1, 4) = 8: .TableCell(tcText, 1, 4) = Format(total, fg_Pict(9, vg_DPr))
    Print #1, "|||||" & .TableCell(tcText, 1, 2) & "|"; .TableCell(tcText, 1, 3) & "|"; .TableCell(tcText, 1, 4)
    .TableBorder = tbNone
    Print #1, " ": Print #1, "|||||" & "_____________________"
    Print #1, " "
    If Tipo = "SP" Then
        Print #1, Space(100) & "Entregado conforme"
    Else
        Print #1, Space(100) & "Recibido conforme"
    End If
    .EndTable
    
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 6
    .TableCell(tcColWidth, , 1) = 7700
    .TableCell(tcColWidth, , 2) = 2800: .TableCell(tcAlign, , 2) = taCenterTop
    .TableCell(tcFontBold) = True
    .TableCell(tcFontUnderline, 5, 2) = True
    .TableCell(tcText, 5, 2) = Space(40)
    .TableCell(tcText, 6, 2) = IIf(Tipo = "SP", "Entregado conforme", "Recibido conforme")
    .TableBorder = tbNone
    .EndTable
    '.CurrentX = 9200
    '.CurrentY = .CurrentY - 400
    '.Text = IIf(tipo = "SP", "_____________________", "____________________")
    .EndDoc
    Close #1
End With
Preview.Show 1
fg_descarga
Exit Function
Error_SalDevBod:
    fg_descarga
    If Err = 55 Then Close #1: GoTo saldevbod1
    If Err = -2147467259 Then GoTo saldevbod2 'Resume
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Mermas(Form As Object)
Dim rutcli As String, NumDoc As Long, i As Long, total As Double
Dim numlin As Long, codmer As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String
Dim codmerma As Long
Dim desmerma As String

On Local Error GoTo Error_Mermas1

fg_carga ""
MsgTitulo = "Mermas"
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
    .text = Chr(13): .text = Chr(13)
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS1.Open "select distinct c.aju_codigo, c.aju_nombre " & _
             "from b_totventas as a with (nolock) " & _
             "inner join b_detventas b with (nolock) on a.tov_rutcli = b.dev_rutcli " & _
                                                   "and a.tov_numdoc = b.dev_numdoc " & _
                                                   "and a.tov_tipdoc = b.dev_tipdoc " & _
             "inner join a_tipoajuste as c with (nolock) on c.aju_codigo = a.tov_codser " & _
             "where a.tov_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).text)) & "' " & _
             "and   a.tov_tipdoc='ME' " & _
             "and   a.tov_numdoc=" & Val(Form.fpLongInteger1(0).text) & "", vg_db, adOpenStatic
    If Not RS1.EOF Then
    
       codmerma = RS1!aju_codigo
       nommerma = RS1!aju_nombre
    
    End If
    RS1.Close: Set RS1 = Nothing

    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3800: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 3700: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 3) = True
    .TableCell(tcText, 1, 1) = "Folio"
    .TableCell(tcText, 1, 2) = ": " & Form.fpLongInteger1(0).text
    .TableCell(tcText, 1, 3) = "Bodega"
    .TableCell(tcText, 1, 4) = ": " & Trim(Left(Form.Combo1(1).List(Form.Combo1(1).ListIndex), 50))
    .TableCell(tcText, 2, 1) = "F. Emisión"
    .TableCell(tcText, 2, 2) = ": " & Form.fpDateTime1(0)
    .TableCell(tcText, 2, 3) = "Tipo de Merma"
    .TableCell(tcText, 2, 4) = ": " & codmerma & " - " & nommerma
    .TableCell(tcText, 3, 1) = "Contrato"
    .TableCell(tcText, 3, 2) = ": " & Trim(LimpiaDato(Form.fpText1(1).text)) & " - " & Trim(Form.fpayuda(1).Caption)
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbBox
    rutcli = Trim(LimpiaDato(Form.fpText1(1).text))
    NumDoc = Form.fpLongInteger1(0).text
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
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
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "select dev.*, uni.uni_nombre from b_detventas dev, b_productos pro, a_unidad uni " & _
             "where dev.dev_rutcli='" & Trim(LimpiaDato(Form.fpText1(1).text)) & "' and dev.dev_tipdoc='ME' " & _
             "and dev.dev_numdoc=" & Val(Form.fpLongInteger1(0).text) & " and dev.dev_codmer=pro.pro_codigo " & _
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
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 6
    .TableCell(tcColWidth, , 1) = 6500
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taCenterTop
    .TableCell(tcFontBold) = True
    .TableCell(tcFontUnderline, 5, 2) = True
    .TableCell(tcText, 5, 2) = Space(40)
    .TableCell(tcText, 6, 2) = "Entregado conforme"
    .TableBorder = tbNone
    .EndTable
'    .FontBold = True
'    .CurrentX = 8800
'    .CurrentY = 14000
'    .Text = "_____________________"
'    .CurrentX = 8950
'    .CurrentY = 14200
'    .Text = "Entregado conforme"
    .EndDoc
    Close #1
End With
Preview.Show 1
fg_descarga
Exit Function
Error_Mermas1:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Traspaso(Form As Object)
Dim rutcli As String, NumDoc As Long, i As Long, total As Double
Dim numlin As Long, codmer As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String
Dim cLin As String
On Local Error GoTo Error_Traspaso
fg_carga ""
MsgTitulo = "Traspasos"
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
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = "Traspaso entre Contratos"
    .TableCell(tcText, 2, 1) = Form.Label1.Caption
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 2, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3800: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 4) = 3700: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 3) = True
    .TableCell(tcText, 1, 1) = "Folio"
    .TableCell(tcText, 1, 2) = Form.fpLongInteger1(0).text
    .TableCell(tcText, 1, 3) = "Bodega"
    .TableCell(tcText, 1, 4) = Trim(Left(Form.Combo1(1).List(Form.Combo1(1).ListIndex), 50))
    .TableCell(tcText, 2, 1) = "F. Emisión"
    .TableCell(tcText, 2, 2) = Form.fpDateTime1(0)
    .TableCell(tcText, 2, 3) = "Tipo Traspaso"
    .TableCell(tcText, 2, 4) = IIf(Form.Option1(1).Value = True, "Entrada", "Salida")
    .TableCell(tcText, 3, 1) = "Contrato"
    .TableCell(tcText, 3, 2) = Trim(LimpiaDato(Form.fpText1(0).text)) & " - " & Trim(Form.fpayuda(0).Caption)
    .TableCell(tcText, 3, 3) = Trim(Form.Label3(0).Caption)
    .TableCell(tcText, 3, 4) = Trim(LimpiaDato(Form.fpText1(1).text)) & " - " & Trim(Form.fpayuda(1).Caption)
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2) & "|" & .TableCell(tcText, 3, 3) & "|" & .TableCell(tcText, 3, 4)
    .TableBorder = tbBox
    rutcli = Trim(LimpiaDato(Form.fpText1(1).text))
    NumDoc = Form.fpLongInteger1(0).text
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    cLin = ""
    .StartTable
    .TableCell(tcCols) = IIf(Form.Option1(1).Value = True, 7, 6): .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = IIf(Form.Option1(1).Value = True, 1500, 2000): .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = IIf(Form.Option1(1).Value = True, 3000, 3500): .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1200: .TableCell(tcAlign, , 4) = taRightTop
    
    If Form.Option1(1).Value = True Then .TableCell(tcColWidth, , 5) = 1200: .TableCell(tcAlign, , 5) = taRightTop
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
    .TableCell(tcColWidth, , 4) = 1200: .TableCell(tcAlign, , 4) = taRightTop
    If Form.Option1(1).Value = True Then .TableCell(tcColWidth, , 5) = 1200: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , IIf(Form.Option1(1).Value = True, 6, 5)) = 1500: .TableCell(tcAlign, , IIf(Form.Option1(1).Value = True, 6, 5)) = taRightTop
    .TableCell(tcColWidth, , IIf(Form.Option1(1).Value = True, 7, 6)) = 1500: .TableCell(tcAlign, , IIf(Form.Option1(1).Value = True, 7, 6)) = taRightTop
    RS1.Open "SELECT dev.*, uni.uni_nombre from b_detventas dev, b_productos pro, a_unidad uni " & _
             "WHERE dev.dev_rutcli = '" & Trim(LimpiaDato(Form.fpText1(0).text)) & "' AND dev.dev_tipdoc = 'TR' " & _
             "AND dev.dev_numdoc = " & Val(Form.fpLongInteger1(0).text) & " AND dev.dev_codmer = pro.pro_codigo " & _
             "AND pro.pro_coduni = uni.uni_codigo ORDER BY dev.dev_numlin", vg_db, adOpenStatic
    i = 1: total = 0
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = RS1!dev_codmer
        .TableCell(tcText, i, 2) = RS1!dev_descri
        .TableCell(tcText, i, 3) = RS1!uni_nombre
        .TableCell(tcText, i, 4) = Format(IIf(Form.Option1(1).Value = True, RS1!dev_canmin, RS1!dev_canmer), fg_Pict(9, vg_DCa))
        'cLin = cLin & .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
        If Form.Option1(1).Value = True Then .TableCell(tcText, i, 5) = Format(RS1!dev_canmer, fg_Pict(9, vg_DCa)): cLin = cLin & "|" & .TableCell(tcText, i, 5) & "|"
        .TableCell(tcText, i, IIf(Form.Option1(1).Value = True, 6, 5)) = Format(RS1!dev_predoc, fg_Pict(9, 2)): cLin = cLin & .TableCell(tcText, i, IIf(Form.Option1(1).Value = True, 6, 5))
        .TableCell(tcText, i, IIf(Form.Option1(1).Value = True, 7, 6)) = Format(RS1!dev_ptotal, fg_Pict(9, vg_DPr)): cLin = cLin & "|" & .TableCell(tcText, i, IIf(Form.Option1(1).Value = True, 7, 6))
        total = total + Format(RS1!dev_ptotal, fg_Pict(9, 2))
        'Print #1, cLin
        RS1.MoveNext: i = i + 1
    Loop
    For j = 1 To i - 1
        cLin = ""
        For K = 1 To .TableCell(tcCols)
            cLin = cLin & .TableCell(tcText, j, K) & "|"
        Next K
        Print #1, cLin
    Next j
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcRows) = i - 1
    .PenColor = &HC0C0C0
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcFontBold, 1, 1, 1, 3) = True
    .TableCell(tcColWidth, 1, 1) = IIf(Form.Option1(1).Value = True, 7900, 7700): .TableCell(tcAlign, , 1) = taLeftTop
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
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 6
    .TableCell(tcColWidth, , 1) = 6500
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taCenterTop
    .TableCell(tcFontBold) = True
    .TableCell(tcFontUnderline, 5, 2) = True
    .TableCell(tcText, 5, 2) = Space(40)
    .TableCell(tcText, 6, 2) = IIf(Form.Option1(1).Value = True, "Recibido conforme", "Entregado conforme")
    .TableBorder = tbNone
    .EndTable
'    .FontBold = True
'    .CurrentX = 8800
'    .CurrentY = 14000
'    .Text = IIf(Form.Option1(1).Value = True, "____________________", "_____________________")
'    .CurrentX = 8950
'    .CurrentY = 14200
'    .Text = IIf(Form.Option1(1).Value = True, "Recibido conforme", "Entregado conforme")
    .EndDoc
    Close #1
End With
Preview.Show 1
fg_descarga
Exit Function
Error_Traspaso:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_VentaDir(Form As Object, Tipo As String)
Dim rutcli As String, NumDoc As Long, i As Long, total As Double
Dim numlin As Long, codmer As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String
On Local Error GoTo Error_VtaDir
fg_carga ""
MsgTitulo = "Venta Directa"
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
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = IIf(Tipo = "FA", "Factura", "Guia Despacho") & " - Venta Directa"
    .TableCell(tcText, 2, 1) = Form.Label1.Caption
    Print #1, .TableCell(tcText, 1, 1)
    Print #1, .TableCell(tcText, 2, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
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
    .TableCell(tcText, 1, 4) = Form.fpLongInteger1(0).text
    .TableCell(tcText, 2, 1) = "Bodega"
    .TableCell(tcText, 2, 2) = Trim(Left(Form.Combo1(1).List(Form.Combo1(1).ListIndex), 50))
    .TableCell(tcText, 2, 3) = "F. Emisión"
    .TableCell(tcText, 2, 4) = Form.fpDateTime1(0)
    .TableCell(tcText, 3, 1) = "Contrato"
    .TableCell(tcText, 3, 2) = Trim(LimpiaDato(Form.fpText1(0).text)) & " - " & Trim(Form.fpayuda(0).Caption)
    .TableCell(tcText, 3, 3) = "Cliente"
    .TableCell(tcText, 3, 4) = Trim(LimpiaDato(Form.fpText1(1).text)) & " - " & Trim(Form.fpayuda(1).Caption)
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2) & "|" & .TableCell(tcText, 3, 3) & "|" & .TableCell(tcText, 3, 4)
    .TableBorder = tbBox
    rutcli = Trim(LimpiaDato(Form.fpText1(1).text))
    NumDoc = Form.fpLongInteger1(0).text
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
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
             "where dev.dev_rutcli='" & Trim(LimpiaDato(Form.fpText1(0).text)) & "' and dev.dev_tipdoc='" & Tipo & "' " & _
             "and dev.dev_numdoc=" & Val(Form.fpLongInteger1(0).text) & " and dev.dev_codmer=pro.pro_codigo " & _
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
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 6
    .TableCell(tcColWidth, , 1) = 6900
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taCenterTop
    .TableCell(tcFontBold) = True
    .TableCell(tcFontUnderline, 5, 2) = True
    .TableCell(tcText, 5, 2) = Space(40)
    .TableCell(tcText, 6, 2) = "Entregado conforme"
    .TableBorder = tbNone
    .EndTable
'    .FontBold = True
'    .CurrentX = 8800
'    .CurrentY = 14000
'    .Text = "_____________________"
'    .CurrentX = 8950
'    .CurrentY = 14200
'    .Text = "Entregado conforme"
    .EndDoc
    Close #1
End With
Preview.Show 1
fg_descarga
Exit Function
Error_VtaDir:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_VentaCafeteria(Form As Object)
Dim rutcli As String, NumDoc As Long, i As Long, total As Double
Dim numlin As Long, codmer As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String
Dim RS1 As New ADODB.Recordset
On Local Error GoTo Error_VtaDir
fg_carga ""
MsgTitulo = "Venta Cafetería"
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
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = "Venta Cafetería"
    .TableCell(tcText, 2, 1) = Form.Label1.Caption
    Print #1, .TableCell(tcText, 1, 1)
    Print #1, .TableCell(tcText, 2, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 3) = True
    .TableCell(tcText, 1, 1) = "Fecha"
    .TableCell(tcText, 1, 2) = Trim(Form.fpDateTime1(0).text)
    .TableCell(tcText, 2, 1) = "Contrato"
    .TableCell(tcText, 2, 2) = Trim(LimpiaDato(Form.fpText1(0).text)) & " - " & Trim(Form.fpayuda(0).Caption)
    .TableCell(tcText, 3, 1) = "Bodega"
    .TableCell(tcText, 3, 2) = Trim(Left(Form.Combo1(0).List(Form.Combo1(0).ListIndex), 50))
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbBox
    .TableBorder = tbNone
    .EndTable
    .FontSize = 7
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1800: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 3200: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taLeftTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1500: .TableCell(tcAlign, , 7) = taRightTop
    
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 200
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Artículo"
    .TableCell(tcText, 1, 3) = "Tipo Pago"
    .TableCell(tcText, 1, 4) = "Cliente"
    .TableCell(tcText, 1, 5) = "Centro Costo"
    .TableCell(tcText, 1, 6) = "Cantidad"
    .TableCell(tcText, 1, 7) = "Precio Venta"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7)
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1800: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 3200: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taLeftTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1500: .TableCell(tcAlign, , 7) = taRightTop
    If vg_tipbase = "1" Then
       RS1.Open "SELECT b_detventascaf.*, b_totpreciocaf.tpc_nombre, b_clientes.cli_nombre FROM (b_detventascaf LEFT JOIN b_clientes ON b_detventascaf.dvc_rutcli=b_clientes.cli_codigo) INNER JOIN b_totpreciocaf ON b_detventascaf.dvc_articulo=b_totpreciocaf.tpc_codigo AND b_detventascaf.dvc_cencos=b_totpreciocaf.tpc_cencos " & _
                "WHERE b_detventascaf.dvc_cencos = '" & Trim(LimpiaDato(Form.fpText1(0).text)) & "' AND b_detventascaf.dvc_fecing = cdate('" & Format(Form.fpDateTime1(0).text, "dd/mm/yyyy") & "') ORDER BY b_detventascaf.dvc_numlin", vg_db, adOpenStatic
    Else
       RS1.Open "SELECT b_detventascaf.*, b_totpreciocaf.tpc_nombre, b_clientes.cli_nombre FROM (b_detventascaf LEFT JOIN b_clientes ON b_detventascaf.dvc_rutcli=b_clientes.cli_codigo) INNER JOIN b_totpreciocaf ON b_detventascaf.dvc_articulo=b_totpreciocaf.tpc_codigo AND b_detventascaf.dvc_cencos=b_totpreciocaf.tpc_cencos " & _
                "WHERE b_detventascaf.dvc_cencos = '" & Trim(LimpiaDato(Form.fpText1(0).text)) & "' AND b_detventascaf.dvc_fecing = '" & Format(Form.fpDateTime1(0).text, "yyyymmdd") & "' ORDER BY b_detventascaf.dvc_numlin", vg_db, adOpenStatic
    End If
    i = 1: total = 0
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = Trim(IIf(IsNull(RS1!dvc_articulo), "", RS1!dvc_articulo))
        .TableCell(tcText, i, 2) = Trim(IIf(IsNull(RS1!tpc_nombre), "", RS1!tpc_nombre))
        .TableCell(tcText, i, 3) = Trim(IIf(IsNull(RS1!dvc_tippag), "", IIf(RS1!dvc_tippag = "CO", "CONTADO", "CREDITO")))
        .TableCell(tcText, i, 4) = fg_PintaRut(Trim(IIf(IsNull(RS1!dvc_rutcli), "", RS1!dvc_rutcli))) & " " & Trim(IIf(IsNull(RS1!cli_nombre), "", RS1!cli_nombre))
        .TableCell(tcText, i, 5) = Trim(IIf(IsNull(RS1!dvc_cencli), "", RS1!dvc_cencli))
        .TableCell(tcText, i, 6) = Trim(IIf(IsNull(RS1!dvc_canart), "", Format(RS1!dvc_canart, fg_Pict(9, vg_DCa))))
        .TableCell(tcText, i, 7) = Trim(IIf(IsNull(RS1!dvc_precio), "", Format(RS1!dvc_precio, fg_Pict(9, vg_DPr))))
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                  .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)
        total = total + (IIf(IsNull(RS1!dvc_precio), 0, RS1!dvc_precio))
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
    .TableCell(tcColWidth, 1, 1) = 8000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, 1, 2) = 1000: .TableCell(tcAlign, , 2) = taLeftTop
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
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 6
    .TableCell(tcColWidth, , 1) = 6000
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taCenterTop
    .TableCell(tcFontBold) = True
    .TableCell(tcFontUnderline, 5, 2) = True
    .TableCell(tcText, 5, 2) = Space(40)
    .TableCell(tcText, 6, 2) = "Entregado conforme"
    .TableBorder = tbNone
    .EndTable
'    .FontBold = True
'    .CurrentX = 8800
'    .CurrentY = 14000
'    .Text = "_____________________"
'    .CurrentX = 8950
'    .CurrentY = 14200
'    .Text = "Entregado conforme"
    .EndDoc
    Close #1
End With
Preview.Show 1
fg_descarga
Exit Function
Error_VtaDir:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_TipAju()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_TipAju
fg_carga ""
MsgTitulo = "Informe Tipo de Ajuste"
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
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.TipoAjuste(2, 0, "", 1), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
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
Preview.Show 1
fg_descarga
Exit Function
Error_TipAju:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_TipMer()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_Mermas
fg_carga ""
MsgTitulo = "Informe Tipo de Mermas"
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
    .text = Chr(13): .text = Chr(13)
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS1.Open RutinaLectura.TipoAjuste(2, 0, "", 0), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
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
Preview.Show 1
fg_descarga
Exit Function
Error_Mermas:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Sector()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_sector
fg_carga ""
MsgTitulo = "Informe Sector"
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Sector"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.Sector(2, 0, ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Function
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 5000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Descripción"
    .TableCell(tcText, 1, 3) = "Orden"
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
            .TableCell(tcText, i, 1) = RS1!sec_codigo
            .TableCell(tcText, i, 2) = RS1!sec_nombre
            .TableCell(tcText, i, 3) = RS1!sec_orden
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
Preview.Show 1
fg_descarga
Exit Function
Error_sector:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Toma1(Sql As String, opfampro As Boolean)
Dim i As Long, codtip As Long
On Local Error GoTo Error_Toma1
fg_carga ""
MsgTitulo = "Toma de Inventario"

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open Sql, vg_db, adOpenStatic
If RS1.EOF Then
    RS1.Close: Set RS1 = Nothing
    fg_descarga
    MsgBox "No existe información...", vbExclamation + vbOKOnly, MsgTitulo
    Exit Function
End If
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
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcFontSize, 1, 1, 3, 2) = 8: .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 2) = False
    .TableCell(tcText, 1, 1) = "Bodega": .TableCell(tcText, 1, 2) = Left(M_TomInv.Combo1(0).text, 50)
    .TableCell(tcText, 2, 1) = "Toma de Inventario": .TableCell(tcText, 2, 2) = M_TomInv.Date1(0).text
    .TableCell(tcText, 3, 1) = "Familia de Producto": .TableCell(tcText, 3, 2) = IIf(I_TomInv.optTIPPRO(0).Value = True, Left(I_TomInv.Combo1(0).text, 50), "Todas")
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 5000: .TableCell(tcAlign, , 2) = taLeftTop
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
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 60000
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 5000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 2000: .TableCell(tcAlign, , 4) = taCenterTop
    i = 1
    codtip = 0
    If Not RS1.EOF Then
        Do While Not RS1.EOF
           If RS1!pro_codtip <> codtip And opfampro Then
              If codtip > 0 Then i = i + 1
              .TableCell(tcColSpan, i, 1) = 5
              .TableCell(tcText, i, 1) = fg_BuscaenArbol(RS1!pro_codtip, "a_tipopro", "tip_codigo")
              .TableCell(tcFontBold, i, 1) = True
              Print #1, .TableCell(tcText, i, 1)
              i = i + 1
              codtip = RS1!pro_codtip
           End If
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
fg_descarga
Exit Function
Error_Toma1:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Toma2(Sql As String, opfampro As Boolean)
Dim i As Long, codtip As Long
On Local Error GoTo Error_Toma2
fg_carga ""
MsgTitulo = "Toma de Inventario"

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open Sql, vg_db, adOpenStatic
If RS1.EOF Then
    RS1.Close: Set RS1 = Nothing
    fg_descarga
    MsgBox "No existe información...", vbExclamation + vbOKOnly, MsgTitulo
    Exit Function
End If
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
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcFontSize, 1, 1, 3, 2) = 8: .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 2) = False
    .TableCell(tcText, 1, 1) = "Bodega": .TableCell(tcText, 1, 2) = Left(M_TomInv.Combo1(0).text, 50)
    .TableCell(tcText, 2, 1) = "Toma de Inventario": .TableCell(tcText, 2, 2) = M_TomInv.Date1(0).text
    .TableCell(tcText, 3, 1) = "Familia de Producto": .TableCell(tcText, 3, 2) = IIf(I_TomInv.optTIPPRO(0).Value = True, Left(I_TomInv.Combo1(0).text, 50), "Todas")
    Print #1, .TableCell(tcText, 1, 1) & "|"; .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|"; .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|"; .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 6: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taLeftTop
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
    .TableCell(tcCols) = 6: .TableCell(tcRows) = 60000
    .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1500: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
    i = 1: codtip = 0
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          If RS1!pro_codtip <> codtip And opfampro Then
             If codtip > 0 Then i = i + 1
             .TableCell(tcColSpan, i, 1) = 5
             .TableCell(tcText, i, 1) = fg_BuscaenArbol(RS1!pro_codtip, "a_tipopro", "tip_codigo")
             .TableCell(tcFontBold, i, 1) = True
             Print #1, .TableCell(tcText, i, 1)
             i = i + 1
             codtip = RS1!pro_codtip
          End If
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
fg_descarga
Exit Function
Error_Toma2:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Toma3(Sql As String, opfampro As Boolean)
Dim i As Long, codtip As Long, sumCuenta As Double, sumTipo As Double, sumCuentaLimDes As Double, sumCuentaAlimen As Double
On Local Error GoTo Error_Toma3
fg_carga ""
MsgTitulo = "Toma de Inventario"

If RS3.State = 1 Then RS3.Close
RS3.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS3.Open Sql, vg_db, adOpenStatic
If RS3.EOF Then
    RS3.Close: Set RS3 = Nothing
    fg_descarga
    MsgBox "No existe información...", vbExclamation + vbOKOnly, MsgTitulo
    Exit Function
End If
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
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcFontSize, 1, 1, 3, 2) = 8: .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 2) = False
    .TableCell(tcText, 1, 1) = "Bodega": .TableCell(tcText, 1, 2) = Left(M_TomInv.Combo1(0).text, 50)
    .TableCell(tcText, 2, 1) = "Toma de Inventario": .TableCell(tcText, 2, 2) = M_TomInv.Date1(0).text
    .TableCell(tcText, 3, 1) = "Familia de Producto": .TableCell(tcText, 3, 2) = IIf(I_TomInv.optTIPPRO(0).Value = True, Left(I_TomInv.Combo1(0).text, 50), "Todas")
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
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
    .StartTable
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 60000
    .TableCell(tcColWidth, , 1) = 200: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 4600: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 500: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1500: .TableCell(tcAlign, , 7) = taRightTop
    sumCuentaLimDes = 0: sumCuentaAlimen = 0: codtip = 0: i = 1
    If Not RS3.EOF Then
        RS1.Open "select cta_codigo, cta_nombre from a_ctacontable order by cta_nombre", vg_db, adOpenStatic
        Do While Not RS1.EOF
            RS3.Filter = "pro_ctacon='" & RS1!cta_codigo & "'"
            If RS3.RecordCount > 0 Then RS3.Find "pro_ctacon='" & RS1!cta_codigo & "'", , adSearchForward
            If Not RS3.EOF Then
                sumCuenta = 0
                If i > 1 Then i = i + 1
                .TableCell(tcColSpan, i, 1) = 7
                .TableCell(tcText, i, 1) = RS1!cta_codigo & "  " & RS1!cta_nombre: .TableCell(tcFontUnderline, i, 1) = True
                .TableCell(tcFontBold, i, 1) = True
                Print #1, .TableCell(tcText, i, 1)
                i = i + 2
                sumTipo = 0: codtip = 0: sumCuenta = 0
                Do While Not RS3.EOF
                   If RS3!pro_codtip <> codtip And opfampro Then
                      If codtip > 0 Then
                         i = i + 1
                         .TableCell(tcFontBold, i, 6) = True: .TableCell(tcText, i, 6) = "Total Familia"
                         .TableCell(tcFontBold, i, 7) = True: .TableCell(tcText, i, 7) = Format(sumTipo, fg_Pict(9, vg_DPr))
                         Print #1, "|" & "|" & "|" & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)
                         i = i + 1
                      End If
                      .TableCell(tcColSpan, i, 1) = 7
                      .TableCell(tcText, i, 1) = fg_BuscaenArbol(RS3!pro_codtip, "a_tipopro", "tip_codigo"): .TableCell(tcFontUnderline, i, 1) = False
                      .TableCell(tcFontBold, i, 1) = True
                      Print #1, .TableCell(tcText, i, 1)
                      i = i + 1
                      codtip = RS3!pro_codtip
                      sumTipo = 0
                   End If
                            
                   .TableCell(tcText, i, 2) = RS3!tin_codpro
                   .TableCell(tcText, i, 3) = RS3!pro_nombre
                   .TableCell(tcText, i, 4) = RS3!uni_nomcor
                   .TableCell(tcText, i, 5) = Format(RS3!tin_stofis, fg_Pict(9, vg_DCa))
                   .TableCell(tcText, i, 6) = Format(RS3!tin_propon, fg_Pict(9, 2))
'                   .TableCell(tcText, i, 7) = Format(Format(RS3!tin_stofis, fg_Pict(9, vg_DCa)) * Format(RS3!tin_propon, fg_Pict(9, vg_DCa)), fg_Pict(9, vg_DPr))
                   .TableCell(tcText, i, 7) = Format(Format(RS3!tin_stofis, fg_Pict(9, vg_DCa)) * Format(RS3!tin_propon, fg_Pict(9, 2)), fg_Pict(9, vg_DPr))
                   Print #1, .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                             .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)
'                   sumTipo = sumTipo + Round(Format(RS3!tin_stofis, fg_Pict(9, vg_DCa)) * Format(RS3!tin_propon, fg_Pict(9, vg_DCa)), vg_DPr)
'                   sumCuenta = sumCuenta + Round(Format(RS3!tin_stofis, fg_Pict(9, vg_DCa)) * Format(RS3!tin_propon, fg_Pict(9, vg_DCa)), vg_DPr)
                   sumTipo = sumTipo + Round(Format(RS3!tin_stofis, fg_Pict(9, vg_DCa)) * Format(RS3!tin_propon, fg_Pict(9, 2)), vg_DPr)
                   sumCuenta = sumCuenta + Round(Format(RS3!tin_stofis, fg_Pict(9, vg_DCa)) * Format(RS3!tin_propon, fg_Pict(9, 2)), vg_DPr)
                            
                   RS3.MoveNext: i = i + 1
                Loop
                If opfampro Then
                   i = i + 1
                   .TableCell(tcFontBold, i, 6) = True: .TableCell(tcText, i, 6) = "Total Familia"
                   .TableCell(tcFontBold, i, 7) = True: .TableCell(tcText, i, 7) = Format(sumTipo, fg_Pict(9, vg_DPr))
                   Print #1, "|" & "|" & "|" & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)
                End If
                If RS1!cta_codigo = GetParametro("ctalimdes") Then sumCuentaLimDes = sumCuentaLimDes + sumCuenta
                If RS1!cta_codigo = GetParametro("ctainsumo") Then sumCuentaAlimen = sumCuentaAlimen + sumCuenta
                i = i + 2
                .TableCell(tcFontBold, i, 6) = True: .TableCell(tcText, i, 6) = "Total Cuenta"
                .TableCell(tcFontBold, i, 7) = True: .TableCell(tcText, i, 7) = Format(sumCuenta, fg_Pict(9, vg_DPr))
                Print #1, .TableCell(tcText, i, 6) & .TableCell(tcText, i, 7)
                i = i + 1
                sumCuenta = 0
            End If
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
    End If
    RS3.Close: Set RS3 = Nothing
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable

    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 5
    .TableCell(tcFontBold) = True
    .TableCell(tcColWidth, , 1) = 1900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 7500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcText, 2, 1) = "Alimentos & Bebidas":       .TableCell(tcText, 2, 2) = Format(sumCuentaAlimen, fg_Pict(9, vg_DPr))
    .TableCell(tcText, 3, 1) = "Limpieza  & Desechables":   .TableCell(tcText, 3, 2) = Format(sumCuentaLimDes, fg_Pict(9, vg_DPr))
    .TableCell(tcFontUnderline, 3, 2) = True
'    .TableCell(tcText, 4, 2) = String(15, "_")
    .TableCell(tcText, 5, 1) = "Totales Generales":         .TableCell(tcText, 5, 2) = Format(sumCuentaLimDes + sumCuentaAlimen, fg_Pict(9, vg_DPr))
    .TableCell(tcFontUnderline, 5, 1) = False
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
fg_descarga
Exit Function
Error_Toma3:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation
    Close #1
    Exit Function
End Function

Public Function I_Toma4(Sql As String, opfampro As Boolean)
Dim i As Long, codtip As Long, sumCuenta As Double, sumTipo As Double, sumCuentaLimDes As Double, sumCuentaAlimen As Double
On Local Error GoTo Error_Toma4
fg_carga ""
MsgTitulo = "Toma de Inventario"

If RS3.State = 1 Then RS3.Close
RS3.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS3.Open Sql, vg_db, adOpenStatic
If RS3.EOF Then
    RS3.Close: Set RS3 = Nothing
    fg_descarga
    MsgBox "No existe información...", vbExclamation + vbOKOnly, MsgTitulo
    Exit Function
End If
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
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcFontSize, 1, 1, 3, 2) = 8: .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 2) = False
    .TableCell(tcText, 1, 1) = "Bodega": .TableCell(tcText, 1, 2) = Left(M_TomInv.Combo1(0).text, 50)
    .TableCell(tcText, 2, 1) = "Toma de Inventario": .TableCell(tcText, 2, 2) = M_TomInv.Date1(0).text
    .TableCell(tcText, 3, 1) = "Familia de Producto": .TableCell(tcText, 3, 2) = IIf(I_TomInv.optTIPPRO(0).Value = True, Left(I_TomInv.Combo1(0).text, 50), "Todas")
    Print #1, .TableCell(tcText, 1, 1) & "|"; .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|"; .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 1, 1) & "|"; .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
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
    .StartTable
    .TableCell(tcCols) = 7: .TableCell(tcRows) = 60000
    .TableCell(tcColWidth, , 1) = 200: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 4600: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 500: .TableCell(tcAlign, , 4) = taCenterTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1500: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1500: .TableCell(tcAlign, , 7) = taRightTop
    
    sumCuentaLimDes = 0: sumCuentaAlimen = 0: codtip = 0: i = 1
    If Not RS3.EOF Then
        RS1.Open "select cta_codigo, cta_nombre from a_ctacontable order by cta_nombre", vg_db, adOpenStatic
        Do While Not RS1.EOF
            RS3.Filter = "pro_ctacon='" & RS1!cta_codigo & "'"
            If RS3.RecordCount > 0 Then RS3.Find "pro_ctacon='" & RS1!cta_codigo & "'", , adSearchForward
            If Not RS3.EOF Then
                sumCuenta = 0
                If i > 1 Then i = i + 1
                .TableCell(tcColSpan, i, 1) = 7
                .TableCell(tcText, i, 1) = RS1!cta_codigo & "  " & RS1!cta_nombre: .TableCell(tcFontUnderline, i, 1) = True
                .TableCell(tcFontBold, i, 1) = True
                Print #1, .TableCell(tcText, i, 1)
                i = i + 2
                sumTipo = 0: codtip = 0: sumCuenta = 0
                Do While Not RS3.EOF
                   If RS3!pro_codtip <> codtip And opfampro Then
                      If codtip > 0 Then
                         i = i + 1
                         .TableCell(tcFontBold, i, 6) = True: .TableCell(tcText, i, 6) = "Total Familia"
                         .TableCell(tcFontBold, i, 7) = True: .TableCell(tcText, i, 7) = Format(sumTipo, fg_Pict(9, vg_DPr))
                         Print #1, "|" & "|" & "|" & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)
                         i = i + 1
                      End If
                      .TableCell(tcColSpan, i, 1) = 7
                      .TableCell(tcText, i, 1) = fg_BuscaenArbol(RS3!pro_codtip, "a_tipopro", "tip_codigo"): .TableCell(tcFontUnderline, i, 1) = False
                      .TableCell(tcFontBold, i, 1) = True
                      Print #1, .TableCell(tcText, i, 1)
                      i = i + 1
                      codtip = RS3!pro_codtip
                      sumTipo = 0
                   End If
                            
                   .TableCell(tcText, i, 2) = RS3!tin_codpro
                   .TableCell(tcText, i, 3) = RS3!pro_nombre
                   .TableCell(tcText, i, 4) = RS3!uni_nomcor
                   .TableCell(tcText, i, 5) = Format(RS3!tin_stosis, fg_Pict(9, vg_DCa))
                   .TableCell(tcText, i, 6) = Format(RS3!tin_propon, fg_Pict(9, 2))
                   .TableCell(tcText, i, 7) = Format(Format(RS3!tin_stosis, fg_Pict(9, vg_DCa)) * Format(RS3!tin_propon, fg_Pict(9, 2)), fg_Pict(9, vg_DPr))
                   Print #1, .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                             .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)
                   sumTipo = sumTipo + Round(Format(RS3!tin_stosis, fg_Pict(9, vg_DCa)) * Format(RS3!tin_propon, fg_Pict(9, 2)), vg_DPr)
                   sumCuenta = sumCuenta + Round(Format(RS3!tin_stosis, fg_Pict(9, vg_DCa)) * Format(RS3!tin_propon, fg_Pict(9, 2)), vg_DPr)
                   RS3.MoveNext: i = i + 1
                Loop
                If opfampro Then
                   i = i + 1
                   .TableCell(tcFontBold, i, 6) = True: .TableCell(tcText, i, 6) = "Total Familia"
                   .TableCell(tcFontBold, i, 7) = True: .TableCell(tcText, i, 7) = Format(sumTipo, fg_Pict(9, vg_DPr))
                   Print #1, "|" & "|" & "|" & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7)
                End If
                If RS1!cta_codigo = GetParametro("ctalimdes") Then sumCuentaLimDes = sumCuentaLimDes + sumCuenta
                If RS1!cta_codigo = GetParametro("ctainsumo") Then sumCuentaAlimen = sumCuentaAlimen + sumCuenta
                i = i + 2
                .TableCell(tcFontBold, i, 6) = True: .TableCell(tcText, i, 6) = "Total Cuenta"
                .TableCell(tcFontBold, i, 7) = True: .TableCell(tcText, i, 7) = Format(sumCuenta, fg_Pict(9, vg_DPr))
                Print #1, .TableCell(tcText, i, 6) & .TableCell(tcText, i, 7)
                i = i + 1
                sumCuenta = 0
            End If
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
    End If
    RS3.Close: Set RS3 = Nothing
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    
    .StartTable
    .TableCell(tcCols) = 3: .TableCell(tcRows) = 5
    .TableCell(tcFontBold) = True
    .TableCell(tcColWidth, , 1) = 1900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 7500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcText, 2, 1) = "Alimentos & Bebidas":       .TableCell(tcText, 2, 2) = Format(sumCuentaAlimen, fg_Pict(9, vg_DPr))
    .TableCell(tcText, 3, 1) = "Limpieza  & Desechables":   .TableCell(tcText, 3, 2) = Format(sumCuentaLimDes, fg_Pict(9, vg_DPr))
    .TableCell(tcFontUnderline, 3, 2) = True
    .TableCell(tcText, 5, 1) = "Totales Generales":         .TableCell(tcText, 5, 2) = Format(sumCuentaLimDes + sumCuentaAlimen, fg_Pict(9, vg_DPr))
    .TableCell(tcFontUnderline, 5, 1) = False
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
fg_descarga
Exit Function
Error_Toma4:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_Toma5(Sql As String, opfampro As Boolean)
Dim i As Long, codtip As Long, ctacon As String, totfis As Double, totsis As Double, totdif As Double, grlfis As Double, grlsis As Double, grldif As Double
On Local Error GoTo Error_Toma5
fg_carga ""
MsgTitulo = "Toma de Inventario"

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open Sql, vg_db, adOpenStatic
If RS1.EOF Then
    RS1.Close: Set RS1 = Nothing
    fg_descarga
    MsgBox "No existe información...", vbExclamation + vbOKOnly, MsgTitulo
    Exit Function
End If
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
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 8500: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcFontSize, 1, 1, 3, 2) = 8: .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 2) = False
    .TableCell(tcText, 1, 1) = "Bodega": .TableCell(tcText, 1, 2) = Left(M_TomInv.Combo1(0).text, 50)
    .TableCell(tcText, 2, 1) = "Toma de Inventario": .TableCell(tcText, 2, 2) = M_TomInv.Date1(0).text
    .TableCell(tcText, 3, 1) = "Familia de Producto": .TableCell(tcText, 3, 2) = IIf(I_TomInv.optTIPPRO(0).Value = True, Left(I_TomInv.Combo1(0).text, 50), "Todas")
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 10: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 800: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2850: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 400: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1000: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcColWidth, , 10) = 1200: .TableCell(tcAlign, , 10) = taRightTop
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
    .TableCell(tcText, 1, 10) = "Total Dif."
    Print #1, .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8) & "|" & .TableCell(tcText, 1, 9) & "|" & .TableCell(tcText, 1, 10)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 10: .TableCell(tcRows) = 60000
    .TableCell(tcColWidth, , 1) = 800: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2850: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 400: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcColWidth, , 9) = 1000: .TableCell(tcAlign, , 9) = taRightTop
    .TableCell(tcColWidth, , 10) = 1200: .TableCell(tcAlign, , 10) = taRightTop
    i = 1: codtip = 0: totfis = 0: totsis = 0: totdif = 0: grlfis = 0: grlsis = 0: grldif = 0: ctacon = ""
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          If RS1!pro_ctacon <> ctacon Then
             If Trim(ctacon) <> "" Then
                .TableCell(tcText, i, 2) = "Total Cta. Contable ": .TableCell(tcFontBold, i, 2) = True
                .TableCell(tcText, i, 6) = Format(totfis, fg_Pict(9, vg_DPr)): .TableCell(tcFontBold, i, 6) = True
                .TableCell(tcText, i, 8) = Format(totsis, fg_Pict(9, vg_DPr)): .TableCell(tcFontBold, i, 8) = True
                .TableCell(tcText, i, 10) = Format(totdif, fg_Pict(9, vg_DPr)): .TableCell(tcFontBold, i, 10) = True
                totfis = 0: totsis = 0: totdif = 0
                i = i + 2
             End If
             .TableCell(tcColSpan, i, 1) = 9
             .TableCell(tcText, i, 1) = RS1!pro_ctacon & Space(10) & Trim(RS1!cta_nombre)
             .TableCell(tcFontBold, i, 1) = True
             Print #1, .TableCell(tcText, i, 1)
             i = i + 1
             ctacon = RS1!pro_ctacon
          End If
          If RS1!pro_codtip <> codtip And opfampro Then
             If codtip > 0 Then i = i + 1
             .TableCell(tcColSpan, i, 1) = 9
             .TableCell(tcText, i, 1) = fg_BuscaenArbol(RS1!pro_codtip, "a_tipopro", "tip_codigo")
             .TableCell(tcFontBold, i, 1) = True
             Print #1, .TableCell(tcText, i, 1)
             i = i + 1
             codtip = RS1!pro_codtip
          End If
            
          .TableCell(tcText, i, 1) = RS1!tin_codpro
          .TableCell(tcText, i, 2) = RS1!pro_nombre
          .TableCell(tcText, i, 3) = RS1!uni_nomcor
          .TableCell(tcText, i, 4) = Format(RS1!tin_propon, fg_Pict(9, 2))
          .TableCell(tcText, i, 5) = Format(RS1!tin_stofis, fg_Pict(9, vg_DCa))
          .TableCell(tcText, i, 6) = Format(Format(RS1!tin_stofis, fg_Pict(9, vg_DCa)) * Format(RS1!tin_propon, fg_Pict(9, 2)), fg_Pict(9, vg_DPr))
          .TableCell(tcText, i, 7) = Format(RS1!tin_stosis, fg_Pict(9, vg_DCa))
          .TableCell(tcText, i, 8) = Format(Format(RS1!tin_stosis, fg_Pict(9, vg_DCa)) * Format(RS1!tin_propon, fg_Pict(9, 2)), fg_Pict(9, vg_DPr))
          .TableCell(tcText, i, 9) = Format(RS1!tin_stofis - RS1!tin_stosis, fg_Pict(9, vg_DCa))
          .TableCell(tcText, i, 10) = Format(((RS1!tin_stofis - RS1!tin_stosis) * RS1!tin_propon), fg_Pict(9, vg_DCa))
          Print #1, .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                    .TableCell(tcText, i, 5) & "|" & .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8) & "|" & .TableCell(tcText, i, 9) & "|" & .TableCell(tcText, i, 10)
          totfis = Round((totfis + (RS1!tin_stofis * RS1!tin_propon)), 0)
          totsis = Round((totsis + (RS1!tin_stosis * RS1!tin_propon)), 0)
          totdif = Round(totdif + ((RS1!tin_stofis * RS1!tin_propon) - (RS1!tin_stosis * RS1!tin_propon)), 0)
          grlfis = Round(grlfis + (RS1!tin_stofis * RS1!tin_propon), 0)
          grlsis = Round(grlsis + (RS1!tin_stosis * RS1!tin_propon), 0)
          grldif = Round(grldif + ((RS1!tin_stofis * RS1!tin_propon) - (RS1!tin_stosis * RS1!tin_propon)), 0)
          RS1.MoveNext: i = i + 1
       Loop
       .TableCell(tcText, i, 2) = "Total Cta. Contable ": .TableCell(tcFontBold, i, 2) = True
       .TableCell(tcText, i, 6) = Format(totfis, fg_Pict(9, vg_DPr)): .TableCell(tcFontBold, i, 6) = True
       .TableCell(tcText, i, 8) = Format(totsis, fg_Pict(9, vg_DPr)): .TableCell(tcFontBold, i, 8) = True
       .TableCell(tcText, i, 10) = Format(totdif, fg_Pict(9, vg_DPr)): .TableCell(tcFontBold, i, 10) = True
       totfis = 0: totsis = 0: totdif = 0
       i = i + 2
       .TableCell(tcText, i, 2) = "Total General ": .TableCell(tcFontBold, i, 2) = True
       .TableCell(tcText, i, 6) = Format(grlfis, fg_Pict(9, vg_DPr)): .TableCell(tcFontBold, i, 6) = True
       .TableCell(tcText, i, 8) = Format(grlsis, fg_Pict(9, vg_DPr)): .TableCell(tcFontBold, i, 8) = True
       .TableCell(tcText, i, 10) = Format(grldif, fg_Pict(9, vg_DPr)): .TableCell(tcFontBold, i, 10) = True
       totfis = 0: totsis = 0: grldif = 0
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
fg_descarga
Exit Function
Error_Toma5:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_SalBodega(cencos As String, codreg As String, codser As String, fecini As String, fecter As String, opspag As Boolean)

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset

Dim i As Long, auxreg As Long, auxser As Long, ndia As Long, nomcen As String
Dim cantfija As Double, cantprodxdia As Double, aAp As String, auxfec As Variant

On Local Error GoTo Error_SalBod

fg_carga ""
MsgTitulo = "Salida de Bodega"
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
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = "Formato de Requisición"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Contrato"
     
     If RS1.State = 1 Then RS1.Close
     RS1.CursorLocation = adUseClient
     vg_db.CursorLocation = adUseClient
    
     RS1.Open RutinaLectura.Cliente(1, cencos, ""), vg_db, adOpenStatic
     If Not RS1.EOF Then .TableCell(tcFontBold, 1, 2) = True: .TableCell(tcText, 1, 2) = RS1!cli_codigo & " " & Trim(RS1!cli_nombre): nomcen = Trim(RS1!cli_nombre)
     RS1.Close: Set RS1 = Nothing
     
    .TableCell(tcText, 2, 1) = "Rango Fecha"
    .TableCell(tcText, 2, 2) = fecini & " - " & fecter
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13)
    
    '------- Rutina validar producto vigente
    ValidarProductoVigente
    '------- Creo tabla temporal y chequeo si existe antes
    
'    aAp = Trim(vg_NUsr) & "_tmp_SalBod1"
'    fg_CheckTmp aAp
'    RS1.Open "SELECT DISTINCT b.reg_codigo, b.reg_nombre, a.ser_codigo, a.ser_nombre, c.min_fecmin INTO " & aAp & " " & _
'             "FROM a_servicio a, a_regimen b, b_minuta c, b_minutadet d " & _
'             "WHERE  c.min_codigo = d.mid_codigo " & _
'             "AND    c.min_codreg = b.reg_codigo " & _
'             "AND    c.min_codser = a.ser_codigo " & _
'             "AND    c.min_cencos = '" & cencos & "' " & _
'             "AND    c.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
'             "AND    c.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
'             "AND    c.min_fecmin >= " & Format(fecini, "yyyymmdd") & " " & _
'             "AND    c.min_fecmin <= " & Format(fecter, "yyyymmdd") & " " & _
'             "AND    d.mid_tipmin = '2'", vg_db, adOpenStatic
'    Set RS1 = Nothing
'
'    vg_db.Execute "INSERT INTO " & aAp & " SELECT DISTINCT b.reg_codigo, b.reg_nombre, a.ser_codigo, a.ser_nombre, c.mfd_fecha AS min_fecmin " & _
'                  "FROM a_servicio a, a_regimen b, b_minutafijadia c " & _
'                  "WHERE c.mfd_codreg = b.reg_codigo " & _
'                  "AND   c.mfd_codser = a.ser_codigo " & _
'                  "AND   c.mfd_cencos = '" & cencos & "' " & _
'                  "AND   c.mfd_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
'                  "AND   c.mfd_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
'                  "AND   c.mfd_fecha >= " & Format(fecini, "yyyymmdd") & " " & _
'                  "AND   c.mfd_fecha <= " & Format(fecter, "yyyymmdd") & " " & _
'                  "AND   c.mfd_tipmin = '2'"
'    Set RS1 = Nothing
'
'    RS1.Open "SELECT DISTINCT reg_codigo, reg_nombre, ser_codigo, ser_nombre, min_fecmin FROM " & aAp & " ORDER BY reg_codigo, ser_codigo, min_fecmin", vg_db, adOpenStatic
     If RS1.State = 1 Then RS1.Close
     RS1.CursorLocation = adUseClient
     vg_db.CursorLocation = adUseClient
     
     Set RS1 = vg_db.Execute("sgp_Sel_FormatoRequisicion '" & codser & "','" & codreg & "', '" & cencos & "'," & Format(fecini, "yyyymmdd") & "," & Format(fecter, "yyyymmdd") & "")
     If RS1.EOF Then
     
        fg_descarga
        RS1.Close
        Set RS1 = Nothing
        Close #1
        
        MsgBox "No existe información con los datos seleccionados", vbExclamation + vbOKOnly, MsgTitulo
        
        Exit Function
    
     End If
    
     Do While Not RS1.EOF
        
        ndia = fg_NumDia(Trim(Left(fg_Fecha_Dia(Trim(Str(RS1!min_fecmin)), 2), Len(fg_Fecha_Dia(Trim(Str(RS1!min_fecmin)), 2)) - 2)))
        
        'Creo tabla temporal y chequeo si existe antes
        aAp = Trim(vg_NUsr) & "_tmp_SalBod"
        fg_CheckTmp aAp
        
        If RS2.State = 1 Then RS2.Close
        RS2.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS2.Open "SELECT red.red_codpro, SUM(mid.mid_numrac*(red.red_canpro/rec.rec_basrac)) AS cantidad into " & aAp & " " & _
                 "FROM b_minuta mi, b_minutadet mid, b_receta rec, b_recetadet red " & _
                 "WHERE rec.rec_codigo = mid.mid_codrec " & _
                 "AND   rec.rec_codigo = red.red_codigo " & _
                 "AND   red.red_tiprec = mid.mid_tiprec AND ((red.red_tiprec<>0 AND red.red_cencos='" & MuestraCasino(1) & "') OR (red.red_tiprec=0 AND red.red_cencos='0')) " & _
                 "AND   mi.min_codigo  = mid.mid_codigo " & _
                 "AND   mi.min_fecmin  = " & RS1!min_fecmin & " " & _
                 "AND   mid.mid_tipmin = '2' " & _
                 "AND   mi.min_cencos  = '" & cencos & "' " & _
                 "AND   mi.min_codreg  = " & RS1!reg_codigo & " " & _
                 "AND   mi.min_codser  = " & RS1!ser_codigo & " " & _
                 "GROUP BY red.red_codpro", vg_db, adOpenStatic
        Set RS2 = Nothing
        
        If RS2.State = 1 Then RS2.Close
        RS2.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS2.Open "SELECT DISTINCT pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, " & _
                 "       SUM(aux.cantidad/pro.pro_facing) AS cantprodxdia, SUM(aux.cantidad) AS cantidad2 " & _
                 "FROM   b_productos pro, b_ingrediente ing, b_productosing pri, a_unidad uni, " & _
                 "       a_unidadmed unm, " & aAp & " aux, b_contlistpreing a " & _
                 "WHERE  pri.pri_coding = aux.red_codpro " & _
                 "AND    ing.ing_codigo = pri.pri_coding AND ing.ing_codigo = a.cpi_coding AND a.cpi_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND    pro.pro_codigo = a.cpi_codped " & _
                 "AND   (pro.pro_fecven > " & Format(Date, "yyyymmdd") & " OR pro.pro_fecven <= 0 OR (pro.pro_codigo IN (SELECT bod.bod_codpro FROM b_bodegas bod WHERE bod.bod_codbod = " & vg_codbod & " AND bod.bod_canmer > 0))) " & _
                 "AND    pri.pri_codpro = a.cpi_codped " & _
                 "AND    ing.ing_unimed = unm.unm_codigo " & _
                 "AND    pro.pro_coduni = uni.uni_codigo " & _
                 "AND    pro.pro_ctrsto = 1 " & _
                 "GROUP BY pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor " & _
                 "ORDER BY pro.pro_codigo", vg_db, adOpenStatic
                     
        If RS3.State = 1 Then RS3.Close
        RS3.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS3.Open "SELECT b.pro_codigo, b.pro_nombre, c.uni_nomcor, a.mfd_canpro " & _
                 "FROM b_minutafijadia a, b_productos b, a_unidad c " & _
                 "WHERE a.mfd_codpro = b.pro_codigo AND b.pro_coduni=c.uni_codigo " & _
                 "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0 OR (b.pro_codigo IN (SELECT bod_codpro FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " AND bod_canmer > 0))) " & _
                 "AND   a.mfd_cencos = '" & cencos & "' " & _
                 "AND   a.mfd_codreg = " & RS1!reg_codigo & " " & _
                 "AND   a.mfd_codser = " & RS1!ser_codigo & " " & _
                 "AND   a.mfd_fecha  = " & RS1!min_fecmin & " AND a.mfd_tipmin='2'", vg_db, adOpenStatic
        
        auxser = 0
        auxreg = 0
        
        If Not RS2.EOF Or Not RS3.EOF Then
           
           If RS1!ser_codigo <> auxser Or RS1!min_fecmin <> auxfec Or RS1!reg_codigo <> auxreg Then
              
              If auxfec > 0 And opspag Then
                 
                 .NewPage
                 .text = Chr(13): .text = Chr(13)
                 .StartTable
                 .TableCell(tcCols) = 2: .TableCell(tcRows) = 2
                 .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
                 .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
                 .TableCell(tcFontBold, , 1) = True
                 .TableCell(tcText, 1, 1) = "Contrato"
                 .TableCell(tcText, 1, 2) = cencos & " " & nomcen
                 .TableCell(tcText, 2, 1) = "Rango Facha"
                 .TableCell(tcText, 2, 2) = fecini & " - " & fecter
                 Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
                 Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
                 .TableBorder = tbNone
                 .EndTable
              
              End If
              
              .StartTable
              .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
              .TableCell(tcFontBold, , 1, , 1) = True: .TableCell(tcFontBold, , 3, , 3) = True
              .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taLeftTop
              .TableCell(tcBackColor, 1) = vbYellow:: .TableCell(tcRowHeight, 2) = 200
              .TableCell(tcText, 1, 1) = "Regimen : " & RS1!reg_codigo & " " & Trim(RS1!reg_nombre) & "- Servicio : " & RS1!ser_codigo & " " & Trim(RS1!ser_nombre) & " - " & "Fecha : " & fg_Ctod1(RS1!min_fecmin)
              Print #1, .TableCell(tcText, 1, 1)
              .TableBorder = tbNone
              .EndTable
              auxreg = RS1!reg_codigo
              auxser = RS1!ser_codigo
              auxfec = RS1!min_fecmin
           
           End If
           
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
           .TableCell(tcText, 2, 5) = "Cant.Real."
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
                 .TableCell(tcText, i, 6) = RS2!uni_nomcor
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
                 .TableCell(tcText, i, 4) = Format(TipoDato(RS3!mfd_canpro, 0), fg_Pict(9, vg_DCa))
                 .TableCell(tcText, i, 5) = "_________"
                 .TableCell(tcText, i, 6) = RS3!uni_nomcor
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
    Close #1
    .EndDoc

End With

Preview.Show 1
fg_descarga
Exit Function

Error_SalBod:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function

End Function

Public Function I_SalBodega2(casnom As String, Regimen As String, codser As Long, fecini As String, fecter As String)
Dim i As Long, Nfecini As Long, Nfecter As Long, Casino As String, nomcas As String, codreg As Long, nomreg As String
Dim sqlSER As String, cantfija As Double, cantprodxdia As Double, Fecha As String
On Local Error GoTo Error_Requisicion
fg_carga ""
Casino = Trim(Mid(casnom, 1, InStr(1, casnom, "|") - 1))
nomcas = Trim(Mid(casnom, InStr(1, casnom, "|") + 1, Len(casnom)))
codreg = Val(Mid(Regimen, 1, InStr(1, Regimen, "|") - 1))
nomreg = Trim(Mid(Regimen, InStr(1, Regimen, "|") + 1, Len(Regimen)))
sqlSER = IIf(codser = 0, " ", " where ser_codigo=" & codser & " ")
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
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 12: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = "Formato de Requisición"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Contrato"
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
    Fecha = fecini
     Do While CDate(Fecha) <= CDate(fecter)
        Nfecini = Val(Format(Fecha, "yyyy") & Right("0" & Format(Fecha, "mm"), 2) & Right("0" & Format(Fecha, "dd"), 2))
        'Dim ndia As Long
        'ndia = fg_NumDia(Trim(Left(fg_Fecha_Dia(Trim(Str(Fecha)), 2), Len(fg_Fecha_Dia(Trim(Str(Fecha)), 2)) - 2)))
        ndia = fg_NumDia(Trim(Left(fg_Fecha_Dia(Trim(Str(Nfecini)), 2), Len(fg_Fecha_Dia(Trim(Str(Nfecini)), 2)) - 2)))
        RS1.Open "SELECT ser_codigo, ser_nombre FROM a_servicio " & sqlSER & " ORDER BY ser_codigo", vg_db, adOpenStatic
        Do While Not RS1.EOF
            RS2.Open "SELECT DISTINCT mi.min_fecmin, mi.min_codser, pro.pro_codigo, pro.pro_nombre, uni.uni_nombre, " & _
                     "(red.red_canpro * mid.mid_numrac) AS cantprodxdia " & _
                     "FROM  b_productos pro, b_minuta mi, b_minutadet mid, b_receta rec, b_recetadet red, a_unidad uni " & _
                     "WHERE pro.pro_codigo=red.red_codpro AND rec.rec_codigo=mid.mid_codrec AND mid.mid_tiprec=red.red_tiprec " & _
                     "AND   rec.rec_codigo=red.red_codigo AND ((red.red_tiprec<>0 AND red.red_cencos='" & MuestraCasino(1) & "') OR (red.red_tiprec=0 AND red.red_cencos='0')) AND mi.min_codigo=mid.mid_codigo " & _
                     "AND   pro.pro_coduni=uni.uni_codigo AND mid.mid_tipmin='2' AND mi.min_fecmin=" & Fecha & " " & _
                     "AND   mi.min_cencos='" & Trim(Casino) & "' AND mi.min_codreg=" & codreg & " " & _
                     "AND   mi.min_codser=" & RS1!ser_codigo, vg_db, adOpenStatic
            
            RS3.Open "SELECT b.pro_codigo, b.pro_nombre, c.uni_nombre, a.mfd_canpro " & _
                     "FROM b_minutafijadia a, b_productos b, a_unidad c " & _
                     "WHERE a.mfd_codpro=b.pro_codigo AND b.pro_coduni=c.uni_codigo " & _
                     "AND   a.mfd_cencos='" & Trim(Casino) & "' " & _
                     "AND   a.mfd_codreg=" & codreg & " " & _
                     "AND   a.mfd_codser=" & RS1!ser_codigo & " " & _
                     "AND   a.mfd_fecha=" & Fecha & " AND a.mfd_tipmin='2'", vg_db, adOpenStatic
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
                .TableCell(tcText, 2, 4) = CDate(Fecha)
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
                .TableCell(tcText, 2, 5) = "Cant.Realizada"
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
                        .TableCell(tcText, i, 4) = Format(TipoDato(RS3!mfd_canpro, 0), fg_Pict(9, vg_DCa))
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
        Fecha = Str(CDate(Fecha) + 1)
    Loop
    .EndDoc
End With
Preview.Show 1
fg_descarga
Exit Function
Error_Requisicion:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_StockxFecha(Formu As Form)
Dim i As Long, sumCuenta As Double, sumTipo As Double, sqlTP As String, sqlBO As String, sqlCU As String, v_codbod As Long, aAp As String
Dim ctacon As String, ctacon2 As String, codtip As Long, codtip2 As Long
On Local Error GoTo Error_StockFecha
fg_carga ""
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
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10800: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Informe de Stock"
    Print #1, .TableCell(tcText, 1, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 2000: .TableCell(tcAlign, , 1) = taLeftMiddle
    .TableCell(tcColWidth, , 2) = 8800: .TableCell(tcAlign, , 2) = taLeftMiddle
    .TableCell(tcFontSize, 1, 1, 1, 1) = 8: .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Bodega": .TableCell(tcText, 1, 2) = Left(I_Stock.Combo1(0).text, 50)
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    .TableBorder = tbNone
    .EndTable
    v_codbod = fg_codigocbo(Formu.Combo1, 0, 10, 0)
    sqlTP = ""
    If Formu.optTIPPRO(0).Value = True Then sqlTP = "and pro.pro_codtip=" & Val(fg_codigocbo(Formu.Combo1, 1, 10, 0)) & " "
    sqlCU = ""
    If Formu.optCUENTA(0).Value = True Then sqlCU = "and pro.pro_ctacon='" & Trim(Mid(Trim(Formu.Combo1(2).List(Formu.Combo1(2).ListIndex)), Len(Trim(Formu.Combo1(2).List(Formu.Combo1(2).ListIndex))) - 10, 10)) & "' "
    sqlBO = "AND bod.bod_codbod=" & v_codbod & " "
    '-------> Insert tabla productospmpdia
    If vg_tipbase = "1" Then
       aAp = Trim(vg_NUsr) & "_tmp_ProductoPMPStockxFecha"
       fg_CheckTmp aAp
       vg_db.Execute "SELECT TOP 1 ppd_cencos, ppd_codpro, 0 AS ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                     "INTO " & aAp & " " & _
                     "FROM b_productospmpdia " & _
                     "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                     "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(Date, "yyyymmdd") & " " & _
                     "AND   ppd_propon > 0 " & _
                     "GROUP BY ppd_cencos, ppd_codpro ORDER BY Max(ppd_fecdia) DESC"
       vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
       vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon"
       vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
    
       RS1.Open "SELECT cta.cta_codigo, cta.cta_nombre, tip.tip_codigo, tip.tip_nombre, " & _
                "       pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, a.ppd_propon, SUM(bod.bod_canmer) AS canmer " & _
                "FROM   a_ctacontable cta, b_bodegas bod, a_unidad uni, b_productos pro, a_tipopro tip, " & aAp & " a " & _
                "WHERE  cta.cta_codigo = pro.pro_ctacon " & _
                "AND    bod.bod_codpro = pro.pro_codigo " & _
                "AND    pro.pro_codigo = a.ppd_codpro " & _
                "AND    a.ppd_cencos   = '" & MuestraCasino(1) & "' " & _
                "AND    tip.tip_codigo = pro.pro_codtip " & _
                "AND    uni.uni_codigo = pro.pro_coduni " & _
                sqlCU & sqlTP & sqlBO & _
                "GROUP  BY cta.cta_codigo, cta.cta_nombre, tip.tip_codigo, tip.tip_nombre, " & _
                "       pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, a.ppd_propon Having round(Sum(bod.bod_canmer)," & vg_DCa & ")>0 " & _
                "ORDER  BY cta.cta_codigo, tip.tip_codigo, pro.pro_nombre", vg_db, adOpenStatic
    Else
'       RS1.Open "SELECT cta.cta_codigo, cta.cta_nombre, tip.tip_codigo, tip.tip_nombre, " & _
'                "       pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, (SELECT TOP 1 ppd_propon FROM b_productospmpdia WHERE ppd_codpro = pro.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND  ppd_fecdia <= " & Format(Date, "yyyymmdd") & " AND ppd_saldo <> 0 ORDER BY ppd_fecdia DESC) AS ppd_propon, SUM(bod.bod_canmer) AS canmer " & _
'                "FROM   a_ctacontable cta, b_bodegas bod, a_unidad uni, b_productos pro, a_tipopro tip " & _
'                "WHERE  cta.cta_codigo = pro.pro_ctacon " & _
'                "AND    bod.bod_codpro = pro.pro_codigo " & _
'                "AND    tip.tip_codigo = pro.pro_codtip " & _
'                "AND    uni.uni_codigo = pro.pro_coduni " & _
'                "AND    (pro.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') or pro.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')) " & _
'                sqlCU & sqlTP & sqlBO & _
'                "GROUP  BY cta.cta_codigo, cta.cta_nombre, tip.tip_codigo, tip.tip_nombre, " & _
'                "       pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor Having round(Sum(bod.bod_canmer), " & vg_DCa & ")>0 " & _
'                "ORDER  BY cta.cta_codigo, tip.tip_codigo, pro.pro_nombre", vg_db, adOpenStatic
       
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS1.Open "SELECT cta.cta_codigo, cta.cta_nombre, tip.tip_codigo, tip.tip_nombre, " & _
                "       pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, (SELECT TOP 1 ppd_propon FROM b_productospmpdia WHERE ppd_codpro = pro.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND  ppd_fecdia <= " & Format(Date, "yyyymmdd") & " ORDER BY ppd_fecdia DESC) AS ppd_propon, SUM(bod.bod_canmer) AS canmer " & _
                "FROM   a_ctacontable cta, b_bodegas bod, a_unidad uni, b_productos pro, a_tipopro tip " & _
                "WHERE  cta.cta_codigo = pro.pro_ctacon " & _
                "AND    bod.bod_codpro = pro.pro_codigo " & _
                "AND    tip.tip_codigo = pro.pro_codtip " & _
                "AND    uni.uni_codigo = pro.pro_coduni " & _
                "AND    (pro.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') or pro.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')) " & _
                sqlCU & sqlTP & sqlBO & _
                "GROUP  BY cta.cta_codigo, cta.cta_nombre, tip.tip_codigo, tip.tip_nombre, " & _
                "       pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor Having round(Sum(bod.bod_canmer), " & vg_DCa & ")>0 " & _
                "ORDER  BY cta.cta_codigo, tip.tip_codigo, pro.pro_nombre", vg_db, adOpenStatic
    End If
    ctacon = ""
    If Not RS1.EOF Then
        .StartTable
        .TableCell(tcCols) = 7: .TableCell(tcRows) = 1
        .TableCell(tcColWidth, , 1) = 200: .TableCell(tcAlign, , 1) = taLeftTop
        .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taLeftTop
        .TableCell(tcColWidth, , 3) = 5000: .TableCell(tcAlign, , 3) = taLeftTop
        .TableCell(tcColWidth, , 4) = 500: .TableCell(tcAlign, , 4) = taCenterTop
        .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
        .TableCell(tcColWidth, , 6) = 1200: .TableCell(tcAlign, , 6) = taRightTop
        .TableCell(tcColWidth, , 7) = 1400: .TableCell(tcAlign, , 7) = taRightTop
        .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 200
        .TableCell(tcText, 1, 2) = "Código"
        .TableCell(tcText, 1, 3) = "Descripción"
        .TableCell(tcText, 1, 4) = "Unidad"
        .TableCell(tcText, 1, 5) = "Stock"
        .TableCell(tcText, 1, 6) = "Precio"
        .TableCell(tcText, 1, 7) = "Total"
        
        Print #1, .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & .TableCell(tcText, 1, 5) & "|" & .TableCell(tcText, 1, 6)
        .TableBorder = tbNone
        .EndTable
    End If
    Do While Not RS1.EOF
        If ctacon <> RS1!cta_codigo Then
            .StartTable
            .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
            .TableCell(tcColWidth, , 1) = 10800: .TableCell(tcAlign, , 1) = taLeftTop
            .TableCell(tcFontBold, 2, 1) = True: .TableCell(tcFontUnderline, 2, 1) = True
            .TableCell(tcText, 2, 1) = RS1!cta_codigo & "  " & RS1!cta_nombre
            Print #1, .TableCell(tcText, 2, 1)
            .TableBorder = tbNone
            .EndTable
            ctacon = RS1!cta_codigo
            sumCuenta = 0
        End If
        If codtip <> RS1!tip_codigo Then
            .StartTable
            .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
            .TableCell(tcFontBold, 2, 1) = True
            .TableCell(tcColWidth, , 1) = 10800: .TableCell(tcAlign, , 1) = taLeftTop
            .TableCell(tcText, 2, 1) = RS1!tip_codigo & "  " & RS1!tip_nombre
            Print #1, .TableCell(tcText, 2, 1)
            .TableBorder = tbNone
            .EndTable
            sumTipo = 0
            codtip = RS1!tip_codigo
            i = 1
            .StartTable
            .TableCell(tcCols) = 7: .TableCell(tcRows) = 1000
            .TableCell(tcColWidth, , 1) = 200: .TableCell(tcAlign, , 1) = taLeftTop
            .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taLeftTop
            .TableCell(tcColWidth, , 3) = 5000: .TableCell(tcAlign, , 3) = taLeftTop
            .TableCell(tcColWidth, , 4) = 500: .TableCell(tcAlign, , 4) = taCenterTop
            .TableCell(tcColWidth, , 5) = 1000: .TableCell(tcAlign, , 5) = taRightTop
            .TableCell(tcColWidth, , 6) = 1200: .TableCell(tcAlign, , 6) = taRightTop
            .TableCell(tcColWidth, , 7) = 1400: .TableCell(tcAlign, , 7) = taRightTop
        End If
        .TableCell(tcText, i, 2) = RS1!pro_codigo
        .TableCell(tcText, i, 3) = RS1!pro_nombre
        .TableCell(tcText, i, 4) = RS1!uni_nomcor
        .TableCell(tcText, i, 5) = Format(RS1!canmer, fg_Pict(9, vg_DCa))
        .TableCell(tcText, i, 6) = Format(RS1!ppd_propon, fg_Pict(9, vg_DPr))
        .TableCell(tcText, i, 7) = Format(RS1!canmer * IIf(IsNull(RS1!ppd_propon), 0, RS1!ppd_propon), fg_Pict(9, vg_DPr))
        sumTipo = sumTipo + Round(RS1!canmer * IIf(IsNull(RS1!ppd_propon), 0, RS1!ppd_propon), vg_DPr)
        RS1.MoveNext: i = i + 1
        If Not RS1.EOF Then codtip2 = RS1!tip_codigo: ctacon2 = RS1!cta_codigo Else codtip2 = 0: ctacon2 = 0
        If ctacon <> ctacon2 Or codtip <> codtip2 Or RS1.EOF Then
            codtip = 0
            sumCuenta = sumCuenta + sumTipo
            .TableCell(tcFontUnderline, i, 7) = True
            .TableCell(tcText, i, 7) = Space(20)
            .TableCell(tcText, i + 1, 6) = "Total Familia"
            .TableCell(tcText, i + 1, 7) = Format(sumTipo, fg_Pict(9, vg_DPr))
            .TableCell(tcRows) = i + 2
            .TableBorder = tbNone
            .EndTable
        End If
        If ctacon <> ctacon2 Or RS1.EOF Then
            ctacon = ""
            .StartTable
            .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
            .TableCell(tcFontBold) = True: .TableCell(tcAlign) = taRightTop
            .TableCell(tcColWidth, , 1) = 9300: .TableCell(tcColWidth, , 2) = 1500
            .TableCell(tcText, 1, 1) = "TOTAL CUENTA"
            .TableCell(tcText, 1, 2) = Format(sumCuenta, fg_Pict(9, vg_DPr))
            Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
            .TableBorder = tbNone
            .EndTable
            .StartTable
            .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
            .TableCell(tcFontBold) = True
            .TableCell(tcColWidth) = 10800: .TableCell(tcAlign) = taRightTop
            .TableCell(tcText, 1, 1) = String(143, "_")
            Print #1, .TableCell(tcText, 1, 1)
            .TableBorder = tbNone
            .EndTable
        End If
    Loop
    RS1.Close: Set RS1 = Nothing
    .EndDoc
    Close #1
    '-------> Borrar tablas temporales
    If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
End With
Preview.Show 1
fg_descarga
Exit Function
Error_StockFecha:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Sub I_UniMed()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_UniEnv
fg_carga ""
MsgTitulo = "Informe de Unidades de Medida"
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
    .TableCell(tcText, 1, 1) = "Unidades de Medida"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open RutinaLectura.UnidadMedida(2, 0, ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Sub
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
Preview.Show 1
fg_descarga
Exit Sub
Error_UniEnv:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Sub I_GrupoPaciente()
Dim i As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long
On Local Error GoTo Error_GrupoPaciente
fg_carga ""
MsgTitulo = "Informe de Departamento"
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
    .TableCell(tcText, 1, 1) = "Departamento"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    RS1.Open "SELECT * from a_grupopaciente order by grp_nombre", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Close #1: fg_descarga: Exit Sub
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Código"
    .TableCell(tcText, 1, 2) = "Nombre"
    .TableCell(tcText, 1, 3) = "Código(s) de Homologación (*)"
    .TableCell(tcText, 1, 4) = "Estado"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 3000: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taLeftTop
    i = 1
    Do While Not RS1.EOF
        .TableCell(tcText, i, 1) = IIf(IsNull(RS1!grp_codigo), "", RS1!grp_codigo)
        .TableCell(tcText, i, 2) = IIf(IsNull(RS1!grp_nombre), "", Trim(RS1!grp_nombre))
        .TableCell(tcText, i, 3) = IIf(IsNull(RS1!grp_othervalue), "", Trim(RS1!grp_othervalue))
        .TableCell(tcText, i, 4) = IIf(IsNull(RS1!grp_estado), "", IIf(RS1!grp_estado = "0", "Activo", "Bloqueado"))
        Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing: Close #1
    .TableCell(tcRows) = i
    .TableBorder = tbNone
    .EndTable
    .EndDoc
End With
Preview.Show 1
fg_descarga
Exit Sub
Error_GrupoPaciente:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Sub
End Sub

Public Function I_FacturaCli(Formu As I_FacCli, Sql As String, sql1 As String)
Dim i As Long, rutcli As String, total As Double, totalgen As Double, auxcodser As Long, auxdes As String
On Local Error GoTo Error_FacturaCli
MsgTitulo = "Facturación Clientes"
fg_carga ""
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 13: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Facturación Clientes" & IIf(Formu.Option1(2).Value = True, " (Resumido)", " (Detallado)")
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Contrato"
    .TableCell(tcText, 1, 2) = Formu.fpText1(1).text & " - " & Formu.fpayuda(0).Caption
    .TableCell(tcText, 2, 1) = "Rango Fecha"
    .TableCell(tcText, 2, 2) = Formu.fpDateTime1(0).text & " - " & Formu.fpDateTime1(1).text
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    .TableBorder = tbNone
    .EndTable
    .StartTable
    .TableCell(tcCols) = IIf(Formu.Option1(2).Value = True, 4, 5): .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 5500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = IIf(Formu.Option1(2).Value = True, Space(10) & "Servicio", Space(10) & "Fecha")
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
'    Clipboard.Clear
'    Clipboard.SetText Sql
    RS1.Open Sql, vg_db, adOpenStatic
    RS2.Open sql1, vg_db, adOpenStatic
    If RS1.EOF And RS2.EOF Then
        RS1.Close: Set RS1 = Nothing
        RS2.Close: Set RS2 = Nothing
        Close #1
        fg_descarga
        MsgBox "No existe información...", vbExclamation + vbOKOnly, MsgTitulo
        fg_descarga
        Exit Function
    End If
    total = 0: totalgen = 0
    rutcli = "": auxcodser = 0
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            If rutcli <> RS1!cli_codigo Then
               i = i + 1
               .TableCell(tcColSpan, i, 1) = 4
               .TableCell(tcText, i, 1) = fg_PintaRut(RS1!cli_codigo) & " - " & Trim(RS1!cli_nombre) & IIf(Formu.Option1(2).Value = True, "", Space(2) & "(" & RS1!ser_codigo & " - " & Trim(RS1!ser_nombre) & ")")
               .TableCell(tcFontBold, i, 1) = True
               Print #1, .TableCell(tcText, i, 1)
               rutcli = RS1!cli_codigo
               auxcodser = RS1!ser_codigo
               i = i + 2
            End If
            If Formu.Option1(3).Value = True And auxcodser <> RS1!ser_codigo Then
               .TableCell(tcFontUnderline, (i - 1), 4) = True
               .TableCell(tcText, i, 4) = Format(total, fg_Pict(9, vg_DPr))
               .TableCell(tcFontBold, i) = True
               totalgen = Round(totalgen + total, 0)
               total = 0
               Print #1, .TableCell(tcText, i, 4)
               i = i + 1
               .TableCell(tcColSpan, i, 1) = 4
               .TableCell(tcText, i, 1) = fg_PintaRut(RS1!cli_codigo) & " - " & Trim(RS1!cli_nombre) & IIf(Formu.Option1(2).Value = True, "", Space(2) & "(" & RS1!ser_codigo & " - " & Trim(RS1!ser_nombre) & ")")
               .TableCell(tcFontBold, i, 1) = True
               Print #1, .TableCell(tcText, i, 1)
               i = i + 1
               rutcli = RS1!cli_codigo
               auxcodser = RS1!ser_codigo
            End If
            .TableCell(tcText, i, 1) = Space(10) & IIf(Formu.Option1(2).Value = True, RS1!ser_codigo & " - " & Trim(RS1!ser_nombre), Mid(RS1!mir_fecmin, 7, 2) & "/" & Mid(RS1!mir_fecmin, 5, 2) & "/" & Mid(RS1!mir_fecmin, 1, 4))
            .TableCell(tcText, i, 2) = Space(10) & Format(RS1!Cantidad, fg_Pict(9, 0))
            .TableCell(tcText, i, 3) = Space(10) & Format(RS1!prv_preven, fg_Pict(9, 2))
            .TableCell(tcText, i, 4) = Space(10) & Format(RS1!prv_preven * RS1!Cantidad, fg_Pict(9, 0))
            total = total + (RS1!prv_preven * RS1!Cantidad)
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
            RS1.MoveNext
            If Not RS1.EOF Then rutcli2 = RS1!cli_codigo Else rutcli2 = ""
            If rutcli <> rutcli2 Or RS1.EOF Then
                .TableCell(tcFontUnderline, i, 4) = True
                i = i + 1
                '.TableCell(tcText, i, 1) = "Total"
                .TableCell(tcText, i, 4) = Format(total, fg_Pict(9, vg_DPr))
                .TableCell(tcFontBold, i) = True
                totalgen = Round(totalgen + total, 0)
                total = 0
                Print #1, .TableCell(tcText, i, 4)
            End If
            i = i + 1
        Loop
 '       .TableCell(tcRows) = i
 '       .TableBorder = tbBottom
 '       .EndTable
 '       .StartTable
 '       .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
 '       .TableCell(tcColWidth, , 1) = 5500: .TableCell(tcAlign, , 1) = taLeftTop
 '       .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taRightTop
 '       .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taRightTop
 '       .TableCell(tcColWidth, , 4) = 1500: .TableCell(tcAlign, , 4) = taRightTop
 '       .TableCell(tcFontBold, 1) = True
 '       .TableCell(tcText, 1, 3) = "Total General"
 '       .TableCell(tcText, 1, 4) = Format(totalgen, fg_Pict(9, vg_DPr))
 '       Print #1, "||" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
 '       .TableBorder = tbNone
 '       .EndTable
    End If
    RS1.Close: Set RS1 = Nothing
    
    
    If Not RS2.EOF Then
        total = 0: rutcli = "": auxcodser = 0: auxcco = "": auxdes = ""
        Do While Not RS2.EOF
            If rutcli <> RS2!cli_codigo Then
               i = i + 1
               .TableCell(tcColSpan, i, 1) = 4
               .TableCell(tcText, i, 1) = fg_PintaRut(RS2!cli_codigo) & " - " & Trim(RS2!cli_nombre) & IIf(Formu.Option1(2).Value = True, "", Space(2) & "(" & RS2!ser_codigo & " - " & Trim(RS2!ser_nombre) & ")")
               .TableCell(tcFontBold, i, 1) = True
               Print #1, .TableCell(tcText, i, 1)
               i = i + 1
               .TableCell(tcColSpan, i, 1) = 4
               .TableCell(tcText, i, 1) = IIf(Formu.Option1(2).Value = True, "", Space(2) & "(" & Trim(RS2!clc_codigo) & " - " & Trim(RS2!clc_nombre) & ")")
               .TableCell(tcFontBold, i, 1) = True
               Print #1, .TableCell(tcText, i, 1)
               rutcli = RS2!cli_codigo
               auxcodser = RS2!ser_codigo
               auxcco = RS2!clc_codigo
               i = i + 1
            End If
            If Formu.Option1(3).Value = True And (auxcodser <> RS2!ser_codigo Or auxcco <> RS2!clc_codigo) Then
               .TableCell(tcFontUnderline, (i - 1), 4) = True
               .TableCell(tcText, i, 4) = Format(total, fg_Pict(9, vg_DPr))
               .TableCell(tcFontBold, i) = True
               totalgen = Round(totalgen + total, 0)
               total = 0
               Print #1, .TableCell(tcText, i, 4)
               If auxcodser <> RS2!ser_codigo Then
                  i = i + 1
                  .TableCell(tcColSpan, i, 1) = 4
                  .TableCell(tcText, i, 1) = fg_PintaRut(RS2!cli_codigo) & " - " & Trim(RS2!cli_nombre) & IIf(Formu.Option1(2).Value = True, "", Space(2) & "(" & RS2!ser_codigo & " - " & Trim(RS2!ser_nombre) & ")")
                  .TableCell(tcFontBold, i, 1) = True
                  Print #1, .TableCell(tcText, i, 1)
               End If
               i = i + 1
               .TableCell(tcColSpan, i, 1) = 4
               .TableCell(tcText, i, 1) = IIf(Formu.Option1(2).Value = True, "", Space(2) & "(" & Trim(RS2!clc_codigo) & " - " & Trim(RS2!clc_nombre) & ")")
               .TableCell(tcFontBold, i, 1) = True
               Print #1, .TableCell(tcText, i, 1)
               i = i + 1
               rutcli = RS2!cli_codigo
               auxcodser = RS2!ser_codigo
               auxcco = RS2!clc_codigo
            End If
            .TableCell(tcColSpan, i, 1) = 3
            .TableCell(tcText, i, 1) = Space(10) & IIf(Formu.Option1(2).Value = True, RS2!ser_codigo & " - " & Trim(RS2!ser_nombre), Mid(RS2!vtc_fecvta, 7, 2) & "/" & Mid(RS2!vtc_fecvta, 5, 2) & "/" & Mid(RS2!vtc_fecvta, 1, 4) & Space(3) & IIf(RS2!vtd_descripcion <> auxdes, Trim(RS2!vtd_descripcion), ""))
            .TableCell(tcText, i, 4) = Space(10) & Format(RS2!vtd_detmon, fg_Pict(9, 0))
            total = total + (RS2!vtd_detmon)
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
            auxdes = RS2!vtd_descripcion
            RS2.MoveNext
            If Not RS2.EOF Then rutcli2 = RS2!cli_codigo Else rutcli2 = ""
            If rutcli <> rutcli2 Or RS2.EOF Then
                .TableCell(tcFontUnderline, i, 4) = True
                i = i + 1
                '.TableCell(tcText, i, 1) = "Total"
                .TableCell(tcText, i, 4) = Format(total, fg_Pict(9, vg_DPr))
                .TableCell(tcFontBold, i) = True
                totalgen = Round(totalgen + total, 0)
                total = 0
                Print #1, .TableCell(tcText, i, 4)
            End If
            i = i + 1
        Loop
    End If
    RS2.Close: Set RS2 = Nothing
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
Preview.Refresh
Preview.Show 1
fg_descarga
Exit Function
Error_FacturaCli:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_GastosA13(Formu As M_GasA13, Sql As String)
Dim i As Long, rutcli As String, total As Double, totalgen As Double
On Local Error GoTo Error_Bodega
MsgTitulo = "Gastos A13"
fg_carga ""
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
    .FontSize = 8
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 13: .TableCell(tcFontBold, 1) = True
    .TableCell(tcText, 1, 1) = "Gastos A13"
    Print #1, .TableCell(tcText, 1, 1) & Chr(13)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 3
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 9000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcFontBold, , 1) = True
    .TableCell(tcText, 1, 1) = "Contrato"
    .TableCell(tcText, 1, 2) = Formu.fpText1(1).text & " - " & Formu.fpayuda(1).Caption
    .TableCell(tcText, 2, 1) = "Fecha(Mes/Año)"
    .TableCell(tcText, 2, 2) = Trim(Formu.fpDateTime1.text)
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2)
    Print #1, .TableCell(tcText, 2, 1) & "|" & .TableCell(tcText, 2, 2)
    Print #1, .TableCell(tcText, 3, 1) & "|" & .TableCell(tcText, 3, 2)
    .TableBorder = tbNone
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taLeftTop
    .TableCell(tcColWidth, , 4) = 5500: .TableCell(tcAlign, , 4) = taLeftTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcText, 1, 1) = "Descripción"
    .TableCell(tcText, 1, 2) = "Valor"
    .TableCell(tcText, 1, 3) = "Valor Proy."
    .TableCell(tcText, 1, 4) = "Cta. Contable"
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    .TableBorder = tbAll
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 3500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 1500: .TableCell(tcAlign, , 2) = taRightTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, , 4) = 5500: .TableCell(tcAlign, , 4) = taLeftTop
    i = 1
    rutcli = ""
    RS1.Open Sql, vg_db, adOpenStatic
    If RS1.EOF Then
        RS1.Close: Set RS1 = Nothing
        Close #1
        fg_descarga
        MsgBox "No existe información...", vbExclamation + vbOKOnly, MsgTitulo
        Exit Function
    End If
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            .TableCell(tcText, i, 1) = RS1!gas_descri & IIf(RS1!gas_codigo >= 1 And RS1!gas_codigo <= 8, " *", "")
            .TableCell(tcText, i, 2) = Format(RS1!gas_valor, fg_Pict(9, 0))
            .TableCell(tcText, i, 3) = Format(RS1!gas_valpro, fg_Pict(9, 0))
            .TableCell(tcText, i, 4) = IIf(IsNull(RS1!cta_codigo), "", RS1!cta_codigo & " - " & RS1!cta_nombre)
            Print #1, .TableCell(tcText, i, 1) & "|" & .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4)
            RS1.MoveNext: i = i + 1
        Loop
    Else
        i = i + 1
    End If
    i = i + 1
    .TableCell(tcText, i, 1) = "* Registros de Sistema"
    RS1.Close: Set RS1 = Nothing
    .TableCell(tcRows) = i
    .PenColor = &HC0C0C0
    .TableBorder = tbAll
    .EndTable
    Close #1
    .EndDoc
End With
Preview.Refresh
Preview.Show 1
fg_descarga
Exit Function
Error_Bodega:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function
End Function

Public Function I_SalDevVentaServiciosEspeciales(Form As Object, Tipo As String)

Dim rutcli As String, NumDoc As Long, i As Long, j As Long, total As Double, aAp As String, titopc As String, codsec As String, coding As String
Dim numlin As Long, codmer As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String, codreg As Long, codser As Long
Dim tipsal As Boolean, totsec As Double, totmin As Double, totsmi As Double
Dim RS As New ADODB.Recordset

On Local Error GoTo Error_SalDevVentaServiciosEspeciales

fg_carga ""

If Tipo = "SE" Then
     
   MsgTitulo = "Informe de Salida Venta Servicios Especiales"

ElseIf Tipo = "DE" Then
   
   MsgTitulo = "Informe de Devolución Venta Servicios Especiales"

End If

'------- Consultar si salida es resumido ó sector
MsgTitulo = MsgTitulo

Preview.Refresh
Preview.Cls
With Preview.VSPrinter
    
    .Styles.Apply "Default"
    .ExportFormat = vpxRTF
'    .ExportFile = App.Path & "\" & vg_NUsr & "Reporte.rtf"
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

saldevbod1:
    
    vg_Archxls = ""
    vg_Archxls = fg_ArchivoTxt
    Open vg_Archxls For Output As #1
    LogoEmp
    .StartTable
    .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 10500: .TableCell(tcAlign, , 1) = taCenterMiddle
    .TableCell(tcFontSize, 1) = 14: .TableCell(tcFontBold, 1) = True
    .TableCell(tcFontSize, 2) = 8: .TableCell(tcFontBold, 2) = True: .TableCell(tcForeColor, 2) = &H4080&
    .TableCell(tcText, 1, 1) = IIf(Tipo = "SE", "Salida de Venta Servicios Especiales", "Devolución Venta Servicios Especiales")
    .TableCell(tcText, 2, 1) = Form.Label1.Caption
    Print #1, .TableCell(tcText, 1, 1) & "|" & .TableCell(tcText, 2, 1)
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 4
    .TableCell(tcColWidth, , 1) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 3300: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 3) = 1500: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 4) = 4200: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcFontBold, , 1) = True: .TableCell(tcFontBold, , 3) = True
    .TableCell(tcText, 1, 1) = "Folio"
    .TableCell(tcText, 1, 2) = Form.fpLongInteger1(0).text
    .TableCell(tcText, 1, 3) = "Contrato"
    .TableCell(tcText, 1, 4) = Trim(Form.fpText1(0).text) & " - " & Trim(Form.fpayuda(0).Caption)
    .TableCell(tcText, 2, 1) = "F. Producción"
    .TableCell(tcText, 2, 2) = Form.fpDateTime1(0)
    .TableCell(tcText, 2, 3) = "Bodega"
    .TableCell(tcText, 2, 4) = Trim(Left(Form.Combo1(0).List(Form.Combo1(0).ListIndex), 50))
    .TableCell(tcText, 3, 1) = "Servicio Especial"
    .TableCell(tcText, 3, 2) = Trim(Form.fpText1(0).text)
    .TableCell(tcText, 3, 3) = IIf(Tipo = "SE", "Comensales", "")
    .TableCell(tcText, 3, 4) = Form.fpDouble1(0).text
    .TableCell(tcText, 4, 1) = IIf(Tipo = "SE", "Precio Venta", "")
    .TableCell(tcText, 4, 2) = Form.fpDouble1(1).Value
    
'    .TableCell(tcText, 3, 4) = IIf(Not vg_tipser, Trim(Form.fpayuda(0).Caption) & " - " & Trim(Form.fpayuda(3).Caption), "")  'Trim(Left(Form.Combo1(0).List(Form.Combo1(0).ListIndex), 50)))
    Print #1, .TableCell(tcText, 1, 1) & "|"; .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4)
    Print #1, .TableCell(tcText, 2, 1) & "|"; .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4)
    Print #1, .TableCell(tcText, 3, 1) & "|"; .TableCell(tcText, 3, 2) & "|" & .TableCell(tcText, 3, 3) & "|" & .TableCell(tcText, 3, 4)
    Print #1, .TableCell(tcText, 4, 1) & "|"; .TableCell(tcText, 4, 2) & "|" & .TableCell(tcText, 4, 3) & "|" & .TableCell(tcText, 4, 4)
    
    .TableBorder = tbBoxRows
    
    rutcli = Trim(LimpiaDato(Form.fpText1(0).text))
    NumDoc = Form.fpLongInteger1(0).text
    
    .TableBorder = tbNone
    .EndTable
    .text = Chr(13): .text = Chr(13)
    .FontSize = 7
    .StartTable
    .TableCell(tcCols) = 8: .TableCell(tcRows) = 2
    .TableCell(tcColWidth, , 1) = 900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 600: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 800: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
    .TableCell(tcBackColor, 1) = vbYellow: .TableCell(tcFontBold, 1) = True: .TableCell(tcRowHeight, 1) = 230
    .TableCell(tcBackColor, 2) = vbYellow: .TableCell(tcFontBold, 2) = True: .TableCell(tcRowHeight, 2) = 230
    .TableCell(tcText, 1, 1) = ""
    .TableCell(tcText, 1, 2) = ""
    .TableCell(tcText, 1, 3) = ""
    .TableCell(tcText, 1, 4) = "Cantidad"
    .TableCell(tcText, 1, 5) = IIf(Tipo = "DE", "Cantidad", "")
    .TableCell(tcText, 1, 6) = ""
    .TableCell(tcText, 1, 7) = IIf(Tipo = "DE", "Total", "")
    .TableCell(tcText, 1, 8) = "Total"
    .TableCell(tcText, 2, 1) = "Código"
    .TableCell(tcText, 2, 2) = "Descripción"
    .TableCell(tcText, 2, 3) = "Unid."
    .TableCell(tcText, 2, 4) = IIf(Tipo = "DE", "Realizada", "Realizada")
    .TableCell(tcText, 2, 5) = IIf(Tipo = "DE", "Devolver", "")
    .TableCell(tcText, 2, 6) = "P.M.P."
    .TableCell(tcText, 2, 7) = IIf(Tipo = "DE", "Realizada", "")
    .TableCell(tcText, 2, 8) = IIf(Tipo = "DE", "Devolver", "Realizada")
    Print #1, .TableCell(tcText, 1, 1) & "|"; .TableCell(tcText, 1, 2) & "|" & .TableCell(tcText, 1, 3) & "|" & .TableCell(tcText, 1, 4) & "|" & _
              .TableCell(tcText, 1, 5) & "|"; .TableCell(tcText, 1, 6) & "|" & .TableCell(tcText, 1, 7) & "|" & .TableCell(tcText, 1, 8)
    Print #1, .TableCell(tcText, 2, 1) & "|"; .TableCell(tcText, 2, 2) & "|" & .TableCell(tcText, 2, 3) & "|" & .TableCell(tcText, 2, 4) & "|" & _
              .TableCell(tcText, 2, 5) & "|"; .TableCell(tcText, 2, 6) & "|" & .TableCell(tcText, 2, 7) & "|" & .TableCell(tcText, 2, 8)
    .TableBorder = tbBox
    .EndTable
    .StartTable
    .TableCell(tcCols) = 8: .TableCell(tcRows) = 10000
    .TableCell(tcColWidth, , 1) = 900: .TableCell(tcAlign, , 1) = taLeftTop
    .TableCell(tcColWidth, , 2) = 4500: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, , 3) = 600: .TableCell(tcAlign, , 3) = taCenterTop
    .TableCell(tcColWidth, , 4) = 800: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcColWidth, , 5) = 800: .TableCell(tcAlign, , 5) = taRightTop
    .TableCell(tcColWidth, , 6) = 1000: .TableCell(tcAlign, , 6) = taRightTop
    .TableCell(tcColWidth, , 7) = 1000: .TableCell(tcAlign, , 7) = taRightTop
    .TableCell(tcColWidth, , 8) = 1000: .TableCell(tcAlign, , 8) = taRightTop
saldevbod2:
    If Tipo = "SE" Then
          
       Set RS = vg_db.Execute("sgp_Sel_InformeDetVentaServiciosEspeciales '" & Trim(LimpiaDato(Form.fpText1(0).text)) & "', '" & Format(Form.fpDateTime1(0).text, "yyyymmdd") & "', '" & Tipo & "', " & vg_codbod & ", " & Val(Form.fpLongInteger1(0).text) & "")
    
    ElseIf Tipo = "DE" Then
          
       Set RS = vg_db.Execute("sgp_Sel_InformeDetDevVentaServiciosEspeciales '" & Trim(LimpiaDato(Form.fpText1(0).text)) & "', '" & Format(Form.fpDateTime1(0).text, "yyyymmdd") & "', '" & Tipo & "', " & vg_codbod & ", " & Val(Form.fpLongInteger1(0).text) & "")
          
    End If
    i = 1
    j = 1
    total = 0
    totsec = 0
    totsmi = 0
    totmin = 0
    
    Do While Not RS.EOF
       
        '------- Producto
        .TableCell(tcText, i, 1) = Trim(RS!CodigoProducto)
        .TableCell(tcText, i, 2) = Trim(RS!NombreProducto)
        .TableCell(tcText, i, 3) = Trim(RS!UnidadProducto)
        .TableCell(tcText, i, 4) = Format(RS!CantidadMercaderia, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
        If Tipo = "DE" Then
                     
           .TableCell(tcText, i, 5) = Format(RS!Cantidaddevolver, fg_Pict(9, vg_DCa))
        
        Else
        
           .TableCell(tcText, i, 5) = ""
        
        End If
        
        .TableCell(tcText, i, 6) = Format(RS!Preciodocumento, fg_Pict(9, 2)) 'vg_DPr))

        If Tipo = "DE" Then
                   
           .TableCell(tcText, i, 7) = Format((RS!CantidadMercaderia * RS!Preciodocumento), fg_Pict(9, vg_DPr))
           
           .TableCell(tcText, i, 8) = Format((RS!Cantidaddevolver * RS!Preciodocumento), fg_Pict(9, vg_DPr))
           
        Else
        
          .TableCell(tcText, i, 8) = Format((RS!CantidadMercaderia * RS!Preciodocumento), fg_Pict(9, vg_DPr))
        
        End If
        
        Print #1, .TableCell(tcText, i, 1) & "|"; .TableCell(tcText, i, 2) & "|" & .TableCell(tcText, i, 3) & "|" & .TableCell(tcText, i, 4) & "|" & _
                  .TableCell(tcText, i, 5) & "|"; .TableCell(tcText, i, 6) & "|" & .TableCell(tcText, i, 7) & "|" & .TableCell(tcText, i, 8)
        
        
        If Tipo = "DE" Then
           
           If RS!Cantidaddevolver <> 0 Then total = total + Format((RS!Cantidaddevolver * RS!Preciodocumento), fg_Pict(9, 2))
           
           If RS!CantidadMercaderia <> 0 Then totmin = (totmin + (RS!CantidadMercaderia * RS!Preciodocumento))
        
        Else
        
           If RS!CantidadMercaderia <> 0 Then total = total + Format(RS!TotalDocumento, fg_Pict(9, 2))

        End If
        
        RS.MoveNext
        i = i + 1
        .TableCell(tcText, i, 1) = ""
    
    Loop
    RS.Close
    Set RS = Nothing
       
    .TableCell(tcRows) = i - 1
    .PenColor = &HC0C0C0
    .TableBorder = tbBottom
    .EndTable
    .StartTable
    .TableCell(tcCols) = 4: .TableCell(tcRows) = 1
    .TableCell(tcFontSize, 1, 1) = 8: .TableCell(tcFontBold, 1, 1, 1, 4) = True
    .TableCell(tcColWidth, 1, 1) = 7600: .TableCell(tcAlign, , 1) = taRightTop
    .TableCell(tcColWidth, 1, 2) = 1000: .TableCell(tcAlign, , 2) = taLeftTop
    .TableCell(tcColWidth, 1, 3) = 1000: .TableCell(tcAlign, , 3) = taRightTop
    .TableCell(tcColWidth, 1, 4) = 1000: .TableCell(tcAlign, , 4) = taRightTop
    .TableCell(tcText, 1, 1) = "Totales "
    
    If Tipo = "DE" Then
    
       .TableCell(tcFontSize, 1, 3) = 8: .TableCell(tcText, 1, 3) = Format(totmin, fg_Pict(9, vg_DPr))
    
    Else
    
      .TableCell(tcFontSize, 1, 3) = 8: .TableCell(tcText, 1, 3) = "" 'Format(totmin, fg_Pict(9, vg_DPr))
    
    End If
    
    .TableCell(tcFontSize, 1, 4) = 8: .TableCell(tcText, 1, 4) = Format(total, fg_Pict(9, vg_DPr))
    Print #1, "|||||" & .TableCell(tcText, 1, 2) & "|"; .TableCell(tcText, 1, 3) & "|"; .TableCell(tcText, 1, 4)
    .TableBorder = tbNone
    Print #1, " ": Print #1, "|||||" & "_____________________"
    Print #1, " "
    
    If Tipo = "SP" Then
        
        Print #1, Space(100) & "Entregado conforme"
    
    Else
        
        Print #1, Space(100) & "Recibido conforme"
    
    End If
    
    .EndTable
    
    .StartTable
    .TableCell(tcCols) = 2: .TableCell(tcRows) = 6
    .TableCell(tcColWidth, , 1) = 7700
    .TableCell(tcColWidth, , 2) = 2800: .TableCell(tcAlign, , 2) = taCenterTop
    .TableCell(tcFontBold) = True
    .TableCell(tcFontUnderline, 5, 2) = True
    .TableCell(tcText, 5, 2) = Space(40)
    .TableCell(tcText, 6, 2) = IIf(Tipo = "SE", "Entregado conforme", "Recibido conforme")
    .TableBorder = tbNone
    .EndTable
    
    '.CurrentX = 9200
    '.CurrentY = .CurrentY - 400
    '.Text = IIf(tipo = "SE", "_____________________", "____________________")
    .EndDoc
    Close #1
    
End With

Preview.Show 1
fg_descarga

Exit Function

Error_SalDevVentaServiciosEspeciales:
    
    fg_descarga
    If Err = 55 Then Close #1: GoTo saldevbod1
    If Err = -2147467259 Then GoTo saldevbod2 'Resume
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Close #1
    Exit Function

End Function




