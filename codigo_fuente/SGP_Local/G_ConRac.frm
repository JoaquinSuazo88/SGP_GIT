VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form G_ConRac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comparativo de Raciones"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   2865
      TabIndex        =   1
      Top             =   75
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
   End
   Begin ChartfxLibCtl.ChartFX Chart1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11475
      _cx             =   20241
      _cy             =   11456
      Build           =   20
      TypeMask        =   1183322113
      Style           =   -1
      Axis(0).Max     =   90
      nSer            =   5
      NumSer          =   5
      ExtCmd          =   30209
      _Data_          =   "G_ConRac.frx":0000
   End
End
Attribute VB_Name = "G_ConRac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 7140
Me.Width = 11760
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
End Sub

Sub LlenarGrafico(cencos As String, codreg As Long, codser As Long, fecini As Long, fecfin As Long)
Dim titgrafico As String, sql1 As String, sql2 As String, numreg As Long, numser As Long
Dim i As Long, j As Long, inddia As Long, fecesf As Long, nrorac As Long, racteo As Long, racrea As Long
Dim totdoc As Double, tdiateo As Double, tdiarea As Double, vCosFij As Double, cospis As Double, costec As Double
Me.Caption = "Comparativo de Raciones": titgrafico = "Comparativo de Raciones " & Mid(fecini, 5, 2) & "/" & Mid(fecini, 1, 4)
'-------> Traer contrato
RS1.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & cencos & "' AND cli_tipo = 0", vg_db, adOpenStatic
If Not RS1.EOF Then titgrafico = titgrafico & " " & RS1!cli_nombre
RS1.Close: Set RS1 = Nothing
'-------> traer regimen
RS1.Open "SELECT reg_nombre FROM a_regimen WHERE reg_codigo = " & codreg & "", vg_db, adOpenStatic
If Not RS1.EOF Then titgrafico = titgrafico & VgLinea & RS1!reg_nombre
RS1.Close: Set RS1 = Nothing
'-------> Traer regimen
RS1.Open "SELECT ser_nombre FROM a_servicio WHERE ser_codigo = " & codser & "", vg_db, adOpenStatic
If Not RS1.EOF Then titgrafico = titgrafico & " " & VgLinea & RS1!ser_nombre
RS1.Close: Set RS1 = Nothing

titgrafico = titgrafico & " " & VgLinea & " Periodo " & fg_Ctod1(fecini) & " - " & fg_Ctod1(fecfin)

inddia = 0
Dim vecdia As Variant, Fecha As Date
Fecha = fg_Ctod1(fecini)
Do While Fecha <= fg_Ctod1(fecfin)
   inddia = inddia + 1
   Fecha = Fecha + 1
Loop
ReDim vecdia(inddia, 8)
Fecha = fg_Ctod1(fecini)
For i = 1 To UBound(vecdia)
    vecdia(i, 1) = 0 'Fecha
    vecdia(i, 2) = 0 '-------> código regimen
    vecdia(i, 3) = 0 '-------> código servicio
    vecdia(i, 4) = 0 '-------> Raciones planificación teórica
    vecdia(i, 5) = 0 '-------> Raciones planificación real
    vecdia(i, 6) = 0 '-------> Raciones producidas
    vecdia(i, 7) = 0 '-------> Raciones control Venta
    vecdia(i, 8) = 0 '-------> raciones Mermas
    Fecha = Fecha + 1
Next i
'-------> Mover raciones teórica - real
RS1.Open "SELECT min_codreg, min_codser, min_fecmin, SUM(min_racteo) AS racteo, SUM(min_racrea) AS racrea " & _
         "FROM b_minuta " & _
         "WHERE min_cencos = '" & cencos & "' " & _
         "AND   min_codreg = " & codreg & " " & _
         "AND   min_codser = " & codser & "  " & _
         "AND   min_fecmin >= " & fecini & " " & _
         "AND   min_fecmin <= " & fecfin & " " & _
         "GROUP BY min_codreg, min_codser, min_fecmin ORDER BY min_codreg, min_codser, min_fecmin", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      For i = 1 To UBound(vecdia)
          If vecdia(i, 1) = fg_Ctod1(RS1!min_fecmin) And vecdia(i, 2) = RS1!min_codreg And vecdia(i, 3) = RS1!min_codser Then
             vecdia(i, 4) = vecdia(i, 4) + RS1!racteo
             vecdia(i, 5) = vecdia(i, 5) + RS1!racrea
             Exit For
          ElseIf vecdia(i, 1) = 0 And vecdia(i, 2) = 0 And vecdia(i, 3) = 0 Then
             vecdia(i, 1) = fg_Ctod1(RS1!min_fecmin)
             vecdia(i, 2) = RS1!min_codreg
             vecdia(i, 3) = RS1!min_codser
             vecdia(i, 4) = RS1!racteo
             vecdia(i, 5) = RS1!racrea
             Exit For
          End If
      Next i
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
'-------> Mover raciones mermas
'RS1.Open "SELECT a.min_codreg, a.min_codser, a.min_fecmin, SUM(b.mid_nummer) AS mid_nummer " & _
'         "FROM b_minuta a, b_minutadet b " & _
'         "WHERE a.min_codigo = b.mid_codigo " & _
'         "AND   a.min_cencos = '" & cencos & "' " & _
'         "AND   a.min_codreg = " & codreg & " " & _
'         "AND   a.min_codser = " & codser & " " & _
'         "AND   a.min_fecmin >= " & fecini & " " & _
'         "AND   a.min_fecmin <= " & fecfin & " " & _
'         "AND   b.mid_tipmin IN ('2') GROUP BY a.min_codreg, a.min_codser, a.min_fecmin ORDER BY a.min_codreg, a.min_codser, a.min_fecmin", vg_db, adOpenStatic
RS1.Open "SELECT DISTINCT a.min_codreg, a.min_codser, a.min_fecmin, a.min_racrea, SUM(b.mid_nummer) AS mid_nummer, SUM((b.mid_cosrec + b.mid_cosdes) * b.mid_nummer) AS cosmer, SUM((b.mid_cosrec + b.mid_cosdes) * b.mid_numrac) AS cosrea " & _
         "FROM b_minuta a, b_minutadet b " & _
         "WHERE a.min_codigo = b.mid_codigo " & _
         "AND   a.min_cencos = '" & cencos & "' " & _
         "AND   a.min_codreg = " & codreg & " " & _
         "AND   a.min_codser = " & codser & "  " & _
         "AND   a.min_fecmin >= " & fecini & " " & _
         "AND   a.min_fecmin <= " & fecfin & " " & _
         "AND   b.mid_tipmin IN ('2') GROUP BY a.min_codreg, a.min_codser, a.min_fecmin, a.min_racrea ORDER BY a.min_codreg, a.min_codser, a.min_fecmin", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      For i = 1 To UBound(vecdia)
          If vecdia(i, 1) = fg_Ctod1(RS1!min_fecmin) And vecdia(i, 2) = RS1!min_codreg And vecdia(i, 3) = RS1!min_codser Then
'             vecdia(i, 8) = vecdia(i, 8) + RS1!mid_nummer
              vecdia(i, 8) = vecdia(i, 8) + 0
              If RS1!min_racrea > 0 And RS1!cosrea > 0 Then
                 vecdia(i, 8) = vecdia(i, 8) + Round(RS1!cosmer / (RS1!cosrea / RS1!min_racrea), 0)
              End If
             Exit For
          ElseIf vecdia(i, 1) = 0 And vecdia(i, 2) = 0 And vecdia(i, 3) = 0 Then
             vecdia(i, 1) = fg_Ctod1(RS1!min_fecmin)
             vecdia(i, 2) = RS1!min_codreg
             vecdia(i, 3) = RS1!min_codser
'             vecdia(i, 8) = RS1!mid_nummer
             vecdia(i, 8) = 0
             If RS1!min_racrea > 0 And RS1!cosrea > 0 Then
                vecdia(i, 8) = Round(RS1!cosmer / (RS1!cosrea / RS1!min_racrea), 0)
             End If
             Exit For
          End If
      Next i
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
'-------> Mover raciones producidas - control ventas
RS1.Open "SELECT mir_codreg, mir_codser, mir_fecmin, mir_rutcli, SUM(mir_nrorac) AS mir_nrorac " & _
         "FROM b_minutaraciones " & _
         "WHERE mir_cencos = '" & cencos & "' " & _
         "AND   mir_codreg = " & codreg & " " & _
         "AND   mir_codser = " & codser & " " & _
         "AND   mir_fecmin >= " & fecini & " " & _
         "AND   mir_fecmin <= " & fecfin & " " & _
         "AND   mir_rutcli NOT IN ('PERSONAL') " & _
         "GROUP BY mir_codreg, mir_codser, mir_fecmin, mir_rutcli ORDER BY mir_codreg, mir_codser, mir_fecmin", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      For i = 1 To UBound(vecdia)
          If vecdia(i, 1) = fg_Ctod1(RS1!mir_fecmin) And vecdia(i, 2) = RS1!mir_codreg And vecdia(i, 3) = RS1!mir_codser Then
             vecdia(i, IIf(RS1!MIR_RUTCLI = "PRODUCIDAS", 6, 7)) = vecdia(i, IIf(RS1!MIR_RUTCLI = "PRODUCIDAS", 6, 7)) + RS1!mir_nrorac
             Exit For
          ElseIf vecdia(i, 1) = 0 And vecdia(i, 2) = 0 And vecdia(i, 3) = 0 Then
             vecdia(i, 1) = fg_Ctod1(RS1!mir_fecmin)
             vecdia(i, 2) = RS1!mir_codreg
             vecdia(i, 3) = RS1!mir_codser
             vecdia(i, IIf(RS1!MIR_RUTCLI = "PRODUCIDAS", 8, 7)) = RS1!mir_nrorac
             Exit For
          End If
      Next i
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
'-------> Buscar nº días
inddia = 0
For i = 1 To UBound(vecdia)
    If vecdia(i, 1) <> 0 Then inddia = inddia + 1
Next i
With Chart1
    .ToolBar = True
    .ToolBarObj.Moveable = False
    .ToolBarObj(0).Visible = False    'Cargar
    .ToolBarObj(1).Visible = False    'Grabar
    .ToolBarObj(2).Visible = True     'Copiar
    .ToolBarObj(3).Visible = False    'Separador
    .ToolBarObj(4).Visible = True     'Tipo de Grafico
    .ToolBarObj(5).Visible = False    'Color
    .ToolBarObj(6).Visible = False    'Separador
    .ToolBarObj(7).Visible = False    'Grilla vertical
    .ToolBarObj(8).Visible = False    'Grilla horizontal
    .ToolBarObj(9).Visible = False     'Cuadro de Leyenda
    .ToolBarObj(10).Visible = False   'Editor de datos
    .ToolBarObj(11).Visible = False   'Propiedades del grafico
    .ToolBarObj(12).Visible = True    'Separador
    .ToolBarObj(13).Visible = True    '2D/3D
    .ToolBarObj(14).Visible = False   'Rotar
    .ToolBarObj(15).Visible = True    'Profundizar
    .ToolBarObj(16).Visible = True    'Separador
    .ToolBarObj(17).Visible = False   'Zoom
    .ToolBarObj(18).Visible = False   'Preview
    .ToolBarObj(19).Visible = True    'Imprimir
    .ToolBarObj(20).Visible = False   'Separador
    .ToolBarObj(21).Visible = False   'Barras de Herramientas
    .AllowEdit = True
    .AllowResize = True
    .AllowDrag = True
    .MenuBar = False
    .ContextMenus = False             'Menus boton derecho
    .DblClk CHART_NONECLK, 1
    .OpenDataEx COD_VALUES, 5, inddia
    .Title(CHART_TOPTIT) = titgrafico
    .Fonts(CHART_TOPFT) = CF_ARIAL
    .SerLegBoxObj.Visible = True
    .SerLegBoxObj.Docked = 515
    .SerLegBoxObj.BorderStyle = 3
    .SerLegBoxObj.Style = 0
    .Axis(AXIS_Y).Title = "Tendencia raciones" '"Titulo Eje Y"
    .Axis(AXIS_X).Title = "Día " '"Titulo Eje X"
    .Axis(AXIS_X).Visible = True
    .Axis(AXIS_X).ClearLabels
    j = 1
    For i = 0 To inddia
        If j <= inddia Then .Axis(AXIS_X).KeyLabel(i) = Mid(vecdia(j, 1), 1, 5) 'j
        j = j + 1
    Next i
    
    For i = 0 To inddia
        .Series(0).Yvalue(i) = 0
        .Series(1).Yvalue(i) = 0
        .Series(2).Yvalue(i) = 0
        .Series(3).Yvalue(i) = 0
        .Series(4).Yvalue(i) = 0
    Next i
    
    auxfec = 0: tdiateo = 0: tdiarea = 0
    .Series(0).Visible = True
    For i = 1 To UBound(vecdia)
        If Trim(vecdia(i, 1)) <> "" Then
           .Series(0).Yvalue(i - 1) = Format(vecdia(i, 4), fg_Pict(9, 0))
           .Series(1).Yvalue(i - 1) = Format(vecdia(i, 5), fg_Pict(9, 0))
           .Series(2).Yvalue(i - 1) = Format(vecdia(i, 6), fg_Pict(9, 0))
           .Series(3).Yvalue(i - 1) = Format(vecdia(i, 7), fg_Pict(9, 0))
           .Series(4).Yvalue(i - 1) = Format(vecdia(i, 8), fg_Pict(9, 0))
        End If
    Next i
    .OpenDataEx COD_COLORS, 5, 0
    For i = 0 To 4
        If i = 0 Then .Series(i).Legend = "Plan. Teórica"
        If i = 1 Then .Series(i).Legend = "Plan. Real"
        If i = 2 Then .Series(i).color = RGB(80, 240, 60): .Series(i).Legend = "Producidas"
        If i = 3 Then .Series(i).color = RGB(40, 240, 202): .Series(i).Legend = "Ctrl. Venta"
        If i = 4 Then .Series(i).Legend = "Mermas"
    Next i
    .CloseData COD_VALUES
    .CloseData COD_COLORS
End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    Me.Hide
    Unload Me
End Select
End Sub
