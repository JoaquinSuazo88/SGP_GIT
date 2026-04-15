VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Begin VB.Form G_TeoRea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costo Teórico - Real - Food Cost"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin ChartfxLibCtl.ChartFX Chart1 
      Height          =   6495
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   10995
      _cx             =   19394
      _cy             =   11456
      Build           =   20
      TypeMask        =   1183322113
      Axis(0).Max     =   90
      _Data_          =   "G_TeoRea.frx":0000
   End
End
Attribute VB_Name = "G_TeoRea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 7080
Me.Width = 11145
fg_centra Me
End Sub

Sub LlenarGrafico(cencos As String, codreg As Long, codser As Long, fecini As Long, fecfin As Long)
Dim titgrafico As String
Dim i As Long, j As Long, inddia As Long
Dim totdoc As Double, tdiateo As Double, tdiarea As Double

titgrafico = "Costo Teorico - Real - Food Cost " & Mid(fecini, 5, 2) & "/" & Mid(fecini, 1, 4)
RS1.Open "select cli_codigo, cli_nombre from b_clientes where cli_codigo='" & cencos & "' and cli_tipo=0", vg_db, adOpenStatic
If Not RS1.EOF Then titgrafico = titgrafico & " " & RS1!cli_nombre
RS1.Close: Set RS1 = Nothing
RS1.Open "select reg_nombre from a_regimen where reg_codigo=" & codreg & "", vg_db, adOpenStatic
If Not RS1.EOF Then titgrafico = titgrafico & " " & RS1!reg_nombre
RS1.Close: Set RS1 = Nothing
If codser > 0 Then
   RS1.Open "select ser_nombre from a_servicio where ser_codigo=" & codser & "", vg_db, adOpenStatic
   If Not RS1.EOF Then titgrafico = titgrafico & " " & RS1!ser_nombre
   RS1.Close: Set RS1 = Nothing
Else
   titgrafico = titgrafico & " " & "Todos Servicios"
End If

'------- Buscar Nº Dìas
RS1.Open "select distinct b_minuta.min_fecmin " & _
         "from  b_minuta, b_minutadet " & _
         "where b_minuta.min_codigo=b_minutadet.mid_codigo " & _
         "and   b_minuta.min_cencos='" & cencos & "' " & _
         "and   b_minuta.min_codreg=" & codreg & " " & _
         "and  (b_minuta.min_codser=" & codser & " or " & codser & "=0) " & _
         "and   val(mid(b_minuta.min_fecmin,1 ,6))=" & Val(Mid(fecini, 1, 6)) & " " & _
         "order by b_minuta.min_fecmin desc ", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
inddia = Mid(RS1!min_fecmin, 7, 2)
RS1.Close: Set RS1 = Nothing

Chart1.Toolbar = True
Chart1.ToolBarObj.Moveable = False
Chart1.ToolBarObj(0).Visible = False    'Cargar
Chart1.ToolBarObj(1).Visible = False    'Grabar
Chart1.ToolBarObj(2).Visible = True     'Copiar
Chart1.ToolBarObj(3).Visible = False    'Separador
Chart1.ToolBarObj(4).Visible = True     'Tipo de Grafico
Chart1.ToolBarObj(5).Visible = False    'Color
Chart1.ToolBarObj(6).Visible = False    'Separador
Chart1.ToolBarObj(7).Visible = False    'Grilla vertical
Chart1.ToolBarObj(8).Visible = False    'Grilla horizontal
Chart1.ToolBarObj(9).Visible = False     'Cuadro de Leyenda
Chart1.ToolBarObj(10).Visible = False   'Editor de datos
Chart1.ToolBarObj(11).Visible = False   'Propiedades del grafico
Chart1.ToolBarObj(12).Visible = True    'Separador
Chart1.ToolBarObj(13).Visible = True    '2D/3D
Chart1.ToolBarObj(14).Visible = False   'Rotar
Chart1.ToolBarObj(15).Visible = True    'Profundizar
Chart1.ToolBarObj(16).Visible = True    'Separador
Chart1.ToolBarObj(17).Visible = False   'Zoom
Chart1.ToolBarObj(18).Visible = False   'Preview
Chart1.ToolBarObj(19).Visible = True    'Imprimir
Chart1.ToolBarObj(20).Visible = False   'Separador
Chart1.ToolBarObj(21).Visible = False   'Barras de Herramientas
Chart1.AllowEdit = True
Chart1.AllowResize = True
Chart1.AllowDrag = True
Chart1.MenuBar = False
Chart1.ContextMenus = False             'Menus boton derecho
Chart1.DblClk CHART_NONECLK, 1
Chart1.OpenDataEx COD_VALUES, 3, inddia
Chart1.TITLE(CHART_TOPTIT) = titgrafico
Chart1.Fonts(CHART_TOPFT) = CF_ARIAL
Chart1.SerLegBoxObj.Visible = True
Chart1.SerLegBoxObj.Docked = 515
Chart1.SerLegBoxObj.BorderStyle = 3
Chart1.SerLegBoxObj.Style = 0
Chart1.Axis(AXIS_Y).TITLE = "Costo " '"Titulo Eje Y"
Chart1.Axis(AXIS_X).TITLE = "Día " '"Titulo Eje X"
Chart1.Axis(AXIS_X).Visible = True
Chart1.Axis(AXIS_X).ClearLabels
j = 1
For i = 0 To inddia
    Chart1.Axis(AXIS_X).KeyLabel(i) = j
    j = j + 1
Next i

RS1.Open "select b_minutadet.mid_tipmin, b_minutadet.mid_numlin, b_minutadet.mid_codrec, " & _
         "b_minutadet.mid_descri, b_minutadet.mid_cosrec, b_minuta.min_fecmin, b_minuta.min_indblo, " & _
         "b_receta.rec_nombre, b_receta.rec_nomfan, b_minutadet.mid_numrac " & _
         "from  b_receta, b_minuta, b_minutadet " & _
         "where b_minuta.min_codigo=b_minutadet.mid_codigo " & _
         "and   b_minutadet.mid_codrec=b_receta.rec_codigo " & _
         "and   b_minuta.min_cencos='" & cencos & "' " & _
         "and   b_minuta.min_codreg=" & codreg & " " & _
         "and  (b_minuta.min_codser=" & codser & " or " & codser & "=0) " & _
         "and   b_minuta.min_fecmin>=" & fecini & " " & _
         "and   b_minuta.min_fecmin<=" & fecfin & " " & _
         "order by b_minuta.min_fecmin, b_minutadet.mid_numlin", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
auxfec = 0: tdiateo = 0: tdiarea = 0
Do While Not RS1.EOF
   If RS1!min_fecmin <> auxfec Then
      If auxfec > 0 Then
         '------- Mover teórico
         Chart1.Series(0).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(tdiateo, fg_Pict(6, 2))
         '------- Mover Real
         Chart1.Series(1).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(tdiarea, fg_Pict(6, 2))
         totdoc = 0
         '------- Traer salida & devolución
         RS2.Open "select b_totventas.tov_codreg, b_totventas.tov_codser, sum(IIf(b_totventas.tov_tipdoc='SP',b_detventas.dev_ptotal,'-' & b_detventas.dev_ptotal)) as totdoc " & _
                  "from b_totventas, b_detventas, b_productos where b_totventas.tov_rutcli=b_detventas.dev_rutcli " & _
                  "and b_totventas.tov_tipdoc=b_detventas.dev_tipdoc and b_totventas.tov_numdoc=b_detventas.dev_numdoc " & _
                  "and b_detventas.dev_codmer=b_productos.pro_codigo and b_productos.pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') " & _
                  "and b_totventas.tov_codreg=" & codreg & " and (b_totventas.tov_codser=" & codser & " or " & codser & "=0) " & _
                  "and (b_totventas.tov_tipdoc='SP' or b_totventas.tov_tipdoc='DP') and b_detventas.dev_canmer<>0 " & _
                  "and b_totventas.tov_estdoc<>'A' and b_totventas.tov_fecpro=cdate('" & fg_Ctod1(auxfec) & "') group by b_totventas.tov_codreg, b_totventas.tov_codser", vg_db, adOpenStatic
         If Not RS2.EOF Then totdoc = RS2!totdoc
         RS2.Close: Set RS2 = Nothing
         Chart1.Series(2).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(totdoc, fg_Pict(6, 2))
      End If
      auxfec = RS1!min_fecmin
      tdiateo = 0: tdiarea = 0
      j = j + 1
   End If
   If RS1!mid_tipmin = "1" Then
      tdiateo = CCur(tdiateo + (RS1!mid_cosrec * RS1!mid_numrac))
   ElseIf RS1!mid_tipmin = "2" Then
      tdiarea = CCur(tdiarea + (RS1!mid_cosrec * RS1!mid_numrac))
   End If
   RS1.MoveNext
Loop
totdoc = 0
'------- Mover costo teórico
Chart1.Series(0).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(tdiateo, fg_Pict(6, 2))
Chart1.Series(1).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(tdiarea, fg_Pict(6, 2))
'------- Traer salida & devolución
RS2.Open "select b_totventas.tov_codreg, b_totventas.tov_codser, sum(IIf(b_totventas.tov_tipdoc='SP',b_detventas.dev_ptotal,'-' & b_detventas.dev_ptotal)) as totdoc " & _
         "from b_totventas, b_detventas, b_productos where b_totventas.tov_rutcli=b_detventas.dev_rutcli " & _
         "and b_totventas.tov_tipdoc=b_detventas.dev_tipdoc and b_totventas.tov_numdoc=b_detventas.dev_numdoc " & _
         "and b_detventas.dev_codmer=b_productos.pro_codigo and b_productos.pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') " & _
         "and b_totventas.tov_codreg=" & codreg & " and (b_totventas.tov_codser=" & codser & " or " & codser & "=0) " & _
         "and (b_totventas.tov_tipdoc='SP' or b_totventas.tov_tipdoc='DP') and b_detventas.dev_canmer<>0 " & _
         "and b_totventas.tov_estdoc<>'A' and b_totventas.tov_fecpro=cdate('" & fg_Ctod1(auxfec) & "') group by b_totventas.tov_codreg, b_totventas.tov_codser", vg_db, adOpenStatic
If Not RS2.EOF Then totdoc = RS2!totdoc
RS2.Close: Set RS2 = Nothing
Chart1.Series(2).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(totdoc, fg_Pict(6, 2))
RS1.Close: Set RS1 = Nothing

Chart1.OpenDataEx COD_COLORS, 3, 0
For i = 0 To 2
    If i = 0 Then Chart1.Series(i).Legend = "Teórico"
    If i = 1 Then Chart1.Series(i).Legend = "Real"
    If i = 2 Then Chart1.Series(i).color = RGB(80, 240, 60): Chart1.Series(i).Legend = "Food Cost"
Next i
Chart1.CloseData COD_VALUES
Chart1.CloseData COD_COLORS
End Sub



