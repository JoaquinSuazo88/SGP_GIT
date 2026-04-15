VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_TomInv 
   Caption         =   "Toma de Inventario"
   ClientHeight    =   7230
   ClientLeft      =   390
   ClientTop       =   2130
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin LpLib.fpList fpList1 
      Height          =   270
      Left            =   3810
      TabIndex        =   6
      Top             =   285
      Width           =   1335
      _Version        =   196608
      _ExtentX        =   2355
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      Columns         =   0
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   1
      BorderDropShadow=   1
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "M_TomInv.frx":0000
   End
   Begin VB.Frame Frame2 
      Height          =   5790
      Left            =   45
      TabIndex        =   8
      Top             =   1395
      Width           =   11520
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5130
         Left            =   90
         TabIndex        =   9
         Top             =   180
         Width           =   11340
         _Version        =   393216
         _ExtentX        =   20003
         _ExtentY        =   9049
         _StockProps     =   64
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         MaxRows         =   50
         SpreadDesigner  =   "M_TomInv.frx":0258
         ScrollBarTrack  =   3
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000A&
         Height          =   435
         Left            =   1905
         TabIndex        =   12
         Top             =   5265
         Width           =   4125
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   13
            Top             =   135
            Width           =   4020
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   90
         TabIndex        =   10
         Top             =   5265
         Width           =   1785
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   11
            Top             =   135
            Width           =   1680
         End
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   6945
         TabIndex        =   14
         Top             =   5340
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   688
         ButtonWidth     =   2963
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar Producto"
               Description     =   "Agregar Productos"
               Object.ToolTipText     =   "Agregar Producto"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar Producto "
               Description     =   "Eliminar Producto "
               Object.ToolTipText     =   "Eliminar Producto "
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10470
         Top             =   5235
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_TomInv.frx":0B48
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_TomInv.frx":0E62
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   990
      Left            =   45
      TabIndex        =   3
      Top             =   405
      Width           =   11520
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cierre de Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   585
         Width           =   1500
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   2550
      End
      Begin EditLib.fpDateTime Date1 
         Height          =   345
         Index           =   0
         Left            =   1785
         TabIndex        =   0
         Top             =   195
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483637
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   16777215
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   3
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   16777215
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   ""
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "dd/mm/yyyy"
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   0
         IncHour         =   0
         IncMinute       =   0
         IncSecond       =   0
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1815
         TabIndex        =   17
         Top             =   630
         Width           =   210
      End
      Begin VB.Label Label2 
         Caption         =   "Comentario Fecha Cierre"
         Height          =   210
         Left            =   2145
         TabIndex        =   16
         Top             =   615
         Width           =   4650
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   4365
         TabIndex        =   7
         Top             =   240
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   510
         TabIndex        =   5
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bodega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3570
         TabIndex        =   4
         Top             =   225
         Width           =   660
      End
   End
End
Attribute VB_Name = "M_TomInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String, Est As Boolean
Dim Msgtitulo As String
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset

Private Sub Check1_Click()
If Est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 6, 0, modo
Label2.Caption = ""
End Sub

Private Sub Combo1_Click(Index As Integer)
If Est Then Exit Sub
Est = True: Date1(0).Text = "": Est = False
MuestraDatosGrilla
On Error Resume Next: vaSpread1.SetFocus
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Date1_Change(Index As Integer)
Dim v_fecinv As Variant, v_codbod As Long
If Est Then Exit Sub
Label2.Caption = ""
Check1.Enabled = False
Est = True: Check1.Value = 0: Est = False
v_codbod = fg_codigocbo(Combo1, 0, 10, 0)
v_fecinv = Format(Date1(0).Text, "yyyymmdd")
RS1.Open "select distinct tin_fectom from b_tomainv where tin_codbod=" & v_codbod & " order by tin_fectom desc", vg_db, adOpenStatic
If Not RS1.EOF Then If Val(v_fecinv) <= RS1!tin_fectom Then Est = True: Date1(0).Text = Str(CDate(fg_Ctod1(RS1!tin_fectom)) + 1): Est = False
RS1.Close: Set RS1 = Nothing
End Sub

Private Sub Date1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Date1_LostFocus(Index As Integer)
If IsDate(Date1(0).Text) = False Then On Error Resume Next: Date1(0).SetFocus
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
On Local Error GoTo Error_Partida
'Dim btnX As Button
Me.Width = 11745
Me.Height = 7635
Me.HelpContextID = vg_OpcM
Msgtitulo = "Toma de Inventario"
fg_centra Me
modo = ""
Gl_Mo_Botones Me, 6
'---Formato de Celdas
vaSpread1.Row = -1
vaSpread1.Col = 4: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 6: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
vaSpread1.Col = 7: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
'---Trae todos los registros de las Bodegas Disponibles
Est = True
Combo1(0).Clear
RS1.Open "select * from a_bodega order by bod_nombre", vg_db, adOpenStatic
Do While Not RS1.EOF
    Combo1(0).AddItem RS1!bod_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!bod_codigo), 10) & ")"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
If Combo1(0).ListCount > 0 Then Combo1(0).ListIndex = 0
Date1(0).Text = ""
Est = False
fpList1.Visible = False
MuestraDatosGrilla
Exit Sub
Error_Partida:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, Msgtitulo
End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then
    Frame1.Left = (Me.Width \ 2) - (Frame1.Width \ 2)
    Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)
ElseIf Me.WindowState = 0 Then
    Frame1.Left = 45
    Frame2.Left = 45
End If
End Sub

Private Sub fpList1_Click()
If Est Then Exit Sub
fpList1.Visible = False
Me.Refresh
Est = True: Date1(0).Text = fpList1.List(fpList1.ListIndex): Est = False
MuestraDatosGrilla
'---- Reviso si hay ajuste y bloqueo --------
'vaSpread1.Row = -1: vaSpread1.Col = 5
'RS1.Open "select count(tov_fecemi) as suma from b_totventas where tov_fecemi=Cdate('" & Date1(0).Text & "') " & _
'         "and tov_codbod=" & Val(fg_codigocbo(Combo1, 0, 10, 0)) & " and tov_tipdoc='AI' and tov_estdoc<>'A'", vg_db, adOpenStatic
'vaSpread1.Lock = IIf(fpList1.ListIndex > 0 Or RS1!suma > 0, True, False)
'RS1.Close: Set RS1 = Nothing
'--------------------------------------------
On Error Resume Next: vaSpread1.SetFocus
End Sub

Private Sub fpList1_LostFocus()
fpList1.Visible = False
End Sub

Private Sub fpList1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Est = True: fpList1.Selected(fpList1.MouseOverRow) = True: Est = False
End Sub

Private Sub Text1_Change(Index As Integer)
vaSpread1_Click Index, 0
vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
vaSpread1.ColUserSortIndicator(Index) = ColUserSortIndicatorAscending
vaSpread1.SortKey(1) = Index: vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).Text), SearchFlagsGreaterOrEqual)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(1) = "": Text1(2) = ""
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
vaSpread1.SetActiveCell 5, vaSpread1.ActiveRow
On Error Resume Next: vaSpread1.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long, v_codbod As Long, CodPro As String, stofis As Double, stosis As Double, v_fecinv As Variant, sqlTMP As String, fecha As String, ciemes As Long
On Error GoTo Man_Error
v_fecinv = Format(Date1(0).Text, "yyyymmdd")
v_codbod = fg_codigocbo(Combo1, 0, 10, 0)
Select Case Button.Index
Case 1
    modo = "A": Gl_Ac_Botones Me, 6, 0, modo
    Date1(0).Text = ""
    Date1(0).Text = Format(Date, "dd/mm/yyyy")
    Date1(0).Enabled = True
    vaSpread1.MaxRows = 0
    On Error Resume Next: vaSpread1.SetFocus
Case 3 'Modifica
    RS1.Open "select max(tin_fectom) as fecha from b_tomainv where tin_codbod=" & Val(fg_codigocbo(M_TomInv.Combo1, 0, 10, 0)), vg_db, adOpenStatic
    If Not RS1.EOF Then
        If fg_Ctod1(RS1!fecha) <> Date1(0).Text Then
            MsgBox "Solo puede modificar el último inventario" & vbCrLf & _
                   "si no se ha generado el ajuste...", vbExclamation + vbOKOnly, Msgtitulo
            RS1.Close: Set RS1 = Nothing: Exit Sub
        End If
    End If
    RS1.Close: Set RS1 = Nothing
    modo = "M"
    vaSpread1.Row = 1: vaSpread1.Col = 5
    vaSpread1.EditMode = True
    Gl_Ac_Botones Me, 6, 0, modo
Case 5 'Borra_Datos
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS1.Open "select distinct tin_fectom from b_tomainv where tin_codbod=" & v_codbod & " order by tin_fectom desc", vg_db, adOpenStatic
    If Not RS1.EOF Then If Str(v_fecinv) <> Str(RS1!tin_fectom) Then MsgBox "Solo puede eliminar la ultima toma de inventario...", vbExclamation + vbOKOnly, Msgtitulo: RS1.Close: Set RS1 = Nothing: Exit Sub
    RS1.Close: Set RS1 = Nothing
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    'Detalle - Devuelve stock
    RS1.Open "select dev.dev_codmer, dev.dev_canmer, aju.aju_tipo from b_totventas tov, b_detventas dev, a_tipoajuste aju " & _
             "where tov.tov_rutcli=dev.dev_rutcli and tov.tov_tipdoc=dev.dev_tipdoc and tov.tov_numdoc=dev.dev_numdoc " & _
             "and tov.tov_codser=aju.aju_codigo and tov.tov_fecemi=Cdate('" & Date1(0).Text & "') and tov_codbod=" & v_codbod & " " & _
             "and tov.tov_tipdoc='AI' and tov.tov_estdoc<>'A' order by dev.dev_numlin", vg_db, adOpenStatic
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            vg_db.Execute "update b_bodegas set bod_canmer=bod_canmer" & IIf(RS1!aju_tipo = "A", "-", "+") & RS1!dev_canmer & " " & _
                          "where bod_codpro='" & RS1!dev_codmer & "' and bod_codbod=" & v_codbod
            RS1.MoveNext
        Loop
    End If
    RS1.Close: Set RS1 = Nothing
    vg_db.Execute "delete b_tomainv from b_tomainv where tin_fectom=" & Val(v_fecinv) & " and tin_codbod=" & v_codbod
    vg_db.Execute "delete dev.* from b_totventas tov inner join b_detventas dev " & _
                  "on (tov.tov_numdoc=dev.dev_numdoc) and (tov.tov_tipdoc=dev.dev_tipdoc) " & _
                  "and (tov.tov_rutcli=dev.dev_rutcli) " & _
                  "where tov.tov_fecemi=Cdate('" & Date1(0).Text & "') and tov.tov_codbod=" & v_codbod & " " & _
                  "and tov.tov_tipdoc='AI'"
    vg_db.Execute "delete from b_totventas where tov_fecemi=Cdate('" & Date1(0).Text & "') and tov_codbod=" & v_codbod & " " & _
                  "and tov_tipdoc='AI'"
    vg_db.CommitTrans
    Est = True
    RS1.Open "select distinct tin_fectom from b_tomainv where tin_codbod=" & v_codbod & " order by tin_fectom desc", vg_db, adOpenStatic
    If Not RS1.EOF Then Date1(0).Text = fg_Ctod1(RS1!tin_fectom) Else Date1(0).Text = ""
    RS1.Close: Set RS1 = Nothing
    Est = False
    modo = ""
    MuestraDatosGrilla
Case 7 'Actualizar
    modo = ""
    MuestraDatosGrilla
Case 10 'Cancelar
    RS1.Open "select count(tov_fecemi) as suma from b_totventas where " & _
         "tov_codbod=" & Val(fg_codigocbo(Combo1, 0, 10, 0)) & " and tov_tipdoc='AI' and tov_estdoc<>'A'", vg_db, adOpenStatic
    If vaSpread1.MaxRows = 0 And modo = "A" And RS1!suma = 0 Then RS1.Close: Set RS1 = Nothing: Exit Sub
    RS1.Close: Set RS1 = Nothing
    If MsgBox("Cancela toma de inventario...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    '---Muestra el ultimo inventario
    Est = True
    RS1.Open "select distinct tin_fectom from b_tomainv where tin_codbod=" & v_codbod & " order by tin_fectom desc", vg_db, adOpenStatic
    If Not RS1.EOF Then Date1(0).Text = fg_Ctod1(RS1!tin_fectom)
    RS1.Close: Set RS1 = Nothing
    Est = False
    modo = "": MuestraDatosGrilla
Case 12 'Confirmar
    If Trim(Date1(0).Text) = "" Then MsgBox "Falta dato importante...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    vg_db.BeginTrans
    If modo = "M" Then
        If vaSpread1.MaxRows > 0 Then
            '--------- Revisa Cierre de Mes ----------
            AñoMes = Format(Date1(0).Value, "yyyy") & Format(Date1(0).Value, "mm")
            'RS1.Open "select tin_fectom, tin_ciemes from b_tomainv where tin_fectom=" & v_fecinv & " and tin_codbod=" & v_codbod & " " & _
                     "group by tin_fectom, tin_ciemes", vg_db, adOpenStatic
            If Check1.Enabled = True And Check1.Value = 1 Then
            'If Not RS1.EOF And Check1.Value = 1 And Check1.Enabled = True Then
                vg_db.Execute "update b_tomainv set tin_ciemes=0 where left(tin_fectom,6)=" & AñoMes & " and tin_codbod=" & v_codbod
                vg_db.Execute "update b_tomainv set tin_ciemes=" & AñoMes & " " & _
                              "where tin_fectom=" & v_fecinv & " and tin_codbod=" & v_codbod
            ElseIf Check1.Enabled = True And Check1.Value = 0 Then
                vg_db.Execute "update b_tomainv set tin_ciemes=0 where left(tin_fectom,6)=" & AñoMes & " and tin_codbod=" & v_codbod
            Else
                vg_db.Execute "update b_tomainv set tin_ciemes=0 where tin_fectom=" & v_fecinv & " and tin_codbod=" & v_codbod
            End If
            'RS1.Close: Set RS1 = Nothing
            '-----------------------------------------
        End If
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Row = i
            vaSpread1.Col = 1: CodPro = vaSpread1.Text
            vaSpread1.Col = 4: stosis = vaSpread1.Text
            vaSpread1.Col = 5: stofis = vaSpread1.Text
            vg_db.Execute "update b_tomainv set tin_stofis=" & stofis & ", tin_stosis=" & stosis & " " & _
                          "where tin_fectom=" & v_fecinv & " and tin_codbod=" & v_codbod & " and tin_codpro='" & CodPro & "'"
        Next
        modo = ""
        MuestraDatosGrilla
    ElseIf modo = "A" Then
        MuestraDatosGrilla
    End If
    vg_db.CommitTrans
    Date1(0).Enabled = False
    modo = "": Gl_Ac_Botones Me, 6, 1, modo
    On Error Resume Next: vaSpread1.SetFocus
Case 15 'Imprimir
    I_TomInv.Show 1
Case 18 'Historico
    If fpList1.Visible = True Then fpList1.Visible = False: On Error Resume Next: vaSpread1.SetFocus: Exit Sub
    v_codbod = fg_codigocbo(Combo1, 0, 10, 0)
    fpList1.Clear
    fpList1.Visible = True
    fpList1.ZOrder
    RS1.Open "select distinct tin_fectom from b_tomainv where tin_codbod=" & v_codbod & " order by tin_fectom desc", vg_db, adOpenStatic
    i = 75
    Do While Not RS1.EOF
        fpList1.AddItem fg_Ctod1(RS1!tin_fectom)
        RS1.MoveNext: i = i + 195
    Loop
    RS1.Close: Set RS1 = Nothing
    fpList1.Height = IIf(i > 2025, 2025, i)
    Est = True: fpList1.Selected(0) = True: Est = False
    On Error Resume Next: fpList1.SetFocus
Case 19 'Filtrar
    If Date1(0).Text = "" Then MsgBox "Debe ingresar fecha...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    vg_codigo = ""
    B_Produc.Show 1
    If vg_codigo <> "|Ok|" Then Exit Sub
    MuestraDatosGrilla
    modo = "M"
Case 21 'Ajustar Inventario
    RS1.Open "select count(tin_codpro) as suma from b_tomainv where tin_fectom=" & v_fecinv & " and tin_codbod=" & v_codbod & " " & _
             "and tin_stosis<>tin_stofis", vg_db, adOpenStatic
    RS2.Open "select count(tov_fecemi) as suma from b_totventas where tov_fecemi=Cdate('" & Date1(0).Text & "') " & _
             "and tov_codbod=" & v_codbod & " and tov_tipdoc='AI' and tov_estdoc<>'A'", vg_db, adOpenStatic
    If RS1!suma = 0 And RS2!suma = 0 Then MsgBox "No existen diferencias en toma de inventario...", vbCritical, Msgtitulo Else M_AjuInv.Show 1
    RS2.Close: Set RS2 = Nothing
    RS1.Close: Set RS1 = Nothing
    modo = ""
    MuestraDatosGrilla
Case 22 'Anular Ajuste Inventario
    If MsgBox("Anula ajuste de inventario...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    'Detalle - Devuelve stock
    RS1.Open "select dev.dev_codmer, dev.dev_canmer, aju.aju_tipo from b_totventas tov, b_detventas dev, a_tipoajuste aju " & _
             "where tov.tov_rutcli=dev.dev_rutcli and tov.tov_tipdoc=dev.dev_tipdoc and tov.tov_numdoc=dev.dev_numdoc " & _
             "and tov.tov_codser=aju.aju_codigo and tov.tov_fecemi=Cdate('" & Date1(0).Text & "') and tov_codbod=" & v_codbod & " " & _
             "and tov.tov_tipdoc='AI' and tov.tov_estdoc<>'A' order by dev.dev_numlin", vg_db, adOpenStatic
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            vg_db.Execute "update b_bodegas set bod_canmer=bod_canmer" & IIf(RS1!aju_tipo = "A", "-", "+") & RS1!dev_canmer & " " & _
                          "where bod_codpro='" & RS1!dev_codmer & "' and bod_codbod=" & v_codbod
            RS1.MoveNext
        Loop
    End If
    RS1.Close: Set RS1 = Nothing
    'Encabezado
    vg_db.Execute "update b_totventas set tov_estdoc='A' where tov_fecemi=Cdate('" & Date1(0).Text & "') " & _
                  "and tov_codbod=" & Val(fg_codigocbo(Combo1, 0, 10, 0)) & " and tov_tipdoc='AI' and tov_estdoc<>'A'"
    vg_db.CommitTrans
    MsgBox "Ajuste anulado...", vbInformation, Msgtitulo
    modo = ""
    MuestraDatosGrilla
Case 24
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then
    MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
    vg_db.RollbackTrans
    Exit Sub
End If
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub MuestraDatosGrilla()
On Local Error GoTo Error_Mover
fg_carga ""
Dim v_fecinv  As Long, v_codbod As Long, aAp As String, i As Long, feccie As Long
Dim sqlTMP As String, fecha As String, sqlPROPON As String, AñoMes As Long
Label2.Caption = ""
Check1.Enabled = False
Est = True: Check1.Value = 0: Est = False
v_codbod = fg_codigocbo(Combo1, 0, 10, 0)
If Trim(Date1(0).Text) = "" Then
    '---Muestra el ultimo inventario
    Est = True
    RS1.Open "select distinct tin_fectom from b_tomainv where tin_codbod=" & v_codbod & " order by tin_fectom desc", vg_db, adOpenStatic
    If Not RS1.EOF Then
        Date1(0).Text = fg_Ctod1(RS1!tin_fectom)
        modo = ""
    Else
        RS1.Close: Set RS1 = Nothing
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        Est = False
        fg_descarga
        Exit Sub
    End If
    RS1.Close: Set RS1 = Nothing
    Est = False
End If
v_fecinv = Format(Date1(0).Text, "yyyymmdd")
'--------- Revisa Cierre de Mes ----------
AñoMes = Format(Date1(0).Value, "yyyy") & Format(Date1(0).Value, "mm")
RS1.Open "select tin_fectom, tin_ciemes from b_tomainv where left(tin_fectom,6)=" & AñoMes & " and tin_codbod=" & v_codbod & " " & _
         "and tin_ciemes=" & AñoMes & " group by tin_fectom, tin_ciemes", vg_db, adOpenStatic
If Not RS1.EOF Then
    Est = True: Check1.Value = 1: Est = False
    Label2.Caption = "Fecha de cierre " & fg_Ctod1(RS1!tin_fectom)
    feccie = RS1!tin_fectom
Else
    Label2.Caption = "No hay cierre"
    feccie = 0
End If
RS1.Close: Set RS1 = Nothing
'-----------------------------------------
'--------- Reviso si hay ajuste ----------
Toolbar2.Enabled = False
sqlPROPON = "tin.tin_propon"
sqlTMP = IIf(vg_codigo = "|Ok|", " and pro.pro_codigo in (select * from " & Trim(vg_NUsr) & "_tmp_filtomainv) ", "")
RS1.Open "select count(tov_fecemi) as suma from b_totventas where tov_fecemi=Cdate('" & Date1(0).Text & "') " & _
         "and tov_codbod=" & Val(fg_codigocbo(Combo1, 0, 10, 0)) & " and tov_tipdoc='AI' and tov_estdoc<>'A'", vg_db, adOpenStatic
RS2.Open "select max(tin_fectom) as fecha from b_tomainv where tin_codbod=" & Val(fg_codigocbo(Combo1, 0, 10, 0)), vg_db, adOpenStatic
If Not RS2.EOF And Not IsNull(RS2!fecha) Then fecha = fg_Ctod1(RS2!fecha) Else fecha = Date1(0).Text
If RS1!suma = 0 And CDate(Date1(0).Text) >= CDate(fecha) Then
    vg_db.BeginTrans
    If feccie = v_fecinv Or feccie = 0 Then
        Check1.Enabled = True
        Label2.Caption = ""
    End If
    If modo = "A" Then
        'Agrega productos si no existen en la bodega
        vg_db.Execute "insert into b_tomainv (tin_fectom, tin_codbod, tin_codpro, tin_stofis, tin_stosis, tin_propon) " & _
                       "select " & v_fecinv & ", " & v_codbod & ", pro.pro_codigo, 0, 0, pro.pro_propon from b_productos pro where pro.pro_codigo not in " & _
                       "(select tin_codpro from b_tomainv where tin_fectom=" & v_fecinv & " and tin_codbod=" & v_codbod & ")" & sqlTMP
    End If
    'Reemplazar todos los stock de sistema en 0
    vg_db.Execute "update b_tomainv set tin_stosis=0 where tin_codbod=" & v_codbod & " and tin_fectom=" & v_fecinv
    'Trae los Stock de la fecha de sistema
    vg_db.Execute "update b_tomainv a inner join b_bodegas b on a.tin_codbod=b.bod_codbod and a.tin_codpro=b.bod_codpro " & _
                  "set a.tin_stosis=b.bod_canmer where a.tin_fectom=" & v_fecinv & " and a.tin_codbod=" & v_codbod
    '--------------------------COMPRAS-------------------------
    'Crea y limpia una tabla temporal
    aAp = Trim(vg_NUsr) & "_tmp_tomainv"
    'Agrego a la tabla temporal las cantidades que aumentan (las que despues de la toma disminuyeron)
    fg_CheckTmp aAp
    vg_db.Execute "select b.dec_codmer, sum(b.dec_canmer) as dec_canmer into " & aAp & " " & _
                  "from b_totcompras a, b_detcompras b " & _
                  "where a.toc_rutpro=b.dec_rutpro and a.toc_tipdoc=b.dec_tipdoc and a.toc_numdoc=b.dec_numdoc " & _
                  "and b.dec_mueinv='S' and b.dec_tipdoc='NC' and a.toc_codbod=" & v_codbod & " " & _
                  "and a.toc_fecemi>cdate('" & Trim(Date1(0).Text) & "') group by b.dec_codmer"
    'Actualizo Stock en la toma
    vg_db.Execute "update b_tomainv a inner join " & aAp & " b on a.tin_codpro=b.dec_codmer " & _
                  "set a.tin_stosis=a.tin_stosis+b.dec_canmer " & _
                  "where a.tin_codpro=b.dec_codmer and a.tin_fectom=" & v_fecinv & " and a.tin_codbod=" & v_codbod
    'Agrego a la tabla temporal las cantidades que disminuyen (las que despues de la toma aumentaron)
    fg_CheckTmp aAp
    vg_db.Execute "select b.dec_codmer, sum(b.dec_canmer) as dec_canmer into " & aAp & " " & _
                  "from b_totcompras a, b_detcompras b " & _
                  "where a.toc_rutpro=b.dec_rutpro and a.toc_tipdoc=b.dec_tipdoc and a.toc_numdoc=b.dec_numdoc " & _
                  "and b.dec_mueinv='S' and not b.dec_tipdoc='NC' and a.toc_codbod=" & v_codbod & " " & _
                  "and a.toc_fecemi>cdate('" & Trim(Date1(0).Text) & "') group by b.dec_codmer"
    'Actualizo Stock en la toma
    vg_db.Execute "update b_tomainv a inner join " & aAp & " b on a.tin_codpro=b.dec_codmer " & _
                  "set a.tin_stosis=a.tin_stosis-b.dec_canmer " & _
                  "where a.tin_codpro=b.dec_codmer and a.tin_fectom=" & v_fecinv & " and a.tin_codbod=" & v_codbod
    '--------------------------VENTAS-------------------------
    'Agrego a la tabla temporal las cantidades que aumentan (las que despues de la toma disminuyeron)
    fg_CheckTmp aAp
    vg_db.Execute "select b.dev_codmer, sum(b.dev_canmer) as dev_canmer into " & aAp & " " & _
                  "from b_totventas a, b_detventas b " & _
                  "where a.tov_rutcli=b.dev_rutcli and a.tov_tipdoc=b.dev_tipdoc and a.tov_numdoc=b.dev_numdoc " & _
                  "and b.dev_mueinv='S' and (b.dev_tipdoc='SP' or b.dev_tipdoc='ME' or b.dev_tipdoc='FA' " & _
                  "or b.dev_tipdoc='GD' or (b.dev_tipdoc='AI' and a.tov_codreg=0) " & _
                  "or (b.dev_tipdoc='TR' and a.tov_codser=0)) and a.tov_codbod=" & v_codbod & " " & _
                  "and a.tov_fecemi>cdate('" & Trim(Date1(0).Text) & "') group by b.dev_codmer"
    'Actualizo Stock en la toma
    vg_db.Execute "update b_tomainv a inner join " & aAp & " b on a.tin_codpro=b.dev_codmer " & _
                  "set a.tin_stosis=a.tin_stosis+b.dev_canmer " & _
                  "where a.tin_codpro=b.dev_codmer and a.tin_fectom=" & v_fecinv & " and a.tin_codbod=" & v_codbod
    'Agrego a la tabla temporal las cantidades que disminuyen (las que despues de la toma aumentaron)
    fg_CheckTmp aAp
    vg_db.Execute "select b.dev_codmer, sum(b.dev_canmer) as dev_canmer into " & aAp & " " & _
                  "from b_totventas a, b_detventas b " & _
                  "where a.tov_rutcli=b.dev_rutcli and a.tov_tipdoc=b.dev_tipdoc and a.tov_numdoc=b.dev_numdoc " & _
                  "and b.dev_mueinv='S' and (b.dev_tipdoc='DP' or (b.dev_tipdoc='TR' and a.tov_codser=1) " & _
                  "or (b.dev_tipdoc='AI' and a.tov_codreg=1)) " & _
                  "and a.tov_codbod=" & v_codbod & " and a.tov_fecemi>cdate('" & Trim(Date1(0).Text) & "') group by b.dev_codmer"
    'Actualizo Stock en la toma
    vg_db.Execute "update b_tomainv a inner join " & aAp & " b on a.tin_codpro=b.dev_codmer " & _
                  "set a.tin_stosis=a.tin_stosis-b.dev_canmer " & _
                  "where a.tin_codpro=b.dev_codmer and a.tin_fectom=" & v_fecinv & " and a.tin_codbod=" & v_codbod
    '----------------------------------------------------------
    vg_db.CommitTrans
    Toolbar2.Enabled = True
    sqlPROPON = "pro.pro_propon"
End If
RS2.Close: Set RS2 = Nothing
RS1.Close: Set RS1 = Nothing
'--------------------------------------------

RS1.Open "select tin.tin_stofis, tin.tin_stosis, pro.pro_codigo, pro.pro_nombre, uni.uni_nombre, " & sqlPROPON & " " & _
         "from b_tomainv tin, b_productos pro, a_unidad uni " & _
         "where uni.uni_codigo=pro.pro_coduni and pro.pro_codigo=tin.tin_codpro " & _
         "and tin_fectom=" & v_fecinv & " and tin_codbod=" & v_codbod & sqlTMP & " order by pro.pro_nombre", vg_db, adOpenStatic
vaSpread1.MaxRows = 0
If Not RS1.EOF Then
    i = 1
    Do While Not RS1.EOF
        vaSpread1.MaxRows = i: vaSpread1.Row = i
        vaSpread1.Col = 1: vaSpread1.Text = RS1!pro_codigo
        vaSpread1.Col = 2: vaSpread1.Text = RS1!pro_nombre
        vaSpread1.Col = 3: vaSpread1.Text = RS1!uni_nombre
        vaSpread1.Col = 4: vaSpread1.Text = IIf(modo = "A", Format(0, fg_Pict(9, vg_DCa)), Format(RS1!tin_stosis, fg_Pict(9, vg_DCa)))
        vaSpread1.Col = 5: vaSpread1.Text = Format(RS1!tin_stofis, fg_Pict(9, vg_DCa))
        vaSpread1.Col = 6: vaSpread1.Text = Format(RS1(5), fg_Pict(9, vg_DPr))
        vaSpread1.Col = 7: vaSpread1.Text = Format(Format(RS1!tin_stosis, fg_Pict(9, vg_DCa)) * Format(RS1(5), fg_Pict(9, vg_DPr)), fg_Pict(9, vg_DPr))
        RS1.MoveNext: i = i + 1
    Loop
    vaSpread1.SetActiveCell 5, 1
End If
RS1.Close: Set RS1 = Nothing
'---- Reviso si hay ajuste y bloqueo --------
vaSpread1.Row = -1: vaSpread1.Col = 5
RS1.Open "select count(tov_fecemi) as suma from b_totventas where tov_fecemi=Cdate('" & Date1(0).Text & "') " & _
         "and tov_codbod=" & Val(fg_codigocbo(Combo1, 0, 10, 0)) & " and tov_tipdoc='AI' and tov_estdoc<>'A'", vg_db, adOpenStatic
RS2.Open "select max(tin_fectom) as fecha from b_tomainv where tin_codbod=" & Val(fg_codigocbo(Combo1, 0, 10, 0)), vg_db, adOpenStatic
If Not RS2.EOF Then fecha = fg_Ctod1(RS2!fecha) Else fecha = Form.Date1(0).Text
vaSpread1.Lock = IIf(RS1!suma = 0 And fecha = Date1(0).Text, False, True)
RS2.Close: Set RS2 = Nothing
RS1.Close: Set RS1 = Nothing
'--------------------------------------------
vg_codigo = ""
Gl_Ac_Botones Me, 6, 1, modo
Date1(0).Enabled = False
fg_descarga
Exit Sub
vg_db.RollbackTrans
Error_Mover:
    fg_descarga
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, Msgtitulo
    Resume Next
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long, v_fecinv As Variant, v_codbod As Long
Select Case Button.Index
Case 1
    vg_nombre = "": vg_codigo = ""
    vg_left = Toolbar2.Width
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "ProInv"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Col = 1: vaSpread1.Row = i
        If Trim(vaSpread1.Text) = Trim(vg_codigo) Then MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Next i
    v_codbod = fg_codigocbo(Combo1, 0, 10, 0)
    v_fecinv = Format(Date1(0).Text, "yyyymmdd")
    '-------------------- Reviso si hay ajuste ----------------
    vg_db.BeginTrans
    'Agrega producto si no existen en la toma
    vg_db.Execute "insert into b_tomainv (tin_fectom, tin_codbod, tin_codpro, tin_stofis, tin_stosis, tin_propon) " & _
                   "select " & v_fecinv & ", " & v_codbod & ", pro.pro_codigo, 0, 0, pro.pro_propon from b_productos pro " & _
                   "where pro.pro_codigo not in (select tin_codpro from b_tomainv where tin_fectom=" & v_fecinv & " " & _
                   "and tin_codbod=" & v_codbod & ") and pro.pro_codigo='" & Trim(vg_codigo) & "'"
    'Trae los Stock de la fecha de sistema
    vg_db.Execute "update b_tomainv a inner join b_bodegas b on a.tin_codbod=b.bod_codbod and a.tin_codpro=b.bod_codpro " & _
                  "set a.tin_stosis=b.bod_canmer where a.tin_fectom=" & v_fecinv & " and a.tin_codbod=" & v_codbod & " " & _
                  "and a.tin_codpro='" & Trim(vg_codigo) & "'"
    vg_db.CommitTrans
    '----------------------------------------------------------
    vaSpread1.Row = vaSpread1.ActiveRow
    RS1.Open "SELECT pro.pro_codigo, pro.pro_propon, pro.pro_nombre, uni.uni_nombre " & _
             "FROM b_productos AS pro, a_unidad AS uni " & _
             "WHERE pro.pro_coduni=uni.uni_codigo and pro.pro_codigo='" & vg_codigo & "'", vg_db, adOpenStatic
    If Not RS1.EOF Then
        i = vaSpread1.MaxRows + 1
        Do While Not RS1.EOF
            vaSpread1.MaxRows = i: vaSpread1.Row = i
            vaSpread1.Col = 1: vaSpread1.Text = RS1!pro_codigo
            vaSpread1.Col = 2: vaSpread1.Text = RS1!pro_nombre
            vaSpread1.Col = 3: vaSpread1.Text = RS1!uni_nombre
            vaSpread1.Col = 4: vaSpread1.Text = Format(0, fg_Pict(9, vg_DCa))
            vaSpread1.Col = 5: vaSpread1.Text = Format(0, fg_Pict(9, vg_DCa))
            vaSpread1.Col = 6: vaSpread1.Text = Format(RS1!pro_propon, fg_Pict(9, vg_DPr))
            vaSpread1.Col = 7: vaSpread1.Text = Format(0, fg_Pict(9, vg_DCa))
            RS1.MoveNext: i = i + 1
        Loop
    End If
    RS1.Close: Set RS1 = Nothing
    modo = ""
    vaSpread1.Col = 5: vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.SetActiveCell 5, vaSpread1.MaxRows
    If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
Case 2
    If vaSpread1.MaxRows = 0 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    vg_db.Execute "delete b_tomainv from b_tomainv where tin_fectom=" & Val(Format(Date1(0).Text, "yyyymmdd")) & " and tin_codbod=" & Val(fg_codigocbo(Combo1, 0, 10, "")) & " " & _
                  "and tin_codpro='" & Trim(vaSpread1.Text) & "'"
    vg_db.CommitTrans
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
    modo = ""
End Select

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If Row <> 0 Then Exit Sub
vaSpread1.Col = -1: vaSpread1.Row = -1: vaSpread1.ForeColor = RGB(0, 0, 0)
vaSpread1.Col = Col: vaSpread1.Row = -1: vaSpread1.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
vaSpread1.Col = Col: vaSpread1.Row = Row
If Round(vaSpread1.Text, vg_DCa) < 0 Then vaSpread1.Text = 0: Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 6, 0, modo
End Sub
