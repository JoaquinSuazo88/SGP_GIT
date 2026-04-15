VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form B_SalBod 
   ClientHeight    =   4815
   ClientLeft      =   2100
   ClientTop       =   2025
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   4815
      Left            =   5985
      TabIndex        =   3
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   8493
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5955
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   5
         Top             =   495
         Width           =   1470
         _Version        =   196608
         _ExtentX        =   2593
         _ExtentY        =   556
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
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
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
         ThreeDTextHighlightColor=   -2147483633
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
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483633
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   4065
         TabIndex        =   6
         Top             =   495
         Width           =   1470
         _Version        =   196608
         _ExtentX        =   2593
         _ExtentY        =   556
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
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
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
         ThreeDTextHighlightColor=   -2147483633
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
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483633
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Termino"
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
         Left            =   3285
         TabIndex        =   4
         Top             =   510
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inicio"
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
         Left            =   585
         TabIndex        =   1
         Top             =   510
         Width           =   480
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   5955
      _Version        =   393216
      _ExtentX        =   10504
      _ExtentY        =   6376
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   10
      OperationMode   =   3
      SelectBlockOptions=   0
      SpreadDesigner  =   "B_SalBod.frx":0000
      ScrollBarTrack  =   3
   End
End
Attribute VB_Name = "B_SalBod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim i As Long, ibusca As Long
Dim lc_codigo As String, lc_tipo As String
Dim icombo As Integer

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me
fg_carga (ss)
icombo = 1

'LlenaDatos
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar "
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
fpDateTime1(1).Enabled = False
icombo = 0
fg_descarga
lc_codigo = vg_codigo: vg_codigo = ""
lc_tipo = vg_nombre: vg_nombre = IIf(lc_tipo = "SP", "Salida de Bodega a Producción", _
                                 IIf(lc_tipo = "DP", "Devolución de Producción a Bodega", _
                                 IIf(lc_tipo = "ME", "Mermas", _
                                 IIf(lc_tipo = "TR", "Traspaso", _
                                 IIf(lc_tipo = "FA", "Factura - Venta directa", "Guia despacho - Venta directa")))))
Me.Caption = vg_nombre
vaSpread1.MaxRows = 0
vaSpread1.col = 3: vaSpread1.Row = 0: vaSpread1.Text = IIf(lc_tipo = "ME", "Tipo Merma", _
                                                       IIf(lc_tipo = "TR", "Tipo Traspaso", _
                                                       IIf(lc_tipo = "FA" Or lc_tipo = "GD", "Casino", "Regimen - Servicio")))
fpDateTime1(0).Text = Format(Date - 30, "dd/mm/yyyy")
fpDateTime1(1).Text = Format(Date, "dd/mm/yyyy")
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, vg_nombre
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
Dim descrip As String
If fpDateTime1(0).Text = "" And Index = 0 Then
    fpDateTime1(1).Enabled = False
    fpDateTime1(1).Text = ""
    Exit Sub
Else
    fpDateTime1(1).Enabled = True
End If
vaSpread1.MaxRows = 0
If CDate(fpDateTime1(0).Value) > CDate(fpDateTime1(1).Value) Then fpDateTime1(1).Text = fpDateTime1(0).Text: Exit Sub
If fpDateTime1(1).Text = "" Then Exit Sub
descrip = ""
If lc_tipo = "TR" Then
    RS1.Open "select distinct tov_numdoc, tov_fecemi, tov_codser, tov_estdoc, tov_tipdoc " & _
             "From b_totventas " & _
             "where tov_rutcli='" & lc_codigo & "' " & _
             "and tov_tipdoc='" & lc_tipo & "' " & _
             "and tov_fecemi>=CDate('" & fpDateTime1(0).Text & "') " & _
             "and tov_fecemi<=CDate('" & fpDateTime1(1).Text & "') " & _
             "order by tov_fecemi", vg_db, adOpenStatic
ElseIf lc_tipo = "FA" Or lc_tipo = "GD" Then
    RS1.Open "select distinct tov.tov_numdoc, tov.tov_fecemi, cli.cli_nombre, tov.tov_estdoc, tov.tov_tipdoc " & _
             "From b_totventas tov, b_clientes cli " & _
             "where tov.tov_rutcli=cli.cli_codigo " & _
             "and tov.tov_tipdoc='" & lc_tipo & "' " & _
             "and tov.tov_fecemi>=CDate('" & fpDateTime1(0).Text & "') " & _
             "and tov.tov_fecemi<=CDate('" & fpDateTime1(1).Text & "') " & _
             "order by tov.tov_fecemi", vg_db, adOpenStatic
ElseIf lc_tipo = "ME" Then
    RS1.Open "select distinct tov.tov_numdoc, tov.tov_fecemi, aju.aju_nombre, tov.tov_estdoc, tov.tov_tipdoc " & _
             "From b_totventas tov, a_tipoajuste aju " & _
             "where tov.tov_rutcli='" & lc_codigo & "' and tov.tov_tipdoc='" & lc_tipo & "' " & _
             "and tov.tov_codser=aju.aju_codigo and tov.tov_fecemi>=CDate('" & fpDateTime1(0).Text & "') " & _
             "and tov.tov_fecemi<=CDate('" & fpDateTime1(1).Text & "') order by tov.tov_fecemi", vg_db, adOpenStatic
ElseIf lc_tipo = "SP" Or lc_tipo = "DP" Then
    RS1.Open "select distinct tov.tov_numdoc, tov.tov_fecpro, ser.ser_nombre, tov.tov_estdoc, reg.reg_nombre " & _
             "From b_totventas tov, a_regimen reg, a_servicio ser " & _
             "where tov.tov_rutcli='" & lc_codigo & "' and tov.tov_tipdoc='" & lc_tipo & "' " & _
             "and tov.tov_codser=ser.ser_codigo and tov.tov_codreg=reg.reg_codigo " & _
             "and tov.tov_fecpro>=CDate('" & fpDateTime1(0).Text & "') " & _
             "and tov.tov_fecpro<=CDate('" & fpDateTime1(1).Text & "') order by tov.tov_fecpro", vg_db, adOpenStatic
End If
If Not RS1.EOF Then
    Do While Not RS1.EOF
        descrip = IIf(lc_tipo = "TR", IIf(RS1(2) = 1, "Recibido", "Entregado"), _
                  IIf(lc_tipo = "SP" Or lc_tipo = "DP", RS1(4) & " - " & RS1(2), RS1(2)))
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.col = 1: vaSpread1.Text = RS1(0)
        vaSpread1.col = 2: vaSpread1.Text = Format(RS1(1), "dd/mm/yyyy")
        vaSpread1.col = 3: vaSpread1.Text = descrip
        vaSpread1.col = 4: vaSpread1.Text = IIf(RS1(3) = "", "ACTIVADA", "ANULADA")
        RS1.MoveNext
    Loop
End If
RS1.Close: Set RS1 = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    MoverDatos
Case 3
    Cerrar
End Select
End Sub

Private Sub vaSpread1_DblClick(ByVal col As Long, ByVal Row As Long)
MoverDatos
End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
MoverDatos
End Sub

Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 27
    Cerrar
End Select
End Sub

Private Sub MoverDatos()
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.col = 1
vg_codigo = Trim(vaSpread1.Text)
Cerrar
End Sub

Sub Cerrar()
Me.Hide
Unload Me
End Sub
