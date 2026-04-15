VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_PVtaCl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precio Venta Cliente"
   ClientHeight    =   5820
   ClientLeft      =   3180
   ClientTop       =   3015
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1965
      Index           =   1
      Left            =   30
      TabIndex        =   5
      Top             =   390
      Width           =   7335
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1665
         TabIndex        =   1
         Top             =   720
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "-2147483647"
         NegFormat       =   0
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   2
         Left            =   1665
         TabIndex        =   2
         Top             =   1080
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   2
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "-2147483647"
         NegFormat       =   0
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1650
         TabIndex        =   0
         Top             =   360
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   10
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   1
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1665
         TabIndex        =   3
         Top             =   1440
         Width           =   1260
         _Version        =   196608
         _ExtentX        =   2222
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
         AllowNull       =   0   'False
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
         Text            =   "06/08/2004"
         DateCalcMethod  =   4
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio de Validez"
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
         Left            =   135
         TabIndex        =   13
         Top             =   1515
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
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
         Index           =   3
         Left            =   135
         TabIndex        =   12
         Top             =   1185
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
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
         Index           =   2
         Left            =   135
         TabIndex        =   11
         Top             =   825
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Casino"
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
         Left            =   135
         TabIndex        =   10
         Top             =   480
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2895
         Picture         =   "M_PVtaCl.frx":0000
         Top             =   285
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2895
         Picture         =   "M_PVtaCl.frx":030A
         Top             =   645
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2895
         Picture         =   "M_PVtaCl.frx":0614
         Top             =   1005
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2895
         Picture         =   "M_PVtaCl.frx":091E
         Top             =   1365
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3330
         TabIndex        =   9
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3330
         TabIndex        =   8
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3330
         TabIndex        =   7
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3330
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3375
         TabIndex        =   14
         Top             =   405
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3375
         TabIndex        =   15
         Top             =   765
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3375
         TabIndex        =   16
         Top             =   1125
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3375
         TabIndex        =   17
         Top             =   1485
         Width           =   1815
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3375
      Left            =   30
      TabIndex        =   4
      Top             =   2430
      Width           =   7335
      _Version        =   393216
      _ExtentX        =   12938
      _ExtentY        =   5953
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
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
      MaxCols         =   4
      MaxRows         =   1
      SpreadDesigner  =   "M_PVtaCl.frx":0C28
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_PVtaCl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim i As Long
Dim Msgtitulo As String
Dim modo As String, codcli As String
Dim Est As Boolean, OpGr As Boolean, accion As Boolean

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 6300
Me.Width = 7470
Msgtitulo = "Precio Venta Cliente"
fg_centra Me
modo = ""
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 3, modo
accion = True
OpGr = False: vaSpread1.MaxRows = 0
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.Text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
fpDateTime1.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub Form_Resize()
Frame1(1).Move IIf(Me.WindowState = 2, 4200, 0), 360, 7335, 1871
vaSpread1.Move 0, 2440, ScaleWidth, ScaleHeight - 2440
Toolbar1.Refresh
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpDateTime1_LostFocus()
fpayuda(3).Caption = fg_Fecha_Dia(Format(fpDateTime1.Text, "yyyymmdd"), 2)
MoverDatos
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    If Index = 1 Then image1_Click 1
    If Index = 2 Then image1_Click 2
End Select
End Sub

Private Sub fpLongInteger1_LostFocus(Index As Integer)
Select Case Index
Case 1
    If Val(fpLongInteger1(1).Value) < 1 Then fpayuda(1).Caption = "": Exit Sub
    RS.Open "select * from a_regimen where reg_codigo=" & Val(fpLongInteger1(1).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(1).Text = "": fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    MoverDatos
Case 2
    If Val(fpLongInteger1(2).Value) < 1 Then fpayuda(2).Caption = "": Exit Sub
    RS.Open "select * from a_servicio where ser_codigo=" & Val(fpLongInteger1(2).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(2).Text = "": fpayuda(2).Caption = "": Exit Sub
    fpayuda(2).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    MoverDatos
End Select
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    image1_Click 0
End Select
End Sub

Private Sub fpText_LostFocus()
If fpText.Text = "" Then fpayuda(0).Caption = "": Exit Sub
RS.Open "select * from b_clientes where cli_codigo='" & fpText.Text & "' and cli_tipo=0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": fpLongInteger1(2).Value = "": fpayuda(3).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
fpLongInteger1(2).Value = "": fpayuda(2).Caption = ""
MoverDatos
End Sub

Private Sub image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casinos", "Casino"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.Text = vg_codigo
    fpayuda(0).Caption = vg_nombre
    fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(2).Value = "": fpayuda(2).Caption = ""
    fpLongInteger1(1).SetFocus
    MoverDatos
Case 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    fpLongInteger1(2).SetFocus
    MoverDatos
Case 2
    vg_left = fpayuda(2).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre
    fpDateTime1.SetFocus
    MoverDatos
Case 3
    If fpText.Text = "" Or Val(fpLongInteger1(1).Value) < 1 Or Val(fpLongInteger1(2).Value) < 0 Or fpDateTime1.Text = "" Then Exit Sub
    B_HistPm.LlenarHistPlan "Histórico Precio Ventas Clientes", fpText.Text, fpLongInteger1(1).Text & "|" & fpLongInteger1(2).Text & "|", 5
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    fpDateTime1.Text = vg_codigo
    MoverDatos
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codigo As Long, Nombre As String, Orden As String
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 1: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows ': vaSpread1.SetFocus
Case 3
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
Case 5
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codcli = fg_DespintaRut(vaSpread1.Text)
    vg_db.BeginTrans
    vg_db.Execute "delete b_preciovta from b_preciovta where prv_cencos='" & fpText.Text & "' and prv_codreg=" & Val(fpLongInteger1(1).Value) & " and prv_codser=" & Val(fpLongInteger1(2).Value) & " and prv_fecvig=" & Format(fpDateTime1.Text, "yyyymmdd") & " and prv_rutcli='" & codcli & "'"
    vg_db.CommitTrans
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    modo = "": Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
Case 7
    MoverDatos
Case 10
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    If modo = "A" Then
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    Else
       Cancela
    End If
    modo = "": Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
Case 12
    GrabaRegistro vaSpread1.ActiveRow
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_PrecioVentaCliente Trim(fpText.Text), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Format(fpDateTime1.Text, "yyyymmdd")
Case 18
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
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
Select Case Col
Case 1
    '------- Validar si existe clientes
    vaSpread1.Row = Row
    vaSpread1.Col = 4: codcli = ""
    If vaSpread1.Text <> "" Then codcli = vaSpread1.Text: vaSpread1.Col = 1: vaSpread1.Text = codcli: Exit Sub
    vaSpread1.Row = Row: vaSpread1.Col = 1: If vaSpread1.Lock = True Then Exit Sub
    '------- Llama  a formulario de busqueda de clientes y carga datos
    vaSpread1.Col = 1
    vg_left = fpayuda(2).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Clientes", "Cliente"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If Trim(vg_codigo) = Trim(vaSpread1.Text) And Row <> i And Trim(vaSpread1.Text) <> "" Then MsgBox "Cliente existe", vbCritical + vbOKOnly, Msgtitulo: vaSpread1.Text = "": vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
    Next i
    vg_codigo = fg_DespintaRut(vg_codigo)
    RS.Open "select cli_codigo, cli_nombre from b_clientes where cli_codigo='" & vg_codigo & "' and cli_tipo=1", vg_db, adOpenStatic
    If RS.EOF Then: RS.Close: Set RS = Nothing: vaSpread1.Text = "": vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
    vaSpread1.Col = 1: vaSpread1.Text = fg_PintaRut(RS!cli_codigo)
    vaSpread1.Col = 2: vaSpread1.Text = RS!cli_nombre
    RS.Close: Set RS = Nothing
    vaSpread1.SetActiveCell 3, vaSpread1.ActiveRow
End Select
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
Select Case Col
Case 1
    vaSpread1.Row = Row
    vaSpread1.Col = 4: codcli = ""
    If vaSpread1.Text <> "" Then codcli = vaSpread1.Text: vaSpread1.Col = 1: vaSpread1.Text = codcli: Exit Sub
Case 3
    If modo = "" Then modo = "M"
    If ChangeMade = True Then Gl_Ac_Botones Me, 1, 0, modo
End Select
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
Dim precio As Double
Select Case Col
Case 1
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    vaSpread1.Col = 4: codcli = ""
    If vaSpread1.Text <> "" Then codcli = vaSpread1.Text: vaSpread1.Col = 1: vaSpread1.Text = codcli: Exit Sub
    vaSpread1.Col = 1
    vaSpread1.Text = UCase(vaSpread1.Text)
    vaSpread1.Text = fg_RutDig(Trim(vaSpread1.Text))
    vaSpread1.Text = fg_PintaRut(vaSpread1.Text)
    
    RS.Open "select cli_codigo, cli_nombre from b_clientes where cli_codigo='" & fg_DespintaRut(vaSpread1.Text) & "' and cli_tipo=1", vg_db, adOpenStatic
    If RS.EOF Then: RS.Close: Set RS = Nothing: vaSpread1.Text = "": vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
    codcli = vaSpread1.Text
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        If Trim(codcli) = Trim(vaSpread1.Text) And Row <> i And Trim(vaSpread1.Text) <> "" Then RS.Close: Set RS = Nothing: vaSpread1.Row = Row: vaSpread1.Text = "": MsgBox "Productos existe", vbCritical + vbOKOnly, Msgtitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
    Next i
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    vaSpread1.Col = 1: vaSpread1.Text = fg_PintaRut(RS!cli_codigo)
    vaSpread1.Col = 2: vaSpread1.Text = RS!cli_nombre
    RS.Close: Set RS = Nothing
    vaSpread1.SetActiveCell 3, vaSpread1.ActiveRow
    vaSpread1.Row = NewRow
    vaSpread1.Col = 1: codcli = fg_PintaRut(vaSpread1.Text)
    vaSpread1.Col = 3: precio = Val(vaSpread1.Text)
Case 2, 3, 4
    If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
       GrabaRegistro Row
    ElseIf Toolbar1.Buttons(12).Visible = False Then
'       Cancela
    End If
End Select
End Sub

Sub MoverDatos()
vaSpread1.MaxRows = 0
If fpText.Text = "" Or Val(fpLongInteger1(1).Value) = 0 Or Val(fpLongInteger1(2).Value) = 0 Or fpDateTime1.Text = "" Then Exit Sub
RS.Open "select b_clientes.cli_codigo, b_clientes.cli_nombre, b_preciovta.prv_preven " & _
        "from  b_clientes, b_preciovta " & _
        "where b_preciovta.prv_rutcli=b_clientes.cli_codigo " & _
        "and   b_preciovta.prv_cencos='" & LimpiaDato(Trim(fpText.Text)) & "' " & _
        "and   b_preciovta.prv_codreg=" & Val(fpLongInteger1(1).Value) & " " & _
        "and   b_preciovta.prv_codser=" & Val(fpLongInteger1(2).Value) & " " & _
        "and   b_preciovta.prv_fecvig=" & Format(fpDateTime1.Text, "yyyymmdd") & " " & _
        "and   b_clientes.cli_tipo=1", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 1: vaSpread1.Text = fg_PintaRut(RS!cli_codigo)
      vaSpread1.Col = 2: vaSpread1.Text = RS!cli_nombre
      vaSpread1.Col = 3: vaSpread1.Text = RS!prv_preven
      vaSpread1.Col = 4: vaSpread1.Text = fg_PintaRut(RS!cli_codigo)
      RS.MoveNext
   Loop
   Gl_Ac_Botones Me, 1, 1, modo
   vaSpread1.SetActiveCell 1, 1
Else
   If fpText.Text = "" Or Val(fpLongInteger1(1).Value) < 1 Or Val(fpLongInteger1(2).Value) < 0 Or fpDateTime1.Text = "" Then
      Gl_Ac_Botones Me, 1, 3, modo
   Else
      Gl_Ac_Botones Me, 1, 2, modo
   End If
End If
RS.Close: Set RS = Nothing
End Sub

Private Sub GrabaRegistro(Fila As Long)
Dim precio As Double
OpGr = True
vaSpread1.Row = Fila
codcli = "": precio = 0
vaSpread1.Col = 1: codcli = fg_DespintaRut(vaSpread1.Text)
vaSpread1.Col = 3: precio = Val(vaSpread1.Text)
If Trim(codcli) = "" Or precio < 1 Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 3, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub
If modo = "A" Then
    vg_db.BeginTrans
      vg_db.Execute "insert into b_preciovta (prv_cencos, prv_codreg, prv_codser, prv_fecvig, prv_rutcli, prv_preven) " & _
                    "values ('" & fpText.Text & "', " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & Val(Format(fpDateTime1.Text, "yyyymmdd")) & ", '" & codcli & "', " & precio & ")"
    vg_db.CommitTrans
    vaSpread1.Col = 1: vaSpread1.Text = fg_PintaRut(codcli)
    vaSpread1.Col = 4: vaSpread1.Text = fg_PintaRut(codcli)
Else
    vg_db.BeginTrans
    vg_db.Execute "update b_preciovta SET prv_preven=" & precio & " where prv_cencos='" & fpText.Text & "' and prv_codreg=" & Val(fpLongInteger1(1).Value) & " and prv_codser=" & Val(fpLongInteger1(2).Value) & " and prv_fecvig=" & Format(fpDateTime1.Text, "yyyymmdd") & " and prv_rutcli='" & codcli & "'"
    vg_db.CommitTrans
End If
modo = "": Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
OpGr = False

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
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Cancela()
OpGr = True
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: codcli = fg_DespintaRut(vaSpread1.Text)
RS.Open "select b_clientes.cli_codigo, b_clientes.cli_nombre, b_preciovta.prv_preven " & _
        "from   b_clientes, b_preciovta " & _
        "where  b_preciovta.prv_rutcli=b_clientes.cli_codigo " & _
        "and    b_preciovta.prv_cencos='" & LimpiaDato(Trim(fpText.Text)) & "' " & _
        "and    b_preciovta.prv_codreg=" & Val(fpLongInteger1(1).Value) & " " & _
        "and    b_preciovta.prv_codser=" & Val(fpLongInteger1(2).Value) & " " & _
        "and    b_preciovta.prv_fecvig=" & Format(fpDateTime1.Text, "yyyymmdd") & " " & _
        "and    b_preciovta.prv_rutcli='" & codcli & "' " & _
        "and    b_clientes.cli_tipo=1", vg_db, adOpenStatic
If Not RS.EOF Then
   vaSpread1.Col = 2: vaSpread1.Text = Trim(RS!cli_nombre)
   vaSpread1.Col = 3: vaSpread1.Text = RS!prv_preven
End If
RS.Close: Set RS = Nothing
OpGr = False
End Sub
