VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_ConRac 
   Caption         =   "Control de Raciones"
   ClientHeight    =   6135
   ClientLeft      =   255
   ClientTop       =   1500
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Datos Raciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   30
      TabIndex        =   6
      Top             =   2310
      Width           =   11715
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2880
         Left            =   150
         TabIndex        =   7
         Top             =   270
         Width           =   11415
         _Version        =   393216
         _ExtentX        =   20135
         _ExtentY        =   5080
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   2
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
         MaxCols         =   33
         MaxRows         =   1
         SpreadDesigner  =   "M_ConRac.frx":0000
         ScrollBarTrack  =   3
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Clientes"
         Height          =   195
         Index           =   2
         Left            =   3510
         TabIndex        =   10
         Top             =   3390
         Width           =   555
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   3120
         Top             =   3420
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dias Bloqueados"
         Height          =   195
         Index           =   1
         Left            =   4965
         TabIndex        =   9
         Top             =   3375
         Width           =   1200
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008484FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   4605
         Top             =   3405
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dias Habilitados"
         Height          =   195
         Index           =   0
         Left            =   6810
         TabIndex        =   8
         Top             =   3375
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   6450
         Top             =   3405
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   60
      TabIndex        =   5
      Top             =   420
      Width           =   11685
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   3180
         TabIndex        =   1
         Top             =   615
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
         Left            =   3180
         TabIndex        =   2
         Top             =   975
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
         Left            =   3165
         TabIndex        =   0
         Top             =   255
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
         Left            =   3180
         TabIndex        =   3
         Top             =   1335
         Width           =   1050
         _Version        =   196608
         _ExtentX        =   1852
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
         Text            =   "10/2004"
         DateCalcMethod  =   4
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/yyyy"
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
         Caption         =   "Fecha Minuta"
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
         Left            =   1770
         TabIndex        =   17
         Top             =   1410
         Width           =   1170
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
         Left            =   1770
         TabIndex        =   16
         Top             =   1080
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
         Left            =   1770
         TabIndex        =   15
         Top             =   720
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
         Left            =   1770
         TabIndex        =   14
         Top             =   375
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   4410
         Picture         =   "M_ConRac.frx":0FDC
         Top             =   180
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   4410
         Picture         =   "M_ConRac.frx":12E6
         Top             =   540
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   4410
         Picture         =   "M_ConRac.frx":15F0
         Top             =   900
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   4845
         TabIndex        =   13
         Top             =   255
         Width           =   3855
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   4845
         TabIndex        =   12
         Top             =   615
         Width           =   3855
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   4845
         TabIndex        =   11
         Top             =   975
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   4890
         TabIndex        =   18
         Top             =   300
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   4890
         TabIndex        =   19
         Top             =   660
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   4890
         TabIndex        =   20
         Top             =   1020
         Width           =   3855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_ConRac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim Msgtitulo As String, modo As String
Dim i As Long, X As Long, v_columnas As Long

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 6630
Me.Width = 11895
fg_centra Me
Msgtitulo = "Control de Raciones"
Me.HelpContextID = vg_OpcM
modo = "": vaSpread1.MaxRows = 0
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 3, 1), modo
OpGr = False: vaSpread1.MaxRows = 0
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.Text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
fpDateTime1.Text = Format(Date, "mm/yyyy")
GenerarTitulo
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpDateTime1_LostFocus()
MoverDatos
End Sub

Private Sub fpLongInteger1_Click(Index As Integer, Button As Integer)
If vaSpread1.MaxRows > 0 Then vaSpread1.MaxRows = 0: Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 3, 1), modo
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    If Index = 1 Then Image1_Click 1
    If Index = 2 Then Image1_Click 2
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
'    MoverDatos
  Case 2
    If Val(fpLongInteger1(2).Value) < 1 Then fpayuda(2).Caption = "": Exit Sub
    RS.Open "select * from a_servicio where ser_codigo=" & Val(fpLongInteger1(2).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(2).Text = "": fpayuda(2).Caption = "": Exit Sub
    fpayuda(2).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
'    MoverDatos
End Select
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    Image1_Click 0
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
'MoverDatos
End Sub

Private Sub Image1_Click(Index As Integer)
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
'    MoverDatos
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
'    MoverDatos
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
'    MoverDatos
Case 3
    If fpText.Text = "" Or Val(fpLongInteger1(1).Value) < 1 Or Val(fpLongInteger1(2).Value) < 0 Or fpDateTime1.Text = "" Then Exit Sub
    B_HistPm.LlenarHistPlan "Histórico Estructura Fija", fpText.Text, fpLongInteger1(1).Text & "|" & fpLongInteger1(2).Text & "|", 3
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    fpDateTime1.Text = vg_codigo
'    accion = False: Combo1(0).ListIndex = vg_auxfecha - 1: accion = True
    MoverDatos
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 1 '------- Agregar registro
    GenerarTitulo
    Dim auxrutcli As String
    auxrutcli = ""
    RS.Open "select cli_codigo, cli_nombre from b_clientes where  cli_tipo=1 order by cli_nombre", vg_db, adOpenStatic
    If Not RS.EOF Then
       Do While Not RS.EOF
          If RS!cli_codigo <> auxrutcli Then vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 1: vaSpread1.Text = fg_PintaRut(RS!cli_codigo): vaSpread1.Col = 2: vaSpread1.Text = RS!cli_nombre: auxrutcli = RS!cli_codigo
          RS.MoveNext
       Loop
    End If
    RS.Close: Set RS = Nothing
   
   '------- Agregar fila totales
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 2: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.Text = "Total Cliente"
   vaSpread1.Col = -1: vaSpread1.BackColor = &HE0E0E0
   For i = 3 To vaSpread1.MaxCols
       vaSpread1.Col = i: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = 1
   Next i
   '------- Fin agregar fila totales
   
   '------- Agregar fila personal
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 1: vaSpread1.Text = "PERSONAL": vaSpread1.Col = 2: vaSpread1.Text = "PERSONAL"
   '------- Fin Agregar fila personal
   
   '-------Bloquea días de cierre en color rojo
   'Dim diablq As Date
   'If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
   'v_columnas = ((CDate(diablq) - CDate(CDate("01/" + Mid(fpDateTime1.Text, 1, 8))) + 1) * 2) - 1
   'If v_columnas > 0 Then
   '   vaSpread1.Row = -1
   '   For i = 3 To v_columnas + 2
   '       vaSpread1.Col = i
   '       vaSpread1.Lock = True
   '       vaSpread1.BackColor = Shape1(0).FillColor
   '   Next i
   '   vaSpread1.SetActiveCell i, 1
   'End If
   '-------Fin Bloqueo de celdas
   '------- Bloquea días de cierre en color rojo
   Dim diablq As Date
   If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
   If Format(Now, "dd/mm/yyyy") > diablq Or Format(CDate(fpDateTime1.Text), "mm/yyyy") < Format(Month(Now) - 1 & "/" & Year(Now), "mm/yyyy") Then
        v_columnas = ((dEoM(Format(Day(Now) & "/" & (Month(Now) - 1) & "/" & Year(Now), "dd/mm/yyyy")) - CDate(CDate("01/" + Mid(fpDateTime1.Text, 1, 8))) + 1) * 2) - 1
   Else
        v_columnas = 0
   End If
        
   If v_columnas > 0 Then
      vaSpread1.Row = -1
      For i = 3 To v_columnas + 2
          vaSpread1.Col = i
          vaSpread1.Lock = True
          vaSpread1.BackColor = Shape1(0).FillColor
      Next i
      vaSpread1.SetActiveCell i, 1
   End If
   '------- Fin Bloqueo de celdas
    
    vaSpread1.Row = 1: vaSpread1.Col = vaSpread1.MaxCols
    If vaSpread1.BackColor <> Shape1(0).FillColor Then modo = "A": Gl_Ac_Botones Me, 1, 0, modo
    vaSpread1.SetActiveCell 1, 1
    If v_columnas < vaSpread1.MaxCols Then fpLongInteger1(1).Enabled = False: Image1(1).Enabled = False: fpLongInteger1(2).Enabled = False: fpDateTime1.Enabled = False: Image1(2).Enabled = False
Case 3 '------- Activar modo modificación
    If vaSpread1.MaxRows < 1 Then Exit Sub
    vaSpread1.Row = 1: vaSpread1.Col = vaSpread1.MaxCols
    If vaSpread1.BackColor <> Shape1(0).FillColor Then modo = "M": Gl_Ac_Botones Me, 1, 0, modo Else Exit Sub
    fpLongInteger1(1).Enabled = False: Image1(1).Enabled = False: fpLongInteger1(2).Enabled = False: fpDateTime1.Enabled = False: Image1(2).Enabled = False
Case 5 '------- Borrar información
    If vaSpread1.ActiveRow < 1 Then MsgBox "No existe información a borrar...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina Documento...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    vg_db.Execute "delete * from b_minutaraciones where mir_cencos='" & fpText.Text & "' and mir_codreg=" & Val(fpLongInteger1(1).Value) & " and mir_codser=" & Val(fpLongInteger1(2).Value) & " and val(mid(mir_fecmin,1,6))=" & Val(Format(fpDateTime1.Text, "yyyymm")) & ""
    vg_db.CommitTrans
    vaSpread1.MaxRows = 0
    modo = "": Gl_Ac_Botones Me, 1, 3, modo
Case 7 '------- Actualizar lista
    MoverDatos
Case 10 '------- Cancelar
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    If modo = "A" Then
        vaSpread1.MaxRows = 0
    Else
       MoverDatos
    End If
    modo = "": Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 4), modo
    fpLongInteger1(1).Enabled = True: Image1(1).Enabled = True: fpLongInteger1(2).Enabled = True: fpDateTime1.Enabled = True: Image1(2).Enabled = True
Case 12 '------- Grabar información
    fg_carga ""
    Dim rutcli As String
    Dim nrorac As Long
    rutcli = ""
    If fpText.Text = "" Or Val(fpLongInteger1(1).Value) < 1 Or Val(fpLongInteger1(2).Value) < 1 Or fpDateTime1.Text = "" Then Exit Sub
    If modo = "A" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i: vaSpread1.Col = 1: rutcli = Trim(fg_DespintaRut(vaSpread1.Text))
           If rutcli <> "" Then
              For X = 3 To vaSpread1.MaxCols
                  vaSpread1.Row = i: vaSpread1.Col = X: nrorac = 0
                  If Trim(vaSpread1.Text) <> "" Then
                     nrorac = Val(vaSpread1.Text)
                     vaSpread1.Row = 0: vaSpread1.Col = X
                     vg_db.BeginTrans
                     vg_db.Execute "insert into b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac) select trim(min_cencos), min_codreg, min_codser, min_fecmin, '" & Trim(rutcli) & "', " & nrorac & " from b_minuta where min_codigo in (select mid_codigo from b_minutadet where mid_tipmin='2') and min_cencos='" & fpText.Text & "' and min_codreg=" & Val(fpLongInteger1(1).Value) & " and min_codser = " & Val(fpLongInteger1(2).Value) & " and min_fecmin=" & Format(Format(fpDateTime1.Text, "yyyy/mm") & "/" & fg_pone_cero(Str(Right(vaSpread1.Text, 2)), 2), "yyyymmdd") & ""
                     vg_db.CommitTrans
                  End If
              Next X
           End If
       Next i
       modo = "M"
    Else
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i: vaSpread1.Col = 1: rutcli = Trim(fg_DespintaRut(vaSpread1.Text))
           If rutcli <> "" Then
              For X = 3 To vaSpread1.MaxCols
                  vaSpread1.Row = i: vaSpread1.Col = X: nrorac = 0
                  nrorac = Val(vaSpread1.Text)
                  vaSpread1.Row = 0: vaSpread1.Col = X
                  RS.Open "select b_minutaraciones.mir_nrorac " & _
                          "from   b_clientes, b_minutaraciones " & _
                          "where  b_minutaraciones.mir_cencos='" & LimpiaDato(Trim(fpText.Text)) & "' " & _
                          "and    b_minutaraciones.mir_codreg=" & Val(fpLongInteger1(1).Value) & " " & _
                          "and    b_minutaraciones.mir_codser=" & Val(fpLongInteger1(2).Value) & " " & _
                          "and    b_minutaraciones.mir_rutcli='" & rutcli & "' " & _
                          "and    b_minutaraciones.mir_fecmin=" & Format(Format(fpDateTime1.Text, "yyyy/mm") & "/" & fg_pone_cero(Str(Right(vaSpread1.Text, 2)), 2), "yyyymmdd") & "", vg_db, adOpenStatic
                  If RS.EOF And nrorac > 0 Then
                     vg_db.BeginTrans
                     vg_db.Execute "insert into b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac) select trim(min_cencos), min_codreg, min_codser, min_fecmin, '" & Trim(rutcli) & "', " & nrorac & " from b_minuta where min_codigo in (select mid_codigo from b_minutadet where mid_tipmin='2') and min_cencos='" & fpText.Text & "' and min_codreg=" & Val(fpLongInteger1(1).Value) & " and min_codser = " & Val(fpLongInteger1(2).Value) & " and min_fecmin=" & Format(Format(fpDateTime1.Text, "yyyy/mm") & "/" & fg_pone_cero(Str(Right(vaSpread1.Text, 2)), 2), "yyyymmdd") & ""
                     vg_db.CommitTrans
                  ElseIf Not RS.EOF Then
                     If RS!mir_nrorac <> nrorac Then
                        vg_db.BeginTrans
                        vg_db.Execute "update b_minutaraciones set mir_nrorac=" & nrorac & " where mir_cencos='" & fpText.Text & "' and mir_codreg=" & Val(fpLongInteger1(1).Value) & " and mir_codser=" & Val(fpLongInteger1(2).Value) & " and mir_fecmin=" & Format(Format(fpDateTime1.Text, "yyyy/mm") & "/" & fg_pone_cero(Str(Right(vaSpread1.Text, 2)), 2), "yyyymmdd") & " and mir_rutcli='" & rutcli & "'"
                        vg_db.CommitTrans
                     End If
                  End If
                  RS.Close: Set RS = Nothing
              Next X
           End If
       Next i
       modo = "M"
    End If
    modo = "": Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 4), modo
    fpLongInteger1(1).Enabled = True: Image1(1).Enabled = True: fpLongInteger1(2).Enabled = True: fpDateTime1.Enabled = True: Image1(2).Enabled = True
    fg_descarga
Case 15 '------- Imprimir
    If vaSpread1.MaxRows < 1 Then Exit Sub
    I_ConRac fpText.Text, Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), fpDateTime1.Text
Case 18 '------- Salir
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

Sub GenerarTitulo()
vaSpread1.MaxRows = 0
vaSpread1.MaxCols = 2 + Left(dEoM("01/" & fpDateTime1.Text), 2)
X = 1
For i = 3 To vaSpread1.MaxCols
    vaSpread1.Row = 0: vaSpread1.Col = i: vaSpread1.Text = fg_Fecha_Dia(Format(Mid(fpDateTime1.Text, 4, 4) & "/" & Mid(fpDateTime1.Text, 1, 2) & "/" & fg_pone_cero(Str(X), 2), "yyyymmdd"), 2): X = X + 1
Next i
vaSpread1.Col = -1: vaSpread1.Row = -1
vaSpread1.BackColor = Shape1(1).FillColor
vaSpread1.Lock = False
vaSpread1.Row = -1
vaSpread1.Col = 1: vaSpread1.BackColor = Shape1(2).FillColor: vaSpread1.Col = 2: vaSpread1.BackColor = Shape1(2).FillColor
End Sub
Sub MoverDatos()
Dim indper As Boolean
If fpText.Text = "" Or Val(fpLongInteger1(1).Value) < 1 Or Val(fpLongInteger1(2).Value) < 1 Or fpDateTime1.Text = "" Then Exit Sub
vaSpread1.MaxRows = 0: indper = False
Dim auxrutcli As String
'----------- Validar si existe planificación real
RS.Open "select count(b_minutadet.mid_codigo) AS nreg From b_minutadet, b_minuta " & _
        "where b_minuta.min_codigo=b_minutadet.mid_codigo " & _
        "and   b_minuta.min_cencos='" & fpText.Text & "' " & _
        "and   b_minuta.min_codreg=" & Val(fpLongInteger1(1).Value) & " " & _
        "and   b_minuta.min_codser=" & Val(fpLongInteger1(2).Value) & " " & _
        "and val(mid(b_minuta.min_fecmin,1,6))=" & Format(fpDateTime1.Text, "yyyymm") & " and b_minutadet.mid_tipmin='2'", vg_db, adOpenStatic
If RS!NReg = 0 Then RS.Close: Set RS = Nothing: MsgBox "No existe planificación real, para este mes...", vbExclamation + vbOKOnly, Msgtitulo: fpLongInteger1(1).Value = "": fpLongInteger1(2).Value = "": Exit Sub
RS.Close: Set RS = Nothing
'----------- Fin validar si existe planificación real

GenerarTitulo
auxrutcli = ""
RS.Open "select b_minutaraciones.mir_fecmin, b_minutaraciones.mir_rutcli, b_minutaraciones.mir_nrorac " & _
        "from   b_minutaraciones " & _
        "where  b_minutaraciones.mir_cencos='" & LimpiaDato(Trim(fpText.Text)) & "' " & _
        "and    b_minutaraciones.mir_codreg=" & Val(fpLongInteger1(1).Value) & " " & _
        "and    b_minutaraciones.mir_codser=" & Val(fpLongInteger1(2).Value) & " " & _
        "and    val(mid(b_minutaraciones.mir_fecmin,1,6))=" & Val(Format(fpDateTime1.Text, "yyyymm")) & "", vg_db, adOpenStatic
If Not RS.EOF Then
    RS1.Open "select cli_codigo, cli_nombre from b_clientes where  cli_tipo=1 order by cli_nombre", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: RS.Close: Set RS = Nothing: Exit Sub
    Do While Not RS1.EOF
       If RS1!cli_codigo <> auxrutcli Then
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 1: vaSpread1.Text = fg_PintaRut(RS1!cli_codigo): vaSpread1.Col = 2: vaSpread1.Text = RS1!cli_nombre: auxrutcli = RS1!cli_codigo
          Do While Not RS.EOF
             If RS1!cli_codigo = Trim(RS!mir_rutcli) Then vaSpread1.Col = 2 + Mid(RS!mir_fecmin, 7, 2): vaSpread1.Text = RS!mir_nrorac
             RS.MoveNext
          Loop
          RS.MoveFirst
       End If
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
   '------- Agregar fila totales
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 2: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.Text = "Total Cliente"
   vaSpread1.Col = -1: vaSpread1.BackColor = &HE0E0E0
   For i = 3 To vaSpread1.MaxCols
       vaSpread1.Col = i: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = 1
   Next i
   '------- Fin agregar fila totales
   
   '------- Sumar totales
   Dim totnrorac As Long
   For i = 3 To vaSpread1.MaxCols
       vaSpread1.Col = i: totnrorac = 0
       For X = 1 To vaSpread1.MaxRows - 1
           vaSpread1.Row = X
           If Trim(vaSpread1.Text) <> "" Then totnrorac = CCur(totnrorac + vaSpread1.Text)
      Next X
      If totnrorac > 0 Then vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = i: vaSpread1.Text = totnrorac
   Next i

   '------- Mover datos personal
   auxrutcli = ""
   Do While Not RS.EOF
      If Trim(RS!mir_rutcli) <> Trim(auxrutcli) And Trim(RS!mir_rutcli) = "PERSONAL" Then vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 1: vaSpread1.Text = "PERSONAL": vaSpread1.Col = 2: vaSpread1.Text = "PERSONAL": auxrutcli = RS!mir_rutcli
      If Trim(RS!mir_rutcli) = "PERSONAL" Then vaSpread1.Col = 2 + Mid(RS!mir_fecmin, 7, 2): vaSpread1.Text = RS!mir_nrorac: indper = True
      RS.MoveNext
   Loop
   
   '------- Fin mover datos personal
   If indper = False Then vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 1: vaSpread1.Text = "PERSONAL": vaSpread1.Col = 2: vaSpread1.Text = "PERSONAL"
   
   '------- Validar si existe datos personal
   
   '------- Bloquea días de cierre en color rojo
   Dim diablq As Date
   If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
   If Format(Now, "dd/mm/yyyy") > diablq Or Format(CDate(fpDateTime1.Text), "mm/yyyy") < Format(Month(Now) - 1 & "/" & Year(Now), "mm/yyyy") Then
        v_columnas = ((dEoM(Format(Day(Now) & "/" & (Month(Now) - 1) & "/" & Year(Now))) - CDate(CDate("01/" + Mid(fpDateTime1.Text, 1, 8))) + 1) * 2) - 1
   Else
        v_columnas = 0
   End If
        
   If v_columnas > 0 Then
      vaSpread1.Row = -1
      For i = 3 To v_columnas + 2
          vaSpread1.Col = i
          vaSpread1.Lock = True
          vaSpread1.BackColor = Shape1(0).FillColor
      Next i
      vaSpread1.SetActiveCell i, 1
   End If
   '------- Fin Bloqueo de celdas
   Gl_Ac_Botones Me, 1, 4, modo
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

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
If modo = "" Then modo = "M"
If ChangeMade = True And modo = "M" Then Gl_Ac_Botones Me, 1, 0, modo: fpLongInteger1(1).Enabled = False: Image1(1).Enabled = False: fpLongInteger1(2).Enabled = False: fpDateTime1.Enabled = False: Image1(2).Enabled = False
vaSpread1.Row = Row
vaSpread1.Col = Col
If Col > 2 Then
   Dim totnrorac As Long
   totnrorac = 0
   For i = 1 To vaSpread1.MaxRows
       vaSpread1.Row = i: vaSpread1.Col = 1
       If Trim(vaSpread1.Text) = "" Then Exit For
       vaSpread1.Col = Col
       If Trim(vaSpread1.Text) <> "" Then totnrorac = CCur(totnrorac + vaSpread1.Text)
   Next i
   If totnrorac > 0 Then vaSpread1.Col = Col: vaSpread1.Text = totnrorac Else vaSpread1.Col = Col: vaSpread1.Text = ""
End If
End Sub
