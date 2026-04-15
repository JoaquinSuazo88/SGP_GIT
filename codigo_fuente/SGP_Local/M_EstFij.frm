VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_EstFij 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estructura Fija del Servicio"
   ClientHeight    =   6765
   ClientLeft      =   2550
   ClientTop       =   1950
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   30
      TabIndex        =   20
      Top             =   2400
      Width           =   7335
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   1695
         TabIndex        =   22
         Top             =   315
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Día de Consumo"
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
         Left            =   165
         TabIndex        =   21
         Top             =   360
         Width           =   1425
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3375
      Left            =   30
      TabIndex        =   11
      Top             =   3240
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
      MaxCols         =   5
      MaxRows         =   1
      SpreadDesigner  =   "M_EstFij.frx":0000
      ScrollBarTrack  =   3
   End
   Begin VB.Frame Frame1 
      Height          =   1965
      Index           =   1
      Left            =   30
      TabIndex        =   5
      Top             =   360
      Width           =   7335
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1635
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
         Left            =   1635
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
         Left            =   1620
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
         Left            =   1635
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3300
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3300
         TabIndex        =   16
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3300
         TabIndex        =   14
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3300
         TabIndex        =   12
         Top             =   360
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2865
         Picture         =   "M_EstFij.frx":040E
         Top             =   1365
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2865
         Picture         =   "M_EstFij.frx":0718
         Top             =   1005
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2865
         Picture         =   "M_EstFij.frx":0A22
         Top             =   645
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2865
         Picture         =   "M_EstFij.frx":0D2C
         Top             =   285
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contrato"
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
         Left            =   150
         TabIndex        =   9
         Top             =   480
         Width           =   735
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
         Left            =   150
         TabIndex        =   8
         Top             =   825
         Width           =   750
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
         Left            =   150
         TabIndex        =   7
         Top             =   1185
         Width           =   705
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
         Left            =   150
         TabIndex        =   6
         Top             =   1515
         Width           =   1425
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3345
         TabIndex        =   13
         Top             =   405
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3345
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
         Left            =   3345
         TabIndex        =   17
         Top             =   1125
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3345
         TabIndex        =   19
         Top             =   1485
         Width           =   1815
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_EstFij"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim i As Long
Dim Msgtitulo As String
Dim modo As String, codpro As String
Dim est As Boolean, OpGr As Boolean, accion As Boolean

Private Sub Combo1_Click(Index As Integer)
If Combo1(0).ListIndex = -1 Or accion = False Then Exit Sub
MoverDatos
End Sub

Private Sub Form_Activate()
fg_descarga
'-------> Traer fecha cierre día
TraerFechaCierre
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 7245
Me.Width = 7560
EspFecha fpDateTime1
Msgtitulo = "Estructura Fija"
fg_centra Me
est = True: modo = ""
Gl_Mo_Botones Me, 2
Gl_Ac_Botones Me, 2, 3, modo
accion = True
Combo1(0).Clear
Combo1(0).AddItem "Lunes" & Space(150) & "(1)"
Combo1(0).AddItem "Martes" & Space(150) & "(2)"
Combo1(0).AddItem "Miércoles" & Space(150) & "(3)"
Combo1(0).AddItem "Jueves" & Space(150) & "(4)"
Combo1(0).AddItem "Viernes" & Space(150) & "(5)"
Combo1(0).AddItem "Sábado" & Space(150) & "(6)"
Combo1(0).AddItem "Domingo" & Space(150) & "(7)"
Combo1(0).ListIndex = -1
OpGr = False: vaSpread1.MaxRows = 0
fg_FormatearCeldaGrillaNuemricaDecimal vaSpread1, 3
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
fpDateTime1.text = Format(Date, "dd/mm/yyyy")
est = False
End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then
'   Frame1.Move 4200, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
End If
Toolbar1.Refresh
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpDateTime1_LostFocus()
fpayuda(3).Caption = fg_Fecha_Dia(Format(fpDateTime1.text, "yyyymmdd"), 2)
MoverDatos
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 1
    RS.Open RutinaLectura.Regimen(2, Val(fpLongInteger1(1).Value), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": vaSpread1.MaxRows = 0: Gl_Ac_Botones Me, 2, 3, modo: Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    MoverDatos
Case 2
    RS.Open RutinaLectura.Servicio(2, Val(fpLongInteger1(2).Value), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "": vaSpread1.MaxRows = 0: Gl_Ac_Botones Me, 2, 3, modo: Exit Sub
    fpayuda(2).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    MoverDatos
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Trim(fpLongInteger1(Index).text) = "" Or Val(fpLongInteger1(Index).Value) < 1 Then fpLongInteger1(Index).text = ""
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    If Index = 1 Then Image1_Click 1
    If Index = 2 Then Image1_Click 2
End Select
End Sub

Private Sub fpText_Change()
If fpText.text = "" Or est Then fpayuda(0).Caption = "": Exit Sub
RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": fpLongInteger1(2).Value = "": fpayuda(3).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
fpLongInteger1(2).Value = "": fpayuda(2).Caption = ""
MoverDatos
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

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.text = vg_codigo
    fpayuda(0).Caption = vg_nombre
    fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(2).Value = "": fpayuda(2).Caption = ""
    fpLongInteger1(1).SetFocus
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
Case 2
    vg_left = fpayuda(2).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Ser"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre
    fpDateTime1.SetFocus
Case 3
    If fpText.text = "" Or Val(fpLongInteger1(1).Value) < 1 Or Val(fpLongInteger1(2).Value) < 0 Or fpDateTime1.text = "" Then Exit Sub
    B_HistPm.LlenarHistPlan "Histórico Estructura Fija", fpText.text, fpLongInteger1(1).text & "|" & fpLongInteger1(2).text & "|", 3
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    fpDateTime1.text = vg_codigo
    accion = False: Combo1(0).ListIndex = vg_auxfecha - 1: accion = True
    MoverDatos
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codigo As Long, Nombre As String, Orden As String
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    If Trim(fpayuda(0).Caption) = "" Or Trim(fpayuda(1).Caption) = "" Or Trim(fpayuda(2).Caption) = "" Or Trim(fpayuda(3).Caption) = "" Then MsgBox "Falta información en encabezado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If CDate(fpDateTime1.text) < CDate(vg_ciedia) Then MsgBox "La fecha de la estructura fija es menor a la fecha cierre diario...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    modo = "A"
    Gl_Ac_Botones Me, 2, 0, modo
    fpText.Enabled = ModCasino: Image1(0).Enabled = ModCasino
    fpLongInteger1(1).Enabled = False: Image1(1).Enabled = False
    fpLongInteger1(2).Enabled = False: Image1(2).Enabled = False
    fpDateTime1.Enabled = False: Image1(3).Enabled = False: Combo1(0).Enabled = False
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 1: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus
Case 3
    If Trim(fpayuda(0).Caption) = "" Or Trim(fpayuda(1).Caption) = "" Or Trim(fpayuda(2).Caption) = "" Or Trim(fpayuda(3).Caption) = "" Then MsgBox "Falta información en encabezado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    modo = "M"
    fpText.Enabled = ModCasino: Image1(0).Enabled = ModCasino
    fpLongInteger1(1).Enabled = False: Image1(1).Enabled = False
    fpLongInteger1(2).Enabled = False: Image1(2).Enabled = False
    fpDateTime1.Enabled = False: Image1(3).Enabled = False: Combo1(0).Enabled = False
    Gl_Ac_Botones Me, 2, 0, modo
Case 5
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codpro = vaSpread1.text
    vg_db.BeginTrans
    vg_db.Execute "DELETE b_minutafija FROM b_minutafija WHERE mif_cencos='" & fpText.text & "' AND mif_codreg=" & Val(fpLongInteger1(1).Value) & " AND mif_codser=" & Val(fpLongInteger1(2).Value) & " AND mif_fecval=" & Format(fpDateTime1.text, "yyyymmdd") & " AND mif_codpro='" & codpro & "' AND mif_dianro=" & Val(fg_codigocbo(Combo1, 0, 1, "")) & ""
    vg_db.CommitTrans
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    modo = "": Gl_Ac_Botones Me, 2, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
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
    fpText.Enabled = ModCasino: Image1(0).Enabled = ModCasino
    fpLongInteger1(1).Enabled = True: Image1(1).Enabled = True
    fpLongInteger1(2).Enabled = True: Image1(2).Enabled = True
    fpDateTime1.Enabled = True: Image1(3).Enabled = True: Combo1(0).Enabled = True
    modo = "": Gl_Ac_Botones Me, 2, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
Case 12
    GrabaRegistro vaSpread1.ActiveRow
Case 15
    M_CpoEsF.LlenarDatos fpText.text, Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Format(fpDateTime1.text, "yyyymmdd"), Val(fg_codigocbo(Combo1, 0, 1, ""))
    M_CpoEsF.Show 1
Case 17
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_EstructuraFija Trim(fpText.text), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Format(fpDateTime1.text, "yyyymmdd")
Case 20
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
vg_db.RollbackTrans
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Or Row < 1 Then Exit Sub
Select Case Col
Case 1
    '------- Validar si existe productos estructura fija
    vaSpread1.Row = Row
    vaSpread1.Col = 5: codpro = ""
    If vaSpread1.text <> "" Then codpro = vaSpread1.text: vaSpread1.Col = 1: vaSpread1.text = codpro: Exit Sub
    vaSpread1.Row = Row: vaSpread1.Col = 1: If vaSpread1.Lock = True Then Exit Sub
    ' llama  a formulario de busqueda de productos y carga datos
    vaSpread1.Col = 1
    vg_nombre = "": vg_codigo = ""
    vg_left = fpayuda(2).Left + 2300
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Pst"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If Trim(vg_codigo) = Trim(vaSpread1.text) And Row <> i And Trim(vaSpread1.text) <> "" Then MsgBox "Productos existe", vbCritical + vbOKOnly, Msgtitulo: vaSpread1.text = "": vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
    Next i
    RS.Open "SELECT a.pro_codigo, a.pro_nombre, b.uni_nomcor FROM b_productos a, a_unidad b WHERE a.pro_coduni = b.uni_codigo AND a.pro_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic
    If RS.EOF Then: RS.Close: Set RS = Nothing: vaSpread1.text = "": vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
    vaSpread1.Col = 1: vaSpread1.text = RS!pro_codigo
    vaSpread1.Col = 2: vaSpread1.text = RS!pro_nombre
    vaSpread1.Col = 4: vaSpread1.text = RS!uni_nomcor
    RS.Close: Set RS = Nothing
    vaSpread1.SetActiveCell 3, vaSpread1.ActiveRow
End Select
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
Select Case Col
Case 1
    vaSpread1.Row = Row
    vaSpread1.Col = 5: codpro = ""
    If Trim(vaSpread1.text) <> "" Then codpro = vaSpread1.text: vaSpread1.Col = 1: vaSpread1.text = codpro: Exit Sub
Case 3
    If ChangeMade = True Then
       fpText.Enabled = ModCasino: Image1(0).Enabled = ModCasino
       fpLongInteger1(1).Enabled = False: Image1(1).Enabled = False
       fpLongInteger1(2).Enabled = False: Image1(2).Enabled = False
       fpDateTime1.Enabled = False: Image1(3).Enabled = False: Combo1(0).Enabled = False
       If modo = "" Then modo = "M"
       Gl_Ac_Botones Me, 2, 0, modo
    End If
End Select
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
Dim canpro As Double
Select Case Col
Case 1
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    vaSpread1.Col = 5: codpro = ""
    If vaSpread1.text <> "" Then codpro = vaSpread1.text: vaSpread1.Col = 1: vaSpread1.text = codpro: Exit Sub
    vaSpread1.Col = 1
    RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, a.pro_ctrsto, b.uni_nomcor FROM b_productos a, a_unidad b, a_tiposervicio c, b_clientes d WHERE (c.tis_codigo = d.cli_codtis OR a.pro_maepro < 1) AND d.cli_codigo = '" & MuestraCasino(1) & "' AND (c.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) AND a.pro_coduni = b.uni_codigo AND a.pro_codigo = '" & vaSpread1.text & "' AND (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) AND a.pro_ctrsto = 1", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1.text = "": vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
    codpro = vaSpread1.text
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        If Trim(codpro) = Trim(vaSpread1.text) And Row <> i And Trim(vaSpread1.text) <> "" Then RS.Close: Set RS = Nothing: vaSpread1.Row = Row: vaSpread1.text = "": MsgBox "Productos existe", vbCritical + vbOKOnly, Msgtitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
    Next i
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    vaSpread1.Col = 1: vaSpread1.text = RS!pro_codigo
    vaSpread1.Col = 2: vaSpread1.text = RS!pro_nombre
    vaSpread1.Col = 4: vaSpread1.text = RS!uni_nomcor
    RS.Close: Set RS = Nothing
    vaSpread1.SetActiveCell 3, vaSpread1.ActiveRow
    vaSpread1.Row = NewRow
    vaSpread1.Col = 1: codpro = vaSpread1.text
    vaSpread1.Col = 3: canpro = Val(vaSpread1.text)
Case 2, 3, 4
    If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
       GrabaRegistro Row
    ElseIf Toolbar1.Buttons(12).Visible = False Then
'       Cancela
    End If
End Select
End Sub

Sub MoverDatos()
With vaSpread1
    .MaxRows = 0
    RS.Open RutinaLectura.MinutaFija(4, Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Format(fpDateTime1.text, "yyyymmdd"), Val(fg_codigocbo(Combo1, 0, 1, "")), "", "", ""), vg_db, adOpenStatic
    If Not RS.EOF Then
       Do While Not RS.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1: .text = RS!pro_codigo
          .Col = 2: .text = RS!pro_nombre
          .Col = 3: .text = RS!mif_canpro
          .Col = 4: .text = RS!uni_nomcor
          .Col = 5: .text = RS!pro_codigo
          RS.MoveNext
       Loop
       Gl_Ac_Botones Me, 2, 1, modo
       .SetActiveCell 1, 1
    Else
       If fpText.text = "" Or Val(fpLongInteger1(1).Value) < 1 Or Val(fpLongInteger1(2).Value) < 0 Or fpDateTime1.text = "" Or Combo1(0).ListIndex = -1 Then
          Gl_Ac_Botones Me, 2, 3, modo
       Else
          Gl_Ac_Botones Me, 2, 2, modo
       End If
    End If
    RS.Close: Set RS = Nothing
End With
End Sub

Private Sub GrabaRegistro(Fila As Long)
Dim canpro As Double
OpGr = True
vaSpread1.Row = Fila
codpro = "": canpro = 0
vaSpread1.Col = 1: codpro = vaSpread1.text
vaSpread1.Col = 3: canpro = Val(vaSpread1.text)
If Trim(fpayuda(0).Caption) = "" Or Trim(fpayuda(1).Caption) = "" Or Trim(fpayuda(2).Caption) = "" Or Trim(fpayuda(3).Caption) = "" Then MsgBox "Falta información en encabezado...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 3, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub
If Trim(codpro) = "" Or (canpro = 0 Or canpro < 0) Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 3, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub
If CDate(fpDateTime1.text) < CDate(vg_ciedia) Then MsgBox "La fecha de la estructura fija es menor a la fecha cierre diario...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
If modo = "A" Then
    vg_db.BeginTrans
    vg_db.Execute "INSERT INTO b_minutafija (mif_cencos, mif_codreg, mif_codser, mif_fecval, mif_codpro, mif_dianro, mif_canpro) " & _
                  "VALUES ('" & fpText.text & "', " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & Val(Format(fpDateTime1.text, "yyyymmdd")) & ", '" & codpro & "', " & Val(fg_codigocbo(Combo1, 0, 1, "")) & ", " & canpro & ")"
    vg_db.CommitTrans
    vaSpread1.Col = 1: vaSpread1.text = codpro
    vaSpread1.Col = 5: vaSpread1.text = codpro
Else
    vg_db.BeginTrans
    vg_db.Execute "UPDATE b_minutafija SET mif_canpro = " & canpro & " WHERE mif_cencos = '" & fpText.text & "' AND mif_codreg = " & Val(fpLongInteger1(1).Value) & " AND mif_codser = " & Val(fpLongInteger1(2).Value) & " AND mif_fecval = " & Format(fpDateTime1.text, "yyyymmdd") & " AND mif_codpro = '" & codpro & "' AND mif_dianro = " & Val(fg_codigocbo(Combo1, 0, 1, "")) & ""
    vg_db.CommitTrans
End If
fpText.Enabled = ModCasino: Image1(0).Enabled = ModCasino
fpLongInteger1(1).Enabled = True: Image1(1).Enabled = True
fpLongInteger1(2).Enabled = True: Image1(2).Enabled = True
fpDateTime1.Enabled = True: Image1(3).Enabled = True: Combo1(0).Enabled = True
modo = "": Gl_Ac_Botones Me, 2, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
OpGr = False

Exit Sub
Man_Error:
vg_db.RollbackTrans
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Cancela()
OpGr = True
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: codpro = vaSpread1.text
RS.Open RutinaLectura.MinutaFija(4, Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Format(fpDateTime1.text, "yyyymmdd"), Val(fg_codigocbo(Combo1, 0, 1, "")), "", "", codpro), vg_db, adOpenStatic
If Not RS.EOF Then
   vaSpread1.Col = 2: vaSpread1.text = Trim(RS!pro_nombre)
   vaSpread1.Col = 3: vaSpread1.text = RS!mif_canpro
   vaSpread1.Col = 4: vaSpread1.text = Trim(RS!uni_nomcor)
End If
RS.Close: Set RS = Nothing
OpGr = False
End Sub
