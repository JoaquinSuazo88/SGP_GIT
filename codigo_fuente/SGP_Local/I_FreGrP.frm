VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_FreGrP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frecuencia De Recetas o Gramos Producto Mensual"
   ClientHeight    =   3885
   ClientLeft      =   2895
   ClientTop       =   2430
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3375
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   8775
      Begin VB.Frame Frame5 
         Caption         =   "Tipo Minuta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   3255
         Begin VB.OptionButton Option5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Real"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1800
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Teórica"
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
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   3255
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   3
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   4
            Top             =   300
            Width           =   735
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   2700
            Picture         =   "I_FreGrP.frx":0000
            Top             =   150
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   3480
         TabIndex        =   11
         Top             =   1440
         Width           =   4575
         Begin VB.OptionButton Option1 
            Caption         =   "Frec. Recetas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   720
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Grs Prod. Mensual"
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
            Left            =   720
            TabIndex        =   6
            Top             =   840
            Width           =   1935
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   7
         Top             =   2850
         Visible         =   0   'False
         Width           =   1365
         _Version        =   393216
         _ExtentX        =   2408
         _ExtentY        =   450
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   13
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "I_FreGrP.frx":030A
         StartingColNumber=   6
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   1
         Top             =   660
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
         BackColor       =   -2147483624
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
         ButtonDefaultAction=   -1  'True
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
         NullColor       =   -2147483624
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "0"
         NegFormat       =   1
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
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   330
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
         BackColor       =   -2147483624
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
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
         NullColor       =   -2147483624
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   5
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   1035
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1658
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
         BackColor       =   -2147483624
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
         ButtonStyle     =   1
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
         AutoAdvance     =   0   'False
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
         NullColor       =   -2147483624
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "10/2021"
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
         AutoMenu        =   0   'False
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
         Index           =   2
         Left            =   3150
         TabIndex        =   17
         Top             =   660
         Width           =   4845
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3150
         TabIndex        =   15
         Top             =   330
         Width           =   4845
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2670
         Picture         =   "I_FreGrP.frx":0780
         Top             =   570
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2670
         Picture         =   "I_FreGrP.frx":0A8A
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha(mm/aa)"
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
         Left            =   465
         TabIndex        =   12
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "Ceco"
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
         Left            =   465
         TabIndex        =   10
         Top             =   405
         Width           =   480
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   465
         TabIndex        =   9
         Top             =   675
         Width           =   960
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3195
         TabIndex        =   18
         Top             =   705
         Width           =   4845
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3195
         TabIndex        =   16
         Top             =   375
         Width           =   4845
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_FreGrP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim est As Boolean
Dim MsgTitulo As String
Public lc_Aux As String

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.Height = 4260
Me.Width = 9165
Me.HelpContextID = vg_OpcM

fg_centra Me
fg_carga ""

MsgTitulo = "Frecuencia De Recetas o Gramos Producto Mensual"
est = True
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico Planificacón Teórica"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

fpText.Enabled = ModCasino
Image1(1).Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(1).Caption = MuestraCasino(2)

vaSpread1(0).MaxRows = 0
fpDateTime1.text = Format(Date, "mm/yyyy")
est = False

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpDateTime1_Change()

On Error GoTo Man_Error

If est Then Exit Sub
MoverDatosVector

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
    
If Val(fpLongInteger1(1).Value) < 1 Then
       
   fpayuda(2).Caption = ""
   Exit Sub
    
End If
    
Set RS = vg_db.Execute("sgp_Sel_RegimenxCodigo " & fpLongInteger1(1).Value & "")

If RS.EOF Then
       
   RS.Close
   Set RS = Nothing
   fpayuda(2).Caption = ""
   Exit Sub
    
End If
    
fpayuda(2).Caption = Trim(RS!reg_nombre)
RS.Close
Set RS = Nothing
MoverDatosVector

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 120
    
    If Index = 0 Then Image1_Click 0
    If Index = 1 Then Image1_Click 2

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Set RS = vg_db.Execute("sgp_Sel_clientes 1, '" & LimpiaDato(fpText.text) & "'")
If RS.EOF Then

   RS.Close
   Set RS = Nothing
   fpayuda(1).Caption = ""
   fpLongInteger1(1).text = ""
   fpayuda(2).Caption = ""
   fpDateTime1.Enabled = True
   Exit Sub

End If

fpayuda(1).Caption = Trim(RS!cli_nombre)
fpText.text = RS!cli_codigo
RS.Close
Set RS = Nothing
 
fpDateTime1.Enabled = True

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo


End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 120
    
    Image1_Click 1

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 2
    
    Option1(2).Value = True
    Option1(3).Value = False
    Image1(3).Enabled = False

Case 3
    
    Option1(2).Value = False
    Option1(3).Value = True
    Image1(3).Enabled = True

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Option1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 120
    
    Image1_Click 2

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 1
    
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    Call B_TabEst.LlenaDatos("b_clientes", "cli_", "Clientes", "Cliente_SitioRemoto")
    Call B_TabEst.Show(1)
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.text = vg_codigo
    fpayuda(1).Caption = vg_nombre
    fpLongInteger1(1).Value = ""
    Let fpayuda(2).Caption = ""
    fpLongInteger1(1).SetFocus

Case 2
    
    vg_left = fpayuda(2).Left + 2300
    vg_nombre = "": vg_codigo = ""
    Call B_TabEst.LlenaDatos("a_regimen", "reg_", "Regimen", "RegBlo")
'    Call B_TabEst.LlenaDatos("a_regimen", "", "Regimen", "RegBlo")
    Call B_TabEst.Show(1)
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre
    fpDateTime1.SetFocus

Case 3
    
    OpcionLectura = "6"
    vg_nombre = "": vg_codigo = ""
    vg_codigo = Trim(LimpiaDato(fpText.text))
    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, vg_codigo, Val(fpLongInteger1(1).Value) & ",", Format(fpDateTime1.text, "yyyymmdd"), Format(fpDateTime1.text, "yyyymmdd"), "1", "", 0, IIf(Option5(0).Value = True, 1, 2)
'    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, 0, Val(fpLongInteger1(1).Value), 0, Format(fpDateTime1.text, "yyyymm"), 0, "6", 0
    B_MTaEst.Show 1
    Me.Refresh
    
    If vg_codigo = "" Then
       
       Exit Sub
    
    End If

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

'On Error GoTo Man_Error

Select Case Button.Index

Case 1
    
    If Not ValidarDatos Then Exit Sub
    
    vg_opimp = 0
    Toolbar1.Enabled = False
    Frame1(0).Enabled = False
    
    If Option1(4).Value = True Then
       
       I_FrecuenciaRecetas Me
    
    ElseIf Option1(1).Value = True Then
       
       I_GramosProductos Me
    
    End If
    
    Toolbar1.Enabled = True
    Frame1(0).Enabled = True

Case 3
    
    Set RS = vg_db.Execute("sgp_Sel_Clientes 1, '" & Trim(LimpiaDato(fpText.text)) & "'")
    If RS.EOF Then
       
       RS.Close
       Set RS = Nothing
       MsgBox "No existe ceco planificado", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    vg_codigo = ""
    B_HistPm.LlenarHistPlan "Histórico Minuta", Trim(LimpiaDato(fpText.text)), 2, 1
'    B_HistPm.LlenarHistPlan "Histórico Minuta", 0, Trim(LimpiaDato(fpText.text)), 5
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codregimen)
    fpDateTime1.text = vg_fecha
    Me.Refresh

Case 5
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
Toolbar1.Enabled = True
Frame1(0).Enabled = True
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub MoverDatosVector()

On Error GoTo Man_Error

If LimpiaDato(Trim(fpText.text)) = "" Or Trim(fpDateTime1.text) = "" And Val(fpLongInteger1(1).Value) = 0 Then Exit Sub

fg_carga ""

Set RS = vg_db.Execute("sgp_Sel_ServicioMinutaMes '" & LimpiaDato(Trim(fpText.text)) & "', " & fpLongInteger1(1).Value & ", " & Format(fpDateTime1.text, "yyyymm") & "")
vaSpread1(0).MaxRows = 0

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
      vaSpread1(0).Row = vaSpread1(0).MaxRows
      
      vaSpread1(0).Col = 2
      vaSpread1(0).Value = RS(0)
      
      vaSpread1(0).Col = 3
      vaSpread1(0).Value = Trim(RS(1))
      
      RS.MoveNext
   
   Loop

End If
RS.Close
Set RS = Nothing
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub ActivarGrillaTodos()

On Error GoTo Man_Error

Dim i As Long

For i = 1 To vaSpread1(0).MaxRows
    
    vaSpread1(0).Row = i
    vaSpread1(0).Col = 1
    vaSpread1(0).CellType = 10
    vaSpread1(0).TypeCheckText = ""
    vaSpread1(0).TypeCheckCenter = True
    vaSpread1(0).text = "1" ' checked

Next i

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Function ValidarDatos() As Boolean

On Error GoTo Man_Error

Dim seleccion As Integer
Dim i         As Long

ValidarDatos = True

'-------> Validar ceco
If Trim(fpayuda(1).Caption) = "" Then

   MsgBox "Debe registrar ceco...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

'-------> Validar regimen
If Trim(fpayuda(2).Caption) = "" Then

   MsgBox "Debe registrar regimen...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

'-------> Validar fechas
If Trim(fpDateTime1.text) = "" Then
   
   MsgBox "Fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

If Option1(2).Value = True Then
    
   ActivarGrillaTodos
    
End If

'-------> Validar servicios
seleccion = 0
For i = 1 To vaSpread1(0).MaxRows
        
    vaSpread1(0).Row = i
    vaSpread1(0).Col = 1
    
    If vaSpread1(0).text = "1" Then
       
       seleccion = 1
       Exit For
    
    End If
    
Next i
    
If seleccion = 0 Then
   
   MsgBox "Servicio debe ser selecionado", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function
   
End If
    
Exit Function
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
   
End Function
