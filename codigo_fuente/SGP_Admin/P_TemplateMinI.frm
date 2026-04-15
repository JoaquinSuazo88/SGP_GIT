VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form P_TemplateMinI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Template Minuta I"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   8595
      Left            =   105
      TabIndex        =   11
      Top             =   90
      Width           =   13710
      Begin VB.OptionButton Option1 
         Caption         =   "Ponderaciones por Estructura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   24
         Top             =   1080
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Plantilla Frecuencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   23
         Top             =   1080
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Index           =   1
         Left            =   2640
         TabIndex        =   22
         Top             =   7800
         Width           =   2475
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   6
            Top             =   135
            Width           =   2370
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   1
         Left            =   1560
         TabIndex        =   21
         Top             =   7800
         Width           =   930
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   5
            Top             =   135
            Width           =   825
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   8640
         TabIndex        =   19
         Top             =   7800
         Width           =   930
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   45
            TabIndex        =   9
            Top             =   135
            Width           =   825
         End
      End
      Begin VB.Frame Frame2 
         Height          =   435
         Left            =   9615
         TabIndex        =   18
         Top             =   7800
         Width           =   3630
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   7
            Left            =   45
            TabIndex        =   10
            Top             =   135
            Width           =   3525
         End
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Index           =   0
         Left            =   6150
         TabIndex        =   16
         Top             =   7800
         Width           =   2475
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   45
            TabIndex        =   8
            Top             =   135
            Width           =   2370
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   0
         Left            =   5175
         TabIndex        =   15
         Top             =   7800
         Width           =   930
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   7
            Top             =   135
            Width           =   825
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5850
         Left            =   105
         TabIndex        =   4
         Top             =   1920
         Width           =   13380
         _Version        =   393216
         _ExtentX        =   23601
         _ExtentY        =   10319
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   9
         SpreadDesigner  =   "P_TemplateMinI.frx":0000
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   2580
         TabIndex        =   0
         Top             =   165
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   2580
         TabIndex        =   1
         Top             =   615
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ButtonStyle     =   2
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
         Text            =   "01/09/2013"
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
      Begin EditLib.fpDateTime FpFecHasta 
         Height          =   315
         Left            =   8445
         TabIndex        =   2
         Top             =   615
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ButtonStyle     =   2
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
         Text            =   "28/09/2013"
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   13065
         TabIndex        =   3
         Top             =   1035
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cargar Información"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   105
         TabIndex        =   20
         Top             =   8265
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha hasta"
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
         Left            =   7260
         TabIndex        =   14
         Top             =   705
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha desde"
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
         Left            =   1305
         TabIndex        =   13
         Top             =   705
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Org. Compras"
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
         Left            =   1305
         TabIndex        =   12
         Top             =   270
         Width           =   1155
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8790
      Left            =   13365
      TabIndex        =   17
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   15505
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "P_TemplateMinI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public lc_Aux    As String
Dim BtnX         As Object
Dim Rs_Receta    As New ADODB.Recordset
Dim VecRecetas() As Variant

Private Sub Check1_Click()

On Error GoTo Man_Error

If Check1.Value = 0 Or Check2.Value = 0 Then

    Option1(0).Visible = False
    Option1(1).Visible = False

ElseIf Check1.Value = 1 And Check2.Value = 1 Then

    Option1(0).Visible = True
    Option1(1).Visible = True

End If

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Check2_Click()

On Error GoTo Man_Error

If Check2.Value = 1 Then
   
   Check3.Visible = True
   Check3.Value = 0

Else
   
   Check3.Visible = False
   Check3.Value = 0

End If

If Check1.Value = 0 Or Check2.Value = 0 Then

    Option1(0).Visible = False
    Option1(1).Visible = False

ElseIf Check1.Value = 1 And Check2.Value = 1 Then

    Option1(0).Visible = True
    Option1(1).Visible = True

End If

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Check5_Click()

On Error GoTo Man_Error

If Check5.Value = 1 Then

   Check6.Visible = True
   Command1.Visible = True
   
Else

   Check6.Visible = False
   Command1.Visible = False

End If

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Command1_Click()

On Error GoTo Man_Error

B_DieTipExcel.Show 1

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub




Private Sub Dir1_Change()

 On Error GoTo errores
 
 Dim nomarc As String
 
   Dir1.Path = fpText1.text   ' Establece la ruta del directorio.
   
   nomarc = fg_ArchivoTXT_1(Dir1.Path & "\")
   Open nomarc For Output As #1
   Print #1, "1"
   Close #1

   If Dir(nomarc) <> "" Then Kill nomarc

   Exit Sub
   
errores:
   
   MsgBox "La carpeta " & fpText1.text & " no está disponible o bien no tiene permiso de escritura.", vbCritical + vbExclamation, "Error"
   fpText1.text = dir_trabajo
   
End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me

Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exportar Excel": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

fpText1.text = dir_trabajo
B_DieTipExcel.MoverDatosTvwDir
vg_RowEnd = 0
FormatearDatos

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FormatearDatos()

On Error GoTo Man_Error

    Let FpFecDesde.text = Format(Date, "dd/mm/yyyy")
    Let FpFecHasta.text = Format(Date, "dd/mm/yyyy")
    Let vaSpread1.MaxRows = 0
    Let fpText.text = ""
    Let Text1(2).text = ""
    Let Text1(3).text = ""
    Let Text1(4).text = ""
    Let Text1(5).text = ""
    Let Text1(6).text = ""
    Let Text1(7).text = ""

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub
vaSpread1.MaxRows = 0

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

If IsDate(FpFecHasta.text) = False Then Exit Sub
vaSpread1.MaxRows = 0

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_Change()

On Error GoTo Man_Error

vaSpread1.MaxRows = 0

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
    
On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_ButtonHit(Button As Integer, NewIndex As Integer)

On Error GoTo Man_Error

    Dim ret As String
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
    ret = Buscar_Carpeta(" ... Seleccione una carpeta ")

    If Trim(ret) <> "" Then
    
       fpText1.text = ret & "\"
       Dir1_Change
       
    End If
    
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i             As Long
Dim X             As Long
Dim indactivo     As Integer
Dim TexBus        As String
Dim EstBuq        As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(Text1(Index).text, ",")

If Index = 2 Then
   
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(6).text = ""
   Text1(7).text = ""

ElseIf Index = 3 Then
   
   Text1(2).text = ""
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(6).text = ""
   Text1(7).text = ""

ElseIf Index = 4 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(5).text = ""
   Text1(6).text = ""
   Text1(7).text = ""

ElseIf Index = 5 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(6).text = ""
   Text1(7).text = ""

ElseIf Index = 6 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(7).text = ""

ElseIf Index = 7 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(6).text = ""

End If

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 8
    vaSpread1.text = 0

Next

Select Case Index

Case 2, 3, 4, 5, 6, 7
    
    vaSpread1.Visible = False
    
    If Trim(Text1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread1.Col = 2
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 8
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 8
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 8
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 8
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 8
                 vaSpread1.text = 0
                 
              End If
           
           End If
        
        Next i
        
        vaSpread1.SetActiveCell Index + 1, 1
        vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread1.ColUserSortIndicator(IIf(Trim(wvarArr8020(X)) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread1.SortKey(1) = IIf(Trim(wvarArr8020(X)) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(Text1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 8
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim varIdColaTrabajo As Integer
Dim varTrabajoxLotes As Boolean

varIdColaTrabajo = 0
varTrabajoxLotes = False

If Check7.Value = 1 Then varTrabajoxLotes = True

Select Case Button.Index
Case 1
    
    If Check4.Value = 0 And Check6.Value = 1 And Check1.Value = 1 Then
       
       If Format(FpFecDesde.text, "mm") <> Format(FpFecHasta.text, "mm") Then
            
          Call MsgBox("Solo debe ser informado mes...", vbCritical, Me.Caption)
          Exit Sub
        
       End If
       
       If varTrabajoxLotes = False Then
          ExportarExcelMinutaBloqueCostoReceta
          MsgBox "Proceso Finalizado Correctamente", vbInformation, MsgTitulo
       Else
          varIdColaTrabajo = ReportePorLotes(1)
          MsgBox "Se ha ingresado la solicitud (N° " & varIdColaTrabajo & ") de generación del reporte por lotes.", vbInformation, MsgTitulo
       End If
    
       
    
    ElseIf Check4.Value = 0 And Check6.Value = 0 And Check1.Value = 1 Then
    
       If varTrabajoxLotes = False Then
          ExportarExcelMinutaBloqueSinCostoReceta
          MsgBox "Proceso Finalizado Correctamente", vbInformation, MsgTitulo
       Else
          varIdColaTrabajo = ReportePorLotes(2)
          MsgBox "Se ha ingresado la solicitud (N° " & varIdColaTrabajo & ") de generación del reporte por lotes.", vbInformation, MsgTitulo
       End If
    
       
    
    
    ElseIf Check4.Value = 1 Then
       
       If varTrabajoxLotes = False Then
          ExportarExcelMinutaBloqueSinCodigoReceta
          MsgBox "Proceso Finalizado Correctamente", vbInformation, MsgTitulo
       Else
          varIdColaTrabajo = ReportePorLotes(3)
          MsgBox "Se ha ingresado la solicitud (N° " & varIdColaTrabajo & ") de generación del reporte por lotes.", vbInformation, MsgTitulo
       End If
       
       
    
    End If
    
Case 3
    
    Me.Hide
    Unload Me

End Select

Exit Sub

Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS  As New ADODB.Recordset
Dim Sql As String
    
    Select Case Button.Index
    
    Case 1 'Mostrar datos en la grilla
        
        Text1(2).text = ""
        Text1(3).text = ""
        Text1(4).text = ""
        Text1(5).text = ""
        Text1(6).text = ""
        Text1(7).text = ""
        
        If Trim(fpText.text) = "" Then
           
           MsgBox "Debe seleccionar Org. Compras...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           Exit Sub
        
        End If
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("sgpadm_Sel_OrgCompras_V02 '" & LimpiaDato(fpText.text) & "'")
        
        If RS.EOF Then
           
           RS.Close
           Set RS = Nothing
           MsgBox "No existe Org. Compras...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           vaSpread1.MaxRows = 0
           Exit Sub
        
        End If
        
        RS.Close
        Set RS = Nothing
        
        If CDate(FpFecDesde.text) > CDate(FpFecHasta.text) Then
            
            Call MsgBox("Fecha Desde No Puede Ser Mayor a Fecha Hasta", vbInformation, Me.Caption)
            Let FpFecDesde.text = Format(Now, "dd/mm/yyyy")
            Call FpFecDesde.SetFocus
            Exit Sub
        
        End If
    
        If CDate(FpFecHasta.text) < CDate(FpFecDesde.text) Then
            
            Call MsgBox("Fecha Hasta No Puede Ser Mayor a Fecha Desde", vbInformation, Me.Caption)
            Let FpFecHasta.text = Format(Now, "dd/mm/yyyy")
            Call FpFecHasta.SetFocus
            Exit Sub
        
        End If
        
        If DateDiff("y", FpFecDesde.text, FpFecHasta.text) > 365 Then
            
            Call MsgBox("Rango De Fecha No Puede Ser Mayor a 12 Meses", vbInformation, Me.Caption)
            vaSpread1.MaxRows = 0
            Exit Sub
        
        End If
        
        vaSpread1.Visible = False
        vaSpread1.MaxRows = 0
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
            
        Sql = ""
        Sql = Sql & LimpiaDato(Trim(fpText.text))
        Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
        Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
        Set RS = vg_db.Execute("sgpadm_Sel_OrgComprasCecoMBloque_V03 " & Sql & "")
        Do While Not RS.EOF
           
           vaSpread1.MaxRows = vaSpread1.MaxRows + 1
           vaSpread1.Row = vaSpread1.MaxRows
           vaSpread1.Col = 1
           vaSpread1.text = "0"
           
           vaSpread1.Col = 2
           vaSpread1.CellType = CellTypeStaticText
           vaSpread1.text = RS!Ceco
           
           vaSpread1.Col = 3
           vaSpread1.CellType = CellTypeStaticText
           vaSpread1.text = RS!Cli_nombre
           
           vaSpread1.Col = 4
           vaSpread1.CellType = CellTypeStaticText
           vaSpread1.text = RS!Regimen
           
           vaSpread1.Col = 5
           vaSpread1.CellType = CellTypeStaticText
           vaSpread1.text = RS!reg_nombre
           
           vaSpread1.Col = 6
           vaSpread1.CellType = CellTypeStaticText
           vaSpread1.text = RS!Servicio
           
           vaSpread1.Col = 7
           vaSpread1.CellType = CellTypeStaticText
           vaSpread1.text = RS!ser_nombre
           
           vaSpread1.Col = 8
           vaSpread1.CellType = CellTypeStaticText
           vaSpread1.text = 0
           
           RS.MoveNext
        
        Loop
        RS.Close
        Set RS = Nothing
        
'        If Check4.Value = 0 Then
'
'           vaSpread1.MaxRows = vaSpread1.MaxRows + 1
'           vaSpread1.Row = vaSpread1.MaxRows
'           vaSpread1.Col = 1
'           vaSpread1.text = "0"
'
'           vaSpread1.Col = 2
'           vaSpread1.CellType = CellTypeStaticText
'           vaSpread1.text = ""
'
'           vaSpread1.Col = 3
'           vaSpread1.CellType = CellTypeStaticText
'           vaSpread1.text = ""
'
'           vaSpread1.Col = 4
'           vaSpread1.CellType = CellTypeStaticText
'           vaSpread1.text = ""
'
'           vaSpread1.Col = 5
'           vaSpread1.CellType = CellTypeStaticText
'           vaSpread1.text = "Recetas"
'
'        End If
        
        vaSpread1.Visible = True
    
    End Select
    
Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

'Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
'
'On Error GoTo Man_Error
'
'Dim i As Long
'
'If BlockCol <> 1 Then Exit Sub
'
'vaSpread1.Col = 1
'
'For i = BlockRow To BlockRow2
'
'    vaSpread1.Row = i
'    vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
'
'Next
'
'Exit Sub
'Man_Error:
'    Call fg_descarga
'    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
'
'End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1
    
    Dim i As Long
    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    If BlockRow = -1 Then Exit Sub

    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    
        End If
        
    Next
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub ExportarExcelMinutaBloqueCostoReceta()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim RS1             As New ADODB.Recordset
Dim Sql             As String
Dim Est             As Boolean
Dim i               As Long
Dim j               As Long
Dim ii              As Long
Dim NomArchivoExcel As String

vg_RowEnd = 0

'-------> Validar Org. Compras
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
            
Set RS = vg_db.Execute("sgpadm_Sel_OrgCompras_V02 '" & LimpiaDato(fpText.text) & "'")
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   MsgBox "No existe Org. Compras...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   vaSpread1.MaxRows = 0
   Exit Sub

End If
RS.Close
Set RS = Nothing

'-------> validar seleción ceco - regimen - servicio
Est = False

For i = 1 To vaSpread1.MaxRows
    
    vaSpread1.Row = i
    vaSpread1.Col = 1
    
    If vaSpread1.text = "1" Then
       
       Est = True
    
    End If

Next i
If Not Est Then MsgBox "Regimen y servicios asociado debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub

'-------> Validar fecha desde - hasta
If CDate(FpFecDesde.text) > CDate(FpFecHasta.text) Then
   
   Call MsgBox("Fecha Desde No Puede Ser Mayor a Fecha Hasta", vbInformation, MsgTitulo)
   Let FpFecDesde.text = Format(Now, "dd/mm/yyyy")
   Call FpFecDesde.SetFocus
   Exit Sub

End If
    
If CDate(FpFecHasta.text) < CDate(FpFecDesde.text) Then
   
   Call MsgBox("Fecha Hasta No Puede Ser Mayor a Fecha Desde", vbInformation, MsgTitulo)
   Let FpFecHasta.text = Format(Now, "dd/mm/yyyy")
   Call FpFecHasta.SetFocus
   Exit Sub

End If

'If DateDiff("ww", FpFecDesde.text, FpFecHasta.text) > 52 Then
If DateDiff("y", FpFecDesde.text, FpFecHasta.text) > 365 Then
   
   Call MsgBox("Rango De Fecha No Puede Ser Mayor a 12 Meses", vbInformation, Me.Caption)
   Exit Sub

End If

Dim oExcel              As Object
Dim oBook               As Object
Dim oSheet              As Object
Dim RowSheet            As Long
Dim oCol                As String
Dim oColCla             As String
Dim oColA               As String
Dim oColFec             As String
Dim oColPor             As String
Dim oColCos             As String
Dim oColRec             As String
Dim Ceco                As String
Dim NomCeco             As String
Dim auxceco             As String
Dim NombreServicioAux   As String
Dim NombreServicio      As String
Dim CodigoServicio      As Long
Dim CodigoRegimen       As Long
Dim Aux_CodigoRegimen   As Long
Dim NombreRegimen       As String
Dim MaxColumna          As Long
Dim DiaColumna          As Long
Dim NumAsc              As Long
Dim FecMin              As Date
Dim FecMax              As Date
Dim Fecha               As Date
Dim IndCol              As Long
Dim IndColA             As Long
Dim IndVec              As Long
Dim IndHoja             As Long
Dim AuxFec              As Long
Dim CodEstructura       As Long
Dim RowEnd              As Long
Dim RowEnd1             As Long
Dim ColEnd              As String
Dim TotDiaRaciones      As Double
Dim TotCol              As Long
Dim EstCeco             As Boolean
Dim oColCpo             As String
Dim CalRacCos           As String
Dim TotCalRacCos        As String
Dim CalComen            As String
Dim SoloNomArchivoExcel As String
Dim LinRecetaTotal      As Long

Dim ClaveExcel          As String
Dim CostoComercial      As Double
Dim CostoPlanificacion  As Double
Dim CostoReceta         As Double
Dim CostoRecetaDet      As Double
Dim Comensales          As Double
Dim CostoBandeja        As Double
Dim EstVacio            As Boolean
Dim LyD                 As Boolean

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
    
fg_carga ""
Bar1(0).Visible = True
Bar1(0).Value = 0
EstCeco = True

ClaveExcel = "Jp123456"
             
Set RS1 = vg_db.Execute("sgpadm_s_parametro 1, 'parhojaexc', ''")
If Not RS1.EOF Then
                
   ClaveExcel = RS1(0)
             
End If
RS1.Close
Set RS1 = Nothing

For i = 1 To vaSpread1.MaxRows
    
    DoEvents
    
    vaSpread1.Row = i
    vaSpread1.Col = 1
    Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
    
    If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
       
       vaSpread1.SetActiveCell 2, vaSpread1.Row
       vaSpread1.Col = 2
       Ceco = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
       vaSpread1.Col = 3
       NomCeco = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
       
       CalRacCos = ""
       TotCalRacCos = ""
       CalComen = ""
       CostoReceta = 0
       Comensales = 0
       CostoRecetaDet = 0
       
       If Ceco <> auxceco Then
       
          If Trim(auxceco) <> "" Then
             
             If Check2.Value = 1 And Check3.Value = 1 Then
          
                oExcel.Sheets("Recetas").Select
                oExcel.Sheets("Recetas").Move Before:=oExcel.Sheets(1)
                
                oSheet.Select
                oExcel.ActiveSheet.Protect password:=ClaveExcel, DrawingObjects:=True, _
                                    Contents:=True, Scenarios:=True, AllowFormattingCells:=True
                                 
             End If
             
             oBook.Close True, NomArchivoExcel
'             oExcel.Visible = True '------->Visualizar
             
             Set oSheet = Nothing
             Set oExcel = Nothing
             Set oBook = Nothing
             
             '------- Copiar archivos en la ruta seleccionada, luego borrar archivos de la carpeta
             If Trim(NomArchivoExcel) <> Trim(fpText1.text & SoloNomArchivoExcel) Then
                
                fso.CopyFile NomArchivoExcel, fpText1.text & SoloNomArchivoExcel, True
             
             End If
             
             If Dir(NomArchivoExcel) <> "" And Trim(NomArchivoExcel) <> Trim(fpText1.text & SoloNomArchivoExcel) Then
             
                Kill NomArchivoExcel
                
             End If
             
             Me.Refresh
             
          End If
          
          '-------> Exportar excel
          Set oExcel = CreateObject("Excel.Application")
          Set oBook = oExcel.Workbooks.Add
          NumAsc = 66
       
          auxceco = Ceco
          EstCeco = True
          
          NomArchivoExcel = dir_trabajo & "ExcelMinutaSGP" & "\" & Ceco & "-" & NomCeco & " " & Format(FpFecDesde.text, "yyyymmdd") & " " & Format(FpFecHasta.text, "yyyymmdd") & " " & Replace(Date, "/", "") & " " & Replace(Time, ":", "") & ".xlsx"
          SoloNomArchivoExcel = Ceco & "-" & NomCeco & " " & Format(FpFecDesde.text, "yyyymmdd") & " " & Format(FpFecHasta.text, "yyyymmdd") & " " & Replace(Date, "/", "") & " " & Replace(Time, ":", "") & ".xlsx"
       
       End If
       
       vaSpread1.Col = 4
       CodigoRegimen = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
       
       vaSpread1.Col = 5
       NombreRegimen = IIf(Trim(vaSpread1.text) = "", "", vaSpread1.text)
       
       vaSpread1.Col = 6
       CodigoServicio = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
       
       vaSpread1.Col = 7
       NombreServicio = IIf(Trim(vaSpread1.text) = "", "", vaSpread1.text)
    
       '-------> Traer costo Comercial
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Sql = ""
       Sql = Sql & " '" & LimpiaDato(Trim(Ceco)) & "'"
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymm")
       Sql = Sql & ",'COMERCIAL'"
       Set RS = vg_db.Execute("sgpadm_Sel_CostoComercialCeco " & Sql & "")

       CostoComercial = 0
       If Not RS.EOF Then
       
          CostoComercial = RS!pcp_valor
       
       End If
       RS.Close
       Set RS = Nothing
       
       '-------> Traer parametro ceco x regimen
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Sql = ""
       Sql = Sql & " '" & LimpiaDato(Trim(Ceco)) & "'"
       Set RS = vg_db.Execute("sgpadm_Sel_ParametroCecoregimen " & Sql & "")

       Aux_CodigoRegimen = CodigoRegimen
       
       If Not RS.EOF Then
       
          Aux_CodigoRegimen = RS!par_codreg
       
       End If
       RS.Close
       Set RS = Nothing
       
       '-------> incluir recetas costo recetas
       If Check5.Value = 1 And EstCeco Then
       
          EstCeco = False
          '-------> Rutinas carga Recetas con Costo
          CargaRecetaCosto Ceco, Aux_CodigoRegimen

          '-------> incluir recetas
          LinRecetaTotal = 0
          If Check5.Value = 1 And Rs_Receta.State = 1 Then
             
             '-------> Add data to cells of the first worksheet in the new workbook
             NombreServicio = "Recetas"
             Set oSheet = Nothing
             Set oSheet = oBook.Worksheets.Add
             NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)

             If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then

                NombreServicioAux = Mid(NombreServicioAux, 1, 31)

             End If
             oSheet.Name = "Recetas" 'Trim(NombreServicioAux)
                
             Fecha = "01/" & Format(FpFecDesde.text, "mm/yyyy")

             '-------> Check version of Excel
             Call encabezado(Rs_Receta, oSheet)
          
             oSheet.Cells(2, 1).CopyFromRecordset Rs_Receta

             MoverDatosExcel oExcel, oSheet, "A", "A", 1, 1, Trim("Código")
             MoverDatosExcel oExcel, oSheet, "B", "B", 1, 1, Trim("Nombre Plato")
             MoverDatosExcel oExcel, oSheet, "C", "C", 1, 1, Trim("Costo Plato")
             MoverDatosExcel oExcel, oSheet, "D", "D", 1, 1, Trim("Categoria Dietetica")
             MoverDatosExcel oExcel, oSheet, "E", "E", 1, 1, Trim("Tipo Plato")
             PonerColorInteriorN oExcel, oSheet, "A", "E", 1, 1, 4

             LinRecetaTotal = Rs_Receta.RecordCount
             IndCol = 2
             PonerColorInteriorN oExcel, oSheet, "A", "E", IndCol, Rs_Receta.RecordCount + 1, 6
                    
             oSheet.Cells.Select
             oSheet.Cells.EntireColumn.AutoFit
             
             Rs_Receta.Close
             Set Rs_Receta = Nothing
             Set oSheet = Nothing
             
          End If
       
       End If
    
       vaSpread1.Col = 7
       NombreServicio = IIf(Trim(vaSpread1.text) = "", "", vaSpread1.text)
       
       '-------> Traer fecha minima - maxima
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Sql = ""
       Sql = Sql & " '" & LimpiaDato(Trim(Ceco)) & "'"
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
       Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
       Set RS = vg_db.Execute("sgpadm_Sel_MinBloqueMinMax " & Sql & "")

       If Not RS.EOF Then
          
          If IsNull(RS!FecMin) Then
             
             GoTo noexiste
          
          End If
          
          DiaColumna = DateDiff("d", CDate(fg_Ctod1(RS!FecMin)), CDate(fg_Ctod1(RS!FecMax))) + 1
          
          MaxColumna = 2
          
          '-------> Raciones
          If Check1.Value = 1 Then
             
             MaxColumna = MaxColumna + 1
          
          End If
          
          '-------> % Ponderación
          If Check2.Value = 1 Then
             
             MaxColumna = MaxColumna + 1
          
          End If
          
          MaxColumna = MaxColumna + 2

          MaxColumna = MaxColumna * DiaColumna
         
          Dim VecDiaExcel() As Variant
          ReDim VecDiaExcel(MaxColumna, 2)
          
          '-------> Setear vector
          For j = 1 To UBound(VecDiaExcel)
              
              VecDiaExcel(j, 1) = Val(0) 'fecha
              VecDiaExcel(j, 2) = "" 'descripción
          
          Next j
          
          FecMin = fg_Ctod1(RS!FecMin)
          FecMax = fg_Ctod1(RS!FecMax)
          
       End If
       RS.Close
       Set RS = Nothing
          
       '-------> Llamar procedimiento
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       RowEnd1 = 0
       vg_RowEnd = 0
       Sql = ""
       Sql = Sql & LimpiaDato(Trim(Ceco))
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
       Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
       Set RS = vg_db.Execute("sgpadm_Sel_MaxLineaMinBloqueExp " & Sql & "")
          
       If Not RS.EOF Then
             
          RowEnd1 = RS!NumLin
          vg_RowEnd = RS!NumLin
          
       End If
       RS.Close
       Set RS = Nothing
       
       '-------> Llamar procedimiento
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Sql = ""
       Sql = Sql & LimpiaDato(Trim(Ceco))
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
       Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
       Set RS = vg_db.Execute("sgpadm_Sel_DetalleMinBloqueExp_V02 " & Sql & "")
       
       If Not RS.EOF Then
          
          vaSpread1.Col = 9
          vaSpread1.text = "Servicio termino correctamente"
          
          '-------> Add data to cells of the first worksheet in the new workbook
          Set oSheet = Nothing
          Set oSheet = oBook.Worksheets.Add
          
          NombreServicioAux = ""
          NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
          
          If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
             
             NombreServicioAux = Mid(CodigoRegimen & "-" & CodigoServicio & NombreServicioAux, 1, 31)
          
          End If
          oSheet.Name = Trim(NombreServicioAux)
          
          '-------> Mover Ceco - Regimen
          MoverDatosExcel oExcel, oSheet, "A", "A", 1, 1, "C.Costo " & Trim(NomCeco) & " - " & Trim(Ceco)
          MoverDatosExcel oExcel, oSheet, "A", "A", 2, 2, "Regimen " & Trim(NombreRegimen) & " - " & CodigoRegimen

          '-------> Mover Nuevo formato
          MoverDatosExcel oExcel, oSheet, "B", "B", 1, 1, "*"

'          '-------> Costos Ppto/C.Comercial
'          MoverDatosExcel oExcel, oSheet, "A", "A", 5, 5, "Costo Ppto/C.Comercial"
'          DibujarLineas oExcel, oSheet, "A", "A", 5, 5
          
          '-------> Costos Costo Planificación
          MoverDatosExcel oExcel, oSheet, "A", "A", 6, 6, "Costo Planificación"
          DibujarLineas oExcel, oSheet, "A", "A", 6, 6
          
          '-------> Costos Costo Sitio
          MoverDatosExcel oExcel, oSheet, "A", "A", 7, 7, "Costo Sitio"
          DibujarLineas oExcel, oSheet, "A", "A", 7, 7

          '-------> Mover titulo excel
          MoverDatosExcel oExcel, oSheet, "A", "A", 8, 8, "Estructura Servicio"
          DibujarLineas oExcel, oSheet, "A", "A", 8, 8
          IndCol = 1
          IndColA = 65
          IndVec = 0
          oCol = ""
          oColA = ""
          oCol = Chr(IndCol + 65)
          IndVec = 1
          TotCol = 1
          
          Do While FecMin <= FecMax
                
             MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, Mid(fg_Fecha_Dia(Format(FecMin, "yyyymmdd"), 1), 1, 4) & " " & FecMin
             DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
             
             VecDiaExcel(IndVec, 1) = Format(FecMin, "yyyymmdd") 'fecha
             VecDiaExcel(IndVec, 2) = oCol 'descripción
             
             '-------> Raciones
             If Check1.Value = 1 Then
                   
                IndVec = IndVec + 1
                IndCol = IndCol + 1
                   
                If Chr(IndCol + 65) = "[" Then
                      
                   oColA = Chr(IndColA)
                   IndColA = IndColA + 1
                   IndCol = 0
                   
                End If
                   
                oCol = oColA & Chr(IndCol + 65)
                VecDiaExcel(IndVec, 2) = oCol 'descripción
                MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, "Rac."
                DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
                
             End If
             
             '-------> % Ponderación
             If Check2.Value = 1 Then
                   
                IndVec = IndVec + 1
                IndCol = IndCol + 1
                   
                If Chr(IndCol + 65) = "[" Then
                      
                   oColA = Chr(IndColA)
                   IndColA = IndColA + 1
                   IndCol = 0
                   
                End If
                   
                oCol = oColA & Chr(IndCol + 65)
                VecDiaExcel(IndVec, 2) = oCol 'descripción
                MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, "% Pond."
                DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
                
             End If
             
             '-------> Costo receta
             IndVec = IndVec + 1
             IndCol = IndCol + 1
                   
             If Chr(IndCol + 65) = "[" Then
                      
                oColA = Chr(IndColA)
                IndColA = IndColA + 1
                IndCol = 0
                   
             End If
                   
             oCol = oColA & Chr(IndCol + 65)
             VecDiaExcel(IndVec, 2) = oCol 'descripción
             MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, "Cto. Plato"
             DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
             
             '-------> Costo receta ponderado
             IndVec = IndVec + 1
             IndCol = IndCol + 1
                   
             If Chr(IndCol + 65) = "[" Then
                      
                oColA = Chr(IndColA)
                IndColA = IndColA + 1
                IndCol = 0
                   
             End If
                   
             oCol = oColA & Chr(IndCol + 65)
             VecDiaExcel(IndVec, 2) = oCol 'descripción
             MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, "Cto. Plato Pon."
             DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
             
             If Check3.Value = 1 Then
                
                IndVec = IndVec + 1
                IndCol = IndCol + 1
                   
                If Chr(IndCol + 65) = "[" Then
                      
                   oColA = Chr(IndColA)
                   IndColA = IndColA + 1
                   IndCol = 0
                   
                End If
                   
                oCol = oColA & Chr(IndCol + 65)
                VecDiaExcel(IndVec, 2) = oCol 'descripción
                MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, "Clave"
                DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
             
             End If
             
             FecMin = FecMin + 1
             IndCol = IndCol + 1
             IndVec = IndVec + 1
             
             If Chr(IndCol + 65) = "[" Then
                   
                oColA = Chr(IndColA)
                IndColA = IndColA + 1
                IndCol = 0
                
             End If
                
             oCol = oColA & Chr(IndCol + 65)

             TotCol = TotCol + 1
             
          Loop
          
          RowSheet = 8
          RowEnd = 0
          AuxFec = 0
          CodEstructura = 0
          EstVacio = True
             
          Do While Not RS.EOF
             
             LyD = RS!ser_LYD
             '-------> Corte x fecha
             If AuxFec <> RS!min_fecmin Then
                
                '-------> Mover comensales totales
                If AuxFec > 0 And Check1.Value = 1 Then
                      
                   MoverDatosExcel oExcel, oSheet, "A", "A", RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, "Comensales"
                   MoverDatosExcel oExcel, oSheet, ColEnd, ColEnd, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, CStr(TotDiaRaciones)
                   
                   CalComen = CalComen & ColEnd & RowSheet + RowEnd1 + 2 & "+"
                   
                   If Check3.Value = 1 Then
                         
                      MoverDatosExcel oExcel, oSheet, oColCla, oColCla, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, LimpiaDato(Trim(Ceco)) & ";" & CodigoRegimen & ";" & CodigoServicio & ";" & CStr(AuxFec)
                      
                   End If
                   
                   If Check2.Value = 1 And Check3.Value = 1 Then
                      
                      '-------> Buscar dia vector dia excel
                      For j = 1 To UBound(VecDiaExcel)
                        
                         If ColEnd = VecDiaExcel(j, 2) Then
                            
                            oColFec = VecDiaExcel(j - 1, 2)
                            oColPor = VecDiaExcel(j + 1, 2)
                            oColCos = VecDiaExcel(j + 2, 2)
                            oColCpo = VecDiaExcel(j + 3, 2)
                            Exit For
                           
                         End If
                      
                      Next j
                        
                      EstVacio = True
                      For ii = 9 To RowSheet + RowEnd1 + 2
                          
                          If (Option1(0).Value = True Or LyD) Then
                          
                             MoverDatosExcelFormulaRaciones oExcel, oSheet, oColFec, oColPor, ColEnd, ColEnd, ii, RowSheet + RowEnd1 + 2, 0, EstVacio
                          
                          ElseIf Option1(1).Value = True Then
                          
                             MoverDatosExcelFormulaPonderacion oExcel, oSheet, oColFec, oColPor, ColEnd, ColEnd, ii, RowSheet + RowEnd1 + 2, 0, EstVacio
                          
                          End If
                          
                          MoverDatosExcelFormulaII oExcel, oSheet, oColFec, oColCpo, ColEnd, oColCos, ii, RowSheet + RowEnd1 + 2, 0, EstVacio
                          
                          If ii < RowSheet + RowEnd1 + 2 Then
                             
                             If Trim(oSheet.Range(ColEnd & ii).Value) <> "" Then
                                
                                CalRacCos = CalRacCos & "iferror(" & ColEnd & ii & "*" & oColCos & ii & ",0)" & "+"
                                
                             End If
                             
                          End If
                          
                          EstVacio = False
                          
                      Next ii
                      
                      EstVacio = True
                      MoverDatosExcelFormulaSum oExcel, oSheet, oColCpo, oColCpo, oColCpo, oColCpo, RowSheet + 1, RowSheet + RowEnd1 + 2, ""
'                      MoverDatosExcel oExcel, oSheet, oColCpo, oColCpo, 5, 5, CStr(CostoComercial)
                      MoverDatosExcel oExcel, oSheet, oColCpo, oColCpo, 6, 6, CStr(Format(CostoPlanificacion, fg_Pict(6, 2)))
                      MoverDatosExcelCostoBandeja oExcel, oSheet, oColCos, RowSheet + RowEnd1 + 4, "(" & Mid(CalRacCos, 1, Len(CalRacCos) - 1) & ")"
                      TotCalRacCos = TotCalRacCos & oColCos & RowSheet + RowEnd1 + 4 & "+"
                      CalRacCos = ""
                      CostoPlanificacion = 0
                                          
                      If (Option1(0).Value = True Or LyD) Then
                         
                         BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + 1, RowSheet + RowEnd1 '+ 2
                         BloquearColumnaExcel oExcel, oSheet, oColFec, RowSheet + 1, RowSheet + RowEnd1 '+ 2
                         'habilitar columna c comensales
                         BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2
                         'poner color gris
                         PonerColorGrisExcel oExcel, oSheet, ColEnd, RowSheet + 1, RowSheet + RowEnd1 '+ 2
                      
                      ElseIf Option1(1).Value = True Then
                      
                         'habilitar %
                         BloquearColumnaExcel oExcel, oSheet, oColPor, RowSheet + 1, RowSheet + RowEnd1 '+ 2
                         'habilitar receta
                         BloquearColumnaExcel oExcel, oSheet, oColFec, RowSheet + 1, RowSheet + RowEnd1 '+ 2
                         'habilitar columna c comensales
                         BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2
                         'poner color gris
                         PonerColorGrisExcel oExcel, oSheet, oColPor, RowSheet + 1, RowSheet + RowEnd1 '+ 2
                      
                      End If
                      
                      FormatearColumnaNumericoExcel oExcel, oSheet, ColEnd
                      FormatearColumnaPorcentajeExcel oExcel, oSheet, oColPor
                       
                    End If
                    
                    RowEnd = 0
                
                End If
                
                EstVacio = True
                RowSheet = 8
                AuxFec = RS!min_fecmin
                TotDiaRaciones = RS!min_racteo
                Comensales = Comensales + RS!min_racteo
             
             End If
             
             '-------> Buscar dia vector dia excel
             For j = 1 To UBound(VecDiaExcel)
                 
                 If RS!min_fecmin = Val(VecDiaExcel(j, 1)) Then
                       
                    oCol = VecDiaExcel(j, 2)
                    IndCol = j
                    
                    If Check3.Value = 1 Then
                          
                       oColCla = VecDiaExcel(j + 5, 2)
                       
                    End If
                    
                    If Check1.Value = 1 Then
                          
                       ColEnd = VecDiaExcel(j + 1, 2)
                       
                    End If
                    Exit For
                    
                 End If
             
             Next j
             
             '-------> Corte x estructura servicio
             '-------> Mover Estructura servicio excel
             If CodEstructura <> RS!ess_codigo Then
                
                CodEstructura = RS!ess_codigo
                MoverDatosExcel oExcel, oSheet, "A", "A", RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, Trim(RS!ess_nombre) & ";" & CodEstructura
                             
             End If

             If Check3.Value = 1 Then
                
                If EstVacio Then
                
                   MoverDatosExcelClave oExcel, oSheet, oColCla, oColCla, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, LimpiaDato(Trim(Ceco)) & ";" & CodigoRegimen & ";" & CodigoServicio & ";" & 0 & ";" & RS!min_fecmin & ";" & RS!mid_numlin
                
                End If
                
                MoverDatosExcel oExcel, oSheet, oColCla, oColCla, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, LimpiaDato(Trim(Ceco)) & ";" & CodigoRegimen & ";" & CodigoServicio & ";" & RS!mid_codrec & ";" & RS!min_fecmin & ";" & RS!mid_numlin
             
             End If
             
             '-------> Mover recetas excel
             MoverDatosExcel oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, Trim(RS!rec_nombre & " " & RS!mid_codrec)
             
             '-------> Mover raciones excel
             If Check1.Value = 1 Then
                
                oCol = VecDiaExcel(IndCol + 1, 2)
                IndCol = IndCol + 1
                MoverDatosExcelValorNumerico oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, RS!mid_numrac, EstVacio
             
             End If
             
             '-------> Mover % ponderación excel
             If Check2.Value = 1 Then
                
                oCol = VecDiaExcel(IndCol + 1, 2)
                MoverDatosExcelValorNumerico oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, RS!mid_porrac & " %", EstVacio
             
             End If
             
             '-------> Mover costo receta
             If LinRecetaTotal > 0 Then
                
                oCol = VecDiaExcel(IndCol + 2, 2)
                oColRec = VecDiaExcel(IndCol - 1, 2)
                MoverDatosExcelBuscarV oExcel, oSheet, oColRec, oCol, RowSheet + RS!mid_numlin, LinRecetaTotal, EstVacio
             
             End If
             
             EstVacio = False
             
             '-------> Mover costo ponderado receta
             oCol = VecDiaExcel(IndCol + 3, 2)
             
             If RS!min_racteo > 0 Then
                
                CostoRecetaDet = 0
                CostoRecetaDet = BuscarCostoReceta(RS!mid_codrec)
                CostoPlanificacion = CostoPlanificacion + ((CostoRecetaDet * RS!mid_numrac) / RS!min_racteo)
                CostoReceta = CostoReceta + (CostoRecetaDet * RS!mid_numrac)
                
             End If
             
             If RS!mid_numlin > RowEnd Then
                
                RowEnd = RS!mid_numlin
                   
             End If
             
             RS.MoveNext
             
          Loop
          
          RS.Close
          Set RS = Nothing
          
          '-------> Mover comensales totales
          If AuxFec > 0 And Check1.Value = 1 Then
             
             MoverDatosExcel oExcel, oSheet, "A", "A", RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, "Comensales"
             MoverDatosExcel oExcel, oSheet, ColEnd, ColEnd, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, CStr(TotDiaRaciones)
          
             CalComen = CalComen & ColEnd & RowSheet + RowEnd1 + 2 & "+"
              
             If Check3.Value = 1 Then
                
                MoverDatosExcel oExcel, oSheet, oColCla, oColCla, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, LimpiaDato(Trim(Ceco)) & ";" & CodigoRegimen & ";" & CodigoServicio & ";" & CStr(AuxFec)
           
             End If
            
             If Check2.Value = 1 And Check3.Value = 1 Then
             
             '-------> Buscar dia vector dia excel
             For j = 1 To UBound(VecDiaExcel)
                 
                 If ColEnd = VecDiaExcel(j, 2) Then
                    
                    oColFec = VecDiaExcel(j - 1, 2)
                    oColPor = VecDiaExcel(j + 1, 2)
                    oColCos = VecDiaExcel(j + 2, 2)
                    oColCpo = VecDiaExcel(j + 3, 2)
                    Exit For
                 
                 End If
             
             Next j
                   
             EstVacio = True
             For ii = 9 To RowSheet + RowEnd1 + 2
                 
                 If (Option1(0).Value = True Or LyD) Then
                          
                    MoverDatosExcelFormulaRaciones oExcel, oSheet, oColFec, oColPor, ColEnd, ColEnd, ii, RowSheet + RowEnd1 + 2, 0, EstVacio
                          
                 ElseIf Option1(1).Value = True Then
                          
                    MoverDatosExcelFormulaPonderacion oExcel, oSheet, oColFec, oColPor, ColEnd, ColEnd, ii, RowSheet + RowEnd1 + 2, 0, EstVacio
                          
                 End If

                 MoverDatosExcelFormulaII oExcel, oSheet, oColFec, oColCpo, ColEnd, oColCos, ii, RowSheet + RowEnd1 + 2, 0, EstVacio
             
                 If ii < RowSheet + RowEnd1 + 2 Then
                          
                    If Trim(oSheet.Range(ColEnd & ii).Value) <> "" Then
                    
                       CalRacCos = CalRacCos & ColEnd & ii & "*" & oColCos & ii & "+"
                    
                    End If
                    
                 End If
                 
                 EstVacio = False
                 
             Next ii
          
             EstVacio = True
             MoverDatosExcelFormulaSum oExcel, oSheet, oColCpo, oColCpo, oColCpo, oColCpo, RowSheet + 1, RowSheet + RowEnd1 + 2, ""
'             MoverDatosExcel oExcel, oSheet, oColCpo, oColCpo, 5, 5, CStr(CostoComercial)
             MoverDatosExcel oExcel, oSheet, oColCpo, oColCpo, 6, 6, CStr(Format(CostoPlanificacion, fg_Pict(6, 2)))
             MoverDatosExcelCostoBandeja oExcel, oSheet, oColCos, RowSheet + RowEnd1 + 4, "(" & Mid(CalRacCos, 1, Len(CalRacCos) - 1) & ")"
             TotCalRacCos = TotCalRacCos & oColCos & RowSheet + RowEnd1 + 4 & "+"
             CalRacCos = ""
             
             CostoPlanificacion = 0
             
             If (Option1(0).Value = True Or LyD) Then
                         
                BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + 1, RowSheet + RowEnd1 '+ 2
                BloquearColumnaExcel oExcel, oSheet, oColFec, RowSheet + 1, RowSheet + RowEnd1 '+ 2
                'habilitar columna c comensales
                BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2
                'poner color gris
                PonerColorGrisExcel oExcel, oSheet, ColEnd, RowSheet + 1, RowSheet + RowEnd1 '+ 2
                      
             ElseIf Option1(1).Value = True Then
                      
                'habilitar %
                BloquearColumnaExcel oExcel, oSheet, oColPor, RowSheet + 1, RowSheet + RowEnd1 '+ 2
                'habilitar receta
                BloquearColumnaExcel oExcel, oSheet, oColFec, RowSheet + 1, RowSheet + RowEnd1 '+ 2
                'habilitar columna c comensales
                BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2
                'poner color gris
                PonerColorGrisExcel oExcel, oSheet, oColPor, RowSheet + 1, RowSheet + RowEnd1 '+ 2
                      
             End If
                      
             FormatearColumnaNumericoExcel oExcel, oSheet, ColEnd
             FormatearColumnaPorcentajeExcel oExcel, oSheet, oColPor
                       
             End If
                    
'             If (Option1(0).Value = True Or LyD) Then
'
'                BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + 1, RowSheet + RowEnd1 + 2
'                BloquearColumnaExcel oExcel, oSheet, oColFec, RowSheet + 1, RowSheet + RowEnd1 + 2
'
'             ElseIf Option1(1).Value = True Then
'
''                 BloquearColumnaExcel oExcel, oSheet, oColPor, RowSheet + 1, RowSheet + RowEnd1 + 2
''                 BloquearColumnaExcel oExcel, oSheet, oColFec, RowSheet + 1, RowSheet + RowEnd1 + 2
'                'habilitar %
'                BloquearColumnaExcel oExcel, oSheet, oColPor, RowSheet + 1, RowSheet + RowEnd1 '+ 2
'                'habilitar receta
'                BloquearColumnaExcel oExcel, oSheet, oColFec, RowSheet + 1, RowSheet + RowEnd1 '+ 2
'                'habilitar columna c comensales
'                BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2
'             End If
'
'             FormatearColumnaNumericoExcel oExcel, oSheet, ColEnd
'             FormatearColumnaPorcentajeExcel oExcel, oSheet, oColPor
'
'            End If
            
             End If
          
             '-------> Dibujar lineas
             DibujarLineas oExcel, oSheet, "A", "A", 8, RowSheet + RowEnd1 + 2
             For j = 1 To UBound(VecDiaExcel)
              
                 oCol = VecDiaExcel(j, 2)
                 DibujarLineas oExcel, oSheet, oCol, oCol, 5, RowSheet + RowEnd1 + 2
          
             Next j
                      
             oSheet.Cells.Select
             oSheet.Cells.EntireColumn.AutoFit
             RowSheet = 8

             'Ocultar columna clave
             If Check3.Value = 1 Then
          
                MoverDatosExcel oExcel, oSheet, "B", "B", 2, 2, "Costo Bandeja Planificado"
                MoverDatosExcel oExcel, oSheet, "B", "B", 3, 3, "Costo Bandeja Sitio"
                
                Select Case Comensales
                Case Is > 0
                   
                   MoverDatosExcel oExcel, oSheet, "F", "F", 2, 2, Format((CostoReceta / Comensales), fg_Pict(6, 2))
                
                Case Is < 0
                
                   MoverDatosExcel oExcel, oSheet, "F", "F", 2, 2, 0
                
                End Select
                
                MoverDatosExcelCostoBandeja oExcel, oSheet, "F", 3, "(" & Mid(TotCalRacCos, 1, Len(TotCalRacCos) - 1) & ")" & "/" & "(" & Mid(CalComen, 1, Len(CalComen) - 1) & ")"
                
                For j = 1 To UBound(VecDiaExcel) Step 6 '4
                 
                    oCol = VecDiaExcel(j + 5, 2) 'VecDiaExcel(j + 3, 2)
                    OcultarColumna oExcel, oSheet, oCol, oCol
                 
                    'Mover clave para poder actualizar
                 
                    MoverDatosExcel oExcel, oSheet, oCol, oCol, 1, 1, ClaveExcel
                    MoverDatosExcel oExcel, oSheet, oCol, oCol, 2, 2, CStr(((TotCol - 1) * 6) + 1)
             
                Next j
             
             End If
          
             'Bloquear protección hoja
             If Check2.Value = 1 And Check3.Value = 1 Then
          
'                If oSheet.Range("a9").Value = "" Then
'
'                   oSheet.Rows("9:9").Select
'                   oExcel.Selection.Delete Shift:=xlUp
'
'                End If
                
                oSheet.Select
                oExcel.ActiveSheet.Protect password:=ClaveExcel, DrawingObjects:=True, _
                                    Contents:=True, Scenarios:=True, AllowFormattingCells:=True
                                 
             End If
             
                         
             fg_descarga
         
         
          Else
          
noexiste:
             RS.Close
             Set RS = Nothing
             fg_descarga
             vaSpread1.Col = 9
             vaSpread1.text = "No existe información"
             auxceco = ""
       
          End If
          
    End If

Next i

If Trim(auxceco) <> "" Then
   
   If Check2.Value = 1 And Check3.Value = 1 Then
             
'      If oSheet.Range("a9").Value = "" Then
'
'         oSheet.Rows("9:9").Select
'         oExcel.Selection.Delete Shift:=xlUp
'
'      End If
                
      oExcel.Sheets("Recetas").Select
      oExcel.Sheets("Recetas").Move Before:=oExcel.Sheets(1)
    
      oSheet.Select
      oExcel.ActiveSheet.Protect password:=ClaveExcel, DrawingObjects:=True, _
                Contents:=True, Scenarios:=True, AllowFormattingCells:=True
          
   End If
   
   oBook.Close True, NomArchivoExcel

   Set oSheet = Nothing
   Set oBook = Nothing
   Set oExcel = Nothing
   
    '------- Copiar archivos en la ruta seleccionada, luego borrar archivos de la carpeta
    If Trim(NomArchivoExcel) <> Trim(fpText1.text & SoloNomArchivoExcel) Then fso.CopyFile NomArchivoExcel, fpText1.text & SoloNomArchivoExcel, True
    If Dir(NomArchivoExcel) <> "" And Trim(NomArchivoExcel) <> Trim(fpText1.text & SoloNomArchivoExcel) Then Kill NomArchivoExcel
       
End If

'oExcel.Visible = True '------->Visualizar

Bar1(0).Value = 0
Bar1(0).Visible = False

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
End Sub

Sub ExportarExcelMinutaBloqueSinCostoReceta()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim RS1             As New ADODB.Recordset
Dim Sql             As String
Dim Est             As Boolean
Dim i               As Long
Dim j               As Long
Dim ii              As Long
Dim NomArchivoExcel As String

'-------> Validar Org. Compras
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
            
Set RS = vg_db.Execute("sgpadm_Sel_OrgCompras_V02 '" & LimpiaDato(fpText.text) & "'")
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   MsgBox "No existe Org. Compras...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   vaSpread1.MaxRows = 0
   Exit Sub

End If
RS.Close
Set RS = Nothing

'-------> validar seleción ceco - regimen - servicio
Est = False

For i = 1 To vaSpread1.MaxRows
    
    vaSpread1.Row = i
    vaSpread1.Col = 1
    
    If vaSpread1.text = "1" Then
       
       Est = True
    
    End If

Next i
If Not Est Then MsgBox "Regimen y servicios asociado debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub

'-------> Validar fecha desde - hasta
If CDate(FpFecDesde.text) > CDate(FpFecHasta.text) Then
   
   Call MsgBox("Fecha Desde No Puede Ser Mayor a Fecha Hasta", vbInformation, MsgTitulo)
   Let FpFecDesde.text = Format(Now, "dd/mm/yyyy")
   Call FpFecDesde.SetFocus
   Exit Sub

End If
    
If CDate(FpFecHasta.text) < CDate(FpFecDesde.text) Then
   
   Call MsgBox("Fecha Hasta No Puede Ser Mayor a Fecha Desde", vbInformation, MsgTitulo)
   Let FpFecHasta.text = Format(Now, "dd/mm/yyyy")
   Call FpFecHasta.SetFocus
   Exit Sub

End If

'If DateDiff("ww", FpFecDesde.text, FpFecHasta.text) > 52 Then
If DateDiff("y", FpFecDesde.text, FpFecHasta.text) > 365 Then
   
   Call MsgBox("Rango De Fecha No Puede Ser Mayor a 12 Meses", vbInformation, Me.Caption)
   Exit Sub

End If

Dim oExcel              As Object
Dim oBook               As Object
Dim oSheet              As Object
Dim RowSheet            As Long
Dim oCol                As String
Dim oColCla             As String
Dim oColA               As String
Dim oColFec             As String
Dim oColPor             As String
Dim Ceco                As String
Dim NomCeco             As String
Dim auxceco             As String
Dim NombreServicioAux   As String
Dim NombreServicio      As String
Dim CodigoServicio      As Long
Dim CodigoRegimen       As Long
Dim NombreRegimen       As String
Dim MaxColumna          As Long
Dim NumAsc              As Long
Dim FecMin              As Date
Dim FecMax              As Date
Dim IndCol              As Long
Dim IndColA             As Long
Dim IndVec              As Long
Dim AuxFec              As Long
Dim CodEstructura       As Long
Dim RowEnd              As Long
Dim RowEnd1             As Long
Dim ColEnd              As String
Dim TotDiaRaciones      As Double
Dim TotCol              As Long

Dim ClaveExcel          As String
Dim SoloNomArchivoExcel As String
Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

fg_carga ""
Bar1(0).Visible = True
Bar1(0).Value = 0

For i = 1 To vaSpread1.MaxRows
    
    DoEvents
    
    vaSpread1.Row = i
    vaSpread1.Col = 1
    Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
    
    If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
       
       vaSpread1.SetActiveCell 2, vaSpread1.Row
       
       vaSpread1.Col = 2
       Ceco = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
       
       vaSpread1.Col = 3
       NomCeco = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
       
       If Ceco <> auxceco Then
       
          If Trim(auxceco) <> "" Then
             
             '-------> incluir recetas
             If Check5.Value = 1 Then
             
                If RS.State = 1 Then RS.Close
                RS.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Sql = ""
                Sql = Sql & " '" & LimpiaDato(Trim(Ceco)) & "' "
                Set RS = vg_db.Execute("sgpadm_Sel_ExportarExcelRecetas_V01 " & Sql & "")
                If Not RS.EOF Then
          
                    '-------> Add data to cells of the first worksheet in the new workbook
                    NombreServicio = "Recetas"
                    Set oSheet = oBook.Worksheets.Add
                    NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
          
                    If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
             
                        NombreServicioAux = Mid(NombreServicioAux, 1, 31)
          
                    End If
                    oSheet.Name = Trim(NombreServicioAux)
             
                    IndCol = 1
             
                    '-------> Check version of Excel
                    Call encabezado(RS, oSheet)
          
                    oSheet.Cells(2, 1).CopyFromRecordset RS
                    
                    MoverDatosExcel oExcel, oSheet, "A", "A", IndCol, IndCol, Trim("Código")
                    MoverDatosExcel oExcel, oSheet, "B", "B", IndCol, IndCol, Trim("Nombre Plato")
                    MoverDatosExcel oExcel, oSheet, "C", "C", IndCol, IndCol, Trim("Categoria Dietetica")
                    MoverDatosExcel oExcel, oSheet, "D", "D", IndCol, IndCol, Trim("Tipo Plato")
                    PonerColorInteriorN oExcel, oSheet, "A", "D", IndCol, IndCol, 4
                
                    IndCol = 2
             
                    PonerColorInteriorN oExcel, oSheet, "A", "D", IndCol, RS.RecordCount + 1, 6
                
                    RS.Close
                    Set RS = Nothing
          
                    oSheet.Cells.Select
                    oSheet.Cells.EntireColumn.AutoFit
             
                    'Bloquear protección hoja
                    If Check2.Value = 1 And Check3.Value = 1 Then
             
                        With oSheet

                        .AutoFilterMode = False

                        .Range("A1:D1").AutoFilter

                        End With
             
                    End If
          
                End If
             
             
             End If
             
             oBook.Close True, NomArchivoExcel
'             oExcel.Visible = True '------->Visualizar
             
             Set oSheet = Nothing
             Set oExcel = Nothing
             Set oBook = Nothing
          
             '------- Copiar archivos en la ruta seleccionada, luego borrar archivos de la carpeta
             If Trim(NomArchivoExcel) <> Trim(fpText1.text & SoloNomArchivoExcel) Then fso.CopyFile NomArchivoExcel, fpText1.text & SoloNomArchivoExcel, True
             If Dir(NomArchivoExcel) <> "" And Trim(NomArchivoExcel) <> Trim(fpText1.text & SoloNomArchivoExcel) Then Kill NomArchivoExcel
             Me.Refresh
             
          End If
          
          '-------> Exportar excel
          Set oExcel = CreateObject("Excel.Application")
          Set oBook = oExcel.Workbooks.Add
          NumAsc = 66
       
          auxceco = Ceco
          
          NomArchivoExcel = dir_trabajo & "ExcelMinutaSGP" & "\" & Ceco & "-" & NomCeco & " " & Format(FpFecDesde.text, "yyyymmdd") & " " & Format(FpFecHasta.text, "yyyymmdd") & " " & Replace(Date, "/", "") & " " & Replace(Time, ":", "") & ".xlsx"
          SoloNomArchivoExcel = Ceco & "-" & NomCeco & " " & Format(FpFecDesde.text, "yyyymmdd") & " " & Format(FpFecHasta.text, "yyyymmdd") & " " & Replace(Date, "/", "") & " " & Replace(Time, ":", "") & ".xlsx"
          
       End If
       
       vaSpread1.Col = 4: CodigoRegimen = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
       vaSpread1.Col = 5: NombreRegimen = IIf(Trim(vaSpread1.text) = "", "", vaSpread1.text)
       vaSpread1.Col = 6: CodigoServicio = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
       vaSpread1.Col = 7: NombreServicio = IIf(Trim(vaSpread1.text) = "", "", vaSpread1.text)
    
       '-------> Traer fecha minima - maxima
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Sql = ""
       Sql = Sql & " '" & LimpiaDato(Trim(Ceco)) & "'"
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
       Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
       Set RS = vg_db.Execute("sgpadm_Sel_MinBloqueMinMax " & Sql & "")

       If Not RS.EOF Then
          
          If IsNull(RS!FecMin) Then
             
'             RS.Close
'             Set RS = Nothing
'             MsgBox "No existe información, para este regimen y servicio", vbExclamation + vbOKOnly, MsgTitulo
'             Exit Sub
              
              GoTo noexiste
          
          End If
          
          MaxColumna = DateDiff("d", CDate(fg_Ctod1(RS!FecMin)), CDate(fg_Ctod1(RS!FecMax))) + 1
          
          '-------> Raciones
          If Check1.Value = 1 Then
             
             MaxColumna = MaxColumna * 2
          
          End If
          
          '-------> % Ponderación
          If Check2.Value = 1 Then
             
             MaxColumna = MaxColumna * 2
          
          End If
          
          Dim VecDiaExcel() As Variant
          ReDim VecDiaExcel(MaxColumna, 2)
          
          '-------> Setear vector
          For j = 1 To UBound(VecDiaExcel)
              
              VecDiaExcel(j, 1) = Val(0) 'fecha
              VecDiaExcel(j, 2) = "" 'descripción
          
          Next j
          
          FecMin = fg_Ctod1(RS!FecMin)
          FecMax = fg_Ctod1(RS!FecMax)
          
       End If
       RS.Close
       Set RS = Nothing
          
       '-------> Llamar procedimiento
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       RowEnd1 = 0
       Sql = ""
       Sql = Sql & LimpiaDato(Trim(Ceco))
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
       Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
       Set RS = vg_db.Execute("sgpadm_Sel_MaxLineaMinBloqueExp " & Sql & "")
          
       If Not RS.EOF Then
             
          RowEnd1 = RS!NumLin
          
       End If
       RS.Close
       Set RS = Nothing
       
       '-------> Llamar procedimiento
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Sql = ""
       Sql = Sql & LimpiaDato(Trim(Ceco))
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
       Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
       Set RS = vg_db.Execute("sgpadm_Sel_DetalleMinBloqueExpSinCostoReceta " & Sql & "")
       
       If Not RS.EOF Then
          
          vaSpread1.Col = 9
          vaSpread1.text = "Servicio termino correctamente"
          
          '-------> Add data to cells of the first worksheet in the new workbook
          Set oSheet = oBook.Worksheets.Add
          NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
          
          If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
             
             NombreServicioAux = Mid(CodigoRegimen & "-" & CodigoServicio & NombreServicioAux, 1, 31)
          
          End If
          oSheet.Name = Trim(NombreServicioAux)
          
          '-------> Mover Ceco - Regimen
          MoverDatosExcel oExcel, oSheet, "A", "A", 1, 1, "C.Costo " & Trim(NomCeco) & " - " & Trim(Ceco)
          MoverDatosExcel oExcel, oSheet, "A", "A", 2, 2, "Regimen " & Trim(NombreRegimen) & " - " & CodigoRegimen

          '-------> Mover titulo excel
          MoverDatosExcel oExcel, oSheet, "A", "A", 5, 5, "Estructura Servicio"
          DibujarLineas oExcel, oSheet, "A", "A", 5, 5
          IndCol = 1
          IndColA = 65
          IndVec = 0
          oCol = ""
          oColA = ""
          oCol = Chr(IndCol + 65)
          IndVec = 1
          TotCol = 1
          Do While FecMin <= FecMax
                
             MoverDatosExcel oExcel, oSheet, oCol, oCol, 5, 5, Mid(fg_Fecha_Dia(Format(FecMin, "yyyymmdd"), 1), 1, 4) & " " & FecMin
             DibujarLineas oExcel, oSheet, oCol, oCol, 5, 5
             
             VecDiaExcel(IndVec, 1) = Format(FecMin, "yyyymmdd") 'fecha
             VecDiaExcel(IndVec, 2) = oCol 'descripción
             
             '-------> Raciones
             If Check1.Value = 1 Then
                   
                IndVec = IndVec + 1
                IndCol = IndCol + 1
                   
                If Chr(IndCol + 65) = "[" Then
                      
                   oColA = Chr(IndColA)
                   IndColA = IndColA + 1
                   IndCol = 0
                   
                End If
                   
                oCol = oColA & Chr(IndCol + 65)
                VecDiaExcel(IndVec, 2) = oCol 'descripción
                MoverDatosExcel oExcel, oSheet, oCol, oCol, 5, 5, "Rac."
                DibujarLineas oExcel, oSheet, oCol, oCol, 5, 5
                
             End If
             
             '-------> % Ponderación
             If Check2.Value = 1 Then
                   
                IndVec = IndVec + 1
                IndCol = IndCol + 1
                   
                If Chr(IndCol + 65) = "[" Then
                      
                   oColA = Chr(IndColA)
                   IndColA = IndColA + 1
                   IndCol = 0
                   
                End If
                   
                oCol = oColA & Chr(IndCol + 65)
                VecDiaExcel(IndVec, 2) = oCol 'descripción
                MoverDatosExcel oExcel, oSheet, oCol, oCol, 5, 5, "% Pond."
                DibujarLineas oExcel, oSheet, oCol, oCol, 5, 5
                
             End If
             
             If Check3.Value = 1 Then
                
                IndVec = IndVec + 1
                IndCol = IndCol + 1
                   
                If Chr(IndCol + 65) = "[" Then
                      
                   oColA = Chr(IndColA)
                   IndColA = IndColA + 1
                   IndCol = 0
                   
                End If
                   
                oCol = oColA & Chr(IndCol + 65)
                VecDiaExcel(IndVec, 2) = oCol 'descripción
                MoverDatosExcel oExcel, oSheet, oCol, oCol, 5, 5, "Clave"
                DibujarLineas oExcel, oSheet, oCol, oCol, 5, 5
             
             End If
             
             FecMin = FecMin + 1
             IndCol = IndCol + 1
             IndVec = IndVec + 1
             
             If Chr(IndCol + 65) = "[" Then
                   
                oColA = Chr(IndColA)
                IndColA = IndColA + 1
                IndCol = 0
                
             End If
                
             oCol = oColA & Chr(IndCol + 65)

             TotCol = TotCol + 1
             
          Loop
          
          RowSheet = 5
          RowEnd = 0
          AuxFec = 0
          CodEstructura = 0
             
          Do While Not RS.EOF
             
             '-------> Corte x fecha
             If AuxFec <> RS!min_fecmin Then
                
                '-------> Mover comensales totales
                If AuxFec > 0 And Check1.Value = 1 Then
                      
                   MoverDatosExcel oExcel, oSheet, "A", "A", RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, "Comensales"
                   MoverDatosExcel oExcel, oSheet, ColEnd, ColEnd, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, CStr(TotDiaRaciones)
                   
                   If Check3.Value = 1 Then
                         
                      MoverDatosExcel oExcel, oSheet, oColCla, oColCla, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, LimpiaDato(Trim(Ceco)) & ";" & CodigoRegimen & ";" & CodigoServicio & ";" & CStr(AuxFec)
                      
                   End If
                   
                   If Check2.Value = 1 And Check3.Value = 1 Then
                      
                      '-------> Buscar dia vector dia excel
                      For j = 1 To UBound(VecDiaExcel)
                        
                         If ColEnd = VecDiaExcel(j, 2) Then
                            
                            oColFec = VecDiaExcel(j - 1, 2)
                            oColPor = VecDiaExcel(j + 1, 2)
                            Exit For
                           
                         End If
                      
                      Next j
                   
                      For ii = 6 To RowSheet + RowEnd + 2
                          
                          MoverDatosExcelFormula oExcel, oSheet, oColFec, oColPor, ColEnd, ColEnd, ii, RowSheet + RowEnd1 + 2, 0
                      
                      Next ii

                      BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + 1, RowSheet + RowEnd1 + 2
                      BloquearColumnaExcel oExcel, oSheet, oColFec, RowSheet + 1, RowSheet + RowEnd1 + 2
                      FormatearColumnaNumericoExcel oExcel, oSheet, ColEnd
                      FormatearColumnaPorcentajeExcel oExcel, oSheet, oColPor
                       
                    End If
                    
                    RowEnd = 0
                
                End If
                
                RowSheet = 5
                AuxFec = RS!min_fecmin
                TotDiaRaciones = RS!min_racteo
             
             End If
             
             '-------> Buscar dia vector dia excel
             For j = 1 To UBound(VecDiaExcel)
                 
                 If RS!min_fecmin = Val(VecDiaExcel(j, 1)) Then
                       
                    oCol = VecDiaExcel(j, 2)
                    IndCol = j
                    
                    If Check3.Value = 1 Then
                          
                       oColCla = VecDiaExcel(j + 3, 2)
                       
                    End If
                    
                    If Check1.Value = 1 Then
                          
                       ColEnd = VecDiaExcel(j + 1, 2)
                       
                    End If
                    Exit For
                    
                 End If
             
             Next j
             
             '-------> Corte x estructura servicio
             '-------> Mover Estructura servicio excel
             If CodEstructura <> RS!ess_codigo Then
                
                MoverDatosExcel oExcel, oSheet, "A", "A", RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, Trim(RS!ess_nombre)
                CodEstructura = RS!ess_codigo
             
             End If

             If Check3.Value = 1 Then
                
                MoverDatosExcel oExcel, oSheet, oColCla, oColCla, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, LimpiaDato(Trim(Ceco)) & ";" & CodigoRegimen & ";" & CodigoServicio & ";" & RS!mid_codrec & ";" & RS!min_fecmin & ";" & RS!mid_numlin
             
             End If
             
             '-------> Mover recetas excel
             MoverDatosExcel oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, Trim(RS!rec_nombre & " " & RS!mid_codrec)
             
             '-------> Mover raciones excel
             If Check1.Value = 1 Then
                
                oCol = VecDiaExcel(IndCol + 1, 2)
                IndCol = IndCol + 1
                MoverDatosExcel oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, RS!mid_numrac
             
             End If
             
             '-------> Mover % ponderación excel
             If Check2.Value = 1 Then
                
                oCol = VecDiaExcel(IndCol + 1, 2)
                MoverDatosExcel oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, RS!mid_porrac & " %"
             
             End If
             
             If RS!mid_numlin > RowEnd Then
                
                RowEnd = RS!mid_numlin
                   
             End If
             
             RS.MoveNext
             
          Loop
          RS.Close
          Set RS = Nothing
          
          '-------> Mover comensales totales
          If AuxFec > 0 And Check1.Value = 1 Then
             
             MoverDatosExcel oExcel, oSheet, "A", "A", RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, "Comensales"
             MoverDatosExcel oExcel, oSheet, ColEnd, ColEnd, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, CStr(TotDiaRaciones)
          
             If Check3.Value = 1 Then
                
                MoverDatosExcel oExcel, oSheet, oColCla, oColCla, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, LimpiaDato(Trim(Ceco)) & ";" & CodigoRegimen & ";" & CodigoServicio & ";" & CStr(AuxFec)
           
             End If
            
             If Check2.Value = 1 And Check3.Value = 1 Then
             
             '-------> Buscar dia vector dia excel
             For j = 1 To UBound(VecDiaExcel)
                 
                 If ColEnd = VecDiaExcel(j, 2) Then
                    
                    oColFec = VecDiaExcel(j - 1, 2)
                    oColPor = VecDiaExcel(j + 1, 2)
                    Exit For
                 
                 End If
             
             Next j
                   
             For ii = 6 To RowSheet + RowEnd1 + 2
                 
                 MoverDatosExcelFormula oExcel, oSheet, oColFec, oColPor, ColEnd, ColEnd, ii, RowSheet + RowEnd1 + 2, 0
             
             Next ii
          
             BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + 1, RowSheet + RowEnd1 + 2
             BloquearColumnaExcel oExcel, oSheet, oColFec, RowSheet + 1, RowSheet + RowEnd1 + 2
             FormatearColumnaNumericoExcel oExcel, oSheet, ColEnd
             FormatearColumnaPorcentajeExcel oExcel, oSheet, oColPor
             
            End If
            
             End If
          
             '-------> Dibujar lineas
             DibujarLineas oExcel, oSheet, "A", "A", 6, RowSheet + RowEnd1 + 2
             For j = 1 To UBound(VecDiaExcel)
              
                 oCol = VecDiaExcel(j, 2)
                 DibujarLineas oExcel, oSheet, oCol, oCol, 6, RowSheet + RowEnd1 + 2
          
             Next j
                      
             oSheet.Cells.Select
             oSheet.Cells.EntireColumn.AutoFit
             RowSheet = 5

             'Ocultar columna clave
             If Check3.Value = 1 Then
          
                ClaveExcel = "Jp123456"
             
                Set RS1 = vg_db.Execute("sgpadm_s_parametro 1, 'parhojaexc', ''")
                If Not RS1.EOF Then
                
                   ClaveExcel = RS1(0)
             
                End If
                RS1.Close
                Set RS1 = Nothing
             
                For j = 1 To UBound(VecDiaExcel) Step 4
                 
                    oCol = VecDiaExcel(j + 3, 2)
                    OcultarColumna oExcel, oSheet, oCol, oCol
                 
                    'Mover clave para poder actualizar
                 
                    MoverDatosExcel oExcel, oSheet, oCol, oCol, 1, 1, ClaveExcel
                    MoverDatosExcel oExcel, oSheet, oCol, oCol, 2, 2, CStr(((TotCol - 1) * 4) + 1)
             
                Next j
             
             End If
          
             'Bloquear protección hoja
             If Check2.Value = 1 And Check3.Value = 1 Then
             
                oSheet.Select
                oExcel.ActiveSheet.Protect password:=ClaveExcel, DrawingObjects:=True, _
                                    Contents:=True, Scenarios:=True, AllowFormattingCells:=True
                                 
              End If
          
             fg_descarga
         
          Else
          
noexiste:
             RS.Close
             Set RS = Nothing
             fg_descarga
             vaSpread1.Col = 9
             vaSpread1.text = "No existe Información"

       '      MsgBox "No existe información", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       
          End If
          
    End If

Next i

If Trim(auxceco) <> "" Then
             
   '-------> incluir recetas
   If Check5.Value = 1 Then
          
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
            
      Sql = ""
      Sql = Sql & " '" & LimpiaDato(Trim(auxceco)) & "' "
      Set RS = vg_db.Execute("sgpadm_Sel_ExportarExcelRecetas_V01 " & Sql & "")
       
      If Not RS.EOF Then
          
         '-------> Add data to cells of the first worksheet in the new workbook
         NombreServicio = "Recetas"
         Set oSheet = oBook.Worksheets.Add
         NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
          
         If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
             
            NombreServicioAux = Mid(NombreServicioAux, 1, 31)
          
         End If
         oSheet.Name = Trim(NombreServicioAux)
             
         '-------> Check version of Excel
         Call encabezado(RS, oSheet)
          
         oSheet.Cells(2, 1).CopyFromRecordset RS
         
         IndCol = 1
             
         MoverDatosExcel oExcel, oSheet, "A", "A", IndCol, IndCol, Trim("Código")
         MoverDatosExcel oExcel, oSheet, "B", "B", IndCol, IndCol, Trim("Nombre Plato")
         MoverDatosExcel oExcel, oSheet, "C", "C", IndCol, IndCol, Trim("Categoria Dietetica")
         MoverDatosExcel oExcel, oSheet, "D", "D", IndCol, IndCol, Trim("Tipo Plato")
         PonerColorInteriorN oExcel, oSheet, "A", "D", IndCol, IndCol, 4
                
         IndCol = 2
             
         PonerColorInteriorN oExcel, oSheet, "A", "D", IndCol, RS.RecordCount + 1, 6
                
         RS.Close
         Set RS = Nothing
          
         oSheet.Cells.Select
         oSheet.Cells.EntireColumn.AutoFit
             
         'Bloquear protección hoja
         If Check2.Value = 1 And Check3.Value = 1 Then
             
            With oSheet

                  .AutoFilterMode = False

                  .Range("A1:D1").AutoFilter

            End With
             
         End If
          
      End If
   
   End If
   
   oBook.Close True, NomArchivoExcel
          
   '------- Copiar archivos en la ruta seleccionada, luego borrar archivos de la carpeta
   If Trim(NomArchivoExcel) <> Trim(fpText1.text & SoloNomArchivoExcel) Then fso.CopyFile NomArchivoExcel, fpText1.text & SoloNomArchivoExcel, True
   If Dir(NomArchivoExcel) <> "" And Trim(NomArchivoExcel) <> Trim(fpText1.text & SoloNomArchivoExcel) Then Kill NomArchivoExcel
   
End If

'oExcel.Visible = True '------->Visualizar
Set oSheet = Nothing
Set oExcel = Nothing
Set oBook = Nothing
Bar1(0).Value = 0
Bar1(0).Visible = False

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Sub ExportarExcelMinutaBloqueSinCodigoReceta()

On Error GoTo Man_Error

Dim RS  As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim Sql As String
Dim Est As Boolean
Dim i   As Long
Dim j   As Long
Dim ii  As Long

'-------> Validar Org. Compras
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
            
Set RS = vg_db.Execute("sgpadm_Sel_OrgCompras_V02 '" & LimpiaDato(fpText.text) & "'")

If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   MsgBox "No existe Org. Compras...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   vaSpread1.MaxRows = 0
   Exit Sub

End If
RS.Close
Set RS = Nothing

'-------> validar seleción regimen
Est = False
For i = 1 To vaSpread1.MaxRows
    
    vaSpread1.Row = i
    vaSpread1.Col = 1
    
    If vaSpread1.text = "1" Then
       
       Est = True
    
    End If
Next i

If Not Est Then MsgBox "Regimen y servicios asociado debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub

'-------> Validar fecha desde - hasta
If CDate(FpFecDesde.text) > CDate(FpFecHasta.text) Then
   
   Call MsgBox("Fecha Desde No Puede Ser Mayor a Fecha Hasta", vbInformation, MsgTitulo)
   Let FpFecDesde.text = Format(Now, "dd/mm/yyyy")
   Call FpFecDesde.SetFocus
   Exit Sub

End If
    
If CDate(FpFecHasta.text) < CDate(FpFecDesde.text) Then
   
   Call MsgBox("Fecha Hasta No Puede Ser Mayor a Fecha Desde", vbInformation, MsgTitulo)
   Let FpFecHasta.text = Format(Now, "dd/mm/yyyy")
   Call FpFecHasta.SetFocus
   Exit Sub

End If

If DateDiff("y", FpFecDesde.text, FpFecHasta.text) > 365 Then
   
   Call MsgBox("Rango De Fecha No Puede Ser Mayor a 12 Meses", vbInformation, Me.Caption)
   Exit Sub

End If

Dim oExcel              As Object
Dim oBook               As Object
Dim oSheet              As Object
Dim RowSheet            As Long
Dim oCol                As String
Dim oColCla             As String
Dim oColA               As String
Dim oColFec             As String
Dim oColPor             As String
Dim NombreServicioAux   As String
Dim NombreServicio      As String
Dim CodigoServicio      As Long
Dim CodigoRegimen       As Long
Dim NombreRegimen       As String
Dim MaxColumna          As Long
Dim NumAsc              As Long
Dim FecMin              As Date
Dim FecMax              As Date
Dim IndCol              As Long
Dim IndColA             As Long
Dim IndVec              As Long
Dim AuxFec              As Long
Dim CodEstructura       As Long
Dim RowEnd              As Long
Dim RowEnd1             As Long
Dim ColEnd              As String
Dim TotDiaRaciones      As Double
Dim TotCol              As Long
Dim ClaveExcel          As String

Dim Ceco                As String
Dim auxceco             As String
Dim NomCeco             As String
Dim NomArchivoExcel     As String
Dim SoloNomArchivoExcel As String
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

''-------> Exportar excel
'Set oExcel = CreateObject("Excel.Application")
'Set oBook = oExcel.Workbooks.Add

'NumAsc = 66

fg_carga ""
Bar1(0).Visible = True
Bar1(0).Value = 0

For i = 1 To vaSpread1.MaxRows
    
    DoEvents
    
    vaSpread1.Row = i
    CodigoServicio = 0
    NombreServicio = ""
    vaSpread1.Col = 4: CodigoServicio = Val(vaSpread1.text)
    vaSpread1.Col = 5: NombreServicio = vaSpread1.text
    
    vaSpread1.Col = 1
    Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
    
    If vaSpread1.text = "1" And CodigoServicio > 0 And NombreServicio <> "Recetas" And vaSpread1.RowHidden = False Then
       
       vaSpread1.Col = 2: Ceco = vaSpread1.text
       vaSpread1.Col = 3: NomCeco = vaSpread1.text
       
       vaSpread1.SetActiveCell 2, vaSpread1.Row
               
       If Ceco <> auxceco Then
       
          
          If Trim(auxceco) <> "" Then
          
             oBook.Close True, NomArchivoExcel

             Set oSheet = Nothing
             Set oExcel = Nothing
             Set oBook = Nothing
          
             '------- Copiar archivos en la ruta seleccionada, luego borrar archivos de la carpeta
             If Trim(NomArchivoExcel) <> Trim(fpText1.text & SoloNomArchivoExcel) Then fso.CopyFile NomArchivoExcel, fpText1.text & SoloNomArchivoExcel, True
             If Dir(NomArchivoExcel) <> "" And Trim(NomArchivoExcel) <> Trim(fpText1.text & SoloNomArchivoExcel) Then Kill NomArchivoExcel
             
             Me.Refresh
             
          End If
          
          '-------> Exportar excel
          Set oExcel = CreateObject("Excel.Application")
          Set oBook = oExcel.Workbooks.Add
          NumAsc = 66
       
          auxceco = Ceco
          
          NomArchivoExcel = dir_trabajo & "ExcelMinutaSGP" & "\" & Ceco & "-" & NomCeco & " " & Format(FpFecDesde.text, "yyyymmdd") & " " & Format(FpFecHasta.text, "yyyymmdd") & " " & Replace(Date, "/", "") & " " & Replace(Time, ":", "") & ".xlsx"
          SoloNomArchivoExcel = Ceco & "-" & NomCeco & " " & Format(FpFecDesde.text, "yyyymmdd") & " " & Format(FpFecHasta.text, "yyyymmdd") & " " & Replace(Date, "/", "") & " " & Replace(Time, ":", "") & ".xlsx"
       
       End If
       
       vaSpread1.Col = 4: CodigoRegimen = vaSpread1.text
       vaSpread1.Col = 5: NombreRegimen = vaSpread1.text
       vaSpread1.Col = 6: CodigoServicio = Val(vaSpread1.text)
       vaSpread1.Col = 7: NombreServicio = vaSpread1.text
    
       '-------> Traer fecha minima - maxima
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Sql = ""
       Sql = Sql & " '" & LimpiaDato(Trim(Ceco)) & "'"
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
       Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
       Set RS = vg_db.Execute("sgpadm_Sel_MinBloqueMinMax " & Sql & "")

       If Not RS.EOF Then
          
          If IsNull(RS!FecMin) Then
             
'             RS.Close
'             Set RS = Nothing
'             MsgBox "No existe información, para este regimen y servicio", vbExclamation + vbOKOnly, MsgTitulo
'             Exit Sub
          
            GoTo noexiste
            
          End If
          
          MaxColumna = DateDiff("d", CDate(fg_Ctod1(RS!FecMin)), CDate(fg_Ctod1(RS!FecMax))) + 1
          
          '-------> Raciones
          If Check1.Value = 1 Then
             
             MaxColumna = MaxColumna * 2
          
          End If
          
          '-------> % Ponderación
          If Check2.Value = 1 Then
             
             MaxColumna = MaxColumna * 2
          
          End If
          Dim VecDiaExcel() As Variant
          ReDim VecDiaExcel(MaxColumna, 2)
          
          '-------> Setear vector
          For j = 1 To UBound(VecDiaExcel)
              
              VecDiaExcel(j, 1) = Val(0) 'fecha
              VecDiaExcel(j, 2) = "" 'descripción
          
          Next j
          
          FecMin = fg_Ctod1(RS!FecMin)
          FecMax = fg_Ctod1(RS!FecMax)
          
       End If
       RS.Close
       Set RS = Nothing
       
       '-------> Llamar procedimiento ultima fila
       RowEnd1 = 0
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Sql = ""
       Sql = Sql & LimpiaDato(Trim(Ceco))
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
       Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
       Set RS = vg_db.Execute("sgpadm_Sel_MaxLineaMinBloqueExp " & Sql & "")
          
       If Not RS.EOF Then
             
          RowEnd1 = RS!NumLin
          
       End If
       RS.Close
       Set RS = Nothing
          
       '-------> Llamar procedimiento
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Sql = ""
       Sql = Sql & LimpiaDato(Trim(Ceco))
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
       Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
       Set RS = vg_db.Execute("sgpadm_Sel_DetalleMinBloqueExpSinCostoReceta " & Sql & "")
       
       If Not RS.EOF Then
          
          vaSpread1.Col = 9
          vaSpread1.text = "Servicio termino correctamente"
          
          '-------> Add data to cells of the first worksheet in the new workbook
          Set oSheet = oBook.Worksheets.Add
          NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
          If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
             
             NombreServicioAux = Mid(CodigoRegimen & "-" & CodigoServicio & NombreServicioAux, 1, 31)
          
          End If
          oSheet.Name = Trim(NombreServicioAux)
          
          '-------> Mover Ceco - Regimen
          MoverDatosExcel oExcel, oSheet, "A", "A", 1, 1, "C.Costo " & Trim(NomCeco) & " - " & Trim(Ceco)
          MoverDatosExcel oExcel, oSheet, "A", "A", 2, 2, "Regimen " & Trim(NombreRegimen) & " - " & CodigoRegimen

          '-------> Mover titulo excel
          MoverDatosExcel oExcel, oSheet, "A", "A", 5, 5, "Estructura Servicio"
          DibujarLineas oExcel, oSheet, "A", "A", 5, 5
          IndCol = 1
          IndColA = 65
          IndVec = 0
          oCol = ""
          oColA = ""
          oCol = Chr(IndCol + 65)
          IndVec = 1
          TotCol = 1
          Do While FecMin <= FecMax
             
             MoverDatosExcel oExcel, oSheet, oCol, oCol, 5, 5, Mid(fg_Fecha_Dia(Format(FecMin, "yyyymmdd"), 1), 1, 4) & " " & FecMin
             DibujarLineas oExcel, oSheet, oCol, oCol, 5, 5
             
             VecDiaExcel(IndVec, 1) = Format(FecMin, "yyyymmdd") 'fecha
             VecDiaExcel(IndVec, 2) = oCol 'descripción
             
             '-------> Raciones
             If Check1.Value = 1 Then
                
                IndVec = IndVec + 1
                IndCol = IndCol + 1
                
                If Chr(IndCol + 65) = "[" Then
                   
                   oColA = Chr(IndColA)
                   IndColA = IndColA + 1
                   IndCol = 0
                
                End If
                
                oCol = oColA & Chr(IndCol + 65)
                VecDiaExcel(IndVec, 2) = oCol 'descripción
                MoverDatosExcel oExcel, oSheet, oCol, oCol, 5, 5, "Rac."
                DibujarLineas oExcel, oSheet, oCol, oCol, 5, 5
             
             End If
             
             '-------> % Ponderación
             If Check2.Value = 1 Then
                
                IndVec = IndVec + 1
                IndCol = IndCol + 1
                
                If Chr(IndCol + 65) = "[" Then
                   
                   oColA = Chr(IndColA)
                   IndColA = IndColA + 1
                   IndCol = 0
                
                End If
                
                oCol = oColA & Chr(IndCol + 65)
                VecDiaExcel(IndVec, 2) = oCol 'descripción
                MoverDatosExcel oExcel, oSheet, oCol, oCol, 5, 5, "% Pond."
                DibujarLineas oExcel, oSheet, oCol, oCol, 5, 5
             
             End If
             
             If Check3.Value = 1 Then
                
                IndVec = IndVec + 1
                IndCol = IndCol + 1
                
                If Chr(IndCol + 65) = "[" Then
                   
                   oColA = Chr(IndColA)
                   IndColA = IndColA + 1
                   IndCol = 0
                
                End If
                
                oCol = oColA & Chr(IndCol + 65)
                VecDiaExcel(IndVec, 2) = oCol 'descripción
                MoverDatosExcel oExcel, oSheet, oCol, oCol, 5, 5, "Clave"
                DibujarLineas oExcel, oSheet, oCol, oCol, 5, 5
             
             End If
             
             FecMin = FecMin + 1
             IndCol = IndCol + 1
             IndVec = IndVec + 1
             
             If Chr(IndCol + 65) = "[" Then
                
                oColA = Chr(IndColA)
                IndColA = IndColA + 1
                IndCol = 0
             
             End If
             oCol = oColA & Chr(IndCol + 65)

            TotCol = TotCol + 1
          
          Loop
          
          RowSheet = 5
          RowEnd = 0
          AuxFec = 0
          CodEstructura = 0
          Do While Not RS.EOF
             
             '-------> Corte x fecha
             If AuxFec <> RS!min_fecmin Then
                
                '-------> Mover comensales totales
                If AuxFec > 0 And Check1.Value = 1 Then
                   
                   MoverDatosExcel oExcel, oSheet, "A", "A", RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, "Comensales"
                   MoverDatosExcel oExcel, oSheet, ColEnd, ColEnd, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, CStr(TotDiaRaciones)
                   
                   If Check3.Value = 1 Then
                      
                      MoverDatosExcel oExcel, oSheet, oColCla, oColCla, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, LimpiaDato(Trim(fpText.text)) & ";" & CodigoRegimen & ";" & CodigoServicio & ";" & CStr(AuxFec)
                   
                   End If
                   
                   
                   If Check2.Value = 1 And Check3.Value = 1 Then
                      
                      '-------> Buscar dia vector dia excel
                      For j = 1 To UBound(VecDiaExcel)
                        
                        If ColEnd = VecDiaExcel(j, 2) Then
                            
                            oColFec = VecDiaExcel(j - 1, 2)
                            oColPor = VecDiaExcel(j + 1, 2)
                            Exit For
                        
                        End If
                      
                      Next j
                   
                      For ii = 6 To RowSheet + RowEnd + 2
                          
                          MoverDatosExcelFormula oExcel, oSheet, oColFec, oColPor, ColEnd, ColEnd, ii, RowSheet + RowEnd1 + 2, 0
                      
                      Next ii

                      BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + 1, RowSheet + RowEnd1 + 2
                      BloquearColumnaExcel oExcel, oSheet, oColFec, RowSheet + 1, RowSheet + RowEnd1 + 2
                      FormatearColumnaNumericoExcel oExcel, oSheet, ColEnd
                    
                    End If
                    
                    RowEnd = 0
                
                End If
                RowSheet = 5
                AuxFec = RS!min_fecmin
                TotDiaRaciones = RS!min_racteo
             
             End If
             
             '-------> Buscar dia vector dia excel
             For j = 1 To UBound(VecDiaExcel)
                 
                 If RS!min_fecmin = Val(VecDiaExcel(j, 1)) Then
                    
                    oCol = VecDiaExcel(j, 2)
                    IndCol = j
                    
                    If Check3.Value = 1 Then
                       
                       oColCla = VecDiaExcel(j + 3, 2)
                    
                    End If
                    
                    If Check1.Value = 1 Then
                       
                       ColEnd = VecDiaExcel(j + 1, 2)
                    
                    End If
                    Exit For
                 
                 End If
             
             Next j
             
             '-------> Corte x estructura servicio
             '-------> Mover Estructura servicio excel
             If CodEstructura <> RS!ess_codigo Then
                
                MoverDatosExcel oExcel, oSheet, "A", "A", RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, Trim(RS!ess_nombre)
                CodEstructura = RS!ess_codigo
             
             End If

             If Check3.Value = 1 Then
                
                MoverDatosExcel oExcel, oSheet, oColCla, oColCla, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, LimpiaDato(Trim(fpText.text)) & ";" & CodigoRegimen & ";" & CodigoServicio & ";" & RS!mid_codrec & ";" & RS!min_fecmin & ";" & RS!mid_numlin
             
             End If
             
             '-------> Mover recetas excel
             MoverDatosExcel oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, Trim(RS!rec_nombre)
             
             '-------> Mover raciones excel
             If Check1.Value = 1 Then
                
                oCol = VecDiaExcel(IndCol + 1, 2)
                IndCol = IndCol + 1
                MoverDatosExcel oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, RS!mid_numrac
             
             End If
             
             '-------> Mover % ponderación excel
             If Check2.Value = 1 Then
                
                oCol = VecDiaExcel(IndCol + 1, 2)
                MoverDatosExcel oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, RS!mid_porrac & " %"
             
             End If
             
             If RS!mid_numlin > RowEnd Then
                
                RowEnd = RS!mid_numlin
             
             End If
             
             RS.MoveNext
          
          Loop
          RS.Close
          Set RS = Nothing
          
          '-------> Mover comensales totales
          If AuxFec > 0 And Check1.Value = 1 Then
             
             MoverDatosExcel oExcel, oSheet, "A", "A", RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, "Comensales"
             MoverDatosExcel oExcel, oSheet, ColEnd, ColEnd, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, CStr(TotDiaRaciones)
          
             If Check3.Value = 1 Then
                
                MoverDatosExcel oExcel, oSheet, oColCla, oColCla, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, LimpiaDato(Trim(fpText.text)) & ";" & CodigoRegimen & ";" & CodigoServicio & ";" & CStr(AuxFec)
             
             End If
            
            If Check2.Value = 1 And Check3.Value = 1 Then
             
             '-------> Buscar dia vector dia excel
             For j = 1 To UBound(VecDiaExcel)
                 
                 If ColEnd = VecDiaExcel(j, 2) Then
                    
                    oColFec = VecDiaExcel(j - 1, 2)
                    oColPor = VecDiaExcel(j + 1, 2)
                    Exit For
                 
                 End If
             
             Next j
                   
             For ii = 6 To RowSheet + RowEnd + 2
                 
                 MoverDatosExcelFormula oExcel, oSheet, oColFec, oColPor, ColEnd, ColEnd, ii, RowSheet + RowEnd1 + 2, 0
             
             Next ii
          
             BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + 1, RowSheet + RowEnd1 + 2
             BloquearColumnaExcel oExcel, oSheet, oColFec, RowSheet + 1, RowSheet + RowEnd1 + 2
             FormatearColumnaNumericoExcel oExcel, oSheet, ColEnd
             
            End If
            
          End If
          
          '-------> Dibujar lineas
          DibujarLineas oExcel, oSheet, "A", "A", 6, RowSheet + RowEnd1 + 2
          For j = 1 To UBound(VecDiaExcel)
              
              oCol = VecDiaExcel(j, 2)
              DibujarLineas oExcel, oSheet, oCol, oCol, 6, RowSheet + RowEnd1 + 2
          
          Next j
                      
          oSheet.Cells.Select
          oSheet.Cells.EntireColumn.AutoFit
          RowSheet = 5

          'Ocultar columna clave
          If Check3.Value = 1 Then
          
             ClaveExcel = "Jp123456"
             
             If RS1.State = 1 Then RS1.Close
             RS1.CursorLocation = adUseClient
             vg_db.CursorLocation = adUseClient
            
             Set RS1 = vg_db.Execute("sgpadm_s_parametro 1, 'parhojaexc', ''")
             
             If Not RS1.EOF Then
                
                ClaveExcel = RS1(0)
             
             End If
             RS1.Close
             Set RS1 = Nothing
             
             For j = 1 To UBound(VecDiaExcel) Step 4
                 
                 oCol = VecDiaExcel(j + 3, 2)
                 OcultarColumna oExcel, oSheet, oCol, oCol
                 
                 'Mover clave para poder actualizar
                 
                 MoverDatosExcel oExcel, oSheet, oCol, oCol, 1, 1, ClaveExcel
                 MoverDatosExcel oExcel, oSheet, oCol, oCol, 2, 2, CStr(((TotCol - 1) * 4) + 1)
             
             Next j
             
          End If
          
          'Bloquear protección hoja
          If Check2.Value = 1 And Check3.Value = 1 Then
             
             oSheet.Select
             oExcel.ActiveSheet.Protect password:=ClaveExcel, DrawingObjects:=True, _
                                 Contents:=True, Scenarios:=True, AllowFormattingCells:=True
                                 
          End If
          
          fg_descarga
       
       Else
          
noexiste:
          RS.Close
          Set RS = Nothing
          fg_descarga
          vaSpread1.Col = 9
          vaSpread1.text = "No existe información"

'          MsgBox "No existe información", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       
       End If
    
    End If

Next i

If Trim(auxceco) <> "" Then
          
   oBook.Close True, NomArchivoExcel

   Set oSheet = Nothing
   Set oExcel = Nothing
   Set oBook = Nothing
          
   '------- Copiar archivos en la ruta seleccionada, luego borrar archivos de la carpeta
   If Trim(NomArchivoExcel) <> Trim(fpText1.text & SoloNomArchivoExcel) Then fso.CopyFile NomArchivoExcel, fpText1.text & SoloNomArchivoExcel, True
   If Dir(NomArchivoExcel) <> "" And Trim(NomArchivoExcel) <> Trim(fpText1.text & SoloNomArchivoExcel) Then Kill NomArchivoExcel
   
   Me.Refresh
             
End If

'oExcel.Visible = True '------->Visualizar
'Set oSheet = Nothing
'Set oExcel = Nothing
'Set oBook = Nothing
Bar1(0).Value = 0
Bar1(0).Visible = False

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Sub CargaRecetaCosto(Ceco As String, Regimen As Long)

On Error GoTo Man_Error

Dim IndRec As Long
Dim Fecha  As Date

Dim XmlDietetica As String
Dim XmlPlato     As String
Dim IndFiltro    As Long
Dim Die          As Long
Dim Pla          As Long

'---------> Armar Xml Categoria Dietetica & Tipo Plato
Let XmlDietetica = ""
Let XmlDietetica = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let XmlDietetica = XmlDietetica & "<Dietetica>"

For IndFiltro = 1 To B_DieTipExcel.TvwDir(0).Nodes.count

    If B_DieTipExcel.TvwDir(0).Nodes.item(IndFiltro).Checked = True And Trim(B_DieTipExcel.TvwDir(0).Nodes.item(IndFiltro).text) <> "*" Then
          
       XmlDietetica = XmlDietetica & " <DetDietetica"
       
       XmlDietetica = XmlDietetica & " Die = " & Chr(34) & Val(Mid(B_DieTipExcel.TvwDir(0).Nodes.item(IndFiltro).text, 1, InStr(B_DieTipExcel.TvwDir(0).Nodes.item(IndFiltro).text, " - ") - 1)) & Chr(34)
       XmlDietetica = XmlDietetica & "/>"

    End If
       
Next IndFiltro

XmlDietetica = XmlDietetica & "</Dietetica>"

Let XmlPlato = ""
Let XmlPlato = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let XmlPlato = XmlPlato & "<Plato>"

For IndFiltro = 1 To B_DieTipExcel.TvwDir(1).Nodes.count

    If B_DieTipExcel.TvwDir(1).Nodes.item(IndFiltro).Checked = True And Trim(B_DieTipExcel.TvwDir(1).Nodes.item(IndFiltro).text) <> "*" Then
          
       XmlPlato = XmlPlato & " <DetPlato"
       
       XmlPlato = XmlPlato & " Pla = " & Chr(34) & Val(Mid(B_DieTipExcel.TvwDir(1).Nodes.item(IndFiltro).text, 1, InStr(B_DieTipExcel.TvwDir(1).Nodes.item(IndFiltro).text, "-") - 1)) & Chr(34)
       XmlPlato = XmlPlato & "/>"

    End If
       
Next IndFiltro

XmlPlato = XmlPlato & "</Plato>"

Fecha = "01/" & Format(FpFecDesde.text, "mm/yyyy")

If Rs_Receta.State = 1 Then Rs_Receta.Close
Rs_Receta.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set Rs_Receta = vg_db.Execute("sgpadm_Sel_XMLResumenCostoReceta_Excel_V02 '" & Ceco & "', " & Regimen & ", 10001, " & Format(Fecha, "yyyymmdd") & ", 1, '1', '" & XmlDietetica & "', '" & XmlPlato & "', '" & Format(FpFecDesde.text, "ddmm") & "', '" & Format(FpFecHasta.text, "ddmm") & "'")

IndRec = 1
'vaSpread2.MaxRows = 0
'vaSpread2.maxcols = 10
ReDim VecRecetas(Rs_Receta.RecordCount, 2)
'vaSpread2.MaxRows =

'zetear variables
For IndRec = 1 To UBound(VecRecetas)

  VecRecetas(IndRec, 1) = 0
  VecRecetas(IndRec, 2) = 0#
    
Next

IndRec = 1
If Not Rs_Receta.EOF Then
        
   DoEvents
   Screen.MousePointer = 11
        
   Do While Not Rs_Receta.EOF
            
      VecRecetas(IndRec, 1) = Rs_Receta!rec_codigo
      VecRecetas(IndRec, 2) = Format(Rs_Receta!promedioreceta, fg_Pict(6, 2))
      
      IndRec = IndRec + 1
            
      Rs_Receta.MoveNext
          
   Loop

End If

fg_descarga
'Rs_Receta.Close
'Set Rs_Receta = Nothing

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Function BuscarCostoReceta(CodRec As String) As Double

On Error GoTo Man_Error

Dim ret    As Long
Dim IndRec As Long
IndRec = 1

BuscarCostoReceta = 0
For IndRec = 1 To UBound(VecRecetas)

    If VecRecetas(IndRec, 1) = CodRec Then
    
        BuscarCostoReceta = VecRecetas(IndRec, 2)
        Exit For
    
    End If

Next

'ret = vaSpread2.SearchCol(1, 1, vaSpread2.MaxRows, CodRec, 4)
'
'If ret > -1 Then
'
'   vaSpread2.Row = ret
'   vaSpread2.Col = 3
'   BuscarCostoReceta = IIf(Trim(vaSpread2.text) = "", 0, vaSpread2.text)
'
'End If

Exit Function
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Function

Sub encabezado(ByRef rst As ADODB.Recordset, ByRef xlWs As Object)

On Error GoTo Man_Error

Dim fldCount As Long
Dim icol     As Long

'-------> Copy field names to the first row of the worksheet
fldCount = rst.Fields.count
For icol = 1 To fldCount
    
    xlWs.Cells(1, icol).Value = rst.Fields(icol - 1).Name

Next

Exit Sub
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo

End Sub

Sub ExportarExcelMinutaBloqueCostoReceta_Backup()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim RS1             As New ADODB.Recordset
Dim Sql             As String
Dim Est             As Boolean
Dim i               As Long
Dim j               As Long
Dim ii              As Long
Dim NomArchivoExcel As String

'-------> Validar Org. Compras
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
            
Set RS = vg_db.Execute("sgpadm_Sel_OrgCompras_V02 '" & LimpiaDato(fpText.text) & "'")
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   MsgBox "No existe Org. Compras...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   vaSpread1.MaxRows = 0
   Exit Sub

End If
RS.Close
Set RS = Nothing

'-------> validar seleción ceco - regimen - servicio
Est = False

For i = 1 To vaSpread1.MaxRows
    
    vaSpread1.Row = i
    vaSpread1.Col = 1
    
    If vaSpread1.text = "1" Then
       
       Est = True
    
    End If

Next i
If Not Est Then MsgBox "Regimen y servicios asociado debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub

'-------> Validar fecha desde - hasta
If CDate(FpFecDesde.text) > CDate(FpFecHasta.text) Then
   
   Call MsgBox("Fecha Desde No Puede Ser Mayor a Fecha Hasta", vbInformation, MsgTitulo)
   Let FpFecDesde.text = Format(Now, "dd/mm/yyyy")
   Call FpFecDesde.SetFocus
   Exit Sub

End If
    
If CDate(FpFecHasta.text) < CDate(FpFecDesde.text) Then
   
   Call MsgBox("Fecha Hasta No Puede Ser Mayor a Fecha Desde", vbInformation, MsgTitulo)
   Let FpFecHasta.text = Format(Now, "dd/mm/yyyy")
   Call FpFecHasta.SetFocus
   Exit Sub

End If

'If DateDiff("ww", FpFecDesde.text, FpFecHasta.text) > 52 Then
If DateDiff("y", FpFecDesde.text, FpFecHasta.text) > 365 Then
   
   Call MsgBox("Rango De Fecha No Puede Ser Mayor a 12 Meses", vbInformation, Me.Caption)
   Exit Sub

End If

Dim oExcel             As Object
Dim oBook              As Object
Dim oSheet             As Object
Dim ObjExcel           As excel.Application
Dim ObjW               As excel.Workbook
Dim RowSheet           As Long
Dim oCol               As String
Dim oColCla            As String
Dim oColA              As String
Dim oColFec            As String
Dim oColPor            As String
Dim oColCos            As String
Dim oColRec            As String
Dim Ceco               As String
Dim NomCeco            As String
Dim auxceco            As String
Dim NombreServicioAux  As String
Dim NombreServicio     As String
Dim CodigoServicio     As Long
Dim CodigoRegimen      As Long
Dim Aux_CodigoRegimen  As Long
Dim NombreRegimen      As String
Dim MaxColumna         As Long
Dim DiaColumna         As Long
Dim NumAsc             As Long
Dim FecMin             As Date
Dim FecMax             As Date
Dim IndCol             As Long
Dim IndColA            As Long
Dim IndVec             As Long
Dim IndHoja            As Long
Dim AuxFec             As Long
Dim CodEstructura      As Long
Dim RowEnd             As Long
Dim RowEnd1            As Long
Dim ColEnd             As String
Dim TotDiaRaciones     As Double
Dim TotCol             As Long
Dim EstCeco            As Boolean
Dim oColCpo            As String
Dim CalRacCos          As String
Dim TotCalRacCos       As String
Dim CalComen           As String

Dim ClaveExcel         As String
Dim CostoComercial     As Double
Dim CostoPlanificacion As Double
Dim CostoReceta        As Double
Dim Comensales         As Double
Dim CostoBandeja       As Double
Dim VecDiaExcel()      As Variant

fg_carga ""
Bar1(0).Visible = True
Bar1(0).Value = 0
EstCeco = True

For i = 1 To vaSpread1.MaxRows
    
    DoEvents
    
    vaSpread1.Row = i
    vaSpread1.Col = 1
    Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
    
    If vaSpread1.text = "1" Then
       
       vaSpread1.SetActiveCell 2, vaSpread1.Row
       vaSpread1.Col = 2
       Ceco = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
       vaSpread1.Col = 3
       NomCeco = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
       
       CalRacCos = ""
       TotCalRacCos = ""
       CalComen = ""
       CostoReceta = 0
       Comensales = 0
       
       If Ceco <> auxceco Then
       
          If Trim(auxceco) <> "" Then
             
             If Check2.Value = 1 And Check3.Value = 1 Then
          
'                       ObjW.Sheets(IndHoja).Select
'                       ObjW.Sheets(IndHoja).Protect password:=ClaveExcel, DrawingObjects:=True, _
'                                              Contents:=True, Scenarios:=True, AllowFormattingCells:=True
                                              
                oSheet.Select
                oExcel.ActiveSheet.Protect password:=ClaveExcel, DrawingObjects:=True, _
                                    Contents:=True, Scenarios:=True, AllowFormattingCells:=True
                                              
                                 
             End If
             
             oBook.Close True, NomArchivoExcel
'             oExcel.Visible = True '------->Visualizar
             
             Set oSheet = Nothing
             Set oExcel = Nothing
             Set ObjW = Nothing
             Set ObjExcel = Nothing
             Set oBook = Nothing
             
             'Set oSheet = Nothing
             'Set oExcel = Nothing
             'Set oBook = Nothing
          
'             Set ObjExcel = New excel.Application
'             Set ObjW = ObjExcel.Workbooks.Open(NomArchivoExcel)
          
'             For IndHoja = 1 To ObjW.Sheets.count
 
'                 If Trim(ObjW.Sheets(IndHoja).Name) <> "Recetas" And Trim(ObjW.Sheets(IndHoja).Name) <> "Hoja1" Then
        
'                    ObjW.Sheets(IndHoja).Cells.Replace What:="Hoja1", Replacement:="recetas", LookAt:=xlPart, _
'                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
                                      
'                 End If
    
'             Next
    
'             ObjW.Save
'             ObjW.Application.DisplayAlerts = False
'             ObjW.Close True
'             ObjExcel.Quit
             
             Set oSheet = Nothing
             Set oBook = Nothing
             Set oExcel = Nothing
             Set ObjExcel = Nothing
             Set ObjW = Nothing
             
'             Me.Refresh
             
          End If
          
          '-------> Exportar excel
          Set oExcel = CreateObject("Excel.Application")
          Set oBook = oExcel.Workbooks.Add
          NumAsc = 66
       
          auxceco = Ceco
          EstCeco = True
          
          NomArchivoExcel = dir_trabajo & "ExcelMinutaSGP" & "\" & Ceco & "-" & NomCeco & " " & Format(FpFecDesde.text, "yyyymmdd") & " " & Format(FpFecHasta.text, "yyyymmdd") & " " & Replace(Date, "/", "") & " " & Replace(Time, ":", "") & ".xlsx"
       
       End If
       
       vaSpread1.Col = 4
       CodigoRegimen = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
       
       vaSpread1.Col = 5
       NombreRegimen = IIf(Trim(vaSpread1.text) = "", "", vaSpread1.text)
       
       vaSpread1.Col = 6
       CodigoServicio = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
       
       vaSpread1.Col = 7
       NombreServicio = IIf(Trim(vaSpread1.text) = "", "", vaSpread1.text)
    
       '-------> Traer costo Comercial
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Sql = ""
       Sql = Sql & " '" & LimpiaDato(Trim(Ceco)) & "'"
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymm")
       Sql = Sql & ",'COMERCIAL'"
       Set RS = vg_db.Execute("sgpadm_Sel_CostoComercialCeco " & Sql & "")

       CostoComercial = 0
       If Not RS.EOF Then
       
          CostoComercial = RS!pcp_valor
       
       End If
       RS.Close
       Set RS = Nothing
       
       '-------> Traer parametro ceco x regimen
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Sql = ""
       Sql = Sql & " '" & LimpiaDato(Trim(Ceco)) & "'"
       Set RS = vg_db.Execute("sgpadm_Sel_ParametroCecoregimen " & Sql & "")

       Aux_CodigoRegimen = CodigoRegimen
       
       If Not RS.EOF Then
       
          Aux_CodigoRegimen = RS!par_codreg
       
       End If
       RS.Close
       Set RS = Nothing
       
       '-------> incluir recetas costo recetas
       If Check5.Value = 1 And EstCeco Then
       
          EstCeco = False
          '-------> Rutinas carga Recetas con Costo
          CargaRecetaCosto Ceco, Aux_CodigoRegimen

          '-------> incluir recetas
          If Check5.Value = 1 Then
             
             '-------> Add data to cells of the first worksheet in the new workbook
             NombreServicio = "Recetas"
             Set oSheet = Nothing
             Set oSheet = oBook.Worksheets.Add
             NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)

             If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then

                NombreServicioAux = Mid(NombreServicioAux, 1, 31)

             End If
             oSheet.Name = "Recetas" 'Trim(NombreServicioAux)
                
             '-------> Check version of Excel
             Call encabezado(Rs_Receta, oSheet)
          
             oSheet.Cells(2, 1).CopyFromRecordset Rs_Receta

             MoverDatosExcel oExcel, oSheet, "A", "A", 1, 1, Trim("Código")
             MoverDatosExcel oExcel, oSheet, "B", "B", 1, 1, Trim("Nombre Plato")
             MoverDatosExcel oExcel, oSheet, "C", "C", 1, 1, Trim("Costo Plato")
             MoverDatosExcel oExcel, oSheet, "D", "D", 1, 1, Trim("Categoria Dietetica")
             MoverDatosExcel oExcel, oSheet, "E", "E", 1, 1, Trim("Tipo Plato")
             PonerColorInteriorN oExcel, oSheet, "A", "E", 1, 1, 4

             IndCol = 2
             PonerColorInteriorN oExcel, oSheet, "A", "E", IndCol, Rs_Receta.RecordCount + 1, 6
                    
             oSheet.Cells.Select
             oSheet.Cells.EntireColumn.AutoFit
             
             Rs_Receta.Close
             Set Rs_Receta = Nothing
             
          End If
       
       End If
    
       vaSpread1.Col = 7
       NombreServicio = IIf(Trim(vaSpread1.text) = "", "", vaSpread1.text)
       
       '-------> Traer fecha minima - maxima
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Sql = ""
       Sql = Sql & " '" & LimpiaDato(Trim(Ceco)) & "'"
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
       Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
       Set RS = vg_db.Execute("sgpadm_Sel_MinBloqueMinMax " & Sql & "")

       If Not RS.EOF Then
          
          If IsNull(RS!FecMin) Then
             
             RS.Close
             Set RS = Nothing
             MsgBox "No existe información, para este regimen y servicio", vbExclamation + vbOKOnly, MsgTitulo
             Exit Sub
          
          End If
          
          DiaColumna = DateDiff("d", CDate(fg_Ctod1(RS!FecMin)), CDate(fg_Ctod1(RS!FecMax))) + 1
          
          MaxColumna = 2
          
          '-------> Raciones
          If Check1.Value = 1 Then
             
             MaxColumna = MaxColumna + 1
          
          End If
          
          '-------> % Ponderación
          If Check2.Value = 1 Then
             
             MaxColumna = MaxColumna + 1
          
          End If
          
          MaxColumna = MaxColumna + 2

          MaxColumna = MaxColumna * DiaColumna
         
          ReDim VecDiaExcel(MaxColumna, 2)
          
          '-------> Setear vector
          For j = 1 To UBound(VecDiaExcel)
              
              VecDiaExcel(j, 1) = Val(0) 'fecha
              VecDiaExcel(j, 2) = "" 'descripción
          
          Next j
          
          FecMin = fg_Ctod1(RS!FecMin)
          FecMax = fg_Ctod1(RS!FecMax)
          
       End If
       RS.Close
       Set RS = Nothing
          
       '-------> Llamar procedimiento
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       RowEnd1 = 0
       Sql = ""
       Sql = Sql & LimpiaDato(Trim(Ceco))
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
       Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
       Set RS = vg_db.Execute("sgpadm_Sel_MaxLineaMinBloqueExp " & Sql & "")
          
       If Not RS.EOF Then
             
          RowEnd1 = RS!NumLin
          
       End If
       RS.Close
       Set RS = Nothing
       
       '-------> Llamar procedimiento
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Sql = ""
       Sql = Sql & LimpiaDato(Trim(Ceco))
       Sql = Sql & ", " & CodigoRegimen
       Sql = Sql & ", " & CodigoServicio
       Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
       Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
       Set RS = vg_db.Execute("sgpadm_Sel_DetalleMinBloqueExpSinCostoReceta " & Sql & "")
       
       If Not RS.EOF Then
          
          vaSpread1.Col = 9
          vaSpread1.text = "Servicio termino correctamente"
          
          '-------> Add data to cells of the first worksheet in the new workbook
          Set oSheet = Nothing
          Set oSheet = oBook.Worksheets.Add
          
          NombreServicioAux = ""
          NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
          
          If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
             
             NombreServicioAux = Mid(CodigoRegimen & "-" & CodigoServicio & NombreServicioAux, 1, 31)
          
          End If
          oSheet.Name = Trim(NombreServicioAux)
          
          '-------> Mover Ceco - Regimen
          MoverDatosExcel oExcel, oSheet, "A", "A", 1, 1, "C.Costo " & Trim(NomCeco) & " - " & Trim(Ceco)
          MoverDatosExcel oExcel, oSheet, "A", "A", 2, 2, "Regimen " & Trim(NombreRegimen) & " - " & CodigoRegimen

'          '-------> Costos Ppto/C.Comercial
'          MoverDatosExcel oExcel, oSheet, "A", "A", 5, 5, "Costo Ppto/C.Comercial"
'          DibujarLineas oExcel, oSheet, "A", "A", 5, 5
          
          '-------> Costos Costo Planificación
          MoverDatosExcel oExcel, oSheet, "A", "A", 6, 6, "Costo Planificación"
          DibujarLineas oExcel, oSheet, "A", "A", 6, 6
          
          '-------> Costos Costo Sitio
          MoverDatosExcel oExcel, oSheet, "A", "A", 7, 7, "Costo Sitio"
          DibujarLineas oExcel, oSheet, "A", "A", 7, 7

          '-------> Mover titulo excel
          MoverDatosExcel oExcel, oSheet, "A", "A", 8, 8, "Estructura Servicio"
          DibujarLineas oExcel, oSheet, "A", "A", 8, 8
          IndCol = 1
          IndColA = 65
          IndVec = 0
          oCol = ""
          oColA = ""
          oCol = Chr(IndCol + 65)
          IndVec = 1
          TotCol = 1
          
          Do While FecMin <= FecMax
                
             MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, Mid(fg_Fecha_Dia(Format(FecMin, "yyyymmdd"), 1), 1, 4) & " " & FecMin
             DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
             
             VecDiaExcel(IndVec, 1) = Format(FecMin, "yyyymmdd") 'fecha
             VecDiaExcel(IndVec, 2) = oCol 'descripción
             
             '-------> Raciones
             If Check1.Value = 1 Then
                   
                IndVec = IndVec + 1
                IndCol = IndCol + 1
                   
                If Chr(IndCol + 65) = "[" Then
                      
                   oColA = Chr(IndColA)
                   IndColA = IndColA + 1
                   IndCol = 0
                   
                End If
                   
                oCol = oColA & Chr(IndCol + 65)
                VecDiaExcel(IndVec, 2) = oCol 'descripción
                MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, "Rac."
                DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
                
             End If
             
             '-------> % Ponderación
             If Check2.Value = 1 Then
                   
                IndVec = IndVec + 1
                IndCol = IndCol + 1
                   
                If Chr(IndCol + 65) = "[" Then
                      
                   oColA = Chr(IndColA)
                   IndColA = IndColA + 1
                   IndCol = 0
                   
                End If
                   
                oCol = oColA & Chr(IndCol + 65)
                VecDiaExcel(IndVec, 2) = oCol 'descripción
                MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, "% Pond."
                DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
                
             End If
             
             '-------> Costo receta
             IndVec = IndVec + 1
             IndCol = IndCol + 1
                   
             If Chr(IndCol + 65) = "[" Then
                      
                oColA = Chr(IndColA)
                IndColA = IndColA + 1
                IndCol = 0
                   
             End If
                   
             oCol = oColA & Chr(IndCol + 65)
             VecDiaExcel(IndVec, 2) = oCol 'descripción
             MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, "Cto. Plato"
             DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
             
             '-------> Costo receta ponderado
             IndVec = IndVec + 1
             IndCol = IndCol + 1
                   
             If Chr(IndCol + 65) = "[" Then
                      
                oColA = Chr(IndColA)
                IndColA = IndColA + 1
                IndCol = 0
                   
             End If
                   
             oCol = oColA & Chr(IndCol + 65)
             VecDiaExcel(IndVec, 2) = oCol 'descripción
             MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, "Cto. Plato Pon."
             DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
             
             
             If Check3.Value = 1 Then
                
                IndVec = IndVec + 1
                IndCol = IndCol + 1
                   
                If Chr(IndCol + 65) = "[" Then
                      
                   oColA = Chr(IndColA)
                   IndColA = IndColA + 1
                   IndCol = 0
                   
                End If
                   
                oCol = oColA & Chr(IndCol + 65)
                VecDiaExcel(IndVec, 2) = oCol 'descripción
                MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, "Clave"
                DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
             
             End If
             
             FecMin = FecMin + 1
             IndCol = IndCol + 1
             IndVec = IndVec + 1
             
             If Chr(IndCol + 65) = "[" Then
                   
                oColA = Chr(IndColA)
                IndColA = IndColA + 1
                IndCol = 0
                
             End If
                
             oCol = oColA & Chr(IndCol + 65)

             TotCol = TotCol + 1
             
          Loop
          
          RowSheet = 8 '7 '5
          RowEnd = 0
          AuxFec = 0
          CodEstructura = 0
             
          Do While Not RS.EOF
             
             '-------> Corte x fecha
             If AuxFec <> RS!min_fecmin Then
                
                '-------> Mover comensales totales
                If AuxFec > 0 And Check1.Value = 1 Then
                      
                   MoverDatosExcel oExcel, oSheet, "A", "A", RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, "Comensales"
                   MoverDatosExcel oExcel, oSheet, ColEnd, ColEnd, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, CStr(TotDiaRaciones)
                   
                   CalComen = CalComen & ColEnd & RowSheet + RowEnd1 + 2 & "+"
                   
                   If Check3.Value = 1 Then
                         
                      MoverDatosExcel oExcel, oSheet, oColCla, oColCla, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, LimpiaDato(Trim(Ceco)) & ";" & CodigoRegimen & ";" & CodigoServicio & ";" & CStr(AuxFec)
                      
                   End If
                   
                   If Check2.Value = 1 And Check3.Value = 1 Then
                      
                      '-------> Buscar dia vector dia excel
                      For j = 1 To UBound(VecDiaExcel)
                        
                         If ColEnd = VecDiaExcel(j, 2) Then
                            
                            oColFec = VecDiaExcel(j - 1, 2)
                            oColPor = VecDiaExcel(j + 1, 2)
                            oColCos = VecDiaExcel(j + 2, 2)
                            oColCpo = VecDiaExcel(j + 3, 2)
                            Exit For
                           
                         End If
                      
                      Next j
                   
                      For ii = 9 To RowSheet + RowEnd1 + 2
                          
                          MoverDatosExcelFormula oExcel, oSheet, oColFec, oColPor, ColEnd, ColEnd, ii, RowSheet + RowEnd1 + 2, 0
                          MoverDatosExcelFormulaIII oExcel, oSheet, oColFec, oColCpo, ColEnd, oColCos, ii, RowSheet + RowEnd1 + 2, 0
                          
                          If ii < RowSheet + RowEnd1 + 2 Then
                             
                             If Trim(oSheet.Range(ColEnd & ii).Value) <> "" Then
                                
                                CalRacCos = CalRacCos & ColEnd & ii & "*" & oColCos & ii & "+"
                                
                             End If
                             
                          End If
                      
                      Next ii
                      
                      MoverDatosExcelFormulaSum oExcel, oSheet, oColCpo, oColCpo, oColCpo, oColCpo, RowSheet + 1, RowSheet + RowEnd1 + 2, ""
'                      MoverDatosExcel oExcel, oSheet, oColCpo, oColCpo, 5, 5, CStr(CostoComercial)
                      MoverDatosExcel oExcel, oSheet, oColCpo, oColCpo, 6, 6, CStr(Format(CostoPlanificacion, fg_Pict(6, 2)))
                      MoverDatosExcelCostoBandeja oExcel, oSheet, oColCos, RowSheet + RowEnd1 + 4, "(" & Mid(CalRacCos, 1, Len(CalRacCos) - 1) & ")"
                      TotCalRacCos = TotCalRacCos & oColCos & RowSheet + RowEnd1 + 4 & "+"
                      CalRacCos = ""
                      CostoPlanificacion = 0
                      
                      BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + 1, RowSheet + RowEnd1 + 2
                      BloquearColumnaExcel oExcel, oSheet, oColFec, RowSheet + 1, RowSheet + RowEnd1 + 2
                      FormatearColumnaNumericoExcel oExcel, oSheet, ColEnd
                      FormatearColumnaPorcentajeExcel oExcel, oSheet, oColPor
                       
                    End If
                    
                    RowEnd = 0
                
                End If
                
                RowSheet = 8
                AuxFec = RS!min_fecmin
                TotDiaRaciones = RS!min_racteo
                Comensales = Comensales + RS!min_racteo
             
             End If
             
             '-------> Buscar dia vector dia excel
             For j = 1 To UBound(VecDiaExcel)
                 
                 If RS!min_fecmin = Val(VecDiaExcel(j, 1)) Then
                       
                    oCol = VecDiaExcel(j, 2)
                    IndCol = j
                    
                    If Check3.Value = 1 Then
                          
                       oColCla = VecDiaExcel(j + 5, 2)
                       
                    End If
                    
                    If Check1.Value = 1 Then
                          
                       ColEnd = VecDiaExcel(j + 1, 2)
                       
                    End If
                    Exit For
                    
                 End If
             
             Next j
             
             '-------> Corte x estructura servicio
             '-------> Mover Estructura servicio excel
             If CodEstructura <> RS!ess_codigo Then
                
                MoverDatosExcel oExcel, oSheet, "A", "A", RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, Trim(RS!ess_nombre)
                CodEstructura = RS!ess_codigo
             
             End If

             If Check3.Value = 1 Then
                
                MoverDatosExcel oExcel, oSheet, oColCla, oColCla, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, LimpiaDato(Trim(Ceco)) & ";" & CodigoRegimen & ";" & CodigoServicio & ";" & RS!mid_codrec & ";" & RS!min_fecmin & ";" & RS!mid_numlin
             
             End If
             
             '-------> Mover recetas excel
             MoverDatosExcel oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, Trim(RS!rec_nombre & " " & RS!mid_codrec)
             
             '-------> Mover raciones excel
             If Check1.Value = 1 Then
                
                oCol = VecDiaExcel(IndCol + 1, 2)
                IndCol = IndCol + 1
                MoverDatosExcel oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, RS!mid_numrac
             
             End If
             
             '-------> Mover % ponderación excel
             If Check2.Value = 1 Then
                
                oCol = VecDiaExcel(IndCol + 1, 2)
                MoverDatosExcel oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, RS!mid_porrac & " %"
             
             End If
             
             '-------> Mover costo receta
             oCol = VecDiaExcel(IndCol + 2, 2)
             oColRec = VecDiaExcel(IndCol - 1, 2)
             'MoverDatosExcel oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, BuscarCostoReceta(RS!mid_codrec)
             MoverDatosExcelBuscarVI oExcel, oSheet, oColRec, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin
             '-------> Mover costo ponderado receta
             oCol = VecDiaExcel(IndCol + 3, 2)
             
             If RS!min_racteo > 0 Then
                
                CostoPlanificacion = CostoPlanificacion + ((BuscarCostoReceta(RS!mid_codrec) * RS!mid_numrac) / RS!min_racteo)
                CostoReceta = CostoReceta + (BuscarCostoReceta(RS!mid_codrec) * RS!mid_numrac)
                
             End If
'             If RS!min_racteo > 0 Then
'
'                MoverDatosExcel oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, Format(((BuscarCostoReceta(RS!mid_codrec) * RS!mid_numrac) / RS!min_racteo), fg_Pict(6, 2))
'
'             Else
'
'                MoverDatosExcel oExcel, oSheet, oCol, oCol, RowSheet + RS!mid_numlin, RowSheet + RS!mid_numlin, 0
'
'             End If
             
             If RS!mid_numlin > RowEnd Then
                
                RowEnd = RS!mid_numlin
                   
             End If
             
             RS.MoveNext
             
          Loop
          RS.Close
          Set RS = Nothing
          
          '-------> Mover comensales totales
          If AuxFec > 0 And Check1.Value = 1 Then
             
             MoverDatosExcel oExcel, oSheet, "A", "A", RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, "Comensales"
             MoverDatosExcel oExcel, oSheet, ColEnd, ColEnd, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, CStr(TotDiaRaciones)
          
             CalComen = CalComen & ColEnd & RowSheet + RowEnd1 + 2 & "+"
              
             If Check3.Value = 1 Then
                
                MoverDatosExcel oExcel, oSheet, oColCla, oColCla, RowSheet + RowEnd1 + 2, RowSheet + RowEnd1 + 2, LimpiaDato(Trim(Ceco)) & ";" & CodigoRegimen & ";" & CodigoServicio & ";" & CStr(AuxFec)
           
             End If
            
             If Check2.Value = 1 And Check3.Value = 1 Then
             
             '-------> Buscar dia vector dia excel
             For j = 1 To UBound(VecDiaExcel)
                 
                 If ColEnd = VecDiaExcel(j, 2) Then
                    
                    oColFec = VecDiaExcel(j - 1, 2)
                    oColPor = VecDiaExcel(j + 1, 2)
                    oColCos = VecDiaExcel(j + 2, 2)
                    oColCpo = VecDiaExcel(j + 3, 2)
                    Exit For
                 
                 End If
             
             Next j
                   
             For ii = 9 To RowSheet + RowEnd1 + 2
                 
                 MoverDatosExcelFormula oExcel, oSheet, oColFec, oColPor, ColEnd, ColEnd, ii, RowSheet + RowEnd1 + 2, 0
                 MoverDatosExcelFormulaIII oExcel, oSheet, oColFec, oColCpo, ColEnd, oColCos, ii, RowSheet + RowEnd1 + 2, 0
             
                 If ii < RowSheet + RowEnd1 + 2 Then
                          
                    If Trim(oSheet.Range(ColEnd & ii).Value) <> "" Then
                    
                       CalRacCos = CalRacCos & ColEnd & ii & "*" & oColCos & ii & "+"
                    
                    End If
                    
                 End If
                 
             Next ii
          
             MoverDatosExcelFormulaSum oExcel, oSheet, oColCpo, oColCpo, oColCpo, oColCpo, RowSheet + 1, RowSheet + RowEnd1 + 2, ""
'             MoverDatosExcel oExcel, oSheet, oColCpo, oColCpo, 5, 5, CStr(CostoComercial)
             MoverDatosExcel oExcel, oSheet, oColCpo, oColCpo, 6, 6, CStr(Format(CostoPlanificacion, fg_Pict(6, 2)))
             MoverDatosExcelCostoBandeja oExcel, oSheet, oColCos, RowSheet + RowEnd1 + 4, "(" & Mid(CalRacCos, 1, Len(CalRacCos) - 1) & ")"
             TotCalRacCos = TotCalRacCos & oColCos & RowSheet + RowEnd1 + 4 & "+"
             CalRacCos = ""
             
             CostoPlanificacion = 0
             BloquearColumnaExcel oExcel, oSheet, ColEnd, RowSheet + 1, RowSheet + RowEnd1 + 2
             BloquearColumnaExcel oExcel, oSheet, oColFec, RowSheet + 1, RowSheet + RowEnd1 + 2
             FormatearColumnaNumericoExcel oExcel, oSheet, ColEnd
             FormatearColumnaPorcentajeExcel oExcel, oSheet, oColPor
             
            End If
            
             End If
          
             '-------> Dibujar lineas
             DibujarLineas oExcel, oSheet, "A", "A", 8, RowSheet + RowEnd1 + 2
             For j = 1 To UBound(VecDiaExcel)
              
                 oCol = VecDiaExcel(j, 2)
                 DibujarLineas oExcel, oSheet, oCol, oCol, 5, RowSheet + RowEnd1 + 2
          
             Next j
                      
             oSheet.Cells.Select
             oSheet.Cells.EntireColumn.AutoFit
             RowSheet = 8

             'Ocultar columna clave
             If Check3.Value = 1 Then
          
                ClaveExcel = "Jp123456"
             
                Set RS1 = vg_db.Execute("sgpadm_s_parametro 1, 'parhojaexc', ''")
                If Not RS1.EOF Then
                
                   ClaveExcel = RS1(0)
             
                End If
                RS1.Close
                Set RS1 = Nothing
             
                MoverDatosExcel oExcel, oSheet, "B", "B", 2, 2, "Costo Bandeja Planificado"
                MoverDatosExcel oExcel, oSheet, "B", "B", 3, 3, "Costo Bandeja Sitio"
                If Comensales > 0 Then
                   
                   MoverDatosExcel oExcel, oSheet, "F", "F", 2, 2, Format((CostoReceta / Comensales), fg_Pict(6, 2))
                
                Else
                
                   MoverDatosExcel oExcel, oSheet, "F", "F", 2, 2, 0
                   
                End If
                MoverDatosExcelCostoBandeja oExcel, oSheet, "F", 3, "(" & Mid(TotCalRacCos, 1, Len(TotCalRacCos) - 1) & ")" & "/" & "(" & Mid(CalComen, 1, Len(CalComen) - 1) & ")"
                
                For j = 1 To UBound(VecDiaExcel) Step 6 '4
                 
                    oCol = VecDiaExcel(j + 5, 2) 'VecDiaExcel(j + 3, 2)
                    OcultarColumna oExcel, oSheet, oCol, oCol
                 
                    'Mover clave para poder actualizar
                 
                    MoverDatosExcel oExcel, oSheet, oCol, oCol, 1, 1, ClaveExcel
                    MoverDatosExcel oExcel, oSheet, oCol, oCol, 2, 2, CStr(((TotCol - 1) * 4) + 1)
             
                Next j
             
             End If
          
             'Bloquear protección hoja

             fg_descarga
         
          Else
          
             RS.Close
             Set RS = Nothing
             fg_descarga
             vaSpread1.Col = 9
             vaSpread1.text = "No existe información"
             'MsgBox "No existe información", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       
          End If
          
    End If

Next i

If Trim(auxceco) <> "" Then
             
'   '-------> incluir recetas
'   If Check5.Value = 1 Then
'
'         '-------> Add data to cells of the first worksheet in the new workbook
'         NombreServicio = "Recetas"
'         Set oSheet = oBook.Worksheets.Add
'         NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
'
'         If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
'
'            NombreServicioAux = Mid(NombreServicioAux, 1, 31)
'
'         End If
'         oSheet.Name = "Recetas" 'Trim(NombreServicioAux)
'
'         '-------> Check version of Excel
'         Call encabezado(Rs_Receta, oSheet)
'
'         oSheet.Cells(2, 1).CopyFromRecordset Rs_Receta
'         IndCol = 1
'
'         MoverDatosExcel oExcel, oSheet, "A", "A", IndCol, IndCol, Trim("Código")
'         MoverDatosExcel oExcel, oSheet, "B", "B", IndCol, IndCol, Trim("Nombre Plato")
'         MoverDatosExcel oExcel, oSheet, "C", "C", IndCol, IndCol, Trim("Costo Plato")
'         MoverDatosExcel oExcel, oSheet, "D", "D", IndCol, IndCol, Trim("Categoria Dietetica")
'         MoverDatosExcel oExcel, oSheet, "E", "E", IndCol, IndCol, Trim("Tipo Plato")
'         PonerColorInteriorN oExcel, oSheet, "A", "E", IndCol, IndCol, 4
'
'         IndCol = 2
'         PonerColorInteriorN oExcel, oSheet, "A", "E", IndCol, Rs_Receta.RecordCount + 1, 6
'
'         oSheet.Cells.Select
'         oSheet.Cells.EntireColumn.AutoFit
'
'         'Bloquear protección hoja
'         If Check2.Value = 1 And Check3.Value = 1 Then
'
'            With oSheet
'
'                  .AutoFilterMode = False
'
'                  .Range("A1:D1").AutoFilter
'
'            End With
'
'         End If
'
'   End If
   
          If Check2.Value = 1 And Check3.Value = 1 Then
             
'             ObjW.Sheets(i).Select
'             ObjW.Sheets(i).Protect password:=ClaveExcel, DrawingObjects:=True, _
'                            Contents:=True, Scenarios:=True, AllowFormattingCells:=True
                                 
                oSheet.Select
                oExcel.ActiveSheet.Protect password:=ClaveExcel, DrawingObjects:=True, _
                                    Contents:=True, Scenarios:=True, AllowFormattingCells:=True
          
          End If
   
   oBook.Close True, NomArchivoExcel
          
   Set oSheet = Nothing
   Set oBook = Nothing
   Set oExcel = Nothing
   Set ObjW = Nothing
   Set ObjExcel = Nothing
   
'   Set ObjExcel = New excel.Application
'   Set ObjW = ObjExcel.Workbooks.Open(NomArchivoExcel)
    
'   For i = 1 To ObjW.Sheets.count
 
'       If Trim(ObjW.Sheets(i).Name) <> "Recetas" And Trim(ObjW.Sheets(i).Name) <> "Hoja1" Then
        
'          ObjW.Sheets(i).Cells.Replace What:="Hoja1", Replacement:="recetas", LookAt:=xlPart, _
'          SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
                                      
 '      End If
    
 '  Next i
    
 '  ObjW.Save
 '  ObjW.Application.DisplayAlerts = False
 '  ObjW.Close True
 '  ObjExcel.Quit
   
   Set oSheet = Nothing
   Set oBook = Nothing
   Set oExcel = Nothing
   Set ObjW = Nothing
   Set ObjExcel = Nothing
'   ActiveWindow.Close
             
             
End If

'oExcel.Visible = True '------->Visualizar
Set oSheet = Nothing
Set oExcel = Nothing
Set oBook = Nothing
Set ObjW = Nothing
Set ObjExcel = Nothing

Bar1(0).Value = 0
Bar1(0).Visible = False

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
End Sub






Function ReportePorLotes(TipoReporte As Integer)

On Error GoTo Man_Error
Dim RS As Recordset
Dim varSql As String
Dim i As Integer

Dim IdColaTrabajo As Integer
Dim varCeco As String
Dim varRegimen As String
Dim varServicio As String
Dim varNumLin As Integer
IdColaTrabajo = 0


'TipoReporte = 1 --> ExportarExcelMinutaBloqueSinCostoReceta
'TipoReporte = 2 --> ExportarExcelMinutaBloqueSinCodigoReceta
'TipoReporte = 3 -->



'Format(FpFecDesde.text, "yyyymmdd")

varSql = ""
varSql = varSql & "'" & vg_NUsr & "', '" & fpText.text & "'," & TipoReporte & ", "
varSql = varSql & "'" & Format(FpFecDesde.text, "yyyymmdd") & "', '" & Format(FpFecHasta.text, "yyyymmdd") & "', " & Check1.Value & ", " & Check2.Value & ", "
varSql = varSql & Check3.Value & ", " & Check4.Value & ", " & IIf(Option1(0).Value = True, 1, 0) & ", " & IIf(Option1(1).Value = True, 1, 0) & ", "
varSql = varSql & Check5.Value & ", " & Check6.Value & ", 1, 0"

Set RS = vg_db.Execute("EXEC SGP_I_IngresaTareasPorLotes " & varSql)
If Not RS.EOF Then
    IdColaTrabajo = RS(0)
End If
RS.Close: Set RS = Nothing

varNumLin = 0
i = 0
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1
    
    If vaSpread1.text = 1 Then
        
        vaSpread1.Col = 2
        varCeco = vaSpread1.text
        
        vaSpread1.Col = 4
        varRegimen = vaSpread1.text
        
        vaSpread1.Col = 6
        varServicio = vaSpread1.text
        
        varNumLin = varNumLin + 1
        
        vg_db.Execute "EXEC SGP_I_IngresaDetalleTareasPorLotes " & IdColaTrabajo & ", " & varNumLin & ", '" & varCeco & "', " & varRegimen & ", " & varServicio
        
        vaSpread1.Col = 9
        vaSpread1.text = "Servicio ingresado a tarea por lotes"
        
    End If

Next i

Call Rellena_CatDietetica_TipoPlato(IdColaTrabajo)

ReportePorLotes = IdColaTrabajo

Exit Function
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo
Resume
End Function



Sub Rellena_CatDietetica_TipoPlato(varIdColaTrabajo As Integer)
Dim RS    As New ADODB.Recordset
Dim isel As Integer
Dim spid As Integer
Dim IndFiltro As Integer

spid = 0
isel = 0
IndFiltro = 0

'-------> Borrar tabla paso_catdie_tipoplato
vg_db.Execute "DELETE paso_catdie_tipoplato WHERE pdp_spid = @@spid AND pdp_user = '" & vg_NUsr & "' AND pdp_idcola = " & varIdColaTrabajo

'-------> Buscar spid
Set RS = vg_db.Execute("SELECT @@spid spid")
If Not RS.EOF Then spid = RS!spid
RS.Close: Set RS = Nothing

For IndFiltro = 1 To B_DieTipExcel.TvwDir(0).Nodes.count

    If B_DieTipExcel.TvwDir(0).Nodes.item(IndFiltro).Checked = True And Trim(B_DieTipExcel.TvwDir(0).Nodes.item(IndFiltro).text) <> "*" Then
          
       vg_db.Execute "INSERT INTO paso_catdie_tipoplato (pdp_spid, pdp_user, pdp_idcola, pdp_codigo, pdp_tipo) VALUES (" & spid & ", '" & vg_NUsr & "', " & varIdColaTrabajo & ", " & Val(Mid(B_DieTipExcel.TvwDir(0).Nodes.item(IndFiltro).text, 1, InStr(B_DieTipExcel.TvwDir(0).Nodes.item(IndFiltro).text, " - ") - 1)) & ", 'CatDie'" & ")"

    End If
       
Next IndFiltro

IndFiltro = 0

For IndFiltro = 1 To B_DieTipExcel.TvwDir(1).Nodes.count

    If B_DieTipExcel.TvwDir(1).Nodes.item(IndFiltro).Checked = True And Trim(B_DieTipExcel.TvwDir(1).Nodes.item(IndFiltro).text) <> "*" Then
                 
       vg_db.Execute "INSERT INTO paso_catdie_tipoplato (pdp_spid, pdp_user, pdp_idcola, pdp_codigo, pdp_tipo) VALUES (" & spid & ", '" & vg_NUsr & "', " & varIdColaTrabajo & ", " & Val(Mid(B_DieTipExcel.TvwDir(1).Nodes.item(IndFiltro).text, 1, InStr(B_DieTipExcel.TvwDir(1).Nodes.item(IndFiltro).text, "-") - 1)) & ", 'TipPla'" & ")"

    End If
       
Next IndFiltro

fg_descarga

End Sub


