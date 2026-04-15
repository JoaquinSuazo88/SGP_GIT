VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form E_TemplateMinI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Template Minuta Bloque"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   8595
      Left            =   105
      TabIndex        =   10
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
         TabIndex        =   23
         Top             =   1080
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
         TabIndex        =   22
         Top             =   1080
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Index           =   1
         Left            =   2640
         TabIndex        =   21
         Top             =   7800
         Width           =   2475
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   5
            Top             =   135
            Width           =   2370
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   1
         Left            =   1560
         TabIndex        =   20
         Top             =   7800
         Width           =   930
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   4
            Top             =   135
            Width           =   825
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   8640
         TabIndex        =   18
         Top             =   7800
         Width           =   930
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   45
            TabIndex        =   8
            Top             =   135
            Width           =   825
         End
      End
      Begin VB.Frame Frame2 
         Height          =   435
         Left            =   9615
         TabIndex        =   17
         Top             =   7800
         Width           =   3630
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   7
            Left            =   45
            TabIndex        =   9
            Top             =   135
            Width           =   3525
         End
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Index           =   0
         Left            =   6150
         TabIndex        =   15
         Top             =   7800
         Width           =   2475
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   45
            TabIndex        =   7
            Top             =   135
            Width           =   2370
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   0
         Left            =   5175
         TabIndex        =   14
         Top             =   7800
         Width           =   930
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   6
            Top             =   135
            Width           =   825
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5850
         Left            =   105
         TabIndex        =   3
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
         SpreadDesigner  =   "E_TemplateMinI.frx":0000
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
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   105
         TabIndex        =   19
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   12840
         TabIndex        =   24
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   270
         Width           =   1155
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8790
      Left            =   13980
      TabIndex        =   16
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "E_TemplateMinI.frx":1A29
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "E_TemplateMinI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public lc_Aux    As String
Dim BtnX         As Object
Dim MsgTitulo As String

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

Me.Height = 9225
Me.Width = 14700

fg_centra Me

MsgTitulo = "Template Minuta Bloque"

Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exportar Excel": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

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


Select Case Button.Index

Case 1
    
    E_TemplateMinutaBloque
    
Case 3
    
    Me.Hide
    Unload Me

End Select

Exit Sub

Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub E_TemplateMinutaBloque()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim Sql             As String
Dim NomArchivoExcel As String
Dim Extension       As String

Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object
Dim i               As Long
Dim MyBuffer        As String
Dim seleccion       As String
Dim Ceco            As String
Dim Regimen         As Long
Dim Servicio        As Long
Dim EstSeleccion    As Boolean

If Not ValidarDatos Then Exit Sub

If vaSpread1.MaxRows = 0 Then
           
   MsgBox "No existe datos selecionado en la grilla...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub
        
End If

EstSeleccion = False

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 1 'Seleccion
    seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
    If seleccion = 1 And vaSpread1.RowHidden = False Then
          
       EstSeleccion = True
       
    End If
    
Next i

If Not EstSeleccion Then
           
   MsgBox "Debe haber a lo menos un dato seleccionado en la grilla...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub
        
End If


Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<Min>"
  
For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 1 'Seleccion
    seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
    If seleccion = 1 And vaSpread1.RowHidden = False Then
          
       Ceco = ""
       vaSpread1.Col = 2
       Ceco = vaSpread1.text
          
       Regimen = 0
       vaSpread1.Col = 4
       Regimen = vaSpread1.text
          
       Servicio = 0
       vaSpread1.Col = 6
       Servicio = vaSpread1.text
          
       MyBuffer = MyBuffer & " <MinDet"
       MyBuffer = MyBuffer & " Cec = " & Chr(34) & Ceco & Chr(34)
       MyBuffer = MyBuffer & " Reg = " & Chr(34) & Regimen & Chr(34)
       MyBuffer = MyBuffer & " Ser = " & Chr(34) & Servicio & Chr(34)
       MyBuffer = MyBuffer & "/>"

    End If
    
Next i

MyBuffer = MyBuffer & "</Min>"

'-------> Validar cantidad registro se sobre pase hoja excel
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Option1(0).Value = True Then

   Set RS = vg_db.Execute("sgpadm_Sel_XmlTemplateExcelPlantillaFrecuencia '" & MyBuffer & "', '" & fpText.text & "', " & Format(FpFecDesde.text, "yyyymmdd") & ", " & Format(FpFecHasta.text, "yyyymmdd") & "")

ElseIf Option1(1).Value = True Then

   Set RS = vg_db.Execute("sgpadm_Sel_XmlTemplateExcelPonderacionesxEstructura '" & MyBuffer & "', '" & fpText.text & "', " & Format(FpFecDesde.text, "yyyymmdd") & ", " & Format(FpFecHasta.text, "yyyymmdd") & "")

End If

If Not RS.EOF Then
  
   If RS.RecordCount > 1020000 Then
      
      RS.Close
      Set RS = Nothing
      
      MsgBox "El resultado sobrepasa maximo de fila en excel, Debera seleccionar menos Ceco", vbCritical
      Exit Sub
   
   End If
  
End If

'-------> Guardar nombre archivo excel
NomArchivoExcel = ""
CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
CD.Filter = "Todos los archivos *.xls,*.xlsx"
On Error Resume Next
CD.ShowSave
           
'-------> JPAZ Permite controlar Boton Cancelar
If Err.Number = 32755 Then
   
   MsgBox "Proceso cancelado"
   Exit Sub

End If
            
If CD.FileName = "" Then
   
   MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
   Exit Sub

Else
   
   Extension = ""
   Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
   
   If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
      MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
      Exit Sub
   End If
   
   NomArchivoExcel = CD.FileName

End If
          
fg_carga ""
  
'-------> Create an instance of Excel and add a workbook
Set xlApp = CreateObject("Excel.Application")
Set xlWb = xlApp.Workbooks.Add
Set xlWs = xlWb.Worksheets("Hoja1")
  
'-------> Display Excel and give user control of Excel's lifetime
xlApp.UserControl = True
    
'-------> Check version of Excel
Call encabezado(RS, xlWs)
          
xlWs.Cells(2, 1).CopyFromRecordset RS

'-------> Auto-fit the column widths and row heights
xlApp.Selection.CurrentRegion.Columns.AutoFit
xlApp.Selection.CurrentRegion.Rows.AutoFit
    
'xlApp.Columns("A:A").Select
'xlApp.Selection.Delete Shift:=xlToLeft
  
xlWb.Close True, NomArchivoExcel

Dim XL As New excel.Application 'Crea el objeto excel
XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
XL.Visible = True
XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
'-------> Close ADO objects
RS.Close
Set RS = Nothing
    
' -- Cerrar Excel
xlApp.Quit
'-------> Release Excel references
Set xlWs = Nothing
Set xlWb = Nothing
Set xlApp = Nothing
  
fg_descarga
MsgBox "Proceso Finalizado", vbInformation, Me.Caption
                
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
        
        If Not ValidarDatos Then Exit Sub
        
        vaSpread1.Visible = False
        vaSpread1.MaxRows = 0
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
            
        Sql = ""
        Sql = Sql & LimpiaDato(Trim(fpText.text))
        Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
        Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
        Set RS = vg_db.Execute("sgpadm_Sel_OrgComprasCecoMBloqueTemplate_V01 " & Sql & "")
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
               
        vaSpread1.Visible = True
    
    End Select
    
Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Function ValidarDatos() As Boolean

On Error GoTo Man_Error

Dim RS  As New ADODB.Recordset

ValidarDatos = True
        
        If Trim(fpText.text) = "" Then
           
           ValidarDatos = False
           MsgBox "Debe seleccionar Org. Compras...", vbExclamation + vbOKOnly, MsgTitulo: Exit Function
           Exit Function
        
        End If
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("sgpadm_Sel_OrgCompras_V02 '" & LimpiaDato(fpText.text) & "'")
        
        If RS.EOF Then
           
           ValidarDatos = False
           RS.Close
           Set RS = Nothing
           MsgBox "No existe Org. Compras...", vbExclamation + vbOKOnly, MsgTitulo: Exit Function
           vaSpread1.MaxRows = 0

           Exit Function
        
        End If
        
        RS.Close
        Set RS = Nothing
        
        If CDate(FpFecDesde.text) > CDate(FpFecHasta.text) Then
            
            ValidarDatos = False
            Call MsgBox("Fecha Desde No Puede Ser Mayor a Fecha Hasta", vbInformation, Me.Caption)
            Let FpFecDesde.text = Format(Now, "dd/mm/yyyy")
            Call FpFecDesde.SetFocus
            Exit Function
        
        End If
    
        If CDate(FpFecHasta.text) < CDate(FpFecDesde.text) Then
            
            ValidarDatos = False
            Call MsgBox("Fecha Hasta No Puede Ser Mayor a Fecha Desde", vbInformation, Me.Caption)
            Let FpFecHasta.text = Format(Now, "dd/mm/yyyy")
            Call FpFecHasta.SetFocus
            Exit Function
        
        End If
        
        If DateDiff("y", FpFecDesde.text, FpFecHasta.text) > 365 Then
            
            Call MsgBox("Rango De Fecha No Puede Ser Mayor a 12 Meses", vbInformation, Me.Caption)
            ValidarDatos = False
            vaSpread1.MaxRows = 0
            Exit Function
        
        End If

Exit Function
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Function

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

