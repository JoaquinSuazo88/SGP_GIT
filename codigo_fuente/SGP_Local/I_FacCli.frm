VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_FacCli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación Cliente"
   ClientHeight    =   3855
   ClientLeft      =   2265
   ClientTop       =   3750
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3855
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   21
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
      _Version        =   393216
      _ExtentX        =   3201
      _ExtentY        =   661
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
      MaxCols         =   3
      MaxRows         =   1
      SpreadDesigner  =   "I_FacCli.frx":0000
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3390
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   375
      Width           =   8520
      Begin VB.Frame Frame1 
         Caption         =   "Servicios"
         Height          =   780
         Index           =   2
         Left            =   4320
         TabIndex        =   22
         Top             =   1650
         Width           =   4095
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   225
            Index           =   7
            Left            =   330
            TabIndex        =   24
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
            Height          =   225
            Index           =   6
            Left            =   2445
            TabIndex        =   23
            Top             =   360
            Width           =   795
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   1
            Left            =   3360
            Picture         =   "I_FacCli.frx":02B1
            Top             =   200
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Regimen"
         Height          =   780
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1650
         Width           =   4095
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
            Height          =   225
            Index           =   5
            Left            =   2445
            TabIndex        =   18
            Top             =   360
            Width           =   795
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   225
            Index           =   4
            Left            =   330
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   3
            Left            =   3360
            Picture         =   "I_FacCli.frx":05BB
            Top             =   200
            Width           =   480
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   8295
         Begin VB.OptionButton Option1 
            Caption         =   "Resumido"
            Height          =   225
            Index           =   2
            Left            =   255
            TabIndex        =   3
            Top             =   1050
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Detallado"
            Height          =   225
            Index           =   3
            Left            =   3810
            TabIndex        =   4
            Top             =   1050
            Width           =   1215
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   1
            Left            =   1860
            TabIndex        =   0
            Top             =   210
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            ButtonStyle     =   0
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
            AutoCase        =   0
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
            CharValidationText=   ""
            MaxLength       =   255
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   0
            Left            =   1860
            TabIndex        =   1
            Top             =   570
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Text            =   "17/08/2004"
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
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
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
            Left            =   5250
            TabIndex        =   2
            Top             =   570
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Text            =   "17/08/2004"
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
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
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
            Caption         =   "Fecha Termino"
            Height          =   195
            Index           =   4
            Left            =   3810
            TabIndex        =   14
            Top             =   615
            Width           =   1275
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio"
            Height          =   195
            Index           =   3
            Left            =   255
            TabIndex        =   13
            Top             =   615
            Width           =   1065
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3795
            TabIndex        =   12
            Top             =   225
            Width           =   4335
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   3250
            Picture         =   "I_FacCli.frx":08C5
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Contrato"
            Height          =   225
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   300
            Width           =   1560
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   3840
            TabIndex        =   15
            Top             =   270
            Width           =   4335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Clientes"
         Height          =   780
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   2490
         Width           =   4095
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   225
            Index           =   0
            Left            =   330
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
            Height          =   225
            Index           =   1
            Left            =   2445
            TabIndex        =   5
            Top             =   360
            Width           =   915
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   2
            Left            =   3360
            Picture         =   "I_FacCli.frx":0BCF
            Top             =   200
            Width           =   480
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
      _Version        =   393216
      _ExtentX        =   3201
      _ExtentY        =   661
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
      MaxCols         =   3
      MaxRows         =   1
      SpreadDesigner  =   "I_FacCli.frx":0ED9
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   375
      Index           =   2
      Left            =   6240
      TabIndex        =   20
      Top             =   3240
      Visible         =   0   'False
      Width           =   1935
      _Version        =   393216
      _ExtentX        =   3413
      _ExtentY        =   661
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
      MaxCols         =   3
      MaxRows         =   1
      SpreadDesigner  =   "I_FacCli.frx":118A
   End
End
Attribute VB_Name = "I_FacCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim MsgTitulo As String, est As Boolean

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Me.Width = 8655
Me.Height = 4335
MsgTitulo = "Facturación Clientes"
est = True
EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)
fg_centra Me
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
'-------------------------Asigna fecha actual del sistema para informe-------------
fpDateTime1(0).text = Date: fpDateTime1(1).text = Date
Option1(0).Value = True
fpText1(1).Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText1(1).text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
est = False
MoverVector
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If est Then Exit Sub
If IsDate(fpDateTime1(0).text) Then
    If fpDateTime1(0).DateValue > fpDateTime1(1).DateValue Then fpDateTime1(1).text = fpDateTime1(0).text: Exit Sub
End If
Select Case Index
Case 0
    If fpDateTime1(0).text = "" Then
       fpDateTime1(1).Enabled = False
       fpDateTime1(1).text = ""
       Exit Sub
    Else
       fpDateTime1(1).Enabled = True
    End If
End Select
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_Change(Index As Integer)
If est Then Exit Sub
RS1.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText1(1).text)), ""), vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(0).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS1!cli_nombre)
RS1.Close: Set RS1 = Nothing
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Image1_Click(Index As Integer)
vg_codigo = 0
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 1920
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    fpText1(1).text = Trim(vg_codigo)
    fpayuda(0).Caption = vg_nombre
Case 2
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpText1(1).text = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Cliente", Me.vaSpread1, fpText1(1).text, 0, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "3", "FacCli", 0, "'1'"
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
Case 1
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpText1(1).text = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, fpText1(1).text, "", Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "0", "FacCli", 2, "'1'"
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
Case 3
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpText1(1).text = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Regimen", Me.vaSpread1, fpText1(1).text, "", Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "0", "FacCli", 1, "'1'"
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    Image1(2).Enabled = False
    With vaSpread1(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1: .text = "1"
        Next i
    End With
Case 1
    Image1(2).Enabled = True
Case 4
    Image1(3).Enabled = False
    With vaSpread1(1)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1: .text = "1"
        Next i
    End With
Case 5
    Image1(3).Enabled = True
Case 7
    Image1(1).Enabled = False
    With vaSpread1(2)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1: .text = "1"
        Next i
    End With
Case 6
    Image1(1).Enabled = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim sql1 As String, sql2 As String, sql3 As String, sql4 As String, sql5 As String, fecini As Long, fecter As Long, codser As String, codreg As String, codcli As String
Dim sqlSE As String, sqlRE As String, cencos As String, nArch As String, i As Long
On Error GoTo Error_Salir
Select Case Button.Index
Case 1
    cencos = Trim(fpText1(1).text)
    codcli = "": est = True
    codreg = "": codser = "": codcli = ""
    With vaSpread1(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .text = "1" Then est = False: .Col = 2: codcli = codcli & "'" & .text & "',"
        Next i
    End With
    If Trim(codcli) = "" Then fg_descarga: MsgBox "Cliente debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    With vaSpread1(1)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .text = "1" Then .Col = 2: codreg = codreg & "" & .text & ","
        Next i
    End With
    If Trim(codreg) = "" Then fg_descarga: MsgBox "Regimen debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    With vaSpread1(2)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .text = "1" Then .Col = 2: codser = codser & "" & .text & ","
        Next i
    End With
    If Trim(codser) = "" Then fg_descarga: MsgBox "Servicio debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fecini = Val(Format(fpDateTime1(0).text, "yyyy") & Format(fpDateTime1(0).text, "mm") & Format(fpDateTime1(0).text, "dd"))
    fecter = Val(Format(fpDateTime1(1).text, "yyyy") & Format(fpDateTime1(1).text, "mm") & Format(fpDateTime1(1).text, "dd"))
    sql1 = " AND ser.ser_codigo IN (" & Mid(codser, 1, Len(codser) - 1) & ")"
    sql2 = IIf(Option1(2).Value = True, "SUM(0) AS mir_fecmin, SUM(mir.mir_nrorac) AS cantidad", "mir.mir_nrorac AS cantidad, mir.mir_fecmin")
    sql3 = IIf(Option1(2).Value = True, "GROUP BY cli.cli_codigo, cli.cli_nombre, ser.ser_codigo, ser.ser_nombre, mir.prv_fecvig, prv.prv_preven ORDER BY cli.cli_codigo, ser.ser_codigo", "ORDER BY cli.cli_codigo, ser.ser_codigo, mir_fecmin")
    sql4 = IIf(Option1(1).Value = True, " AND mir.mir_rutcli IN (" & Mid(codcli, 1, Len(codcli) - 1) & ")", "")
    est = False
    nArch = Trim(vg_NUsr) & "_tmp_fact1"
    fg_CheckTmp (nArch)
    vg_db.Execute "SELECT mir.mir_cencos, mir.mir_codreg, mir.mir_codser, mir.mir_fecmin, mir.mir_rutcli, 0 AS mir_codcco, mir.mir_nrorac, max(prv.prv_fecvig) as prv_fecvig into " & nArch & " " & _
                  "FROM b_minutaraciones mir INNER JOIN b_preciovta prv ON (mir.mir_rutcli = prv.prv_rutcli) AND (mir.mir_codser = prv.prv_codser) AND (mir.mir_codreg = prv.prv_codreg) AND (mir.mir_cencos = prv.prv_cencos AND mir.mir_fecmin >= prv.prv_fecvig) " & _
                  "WHERE mir.mir_cencos = '" & cencos & "' " & _
                  "AND mir.mir_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                  "AND mir.mir_fecmin >= " & fecini & " " & _
                  "AND mir.mir_fecmin <= " & fecter & " " & _
                  "AND mir.mir_nrorac > 0 " & sql4 & " " & _
                  "GROUP BY mir.mir_cencos, mir.mir_codreg, mir.mir_codser, mir.mir_fecmin, mir.mir_rutcli, mir.mir_nrorac"

    sql5 = "SELECT f.reg_codigo, f.reg_nombre, e.ser_codigo, e.ser_nombre " & IIf(Option1(2).Value = False, ",a.vtc_fecvta, c.cli_codigo, c.cli_nombre, d.clc_codigo, d.clc_nombre, b.vtd_descripcion", ", 0 AS vtc_fecvta, c.cli_codigo, c.cli_nombre, '' AS clc_codigo, '' AS clc_nombre, '' AS vtd_descripcion") & " , SUM(b.vtd_detmon) AS vtd_detmon " & _
           "FROM b_ventacontado a, b_ventacontadodet b, b_clientes c, b_clientecencos d, a_servicio e, a_regimen f " & _
           "WHERE a.vtc_codigo = b.vtd_codigo " & _
           "AND   a.vtc_codreg = f.reg_codigo " & _
           "AND   a.vtc_codser = e.ser_codigo " & _
           "AND   b.vtd_codcli = c.cli_codigo " & _
           "AND   b.vtd_codcli = d.clc_codcli " & _
           "AND   b.vtd_codcco = d.clc_codigo " & _
           "AND   a.vtc_cencos = '" & cencos & "' " & _
           "AND   a.vtc_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
           "AND   a.vtc_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
           "AND   a.vtc_fecvta >= " & fecini & " " & _
           "AND   a.vtc_fecvta <= " & fecter & " GROUP BY f.reg_codigo, f.reg_nombre, e.ser_codigo, e.ser_nombre " & IIf(Option1(2).Value = False, ",a.vtc_fecvta, c.cli_codigo, c.cli_nombre, d.clc_codigo, d.clc_nombre, vtd_descripcion", ", c.cli_codigo, c.cli_nombre") & " " & _
           "ORDER BY f.reg_codigo, e.ser_codigo " & IIf(Option1(2).Value = False, ", d.clc_codigo, a.vtc_fecvta", "") & ""
           
    sql1 = "SELECT cli.cli_codigo, cli.cli_nombre, ser.ser_codigo, ser.ser_nombre, mir.prv_fecvig, prv.prv_preven, " & sql2 & " " & _
           "FROM a_servicio ser, " & nArch & " mir, b_preciovta prv, b_clientes cli, a_regimen reg " & _
           "WHERE mir.mir_rutcli = cli.cli_codigo " & _
           "AND   mir.mir_codser = ser.ser_codigo " & _
           "AND   mir.mir_codreg = reg.reg_codigo " & sql1 & " " & _
           "AND   mir.mir_rutcli = prv.prv_rutcli AND mir.mir_codser = prv.prv_codser AND mir.mir_codreg = prv.prv_codreg AND mir.mir_cencos = prv.prv_cencos AND mir.prv_fecvig = prv.prv_fecvig " & _
           "" & sql3 & ""
    
    I_FacturaCli Me, sql1, sql5
Case 3
    Me.Hide
    Unload Me
End Select
Exit Sub
Error_Salir:
    MsgBox "Error:" & Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    fg_descarga
End Sub

Sub MoverVector()
'-------> Mover Clientes
With vaSpread1(0)
    .MaxRows = 0
'    RS1.Open RutinaLectura.Cliente(2, fpText1(1).text, ""), vg_db, adOpenStatic
    RS1.Open RutinaLectura.Cliente(2, "", ""), vg_db, adOpenStatic
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1: .text = "1"
          .Col = 2: .text = RS1!cli_codigo
          .Col = 3: .text = Trim(RS1!cli_nombre)
          RS1.MoveNext
       Loop
    End If
    RS1.Close: Set RS1 = Nothing
End With
'-------> Mover regimen
With vaSpread1(1)
    .MaxRows = 0
    RS1.Open RutinaLectura.Regimen(1, 0, ""), vg_db, adOpenStatic
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1: .text = "1"
          .Col = 2: .text = RS1!reg_codigo
          .Col = 3: .text = Trim(RS1!reg_nombre)
          RS1.MoveNext
       Loop
    End If
    RS1.Close: Set RS1 = Nothing
End With
'-------> Mover servicio
With vaSpread1(2)
    .MaxRows = 0
    RS1.Open RutinaLectura.Servicio(1, 0, ""), vg_db, adOpenStatic
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1: .text = "1"
          .Col = 2: .text = RS1!ser_codigo
          .Col = 3: .text = Trim(RS1!ser_nombre)
          RS1.MoveNext
       Loop
    End If
    RS1.Close: Set RS1 = Nothing
End With
End Sub
