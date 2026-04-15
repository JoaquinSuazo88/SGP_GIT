VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_FCost 
   Caption         =   "Food Cost"
   ClientHeight    =   3165
   ClientLeft      =   3495
   ClientTop       =   2625
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3165
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2625
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   7875
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   1710
         Width           =   3735
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
            Left            =   2280
            TabIndex        =   20
            Top             =   300
            Width           =   735
         End
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
            TabIndex        =   19
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   3
            Left            =   3000
            Picture         =   "I_FCost.frx":0000
            Top             =   165
            Width           =   480
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
         Height          =   735
         Index           =   0
         Left            =   4050
         TabIndex        =   13
         Top             =   1710
         Width           =   3735
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
            Index           =   0
            Left            =   120
            TabIndex        =   6
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
            Index           =   1
            Left            =   2280
            TabIndex        =   7
            Top             =   300
            Width           =   735
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   2
            Left            =   3000
            Picture         =   "I_FCost.frx":030A
            Top             =   165
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   90
         TabIndex        =   9
         Top             =   120
         Width           =   7700
         Begin VB.OptionButton Option2 
            Caption         =   "Total Costo"
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
            Index           =   2
            Left            =   5160
            TabIndex        =   5
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Costo Desechable"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   2640
            TabIndex        =   4
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Costo Alimentación"
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
            Left            =   120
            TabIndex        =   3
            Top             =   960
            Value           =   -1  'True
            Width           =   1455
         End
         Begin EditLib.fpText fpText 
            Height          =   315
            Left            =   1335
            TabIndex        =   0
            Top             =   210
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
            Index           =   0
            Left            =   1350
            TabIndex        =   1
            Top             =   560
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
            Text            =   "13/07/2004"
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
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   1
            Left            =   5205
            TabIndex        =   2
            Top             =   560
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
            Text            =   "13/07/2004"
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
            Index           =   0
            Left            =   3120
            TabIndex        =   16
            Top             =   210
            Width           =   4335
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
            Index           =   7
            Left            =   120
            TabIndex        =   12
            Top             =   285
            Width           =   735
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   2580
            Picture         =   "I_FCost.frx":0614
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicial"
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
            Left            =   120
            TabIndex        =   11
            Top             =   630
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Final"
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
            Left            =   4005
            TabIndex        =   10
            Top             =   630
            Width           =   1005
         End
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3165
            TabIndex        =   17
            Top             =   255
            Width           =   4335
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   135
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   2500
         Visible         =   0   'False
         Width           =   615
         _Version        =   393216
         _ExtentX        =   1085
         _ExtentY        =   238
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
         MaxCols         =   4
         MaxRows         =   100
         SpreadDesigner  =   "I_FCost.frx":091E
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   135
      Index           =   1
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _Version        =   393216
      _ExtentX        =   1085
      _ExtentY        =   238
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
      MaxCols         =   4
      MaxRows         =   100
      SpreadDesigner  =   "I_FCost.frx":0FD2
   End
End
Attribute VB_Name = "I_FCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim i As Integer, est As Boolean
Dim MsgTitulo As String
Public lc_Aux As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 3660
Me.Width = 8175
Me.HelpContextID = vg_OpcM
Option2(0).Visible = True
Option2(1).Visible = True
est = True
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
If lc_Aux = "CosTot" Then
   MsgTitulo = "Costos Totales del Período"
   Me.Caption = "Costos Totales del Período"
   Option2(0).Visible = True
   Option2(1).Visible = True
   Option2(2).Visible = True: Option2(2).Value = True
ElseIf lc_Aux = "FooCos" Then
   MsgTitulo = "Food Cost"
   Me.Caption = "Food Cost"
   Option2(2).Visible = True
ElseIf lc_Aux = "CosSec" Then
   MsgTitulo = "CosSec x Sector"
   Me.Caption = "Costo x Sector"
   Option2(0).Caption = "Detalle"
   Option2(1).Caption = "Resumido"
   Option2(2).Visible = False
ElseIf lc_Aux = "InNPla" Then
   MsgTitulo = "Insumos no Planificados en Salida Bodega"
   Me.Caption = "Insumos no Planificados en Salida Bodega"
   Option2(0).Visible = False
   Option2(1).Visible = False
   Option2(2).Visible = False
ElseIf lc_Aux = "CosPer" Then
   MsgTitulo = "Costo Detalle Periodo Realizado"
   Me.Caption = "Costo Detalle Periodo Realizado"
   Option2(0).Caption = "Planif. Teórico"
   Option2(1).Caption = "Planif. Real"
   Option2(2).Caption = "Salida Prod."
   Option2(0).Visible = False
   Option2(1).Visible = False
   Option2(2).Visible = False
'   Option1(0).Value = True
ElseIf lc_Aux = "CurABC" Then
   MsgTitulo = "Curva ABC"
   Me.Caption = "Curva ABC"
   Option2(0).Caption = "Planif. Teórico"
   Option2(1).Caption = "Planif. Real"
   Option2(2).Caption = "Salida Prod."
   Option1(0).Value = True
ElseIf lc_Aux = "CocABC" Then
   MsgTitulo = "Comparativo Curva ABC"
   Me.Caption = "Comparativo Curva ABC"
   Option2(0).Caption = "Planif. Teórico"
   Option2(1).Caption = "Planif. Real"
   Option2(2).Caption = "Salida Prod."
   Option1(0).Value = True
ElseIf lc_Aux = "ConRac" Then
   MsgTitulo = "Comparativo de Raciones"
   Me.Caption = "Comparativo de Raciones"
   Option2(0).Visible = False
   Option2(1).Visible = False
   Option2(2).Visible = False
   Option1(0).Value = True
Else
   MsgTitulo = "Raciones no Vendidas"
   Me.Caption = "Raciones no Vendidas"
   Option2(0).Caption = "Detalle"
   Option2(1).Caption = "Resumido"
   Option2(2).Visible = False
End If
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico Planificacón Teórica"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
est = False
MoverDatoGrilla
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_Change()
RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
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
Case 2
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpText.text = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, fpText.text, "", Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "0", lc_Aux, 1, "'1'"
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
Case 3
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpText.text = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Regimen", Me.vaSpread1, fpText.text, "", Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "0", lc_Aux, 0, "'1'"
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
'    MoverDatoGrilla
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    Image1(2).Enabled = False
    For i = 1 To vaSpread1(1).MaxRows
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1: vaSpread1(1).text = "1"
    Next i
Case 1
    Image1(2).Enabled = True
Case 2
    Image1(3).Enabled = False
    For i = 1 To vaSpread1(0).MaxRows
        vaSpread1(0).Row = i
        vaSpread1(0).Col = 1: vaSpread1(0).text = "1"
    Next i
Case 3
    Image1(3).Enabled = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codser As String, codreg As String
Select Case Button.Index
Case 1
    
    RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda(0).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    If fpDateTime1(0).Value > fpDateTime1(1).Value Then MsgBox "Fecha origen Mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Mid(fpDateTime1(0).text, 4, 2) <> Mid(fpDateTime1(1).text, 4, 2) And lc_Aux <> "ConRac" Then MsgBox "Mes origen mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Mid(fpDateTime1(0).text, 7, 4) <> Mid(fpDateTime1(1).text, 7, 4) And lc_Aux <> "ConRac" Then MsgBox "Año origen mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Dim nroser As Long
    codser = "": nroser = 0
    codreg = "": codser = ""
    For i = 1 To vaSpread1(0).MaxRows
        vaSpread1(0).Row = i
        vaSpread1(0).Col = 1
        If vaSpread1(0).text = "1" Then vaSpread1(0).Col = 2: codreg = codreg & "" & vaSpread1(0).text & ","
    Next i
    For i = 1 To vaSpread1(1).MaxRows
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1
        If vaSpread1(1).text = "1" Then vaSpread1(1).Col = 2: codser = codser & "" & vaSpread1(1).text & ",": nroser = nroser + 1
    Next i
    If Trim(codreg) = "" Then fg_descarga: MsgBox "Regimen debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Trim(codser) = "" Then fg_descarga: MsgBox "Servicio debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If lc_Aux = "CosTot" Then
       I_CostosTotPeriodo fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), IIf(Option2(0).Value = True, 0, IIf(Option2(1).Value = True, 1, 2)), nroser
    ElseIf lc_Aux = "FooCos" Then
       I_FoodCost fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), IIf(Option2(0).Value = True, 0, IIf(Option2(1).Value = True, 1, 2))
    ElseIf lc_Aux = "CosSec" Then
       I_CostoxSector fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), IIf(Option2(0).Value = True, 0, IIf(Option2(1).Value = True, 1, 2))
    ElseIf lc_Aux = "InNPla" Then
       I_InsumoNoPlanifSalBod fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd"))
    ElseIf lc_Aux = "CosPer" Then
       I_CostoDetPeriodoRealizado fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), IIf(Option2(0).Value = True, "1", "2")
    ElseIf lc_Aux = "CurABC" Then
       I_CurvaABC fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), IIf(Option2(0).Value = True, "1", IIf(Option2(1).Value = True, "2", "0"))
    ElseIf lc_Aux = "CocABC" Then
       I_ComparativoCurvaABC fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), IIf(Option2(0).Value = True, "1", IIf(Option2(1).Value = True, "2", "0"))
    ElseIf lc_Aux = "ConRac" Then
       vg_opgra = 2 '-------> Control de raciones
       I_ComparativodeRaciones fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd"))
    Else
       I_MermaPreparacion fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), IIf(Option2(0).Value = True, 1, 2)
    End If
Case 3 '-------> Historico Planificación
    RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda(0).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    vg_codigo = ""
    B_HistPm.LlenarHistPlan "Histórico Planificación", fpText.text, 1, 7
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    fpDateTime1(0).text = "01/" & vg_auxfecha: fpDateTime1(1).text = dEoM("01/" & vg_auxfecha)
    MoverDatoGrilla
    Option1(0).SetFocus
    Me.Refresh
Case 5
    Me.Hide
    Unload Me
End Select
End Sub

Sub MoverDatoGrilla()
'-------> Mover regimen
With vaSpread1(0)
    .MaxRows = 0
    RS.Open RutinaLectura.Regimen(1, 0, ""), vg_db, adOpenStatic
    If Not RS.EOF Then
       Do While Not RS.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1: .text = "1"
          .Col = 2: .text = RS!reg_codigo
          .Col = 3: .text = Trim(RS!reg_nombre)
          RS.MoveNext
       Loop
    End If
    RS.Close: Set RS = Nothing
End With
'-------> Mover servicio
With vaSpread1(1)
    .MaxRows = 0
    RS.Open RutinaLectura.Servicio(1, 0, ""), vg_db, adOpenStatic
    If Not RS.EOF Then
       Do While Not RS.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1: .text = "1"
          .Col = 2: .text = RS!ser_codigo
          .Col = 3: .text = Trim(RS!ser_nombre)
          RS.MoveNext
       Loop
    End If
    RS.Close: Set RS = Nothing
End With
End Sub
