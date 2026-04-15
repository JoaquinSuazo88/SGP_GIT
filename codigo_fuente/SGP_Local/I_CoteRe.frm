VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_CoteRe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costo Plan. Teórico - Plan. Real - Realizado"
   ClientHeight    =   4230
   ClientLeft      =   3075
   ClientTop       =   2820
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3705
      Index           =   0
      Left            =   30
      TabIndex        =   6
      Top             =   390
      Width           =   7875
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
         Index           =   1
         Left            =   4040
         TabIndex        =   22
         Top             =   2820
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
            Index           =   5
            Left            =   2280
            TabIndex        =   24
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
            Index           =   4
            Left            =   240
            TabIndex        =   23
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   2
            Left            =   3000
            Picture         =   "I_CoteRe.frx":0000
            Top             =   160
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   90
         TabIndex        =   13
         Top             =   960
         Width           =   7700
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Caption         =   "Solamente Costo Totales"
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
            Height          =   270
            Index           =   0
            Left            =   4920
            TabIndex        =   19
            Top             =   1440
            Width           =   2610
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
            Left            =   3000
            TabIndex        =   4
            Top             =   960
            Width           =   1695
         End
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
            Left            =   6720
            TabIndex        =   5
            Top             =   960
            Width           =   855
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
            Top             =   550
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
            Top             =   550
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
            TabIndex        =   17
            Top             =   210
            Width           =   4335
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
            Left            =   4040
            TabIndex        =   16
            Top             =   620
            Width           =   1005
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
            TabIndex        =   15
            Top             =   620
            Width           =   1110
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   2620
            Picture         =   "I_CoteRe.frx":030A
            Top             =   120
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
            Index           =   7
            Left            =   120
            TabIndex        =   14
            Top             =   285
            Width           =   735
         End
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3165
            TabIndex        =   18
            Top             =   255
            Width           =   4335
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   90
         TabIndex        =   10
         Top             =   150
         Width           =   7700
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "I_CoteRe.frx":0614
            Left            =   2040
            List            =   "I_CoteRe.frx":0616
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Informes"
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
            Left            =   960
            TabIndex        =   12
            Top             =   300
            Width           =   735
         End
      End
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
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   2820
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
            Index           =   3
            Left            =   240
            TabIndex        =   21
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
            Index           =   2
            Left            =   2280
            TabIndex        =   20
            Top             =   300
            Width           =   735
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   1
            Left            =   3000
            Picture         =   "I_CoteRe.frx":0618
            Top             =   160
            Width           =   480
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   135
         Index           =   0
         Left            =   6660
         TabIndex        =   8
         Top             =   3600
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
         SpreadDesigner  =   "I_CoteRe.frx":0922
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7965
      _ExtentX        =   14049
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
      TabIndex        =   25
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
      SpreadDesigner  =   "I_CoteRe.frx":0FD6
   End
End
Attribute VB_Name = "I_CoteRe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim i As Integer, isel As Integer, tipmin As String
Dim MsgTitulo As String, opcion As String, est As Boolean
Public lc_Aux As String

Private Sub Combo1_Click(Index As Integer)
Select Case Val(fg_codigocbo(Combo1, 0, 1, ""))
Case 0, 3
    codreg = ""
    tipmin = "'1'"
    Check1(0).Enabled = True
    Me.Caption = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 0, "Plan. Teórico & Realizado", "Plan. Teórico & Realziado Acumulado")
Case 1, 4
    codreg = ""
    tipmin = "'2'"
    Check1(0).Enabled = True
    Me.Caption = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 1, "Plan. Real & Realizado", "Plan. Real & Realziado Acumulado")
Case 2, 5
    codreg = ""
    tipmin = "'1','2'"
    Check1(0).Enabled = True
    Me.Caption = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 2, "Plan. Reórico & Plan. Real & Realizado", "Plan. Teórico & Plan. Real & Realziado Acumulado")
Case 6
    codreg = ""
    tipmin = "'1'"
    Check1(0).Enabled = False
    Me.Caption = "Comparativo Plan. Teórico & Negociado"
End Select
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 4710
Me.Width = 8085
Me.HelpContextID = vg_OpcM
fg_centra Me
est = True
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
'If opcion = "G" Then Set btnX = Toolbar1.Buttons.Add(, "A_Grafico", , tbrDefault, "A_Grafico"): btnX.Visible = True: btnX.ToolTipText = "Gráfico": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): BtnX.Visible = False: BtnX.Enabled = False: BtnX.ToolTipText = "Enviar SGP Inf.": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico Planificacón Teórica"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
With Combo1(0)
    .Clear
    If lc_Aux = "CoTeRe" Then
       Me.Caption = "Costo Plan. Teórico - Plan. Real - Realizado"
       .AddItem "Plan. Teórico & Realizado" & Space(150) & "(0)"
       .AddItem "Plan. Real & Realizado" & Space(150) & "(1)"
       .AddItem "Plan. Teórico & Plan. Real & Realizado" & Space(150) & "(2)"
       .AddItem "Plan. Teórico & Realizado Acumulado" & Space(150) & "(3)"
       .AddItem "Plan. Real & Realizado Acumulado" & Space(150) & "(4)"
       .AddItem "Plan. Teórico & Plan. Real & Realizado Acumulado" & Space(150) & "(5)"
    Else
       Me.Caption = "Comparativo Costo Plan. Teórico & Negociado"
       .AddItem "Comparativo Plan. Teórico & Negociado" & Space(150) & "(6)"
    End If
    .ListIndex = 0
End With
codreg = ""
est = False
tipmin = "'1'"
MoverDatoGrilla
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_Change()
RS.Open RutinaLectura.Cliente(1, Trim(LimpiaDato(fpText.text)), ""), vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
codreg = ""
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
Dim i As Long, auxreg As String

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

Case 1
    
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpText.text = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Regimen", Me.vaSpread1, fpText.text, "", Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "0", "", 0, IIf(lc_Aux = "PlaTei", "'1'", "'2'")
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    vg_codigo = ""

Case 2
    
    If fpText.text = "" Then Exit Sub
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, fpText.text, "", Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "0", "", 1, IIf(lc_Aux = "PlaTei", "'1'", "'2'")
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub

End Select

End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    Image1(2).Enabled = False
Case 1
    Image1(2).Enabled = True
Case 2
    Image1(1).Enabled = True
Case 3
    Image1(1).Enabled = False
    MoverDatoGrilla
Case 4
    Image1(2).Enabled = False
    MoverDatoGrilla
Case 5
    Image1(2).Enabled = True
'    MoverDatoGrilla
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim opnomrec As Boolean
Dim codser As String, codreg As String, opcosto As Integer, tipmin As String, numreg As Long, numser As Long
Select Case Button.Index
Case 1 '-------> Generar Informe
    If Combo1(0).ListIndex = -1 Or Combo1(0).text = "" Then Exit Sub
    RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda(0).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    If fpDateTime1(0).Value > fpDateTime1(1).Value Then MsgBox "Fecha origen Mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Mid(fpDateTime1(0).text, 4, 2) <> Mid(fpDateTime1(1).text, 4, 2) Then MsgBox "Mes origen mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Mid(fpDateTime1(0).text, 7, 4) <> Mid(fpDateTime1(1).text, 7, 4) Then MsgBox "Año origen mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    codreg = "": codser = ""
    numreg = 0: numser = 0
    With vaSpread1(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .text = "1" Then .Col = 2: codreg = codreg & "" & .text & ",": numreg = numreg + 1
        Next i
    End With
    With vaSpread1(1)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .text = "1" Then .Col = 2: codser = codser & "" & .text & ",": numser = numser + 1
        Next i
    End With
    If Trim(codreg) = "" Then fg_descarga: MsgBox "Regimen debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Trim(codser) = "" Then fg_descarga: MsgBox "Servicio debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    opcosto = IIf(Option2(0).Value = True, 0, IIf(Option2(1).Value = True, 1, 2))
    Frame1(0).Enabled = False
    Toolbar1.Enabled = False
    Select Case Val(fg_codigocbo(Combo1, 0, 1, ""))
    Case 0, 1, 2
        vg_opgra = 1 '------> gráfico food cost
        tipmin = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 2, "'1','2'", IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 0, "'1'", "'2'"))
        vg_codcasino = fpText.text
        vg_codreg = codreg
        vg_codser = codser
        vg_fecini = Val(Format(fpDateTime1(0).text, "yyyymmdd"))
        vg_fecfin = Val(Format(fpDateTime1(1).text, "yyyymmdd"))
        vg_op1 = Combo1(0).ListIndex
        vg_op2 = IIf(Option2(2).Value = True, 2, IIf(Option2(0).Value = True, 0, 1))
        vg_op3 = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 2, "'1','2'", IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 0, "'1'", "'2'"))
        I_CostoTeoricoRealFood fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), tipmin, Val(fg_codigocbo(Combo1, 0, 1, "")), IIf(Check1(0).Value = 1, True, False), opcosto, numreg, numser 'opnomrec
        vg_codcasino = "": vg_codreg = "":       vg_codser = "": vg_fecini = 0: vg_fecfin = 0: vg_op1 = "": vg_op2 = "": vg_op3 = ""
    Case 3, 4, 5
        tipmin = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 5, "'1','2'", IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 3, "'1'", "'2'"))
        I_CostoTeoricoRealFoodAcum fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), tipmin, Val(fg_codigocbo(Combo1, 0, 1, "")), IIf(Check1(0).Value = 1, True, False), opcosto, numreg, numser 'opnomrec
    Case 6
        vg_opgra = 0
        tipmin = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 6, "'1'", "'2'")
        I_CostoTeoricoNegociado fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), tipmin, Val(fg_codigocbo(Combo1, 0, 1, "")), IIf(Check1(0).Value = 1, True, False), opcosto, numreg, numser
    End Select
    Frame1(0).Enabled = True
    Toolbar1.Enabled = True
Case 3 '-------> Generar envío sgp inf
    If Combo1(0).ListIndex = -1 Or Combo1(0).text = "" Then Exit Sub
    RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda(0).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    If fpDateTime1(0).Value > fpDateTime1(1).Value Then MsgBox "Fecha origen Mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Mid(fpDateTime1(0).text, 4, 2) <> Mid(fpDateTime1(1).text, 4, 2) Then MsgBox "Mes origen mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Mid(fpDateTime1(0).text, 7, 4) <> Mid(fpDateTime1(1).text, 7, 4) Then MsgBox "Año origen mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    codreg = "": codser = ""
    numreg = 0: numser = 0
    With vaSpread1(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .text = "1" Then .Col = 2: codreg = codreg & "" & .text & ",": numreg = numreg + 1
        Next i
    End With
    With vaSpread1(1)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .text = "1" Then .Col = 2: codser = codser & "" & .text & ",": numser = numser + 1
        Next i
    End With
    If Trim(codreg) = "" Then fg_descarga: MsgBox "Regimen debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Trim(codser) = "" Then fg_descarga: MsgBox "Servicio debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Combo1(0).ListIndex = 2
    Option2(2).Value = True
    tipmin = "'1','2'"
    opcosto = 2
    Frame1(0).Enabled = False
    Toolbar1.Enabled = False
    E_CostoTeoricoRealFood fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), tipmin, Val(fg_codigocbo(Combo1, 0, 1, "")), IIf(Check1(0).Value = 1, True, False), opcosto, numreg, numser 'opnomrec
    Frame1(0).Enabled = True
    Toolbar1.Enabled = True
Case 5 '-------> Historico Planificación
    RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda(0).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    vg_codigo = ""
    B_HistPm.LlenarHistPlan "Histórico Planificación", fpText.text, 1, 2
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    fpDateTime1(0).text = "01/" & vg_auxfecha: fpDateTime1(1).text = dEoM("01/" & vg_auxfecha)
'    MoverDatoGrilla
    Me.Refresh
Case 7 '-------> Salir
    Me.Hide
    Unload Me
End Select
End Sub

Sub MoverDatoGrilla()
fg_carga ""
With vaSpread1(0)
.MaxRows = 0
    RS.Open RutinaLectura.Regimen(1, 0, ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: Exit Sub
    Do While Not RS.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1: .text = "1"
       .Col = 2: .text = RS!reg_codigo
       .Col = 3: .text = Trim(RS!reg_nombre)
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
End With
With vaSpread1(1)
    .MaxRows = 0
    RS.Open RutinaLectura.Servicio(1, 0, ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: Exit Sub
    Do While Not RS.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1: .text = "1"
       .Col = 2: .text = RS!ser_codigo
       .Col = 3: .text = Trim(RS!ser_nombre)
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
End With
fg_descarga
End Sub

Sub Inicio(tfor As String, op As String)
MsgTitulo = tfor
opcion = op
Me.Caption = tfor
End Sub

Function E_CostoTeoricoRealFood(cencos As String, codreg As String, codser As String, fecini As Long, fecfin As Long, tipmin As String, opinf As Integer, opinc As Boolean, opcosto As Integer, numreg As Long, numser As Long)
Dim i As Long, j As Long, NotA As Long, AsiA As Long, CumA As Long, CerA As Long, auxfec As Long, auxfecrac As Long, ancpag As Long, fecesf As Long
Dim nomser As String, cLin As String, nomtit As String, sql1 As String, estfij As Boolean, auxreg As Long, auxser As Long, sql2 As String, sql3 As String
Dim costeo As Double, cosrea As Double, racteo As Double, racrea As Double, totteo As Double, totdoc As Double
Dim tcoteo As Double, tcorea As Double, trateo As Double, trarea As Double, totrea As Double
Dim cosfod As Double, tcofod As Double, trafod As Double, tdiateo As Double, tdiarea As Double, vCosFij As Double, vFcUnit As Double, vPtUnit As Double, vPrUnit As Double
Dim tgracpt As Double, tgracpr As Double, tgracre As Double, tgrcospt As Double, tgrcospr As Double, tgrcosre As Double
Dim vec_cosmin() As Variant, cosali As Double, CosDes As Double, aAp As String
On Local Error GoTo Error_CostTeoReaFood
fg_carga ""
ancpag = 13500
If vg_tipbase = "1" Then
   '-------> Insert tabla productospmpdia
   aAp = Trim(vg_NUsr) & "_tmp_ProductoCTRFePMP"
   fg_CheckTmp aAp
   vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, 0 AS ppd_upreco, null AS ppd_fecuco, Max(ppd_fecdia) AS ppd_fecdia " & _
                 "INTO " & aAp & " " & _
                 "FROM b_productospmpdia " & _
                 "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                 "AND   ppd_propon > 0 " & _
                 "GROUP BY ppd_cencos, ppd_codpro"
   vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
   vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon, " & aAp & ".ppd_upreco=b_productospmpdia.ppd_upreco, " & aAp & ".ppd_fecuco=b_productospmpdia.ppd_fecuco"
   vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_upreco, ppd_fecuco, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_upreco, ppd_fecuco, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
End If
sql1 = "(c.pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') OR c.pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "'))"
numreg = (numser * 31)
ReDim vec_cosmin(numser * 31, 15)
For i = 1 To UBound(vec_cosmin)
    vec_cosmin(i, 1) = 0 '-------> Fecha ddmmyyyy
    vec_cosmin(i, 2) = 0 '-------> Código regimen
    vec_cosmin(i, 3) = "" '-------> Nombre regimen
    vec_cosmin(i, 4) = 0 '-------> Código servicio
    vec_cosmin(i, 5) = "" '-------> Nombre servicio
    vec_cosmin(i, 6) = 0 '-------> Raciones teorica
    vec_cosmin(i, 7) = 0 '-------> Monto teorica alimentación
    vec_cosmin(i, 8) = 0 '-------> Monto teorica desechable
    vec_cosmin(i, 9) = 0 '-------> raciones real
    vec_cosmin(i, 10) = 0 '-------> Monto real alimentación
    vec_cosmin(i, 11) = 0 '-------> Monto real desechable
    vec_cosmin(i, 12) = 0 '-------> raciones realizada
    vec_cosmin(i, 13) = 0 '-------> Monto realizada alimentación
    vec_cosmin(i, 14) = 0 '-------> Monto realizada desechable
Next i
'-------> Mover costo planificación teorica vs real
RS1.Open "SELECT b.min_codreg, b.min_codser, c.mid_tipmin, b.min_fecmin, b.min_indblo, b.min_racteo, b.min_racrea, " & _
         "ROUND(SUM(c.mid_cosrec*c.mid_numrac),2) AS mid_cosrec, ROUND(SUM(c.mid_cosdes*c.mid_numrac),2) AS mid_cosdes FROM b_minuta b, b_minutadet c " & _
         "WHERE b.min_codigo = c.mid_codigo " & _
         "AND   b.min_cencos = '" & cencos & "' " & _
         "AND   b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
         "AND   b.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
         "AND   b.min_fecmin >= " & fecini & " " & _
         "AND   b.min_fecmin <= " & fecfin & " " & _
         "AND   c.mid_tipmin IN (" & Mid(tipmin, 1, Len(tipmin)) & ") " & _
         "GROUP BY b.min_codreg, b.min_codser, c.mid_tipmin, b.min_fecmin, b.min_indblo, b.min_racteo, b.min_racrea " & _
         "ORDER BY b.min_codreg, b.min_codser, b.min_fecmin, c.mid_tipmin", vg_db, adOpenStatic
If RS1.EOF Then Close #1: RS1.Close: Set RS1 = Nothing: Exit Function
auxreg = 0: auxser = 0: estfij = False
Do While Not RS1.EOF
   DoEvents
   If RS1!min_codreg <> auxreg Or RS1!min_codser <> auxser Then
      '-------> Traer estructura fija
      estfij = False
      RS2.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
               "WHERE mfd_cencos = '" & cencos & "' " & _
               "AND   mfd_codreg = " & RS1!min_codreg & " " & _
               "AND   mfd_codser = " & RS1!min_codser & " " & _
               "AND   mfd_fecha >= " & fecini & " AND mfd_fecha <= " & fecfin & " AND mfd_tipmin IN (" & Mid(tipmin, 1, Len(tipmin)) & ")", vg_db, adOpenStatic
      If Not RS2.EOF Then estfij = True
      RS2.Close: Set RS2 = Nothing

      '-------> Buscar fecha estructura fija
      fecesf = 0
      If Not estfij Then
         RS2.Open "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija " & _
                  "WHERE mif_cencos = '" & cencos & "' " & _
                  "AND   mif_codreg = " & RS1!min_codreg & " " & _
                  "AND   mif_codser = " & RS1!min_codser & "", vg_db, adOpenStatic
         If Not RS2.EOF Then fecesf = IIf(IsNull(RS2!fecval), 0, RS2!fecval)
         RS2.Close: Set RS2 = Nothing
      End If
      auxreg = RS1!min_codreg
      auxser = RS1!min_codser
   End If

   vCosFij = 0: cosali = 0: CosDes = 0
   If estfij Then
      '-------> Calcular datos desde tabla estructura fija día
      RS3.Open "SELECT c.pro_ctacon, SUM(a.mfd_canpro*a.mfd_cospro) AS cosfij " & _
               "FROM b_minutafijadia a, b_productos c " & _
               "WHERE a.mfd_codpro = c.pro_codigo " & _
               "AND   a.mfd_cencos = '" & cencos & "' " & _
               "AND   a.mfd_codreg = " & RS1!min_codreg & " " & _
               "AND   a.mfd_codser = " & RS1!min_codser & " " & _
               "AND   a.mfd_fecha = " & RS1!min_fecmin & " AND a.mfd_tipmin = '" & RS1!mid_tipmin & "' " & _
               "AND   " & sql1 & " GROUP BY c.pro_ctacon", vg_db, adOpenStatic
      If Not RS3.EOF Then
         Do While Not RS3.EOF
            If RS3!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
               cosali = cosali + RS3!cosfij
            ElseIf RS3!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
               CosDes = CosDes + RS3!cosfij
            End If
            RS3.MoveNext
         Loop
      End If
'      And Not IsNull(RS3!cosfij) Then vCosFij = RS3!cosfij
      RS3.Close: Set RS3 = Nothing
   ElseIf Not estfij And fecesf > 0 Then
      '-------> Calcular datos desde tabla estructura fija
      If vg_tipbase = "1" Then
         RS3.Open "SELECT c.pro_ctacon, SUM(b.ppd_propon*a.mif_canpro) AS cosfij " & _
                  "FROM  b_minutafija a, " & aAp & " b, b_productos c " & _
                  "WHERE a.mif_codpro = c.pro_codigo AND c.pro_codigo = b.ppd_codpro AND b.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                  "AND   a.mif_cencos = '" & cencos & "' " & _
                  "AND   a.mif_codreg = " & RS1!min_codreg & " " & _
                  "AND   a.mif_codser = " & RS1!min_codser & " " & _
                  "AND   a.mif_fecval = " & fecesf & " " & _
                  "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(RS1!min_fecmin & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(RS1!min_fecmin & Right("0" & i, 2), 2)) - 2))) & " " & _
                  "AND   " & sql1 & " GROUP BY c.pro_ctacon", vg_db, adOpenStatic
      Else
         RS3.Open "SELECT c.pro_ctacon, SUM(b.ppd_propon*a.mif_canpro) AS cosfij " & _
                  "FROM  b_minutafija a, b_productospmpdia b, b_productos c " & _
                  "WHERE a.mif_codpro = c.pro_codigo AND c.pro_codigo = b.ppd_codpro AND b.ppd_cencos = '" & MuestraCasino(1) & "' AND b.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
                  "AND   a.mif_cencos = '" & cencos & "' " & _
                  "AND   a.mif_codreg = " & RS1!min_codreg & " " & _
                  "AND   a.mif_codser = " & RS1!min_codser & " " & _
                  "AND   a.mif_fecval = " & fecesf & " " & _
                  "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(RS1!min_fecmin & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(RS1!min_fecmin & Right("0" & i, 2), 2)) - 2))) & " " & _
                  "AND   " & sql1 & " GROUP BY c.pro_ctacon", vg_db, adOpenStatic
      End If
      
      If Not RS3.EOF Then
         Do While Not RS3.EOF
            If RS3!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
               cosali = cosali + RS3!cosfij
            ElseIf RS3!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
               CosDes = CosDes + RS3!cosfij
            End If
            RS3.MoveNext
         Loop
      End If
'      If Not RS3.EOF And Not IsNull(RS3!cosfij) Then vCosFij = RS3!cosfij
      RS3.Close: Set RS3 = Nothing
   End If

   For i = 1 To UBound(vec_cosmin)
       If vec_cosmin(i, 1) = RS1!min_fecmin And vec_cosmin(i, 2) = RS1!min_codreg And vec_cosmin(i, 4) = RS1!min_codser Then
          vec_cosmin(i, IIf(RS1!mid_tipmin = "1", 6, 9)) = vec_cosmin(i, IIf(RS1!mid_tipmin = "1", 6, 9)) + IIf(RS1!mid_tipmin = "1", IIf(IsNull(RS1!min_racteo), 0, RS1!min_racteo), IIf(IsNull(RS1!min_racrea), 0, RS1!min_racrea))
          vec_cosmin(i, IIf(RS1!mid_tipmin = "1", 7, 10)) = vec_cosmin(i, IIf(RS1!mid_tipmin = "1", 7, 10)) + IIf(IsNull(RS1!mid_cosrec), 0, RS1!mid_cosrec) + cosali
          vec_cosmin(i, IIf(RS1!mid_tipmin = "1", 8, 11)) = vec_cosmin(i, IIf(RS1!mid_tipmin = "1", 8, 11)) + IIf(IsNull(RS1!mid_cosdes), 0, RS1!mid_cosdes) + CosDes
          Exit For
       ElseIf vec_cosmin(i, 1) = 0 And vec_cosmin(i, 2) = 0 And vec_cosmin(i, 4) = 0 Then
          vec_cosmin(i, 1) = RS1!min_fecmin
          vec_cosmin(i, 2) = RS1!min_codreg
          vec_cosmin(i, 4) = RS1!min_codser
          vec_cosmin(i, IIf(RS1!mid_tipmin = "1", 6, 9)) = IIf(RS1!mid_tipmin = "1", IIf(IsNull(RS1!min_racteo), 0, RS1!min_racteo), IIf(IsNull(RS1!min_racrea), 0, RS1!min_racrea))
          vec_cosmin(i, IIf(RS1!mid_tipmin = "1", 7, 10)) = IIf(IsNull(RS1!mid_cosrec), 0, RS1!mid_cosrec) + cosali
          vec_cosmin(i, IIf(RS1!mid_tipmin = "1", 8, 11)) = IIf(IsNull(RS1!mid_cosdes), 0, RS1!mid_cosdes) + CosDes
          Exit For
       End If
   Next i
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
'-------> Traer costo salida & devolución
totdoc = 0: vFcUnit = 0
sql2 = IIf(vg_opbase = "1", " cdate('" & fg_Ctod1(fecini) & "') ", " '" & Format(fg_Ctod1(fecini), "yyyymmdd") & "' ")
sql2 = IIf(vg_opbase = "1", " cdate('" & fg_Ctod1(fecfin) & "') ", " '" & Format(fg_Ctod1(fecfin), "yyyymmdd") & "' ")
RS1.Open "SELECT a.tov_fecpro, a.tov_codreg, a.tov_codser, c.pro_ctacon, " & _
         "SUM(IIf(a.tov_tipdoc='SP',b.dev_ptotal,'-' & b.dev_ptotal)) AS dev_ptotal " & _
         "FROM  b_totventas a, b_detventas b, b_productos c " & _
         "WHERE a.tov_rutcli = b.dev_rutcli " & _
         "AND   a.tov_tipdoc = b.dev_tipdoc " & _
         "AND   a.tov_numdoc = b.dev_numdoc " & _
         "AND   b.dev_codmer = c.pro_codigo " & _
         "AND   " & sql1 & " " & _
         "AND   a.tov_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
         "AND   a.tov_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
         "AND  (a.tov_tipdoc = 'SP' OR a.tov_tipdoc = 'DP') " & _
         "AND   b.dev_canmer <> 0 " & _
         "AND   a.tov_estdoc <> 'A' AND a.tov_estdoc<>'P' AND a.tov_codbod=" & vg_codbod & " " & _
         "AND   a.tov_fecpro >= " & sql2 & " " & _
         "AND   a.tov_fecpro <= " & sql3 & " " & _
         "GROUP BY a.tov_fecpro, a.tov_codreg, a.tov_codser, c.pro_ctacon ORDER BY a.tov_codreg, a.tov_codser, a.tov_fecpro", vg_db, adOpenStatic
Do While Not RS1.EOF
   DoEvents
   For i = 1 To UBound(vec_cosmin)
       If vec_cosmin(i, 1) = Val(Format(RS1!tov_fecpro, "yyyymmdd")) And vec_cosmin(i, 2) = RS1!tov_codreg And vec_cosmin(i, 4) = RS1!tov_codser Then
          vec_cosmin(i, IIf(RS1!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")), 13, 14)) = vec_cosmin(i, IIf(RS1!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")), 13, 14)) + RS1!dev_ptotal
          Exit For
       ElseIf vec_cosmin(i, 1) = 0 And vec_cosmin(i, 2) = 0 And vec_cosmin(i, 4) = 0 Then
          vec_cosmin(i, 1) = Format(RS1!tov_fecpro, "yyyymmdd")
          vec_cosmin(i, 2) = RS1!tov_codreg
          vec_cosmin(i, 4) = RS1!tov_codser
          vec_cosmin(i, IIf(RS1!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")), 13, 14)) = RS1!dev_ptotal
          Exit For
       End If
   Next i
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
'-------> Mover raciones
If opinc = False Then
   RS1.Open "SELECT mir_fecmin, mir_codreg, mir_codser, SUM(mir_nrorac) AS mir_nrorac FROM b_minutaraciones " & _
            "WHERE mir_cencos='" & cencos & "' " & _
            "AND   mir_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
            "AND   mir_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
            "AND  (mir_rutcli = 'PRODUCIDAS') " & _
            "AND   mir_fecmin>=" & fecini & " AND mir_fecmin<=" & fecfin & " GROUP BY mir_fecmin, mir_codreg, mir_codser", vg_db, adOpenStatic
   Do While Not RS1.EOF
      DoEvents
      For i = 1 To UBound(vec_cosmin)
       If vec_cosmin(i, 1) = RS1!mir_fecmin And vec_cosmin(i, 2) = RS1!mir_codreg And vec_cosmin(i, 4) = RS1!mir_codser Then
          vec_cosmin(i, 12) = vec_cosmin(i, 12) + IIf(IsNull(RS1!mir_nrorac), 0, RS1!mir_nrorac)
          Exit For
       ElseIf vec_cosmin(i, 1) = 0 And vec_cosmin(i, 2) = 0 And vec_cosmin(i, 4) = 0 Then
          vec_cosmin(i, 1) = RS1!mir_fecmin
          vec_cosmin(i, 2) = RS1!mir_codreg
          vec_cosmin(i, 4) = RS1!mir_codser
          vec_cosmin(i, 12) = IIf(IsNull(RS1!mir_nrorac), 0, RS1!mir_nrorac)
          Exit For
       End If
      Next i
      RS1.MoveNext
   Loop
   RS1.Close: Set RS1 = Nothing
End If
'-------> Proceso de grabado
If vec_cosmin(1, 1) < 1 Then Exit Function
Dim cospis As Double, costec As Double, nomreg As String
auxreg = 0: auxser = 0
vg_db.Execute "DELETE p_costrr FROM p_costrr WHERE trr_cencos = '" & cencos & "' AND trr_usuario = '" & vg_NUsr & "'"
For i = 1 To UBound(vec_cosmin)
    DoEvents
    If vec_cosmin(i, 1) > 0 Then
        If vec_cosmin(i, 2) <> auxreg Or vec_cosmin(i, 4) <> auxser Then
           cospis = 0: costec = 0
           RS1.Open "SELECT * FROM b_costopatron WHERE cpa_cencos = '" & cencos & "' AND cpa_codreg = " & vec_cosmin(i, 2) & " AND cpa_codser = " & vec_cosmin(i, 4) & " AND cpa_anomes = " & Mid(vec_cosmin(i, 1), 1, 6) & "", vg_db, adOpenStatic
           If Not RS1.EOF Then
              Do While Not RS1.EOF
                 If RS1!cpa_descripcion = "PISO" Then
                    cospis = RS1!cpa_valor
                 ElseIf RS1!cpa_descripcion = "TECHO" Then
                    costec = RS1!cpa_valor
                 End If
                 RS1.MoveNext
              Loop
           End If
           RS1.Close: Set RS1 = Nothing
           nomreg = ""
           RS1.Open "SELECT * FROM a_regimen WHERE reg_codigo = " & vec_cosmin(i, 2) & "", vg_db, adOpenStatic
           If Not RS1.EOF Then nomreg = Trim(RS1!reg_nombre)
           RS1.Close: Set RS1 = Nothing
           nomser = ""
           RS1.Open "SELECT * FROM a_servicio WHERE ser_codigo = " & vec_cosmin(i, 4) & "", vg_db, adOpenStatic
           If Not RS1.EOF Then nomser = Trim(RS1!ser_nombre)
           RS1.Close: Set RS1 = Nothing
           auxreg = vec_cosmin(i, 2)
           auxser = vec_cosmin(i, 4)
        End If
        vg_db.Execute "INSERT INTO p_costrr VALUES ('" & vg_NUsr & "', '" & MuestraCasino(1) & "', " & vec_cosmin(i, 2) & ", " & vec_cosmin(i, 4) & ", '" & fg_Ctod1(vec_cosmin(i, 1)) & "', '" & MuestraCasino(2) & "', '" & nomreg & "', '" & nomser & "', " & cospis & ", " & costec & ", " & vec_cosmin(i, 6) & ", " & vec_cosmin(i, 7) & ", " & vec_cosmin(i, 8) & ", " & vec_cosmin(i, 9) & ", " & vec_cosmin(i, 10) & ", " & vec_cosmin(i, 11) & ", " & vec_cosmin(i, 12) & ", " & vec_cosmin(i, 13) & ", " & vec_cosmin(i, 14) & ")"
    End If
Next i
fg_descarga
MsgBox "Generación envió Finalizado Sin Problema", vbExclamation + vbOKOnly, MsgTitulo
Exit Function
Error_CostTeoReaFood:
    fg_descarga
    Resume Next
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Exit Function
End Function
