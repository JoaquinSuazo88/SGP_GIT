VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_ApoNutSansis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aporte Nutricional Sansis"
   ClientHeight    =   5265
   ClientLeft      =   4980
   ClientTop       =   2490
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4575
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   8535
      Begin FPSpread.vaSpread vaSpread3 
         Height          =   135
         Left            =   1200
         TabIndex        =   34
         Top             =   3720
         Visible         =   0   'False
         Width           =   1455
         _Version        =   393216
         _ExtentX        =   2566
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
         SpreadDesigner  =   "I_ApoNutSansis.frx":0000
      End
      Begin VB.CheckBox IncluyeGrsCero 
         Caption         =   "Incluye Grs Cero"
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
         Left            =   6720
         TabIndex        =   33
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox SaltoPagina 
         Caption         =   "Salto Página"
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
         Left            =   4560
         TabIndex        =   32
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Frame Frame4 
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
         TabIndex        =   31
         Top             =   1440
         Width           =   3855
         Begin VB.OptionButton Option4 
            Caption         =   "Detallado"
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
            Left            =   2400
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Resumido "
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
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pavb"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4560
         TabIndex        =   30
         Top             =   2820
         Width           =   3855
         Begin VB.OptionButton Option3 
            Caption         =   "Con Pavb"
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
            Left            =   2220
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Sin Pavb"
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
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   255
         Left            =   4440
         TabIndex        =   24
         Top             =   960
         Visible         =   0   'False
         Width           =   375
         _Version        =   393216
         _ExtentX        =   661
         _ExtentY        =   450
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
         SpreadDesigner  =   "I_ApoNutSansis.frx":01FE
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Aporte Nutricional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   2820
         Width           =   3810
         Begin VB.OptionButton Option2 
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
            Left            =   240
            TabIndex        =   8
            Top             =   340
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option2 
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
            Left            =   2400
            TabIndex        =   9
            Top             =   340
            Width           =   735
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   3120
            Picture         =   "I_ApoNutSansis.frx":4664
            Top             =   210
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   675
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   2100
         Width           =   3810
         Begin VB.OptionButton Option2 
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
            Left            =   240
            TabIndex        =   6
            Top             =   340
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option2 
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
            Left            =   2400
            TabIndex        =   7
            Top             =   340
            Width           =   735
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   3180
            Picture         =   "I_ApoNutSansis.frx":496E
            Top             =   210
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Opción Casino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4560
         TabIndex        =   18
         Top             =   2100
         Width           =   3885
         Begin VB.OptionButton Option1 
            Caption         =   "Sin Código"
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
            Left            =   240
            TabIndex        =   10
            Top             =   340
            Value           =   -1  'True
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Con Código"
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
            Left            =   2220
            TabIndex        =   11
            Top             =   340
            Width           =   1365
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   375
         Left            =   6240
         TabIndex        =   14
         Top             =   3600
         Visible         =   0   'False
         Width           =   1920
         _Version        =   393216
         _ExtentX        =   3387
         _ExtentY        =   661
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
         SpreadDesigner  =   "I_ApoNutSansis.frx":4C78
         StartingColNumber=   6
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   915
         Width           =   1290
         _Version        =   196608
         _ExtentX        =   2284
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
         ThreeDTextHighlightColor=   -2147483633
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
         InvalidColor    =   -2147483643
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
         Text            =   "08/2025"
         DateCalcMethod  =   1
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
         ThreeDFrameColor=   -2147483633
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime FpFecHasta 
         Height          =   315
         Left            =   7125
         TabIndex        =   3
         Top             =   915
         Width           =   1290
         _Version        =   196608
         _ExtentX        =   2284
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
         ThreeDTextHighlightColor=   -2147483633
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
         InvalidColor    =   -2147483643
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
         Text            =   "08/2025"
         DateCalcMethod  =   1
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
         ThreeDFrameColor=   -2147483633
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   240
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
      Begin EditLib.fpLongInteger Regimen 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   575
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
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
         NullColor       =   16777215
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3510
         TabIndex        =   28
         Top             =   585
         Width           =   4845
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3060
         Picture         =   "I_ApoNutSansis.frx":50DC
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3060
         Picture         =   "I_ApoNutSansis.frx":53E6
         Top             =   150
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   5415
         TabIndex        =   20
         Top             =   975
         Width           =   1605
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   975
         Width           =   1605
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
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   630
         Width           =   1215
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
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   16
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3510
         TabIndex        =   26
         Top             =   240
         Width           =   4845
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3555
         TabIndex        =   27
         Top             =   285
         Width           =   4845
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3555
         TabIndex        =   29
         Top             =   615
         Width           =   4845
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_ApoNutSansis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgTitulo As String
Public lc_Aux As String

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Me.Height = 5700
Me.Width = 8925

Me.HelpContextID = vg_OpcM
fg_centra Me
fg_carga ""

MsgTitulo = "Aporte Nutricional Sansis"

Toolbar1.ImageList = Partida.IL1
'Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exportar Excel": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico Planificacón Teórica"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0

FpFecDesde.text = Format(Date, "mm/yyyy")
FpFecHasta.text = Format(Date, "mm/yyyy")

'------- Llenar Tabla Diéteticas
Set RS = vg_db.Execute("sgpadm_s_nutriente 1, 0, ''")

If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe maestro nutrientes", vbExclamation + vbOKOnly, MsgTitulo
   Me.Hide
   Unload Me

End If

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      
      If RS!nut_indpri > 0 Then
         
         vaSpread2.Col = 1
         vaSpread2.CellType = 10
         vaSpread2.TypeCheckText = ""
         vaSpread2.TypeCheckCenter = True
         vaSpread2.text = "1" ' checked
      
      Else
         
         vaSpread2.Col = 1
         vaSpread2.CellType = 10
         vaSpread2.TypeCheckText = ""
         vaSpread2.TypeCheckCenter = True
         vaSpread2.text = " " ' checked
      
      End If
      
      vaSpread2.Col = 2: vaSpread2.Value = RS(0)
      vaSpread2.Col = 3: vaSpread2.Value = Trim(RS(1))
      
      RS.MoveNext
   
   Loop

End If
RS.Close
Set RS = Nothing
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub
MoverDatosVector

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
MoverDatosVector

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

Private Sub fpText1_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Set RS = vg_db.Execute("sgpadm_s_cliente_V02 45, '" & LimpiaDato(fpText1.text) & "', ''")
If RS.EOF Then

   RS.Close
   Set RS = Nothing
   fpayuda(0).Caption = ""
   Regimen.text = ""
   fpayuda(1).Caption = ""
   FpFecDesde.Enabled = True
   FpFecHasta.Enabled = True
   Exit Sub

End If

fpayuda(0).Caption = Trim(RS!Cli_nombre)
fpText1.text = RS!Cli_codigo
RS.Close
Set RS = Nothing
 
FpFecDesde.Enabled = True
FpFecHasta.Enabled = True

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Option2_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0
    Option2(0).Value = True
    Option2(1).Value = False
    Image1(2).Enabled = False

Case 1
    Option2(0).Value = False
    Option2(1).Value = True
    Image1(2).Enabled = True

Case 2
    Option2(2).Value = True
    Option2(3).Value = False
    Image1(3).Enabled = False

Case 3
    Option2(2).Value = False
    Option2(3).Value = True
    Image1(3).Enabled = True

End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0
    
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    Call B_TabEst.LlenaDatos("b_clientes", "cli_", "Clientes", "Cliente_SitioRemoto")
    Call B_TabEst.Show(1)
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText1.text = vg_codigo
    fpayuda(0).Caption = vg_nombre
    Regimen.Value = ""
    Let fpayuda(1).Caption = ""
    Regimen.SetFocus

Case 1
    
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    Call B_TabEst.LlenaDatos("a_regimen", "", "Regimen", "RegBlo")
    Call B_TabEst.Show(1)
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    Regimen.Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    FpFecDesde.SetFocus

Case 2
    
    OpcionLectura = "5"
    vg_nombre = "": vg_codigo = ""
    vg_codigo = Trim(LimpiaDato(fpText1.text))
    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, 0, Val(Regimen.Value), 0, Format(FpFecDesde.text, "yyyymmdd"), Format(FpFecHasta.text, "yyyymmdd"), "5", 0
    B_MTaEst.Show 1
    Me.Refresh
    
    If vg_codigo = "" Then
       
       Exit Sub
    
    End If

Case 3
    
    vg_nombre = "": vg_codigo = ""
    
    If Trim(LimpiaDato(fpText1.text)) = "" Or Regimen.Value = "" Then
    
        Exit Sub
    
    End If
    
    B_MTaEst.LlenaDatos "Nutrientes", Me.vaSpread2, 0, 0, 0, 0, 0, "2", ""
    B_MTaEst.Show 1
    Me.Refresh
    
    If vg_codigo = "" Then
       
       Exit Sub
    
    End If
    
End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Regimen_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
    
If Val(Regimen.Value) < 1 Then
       
   fpayuda(1).Caption = ""
   Exit Sub
    
End If
    
If vg_Indppr = 1 Or vg_Indppr = 2 Then
      
   Set RS = vg_db.Execute("SELECT * FROM a_regimen With(NoLock) WHERE reg_codigo=" & Regimen.Value & " and reg_indppr='" & vg_Indppr & "'")
    
Else
      
   Set RS = vg_db.Execute("SELECT * FROM a_regimen With(NoLock) WHERE reg_codigo=" & Regimen.Value & "")
    
End If
    
If RS.EOF Then
       
   RS.Close
   Set RS = Nothing
   fpayuda(1).Caption = ""
   Exit Sub
    
End If
    
fpayuda(1).Caption = Trim(RS!reg_nombre)
RS.Close
Set RS = Nothing
MoverDatosVector

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Regimen_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim iselecc As Integer
Dim i       As Long
Dim RS      As New ADODB.Recordset

Select Case Button.Index

Case 1

    If Not ValidarDatos Then Exit Sub
    
    fg_carga ""
    Toolbar1.Enabled = False
    Frame1(0).Enabled = False
    vg_opimp = 0

    'I_AporteNutricionalSansis Me
    Exportar_Excel_Aporte
    
    Toolbar1.Enabled = True
    Frame1(0).Enabled = True

Case 3
    
    Set RS = vg_db.Execute("sgpadm_Sel_CecoMinutaBloque '" & Trim(LimpiaDato(fpText1.text)) & "'")
    If RS.EOF Then
       
       RS.Close
       Set RS = Nothing
       MsgBox "No existe ceco planificado", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    vg_codigo = ""
    B_HistPm.LlenarHistPlan "Histórico Minuta", 0, Trim(LimpiaDato(fpText1.text)), 5
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    Regimen.Value = vg_codregimen
    FpFecDesde.text = Format(dBoM("01/" & vg_fecha), "mm/yyyy")
    FpFecHasta.text = Format(dEoM("27/" & vg_fecha), "mm/yyyy")


Case 5
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub Exportar_Excel_Aporte()

On Error GoTo Man_Error
    
    Dim FecIni           As Date
    Dim i                As Long
    Dim Idietetico       As Long
    Dim indfin           As Long
    Dim RS               As New ADODB.Recordset
    Dim xx               As Long
    Dim JJ               As Long
    Dim PosAscAux        As String
    Dim NumAscAux        As Long
    Dim CantNut          As Long
    Dim NumCol           As Long
    Dim NumAsc           As Long
    Dim NumLinExcel      As Long
    Dim ColumnaExcel     As String
    Dim totsercanservida As Double
    Dim totdiacanservida As Double
    Dim canservida       As Double

    Dim nomservicio As String, NombreServicioAux As String
    Dim NomIngrediente As String
    Dim MnitmRef As Long, ctrfecha As Long, MnitmNo As Long
    Dim p As Long, z As Long, ipventa As Long, CodServicio As Long, auxcodservicio As Long
    Dim iprocesa As Long, totreg As Long, j As Long, idetapo As Long
    Dim numdia As Integer, SwSalto As Integer, SwTotal As Integer, exdiet As Integer, npage As Integer, iglosa As Integer
    Dim dAporte As String, vAporte As String, sAporte As String
    Dim DietItemYldVal1 As Double, NtrntVal As Double, DietItemConvVal As Double
    Dim cantcalorias As Double, cantproteinas As Double, cantlipidos As Double, canthidratos As Double, cantacgrsat As Double
    Dim porsolida As Double, porliquida As Double, totserporsolida As Double, totserporliquida As Double, totdiaporsolida As Double
    Dim totdiaporliquida As Double
    Dim pneto As Double, totserpneto As Double, totdiapneto As Double, canpavb As Double, totpavb As Double
    Dim pbruto As Double, totserpbruto As Double, totdiapbruto As Double
    Dim pNetoApr As Double, totserpnetoApr As Double, totdiapNetoApr As Double
    Dim cantotp As Double, cantotg As Double, cantotcho As Double, cantotdensidad As Double, cantotsacarosa As Double, cantotagrs As Double
    Dim cantotcalorias As Double, cantotproteinas As Double, cantotlipidos As Double, cantothidratos As Double, cantotacgrsat As Double

    Dim VecDie()     As String
    Dim VecCho()     As String
    
    Dim oExcel       As Object
    Dim oBook        As Object
    Dim oSheet       As Object

    '-------> Definir vector
    Dim MatrizCodDietetico()   As Long
    Dim MatrizMenuAporte()     As Double
    Dim Matrizaporteservicio() As Double
    Dim MatrizMenuTotAporte()  As Double
    Dim matrizcatr()           As Long
    Dim matrizglosa()          As String
    Dim VecDetAporte()         As Variant

    fg_carga ""
   
    '-------> Start a new workbook in Excel
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    
       Idietetico = 1
                       
       CantNut = 0
       For i = 1 To vaSpread2.MaxRows
        
           vaSpread2.Row = i
           vaSpread2.Col = 1
        
           If vaSpread2.text = "1" Then
           
              NumCol = NumCol + 1
              CantNut = CantNut + 1
           
           End If
    
       Next i
    
       ReDim VecDie(CantNut, 2)
       
       j = 1
       For i = 1 To vaSpread2.MaxRows
        
           vaSpread2.Row = i
           vaSpread2.Col = 1
        
           If vaSpread2.text = "1" Then
           
              vaSpread2.Col = 2
              VecDie(j, 1) = vaSpread2.text
              VecDie(j, 2) = ""
           
              j = j + 1
        
           End If
    
       Next i
       
       xx = 1
       JJ = 1
       NumAscAux = 69
       PosAscAux = ""
       NumAsc = 69

       For i = 0 To vaSpread2.MaxRows
                
           vaSpread2.Row = i
           vaSpread2.Col = 1
                
           If vaSpread2.text = "1" Then
                   
              vaSpread2.Col = 2
              ReDim Preserve MatrizCodDietetico(Idietetico)
              MatrizCodDietetico(Idietetico) = Val(vaSpread2.text)
                   
              vaSpread2.Col = 3
              Idietetico = Idietetico + 1
                
              ColumnaExcel = PosAscAux + Chr(xx + NumAsc)
              VecDie(JJ, 2) = PosAscAux + Chr(xx + NumAsc)

              xx = xx + 1
              JJ = JJ + 1
              If xx + NumAsc >= 90 Then
                  
                 NumAscAux = NumAscAux + 1
                 NumAsc = 64
                 PosAscAux = Chr(NumAscAux)
                   
                 xx = 1
                 
              End If
                
           End If
            
       Next i
                  
       xx = 1
       NumAsc = Asc(ColumnaExcel)
       If Option3(0).Value = True Then
       
          ReDim VecCho(4, 2)
          
          For i = 1 To 4
          
              Select Case i
                
                Case 1
                
                    VecCho(i, 1) = "P%"
                    VecCho(i, 2) = PosAscAux + Chr(xx + NumAsc)
      
                Case 2
              
                    VecCho(i, 1) = "G%"
                    VecCho(i, 2) = PosAscAux + Chr(xx + NumAsc)
                
                Case 3
               
                    VecCho(i, 1) = "Cho%"
                    VecCho(i, 2) = PosAscAux + Chr(xx + NumAsc)
                     
                Case 4
             
                    VecCho(i, 1) = "AGS"
                    VecCho(i, 2) = PosAscAux + Chr(xx + NumAsc)
                     
              End Select
          
              xx = xx + 1
              If xx + NumAsc >= 90 Then
                  
                 NumAscAux = NumAscAux + 1
                 NumAsc = 64
                 PosAscAux = Chr(NumAscAux)
                   
                 xx = 1
                 
              End If
          
          Next i
       
       Else
       
          ReDim VecCho(6, 2)
       
          For i = 1 To 6
          
              Select Case i
                
                Case 1
                
                    VecCho(i, 1) = "Pavb"
                    VecCho(i, 2) = PosAscAux + Chr(xx + NumAsc)
      
                Case 2
              
                    VecCho(i, 1) = "Pavb%"
                    VecCho(i, 2) = PosAscAux + Chr(xx + NumAsc)
                
                Case 3
               
                    VecCho(i, 1) = "P%"
                    VecCho(i, 2) = PosAscAux + Chr(xx + NumAsc)
                     
                Case 4
             
                    VecCho(i, 1) = "G%"
                    VecCho(i, 2) = PosAscAux + Chr(xx + NumAsc)
                     
                Case 5
             
                    VecCho(i, 1) = "Cho%"
                    VecCho(i, 2) = PosAscAux + Chr(xx + NumAsc)
                     
                Case 6
             
                    VecCho(i, 1) = "AGS %"
                    VecCho(i, 2) = PosAscAux + Chr(xx + NumAsc)
                     
              End Select
          
              xx = xx + 1
              If xx + NumAsc >= 90 Then
                  
                 NumAscAux = NumAscAux + 1
                 NumAsc = 64
                 PosAscAux = Chr(NumAscAux)
                   
                 xx = 1
                 
              End If
          
          Next i
       
       End If
       
       Idietetico = Idietetico - 1
       
       If Idietetico < 1 Then Idietetico = 1
       
       indfin = (((Idietetico) * 3) + 2)
            
       ReDim Preserve MatrizMenuAporte(Idietetico)
       ReDim Preserve Matrizaporteservicio(Idietetico)
       ReDim Preserve MatrizMenuTotAporte(Idietetico)
       ReDim VecDetAporte(5000, 6 + Idietetico)
            
       For i = 1 To UBound(VecDetAporte)
                
           VecDetAporte(i, 1) = 0
           VecDetAporte(i, 2) = ""
                
           For j = 7 To 3 + Idietetico
                    
               VecDetAporte(i, j) = 0
                
           Next j
            
       Next i
                        
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Set RS = vg_db.Execute("sgpadm_Sel_CountAportes")
            
       If Not RS.EOF Then
       
          totreg = RS!nReg
          ReDim matrizdiet(RS!nReg, indfin) As Double
            
       End If
       
       RS.Close
       Set RS = Nothing
            
       MnitmRef = 0
       ctrfecha = 0
       numdia = 1
       SwSalto = 0
       SwTotal = 0
       exdiet = 0
       CodServicio = 0
       totserporsolida = 0
       totserporliquida = 0
       totdiaporsolida = 0
       totdiaporliquida = 0
       pneto = 0
       totdiapneto = 0
       totdiapNetoApr = 0
       totserpneto = 0
       totserpnetoApr = 0
            
       totsercanservida = 0
       totdiacanservida = 0
       canservida = 0
            
       pbruto = 0
       pNetoApr = 0
       totdiapbruto = 0
       totserpbruto = 0
            
       porsolida = 0
       porliquida = 0
       totserporsolida = 0
       totserporliquida = 0
       totdiaporsolida = 0
       totdiaporliquida = 0
       cantotp = 0
       cantotg = 0
       cantotcho = 0
       cantotdensidad = 0
       cantotsacarosa = 0
       cantotagrs = 0
       cantotcalorias = 0
       cantotproteinas = 0
       cantotlipidos = 0
       cantothidratos = 0
       cantotacgrsat = 0
       iprocesa = 0
       
       ReDim Preserve matrizglosa(0)
       
       iglosa = 1
            
       YCurrentPage = 1
       WsLinea = 8
            
       Dim MyBufferServicio As String
       Let MyBufferServicio = ""
       Let MyBufferServicio = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
       Let MyBufferServicio = MyBufferServicio & "<Servicio>"
           
       For i = 1 To vaSpread1.MaxRows
               
           vaSpread1.Row = i
           vaSpread1.Col = 1
               
           If vaSpread1.text = "1" Then
                  
              vaSpread1.Col = 2
              MyBufferServicio = MyBufferServicio & " <Ser"
              MyBufferServicio = MyBufferServicio & " Ser = " & Chr(34) & vaSpread1.text & Chr(34)
              Let MyBufferServicio = MyBufferServicio & "/>"
               
           End If
           
       Next i
       Let MyBufferServicio = MyBufferServicio & "</Servicio>"
            
       '-------> Mover aportes nutriconales
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       Set RS = vg_db.Execute("sgpadm_Sel_AporteMinutasSansis_V03 '" & MyBufferServicio & "', '" & Trim(fpText1.text) & "', " & Regimen.Value & ", " & Format(FpFecDesde.text, "yyyymm") & ", " & Format(FpFecHasta.text, "yyyymm") & "")
                   
       If RS.EOF Then
       
          RS.Close
          Set RS = Nothing
          Exit Sub
          
       End If
       
       vaSpread3.MaxRows = 0
       vaSpread3.maxcols = 4
       Do While Not RS.EOF
              
          vaSpread3.MaxRows = vaSpread3.MaxRows + 1
          vaSpread3.Row = vaSpread3.MaxRows
          vaSpread3.Col = 1
          vaSpread3.text = RS!red_codpro
              
          vaSpread3.Col = 2
          vaSpread3.text = RS!pnu_codapo
              
          vaSpread3.Col = 3
          vaSpread3.text = RS!pnu_canapo
            
          vaSpread3.Col = 4
          vaSpread3.text = RS!ing_facnut
              
          RS.MoveNext
           
       Loop
       RS.Close
       Set RS = Nothing
            
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       Set RS = vg_db.Execute("sgpadm_Sel_XmlInfAporteNutricionalSANSIS_V05 '" & MyBufferServicio & "','" & Trim(fpText1.text) & "', " & Regimen.Value & ", " & Format(FpFecDesde.text, "yyyymm") & ", " & Format(FpFecHasta.text, "yyyymm") & ", '" & IncluyeGrsCero.Value & "'")
            
       If Not RS.EOF Then
               
           '-------> Crear Nueva Hoja Excel
           Set oSheet = oBook.Worksheets.Add
           NombreServicioAux = "Aportes Sansis"
           If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
              
              NombreServicioAux = "Aportes Sansis"
           
           End If
           
           oSheet.Name = NombreServicioAux
               
           '-------> Impresión titulo informe
           MoverDatosExcel oExcel, oSheet, "D", "D", 1, 1, "Aporte Nutricional " & IIf(Option4(0).Value = 1, "Resumido ", "Detallado")
          
           MoverDatosExcel oExcel, oSheet, "A", "A", 3, 3, "Casino  : " & IIf(Option1(1).Value = True, Trim(fpText1.text) & " - " & Trim(fpayuda(0).Caption), Trim(fpayuda(0).Caption))
           MoverDatosExcel oExcel, oSheet, "A", "A", 4, 4, "Regimen : " & Trim(fpayuda(1).Caption)
           MoverDatosExcel oExcel, oSheet, "A", "A", 5, 5, "Correspondiente al periodo de : " & Format(FpFecDesde.text, "mm/yyyy") & " Hasta " & Format(FpFecHasta.text, "mm/yyyy")
           
           NumLinExcel = 7
                          
           Do While Not RS.EOF
                  
              iprocesa = 1
                  
              For ipventa = 1 To vaSpread1.MaxRows
                      
                  vaSpread1.Row = ipventa
                  vaSpread1.Col = 2
                  auxcodservicio = Val(vaSpread1.Value)
                  vaSpread1.Col = 1
                      
                  If vaSpread1.Value = "1" And auxcodservicio = RS![Codigo Servicio] Then
                         
                     iprocesa = 0
                     Exit For
                  
                  End If
                  
              Next ipventa
                  
              If iprocesa = 0 Then
                     
                 If RS![Fecha Minuta] <> ctrfecha Then
                      
                        If SwSalto = 1 Then
                           
                           numdia = numdia + 1
                           
                           If SwTotal > 0 Then
                                                             
                               If Option4(1).Value = True Then
                                  
                                  NumLinExcel = NumLinExcel + 2
                               
                               End If
                               
                               MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, vAporte
                
                               MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(porsolida, fg_Pict(4, 2))
                
                               MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(pbruto, fg_Pict(4, 2))
                
                               MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(pNetoApr, fg_Pict(4, 2))
           
                               MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(pneto, fg_Pict(4, 2))
                              
                              For i = 1 To Idietetico

                                   MoverDatosExcel oExcel, oSheet, VecDie(i, 2), VecDie(i, 2), NumLinExcel, NumLinExcel, Format(MatrizMenuAporte(i), fg_Pict(4, 2))

                              Next i
                              
                              NumLinExcel = NumLinExcel + 1
                                                                
                              ReDim Preserve matrizglosa(iglosa)
                              iglosa = iglosa + 1
                              vAporte = ""
                              
                              For i = 1 To Idietetico
                                  
                                  MatrizMenuAporte(i) = 0
                              
                              Next i
                              
                              If Option4(1).Value = True Then
                                 
                                 '-------> Imprimir detalle
                                 For i = 1 To UBound(VecDetAporte)
                                     
                                     If VecDetAporte(i, 1) > 0 Then
                                        
                                        MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "  " & VecDetAporte(i, 2)
                
                                        MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 3), fg_Pict(4, 2))
                
                                        MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 4), fg_Pict(4, 2))
                
                                        MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 5), fg_Pict(4, 2))
           
                                        MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 6), fg_Pict(4, 2))

                                        For j = 7 To Idietetico + 6
                                            
                                            MoverDatosExcel oExcel, oSheet, VecDie(j - 6, 2), VecDie(j - 6, 2), NumLinExcel, NumLinExcel, Format(VecDetAporte(i, j), fg_Pict(4, 2))
                                        
                                        
                                        Next j
                                        
                                        NumLinExcel = NumLinExcel + 1
                                        ReDim Preserve matrizglosa(iglosa)
                                        iglosa = iglosa + 1
                                        vAporte = ""
                                      
                                      Else
                                         
                                         Exit For
                                      
                                      End If
                                  
                                  Next i
                                  
                                  For i = 1 To UBound(VecDetAporte)
                                      
                                      VecDetAporte(i, 1) = 0
                                      VecDetAporte(i, 2) = ""
                                      VecDetAporte(i, 3) = 0
                                      VecDetAporte(i, 4) = 0
                                      VecDetAporte(i, 5) = 0
                                      VecDetAporte(i, 6) = 0
                                      
                                      For j = 7 To Idietetico + 6
                                          
                                          VecDetAporte(i, j) = 0
                                      
                                      Next j
                                  
                                  Next i
                                  
                                  ReDim Preserve matrizglosa(iglosa)
                                  matrizglosa(iglosa) = ""
                                  iglosa = iglosa + 1
                              
                              End If
                              
                              vAporte = ""
                           
                           End If
                           
                           ReDim Preserve matrizglosa(iglosa)
                           matrizglosa(iglosa) = ""
                           iglosa = iglosa + 1
                           
                           NumLinExcel = NumLinExcel + 1
                           
                           MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "Total "
                
                           MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(totserporsolida, fg_Pict(4, 2))
                
                           MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(totserpbruto, fg_Pict(4, 2))
                
                           MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(totserpnetoApr, fg_Pict(4, 2))
           
                           MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(totserpneto, fg_Pict(4, 2))
                           
                           
                           For i = 1 To Idietetico
                               
                               MoverDatosExcel oExcel, oSheet, VecDie(i, 2), VecDie(i, 2), NumLinExcel, NumLinExcel, Format(Matrizaporteservicio(i), fg_Pict(4, 2))
                               
                           Next i
                           
                           cantotp = 0
                           cantotg = 0
                           cantotcho = 0
                           cantotagrs = 0
                           
                           If cantproteinas > 0 And cantcalorias > 0 Then
                              
                              cantotp = CCur(((cantproteinas * 4) / cantcalorias) * 100)
                           
                           End If
                           
                           If cantlipidos > 0 And cantcalorias > 0 Then
                              
                              cantotg = CCur(((cantlipidos * 9) / cantcalorias) * 100)
                           
                           End If
                           
                           If canthidratos > 0 And cantcalorias > 0 Then
                              
                              cantotcho = CCur(((canthidratos * 4) / cantcalorias) * 100)
                           
                           End If
                           
                           If cantacgrsat > 0 And cantcalorias > 0 Then
                              
                              cantotagrs = CCur(((cantacgrsat * 9) / cantcalorias) * 100)
                           
                           End If
                           
                           If Option3(1).Value = True And cantproteinas > 0 Then

                              '-------> Mover Alternativa %
                              For i = 1 To 6
                              
                                  Select Case i
                                                                        
                                    Case 1
                              
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(canpavb, fg_Pict(4, 2))
                                        
                                    Case 2
                                        
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(CCur((canpavb / cantproteinas) * 100), fg_Pict(4, 2))

                                    Case 3
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotp, fg_Pict(4, 2))
                                    
                                    Case 4
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotg, fg_Pict(4, 2))
                                    
                                    Case 5
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotcho, fg_Pict(4, 2))
                                    
                                    Case 6
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotagrs, fg_Pict(4, 2))
                                        
                                  End Select
        
                     
                              Next i
                                                      
                           Else
                           
                              '-------> Mover Alternativa %
                              For i = 1 To 4
                              
                                  Select Case i
                                                                        
                                    Case 1
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotp, fg_Pict(4, 2))
                                    
                                    Case 2
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotg, fg_Pict(4, 2))
                                    
                                    Case 3
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotcho, fg_Pict(4, 2))
                                    
                                    Case 4
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotagrs, fg_Pict(4, 2))
                                        
                                  End Select
                     
                              Next i
                           
                           End If
                           
                           NumLinExcel = NumLinExcel + 1
                           
                           ReDim Preserve matrizglosa(iglosa)
                           iglosa = iglosa + 1
                           totserporsolida = 0
                           totserporliquida = 0
                           totserpneto = 0
                           pneto = 0
                           pNetoApr = 0
                           totserpbruto = 0
                           totserpnetoApr = 0
                           pbruto = 0
                           pNetoApr = 0
                           canpavb = 0
                           cantcalorias = 0
                           cantproteinas = 0
                           cantlipidos = 0
                           canthidratos = 0
                           cantacgrsat = 0
                           ReDim Preserve matrizglosa(iglosa)
                           matrizglosa(iglosa) = ""
                           iglosa = iglosa + 1
                           
                           For i = 1 To Idietetico
                               
                               Matrizaporteservicio(i) = 0
                           
                           Next i
                           
                           sAporte = ""
                           
                           ' *** Imprimir totales del día aportes nutricionales *** '
                           ReDim Preserve matrizglosa(iglosa)
                           matrizglosa(iglosa) = ""
                           iglosa = iglosa + 1
                                                     
                           NumLinExcel = NumLinExcel + 1
                           
                           MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "A. Nutricional del Día "
                
                           MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(totdiaporsolida, fg_Pict(4, 2))
                
                           MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(totdiapbruto, fg_Pict(4, 2))
                
                           MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(totdiapNetoApr, fg_Pict(4, 2))
           
                           MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(totdiapneto, fg_Pict(4, 2))
                           
                           For i = 1 To Idietetico
                               
                               MoverDatosExcel oExcel, oSheet, VecDie(i, 2), VecDie(i, 2), NumLinExcel, NumLinExcel, Format(MatrizMenuTotAporte(i), fg_Pict(4, 2))
                           
                           Next i
                           
                           cantotp = 0
                           cantotg = 0
                           cantotcho = 0
                           cantotagrs = 0
                           
                           If cantotproteinas > 0 And cantotcalorias > 0 Then
                              
                              cantotp = CCur(((cantotproteinas * 4) / cantotcalorias) * 100)
                              
                           End If
                           
                           If cantotlipidos > 0 And cantotcalorias > 0 Then
                              
                              cantotg = CCur(((cantotlipidos * 9) / cantotcalorias) * 100)
                           
                           End If
                           
                           If cantothidratos > 0 And cantotcalorias > 0 Then
                              
                              cantotcho = CCur(((cantothidratos * 4) / cantotcalorias) * 100)
                           
                           End If
                           
                           If cantotacgrsat > 0 And cantotcalorias > 0 Then
                              
                              cantotagrs = CCur(((cantotacgrsat * 9) / cantotcalorias) * 100)
                           
                           End If
                           
                           If Option3(1).Value = True And cantotproteinas > 0 Then
                              
                              '-------> Mover Alternativa %
                              For i = 1 To 6
                              
                                  Select Case i
                                                                        
                                    Case 1
                              
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(totpavb, fg_Pict(4, 2))
                                        
                                    Case 2
                                        
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(CCur((totpavb / cantotproteinas) * 100), fg_Pict(4, 2))

                                    Case 3
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotp, fg_Pict(4, 2))
                                    
                                    Case 4
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotg, fg_Pict(4, 2))
                                    
                                    Case 5
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotcho, fg_Pict(4, 2))
                                    
                                    Case 6
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotagrs, fg_Pict(4, 2))
                                        
                                  End Select
        
                     
                              Next i
                                                      
                           Else
                           
                              '-------> Mover Alternativa %
                              For i = 1 To 4
                              
                                  Select Case i
                                                                        
                                    Case 1
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotp, fg_Pict(4, 2))
                                    
                                    Case 2
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotg, fg_Pict(4, 2))
                                    
                                    Case 3
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotcho, fg_Pict(4, 2))
                                    
                                    Case 4
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotagrs, fg_Pict(4, 2))
                                        
                                  End Select
                     
                              Next i
                           
                           End If
                           
                           NumLinExcel = NumLinExcel + 1
                           
                           ReDim Preserve matrizglosa(iglosa)
                           iglosa = iglosa + 1
                           
                           vAporte = ""
                           cantotcalorias = 0
                           cantotproteinas = 0
                           cantotlipidos = 0
                           cantothidratos = 0
                           cantotacgrsat = 0
                           totserporsolida = 0
                           totserporliquida = 0
                           totdiaporliquida = 0
                           totdiaporsolida = 0
                           totserpneto = 0
                           totsercanservida = 0
                           totdiapneto = 0
                           totdiacanservida = 0
                           pneto = 0
                           canservida = 0
                           totserpbruto = 0
                           totserpnetoApr = 0
                           totdiapbruto = 0
                           totdiapNetoApr = 0
                           pbruto = 0
                           pNetoApr = 0
                           
                           totpavb = 0
                           canpavb = 0
                           ' *** Fin imprimir totales del día aportes nutricionales *** '
                           
                           For i = 1 To Idietetico
                               
                               MatrizMenuTotAporte(i) = 0
                           
                           Next i
                           ReDim Preserve matrizglosa(0)
                           iglosa = 1
                           
                        End If
            
                        ReDim Preserve matrizglosa(iglosa)
                        iglosa = iglosa + 1
                       
                        NumLinExcel = NumLinExcel + 1
                        
                        MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "Fecha    Servicio"
                        DibujarLineas oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel
                
                        MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, "G/V Servir"
                        DibujarLineas oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel
                
                        MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, "Bruta"
                        DibujarLineas oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel
                
                        MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, "G/V Neto"
                        DibujarLineas oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel
           
                        MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, "G/V Neto Nut"
                        DibujarLineas oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel
                        
                       '-------> Mover aportes nutricionales
                       JJ = 1
                       For j = 1 To vaSpread2.MaxRows
               
                           vaSpread2.Row = j
                           vaSpread2.Col = 1
               
                           If vaSpread2.text = "1" Then
                  
                              vaSpread2.Col = 3
                              MoverDatosExcel oExcel, oSheet, VecDie(JJ, 2), VecDie(JJ, 2), NumLinExcel, NumLinExcel, vaSpread2.text
                              DibujarLineas oExcel, oSheet, VecDie(JJ, 2), VecDie(JJ, 2), NumLinExcel, NumLinExcel
                              ColumnaExcel = VecDie(JJ, 2)
                              JJ = JJ + 1
               
                           End If
           
                       Next j
           
                       '-------> Mover Alternativa %
                       For i = 1 To IIf(Option3(0).Value = True, 4, 6)
        
                           MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, VecCho(i, 1)
                           DibujarLineas oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel
                     
                        Next i
                        
                        NumLinExcel = NumLinExcel + 1
                        
                        MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "Fecha : " & CStr(Mid(RS![Fecha Minuta], 7, 2)) & "/" & CStr(Mid(RS![Fecha Minuta], 5, 2)) & "/" & CStr(Mid(RS![Fecha Minuta], 1, 4))
                        
                        NumLinExcel = NumLinExcel + 1
                        
                        ReDim Preserve matrizglosa(iglosa)
                        iglosa = iglosa + 1
                        
                        MnitmRef = 0
                        MnitmNo = 0
                        SwTotal = 0
                        CodServicio = 0
                        ctrfecha = RS![Fecha Minuta]
                        SwSalto = 1
                        
                  End If
                     
                     If RS![Codigo Servicio] <> CodServicio Then
                        
                        If CodServicio > 0 Then
                           
                           If Option4(1).Value = True Then
                           
                              NumLinExcel = NumLinExcel + 1
                           
                           End If
                           
                           MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, vAporte
                
                           MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(porsolida, fg_Pict(4, 2))
                
                           MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(pbruto, fg_Pict(4, 2))
                
                           MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(pNetoApr, fg_Pict(4, 2))
           
                           MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(pneto, fg_Pict(4, 2))
                           
                           For i = 1 To Idietetico

                               MoverDatosExcel oExcel, oSheet, VecDie(i, 2), VecDie(i, 2), NumLinExcel, NumLinExcel, Format(MatrizMenuAporte(i), fg_Pict(4, 2))
                           
                           Next i
                           
                           NumLinExcel = NumLinExcel + 1
                           
                           ReDim Preserve matrizglosa(iglosa)
                           iglosa = iglosa + 1
                           vAporte = ""
                           
                           For i = 1 To Idietetico
                               
                               MatrizMenuAporte(i) = 0
                           
                           Next i
                           
                           SwTotal = 0
                           If Option4(1).Value = True Then
                              
                              '-------> Imprimir detalle
                              
                              For i = 1 To UBound(VecDetAporte)
                                  
                                  If VecDetAporte(i, 1) > 0 Then
                                     
                                     MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "  " & VecDetAporte(i, 2)
                
                                     MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 3), fg_Pict(4, 2))
                
                                     MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 4), fg_Pict(4, 2))
                
                                     MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 5), fg_Pict(4, 2))
           
                                     MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 6), fg_Pict(4, 2))
                                     
                                     For j = 7 To Idietetico + 6
                                         
                                         MoverDatosExcel oExcel, oSheet, VecDie(j - 6, 2), VecDie(j - 6, 2), NumLinExcel, NumLinExcel, Format(VecDetAporte(i, j), fg_Pict(4, 2))
                                     
                                     Next j
                                     
                                     NumLinExcel = NumLinExcel + 1
                                     ReDim Preserve matrizglosa(iglosa)
                                     iglosa = iglosa + 1
                                     vAporte = ""
                                  
                                  Else
                                     
                                     Exit For
                                  
                                  End If
                              
                              Next i
                              
                              For i = 1 To UBound(VecDetAporte)
                                  
                                  VecDetAporte(i, 1) = 0
                                  VecDetAporte(i, 2) = ""
                                  VecDetAporte(i, 3) = 0
                                  VecDetAporte(i, 4) = 0
                                  VecDetAporte(i, 5) = 0
                                  VecDetAporte(i, 6) = 0
                                  
                                  For j = 7 To Idietetico + 6
                                      
                                      VecDetAporte(i, j) = 0
                                  
                                  Next j
                              
                              Next i
                              
                              ReDim Preserve matrizglosa(iglosa)
                              matrizglosa(iglosa) = ""
                              iglosa = iglosa + 1
                           
                           End If
                           vAporte = ""
                             
                           ReDim Preserve matrizglosa(iglosa)
                           matrizglosa(iglosa) = ""
                           iglosa = iglosa + 1
                           
                           NumLinExcel = NumLinExcel + 1
                           
                           MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "Total "
                
                           MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(totserporsolida, fg_Pict(4, 2))
                
                           MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(totserpbruto, fg_Pict(4, 2))
                
                           MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(totserpnetoApr, fg_Pict(4, 2))
           
                           MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(totserpneto, fg_Pict(4, 2))
                                        
                           For i = 1 To Idietetico
                               
                                MoverDatosExcel oExcel, oSheet, VecDie(i, 2), VecDie(i, 2), NumLinExcel, NumLinExcel, Format(Matrizaporteservicio(i), fg_Pict(4, 2))
                           
                           Next i
                           
                           cantotp = 0
                           cantotg = 0
                           cantotcho = 0
                           cantotagrs = 0
                           
                           If cantproteinas > 0 And cantcalorias > 0 Then
                              
                              cantotp = CCur(((cantproteinas * 4) / cantcalorias) * 100)
                           
                           End If
                           
                           If cantlipidos > 0 And cantcalorias > 0 Then
                              
                              cantotg = CCur(((cantlipidos * 9) / cantcalorias) * 100)
                           
                           End If
                           
                           If canthidratos > 0 And cantcalorias > 0 Then
                              
                              cantotcho = CCur(((canthidratos * 4) / cantcalorias) * 100)
                           
                           End If
                           
                           If cantacgrsat > 0 And cantcalorias > 0 Then
                              
                              cantotagrs = CCur(((cantacgrsat * 9) / cantcalorias) * 100)
                           
                           End If
                           
                           If Option3(1).Value = True And cantproteinas > 0 Then
                              
                              '-------> Mover Alternativa %
                              For i = 1 To 6
                              
                                  Select Case i
                                                                        
                                    Case 1
                              
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(totpavb, fg_Pict(4, 2))
                                        
                                    Case 2
                                        
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(CCur((totpavb / cantotproteinas) * 100), fg_Pict(4, 2))

                                    Case 3
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotp, fg_Pict(4, 2))
                                    
                                    Case 4
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotg, fg_Pict(4, 2))
                                    
                                    Case 5
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotcho, fg_Pict(4, 2))
                                    
                                    Case 6
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotagrs, fg_Pict(4, 2))
                                        
                                  End Select
        
                     
                              Next i
                                                      
                           Else
                           
                              '-------> Mover Alternativa %
                              For i = 1 To 4
                              
                                  Select Case i
                                                                        
                                    Case 1
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotp, fg_Pict(4, 2))
                                    
                                    Case 2
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotg, fg_Pict(4, 2))
                                    
                                    Case 3
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotcho, fg_Pict(4, 2))
                                    
                                    Case 4
                                    
                                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotagrs, fg_Pict(4, 2))
                                        
                                  End Select
                     
                              Next i
                           
                           
                           End If
                           
                           NumLinExcel = NumLinExcel + 1
                           
                           ReDim Preserve matrizglosa(iglosa)
                           iglosa = iglosa + 1
                           ReDim Preserve matrizglosa(iglosa)
                           matrizglosa(iglosa) = ""
                           iglosa = iglosa + 1
                           
                           For i = 1 To Idietetico
                               
                               Matrizaporteservicio(i) = 0
                           
                           Next i
                           
                           totserporsolida = 0
                           totserporliquida = 0
                           totserpneto = 0
                           pneto = 0
                           pNetoApr = 0
                           totserpbruto = 0
                           totserpnetoApr = 0
                           pbruto = 0
                           
                           canpavb = 0
                           cantcalorias = 0
                           cantproteinas = 0
                           cantlipidos = 0
                           canthidratos = 0
                           cantacgrsat = 0
                           sAporte = ""
        ' ****
                           MnitmRef = 0
                           MnitmNo = 0
                           SwTotal = 0
                           canpavb = 0
                        
                        End If
                        
                        CodServicio = RS![Codigo Servicio]
                        
                        NumLinExcel = NumLinExcel + 1
                        
                        MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, " " & Trim(RS![Nombre Servicio])
                        
                        NumLinExcel = NumLinExcel + 1
                        
                        ReDim Preserve matrizglosa(iglosa)
                        iglosa = iglosa + 1
                        nomservicio = Trim(RS![Nombre Servicio])
                     
                     End If
                         
                     If RS![Codigo Minuta] <> MnitmRef Or RS![Nro. Linea] <> MnitmNo Then
                        
                        If SwTotal > 0 Then
                           
                           If Option4(1).Value = True Then
                           
                              NumLinExcel = NumLinExcel + 1
                           
                           End If
                           
                           MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, vAporte
                
                           MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(porsolida, fg_Pict(4, 2))
                
                           MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(pbruto, fg_Pict(4, 2))
                
                           MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(pNetoApr, fg_Pict(4, 2))
           
                           MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(pneto, fg_Pict(4, 2))
                                        
                           For i = 1 To Idietetico
                               
                               MoverDatosExcel oExcel, oSheet, VecDie(i, 2), VecDie(i, 2), NumLinExcel, NumLinExcel, Format(MatrizMenuAporte(i), fg_Pict(4, 2))
                           
                           Next i
                           
                           NumLinExcel = NumLinExcel + 1
                           
                           ReDim Preserve matrizglosa(iglosa)
                           iglosa = iglosa + 1
                           vAporte = ""
                           
                           For i = 1 To Idietetico
                               
                               MatrizMenuAporte(i) = 0
                           
                           Next i
                           
                           If Option4(1).Value = True Then
                              '-------> Imprimir detalle
                              
                              For i = 1 To UBound(VecDetAporte)
                                  
                                  If VecDetAporte(i, 1) > 0 Then
                                     
                                     MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, " " & VecDetAporte(i, 2)
                
                                     MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 3), fg_Pict(4, 2))
                
                                     MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 4), fg_Pict(4, 2))
                
                                     MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 5), fg_Pict(4, 2))
           
                                     MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 6), fg_Pict(4, 2))
                                     
                                     For j = 7 To Idietetico + 6
                                         
                                         MoverDatosExcel oExcel, oSheet, VecDie(j - 6, 2), VecDie(j - 6, 2), NumLinExcel, NumLinExcel, Format(VecDetAporte(i, j), fg_Pict(4, 2))
                                     
                                     Next j
                                     
                                     NumLinExcel = NumLinExcel + 1
                                     
                                     ReDim Preserve matrizglosa(iglosa)
                                     iglosa = iglosa + 1
                                     vAporte = ""
                                  
                                  Else
                                     
                                     Exit For
                                  
                                  End If
                              
                              Next i
                              
                              For i = 1 To UBound(VecDetAporte)
                                  
                                  VecDetAporte(i, 1) = 0
                                  VecDetAporte(i, 2) = ""
                                  VecDetAporte(i, 3) = 0
                                  VecDetAporte(i, 4) = 0
                                  VecDetAporte(i, 5) = 0
                                  VecDetAporte(i, 6) = 0
                                  
                                  For j = 7 To Idietetico + 6
                                      
                                      VecDetAporte(i, j) = 0
                                  
                                  Next j
                              
                              Next i
                              
                              ReDim Preserve matrizglosa(iglosa)
                              matrizglosa(iglosa) = ""
                              iglosa = iglosa + 1
                           
                           End If
                           vAporte = ""
                        
                        End If
                        
                        MnitmRef = RS![Codigo Minuta]
                        MnitmNo = RS![Nro. Linea]
                        porsolida = 0
                        porliquida = 0
                        pneto = 0
                        pNetoApr = 0
                        pbruto = 0
                        vAporte = Trim(RS![Nombre Receta]) 'Fg_SacaParentesis(Trim(RS![Nombre Receta]))
                        SwTotal = 1
                        idetapo = 1
                     
                     End If
                     
                     porsolida = CCur(porsolida + RS!canservida)
                     totserporsolida = CCur(totserporsolida + RS!canservida)
                     totdiaporsolida = CCur(totdiaporsolida + RS!canservida)
                        
                     pneto = CCur(pneto + Format(RS!pneto, fg_Pict(4, 2)))
                     totserpneto = CCur(totserpneto + RS!pneto)
                     totdiapneto = CCur(totdiapneto + RS!pneto)
                         
                     pNetoApr = CCur(pNetoApr + Format(RS!pNetoApr, fg_Pict(4, 2)))
                     totserpnetoApr = CCur(totserpnetoApr + RS!pNetoApr)
                     totdiapNetoApr = CCur(totdiapNetoApr + RS!pNetoApr)
                         
                     pbruto = CCur(pbruto + RS!wsnumporcion)
                     totserpbruto = CCur(totserpbruto + RS!wsnumporcion)
                     totdiapbruto = CCur(totdiapbruto + RS!wsnumporcion)
                         
                     cantcalorias = CCur(cantcalorias + RS!calorias)
                     cantproteinas = CCur(cantproteinas + RS!proteinas)
                     cantlipidos = CCur(cantlipidos + RS!lipidos)
                     canthidratos = CCur(canthidratos + RS!Hidratos)
                     cantacgrsat = CCur(cantacgrsat + RS!Acgrsat)
                         
                     cantotcalorias = CCur(cantotcalorias + RS!calorias)
                     cantotproteinas = CCur(cantotproteinas + RS!proteinas)
                     cantotlipidos = CCur(cantotlipidos + RS!lipidos)
                     cantothidratos = CCur(cantothidratos + RS!Hidratos)
                     cantotacgrsat = CCur(cantotacgrsat + RS!Acgrsat)
                         
                     If RS!indpavb = 1 Then
                        
                        canpavb = CCur(canpavb + RS!proteinas)
                        totpavb = CCur(totpavb + RS!proteinas)
                     
                     End If
                     
        'aqui reemplazar
        
                      Trim (CStr(RS![Codigo Ingrediente]))
                      ind_ini = vaSpread3.SearchCol(1, -1, vaSpread3.MaxRows, RS![Codigo Ingrediente], SearchFlagsEqual)
                      codpro = ""
                      
                      For ind_par = ind_ini To vaSpread3.MaxRows
                          
                          vaSpread3.Row = ind_par
                          vaSpread3.Col = 1
                          
                          If vaSpread3.text <> RS![Codigo Ingrediente] Then Exit For
                          
                          vaSpread3.Col = 2
                          codapo = vaSpread3.text
                          
                          vaSpread3.Col = 3
                          NtrntVal = vaSpread3.text
                          
                          vaSpread3.Col = 4
                          DietItemConvVal = vaSpread3.text
                          
                          VecDetAporte(idetapo, 1) = RS![Codigo Ingrediente]
                          VecDetAporte(idetapo, 2) = IIf(IsNull(RS![Nombre Ingrediente]), "", Trim(RS![Nombre Ingrediente]))
                          VecDetAporte(idetapo, 3) = IIf(IsNull(RS!canservida), 0, Format(RS!canservida, fg_Pict(4, 2)))
                          VecDetAporte(idetapo, 4) = IIf(IsNull(RS!wsnumporcion), 0, Format(RS!wsnumporcion, fg_Pict(4, 2)))
                          VecDetAporte(idetapo, 5) = IIf(IsNull(RS!pNetoApr), 0, Format(RS!pNetoApr, fg_Pict(4, 2)))
                          VecDetAporte(idetapo, 6) = IIf(IsNull(RS!pneto), 0, Format(RS!pneto, fg_Pict(4, 2)))
                          
                          DietItemYldVal1 = RS![Porcentaje Nutricional]
                          
                          For j = 1 To Idietetico
                              
                              If MatrizCodDietetico(j) = codapo Then
                                 
                                 MatrizMenuAporte(j) = CCur(MatrizMenuAporte(j) + ((((DietItemYldVal1 / 100) * (NtrntVal * (RS!wsnumporcion))) / DietItemConvVal)))
                                 MatrizMenuTotAporte(j) = CCur(MatrizMenuTotAporte(j) + ((((DietItemYldVal1 / 100) * (NtrntVal * (RS!wsnumporcion))) / DietItemConvVal)))
                                 Matrizaporteservicio(j) = CCur(Matrizaporteservicio(j) + ((((DietItemYldVal1 / 100) * (NtrntVal * (RS!wsnumporcion))) / DietItemConvVal)))
                                 
                                 VecDetAporte(idetapo, j + 6) = CCur(((DietItemYldVal1 / 100) * (NtrntVal * RS!wsnumporcion) / DietItemConvVal)) + VecDetAporte(idetapo, j + 6)
                              
                              End If
                          
                          Next j
                      
                      Next ind_par
                  
                  End If
                  
              RS.MoveNext
              idetapo = idetapo + 1
                  
           Loop
               
           If SwTotal > 0 Then
                  
              If Option4(1).Value = True Then
              
                 NumLinExcel = NumLinExcel + 1
                 
              End If
              
              MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, vAporte
                
              MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(porsolida, fg_Pict(4, 2))
                
              MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(pbruto, fg_Pict(4, 2))
                
              MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(pNetoApr, fg_Pict(4, 2))
           
              MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(pneto, fg_Pict(4, 2))
                  
              For i = 1 To Idietetico
                      
                  MoverDatosExcel oExcel, oSheet, VecDie(i, 2), VecDie(i, 2), NumLinExcel, NumLinExcel, Format(MatrizMenuAporte(i), fg_Pict(4, 2))
                   
              Next i
                  
              NumLinExcel = NumLinExcel + 1
              
              ReDim Preserve matrizglosa(iglosa)
              iglosa = iglosa + 1
              vAporte = ""
                  
              For i = 1 To Idietetico
                      
                  MatrizMenuAporte(i) = 0
                  
              Next i
                  
              If Option4(1).Value = True Then
                 '-------> Imprimir detalle
                     
                 For i = 1 To UBound(VecDetAporte)
                         
                     If VecDetAporte(i, 1) > 0 Then
                         
                        MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, " " & VecDetAporte(i, 2)
                
                        MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 3), fg_Pict(4, 2))
                
                        MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 4), fg_Pict(4, 2))
                
                        MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 5), fg_Pict(4, 2))
           
                        MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(VecDetAporte(i, 6), fg_Pict(4, 2))
                                        
                        For j = 7 To Idietetico + 6
                                
                            MoverDatosExcel oExcel, oSheet, VecDie(j - 6, 2), VecDie(j - 6, 2), NumLinExcel, NumLinExcel, Format(VecDetAporte(i, j), fg_Pict(4, 2))
                            
                        Next j
                            
                        NumLinExcel = NumLinExcel + 1
                        
                        ReDim Preserve matrizglosa(iglosa)
                        iglosa = iglosa + 1
                        vAporte = ""
                       
                     Else
                            
                        Exit For
                         
                     End If
                     
                 Next i
                     
                 For i = 1 To UBound(VecDetAporte)
                         
                     VecDetAporte(i, 1) = 0
                     VecDetAporte(i, 2) = ""
                     VecDetAporte(i, 3) = 0
                     VecDetAporte(i, 4) = 0
                     VecDetAporte(i, 5) = 0
                     VecDetAporte(i, 6) = 0
                         
                     For j = 7 To Idietetico + 6
                             
                         VecDetAporte(i, j) = 0
                         
                     Next j
                     
                 Next i
                     
                 ReDim Preserve matrizglosa(iglosa)
                 matrizglosa(iglosa) = ""
                 iglosa = iglosa + 1
                 
              End If
                 
              vAporte = ""
               
           End If
                     
           ReDim Preserve matrizglosa(iglosa)
           matrizglosa(iglosa) = ""
           iglosa = iglosa + 1
        
           NumLinExcel = NumLinExcel + 1
           
           MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "Total "
                
           MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(totserporsolida, fg_Pict(4, 2))
                
           MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(totserpbruto, fg_Pict(4, 2))
                
           MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(totserpnetoApr, fg_Pict(4, 2))
           
           MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(totserpneto, fg_Pict(4, 2))
               
           For i = 1 To Idietetico
                                 
               MoverDatosExcel oExcel, oSheet, VecDie(i, 2), VecDie(i, 2), NumLinExcel, NumLinExcel, Format(Matrizaporteservicio(i), fg_Pict(4, 2))
               
           Next i
               
           cantotp = 0
           cantotg = 0
           cantotcho = 0
           cantotagrs = 0
               
           If cantproteinas > 0 And cantcalorias > 0 Then
                  
              cantotp = CCur(((cantproteinas * 4) / cantcalorias) * 100)
               
           End If
               
           If cantlipidos > 0 And cantcalorias > 0 Then
                  
              cantotg = CCur(((cantlipidos * 9) / cantcalorias) * 100)
               
           End If
               
           If canthidratos > 0 And cantcalorias > 0 Then
                  
              cantotcho = CCur(((canthidratos * 4) / cantcalorias) * 100)
               
           End If
               
           If cantacgrsat > 0 And cantcalorias > 0 Then
                  
              cantotagrs = CCur(((cantacgrsat * 9) / cantcalorias) * 100)
               
           End If
                             
           If Option3(1).Value = True And cantproteinas > 0 Then
                  
              '-------> Mover Alternativa %
              For i = 1 To 6
                              
                  Select Case i
                                                                        
                    Case 1
                              
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(totpavb, fg_Pict(4, 2))
                           
                    Case 2
                                        
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(CCur((totpavb / cantotproteinas) * 100), fg_Pict(4, 2))

                    Case 3
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotp, fg_Pict(4, 2))
                                    
                    Case 4
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotg, fg_Pict(4, 2))
                                    
                    Case 5
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotcho, fg_Pict(4, 2))
                                    
                    Case 6
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotagrs, fg_Pict(4, 2))
                                        
                  End Select
        
                     
              Next i
                                                      
           Else
                           
              '-------> Mover Alternativa %
              For i = 1 To 4
                              
                Select Case i
                                                                        
                    Case 1
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotp, fg_Pict(4, 2))
                                    
                    Case 2
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotg, fg_Pict(4, 2))
                                    
                    Case 3
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotcho, fg_Pict(4, 2))
                                    
                    Case 4
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotagrs, fg_Pict(4, 2))
                                        
                End Select
                     
              Next i
               
           End If
               
           NumLinExcel = NumLinExcel + 1
           
           ReDim Preserve matrizglosa(iglosa)
           iglosa = iglosa + 1
           ReDim Preserve matrizglosa(iglosa)
           matrizglosa(iglosa) = ""
           iglosa = iglosa + 1
           totserporsolida = 0
           totserporliquida = 0
           totserpneto = 0
           totserpnetoApr = 0
           totserpbruto = 0
           cantcalorias = 0
           cantproteinas = 0
           cantlipidos = 0
           canthidratos = 0
           cantacgrsat = 0
                               
           For i = 1 To Idietetico
                   
               Matrizaporteservicio(i) = 0
               
           Next i
           sAporte = ""
        
           ' *** Imprimir totales del día aportes nutricionales *** '
           ReDim Preserve matrizglosa(iglosa)
           matrizglosa(iglosa) = ""
           iglosa = iglosa + 1
                      
           NumLinExcel = NumLinExcel + 1
           
           MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "A. Nutricional del Día"
                
           MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(totdiaporsolida, fg_Pict(4, 2))
                
           MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(totdiapbruto, fg_Pict(4, 2))
                
           MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(totdiapNetoApr, fg_Pict(4, 2))
           
           MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(totdiapneto, fg_Pict(4, 2))
               
           For i = 1 To Idietetico
                   
               MoverDatosExcel oExcel, oSheet, VecDie(i, 2), VecDie(i, 2), NumLinExcel, NumLinExcel, Format(MatrizMenuTotAporte(i), fg_Pict(4, 2))
               
           Next i
               
           cantotp = 0
           cantotg = 0
           cantotcho = 0
           cantotagrs = 0
               
           If cantotproteinas > 0 And cantotcalorias > 0 Then
                  
              cantotp = CCur(((cantotproteinas * 4) / cantotcalorias) * 100)
               
           End If
               
           If cantotlipidos > 0 And cantotcalorias > 0 Then
                  
              cantotg = CCur(((cantotlipidos * 9) / cantotcalorias) * 100)
               
           End If
               
           If cantothidratos > 0 And cantotcalorias > 0 Then
                  
              cantotcho = CCur(((cantothidratos * 4) / cantotcalorias) * 100)
               
           End If
               
           If cantotacgrsat > 0 And cantotcalorias > 0 Then
                  
              cantotagrs = CCur(((cantotacgrsat * 9) / cantotcalorias) * 100)
               
           End If
               
           If Option3(1).Value = True And cantotproteinas > 0 Then
                  
              '-------> Mover Alternativa %
              For i = 1 To 6
                              
                  Select Case i
                                                                        
                    Case 1
                              
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(totpavb, fg_Pict(4, 2))
                           
                    Case 2
                                        
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(CCur((totpavb / cantotproteinas) * 100), fg_Pict(4, 2))

                    Case 3
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotp, fg_Pict(4, 2))
                                    
                    Case 4
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotg, fg_Pict(4, 2))
                                    
                    Case 5
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotcho, fg_Pict(4, 2))
                                    
                    Case 6
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotagrs, fg_Pict(4, 2))
                                        
                  End Select
        
                     
              Next i
                                                      
           Else
                           
              '-------> Mover Alternativa %
              For i = 1 To 4
                              
                Select Case i
                                                                        
                    Case 1
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotp, fg_Pict(4, 2))
                                    
                    Case 2
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotg, fg_Pict(4, 2))
                                    
                    Case 3
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotcho, fg_Pict(4, 2))
                                    
                    Case 4
                                    
                        MoverDatosExcel oExcel, oSheet, VecCho(i, 2), VecCho(i, 2), NumLinExcel, NumLinExcel, Format(cantotagrs, fg_Pict(4, 2))
                                        
                End Select
                     
              Next i
               
           End If
               
           ReDim Preserve matrizglosa(iglosa)
           iglosa = iglosa + 1
                      
           vAporte = ""
           WsLinea = WsLinea + iglosa - 1
               
           ReDim Preserve matrizglosa(0)
           iglosa = 1
               
           cantotcalorias = 0
           cantotproteinas = 0
           cantotlipidos = 0
           cantothidratos = 0
           cantotacgrsat = 0
           totserporsolida = 0
           totserporliquida = 0
           totdiaporliquida = 0
           totdiaporsolida = 0
           totserpneto = 0
           totdiapneto = 0
           totserpnetoApr = 0
           totdiapNetoApr = 0
           totserpbruto = 0
           totdiapbruto = 0
               
           For i = 1 To Idietetico
                   
               MatrizMenuTotAporte(i) = 0
               
           Next i
                      
       End If
       RS.Close
       Set RS = Nothing

      oExcel.Visible = True '------->Visualizar
      Set oSheet = Nothing
      Set oExcel = Nothing
      Set oBook = Nothing

Exit Sub
Man_Error:
    Call fg_descarga
    
    oExcel.DisplayAlerts = False
    oExcel.Quit
    oExcel.DisplayAlerts = True
    
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub MoverDatosVector()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If LimpiaDato(Trim(fpText1.text)) = "" Or FpFecDesde.text = "" Or FpFecHasta.text = "" Or Val(Regimen.Value) = 0 Then

    Exit Sub

End If

fg_carga ""
Set RS = vg_db.Execute("sgpadm_Sel_ServicioMinutaBloque '" & LimpiaDato(Trim(fpText1.text)) & "', " & Regimen.Value & ", " & Format(FpFecDesde.text, "yyyymmdd") & ", " & Format(FpFecHasta.text, "yyyymmdd") & "")
vaSpread1.MaxRows = 0

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      
      vaSpread1.Col = 2
      vaSpread1.Value = RS(0)
      
      vaSpread1.Col = 3
      vaSpread1.Value = Trim(RS(1))
      
      RS.MoveNext
   
   Loop

End If
RS.Close
Set RS = Nothing
fg_descarga

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub ActivarGrillaServicio()

On Error GoTo Man_Error

Dim i As Long

For i = 1 To vaSpread1.MaxRows
    
    vaSpread1.Row = i
    vaSpread1.Col = 1
    vaSpread1.CellType = 10
    vaSpread1.TypeCheckText = ""
    vaSpread1.TypeCheckCenter = True
    vaSpread1.text = "1" ' checked

Next i

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub ActivarGrillaAportes()

On Error GoTo Man_Error

Dim iselecc As Integer
Dim i       As Long
iselecc = 0

For i = 1 To vaSpread2.MaxRows
    
    vaSpread2.Row = i: vaSpread2.Col = 1
    
    If vaSpread2.text = "1" Then iselecc = 1: Exit For

Next i

If iselecc = 0 Then
   
   For i = 1 To vaSpread2.MaxRows
       
       vaSpread2.Row = i
       vaSpread2.Col = 1
       vaSpread2.CellType = 10
       vaSpread2.TypeCheckText = ""
       vaSpread2.TypeCheckCenter = True
       vaSpread2.text = "1" ' checked
   
   Next i

End If

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Function ValidarDatos() As Boolean

Dim seleccion As Integer
Dim i As Long

ValidarDatos = True

'-------> Validar regimen
If Trim(fpayuda(0).Caption) = "" Then

   MsgBox "Debe registrar ceco...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

'-------> Validar regimen
If Trim(fpayuda(1).Caption) = "" Then

   MsgBox "Debe registrar regimen...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

'-------> Validar fechas
If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
   
   MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

If FpFecDesde.text > FpFecHasta.text Then
   
   MsgBox "Fecha Origen No Puede Ser Mayor Que Fecha Destino", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

If DateDiff("d", "01" & "/" & FpFecDesde.text, dEoM("27" & "/" & FpFecHasta.text)) > 98 Then
        
   Call MsgBox("Sobre pasa los 98 días corresponde a 14 semana", vbInformation, Me.Caption)
   Let ValidarDatos = False
   Exit Function
        
End If
    
If DateDiff("m", "01" & "/" & FpFecDesde.text, dEoM("27" & "/" & FpFecHasta.text)) + 1 > 3 Then
        
   Call MsgBox("Rango De Fecha No Puede Ser Mayor a 3 Meses", vbInformation, Me.Caption)
   Let ValidarDatos = False
   Exit Function
        
End If

If Option2(0).Value = True Then ActivarGrillaServicio
If Option2(2).Value = True Then ActivarGrillaAportes
    
'-------> Validar servicios
seleccion = 0
For i = 1 To vaSpread1.MaxRows
        
    vaSpread1.Row = i: vaSpread1.Col = 1
    
    If vaSpread1.text = "1" Then
       
       seleccion = 1
       Exit For
    
    End If
    
Next i
    
If seleccion = 0 Then
   
   MsgBox "Seleccione Opción Dentro Grilla", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function
   
End If
    
'-------> Validar aportes nutricionales
seleccion = 0
    
For i = 1 To vaSpread2.MaxRows
        
    vaSpread2.Row = i: vaSpread2.Col = 1
    
    If vaSpread2.text = "1" Then
       
       seleccion = 1
       Exit For
    
    End If

Next i
    
If seleccion = 0 Then
   
   MsgBox "Debe Seleccionar A lo Menos Un Aporte Nutricional", vbExclamation + vbOKOnly, "Aporte Nutricional"
   ValidarDatos = False
   Exit Function

End If

End Function

