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
         Height          =   735
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   3855
         Begin VB.OptionButton Option5 
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
            Height          =   195
            Index           =   1
            Left            =   2400
            TabIndex        =   36
            Top             =   360
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
            TabIndex        =   35
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
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
         Top             =   2160
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
         Index           =   0
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
         SpreadDesigner  =   "I_ApoNutSansis.frx":0000
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
         Top             =   3540
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
            Picture         =   "I_ApoNutSansis.frx":4466
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
            Picture         =   "I_ApoNutSansis.frx":4770
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
         Index           =   0
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
         SpreadDesigner  =   "I_ApoNutSansis.frx":4A7A
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
         AllowNull       =   -1  'True
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
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
         Text            =   ""
         DateCalcMethod  =   1
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
         AutoMenu        =   -1  'True
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
         AllowNull       =   -1  'True
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
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
         Text            =   ""
         DateCalcMethod  =   1
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
         AutoMenu        =   -1  'True
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
         Picture         =   "I_ApoNutSansis.frx":4EDE
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3060
         Picture         =   "I_ApoNutSansis.frx":51E8
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
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

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
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico Planificacón Teórica"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

fpText1.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText1.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)

vaSpread1(0).MaxRows = 0
vaSpread2(0).MaxRows = 0

FpFecDesde.text = Format(Date, "dd/mm/yyyy")
FpFecHasta.text = Format(Date, "dd/mm/yyyy")

'------- Llenar Tabla Diéteticas
Set RS = vg_db.Execute("sgp_Sel_Nutriente 1, 0, ''")

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
      
      vaSpread2(0).MaxRows = vaSpread2(0).MaxRows + 1
      vaSpread2(0).Row = vaSpread2(0).MaxRows
      
      If RS!nut_indpri > 0 Then
         
         vaSpread2(0).Col = 1
         vaSpread2(0).CellType = 10
         vaSpread2(0).TypeCheckText = ""
         vaSpread2(0).TypeCheckCenter = True
         vaSpread2(0).text = "1" ' checked
      
      Else
         
         vaSpread2(0).Col = 1
         vaSpread2(0).CellType = 10
         vaSpread2(0).TypeCheckText = ""
         vaSpread2(0).TypeCheckCenter = True
         vaSpread2(0).text = " " ' checked
      
      End If
      
      vaSpread2(0).Col = 2: vaSpread2(0).Value = RS(0)
      vaSpread2(0).Col = 3: vaSpread2(0).Value = Trim(RS(1))
      
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

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub
MoverDatosVector

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

If IsDate(FpFecHasta.text) = False Then Exit Sub
MoverDatosVector

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpText1_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Set RS = vg_db.Execute("sgp_Sel_Clientes 1, '" & LimpiaDato(fpText1.text) & "'")
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

fpayuda(0).Caption = Trim(RS!cli_nombre)
fpText1.text = RS!cli_codigo
RS.Close
Set RS = Nothing
 
FpFecDesde.Enabled = True
FpFecHasta.Enabled = True

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
    
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
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0
    
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = ""
    vg_codigo = ""
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
    Call B_TabEst.LlenaDatos("a_regimen", "reg_", "Regimen", "RegBlo")
    Call B_TabEst.Show(1)
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    Regimen.Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    FpFecDesde.SetFocus

Case 2
    
    vg_left = fpayuda(0).Left + 2300
    OpcionLectura = "5"
    vg_nombre = ""
    vg_codigo = ""
    vg_codigo = Trim(LimpiaDato(fpText1.text))
    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, vg_codigo, Regimen.text & ",", Format(FpFecDesde.text, "yyyymmdd"), Format(FpFecHasta.text, "yyyymmdd"), "1", "", 0, IIf(Option5(0).Value = True, 1, 2)
    B_MTaEst.Show 1
    Me.Refresh
    
    If vg_codigo = "" Then
       
       Exit Sub
    
    End If

Case 3
    
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = ""
    vg_codigo = ""
    
    If Trim(LimpiaDato(fpText1.text)) = "" Or Regimen.Value = "" Then
    
        Exit Sub
    
    End If
    
    B_MTaEst.LlenaDatos "Nutrientes", Me.vaSpread2, fpText1.text, "", Format(FpFecDesde.text, "yyyymmdd"), Format(FpFecHasta.text, "yyyymmdd"), "2", lc_Aux, 0, 2
    B_MTaEst.Show 1
    Me.Refresh
    
    If vg_codigo = "" Then
       
       Exit Sub
    
    End If
    
End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Regimen_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
    
If Val(Regimen.Value) < 1 Then
       
   fpayuda(1).Caption = ""
   Exit Sub
    
End If
    
Set RS = vg_db.Execute("sgp_Sel_RegimenxCodigo " & Regimen.Value & "")

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
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Regimen_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
    
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

    I_AporteNutricionalSansis Me
    
    Toolbar1.Enabled = True
    Frame1(0).Enabled = True

Case 3
    
    Set RS = vg_db.Execute("sgp_Sel_Clientes 1, '" & Trim(LimpiaDato(fpText1.text)) & "'")
    If RS.EOF Then
       
       RS.Close
       Set RS = Nothing
       MsgBox "No existe ceco planificado", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    vg_codigo = ""
    B_HistPm.LlenarHistPlan "Histórico Minuta", Trim(LimpiaDato(fpText1.text)), 2, 1
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    Regimen.Value = vg_codregimen
    FpFecDesde.text = dBoM("01/" & vg_fecha)
    FpFecHasta.text = dEoM("27/" & vg_fecha)
    
Case 5
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub MoverDatosVector()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If LimpiaDato(Trim(fpText1.text)) = "" Or FpFecDesde.text = "" Or FpFecHasta.text = "" Or Val(Regimen.Value) = 0 Then

    Exit Sub

End If

fg_carga ""
Set RS = vg_db.Execute("sgp_Sel_ServicioMinuta '" & LimpiaDato(Trim(fpText1.text)) & "', " & Regimen.Value & ", " & Format(FpFecDesde.text, "yyyymmdd") & ", " & Format(FpFecHasta.text, "yyyymmdd") & "")
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
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub ActivarGrillaServicio()

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
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub ActivarGrillaAportes()

On Error GoTo Man_Error

Dim iselecc As Integer
Dim i       As Long
iselecc = 0

For i = 1 To vaSpread2(0).MaxRows
    
    vaSpread2(0).Row = i: vaSpread2(0).Col = 1
    
    If vaSpread2(0).text = "1" Then iselecc = 1: Exit For

Next i

If iselecc = 0 Then
   
   For i = 1 To vaSpread2(0).MaxRows
       
       vaSpread2(0).Row = i
       vaSpread2(0).Col = 1
       vaSpread2(0).CellType = 10
       vaSpread2(0).TypeCheckText = ""
       vaSpread2(0).TypeCheckCenter = True
       vaSpread2(0).text = "1" ' checked
   
   Next i

End If

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

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

If Option2(0).Value = True Then ActivarGrillaServicio
If Option2(2).Value = True Then ActivarGrillaAportes
    
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
   
   MsgBox "Seleccione Opción Dentro Grilla", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function
   
End If
    
'-------> Validar aportes nutricionales
seleccion = 0
    
For i = 1 To vaSpread2(0).MaxRows
        
    vaSpread2(0).Row = i
    vaSpread2(0).Col = 1
    
    If vaSpread2(0).text = "1" Then
       
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

