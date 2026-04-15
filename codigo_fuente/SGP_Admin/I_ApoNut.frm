VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form I_ApoNut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aporte Nutricional"
   ClientHeight    =   4635
   ClientLeft      =   4980
   ClientTop       =   2490
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3975
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   8175
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
         TabIndex        =   33
         Top             =   1800
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
            TabIndex        =   35
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
            TabIndex        =   34
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
         Left            =   4080
         TabIndex        =   30
         Top             =   3180
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
            Left            =   2520
            TabIndex        =   32
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
            TabIndex        =   31
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   255
         Left            =   4440
         TabIndex        =   22
         Top             =   1320
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
         SpreadDesigner  =   "I_ApoNut.frx":0000
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   3600
         TabIndex        =   21
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Aporte Nutricional"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
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
         TabIndex        =   20
         Top             =   3180
         Width           =   3810
         Begin VB.OptionButton Option2 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   9
            Top             =   340
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Lista"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   10
            Top             =   340
            Width           =   735
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   3120
            Picture         =   "I_ApoNut.frx":4466
            Top             =   210
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Servicio"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   2460
         Width           =   3810
         Begin VB.OptionButton Option2 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   340
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Lista"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   6
            Top             =   340
            Width           =   735
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   3180
            Picture         =   "I_ApoNut.frx":4770
            Top             =   210
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Opción Casino"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4080
         TabIndex        =   15
         Top             =   2460
         Width           =   3885
         Begin VB.OptionButton Option1 
            Caption         =   "Sin Código"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   340
            Value           =   -1  'True
            Width           =   1180
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Con Código"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2460
            TabIndex        =   8
            Top             =   340
            Width           =   1240
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   3240
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
         SpreadDesigner  =   "I_ApoNut.frx":4A7A
         StartingColNumber=   6
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   2
         Top             =   930
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
         TabIndex        =   1
         Top             =   585
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
         Index           =   0
         Left            =   1800
         TabIndex        =   3
         Top             =   1275
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
         BackColor       =   -2147483624
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   6765
         TabIndex        =   4
         Top             =   1275
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
         BackColor       =   -2147483624
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   0
         Top             =   240
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
         Index           =   2
         Left            =   3150
         TabIndex        =   28
         Top             =   930
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
         TabIndex        =   26
         Top             =   585
         Width           =   4845
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2700
         Picture         =   "I_ApoNut.frx":4EDE
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2700
         Picture         =   "I_ApoNut.frx":51E8
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2700
         Picture         =   "I_ApoNut.frx":54F2
         Top             =   150
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha (dd/mm/aaaa)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   5055
         TabIndex        =   18
         Top             =   1335
         Width           =   1605
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha (dd/mm/aaaa)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   1335
         Width           =   1605
      End
      Begin VB.Label Label2 
         Caption         =   "Punto Venta"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   975
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Casino"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   14
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Segmento"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   13
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
         Left            =   3150
         TabIndex        =   24
         Top             =   240
         Width           =   4845
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3195
         TabIndex        =   25
         Top             =   285
         Width           =   4845
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3195
         TabIndex        =   27
         Top             =   615
         Width           =   4845
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3195
         TabIndex        =   29
         Top             =   975
         Width           =   4845
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_ApoNut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim iaporte As Integer, i As Integer, imarca As Integer, iselecc As Integer, Msgtitulo As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
On Error GoTo Man_Error
Me.Height = 5145
Me.Width = 8475
Me.HelpContextID = vg_OpcM
fg_centra Me
fg_carga ""
Msgtitulo = "Aporte Nutricional"
Toolbar1.ImageList = partida.IL1
Toolbar1.Buttons.Clear
'Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar ": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set btnX = Toolbar1.Buttons.Add(, "Vista Previa", , tbrDefault, "Vista Previa"): btnX.Visible = True: btnX.ToolTipText = "Vista Previa ": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): btnX.Visible = True: btnX.ToolTipText = "Historico Minutas"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
vaSpread1.MaxRows = 0: vaSpread2.MaxRows = 0: imarca = 0
fpDateTime1(0).Text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).Text = Format(Date, "dd/mm/yyyy")

'------- Llenar Tabla Diéteticas
Set RS = vg_db.Execute("min_s_nutriente 10, 0, ''")
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      If RS!Ntrnt_Main_Ind > 0 Then
         vaSpread2.Col = 1
         vaSpread2.CellType = 10
         vaSpread2.TypeCheckText = ""
         vaSpread2.TypeCheckCenter = True
         vaSpread2.Text = "1" ' checked
      Else
         vaSpread2.Col = 1
         vaSpread2.CellType = 10
         vaSpread2.TypeCheckText = ""
         vaSpread2.TypeCheckCenter = True
         vaSpread2.Text = " " ' checked
      End If
      vaSpread2.Col = 2: vaSpread2.Value = RS!Ntrnt_Code
      vaSpread2.Col = 3: vaSpread2.Value = Trim(RS!Ntrnt_Name)
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing: fg_descarga
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub

Private Sub Form_Unload(Cancel As Integer)
'registrar Log sistema salir
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "")
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
MoverDatosVector
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Select Case Index
Case 0
    Set RS = vg_db.Execute("min_s_segmento 4, " & Val(fpLongInteger1(0).Value) & ", ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "":  Exit Sub
    fpayuda(0).Caption = Trim(RS!Unit_Dfnd_Desc)
    RS.Close: Set RS = Nothing
    fpText.Text = "": fpayuda(1).Caption = "": fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": vaSpread1.MaxRows = 0
Case 1
    Set RS = vg_db.Execute("min_s_puntoventa 7, " & Val(fpLongInteger1(1).Value) & ", ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "":  Exit Sub
    fpayuda(2).Caption = Trim(RS!Sls_Locn_Name)
    RS.Close: Set RS = Nothing
    MoverDatosVector
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    If Index = 0 Then image1_Click 0
    If Index = 1 Then image1_Click 2
End Select
End Sub

Private Sub fpText_Change()
Set RS = vg_db.Execute("min_s_casino 6, " & Val(fpLongInteger1(0).Value) & ", '" & "00000" & LimpiaDato(Trim(fpText.Text)) & "'")
If RS.EOF Then fpayuda(1).Caption = "": RS.Close: Set RS = Nothing: Exit Sub
fpayuda(1).Caption = Trim(RS!Nombre_Casino)
RS.Close: Set RS = Nothing
vaSpread1.MaxRows = 0: fpLongInteger1(1).Value = "": fpayuda(2).Caption = ""
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    image1_Click 1
End Select
End Sub

Private Sub Option2_Click(Index As Integer)
Select Case Index
Case 0
    Option2(0).Value = True
    Option2(1).Value = False
    Image1(3).Enabled = False
Case 1
    Option2(0).Value = False
    Option2(1).Value = True
    Image1(3).Enabled = True
Case 2
    Option2(2).Value = True
    Option2(3).Value = False
    Image1(4).Enabled = False
Case 3
    Option2(2).Value = False
    Option2(3).Value = True
    Image1(4).Enabled = True
End Select
End Sub

Private Sub image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_codigo = "": vg_left = fpayuda(0).Left
    B_TabEst.LlenaDatos "Segemento", 0, 2
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Then Exit Sub
    fpLongInteger1(0).Value = vg_codigo
    fpText.Text = "": fpayuda(1).Caption = "": fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": vaSpread1.MaxRows = 0
Case 1
    vg_codigo = ""
    vg_left = fpayuda(1).Left
    B_TabEst.LlenaDatos "Casino", Val(fpLongInteger1(0).Value), 1
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Then Exit Sub
    fpText.Text = vg_codigo
    vaSpread1.MaxRows = 0: fpLongInteger1(1).Value = "": fpayuda(2).Caption = ""
Case 2
    vg_codigo = ""
    vg_left = fpayuda(2).Left
    B_TabEst.LlenaDatos "Punto Venta", 0, 3
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Then Exit Sub
    fpLongInteger1(1).Value = vg_codigo
Case 3
    B_MServi.LlenaDatosSer Me, Trim(fpText.Text), Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Format(fpDateTime1(0).Text, "yyyy"), Format(fpDateTime1(0).Text, "mm")
    B_MServi.Show 1
    Me.Refresh
Case 4
    vg_opnutriente = 1
    B_MNutri.Show 1
    Me.Refresh
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim iselecc As Integer, i As Long
Select Case Button.Index
Case 1
    Set RS = vg_db.Execute("min_s_segmento 4, " & Val(fpLongInteger1(0).Value) & ", ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Segmento", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Set RS = vg_db.Execute("min_s_casino 6, " & Val(fpLongInteger1(0).Value) & ", '" & "00000" & LimpiaDato(Trim(fpText.Text)) & "'")
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Casino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Set RS = vg_db.Execute("min_s_puntoventa 7, " & Val(fpLongInteger1(1).Value) & ", ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe P. Venta", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If fpDateTime1(0).Text > fpDateTime1(1).Text Then MsgBox "Fecha Origen No Puede Ser Mayor Que Fecha Destino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Mid(fpDateTime1(0).Text, 4, 2) <> Mid(fpDateTime1(1).Text, 4, 2) Then MsgBox "Mes Origen Tiene Que Ser Igual Mes Destino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Mid(fpDateTime1(0).Text, 7, 4) <> Mid(fpDateTime1(1).Text, 7, 4) Then MsgBox "Año Origen Tiene Que Ser Igual Año Destino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If fpDateTime1(0).Text = "" Then Exit Sub
    If fpDateTime1(1).Text = "" Then Exit Sub
    If Option2(0).Value = True Then ActivarGrillaServicio
    If Option2(2).Value = True Then ActivarGrillaAportes
    iselecc = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        If vaSpread1.Text = "1" Then iselecc = 1: Exit For
    Next i
    If iselecc = 0 Then MsgBox "Seleccione Opción Dentro Grilla", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    iselecc = 0
    For i = 1 To vaSpread2.MaxRows
        vaSpread2.Row = i: vaSpread2.Col = 1
        If vaSpread2.Text = "1" Then iselecc = 1: Exit For
    Next i
    If iselecc = 0 Then MsgBox "Debe Seleccionar A lo Menos Un Aporte Nutricional", vbExclamation + vbOKOnly, "Aporte Nutricional": Exit Sub
    fg_carga ""
    Set RS = vg_db.Execute("min_s_minutas 3, 0, '" & "00000" & Trim(fpText.Text) & "', " & Val(fpLongInteger1(0).Value) & ", 0, " & Val(fpLongInteger1(1).Value) & ", 0, '" & Format(fpDateTime1(0).Text, "yyyy") & "', '" & Format(fpDateTime1(0).Text, "mm") & "'")
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No Existe Información", vbExclamation + vbOKOnly, "Aporte Nutricional": Exit Sub
    RS.Close: Set RS = Nothing: fg_descarga
    Toolbar1.Enabled = False
    Frame1(0).Enabled = False
    vg_opimp = 0
    'registrar Log sistema confirma
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Aceptar"), Me.HelpContextID, "", "")
    I_AporteNutricional Me
    Toolbar1.Enabled = True
    Frame1(0).Enabled = True
Case 3
    Set RS = vg_db.Execute("min_s_casino 6, " & Val(fpLongInteger1(0).Value) & ", '" & "00000" & LimpiaDato(Trim(fpText.Text)) & "'")
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Casino", vbExclamation + vbOKOnly, "Listado Menu Estudio": Exit Sub
    vg_codigo = ""
    B_HistPm.LlenarDatos "00000" & LimpiaDato(Trim(fpText.Text)), Val(fpLongInteger1(0).Value)
    B_HistPm.Show 1
    If Trim(vg_codigo) = "" Then Exit Sub
    Dim StrImp As String, StrImpb As String
    StrImp = Trim(vg_codigo): i = 1
    Do While InStr(StrImp, ";") <> 0
       If i = 1 Then
          fpLongInteger1(1).Value = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
       ElseIf i = 3 Then
          fpDateTime1(0).Text = "01" & "/" & Mid(StrImp, 1, InStr(StrImp, ";") - 1)
          fpDateTime1(1).Text = fg_mes(Mid(StrImp, 1, InStr(StrImp, ";") - 1)) & "/" & Mid(StrImp, 1, InStr(StrImp, ";") - 1)
       End If
       StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), ""): i = i + 1
    Loop
Case 5
    Me.Hide
    Unload Me
End Select
End Sub

Sub MoverDatosVector()
If LimpiaDato(Trim(fpText.Text)) = "" Or fpDateTime1(0).Text = "" Or fpDateTime1(1).Text = "" Or Val(fpLongInteger1(1).Value) = 0 Or Val(fpLongInteger1(1).Value) = 0 Then Exit Sub
fg_carga ""
Set RS = vg_db.Execute("min_s_minutas 3, 0,  '" & "00000" & LimpiaDato(Trim(fpText.Text)) & "', " & Val(fpLongInteger1(0).Value) & ", 0, " & Val(fpLongInteger1(1).Value) & ", 0, '" & Format(fpDateTime1(0).Text, "yyyy") & "', '" & Format(fpDateTime1(0).Text, "mm") & "'")
vaSpread1.MaxRows = 0
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 2
      vaSpread1.Value = RS!Serv_No
      vaSpread1.Col = 3
      vaSpread1.Value = Trim(RS!Serv_Name)
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing: fg_descarga
End Sub

Sub ActivarGrillaServicio()
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1
    vaSpread1.CellType = 10
    vaSpread1.TypeCheckText = ""
    vaSpread1.TypeCheckCenter = True
    vaSpread1.Text = "1" ' checked
Next i
End Sub

Sub ActivarGrillaAportes()
iselecc = 0
For i = 1 To vaSpread2.MaxRows
    vaSpread2.Row = i: vaSpread2.Col = 1
    If vaSpread2.Text = "1" Then iselecc = 1: Exit For
Next i
If iselecc = 0 Then
   For i = 1 To vaSpread2.MaxRows
       vaSpread2.Row = i
       vaSpread2.Col = 1
       vaSpread2.CellType = 10
       vaSpread2.TypeCheckText = ""
       vaSpread2.TypeCheckCenter = True
       vaSpread2.Text = "1" ' checked
   Next i
End If
End Sub
