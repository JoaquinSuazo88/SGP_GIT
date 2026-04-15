VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form C_ConIng 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumo Ingrediente"
   ClientHeight    =   8940
   ClientLeft      =   2295
   ClientTop       =   2490
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   120
      TabIndex        =   13
      Top             =   2650
      Width           =   9015
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5535
         Left            =   180
         TabIndex        =   7
         Top             =   240
         Width           =   8655
         _Version        =   393216
         _ExtentX        =   15266
         _ExtentY        =   9763
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
         MaxCols         =   5
         SpreadDesigner  =   "C_ConIng.frx":0000
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   2280
         TabIndex        =   15
         Top             =   5880
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Generando Consumo"
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
         Left            =   195
         TabIndex        =   20
         Top             =   5880
         Visible         =   0   'False
         Width           =   1770
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   9015
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   1
         Left            =   1755
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1750
         Width           =   1440
      End
      Begin VB.Frame Frame3 
         Height          =   675
         Left            =   3240
         TabIndex        =   16
         Top             =   1750
         Width           =   5295
         Begin VB.OptionButton Option1 
            Caption         =   "Resumido"
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
            Left            =   3480
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Detallado "
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
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1755
         TabIndex        =   4
         Top             =   2145
         Width           =   1425
         _Version        =   196608
         _ExtentX        =   2514
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
         Text            =   "05/2010"
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
         Left            =   1755
         TabIndex        =   0
         Top             =   435
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
         Index           =   1
         Left            =   1755
         TabIndex        =   1
         Top             =   770
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
         Left            =   1755
         TabIndex        =   3
         Top             =   1430
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
         Index           =   3
         Left            =   1755
         TabIndex        =   2
         Top             =   1100
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         Index           =   6
         Left            =   480
         TabIndex        =   27
         Top             =   1860
         Width           =   390
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
         Index           =   5
         Left            =   480
         TabIndex        =   25
         Top             =   1150
         Width           =   705
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2685
         Picture         =   "C_ConIng.frx":1906
         Top             =   1000
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   3240
         TabIndex        =   24
         Top             =   1100
         Width           =   5205
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
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
         Index           =   4
         Left            =   480
         TabIndex        =   22
         Top             =   1480
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2685
         Picture         =   "C_ConIng.frx":1C10
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3240
         TabIndex        =   21
         Top             =   1430
         Width           =   5205
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3240
         TabIndex        =   18
         Top             =   770
         Width           =   5205
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2685
         Picture         =   "C_ConIng.frx":1F1A
         Top             =   660
         Width           =   480
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
         Left            =   480
         TabIndex        =   17
         Top             =   820
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
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
         Left            =   480
         TabIndex        =   11
         Top             =   2205
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Subsegmento"
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
         Left            =   480
         TabIndex        =   10
         Top             =   495
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2685
         Picture         =   "C_ConIng.frx":2224
         Top             =   360
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3240
         TabIndex        =   9
         Top             =   435
         Width           =   5205
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3240
         TabIndex        =   12
         Top             =   480
         Width           =   5235
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3240
         TabIndex        =   19
         Top             =   810
         Width           =   5235
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   3240
         TabIndex        =   23
         Top             =   1470
         Width           =   5235
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   7
         Left            =   3240
         TabIndex        =   26
         Top             =   1140
         Width           =   5235
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8940
      Left            =   9285
      TabIndex        =   14
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   15769
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "C_ConIng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim Msgtitulo As String, texcol As String
Dim vecsub As Variant

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_centra Me
Me.Height = 8910
Me.Width = 9870
fg_centra Me
Me.HelpContextID = vg_OpcM
Msgtitulo = "Resumen Ingrediente"
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar ": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 2) = "00", False, True)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): btnX.Visible = True: btnX.ToolTipText = "Exporta Excel ": btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
fpDateTime1(0).text = Format(Date, "mm/yyyy")
vaSpread1.MaxRows = 0
vaSpread1.MaxCols = 0
vaSpread1.Visible = False

OpUsuario = vg_Indppr
If IsNull(OpUsuario) Or Trim(OpUsuario) = "" Then
    MsgBox "Contactese con el Administrador del Sistema...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
Else
    Select Case OpUsuario
    Case "1"
        Combo2(1).Clear
        Combo2(1).AddItem "Real" & Space(150) & "(1)"
        Combo2(1).ListIndex = fg_buscacbo(Combo2, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
        vg_IndpprSelec = fg_buscacbo(Combo2, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
    Case "2"
        Combo2(1).Clear
        Combo2(1).AddItem "Propuesta" & Space(150) & "(2)"
        Combo2(1).ListIndex = fg_buscacbo(Combo2, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
        vg_IndpprSelec = fg_buscacbo(Combo2, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
    Case "3"
        Combo2(1).Clear
        Combo2(1).AddItem "Real" & Space(150) & "(1)"
        Combo2(1).AddItem "Propuesta" & Space(150) & "(2)"
        Combo2(1).ListIndex = 0
        vg_IndpprSelec = 1
    End Select
End If

End Sub

Private Sub fpDateTime1_Change(Index As Integer)
vaSpread1.MaxRows = 0: Toolbar1.Buttons(4).Enabled = False
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    If Index = 0 Then Image1_Click 0
End Select
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Select Case Index
Case 0
    vaSpread1.MaxRows = 0: Toolbar1.Buttons(4).Enabled = False
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      Set RS = vg_db.Execute("sgpadm_s_subsegmento 10, " & Val(fpLongInteger1(0).Value) & ", '', '" & vg_Indppr & "'")
    Else
      Set RS = vg_db.Execute("sgpadm_s_subsegmento 1, " & Val(fpLongInteger1(0).Value) & ", '', ''")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing
Case 1
'    Set RS = vg_db.Execute("sgpadm_s_regimen 1, " & Val(fpLongInteger1(1).Value) & ",''")
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & " AND reg_indppr = '" & vg_Indppr & "'")
    Else
      Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & "")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
Case 2
    Set RS = vg_db.Execute("sgpadm_s_zona 1, " & Val(fpLongInteger1(2).Value) & ",''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(3).Caption = "": Exit Sub
    fpayuda(3).Caption = Trim(RS!Zon_nombre)
    RS.Close: Set RS = Nothing
Case 3
'    Set RS = vg_db.Execute("sgpadm_s_servicio 1, '', " & Val(fpLongInteger1(3).Value) & ",''")
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      Set RS = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(3).Value) & " AND ser_indppr = '" & vg_Indppr & "'")
    Else
      Set RS = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(3).Value) & "")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(5).Caption = "": Exit Sub
    fpayuda(5).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
End Select
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_subsegmento", "sub_", "Subsegmento", "Sub"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
    fpLongInteger1(1).SetFocus
Case 1
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    fpLongInteger1(3).SetFocus
Case 2
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_zona", "zon_", "Zona", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(3).Caption = vg_nombre
    fpDateTime1(0).SetFocus
Case 3
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Ser"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(3).Value = Val(vg_codigo)
    fpayuda(5).Caption = vg_nombre
    fpLongInteger1(2).SetFocus
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
vaSpread1.MaxRows = 0: Toolbar1.Buttons(4).Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 2
    Dim i As Long, j As Long, x As Long, z As Long, auxfec As Long, auxsub As Long, coding As String, nomsub As String, indtop As Long, esttop As Boolean, y As Long
'    If Trim(fpayuda(3).Caption) = "" Then MsgBox "No existe zona", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Set RS = vg_db.Execute("sgpadm_s_planifminuta 5, " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", 0, 0, " & Format(fpDateTime1(0).Value, "yyyymm") & ", 0, 0,'" & Val(fg_codigocbo(Combo2, 1, 1, "")) & "'")
    If RS.EOF Or RS!nReg = 0 Then RS.Close: Set RS = Nothing: MsgBox "No existe subsegmento, para este periodo.", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    vaSpread1.MaxCols = IIf(Option1(0).Value = True, 9, 8) + RS!nReg + IIf(Option1(0).Value = True, 2, 1)
    ReDim vecsub(RS!nReg, 4)
    RS.Close: Set RS = Nothing
    Set RS = vg_db.Execute("sgpadm_s_planifminuta 6, " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", 0, 0, " & Format(fpDateTime1(0).Value, "yyyymm") & ", 0, 0,'" & Val(fg_codigocbo(Combo2, 1, 1, "")) & "'")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(0).Value = "": fpayuda(0).Caption = "": MsgBox "No existe subsegmento, para este periodo.", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    fg_carga ""
    vaSpread1.Visible = False
    i = 65: x = 64: j = 1: texcol = "": z = 65
    Do Until j > vaSpread1.MaxCols
       If j = vaSpread1.MaxCols Then texcol = Trim(texcol & Chr(i))
       i = i + 1: j = j + 1: z = z + 1
       If i = 90 Then i = 65: x = x + 1: texcol = Chr(i)
    Loop
    If texcol = "A" Or texcol = "" And vaSpread1.MaxCols > 1 Then texcol = "": texcol = Chr(z)
    j = 1: i = IIf(Option1(0).Value = True, 10, 9): vaSpread1.Row = 0
    If Option1(0).Value = True Then
       vaSpread1.Col = 1
       vaSpread1.ColWidth(i) = 8
       vaSpread1.text = "Fecha"
       vaSpread1.ColHidden = False
    End If
    
    vaSpread1.Col = IIf(Option1(0).Value = True, 2, 1)
    vaSpread1.ColWidth(i) = 8
    vaSpread1.text = "Cód Ing."
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = IIf(Option1(0).Value = True, 3, 2)
    vaSpread1.ColWidth(i) = 20
    vaSpread1.text = "Ingrediente"
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = IIf(Option1(0).Value = True, 4, 3)
    vaSpread1.ColWidth(i) = 8
    vaSpread1.text = "Tipo Ingrediente"
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = IIf(Option1(0).Value = True, 5, 4)
    vaSpread1.ColWidth(i) = 4
    vaSpread1.text = "U.M."
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = IIf(Option1(0).Value = True, 6, 5)
    vaSpread1.ColWidth(i) = 15
    vaSpread1.text = "Familia 1"
    vaSpread1.ColHidden = False

    vaSpread1.Col = IIf(Option1(0).Value = True, 7, 6)
    vaSpread1.ColWidth(i) = 15
    vaSpread1.text = "Familia 2"
    vaSpread1.ColHidden = False

    vaSpread1.Col = IIf(Option1(0).Value = True, 8, 7)
    vaSpread1.ColWidth(i) = 15
    vaSpread1.text = "Familia 3"
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = IIf(Option1(0).Value = True, 9, 8)
    vaSpread1.ColWidth(i) = 15
    vaSpread1.text = "Familia 4"
    vaSpread1.ColHidden = False
    auxsub = 0
    Do While Not RS.EOF
       nomsub = ""
       If RS!sub_codigo <> auxsub Then
          auxsub = RS!sub_codigo
          nomsub = Trim(RS!sub_nombre)
       End If
       vaSpread1.Col = i
       vaSpread1.ColWidth(i) = 10
       vaSpread1.text = nomsub & " " & Trim(RS!reg_nombre)
       vaSpread1.ColHidden = False
       
       vecsub(j, 1) = RS!sub_codigo
       vecsub(j, 2) = RS!reg_codigo
       vecsub(j, 3) = i
       vecsub(j, 4) = 0
       RS.MoveNext: i = i + 1: j = j + 1
    Loop
    RS.Close: Set RS = Nothing
    
    vaSpread1.Col = IIf(Option1(0).Value = True, (vaSpread1.MaxCols - 1), vaSpread1.MaxCols)
    vaSpread1.ColWidth(i) = 10
    vaSpread1.text = IIf(Option1(0).Value = True, "Total Día", "Total Mes")
    vaSpread1.ColHidden = False
    
    If Option1(0).Value = True Then
       vaSpread1.Col = vaSpread1.MaxCols
       vaSpread1.ColWidth(i) = 12
       vaSpread1.text = "Total Mes"
       vaSpread1.ColHidden = False
    End If
    vaSpread1.Visible = Visible
    auxfec = 0: coding = ""
    If Option1(0).Value = True Then
       Set RS = vg_db.Execute("sgpadm_s_detconsumoingrser " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(3).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & Format(fpDateTime1(0).Value, "yyyymm") & "")
    Else
       Set RS = vg_db.Execute("sgpadm_s_resconsumoingrser " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(3).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & Format(fpDateTime1(0).Value, "yyyymm") & "")
    End If
    If RS.EOF Then fg_descarga: RS.Close: Set RS = Nothing: Toolbar1.Buttons(4).Enabled = False: vaSpread1.MaxRows = 0: vaSpread1.MaxCols = 0: Exit Sub
    vaSpread1.Visible = False
    vaSpread1.MaxRows = 0
    vaSpread1.MaxRows = RS!totreg
    Label2(3).Caption = "Generando Consumo": Label2(3).Visible = True: Bar1(0).Visible = True: Bar1(0).Value = 0: j = 1: y = 1
    Do While Not RS.EOF
       DoEvents
       If RS!min_fecmin <> auxfec Or Trim(RS!ing_codigo) <> coding Then
          Bar1(0).Value = Val((j / RS!totreg) * 100)
'          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
'          vaSpread1.Row = vaSpread1.MaxRows: j = vaSpread1.MaxRows
          vaSpread1.Row = y
          auxfec = RS!min_fecmin
          coding = Trim(RS!ing_codigo)
          x = x + 1
          If Option1(0).Value = True Then
             vaSpread1.Col = 1
             vaSpread1.ColWidth(1) = 8
             vaSpread1.text = Mid(RS!min_fecmin, 7, 2) & "/" & Mid(RS!min_fecmin, 5, 2) & "/" & Mid(RS!min_fecmin, 1, 4)
             vaSpread1.ColHidden = False
             vaSpread1.TypeHAlign = TypeHAlignCenter
             vaSpread1.BackColor = &HC0FFC0
          End If
          
          vaSpread1.Col = IIf(Option1(0).Value = True, 2, 1)
          vaSpread1.ColWidth(IIf(Option1(0).Value = True, 2, 1)) = 6
          vaSpread1.text = Val(RS!ing_codigo) 'Trim(RS!ing_codigo)
          vaSpread1.ColHidden = False
          vaSpread1.BackColor = &HC0FFC0
          
          vaSpread1.Col = IIf(Option1(0).Value = True, 3, 2)
          vaSpread1.ColWidth(IIf(Option1(0).Value = True, 3, 2)) = 20
          vaSpread1.text = Trim(RS!ing_nombre)
          vaSpread1.ColHidden = False
          vaSpread1.BackColor = &HC0FFC0
       
          vaSpread1.Col = IIf(Option1(0).Value = True, 4, 3)
          vaSpread1.ColWidth(IIf(Option1(0).Value = True, 4, 3)) = 8
          vaSpread1.text = IIf(Trim(RS!ing_indppr) = "1", "Real", "Propuesta")
          vaSpread1.ColHidden = False
          vaSpread1.TypeHAlign = TypeHAlignCenter
          vaSpread1.BackColor = &HC0FFC0
       
          vaSpread1.Col = IIf(Option1(0).Value = True, 5, 4)
          vaSpread1.ColWidth(IIf(Option1(0).Value = True, 5, 4)) = 4
          vaSpread1.text = Trim(RS!unm_nomcor)
          vaSpread1.ColHidden = False
          vaSpread1.TypeHAlign = TypeHAlignCenter
          vaSpread1.BackColor = &HC0FFC0
       
          vaSpread1.Col = IIf(Option1(0).Value = True, 6, 5)
          vaSpread1.ColWidth(IIf(Option1(0).Value = True, 6, 5)) = 15
          vaSpread1.text = IIf(IsNull(RS!tip_nombre4) Or Trim(RS!tip_nombre4) = "", Trim(RS!tip_nombre3), Trim(RS!tip_nombre4))
          vaSpread1.ColHidden = False
          vaSpread1.BackColor = &HC0FFC0
          
          vaSpread1.Col = IIf(Option1(0).Value = True, 7, 6)
          vaSpread1.ColWidth(IIf(Option1(0).Value = True, 7, 6)) = 15
          vaSpread1.text = IIf(IsNull(RS!tip_nombre4) Or Trim(RS!tip_nombre4) = "", Trim(RS!tip_nombre2), Trim(RS!tip_nombre3)) 'Trim(RS!tip_nombre2)
          vaSpread1.ColHidden = False
          vaSpread1.BackColor = &HC0FFC0
          
          vaSpread1.Col = IIf(Option1(0).Value = True, 8, 7)
          vaSpread1.ColWidth(IIf(Option1(0).Value = True, 6, 5)) = 15
          vaSpread1.text = IIf(IsNull(RS!tip_nombre4) Or Trim(RS!tip_nombre4) = "", Trim(RS!tip_nombre1), Trim(RS!tip_nombre2)) 'Trim(RS!tip_nombre3)
          vaSpread1.ColHidden = False
          vaSpread1.BackColor = &HC0FFC0
          
          vaSpread1.Col = IIf(Option1(0).Value = True, 9, 8)
          vaSpread1.ColWidth(IIf(Option1(0).Value = True, 6, 5)) = 15
          vaSpread1.text = IIf(IsNull(RS!tip_nombre4) Or Trim(RS!tip_nombre4) = "", "", Trim(RS!tip_nombre1)) 'Trim(RS!tip_nombre1)
          vaSpread1.ColHidden = False
          vaSpread1.BackColor = &HC0FFC0
          
          y = y + 1
       End If
       For i = 1 To UBound(vecsub)
           If RS!min_subseg = vecsub(i, 1) And RS!min_codreg = vecsub(i, 2) Then
              vaSpread1.Col = vecsub(i, 3)
              vaSpread1.TypeHAlign = TypeHAlignRight
              vaSpread1.text = Format(RS!red_canpro, fg_Pict(6, 2))
              vaSpread1.Col = IIf(Option1(0).Value = True, vaSpread1.MaxCols - 1, vaSpread1.MaxCols)
              vaSpread1.TypeHAlign = TypeHAlignRight
              vaSpread1.Value = Format((IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.Value) + RS!red_canpro), fg_Pict(6, 2))
              vecsub(i, 4) = (vecsub(i, 4) + RS!red_canpro)
              Exit For
           End If
       Next i
       RS.MoveNext: j = j + 1
    Loop
    RS.Close: Set RS = Nothing
    vaSpread1.MaxRows = y - 1
    Label2(3).Visible = False: Bar1(0).Visible = False
    indtop = 1: esttop = False
    If Option1(0).Value = True Then
       j = 1
       Label2(3).Caption = "Buscando Consumo": Label2(3).Visible = True: Bar1(0).Visible = True: Bar1(0).Value = 0
       Set RS = vg_db.Execute("sgpadm_s_resconsumoingrser " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(3).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & Format(fpDateTime1(0).Value, "yyyymm") & "")
       Do While Not RS.EOF
          DoEvents
          ret2 = 0
          For i = 0 To vaSpread1.MaxRows
              ret2 = vaSpread1.SearchCol(IIf(Option1(0).Value = True, 2, 1), i, vaSpread1.MaxRows, Val(RS!ing_codigo), 4)
              If ret2 > -1 Then
                 i = ret2 + 1
                 vaSpread1.Row = ret2
                 vaSpread1.Col = vaSpread1.MaxCols
                 vaSpread1.TypeHAlign = TypeHAlignRight
                 vaSpread1.Value = Format((IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.Value) + RS!red_canpro), fg_Pict(6, 2))
              Else
                 Exit For
              End If
          Next i
          Bar1(0).Value = Val((j / RS!totreg) * 100)
          RS.MoveNext: j = j + 1
       Loop
       RS.Close: Set RS = Nothing
    End If
    Label2(3).Visible = False: Bar1(0).Visible = False
    vaSpread1.Visible = True
    Toolbar1.Buttons(4).Enabled = True
    fg_descarga
Case 4
    Dim NashXl As Excel.Application
    Dim irow As Long, irow2 As Long
    fg_carga ""
    Set NashXl = CreateObject("excel.application")
    Set NashXl = New Excel.Application
    NashXl.SheetsInNewWorkbook = 1
    NashXl.Workbooks.Add
    NashXl.Range("A1").Select
    NashXl.ActiveCell.FormulaR1C1 = Label2(0).Caption & ": (" & IIf(Val(fpLongInteger1(0).Value) = 0, "Todos", Val(fpLongInteger1(0).Value)) & ") " & fpayuda(0).Caption
    NashXl.Range("A2").Select
    NashXl.ActiveCell.FormulaR1C1 = Label2(2).Caption & ": (" & IIf(Val(fpLongInteger1(1).Value) = 0, "Todos", Val(fpLongInteger1(1).Value)) & ") " & fpayuda(1).Caption
    NashXl.Range("A3").Select
    NashXl.ActiveCell.FormulaR1C1 = Label2(5).Caption & ": (" & IIf(Val(fpLongInteger1(3).Value) = 0, "Todos", Val(fpLongInteger1(3).Value)) & ") " & fpayuda(5).Caption
    NashXl.Range("A4").Select
    NashXl.ActiveCell.FormulaR1C1 = Label2(4).Caption & ": (" & IIf(Val(fpLongInteger1(2).Value) = 0, "Todos", Val(fpLongInteger1(2).Value)) & ") " & fpayuda(3).Caption
    NashXl.Range("A5").Select
    NashXl.ActiveCell.FormulaR1C1 = Label2(1).Caption & ": " & Format(fpDateTime1(0).Value, "mm/yyyy") & "  " & IIf(Option1(0).Value = True, "Detallado", "Resumido")
    vaSpread1.AllowMultiBlocks = True
    vaSpread1.SetSelection 1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows + 3
    vaSpread1.ClipboardCopy
    irow = vaSpread1.MaxRows + 6
    '------- Pegar vaspread1(1) - Planilla Excel
    NashXl.Range("A6").Select
    NashXl.ActiveSheet.Paste
    
    NashXl.Range("A6:" & texcol & "6").Select
    With NashXl.Selection.Interior
         .ColorIndex = 15
         .Pattern = xlSolid
    End With
    NashXl.Columns("A:A").Select
    NashXl.Selection.NumberFormat = "mm/dd/yyyy"
    '------- Dibujar marco
    NashXl.Range("A6:" & texcol & "" & irow).Select
    NashXl.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    NashXl.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With NashXl.Selection.Borders(xlEdgeLeft)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
    End With
    With NashXl.Selection.Borders(xlEdgeTop)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
    End With
    With NashXl.Selection.Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
    End With
    With NashXl.Selection.Borders(xlEdgeRight)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
    End With
    With NashXl.Selection.Borders(xlEdgeRight) '(xlInsideVertical)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
    End With
    With NashXl.Selection.Borders(xlInsideHorizontal)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
    End With
    NashXl.Range(IIf(Option1(0).Value = True, "G6", "F6") & ":" & texcol & "" & irow).Select
    NashXl.Selection.NumberFormat = "#,##0.00"
    
'    NashXl.Selection.Font.Bold = True
'    NashXl.Range("B" & irow & ":" & "B" & 2).Select
'    NashXl.Selection.NumberFormat = "#,##0.00"
    '------- Ajustar columna
    NashXl.Cells.Select
    NashXl.Cells.EntireColumn.AutoFit
    vaSpread1.AllowMultiBlocks = False: vaSpread1.SetSelection 1, 0, vaSpread1.MaxCols, vaSpread1.MaxRows
    fg_descarga
    NashXl.Visible = True
Case 6
    Me.Hide
    Unload Me
End Select
End Sub

