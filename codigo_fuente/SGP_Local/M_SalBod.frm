VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_SalBod 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salida de Bodega para Producción"
   ClientHeight    =   7440
   ClientLeft      =   1995
   ClientTop       =   2280
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   345
      Left            =   9330
      TabIndex        =   32
      Top             =   1260
      Width           =   2355
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Oculta Ingrediente"
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
         Index           =   0
         Left            =   150
         TabIndex        =   33
         Top             =   120
         Width           =   2010
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   0
      TabIndex        =   12
      Top             =   315
      Width           =   11820
      Begin VB.OptionButton Option1 
         Caption         =   "Sector"
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
         Left            =   10800
         TabIndex        =   35
         Top             =   720
         Width           =   975
      End
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
         Index           =   0
         Left            =   9480
         TabIndex        =   34
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Height          =   45
         Left            =   30
         TabIndex        =   30
         Top             =   2265
         Visible         =   0   'False
         Width           =   8520
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   6075
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   200
         Width           =   2325
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   1350
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   9270
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   200
         Width           =   1155
         _Version        =   196608
         _ExtentX        =   2037
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   2
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "-2147483647"
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1395
         TabIndex        =   7
         Top             =   200
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
         AllowNull       =   -1  'True
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
         Text            =   ""
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
         Left            =   3915
         TabIndex        =   1
         Top             =   200
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
         AllowNull       =   -1  'True
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
         Text            =   ""
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
      Begin EditLib.fpDoubleSingle fpDouble1 
         Height          =   315
         Index           =   0
         Left            =   2070
         TabIndex        =   8
         Top             =   2400
         Visible         =   0   'False
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         DecimalPlaces   =   2
         DecimalPoint    =   "."
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
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
      Begin EditLib.fpDoubleSingle fpDouble1 
         Height          =   315
         Index           =   1
         Left            =   5385
         TabIndex        =   9
         Top             =   2400
         Visible         =   0   'False
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         DecimalPlaces   =   2
         DecimalPoint    =   "."
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1395
         TabIndex        =   4
         Top             =   630
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
         Left            =   1395
         TabIndex        =   5
         Top             =   990
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
      Begin VB.Label Label3 
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
         Index           =   2
         Left            =   75
         TabIndex        =   40
         Top             =   1060
         Width           =   705
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2280
         Picture         =   "M_SalBod.frx":0000
         Top             =   885
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   2760
         TabIndex        =   39
         Top             =   990
         Width           =   5415
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2760
         TabIndex        =   37
         Top             =   630
         Width           =   5415
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2280
         Picture         =   "M_SalBod.frx":030A
         Top             =   525
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Raciones del Personal"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Index           =   1
         Left            =   4065
         TabIndex        =   29
         Top             =   2310
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Raciones Facturables"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   0
         Left            =   435
         TabIndex        =   28
         Top             =   2310
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   10530
         TabIndex        =   22
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bodega"
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
         Index           =   8
         Left            =   5355
         TabIndex        =   19
         Top             =   285
         Width           =   660
      End
      Begin VB.Label Label3 
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
         Index           =   6
         Left            =   75
         TabIndex        =   18
         Top             =   1515
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2715
         Picture         =   "M_SalBod.frx":0614
         Top             =   1350
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº Doc."
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
         Left            =   8550
         TabIndex        =   17
         Top             =   285
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Emisión"
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
         Left            =   75
         TabIndex        =   0
         Top             =   285
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Prod."
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
         Index           =   9
         Left            =   2835
         TabIndex        =   16
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Régimen"
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
         Index           =   10
         Left            =   75
         TabIndex        =   15
         Top             =   720
         Width           =   750
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3165
         TabIndex        =   13
         Top             =   1455
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3210
         TabIndex        =   14
         Top             =   1485
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   6135
         TabIndex        =   20
         Top             =   250
         Width           =   2310
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2805
         TabIndex        =   38
         Top             =   675
         Width           =   5415
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   2805
         TabIndex        =   41
         Top             =   1035
         Width           =   5415
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5070
      Left            =   180
      TabIndex        =   27
      Top             =   1725
      Width           =   11460
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   1455
         Left            =   0
         TabIndex        =   36
         Top             =   195
         Width           =   11385
         _Version        =   393216
         _ExtentX        =   20082
         _ExtentY        =   2566
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
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
         MaxRows         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "M_SalBod.frx":091E
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3165
         Left            =   0
         TabIndex        =   10
         Top             =   1755
         Width           =   11385
         _Version        =   393216
         _ExtentX        =   20082
         _ExtentY        =   5583
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         MaxCols         =   10
         MaxRows         =   20
         ProcessTab      =   -1  'True
         SelectBlockOptions=   0
         SpreadDesigner  =   "M_SalBod.frx":0C3D
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   180
      TabIndex        =   23
      Top             =   6750
      Width           =   11460
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   720
         Left            =   1230
         TabIndex        =   11
         Top             =   180
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1270
         ButtonWidth     =   2302
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agr. Prod."
               Description     =   "Agregar Productos"
               Object.ToolTipText     =   "Agregar Producto"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Elim. Prod. "
               Description     =   "Eliminar Producto "
               Object.ToolTipText     =   "Eliminar Producto "
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   225
         Top             =   90
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_SalBod.frx":13B3
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_SalBod.frx":16CD
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ingrediente"
         Height          =   195
         Index           =   3
         Left            =   5295
         TabIndex        =   31
         Top             =   255
         Width           =   795
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   4905
         Top             =   285
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   2295
         Top             =   285
         Width           =   300
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000013&
         Caption         =   "Igresados por el Usuario"
         Height          =   450
         Index           =   0
         Left            =   2685
         TabIndex        =   26
         Top             =   165
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   8205
         Top             =   285
         Width           =   300
      End
      Begin VB.Label Label5 
         Caption         =   "Producto"
         Height          =   210
         Index           =   2
         Left            =   8565
         TabIndex        =   25
         Top             =   255
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "Sobrepasa Stock actual"
         Height          =   450
         Index           =   1
         Left            =   6930
         TabIndex        =   24
         Top             =   165
         Width           =   900
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008484FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   6525
         Top             =   285
         Width           =   300
      End
   End
End
Attribute VB_Name = "M_SalBod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset
Dim RS5 As New ADODB.Recordset
Dim RS6 As New ADODB.Recordset
Dim RS7 As New ADODB.Recordset
Dim est As Boolean, est1 As Boolean, modo As String
'Dim btnX As Button

Private Sub Check1_Click(Index As Integer)
Dim codsec As String
On Error GoTo Man_Error
'-------> Actualizar parametro salida producción
vg_db.BeginTrans
vg_db.Execute "UPDATE a_param SET par_valor = '" & IIf(Check1(0).Value = 1, 1, 0) & "' WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'ingsalpro'"
vg_db.CommitTrans
If vaSpread1.MaxRows < 1 And Option1(1).Value = True Then Exit Sub
If Frame2.Enabled = True And vaSpread1.MaxRows > 1 Then vaSpread1_EditMode vaSpread1.ActiveCol, vaSpread1.ActiveRow, 0, True
vaSpread1.Visible = False
codsec = 0
vaSpread2.Row = vaSpread2.ActiveRow
vaSpread2.Col = 1: codsec = vaSpread2.text
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 10
    If codsec = vaSpread1.text And Option1(1).Value = True Then
       vaSpread1.Col = 5
       If Trim(vaSpread1.text) = "" Then vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
    ElseIf Option1(0).Value = True Then
       vaSpread1.Col = 5
       If Trim(vaSpread1.text) = "" Then vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
    End If
Next i
vaSpread1.SetActiveCell 1, 1
vaSpread1.Visible = True
Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Sub MostrarDetalle()

Dim feprod As Long, codser As Long, codreg As Long, codbod As Long, numdia As Long, canreal As Double, Cantidad As Double, aAp As String, aAp1 As String
Dim codsec As Long, sql1 As String
Dim canrea As Double, canbod As Double, color As Variant, estado As String, codpro As String

If Trim(fpDateTime1(1).text) = "" Or (Trim(fpayuda(0).Caption) = "" And Not vg_tipser) Or (Trim(fpayuda(3).Caption) = "" And Not vg_tipser) Then Exit Sub
'-------> Validar si el contrato tiene asignado inventario rotativo
If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
   Format(fpDateTime1(1).Value, "dd/mm/yyyy") > CDate(vg_ciedia) Then MsgBox "Tiene que realizar cierre diario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
Label1.Caption = ""
feprod = Val(fpDateTime1(1).Year & Right("0" & fpDateTime1(1).Month, 2) & Right("0" & fpDateTime1(1).Day, 2))
If CierrePeriodo(feprod, vg_codbod, 6) Then
'   MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, Msgtitulo
   Exit Sub
End If
codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
'If Not vg_tipser And (Combo1(0).ListIndex = -1 Or Combo1(0).text = "") Then fg_descarga: Exit Sub
If vg_tipser Then fg_descarga: Exit Sub
If Combo1(1).ListIndex = -1 Or Combo1(1).text = "" Then est = True: Combo1(1).ListIndex = 0: est = False
codreg = 0: codser = 0
If Not vg_tipser Then codreg = Val(fpLongInteger1(1).Value) 'Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 22, 10))
If Not vg_tipser Then codser = Val(fpLongInteger1(2).Value) 'Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 10, 10))
'-------> Chequear si existe parametro en planificación o bien estructura fija
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT DISTINCT a.min_cencos " & _
         "FROM b_minuta a, b_minutadet b " & _
         "WHERE a.min_codigo = b.mid_codigo " & _
         "AND   b.mid_tipmin = '2' " & _
         "AND   a.min_fecmin = " & feprod & " " & _
         "AND   a.min_cencos = '" & Trim(fpText1(1).text) & "' " & _
         "AND   a.min_codreg = " & codreg & " " & _
         "AND   a.min_codser = " & codser & "", vg_db, adOpenStatic
If RS1.EOF Then
   RS1.Close: Set RS1 = Nothing
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   RS1.Open "SELECT DISTINCT mfd_cencos " & _
            "FROM b_minutafijadia " & _
            "WHERE mfd_cencos = '" & Trim(fpText1(1).text) & "' " & _
            "AND   mfd_codreg = " & codreg & " " & _
            "AND   mfd_codser = " & codser & " " & _
            "AND   mfd_fecha  = " & feprod & " AND mfd_tipmin = '2'", vg_db, adOpenStatic
   If RS1.EOF Then: RS1.Close: Set RS1 = Nothing: vaSpread2.MaxRows = 0: vaSpread1.MaxRows = 0: Exit Sub

End If
RS1.Close: Set RS1 = Nothing
sql1 = IIf(vg_tipbase = "1", " CDate('" & fpDateTime1(1).text & "') ", " '" & Format(fpDateTime1(1).text, "yyyymmdd") & "' ")

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT COUNT(tov_fecpro) AS suma FROM b_totventas WHERE " & _
         "tov_rutcli = '" & Trim(LimpiaDato(fpText1(1).text)) & "' AND tov_tipdoc='SP' " & _
         "AND tov_fecpro = " & sql1 & " " & _
         "AND tov_codreg = " & codreg & " AND tov_codser = " & codser & " AND (tov_estdoc <> 'A' AND tov_estdoc <> 'P') AND tov_codbod = " & vg_codbod & "", vg_db, adOpenStatic

Do While Not RS1.EOF
   
   If RS1!Suma > 0 Then
      
      modo = "A"
      RS1.Close: Set RS1 = Nothing
      Gl_Ac_Botones Me, 12, 6, ""
      fpLongInteger1(0).text = TraerCorrelativo(vg_codbod, "SP") 'MuestraFolio(Trim(fpText1(1).text))
      vaSpread1.MaxRows = 0
      Frame2.Enabled = True
      If vaSpread1.Enabled = True Then vaSpread1.SetFocus
      Frame1.Enabled = False
      If Option1(0).Value = True Then
         
         Toolbar2_ButtonClick Toolbar2.Buttons.Item(1)
      
      ElseIf Option1(1).Value = True Then
         
         vaSpread2.Visible = False
         vaSpread2.MaxRows = 0
         
         If RS1.State = 1 Then RS1.Close
         RS1.CursorLocation = adUseClient
         vg_db.CursorLocation = adUseClient
         
         RS1.Open "SELECT DISTINCT a.* FROM a_sector a, a_estservicio b WHERE a.sec_codigo = b.ess_codsec AND b.ess_codser = " & codser & " AND b.ess_cencos = '" & Trim(LimpiaDato(fpText1(1).text)) & "' ORDER BY a.sec_orden", vg_db, adOpenStatic
         If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: Me.MousePointer = 0: MsgBox "No existe relación estructura servicio & sector, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
         Do While Not RS1.EOF
            vaSpread2.MaxRows = vaSpread2.MaxRows + 1
            vaSpread2.Row = vaSpread2.MaxRows
            vaSpread2.Col = 1: vaSpread2.text = RS1!sec_codigo
            vaSpread2.Col = 2: vaSpread2.text = Trim(RS1!sec_nombre)
            RS1.MoveNext
         Loop
         vaSpread2.MaxRows = vaSpread2.MaxRows + 1
         vaSpread2.Row = vaSpread2.MaxRows
         vaSpread2.Col = 1: vaSpread2.Value = "estfij"
         vaSpread2.Col = 2: vaSpread2.Value = "Estructura Fija"
         vaSpread2.Visible = True
         RS1.Close: Set RS1 = Nothing
         Me.MousePointer = 0: fg_descarga
      
      End If
      Exit Sub
   
   Else
      
      RS1.Close: Set RS1 = Nothing
      modo = "M"
      sql1 = IIf(vg_tipbase = "1", " CDate('" & fpDateTime1(1).text & "') ", " '" & Format(fpDateTime1(1).text, "yyyymmdd") & "' ")
      
      If RS1.State = 1 Then RS1.Close
      RS1.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      RS1.Open "SELECT tov_estdoc FROM b_totventas WHERE " & _
               "tov_rutcli     = '" & Trim(LimpiaDato(fpText1(1).text)) & "' AND tov_tipdoc = 'SP' " & _
               "AND tov_fecpro = " & sql1 & " " & _
               "AND tov_codreg = " & codreg & " AND tov_codser = " & codser & " AND tov_estdoc = 'P' AND tov_codbod = " & vg_codbod & "", vg_db, adOpenStatic
      If Not RS1.EOF Then
         
         RS1.Close: Set RS1 = Nothing
         CargarSalidaPendientes
         modo = "M"
         Exit Sub
      
      Else
         
         Exit Do
      
      End If
   
   End If
   RS1.MoveNext

Loop
Me.MousePointer = 11
RS1.Close: Set RS1 = Nothing
vaSpread2.MaxRows = 0
Gl_Ac_Botones Me, 12, 6, ""
fpLongInteger1(0).text = TraerCorrelativo(vg_codbod, "SP")
numdia = fg_NumDia(Trim(Left(fg_Fecha_Dia(Trim(Str(feprod)), 2), Len(fg_Fecha_Dia(Trim(Str(feprod)), 2)) - 2)))
'-------> Validar productos vigente salida producción
ValidarProductoVigente

If Option1(1).Value = True Then
   '-------> Validar si la estructura de servicio no tiene asignado un sector
   Dim GlosaEstructura As String
   GlosaEstructura = ""
  
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   RS1.Open "SELECT DISTINCT d.ess_codigo, d.ess_nombre " & _
            "FROM b_minuta a, b_minutadet b, a_servicio c, a_estservicio d " & _
            "WHERE a.min_codigo = b.mid_codigo " & _
            "AND   a.min_codser = c.ser_codigo " & _
            "AND   a.min_codser = " & codser & " " & _
            "AND   b.mid_estser = d.ess_codigo " & _
            "AND   a.min_cencos = d.ess_cencos AND c.ser_codigo = d.ess_codser and (d.ess_codsec IS NULL OR d.ess_codsec = 0) " & _
            "AND   a.min_cencos = '" & Trim(fpText1(1).text) & "' AND a.min_codreg = " & codreg & " " & _
            "AND   a.min_fecmin = " & feprod & " AND b.mid_tipmin = '2' AND b.mid_numrac > 0 " & _
            "", vg_db, adOpenStatic
   If Not RS1.EOF Then
      
      Do While Not RS1.EOF
         
         GlosaEstructura = Trim(GlosaEstructura) & RS1!ess_codigo & " - " & Trim(RS1!ess_nombre) & VgLinea
         RS1.MoveNext
      
      Loop
      
      RS1.Close
      Set RS1 = Nothing
      Me.MousePointer = 0
      fg_descarga
      MsgBox "Una de las estructuras de servicio no tiene asignado sector: " & VgLinea & VgLinea & GlosaEstructura & VgLinea & "Asigne la sector ...", vbCritical + vbOKOnly, MsgTitulo
      Exit Sub
   
   End If
   
   RS1.Close: Set RS1 = Nothing

End If
'------------------------------------ MINUTA ---------------------------------
'-------> Creo tabla temporal y chequeo si existe antes
aAp = Trim(vg_NUsr) & "_tmp_SalBod"
fg_CheckTmp aAp

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Option1(0).Value = True Then
   
   RS1.Open "SELECT red.red_codpro, 0 AS sec_codigo, 0 AS sec_orden, '' AS sec_nombre, SUM(isnull(mid.mid_numrac,0)*(isnull(red.red_canpro,0)/isnull(rec.rec_basrac,0))) AS cantidad INTO " & aAp & " " & _
            "FROM b_minuta mi, b_minutadet mid, b_receta rec, b_recetadet red " & _
            "WHERE rec.rec_codigo = mid.mid_codrec " & _
            "AND rec.rec_codigo = red.red_codigo AND red.red_tiprec = mid.mid_tiprec AND ((red.red_tiprec <> 0 AND red.red_cencos = '" & MuestraCasino(1) & "') OR (red.red_tiprec = 0 AND red.red_cencos = '0')) AND mi.min_codigo = mid.mid_codigo " & _
            "AND mi.min_fecmin = " & feprod & " AND mid.mid_tipmin = '2' AND mi.min_cencos = '" & Trim(fpText1(1).text) & "' " & _
            "AND mi.min_codreg = " & codreg & " AND mi.min_codser = " & codser & " " & _
            "GROUP BY red.red_codpro", vg_db, adOpenStatic

Else
   
   RS1.Open "SELECT red.red_codpro, sec.sec_codigo, sec.sec_orden, sec.sec_nombre, SUM(isnull(mid.mid_numrac,0)*(isnull(red.red_canpro,0)/isnull(rec.rec_basrac,0))) AS cantidad INTO " & aAp & " " & _
            "FROM b_minuta mi, b_minutadet mid, b_receta rec, b_recetadet red, a_servicio ser, a_estservicio ess, a_sector sec " & _
            "WHERE rec.rec_codigo = mid.mid_codrec " & _
            "AND rec.rec_codigo = red.red_codigo AND red.red_tiprec = mid.mid_tiprec AND ((red.red_tiprec <> 0 AND red.red_cencos = '" & MuestraCasino(1) & "') OR (red.red_tiprec = 0 AND red.red_cencos = '0')) AND mi.min_codigo = mid.mid_codigo " & _
            "AND mi.min_codser = ser.ser_codigo AND mi.min_codser = " & codser & " AND mid.mid_estser = ess.ess_codigo AND mi.min_cencos = ess.ess_cencos AND ess.ess_codser = ser.ser_codigo and ess.ess_codsec = sec.sec_codigo " & _
            "AND mi.min_cencos = '" & Trim(fpText1(1).text) & "' AND mi.min_codreg = " & codreg & " AND mi.min_fecmin = " & feprod & " AND mid.mid_tipmin = '2'  " & _
            "GROUP BY red.red_codpro, sec.sec_codigo, sec.sec_orden, sec.sec_nombre ORDER BY sec.sec_orden", vg_db, adOpenStatic

End If
Set RS1 = Nothing

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'-------> Creo en cero los productos que no existen en bodega ---------------------
RS1.Open "SELECT isnull(a.cpi_codped,'') AS codigo FROM " & aAp & " aux, b_contlistpreing a, b_productosing b WHERE a.cpi_coding = b.pri_coding AND a.cpi_cencos = '" & MuestraCasino(1) & "' AND aux.red_codpro = b.pri_coding " & _
         "UNION " & _
         "SELECT a.mfd_codpro AS codigo  " & _
         "FROM b_minutafijadia a, b_productos b " & _
         "WHERE a.mfd_codpro = b.pro_codigo " & _
         "AND   a.mfd_cencos = '" & Trim(fpText1(1).text) & "' " & _
         "AND   a.mfd_codreg = " & codreg & " " & _
         "AND   a.mfd_codser = " & codser & " " & _
         "AND   a.mfd_fecha = " & feprod & " AND a.mfd_tipmin = '2'", vg_db, adOpenStatic
Do While Not RS1.EOF
   ValidaBod codbod, RS1!codigo
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
'-----------------------------------------------------------------------------
If vg_tipbase = "1" Then
   '-------> Insert tabla productospmpdia
   aAp1 = Trim(vg_NUsr) & "_tmp_ProductoPMPSalBodMosDet"
   fg_CheckTmp aAp1
   vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                 "INTO " & aAp1 & " " & _
                 "FROM b_productospmpdia " & _
                 "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(fpDateTime1(1).text), "yyyymmdd") & " " & _
                 "AND   ppd_propon > 0 " & _
                 "GROUP BY ppd_cencos, ppd_codpro ORDER BY Max(ppd_fecdia) DESC"
   vg_db.Execute "ALTER TABLE " & aAp1 & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
   vg_db.Execute "UPDATE " & aAp1 & " INNER JOIN b_productospmpdia ON (" & aAp1 & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp1 & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp1 & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp1 & ".ppd_propon = b_productospmpdia.ppd_propon"
'   vg_db.Execute "INSERT INTO " & aAp1 & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp1 & ")"
   vg_db.Execute "INSERT INTO " & aAp1 & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(fpDateTime1(1).text) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp1 & ")"
End If

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

sqlparam = IIf(Option1(0).Value = True, "pro.pro_codigo", "aux.sec_orden")
If vg_tipbase = "1" Then
   
   RS1.Open "SELECT   pro.pro_codigo, a.ppd_propon, pro.pro_nombre, uni.uni_nomcor, ing.ing_codigo, ing.ing_nombre, " & _
            "         round(bod.bod_canmer, " & vg_DCa & ") as bod_canmer, unm.unm_nomcor, aux.sec_codigo, aux.sec_orden, aux.sec_nombre, SUM(aux.cantidad/pro.pro_facing) As cantidad, Sum(aux.cantidad) As cantidad2 " & _
            "FROM     b_productos pro, b_ingrediente ing, b_productosing pri, a_unidad uni, " & _
            "         b_bodegas bod, a_unidadmed unm, " & aAp & " aux, " & aAp1 & " a, b_contlistpreing b " & _
            "WHERE    pri.pri_coding = aux.red_codpro " & _
            "AND      ing.ing_codigo = pri.pri_coding " & _
            "AND      pro.pro_codigo = b.cpi_codped AND ing.ing_codigo = b.cpi_coding AND b.cpi_cencos = '" & MuestraCasino(1) & "' AND pro.pro_codigo = a.ppd_codpro AND a.ppd_cencos = '" & MuestraCasino(1) & "' " & _
            "AND      pri.pri_codpro = b.cpi_codped " & _
            "AND      ing.ing_unimed = unm.unm_codigo " & _
            "AND      pro.pro_coduni = uni.uni_codigo " & _
            "AND      bod.bod_codpro = pro.pro_codigo " & _
            "AND     (pro.pro_fecven > " & Format(Date, "yyyymmdd") & " OR pro.pro_fecven <= 0 OR bod.bod_canmer > 0) " & _
            "AND      bod.bod_codbod = " & vg_codbod & " " & _
            "AND      pro.pro_ctrsto = 1 " & _
            "GROUP BY pro.pro_codigo, a.ppd_propon, pro.pro_nombre, uni.uni_nomcor, ing.ing_codigo, " & _
            "         ing.ing_nombre, bod.bod_canmer, unm.unm_nomcor, aux.sec_codigo, aux.sec_orden, aux.sec_nombre " & _
            "ORDER BY " & sqlparam & "", vg_db, adOpenStatic

Else
   
   RS1.Open "SELECT   pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, ing.ing_codigo, ing.ing_nombre, " & _
            "         round(bod.bod_canmer," & vg_DCa & ") as bod_canmer, unm.unm_nomcor, aux.sec_codigo, aux.sec_orden, aux.sec_nombre, SUM(aux.cantidad/pro.pro_facing) As cantidad, Sum(aux.cantidad) As cantidad2, " & _
            "        (SELECT TOP 1 ppd_propon FROM b_productospmpdia WHERE ppd_codpro = pro.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(fpDateTime1(1).text), "yyyymmdd") & " ORDER BY ppd_fecdia DESC) AS ppd_propon " & _
            "FROM     b_productos pro, b_ingrediente ing, b_productosing pri, a_unidad uni, " & _
            "         b_bodegas bod, a_unidadmed unm, " & aAp & " aux, b_contlistpreing b " & _
            "WHERE    pri.pri_coding = aux.red_codpro " & _
            "AND      ing.ing_codigo = pri.pri_coding " & _
            "AND      pro.pro_codigo = b.cpi_codped AND ing.ing_codigo = b.cpi_coding AND b.cpi_cencos = '" & MuestraCasino(1) & "' " & _
            "AND      pri.pri_codpro = b.cpi_codped " & _
            "AND      ing.ing_unimed = unm.unm_codigo " & _
            "AND      pro.pro_coduni = uni.uni_codigo " & _
            "AND      bod.bod_codpro = pro.pro_codigo " & _
            "AND     (pro.pro_fecven > " & Format(Date, "yyyymmdd") & " OR pro.pro_fecven <= 0 OR bod.bod_canmer > 0) " & _
            "AND      bod.bod_codbod = " & vg_codbod & " " & _
            "AND      pro.pro_ctrsto = 1 " & _
            "GROUP BY pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, ing.ing_codigo, " & _
            "         ing.ing_nombre, bod.bod_canmer, unm.unm_nomcor, aux.sec_codigo, aux.sec_orden, aux.sec_nombre " & _
            "ORDER BY " & sqlparam & "", vg_db, adOpenStatic

End If
vaSpread1.MaxRows = 0
vaSpread1.Visible = False
vaSpread2.Visible = False
i = 0: codsec = 0
Do While Not RS1.EOF
    
    i = i + 1
    
    If codsec <> RS1!sec_codigo Then
       
       vaSpread2.MaxRows = vaSpread2.MaxRows + 1
       vaSpread2.Row = vaSpread2.MaxRows
       vaSpread2.Col = 1: vaSpread2.Value = RS1!sec_codigo
       vaSpread2.Col = 2: vaSpread2.Value = Trim(RS1!sec_nombre)
       codsec = RS1!sec_codigo
    
    End If
    
    '-------> Ingrediente
    Cantidad = IIf(IsNull(RS1!cantidad2), 0, RS1!cantidad2)
    vaSpread1.MaxRows = i
    vaSpread1.Row = i
    vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
    vaSpread1.RowHidden = IIf((vaSpread2.Row = 1 Or Option1(0).Value = True) And Check1(0).Value = 0, False, True)
    vaSpread1.Col = 1: vaSpread1.text = Trim(RS1!ing_codigo)
    vaSpread1.Col = 2: vaSpread1.text = Trim(RS1!ing_nombre)
    vaSpread1.Col = 3: vaSpread1.text = Trim(RS1!unm_nomcor)
    vaSpread1.Col = 4: vaSpread1.text = Format(Round(Cantidad, vg_DCa), fg_Pict(9, vg_DCa))
    vaSpread1.Col = 8: vaSpread1.text = "NI" 'No bloquedo - Ingrediente
    vaSpread1.Col = 10: vaSpread1.text = RS1!sec_codigo
    vaSpread1.Col = -1
    vaSpread1.FontBold = True: vaSpread1.Lock = True
    vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(3).FillColor
    
    '-------> Producto
    i = i + 1
    Cantidad = IIf(IsNull(RS1!Cantidad), 0, RS1!Cantidad)
    If Cantidad > 0 And Cantidad < 0.5 Then canreal = Format(0.5, fg_Pict(9, vg_DCa)) Else canreal = Format(Round(Cantidad, 0), fg_Pict(9, vg_DPr))
    vaSpread1.MaxRows = i
    vaSpread1.Row = i
    vaSpread1.RowHidden = IIf(vaSpread2.Row = 1 Or Option1(0).Value = True, False, True)
    vaSpread1.Col = 1: vaSpread1.text = Trim(RS1!pro_codigo)
    vaSpread1.Col = 2: vaSpread1.text = Trim(RS1!pro_nombre)
    vaSpread1.Col = 3: vaSpread1.text = Trim(RS1!uni_nomcor)
    vaSpread1.Col = 4: vaSpread1.text = Format(Cantidad, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
    vaSpread1.Col = 5: vaSpread1.ForeColor = &HFF0000: vaSpread1.text = Format(Round(canreal, vg_DCa), fg_Pict(9, vg_DCa))
    vaSpread1.Col = 6: vaSpread1.text = Format(RS1!ppd_propon, fg_Pict(9, vg_DPr))
    vaSpread1.Col = 7: vaSpread1.text = Format(Round(canreal, vg_DCa) * RS1!ppd_propon, fg_Pict(9, vg_DPr))
    vaSpread1.Col = 8: vaSpread1.text = "NM" 'No bloquedo - Viene de la Minuta
    vaSpread1.Col = 9: vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa))
    vaSpread1.Col = 10: vaSpread1.text = RS1!sec_codigo
    '-------> Revisa color
    vaSpread1.Col = 5: canrea = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
    vaSpread1.Col = 8: color = Right(vaSpread1.text, 1)
    vaSpread1.Col = 9: canbod = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
    '-------> Mover sectores totales
    If vaSpread2.MaxRows > 0 And canreal <> 0 Then vaSpread2.Col = 3: vaSpread2.TypeHAlign = TypeHAlignRight: vaSpread2.text = Format(IIf(Trim(vaSpread2.text) = "", 0, vaSpread2.Value) + (Round(canreal, vg_DCa) * RS1!ppd_propon), fg_Pict(9, vg_DPr))
    canaux = 0
    For z = 1 To vaSpread1.MaxRows
        vaSpread1.Row = z
        vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
        vaSpread1.Col = 1
        If RS1!pro_codigo = vaSpread1.text And color2 <> "I" Then
            vaSpread1.Col = 5: canaux = canaux + Format(IIf(vaSpread1.text = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
        End If
    Next z
    canrea = IIf(canaux > 0, canaux, canrea)
    
    If canbod - canrea >= 0 Then
       
       For z = 1 To vaSpread1.MaxRows
           vaSpread1.Row = z
           vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
           vaSpread1.Col = 1
           If RS1!pro_codigo = vaSpread1.text And color2 <> "I" Then
               vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(IIf(color = "M", 2, IIf(color = "U", 0, 3))).FillColor
               vaSpread1.Col = 8: vaSpread1.text = "N" & color 'No Bloqueado - Depende
           End If
       Next z
    
    Else
       
       For z = 1 To vaSpread1.MaxRows
           vaSpread1.Row = z
           vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
           vaSpread1.Col = 1
           If RS1!pro_codigo = vaSpread1.text And color2 <> "I" Then
              vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
              vaSpread1.Col = 8: vaSpread1.text = "S" & color 'Bloqueado - Depende
           End If
       Next z
    
    End If
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
'---------------------------------- FIN MINUTA ---------------------------------
'----------------------------- ESTRUCTURA FIJA ---------------------------------

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If vg_tipbase = "1" Then
   
   RS1.Open "SELECT b.pro_codigo, e.ppd_propon, b.pro_nombre, c.uni_nomcor, round(d.bod_canmer, " & vg_DCa & ") as bod_canmer, a.mfd_canpro AS cantidad " & _
            "FROM b_minutafijadia a, b_productos b, a_unidad c, b_bodegas d, " & aAp1 & " e " & _
            "WHERE a.mfd_codpro = b.pro_codigo AND b.pro_codigo = e.ppd_codpro AND e.ppd_cencos='" & MuestraCasino(1) & "' " & _
            "AND   b.pro_codigo = d.bod_codpro " & _
            "AND   b.pro_coduni = c.uni_codigo " & _
            "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0 OR d.bod_canmer > 0) " & _
            "AND   d.bod_codbod = " & vg_codbod & " " & _
            "AND   a.mfd_cencos = '" & Trim(fpText1(1).text) & "' " & _
            "AND   a.mfd_codreg = " & codreg & " " & _
            "AND   a.mfd_codser = " & codser & " " & _
            "AND   a.mfd_fecha  = " & feprod & " AND a.mfd_tipmin = '2' " & _
            "AND   b.pro_ctrsto = 1 AND d.bod_codbod = " & vg_codbod & "", vg_db, adOpenStatic

Else
   
   RS1.Open "SELECT b.pro_codigo, b.pro_nombre, c.uni_nomcor, round(d.bod_canmer, " & vg_DCa & ") as bod_canmer, a.mfd_canpro AS cantidad, " & _
            "        (SELECT TOP 1 ppd_propon FROM b_productospmpdia WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(fpDateTime1(1).text), "yyyymmdd") & " ORDER BY ppd_fecdia DESC) AS ppd_propon " & _
            "FROM b_minutafijadia a, b_productos b, a_unidad c, b_bodegas d " & _
            "WHERE a.mfd_codpro = b.pro_codigo " & _
            "AND   b.pro_codigo = d.bod_codpro " & _
            "AND   b.pro_coduni = c.uni_codigo " & _
            "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0 OR d.bod_canmer > 0) " & _
            "AND   d.bod_codbod = " & vg_codbod & " " & _
            "AND   a.mfd_cencos = '" & Trim(fpText1(1).text) & "' " & _
            "AND   a.mfd_codreg = " & codreg & " " & _
            "AND   a.mfd_codser = " & codser & " " & _
            "AND   a.mfd_fecha = " & feprod & " AND a.mfd_tipmin = '2' " & _
            "AND   b.pro_ctrsto = 1 AND d.bod_codbod = " & vg_codbod & "", vg_db, adOpenStatic

End If
z = 1
If Not RS1.EOF Then
    i = i + 1
    '-------> Ingrediente
    If z = 1 And Option1(1).Value = True Then
       
       vaSpread2.MaxRows = vaSpread2.MaxRows + 1
       vaSpread2.Row = vaSpread2.MaxRows
       vaSpread2.Col = 1: vaSpread2.Value = "estfij"
       vaSpread2.Col = 2: vaSpread2.Value = "Estructura Fija"
       z = z + 1
    
    End If
    
    vaSpread1.MaxRows = i
    vaSpread1.Row = i
    vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
    vaSpread1.RowHidden = IIf((vaSpread2.Row = 1 Or Option1(0).Value = True) And Check1(0).Value = 0, False, True)
    vaSpread1.Col = 1: vaSpread1.text = ""
    vaSpread1.Col = 2: vaSpread1.text = "Estructura Fija"
    vaSpread1.Col = 8: vaSpread1.text = "NI" 'No bloquedo - Ingrediente
    vaSpread1.Col = 10: vaSpread1.text = "estfij"
    vaSpread1.Col = -1
    vaSpread1.FontBold = True: vaSpread1.Lock = True
    vaSpread1.BackColor = Shape1(3).FillColor
    vaSpread1.Col = 1
    vaSpread1.ForeColor = Shape1(3).FillColor

End If
Do While Not RS1.EOF
    '-------> Producto
    i = i + 1
    Cantidad = IIf(IsNull(RS1!Cantidad), 0, RS1!Cantidad)
    If Cantidad > 0 And Cantidad < 0.5 Then canreal = Format(0.5, fg_Pict(9, vg_DCa)) Else canreal = Format(Round(Cantidad, 2), fg_Pict(9, vg_DPr))
    vaSpread1.MaxRows = i
    vaSpread1.Row = i
    vaSpread1.RowHidden = IIf(vaSpread2.Row = 1 Or Option1(0).Value = True, False, True)
    vaSpread1.Col = 1: vaSpread1.text = Trim(RS1!pro_codigo)
    vaSpread1.Col = 2: vaSpread1.text = Trim(RS1!pro_nombre)
    vaSpread1.Col = 3: vaSpread1.text = Trim(RS1!uni_nomcor)
    vaSpread1.Col = 4: vaSpread1.text = Format(Cantidad, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa))) 'vg_DCa))
    vaSpread1.Col = 5: vaSpread1.ForeColor = &HFF0000: vaSpread1.text = Format(Round(canreal, vg_DCa), fg_Pict(9, vg_DCa))
    vaSpread1.Col = 6: vaSpread1.text = Format(RS1!ppd_propon, fg_Pict(9, vg_DPr))
    vaSpread1.Col = 7: vaSpread1.text = Format(Round(canreal, vg_DCa) * RS1!ppd_propon, fg_Pict(9, vg_DPr))
    vaSpread1.Col = 8: vaSpread1.text = "NM" 'No bloquedo - Viene de la Minuta
    vaSpread1.Col = 9: vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa))
    vaSpread1.Col = 10: vaSpread1.text = "estfij"
    '-------> Mover sectores totales
    If vaSpread2.MaxRows > 0 And canreal <> 0 Then vaSpread2.Col = 3: vaSpread2.TypeHAlign = TypeHAlignRight: vaSpread2.text = Format(IIf(Trim(vaSpread2.text) = "", 0, vaSpread2.Value) + (Round(canreal, vg_DCa) * RS1!ppd_propon), fg_Pict(9, vg_DPr))
    '-------> Revisa color
    vaSpread1.Col = 5: canrea = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
    vaSpread1.Col = 8: color = Right(vaSpread1.text, 1)
    vaSpread1.Col = 9: canbod = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
    
    canaux = 0
    For z = 1 To vaSpread1.MaxRows
        vaSpread1.Row = z
        vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
        vaSpread1.Col = 1
        If RS1!pro_codigo = vaSpread1.text And color2 <> "I" Then
            vaSpread1.Col = 5: canaux = canaux + Format(IIf(vaSpread1.text = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
        End If
    Next z
    canrea = IIf(canaux > 0, canaux, canrea)
    
    If canbod - canrea >= 0 Then
        
        For z = 1 To vaSpread1.MaxRows
            vaSpread1.Row = z
            vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
            vaSpread1.Col = 1
            If RS1!pro_codigo = vaSpread1.text And color2 <> "I" Then
                vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(IIf(color = "M", 2, IIf(color = "U", 0, 3))).FillColor
                vaSpread1.Col = 8: vaSpread1.text = "N" & color '-------> No Bloqueado - Depende
            End If
        Next z
        
    Else
        
        For z = 1 To vaSpread1.MaxRows
            vaSpread1.Row = z
            vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
            vaSpread1.Col = 1
            If RS1!pro_codigo = vaSpread1.text And color2 <> "I" Then
                vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
                vaSpread1.Col = 8: vaSpread1.text = "S" & color '-------> Bloqueado - Depende
            End If
        Next z
    
    End If
    
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
'----------------------------- FIN ESTRUCTURA FIJA ------------------------------
'-------> Borrar tablas temporales
If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
If Trim(aAp1) <> "" Then vg_db.Execute "DROP TABLE " & aAp1 & ""
If vaSpread1.MaxRows > 0 Then Frame1.Enabled = False
modo = "A"
Frame2.Enabled = True
If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
vaSpread2.Visible = IIf(Option1(0).Value = True, False, True)
vaSpread1.Visible = True
Me.MousePointer = 0
SendKeys "{Tab}"

End Sub

Private Sub Combo1_Click(Index As Integer)
Dim feprod As Long, codser As Long, codreg As Long, codbod As Long, numdia As Long, canreal As Double, Cantidad As Double, aAp As String
Dim codsec As Long
Dim canrea As Double, canbod As Double, color As Variant, estado As String, codpro As String
If est Then Exit Sub
fg_carga ""
codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
est1 = False
Select Case Index
Case 0
    MostrarDetalle
Case 1
    
    If vaSpread1.MaxRows = 0 Then fg_descarga: Exit Sub
    fg_carga ""
    vaSpread1.Visible = False
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 8: color = Right(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), 1)
        If color <> "I" Then
           
           vaSpread1.Col = 1: codpro = Trim(LimpiaDato(vaSpread1.text))
           
           If RS2.State = 1 Then RS2.Close
           RS2.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           RS2.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                    "WHERE  bod.bod_codpro = pro.pro_codigo " & _
                    "AND    bod.bod_codbod = " & vg_codbod & " " & _
                    "AND    pro.pro_ctrsto = 1 " & _
                    "AND    pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "'", vg_db, adOpenStatic
           vaSpread1.Col = 9
           If Not RS2.EOF Then vaSpread1.text = Format(RS2!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
           RS2.Close: Set RS2 = Nothing
        End If
        '-------> Revisa color
        vaSpread1.Col = 5: canrea = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
        vaSpread1.Col = 8: color = Right(vaSpread1.text, 1)
        vaSpread1.Col = 9: canbod = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
        If color <> "I" Then
            canaux = 0
            For z = 1 To vaSpread1.MaxRows
                vaSpread1.Row = z
                vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
                vaSpread1.Col = 1
                If codpro = vaSpread1.text And color2 <> "I" Then
                    vaSpread1.Col = 5: canaux = canaux + Format(IIf(vaSpread1.text = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
                End If
            Next z
            canrea = IIf(canaux > 0, canaux, canrea)
            
            If canbod - canrea >= 0 Then
                For z = 1 To vaSpread1.MaxRows
                    vaSpread1.Row = z
                    vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
                    vaSpread1.Col = 1
                    If codpro = vaSpread1.text And color2 <> "I" Then
                        vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(IIf(color = "M", 2, IIf(color = "U", 0, 3))).FillColor
                        vaSpread1.Col = 8: vaSpread1.text = "N" & color '-------> No Bloqueado - Depende
                    End If
                Next z
            Else
                For z = 1 To vaSpread1.MaxRows
                    vaSpread1.Row = z
                    vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
                    vaSpread1.Col = 1
                    If codpro = vaSpread1.text And color2 <> "I" Then
                        vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
                        vaSpread1.Col = 8: vaSpread1.text = "S" & color '-------> Bloqueado - Depende
                    End If
                Next z
            End If
        
        End If
    Next i
    vaSpread1.Visible = True
    fg_descarga
End Select
fg_descarga
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 7920
Me.Width = 11925
fpDateTime1(1).DateTimeFormat = UserDefined
fpDateTime1(1).UserDefinedFormat = "dd/mm/yyyy"
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")

EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)
fg_centra Me
est = False
Me.HelpContextID = vg_OpcM
MsgTitulo = "Salida Producción"
Dim X As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
Gl_Mo_Botones Me, 12
vaSpread1.Row = -1
vaSpread1.Col = 4: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 6: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
vaSpread1.Col = 7: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
'-------> Cargar Combo Bodega
CargarDatoCombo Combo1, 1, "b_clientes", "cli_", "CliBod", "N"
Limpia 2
'-------> Ocultar regimen y servicio cuando el contrato es FM
If vg_tipser Then
   Shape1(3).Visible = False
   Label5(3).Visible = False
   Label3(10).Visible = False
   fpLongInteger1(1).Visible = False
   fpLongInteger1(2).Visible = False
   Image1(0).Visible = False
   Image1(2).Visible = False
   fpayuda(0).Visible = False
   fpayuda(3).Visible = False
   fpayuda(2).Visible = False
   fpayuda(5).Visible = False
   Label3(2).Visible = False
'   Combo1(0).Visible = False
   fpayuda(3).Visible = False
   Option1(0).Visible = False
   Option1(1).Visible = False
   Check1(0).Visible = False
   Frame5.Visible = False
   Option1(0).Value = True: Option1(1).Value = False: Check1(0).Value = 1
End If
If Not vg_tipser Then
   Check1(0).Value = IIf(0 = (fg_CambiaChar(GetParametro("ingsalpro"), ";", "','")), 0, 1)
   If 0 = (fg_CambiaChar(GetParametro("salressec"), ";", "','")) Then
      vaSpread2.Visible = False
      vaSpread1.Top = vaSpread2.Top
      vaSpread1.Height = 3165 + vaSpread2.Height
      Option1(0).Value = True: Option1(1).Value = False
   Else
      vaSpread2.Visible = True
      vaSpread1.Top = 1755
      vaSpread1.Height = 3165
      Option1(0).Value = False: Option1(1).Value = True
   End If
Else
    vaSpread2.Visible = False
    vaSpread1.Top = vaSpread2.Top
    vaSpread1.Height = 3165 + vaSpread2.Height
End If
TraerFechaCierre
est = False: est1 = False
End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then
    Frame3.Left = (Me.Width \ 2) - (Frame3.Width \ 2)
    Frame1.Left = (Me.Width \ 2) - (Frame1.Width \ 2)
    Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)
    Frame5.Left = 10500
ElseIf Me.WindowState = 0 Then
    Frame3.Left = 180
    Frame1.Left = 0
    Frame2.Left = 180
    Frame5.Left = 8970
End If
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If est Then Exit Sub
If Trim(fpDateTime1(0).text) = "" Or Trim(fpDateTime1(1).text) = "" Then Exit Sub
If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then Exit Sub
Select Case Index
Case 1
    vaSpread1.MaxRows = 0
    If vg_tipser Then MostrarDetalle
End Select
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpDouble1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 1
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & "", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS1!reg_nombre)
    RS1.Close: Set RS1 = Nothing
    MostrarDetalle

Case 2
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS1.Open "SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(2).Value) & " AND ser_activo = '1'", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(3).Caption = "": Exit Sub
    fpayuda(3).Caption = Trim(RS1!ser_nombre)
    RS1.Close: Set RS1 = Nothing
    MostrarDetalle

End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_LostFocus(Index As Integer)
If fpText1(1).text = "" Then Exit Sub

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo = '" & fpText1(1).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
       fpayuda(Index).Caption = RS1!cli_nombre
       Gl_Ac_Botones Me, 12, 2, ""
       fpText1(1).Enabled = False
       RS1.MoveNext
    Loop
Else
   RS1.Close: Set RS1 = Nothing
   MsgBox "Contrato no existe...", vbExclamation + vbOKOnly, MsgTitulo
   Limpia 2
   If fpText1(1).Enabled = True Then fpText1(1).SetFocus
   Exit Sub
End If
RS1.Close: Set RS1 = Nothing
fpLongInteger1(0).text = TraerCorrelativo(vg_codbod, "SP") 'MuestraFolio(Trim(fpText1(1).text))
End Sub

Private Sub Image1_Click(Index As Integer)
vg_codigo = 0
Select Case Index
Case 0 '-------> Regimen
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
'    If fpLongInteger1(2).Enable = True Then fpLongInteger1(2).SetFocus
Case 1 '-------> Contraro
    vg_left = fpayuda(Index).Left + 1920
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    If Trim(vg_codigo) <> fpText1(Index) Then Limpia 2
    fpText1(Index) = Trim(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    fpText1_LostFocus 1
    If fpDateTime1(Index - 1).Enabled = True Then fpDateTime1(Index - 1).SetFocus
    Gl_Ac_Botones Me, 12, 2, ""
Case 2 '-------> Servicio
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Ser"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpayuda(3).Caption = vg_nombre
    fpLongInteger1(2).Value = Val(vg_codigo)

End Select
End Sub

Private Sub Option1_Click(Index As Integer)
On Error GoTo Man_Error
If Option1(0).Value = True Then
   vaSpread2.Visible = False
   vaSpread1.Top = vaSpread2.Top
   vaSpread1.Height = vaSpread1.Height + vaSpread2.Height
Else
   vaSpread2.Visible = True
   vaSpread1.Top = 1755
   vaSpread1.Height = 3165
End If
If est Then Exit Sub
'-------> Actualizar parametro salida producción
vg_db.BeginTrans
vg_db.Execute "UPDATE a_param SET par_valor='" & IIf(Option1(0).Value = True, 0, 1) & "' WHERE par_cencos='" & MuestraCasino(1) & "' AND par_codigo='salressec'"
vg_db.CommitTrans
Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim rutcli As String, tipdoc As String, NumDoc As Long, Fecha As Long, codbod  As Long, fecemi As Date, fecpro As Date, codreg As Long, codser As Long, i As Long, canact As Double, aAp  As String, estdoc As String
Dim numlin As Long, codmer As String, coding As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String, total As Double, diablq As Date, color As String, codsec As String, totdec As Double
Dim sql1 As String
Dim NumRac As Long

On Error GoTo Man_Error

MsgTitulo = "Salida Producción"
est1 = False
codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
fecpro = Format(fpDateTime1(1).Value, "dd/mm/yyyy")
If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
TraerFechaCierre

Select Case Button.Index

Case 1, 6 '-------> Nuevo-Cancelar
    
    modo = IIf(Button.Index = 1, "A", "")
    If Button.Index = 6 And vaSpread1.MaxRows > 0 Then If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    Limpia IIf(Button.Index = 1, 6, 2)
    If fpText1(1).Enabled = True Then fpText1(1).SetFocus
    Frame2.Enabled = IIf(vg_tipser, True, False)
    If vg_tipser Then vaSpread1.SetFocus

Case 8, 15 '-------> Graba
    
    If Button.Index = 15 Then modo = "M": est1 = True
    If vaSpread1.MaxRows > 1 Then vaSpread1_EditMode vaSpread1.ActiveCol, vaSpread1.ActiveRow, 0, True
    If Trim(fpText1(1).text) = "" Or Trim(fpLongInteger1(0).text) = "" Or Trim(fpDateTime1(0).text) = "" _
    Or Trim(Combo1(1).text) = "" Or Trim(fpDateTime1(1).text) = "" Or (Trim(fpayuda(0).Caption) = "" And Not vg_tipser) Or (Trim(fpayuda(3).Caption) = "" And Not vg_tipser) Then MsgBox "Debe ingresar dato importante...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    '-------> Validar si el contrato tiene asignado inventario rotativo
    If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
       Format(fpDateTime1(1).Value, "dd/mm/yyyy") > CDate(vg_ciedia) Then MsgBox "Tiene que realizar cierre diario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    If CierrePeriodo(Format(fpDateTime1(1).text, "yyyymmdd"), codbod, 6) Then
    
       MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CierrePeriodo(Format(fpDateTime1(1).text, "yyyymmdd"), codbod, 0) Then
    
       MsgBox "Documento no corresponde al periodo : " & VgLinea & VgLinea & CierreFecha, vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CierrePeriodo(Format(fpDateTime1(1).text, "yyyymmdd"), codbod, 6) Then
    
       MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    'Validar inventario calendarizado 20201001
    If CierrePeriodo(Format(fpDateTime1(1).text, "yyyymmdd"), codbod, 38) Then
        
       MsgBox "Se esta realizando la toma de inventario en estos momento...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
    End If
        
    'Validar ingreso documento inventario calendarizado 20201001
    If CierrePeriodo(Format(fpDateTime1(1).text, "yyyymmdd"), codbod, 40) Then
        
       MsgBox "No puede ingresar documento, antes de un inventario calendarizado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
    End If
    
    If CierrePeriodo(Format(fpDateTime1(1).text, "yyyymmdd"), codbod, 8) Then
    
       MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CDate(fpDateTime1(1).text) < CDate(vg_ciedia) Then
    
       MsgBox "Día se encuentra cerrado, no es posible ingresar...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    Toolbar1.Enabled = False
    If Button.Index = 15 Then
       
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS1.Open "SELECT tov_estdoc FROM b_totventas WHERE tov_rutcli = '" & Trim(LimpiaDato(fpText1(1).text)) & "' " & _
                "AND tov_tipdoc = 'SP' AND tov_numdoc = " & fpLongInteger1(0).Value & " AND tov_estdoc = 'P' AND tov_codbod = " & vg_codbod & "", vg_db, adOpenStatic
       If RS1.EOF Then
          
          RS1.Close: Set RS1 = Nothing
          Gl_Ac_Botones Me, 12, 3, ""
          Toolbar1.Buttons(15).Enabled = False
          Label1.Caption = "": Frame1.Enabled = False: Frame2.Enabled = False: vaSpread1.Col = -1: vaSpread1.Row = -1: vaSpread1.Lock = True
          MsgBox "Documento fue cerrado por otro usuario, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo: Toolbar1.Enabled = True: Exit Sub
       
       End If
       RS1.Close: Set RS1 = Nothing
       
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i: vaSpread1.Col = 8
           If Left(vaSpread1.text, 1) = "S" Then est1 = False: MsgBox "Existe una cantidad que exede el Stock...", vbExclamation + vbOKOnly, MsgTitulo: Toolbar1.Enabled = True: Exit Sub
       Next i
    
    End If
    
    total = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 7: ptotal = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DPr))
        total = total + ptotal
    Next i
    
    If total = 0 Then MsgBox "El total del documento debe ser mayor a 0...", vbExclamation + vbOKOnly, MsgTitulo: Toolbar1.Enabled = True: Exit Sub
    '-------> validar si graba documentos con rebaja de bodega
    If Button.Index = 15 Then If MsgBox("Esta Seguro Cerrar Salida...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then est1 = False: Toolbar1.Enabled = True: Exit Sub

paso:
    rutcli = Trim(LimpiaDato(fpText1(1).text))
    tipdoc = "SP"
    NumDoc = fpLongInteger1(0).text
    codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
    fecemi = Format(fpDateTime1(0).text, "dd/mm/yyyy")
    fecpro = Format(fpDateTime1(1).text, "dd/mm/yyyy")
    Fecha = Val(fpDateTime1(1).Year & Right("0" & fpDateTime1(1).Month, 2) & Right("0" & fpDateTime1(1).Day, 2))
    If vg_tipser Then codreg = 0 Else codreg = Val(fpLongInteger1(1).Value) 'Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 22, 10))
    If vg_tipser Then codser = 0 Else codser = Val(fpLongInteger1(2).Value) 'Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 10, 10))
    est1 = False
    
    If fpLongInteger1(0).text < 1 Or modo = "A" Then
       
       NumDoc = TraerCorrelativo(codbod, "SP")
       vg_db.Execute "UPDATE b_parametros SET par_correlativo = " & NumDoc & " WHERE par_codbod = " & codbod & " AND par_tipdoc = 'SP'"
       fpLongInteger1(0).text = NumDoc
       DoEvents
    
    End If
    
    vg_db.BeginTrans
    '-------> Borrar datos si datos existen detalle y encabezado salida produccción
    If modo <> "A" Then
       
       vg_db.Execute "DELETE b_detventasimp FROM b_detventasimp WHERE imd_rutdoc = '" & rutcli & "' AND imd_tipdoc = '" & tipdoc & "' AND imd_numdoc = " & NumDoc & ""
       vg_db.Execute "DELETE b_detventas FROM b_detventas WHERE dev_rutcli = '" & rutcli & "' AND dev_tipdoc = '" & tipdoc & "' AND dev_numdoc = " & NumDoc & ""
       vg_db.Execute "DELETE b_totventas FROM b_totventas WHERE tov_rutcli = '" & rutcli & "' AND tov_tipdoc = '" & tipdoc & "' AND tov_numdoc = " & NumDoc & " AND tov_codbod = " & vg_codbod & ""
    
    End If
    
    '-------> traer raciones reales
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS1.Open "SELECT distinct isnull(min_racrea,0) as min_racrea FROM b_minuta WHERE min_cencos = '" & fpText1(1).text & "' AND min_codreg = " & codreg & " and min_codser = " & codser & " and min_fecmin = " & Format(fpDateTime1(1).text, "yyyymmdd") & "", vg_db, adOpenStatic
    If Not RS1.EOF Then
    
       NumRac = RS1!min_racrea
       
    End If
    RS1.Close
    Set RS1 = Nothing
    
    '-------> Encabezado
    vg_db.Execute "INSERT INTO b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_numinf, tov_racionesproduccion, tov_fechacreacion, tov_fechamodificacion, tov_usuariocreacion, tov_usuariomodificacion) " & _
                  "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & codbod & ", '" & Format(fpDateTime1(0).text, "yyyymmdd") & "', '" & Format(fpDateTime1(1).text, "yyyymmdd") & "', " & codreg & ", " & codser & ", 0, '', '', 0, " & NumRac & ", getdate(), getdate(), '" & vg_NUsr & "', '" & vg_NUsr & "')"
    
    '-------> Detalle
    total = 0: numlin = 1
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1: codmer = Trim(LimpiaDato(vaSpread1.text))
        vaSpread1.Col = 2: descri = Trim(LimpiaDato(vaSpread1.text))
        vaSpread1.Col = 4: canmin = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), IIf(vg_pais = "CL", 3, vg_DCa)) 'vg_DCa)
        vaSpread1.Col = 5: canmer = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DCa)
        vaSpread1.Col = 6: predoc = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DPr)
        vaSpread1.Col = 7: ptotal = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DPr)
        vaSpread1.Col = 8: color = Right(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), 1)
        vaSpread1.Col = 10: codsec = Trim(vaSpread1.text)
        total = total + ptotal
        If color = "I" Then '-------> Rescata el Ingrediente
           vaSpread1.Col = 1: coding = Trim(LimpiaDato(vaSpread1.text))
        End If
        If color <> "I" Then '-------> No entra si es ingrediente
            If color <> "U" Or canmer <> 0 Then '-------> No entra si es ingresado x usuario y valor es cero
               vg_db.Execute "INSERT INTO b_detventas (dev_rutcli, dev_tipdoc, dev_numdoc, dev_numlin, dev_codmer, dev_canmin, dev_canmer, dev_predoc, dev_ptotal, dev_descri, dev_mueinv, dev_porcen, dev_precos, dev_coding, dev_codsec) " & _
                             "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & numlin & ", '" & codmer & "', " & canmin & ", " & canmer & ", " & predoc & ", " & ptotal & ", '" & descri & "', 'S', 0, " & predoc & ", '" & coding & "', " & IIf(codsec = "0" Or Option1(0).Value = True, "Null", IIf(codsec = "estfij", -1, codsec)) & ")"
'               '-------> Actualizar costo receta planificación
'               vg_db.Execute "UPDATE ((b_minuta INNER JOIN b_minutacosto ON (b_minutacosto.mic_cencos=b_minuta.min_cencos) AND (b_minuta.min_fecmin=b_minutacosto.mic_fecval)) INNER JOIN b_ingrediente ON b_minutacosto.mic_codpro=b_ingrediente.ing_codigo) INNER JOIN b_productos ON b_ingrediente.ing_codcom=b_productos.pro_codigo SET b_minutacosto.mic_cospro=" & predoc / b_productos.pro_facing & " " & _
'                             "WHERE b_minutacosto.mic_codpro='" & coding & "' AND b_minutacosto.mic_tipmin='1' AND b_minuta.min_cencos='" & rutcli & "' AND b_minuta.min_codreg=" & codreg & " AND b_minuta.min_codser=" & codser & " AND b_minuta.min_fecmin=" & Format(fpDateTime1(1).text, "yyyymmdd") & ""
               '-------> Control de Stock
               If Button.Index = 15 Then
                  '-------> Validar stock es negativo
                  
                  If RS2.State = 1 Then RS2.Close
                  RS2.CursorLocation = adUseClient
                  vg_db.CursorLocation = adUseClient
              
                  RS2.Open "SELECT a.*, b.bod_canmer FROM b_productos a, b_bodegas b WHERE a.pro_codigo = b.bod_codpro AND b.bod_codbod = " & codbod & " AND  a.pro_codigo = '" & codmer & "' AND a.pro_ctrsto = 1", vg_db, adOpenStatic
                  If Not RS2.EOF Then
'                     If (Round(RS2!bod_canmer, 2) - canmer) < 0 Then
                     If (Round(RS2!bod_canmer, vg_DCa) - canmer) < 0 Then
                        RS2.Close: Set RS2 = Nothing
                        vg_db.RollbackTrans
                        For j = 1 To vaSpread1.MaxRows
                            vaSpread1.Row = j
                            vaSpread1.Col = 8: color = Right(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), 1)
                            If color <> "I" Then
                               vaSpread1.Col = 1
                               
                               If RS1.State = 1 Then RS1.Close
                               RS1.CursorLocation = adUseClient
                               vg_db.CursorLocation = adUseClient
          
                               
                               RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                                        "WHERE  bod.bod_codpro = pro.pro_codigo " & _
                                        "AND    pro.pro_ctrsto = 1 " & _
                                        "AND    bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                                        "AND    pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "'", vg_db, adOpenStatic
                               vaSpread1.Col = 9
                               If Not RS1.EOF Then vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
                               RS1.Close: Set RS1 = Nothing
                            End If
                        Next j
                        MsgBox "Existen productos con diferencia en la bodega, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo
                        Toolbar1.Enabled = True
                        Exit Sub
                     End If
                     vg_db.Execute "UPDATE b_bodegas SET bod_canmer = bod_canmer-" & canmer & " WHERE bod_codpro = '" & codmer & "' AND bod_codbod = " & vg_codbod & ""
                  End If
                 RS2.Close: Set RS2 = Nothing
               End If
               numlin = numlin + 1
            End If
        End If
    Next i
    
    '-------> Grabar total
    vg_db.Execute "UPDATE b_totventas SET tov_totdoc = " & total & ", tov_estdoc = '" & IIf(Button.Index = 15, "", "P") & "' WHERE tov_rutcli = '" & Trim(LimpiaDato(fpText1(1).text)) & "' " & _
                  "AND tov_tipdoc = 'SP' AND tov_numdoc = " & fpLongInteger1(0).Value & " AND tov_codbod = " & vg_codbod & ""
    vg_db.CommitTrans
    Gl_Ac_Botones Me, 12, 3, ""
    Label1.Caption = IIf(Button.Index = 15, "", "PENDIENTE")
    Frame1.Enabled = False 'True 'False
    Frame2.Enabled = True 'False
    If Button.Index = 15 Then Frame1.Enabled = False: Frame2.Enabled = False: vaSpread1.Col = -1: vaSpread1.Row = -1: vaSpread1.Lock = True
    
    '-------> Revisa Stock
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 8: color = Right(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), 1)
        If color <> "I" Then
           vaSpread1.Col = 1
           
           If RS1.State = 1 Then RS1.Close
           RS1.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient

           RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                    "WHERE  bod.bod_codpro = pro.pro_codigo " & _
                    "AND    pro.pro_ctrsto = 1 " & _
                    "AND    bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                    "AND    pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "'", vg_db, adOpenStatic
           vaSpread1.Col = 9
           If Not RS1.EOF Then vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
           RS1.Close: Set RS1 = Nothing
        End If
    Next i
    If Button.Index = 15 Then I_SalDevBod Me, "SP": Toolbar1.Buttons(15).Enabled = False
    Toolbar1.Enabled = True
    modo = "M"
    
Case 3 '-------> Anular

    If CierrePeriodo(Format(fpDateTime1(1).text, "yyyymmdd"), codbod, 0) Then MsgBox "Periodo esta cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(1).text, "yyyymmdd"), codbod, 6) Then MsgBox "No puede anular documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CDate(fpDateTime1(1).text) < CDate(vg_ciedia) Then MsgBox "No puede anular documento, día esta cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
    '-------> Validar si existe devolucion de producción
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    NumDoc = fpLongInteger1(0).text
    codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
    fecemi = Format(fpDateTime1(0).text, "dd/mm/yyyy")
    fecpro = Format(fpDateTime1(1).text, "dd/mm/yyyy")
    Fecha = Val(fpDateTime1(1).Year & Right("0" & fpDateTime1(1).Month, 2) & Right("0" & fpDateTime1(1).Day, 2))
    If vg_tipser Then codreg = 0 Else codreg = Val(fpLongInteger1(1).Value)
    If vg_tipser Then codser = 0 Else codser = Val(fpLongInteger1(2).Value)
     
    Set RS1 = vg_db.Execute("sgp_Sel_ValidarDevolucionProduccion '" & Trim(LimpiaDato(fpText1(1).text)) & "', " & codreg & ", " & codser & ", " & codbod & ", '" & Fecha & "', " & NumDoc & " ")
    
    If Not RS1.EOF Then
    
        RS1.Close
        Set RS1 = Nothing

       MsgBox "No puede anular documento, ya que existen devolución producción. debe anular la devolución producción...", vbExclamation + vbOKOnly, MsgTitulo
       
       Exit Sub
    
    End If
    RS1.Close
    Set RS1 = Nothing
    
    If MsgBox("Anula documento...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
    estdoc = "": totdec = 0
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "SELECT tov_estdoc, tov_totdoc FROM b_totventas WHERE tov_rutcli = '" & Trim(LimpiaDato(fpText1(1).text)) & "' " & _
             "AND tov_tipdoc = 'SP' AND tov_numdoc = " & fpLongInteger1(0).Value & " AND tov_codbod = " & vg_codbod & "", vg_db, adOpenStatic
    If Not RS1.EOF Then estdoc = RS1!tov_estdoc: totdec = RS1!tov_totdoc
    RS1.Close: Set RS1 = Nothing
    vg_db.BeginTrans
    
    '-------> Encabezado
    vg_db.Execute "UPDATE b_totventas SET tov_estdoc = 'A' WHERE tov_rutcli = '" & Trim(LimpiaDato(fpText1(1).text)) & "' " & _
                  "AND tov_tipdoc = 'SP' AND tov_numdoc = " & fpLongInteger1(0).Value & " AND tov_codbod = " & vg_codbod & ""
    
    '-------> Detalle
    If estdoc <> "P" And totdec > 0 Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i: numlin = i
           vaSpread1.Col = 1: codmer = Trim(LimpiaDato(vaSpread1.text))
           vaSpread1.Col = 5: canmer = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
           vaSpread1.Col = 8: color = Right(vaSpread1.text, 1)
           If color <> "I" Then '-------> No entra si es ingrediente
              
              If RS1.State = 1 Then RS1.Close
              RS1.CursorLocation = adUseClient
              vg_db.CursorLocation = adUseClient
              
              '-------> Control de Stock
              RS1.Open "SELECT * FROM b_productos WHERE pro_codigo = '" & codmer & "' AND pro_ctrsto = 1", vg_db, adOpenStatic
              If Not RS1.EOF Then
                 vg_db.Execute "UPDATE b_bodegas SET bod_canmer = bod_canmer+" & canmer & " WHERE bod_codpro = '" & codmer & "' AND bod_codbod = " & vg_codbod
              End If
              RS1.Close: Set RS1 = Nothing: numlin = numlin + 1
           
           End If
       
       Next i
    
    End If
    
    vg_db.CommitTrans
    '-------> Revisa Stock
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 8: color = Right(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), 1)
        If color <> "I" Then
           vaSpread1.Col = 1
           
           If RS1.State = 1 Then RS1.Close
           RS1.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
       
           
           RS1.Open "SELECT bod.bod_canmer from b_productos AS pro, b_bodegas AS bod " & _
                    "WHERE bod.bod_codpro = pro.pro_codigo " & _
                    "AND   pro.pro_ctrsto = 1 " & _
                    "AND   bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                    "AND   pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "'", vg_db, adOpenStatic
           
           vaSpread1.Col = 9
           If Not RS1.EOF Then vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
           RS1.Close: Set RS1 = Nothing
        
        End If
    
    Next i
    Label1.Caption = "ANULADA"
    Gl_Ac_Botones Me, 12, IIf(Label1.Caption = "ANULADA", 4, 3), ""
    modo = ""

Case 11 '-------> Busqueda
    
    If Trim(fpText1(1).text) = "" Then MsgBox "Debe seleccionar contrato...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    vg_codigo = Trim(fpText1(1).text)
    vg_nombre = "SP"
    B_SalBod.Show 1
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    Me.MousePointer = 11
    Me.Refresh
    Frame2.Enabled = False
    Frame1.Enabled = False
    vaSpread1.Col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    vaSpread1.MaxRows = 0
    vaSpread2.MaxRows = 0
    est = True
    
    If RS2.State = 1 Then RS2.Close
    RS2.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    '-------> Consultar si salida es resumido ó sector
    RS2.Open "SELECT DISTINCT dev_codsec FROM b_detventas WHERE  dev_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' AND dev_tipdoc = 'SP' AND dev_numdoc = " & Val(vg_codigo) & "", vg_db, adOpenStatic
    If RS2.EOF Then RS2.Close: Set RS2 = Nothing: MsgBox "No existe salida producción...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If IsNull(RS2!dev_codsec) Then Option1(0).Value = True: Option1(1).Value = False Else Option1(1).Value = False: Option1(1).Value = True
    RS2.Close: Set RS2 = Nothing
    If Option1(0).Value = True Then
       vaSpread1.Top = vaSpread2.Top
       vaSpread1.Height = 3165 + vaSpread2.Height
    Else
       vaSpread1.Top = 1755
       vaSpread1.Height = 3165
    End If
    est = False
    
    If RS2.State = 1 Then RS2.Close
    RS2.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    If Not vg_tipser Then
       
       RS2.Open "SELECT tov.tov_numdoc, tov.tov_fecemi, tov.tov_codbod, tov.tov_fecpro, tov.tov_codser, tov.tov_codreg, tov.tov_estdoc, ser.ser_nombre, reg.reg_nombre " & _
                "FROM  b_totventas tov, b_clientes cli, a_servicio ser, a_regimen reg " & _
                "WHERE tov.tov_rutcli = cli.cli_codigo AND ser.ser_codigo = tov.tov_codser " & _
                "AND   reg.reg_codigo = tov.tov_codreg AND tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
                "AND   tov.tov_tipdoc = 'SP' " & _
                "AND   tov.tov_numdoc = " & Val(vg_codigo) & " AND tov.tov_codbod = " & vg_codbod & "", vg_db, adOpenStatic
    Else
       
       RS2.Open "SELECT tov.tov_numdoc, tov.tov_fecemi, tov.tov_codbod, tov.tov_fecpro, tov.tov_codser, tov.tov_codreg, tov.tov_estdoc " & _
                "FROM  b_totventas tov, b_clientes cli " & _
                "WHERE tov.tov_rutcli = cli.cli_codigo " & _
                "AND   tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
                "AND   tov.tov_tipdoc = 'SP' " & _
                "AND   tov.tov_numdoc = " & Val(vg_codigo) & " AND tov.tov_codbod = " & vg_codbod & "", vg_db, adOpenStatic
    End If
    
    If Not RS2.EOF Then
        Do While Not RS2.EOF
            est = True
            fpLongInteger1(0).text = RS2!tov_numdoc
            fpDateTime1(0).text = RS2!tov_fecemi
            Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 10, fg_pone_cero(Str(RS2!tov_codbod), 10))
            fpDateTime1(1).text = RS2!tov_fecpro
            If Not vg_tipser Then
               fpLongInteger1(1).Value = RS2!tov_codreg
               fpLongInteger1(2).Value = RS2!tov_codser
               fpayuda(0).Caption = Trim(RS2!reg_nombre)
               fpayuda(3).Caption = Trim(RS2!ser_nombre)
'               Combo1(0).Clear
'               Combo1(0).AddItem RS2!reg_nombre & " - " & RS2!ser_nombre & Space(150) & "(" & fg_pone_cero(Str(RS2!tov_codreg), 10) & ")(" & fg_pone_cero(Str(RS2!tov_codser), 10) & ")"
'               Combo1(0).ListIndex = 0
            End If
            Label1.Caption = IIf(RS2!tov_estdoc = "", "", IIf(RS2!tov_estdoc = "A", "ANULADA", "PENDIENTE"))
            est = False
            RS2.MoveNext
        Loop
    End If
    RS2.Close: Set RS2 = Nothing
    codreg = 0: codser = 0
    If Not vg_tipser Then codreg = Val(fpLongInteger1(1).Value) 'Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 22, 10))
    If Not vg_tipser Then codser = Val(fpLongInteger1(2).Value) 'Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 10, 10))
    Fecha = Val(fpDateTime1(1).Year & Right("0" & fpDateTime1(1).Month, 2) & Right("0" & fpDateTime1(1).Day, 2))
    
    If RS2.State = 1 Then RS2.Close
    RS2.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    If vg_tipbase = "1" Then
       RS2.Open "SELECT DISTINCT tov_numdoc FROM b_totventas WHERE tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' AND tov_tipdoc = 'SP' AND tov_fecpro = cdate('" & fpDateTime1(1).text & "') AND tov_codreg=" & codreg & " AND tov_codser=" & codser & " AND tov_estdoc = 'P' AND tov_numdoc=" & Val(vg_codigo) & " AND tov_codbod=" & vg_codbod & "", vg_db, adOpenStatic
    Else
       RS2.Open "SELECT DISTINCT tov_numdoc FROM b_totventas WHERE tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' AND tov_tipdoc = 'SP' AND tov_fecpro = '" & Format(fpDateTime1(1).text, "yyyymmdd") & "' AND tov_codreg = " & codreg & " AND tov_codser = " & codser & " AND tov_estdoc = 'P' AND tov_numdoc = " & Val(vg_codigo) & " AND tov_codbod = " & vg_codbod & "", vg_db, adOpenStatic
    End If
    modo = ""
    If Not RS2.EOF Then
       fg_descarga
       Frame2.Enabled = True
       vg_codigo = RS2!tov_numdoc
       RS2.Close: Set RS2 = Nothing
       CargarSalidaPendientes
       modo = "M"
       Exit Sub
    Else
       RS2.Close: Set RS2 = Nothing
    End If
'                   "AND  (dev.dev_coding='' OR ISNULL(dev.dev_coding)) ORDER BY dev.dev_numlin", vg_db, adOpenStatic
    
    If RS3.State = 1 Then RS3.Close
    RS3.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    If Option1(0).Value = True Then
       sql1 = IIf(vg_tipbase = "1", " ORDER BY dev.dev_numlin ", "")
       RS3.Open "SELECT ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 as sec_codigo, '' as sec_nombre, 0 as sec_orden, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
                "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
                "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_unidad uni " & _
                "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
                "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_coding = ing.ing_codigo " & _
                "AND   ing.ing_unimed = unm.unm_codigo AND dev.dev_codmer = pro.pro_codigo " & _
                "AND   pro.pro_coduni = uni.uni_codigo " & _
                "AND   tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' AND tov.tov_numdoc = " & Val(vg_codigo) & " " & _
                "AND   tov.tov_tipdoc = 'SP' AND tov.tov_codbod = " & vg_codbod & " " & _
                "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin " & sql1 & " " & _
                "UNION ALL " & _
                "SELECT '' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' as unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 AS sec_codigo, '' AS sec_nombre, 0 AS sec_orden, 0 AS canmin, dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
                "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
                "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
                "AND   tov.tov_numdoc = dev.dev_numdoc AND  dev.dev_codmer = pro.pro_codigo " & _
                "AND   pro.pro_coduni = uni.uni_codigo AND  tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
                "AND   dev.dev_numdoc = " & Val(vg_codigo) & " AND  tov.tov_tipdoc = 'SP' AND tov.tov_codbod = " & vg_codbod & " " & _
                "AND  (dev.dev_coding = '' OR (dev.dev_coding) IS NULL OR dev.dev_codsec = -1) ORDER BY dev.dev_numlin", vg_db, adOpenStatic
    Else
       sql1 = IIf(vg_tipbase = "1", " ORDER BY sec.sec_orden, dev.dev_numlin ", "")
       RS3.Open "SELECT ing.ing_codigo, ing.ing_nombre,unm.unm_nomcor, sec.sec_codigo, sec.sec_nombre, sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
                "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
                "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_sector sec, a_unidad uni " & _
                "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
                "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_coding = ing.ing_codigo " & _
                "AND   ing.ing_unimed = unm.unm_codigo AND dev.dev_codmer = pro.pro_codigo " & _
                "AND   dev.dev_codsec = sec.sec_codigo AND pro.pro_coduni = uni.uni_codigo " & _
                "AND   tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' AND tov.tov_numdoc = " & Val(vg_codigo) & " " & _
                "AND   tov.tov_tipdoc = 'SP' AND tov.tov_codbod = " & vg_codbod & " " & _
                "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor,  sec.sec_codigo, sec.sec_nombre,  sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin " & sql1 & " " & _
                "UNION ALL " & _
                "SELECT '' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' as unm_nomcor, -1 AS sec_codigo, 'Estructura Fija' AS sec_nombre, 999999999 AS sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 AS canmin, dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
                "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
                "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
                "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_codmer = pro.pro_codigo " & _
                "AND   pro.pro_coduni = uni.uni_codigo AND  tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
                "AND   dev.dev_numdoc = " & Val(vg_codigo) & " AND  tov.tov_tipdoc = 'SP' AND tov.tov_codbod = " & vg_codbod & " " & _
                "AND  (dev.dev_coding = '' OR (dev.dev_coding) IS NULL OR dev.dev_codsec = -1) ORDER BY sec_orden, dev.dev_numlin", vg_db, adOpenStatic
    End If
    vaSpread1.Visible = False
    vaSpread2.Visible = False
    codsec = ""
    Do While Not RS3.EOF
        If codsec <> RS3!sec_codigo And Option1(1).Value = True Then
           vaSpread2.MaxRows = vaSpread2.MaxRows + 1
           vaSpread2.Row = vaSpread2.MaxRows
           vaSpread2.Col = 1: vaSpread2.Value = IIf(RS3!sec_codigo = -1, "estfij", RS3!sec_codigo)
           vaSpread2.Col = 2: vaSpread2.Value = Trim(RS3!sec_nombre)
           
           codsec = RS3!sec_codigo
           coding = 0
        End If
        '-------> Ingrediente
        If coding <> RS3!ing_codigo Then
           vaSpread1.MaxRows = vaSpread1.MaxRows + 1
           vaSpread1.Row = vaSpread1.MaxRows
           vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
           vaSpread1.RowHidden = IIf((vaSpread2.Row = 1 Or Option1(0).Value = True) And Check1(0).Value = 0, False, True)
           vaSpread1.Col = 1: vaSpread1.text = Trim(RS3!ing_codigo)
           vaSpread1.Col = 2: vaSpread1.text = Trim(RS3!ing_nombre)
           vaSpread1.Col = 3: vaSpread1.text = Trim(RS3!unm_nomcor)
           vaSpread1.Col = 4: vaSpread1.text = IIf(RS3!ing_codigo = "", "", Format(RS3!canmin, fg_Pict(9, vg_DCa)))
           vaSpread1.Col = 8: vaSpread1.text = "NI" '-------> No bloquedo - Ingrediente
           vaSpread1.Col = 10: vaSpread1.text = IIf(RS3!sec_codigo = -1, "estfij", RS3!sec_codigo)
           vaSpread1.Col = -1
           vaSpread1.FontBold = True
           vaSpread1.BackColor = Shape1(3).FillColor
           coding = RS3!ing_codigo
        End If
        '-------> Producto
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.RowHidden = IIf(vaSpread2.Row = 1 Or Option1(0).Value = True, False, True)
        vaSpread1.Col = 1: vaSpread1.text = Trim(RS3!pro_codigo)
        vaSpread1.Col = 2: vaSpread1.text = Trim(RS3!pro_nombre)
        vaSpread1.Col = 3: vaSpread1.text = Trim(RS3!uni_nomcor)
        vaSpread1.Col = 4: vaSpread1.text = Format(RS3!dev_canmin, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa))) 'vg_DCa))
        vaSpread1.Col = 5: vaSpread1.text = Format(RS3!dev_canmer, fg_Pict(9, vg_DCa))
        vaSpread1.Col = 6: vaSpread1.text = Format(RS3!dev_predoc, fg_Pict(9, vg_DPr))
        vaSpread1.Col = 7: vaSpread1.text = Format(RS3!dev_ptotal, fg_Pict(9, vg_DPr))
        vaSpread1.Col = 10: vaSpread1.text = IIf(RS3!sec_codigo = -1, "estfij", RS3!sec_codigo)
        vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(2).FillColor
        vaSpread1.Col = 9: vaSpread1.text = 0
        '-------> Mover sectores totales
        If vaSpread2.MaxRows > 0 And RS3!dev_canmer <> 0 Then vaSpread2.Col = 3: vaSpread2.TypeHAlign = TypeHAlignRight: vaSpread2.text = Format(IIf(Trim(vaSpread2.text) = "", 0, vaSpread2.Value) + (RS3!dev_ptotal), fg_Pict(9, vg_DPr))
        RS3.MoveNext
    Loop
    RS3.Close: Set RS3 = Nothing
    Me.MousePointer = 0
    vaSpread1.Visible = True
    vaSpread2.Visible = IIf(Option1(0).Value = True, False, True)
    vg_codigo = ""
    Frame5.Enabled = True: Check1(0).Enabled = True
    Gl_Ac_Botones Me, 12, IIf(Label1.Caption = "ANULADA", 4, 3), ""
    Toolbar1.Buttons(15).Enabled = False: Toolbar1.Buttons(15).ToolTipText = ""

Case 12 '-------> Imprimir
    
    If vaSpread1.MaxRows < 1 Then Exit Sub
    I_SalDevBod Me, "SP"

Case 17 '-------> Salir
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
If (Err = -2147467259 Or Err = -2147217900) And modo = "A" Then vg_db.RollbackTrans: GoTo paso
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Function MuestraFolio(Casino As String) As String

Dim sql1 As String
MuestraFolio = ""
If Trim(Casino) = "" Then Exit Function

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

sql1 = IIf(vg_tipbase = "1", " HOLDLOCK ", " WITH (HOLDLOCK) ")
RS1.Open "SELECT tov_numdoc FROM b_totventas " & sql1 & " WHERE tov_tipdoc = 'SP' AND tov_codbod = " & vg_codbod & " ORDER BY tov_numdoc DESC", vg_db, adOpenStatic
If Not RS1.EOF Then RS1.MoveFirst: MuestraFolio = RS1!tov_numdoc + 1 Else MuestraFolio = 1
RS1.Close: Set RS1 = Nothing

End Function

Sub Limpia(op As Integer)
est = True
If 0 = (fg_CambiaChar(GetParametro("salressec"), ";", "','")) Then Option1(0).Value = True: Option1(1).Value = False Else Option1(0).Value = False: Option1(1).Value = True
Label1.Caption = ""
Frame1.Enabled = True
fpDouble1(0) = Format(0, fg_Pict(0, vg_DCa))
fpDouble1(1) = Format(0, fg_Pict(0, vg_DCa))
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = ""
fpLongInteger1(1).text = ""
fpLongInteger1(2).text = ""
fpayuda(0).Caption = ""
fpayuda(3).Caption = ""
'Combo1(0).ListIndex = -1
Combo1(1).ListIndex = IIf(Combo1(1).listcount = 1, 0, -1)
vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0
vaSpread2.Row = -1: vaSpread2.Col = -1:
vaSpread2.BackColor = Shape1(2).FillColor
Frame2.Enabled = False
vaSpread1.Col = -1: vaSpread1.Row = -1
vaSpread1.Lock = True
vaSpread1.Col = 5: vaSpread1.Row = -1
vaSpread1.Lock = False
fpText1(1).Enabled = ModCasino
Image1(1).Enabled = ModCasino
fpText1(1).text = MuestraCasino(1)
fpayuda(1).Caption = MuestraCasino(2)
fpLongInteger1(0).text = TraerCorrelativo(vg_codbod, "SP") 'MuestraFolio(Trim(fpText1(1).text))
Gl_Ac_Botones Me, 12, op, ""
est = False
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long, indsec As Long, codpro As String, color As String, coding As String, texto As String, auxpro As String, auxing As String, codsec As String
Dim propon As Double
Select Case Button.Index
Case 1
    If Trim(Combo1(1).text) = "" Then MsgBox "Debe seleccionar bodega...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Toolbar2.Enabled = False
    If vaSpread2.MaxRows > 1 Then
       vaSpread2.Row = vaSpread2.ActiveRow
       vaSpread2.Col = 1
       codsec = vaSpread2.text
    Else
       codsec = "0"
    End If
    vg_nombre = "": vg_codigo = "": vg_bodega = 0: vg_bodega = Val(fg_codigocbo(Combo1, 1, 10, ""))
    vg_left = fpayuda(1).Left + 1920
    vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 2
    texto = ""
    If vaSpread1.MaxRows > 0 Then texto = Trim(Mid(Trim(vaSpread1.text), 1, IIf(InStr(1, vaSpread1.text, " ") > 0, _
                                  InStr(1, vaSpread1.text, " "), Len(Trim(vaSpread1.text)))))
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Pbo"
    B_TabEst.Text1.text = texto
    B_TabEst.Show 1
    If vg_codigo = "" Then Toolbar2.Enabled = True: Exit Sub
    Toolbar2.Enabled = True
    auxpro = vg_codigo: indsec = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        codpro = "": color = ""
        vaSpread1.Col = 1: codpro = Trim(vaSpread1.text)
        vaSpread1.Col = 8: color = Right(Trim(vaSpread1.text), 1)
        vaSpread1.Col = 10
        If Trim(codpro) = Trim(vg_codigo) And color <> "I" Then
           If Option1(0).Value = True Then
              MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           ElseIf Option1(1).Value = True And Trim(codpro) = Trim(vg_codigo) And color <> "I" And codsec = vaSpread1.text Then
              MsgBox "El producto ya existe sector...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           End If
        End If
        If codsec = vaSpread1.text And Option1(1).Value = True Then indsec = i
    Next i
    vaSpread1.Row = vaSpread1.ActiveRow
    '-------> validar si existe más de un ingrediente
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "SELECT COUNT(pri_coding) AS nreg FROM b_productosing WHERE pri_codpro = '" & vg_codigo & "'", vg_db, adOpenStatic
    If RS1.EOF Or IsNull(RS1!nreg) Or RS1!nreg = 0 Then RS1.Close: Set RS1 = Nothing: MsgBox "No hay ingrediente asignado al producto...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If RS1!nreg > 1 Then
       
       vg_nombre = ""
       vg_left = fpayuda(1).Left + 2300
       B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "Proing"
       B_TabEst.Show 1
       If vg_codigo = "" Then RS1.Close: Set RS1 = Nothing: Exit Sub
       auxing = vg_codigo
    
    ElseIf RS1!nreg = 1 Then
        
        If RS2.State = 1 Then RS2.Close
        RS2.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        RS2.Open "SELECT pri_coding FROM b_productosing WHERE pri_codpro = '" & vg_codigo & "'", vg_db, adOpenStatic
        If Not RS2.EOF Then auxing = RS2!pri_coding
        RS2.Close: Set RS2 = Nothing
    
    End If
    RS1.Close: Set RS1 = Nothing
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "SELECT pro.pro_codigo, pro.pro_nombre, ing.ing_codigo, uni.uni_nomcor, ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor " & _
             "FROM b_productos pro, a_unidad uni, b_productosing pri, b_ingrediente ing, a_unidadmed unm " & _
             "WHERE pro.pro_coduni = uni.uni_codigo " & _
             "AND   pri.pri_coding = ing.ing_codigo " & _
             "AND   pri.pri_codpro = pro.pro_codigo " & _
             "AND   ing.ing_unimed = unm.unm_codigo " & _
             "AND   pro.pro_codigo = '" & auxpro & "' " & _
             "AND   pro.pro_ctrsto = 1 AND ing.ing_codigo = '" & auxing & "'", vg_db, adOpenStatic
    
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            For i = 1 To vaSpread1.MaxRows
                vaSpread1.Row = i
                coding = "": color = ""
                vaSpread1.Col = 8: color = Right(Trim(vaSpread1.text), 1)
                If color = "I" Then
                    vaSpread1.Col = 1: coding = Trim(vaSpread1.text)
                    vaSpread1.Col = 10
                    If coding = RS1!ing_codigo Then
                       If Option1(0).Value = True Then
                          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
                          vaSpread1.Row = i + 1
                          vaSpread1.InsertRows vaSpread1.Row, 1
                          Exit For
                       ElseIf Option1(1).Value = True And codsec = vaSpread1.text Then
                          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
                          vaSpread1.Row = i + 1: indsec = vaSpread1.Row
                          vaSpread1.InsertRows vaSpread1.Row, 1
                          Exit For
                       End If
                    End If
                End If
            Next i
            If i > vaSpread1.MaxRows Then
               If codsec <> "estfij" Then
                  vaSpread1.MaxRows = vaSpread1.MaxRows + 2
                  vaSpread1.Row = IIf(Option1(1).Value = True, indsec + 1, i)
                  vaSpread1.InsertRows vaSpread1.Row, 2
                  vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
                  vaSpread1.Col = 1: vaSpread1.text = RS1!ing_codigo
                  vaSpread1.Col = 2: vaSpread1.text = RS1!ing_nombre
                  vaSpread1.Col = 3: vaSpread1.text = RS1!unm_nomcor
                  vaSpread1.Col = 4: vaSpread1.text = Format(0, fg_Pict(9, vg_DCa))
                  vaSpread1.Col = 8: vaSpread1.text = "NI" '-------> No bloquedo - Ingrediente
                  vaSpread1.Col = 10: vaSpread1.text = codsec
                  vaSpread1.Col = -1: vaSpread1.FontBold = True: vaSpread1.Lock = True
                  vaSpread1.BackColor = Shape1(3).FillColor
                  vaSpread1.Row = IIf(Option1(1).Value = True, indsec + 2, i + 1)
               ElseIf codsec = "estfij" Then
                  vaSpread1.MaxRows = vaSpread1.MaxRows + 1
                  vaSpread1.Row = IIf(Option1(1).Value = True, indsec + 1, i)
                  vaSpread1.InsertRows vaSpread1.Row, 1
                  vaSpread1.Row = IIf(Option1(1).Value = True, indsec + 1, i + 1)
               End If
            End If
                        
            vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(0).FillColor
            vaSpread1.Col = 1: vaSpread1.text = RS1!pro_codigo
            vaSpread1.Col = 2: vaSpread1.text = RS1!pro_nombre
            vaSpread1.Col = 3: vaSpread1.text = RS1!uni_nomcor
            vaSpread1.Col = 4: vaSpread1.text = Format(0, fg_Pict(9, vg_DCa))
            vaSpread1.Col = 5: vaSpread1.ForeColor = &HFF0000: vaSpread1.text = Format(0, fg_Pict(9, vg_DCa))
            '-------> Traer propon
            
            If RS2.State = 1 Then RS2.Close
            RS2.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
        
            propon = 0
            RS2.Open "SELECT TOP 1 ppd_cencos, ppd_codpro, ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                     "FROM  b_productospmpdia " & _
                     "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                     "AND   ppd_codpro = '" & RS1!pro_codigo & "' " & _
                     "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(fpDateTime1(1).text), "yyyymmdd") & " " & _
                     "GROUP BY ppd_cencos, ppd_codpro, ppd_propon " & _
                     "HAVING (ppd_propon) > 0 ORDER BY Max(ppd_fecdia) DESC", vg_db, adOpenStatic
            If Not RS2.EOF Then propon = RS2!ppd_propon
            RS2.Close: Set RS2 = Nothing
            vaSpread1.Col = 6: vaSpread1.text = Format(propon, fg_Pict(9, vg_DPr))
            vaSpread1.Col = 7: vaSpread1.text = Format(0, fg_Pict(9, vg_DPr))
            vaSpread1.Col = 8: vaSpread1.text = "NU" '-------> No bloquedo - Usuario agrega
            vaSpread1.Col = 10: vaSpread1.text = codsec
            '-------> Trae Stock
            If RS2.State = 1 Then RS2.Close
            RS2.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient

            
            RS2.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                     "WHERE bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                     "AND   bod.bod_codpro = pro.pro_codigo " & _
                     "AND   pro.pro_codigo = '" & Trim(RS1!pro_codigo) & "' " & _
                     "AND   pro.pro_ctrsto = 1", vg_db, adOpenStatic
            vaSpread1.Col = 9
            If Not RS2.EOF Then vaSpread1.text = Format(RS2!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
            RS2.Close: Set RS2 = Nothing
            RS1.MoveNext
            'i = i + 1
        Loop
    End If
    RS1.Close: Set RS1 = Nothing
    vaSpread1.Col = 5
    vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus 'vaSpread1.MaxRows
Case 2
    Dim pos1 As Long, pos2 As Long, candif As Long
    If vaSpread1.MaxRows = 0 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 8: color = Right(Trim(vaSpread1.text), 1)
    If color = "I" Then Exit Sub
    '-------> si cantidad planificada es > 0 no eliminar producto
    vaSpread1.Col = 4: If vaSpread1.text > 0 Then Exit Sub
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    For i = vaSpread1.ActiveRow To 1 Step -1
        color = ""
        vaSpread1.Row = i: vaSpread1.Col = 8: color = Right(Trim(vaSpread1.text), 1)
        If color = "I" Then pos1 = i: Exit For
    Next i
    For i = pos1 + 1 To vaSpread1.MaxRows
        color = ""
        vaSpread1.Row = i: vaSpread1.Col = 8: color = Right(Trim(vaSpread1.text), 1)
        pos2 = i
        If color = "I" Then pos2 = i - 1: Exit For
    Next i
    candif = pos2 - pos1
    
    vaSpread1.Row = pos1: vaSpread1.Col = 1
    vaSpread1.DeleteRows IIf(candif = 1, pos1, vaSpread1.ActiveRow), IIf(candif = 1, 2, 1)
    vaSpread1.MaxRows = vaSpread1.MaxRows - IIf(candif = 1, 2, 1)
End Select
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
Dim canrea As Double, canbod As Double, canaux As Double, propon As Double, codmer As String, color As Variant, i As Long
Dim color2 As String, codsec As String, auxsec As String, totcos As Double
codsec = "0": auxsec = "0"
If ChangeMade = False Then Exit Sub
If Not est1 Then Gl_Ac_Botones Me, 12, 6, ""
If vaSpread2.MaxRows > 0 Then vaSpread2.Row = vaSpread2.ActiveRow: vaSpread2.Col = 1: codsec = vaSpread2.text
vaSpread1.Row = Row
vaSpread1.Col = 1: codmer = vaSpread1.text
vaSpread1.Col = 5: canrea = Format(IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
vaSpread1.Col = 6: propon = Format(IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
vaSpread1.Col = 7: vaSpread1.text = Format(canrea * propon, fg_Pict(9, vg_DCa))
vaSpread1.Col = 8: color = Right(vaSpread1.text, 1)
vaSpread1.Col = 9: canbod = Format(IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
vaSpread1.Col = 10: auxsec = vaSpread1.text
If color <> "I" Then
    canaux = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
        vaSpread1.Col = 10: auxsec = vaSpread1.text
        vaSpread1.Col = 1
        If codmer = vaSpread1.text And color2 <> "I" And Option1(0).Value = True Then
            vaSpread1.Col = 5: canaux = canaux + Format(IIf(vaSpread1.text = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
        ElseIf codmer = vaSpread1.text And color2 <> "I" And Option1(1).Value = True Then
            vaSpread1.Col = 5: canaux = canaux + Format(IIf(vaSpread1.text = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
        End If
    Next i
    canrea = IIf(canaux > 0, canaux, canrea)
    vaSpread1.Row = Row
    If canbod - canrea >= 0 Then
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Row = i
            vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
            vaSpread1.Col = 10: auxsec = vaSpread1.text
            vaSpread1.Col = 1
            If codmer = vaSpread1.text And color2 <> "I" And Option1(0).Value = True Then
               vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(IIf(color = "M", 2, IIf(color = "U", 0, 3))).FillColor
               vaSpread1.Col = 8: vaSpread1.text = "N" & color '-------> No Bloqueado - Depende
            ElseIf codmer = vaSpread1.text And color2 <> "I" And Option1(1).Value = True Then
               vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(IIf(color = "M", 2, IIf(color = "U", 0, 3))).FillColor
               vaSpread1.Col = 8: vaSpread1.text = "N" & color '-------> No Bloqueado - Depende
            End If
        Next i
        Exit Sub
    End If
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
        vaSpread1.Col = 10: auxsec = vaSpread1.text
        vaSpread1.Col = 5: canrea = Format(IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
        vaSpread1.Col = 1
        If codmer = vaSpread1.text And color2 <> "I" And Option1(0) = True Then
           vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
           vaSpread1.Col = 8: vaSpread1.text = "S" & color '-------> Bloqueado - Depende
        ElseIf codmer = vaSpread1.text And color2 <> "I" And Option1(1).Value = True Then
           vaSpread1.Col = -1: vaSpread1.BackColor = IIf(canrea > 0, Shape1(1).FillColor, Shape1(2).FillColor)
           vaSpread1.Col = 8: vaSpread1.text = IIf(canrea > 0, "S" & color, "N" & color) '-------> Bloqueado - Depende
        End If
    Next i
End If
End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
Dim i As Long, color As String
If KeyAscii <> 13 Then Exit Sub
For i = vaSpread1.ActiveRow + 1 To vaSpread1.MaxRows
    vaSpread1.Row = i: vaSpread1.Col = 8: color = Right(vaSpread1.text, 1)
    If color <> "I" Then vaSpread1.SetActiveCell vaSpread1.ActiveCol, i - 1: Exit Sub
Next i
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim color As String, canrea As Double, propon As Double, codtot As Double
'-------> Calcular costo sectores
If vaSpread2.MaxRows < 1 Then Exit Sub
vaSpread2.Row = vaSpread2.ActiveRow: vaSpread2.Col = 1: codsec = vaSpread2.text
totcos = 0: canrea = 0: propon = 0
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 8: color = Right(vaSpread1.text, 1)
    vaSpread1.Col = 10
    If Trim(vaSpread1.text) = codsec And color <> "I" Then
       vaSpread1.Col = 5: canrea = Format(IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
       vaSpread1.Col = 6: propon = Format(IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
       totcos = totcos + (canrea * propon)
    End If
Next i
vaSpread2.Col = 3: vaSpread2.text = Format(totcos, fg_Pict(9, vg_DPr))
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If Row = 0 Then Exit Sub
Dim Stock As String, Nombre As String, color As String
vaSpread1.Row = Row
vaSpread1.Col = 8: color = Right(vaSpread1.text, 1)
If color = "I" Then Exit Sub
TipWidth = 4000
ShowTip = True
MultiLine = 2
vaSpread1.Col = 9: Stock = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
vaSpread1.Col = 2: Nombre = vaSpread1.text
TipText = "Bodega   : " & Trim(Left(Combo1(1).text, 50)) & vbCrLf & _
          "Producto : " & Trim(Nombre) & vbCrLf & _
          "Stock       : " & Format(Trim(Stock), fg_Pict(9, vg_DCa))
End Sub

Private Sub vaSpread2_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread2.MaxRows < 1 Or Row < 1 Then Exit Sub
Dim codsec As String, esthidden As Boolean
esthidden = True
vaSpread2.Row = vaSpread2.ActiveRow
vaSpread2.Col = 1: codsec = vaSpread2.text
vaSpread1.Visible = False
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 10
    If codsec = vaSpread1.text Then
       vaSpread1.Col = 5
       If Trim(vaSpread1.text) = "" Then vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False) Else vaSpread1.RowHidden = False
       If esthidden Then vaSpread1.Col = 1: vaSpread1.SetActiveCell vaSpread1.ActiveCol, vaSpread1.Row: esthidden = False
    Else
       vaSpread1.RowHidden = True
    End If
Next i
vaSpread1.Visible = True
End Sub

Private Sub vaSpread2_KeyUp(KeyCode As Integer, Shift As Integer)
vaSpread2_Click vaSpread2.ActiveCol, vaSpread2.ActiveRow
End Sub

Sub CargarSalidaPendientes()
Dim rutcli As String, tipdoc As String, NumDoc As Long, Fecha As Long, codbod  As Long, fecemi As Date, fecpro As Date, codreg As Long, codser As Long, i As Long, canact As Double, aAp  As String
Dim numlin As Long, codmer As String, coding As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String, total As Double, diablq As Date, color As String, codsec As String
Dim sql1 As String
If Trim(fpText1(1).text) = "" Then MsgBox "Debe seleccionar contrato...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
codreg = 0: codser = 0: fecpro = Date: est1 = False
If Val(fpLongInteger1(1).Value) > 0 Then codreg = Val(fpLongInteger1(1).Value) 'Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 22, 10))
If Val(fpLongInteger1(2).Value) > 0 Then codser = Val(fpLongInteger1(2).Value) 'Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 10, 10))
vg_codigo = Trim(fpText1(1).text)
vg_nombre = "SP"
If Trim(fpDateTime1(1).text) = "" Then fecpro = CDate(fpDateTime1(0).text) Else fecpro = CDate(fpDateTime1(1).text)

If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If vg_tipbase = "1" Then
   RS2.Open "SELECT DISTINCT tov_numdoc FROM b_totventas WHERE tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' AND tov_tipdoc = 'SP' AND tov_fecpro = cdate('" & fecpro & "') AND tov_codreg = " & codreg & " AND tov_codser = " & codser & " AND tov_estdoc = 'P' AND tov_codbod = " & vg_codbod & "", vg_db, adOpenStatic
Else
   RS2.Open "SELECT DISTINCT tov_numdoc FROM b_totventas WHERE tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' AND tov_tipdoc = 'SP' AND tov_fecpro = '" & Format(fecpro, "yyyymmdd") & "' AND tov_codreg = " & codreg & " AND tov_codser = " & codser & " AND tov_estdoc = 'P' AND tov_codbod = " & vg_codbod & "", vg_db, adOpenStatic
End If
If RS2.EOF Then RS2.Close: Set RS2 = Nothing: Exit Sub
fg_descarga
Frame2.Enabled = True
vg_codigo = RS2!tov_numdoc
RS2.Close: Set RS2 = Nothing
vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0
est = True
'-------> Consultar si salida es resumido ó sector
If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS2.Open "SELECT DISTINCT dev_codsec FROM b_detventas WHERE  dev_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' AND dev_tipdoc = 'SP' AND dev_numdoc = " & Val(vg_codigo) & "", vg_db, adOpenStatic
If RS2.EOF Then RS2.Close: Set RS2 = Nothing: MsgBox "No existe salida producción...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
If IsNull(RS2!dev_codsec) Then Option1(0).Value = True: Option1(1).Value = False Else Option1(1).Value = False: Option1(1).Value = True
RS2.Close: Set RS2 = Nothing
If Option1(0).Value = True Then
   vaSpread1.Top = vaSpread2.Top
   vaSpread1.Height = 3165 + vaSpread2.Height
Else
   vaSpread1.Top = 1755
   vaSpread1.Height = 3165
End If
est = False

If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Not vg_tipser Then
   RS2.Open "SELECT tov.tov_numdoc, tov.tov_fecemi, tov.tov_codbod, tov.tov_fecpro, tov.tov_codser, tov.tov_codreg, tov.tov_estdoc, ser.ser_nombre, reg.reg_nombre " & _
            "FROM  b_totventas tov, b_clientes cli, a_servicio ser, a_regimen reg " & _
            "WHERE tov.tov_rutcli = cli.cli_codigo AND ser.ser_codigo = tov.tov_codser " & _
            "AND   reg.reg_codigo = tov.tov_codreg AND tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
            "AND   tov.tov_tipdoc = 'SP' " & _
            "AND   tov.tov_numdoc = " & Val(vg_codigo) & " AND tov.tov_codbod = " & vg_codbod & "", vg_db, adOpenStatic
Else
   RS2.Open "SELECT tov.tov_numdoc, tov.tov_fecemi, tov.tov_codbod, tov.tov_fecpro, tov.tov_codser, tov.tov_codreg, tov.tov_estdoc " & _
            "FROM  b_totventas tov, b_clientes cli " & _
            "WHERE tov.tov_rutcli = cli.cli_codigo " & _
            "AND   tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
            "AND   tov.tov_tipdoc = 'SP' " & _
            "AND   tov.tov_numdoc = " & Val(vg_codigo) & " AND tov.tov_codbod = " & vg_codbod & "", vg_db, adOpenStatic
End If
If Not RS2.EOF Then
   Do While Not RS2.EOF
      est = True
      fpLongInteger1(0).text = RS2!tov_numdoc
      fpDateTime1(0).text = RS2!tov_fecemi
      Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 10, fg_pone_cero(Str(RS2!tov_codbod), 10))
      fpDateTime1(1).text = RS2!tov_fecpro
'      If Not vg_tipser Then
'         Combo1(0).Clear
'         Combo1(0).AddItem RS2!reg_nombre & " - " & RS2!ser_nombre & Space(150) & "(" & fg_pone_cero(Str(RS2!tov_codreg), 10) & ")(" & fg_pone_cero(Str(RS2!tov_codser), 10) & ")"
'         Combo1(0).ListIndex = 0
'      End If
      Label1.Caption = IIf(RS2!tov_estdoc = "", "", IIf(RS2!tov_estdoc = "A", "ANULADA", "PENDIENTE"))
      vg_codigo = RS2!tov_numdoc
      est = False
      RS2.MoveNext
   Loop
End If
RS2.Close: Set RS2 = Nothing
vaSpread1.Col = -1: vaSpread1.Row = -1
vaSpread1.Lock = False
Frame2.Enabled = True
Gl_Ac_Botones Me, 12, 3, ""

If RS3.State = 1 Then RS3.Close
RS3.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Option1(0).Value = True Then
   sql1 = IIf(vg_tipbase = "1", " ORDER BY dev.dev_numlin ", "")
   RS3.Open "SELECT ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 AS sec_codigo, '' AS sec_nombre, 0 AS sec_orden, " & _
            "(SELECT DISTINCT round(bod_canmer, " & vg_DCa & ") FROM b_bodegas WHERE tov.tov_codbod = bod_codbod AND pro.pro_codigo = bod_codpro AND bod_codbod = " & vg_codbod & ") AS bod_canmer, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
            "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
            "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_unidad uni " & _
            "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
            "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_coding = ing.ing_codigo " & _
            "AND   ing.ing_unimed = unm.unm_codigo AND dev.dev_codmer = pro.pro_codigo " & _
            "AND   pro.pro_coduni = uni.uni_codigo " & _
            "AND   tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' AND tov.tov_numdoc = " & Val(vg_codigo) & " " & _
            "AND   tov.tov_tipdoc = 'SP' AND tov.tov_codbod = " & vg_codbod & " " & _
            "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, tov.tov_codbod " & sql1 & " " & _
            "UNION ALL " & _
            "SELECT '' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' as unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 AS sec_codigo, '' AS sec_nombre, 0 AS sec_orden, " & _
            "(SELECT DISTINCT bod_canmer FROM b_bodegas WHERE  tov.tov_codbod=bod_codbod AND pro.pro_codigo=bod_codpro AND bod_codbod=" & vg_codbod & ") AS bod_canmer, 0 AS canmin, dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
            "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
            "WHERE tov.tov_rutcli=dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
            "AND   tov.tov_numdoc=dev.dev_numdoc AND dev.dev_codmer=pro.pro_codigo " & _
            "AND   pro.pro_coduni=uni.uni_codigo AND tov.tov_rutcli='" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
            "AND   dev.dev_numdoc=" & Val(vg_codigo) & " AND  tov.tov_tipdoc='SP' AND tov.tov_codbod=" & vg_codbod & " " & _
            "AND  (dev.dev_coding='' OR (dev.dev_coding) IS NULL OR dev.dev_codsec=-1) ORDER BY dev.dev_numlin", vg_db, adOpenStatic
Else
   sql1 = IIf(vg_tipbase = "1", " ORDER BY sec.sec_orden, dev.dev_numlin ", "")
   RS3.Open "SELECT ing.ing_codigo, ing.ing_nombre,unm.unm_nomcor, sec.sec_codigo, sec.sec_nombre, sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, " & _
            "(SELECT DISTINCT round(bod_canmer, " & vg_DCa & ") FROM b_bodegas WHERE tov.tov_codbod = bod_codbod AND pro.pro_codigo = bod_codpro AND bod_codbod = " & vg_codbod & ") AS bod_canmer, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
            "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
            "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_sector sec, a_unidad uni " & _
            "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
            "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_coding = ing.ing_codigo " & _
            "AND   ing.ing_unimed = unm.unm_codigo AND dev.dev_codmer = pro.pro_codigo " & _
            "AND   dev.dev_codsec = sec.sec_codigo AND pro.pro_coduni = uni.uni_codigo " & _
            "AND   tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' AND tov.tov_numdoc = " & Val(vg_codigo) & " " & _
            "AND   tov.tov_tipdoc = 'SP' AND tov.tov_codbod = " & vg_codbod & " " & _
            "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, sec.sec_codigo, sec.sec_nombre,  sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, tov.tov_codbod " & sql1 & " " & _
            "UNION ALL " & _
            "SELECT '' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' as unm_nomcor, -1 AS sec_codigo, 'Estructura Fija' AS sec_nombre, 999999999 AS sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, (SELECT DISTINCT bod_canmer FROM b_bodegas WHERE tov.tov_codbod=bod_codbod AND pro.pro_codigo=bod_codpro AND bod_codbod=" & vg_codbod & ") AS bod_canmer, 0 AS canmin,  dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
            "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
            "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
            "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_codmer = pro.pro_codigo " & _
            "AND   pro.pro_coduni = uni.uni_codigo AND tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
            "AND   dev.dev_numdoc = " & Val(vg_codigo) & " AND  tov.tov_tipdoc = 'SP' AND tov.tov_codbod = " & vg_codbod & "  " & _
            "AND  (dev.dev_coding = '' OR (dev.dev_coding) IS NULL OR dev.dev_codsec=-1) ORDER BY sec_orden, dev.dev_numlin", vg_db, adOpenStatic
End If
vaSpread1.Visible = False
vaSpread2.Visible = False
codsec = ""
Do While Not RS3.EOF
   If codsec <> RS3!sec_codigo And Option1(1).Value = True Then
      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      vaSpread2.Col = 1: vaSpread2.Value = IIf(RS3!sec_codigo = -1, "estfij", RS3!sec_codigo)
      vaSpread2.Col = 2: vaSpread2.Value = Trim(RS3!sec_nombre)
      codsec = RS3!sec_codigo
      coding = 0
   End If
   '-------> Ingrediente
   If coding <> RS3!ing_codigo Then
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
      vaSpread1.RowHidden = IIf((vaSpread2.Row = 1 Or Option1(0).Value = True) And Check1(0).Value = 0, False, True)
      vaSpread1.Col = 1: vaSpread1.Lock = True: vaSpread1.text = Trim(RS3!ing_codigo)
      vaSpread1.Col = 2: vaSpread1.Lock = True: vaSpread1.text = Trim(RS3!ing_nombre)
      vaSpread1.Col = 3: vaSpread1.Lock = True: vaSpread1.text = Trim(RS3!unm_nomcor)
      vaSpread1.Col = 4: vaSpread1.Lock = True: vaSpread1.text = IIf(RS3!ing_codigo = "", "", Format(RS3!canmin, fg_Pict(9, vg_DCa)))
      vaSpread1.Col = 5: vaSpread1.Lock = True: vaSpread1.text = ""
      vaSpread1.Col = 6: vaSpread1.Lock = True: vaSpread1.text = ""
      vaSpread1.Col = 7: vaSpread1.Lock = True: vaSpread1.text = ""
      vaSpread1.Col = 8: vaSpread1.text = "NI" 'No bloquedo - Ingrediente
      vaSpread1.Col = 10: vaSpread1.text = IIf(RS3!sec_codigo = -1, "estfij", RS3!sec_codigo)
      vaSpread1.Col = -1
      vaSpread1.FontBold = True
      vaSpread1.BackColor = Shape1(3).FillColor
      coding = RS3!ing_codigo
   End If
   '-------> Producto
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.RowHidden = IIf(vaSpread2.Row = 1 Or Option1(0).Value = True, False, True)
   vaSpread1.Col = 1: vaSpread1.Lock = True: vaSpread1.text = Trim(RS3!pro_codigo)
   vaSpread1.Col = 2: vaSpread1.Lock = True: vaSpread1.text = Trim(RS3!pro_nombre)
   vaSpread1.Col = 3: vaSpread1.Lock = True: vaSpread1.text = Trim(RS3!uni_nomcor)
   vaSpread1.Col = 4: vaSpread1.Lock = True: vaSpread1.text = Format(RS3!dev_canmin, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa))) 'vg_DCa))
   vaSpread1.Col = 5: vaSpread1.ForeColor = &HFF0000: vaSpread1.text = Format(RS3!dev_canmer, fg_Pict(9, vg_DCa))
   vaSpread1.Col = 6: vaSpread1.Lock = True: vaSpread1.text = Format(RS3!dev_predoc, fg_Pict(9, vg_DPr))
   vaSpread1.Col = 7: vaSpread1.Lock = True: vaSpread1.text = Format(RS3!dev_ptotal, fg_Pict(9, vg_DPr))
   vaSpread1.Col = 8: vaSpread1.text = "NM" 'No bloquedo - Viene de la Minuta
   vaSpread1.Col = 10: vaSpread1.text = IIf(RS3!sec_codigo = -1, "estfij", RS3!sec_codigo)
   vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(2).FillColor
   vaSpread1.Col = 9: vaSpread1.text = Format(RS3!bod_canmer, fg_Pict(9, vg_DCa))
   '-------> Mover sectores totales
   If vaSpread2.MaxRows > 0 And RS3!dev_canmer <> 0 Then vaSpread2.Col = 3: vaSpread2.TypeHAlign = TypeHAlignRight: vaSpread2.text = Format(IIf(Trim(vaSpread2.text) = "", 0, vaSpread2.Value) + (RS3!dev_ptotal), fg_Pict(9, vg_DPr))
   '-------> Revisa color
   vaSpread1.Col = 5: canrea = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
   vaSpread1.Col = 8: color = Right(vaSpread1.text, 1)
   vaSpread1.Col = 9: canbod = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
        
   canaux = 0
   For z = 1 To vaSpread1.MaxRows
       vaSpread1.Row = z
       vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
       vaSpread1.Col = 1
       If RS3!pro_codigo = vaSpread1.text And color2 <> "I" Then
          vaSpread1.Col = 5: canaux = canaux + Format(IIf(vaSpread1.text = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
       End If
   Next z
   canrea = IIf(canaux > 0, canaux, canrea)
        
   If canbod - canrea >= 0 Then
      For z = 1 To vaSpread1.MaxRows
          vaSpread1.Row = z
          vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
          vaSpread1.Col = 1
          If RS3!pro_codigo = vaSpread1.text And color2 <> "I" Then
             vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(IIf(color = "M", 2, IIf(color = "U", 0, 3))).FillColor
             vaSpread1.Col = 8: vaSpread1.text = "N" & color 'No Bloqueado - Depende
          End If
      Next z
   Else
      For z = 1 To vaSpread1.MaxRows
          vaSpread1.Row = z
          vaSpread1.Col = 8: color2 = Right(vaSpread1.text, 1)
          vaSpread1.Col = 1
          If RS3!pro_codigo = vaSpread1.text And color2 <> "I" Then
             vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
             vaSpread1.Col = 8: vaSpread1.text = "S" & color 'Bloqueado - Depende
          End If
      Next z
   End If
   RS3.MoveNext
Loop
RS3.Close: Set RS3 = Nothing
Me.MousePointer = 0
Frame1.Enabled = False
vaSpread1.Visible = True
vaSpread2.Visible = IIf(Option1(0).Value = True, False, True)
vg_codigo = ""
Frame5.Enabled = True: Check1(0).Enabled = True
End Sub
