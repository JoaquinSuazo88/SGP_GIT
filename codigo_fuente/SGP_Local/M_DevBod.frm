VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_DevBod 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolución de Producción para Bodega"
   ClientHeight    =   7320
   ClientLeft      =   1740
   ClientTop       =   1950
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   345
      Left            =   9000
      TabIndex        =   31
      Top             =   1200
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
         TabIndex        =   32
         Top             =   120
         Width           =   2010
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   11460
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   810
         Width           =   3885
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   6390
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   470
         Width           =   2325
      End
      Begin VB.Frame Frame4 
         Height          =   45
         Left            =   30
         TabIndex        =   8
         Top             =   2265
         Visible         =   0   'False
         Width           =   8520
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
         Left            =   9000
         TabIndex        =   7
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
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
         Left            =   10320
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   1350
         TabIndex        =   10
         Top             =   120
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
         Left            =   8670
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   150
         Width           =   1515
         _Version        =   196608
         _ExtentX        =   2672
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
         Left            =   1350
         TabIndex        =   13
         Top             =   470
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
         Left            =   3990
         TabIndex        =   14
         Top             =   470
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
         TabIndex        =   15
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
         TabIndex        =   16
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   6450
         TabIndex        =   29
         Top             =   510
         Width           =   2310
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   1395
         TabIndex        =   27
         Top             =   885
         Width           =   3885
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3165
         TabIndex        =   26
         Top             =   135
         Width           =   3975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Rég. - Serv."
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
         TabIndex        =   25
         Top             =   900
         Width           =   1050
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
         TabIndex        =   24
         Top             =   520
         Width           =   1050
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
         TabIndex        =   23
         Top             =   520
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
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
         Left            =   7395
         TabIndex        =   22
         Top             =   195
         Width           =   1365
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2715
         Picture         =   "M_DevBod.frx":0000
         Top             =   30
         Width           =   480
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
         TabIndex        =   21
         Top             =   195
         Width           =   735
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
         Left            =   5595
         TabIndex        =   20
         Top             =   520
         Width           =   660
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
         Left            =   10290
         TabIndex        =   19
         Top             =   210
         Width           =   1095
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
         TabIndex        =   18
         Top             =   2310
         Visible         =   0   'False
         Width           =   1650
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
         TabIndex        =   17
         Top             =   2310
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3210
         TabIndex        =   28
         Top             =   165
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5730
      Left            =   15
      TabIndex        =   1
      Top             =   1590
      Width           =   11460
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3645
         Left            =   105
         TabIndex        =   2
         Top             =   1785
         Width           =   11295
         _Version        =   393216
         _ExtentX        =   19923
         _ExtentY        =   6429
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
         SpreadDesigner  =   "M_DevBod.frx":030A
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   1455
         Left            =   105
         TabIndex        =   30
         Top             =   240
         Width           =   11295
         _Version        =   393216
         _ExtentX        =   19923
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
         SpreadDesigner  =   "M_DevBod.frx":0A4F
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ingrediente"
         Height          =   195
         Index           =   3
         Left            =   555
         TabIndex        =   4
         Top             =   5490
         Width           =   795
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   165
         Top             =   5520
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   1635
         Top             =   5520
         Width           =   300
      End
      Begin VB.Label Label5 
         Caption         =   "Producto"
         Height          =   210
         Index           =   2
         Left            =   1995
         TabIndex        =   3
         Top             =   5490
         Width           =   915
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_DevBod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset
Dim est As Boolean

Private Sub Check1_Click(Index As Integer)
Dim codsec As String
On Error GoTo Man_Error
'------- Actualizar parametro devolución producción
vg_db.BeginTrans
vg_db.Execute "UPDATE a_param SET par_valor = '" & IIf(Check1(0).Value = 1, 1, 0) & "' WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'ingdevpro'"
vg_db.CommitTrans
If vaSpread1.MaxRows < 1 And Option1(1).Value = True Then Exit Sub
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

Private Sub Combo1_Click(Index As Integer)
Dim feprod As Long, codser As Long, fil As Long, codreg As Long, aAp As String, codsec As String, coding As String, opcsal As String
Dim sql1 As String, sql2 As String
If est Then Exit Sub
Select Case Index
Case 0
    '-------> Validar si el contrato tiene asignado inventario rotativo
    If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
       Format(fpDateTime1(1).Value, "dd/mm/yyyy") > CDate(vg_ciedia) Then MsgBox "Tiene que realizar cierre diario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Combo1(0).ListIndex = -1 Or Combo1(0).text = "" Then Exit Sub
    codreg = Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 22, 10))
    codser = Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 10, 10))
    opcsal = Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 2, 1))
    sql1 = IIf(vg_tipbase = "1", " CDate('" & fpDateTime1(1).text & "') ", " '" & Format(fpDateTime1(1).text, "yyyymmdd") & "' ")
    RS3.Open "SELECT tov_numdoc FROM b_totventas where tov_rutcli = '" & Trim(LimpiaDato(fpText1(1).text)) & "' AND tov_tipdoc = 'DP' " & _
             "AND tov_fecpro = " & sql1 & " AND tov_codbod = " & vg_codbod & " AND tov_estdoc <> 'A' AND tov_estdoc <> 'P' " & _
             "AND tov_codreg = " & codreg & " AND tov_codser = " & codser & "", vg_db, adOpenStatic
    If Not RS3.EOF Then
        MsgBox "Devolución ya realizada...", vbExclamation + vbOKOnly, MsgTitulo
        DevExiste RS3!tov_numdoc
        RS3.Close: Set RS3 = Nothing
        Exit Sub
    End If
    RS3.Close: Set RS3 = Nothing
    Me.MousePointer = 11
    Gl_Ac_Botones Me, 4, 6, ""
    If opcsal = "0" Then
       Option1(0).Value = True
       Option1(1).Value = False
    ElseIf opcsal = "1" Then
       Option1(1).Value = True
       Option1(0).Value = False
    End If
    If Option1(0).Value = True Then
       sql1 = IIf(vg_tipbase = "1", " CDate('" & fpDateTime1(1).text & "') ", " '" & Format(fpDateTime1(1).text, "yyyymmdd") & "' ")
       sql2 = IIf(vg_tipbase = "1", " ORDER BY dev.dev_numlin ", "")
       RS3.Open "SELECT ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 as sec_codigo, '' as sec_nombre, 0 as sec_orden, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
                "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
                "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_unidad uni " & _
                "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
                "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_coding=ing.ing_codigo " & _
                "AND   ing.ing_unimed = unm.unm_codigo AND dev.dev_codmer=pro.pro_codigo " & _
                "AND   pro.pro_coduni = uni.uni_codigo " & _
                "AND   tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
                "AND   tov.tov_fecpro = " & sql1 & " " & _
                "AND   tov.tov_codser = " & codser & " " & _
                "AND   tov.tov_codreg = " & codreg & " " & _
                "AND   tov.tov_codbod = " & vg_codbod & " AND tov.tov_tipdoc = 'SP' AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' AND dev.dev_canmer <> 0 " & _
                "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin " & sql2 & " " & _
                "UNION ALL " & _
                "SELECT '' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' as unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 AS sec_codigo, '' AS sec_nombre, 0 AS sec_orden, 0 AS canmin,  dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
                "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
                "WHERE tov.tov_rutcli = dev.dev_rutcli AND  tov.tov_tipdoc = dev.dev_tipdoc " & _
                "AND   tov.tov_numdoc = dev.dev_numdoc AND  dev.dev_codmer = pro.pro_codigo " & _
                "AND   pro.pro_coduni = uni.uni_codigo AND  tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
                "AND   tov.tov_fecpro = " & sql1 & " " & _
                "AND   tov.tov_codser = " & codser & " " & _
                "AND   tov.tov_codreg = " & codreg & " " & _
                "AND   tov.tov_tipdoc = 'SP'  " & _
                "AND   tov.tov_codbod = " & vg_codbod & " AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' AND (dev.dev_coding = '' OR (dev.dev_coding) IS NULL OR dev.dev_codsec = -1) AND dev.dev_canmer<>0 ORDER BY ing_codigo, dev.dev_numlin", vg_db, adOpenStatic
    ElseIf Option1(1).Value = True Then
       sql1 = IIf(vg_tipbase = "1", " CDate('" & fpDateTime1(1).text & "') ", " '" & Format(fpDateTime1(1).text, "yyyymmdd") & "' ")
       sql2 = IIf(vg_tipbase = "1", " ORDER BY sec.sec_orden, dev.dev_numlin ", "")
       RS3.Open "SELECT ing.ing_codigo, ing.ing_nombre,unm.unm_nomcor, sec.sec_codigo, sec.sec_nombre, sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
                "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
                "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_sector sec, a_unidad uni " & _
                "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
                "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_coding=ing.ing_codigo " & _
                "AND   ing.ing_unimed = unm.unm_codigo AND dev.dev_codmer=pro.pro_codigo " & _
                "AND   dev.dev_codsec = sec.sec_codigo AND pro.pro_coduni=uni.uni_codigo " & _
                "AND   tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
                "AND   tov.tov_fecpro = " & sql1 & " " & _
                "AND   tov.tov_codser = " & codser & " " & _
                "AND   tov.tov_codreg = " & codreg & " " & _
                "AND   tov.tov_codbod = " & vg_codbod & " AND tov.tov_tipdoc = 'SP' AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' AND dev.dev_canmer <> 0 " & _
                "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor,  sec.sec_codigo, sec.sec_nombre,  sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin " & sql2 & " " & _
                "UNION ALL " & _
                "SELECT '' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' as unm_nomcor, -1 AS sec_codigo, 'Estructura Fija' AS sec_nombre, 999999999 AS sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 AS canmin,  dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
                "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
                "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
                "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_codmer=pro.pro_codigo " & _
                "AND   pro.pro_coduni = uni.uni_codigo AND tov.tov_rutcli='" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
                "AND   tov.tov_fecpro = " & sql1 & " " & _
                "AND   tov.tov_codser = " & codser & " " & _
                "AND   tov.tov_codreg = " & codreg & " " & _
                "AND   tov.tov_codbod = " & vg_codbod & " AND tov.tov_tipdoc='SP'  " & _
                "AND   tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' AND (dev.dev_coding = '' OR (dev.dev_coding) IS NULL OR dev.dev_codsec = -1) AND dev.dev_canmer <> 0 ORDER BY sec_orden, dev.dev_numlin", vg_db, adOpenStatic
    End If
   If RS3.EOF Then
        RS3.Close: Set RS3 = Nothing
        MsgBox "No existe salida a producción...", vbExclamation + vbOKOnly, MsgTitulo
        Me.MousePointer = 0
        Exit Sub
    End If
    vaSpread1.Visible = False: vaSpread1.MaxRows = 0
    vaSpread2.Visible = False
    i = 0: codsec = "0": coding = ""
    Do While Not RS3.EOF
       '------ Sector
       If codsec <> RS3!sec_codigo Then
          vaSpread2.MaxRows = vaSpread2.MaxRows + 1
          vaSpread2.Row = vaSpread2.MaxRows
          vaSpread2.Col = 1: vaSpread2.Value = IIf(RS3!sec_codigo = -1, "estfij", RS3!sec_codigo)
          vaSpread2.Col = 2: vaSpread2.Value = Trim(RS3!sec_nombre)
          codsec = RS3!sec_codigo
       End If
       '------ Ingrediente
       If coding <> RS3!ing_codigo Then
          i = i + 1: vaSpread1.MaxRows = i
          vaSpread1.Row = i
          vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
          vaSpread1.RowHidden = IIf((vaSpread2.Row = 1 Or Option1(0).Value = True) And Check1(0).Value = 0, False, True)
          vaSpread1.Col = 1: vaSpread1.text = RS3!ing_codigo
          vaSpread1.Col = 2: vaSpread1.text = RS3!ing_nombre
          vaSpread1.Col = 3: vaSpread1.text = RS3!unm_nomcor
          vaSpread1.Col = 4: vaSpread1.text = IIf(RS3!ing_codigo = "", "", Format(RS3!canmin, fg_Pict(9, vg_DCa))) 'RS3!canmer
          vaSpread1.Col = 8: vaSpread1.text = "NI" 'No bloquedo - Ingrediente
          vaSpread1.Col = 10: vaSpread1.text = IIf(RS3!sec_codigo = -1, "estfij", RS3!sec_codigo)
          vaSpread1.Col = -1
          vaSpread1.FontBold = True: vaSpread1.Lock = True
          vaSpread1.BackColor = Shape1(0).FillColor
          coding = RS3!ing_codigo
       End If
       '------- Productos
       i = i + 1: vaSpread1.MaxRows = i
       vaSpread1.Row = i
       vaSpread1.RowHidden = IIf(vaSpread2.Row = 1 Or Option1(0).Value = True, False, True)
       vaSpread1.Col = 1: vaSpread1.text = RS3!pro_codigo
       vaSpread1.Col = 2: vaSpread1.text = RS3!pro_nombre
       vaSpread1.Col = 3: vaSpread1.text = RS3!uni_nomcor
       vaSpread1.Col = 4: vaSpread1.text = Format(RS3!dev_canmer, fg_Pict(9, vg_DCa))
       vaSpread1.Col = 5: vaSpread1.ForeColor = &HFF0000: vaSpread1.text = Format(0, fg_Pict(9, vg_DCa))
       vaSpread1.Col = 6: vaSpread1.text = Format(RS3!dev_predoc, fg_Pict(9, vg_DPr))
       vaSpread1.Col = 7: vaSpread1.text = Format(0, fg_Pict(9, vg_DPr))
       vaSpread1.Col = 8: vaSpread1.text = "NP" 'No bloquedo - Producto
       vaSpread1.Col = 10: vaSpread1.text = IIf(RS3!sec_codigo = -1, "estfij", RS3!sec_codigo)
       vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
       '------- Trae Stock
       RS2.Open "SELECT bod.bod_canmer FROM b_productos pro, b_bodegas bod " & _
                "WHERE  bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                "AND    bod.bod_codpro = pro.pro_codigo AND pro.pro_codigo = '" & Trim(RS3!pro_codigo) & "'", vg_db, adOpenStatic
       vaSpread1.Col = 9
       If Not RS2.EOF Then vaSpread1.text = Format(RS2!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
       RS2.Close: Set RS2 = Nothing
       RS3.MoveNext
    Loop
    RS3.Close: Set RS3 = Nothing
    Me.MousePointer = 0
    Frame1.Enabled = False
    vaSpread2.Visible = True
    vaSpread1.Visible = True
    If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
Case 1
    If vaSpread1.MaxRows = 0 Then Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        RS2.Open "SELECT bod.bod_canmer FROM b_productos pro, b_bodegas bod WHERE bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "AND    bod.bod_codpro = pro.pro_codigo AND pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "'", vg_db, adOpenStatic
        vaSpread1.Col = 9
        If Not RS2.EOF Then vaSpread1.text = Format(RS2!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
        RS2.Close: Set RS2 = Nothing
    Next i
End Select
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 7830
Me.Width = 11625
fg_centra Me
est = False
Me.HelpContextID = vg_OpcM
MsgTitulo = "Salida Producción"
EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)
Dim X As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
Gl_Mo_Botones Me, 4
vaSpread1.Row = -1
vaSpread1.Col = 4: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 6: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
vaSpread1.Col = 7: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
'-------> Cargar Combo Bodega
CargarDatoCombo Combo1, 1, "b_clientes", "cli_", "CliBod", "N"
Check1(0).Value = IIf(0 = (fg_CambiaChar(GetParametro("ingdevpro"), ";", "','")), 0, 1)
If 0 = (fg_CambiaChar(GetParametro("salressec"), ";", "','")) Then
   vaSpread2.Visible = False
   vaSpread1.Top = vaSpread2.Top
   vaSpread1.Height = vaSpread1.Height + vaSpread2.Height
   Option1(0).Value = True: Option1(1).Value = False
Else
   vaSpread2.Visible = True
   vaSpread1.Top = 1785
   vaSpread1.Height = 3645
   Option1(0).Value = False: Option1(1).Value = True
End If
'Limpia
Limpia 2
End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then
    Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)
    Frame1.Left = (Me.Width \ 2) - (Frame1.Width \ 2)
    Frame5.Left = 10550
ElseIf Me.WindowState = 0 Then
    Frame2.Left = 0 '15
    Frame1.Left = 0 '825
    Frame5.Left = 9000 '6960
End If
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 1
    Combo1(0).Clear
    vaSpread1.MaxRows = 0
End Select
End Sub

Private Sub fpDateTime1_GotFocus(Index As Integer)
Select Case Index
Case 1
    Toolbar1.Buttons(8).Enabled = False
End Select
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpDateTime1_LostFocus(Index As Integer)
Dim Tipo As String, sql1 As String, sql2 As String
Select Case Index
Case 1
    Toolbar1.Buttons(8).Enabled = True
    If Trim(fpDateTime1(1).text) = "" Or Trim(fpText1(1).text) = "" Then Exit Sub
    sql1 = IIf(vg_tipbase = "1", " iif(isnull(dev.dev_codsec),'Resumido','Sector') AS sec_nombre ", " CASE WHEN (dev.dev_codsec) is null THEN 'Resumido' ELSE 'Sector' END AS sec_nombre ")
    sql2 = IIf(vg_tipbase = "1", " CDate('" & fpDateTime1(1).text & "') ", " '" & Format(fpDateTime1(1).text, "yyyymmdd") & "' ")
    RS1.Open "SELECT DISTINCT tov.tov_codser, ser.ser_nombre,tov.tov_codreg, reg.reg_nombre, " & sql1 & " " & _
             "FROM    b_totventas tov, b_detventas dev, a_servicio ser, a_regimen reg " & _
             "WHERE   tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
             "AND     tov.tov_numdoc = dev.dev_numdoc AND tov.tov_codser = ser.ser_codigo and tov.tov_codreg = reg.reg_codigo " & _
             "AND     tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
             "AND     tov.tov_tipdoc = 'SP' " & _
             "AND     tov.tov_fecpro = " & sql2 & " " & _
             "AND     tov.tov_codbod = " & vg_codbod & " AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P'", vg_db, adOpenStatic
    Combo1(0).Clear
    Do While Not RS1.EOF
        Combo1(0).AddItem RS1!reg_nombre & " - " & RS1!ser_nombre & " - " & RS1!sec_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!tov_codreg), 10) & ")(" & fg_pone_cero(Str(RS1!tov_codser), 10) & ")(" & IIf(RS1!sec_nombre = "Sector", "1", "0") & ")"""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    If Combo1(0).listcount = 0 Then MsgBox "No existe salida a producción...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
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
RS1.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo='" & fpText1(1).text & "' AND cli_tipo=0", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      fpayuda(Index).Caption = RS1!cli_nombre
      Gl_Ac_Botones Me, 4, 2, ""
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
'fpLongInteger1(0).text = MuestraFolio(Trim(fpText1(1).text))
fpLongInteger1(0).text = TraerCorrelativo(vg_codbod, "DP")
End Sub

Private Sub Image1_Click(Index As Integer)
vg_codigo = 0
Select Case Index
Case 1
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
    Gl_Ac_Botones Me, 4, 2, ""
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
On Error GoTo Man_Error
If Option1(0).Value = True Then
   vaSpread2.Visible = False
   vaSpread1.Top = vaSpread2.Top
   vaSpread1.Height = 3645 + vaSpread2.Height
Else
   vaSpread2.Visible = True
   vaSpread1.Top = 1785
   vaSpread1.Height = 3645
End If
If est Then Exit Sub
'------- Actualizar parametro devolución producción
vg_db.BeginTrans
vg_db.Execute "UPDATE a_param SET par_valor = '" & IIf(Option1(0).Value = True, 0, 1) & "' WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'salressec'"
vg_db.CommitTrans
Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim rutcli As String, tipdoc As String, NumDoc As Long, codbod   As Long, fecemi As Date, fecpro As Date, codreg As Long, codser As Long, i As Long, canact As Double
Dim numlin As Long, codmer As String, coding As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String, diablq As Date, color As String, codsec As String
On Error GoTo Man_Error
codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
fecpro = Format(fpDateTime1(1).Value, "dd/mm/yyyy")
If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
TraerFechaCierre
Select Case Button.Index
Case 1, 6 '-------> Nuevo
'    Limpia
    If Button.Index = 6 And vaSpread1.MaxRows > 0 Then If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    Limpia IIf(Button.Index = 1, 6, 2)
    If fpText1(1).Enabled = True Then fpText1(1).SetFocus
Case 8 '-------> Graba
    If Trim(fpText1(1).text) = "" Or Trim(fpLongInteger1(0).text) = "" Or Trim(Combo1(0).text) = "" Or Trim(fpDateTime1(0).text) = "" _
    Or Trim(Combo1(1).text) = "" Or Trim(fpDateTime1(1).text) = "" Then MsgBox "Debe ingresar dato importante...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
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
    
       MsgBox "Día se encuentra cerrado, no es posible ingresar...", vbExclamation + vbQuestion, MsgTitulo: Exit Sub
    
    End If
    
    total = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 7: ptotal = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DPr)
        total = total + ptotal
    Next i
    If total = 0 Then MsgBox "El total del documento debe ser mayor a 0...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
paso:
    rutcli = Trim(LimpiaDato(fpText1(1).text))
    tipdoc = "DP"
'    fpLongInteger1(0).text = MuestraFolio(Trim(fpText1(1).text))
'    numdoc = fpLongInteger1(0).text
    codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
    fecemi = Format(fpDateTime1(0).text, "dd/mm/yyyy")
    fecpro = Format(fpDateTime1(1).text, "dd/mm/yyyy")
    codreg = Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 22, 10))
    codser = Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 10, 10))
    NumDoc = TraerCorrelativo(codbod, "DP")
'    vg_db.BeginTrans
    vg_db.Execute "UPDATE b_parametros SET par_correlativo = " & NumDoc & " WHERE par_codbod = " & codbod & " AND par_tipdoc = 'DP'"
'    vg_db.CommitTrans
    fpLongInteger1(0).text = NumDoc
    DoEvents
    
    vg_db.BeginTrans
    '------- Encabezado
    If vg_tipbase = "1" Then
       vg_db.Execute "INSERT INTO b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_numinf) " & _
                     "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & codbod & ", CDate('" & _
                  Format(fpDateTime1(0).text, "dd/mm/yyyy") & "'), CDate('" & Format(fpDateTime1(1).text, "dd/mm/yyyy") & "'), " & codreg & ", " & codser & ", 0, '', '', 0)"
    Else
       vg_db.Execute "INSERT INTO b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_numinf) " & _
                     "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & codbod & ", '" & _
                  Format(fpDateTime1(0).text, "yyyymmdd") & "', '" & Format(fpDateTime1(1).text, "yyyymmdd") & "', " & codreg & ", " & codser & ", 0, '', '', 0)"
    End If
    '------- Detalle
    total = 0
    numlin = 1
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1: codmer = Trim(LimpiaDato(vaSpread1.text))
        vaSpread1.Col = 2: descri = Trim(LimpiaDato(vaSpread1.text))
        vaSpread1.Col = 4: canmin = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DCa)
        vaSpread1.Col = 5: canmer = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DCa)
        vaSpread1.Col = 6: predoc = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DPr)
        vaSpread1.Col = 7: ptotal = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DPr)
        vaSpread1.Col = 8: color = Right(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), 1)
        vaSpread1.Col = 10: codsec = vaSpread1.Value
        If color = "I" Then 'Rescata el Ingrediente
            vaSpread1.Col = 1: coding = Trim(LimpiaDato(vaSpread1.text))
        End If
        If color <> "I" Then '------- No entra si es ingrediente
            If canmer > 0 Then
                total = total + ptotal
                vg_db.Execute "INSERT INTO b_detventas (dev_rutcli, dev_tipdoc, dev_numdoc, dev_numlin, dev_codmer, dev_canmin, dev_canmer, dev_predoc, dev_ptotal, dev_descri, dev_mueinv, dev_porcen, dev_precos, dev_coding, dev_codsec) " & _
                              "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & numlin & ", '" & codmer & "', " & canmin & ", " & canmer & ", " & predoc & ", " & ptotal & ", '" & descri & "', 'S', 0, " & predoc & ", '" & coding & "', " & IIf(codsec = "0" Or Option1(0).Value = True, "Null", IIf(codsec = "estfij", -1, codsec)) & ")"
                '------- Control de Stock
                ValidaBod codbod, Trim(LimpiaDato(codmer))
                canact = 0
                RS1.Open "SELECT bod_canmer FROM b_bodegas WHERE bod_codpro = '" & Trim(LimpiaDato(codmer)) & "' AND bod_codbod = " & vg_codbod, vg_db, adOpenStatic
                If Not RS1.EOF Then
                   Do While Not RS1.EOF
                      canact = RS1!bod_canmer + canmer
                      RS1.MoveNext
                   Loop
                   vg_db.Execute "UPDATE b_bodegas SET bod_canmer = " & canact & " " & _
                                 "WHERE bod_codpro = '" & Trim(LimpiaDato(codmer)) & "' AND bod_codbod = " & vg_codbod
                End If
                RS1.Close: Set RS1 = Nothing
                numlin = numlin + 1
            End If
        End If
    Next i
    '-------> Total
    vg_db.Execute "UPDATE b_totventas SET tov_totdoc = " & total & " WHERE tov_rutcli = '" & Trim(LimpiaDato(fpText1(1).text)) & "' " & _
                  "AND tov_tipdoc = 'DP' AND tov_numdoc = " & fpLongInteger1(0).Value & " AND tov_codbod = " & vg_codbod & ""
    vg_db.CommitTrans
    Gl_Ac_Botones Me, 4, 3, ""
    Frame1.Enabled = False
    vaSpread1.Col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    '-------> Revisa Stock
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "AND   bod.bod_codpro = pro.pro_codigo AND pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "'", vg_db, adOpenStatic
        vaSpread1.Col = 9
        If Not RS1.EOF Then vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
        RS1.Close: Set RS1 = Nothing
    Next i
    I_SalDevBod Me, "DP"
Case 3 '5 '-------> Anular
    If CierrePeriodo(Format(fpDateTime1(1).text, "yyyymmdd"), codbod, 0) Then MsgBox "Periodo esta cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(1).text, "yyyymmdd"), codbod, 6) Then MsgBox "No puede anular documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CDate(fpDateTime1(1).text) < CDate(vg_ciedia) Then MsgBox "No puede anular documento, día esta cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Anula documento...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    '-------> Encabezado
    vg_db.Execute "UPDATE b_totventas SET tov_estdoc = 'A' WHERE tov_rutcli = '" & Trim(LimpiaDato(fpText1(1).text)) & "' " & _
                  "AND    tov_tipdoc = 'DP' AND tov_numdoc = " & fpLongInteger1(0).Value & " AND tov_codbod = " & vg_codbod & ""
    Label1.Caption = "ANULADA"
    '-------> Detalle
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: numlin = i
        vaSpread1.Col = 1: codmer = Trim(LimpiaDato(vaSpread1.text))
        vaSpread1.Col = 5: canmer = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DCa)
        vaSpread1.Col = 8: color = Right(vaSpread1.text, 1)
        If color <> "I" Then '-------> No entra si es ingrediente
            '-------> Control de Stock
            canact = 0
            RS1.Open "SELECT bod_canmer FROM b_bodegas WHERE bod_codpro = '" & codmer & "' AND bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")), vg_db, adOpenStatic
            If Not RS1.EOF Then
                Do While Not RS1.EOF
                    canact = RS1!bod_canmer - canmer
                    RS1.MoveNext
                Loop
                vg_db.Execute "UPDATE b_bodegas SET bod_canmer = " & canact & " " & _
                              "WHERE bod_codpro = '" & Trim(LimpiaDato(codmer)) & "' AND bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, ""))
            End If
            RS1.Close: Set RS1 = Nothing
        End If
    Next i
    '-------> Revisa Stock
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "AND   bod.bod_codpro = pro.pro_codigo AND pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "'", vg_db, adOpenStatic
        vaSpread1.Col = 9
        If Not RS1.EOF Then vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
        RS1.Close: Set RS1 = Nothing
    Next i
    vg_db.CommitTrans
    Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""
Case 11 '8 '-------> Busqueda
    If Trim(fpText1(1).text) = "" Then MsgBox "Debe seleccionar contrato...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    vg_codigo = Trim(fpText1(1).text)
    vg_nombre = "DP"
    B_SalBod.Show 1
    Me.MousePointer = 11
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Me.MousePointer = 0: Exit Sub
    DevExiste Val(vg_codigo)
    vg_codigo = ""
    Me.MousePointer = 0
Case 12 '9 '-------> Imprimir
    I_SalDevBod Me, "DP"
Case 15 '12 '-------> Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
vg_swpegreceta = 0
If Err = -2147467259 Then vg_db.RollbackTrans: GoTo paso
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
'Resume Next
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Sub DevExiste(codigo As Long)
Dim aAp As String, codsec As String, coding As String, sql1 As String, sql2 As String
Frame1.Enabled = False
est = True
'------- Consultar si salida es resumido ó sector
RS2.Open "SELECT DISTINCT dev_codsec FROM b_detventas WHERE  dev_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' AND dev_tipdoc = 'DP' AND dev_numdoc = " & Val(vg_codigo) & "", vg_db, adOpenStatic
If RS2.EOF Then RS2.Close: Set RS2 = Nothing: MsgBox "No existe devolución producción...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
If IsNull(RS2!dev_codsec) Then Option1(0).Value = True: Option1(1).Value = False Else Option1(1).Value = False: Option1(1).Value = True
RS2.Close: Set RS2 = Nothing
est = False
vaSpread1.Col = -1: vaSpread1.Row = -1
vaSpread1.Lock = True
vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0
vaSpread1.Visible = False
vaSpread2.Visible = False
RS2.Open "SELECT tov.tov_numdoc, tov.tov_fecemi, tov.tov_codbod, tov.tov_fecpro, tov.tov_codser, " & _
         "       tov.tov_codreg, tov.tov_estdoc, ser.ser_nombre, reg.reg_nombre " & _
         "FROM   b_totventas tov, b_clientes cli, a_servicio ser, a_regimen reg " & _
         "WHERE  tov.tov_rutcli = cli.cli_codigo " & _
         "AND    ser.ser_codigo = tov.tov_codser  " & _
         "AND    reg.reg_codigo = tov.tov_codreg " & _
         "AND    tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
         "AND    tov.tov_tipdoc = 'DP' AND tov.tov_codbod = " & vg_codbod & " " & _
         "AND    tov.tov_numdoc = " & codigo, vg_db, adOpenStatic
If Not RS2.EOF Then
    Do While Not RS2.EOF
        est = True
        fpLongInteger1(0).text = RS2!tov_numdoc
        fpDateTime1(0).text = RS2!tov_fecemi
        Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 10, fg_pone_cero(Str(RS2!tov_codbod), 10))
        fpDateTime1(1).text = RS2!tov_fecpro
        Combo1(0).Clear
        Combo1(0).AddItem RS2!reg_nombre & " - " & RS2!ser_nombre & Space(150) & "(" & fg_pone_cero(Str(RS2!tov_codreg), 10) & ")(" & fg_pone_cero(Str(RS2!tov_codser), 10) & ")"
        Combo1(0).ListIndex = 0
        Label1.Caption = IIf(RS2!tov_estdoc = "A", "ANULADA", "")
        RS2.MoveNext
        est = False
    Loop
End If
RS2.Close: Set RS2 = Nothing
If Option1(0).Value = True Then
    sql1 = IIf(vg_tipbase = "1", " ORDER BY dev.dev_numlin ", "")
    RS4.Open "SELECT ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 as sec_codigo, '' as sec_nombre, 0 as sec_orden, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
             "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
             "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_unidad uni " & _
             "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
             "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_coding = ing.ing_codigo " & _
             "AND   ing.ing_unimed = unm.unm_codigo AND dev.dev_codmer = pro.pro_codigo " & _
             "AND   pro.pro_coduni = uni.uni_codigo " & _
             "AND   tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
             "AND   dev.dev_numdoc = " & Val(fpLongInteger1(0).text) & _
             "AND   tov.tov_codbod = " & vg_codbod & " AND tov.tov_tipdoc = 'DP' AND dev.dev_canmer <> 0 " & _
             "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin " & sql1 & " " & _
             "UNION ALL " & _
             "SELECT 'estfij' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' as unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, -1 AS sec_codigo, '' AS sec_nombre, 0 AS sec_orden, 0 AS canmin,  dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
             "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
             "WHERE tov.tov_rutcli = dev.dev_rutcli AND  tov.tov_tipdoc = dev.dev_tipdoc " & _
             "AND   tov.tov_numdoc = dev.dev_numdoc AND  dev.dev_codmer = pro.pro_codigo " & _
             "AND   pro.pro_coduni = uni.uni_codigo AND  tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
             "AND   dev.dev_numdoc = " & Val(fpLongInteger1(0).text) & _
             "AND   tov.tov_codbod = " & vg_codbod & " AND tov.tov_tipdoc = 'DP'  " & _
             "AND  (dev.dev_coding = '' OR (dev.dev_coding) IS NULL OR dev.dev_codsec = -1) AND dev.dev_canmer <> 0 ORDER BY ing_codigo, dev.dev_numlin", vg_db, adOpenStatic
ElseIf Option1(1).Value = True Then
   sql1 = IIf(vg_tipbase = "1", " ORDER BY sec.sec_orden, dev.dev_numlin ", "")
   RS4.Open "SELECT ing.ing_codigo, ing.ing_nombre,unm.unm_nomcor, sec.sec_codigo, sec.sec_nombre, sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
            "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
            "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_sector sec, a_unidad uni " & _
            "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
            "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_coding = ing.ing_codigo " & _
            "AND   ing.ing_unimed = unm.unm_codigo AND dev.dev_codmer = pro.pro_codigo " & _
            "AND   dev.dev_codsec = sec.sec_codigo AND pro.pro_coduni = uni.uni_codigo " & _
            "AND   tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
            "AND   dev.dev_numdoc = " & Val(fpLongInteger1(0).text) & _
            "AND   tov.tov_codbod = " & vg_codbod & " AND tov.tov_tipdoc = 'DP' AND dev.dev_canmer <> 0 " & _
            "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor,  sec.sec_codigo, sec.sec_nombre, sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin " & sql1 & " " & _
            "UNION ALL " & _
            "SELECT 'estfij' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' as unm_nomcor, -1 AS sec_codigo, 'Estructura Fija' AS sec_nombre, 999999999 AS sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 AS canmin,  dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
            "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
            "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
            "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_codmer = pro.pro_codigo " & _
            "AND   pro.pro_coduni = uni.uni_codigo AND tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
            "AND   dev.dev_numdoc = " & Val(fpLongInteger1(0).text) & _
            "AND   tov.tov_codbod = " & vg_codbod & " AND tov.tov_tipdoc = 'DP'  " & _
             "AND  (dev.dev_coding = '' OR (dev.dev_coding) IS NULL OR dev.dev_codsec = -1) AND dev.dev_canmer <> 0 ORDER BY sec_orden, dev.dev_numlin", vg_db, adOpenStatic
End If
coding = "": codsec = "0": i = 0
Do While Not RS4.EOF
   '------ Sector
   If codsec <> RS4!sec_codigo Then
      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      vaSpread2.Col = 1: vaSpread2.Value = IIf(RS4!sec_codigo = -1, "estfij", RS4!sec_codigo)
      vaSpread2.Col = 2: vaSpread2.Value = Trim(RS4!sec_nombre)
      codsec = RS4!sec_codigo
   End If
   '------- Ingredientes
   If coding <> RS4!ing_codigo Then
      i = i + 1: vaSpread1.MaxRows = i
      vaSpread1.Row = i
      vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
      vaSpread1.RowHidden = IIf((vaSpread2.Row = 1 Or Option1(0).Value = True) And Check1(0).Value = 0, False, True)
      vaSpread1.Col = 1: vaSpread1.text = RS4!ing_codigo
      vaSpread1.Col = 2: vaSpread1.text = RS4!ing_nombre
      vaSpread1.Col = 3: vaSpread1.text = RS4!unm_nomcor
      vaSpread1.Col = 4: vaSpread1.text = IIf(RS4!ing_codigo = "", "", Format(RS4!canmin, fg_Pict(9, vg_DCa)))
      vaSpread1.Col = 8: vaSpread1.text = "NI" 'No bloquedo - Ingrediente
      vaSpread1.Col = 10: vaSpread1.text = IIf(RS4!sec_codigo = -1, "estfij", RS4!sec_codigo)
      vaSpread1.Col = -1
      vaSpread1.FontBold = True: vaSpread1.Lock = True
      vaSpread1.BackColor = Shape1(0).FillColor
      coding = RS4!ing_codigo
   End If
   '------- Productos
   i = i + 1: vaSpread1.MaxRows = i
   vaSpread1.Row = i
   vaSpread1.RowHidden = IIf(vaSpread2.Row = 1 Or Option1(0).Value = True, False, True)
   vaSpread1.Col = 1: vaSpread1.text = RS4!pro_codigo
   vaSpread1.Col = 2: vaSpread1.text = RS4!pro_nombre
   vaSpread1.Col = 3: vaSpread1.text = RS4!uni_nomcor
   vaSpread1.Col = 4: vaSpread1.text = Format(RS4!dev_canmin, fg_Pict(9, vg_DCa))
   vaSpread1.Col = 5: vaSpread1.text = Format(RS4!dev_canmer, fg_Pict(9, vg_DCa))
   vaSpread1.Col = 6: vaSpread1.text = Format(RS4!dev_predoc, fg_Pict(9, vg_DPr))
   vaSpread1.Col = 7: vaSpread1.text = Format(RS4!dev_ptotal, fg_Pict(9, vg_DPr))
   vaSpread1.Col = 8: vaSpread1.text = "NP" 'No bloquedo - Producto
   vaSpread1.Col = 10: vaSpread1.text = IIf(RS4!sec_codigo = -1, "estfij", RS4!sec_codigo)
   vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
   '------- Mover sectores totales
   If vaSpread2.MaxRows > 0 Then vaSpread2.Col = 3: vaSpread2.TypeHAlign = TypeHAlignRight: vaSpread2.text = Format(IIf(Trim(vaSpread2.text) = "", 0, vaSpread2.Value) + (RS4!dev_ptotal), fg_Pict(9, vg_DPr))
   '------- Trae Stock
   RS2.Open "SELECT bod.bod_canmer FROM b_productos pro, b_bodegas bod " & _
            "WHERE bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
            "AND   bod.bod_codpro = pro.pro_codigo AND pro.pro_codigo = '" & Trim(RS4!pro_codigo) & "'", vg_db, adOpenStatic
   vaSpread1.Col = 9
   If Not RS2.EOF Then vaSpread1.text = Format(RS2!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
   RS2.Close: Set RS2 = Nothing
   RS4.MoveNext
Loop
RS4.Close: Set RS4 = Nothing
vaSpread2.Visible = True
vaSpread1.Visible = True
Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""
End Sub

Function MuestraFolio(Casino As String) As String
MuestraFolio = ""
If Trim(Casino) = "" Then Exit Function
RS1.Open "SELECT tov_numdoc FROM b_totventas WHERE tov_tipdoc = 'DP' AND tov_codbod = " & vg_codbod & " ORDER BY tov_numdoc DESC", vg_db, adOpenStatic
If Not RS1.EOF Then RS1.MoveFirst: MuestraFolio = RS1!tov_numdoc + 1 Else MuestraFolio = 1
RS1.Close: Set RS1 = Nothing
End Function

Sub Limpia(op As Integer)
Label1.Caption = ""
Frame1.Enabled = True
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = ""
Combo1(0).ListIndex = -1
Combo1(1).ListIndex = IIf(Combo1(1).listcount = 1, 0, -1)
vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0
vaSpread2.Row = -1: vaSpread2.Col = -1:
vaSpread2.BackColor = Shape1(1).FillColor
vaSpread1.Col = -1: vaSpread1.Row = -1
vaSpread1.Lock = True
vaSpread1.Col = 5: vaSpread1.Row = -1
vaSpread1.Lock = False
fpText1(1).Enabled = ModCasino
Image1(1).Enabled = ModCasino
fpText1(1).text = MuestraCasino(1)
fpayuda(1).Caption = MuestraCasino(2)
'fpLongInteger1(0).text = MuestraFolio(Trim(fpText1(1).text))
fpLongInteger1(0).text = TraerCorrelativo(vg_codbod, "DP")
'Gl_Ac_Botones Me, 4, 2, ""
Gl_Ac_Botones Me, 4, op, ""
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
Dim canrea As Double, propon As Double, codmer As String, cansto As Double, cansal As Double
vaSpread1.Row = Row
vaSpread1.Col = 4
cansal = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
vaSpread1.Col = 5: canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
If ChangeMade = True And canrea > cansal Then vaSpread1.text = Format(0, fg_Pict(9, vg_DCa)): Exit Sub
vaSpread1.Col = 1: codmer = vaSpread1.text
vaSpread1.Col = 5: canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
vaSpread1.Col = 6: propon = Format(vaSpread1.text, fg_Pict(9, vg_DPr))
vaSpread1.Col = 7: vaSpread1.text = Format(canrea * propon, fg_Pict(9, vg_DPr))
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
If vaSpread1.MaxRows < 1 Then Exit Sub
Dim color As String, canrea As Double, propon As Double, codtot As Double

Dim codmer As String, cansto As Double, cansal As Double
vaSpread1.Row = Row
vaSpread1.Col = 8: color = Right(vaSpread1.text, 1)
If color = "I" Then Exit Sub
Select Case Col
Case 4
    vaSpread1.Row = Row
    vaSpread1.Col = 4
    cansal = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
    vaSpread1.Col = 5: canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
    If canrea > cansal Then vaSpread1.text = Format(0, fg_Pict(9, vg_DCa)): Exit Sub
    vaSpread1.Col = 1: codmer = vaSpread1.text
    vaSpread1.Col = 5: canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
    vaSpread1.Col = 6: propon = Format(vaSpread1.text, fg_Pict(9, vg_DPr))
    vaSpread1.Col = 7: vaSpread1.text = Format(canrea * propon, fg_Pict(9, vg_DPr))
End Select
'------- Calcular costo sectores
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
vaSpread2.Col = 3: vaSpread2.TypeHAlign = TypeHAlignRight: vaSpread2.text = IIf(totcos > 0, Format(totcos, fg_Pict(9, vg_DPr)), "")
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
