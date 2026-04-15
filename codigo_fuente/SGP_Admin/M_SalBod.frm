VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_SalBod 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salida de Bodega para Producción"
   ClientHeight    =   7230
   ClientLeft      =   2055
   ClientTop       =   1560
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   825
      TabIndex        =   9
      Top             =   315
      Width           =   8580
      Begin VB.Frame Frame4 
         Height          =   45
         Left            =   30
         TabIndex        =   30
         Top             =   2265
         Visible         =   0   'False
         Width           =   8520
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1140
         Width           =   3885
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   2070
         TabIndex        =   0
         Top             =   480
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
         ThreeDInsideHighlightColor=   -2147483633
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
         ThreeDTextHighlightColor=   -2147483633
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1875
         Width           =   3885
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   2070
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   150
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   2070
         TabIndex        =   1
         Top             =   810
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
         Left            =   2070
         TabIndex        =   3
         Top             =   1530
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
         TabIndex        =   5
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDoubleSingle fpDouble1 
         Height          =   315
         Index           =   1
         Left            =   5385
         TabIndex        =   6
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
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
         Left            =   3930
         TabIndex        =   22
         Top             =   210
         Width           =   2055
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
         Left            =   315
         TabIndex        =   19
         Top             =   1170
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Casino"
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
         Left            =   315
         TabIndex        =   18
         Top             =   510
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3435
         Picture         =   "M_SalBod.frx":0000
         Top             =   390
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Folio"
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
         Left            =   315
         TabIndex        =   17
         Top             =   195
         Width           =   420
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
         Left            =   315
         TabIndex        =   16
         Top             =   825
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Producción"
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
         Left            =   315
         TabIndex        =   15
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Régimen - Servicio"
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
         Left            =   315
         TabIndex        =   14
         Top             =   1920
         Width           =   1620
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3885
         TabIndex        =   12
         Top             =   495
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   2115
         TabIndex        =   11
         Top             =   1935
         Width           =   3885
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3930
         TabIndex        =   13
         Top             =   540
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   2130
         TabIndex        =   20
         Top             =   1185
         Width           =   3870
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3990
      Left            =   30
      TabIndex        =   27
      Top             =   2565
      Width           =   10080
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3645
         Left            =   120
         TabIndex        =   7
         Top             =   195
         Width           =   9825
         _Version        =   393216
         _ExtentX        =   17330
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
         MaxCols         =   9
         MaxRows         =   20
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_SalBod.frx":030A
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   15
      TabIndex        =   23
      Top             =   6510
      Width           =   10080
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   750
         TabIndex        =   8
         Top             =   180
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   688
         ButtonWidth     =   2963
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar Producto"
               Description     =   "Agregar Productos"
               Object.ToolTipText     =   "Agregar Productos"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar Producto "
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
               Picture         =   "M_SalBod.frx":0A1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_SalBod.frx":0D36
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
         BackColor       =   &H8000000A&
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
Dim Est As Boolean
'Dim btnX As Button

Private Sub Combo1_Click(Index As Integer)
Dim feprod As Long, codser As Long, codreg As Long, CodBod As Long, NumDia As Long, canreal As Double, cantidad As Double, aAp As String
Dim canrea As Double, canbod As Double, color As Variant, estado As String, CodPro As String
If Est Then Exit Sub
fg_carga ""
CodBod = Val(fg_codigocbo(Combo1, 1, 10, ""))
Select Case Index
Case 0
    If Combo1(0).ListIndex = -1 Or Combo1(0).Text = "" Then fg_descarga: Exit Sub
    If Combo1(1).ListIndex = -1 Or Combo1(1).Text = "" Then Est = True: Combo1(1).ListIndex = 0: Est = False
    feprod = Val(fpDateTime1(1).Year & Right("0" & fpDateTime1(1).Month, 2) & Right("0" & fpDateTime1(1).Day, 2))
    codreg = Val(Mid(Combo1(0), Len(Trim(Combo1(0).Text)) - 22, 10))
    codser = Val(Mid(Combo1(0), Len(Trim(Combo1(0).Text)) - 10, 10))
    'Muestra Raciones--------------
'    RS3.Open "select mir_racfac, mir_racper from b_minutaraciones where mir_cencos='" & Trim(fpText1(1).Text) & "' " & _
'             "and mir_codreg=" & codreg & " and mir_codser=" & codser & " and mir_fecmin=" & feprod, vg_db, adOpenStatic
'    If Not RS3.EOF Then
'        fpDouble1(0) = RS3!mir_racfac
'        fpDouble1(1) = RS3!mir_racper
'    End If
'    RS3.Close: Set RS3 = Nothing
    '-----------------------------
    RS1.Open "SELECT count(tov_fecpro) as suma from b_totventas where " & _
             "tov_rutcli='" & Trim(LimpiaDato(fpText1(1).Text)) & "' and tov_tipdoc='SP' " & _
             "and tov_fecpro=CDate('" & fpDateTime1(1).Text & "') " & _
             "and tov_codreg=" & codreg & " and tov_codser=" & codser & " and tov_estdoc<>'A'", vg_db, adOpenStatic
    Do While Not RS1.EOF
        If RS1!suma > 0 Then
            RS1.Close: Set RS1 = Nothing
            vaSpread1.MaxRows = 0
            Frame2.Enabled = True
            If vaSpread1.Enabled = True Then vaSpread1.SetFocus
            Toolbar2_ButtonClick Toolbar2.Buttons.Item(1)
            Exit Sub
        End If
        RS1.MoveNext
    Loop
    Me.MousePointer = 11
    RS1.Close: Set RS1 = Nothing
    NumDia = fg_NumDia(Trim(Left(fg_Fecha_Dia(Trim(Str(feprod)), 2), Len(fg_Fecha_Dia(Trim(Str(feprod)), 2)) - 2)))
    '------------------------------------ MINUTA ---------------------------------
    'Creo tabla temporal y chequeo si existe antes
    aAp = Trim(vg_NUsr) & "_tmp_SalBod"
    fg_CheckTmp aAp
    RS1.Open "select red.red_codpro, sum (mid.mid_numrac*(red.red_canpro/rec.rec_basrac)) AS cantidad into " & aAp & " " & _
             "from b_minuta mi, b_minutadet mid, b_receta rec, b_recetadet red " & _
             "where rec.rec_codigo=mid.mid_codrec " & _
             "and rec.rec_codigo=red.red_codigo and mi.min_codigo=mid.mid_codigo " & _
             "and mi.min_fecmin=" & feprod & " and mid.mid_tipmin='2' and mi.min_cencos='" & Trim(fpText1(1).Text) & "' " & _
             "and mi.min_codreg=" & codreg & " and mi.min_codser=" & codser & " " & _
             "group by red.red_codpro", vg_db, adOpenStatic
    Set RS1 = Nothing
    '---- Creo en cero los productos que no existen en bodega ---------------------
    RS1.Open "Select ing.ing_codped As codigo From " & aAp & " aux, b_ingrediente ing Where ing.ing_codigo=aux.red_codpro " & _
             "Union " & _
             "Select mif_codpro As codigo From b_minutafija As mif " & _
             "Where mif.mif_cencos='" & Trim(fpText1(1).Text) & "' And mif.mif_codreg=" & codreg & " " & _
             "And mif.mif_codser=" & codser & " And mif.mif_dianro=" & NumDia & " " & _
             "And mif.mif_fecval=(Select max(mif2.mif_fecval) From b_minutafija As mif2)", vg_db, adOpenStatic
    Do While Not RS1.EOF
        ValidaBod CodBod, RS1!Codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    '-----------------------------------------------------------------------------
    RS1.Open "Select   pro.pro_codigo, pro.pro_propon, pro.pro_nombre, uni.uni_nomcor, ing.ing_codigo, ing.ing_nombre, " & _
             "         bod.bod_canmer, unm.unm_nomcor, Sum(aux.cantidad/pro.pro_facing) As cantidad, Sum(aux.cantidad) As cantidad2 " & _
             "From     b_productos pro, b_ingrediente ing, b_productosing pri, a_unidad uni, " & _
             "         b_bodegas bod, a_unidadmed unm, " & aAp & " aux " & _
             "Where    pri.pri_coding=aux.red_codpro " & _
             "And      ing.ing_codigo=pri.pri_coding " & _
             "And      pro.pro_codigo=ing.ing_codped " & _
             "And      ing.ing_unimed=unm.unm_codigo " & _
             "And      pro.pro_coduni=uni.uni_codigo " & _
             "And      bod.bod_codpro=pro.pro_codigo " & _
             "And      bod.bod_codbod=" & CodBod & " " & _
             "Group by pro.pro_codigo, pro.pro_propon, pro.pro_nombre, uni.uni_nomcor, ing.ing_codigo, " & _
             "         ing.ing_nombre, bod.bod_canmer, unm.unm_nomcor " & _
             "Order by ing.ing_codigo, pro.pro_codigo", vg_db, adOpenStatic
    vaSpread1.MaxRows = 0
    vaSpread1.Visible = False
    i = 0
    Do While Not RS1.EOF
        i = i + 1
        'Ingrediente
        cantidad = IIf(IsNull(RS1!cantidad2), 0, RS1!cantidad2)
        vaSpread1.MaxRows = i
        vaSpread1.Row = i
        vaSpread1.col = 1: vaSpread1.Text = RS1!ing_codigo
        vaSpread1.col = 2: vaSpread1.Text = RS1!ing_nombre
        vaSpread1.col = 3: vaSpread1.Text = RS1!unm_nomcor
        vaSpread1.col = 4: vaSpread1.Text = Format(Round(cantidad, vg_DCa), fg_Pict(9, vg_DCa))
        vaSpread1.col = 8: vaSpread1.Text = "NI" 'No bloquedo - Ingrediente
        vaSpread1.col = -1
        vaSpread1.FontBold = True: vaSpread1.Lock = True
        vaSpread1.col = -1: vaSpread1.BackColor = Shape1(3).FillColor
            
        'Producto
        i = i + 1
        cantidad = IIf(IsNull(RS1!cantidad), 0, RS1!cantidad)
        If cantidad > 0 And cantidad < 0.5 Then canreal = Format(0.5, fg_Pict(9, vg_DCa)) Else canreal = Format(Round(cantidad, 0), fg_Pict(9, vg_DPr))
        vaSpread1.MaxRows = i
        vaSpread1.Row = i
        vaSpread1.col = 1: vaSpread1.Text = RS1!pro_codigo
        vaSpread1.col = 2: vaSpread1.Text = RS1!pro_nombre
        vaSpread1.col = 3: vaSpread1.Text = RS1!uni_nomcor
        vaSpread1.col = 4: vaSpread1.Text = Format(cantidad, fg_Pict(9, vg_DCa))
        vaSpread1.col = 5: vaSpread1.Text = Format(Round(canreal, vg_DCa), fg_Pict(9, vg_DCa))
        vaSpread1.col = 6: vaSpread1.Text = Format(RS1!pro_propon, fg_Pict(9, vg_DPr))
        vaSpread1.col = 7: vaSpread1.Text = Format(Round(canreal, vg_DCa) * RS1!pro_propon, fg_Pict(9, vg_DPr))
        vaSpread1.col = 8: vaSpread1.Text = "NM" 'No bloquedo - Viene de la Minuta
        vaSpread1.col = 9: vaSpread1.Text = RS1!bod_canmer
        
        'REvisa color
        vaSpread1.col = 5: canrea = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
        vaSpread1.col = 8: color = Right(vaSpread1.Text, 1)
        vaSpread1.col = 9: canbod = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
        
        canaux = 0
        For z = 1 To vaSpread1.MaxRows
            vaSpread1.Row = z
            vaSpread1.col = 8: color2 = Right(vaSpread1.Text, 1)
            vaSpread1.col = 1
            If RS1!pro_codigo = vaSpread1.Text And color2 <> "I" Then
                vaSpread1.col = 5: canaux = canaux + Format(IIf(vaSpread1.Text = "", 0, vaSpread1.Text), fg_Pict(9, vg_DCa))
            End If
        Next z
        canrea = IIf(canaux > 0, canaux, canrea)
        
        If canbod - canrea >= 0 Then
            For z = 1 To vaSpread1.MaxRows
                vaSpread1.Row = z
                vaSpread1.col = 8: color2 = Right(vaSpread1.Text, 1)
                vaSpread1.col = 1
                If RS1!pro_codigo = vaSpread1.Text And color2 <> "I" Then
                    vaSpread1.col = -1: vaSpread1.BackColor = Shape1(IIf(color = "M", 2, IIf(color = "U", 0, 3))).FillColor
                    vaSpread1.col = 8: vaSpread1.Text = "N" & color 'No Bloqueado - Depende
                End If
            Next z
            
        Else
            For z = 1 To vaSpread1.MaxRows
                vaSpread1.Row = z
                vaSpread1.col = 8: color2 = Right(vaSpread1.Text, 1)
                vaSpread1.col = 1
                If RS1!pro_codigo = vaSpread1.Text And color2 <> "I" Then
                    vaSpread1.col = -1: vaSpread1.BackColor = Shape1(1).FillColor
                    vaSpread1.col = 8: vaSpread1.Text = "S" & color 'Bloqueado - Depende
                End If
            Next z
        End If
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    '---------------------------------- FIN MINUTA ---------------------------------
    '----------------------------- ESTRUCTURA FIJA ---------------------------------
    RS1.Open "Select   pro.pro_codigo, pro.pro_propon, pro.pro_nombre, uni.uni_nomcor, bod.bod_canmer, mif.mif_canpro As cantidad " & _
             "From     b_productos pro, a_unidad uni, b_bodegas bod, b_minutafija mif " & _
             "Where    pro.pro_coduni=uni.uni_codigo " & _
             "And      bod.bod_codpro=pro.pro_codigo " & _
             "And      mif.mif_codpro=pro.pro_codigo " & _
             "And      bod.bod_codbod=" & CodBod & " " & _
             "And      mif.mif_cencos='" & Trim(fpText1(1).Text) & "' " & _
             "And      mif.mif_codreg=" & codreg & " " & _
             "And      mif.mif_codser=" & codser & " " & _
             "And      mif.mif_dianro=" & NumDia & " " & _
             "And      mif.mif_fecval=(select max(mif2.mif_fecval) from b_minutafija as mif2) " & _
             "Group by pro.pro_codigo, pro.pro_propon, pro.pro_nombre, uni.uni_nomcor, bod.bod_canmer, mif.mif_canpro " & _
             "Order by pro.pro_codigo", vg_db, adOpenStatic
    If Not RS1.EOF Then
        i = i + 1
        'Ingrediente
        vaSpread1.MaxRows = i
        vaSpread1.Row = i
        vaSpread1.col = 1: vaSpread1.Text = ""
        vaSpread1.col = 2: vaSpread1.Text = "Estructura Fija"
        vaSpread1.col = 8: vaSpread1.Text = "NI" 'No bloquedo - Ingrediente
        vaSpread1.col = -1
        vaSpread1.FontBold = True: vaSpread1.Lock = True
        vaSpread1.BackColor = Shape1(3).FillColor
        vaSpread1.col = 1
        vaSpread1.ForeColor = Shape1(3).FillColor
    End If
    Do While Not RS1.EOF
        'Producto
        i = i + 1
        cantidad = IIf(IsNull(RS1!cantidad), 0, RS1!cantidad)
        If cantidad > 0 And cantidad < 0.5 Then canreal = Format(0.5, fg_Pict(9, vg_DCa)) Else canreal = Format(Round(cantidad, 0), fg_Pict(9, vg_DPr))
        vaSpread1.MaxRows = i
        vaSpread1.Row = i
        vaSpread1.col = 1: vaSpread1.Text = RS1!pro_codigo
        vaSpread1.col = 2: vaSpread1.Text = RS1!pro_nombre
        vaSpread1.col = 3: vaSpread1.Text = RS1!uni_nomcor
        vaSpread1.col = 4: vaSpread1.Text = Format(cantidad, fg_Pict(9, vg_DCa))
        vaSpread1.col = 5: vaSpread1.Text = Format(Round(canreal, vg_DCa), fg_Pict(9, vg_DCa))
        vaSpread1.col = 6: vaSpread1.Text = Format(RS1!pro_propon, fg_Pict(9, vg_DPr))
        vaSpread1.col = 7: vaSpread1.Text = Format(Round(canreal, vg_DCa) * RS1!pro_propon, fg_Pict(9, vg_DPr))
        vaSpread1.col = 8: vaSpread1.Text = "NM" 'No bloquedo - Viene de la Minuta
        vaSpread1.col = 9: vaSpread1.Text = RS1!bod_canmer
        
        'REvisa color
        vaSpread1.col = 5: canrea = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
        vaSpread1.col = 8: color = Right(vaSpread1.Text, 1)
        vaSpread1.col = 9: canbod = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
        
        canaux = 0
        For z = 1 To vaSpread1.MaxRows
            vaSpread1.Row = z
            vaSpread1.col = 8: color2 = Right(vaSpread1.Text, 1)
            vaSpread1.col = 1
            If RS1!pro_codigo = vaSpread1.Text And color2 <> "I" Then
                vaSpread1.col = 5: canaux = canaux + Format(IIf(vaSpread1.Text = "", 0, vaSpread1.Text), fg_Pict(9, vg_DCa))
            End If
        Next z
        canrea = IIf(canaux > 0, canaux, canrea)
        
        If canbod - canrea >= 0 Then
            For z = 1 To vaSpread1.MaxRows
                vaSpread1.Row = z
                vaSpread1.col = 8: color2 = Right(vaSpread1.Text, 1)
                vaSpread1.col = 1
                If RS1!pro_codigo = vaSpread1.Text And color2 <> "I" Then
                    vaSpread1.col = -1: vaSpread1.BackColor = Shape1(IIf(color = "M", 2, IIf(color = "U", 0, 3))).FillColor
                    vaSpread1.col = 8: vaSpread1.Text = "N" & color 'No Bloqueado - Depende
                End If
            Next z
            
        Else
            For z = 1 To vaSpread1.MaxRows
                vaSpread1.Row = z
                vaSpread1.col = 8: color2 = Right(vaSpread1.Text, 1)
                vaSpread1.col = 1
                If RS1!pro_codigo = vaSpread1.Text And color2 <> "I" Then
                    vaSpread1.col = -1: vaSpread1.BackColor = Shape1(1).FillColor
                    vaSpread1.col = 8: vaSpread1.Text = "S" & color 'Bloqueado - Depende
                End If
            Next z
        End If
        
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    '----------------------------- FIN ESTRUCTURA FIJA ------------------------------
    
    
    Frame2.Enabled = True
    If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
    vaSpread1.Visible = True
    Me.MousePointer = 0
Case 1
    If vaSpread1.MaxRows = 0 Then fg_descarga: Exit Sub
    fg_carga ""
    vaSpread1.Visible = False
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.col = 1
        CodPro = Trim(LimpiaDato(vaSpread1.Text))
        RS2.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod=" & CodBod & " " & _
                 "and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & Trim(LimpiaDato(vaSpread1.Text)) & "'", vg_db, adOpenStatic
        vaSpread1.col = 9
        If Not RS2.EOF Then vaSpread1.Text = RS2!bod_canmer Else vaSpread1.Text = 0
        RS2.Close: Set RS2 = Nothing
        'REvisa color
        vaSpread1.col = 5: canrea = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
        vaSpread1.col = 8: color = Right(vaSpread1.Text, 1)
        vaSpread1.col = 9: canbod = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
        If color <> "I" Then
        
            canaux = 0
            For z = 1 To vaSpread1.MaxRows
                vaSpread1.Row = z
                vaSpread1.col = 8: color2 = Right(vaSpread1.Text, 1)
                vaSpread1.col = 1
                If CodPro = vaSpread1.Text And color2 <> "I" Then
                    vaSpread1.col = 5: canaux = canaux + Format(IIf(vaSpread1.Text = "", 0, vaSpread1.Text), fg_Pict(9, vg_DCa))
                End If
            Next z
            canrea = IIf(canaux > 0, canaux, canrea)
            
            If canbod - canrea >= 0 Then
                For z = 1 To vaSpread1.MaxRows
                    vaSpread1.Row = z
                    vaSpread1.col = 8: color2 = Right(vaSpread1.Text, 1)
                    vaSpread1.col = 1
                    If CodPro = vaSpread1.Text And color2 <> "I" Then
                        vaSpread1.col = -1: vaSpread1.BackColor = Shape1(IIf(color = "M", 2, IIf(color = "U", 0, 3))).FillColor
                        vaSpread1.col = 8: vaSpread1.Text = "N" & color 'No Bloqueado - Depende
                    End If
                Next z
                
            Else
                For z = 1 To vaSpread1.MaxRows
                    vaSpread1.Row = z
                    vaSpread1.col = 8: color2 = Right(vaSpread1.Text, 1)
                    vaSpread1.col = 1
                    If CodPro = vaSpread1.Text And color2 <> "I" Then
                        vaSpread1.col = -1: vaSpread1.BackColor = Shape1(1).FillColor
                        vaSpread1.col = 8: vaSpread1.Text = "S" & color 'Bloqueado - Depende
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
Me.Height = 7635
Me.Width = 10245
fg_centra Me
Est = False
Me.HelpContextID = vg_OpcM
Msgtitulo = "Salida Producción"
Dim X As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
Gl_Mo_Botones Me, 4
vaSpread1.Row = -1
vaSpread1.col = 4: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.col = 6: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
vaSpread1.col = 7: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
RS1.Open "select bod_nombre, bod_codigo from a_bodega", vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
        Combo1(1).AddItem RS1!bod_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!bod_codigo), 10) & ")"
        RS1.MoveNext
    Loop
End If
RS1.Close: Set RS1 = Nothing
Limpia
End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then
    Frame3.Left = (Me.Width \ 2) - (Frame3.Width \ 2)
    Frame1.Left = (Me.Width \ 2) - (Frame1.Width \ 2)
    Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)
ElseIf Me.WindowState = 0 Then
    Frame3.Left = 30
    Frame1.Left = 825
    Frame2.Left = 15
End If
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
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
Dim feprod As Long, codreg As Long, codser As Long, aAp As String
Select Case Index
Case 1
    Toolbar1.Buttons(8).Enabled = True
    Combo1(0).Clear
    vaSpread1.MaxRows = 0
    If Trim(fpDateTime1(1).Text) = "" Or Trim(fpText1(1).Text) = "" Then Exit Sub
    feprod = Val(fpDateTime1(1).Year & Right("0" & fpDateTime1(1).Month, 2) & Right("0" & fpDateTime1(1).Day, 2))
    
    
    NumDia = fg_NumDia(Trim(Left(fg_Fecha_Dia(Trim(Str(feprod)), 2), Len(fg_Fecha_Dia(Trim(Str(feprod)), 2)) - 2)))
    aAp = Trim(vg_NUsr) & "_tmp_SalBod"
    'Creo tabla temporal y chequeo si existe antes
    fg_CheckTmp aAp
    RS1.Open "select distinct ser.ser_codigo, reg.reg_codigo, ser.ser_nombre, reg.reg_nombre into " & aAp & " " & _
             "from b_minuta [min], b_minutadet mid, a_servicio ser, a_regimen reg " & _
             "where min.min_codigo=mid.mid_codigo and min.min_codser=ser.ser_codigo and min.min_codreg=reg.reg_codigo " & _
             "and mid.mid_tipmin='2' and min.min_fecmin=" & feprod & " and min.min_cencos='" & Trim(fpText1(1).Text) & "'", vg_db, adOpenStatic
    Set RS1 = Nothing
    vg_db.Execute "insert into " & aAp & " select distinct ser.ser_codigo, reg.reg_codigo, ser.ser_nombre, reg.reg_nombre " & _
                  "from b_minutafija as mif, a_servicio ser, a_regimen reg " & _
                  "where mif.mif_codser=ser.ser_codigo and mif.mif_codreg=reg.reg_codigo " & _
                  "and mif.mif_cencos='" & Trim(fpText1(1).Text) & "' and mif.mif_dianro=" & NumDia & " " & _
                  "and mif.mif_fecval=(select max(mif2.mif_fecval) from b_minutafija as mif2)"
                  
    RS1.Open "select ser_codigo, reg_codigo, ser_nombre, reg_nombre " & _
             "from " & aAp & " " & _
             "group by ser_codigo, reg_codigo, ser_nombre, reg_nombre ", vg_db, adOpenStatic
    
    Combo1(0).Clear
    Do While Not RS1.EOF
        Combo1(0).AddItem RS1!reg_nombre & " - " & RS1!ser_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!reg_codigo), 10) & ")(" & fg_pone_cero(Str(RS1!ser_codigo), 10) & ")"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    If Combo1(0).ListCount = 0 Then MsgBox "No existe planificación...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    'If Combo1(0).ListCount = 1 Then Combo1(0).ListIndex = 0
End Select
End Sub

Private Sub fpDouble1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
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
If fpText1(1).Text = "" Then Exit Sub
RS1.Open "select cli_nombre from b_clientes where cli_codigo='" & fpText1(1).Text & "' and cli_tipo=0", vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
        fpayuda(Index).Caption = RS1!cli_nombre
        Gl_Ac_Botones Me, 4, 2, ""
        fpText1(1).Enabled = False
        RS1.MoveNext
    Loop
Else
    RS1.Close: Set RS1 = Nothing
    MsgBox "Casino no existe...", vbExclamation + vbOKOnly, Msgtitulo
    Limpia
    If fpText1(1).Enabled = True Then fpText1(1).SetFocus
    Exit Sub
End If
RS1.Close: Set RS1 = Nothing
fpLongInteger1(0).Text = MuestraFolio(Trim(fpText1(1).Text))
End Sub

Private Sub Image1_Click(Index As Integer)
vg_codigo = 0
Select Case Index
Case 1
    vg_left = fpayuda(Index).Left + 1920
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "Casino"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    If Trim(vg_codigo) <> fpText1(Index) Then Limpia
    fpText1(Index) = Trim(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    fpText1_LostFocus 1
    If fpDateTime1(Index - 1).Enabled = True Then fpDateTime1(Index - 1).SetFocus
    Gl_Ac_Botones Me, 4, 2, ""
End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim rutcli As String, TipDoc As String, numdoc As Long, fecha As Long, CodBod  As Long, fecemi As Date, fecpro As Date, codreg As Long, codser As Long, i As Long, canact As Long, aAp  As String
Dim numlin As Long, CodMer As String, coding As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String, total As Double, diablq As Date, color As String
On Error GoTo Man_Error
If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
Select Case Button.Index
Case 1 'Nuevo
    Limpia
    If fpText1(1).Enabled = True Then fpText1(1).SetFocus
Case 3 'Graba
    If Trim(fpText1(1).Text) = "" Or Trim(fpLongInteger1(0).Text) = "" Or Trim(Combo1(0).Text) = "" Or Trim(fpDateTime1(0).Text) = "" _
    Or Trim(Combo1(1).Text) = "" Or Trim(fpDateTime1(1).Text) = "" Then MsgBox "Debe ingresar dato importante...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    'If Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(fpDateTime1(1).Text), "mm/yyyy") < Format(Now, "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If (Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(fpDateTime1(1).Text), "mm/yyyy") < Format(Now, "mm/yyyy")) Or Format(CDate(fpDateTime1(1).Text), "mm/yyyy") < Format(Month(Now) - 1 & "/" & Year(Now), "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.col = 8
        If Left(vaSpread1.Text, 1) = "S" Then MsgBox "Existe una cantidad que exede el Stock...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Next i
    total = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.col = 7: ptotal = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DPr))
        total = total + ptotal
    Next i
    If total = 0 Then MsgBox "El total del documento debe ser mayor a 0...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    vg_db.BeginTrans
    rutcli = Trim(LimpiaDato(fpText1(1).Text))
    TipDoc = "SP"
    fpLongInteger1(0).Text = MuestraFolio(Trim(fpText1(1).Text))
    numdoc = fpLongInteger1(0).Text
    CodBod = Val(fg_codigocbo(Combo1, 1, 10, ""))
    fecemi = Format(fpDateTime1(0).Text, "dd/mm/yyyy")
    fecpro = Format(fpDateTime1(1).Text, "dd/mm/yyyy")
    fecha = Val(fpDateTime1(1).Year & Right("0" & fpDateTime1(1).Month, 2) & Right("0" & fpDateTime1(1).Day, 2))
    codreg = Val(Mid(Combo1(0), Len(Trim(Combo1(0).Text)) - 22, 10))
    codser = Val(Mid(Combo1(0), Len(Trim(Combo1(0).Text)) - 10, 10))
    'Encabezado
    vg_db.Execute "insert into b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_numinf) " & _
                  "values ('" & rutcli & "', '" & TipDoc & "', " & numdoc & ", " & CodBod & ", CDate('" & _
                  Format(fpDateTime1(0).Text, "dd/mm/yyyy") & "'), CDate('" & Format(fpDateTime1(1).Text, "dd/mm/yyyy") & "'), " & codreg & ", " & codser & ", 0, '', '', 0)"
    'Detalle
    total = 0
    numlin = 1
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.col = 1: CodMer = Trim(LimpiaDato(vaSpread1.Text))
        vaSpread1.col = 2: descri = Trim(LimpiaDato(vaSpread1.Text))
        vaSpread1.col = 4: canmin = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DCa)
        vaSpread1.col = 5: canmer = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DCa)
        vaSpread1.col = 6: predoc = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DPr)
        vaSpread1.col = 7: ptotal = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DPr)
        vaSpread1.col = 8: color = Right(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), 1)
        total = total + ptotal
        If color = "I" Then 'Rescata el Ingrediente
            vaSpread1.col = 1: coding = Trim(LimpiaDato(vaSpread1.Text))
        End If
        If color <> "I" Then 'No entra si es ingrediente
            If color <> "U" Or canmer <> 0 Then 'No entra si es ingresado x usuario y valor es cero
                vg_db.Execute "insert into b_detventas (dev_rutcli, dev_tipdoc, dev_numdoc, dev_numlin, dev_codmer, dev_canmin, dev_canmer, dev_predoc, dev_ptotal, dev_descri, dev_mueinv, dev_porcen, dev_precos, dev_coding) " & _
                              "values ('" & rutcli & "', '" & TipDoc & "', " & numdoc & ", " & numlin & ", '" & CodMer & "', " & canmin & ", " & canmer & ", " & predoc & ", " & ptotal & ", '" & descri & "', 'S', 0, " & predoc & ", '" & coding & "')"
                'Control de Stock
                canact = 0
                RS1.Open "select bod_canmer from b_bodegas where bod_codpro='" & Trim(LimpiaDato(CodMer)) & "' " & _
                         "and bod_codbod=" & CodBod, vg_db, adOpenStatic
                If Not RS1.EOF Then
                    Do While Not RS1.EOF
                        canact = RS1!bod_canmer - canmer
                        RS1.MoveNext
                    Loop
                    vg_db.Execute "update b_bodegas set bod_canmer=" & canact & " " & _
                                  "where bod_codpro='" & Trim(LimpiaDato(CodMer)) & "' and bod_codbod=" & CodBod
                End If
                RS1.Close: Set RS1 = Nothing: numlin = numlin + 1
            End If
        End If
    Next i
    'Total
    vg_db.Execute "update b_totventas set tov_totdoc=" & total & " where tov_rutcli='" & Trim(LimpiaDato(fpText1(1).Text)) & "' " & _
                  "and tov_tipdoc='SP' and tov_numdoc=" & fpLongInteger1(0).Value
    vg_db.CommitTrans
    Gl_Ac_Botones Me, 4, 3, ""
    Frame1.Enabled = False
    vaSpread1.col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    Frame2.Enabled = False
    'Revisa Stock
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.col = 1
        RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & Trim(LimpiaDato(vaSpread1.Text)) & "'", vg_db, adOpenStatic
        vaSpread1.col = 9
        If Not RS1.EOF Then vaSpread1.Text = RS1!bod_canmer Else vaSpread1.Text = 0
        RS1.Close: Set RS1 = Nothing
    Next i
    'Graba Raciones--------------
'    RS3.Open "select mir_racfac, mir_racper from b_minutaraciones where mir_cencos='" & Trim(fpText1(1).Text) & "' " & _
'             "and mir_codreg=" & codreg & " and mir_codser=" & codser & " and mir_fecmin=" & Fecha, vg_db, adOpenStatic
'    If RS3.EOF Then
'        vg_db.Execute "insert into b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_racfac, mir_racper) " & _
'                      "values ('" & rutcli & "', " & codreg & ", " & codser & ", " & Fecha & ", " & fpDouble1(0).Value & ", " & fpDouble1(1).Value & ")"
'    Else
'        vg_db.Execute "update b_minutaraciones set mir_racfac=" & fpDouble1(0).Value & ", mir_racper=" & fpDouble1(1).Value & " " & _
'                      "where mir_cencos='" & Trim(fpText1(1).Text) & "' and mir_codreg=" & codreg & " and mir_codser=" & codser & " and mir_fecmin=" & Fecha
'    End If
'    RS3.Close: Set RS3 = Nothing
    '-----------------------------
    I_SalDevBod Me, "SP"
Case 5 'Anular
    'If Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(fpDateTime1(1).Text), "mm/yyyy") < Format(Now, "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If (Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(fpDateTime1(1).Text), "mm/yyyy") < Format(Now, "mm/yyyy")) Or Format(CDate(fpDateTime1(1).Text), "mm/yyyy") < Format(Month(Now) - 1 & "/" & Year(Now), "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Anula documento...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    'Encabezado
    vg_db.Execute "update b_totventas set tov_estdoc='A' where tov_rutcli='" & Trim(LimpiaDato(fpText1(1).Text)) & "' " & _
                  "and tov_tipdoc='SP' and tov_numdoc=" & fpLongInteger1(0).Value
    'Detalle
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: numlin = i
        vaSpread1.col = 1: CodMer = Trim(LimpiaDato(vaSpread1.Text))
        vaSpread1.col = 5: canmer = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DCa))
        vaSpread1.col = 8: color = Right(vaSpread1.Text, 1)
        If color <> "I" Then 'No entra si es ingrediente
            'Control de Stock
            canact = 0
            RS1.Open "select bod_canmer from b_bodegas where bod_codpro='" & CodMer & "' and bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")), vg_db, adOpenStatic
            If Not RS1.EOF Then
                Do While Not RS1.EOF
                    canact = RS1!bod_canmer + canmer
                    RS1.MoveNext
                Loop
                vg_db.Execute "update b_bodegas set bod_canmer=" & canact & " " & _
                              "where bod_codpro='" & Trim(LimpiaDato(CodMer)) & "' and bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, ""))
            End If
            RS1.Close: Set RS1 = Nothing
        End If
    Next i
    'Revisa Stock
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.col = 1
        RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & Trim(LimpiaDato(vaSpread1.Text)) & "'", vg_db, adOpenStatic
        vaSpread1.col = 9
        If Not RS1.EOF Then vaSpread1.Text = RS1!bod_canmer Else vaSpread1.Text = 0
        RS1.Close: Set RS1 = Nothing
    Next i
    Label1.Caption = "ANULADA"
    vg_db.CommitTrans
    Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""
Case 8 'Busqueda
    If Trim(fpText1(1).Text) = "" Then MsgBox "Debe seleccionar casino...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    vg_codigo = Trim(fpText1(1).Text)
    vg_nombre = "SP"
    B_SalBod.Show 1
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    Me.MousePointer = 11
    Me.Refresh
    Frame2.Enabled = False
    Frame1.Enabled = False
    vaSpread1.col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    vaSpread1.MaxRows = 0
    RS2.Open "select tov.tov_numdoc, tov.tov_fecemi, tov.tov_codbod, tov.tov_fecpro, tov.tov_codser, tov.tov_codreg, tov.tov_estdoc, ser.ser_nombre, reg.reg_nombre " & _
             "from b_totventas tov, b_clientes cli, a_servicio ser, a_regimen reg " & _
             "where tov.tov_rutcli=cli.cli_codigo and ser.ser_codigo=tov.tov_codser " & _
             "and reg.reg_codigo=tov.tov_codreg and tov.tov_rutcli='" & LimpiaDato(Trim(fpText1(1).Text)) & "' " & _
             "and tov.tov_tipdoc='SP' " & _
             "and tov.tov_numdoc=" & Val(vg_codigo), vg_db, adOpenStatic
    If Not RS2.EOF Then
        Do While Not RS2.EOF
            Est = True
            fpLongInteger1(0).Text = RS2!tov_numdoc
            fpDateTime1(0).Text = RS2!tov_fecemi
            Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 10, fg_pone_cero(Str(RS2!tov_codbod), 10))
            fpDateTime1(1).Text = RS2!tov_fecpro
            Combo1(0).Clear
            Combo1(0).AddItem RS2!reg_nombre & " - " & RS2!ser_nombre & Space(150) & "(" & fg_pone_cero(Str(RS2!tov_codreg), 10) & ")(" & fg_pone_cero(Str(RS2!tov_codser), 10) & ")"
            Combo1(0).ListIndex = 0
            Label1.Caption = IIf(RS2!tov_estdoc = "", "", "ANULADA")
            Est = False
            RS2.MoveNext
        Loop
    End If
    RS2.Close: Set RS2 = Nothing
    'Muestra Raciones--------------
    codreg = Val(Mid(Combo1(0), Len(Trim(Combo1(0).Text)) - 22, 10))
    codser = Val(Mid(Combo1(0), Len(Trim(Combo1(0).Text)) - 10, 10))
    fecha = Val(fpDateTime1(1).Year & Right("0" & fpDateTime1(1).Month, 2) & Right("0" & fpDateTime1(1).Day, 2))
'    RS3.Open "select mir_racfac, mir_racper from b_minutaraciones where mir_cencos='" & Trim(fpText1(1).Text) & "' " & _
'             "and mir_codreg=" & codreg & " and mir_codser=" & codser & " and mir_fecmin=" & Fecha, vg_db, adOpenStatic
'    If Not RS3.EOF Then
'        fpDouble1(0) = RS3!mir_racfac
'        fpDouble1(1) = RS3!mir_racper
'    End If
'    RS3.Close: Set RS3 = Nothing
'-----------------------------
    aAp = Trim(vg_NUsr) & "_tmp_SalBod"
    fg_CheckTmp aAp
    RS3.Open "Select   ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, sum(dev.dev_canmin * pro.pro_facing) as canmin, max(dev_numlin) as num " & _
             "Into " & aAp & " " & _
             "From     b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm " & _
             "Where    ing.ing_codigo=dev.dev_coding " & _
             "And      ing.ing_unimed=unm.unm_codigo " & _
             "And      dev.dev_codmer=pro.pro_codigo " & _
             "And      dev.dev_rutcli='" & LimpiaDato(Trim(fpText1(1).Text)) & "' " & _
             "And      dev.dev_tipdoc='SP' " & _
             "And      dev.dev_numdoc=" & Val(vg_codigo) & " " & _
             "Group by ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor " & _
             "Order by max(dev.dev_numlin)", vg_db, adOpenStatic
    Set RS3 = Nothing
    vg_db.Execute "Insert Into " & aAp & " " & _
                  "Select   '' as ing_codigo, 'Estructura Fija' as ing_nombre, '' as unm_nomcor, 0 as canmin, max(dev_numlin) as num " & _
                  "From     b_detventas " & _
                  "Where    dev_rutcli='" & LimpiaDato(Trim(fpText1(1).Text)) & "' " & _
                  "And      dev_tipdoc='SP' " & _
                  "And      dev_numdoc=" & Val(vg_codigo) & " " & _
                  "And      dev_coding=''"
    RS3.Open "Select ing_codigo, ing_nombre, unm_nomcor, canmin, num From " & aAp & " Order by num", vg_db, adOpenStatic
    vaSpread1.Visible = False
    i = 1
    Do While Not RS3.EOF
        vaSpread1.MaxRows = i
        vaSpread1.Row = i
        vaSpread1.col = 1: vaSpread1.Text = RS3!ing_codigo
        vaSpread1.col = 2: vaSpread1.Text = RS3!ing_nombre
        vaSpread1.col = 3: vaSpread1.Text = RS3!unm_nomcor
        vaSpread1.col = 4: vaSpread1.Text = IIf(RS3!ing_codigo = "", "", Format(RS3!canmin, fg_Pict(9, vg_DCa)))
        vaSpread1.col = 8: vaSpread1.Text = "NI" 'No bloquedo - Ingrediente
        vaSpread1.col = -1
        vaSpread1.FontBold = True
        vaSpread1.BackColor = Shape1(3).FillColor
        i = i + 1
        RS1.Open "Select   dev.dev_codmer, dev.dev_canmin, dev.dev_canmer, dev.dev_predoc, " & _
                 "         dev.dev_ptotal, dev.dev_descri, uni.uni_nomcor " & _
                 "From     b_detventas dev, b_productos pro, a_unidad uni " & _
                 "Where    dev.dev_codmer=pro.pro_codigo " & _
                 "And      pro.pro_coduni=uni.uni_codigo " & _
                 "And      dev.dev_rutcli='" & LimpiaDato(Trim(fpText1(1).Text)) & "'" & _
                 "And      dev.dev_tipdoc='SP' " & _
                 "And      dev.dev_numdoc=" & Val(vg_codigo) & " " & _
                 "And      dev.dev_coding='" & RS3!ing_codigo & "' " & _
                 "Group by dev.dev_codmer, dev.dev_canmin, dev.dev_canmer, dev.dev_predoc, " & _
                 "         dev.dev_ptotal, dev.dev_descri, uni.uni_nomcor " & _
                 "Order by max(dev.dev_numlin)", vg_db, adOpenStatic 'pro.pro_codigo", vg_db, adOpenStatic
        Do While Not RS1.EOF
            vaSpread1.MaxRows = i
            vaSpread1.Row = i
            vaSpread1.col = 1: vaSpread1.Text = RS1!dev_codmer
            vaSpread1.col = 2: vaSpread1.Text = RS1!dev_descri
            vaSpread1.col = 3: vaSpread1.Text = RS1!uni_nomcor
            vaSpread1.col = 4: vaSpread1.Text = Format(RS1!dev_canmin, fg_Pict(9, vg_DCa))
            vaSpread1.col = 5: vaSpread1.Text = Format(RS1!dev_canmer, fg_Pict(9, vg_DCa))
            vaSpread1.col = 6: vaSpread1.Text = Format(RS1!dev_predoc, fg_Pict(9, vg_DPr))
            vaSpread1.col = 7: vaSpread1.Text = Format(RS1!dev_ptotal, fg_Pict(9, vg_DPr))
            vaSpread1.col = -1: vaSpread1.BackColor = Shape1(2).FillColor
            'Trae Stock
            RS2.Open "Select bod.bod_canmer From b_productos pro, b_bodegas bod " & _
                     "Where bod.bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                     "And bod.bod_codpro=pro.pro_codigo And pro.pro_codigo='" & Trim(RS1!dev_codmer) & "'", vg_db, adOpenStatic
            vaSpread1.col = 9
            If Not RS2.EOF Then vaSpread1.Text = RS2!bod_canmer Else vaSpread1.Text = 0
            RS2.Close: Set RS2 = Nothing
            RS1.MoveNext: i = i + 1
        Loop
        RS3.MoveNext
        RS1.Close: Set RS1 = Nothing
    Loop
    RS3.Close: Set RS3 = Nothing
    Me.MousePointer = 0
    vaSpread1.Visible = True
    vg_codigo = ""
    Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""
Case 9 'Imprimir
    I_SalDevBod Me, "SP"
Case 12 'Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
vg_swpegreceta = 0
If Err = 3034 Then Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Function MuestraFolio(Casino As String) As String
MuestraFolio = ""
If Trim(Casino) = "" Then Exit Function
RS1.Open "select tov_numdoc from b_totventas where tov_tipdoc='SP' order by tov_numdoc desc", vg_db, adOpenStatic
If Not RS1.EOF Then
    RS1.MoveFirst
    MuestraFolio = RS1!tov_numdoc + 1
Else
    MuestraFolio = 1
End If
RS1.Close: Set RS1 = Nothing
End Function

Sub Limpia()
Label1.Caption = ""
Frame1.Enabled = True
fpDouble1(0) = Format(0, fg_Pict(0, vg_DCa))
fpDouble1(1) = Format(0, fg_Pict(0, vg_DCa))
fpDateTime1(0).Text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).Text = ""
Combo1(0).ListIndex = -1
Combo1(1).ListIndex = IIf(Combo1(1).ListCount = 1, 0, -1)
vaSpread1.MaxRows = 0
Frame2.Enabled = False
vaSpread1.col = -1: vaSpread1.Row = -1
vaSpread1.Lock = True
vaSpread1.col = 5: vaSpread1.Row = -1
vaSpread1.Lock = False
fpText1(1).Enabled = ModCasino
Image1(1).Enabled = ModCasino
fpText1(1).Text = MuestraCasino(1)
fpayuda(1).Caption = MuestraCasino(2)
fpLongInteger1(0).Text = MuestraFolio(Trim(fpText1(1).Text))
Gl_Ac_Botones Me, 4, 2, ""
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long, CodPro As String, color As String, coding As String, texto As String, auxpro As String, auxing As String
Select Case Button.Index
Case 1
    If Trim(Combo1(1).Text) = "" Then MsgBox "Debe seleccionar bodega...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    
    vg_nombre = "": vg_codigo = ""
    vg_left = fpayuda(1).Left + 1920
    vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.col = 2
    texto = ""
    If vaSpread1.MaxRows > 0 Then texto = Trim(Mid(Trim(vaSpread1.Text), 1, IIf(InStr(1, vaSpread1.Text, " ") > 0, _
                                  InStr(1, vaSpread1.Text, " "), Len(Trim(vaSpread1.Text)))))
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Gen"
    B_TabEst.Text1.Text = texto
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    auxpro = vg_codigo
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        CodPro = "": color = ""
        vaSpread1.col = 1: CodPro = Trim(vaSpread1.Text)
        vaSpread1.col = 8: color = Right(Trim(vaSpread1.Text), 1)
        If Trim(CodPro) = Trim(vg_codigo) And color <> "I" Then MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Next i
    vaSpread1.Row = vaSpread1.ActiveRow
    '------- validar si existe mas de un ingrediente
    RS1.Open "select count(pri_coding) as nreg from b_productosing where pri_codpro='" & vg_codigo & "'", vg_db, adOpenStatic
    If RS1.EOF Or IsNull(RS1!NReg) Or RS1!NReg = 0 Then RS1.Close: Set RS1 = Nothing: Exit Sub
    If RS1!NReg > 1 Then
       vg_nombre = ""
       vg_left = fpayuda(1).Left + 2300
       B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "Proing"
       B_TabEst.Show 1
       If vg_codigo = "" Then Exit Sub
       auxing = vg_codigo
    ElseIf RS1!NReg = 1 Then
        RS2.Open "select pri_coding from b_productosing where pri_codpro='" & vg_codigo & "'", vg_db, adOpenStatic
        If Not RS2.EOF Then auxing = RS2!pri_coding
        RS2.Close: Set RS2 = Nothing
    End If
    RS1.Close: Set RS1 = Nothing
    RS1.Open "SELECT pro.pro_codigo, pro.pro_propon, pro.pro_nombre, ing.ing_codigo, uni.uni_nomcor, ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor " & _
             "FROM b_productos pro, a_unidad uni, b_productosing pri, b_ingrediente ing, a_unidadmed unm " & _
             "WHERE pro.pro_coduni=uni.uni_codigo and pri.pri_coding=ing.ing_codigo and pri.pri_codpro=pro.pro_codigo and ing.ing_unimed=unm.unm_codigo " & _
             "and pro.pro_codigo='" & auxpro & "' and ing.ing_codigo='" & auxing & "'", vg_db, adOpenStatic
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            For i = 1 To vaSpread1.MaxRows
                vaSpread1.Row = i
                coding = "": color = ""
                vaSpread1.col = 8: color = Right(Trim(vaSpread1.Text), 1)
                If color = "I" Then
                    vaSpread1.col = 1: coding = Trim(vaSpread1.Text)
                    If coding = RS1!ing_codigo Then
                        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
                        vaSpread1.Row = i + 1
                        vaSpread1.InsertRows vaSpread1.Row, 1
                        Exit For
                    End If
                End If
            Next i
            If i > vaSpread1.MaxRows Then
                vaSpread1.MaxRows = vaSpread1.MaxRows + 2
                vaSpread1.Row = i
                vaSpread1.InsertRows vaSpread1.Row, 2
                
                vaSpread1.col = 1: vaSpread1.Text = RS1!ing_codigo
                vaSpread1.col = 2: vaSpread1.Text = RS1!ing_nombre
                vaSpread1.col = 3: vaSpread1.Text = RS1!unm_nomcor
                vaSpread1.col = 4: vaSpread1.Text = Format(0, fg_Pict(9, vg_DCa))
                vaSpread1.col = 8: vaSpread1.Text = "NI" 'No bloquedo - Ingrediente
                vaSpread1.col = -1: vaSpread1.FontBold = True: vaSpread1.Lock = True
                vaSpread1.BackColor = Shape1(3).FillColor
                vaSpread1.Row = i + 1
            End If
                        
            vaSpread1.col = -1: vaSpread1.BackColor = Shape1(0).FillColor
            vaSpread1.col = 1: vaSpread1.Text = RS1!pro_codigo
            vaSpread1.col = 2: vaSpread1.Text = RS1!pro_nombre
            vaSpread1.col = 3: vaSpread1.Text = RS1!uni_nomcor
            vaSpread1.col = 4: vaSpread1.Text = Format(0, fg_Pict(9, vg_DCa))
            vaSpread1.col = 5: vaSpread1.Text = Format(0, fg_Pict(9, vg_DCa))
            vaSpread1.col = 6: vaSpread1.Text = Format(RS1!pro_propon, fg_Pict(9, vg_DPr))
            vaSpread1.col = 7: vaSpread1.Text = Format(0, fg_Pict(9, vg_DPr))
            vaSpread1.col = 8: vaSpread1.Text = "NU" 'No bloquedo - Usuario agrega
            'Trae Stock
            RS2.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                     "WHERE bod.bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                     "and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & Trim(RS1!pro_codigo) & "'", vg_db, adOpenStatic
            vaSpread1.col = 9
            If Not RS2.EOF Then vaSpread1.Text = RS2!bod_canmer Else vaSpread1.Text = 0
            RS2.Close: Set RS2 = Nothing
            RS1.MoveNext
            'i = i + 1
        Loop
    End If
    RS1.Close: Set RS1 = Nothing
    vaSpread1.col = 5
    vaSpread1.SetActiveCell 5, vaSpread1.Row 'vaSpread1.MaxRows
Case 2
    Dim pos1 As Long, pos2 As Long, candif As Long
    If vaSpread1.MaxRows = 0 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.col = 8: color = Right(Trim(vaSpread1.Text), 1)
    If color = "I" Then Exit Sub
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    For i = vaSpread1.ActiveRow To 1 Step -1
        color = ""
        vaSpread1.Row = i: vaSpread1.col = 8: color = Right(Trim(vaSpread1.Text), 1)
        If color = "I" Then pos1 = i: Exit For
    Next i
    For i = pos1 + 1 To vaSpread1.MaxRows
        color = ""
        vaSpread1.Row = i: vaSpread1.col = 8: color = Right(Trim(vaSpread1.Text), 1)
        pos2 = i
        If color = "I" Then pos2 = i - 1: Exit For
    Next i
    candif = pos2 - pos1
    
    vaSpread1.Row = pos1: vaSpread1.col = 1
    vaSpread1.DeleteRows IIf(candif = 1, pos1, vaSpread1.ActiveRow), IIf(candif = 1, 2, 1)
    vaSpread1.MaxRows = vaSpread1.MaxRows - IIf(candif = 1, 2, 1)
End Select
End Sub

Private Sub vaSpread1_EditChange(ByVal col As Long, ByVal Row As Long)
Dim canrea As Double, canbod As Double, canaux As Double, propon As Double, CodMer As String, color As Variant, i As Long
Dim color2 As String
vaSpread1.Row = Row
vaSpread1.col = 1: CodMer = vaSpread1.Text
vaSpread1.col = 5: canrea = Format(vaSpread1.Text, fg_Pict(9, vg_DCa))
vaSpread1.col = 6: propon = Format(vaSpread1.Text, fg_Pict(9, vg_DCa))
vaSpread1.col = 7: vaSpread1.Text = Format(canrea * propon, fg_Pict(9, vg_DCa))
vaSpread1.col = 8: color = Right(vaSpread1.Text, 1)
vaSpread1.col = 9: canbod = Format(vaSpread1.Text, fg_Pict(9, vg_DCa))
If color <> "I" Then
    canaux = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.col = 8: color2 = Right(vaSpread1.Text, 1)
        vaSpread1.col = 1
        If CodMer = vaSpread1.Text And color2 <> "I" Then
            vaSpread1.col = 5: canaux = canaux + Format(IIf(vaSpread1.Text = "", 0, vaSpread1.Text), fg_Pict(9, vg_DCa))
        End If
    Next i
    canrea = IIf(canaux > 0, canaux, canrea)
    vaSpread1.Row = Row
    If canbod - canrea >= 0 Then
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Row = i
            vaSpread1.col = 8: color2 = Right(vaSpread1.Text, 1)
            vaSpread1.col = 1
            If CodMer = vaSpread1.Text And color2 <> "I" Then
                vaSpread1.col = -1: vaSpread1.BackColor = Shape1(IIf(color = "M", 2, IIf(color = "U", 0, 3))).FillColor
                vaSpread1.col = 8: vaSpread1.Text = "N" & color 'No Bloqueado - Depende
            End If
        Next i
        Exit Sub
    End If
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.col = 8: color2 = Right(vaSpread1.Text, 1)
        vaSpread1.col = 1
        If CodMer = vaSpread1.Text And color2 <> "I" Then
            vaSpread1.col = -1: vaSpread1.BackColor = Shape1(1).FillColor
            vaSpread1.col = 8: vaSpread1.Text = "S" & color 'Bloqueado - Depende
        End If
    Next i
End If
End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
Dim i As Long, color As String
If KeyAscii <> 13 Then Exit Sub
For i = vaSpread1.ActiveRow + 1 To vaSpread1.MaxRows
    vaSpread1.Row = i: vaSpread1.col = 8: color = Right(vaSpread1.Text, 1)
    If color <> "I" Then vaSpread1.SetActiveCell vaSpread1.ActiveCol, i - 1: Exit Sub
Next i
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If Row = 0 Then Exit Sub
Dim Stock As String, Nombre As String, color As String
vaSpread1.Row = Row
vaSpread1.col = 8: color = Right(vaSpread1.Text, 1)
If color = "I" Then Exit Sub
TipWidth = 4000
ShowTip = True
MultiLine = 2
vaSpread1.col = 9: Stock = vaSpread1.Text
vaSpread1.col = 2: Nombre = vaSpread1.Text
TipText = "Bodega   : " & Trim(Left(Combo1(1).Text, 50)) & vbCrLf & _
          "Producto : " & Trim(Nombre) & vbCrLf & _
          "Stock       : " & Format(Trim(Stock), fg_Pict(9, vg_DCa))
End Sub

