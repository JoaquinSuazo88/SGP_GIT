VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_GuiVta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Guía Venta Sap"
   ClientHeight    =   7815
   ClientLeft      =   960
   ClientTop       =   2040
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Width           =   11100
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   0
         Left            =   5730
         TabIndex        =   16
         Top             =   1065
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1770
         TabIndex        =   17
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
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   1080
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
         Left            =   90
         TabIndex        =   19
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
         BackColor       =   16777215
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         MaxValue        =   "999999999"
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   5640
         TabIndex        =   20
         TabStop         =   0   'False
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
         ButtonStyle     =   1
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
         Text            =   "01/2000"
         DateCalcMethod  =   0
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
         Index           =   2
         Left            =   7440
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
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
         ControlType     =   2
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
         ButtonColor     =   -2147483633
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fpDateTime1 
         DataField       =   "º"
         Height          =   315
         Index           =   3
         Left            =   9360
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
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
         ControlType     =   2
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
         ButtonColor     =   -2147483633
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
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
         Index           =   1
         Left            =   90
         TabIndex        =   31
         Top             =   255
         Width           =   1245
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   1920
         TabIndex        =   30
         Top             =   1080
         Width           =   3345
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   1470
         Picture         =   "M_GuiVta.frx":0000
         Top             =   975
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
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
         Left            =   5730
         TabIndex        =   29
         Top             =   855
         Width           =   750
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   7110
         Picture         =   "M_GuiVta.frx":030A
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Left            =   90
         TabIndex        =   28
         Top             =   870
         Width           =   600
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
         Left            =   1770
         TabIndex        =   27
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   7560
         TabIndex        =   26
         Top             =   1065
         Width           =   3345
      End
      Begin VB.Label Label3 
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
         Index           =   9
         Left            =   5640
         TabIndex        =   25
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Facturación Inicial"
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
         Left            =   7440
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Facturación Final"
         DataField       =   "º"
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
         Index           =   11
         Left            =   9360
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   1965
         TabIndex        =   33
         Top             =   1110
         Width           =   3345
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   7605
         TabIndex        =   32
         Top             =   1095
         Width           =   3345
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5715
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   9795
      Begin VB.Frame Frame4 
         Height          =   5655
         Left            =   600
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   8175
         Begin VB.TextBox Text1 
            Height          =   5175
            Index           =   0
            Left            =   480
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   7095
         End
         Begin VB.Label label 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   5145
            Index           =   0
            Left            =   480
            TabIndex        =   13
            Top             =   360
            Width           =   7185
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Glosa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   4200
         Width           =   9555
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   3
            Left            =   2280
            TabIndex        =   6
            Top             =   240
            Width           =   4515
            _Version        =   196608
            _ExtentX        =   7964
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
            MaxLength       =   40
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   5
            Left            =   2280
            TabIndex        =   8
            Top             =   960
            Width           =   4515
            _Version        =   196608
            _ExtentX        =   7964
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
            MaxLength       =   40
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   4
            Left            =   2280
            TabIndex        =   7
            Top             =   600
            Width           =   4515
            _Version        =   196608
            _ExtentX        =   7964
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
            MaxLength       =   40
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "3.-"
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
            Left            =   1920
            TabIndex        =   11
            Top             =   1020
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "2.-"
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
            Left            =   1920
            TabIndex        =   10
            Top             =   645
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "1.-"
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
            Left            =   1920
            TabIndex        =   9
            Top             =   300
            Width           =   240
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3165
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   9525
         _Version        =   393216
         _ExtentX        =   16801
         _ExtentY        =   5583
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
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
         MaxCols         =   12
         MaxRows         =   20
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_GuiVta.frx":0614
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total General"
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
         Left            =   6720
         TabIndex        =   4
         Top             =   3600
         Width           =   1170
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   255
         Index           =   8
         Left            =   9180
         TabIndex        =   3
         Top             =   3600
         Width           =   180
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_GuiVta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim est As Boolean, fecvti As Long, fecvtf As Long, fecper As Long

Private Sub Form_Activate()
fg_descarga
TraerFechaCierre
End Sub

Private Sub Form_Load()
Me.Height = 8295
Me.Width = 11385
fg_centra Me
est = False
Me.HelpContextID = vg_OpcM
MsgTitulo = "Generación Guia Venta SAP"
Dim X As Boolean
vaSpread1.MaxRows = 0
fpDateTime1(1).text = Format(Date, "mm/yyyy")
FechaFacturacion
Gl_Mo_Botones Me, 14
Limpia 2
End Sub

Sub FechaFacturacion()
Dim ciedia As Long
'-------> Traer fecha del periodo
RS1.Open "SELECT * FROM b_cierreperiodo WHERE cie_periodo = " & Format(fpDateTime1(1).text, "yyyymm") & " AND cie_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
If Not RS1.EOF Then
   fpDateTime1(1).text = Mid(RS1!cie_periodo, 5, 2) & "/" & Mid(RS1!cie_periodo, 1, 4)
   fpDateTime1(2).text = Mid(RS1!cie_fecini, 7, 2) & "/" & Mid(RS1!cie_fecini, 5, 2) & "/" & Mid(RS1!cie_fecini, 1, 4)
   fpDateTime1(3).text = Mid(RS1!cie_fecter, 7, 2) & "/" & Mid(RS1!cie_fecter, 5, 2) & "/" & Mid(RS1!cie_fecter, 1, 4)
   fecvti = RS1!cie_fecini
   fecvtf = RS1!cie_fecter
   fecper = RS1!cie_periodo
Else
   RS1.Close: Set RS1 = Nothing
   fpDateTime1(2).text = ""
   fpDateTime1(3).text = ""
   fecvti = 0
   fecvtf = 0
   fecper = Format(Date, "yyyymm")
   MsgBox "Periodo no existe...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub
End If
RS1.Close: Set RS1 = Nothing
'-------> Traer día cierre si existen cierre venta igual 2
ciedia = 0
RS1.Open "SELECT MIN(cli_ciedia) AS cli_ciedia FROM b_clientes WHERE cli_tipo = 1 AND cli_cievta = '2' AND cli_activo = '1'", vg_db, adOpenStatic
If Not RS1.EOF Then
   ciedia = IIf(IsNull(RS1!cli_ciedia) Or Trim(RS1!cli_ciedia) = "", 0, RS1!cli_ciedia)
   If ciedia > 0 Then
      fecvti = Format(BoM("01/" & Mid(fecper, 5, 2) & "/" & Mid(fecper, 1, 4)), "yyyymm") & fg_pone_cero(ciedia + 1, 2)
      fecvtf = fecper & fg_pone_cero(ciedia, 2)
      fpDateTime1(2).text = Mid(fecvti, 7, 2) & "/" & Mid(fecvti, 5, 2) & "/" & Mid(fecvti, 1, 4)
      fpDateTime1(3).text = Mid(fecvtf, 7, 2) & "/" & Mid(fecvtf, 5, 2) & "/" & Mid(fecvtf, 1, 4)
   End If
End If
RS1.Close: Set RS1 = Nothing

'Pendiente'-------> Traer día cierre si existen cierre venta = 1
'ciedia = 0
'RS1.Open "SELECT DISTINCT cli_ciedia FROM b_clientes WHERE cli_tipo=1 AND cli_cievta='1' AND cli_activo='1'", vg_db, adOpenStatic
'If Not RS1.EOF Then
'   ciedia = IIf(IsNull(RS1!cli_ciedia) Or Trim(RS1!cli_ciedia) = "", 0, RS1!cli_ciedia)
'   If ciedia > 0 Then
'      fecvti = Format(BoM("01/" & Mid(fecper, 5, 2) & "/" & Mid(fecper, 1, 4)), "yyyymm") & fg_pone_cero(ciedia + 1, 2)
'      fecvtf = fecper & fg_pone_cero(ciedia, 2)
'      fpDateTime1(2).text = Mid(fecvti, 7, 2) & "/" & Mid(fecvti, 5, 2) & "/" & Mid(fecvti, 1, 4)
'      fpDateTime1(3).text = Mid(fecvtf, 7, 2) & "/" & Mid(fecvtf, 5, 2) & "/" & Mid(fecvtf, 1, 4)
'   End If
'End If
'RS1.Close: Set RS1 = Nothing

End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then
    Frame3.Left = (Me.Width \ 2) - (vaSpread1.Width \ 2)
    Frame1.Left = (Me.Width \ 2) - (Frame1.Width \ 2)
'    Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)
ElseIf Me.WindowState = 0 Then
    Frame3.Left = 720
    Frame1.Left = 30
    Frame2.Left = 120 '720 '360
End If
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
Select Case Index
Case 1
    If Trim(fpDateTime1(1).text) <> "" Then
       FechaFacturacion
       TraerDatosPeriodo fg_DespintaRut(LimpiaDato(Trim(fpText1(1).text)))
    End If
End Select
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
If est Then Exit Sub
BuscaDoc Val(fpLongInteger1(0).Value)
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 1
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Select
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_LostFocus(Index As Integer)
If fpText1(Index).text = "" Then Exit Sub
Select Case Index
Case 0
    RS1.Open "SELECT scl_direccion FROM b_sucursalcliente WHERE scl_codcli = '" & fg_DespintaRut(LimpiaDato(Trim(fpText1(1).text))) & "' AND scl_codigo = '" & LimpiaDato(Trim(fpText1(0).text)) & "'", vg_db, adOpenStatic
    If Not RS1.EOF Then
        Do While Not RS1.EOF
           fpayuda(Index).Caption = Trim(RS1!scl_direccion)
           RS1.MoveNext
        Loop
    Else
        RS1.Close: Set RS1 = Nothing
        fpText1(0).text = "": fpayuda(0).Caption = ""
        MsgBox "Sucursal no existe...", vbExclamation + vbOKOnly, MsgTitulo
        If fpText1(0).Enabled = True Then fpText1(0).SetFocus
        Exit Sub
    End If
    RS1.Close: Set RS1 = Nothing
Case 1
    RS1.Open "SELECT cli_nombre, cli_codigo FROM b_clientes WHERE cli_codigo = '" & fg_DespintaRut(LimpiaDato(Trim(fpText1(1).text))) & "' AND cli_tipo = 1 AND cli_clisap = '1' AND cli_activo = '1'", vg_db, adOpenStatic
    If Not RS1.EOF Then
        fpText1(1).text = fg_PintaRut(RS1!cli_codigo)
        fpayuda(Index).Caption = Trim(RS1!cli_nombre)
    Else
        RS1.Close: Set RS1 = Nothing
        fpText1(1).text = "": fpayuda(1).Caption = ""
        MsgBox "Cliente no existe...", vbExclamation + vbOKOnly, MsgTitulo
        Exit Sub
    End If
    RS1.Close: Set RS1 = Nothing
    TraerDatosPeriodo fg_DespintaRut(LimpiaDato(Trim(fpText1(1).text)))
'Case 2
'    BuscaDoc fpText1(2).text
End Select
End Sub

Private Sub Image1_Click(Index As Integer)
Dim Variable As String
vg_codigo = IIf(Index = 0, fg_DespintaRut(Trim(LimpiaDato(fpText1(1).text))), "")
vg_left = fpayuda(Index).Left + 1920
Variable = IIf(Index = 0, "Sucursal", "Cliente")
B_TabEst.LlenaDatos IIf(Index = 0, "b_sucursalcliente", "b_clientes"), IIf(Index = 0, "scl_", "cli_"), Variable, IIf(Index = 1, "CliSap", Variable)
B_TabEst.Show 1
Me.Refresh
If Trim(vg_codigo) = "" Then Exit Sub
fpText1(Index) = IIf(Index = 1, fg_PintaRut(vg_codigo), Trim(vg_codigo))
fpayuda(Index).Caption = vg_nombre
Select Case Index
Case 0
    If Trim(fpText1(Index).text) = "" Then Exit Sub
    If fpText1(3).Enabled = True Then fpText1(3).SetFocus
Case 1
    TraerDatosPeriodo fg_DespintaRut(LimpiaDato(Trim(fpText1(1).text)))
    If fpText1(0).Enabled = True Then fpText1(0).SetFocus
    fpText1_LostFocus 0
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim rutcli As String, codsuc As String, NumDoc As Long, i As Long, canact As Double, grlvta As Long, codcli As String, sql1 As String, sql2 As String, sql3 As String
Dim numlin As Long, codser As Long, codreg As Long, nomser As String, desser As String, codsap As String, racsgp As Long, racgui As Long, presgp As Double, pregui As Double
On Error GoTo Man_Error
Frame4.Visible = False
Text1(0).Visible = True
fecpro = Format(fpDateTime1(0).Value, "dd/mm/yyyy")
If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
Select Case Button.Index
Case 1, 6 '-------> Nuevo
    If Button.Index = 6 And vaSpread1.MaxRows > 0 Then If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    Limpia IIf(Button.Index = 1, 6, 2)
    If fpLongInteger1(0).Enabled = True Then fpLongInteger1(0).SetFocus
Case 8 '-------> Graba
    If vaSpread1.MaxRows < 1 Or Trim(fpText1(0).text) = "" Or Trim(fpayuda(1).Caption) = "" Or Trim(fpayuda(0).Caption) = "" Or Trim(fpText1(1).text) = "" Or Trim(fpDateTime1(0).text) = "" Then MsgBox "Debe ingresar dato importante...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
'    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), vg_codbod, 0) Then MsgBox "Documento no corresponde al periodo : " & VgLinea & VgLinea & CierreFecha, vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
'    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), vg_codbod, 6) Then MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
'    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), vg_codbod, 8) Then MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Label3(0).Caption = "0" Then MsgBox "El total del documento debe ser mayor a 0...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 3
        If Trim(vaSpread1.text) = "" Then MsgBox "Descripción de servicio, debe ser distinto de blanco...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        vaSpread1.Col = 4
        If Trim(vaSpread1.text) = "" Then MsgBox "Còdigo de SAP, debe ser distinto de blanco...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Next i
    If MsgBox("Desea grabar...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
'    '-------> Traer fecha del periodo
'    fecper = 0
'    RS1.Open "SELECT * FROM b_cierreperiodo WHERE cie_estado=1 AND cie_cencos='" & MuestraCasino(1) & "'", vg_db, adOpenStatic
'    If Not RS1.EOF Then fecper = RS1!cie_periodo
'    RS1.Close: Set RS1 = Nothing
    vg_db.BeginTrans
    rutcli = fg_DespintaRut(Trim(LimpiaDato(fpText1(1).text)))
    codsuc = Trim(LimpiaDato(fpText1(0).text))
    NumDoc = Val(fpLongInteger1(0).Value)
    '-------> Encabezado
    sql1 = IIf(vg_tipbase = "1", " '" & CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")) & "' ", " '" & Format(fpDateTime1(0).text, "yyyymmdd") & "' ")
    sql2 = IIf(vg_tipbase = "1", " '" & CDate(Format(fpDateTime1(2).text, "dd/mm/yyyy")) & "' ", " '" & Format(fpDateTime1(2).text, "yyyymmdd") & "' ")
    sql3 = IIf(vg_tipbase = "1", " '" & CDate(Format(fpDateTime1(3).text, "dd/mm/yyyy")) & "' ", " '" & Format(fpDateTime1(3).text, "yyyymmdd") & "' ")
    vg_db.Execute "INSERT INTO b_totguiavta (tgv_rutcli, tgv_codsuc, tgv_numdoc, tgv_fecing, tgv_perfac, tgv_fecini, tgv_fecfin, tgv_glosa1, tgv_glosa2, tgv_glosa3) " & _
                  "VALUES ('" & rutcli & "', '" & codsuc & "', " & NumDoc & ", " & sql1 & ", " & Format(fpDateTime1(1).text, "yyyymm") & ", " & sql2 & ", " & sql3 & ", '" & Trim(LimpiaDato(fpText1(3).text)) & "', '" & Trim(LimpiaDato(fpText1(4).text)) & "', '" & Trim(LimpiaDato(fpText1(5).text)) & "' )"
    '-------> Detalle
    total = 0
    numlin = 1
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1: codser = vaSpread1.text
        vaSpread1.Col = 2: nomser = Trim(LimpiaDato(vaSpread1.text))
        vaSpread1.Col = 3: desser = Trim(LimpiaDato(Mid(vaSpread1.text, 1, 40)))
        vaSpread1.Col = 4: codsap = Trim(LimpiaDato(Mid(vaSpread1.text, 1, 40)))
        vaSpread1.Col = 5: racsgp = vaSpread1.text
        vaSpread1.Col = 6: racgui = vaSpread1.text
        vaSpread1.Col = 7: presgp = vaSpread1.text
        vaSpread1.Col = 8: pregui = vaSpread1.text
        vaSpread1.Col = 10: codreg = vaSpread1.text
        vaSpread1.Col = 12: codcli = vaSpread1.text
        vg_db.Execute "INSERT INTO b_detguiavta (dgv_rutcli, dgv_codsuc, dgv_numdoc, dgv_numlin, dgv_codreg, dgv_codser, dgv_nomser, dgv_desser, dgv_codsap, dgv_racsgp, dgv_racguia, dgv_presgp, dgv_preguia, dgv_codcli) " & _
                      "VALUES ('" & rutcli & "', '" & codsuc & "', " & NumDoc & ", " & i & ", " & codreg & ", " & codser & ", '" & Mid(nomser, 1, 40) & "', '" & Mid(desser, 1, 40) & "', '" & codsap & "', " & racsgp & ", " & racgui & ", " & presgp & ", " & pregui & ", '" & codcli & "')"
        '-------> Grabar minuta raciones
        vg_db.Execute "UPDATE b_minutaraciones SET mir_nroguia = " & NumDoc & ", mir_codcli = '" & rutcli & "' WHERE mir_cencos = '" & MuestraCasino(1) & "' AND mir_codreg = " & codreg & " AND mir_codser = " & codser & " AND mir_fecmin >= " & fecvti & " AND mir_fecmin <= " & fecvtf & " AND mir_rutcli = '" & codcli & "'"
    Next i
    vg_db.CommitTrans
    Gl_Ac_Botones Me, 14, 3, ""
    '-------> Validar si tienes activada la opción de envío
    Toolbar1.Buttons(15).Enabled = ValidarOpEnvio(MuestraCasino(1), 3)
    Toolbar1.Buttons(15).ToolTipText = IIf(ValidarOpEnvio(MuestraCasino(1), 3), "Enviar Guía Venta SAP", "")
    Frame1.Enabled = False
    Frame2.Enabled = False
    vaSpread1.Col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    I_GuiaVentaSAP Me
    MsgTitulo = "Generación Guia Venta SAP"
Case 3 '-------> Eliminar
    If MsgBox("Elimina documento...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
'    '------- Traer fecha del periodo
'    fecper = 0
'    RS1.Open "SELECT * FROM b_cierreperiodo WHERE cie_estado=1 AND cie_cencos='" & MuestraCasino(1) & "'", vg_db, adOpenStatic
'    If Not RS1.EOF Then fecper = RS1!cie_periodo
'    RS1.Close: Set RS1 = Nothing
    vg_db.BeginTrans
    '------- Grabar minuta raciones
    vg_db.Execute "UPDATE b_minutaraciones SET mir_nroguia=0, mir_codcli='' WHERE mir_cencos='" & MuestraCasino(1) & "' AND mir_nroguia=" & Val(fpLongInteger1(0).Value) & " AND mir_codcli='" & fg_DespintaRut(Trim(LimpiaDato(fpText1(1).text))) & "' AND mir_fecmin>=" & fecvti & " AND mir_fecmin<=" & fecvtf & ""
    vg_db.Execute "DELETE b_detguiavta FROM b_detguiavta WHERE dgv_rutcli='" & fg_DespintaRut(Trim(LimpiaDato(fpText1(1).text))) & "' " & _
                  "AND dgv_numdoc=" & Val(fpLongInteger1(0).Value) & ""
    vg_db.Execute "DELETE b_totguiavta FROM b_totguiavta WHERE tgv_rutcli='" & fg_DespintaRut(Trim(LimpiaDato(fpText1(1).text))) & "' " & _
                  "AND tgv_numdoc=" & Val(fpLongInteger1(0).Value) & ""
    vg_db.CommitTrans
    Limpia 2
Case 11 '-------> Busqueda
    vg_codigo = fg_DespintaRut(Trim(LimpiaDato(fpText1(1).text)))
    vg_nombre = "GV"
    B_SalBod.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    BuscaDoc Val(vg_codigo)
Case 12 '-------> Imprimir
    est = True
    I_GuiaVentaSAP Me
    est = False
    MsgTitulo = "Generación Guia Venta SAP"
Case 15 '-------> Enviar archivo guia venta SAP
'    If Not isNetwork(NETWORK_ALIVE_LAN) Then MsgBox "No hay conexión a internet, proceso cancelado", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
     If Not isInternetConnected(False, False, False) Then MsgBox "No hay conexión a internet, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Toolbar1.Enabled = False
    If Not GenerarArcSap(Val(fpLongInteger1(0).Value)) Then
       Text1(0).text = Text1(0).text & FechaHora & "Generación envió Finalizado Con Problema" & VgLinea
       I_EnvioSap "2"
       Toolbar1.Enabled = True
       Exit Sub
    Else
       Text1(0).text = Text1(0).text & FechaHora & "Generación envió Finalizado Sin Problema" & VgLinea
       I_EnvioSap "2"
       Toolbar1.Buttons(15).Enabled = False: Toolbar1.Buttons(15).ToolTipText = ""
    End If
    Toolbar1.Enabled = True
Case 17 '-------> Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Sub BuscaDoc(codigo As Long)
Dim NumDoc As Long
NumDoc = Trim(codigo)
'-------> Encabezado
RS2.Open "SELECT a.tgv_numdoc, a.tgv_fecing, a.tgv_rutcli, a.tgv_codsuc, b.cli_nombre, c.scl_direccion, a.tgv_perfac, a.tgv_fecini, tgv_fecfin, a.tgv_glosa1, a.tgv_glosa2, a.tgv_glosa3 " & _
         "FROM b_totguiavta a, b_clientes b, b_sucursalcliente c " & _
         "WHERE a.tgv_numdoc = " & codigo & " " & _
         "AND   a.tgv_rutcli = b.cli_codigo " & _
         "AND   a.tgv_codsuc = c.scl_codigo " & _
         "AND   c.scl_codcli = b.cli_codigo", vg_db, adOpenStatic
If Not RS2.EOF Then
    Frame1.Enabled = False
    Frame2.Enabled = False
    vaSpread1.Col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    vaSpread1.MaxRows = 0
    Do While Not RS2.EOF
        est = True
        fpLongInteger1(0).Value = RS2!tgv_numdoc
        fpDateTime1(0).text = RS2!tgv_fecing
        fpText1(1).text = fg_PintaRut(RS2!tgv_rutcli)
        fpText1_LostFocus 1
        fpText1(0).text = Trim(RS2!tgv_codsuc)
        fpText1_LostFocus 0
        fpText1(3).text = Trim(RS2!tgv_glosa1)
        fpText1(4).text = Trim(RS2!tgv_glosa2)
        fpText1(5).text = Trim(RS2!tgv_glosa3)
        fpDateTime1(1).text = Mid(RS2!tgv_perfac, 5, 2) & "/" & Mid(RS2!tgv_perfac, 1, 4)
        fpDateTime1(2).text = IIf(IsNull(RS2!tgv_fecini), "", RS2!tgv_fecini)
        fpDateTime1(3).text = IIf(IsNull(RS2!tgv_fecfin), "", RS2!tgv_fecfin)
        est = False
        RS2.MoveNext
    Loop
Else
    Gl_Ac_Botones Me, 14, 6, ""
    Toolbar1.Buttons(15).Enabled = False
    Toolbar1.Buttons(15).ToolTipText = ""
    RS2.Close: Set RS2 = Nothing
    Exit Sub
End If
RS2.Close: Set RS2 = Nothing
'-------> Detalle
RS2.Open "SELECT a.*, b.* " & _
         "FROM b_detguiavta a, a_regimen b " & _
         "WHERE a.dgv_codreg = b.reg_codigo AND a.dgv_rutcli = '" & fg_DespintaRut(LimpiaDato(Trim(fpText1(1).text))) & "'" & _
         "AND   a.dgv_numdoc = " & codigo & " ORDER BY a.dgv_numlin", vg_db, adOpenStatic
If Not RS2.EOF Then
    i = 1
    Frame3.Caption = "Regimen : " & RS2!reg_codigo & " - " & Trim(RS2!reg_nombre)
    Do While Not RS2.EOF
        vaSpread1.MaxRows = i
        vaSpread1.Row = i
        vaSpread1.Col = 1: vaSpread1.text = RS2!dgv_codser
        vaSpread1.Col = 2: vaSpread1.text = Trim(RS2!dgv_nomser)
        vaSpread1.Col = 3: vaSpread1.text = Trim(RS2!dgv_desser)
        vaSpread1.Col = 4: vaSpread1.text = Trim(RS2!dgv_codsap)
        vaSpread1.Col = 5: vaSpread1.text = Format(RS2!dgv_racsgp, fg_Pict(6, 0))
        vaSpread1.Col = 6: vaSpread1.text = Format(RS2!dgv_racguia, fg_Pict(6, 0))
        vaSpread1.Col = 7: vaSpread1.text = Format(RS2!dgv_presgp, fg_Pict(6, 2))
        vaSpread1.Col = 8: vaSpread1.text = Format(RS2!dgv_preguia, fg_Pict(6, 2))
        vaSpread1.Col = 9: vaSpread1.text = Format((RS2!dgv_racguia * RS2!dgv_preguia), fg_Pict(9, 0))
        vaSpread1.Col = 10: vaSpread1.text = RS2!reg_codigo
        vaSpread1.Col = 11: vaSpread1.text = Trim(RS2!reg_nombre)
        vaSpread1.Col = 12: vaSpread1.text = RS2!dgv_codcli
        RS2.MoveNext: i = i + 1
    Loop
End If
RS2.Close: Set RS2 = Nothing
TotalGuia
vg_codigo = ""
Gl_Ac_Botones Me, 14, 3, ""
'-------> Validar si tienes activada la opción de envío
Toolbar1.Buttons(15).Enabled = ValidarOpEnvio(MuestraCasino(1), 3)
Toolbar1.Buttons(15).ToolTipText = IIf(ValidarOpEnvio(MuestraCasino(1), 3), "Enviar Guía Venta SAP", "")
If Toolbar1.Buttons(15).Enabled = False Then Exit Sub
Dim sql1 As String
'sql1 = IIf(vg_tipbase = "1", " '" & "G" & numdoc & "' ", " '" & "G" + numdoc & "' ")
sql1 = IIf(vg_tipbase = "1", " '" & "G" & NumDoc & "' ", " '" & "G" & NumDoc & "' ")
RS1.Open "SELECT DISTINCT cencos from log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND tipo_proceso = '4' AND tipo_documento = 'GD' AND num_documento = " & sql1 & " AND estado = '1'", vg_db, adOpenStatic
Toolbar1.Buttons(15).Enabled = IIf(RS1.EOF, True, False)
Toolbar1.Buttons(15).ToolTipText = IIf(RS1.EOF, "Enviar Guía Venta SAP", "")
RS1.Close: Set RS1 = Nothing
End Sub

Sub Limpia(op As Integer)
est = True
'Label1.Caption = ""
Frame3.Caption = ""
Frame1.Enabled = True
fpLongInteger1(0).Enabled = True
Frame2.Enabled = True
fpText1(5).text = ""
fpText1(4).text = ""
fpText1(3).text = ""
fpLongInteger1(0).Value = ""
fpText1(1).text = ""
fpText1(0).text = ""
fpayuda(1).Caption = ""
fpayuda(0).Caption = ""
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
vaSpread1.MaxRows = 0
vaSpread1.Row = -1
vaSpread1.Col = -1
vaSpread1.Lock = False
Label3(8).Caption = "0"
Gl_Ac_Botones Me, 14, op, ""
est = False
End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
Dim nrorac As Long, preguia As Double
vaSpread1.Row = Row
vaSpread1.Col = 6: nrorac = Format(vaSpread1.text, fg_Pict(9, vg_DPr))
vaSpread1.Col = 8: preguia = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
Select Case Col
Case 6, 8
     vaSpread1.Col = 9: vaSpread1.text = Format(nrorac * preguia, fg_Pict(9, vg_DPr))
     TotalGuia
     vaSpread1.Row = Row
End Select
End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Or vaSpread1.MaxRows = 0 Then Exit Sub
'If vaSpread1.ActiveCol = 4 Then vaSpread1.SetActiveCell 5, vaSpread1.ActiveRow - 1: Exit Sub
'If vaSpread1.ActiveCol = 5 Then vaSpread1.SetActiveCell 6, vaSpread1.ActiveRow - 1: Exit Sub
'If vaSpread1.ActiveCol = 6 And vaSpread1.ActiveRow - 1 <> vaSpread1.MaxRows Then vaSpread1.SetActiveCell 4, vaSpread1.ActiveRow
End Sub

Sub TraerDatosPeriodo(rutcli As String)
If est Then Exit Sub
Dim nArch As String, fecpin As Long, fecpfi As Long, fanomes As Long, sql1 As String, sql2 As String
''-------> Traer fecha del periodo
fecpin = 0: fecpfin = 0: fanomes = 0
RS1.Open "SELECT * FROM b_cierreperiodo WHERE cie_periodo = " & Format(fpDateTime1(1).text, "yyyymm") & " AND cie_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
If Not RS1.EOF Then fecper = RS1!cie_periodo: fecpin = RS1!cie_fecini: fecpfi = RS1!cie_fecter
RS1.Close: Set RS1 = Nothing
If fecpin = 0 Or fecpfi = 0 Then MsgBox "No existe fecha inicio y termino del periodo contable, proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub

fanomes = Format(BoM("01/" & Mid(fecper, 5, 2) & "/" & Mid(fecper, 1, 4)), "yyyymm")
'-------> Validar si fue procesado
RS1.Open "SELECT DISTINCT a.mir_nroguia FROM b_minutaraciones a WHERE a.mir_codcli = '" & rutcli & "' AND a.mir_cencos = '" & MuestraCasino(1) & "' AND a.mir_fecmin >= " & fecvti & " AND a.mir_fecmin <= " & fecvtf & " AND (a.mir_nroguia) IS NOT NULL AND a.mir_nroguia > 0", vg_db, adOpenStatic
If Not RS1.EOF Then MsgBox "Cliente fue procesado, con Nº documento : " & Trim(RS1!mir_nroguia), vbExclamation + vbOKOnly, MsgTitulo: RS1.Close: Set RS1 = Nothing: fpText1(1).text = "": fpayuda(1).Caption = "": Exit Sub
RS1.Close: Set RS1 = Nothing

nArch = Trim(vg_NUsr) & "_tmp_guiavta"
fg_CheckTmp (nArch)
sql1 = IIf(vg_tipbase = "1", " iif(cli.cli_cievta='1', " & fecpin & ", " & fanomes & " & Right('00' + (cli.cli_ciedia + 1), 2)) ", " CASE WHEN cli.cli_cievta = '1' THEN " & fecpin & " ELSE CASE WHEN cli.cli_ciedia < 10 THEN convert(int," & fanomes & " + '0' + convert(varchar(1),(cli.cli_ciedia + 1))) ELSE convert(int, " & fanomes & " + convert(varchar(02),(cli.cli_ciedia + 1))) END END ")
sql2 = IIf(vg_tipbase = "1", " iif(cli.cli_cievta='1', " & fecpfi & ", " & fecper & " & Right('00' + cli.cli_ciedia, 2)) ", " CASE WHEN cli.cli_cievta = '1' THEN " & fecpfi & " ELSE CASE WHEN cli.cli_ciedia < 10 THEN  convert(int, " & fecper & " + '0' + convert(varchar(1),cli.cli_ciedia)) ELSE convert(int, " & fecper & " + convert(varchar(02),cli.cli_ciedia))  END END ")
vg_db.Execute "SELECT mir.mir_cencos, mir.mir_codreg, mir.mir_codser, mir.mir_fecmin, mir.mir_rutcli, 0 AS mir_codcco, mir.mir_nrorac, MAX(prv.prv_fecvig) AS prv_fecvig INTO " & nArch & " " & _
              "FROM b_minutaraciones mir, b_preciovta prv, b_clientes cli " & _
              "WHERE mir.mir_rutcli = prv.prv_rutcli " & _
              "AND   mir.mir_codser = prv.prv_codser " & _
              "AND   mir.mir_codreg = prv.prv_codreg " & _
              "AND   mir.mir_cencos = prv.prv_cencos " & _
              "AND   mir.mir_fecmin >= prv.prv_fecvig " & _
              "AND   mir.mir_rutcli = cli.cli_codigo " & _
              "AND   mir.mir_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   mir.mir_fecmin >= " & sql1 & " " & _
              "AND   mir.mir_fecmin <= " & sql2 & " " & _
              "AND   mir.mir_nrorac > 0 " & _
              "AND  ((mir.mir_nroguia) IS NULL OR mir.mir_nroguia = 0) " & _
              "GROUP BY mir.mir_cencos, mir.mir_codreg, mir.mir_codser, mir.mir_fecmin, mir.mir_rutcli, mir.mir_nrorac"
'vg_db.Execute "SELECT mir.mir_cencos, mir.mir_codreg, mir.mir_codser, mir.mir_fecmin, mir.mir_rutcli, 0 AS mir_codcco, mir.mir_nrorac, MAX(prv.prv_fecvig) AS prv_fecvig INTO " & nArch & " " & _
'              "FROM b_minutaraciones mir INNER JOIN b_preciovta prv ON (mir.mir_rutcli=prv.prv_rutcli) AND (mir.mir_codser=prv.prv_codser) AND (mir.mir_codreg=prv.prv_codreg) AND (mir.mir_cencos=prv.prv_cencos AND mir.mir_fecmin>=prv.prv_fecvig) " & _
'              "WHERE mir.mir_cencos='" & MuestraCasino(1) & "' " & _
'              "AND   mir.mir_fecmin>=" & fecvti & " AND mir.mir_fecmin<=" & fecvtf & " " & _
'              "AND   mir.mir_nrorac>0 " & _
'              "AND (ISNULL(mir.mir_nroguia) OR mir.mir_nroguia=0) " & _
'              "GROUP BY mir.mir_cencos, mir.mir_codreg, mir.mir_codser, mir.mir_fecmin, mir.mir_rutcli, mir.mir_nrorac"

RS1.Open "SELECT d.reg_codigo, d.reg_nombre, c.ser_codigo, c.ser_nombre, c.ser_orden, c.ser_codsap, e.cli_codigo, e.cli_nombre, b.prv_preven, SUM(a.mir_nrorac) AS mir_nrorac " & _
         "FROM " & nArch & " a, b_preciovta b, a_servicio c, a_regimen d, b_clientes e " & _
         "WHERE a.mir_codreg = d.reg_codigo " & _
         "AND   a.mir_codser = c.ser_codigo " & _
         "AND   a.mir_rutcli = e.cli_codigo " & _
         "AND  (e.cli_codcli = '" & rutcli & "' OR e.cli_codigo = '" & rutcli & "') " & _
         "AND   a.mir_cencos = b.prv_cencos " & _
         "AND   a.mir_codreg = b.prv_codreg " & _
         "AND   a.mir_codser = b.prv_codser " & _
         "AND   a.mir_rutcli = b.prv_rutcli " & _
         "AND   a.prv_fecvig = b.prv_fecvig " & _
         "AND   c.ser_facturable = '1' AND e.cli_activo = '1' " & _
         "GROUP BY d.reg_codigo, d.reg_nombre, c.ser_codigo, c.ser_nombre, c.ser_orden, c.ser_codsap, e.cli_codigo, e.cli_nombre, b.prv_preven ORDER BY d.reg_codigo, c.ser_orden", vg_db, adOpenStatic
vaSpread1.MaxRows = 0
If Not RS1.EOF And Trim(rutcli) <> "" Then
   Frame3.Caption = "Regimen : " & RS1!reg_codigo & " - " & Trim(RS1!reg_nombre)
   Do While Not RS1.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 1: vaSpread1.text = RS1!ser_codigo
      vaSpread1.Col = 2: vaSpread1.text = Trim(RS1!ser_nombre) & " (" & Trim(RS1!cli_nombre) & ")"
      vaSpread1.Col = 3: vaSpread1.ForeColor = &HFF0000: vaSpread1.Lock = False: vaSpread1.text = Trim(RS1!ser_nombre) & " (" & Trim(RS1!cli_nombre) & ")"
      vaSpread1.Col = 4: vaSpread1.ForeColor = &HFF0000: vaSpread1.Lock = False: vaSpread1.text = IIf(IsNull(RS1!ser_codsap), "", Trim(RS1!ser_codsap))
      vaSpread1.Col = 5: vaSpread1.text = Format(RS1!mir_nrorac, fg_Pict(6, 0))
      vaSpread1.Col = 6: vaSpread1.ForeColor = &HFF0000: vaSpread1.Lock = False: vaSpread1.text = Format(RS1!mir_nrorac, fg_Pict(6, 0))
      vaSpread1.Col = 7: vaSpread1.text = Format(RS1!prv_preven, fg_Pict(6, 2))
      vaSpread1.Col = 8: vaSpread1.ForeColor = &HFF0000: vaSpread1.Lock = False: vaSpread1.text = Format(RS1!prv_preven, fg_Pict(6, 2))
      vaSpread1.Col = 9: vaSpread1.text = Format(RS1!prv_preven * RS1!mir_nrorac, fg_Pict(9, 0))
      vaSpread1.Col = 10: vaSpread1.text = RS1!reg_codigo
      vaSpread1.Col = 11: vaSpread1.text = Trim(RS1!reg_nombre)
      vaSpread1.Col = 12: vaSpread1.text = RS1!cli_codigo
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
TotalGuia
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread1.MaxRows < 1 Or est Then Exit Sub
Dim codreg As Long, nomreg As String
vaSpread1.Row = NewRow
vaSpread1.Col = 10: codreg = Val(vaSpread1.text)
vaSpread1.Col = 11: nomreg = vaSpread1.text
Frame3.Caption = "Regimen : " & codreg & " - " & nomreg
End Sub

Sub TotalGuia()
Dim grlvta As Long
grlvta = 0
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i: vaSpread1.Col = 9: grlvta = (grlvta + vaSpread1.text)
Next i
Label3(8).Caption = Format(grlvta, fg_Pict(9, vg_DPr))
End Sub

Function GenerarArcSap(NumDoc As Long) As Boolean
Dim cpedvta As String, corgvta As String, ccandis As String, codsec As String, cofivta As String, cpedvta1 As String, codcli As String, cdesmer As String, censum As String, fecent As String
Dim codmat As String, desmat As String, Cantidad As String, prevta As String, moneda As String, glosa1 As String, glosa2 As String, glosa3 As String
Dim nomarch  As String, numlin As Long, codigo As Long, tipmon As String, numero As Long
Dim parametro1 As String, parametro2 As String, parametro3 As String, vRet As Variant
Dim Segundos As Byte
On Error GoTo Man_EnvioGuiaVenta
DoEvents
GenerarArcSap = False
cpedvta = "": corgvta = "": ccandis = "": codsec = "": cofivta = "": cpedvta1 = "": codcli = "": cdesmer = "": censum = "": fecent = ""
codmat = "": desmat = "": Cantidad = "": prevta = "": moneda = "": glosa1 = "": glosa2 = "": glosa3 = ""

'-------> Abrir mensaje de text
Frame4.Visible = True
Text1(0).Visible = True
'Text1(0).text = FechaHora & "PC : " & Environ("COMPUTERNAME") & VgLinea
Text1(0).text = FechaHora & "CENCO : " & MuestraCasino(1) & " - " & MuestraCasino(2) & VgLinea
Text1(0).text = Text1(0).text & FechaHora & "USUARIO : " & Environ("USERNAME") & VgLinea
Text1(0).text = Text1(0).text & FechaHora & "Inicio del Proceso. Guía Ventas : " & Val(fpLongInteger1(0).Value) & VgLinea

'-------> Validar si existe usuario sap
RS1.Open "SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'sapusu'", vg_db, adOpenStatic
If RS1.EOF Then
   RS1.Close: Set RS1 = Nothing
   Text1(0).text = Text1(0).text & FechaHora & "No tiene creado usuario, para Web Service" & VgLinea
   GenerarArcInvSap = False: Exit Function
ElseIf IsNull(RS1!par_valor) Or Trim(RS1!par_valor) = "" Then
   RS1.Close: Set RS1 = Nothing
   Text1(0).text = Text1(0).text & FechaHora & "usuario fue borrado, para Web Service" & VgLinea
   GenerarArcInvSap = False: Exit Function
End If
RS1.Close: Set RS1 = Nothing
'-------> Validar si existe password sap
RS1.Open "SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'sappas'", vg_db, adOpenStatic
If RS1.EOF Then
   RS1.Close: Set RS1 = Nothing
   Text1(0).text = Text1(0).text & FechaHora & "No tiene creado password, para Web Service" & VgLinea
   GenerarArcInvSap = False: Exit Function
ElseIf IsNull(RS1!par_valor) Or Trim(RS1!par_valor) = "" Then
   RS1.Close: Set RS1 = Nothing
   Text1(0).text = Text1(0).text & FechaHora & "Password fue borrada, para Web Service" & VgLinea
   GenerarArcInvSap = False: Exit Function
End If
RS1.Close: Set RS1 = Nothing
'-------> Traer guias de ventas
RS1.Open "SELECT a.*, b.*, c.* " & _
         "FROM b_totguiavta a, b_detguiavta b, a_regimen c " & _
         "WHERE a.tgv_rutcli = b.dgv_rutcli " & _
         "AND   a.tgv_codsuc = b.dgv_codsuc " & _
         "AND   a.tgv_numdoc = b.dgv_numdoc " & _
         "AND   b.dgv_codreg = c.reg_codigo " & _
         "AND   a.tgv_rutcli = '" & fg_DespintaRut(LimpiaDato(Trim(fpText1(1).text))) & "'" & _
         "AND   a.tgv_numdoc = " & NumDoc & " AND b.dgv_preguia > 0 AND (a.tgv_envsap = '0' OR (a.tgv_envsap) IS NULL) ORDER BY b.dgv_numlin", vg_db, adOpenStatic
If RS1.EOF Then
   fg_descarga
   RS1.Close: Set RS1 = Nothing
   Text1(0).text = Text1(0).text & FechaHora & "No existe Información a procesar." & VgLinea
   GenerarArcSap = False
   Exit Function
End If

codigo = 0
RS2.Open "SELECT gvt_codigo FROM sap_guiavta ORDER BY gvt_codigo DESC", vg_db, adOpenStatic
If Not RS2.EOF Then RS2.MoveFirst: codigo = RS2!gvt_codigo + 1 Else codigo = 1
RS2.Close: Set RS2 = Nothing

'-------> Mover parametro Web Service
parametro1 = "4"
parametro2 = codigo
parametro3 = MuestraCasino(1)

numlin = 0
Do While Not RS1.EOF
   DoEvents
   '------> Grabar texto encabezado
   cpedvta = "G" & Trim(RS1!tgv_numdoc)
   codcli = Trim(RS1!tgv_rutcli)
   cdesmer = Trim(RS1!tgv_codsuc)
   censum = Trim(MuestraCasino(1))
   fecent = Format(RS1!tgv_fecing, "ddmmyyyy")
   '------> Detalle
   codmat = Trim(Mid(RS1!dgv_codsap, 1, 18))
   desmat = Trim(Mid(RS1!dgv_desser, 1, 40))
   Cantidad = RS1!dgv_racguia
   prevta = RS1!dgv_preguia
   tipmon = vg_tipmonsap '"CLP"
   glosa1 = IIf(IsNull(RS1!tgv_glosa1), "", Trim(RS1!tgv_glosa1))
   glosa2 = IIf(IsNull(RS1!tgv_glosa2), "", Trim(RS1!tgv_glosa2))
   glosa3 = IIf(IsNull(RS1!tgv_glosa3), "", Trim(RS1!tgv_glosa3))
   numlin = RS1!dgv_numlin
   vg_db.Execute "INSERT INTO sap_guiavta VALUES ('" & codigo & "', " & numlin & ", '" & cpedvta & "', '" & codcli & "', " & _
                 "'" & cdesmer & "', '" & censum & "', '" & fecent & "', '" & codmat & "', '" & desmat & "', " & _
                 "'" & Cantidad & "', '" & prevta & "', '" & tipmon & "', '" & glosa1 & "', '" & glosa2 & "', '" & glosa3 & "')"
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing

'------> Grabar log proceso
Dim fecenv As String
fecenv = IIf(vg_tipbase = "1", Format(Date, "dd-mm-yyyy") & " " & Format(Time, "h:m:s"), Format(Date, "yyyymmdd") & " " & Format(Time, "h:m:s"))
vg_db.BeginTrans
'RS1.Open "SELECT DISTINCT a.* " & _
'         "FROM b_totguiavta a, b_detguiavta b " & _
'         "WHERE a.tgv_rutcli = b.dgv_rutcli " & _
'         "AND   a.tgv_codsuc = b.dgv_codsuc " & _
'         "AND   a.tgv_numdoc = b.dgv_numdoc " & _
'         "AND   a.tgv_rutcli = '" & fg_DespintaRut(LimpiaDato(Trim(fpText1(1).text))) & "'" & _
'         "AND   a.tgv_numdoc = " & numdoc & "", vg_db, adOpenStatic
'If Not RS1.EOF Then
'   Do While Not RS1.EOF
numero = 0
RS2.Open "SELECT numero FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' ORDER BY numero DESC", vg_db, adOpenStatic
If Not RS2.EOF Then RS2.MoveFirst: numero = RS2!numero + 1 Else numero = 1
RS2.Close: Set RS2 = Nothing
vg_db.Execute "INSERT INTO log_procesos (cencos, numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mensaje, envio) " & _
              "VALUES ('" & MuestraCasino(1) & "', " & numero & ", '" & fecenv & "', '4', '" & fg_DespintaRut(LimpiaDato(Trim(fpText1(1).text))) & "', 'GD' , '" & "G" & NumDoc & "',  null, '0', '', " & codigo & ")"
'      RS1.MoveNext
'   Loop
'End If
'RS1.Close: Set RS1 = Nothing
vg_db.CommitTrans

'-------> Proceso envio Web Service
DoEvents
If vg_tipbase = "1" Then
'   vRet = Shell(Trim(dir_trabajo) & "WsSapPortal.exe " & Trim(parametro1) & "|" & Trim(parametro2) & "|" & Trim(parametro3) & "|" & Trim(dir_trabajo) & "|" & "" & "|" & "" & "|" & "" & "|" & "" & "|")
   vRet = Shell(Trim(dir_trabajo) & "WsSapPortal.exe " & Trim(parametro1) & "|" & Trim(parametro2) & "|" & Trim(parametro3) & "|" & LCase(App.Path) & "\" & "|" & "" & "|" & "" & "|" & "" & "|" & "" & "|")
Else
'   vRet = Shell(Trim(dir_trabajo) & "WsSapPortal.exe " & Trim(parametro1) & "|" & Trim(parametro2) & "|" & Trim(parametro3) & "|" & Trim(dir_trabajo) & "|" & vg_SqlNSvr & "|" & vg_SqlBase & "|" & vg_SqlNUsr & "|" & vg_SqlPass & "|")
   vRet = Shell(Trim(dir_trabajo) & "WsSapPortal.exe " & Trim(parametro1) & "|" & Trim(parametro2) & "|" & Trim(parametro3) & "|" & LCase(App.Path) & "\" & "|" & vg_SqlNSvr & "|" & vg_SqlBase & "|" & vg_SqlNUsr & "|" & vg_SqlPass & "|")
End If
If vRet = 0 Then Text1(0).text = Text1(0).text & "Proceso cancelado, no hay comunicación con Web Service": GenerarArcSap = False: Exit Function
DoEvents
RS1.Open "SELECT * FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND numero = " & numero & " AND tipo_proceso = '4' AND estado = '0'", vg_db, adOpenStatic
Do While Not RS1.EOF
   DoEvents
   RS1.Close: Set RS1 = Nothing
   RS1.Open "SELECT * FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND numero = " & numero & " AND tipo_proceso = '4' AND estado = '0'", vg_db, adOpenStatic
Loop
RS1.Close: Set RS1 = Nothing

'Dim xx As String
'Segundos = 0
'RS1.Open "SELECT * FROM log_procesos WHERE cencos='" & MuestraCasino(1) & "' AND numero=" & numero & " AND tipo_proceso='4'", vg_db, adOpenStatic
'If Not RS1.EOF Then
'   Segundos = Format(RS1!Fecha, "SS")
'   If Segundos + 10 > 60 Then
'      Segundos = (Segundos + 10) - 60
'   Else
'      Segundos = Segundos + 10
'   End If
'End If
'RS1.Close: Set RS1 = Nothing
'xx = "0"
'Do While xx <> "1"
'   If Format(Time, "SS") = fg_pone_cero(Segundos, 2) Then
'      RS1.Open "SELECT * FROM log_procesos WHERE cencos='" & MuestraCasino(1) & "' AND numero=" & numero & " AND tipo_proceso='4' AND estado<>'0'", vg_db, adOpenStatic
'      If Not RS1.EOF Then xx = "1"
'      RS1.Close: Set RS1 = Nothing
'   End If
'Loop
'-------> Proceso de estado de envio
RS1.Open "SELECT * FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND numero = " & numero & " AND tipo_proceso = '4'", vg_db, adOpenStatic
If Not RS1.EOF Then
   StrMensaje = Trim(RS1!mensaje)
   If Len(StrMensaje) <> 0 Then
      Text1(0).text = Text1(0).text & VgLinea
      Text1(0).text = Text1(0).text & FechaHora & IIf(RS1!estado = "3", "Mensaje Error : ", "Mensaje SAP : ") & VgLinea
      Text1(0).text = Text1(0).text & FechaHora & "---------------------------------------------------------" & VgLinea
      Do While InStr(StrMensaje, ";") <> 0 And InStr(StrMensaje, ";") <> 1
         If StrMensaje <> "" Then
            nommen = Mid(StrMensaje, 1, InStr(StrMensaje, "|") - 1)
            StrMensaje = Mid(StrMensaje, InStr(StrMensaje, "|") + 1)
            Text1(0).text = Text1(0).text & FechaHora & Trim(nommen) & VgLinea
            If InStr(nommen, "timed out") <> 0 Or InStr(nommen, "No esta conectado a la internet") <> 0 Then RS1.Close: Set RS1 = Nothing: GenerarArcSap = False: Exit Function
         End If
      Loop
      Text1(0).text = Text1(0).text & FechaHora & "---------------------------------------------------------" & VgLinea
      Text1(0).text = Text1(0).text & VgLinea
      If (RS1!estado = "2" Or RS1!estado = "0" Or RS1!estado = "3") Then RS1.Close: Set RS1 = Nothing: GenerarArcSap = False: Exit Function
   End If
End If
RS1.Close: Set RS1 = Nothing
'-------> Grabar tabla b_totcompras si se genero sin problema
vg_db.Execute "UPDATE b_totguiavta SET tgv_envsap = '1' WHERE tgv_rutcli = '" & fg_DespintaRut(LimpiaDato(Trim(fpText1(1).text))) & "' AND tgv_codsuc = '" & Trim(fpText1(0).text) & "' AND tgv_numdoc = " & Val(fpLongInteger1(0).Value) & " AND tgv_fecing = '" & CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")) & "'"
GenerarArcSap = True

Exit Function
Man_EnvioGuiaVenta:
If Err = 53 Then
   RS1.Close: Set RS1 = Nothing
   fpText1(0).text = fpText1(0).text & FechaHora & "No existe Ejecutable de envio..." & VgLinea
   vg_db.Execute "UPDATE log_procesos SET estado='4', mensaje='No existe ejecutable, para procesar Web Service' WHERE cencos='" & MuestraCasino(1) & "' AND tipo_proceso='4' AND numero=" & numero & ""
   GenerarArcSap = False
   Exit Function
End If
If Err = 3034 Then vg_db.RollbackTrans: Exit Function
RS1.Close: Set RS1 = Nothing
If Err.Number = -2147467259 Then MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error" Else MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
vg_db.RollbackTrans
End Function
