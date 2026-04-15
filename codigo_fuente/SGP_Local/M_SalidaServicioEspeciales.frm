VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_SalidaServicioEspeciales 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salida de Servicio Especiales"
   ClientHeight    =   7440
   ClientLeft      =   1995
   ClientTop       =   2280
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   2130
      Left            =   0
      TabIndex        =   10
      Top             =   435
      Width           =   12540
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   8595
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   200
         Width           =   2325
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   0
         Top             =   240
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
         Left            =   1800
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   580
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
         Left            =   1800
         TabIndex        =   3
         Top             =   945
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
         Left            =   1800
         TabIndex        =   5
         Top             =   1680
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
         Left            =   6105
         TabIndex        =   6
         Top             =   1680
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   4
         Top             =   1320
         Width           =   7995
         _Version        =   196608
         _ExtentX        =   14102
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3120
         Picture         =   "M_SalidaServicioEspeciales.frx":0000
         Top             =   170
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Precio Venta"
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
         Left            =   4800
         TabIndex        =   26
         Top             =   1770
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Comensales"
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
         TabIndex        =   25
         Top             =   1770
         Width           =   1020
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   9840
         Picture         =   "M_SalidaServicioEspeciales.frx":030A
         Top             =   1245
         Width           =   480
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
         Left            =   11250
         TabIndex        =   19
         Top             =   330
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
         Left            =   7755
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   350
         Width           =   735
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
         Left            =   75
         TabIndex        =   14
         Top             =   680
         Width           =   690
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
         Index           =   7
         Left            =   75
         TabIndex        =   9
         Top             =   1005
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Servicio Especiales"
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
         TabIndex        =   13
         Top             =   1380
         Width           =   1680
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3525
         TabIndex        =   11
         Top             =   255
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3570
         TabIndex        =   12
         Top             =   285
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   8655
         TabIndex        =   17
         Top             =   255
         Width           =   2310
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4350
      Left            =   540
      TabIndex        =   24
      Top             =   2445
      Width           =   11460
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3885
         Left            =   120
         TabIndex        =   7
         Top             =   315
         Width           =   11265
         _Version        =   393216
         _ExtentX        =   19870
         _ExtentY        =   6853
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
         MaxCols         =   8
         MaxRows         =   20
         ProcessTab      =   -1  'True
         SelectBlockOptions=   0
         SpreadDesigner  =   "M_SalidaServicioEspeciales.frx":0614
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   540
      TabIndex        =   20
      Top             =   6750
      Width           =   11460
      Begin VB.Frame Frame4 
         Height          =   450
         Left            =   9840
         TabIndex        =   28
         Top             =   120
         Width           =   1425
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   45
            TabIndex        =   29
            Top             =   135
            Width           =   1230
         End
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   1230
         TabIndex        =   8
         Top             =   180
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   688
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
         Left            =   0
         Top             =   0
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
               Picture         =   "M_SalidaServicioEspeciales.frx":0C06
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_SalidaServicioEspeciales.frx":0F20
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Total Grl."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   8760
         TabIndex        =   27
         Top             =   360
         Width           =   915
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
         TabIndex        =   23
         Top             =   165
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   7125
         Top             =   285
         Width           =   300
      End
      Begin VB.Label Label5 
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   7485
         TabIndex        =   22
         Top             =   255
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "Sobrepasa Stock actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   5850
         TabIndex        =   21
         Top             =   165
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008484FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   5445
         Top             =   285
         Width           =   300
      End
   End
End
Attribute VB_Name = "M_SalidaServicioEspeciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim est As Boolean
Dim est1 As Boolean
Dim modo As String
Dim MsgTitulo As String


Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.Height = 7875
Me.Width = 12720

fpDateTime1(0).DateTimeFormat = UserDefined
fpDateTime1(0).UserDefinedFormat = "dd/mm/yyyy"
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")

EspFecha fpDateTime1(0)

fg_centra Me
est = False
Me.HelpContextID = vg_OpcM
MsgTitulo = "Salida Ventas Servicios Especiales"

Dim X As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
Gl_Mo_Botones Me, 12

vaSpread1.Row = -1
vaSpread1.Col = 4: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
'vaSpread1.Col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = 2 'vg_DCa
vaSpread1.Col = 6: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
vaSpread1.Col = 8: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr

'-------> Cargar Combo Bodega
CargarDatoCombo Combo1, 0, "b_clientes", "cli_", "CliBod", "N"
Limpia 2

TraerFechaCierre

est = False
est1 = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Resize()

On Error GoTo Man_Error


Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpDateTime1_Change(Index As Integer)

On Error GoTo Man_Error

If est Then Exit Sub
If Trim(fpDateTime1(0).text) = "" Then Exit Sub
If Not IsDate(fpDateTime1(0).text) Then Exit Sub

Select Case Index

    Case 0
        
        vaSpread1.MaxRows = 0
'        If vg_tipser Then MostrarDetalle

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpDouble1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_LostFocus(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If Index = 1 Then Exit Sub

If fpText1(0).text = "" Then Exit Sub

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_TraerCecoSalVentaServiciosEspeciales '" & fpText1(0).text & "'")
If Not RS.EOF Then
    
    Do While Not RS.EOF
       
       fpayuda(Index).Caption = RS!cli_nombre
       Gl_Ac_Botones Me, 12, 2, ""
       fpText1(0).Enabled = False
       RS.MoveNext
    
    Loop

Else
   
   RS.Close
   Set RS = Nothing
   
   MsgBox "Contrato no existe...", vbExclamation + vbOKOnly, MsgTitulo
   Limpia 2
   If fpText1(0).Enabled = True Then fpText1(0).SetFocus
   Exit Sub

End If
RS.Close
Set RS = Nothing
fpLongInteger1(0).text = TraerCorrelativo(vg_codbod, "SE") 'MuestraFolio(Trim(fpText1(1).text))

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

vg_codigo = 0

Select Case Index

    Case 0 '-------> Contrato
        
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
    
    Case 1 '-------> Servicio especiales
        
        vg_left = fpayuda(0).Left + 2300
        vg_nombre = ""
        vg_codigo = ""
        vg_modrec = True
        B_TabEst.LlenaDatos "a_servicio", "ser_", "Venta Servicios Especiales", "VtaSerEsp"
        B_TabEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        fpText1(Index) = Trim(vg_nombre)
'        fpayuda(Index).Caption = vg_nombre

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim rutcli As String, tipdoc As String, NumDoc As Long, Fecha As Long, codbod  As Long, fecemi As Date, fecpro As String, codreg As Long, codser As Long, i As Long, canact As Double, aAp  As String, estdoc As String
Dim numlin As Long, codmer As String, coding As String, canmin As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String, total As Double, diablq As Date, color As String, codsec As String, totdec As Double
Dim sql1             As String
Dim ServicioEspecial As String
Dim Comensales       As Double
Dim PrecioServicio   As Double
Dim RS               As New ADODB.Recordset
Dim MyBuffer         As String
Dim NomArchivoExcel  As String

Dim xlApp    As Object
Dim xlWb     As Object
Dim xlWs     As Object
Dim XL       As New Excel.Application 'Crea el objeto excel

MsgTitulo = "Salida Servicios Especiales"

est1 = False
codbod = Val(fg_codigocbo(Combo1, 0, 10, ""))
fecpro = Format(fpDateTime1(0).Value, "dd/mm/yyyy")

If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then

   diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy")

Else

   diablq = 0

End If

TraerFechaCierre

Select Case Button.Index

    Case 1, 6 '-------> Nuevo-Cancelar
        
        If Button.Index = 6 And vaSpread1.MaxRows > 0 Then
        
           If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
           
              Exit Sub
              
           End If
           
        End If
        
        modo = IIf(Button.Index = 1, "A", "")
        
        Limpia IIf(Button.Index = 1, 6, 2)
        If fpText1(0).Enabled = True Then fpText1(0).SetFocus
        Frame2.Enabled = IIf(Button.Index = 1, True, False)
        'If vg_tipser Then vaSpread1.SetFocus
    
    Case 8, 15 '-------> Graba
        
        If Button.Index = 15 Then
        
           modo = "M"
           est1 = True
        
        End If
        
        If vaSpread1.MaxRows > 1 Then vaSpread1_EditMode vaSpread1.ActiveCol, vaSpread1.ActiveRow, 0, True
        
        If Trim(fpText1(0).text) = "" Or (Trim(fpayuda(0).Caption) = "") Then
        
           MsgBox "Debe ingresar contrato...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If Trim(fpLongInteger1(0).text) = "" Then
        
           MsgBox "Debe ingresar numero documento...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If Trim(fpDateTime1(0).text) = "" Then
        
           MsgBox "Debe ingresar fecha producción...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If Trim(fpDateTime1(0).text) = "" Then
        
           MsgBox "Debe ingresar fecha producción...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If Trim(Combo1(0).text) = "" Then
        
           MsgBox "Debe ingresar bodega...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If fpDouble1(0).Value < 1 Then
        
           MsgBox "Debe ingresar comensales...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If fpDouble1(1).Value < 1 Then
        
           MsgBox "Debe ingresar precio venta...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If Trim(fpText1(1).text) = "" Then
        
           MsgBox "Debe ingresar descripción venta especial...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If


'        If Trim(fpText1(0).text) = "" Or Trim(fpLongInteger1(0).text) = "" Or Trim(fpDateTime1(0).text) = "" _
'        Or Trim(Combo1(0).text) = "" Or (Trim(fpayuda(0).Caption) = "") Or (Trim(fpText1(0).text) = "" Or fpDouble1(0).Value < 1 Or fpDouble1(1).Value < 1) Then
'
'           MsgBox "Debe ingresar dato importante...", vbExclamation + vbOKOnly, MsgTitulo
'           Exit Sub
'
'        End If
        '-------> Validar si el contrato tiene asignado inventario rotativo
        If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
           Format(fpDateTime1(0).Value, "dd/mm/yyyy") > CDate(vg_ciedia) Then
           
           MsgBox "Tiene que realizar cierre diario...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then
        
           MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then
        
           MsgBox "Documento no corresponde al periodo : " & VgLinea & VgLinea & CierreFecha, vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then
        
           MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        'Validar inventario calendarizado 20201001
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 38) Then
            
           MsgBox "Se esta realizando la toma de inventario en estos momento...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
               
        End If
            
        
        'Validar ingreso documento inventario calendarizado 20201001
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 40) Then
            
           MsgBox "No puede ingresar documento, antes de un inventario calendarizado...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
               
        End If
        
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 8) Then
        
           MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Then
        
           MsgBox "Día se encuentra cerrado, no es posible ingresar...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        'validar que no exista un documento devolución
        rutcli = Trim(LimpiaDato(fpText1(0).text))
        ServicioEspecial = Trim(LimpiaDato(fpText1(1).text))
        tipdoc = "SE"
        NumDoc = fpLongInteger1(0).text
        codbod = Val(fg_codigocbo(Combo1, 0, 10, ""))
        fecpro = Format(fpDateTime1(0).text, "yyyymmdd")
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("sgp_Sel_ValidarDocumentoDEVentaServiciosEspeciales '" & rutcli & "', '" & fecpro & "', " & codbod & ", '" & ServicioEspecial & "', '" & tipdoc & "', " & NumDoc & "")
        
        If Not RS.EOF Then
        
              RS.Close
              Set RS = Nothing
              MsgBox "Documento tiene una devolucion realizada, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo
              Toolbar1.Enabled = True
              Exit Sub
        
        End If
        
        RS.Close
        Set RS = Nothing
           
        Toolbar1.Enabled = False
        If Button.Index = 15 Then
           
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient

           Set RS = vg_db.Execute("sgp_Sel_DocCerradoxUsuarioSalVentaServiciosEspeciales '" & Trim(LimpiaDato(fpText1(0).text)) & "', " & fpLongInteger1(0).Value & ", " & vg_codbod & "")
           
           If RS.EOF Then
              
              RS.Close
              Set RS = Nothing
              Gl_Ac_Botones Me, 12, 3, ""
              Toolbar1.Buttons(15).Enabled = False
              Label1.Caption = ""
              Frame1.Enabled = False
              Frame2.Enabled = False
              vaSpread1.Col = -1
              vaSpread1.Row = -1
              vaSpread1.Lock = True
              MsgBox "Documento fue cerrado por otro usuario, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo
              Toolbar1.Enabled = True
              Exit Sub
           
           End If
           RS.Close
           Set RS = Nothing
           
           For i = 1 To vaSpread1.MaxRows
               
               vaSpread1.Row = i
               vaSpread1.Col = 7
               
               If Left(vaSpread1.text, 1) = "S" Then
                  
                  est1 = False
                  MsgBox "Existe una cantidad que exceden el Stock...", vbExclamation + vbOKOnly, MsgTitulo
                  Toolbar1.Enabled = True
                  Exit Sub
                  
              End If
           
               vaSpread1.Col = 4
               
               If vaSpread1.text <= 0 Then
                  
                  est1 = False
                  MsgBox "Existe una cantidad con valor cero...", vbExclamation + vbOKOnly, MsgTitulo
                  Toolbar1.Enabled = True
                  Exit Sub
                  
              End If
           
           Next i
           
        End If
        
        total = 0
        
        For i = 1 To vaSpread1.MaxRows

            vaSpread1.Row = i
            vaSpread1.Col = 6
            ptotal = Format(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), fg_Pict(9, vg_DPr))
            total = total + ptotal

        Next i
        
        If total = 0 Then
           
           MsgBox "El total del documento debe ser mayor a cero...", vbExclamation + vbOKOnly, MsgTitulo
           Toolbar1.Enabled = True
           Exit Sub
        
        End If
        
        '-------> validar si graba documentos con rebaja de bodega
        If Button.Index = 15 Then
        
           If MsgBox("Esta Seguro Cerrar Salida Venta Servicio Especial...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
              
              est1 = False
              Toolbar1.Enabled = True
              Exit Sub
              
          End If
          
       End If
paso:
        
        rutcli = Trim(LimpiaDato(fpText1(0).text))
        tipdoc = "SE"
        NumDoc = fpLongInteger1(0).text
        codbod = Val(fg_codigocbo(Combo1, 0, 10, ""))
        fecpro = Format(fpDateTime1(0).text, "yyyymmdd")
        ServicioEspecial = Trim(LimpiaDato(fpText1(1).text))
        Comensales = fpDouble1(0).text
        PrecioServicio = fpDouble1(1).text
        
        est1 = False
        
        '-------> Detalle
        Let MyBuffer = ""
        Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        Let MyBuffer = MyBuffer & "<GrabaVenta>"
    
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            
            vaSpread1.Col = 1
            codmer = Trim(LimpiaDato(vaSpread1.text))
            
            codmer = Replace(Trim(codmer), Chr(34), "&quot;")
            codmer = Replace(Trim(codmer), Chr(38), "&amp;")
            codmer = Replace(Trim(codmer), Chr(39), "&apos;")
            codmer = Replace(Trim(codmer), Chr(60), "&lt;")
            codmer = Replace(Trim(codmer), Chr(62), "&gt;")
              
            vaSpread1.Col = 4
            canmer = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DCa)
            
            vaSpread1.Col = 5
            predoc = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), 2) 'vg_DPr)
                                     
            MyBuffer = MyBuffer & " <Venta"
            MyBuffer = MyBuffer & " NLin = " & Chr(34) & i & Chr(34)
            MyBuffer = MyBuffer & " IdProd = " & Chr(34) & codmer & Chr(34)
            MyBuffer = MyBuffer & " CanMer = " & Chr(34) & canmer & Chr(34)
            MyBuffer = MyBuffer & " Precio = " & Chr(34) & predoc & Chr(34)
            MyBuffer = MyBuffer & "/>"
        
        Next i
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        MyBuffer = MyBuffer & "</GrabaVenta>"
        Set RS = vg_db.Execute("sgp_Ins_XmlSalidaVentaServiciosEspeciales '" & MyBuffer & "', '" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", '" & fecpro & "', " & codbod & ", '" & ServicioEspecial & "', " & Comensales & ", " & PrecioServicio & ", '" & vg_NUsr & "', '" & modo & "', '" & IIf(Button.Index = 15, "", "P") & "'")
        If Not RS.EOF Then
           
           If RS(0) > 0 Then
              
           
              MsgBox RS(1) & VgLinea & VgLinea & "            Adjunto archivo con error", vbCritical, MsgTitulo

              '-------> Create an instance of Excel and add a workbook
              Set xlApp = CreateObject("Excel.Application")
              Set xlWb = xlApp.Workbooks.Add
              Set xlWs = xlWb.Worksheets("Hoja1")
        
              '-------> Display Excel and give user control of Excel's lifetime
              xlApp.UserControl = True
        
              '-------> Check version of Excel
              Call encabezado(RS, xlWs)
              xlWs.Cells(2, 1).CopyFromRecordset RS
              '-------> Auto-fit the column widths and row heights
              xlApp.Selection.CurrentRegion.Columns.AutoFit
              xlApp.Selection.CurrentRegion.Rows.AutoFit
'              xlApp.Columns("A:A").Select
'              xlApp.Selection.Delete Shift:=xlToLeft
              
              NomArchivoExcel = fg_ArchivoXls("ReporteError_VentasServiciosEspeciales")
              
              xlWb.Close True, NomArchivoExcel
              XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
              XL.Visible = True
              XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
              
              '-- Cerrar Excel
              xlApp.Quit
        
              '-------> Release Excel references
              Set xlWs = Nothing
              Set xlWb = Nothing
              Set xlApp = Nothing
           
              RS.Close
              Set RS = Nothing
              Toolbar1.Enabled = True
              Exit Sub
              
           End If
        
           fpLongInteger1(0).text = RS(3)
        
        End If
        RS.Close
        Set RS = Nothing
       
        Gl_Ac_Botones Me, 12, 3, ""
        Toolbar1.Buttons(15).Enabled = IIf(Label1.Caption = "ANULADA", False, True)
        Toolbar1.Buttons(15).ToolTipText = IIf(Label1.Caption = "ANULADA", "", "Cerrar Salida Venta Servicios Especiales")
        
        Label1.Caption = IIf(Button.Index = 15, "", "PENDIENTE")
        Frame1.Enabled = False
        Frame2.Enabled = True
        
        If Button.Index = 15 Then
        
           Frame1.Enabled = False
           Frame2.Enabled = False
           vaSpread1.Col = -1
           vaSpread1.Row = -1
           vaSpread1.Lock = True
           
        End If
        
        '-------> Revisa Stock
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 7
            color = Right(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), 1)
                         
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
        
            vaSpread1.Col = 1
            Set RS = vg_db.Execute("SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                     "WHERE  bod.bod_codpro = pro.pro_codigo " & _
                     "AND    pro.pro_ctrsto = 1 " & _
                     "AND    bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 0, 10, "")) & " " & _
                     "AND    pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "'")
               
            vaSpread1.Col = 8
            If Not RS.EOF Then
               
               vaSpread1.text = Format(RS!bod_canmer, fg_Pict(9, vg_DCa))
               
            Else
               
               vaSpread1.text = 0
               
            End If
               
            RS.Close
            Set RS = Nothing
                   
        Next i
        
        If Button.Index = 15 Then
        
           I_SalDevVentaServiciosEspeciales Me, "SE"
           Toolbar1.Buttons(15).Enabled = False
        
        End If
        
        Toolbar1.Enabled = True
        modo = "M"
        
    Case 3 '-------> Anular
        
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then
        
           MsgBox "Periodo esta cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then
        
           MsgBox "No puede anular documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Then
        
           MsgBox "No puede anular documento, día esta cerrado...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        'Validar inventario calendarizado 20201001
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 38) Then
            
           MsgBox "Se esta realizando la toma de inventario en estos momento...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
               
        End If
              
        'validar que no exista un documento devolución ventas servicios especiales
        rutcli = Trim(LimpiaDato(fpText1(0).text))
        ServicioEspecial = Trim(LimpiaDato(fpText1(1).text))
        tipdoc = "SE"
        NumDoc = fpLongInteger1(0).text
        codbod = Val(fg_codigocbo(Combo1, 0, 10, ""))
        fecpro = Format(fpDateTime1(0).text, "yyyymmdd")
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        Set RS = vg_db.Execute("sgp_Sel_ValidarDocumentoDEVentaServiciosEspeciales '" & rutcli & "', '" & fecpro & "', " & codbod & ", '" & ServicioEspecial & "', '" & tipdoc & "', " & NumDoc & "")
        
        If Not RS.EOF Then
        
              RS.Close
              Set RS = Nothing
              MsgBox "Documento tiene una devolucion realizada, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo
              Toolbar1.Enabled = True
              Exit Sub
        
        End If
        
        RS.Close
        Set RS = Nothing
        
        If MsgBox("Anula documento...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        
        codbod = Val(fg_codigocbo(Combo1, 0, 10, ""))
        estdoc = ""
        totdec = 0
              
        Let MyBuffer = ""
        Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        Let MyBuffer = MyBuffer & "<GrabaVenta>"
           
        For i = 1 To vaSpread1.MaxRows
               
            vaSpread1.Row = i
            
            vaSpread1.Col = 1
            codmer = Trim(LimpiaDato(vaSpread1.text))
            
            codmer = Replace(Trim(codmer), Chr(34), "&quot;")
            codmer = Replace(Trim(codmer), Chr(38), "&amp;")
            codmer = Replace(Trim(codmer), Chr(39), "&apos;")
            codmer = Replace(Trim(codmer), Chr(60), "&lt;")
            codmer = Replace(Trim(codmer), Chr(62), "&gt;")
              
            vaSpread1.Col = 4
            canmer = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DCa)
            
            vaSpread1.Col = 5
            predoc = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), 2) 'vg_DPr)
                                     
            MyBuffer = MyBuffer & " <Venta"
            MyBuffer = MyBuffer & " NLin = " & Chr(34) & i & Chr(34)
            MyBuffer = MyBuffer & " IdProd = " & Chr(34) & codmer & Chr(34)
            MyBuffer = MyBuffer & " CanMer = " & Chr(34) & canmer & Chr(34)
            MyBuffer = MyBuffer & " Precio = " & Chr(34) & predoc & Chr(34)
            MyBuffer = MyBuffer & "/>"
                          
        Next i
               
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        MyBuffer = MyBuffer & "</GrabaVenta>"
        Set RS = vg_db.Execute("sgp_Upd_XmlSalidaVentaServiciosEspeciales '" & MyBuffer & "', '" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", '" & fecpro & "', " & codbod & ", '" & vg_NUsr & "'")
        If Not RS.EOF Then
           
           If RS(0) > 0 Then
              
              MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
           
              RS.Close
              Set RS = Nothing
           
              Exit Sub
              
           End If
        
        End If
        RS.Close
        Set RS = Nothing
        
        '-------> Revisa Stock
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 7
            color = Right(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), 1)
            
            vaSpread1.Col = 1
               
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
               
            Set RS = vg_db.Execute("SELECT bod.bod_canmer from b_productos AS pro, b_bodegas AS bod " & _
                    "WHERE bod.bod_codpro = pro.pro_codigo " & _
                    "AND   pro.pro_ctrsto = 1 " & _
                    "AND   bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 0, 10, "")) & " " & _
                    "AND   pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "'")
               
            vaSpread1.Col = 8
               
            If Not RS.EOF Then
               
               vaSpread1.text = Format(RS!bod_canmer, fg_Pict(9, vg_DCa))
            
            Else
               
               vaSpread1.text = 0
               
            End If
            
            RS.Close
            Set RS = Nothing
        
        Next i
        Label1.Caption = "ANULADA"
        Gl_Ac_Botones Me, 12, IIf(Label1.Caption = "ANULADA", 4, 3), ""
        modo = ""
    
    Case 11 '-------> Busqueda
        
        If Trim(fpText1(0).text) = "" Then MsgBox "Debe seleccionar contrato...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        vg_codigo = Trim(fpText1(0).text)
        vg_nombre = "SE"
        B_SalBod.Show 1
        If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
        Me.MousePointer = 11
        Me.Refresh
        
        Frame2.Enabled = False
        Frame1.Enabled = True
        
        vaSpread1.Col = -1
        vaSpread1.Row = -1
        vaSpread1.Lock = True
        vaSpread1.MaxRows = 0
        
        est = True
        
        est = False
           
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("sgp_Sel_DetalleSalVentaServiciosEspeciales '" & LimpiaDato(Trim(fpText1(0).text)) & "', 'SE', " & vg_codbod & ", " & Val(vg_codigo) & "")
       
        If Not RS.EOF Then
            
           est = True
                
           fpLongInteger1(0).text = RS!numerodocumento
           fpDateTime1(0).text = RS!fechaproduccion
           Combo1(0).ListIndex = fg_buscacbo(Combo1, 0, 10, fg_pone_cero(Str(RS!Bodega), 10))
           Label1.Caption = IIf(RS!estadodocumento = "", "", IIf(RS!estadodocumento = "A", "ANULADA", "PENDIENTE"))
           fpText1(1).text = RS!vtaServicioespeciales
           fpDouble1(0).Value = RS!Comensales
           fpDouble1(1).Value = RS!PrecioServicio
           
           If RS!estadodocumento = "P" Then
           
              modo = "M"
              Frame2.Enabled = True

           Else
           
              Frame2.Enabled = False
           
           End If
           
           est = False
            
            Do While Not RS.EOF
                
                
               vaSpread1.MaxRows = vaSpread1.MaxRows + 1
               vaSpread1.Row = vaSpread1.MaxRows
               
               vaSpread1.Col = 1
               vaSpread1.text = Trim(RS!CodigoProducto)
               
               vaSpread1.Col = 2
               vaSpread1.text = Trim(RS!NombreProducto)
               
               vaSpread1.Col = 3
               vaSpread1.text = Trim(RS!UnidadProducto)
               
               vaSpread1.Col = 4
               vaSpread1.ForeColor = &HFF0000
               vaSpread1.Lock = IIf(RS!estadodocumento = "P", False, True)
               vaSpread1.text = Format(RS!CantidadMercaderia, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa))) 'vg_DCa))
               
               vaSpread1.Col = 5
               vaSpread1.text = Format(RS!Preciodocumento, fg_Pict(9, 2)) 'vg_DCa))
               
               vaSpread1.Col = 6
               vaSpread1.text = Format(RS!TotalDocumento, fg_Pict(9, vg_DPr))
                             
               vaSpread1.Col = 7
               vaSpread1.text = "N"
               
               vaSpread1.Col = 8
               vaSpread1.text = Format(RS!cantidadbodega, fg_Pict(9, vg_DCa))
                
               If RS!estadodocumento = "P" And (RS!cantidadbodega - RS!CantidadMercaderia) < 0 Then
               
                  vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
                  
                  vaSpread1.Col = 7
                  vaSpread1.text = "S"
                  
               End If
               
               RS.MoveNext
                
            Loop
        
        End If
        RS.Close
        Set RS = Nothing
              
        '------- Total General ---------
        Dim Cantidad As Double
        Dim Precio As Double
        Dim subtot As Double
        
        subtot = 0
        
        For i = 1 To vaSpread1.MaxRows
                    
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            Cantidad = 0
            Precio = 0
            
            If Trim(vaSpread1.text) <> "" Then
                    
               vaSpread1.Col = 4
               Cantidad = IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0)
               
               vaSpread1.Col = 5
               Precio = IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0)
               
               vaSpread1.Col = 6
               subtot = subtot + Format((Cantidad * Precio), fg_Pict(9, vg_DPr))
                    
            End If
                    
        Next
        Label2.Caption = Format(subtot, fg_Pict(9, vg_DPr))
        
        Me.MousePointer = 0
        vaSpread1.Visible = True
        vg_codigo = ""
'        Frame5.Enabled = True
        Gl_Ac_Botones Me, 12, IIf(Label1.Caption = "ANULADA", 4, 3), ""
        Toolbar1.Buttons(15).Enabled = IIf(Label1.Caption = "ANULADA" Or Label1.Caption = "", False, True)
        Toolbar1.Buttons(15).ToolTipText = IIf(Label1.Caption = "ANULADA" Or Label1.Caption = "", "", "Cerrar Salida Venta Servicios Especiales")
    
        fpDateTime1(0).Enabled = False
        fpText1(1).Enabled = False
        Image1(1).Enabled = False
        
        fpDouble1(0).Enabled = IIf(Label1.Caption = "ANULADA" Or Label1.Caption = "", False, True)
        fpDouble1(1).Enabled = IIf(Label1.Caption = "ANULADA" Or Label1.Caption = "", False, True)
   
    Case 12 '-------> Imprimir
        
        If vaSpread1.MaxRows < 1 Then Exit Sub
        I_SalDevVentaServiciosEspeciales Me, "SE"

    Case 17 '-------> Salir
        
        Me.Hide
        Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Function MuestraFolio(Casino As String) As String

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
        
        
Dim sql1 As String
MuestraFolio = ""
If Trim(Casino) = "" Then Exit Function
sql1 = IIf(vg_tipbase = "1", " HOLDLOCK ", " WITH (HOLDLOCK) ")

Set RS = vg_db.Execute("SELECT tos_numero_documento FROM b_totventaserviciosespciales " & sql1 & " WHERE tos_tipo_documento = 'SE' AND tos_IdBodega = " & vg_codbod & " ORDER BY tos_numero_documento DESC")
If Not RS.EOF Then
   
   RS.MoveFirst
   MuestraFolio = RS!tos_numero_documento + 1

Else

   MuestraFolio = 1
   
End If

RS.Close
Set RS = Nothing

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Function

Sub Limpia(op As Integer)

On Error GoTo Man_Error

est = True
Label1.Caption = ""
Label2.Caption = 0
Frame1.Enabled = True
Frame2.Enabled = False

fpLongInteger1(0).text = TraerCorrelativo(vg_codbod, "SE")
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(0).Enabled = True

fpText1(0).Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText1(0).text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)

fpText1(1).text = ""
fpText1(1).Enabled = True
Image1(1).Enabled = True

fpDouble1(0) = Format(0, fg_Pict(0, vg_DCa))
fpDouble1(0).Enabled = True

fpDouble1(1) = Format(0, fg_Pict(0, vg_DCa))
fpDouble1(1).Enabled = True

Combo1(0).ListIndex = IIf(Combo1(0).listcount = 1, 0, -1)

vaSpread1.MaxRows = 0
vaSpread1.Col = -1
vaSpread1.Row = -1
vaSpread1.Lock = True

vaSpread1.Col = 4
vaSpread1.Row = -1
vaSpread1.Lock = False

Gl_Ac_Botones Me, 12, op, ""
est = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim i As Long, indsec As Long, codpro As String, color As String, coding As String, texto As String, auxpro As String, auxing As String, codsec As String
Dim Cantidad As Double
Dim Precio As Double
Dim propon As Double
Dim subtot As Double
Dim RS As New ADODB.Recordset

Select Case Button.Index

Case 1
    
    If Trim(Combo1(0).text) = "" Then MsgBox "Debe seleccionar bodega...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Toolbar2.Enabled = False
    
    vg_nombre = ""
    vg_codigo = ""
    vg_bodega = 0
    vg_bodega = Val(fg_codigocbo(Combo1, 0, 10, ""))
    vg_left = fpayuda(0).Left + 1920
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 2
    
    texto = ""
    
    If vaSpread1.MaxRows > 0 Then
    
       texto = Trim(Mid(Trim(vaSpread1.text), 1, IIf(InStr(1, vaSpread1.text, " ") > 0, InStr(1, vaSpread1.text, " "), Len(Trim(vaSpread1.text)))))
           
    End If
    
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Pbo"
    B_TabEst.Text1.text = "" 'texto
    B_TabEst.Show 1
    If vg_codigo = "" Then Toolbar2.Enabled = True: Exit Sub
    Toolbar2.Enabled = True
    auxpro = vg_codigo
    indsec = 0
       
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        codpro = ""
        color = ""
        vaSpread1.Col = 1
        codpro = Trim(vaSpread1.text)
        
        If Trim(codpro) = Trim(vg_codigo) Then
                        
           MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
                  
        End If
           
    Next i
               
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
       
    Set RS = vg_db.Execute("sgp_Sel_ProdUnidadSalventaServiciosEspeciales '" & vg_codigo & "'")
    
    
    If Not RS.EOF Then
       
       vaSpread1.Col = -1
       vaSpread1.BackColor = Shape1(0).FillColor
       
       vaSpread1.Col = 1
       vaSpread1.text = RS!pro_codigo
       
       vaSpread1.Col = 2
       vaSpread1.text = RS!pro_nombre
       
       vaSpread1.Col = 3
       vaSpread1.text = RS!uni_nomcor
       
       vaSpread1.Col = 4
       vaSpread1.ForeColor = &HFF0000
       vaSpread1.Lock = False
       vaSpread1.text = Format(0, fg_Pict(9, vg_DCa))
       
       vaSpread1.Col = 5
       vaSpread1.text = Format(0, fg_Pict(9, vg_DCa))
    
    End If
    
    RS.Close
    Set RS = Nothing
    
    vaSpread1.Row = vaSpread1.MaxRows
    
    '-------> Traer propon
    propon = 0
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
       
    Set RS = vg_db.Execute("sgp_Sel_PrecioProdSalventaServiciosEspeciales '" & MuestraCasino(1) & "', '" & vg_codigo & "', " & Format(CDate(fpDateTime1(0).text), "yyyymmdd") & ", " & Format(CDate(vg_ciedia), "yyyymmdd") & "")
    If Not RS.EOF Then propon = RS!ppd_propon
    RS.Close: Set RS = Nothing
            
    vaSpread1.Row = vaSpread1.MaxRows
    
    vaSpread1.Col = 5
    vaSpread1.text = Format(propon, fg_Pict(9, 2)) 'vg_DPr))
    
    vaSpread1.Col = 6
    vaSpread1.text = Format(0, fg_Pict(9, vg_DPr))
            
    '-------> Trae Stock
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgp_Sel_SaldoBodegaSalventaServiciosEspeciales " & Val(fg_codigocbo(Combo1, 0, 10, "")) & ", '" & Trim(vg_codigo) & "'")
    
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = 8
    If Not RS.EOF Then
    
       vaSpread1.text = Format(RS!bod_canmer, fg_Pict(9, vg_DCa))
       
    Else
    
       vaSpread1.text = 0
       
    End If
    RS.Close
    Set RS = Nothing
    
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = 4
    vaSpread1.SetActiveCell 4, vaSpread1.Row
    vaSpread1.SetFocus

Case 2
    
    If vaSpread1.MaxRows = 0 Then Exit Sub
        
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1

    '------- Total General ---------
    subtot = 0
        
    For i = 1 To vaSpread1.MaxRows
                    
        vaSpread1.Row = i
        vaSpread1.Col = 1
        
        Cantidad = 0
        Precio = 0
        
        If Trim(vaSpread1.text) <> "" Then
                    
           vaSpread1.Col = 4
           Cantidad = IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0)
               
           vaSpread1.Col = 5
           Precio = IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0)
               
           vaSpread1.Col = 6
           subtot = subtot + Format((Cantidad * Precio), fg_Pict(9, vg_DPr))
                    
        End If
                    
    Next
    Label2.Caption = Format(subtot, fg_Pict(9, vg_DPr))

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim i As Long
Dim subtot As Double
Dim Cantidad As Double
Dim Precio As Double

Select Case Col

    Case 4
    
        '------- Total General ---------
        subtot = 0
        
        For i = 1 To vaSpread1.MaxRows
                    
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            Cantidad = 0
            Precio = 0
            
            If Trim(vaSpread1.text) <> "" Then
                    
               vaSpread1.Col = 4
               Cantidad = IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0)
               
               vaSpread1.Col = 5
               Precio = IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0)
               
               vaSpread1.Col = 6
               subtot = subtot + Format((Cantidad * Precio), fg_Pict(9, vg_DPr))
                    
            End If
                    
        Next
        Label2.Caption = Format(subtot, fg_Pict(9, vg_DPr))

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

'On Error GoTo Man_Error

Dim canrea As Double, canbod As Double, canaux As Double, propon As Double, codmer As String, color As Variant, i As Long
Dim color2 As String, codsec As String, auxsec As String, totcos As Double
codsec = "0"
auxsec = "0"

If ChangeMade = False Then Exit Sub
If Not est1 Then Gl_Ac_Botones Me, 12, 6, ""

vaSpread1.Row = Row
vaSpread1.Col = 1: codmer = vaSpread1.text
vaSpread1.Col = 4: canrea = Format(IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
vaSpread1.Col = 5: propon = Format(IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
vaSpread1.Col = 6: vaSpread1.text = Format(canrea * propon, fg_Pict(9, vg_DCa))
vaSpread1.Col = 8: canbod = Format(IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))

canaux = 0
    
For i = 1 To vaSpread1.MaxRows
        
    vaSpread1.Row = i
    vaSpread1.Col = 1
    
    If codmer = vaSpread1.text Then
       
       vaSpread1.Col = 4
       canaux = canaux + Format(IIf(vaSpread1.text = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
       
    End If
    
Next i
    
canrea = IIf(canaux > 0, canaux, canrea)
vaSpread1.Row = Row
    
If (canbod - canrea) >= 0 Then
        
   For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If codmer = vaSpread1.text Then
               
               vaSpread1.Col = -1
               vaSpread1.BackColor = Shape1(2).FillColor
               vaSpread1.Col = 7: vaSpread1.text = "" '-------> No Bloqueado

            
            End If
   
   Next i
   Exit Sub

End If
    
For i = 1 To vaSpread1.MaxRows
        
    vaSpread1.Row = i
    
    vaSpread1.Col = 4
    canrea = Format(IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text), fg_Pict(9, vg_DCa))
    
    vaSpread1.Col = 1
        
    If codmer = vaSpread1.text Then
           
       vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
       vaSpread1.Col = 7: vaSpread1.text = "S" '-------> Bloqueado
   
   End If

Next i

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

Dim i As Long, color As String
If KeyAscii <> 13 Then Exit Sub

For i = vaSpread1.ActiveRow + 1 To vaSpread1.MaxRows
    
    vaSpread1.Row = i: vaSpread1.Col = 8: color = Right(vaSpread1.text, 1)
    If color <> "I" Then vaSpread1.SetActiveCell vaSpread1.ActiveCol, i - 1: Exit Sub

Next i

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

On Error GoTo Man_Error

If Row = 0 Then Exit Sub

Dim Stock As String, Nombre As String, color As String
vaSpread1.Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
vaSpread1.Col = 8: Stock = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
vaSpread1.Col = 2: Nombre = vaSpread1.text
TipText = "Bodega   : " & Trim(Left(Combo1(0).text, 50)) & vbCrLf & _
          "Producto : " & Trim(Nombre) & vbCrLf & _
          "Stock       : " & Format(Trim(Stock), fg_Pict(9, vg_DCa))

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

