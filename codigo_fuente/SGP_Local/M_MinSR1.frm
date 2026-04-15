VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_MinSR1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minuta Bloque"
   ClientHeight    =   7845
   ClientLeft      =   2670
   ClientTop       =   2310
   ClientWidth     =   13485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7845
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2775
      Index           =   1
      Left            =   720
      TabIndex        =   19
      Top             =   120
      Width           =   11655
      Begin VB.OptionButton optPrecioLista 
         Caption         =   "Precio Lista"
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
         Left            =   8280
         TabIndex        =   36
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton optPrecioConvenio 
         Caption         =   "Precio Convenio"
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
         Left            =   1920
         TabIndex        =   5
         Top             =   2160
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton optPrecioGenerico 
         Caption         =   "Precio Sitio"
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
         Left            =   5400
         TabIndex        =   6
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1980
         TabIndex        =   1
         Top             =   735
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1676
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
         Left            =   1980
         TabIndex        =   2
         Top             =   1155
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1676
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
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   2
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1980
         TabIndex        =   0
         Top             =   315
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
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   1980
         TabIndex        =   3
         Top             =   1575
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         ButtonStyle     =   2
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
         Text            =   "01/09/2013"
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
      Begin EditLib.fpDateTime FpFecHasta 
         Height          =   315
         Left            =   9165
         TabIndex        =   4
         Top             =   1575
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         ButtonStyle     =   2
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
         Text            =   "28/09/2013"
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   2
         Left            =   1635
         TabIndex        =   34
         Top             =   2280
         Visible         =   0   'False
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1676
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
         Caption         =   "Bloque Minuta"
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
         Left            =   360
         TabIndex        =   35
         Top             =   2385
         Visible         =   0   'False
         Width           =   1230
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
         Index           =   0
         Left            =   705
         TabIndex        =   27
         Top             =   420
         Width           =   735
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
         Left            =   705
         TabIndex        =   26
         Top             =   840
         Width           =   750
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
         Index           =   3
         Left            =   705
         TabIndex        =   25
         Top             =   1260
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha desde"
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
         Left            =   705
         TabIndex        =   24
         Top             =   1665
         Width           =   1110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   3720
         TabIndex        =   23
         Top             =   315
         Width           =   6735
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3720
         TabIndex        =   22
         Top             =   735
         Width           =   6735
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3720
         TabIndex        =   21
         Top             =   1155
         Width           =   6735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha hasta"
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
         Left            =   7860
         TabIndex        =   20
         Top             =   1665
         Width           =   1065
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3270
         Picture         =   "M_MinSR1.frx":0000
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3270
         Picture         =   "M_MinSR1.frx":030A
         Top             =   660
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   3270
         Picture         =   "M_MinSR1.frx":0614
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
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
         Left            =   3765
         TabIndex        =   28
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
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
         Index           =   1
         Left            =   3765
         TabIndex        =   29
         Top             =   780
         Width           =   6735
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
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
         Index           =   2
         Left            =   3765
         TabIndex        =   30
         Top             =   1200
         Width           =   6735
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3615
      Left            =   285
      TabIndex        =   7
      Top             =   3240
      Width           =   12255
      _Version        =   393216
      _ExtentX        =   21616
      _ExtentY        =   6376
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
      MaxCols         =   10
      SpreadDesigner  =   "M_MinSR1.frx":091E
   End
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   12615
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   1560
         TabIndex        =   18
         Top             =   3840
         Width           =   1035
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   8
            Top             =   135
            Width           =   930
         End
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   2640
         TabIndex        =   17
         Top             =   3840
         Width           =   3525
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   9
            Top             =   135
            Width           =   3420
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   6240
         TabIndex        =   16
         Top             =   3840
         Width           =   3525
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   10
            Top             =   135
            Width           =   3420
         End
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Left            =   10080
         TabIndex        =   15
         Top             =   3840
         Width           =   1035
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   11
            Top             =   135
            Width           =   930
         End
      End
      Begin VB.Frame Frame7 
         Height          =   435
         Left            =   11160
         TabIndex        =   14
         Top             =   3840
         Width           =   1035
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   12
            Top             =   135
            Width           =   930
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celda Bloqueada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   9240
         TabIndex        =   33
         Top             =   4440
         Width           =   1395
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   8880
         Top             =   4470
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celda Habilitada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   11055
         TabIndex        =   32
         Top             =   4440
         Width           =   1365
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   10695
         Top             =   4470
         Width           =   300
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7845
      Left            =   12855
      TabIndex        =   31
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   13838
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_MinSR1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private NomFor As String
Private BtnX   As Variant
Public lc_Aux  As String
Dim modo       As String
Dim est        As Boolean

Private Sub Form_Activate()
    Call fg_descarga
End Sub

Private Sub Form_Load()
    Call fg_carga("")
    Me.HelpContextID = vg_OpcM
    Call fg_centra(Me)
    Let Me.Height = 8205
    Let Me.Width = 13575
    modo = ""
    Gl_Mo_Botones Me, 19
    Gl_Ac_Botones Me, 1, 16, modo
    est = True
    Call FormatearDatos
    Call fg_descarga
    fpText.Enabled = ModCasino
    Image1(0).Enabled = ModCasino
    fpText.text = MuestraCasino(1)
    fpayuda(0).Caption = MuestraCasino(2)
    optPrecioConvenio.Value = True
    est = False
End Sub

Private Sub FormatearDatos()
    Let FpFecDesde.text = Format(Date, "dd/mm/yyyy")
    Let FpFecHasta.text = Format(Date, "dd/mm/yyyy")
    Let vaSpread1.MaxRows = 0
    Let fpText.text = ""
    Let fpLongInteger1(0).Value = ""
    Let fpLongInteger1(1).Value = ""
    Let Text1(0).text = ""
    Let Text1(1).text = ""
    Let Text1(2).text = ""
    Let Text1(3).text = ""
    Let Text1(4).text = ""
End Sub

Private Sub MoverGrilla(cencos As String)
Dim RS As New ADODB.Recordset

'If Lc_Aux = "PlaTeo" Then
'   Set RS = vg_db.Execute("sgp_Sel_ListarMinutaBloquexCecoTeoReal '" & cencos & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & "")
'Else
   Set RS = vg_db.Execute("sgp_Sel_ListarMinutaBloquexCeco '" & cencos & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & "")
'End If
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(0).FillColor  'Amarillo
Do While Not RS.EOF = True
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   If RS!IdEstadoMinuta <> 11 Then
      vaSpread1.Col = -1
      vaSpread1.BackColor = Shape1(1).FillColor ' Rojo
   End If
   vaSpread1.Col = 2
   vaSpread1.text = CStr(RS!ID_Bloque)
   vaSpread1.Col = 3
   vaSpread1.text = RS!reg_codigo & " - " & Trim(RS!reg_nombre)
   vaSpread1.Col = 4
   vaSpread1.text = RS!ser_codigo & " - " & Trim(RS!ser_nombre)
   vaSpread1.Col = 5
   vaSpread1.text = Format(RS!FechaDesde, "dd/mm/yyyy")
   vaSpread1.Col = 6
   vaSpread1.text = Format(RS!FechaHasta, "dd/mm/yyyy")
   vaSpread1.Col = 7
   vaSpread1.text = RS!reg_codigo
   vaSpread1.Col = 8
   vaSpread1.text = RS!ser_codigo
   vaSpread1.Col = 9
   vaSpread1.text = Trim(RS!reg_nombre)
   vaSpread1.Col = 10
   vaSpread1.text = Trim(RS!ser_nombre)
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
        FpFecDesde.Enabled = True
        FpFecHasta.Enabled = True

End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
End Sub

Private Function ValidaDatos() As Boolean
Dim Dias    As Long
Dim i       As Long
Dim Fecha   As String
Dim mes     As String
Dim Año     As String

    Let ValidaDatos = True
     
    If Len(fpText.text) = 0 Then
        Call MsgBox("Debe Ingresar Centro De Costo", vbInformation, Me.Caption)
        Let ValidaDatos = False
        Exit Function
    End If
    
    If Len(fpLongInteger1(0).text) = 0 Then
        Call MsgBox("Debe Ingresar Regimen", vbInformation, Me.Caption)
        Call fpLongInteger1(0).SetFocus
        Let ValidaDatos = False
        Exit Function
    End If
    
    If Len(fpLongInteger1(1).text) = 0 Then
        Call MsgBox("Debe Ingresar Servicio", vbInformation, Me.Caption)
        Call fpLongInteger1(1).SetFocus
        Let ValidaDatos = False
        Exit Function
    End If
    
    If CDate(FpFecDesde.text) > CDate(FpFecHasta.text) Then
        Call MsgBox("Fecha Desde No Puede Ser Mayor a Fecha Hasta", vbInformation, Me.Caption)
        Let FpFecDesde.text = Format(Now, "dd/mm/yyyy")
        Call FpFecDesde.SetFocus
        Let ValidaDatos = False
        Exit Function
    End If
    
    If CDate(FpFecHasta.text) < CDate(FpFecDesde.text) Then
        Call MsgBox("Fecha Hasta No Puede Ser Mayor a Fecha Desde", vbInformation, Me.Caption)
        Let FpFecHasta.text = Format(Now, "dd/mm/yyyy")
        Call FpFecHasta.SetFocus
        Let ValidaDatos = False
        Exit Function
    End If
    
    If DateDiff("m", FpFecDesde.text, FpFecHasta.text) > 2 Then
        Call MsgBox("Rango De Fecha No Puede Ser Mayor a 3 Meses", vbInformation, Me.Caption)
        Let ValidaDatos = False
        Exit Function
        
    End If
End Function

Private Sub FpFecDesde_Change()
If IsDate(FpFecDesde.text) = False Then Exit Sub
End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
End Sub

Private Sub FpFecHasta_Change()
If IsDate(FpFecHasta.text) = False Then Exit Sub
End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
'Dim sql      As String
'Dim sdql_mvi As String
'Dim RS       As New ADODB.Recordset
'Select Case Index
'    Case 0
'        Set RS = vg_db.Execute("sgp_Sel_RegimenxCodigo " & Val(fpLongInteger1(0).Value) & "")
'        If RS.EOF = True Then
'            RS.Close
'            Set RS = Nothing
'            fpayuda(1).Caption = ""
'            Call MoverGrilla(fpText.text)
'            Exit Sub
'        End If
'        fpayuda(1).Caption = Trim(RS!reg_nombre)
'        RS.Close: Set RS = Nothing
'        Call MoverGrilla(fpText.text)
'
'    Case 1
'        Set RS = vg_db.Execute("sgp_Sel_ServicioxCodigo " & Val(fpLongInteger1(1).Value) & "")
'        If RS.EOF Then
'            RS.Close:
'            Set RS = Nothing
'            fpayuda(2).Caption = ""
'            Call MoverGrilla(fpText.text)
'            Exit Sub
'        End If
'        fpayuda(2).Caption = Trim(RS!ser_nombre)
'        RS.Close: Set RS = Nothing
'        Call MoverGrilla(fpText.text)
'End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
Call MoverGrilla(fpText.text)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    If Index = 0 Then Image1_Click 1
    If Index = 1 Then Image1_Click 2
End Select
End Sub

Private Sub fpLongInteger1_LostFocus(Index As Integer)
Dim Sql      As String
Dim sdql_mvi As String
Dim RS       As New ADODB.Recordset
Select Case Index
    Case 0
        Set RS = vg_db.Execute("sgp_Sel_RegimenxCodigo " & IIf(Val(fpLongInteger1(0).Value) = 0, -1, Val(fpLongInteger1(0).Value)) & "")
        If RS.EOF = True Then
            RS.Close
            Set RS = Nothing
            fpayuda(1).Caption = ""
            Exit Sub
        End If
        fpayuda(1).Caption = Trim(RS!reg_nombre)
        RS.Close: Set RS = Nothing
    Case 1
        Set RS = vg_db.Execute("sgp_Sel_ServicioxCodigo " & IIf(Val(fpLongInteger1(1).Value) = 0, -1, Val(fpLongInteger1(1).Value)) & "")
        If RS.EOF Then
            RS.Close:
            Set RS = Nothing
            fpayuda(2).Caption = ""
            Exit Sub
        End If
        fpayuda(2).Caption = Trim(RS!ser_nombre)
        RS.Close: Set RS = Nothing
End Select
Call MoverGrilla(fpText.text)
End Sub

Private Sub fpText_Change()
Dim RS As New ADODB.Recordset

    RS.Open "SELECT cli_codigo, cli_nombre " & _
            "FROM b_clientes " & _
            "WHERE cli_codigo = '" & fpText.text & "' " & _
            "AND   cli_tipo   = 0 " & _
            "AND   cli_tipominuta in ('1')", vg_db, adOpenStatic
    If RS.EOF Then
        RS.Close
        Set RS = Nothing
        fpayuda(0).Caption = ""
        fpLongInteger1(0).Value = ""
        fpayuda(1).Caption = ""
        fpLongInteger1(1).Value = ""
        fpayuda(2).Caption = ""
        vaSpread1.MaxRows = 0
        FpFecDesde.Enabled = True
        FpFecHasta.Enabled = True
        Exit Sub
    End If
    fpayuda(0).Caption = Trim(RS!cli_nombre)
    fpText.text = RS!cli_codigo
    RS.Close
    Set RS = Nothing
    fpLongInteger1(0).Value = ""
    fpayuda(1).Caption = ""
    
    fpLongInteger1(1).Value = ""
    fpayuda(2).Caption = ""
    FpFecDesde.Enabled = True
    FpFecHasta.Enabled = True
    Call MoverGrilla(fpText.text)

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
Dim RS As New ADODB.Recordset
    Select Case Index
        Case 0
            vg_left = fpayuda(0).Left + 2300
            vg_nombre = "": vg_codigo = ""
            Call B_TabEst.LlenaDatos("b_clientes", "cli_", "Clientes", "Cliente_SitioRemoto")
            Call B_TabEst.Show(1)
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpText.text = vg_codigo
            fpayuda(0).Caption = vg_nombre
            fpLongInteger1(0).Value = ""
            Let fpayuda(1).Caption = ""
            fpLongInteger1(1).Value = ""
            Let fpayuda(2).Caption = ""
            fpLongInteger1(0).SetFocus
        Case 1
            vg_left = fpayuda(1).Left + 2300
            vg_nombre = "": vg_codigo = ""
            Call B_TabEst.LlenaDatos("a_regimen", "reg_", "Regimen", "RegBlo")
            Call B_TabEst.Show(1)
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(0).Value = Val(vg_codigo)
            fpLongInteger1(0).SetFocus
            fpayuda(1).Caption = vg_nombre
            fpLongInteger1(1).SetFocus
        Case 2
            Let vg_left = fpayuda(2).Left + 2300
            Let vg_nombre = ""
            Let vg_codigo = ""
            Call B_TabEst.LlenaDatos("a_servicio", "ser_", "Servicio", "SerBlo")
            Call B_TabEst.Show(1)
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(1).Value = "0"
            fpLongInteger1(1).Value = Val(vg_codigo)
            fpayuda(2).Caption = vg_nombre
            fpLongInteger1(1).SetFocus
            FpFecDesde.Enabled = True
            FpFecHasta.Enabled = True
            Call FpFecDesde.SetFocus
    End Select
End Sub

Private Sub Text1_Change(Index As Integer)
Dim Col As Long
Dim i As Long
Dim indactivo As Long
Col = IIf(Index = 0, 1, IIf(Index = 1, 2, IIf(Index = 2, 3, IIf(Index = 3, 4, 5))))
vaSpread1.Visible = False
If Trim(Text1(Index).text) <> "" Then
   For i = 1 To vaSpread1.MaxRows
       vaSpread1.Row = i
       vaSpread1.Col = Col
       indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(Text1(Index).text) & "*"
       vaSpread1.Col = Col
       If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
          If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Else
          If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
       End If
   Next i
   vaSpread1.SetActiveCell Col, 1
End If
vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
If Trim(Text1(Index).text) = "" Then
   For i = 1 To vaSpread1.MaxRows
       vaSpread1.Row = i
       If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
   Next
   vaSpread1.SetActiveCell Col, vaSpread1.SearchCol(Col, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
   vaSpread1.SetActiveCell Col, 1
End If
vaSpread1.Visible = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim RS        As New ADODB.Recordset
Dim RS1       As New ADODB.Recordset
Dim tipmin    As String
Dim Sql       As String
Dim i         As Long
Dim estsel    As Boolean
Dim IdBloque  As Long

    Select Case Button.Index
        Case 1, 3
            If ValidaDatos = False Then Exit Sub
            
            '------->Validar Minuta Bloque
            Sql = ""
            Sql = LimpiaDato(Trim(fpText.text))
            Sql = Sql & ", " & fpLongInteger1(0).Value & ", " & fpLongInteger1(1).Value & ", " & Format(FpFecDesde.text, "yyyymmdd") & ", " & Format(FpFecHasta.text, "yyyymmdd")
            If Button.Index = 1 Then
               vg_IDBloque = 0
               Set RS = vg_db.Execute("sgp_Sel_ValidarMinutaBloqueNuevo " & Sql & "")
               If Not RS.EOF Then
                  MsgBox "Datos ingresados ya existe con el bloque : " & RS!ID_Bloque, vbExclamation + vbOKOnly, Me.Caption
                  RS.Close
                  Set RS = Nothing
                  Exit Sub
               End If
               RS.Close
               Set RS = Nothing
            ElseIf Button.Index = 3 Then
               'validar que haya un datos seleccionado en la grilla
               estsel = False
               For i = 1 To vaSpread1.MaxRows
                   vaSpread1.Row = i
                   vaSpread1.Col = 1
                   If vaSpread1.text = "1" Then
                      vaSpread1.Col = 2
                      vg_IDBloque = Val(vaSpread1.text)
                      estsel = True
                   End If
               Next i
               If Not estsel Then
                  MsgBox "Seleccione un bloque del detalle de la grilla", vbExclamation + vbOKOnly, Me.Caption
                  Exit Sub
               End If
            
               Set RS = vg_db.Execute("sgp_Sel_ValidarMinutaBloqueModificado " & Sql & "")
               If RS.EOF Then
                  MsgBox "Datos modificar no corresponde", vbExclamation + vbOKOnly, Me.Caption
                  RS.Close
                  Set RS = Nothing
                  Exit Sub
               End If
               If vg_IDBloque = 0 Then vg_IDBloque = RS!ID_Bloque
               RS.Close
               Set RS = Nothing
           
            End If
            '-------> Validar clientes
            Set RS = vg_db.Execute("select cli_codigo, cli_nombre from b_clientes WITH (NoLock) where cli_codigo = '" & fpText.text & "' and cli_tipo = 0 and cli_tipominuta = 1")
            If RS.EOF Then
                RS.Close
                Set RS = Nothing
                fpText.text = ""
                fpayuda(0).Caption = ""
                fpLongInteger1(0).Value = ""
                fpayuda(1).Caption = ""
                fpLongInteger1(1).Value = ""
                fpayuda(2).Caption = ""
                MsgBox "No existe contrato", vbExclamation + vbOKOnly, Me.Caption
                Exit Sub
            End If
            fpayuda(0).Caption = RS!cli_nombre
            RS.Close
            Set RS = Nothing
            '-------> Validar regimen
            Set RS = vg_db.Execute("select reg_codigo, reg_nombre from a_regimen WITH (NoLock) where reg_codigo = " & Val(fpLongInteger1(0).Value) & " and reg_codigo >9999")
            If RS.EOF Then
                RS.Close
                Set RS = Nothing
                MsgBox "No Existe Regimen", vbExclamation + vbOKOnly, Me.Caption
                Exit Sub
            End If
            fpayuda(1).Caption = RS!reg_nombre
            RS.Close
            Set RS = Nothing

            '-------> Validar servicio
            Set RS = vg_db.Execute("select ser_codigo, ser_nombre from a_servicio WITH (NoLock) where ser_codigo = " & Val(fpLongInteger1(1).Value) & " and ser_codigo >9999")
            If RS.EOF Then
                RS.Close
                Set RS = Nothing
                MsgBox "No Existe Servicio", vbExclamation + vbOKOnly, Me.Caption
                Exit Sub
            End If
            fpayuda(2).Caption = RS!ser_nombre
            RS.Close
            Set RS = Nothing

            '-------> Validar estructura
            Set RS = vg_db.Execute("SELECT * FROM a_estservicio With (NoLock) WHERE ess_codser = " & Val(fpLongInteger1(1).Value) & " ORDER BY ess_orden")
            If RS.EOF Then
               RS.Close
               Set RS = Nothing
               MsgBox "No Existe estructura de servicio", vbExclamation + vbOKOnly, Me.Caption
               Exit Sub
            End If
            RS.Close
            Set RS = Nothing

            
            '*****************---->Validar minuta en uso <---------------------------
            '------ Esta funcion crea una tabla temporal concatenando los parametros ingresaods
            '------ para la minuta, de esta manera permanece una tabla temporal identificando
            '------ que alguien se encuentra conectado a esa minuta, si alguien
            '------ mas quiere acceder, se dara un aviso que esta en uso
            '------ esta tabla temporal se destruye cuando se cierra este formulario (evento Unload)
            '------ y tambien si el usuario cierra la sesion SQL Server la destruye automaticamente.
            '----------------------------------------------------------------------
                
'                Dim RSTempCheck As New ADODB.Recordset
'                Dim RSTem As New ADODB.Recordset
'                Dim RSinsert As New ADODB.Recordset
'                Dim NameTemp As String
'                NameTemp = LimpiaDato(Trim(fpText.text)) & Val(fpLongInteger1(0).Value) & Val(fpLongInteger1(1).Value) & Format(FpFecDesde.text, "yyyymmdd") & Format(FpFecDesde.text, "yyyymmdd")
            
'                Set RSTempCheck = vg_db.Execute("select * from tempdb.dbo.sysobjects where xtype = 'U' and name = '##ValidaMinutaSitioRemoto_" & NameTemp & "'")
                
'                If RSTempCheck.EOF And RSTempCheck.BOF Then
'                    Set RSTem = vg_db.Execute("CREATE TABLE ##ValidaMinutaSitioRemoto_" & NameTemp & " (usu_codigo VarChar(20))")
'                    Set RSinsert = vg_db.Execute("INSERT INTO ##ValidaMinutaSitioRemoto_" & NameTemp & " (usu_codigo) values ('" & vg_NUsr & "')")
'            '        Set RS = Nothing
'            '        Set RSTem = Nothing
'                Else
'                    Set RS = vg_db.Execute("SELECT usu_codigo from ##ValidaMinutaSitioRemoto_" & NameTemp & " ")
'                    If Not (RS.EOF = True And RS.BOF = True) Then
'                        RS.MoveFirst
'                        MsgBox "La minuta con los parametros ingresados, actualmente esta siendo usada por el usuario: '" & UCase(RS!usu_codigo) & "', podra ingresar cuando el usuario termine de trabajar en ella"
'                        RS.Close: Set RS = Nothing
'                        Exit Sub
'                    End If
'                    RS.Close: Set RS = Nothing
'                End If
            
            'RSTempCheck.Close
            'Set RSTempCheck = Nothing
                        
            vg_codcasino = LimpiaDato(Trim(fpText.text))
            vg_codregimen = Val(fpLongInteger1(0).Value)
            vg_codservicio = Val(fpLongInteger1(1).Value)
            Let Vg_FechaDesde = Format(FpFecDesde.text, "yyyymmdd")
            Let Vg_FechaHasta = Format(FpFecHasta.text, "yyyymmdd")
            
            Let NomFor = "MINTEO"
            Let tipmin = "1"
            Unload B_Receta
               
            Unload M_MinSR2
            
            Call M_MinSR2.Show(1)
               
'            DropTebleTmp (LimpiaDato(Trim(fpText.text)) & Val(fpLongInteger1(0).Value) & Val(fpLongInteger1(1).Value) & Mid(FpFecDesde.text, 4, 4) & Mid(FpFecDesde.text, 1, 2))
        Case 5 'Eliminar registro minuta bloque
            estsel = False
            IdBloque = 0
            For i = 1 To vaSpread1.MaxRows
                vaSpread1.Row = i
                vaSpread1.Col = 1
                If vaSpread1.text = "1" Then
                   vaSpread1.SetActiveCell 1, vaSpread1.Row: vaSpread1.SetFocus
                   '-------> Validar si la minuta esta bloqueada
                   If vaSpread1.BackColor = Shape1(1).FillColor Then
                      MsgBox "Minuta bloque esta bloqueada, proceso cancelado", vbInformation + vbOKOnly, MsgTitulo
                      Exit Sub
                   End If
                   If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
                   vaSpread1.Col = 2
                   IdBloque = vaSpread1.text
                   Sql = ""
                   Sql = LimpiaDato(Trim(fpText.text)) & ", " & IdBloque
                   Set RS = vg_db.Execute("sgp_Del_MinutaBloque " & Sql & "")
                   If Not RS.EOF Then
                      If UCase(RS(0)) = "OK" Then
                         MsgBox "Registro eliminado exitosamente", vbInformation + vbOKOnly, MsgTitulo
                         estsel = True
                         vaSpread1.DeleteRows vaSpread1.Row, 1
                         vaSpread1.MaxRows = vaSpread1.MaxRows - 1
                      Else
                         MsgBox "Registro finalizo con error " & RS(0), vbInformation + vbOKOnly, MsgTitulo
                      End If
                   End If
                End If
            Next i
            If Not estsel Then MsgBox "No existe minuta bloque seleccionada", vbExclamation + vbOKOnly, Me.Caption
        Case 8 'Salir
            Unload B_Receta
            Me.Hide
            Unload Me
            Unload M_MinSR1
    End Select
End Sub

'Sub DropTebleTmp(NameTable As String)
''*****************----> Destruye Tabla temporal<---------------------------
''---- Destruye tabla temporal, de manera que desbloquee el acceso a la minuta

'    Dim RSTempCheck As New ADODB.Recordset
'    Dim RSTem As New ADODB.Recordset

'    Set RSTempCheck = vg_db.Execute("select * from tempdb.dbo.sysobjects where xtype = 'U' and name = '##ValidaMinutaSitioRemoto_" & NameTable & "'")
'        If Not (RSTempCheck.EOF = True And RSTempCheck.BOF = True) Then
'            Set RSTem = vg_db.Execute("Drop Table ##ValidaMinutaSitioRemoto_" & NameTable & " ")
'        End If

'    RSTempCheck.Close
'    Set RSTempCheck = Nothing
'    Set RSTem = Nothing
'End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim estsel As Boolean
If ButtonDown = 0 Then
   FpFecDesde.Enabled = True
   FpFecHasta.Enabled = True
End If
If est Or ButtonDown = 0 Or vaSpread1.MaxRows < 1 Then Exit Sub
Dim i As Long
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1
    If i <> Row Then
       If vaSpread1.text = "1" Then
          est = True
          vaSpread1.text = "0"
          est = False
        End If
    End If
Next i

vaSpread1.Row = Row
vaSpread1.Col = 2 'Mover Bloque Minuta
vg_IDBloque = Val(vaSpread1.text)
fpLongInteger1(2).Value = vaSpread1.text
vaSpread1.Col = 7 'Mover Regimen
fpLongInteger1(0).Value = vaSpread1.text
vaSpread1.Col = 9 'Mover descripción Regimen
fpayuda(1).Caption = vaSpread1.text
vaSpread1.Col = 8 'Mover Servicio
fpLongInteger1(1).Value = vaSpread1.text
vaSpread1.Col = 10 'Mover descripción Servicio
fpayuda(2).Caption = vaSpread1.text
vaSpread1.Col = 5 'Mover Fecha Desde
FpFecDesde.text = vaSpread1.text
vaSpread1.Col = 6 'Mover Fecha Hasta
FpFecHasta.text = vaSpread1.text
FpFecDesde.Enabled = False
FpFecHasta.Enabled = False
End Sub

'Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
'If vaSpread1.MaxRows < 1 Then Exit Sub
'vaSpread1.Row = Row
'vaSpread1.Col = 2 'Mover Bloque Minuta
'fpLongInteger1(2).Value = vaSpread1.text
'vaSpread1.Col = 7 'Mover Regimen
'fpLongInteger1(0).Value = vaSpread1.text
'vaSpread1.Col = 8 'Mover Servicio
'fpLongInteger1(1).Value = vaSpread1.text
'vaSpread1.Col = 5 'Mover Fecha Desde
'FpFecDesde.text = vaSpread1.text
'vaSpread1.Col = 6 'Mover Fecha Hasta
'FpFecHasta.text = vaSpread1.text
'End Sub
