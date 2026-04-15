VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_DevolucionServicioEspeciales 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolución de Servicios Especiales"
   ClientHeight    =   7740
   ClientLeft      =   1740
   ClientTop       =   1950
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   11460
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   840
         Width           =   8925
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4830
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   470
         Width           =   2325
      End
      Begin VB.Frame Frame4 
         Height          =   45
         Left            =   30
         TabIndex        =   5
         Top             =   2265
         Visible         =   0   'False
         Width           =   8520
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   0
         Left            =   1350
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2760
         Picture         =   "M_DevolucionSalidaEspeciales.frx":0000
         Top             =   50
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   4890
         TabIndex        =   22
         Top             =   510
         Width           =   2310
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3165
         TabIndex        =   20
         Top             =   135
         Width           =   3975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ser. Esp."
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
         TabIndex        =   19
         Top             =   900
         Width           =   795
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
         Index           =   7
         Left            =   75
         TabIndex        =   18
         Top             =   525
         Width           =   1050
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
         TabIndex        =   17
         Top             =   195
         Width           =   1365
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
         TabIndex        =   16
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
         Left            =   3915
         TabIndex        =   15
         Top             =   525
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   21
         Top             =   165
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6090
      Left            =   15
      TabIndex        =   1
      Top             =   1590
      Width           =   11460
      Begin VB.Frame Frame3 
         Height          =   450
         Left            =   9840
         TabIndex        =   25
         Top             =   5400
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
            TabIndex        =   26
            Top             =   135
            Width           =   1230
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5085
         Left            =   105
         TabIndex        =   2
         Top             =   225
         Width           =   11295
         _Version        =   393216
         _ExtentX        =   19923
         _ExtentY        =   8969
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
         SpreadDesigner  =   "M_DevolucionSalidaEspeciales.frx":030A
         ScrollBarTrack  =   3
         ClipboardOptions=   0
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
         Left            =   8880
         TabIndex        =   24
         Top             =   5640
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   195
         Top             =   5520
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
         Left            =   555
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
Attribute VB_Name = "M_DevolucionServicioEspeciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim est As Boolean
Dim MsgTitulo As String

Private Sub Combo1_Click(Index As Integer)

On Error GoTo Man_Error

Dim feprod As Long, codser As Long, fil As Long, codreg As Long, aAp As String, codsec As String, coding As String, opcsal As String
Dim sql1 As String, sql2 As String
Dim NumDoc As Long
Dim NumDocAso As Long
Dim i As Long
Dim RS     As New ADODB.Recordset

If est Then Exit Sub

NumDocAso = 0

Select Case Index

    Case 0
        
        '-------> Validar si el contrato tiene asignado inventario rotativo
        If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
           Format(fpDateTime1(0).Value, "dd/mm/yyyy") > CDate(vg_ciedia) Then
           
           MsgBox "Tiene que realizar cierre diario...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If Combo1(0).ListIndex = -1 Or Combo1(0).text = "" Then Exit Sub
              
        sql1 = Format(fpDateTime1(0).text, "yyyymmdd")
        NumDoc = Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 3, 10)) 'Trim(Combo1(0).text)
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        Set RS = vg_db.Execute("sgp_Sel_ListaDevCerradaVentaServiciosEspeciales '" & Trim(LimpiaDato(fpText1(0).text)) & "', " & vg_codbod & ", '" & sql1 & "', " & NumDoc & "")
        
        If Not RS.EOF Then
            
            MsgBox "Devolución ya fue realizada...", vbExclamation + vbOKOnly, MsgTitulo
            DevExiste RS!tos_numero_documento
            RS.Close
            Set RS = Nothing
            Exit Sub
        
        End If
        RS.Close
        Set RS = Nothing
        
        Me.MousePointer = 11
        Gl_Ac_Botones Me, 4, 6, ""
        
        sql1 = Format(fpDateTime1(0).text, "yyyymmdd")
        sql2 = Trim(Combo1(0).text)
                
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("sgp_Sel_DetalleSalDevVentaServiciosEspeciales '" & LimpiaDato(Trim(fpText1(0).text)) & "', 'SE', '" & sql1 & "', " & vg_codbod & ", " & NumDoc & " ")
        
        If RS.EOF Then
            
            RS.Close
            Set RS = Nothing
            
            MsgBox "No existe salida ventas servicio especiales...", vbExclamation + vbOKOnly, MsgTitulo
            Me.MousePointer = 0
            Exit Sub
        
        End If
        vaSpread1.Visible = False
        vaSpread1.MaxRows = 0
        
        i = 0
        codsec = "0"
        coding = ""
        
        Do While Not RS.EOF
           
           '------- Productos
           i = i + 1
           
           vaSpread1.MaxRows = i
           vaSpread1.Row = i
           
           vaSpread1.Col = 1
           vaSpread1.text = RS!pro_codigo
           
           vaSpread1.Col = 2
           vaSpread1.text = RS!pro_nombre
           
           vaSpread1.Col = 3
           vaSpread1.text = RS!uni_nomcor
           
           vaSpread1.Col = 4
           vaSpread1.text = Format(RS!des_Cantidad_Mercaderia, fg_Pict(9, vg_DCa))
           
           vaSpread1.Col = 5
           vaSpread1.ForeColor = &HFF0000
           vaSpread1.text = Format(0, fg_Pict(9, vg_DCa))
           
           vaSpread1.Col = 6
           vaSpread1.text = Format(RS!des_Precio_Documento, fg_Pict(9, 2)) 'vg_DPr))
           
           vaSpread1.Col = 7
           vaSpread1.text = Format(0, fg_Pict(9, vg_DPr))
           
           vaSpread1.Col = 8
           vaSpread1.text = "NP" 'No bloquedo - Producto
           
           vaSpread1.Col = -1
           vaSpread1.BackColor = Shape1(1).FillColor
                     
           vaSpread1.Col = 9
           vaSpread1.text = Format(RS!bod_canmer, fg_Pict(9, vg_DCa))
                        
           RS.MoveNext
        
        Loop
        RS.Close
        Set RS = Nothing
        
        Me.MousePointer = 0
        Frame1.Enabled = False
        vaSpread1.Visible = True
        
        If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
               
    Case 1
        
        If vaSpread1.MaxRows = 0 Then Exit Sub
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute("SELECT bod.bod_canmer FROM b_productos pro, b_bodegas bod WHERE bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                                   "AND    bod.bod_codpro = pro.pro_codigo AND pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "'")
            
            vaSpread1.Col = 9
            
            If Not RS.EOF Then vaSpread1.text = Format(RS!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
            
            RS.Close
            Set RS = Nothing
        
        Next i

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

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

Me.Height = 8175
Me.Width = 11625

fg_centra Me
est = False

Me.HelpContextID = vg_OpcM
MsgTitulo = "Devolución Venta Servicios Especiales"

EspFecha fpDateTime1(0)

Dim X As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
Gl_Mo_Botones Me, 4

vaSpread1.Row = -1
vaSpread1.Col = 4: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
'vaSpread1.Col = 6: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
vaSpread1.Col = 6: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = 2
vaSpread1.Col = 7: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr

'-------> Cargar Combo Bodega
CargarDatoCombo Combo1, 1, "b_clientes", "cli_", "CliBod", "N"

'Limpia
Limpia 2

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpDateTime1_Change(Index As Integer)

On Error GoTo Man_Error

If est Then Exit Sub

Select Case Index

Case 0

    Combo1(0).Clear
    vaSpread1.MaxRows = 0

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpDateTime1_GotFocus(Index As Integer)

On Error GoTo Man_Error

Select Case Index
Case 0
    
    Toolbar1.Buttons(8).Enabled = False

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

Private Sub fpDateTime1_LostFocus(Index As Integer)

On Error GoTo Man_Error

Dim Tipo As String, sql1 As String, sql2 As String
Dim RS As New ADODB.Recordset

Select Case Index

Case 0
    
    Toolbar1.Buttons(8).Enabled = True
    If Trim(fpDateTime1(0).text) = "" Or Trim(fpText1(0).text) = "" Then
    
       Exit Sub
    
    End If
       
    sql2 = Format(fpDateTime1(0).text, "yyyymmdd")
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS = vg_db.Execute("sgp_Sel_ListaSalCerradaVentaServiciosEspeciales '" & LimpiaDato(Trim(fpText1(0).text)) & "', " & vg_codbod & ", '" & sql2 & "'")
    
    Combo1(0).Clear
    
    Do While Not RS.EOF
        
        Combo1(0).AddItem "Nro. Doc. - " & RS!tos_numero_documento & " - " & RS!tos_Venta_servicio_Especiales & Space(150) & "(" & fg_pone_cero(Str(RS!tos_numero_documento), 10) & ")"
        RS.MoveNext

    Loop
    RS.Close
    Set RS = Nothing
    
    If Combo1(0).listcount = 0 Then
    
       MsgBox "No existe salida ventas servicios especiales...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

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

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If fpText1(0).text = "" Then Exit Sub

Set RS = vg_db.Execute("SELECT cli_nombre FROM b_clientes WHERE cli_codigo='" & fpText1(0).text & "' AND cli_tipo=0")

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      fpayuda(Index).Caption = RS!cli_nombre
      Gl_Ac_Botones Me, 4, 2, ""
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

fpLongInteger1(0).text = TraerCorrelativo(vg_codbod, "DE")

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

vg_codigo = 0

Select Case Index

    Case 0
        
        vg_left = fpayuda(Index).Left + 1920
        B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
        B_TabEst.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
        If Trim(vg_codigo) <> fpText1(Index) Then Limpia 2
        fpText1(Index) = Trim(vg_codigo)
        fpayuda(Index).Caption = vg_nombre
        fpText1_LostFocus 1
        If fpDateTime1(0).Enabled = True Then fpDateTime1(0).SetFocus
        Gl_Ac_Botones Me, 4, 2, ""
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim rutcli As String, tipdoc As String, NumDoc As Long, codbod   As Long, fecemi As Date, fecpro As String, codreg As Long, codser As Long, i As Long, canact As Double
Dim numlin As Long, codmer As String, coding As String, candev As Double, canmer As Double, predoc As Double, ptotal As Double, descri As String, diablq As Date, color As String, codsec As String
Dim NumDocAso        As Long
Dim RS               As New ADODB.Recordset
Dim MyBuffer         As String
Dim ServicioEspecial As String
Dim canmin           As Double
Dim total            As Double

codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
fecpro = Format(fpDateTime1(0).Value, "dd/mm/yyyy")
If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
TraerFechaCierre

Select Case Button.Index

Case 1, 6 '-------> Nuevo

'    Limpia
    If Button.Index = 6 And vaSpread1.MaxRows > 0 Then If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    Limpia IIf(Button.Index = 1, 6, 2)
    If fpText1(0).Enabled = True Then fpText1(0).SetFocus

Case 8 '-------> Graba
    
    If Trim(fpText1(0).text) = "" Or Trim(fpLongInteger1(0).text) = "" Or Trim(Combo1(0).text) = "" Or Trim(fpDateTime1(0).text) = "" _
    Or Trim(Combo1(1).text) = "" Then
    
       MsgBox "Debe ingresar dato importante...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then
    
       MsgBox "Documento no corresponde al periodo : " & VgLinea & VgLinea & CierreFecha, vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then
    
       MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    'Validar inventario calendarizado 20201001
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 38) Then
        
       MsgBox "Se esta realizando la toma de inventario en estos momento...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
    End If
        
    'Validar ingreso documento inventario calendarizado 20201001
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 40) Then
        
       MsgBox "No puede ingresar documento, antes de un inventario calendarizado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
    End If
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 8) Then
    
       MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Then
    
       MsgBox "Día se encuentra cerrado, no es posible ingresar...", vbExclamation + vbQuestion, MsgTitulo: Exit Sub
    
    End If
    
    rutcli = Trim(LimpiaDato(fpText1(0).text))
    tipdoc = "DE"
    codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
    fecpro = Format(fpDateTime1(0).text, "yyyymmdd")
'    ServicioEspecial = Trim(Trim(Combo1(0).text))
    NumDocAso = Val(Mid(Combo1(0), Len(Trim(Combo1(0).text)) - 2, 10)) 'Trim(Combo1(0).text)
    NumDoc = TraerCorrelativo(codbod, "DE")
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS = vg_db.Execute("sgp_Sel_DocCerradoxUsuarioDevVentaServiciosEspeciales '" & Trim(LimpiaDato(fpText1(0).text)) & "', " & NumDoc & ", " & codbod & "")
           
    If Not RS.EOF Then
              
       RS.Close
       Set RS = Nothing
       Gl_Ac_Botones Me, 12, 3, ""
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
    
    '-------> Validar cantidad devolver
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        
        vaSpread1.Col = 4
        canmin = vaSpread1.text
        
        vaSpread1.Col = 5
        candev = vaSpread1.text
        
        If candev > canmin Then
        
           MsgBox "Cantidad devolver es mayor cantidad salida...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub

        End If
        
    Next i
    
    total = 0
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 7
        ptotal = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DPr)
        total = total + ptotal
    
    Next i
    
    If total = 0 Then
    
       MsgBox "El total del documento debe ser mayor a cero...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub

    End If
    
paso:
    
    DoEvents
        
    '------- Detalle
    total = 0
    numlin = 1
    
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
        
        vaSpread1.Col = 2
        descri = Trim(LimpiaDato(vaSpread1.text))
        
        vaSpread1.Col = 4
        canmer = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DCa)
        
        vaSpread1.Col = 5
        candev = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DCa)
        
        vaSpread1.Col = 6
        predoc = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), 2) 'vg_DPr)
        
        vaSpread1.Col = 7
        ptotal = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DPr)
        
        vaSpread1.Col = 8
        color = Right(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), 1)
        
'        If candev > 0 Then
        
           MyBuffer = MyBuffer & " <Venta"
           MyBuffer = MyBuffer & " NLin = " & Chr(34) & numlin & Chr(34)
           MyBuffer = MyBuffer & " IdProd = " & Chr(34) & codmer & Chr(34)
           MyBuffer = MyBuffer & " CanMer = " & Chr(34) & canmer & Chr(34)
           MyBuffer = MyBuffer & " CanDev = " & Chr(34) & candev & Chr(34)
           MyBuffer = MyBuffer & " Precio = " & Chr(34) & predoc & Chr(34)
           MyBuffer = MyBuffer & "/>"
        
           numlin = numlin + 1
           
'        End If
    
    Next i
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    MyBuffer = MyBuffer & "</GrabaVenta>"
    Set RS = vg_db.Execute("sgp_Ins_XmlDevolucionVentaServiciosEspeciales '" & MyBuffer & "', '" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", '" & fecpro & "', " & codbod & ", '" & vg_NUsr & "', " & NumDocAso & "")
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          MsgBox RS(0) & " - " & RS(1) & VgLinea & " Proceso termino con problemas...", vbCritical, MsgTitulo
      
          RS.Close
          Set RS = Nothing
          
          Toolbar1.Enabled = True
          Exit Sub
          
       End If
    
       fpLongInteger1(0).text = RS(3)
    
    End If
    RS.Close
    Set RS = Nothing
       
    Gl_Ac_Botones Me, 4, 3, ""
    Frame1.Enabled = False
    vaSpread1.Col = -1
    vaSpread1.Row = -1
    vaSpread1.Lock = True
    
    '-------> Revisa Stock
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "AND   bod.bod_codpro = pro.pro_codigo AND pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "'")
        
        vaSpread1.Col = 9
        If Not RS.EOF Then
        
           vaSpread1.text = Format(RS!bod_canmer, fg_Pict(9, vg_DCa))
           
        Else
           vaSpread1.text = 0
           
        End If
        RS.Close
        Set RS = Nothing
    
    Next i
    
    I_SalDevVentaServiciosEspeciales Me, "DE"

Case 3 '5 '-------> Anular
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then
    
       MsgBox "Periodo esta cerrado...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
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
        
    'Validar ingreso documento inventario calendarizado 20201001
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 40) Then
        
       MsgBox "No puede ingresar documento, antes de un inventario calendarizado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
    End If
    
    If MsgBox("Anula documento...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
       
    '-------> Detalle
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<GrabaVenta>"
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        numlin = i
        
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
        candev = Round(IIf(vaSpread1.Value = "", 0, vaSpread1.Value), vg_DCa)
        
        predoc = 0
        
        vaSpread1.Col = 8
        color = Right(vaSpread1.text, 1)
        
        MyBuffer = MyBuffer & " <Venta"
        MyBuffer = MyBuffer & " NLin = " & Chr(34) & i & Chr(34)
        MyBuffer = MyBuffer & " IdProd = " & Chr(34) & codmer & Chr(34)
        MyBuffer = MyBuffer & " CanMer = " & Chr(34) & canmer & Chr(34)
        MyBuffer = MyBuffer & " CanDev = " & Chr(34) & candev & Chr(34)
        MyBuffer = MyBuffer & " Precio = " & Chr(34) & predoc & Chr(34)
        MyBuffer = MyBuffer & "/>"
            
    Next i
    
    rutcli = Trim(LimpiaDato(fpText1(0).text))
    tipdoc = "DE"
    codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
    fecpro = Format(fpDateTime1(0).text, "yyyymmdd")
    NumDoc = fpLongInteger1(0).text
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
        
    MyBuffer = MyBuffer & "</GrabaVenta>"
    Set RS = vg_db.Execute("sgp_Upd_XmlDevolucionVentaServiciosEspeciales '" & MyBuffer & "', '" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", '" & fecpro & "', " & codbod & ", '" & vg_NUsr & "'")
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
        vaSpread1.Col = 1
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "AND   bod.bod_codpro = pro.pro_codigo AND pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "'")
        
        vaSpread1.Col = 9
        
        If Not RS.EOF Then
           
           vaSpread1.text = Format(RS!bod_canmer, fg_Pict(9, vg_DCa))
        
        Else
           
           vaSpread1.text = 0
        
        End If
        
        RS.Close
        Set RS = Nothing
    
    Next i
    
    Label1.Caption = "ANULADA"
        
    Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""

Case 11 '8 '-------> Busqueda
    
    If Trim(fpText1(0).text) = "" Then MsgBox "Debe seleccionar contrato...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    vg_codigo = Trim(fpText1(0).text)
    vg_nombre = "DE"
    B_SalBod.Show 1
    Me.MousePointer = 11
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Me.MousePointer = 0: Exit Sub
    DevExiste Val(vg_codigo)
    vg_codigo = ""
    Me.MousePointer = 0
    
Case 12 '9 '-------> Imprimir

    I_SalDevVentaServiciosEspeciales Me, "DE"
    
Case 15 '12 '-------> Salir
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub DevExiste(codigo As Long)

On Error GoTo Man_Error

Dim aAp As String, codsec As String, coding As String, sql1 As String, sql2 As String
Dim RS As New ADODB.Recordset
Dim i As Long
Frame1.Enabled = False
est = True

est = False

vaSpread1.Col = -1: vaSpread1.Row = -1
vaSpread1.Lock = True
vaSpread1.MaxRows = 0
vaSpread1.Visible = False

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_DetalleDevExistenteVentaServiciosEspeciales '" & LimpiaDato(Trim(fpText1(0).text)) & "', 'DE', " & vg_codbod & ", " & codigo & "")

If Not RS.EOF Then
    
   est = True
   fpLongInteger1(0).text = RS!tos_numdoc
   fpDateTime1(0).text = RS!tos_fecemi
   Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 10, fg_pone_cero(Str(RS!tos_codbod), 10))
   Combo1(0).Clear
   Combo1(0).AddItem RS!tos_Venta_servicio_Especiales
   Combo1(0).ListIndex = 0
   Label1.Caption = IIf(RS!tos_estdoc = "A", "ANULADA", "")
   est = False
    
    Do While Not RS.EOF
    
        '------- Productos
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 1
        vaSpread1.text = RS!pro_codigo
        
        vaSpread1.Col = 2
        vaSpread1.text = RS!pro_nombre
        
        vaSpread1.Col = 3
        vaSpread1.text = RS!uni_nomcor
        
        vaSpread1.Col = 4
        vaSpread1.text = Format(RS!dev_canmin, fg_Pict(9, vg_DCa))
        
        vaSpread1.Col = 5
        vaSpread1.text = Format(RS!dev_canmer, fg_Pict(9, vg_DCa))
        
        vaSpread1.Col = 6
        vaSpread1.text = Format(RS!dev_predoc, fg_Pict(9, 2)) 'vg_DPr))
        
        vaSpread1.Col = 7
        vaSpread1.text = Format(RS!dev_ptotal, fg_Pict(9, vg_DPr))
        
        vaSpread1.Col = 8
        vaSpread1.text = "NP" 'No bloquedo - Producto
        
        vaSpread1.Col = -1
        vaSpread1.BackColor = Shape1(1).FillColor
        
        vaSpread1.Col = 9
        vaSpread1.text = Format(RS!bod_canmer, fg_Pict(9, vg_DCa))
    
        RS.MoveNext
        
    Loop

End If
RS.Close
Set RS = Nothing

'------- Total General ---------
Dim subtot As Double
Dim Cantidad As Double
Dim Precio As Double

subtot = 0

For i = 1 To vaSpread1.MaxRows
            
    vaSpread1.Row = i
    vaSpread1.Col = 1
    
    Cantidad = 0
    Precio = 0
    
    If Trim(vaSpread1.text) <> "" Then
            
       vaSpread1.Col = 5
       Cantidad = IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0)
       
       vaSpread1.Col = 6
       Precio = IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0)
       
       subtot = subtot + Format((Cantidad * Precio), fg_Pict(9, vg_DPr))
            
    End If
            
Next
Label2.Caption = Format(subtot, fg_Pict(9, vg_DPr))

vaSpread1.Visible = True
Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Function MuestraFolio(Casino As String) As String

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

MuestraFolio = ""
If Trim(Casino) = "" Then Exit Function

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT tos_numero_documento FROM b_totventaserviciosespeciales WHERE tos_tipo_documento = 'DE' AND tos_IdBodega = " & vg_codbod & " ORDER BY tos_numero_documento DESC")
If Not RS.EOF Then RS.MoveFirst: MuestraFolio = RS!tos_numero_documento + 1 Else MuestraFolio = 1
RS.Close: Set RS = Nothing

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Function

Sub Limpia(op As Integer)

On Error GoTo Man_Error

Label1.Caption = ""
Label2.Caption = 0
Frame1.Enabled = True
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")

Combo1(0).ListIndex = -1
Combo1(1).ListIndex = IIf(Combo1(1).listcount = 1, 0, -1)

vaSpread1.MaxRows = 0
vaSpread1.Col = -1: vaSpread1.Row = -1
vaSpread1.Lock = True
vaSpread1.Col = 5: vaSpread1.Row = -1
vaSpread1.Lock = False

fpText1(0).Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText1(0).text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)

fpLongInteger1(0).text = TraerCorrelativo(vg_codbod, "DE")

Gl_Ac_Botones Me, 4, op, ""

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

    Case 5
    
        '------- Total General ---------
        subtot = 0
        
        For i = 1 To vaSpread1.MaxRows
                    
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            Cantidad = 0
            Precio = 0
    
            If Trim(vaSpread1.text) <> "" Then
                    
               vaSpread1.Col = 5
               Cantidad = IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0)
               
               vaSpread1.Col = 6
               Precio = IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0)
               
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

On Error GoTo Man_Error

Dim canrea As Double, propon As Double, codmer As String, cansto As Double, cansal As Double

vaSpread1.Row = Row
vaSpread1.Col = 4
cansal = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
vaSpread1.Col = 5
canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))

If ChangeMade = True And canrea > cansal Then vaSpread1.text = Format(0, fg_Pict(9, vg_DCa)): Exit Sub

vaSpread1.Col = 1
codmer = vaSpread1.text

vaSpread1.Col = 5
canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))

vaSpread1.Col = 6
propon = Format(vaSpread1.text, fg_Pict(9, 2)) 'vg_DPr))

vaSpread1.Col = 7
vaSpread1.text = Format(canrea * propon, fg_Pict(9, vg_DPr))

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
    
    vaSpread1.Row = i
    vaSpread1.Col = 8
    color = Right(vaSpread1.text, 1)

Next i

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
Dim color As String, canrea As Double, propon As Double, codtot As Double

Dim codmer As String, cansto As Double, cansal As Double
vaSpread1.Row = Row

Select Case Col

Case 4
    
    vaSpread1.Row = Row
    vaSpread1.Col = 4
    cansal = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
    vaSpread1.Col = 5: canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
    
    If canrea > cansal Then vaSpread1.text = Format(0, fg_Pict(9, vg_DCa)): Exit Sub
    
    vaSpread1.Col = 1: codmer = vaSpread1.text
    
    vaSpread1.Col = 5: canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
    
    vaSpread1.Col = 6: propon = Format(vaSpread1.text, fg_Pict(9, 2)) 'vg_DPr))
    
    vaSpread1.Col = 7: vaSpread1.text = Format(canrea * propon, fg_Pict(9, vg_DPr))

End Select

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
vaSpread1.Col = 9: Stock = Format(vaSpread1.text, fg_Pict(9, vg_DCa))

vaSpread1.Col = 2: Nombre = vaSpread1.text

TipText = "Bodega   : " & Trim(Left(Combo1(1).text, 50)) & vbCrLf & _
          "Producto : " & Trim(Nombre) & vbCrLf & _
          "Stock       : " & Format(Trim(Stock), fg_Pict(9, vg_DCa))

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

