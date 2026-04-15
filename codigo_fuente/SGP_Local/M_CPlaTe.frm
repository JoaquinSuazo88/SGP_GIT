VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_CPlaTe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Planificacón Teórica"
   ClientHeight    =   4380
   ClientLeft      =   9045
   ClientTop       =   2685
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4380
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos Destino"
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
      Height          =   2295
      Index           =   1
      Left            =   0
      TabIndex        =   17
      Top             =   2040
      Width           =   7475
      Begin VB.OptionButton Option1 
         Caption         =   "Mantener Rac. en Destino"
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
         Left            =   4680
         TabIndex        =   27
         Top             =   1920
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Usar Rac. Origen"
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
         Left            =   165
         TabIndex        =   26
         Top             =   1920
         Value           =   -1  'True
         Width           =   2655
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   2
         Left            =   1440
         TabIndex        =   6
         Top             =   750
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   3
         Left            =   1440
         TabIndex        =   7
         Top             =   1090
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   1440
         TabIndex        =   5
         Top             =   400
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   2
         Left            =   1440
         TabIndex        =   8
         Top             =   1440
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         Text            =   "13/07/2004"
         DateCalcMethod  =   3
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
         Index           =   3
         Left            =   5430
         TabIndex        =   24
         Top             =   1440
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         Text            =   "13/07/2004"
         DateCalcMethod  =   3
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   3360
         TabIndex        =   42
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   3360
         TabIndex        =   40
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3360
         TabIndex        =   38
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   11
         Left            =   6840
         TabIndex        =   31
         Top             =   1530
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   10
         Left            =   2880
         TabIndex        =   30
         Top             =   1530
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Final"
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
         Index           =   7
         Left            =   4275
         TabIndex        =   25
         Top             =   1530
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
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
         Left            =   165
         TabIndex        =   22
         Top             =   1530
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   165
         TabIndex        =   20
         Top             =   1180
         Width           =   750
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   165
         TabIndex        =   19
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   165
         TabIndex        =   18
         Top             =   465
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2685
         Picture         =   "M_CPlaTe.frx":0000
         Top             =   300
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   2685
         Picture         =   "M_CPlaTe.frx":030A
         Top             =   675
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   2685
         Picture         =   "M_CPlaTe.frx":0614
         Top             =   1020
         Width           =   480
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3405
         TabIndex        =   39
         Top             =   405
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   3405
         TabIndex        =   41
         Top             =   765
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   3405
         TabIndex        =   43
         Top             =   1125
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos Origen"
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
      Height          =   1815
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7455
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   750
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   1090
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
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
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         Text            =   "13/07/2004"
         DateCalcMethod  =   3
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   400
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   5430
         TabIndex        =   4
         Top             =   1440
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         Text            =   "13/07/2004"
         DateCalcMethod  =   3
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3240
         TabIndex        =   36
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3240
         TabIndex        =   34
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3240
         TabIndex        =   32
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   9
         Left            =   6840
         TabIndex        =   29
         Top             =   1530
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   8
         Left            =   2880
         TabIndex        =   28
         Top             =   1530
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
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
         Left            =   4275
         TabIndex        =   23
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
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
         Left            =   210
         TabIndex        =   21
         Top             =   1530
         Width           =   1110
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
         Index           =   2
         Left            =   210
         TabIndex        =   16
         Top             =   1180
         Width           =   705
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
         Index           =   0
         Left            =   210
         TabIndex        =   15
         Top             =   840
         Width           =   750
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
         Index           =   3
         Left            =   210
         TabIndex        =   14
         Top             =   465
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2685
         Picture         =   "M_CPlaTe.frx":091E
         Top             =   300
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2685
         Picture         =   "M_CPlaTe.frx":0C28
         Top             =   675
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2685
         Picture         =   "M_CPlaTe.frx":0F32
         Top             =   1020
         Width           =   480
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3285
         TabIndex        =   33
         Top             =   405
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3285
         TabIndex        =   35
         Top             =   765
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3285
         TabIndex        =   37
         Top             =   1125
         Width           =   3975
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1035
      Left            =   1560
      ScaleHeight     =   975
      ScaleWidth      =   4785
      TabIndex        =   10
      Top             =   4800
      Visible         =   0   'False
      Width           =   4845
      Begin MSComctlLib.ProgressBar gauge 
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Procesando Necedidad De Insumos"
         ForeColor       =   &H80000018&
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   4515
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   4380
      Left            =   7485
      TabIndex        =   9
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   7726
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_CPlaTe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
fg_centra Me
EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)
EspFecha fpDateTime1(2)
EspFecha fpDateTime1(3)
Msgtitulo = "Copiar Planificación Teórica"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpText(0).Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText(0).text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
fpText(1).Enabled = ModCasino
Image1(3).Enabled = ModCasino
fpText(1).text = MuestraCasino(1)
fpayuda(3).Caption = MuestraCasino(2)
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
Label1(8).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(0).text, "yyyymmdd"), 1), 1, 4)
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
Label1(9).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(1).text, "yyyymmdd"), 1), 1, 4)
fpDateTime1(2).text = Format(Date, "dd/mm/yyyy")
Label1(10).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(2).text, "yyyymmdd"), 1), 1, 4)
fpDateTime1(3).text = Format(Date, "dd/mm/yyyy")
Label1(11).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(3).text, "yyyymmdd"), 1), 1, 4)
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
Select Case Index
Case 0
    Label1(8).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(0).text, "yyyymmdd"), 1), 1, 4)
Case 1
    Label1(9).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(1).text, "yyyymmdd"), 1), 1, 4)
Case 2
    Label1(10).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(2).text, "yyyymmdd"), 1), 1, 4)
Case 3
    Label1(11).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(3).text, "yyyymmdd"), 1), 1, 4)
End Select
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Select Case Index
Case 0
    RS.Open "SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    If ((Val(fpLongInteger1(0).Value) < 10000 And Val(fpLongInteger1(1).Value) < 10000) And (Val(fpLongInteger1(2).Value) > 9999 Or Val(fpLongInteger1(3).Value) > 9999)) Then Toolbar1.Buttons(2).Enabled = falso Else Toolbar1.Buttons(2).Enabled = True
Case 1
    RS.Open "SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(1).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "": Exit Sub
    fpayuda(2).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    If ((Val(fpLongInteger1(0).Value) < 10000 And Val(fpLongInteger1(1).Value) < 10000) And (Val(fpLongInteger1(2).Value) > 9999 Or Val(fpLongInteger1(3).Value) > 9999)) Then Toolbar1.Buttons(2).Enabled = falso Else Toolbar1.Buttons(2).Enabled = True
Case 2
    RS.Open "SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(2).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(4).Caption = "": Exit Sub
    fpayuda(4).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    If ((Val(fpLongInteger1(0).Value) < 10000 And Val(fpLongInteger1(1).Value) < 10000) And (Val(fpLongInteger1(2).Value) > 9999 Or Val(fpLongInteger1(3).Value) > 9999)) Then Toolbar1.Buttons(2).Enabled = falso Else Toolbar1.Buttons(2).Enabled = True
Case 3
    RS.Open "SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(3).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(5).Caption = "": Exit Sub
    fpayuda(5).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    If ((Val(fpLongInteger1(0).Value) < 10000 And Val(fpLongInteger1(1).Value) < 10000) And (Val(fpLongInteger1(2).Value) > 9999 Or Val(fpLongInteger1(3).Value) > 9999)) Then Toolbar1.Buttons(2).Enabled = falso Else Toolbar1.Buttons(2).Enabled = True
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    If Index = 1 Then Image1_Click 1
    If Index = 2 Then Image1_Click 2
    If Index = 4 Then Image1_Click 4
    If Index = 5 Then Image1_Click 5
End Select
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    If Index = 0 Then Image1_Click 0
    If Index = 3 Then Image1_Click 3
End Select
End Sub

Private Sub fpText_LostFocus(Index As Integer)
Select Case Index
Case 0
    If fpText(0).text = "" Then fpayuda(0).Caption = "": Exit Sub
    RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & fpText(0).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": fpLongInteger1(0).Value = "": fpayuda(1).Caption = "": fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!cli_nombre)
    RS.Close: Set RS = Nothing
    fpLongInteger1(0).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(1).Value = "": fpayuda(2).Caption = ""
Case 1
    If fpText(1).text = "" Then fpayuda(3).Caption = "": Exit Sub
    RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & fpText(1).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(3).Caption = "": fpLongInteger1(2).Value = "": fpayuda(4).Caption = "": fpLongInteger1(3).Value = "": fpayuda(5).Caption = "": Exit Sub
    fpayuda(3).Caption = Trim(RS!cli_nombre)
    RS.Close: Set RS = Nothing
    fpLongInteger1(2).Value = "": fpayuda(4).Caption = ""
    fpLongInteger1(3).Value = "": fpayuda(5).Caption = ""
End Select
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(0).text = vg_codigo
    fpayuda(0).Caption = vg_nombre
    fpLongInteger1(0).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(1).Value = "": fpayuda(2).Caption = ""
    fpLongInteger1(0).SetFocus
Case 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    fpLongInteger1(1).SetFocus
Case 2
    vg_left = fpayuda(2).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre
    fpDateTime1(0).SetFocus
Case 3
    vg_left = fpayuda(3).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(1).text = vg_codigo
    fpayuda(3).Caption = vg_nombre
    fpLongInteger1(2).Value = "": fpayuda(4).Caption = ""
    fpLongInteger1(3).Value = "": fpayuda(5).Caption = ""
    fpLongInteger1(2).SetFocus
Case 4
    vg_left = fpayuda(4).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(4).Caption = vg_nombre
    fpLongInteger1(3).SetFocus
Case 5
    vg_left = fpayuda(5).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(3).Value = Val(vg_codigo)
    fpayuda(5).Caption = vg_nombre
    fpDateTime1(2).SetFocus
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim fecori1 As Long, fecori2 As Long, fecdes1 As Long, fecdes2 As Long, vdia As Long, indice As Long, tiprec As Long
Dim auxfeco As String, auxfecd As String, vaux1 As Long, vaux2 As Long, diatop As Long, est As Boolean, enumrac As Long, sql1 As String, sql2 As String
On Error GoTo Man_Error
Select Case Button.Index
Case 2
    '------- Validar datos origen
    RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText(0).text = "": fpayuda(0).Caption = "": fpLongInteger1(0).Value = "": fpayuda(1).Caption = "": fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    RS.Open "SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(0).Value = "": fpayuda(1).Caption = "": MsgBox "No Existe Regimen", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    RS.Open "SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(1).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set ConSql = Nothing: fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": MsgBox "No Existe Servicio", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    If fpText(0).text = fpText(1).text And Val(fpLongInteger1(0).Value) = Val(fpLongInteger1(2).Value) And Val(fpLongInteger1(1).Value) = Val(fpLongInteger1(3).Value) And fpDateTime1(0).text = fpDateTime1(2).text And fpDateTime1(1).text = fpDateTime1(3).text Then: MsgBox "Datos origen, beben ser distinto datos destino", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    If fpDateTime1(0).text = "" Or fpDateTime1(1).text = "" Or fpDateTime1(2).text = "" Or fpDateTime1(3).text = "" Then MsgBox "Fecha no definida", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If (Val(Mid(fpDateTime1(1).text, 1, 2)) - Val(Mid(fpDateTime1(0).text, 1, 2))) > (Val(Mid(fpDateTime1(3).text, 1, 2)) - Val(Mid(fpDateTime1(2).text, 1, 2))) Then MsgBox "Fecha origen supera nº dìas", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Val(Format(fpDateTime1(0).text, "ddmmyyyy")) > Val(Format(fpDateTime1(1).text, "ddmmyyyy")) Or Val(Format(fpDateTime1(2).text, "ddmmyyyy")) > Val(Format(fpDateTime1(3).text, "ddmmyyyy")) Then MsgBox "Fecha no coincide", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Val(Format(fpDateTime1(0).text, "ddmmyyyy")) > Val(Format(fpDateTime1(1).text, "ddmmyyyy")) Or Val(Format(fpDateTime1(2).text, "ddmmyyyy")) > Val(Format(fpDateTime1(3).text, "ddmmyyyy")) Then MsgBox "Fecha no coincide", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Val(Format(fpDateTime1(0).text, "mm")) <> Val(Format(fpDateTime1(1).text, "mm")) Or Val(Format(fpDateTime1(2).text, "mm")) <> Val(Format(fpDateTime1(3).text, "mm")) Then MsgBox "Fecha no coincide", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If (Val(Mid(fpDateTime1(1).text, 1, 2)) - Val(Mid(fpDateTime1(0).text, 1, 2))) > (Val(Mid(fpDateTime1(3).text, 1, 2)) - Val(Mid(fpDateTime1(2).text, 1, 2))) Then MsgBox "Fecha origen supera nº dìas", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    '------- Validar datos destino
    RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText(1).text)) & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText(1).text = "": fpayuda(3).Caption = "": fpLongInteger1(2).Value = "": fpayuda(4).Caption = "": fpLongInteger1(3).Value = "": fpayuda(5).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    RS.Open "SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(2).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(2).Value = "": fpayuda(4).Caption = "": MsgBox "No Existe Regimen", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    RS.Open "SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(3).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set ConSql = Nothing: fpLongInteger1(3).Value = "": fpayuda(5).Caption = "": MsgBox "No Existe Servicio", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '------- Validar datos destino bloqueado
    fecdes1 = Format(fpDateTime1(2).text, "yyyymm")
    sql1 = IIf(vg_tipbase = "1", " val(mid(a.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) ")
    If vg_tipmin Then
       '-------> Validar si existe una planificación teorica el mes a copiar
        RS.Open "SELECT COUNT(a.min_codigo) AS nreg " & _
                "FROM  b_minuta a, b_minutadet b " & _
                "WHERE a.min_codigo = b.mid_codigo " & _
                "AND   a.min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
                "AND   a.min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
                "AND   a.min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
                "AND   " & sql1 & " = " & fecdes1 & " " & _
                "AND   a.min_indblo IN (0) " & _
                "AND   b.mid_tipmin = '1'", vg_db, adOpenStatic
        If Not RS.EOF And RS!nreg > 0 Then RS.Close: Set RS = Nothing: MsgBox "No es posible copiar, ya que la minuta corresponde minuta normal, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
        RS.Close: Set RS = Nothing
    End If
    sql2 = IIf(vg_tipmin, 2, 1)
    If vg_tipmin Then
       RS.Open "SELECT COUNT(a.min_codigo) AS nreg " & _
               "FROM  b_minuta a, b_minutadet b " & _
               "WHERE a.min_codigo = b.mid_codigo " & _
               "AND   a.min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
               "AND   a.min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
               "AND   a.min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
               "AND   " & sql1 & " = " & fecdes1 & " " & _
               "AND   a.min_indblo In (2,1) " & _
               "AND   b.mid_tipmin = '1'", vg_db, adOpenStatic
    Else
'       RS.Open "SELECT COUNT(a.min_codigo) AS nreg " & _
'               "FROM  b_minuta a, b_minutadet b " & _
'               "WHERE a.min_codigo = b.mid_codigo " & _
'               "AND   a.min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
'               "AND   " & sql1 & " = " & fecdes1 & " " & _
'               "AND   a.min_indblo In (2,1) " & _
'               "AND   b.mid_tipmin = '1'", vg_db, adOpenStatic
       
       RS.Open "SELECT COUNT(a.min_codigo) AS nreg " & _
               "FROM  b_minuta a, b_minutadet b " & _
               "WHERE a.min_codigo = b.mid_codigo " & _
               "AND   a.min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
               "AND   a.min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
               "AND   a.min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
               "AND   " & sql1 & " = " & fecdes1 & " " & _
               "AND   a.min_indblo In (2,1) " & _
               "AND   b.mid_tipmin = '1'", vg_db, adOpenStatic
    End If
    If Not RS.EOF And RS!nreg > 0 Then RS.Close: Set RS = Nothing: MsgBox "Minuta esta bloqueda, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    sql2 = IIf(vg_tipmin, 2, 1)
    RS.Open "SELECT COUNT(a.min_codigo) AS nreg " & _
            "FROM  b_minuta a, b_minutadet b " & _
            "WHERE a.min_codigo = b.mid_codigo " & _
            "AND   a.min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
            "AND   a.min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
            "AND   a.min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
            "AND   " & sql1 & " = " & fecdes1 & " " & _
            "AND   a.min_indblo IN ( 2,1) " & _
            "AND   b.mid_tipmin = '1'", vg_db, adOpenStatic
    If Not RS.EOF And RS!nreg > 0 Then RS.Close: Set RS = Nothing: MsgBox "Minuta esta bloqueda, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '------- Validar si existe datos origen
    fecori1 = Mid(fpDateTime1(0).text, 7, 4) & Mid(fpDateTime1(0).text, 4, 2)
    sql1 = IIf(vg_tipbase = "1", " val(mid(min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),min_fecmin),1,6)) ")
    RS.Open "SELECT DISTINCT min_cencos, min_codreg, min_codser " & _
            "FROM  b_minuta " & _
            "WHERE min_cencos = '" & LimpiaDato(Trim(fpText(0).text)) & "' " & _
            "AND   min_codreg = " & Val(fpLongInteger1(0).Value) & " " & _
            "AND   min_codser = " & Val(fpLongInteger1(1).Value) & " " & _
            "AND   " & sql1 & " = " & fecori1 & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe datos origen, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '------- Grabando plantilla contrato origen hacia origen
    vdia = 999999: indice = 0
    fecori1 = Format(fpDateTime1(0).text, "yyyymmdd")
    fecori2 = Format(fpDateTime1(1).text, "yyyymmdd")
    fecdes1 = Format(fpDateTime1(2).text, "yyyymmdd")
    fecdes2 = Format(fpDateTime1(3).text, "yyyymmdd")
    diatop = Format(fpDateTime1(3).text, "yyyymmdd")
    '------- validar si Existe Datos Destino
    RS.Open "SELECT DISTINCT min_cencos, min_codreg, min_codser " & _
             "FROM  b_minuta " & _
             "WHERE min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
             "AND   min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
             "AND   min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
             "AND   min_fecmin >= " & fecdes1 & " and min_fecmin <= " & fecdes2 & "", vg_db, adOpenStatic
    If Not RS.EOF Then If MsgBox("Existe información contrato destino. se borrara la información existente ...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then RS.Close: Set RS = Nothing: Exit Sub
    RS.Close: Set RS = Nothing
    '------- Fin validar si existe datos destino
    RS.Open "SELECT a.*, b.* " & _
            "FROM  b_minuta a, b_minutadet b, b_receta c " & _
            "WHERE a.min_codigo = b.mid_codigo " & _
            "AND   b.mid_codrec = c.rec_codigo " & _
            "AND  (c.rec_fecvig > " & Format(Date, "yyyymmdd") & " OR c.rec_fecvig <= 0 OR (c.rec_fecvig) IS NULL) " & _
            "AND   a.min_cencos = '" & LimpiaDato(Trim(fpText(0).text)) & "' " & _
            "AND   a.min_codreg = " & Val(fpLongInteger1(0).Value) & " " & _
            "AND   a.min_codser = " & Val(fpLongInteger1(1).Value) & " " & _
            "AND   a.min_fecmin >= " & fecori1 & " " & _
            "AND   a.min_fecmin <= " & fecori2 & " " & _
            "AND   b.mid_tipmin = '1' ORDER BY a.min_fecmin", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
    fg_carga ""
    est = True
    If DatePart("w", fg_Ctod1(RS!min_fecmin), 2) <> DatePart("w", fg_Ctod1(fecdes1), 2) Then
       If MsgBox("No coincide día de la semana. ¿ Desea copiar ? ...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then RS.Close: Set RS = Nothing: fg_descarga: Exit Sub Else est = False
    End If
    vg_db.BeginTrans
    Do While Not RS.EOF
       If RS!min_fecmin <> vdia Then
          If Not est And fecdes1 > diatop Then GoTo paso
          If est Then
             auxfeco = Mid(RS!min_fecmin, 7, 2) & "/" & Mid(RS!min_fecmin, 5, 2) & "/" & Mid(RS!min_fecmin, 1, 4)
             auxfecd = Mid(fecdes1, 7, 2) & "/" & Mid(fecdes1, 5, 2) & "/" & Mid(fecdes1, 1, 4)
     
             vaux1 = DatePart("w", auxfeco, 2)
             vaux2 = DatePart("w", auxfecd, 2)
       
             Do While (vaux1 <> vaux2)
                fecdes1 = (fecdes1 + 1)
                If fecdes1 > diatop Then GoTo paso
                auxfecd = Mid(fecdes1, 7, 2) & "/" & Mid(fecdes1, 5, 2) & "/" & Mid(fecdes1, 1, 4)
                vaux2 = DatePart("w", auxfecd, 2)
             Loop
          End If
          indice = 0
          '------- actualizar nro. raciones totales
          enumrac = 0
          If Option1(1).Value = True Then
             RS1.Open "SELECT sra_serdia, SUM(sra_raciones) AS raciones FROM a_serviciorac WHERE sra_cencos = '" & MuestraCasino(1) & "' AND sra_codser = " & Val(fpLongInteger1(3).Value) & " AND sra_serdia = " & IIf(DatePart("w", fg_Ctod1(fecdes1), 2) = 1, 7, DatePart("w", fg_Ctod1(fecdes1), 2) - 1) & "  GROUP BY sra_serdia", vg_db, adOpenStatic
             If Not RS1.EOF Then enumrac = RS1!raciones
             RS1.Close: Set RS1 = Nothing
          End If
          sql2 = IIf(vg_tipmin, 11, 0)
          RS1.Open "SELECT min_codigo FROM b_minuta " & _
                   "WHERE min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
                   "AND   min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
                   "AND   min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
                   "AND   min_fecmin = " & fecdes1 & " " & _
                   "AND   min_indblo = " & sql2 & "", vg_db, adOpenStatic
          If Not RS1.EOF Then
             indice = RS1!min_codigo
             RS1.Close: Set RS1 = Nothing
'             vg_db.Execute "DELETE b_minuta FROM b_minuta " & _
'                           "WHERE mid_cencos='" & vg_codcasino & "' " & _
'                           "AND   mid_codreg=" & vg_codregimen & " " & _
'                           "AND   mid_codser=" & vg_codservicio & " " & _
'                           "AND   min_codigo=" & indiceminutas & " " & _
'                           "AND   min_fecmin=" & Val(wsfecha) & ""
             vg_db.Execute "DELETE b_minutadet FROM b_minutadet WHERE mid_codigo = " & indice & " AND mid_tipmin = '1'"
             vg_db.Execute "UPDATE b_minuta SET min_racteo = " & IIf(Option1(0).Value = True, IIf(IsNull(RS!min_racteo), 0, RS!min_racteo), enumrac) & " WHERE min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' AND min_codreg = " & Val(fpLongInteger1(2).Value) & " AND min_codser = " & Val(fpLongInteger1(3).Value) & " AND min_fecmin = " & fecdes1 & " AND min_codigo = " & indice & ""
          Else
             RS1.Close: Set RS1 = Nothing
             RS1.Open "SELECT min_codigo FROM b_minuta ORDER BY min_codigo DESC", vg_db, adOpenStatic
             If Not RS1.EOF Then RS1.MoveFirst: indice = RS1!min_codigo + 1 Else indice = 1
             RS1.Close: Set RS1 = Nothing
             sql2 = IIf(vg_tipmin, 11, 0)
             vg_db.Execute "INSERT INTO b_minuta (min_codigo, min_cencos, min_codreg, min_codser, min_fecmin, min_indblo, min_racteo, min_racrea) " & _
                           "VALUES (" & indice & ", '" & LimpiaDato(Trim(fpText(1).text)) & "', " & Val(fpLongInteger1(2).Value) & ", " & _
                           "" & Val(fpLongInteger1(3).Value) & ", " & fecdes1 & ", " & sql2 & ", " & IIf(Option1(0).Value = True, IIf(IsNull(RS!min_racteo), 0, RS!min_racteo), enumrac) & ", 0)"
          End If
          vdia = RS!min_fecmin
          If Not est Then
             fecdes1 = fecdes1 + 1
          End If
       End If
       '------- Traer tipo receta
       tiprec = 0
       RS1.Open "SELECT DISTINCT red_tiprec FROM b_recetadet WHERE red_codigo = " & RS!mid_codrec & " AND ((red_tiprec <> 0 AND red_cencos = '" & MuestraCasino(1) & "') OR (red_tiprec = 0 AND red_cencos = '0')) ORDER BY red_tiprec", vg_db, adOpenStatic
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             If RS1!red_tiprec = -1 Then
                tiprec = IIf((fpLongInteger1(2).Value) < 10000, RS1!red_tiprec, 0)
             ElseIf RS1!red_tiprec = Val(fpLongInteger1(2).Value) And RS1!red_tiprec = RS!mid_tiprec Then
                tiprec = RS1!red_tiprec
                Exit Do
             ElseIf RS1!red_tiprec <> Val(fpLongInteger1(2).Value) And RS1!red_tiprec = RS!mid_tiprec Then
                tiprec = RS!mid_tiprec
                Exit Do
             End If
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       RS1.Open "SELECT * FROM b_minutadet WHERE mid_codigo = " & indice & " AND mid_tipmin = '1' AND mid_numlin = " & RS!mid_numlin & "", vg_db, adOpenStatic
       If RS1.EOF Then
          vg_db.Execute "INSERT INTO b_minutadet (mid_codigo, mid_tipmin, mid_numlin, mid_estser, mid_codrec, mid_numrac, mid_descri, mid_cosrec, mid_tiprec, mid_nummer, mid_rec5eta, mid_cosdes, mid_modmina, mid_modminb) " & _
                        "VALUES (" & indice & ", '1', " & RS!mid_numlin & ", " & RS!mid_estser & ", " & RS!mid_codrec & ", " & IIf(Option1(0).Value = True, IIf(IsNull(RS!mid_numrac), 0, RS!mid_numrac), 0) & ", '" & RS!mid_descri & "', " & fg_CalCtoRecInv(RS!mid_codrec, tiprec, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','"))) & ", " & tiprec & ", 0, '" & IIf((fpLongInteger1(2).Value) < 10000, 0, RS!mid_rec5eta) & "', " & fg_CalCtoRecInv(RS!mid_codrec, tiprec, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))) & ", '0', '0')"
       Else
          vg_db.Execute "UPDATE b_minutadet SET mid_modmina= '0', mid_modminb= '0', mid_estser = " & RS!mid_estser & ", mid_codrec = " & RS!mid_codrec & ", mid_numrac = " & RS!mid_numrac & ", mid_descri = '" & RS!mid_descri & "', mid_cosrec = " & fg_CalCtoRecInv(RS!mid_codrec, tiprec, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','"))) & ", mid_tiprec=" & tiprec & ", mid_rec5eta='" & IIf((fpLongInteger1(2).Value) < 10000, 0, 1) & "', mid_cosdes=" & fg_CalCtoRecInv(RS!mid_codrec, tiprec, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))) & " WHERE mid_codigo=" & indice & " AND mid_tipmin='1' AND mid_numlin=" & RS!mid_numlin & ""
       End If
       RS1.Close: Set RS1 = Nothing
       RS.MoveNext
    Loop
paso:
    RS.Close: Set RS = Nothing
    vg_db.CommitTrans
    fg_descarga
    Picture1.Visible = False: Label1(5).Visible = False: gauge.Visible = False
    MsgBox "Copia Finalizada Sin Problema", vbInformation + vbOKOnly, Msgtitulo
Case 4
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub
