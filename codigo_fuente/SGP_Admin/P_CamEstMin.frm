VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form P_CamEstMin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Estado Minuta Bloque"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   10560
      TabIndex        =   6
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   8400
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   12375
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1995
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opcional"
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
         Left            =   10680
         TabIndex        =   20
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opcional"
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
         Left            =   10680
         TabIndex        =   19
         Top             =   840
         Width           =   765
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   3270
         Picture         =   "P_CamEstMin.frx":0000
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3270
         Picture         =   "P_CamEstMin.frx":030A
         Top             =   660
         Width           =   480
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
         TabIndex        =   15
         Top             =   1665
         Width           =   1065
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
         TabIndex        =   14
         Top             =   1155
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
         TabIndex        =   13
         Top             =   735
         Width           =   6735
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
         TabIndex        =   12
         Top             =   1665
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
         Index           =   3
         Left            =   705
         TabIndex        =   11
         Top             =   1260
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
         Index           =   2
         Left            =   705
         TabIndex        =   10
         Top             =   840
         Width           =   750
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3285
         Picture         =   "P_CamEstMin.frx":0614
         Top             =   240
         Width           =   480
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
         Left            =   3735
         TabIndex        =   9
         Top             =   315
         Width           =   6735
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
         Left            =   720
         TabIndex        =   8
         Top             =   420
         Width           =   735
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
         Left            =   3780
         TabIndex        =   18
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
         TabIndex        =   16
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
         TabIndex        =   17
         Top             =   1200
         Width           =   6735
      End
   End
End
Attribute VB_Name = "P_CamEstMin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public lc_Aux As String

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String

Select Case Index
Case 0
               
    If ValidaDatos = False Then Exit Sub
           
    '--> Validar si existe minuta
    Sql = ""
    Sql = " '" & LimpiaDato(Trim(fpText.text)) & "', "
    Sql = Sql & " " & IIf(fpLongInteger1(0).Value = "", 0, fpLongInteger1(0).Value) & ", "
    Sql = Sql & " " & IIf(fpLongInteger1(1).Value = "", 0, fpLongInteger1(1).Value) & " , "
    Sql = Sql & " " & Format(FpFecDesde.text, "yyyymmdd") & " , "
    Sql = Sql & " " & Format(FpFecHasta.text, "yyyymmdd") & ""
    
    Set RS = vg_db.Execute("sgpadm_Sel_ValidarMinutaBloques " & Sql & "")
    If RS.EOF Then
       
       MsgBox "No existe minuta modificar estado...", vbExclamation + vbOKOnly, Me.Caption
       RS.Close
       Set RS = Nothing
       Exit Sub
    
    End If
    RS.Close
    Set RS = Nothing

    If MsgBox("Esta seguro realizar cambio...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    
    '--> Actualizar cambio de estado minuta
    Sql = ""
    Sql = " '" & LimpiaDato(Trim(fpText.text)) & "', "
    Sql = Sql & " " & IIf(fpLongInteger1(0).Value = "", 0, fpLongInteger1(0).Value) & ", "
    Sql = Sql & " " & IIf(fpLongInteger1(1).Value = "", 0, fpLongInteger1(1).Value) & " , "
    Sql = Sql & " " & Format(FpFecDesde.text, "yyyymmdd") & " , "
    Sql = Sql & " " & Format(FpFecHasta.text, "yyyymmdd") & ", "
    Sql = Sql & " '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "' "
    
    If lc_Aux = "CamEstMin" Then
       
       Set RS = vg_db.Execute("sgpadm_Upd_DesbloquearMinutaBloque " & Sql & "")
    
    Else
       
       Set RS = vg_db.Execute("sgpadm_Upd_PorcentajePonderacionMinutaBloque " & Sql & "")
    
    End If
    
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, Me.Caption
       
       Else
          
          MsgBox "Proceso finalizado sin problema", vbInformation + vbOKOnly, Me.Caption
       
       End If
    
    End If
    RS.Close
    Set RS = Nothing

Case 1

    Unload Me
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Private Function ValidaDatos() As Boolean

On Error GoTo Man_Error

Dim RS      As New ADODB.Recordset
Dim Dias    As Long
Dim i       As Long
Dim Fecha   As String
Dim mes     As String
Dim Año     As String

Let ValidaDatos = True
 
'-------> Validar Ceco
Set RS = vg_db.Execute("sgpadm_s_cliente_V02 45, '" & LimpiaDato(fpText.text) & "', ''")
If RS.EOF Then
   RS.Close
   Set RS = Nothing
   fpayuda(0).Caption = ""
   MsgBox "No existe Contrato...", vbExclamation + vbOKOnly, MsgTitulo
   Let ValidaDatos = False
   Exit Function
End If
RS.Close
Set RS = Nothing

'-------> Validar Regimen
If Trim(fpLongInteger1(0).text) <> "" Then
'    Call MsgBox("Debe Ingresar Regimen", vbInformation, Me.Caption)
    Call fpLongInteger1(0).SetFocus
    Let ValidaDatos = True
    Exit Function
End If

'-------> Validar Servicio
If Trim(fpLongInteger1(1).text) <> "" Then
'    Call MsgBox("Debe Ingresar Servicio", vbInformation, Me.Caption)
    Call fpLongInteger1(1).SetFocus
    Let ValidaDatos = True
    Exit Function
End If

'-------> Validar Fecha Desde & Hasta
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

If DateDiff("m", FpFecDesde.text, FpFecHasta.text) > 3 Then
    Call MsgBox("Rango De Fecha No Puede Ser Mayor a 3 Meses", vbInformation, Me.Caption)
    Let ValidaDatos = False
    Exit Function
End If
    
Exit Function
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  

End Function

Private Sub Form_Activate()
    
    Call fg_descarga

End Sub

Private Sub Form_Load()
    
On Error GoTo Man_Error
    
Me.Caption = IIf(lc_Aux = "CamEstMin", "Estado Minuta Bloque", "% Ponderación")
    
Call fg_carga("")
Me.HelpContextID = vg_OpcM
Call fg_centra(Me)

Let FpFecDesde.text = Format(Date, "dd/mm/yyyy")
Let FpFecHasta.text = Format(Date, "dd/mm/yyyy")
   
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)
    
On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

If IsDate(FpFecHasta.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Select Case Index
    Case 0
        Set RS = vg_db.Execute("sgpadm_Sel_RegimenBloque " & IIf(Val(fpLongInteger1(0).Value) = 0, -1, Val(fpLongInteger1(0).Value)) & "")
        If RS.EOF Then
            RS.Close
            Set RS = Nothing
            fpayuda(1).Caption = ""
            Exit Sub
        End If
        fpayuda(1).Caption = Trim(RS!reg_nombre)
        RS.Close: Set RS = Nothing
       
    Case 1
        Set RS = vg_db.Execute("sgpadm_Sel_ServicioBloque " & IIf(Val(fpLongInteger1(1).Value) = 0, -1, Val(fpLongInteger1(1).Value)) & "")
        If RS.EOF Then
            RS.Close
            Set RS = Nothing
            fpayuda(2).Caption = ""
            Exit Sub
        End If
        fpayuda(2).Caption = Trim(RS!ser_nombre)
        RS.Close: Set RS = Nothing
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
    
On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub fpText_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

    Set RS = vg_db.Execute("sgpadm_s_cliente_V02 45, '" & LimpiaDato(fpText.text) & "', ''")
    If RS.EOF Then
        RS.Close
        Set RS = Nothing
        fpayuda(0).Caption = ""
        fpLongInteger1(0).Value = ""
        fpayuda(1).Caption = ""
        fpLongInteger1(1).Value = ""
        fpayuda(2).Caption = ""
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

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub


Private Sub fpText_KeyPress(KeyAscii As Integer)
    
On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

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
    Call B_TabEst.LlenaDatos("a_regimen", "", "Regimen", "RegBlo")
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
    Call B_TabEst.LlenaDatos("a_servicio", "", "Servicio", "SerBlo")
    Call B_TabEst.Show(1)
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpLongInteger1(1).SetFocus
    fpayuda(2).Caption = vg_nombre
    FpFecDesde.Enabled = True
    FpFecHasta.Enabled = True
    
    Call FpFecDesde.SetFocus
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub
