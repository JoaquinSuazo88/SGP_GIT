VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form m_copia_min_seg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Planificación"
   ClientHeight    =   8865
   ClientLeft      =   2910
   ClientTop       =   1230
   ClientWidth     =   9180
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8865
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos Origen"
      ForeColor       =   &H80000008&
      Height          =   2745
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   8250
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   3810
         TabIndex        =   43
         Top             =   120
         Width           =   3435
         Begin VB.OptionButton Option1 
            Caption         =   "Copia Varios Servicios"
            Height          =   435
            Index           =   1
            Left            =   1770
            TabIndex        =   45
            Top             =   240
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Copia Solo un Servicio"
            Height          =   555
            Index           =   0
            Left            =   150
            TabIndex        =   44
            Top             =   180
            Value           =   -1  'True
            Width           =   1125
         End
      End
      Begin VB.ComboBox cboTipoMinuta 
         Height          =   315
         ItemData        =   "m_copia_min_seg.frx":0000
         Left            =   1320
         List            =   "m_copia_min_seg.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   250
         Width           =   2295
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   4
         Top             =   2370
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1395
         TabIndex        =   2
         Top             =   1485
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
         Index           =   2
         Left            =   1395
         TabIndex        =   3
         Top             =   1905
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
         Left            =   1395
         TabIndex        =   1
         Top             =   1065
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
      Begin EditLib.fpLongInteger txtDias 
         Height          =   315
         Index           =   4
         Left            =   3915
         TabIndex        =   40
         Top             =   2310
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   5835
         TabIndex        =   5
         Top             =   2370
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   42
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dias"
         Height          =   195
         Index           =   7
         Left            =   3360
         TabIndex        =   41
         Top             =   2415
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Minuta"
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   39
         Top             =   330
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[Mensaje]"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   1170
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   34
         Top             =   1590
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   33
         Top             =   2010
         Width           =   705
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
         Left            =   3135
         TabIndex        =   13
         Top             =   1065
         Width           =   4935
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
         Left            =   3135
         TabIndex        =   14
         Top             =   1485
         Width           =   4935
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
         Left            =   3135
         TabIndex        =   15
         Top             =   1905
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2685
         Picture         =   "m_copia_min_seg.frx":0004
         Top             =   990
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2685
         Picture         =   "m_copia_min_seg.frx":030E
         Top             =   1410
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2685
         Picture         =   "m_copia_min_seg.frx":0618
         Top             =   1830
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   32
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   8
         Left            =   2880
         TabIndex        =   31
         Top             =   2400
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos Destino"
      ForeColor       =   &H80000008&
      Height          =   1995
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   2835
      Width           =   8220
      Begin VB.CommandButton btnProcesarEstructuras 
         Caption         =   "Procesar Estructuras"
         Height          =   255
         Left            =   6180
         TabIndex        =   11
         Top             =   1650
         Width           =   1935
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   2
         Left            =   1395
         TabIndex        =   9
         Top             =   1620
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
         Left            =   1395
         TabIndex        =   19
         Top             =   1980
         Visible         =   0   'False
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1395
         TabIndex        =   7
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
         Index           =   3
         Left            =   1395
         TabIndex        =   8
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1395
         TabIndex        =   6
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   4
         Left            =   5130
         TabIndex        =   10
         Top             =   1620
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Largo de Días"
         Height          =   195
         Index           =   8
         Left            =   3810
         TabIndex        =   46
         Top             =   1700
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contrato"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   38
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   37
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   36
         Top             =   1260
         Width           =   705
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
         Index           =   5
         Left            =   3135
         TabIndex        =   16
         Top             =   315
         Width           =   4935
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
         Index           =   4
         Left            =   3135
         TabIndex        =   17
         Top             =   735
         Width           =   4935
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
         Index           =   3
         Left            =   3135
         TabIndex        =   18
         Top             =   1155
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   2685
         Picture         =   "m_copia_min_seg.frx":0922
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   2685
         Picture         =   "m_copia_min_seg.frx":0C2C
         Top             =   660
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2685
         Picture         =   "m_copia_min_seg.frx":0F36
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fec. Destino"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   29
         Top             =   1700
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   28
         Top             =   2010
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   10
         Left            =   2880
         TabIndex        =   27
         Top             =   1650
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   11
         Left            =   6650
         TabIndex        =   26
         Top             =   1650
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estructura Servicio Origen && Destino"
      Height          =   3585
      Left            =   150
      TabIndex        =   24
      Top             =   5040
      Width           =   8160
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2985
         Left            =   180
         TabIndex        =   12
         Top             =   360
         Width           =   7845
         _Version        =   393216
         _ExtentX        =   13838
         _ExtentY        =   5265
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   5
         MaxRows         =   10
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "m_copia_min_seg.frx":1240
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   1560
      ScaleHeight     =   975
      ScaleWidth      =   4785
      TabIndex        =   21
      Top             =   8430
      Visible         =   0   'False
      Width           =   4845
      Begin MSComctlLib.ProgressBar gauge 
         Height          =   330
         Left            =   120
         TabIndex        =   22
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   4515
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8865
      Left            =   8550
      TabIndex        =   20
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   15637
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "m_copia_min_seg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS            As New ADODB.Recordset
Dim RS1           As New ADODB.Recordset
Dim indsel        As Long
Dim vg_AuxIndppr  As String
Dim Est           As Boolean
Private MsgTitulo As String

Private Sub LimpiarControles()

On Error GoTo Man_Error

'ORIGEN
fpText = ""
fpLongInteger1(0) = ""
fpLongInteger1(1) = ""
fpLongInteger1(2) = ""
fpDateTime1(0) = Date
fpDateTime1(1) = Date + 1
'txtDias(4) = ""

'DESTINO
fpText1 = ""
fpLongInteger1(0) = ""
fpLongInteger1(3) = ""
fpDateTime1(2) = Date

'grilla
vaSpread1.MaxRows = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub btnProcesarEstructuras_Click()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim cadena As String
Dim cadenacodigoestruturadestino As String
Dim Sql As String

'VALIDACION GRAL DE CONTROLES
If ValidaControles = True Then
    
    MsgBox "Falta(n) ingresar\seleccionar valores en pantalla", vbExclamation, Me.Caption
    Exit Sub

End If

'-------> validar si existe información
Sql = ""

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = Sql & "sgpadm_Sel_ValidarMinutaBloque_SubSegmento " & IIf(Label2(0).Caption = "Contrato", 1, 2) & ", '" & Trim(LimpiaDato(fpText)) & "', " & fpLongInteger1(1).Value & ", " & IIf(Option1(0).Value = True, fpLongInteger1(2).Value, 0) & ", " & Format(fpDateTime1(0).text, "yyyymmdd") & ", " & Format(fpDateTime1(1).text, "yyyymmdd") & ""

Set RS = vg_db.Execute(Sql)
If RS.EOF Then
   
   RS.Close: Set RS = Nothing
   MsgBox "No existe minuta origen", vbExclamation + vbOKOnly, Me.Caption
   Exit Sub

End If
RS.Close
Set RS = Nothing

vaSpread1.MaxRows = 0
Dim i As Integer
i = 1

If Option1(0).Value = True Then
   '********************************************************************************************************
   'genera el codigo de la col derecha
   '********************************************************************************************************

   '********************************************************************************************************
   'BLOQUE DESTINO
   Sql_MVI = " sgpadm_Sel_TraeEstructMinutaBloque_V02 "
   Sql_MVI = Sql_MVI & " 'BloqueDes'"
   Sql_MVI = Sql_MVI & " , '" & Trim(fpText1) & "'" 'cencoorigen
   Sql_MVI = Sql_MVI & " ,'' " 'codigoSubsegmento
   Sql_MVI = Sql_MVI & " , 0" 'codigoregimen
   Sql_MVI = Sql_MVI & " , " & Trim(fpLongInteger1(3)) 'codigoservicio
   Sql_MVI = Sql_MVI & " , 0" 'fechaorigenini
   Sql_MVI = Sql_MVI & " , 0" 'fechaorigenfin
   '********************************************************************************************************
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient

   Set RS = vg_db.Execute(Sql_MVI)
   cadena = ""
   cadenacodigoestruturadestino = ""
   If RS.EOF = False Then
      
      While Not RS.EOF
           
         cadena = cadena & Chr(9) & Trim(RS!ess_nombre)
         cadenacodigoestruturadestino = cadenacodigoestruturadestino & Chr(9) & RS!ess_codigo
            
         RS.MoveNext
         i = i + 1
        
      Wend
   
   End If
   RS.Close
   Set RS = Nothing

   If cadena = "" Then
      
      MsgBox "No hay estructura de servicios para esta selección", vbExclamation, Me.Caption
      Exit Sub
   
   End If

   'cadena = "Mastiff" & Chr(9) & "Sheepdog" & Chr(9) & "Terrier" & Chr(9) & "Spaniel" & Chr(9) & "Pointer" & Chr(9) & "Coonhound"

   '********************************************************************************************************
   'genera el codigo de la col izq.
   '********************************************************************************************************
End If
    
'********************************************************************************************************
'BLOQUE ORIGEN
'********************************************************************************************************
Sql_MVI = ""
If Option1(0).Value = True Then
   
   If cboTipoMinuta = "Bloque" Then
    
      Sql_MVI = " sgpadm_Sel_TraeEstructMinutaBloque_V02 "
      Sql_MVI = Sql_MVI & " 'BloqueOri'"
      Sql_MVI = Sql_MVI & " , '" & Trim(fpText) & "'" 'cencoorigen
      Sql_MVI = Sql_MVI & " ,'' " 'codigoSubsegmento
      Sql_MVI = Sql_MVI & " , " & Trim(fpLongInteger1(1)) 'codigoregimen
      Sql_MVI = Sql_MVI & " , " & Trim(fpLongInteger1(2)) 'codigoservicio
      Sql_MVI = Sql_MVI & " , '" & Format(fpDateTime1(0), "yyyymmdd") & "'" 'fechaorigenini
      Sql_MVI = Sql_MVI & " , '" & Format(fpDateTime1(1), "yyyymmdd") & "'" 'fechaorigenfin
    
   Else
    
      Sql_MVI = " sgpadm_Sel_TraeEstructMinutaBloque_V02 "
      Sql_MVI = Sql_MVI & " 'SegOri'"
      Sql_MVI = Sql_MVI & " , ''" 'cencoorigen
      Sql_MVI = Sql_MVI & " ,'" & Trim(fpText) & "'" 'codigoSubsegmento
      Sql_MVI = Sql_MVI & " , " & Trim(fpLongInteger1(1)) 'codigoregimen
      Sql_MVI = Sql_MVI & " , " & Trim(fpLongInteger1(2)) 'codigoservicio
      Sql_MVI = Sql_MVI & " , '" & Format(fpDateTime1(0), "yyyymmdd") & "'" 'fechaorigenini
      Sql_MVI = Sql_MVI & " , '" & Format(fpDateTime1(1), "yyyymmdd") & "'" 'fechaorigenfin

   End If

Else
   
   Sql_MVI = " sgpadm_Sel_TraeServicioMinutaBloque_SubSegmento_V02 "
   Sql_MVI = Sql_MVI & " '" & IIf(cboTipoMinuta = "Bloque", "BloqueOri", "SegOri") & "' "
   Sql_MVI = Sql_MVI & " ,'" & Trim(fpText) & "'" 'Ceco o bien Subsegmento
   Sql_MVI = Sql_MVI & " , " & Trim(fpLongInteger1(1)) 'codigoregimen
   Sql_MVI = Sql_MVI & " , " & Format(fpDateTime1(0), "yyyymmdd") & "" 'fechaorigenini
   Sql_MVI = Sql_MVI & " , " & Format(fpDateTime1(1), "yyyymmdd") & "" 'fechaorigenfin
   Sql_MVI = Sql_MVI & " , '" & LimpiaDato(Trim(fpText1)) & "'" 'Ceco Origen
   Sql_MVI = Sql_MVI & " , " & fpLongInteger1(0) & "" 'Regimen Origen
   Sql_MVI = Sql_MVI & " , " & Format(fpDateTime1(2), "yyyymmdd") & "" 'Fecha Origen

End If
'********************************************************************************************************
    
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql_MVI)

If RS.EOF = False Then
    
    While Not RS.EOF
        
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 2
        vaSpread1.text = CStr(RS(0))
        
        If Option1(0).Value = True Then
           
           vaSpread1.Col = 3
           vaSpread1.text = CStr(RS(1))
           
           vaSpread1.Col = 4
           vaSpread1.TypeComboBoxList = cadenacodigoestruturadestino
        
           vaSpread1.Col = 5
           vaSpread1.TypeComboBoxList = cadena
           vaSpread1.Col = 4
        
           '-------> coloca el mismo codigo al lado derecho e izq., siempre y cuando sean el mismo servicio origen y destino
           If fpLongInteger1(2) = fpLongInteger1(3) Then
              vaSpread1.Col = 4
              
              For i = 0 To vaSpread1.TypeComboBoxCount
                  
                  vaSpread1.TypeComboBoxCurSel = i
                 
                 If RS!ess_codigo = Val(vaSpread1.text) Then
                    
                    vaSpread1.Col = 5
                    vaSpread1.TypeComboBoxCurSel = i
                    Exit For
                 
                 End If
              
              Next i
           
           End If
        
        Else
           
           vaSpread1.Col = 3
           vaSpread1.text = CStr(RS(0)) & " - " & CStr(RS(1))
           
           vaSpread1.Col = 5
           vaSpread1.text = ""
'           vaSpread1.ForeColor = &HFFFF&
           If RS(2) > 0 Then
              
              vaSpread1.text = IIf(RS(2) = 0, "", "Minuta Esta Bloqueada")
              vaSpread1.ForeColor = &HFF&
           
           End If
           
           If IsNull(RS(2)) Or RS(2) = 0 Then
              
              '-------> Validar diías dentro de una minuta
              Sql = ""
              Sql = Trim(LimpiaDato(fpText1.text))
              Sql = Sql & ", " & fpLongInteger1(0).Value & ", " & RS(0) & ", " & Format(fpDateTime1(2).text, "yyyymmdd") & ", " & Val(fpLongInteger1(4).text) & ""
              
              If RS1.State = 1 Then RS1.Close
              RS1.CursorLocation = adUseClient
              vg_db.CursorLocation = adUseClient
              
              Set RS1 = vg_db.Execute("sgpadm_Sel_ValidarMinutaBloqueLargoDias " & Sql & "")
              
              If Not RS1.EOF Then
                 
                 If (RS1(0) = "S" And RS1(1) = "N") Or (RS1(0) = "N" And RS1(1) = "S") Then
                    
                    vaSpread1.Col = 5
                    vaSpread1.text = "Existen minuta bloque para este periodo"
                    vaSpread1.ForeColor = &HFF&
                 
                 ElseIf RS1(2) <> RS1(3) And RS1(2) > 0 And RS1(3) > 0 Then
                    
                    vaSpread1.Col = 5
                    vaSpread1.text = "Existen minuta bloque para este periodo"
                    vaSpread1.ForeColor = &HFF&
                 
                 End If
             
             End If
             
             RS1.Close
             Set RS1 = Nothing
           
           End If
        
        End If
        
        RS.MoveNext
        i = i + 1
    
    Wend

End If
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub cboTipoMinuta_Click()

On Error GoTo Man_Error

    SeleccTipoMinuta_MVI = cboTipoMinuta.text
    
    Call LimpiarControles
    
    If Me.cboTipoMinuta = "Bloque" Or Me.cboTipoMinuta = "" Then
        
        Label2(0).Caption = "Contrato"
    
    Else
        
        Label2(0).Caption = "Sub.Segto."
    
    End If
    
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Public Function CambiaDias_A_Fecha(ByVal Dias As Integer, ByVal Fecha As Date) As String

On Error GoTo Man_Error

    CambiaDias_A_Fecha = DateAdd("d", Dias, Fecha)

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Function

Private Sub Form_Activate()

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub LlenarcboTipoMinuta()
    
On Error GoTo Man_Error

    cboTipoMinuta.Clear
    cboTipoMinuta.AddItem ("Bloque")
    cboTipoMinuta.AddItem ("Segmento")

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

swUp_CECO = False 'MVA - MVI - USO GRAL.
swUp_REG = False 'MVA - MVI - USO GRAL.
swUp_SERV = False 'MVA - MVI - USO GRAL.

swEsCopia = True

LlenarcboTipoMinuta

Me.HelpContextID = vg_OpcM
fg_centra Me
EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)
EspFecha fpDateTime1(2)
EspFecha fpDateTime1(3)
MsgTitulo = "Copiar Planificación Bloque"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
Label1(8).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(0).text, "yyyymmdd"), 1), 1, 4)
'fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
'Label1(9).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(1).text, "yyyymmdd"), 1), 1, 4)
fpDateTime1(2).text = Format(Date, "dd/mm/yyyy")
Label1(10).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(2).text, "yyyymmdd"), 1), 1, 4)
fpDateTime1(3).text = Format(Date, "dd/mm/yyyy")
Label1(11).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(3).text, "yyyymmdd"), 1), 1, 4)
vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &HC0FFFF
vg_AuxIndppr = vg_Indppr
OpUsuario = vg_IndpprSelec

If IsNull(OpUsuario) Or Trim(OpUsuario) = "" Then
    
    'MsgBox "Contactese con el Administrador del Sistema...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
Else
    
    Me.HelpContextID = 1110010
    
    If Mid(ValidaPerfil(Me), 1, 1) = "1" Then
       
       vg_Indppr = 3
    
    Else
        
        Select Case OpUsuario
        Case "1"
        Case "2"
        End Select
    
    End If
End If

Me.HelpContextID = vg_OpcM
fpDateTime1(1) = Date + 1
fpDateTime1(3) = Date + 1

'Call cboTipoMinuta_Click

Me.cboTipoMinuta.text = Me.cboTipoMinuta.List(0)

Est = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Man_Error

vg_Indppr = vg_AuxIndppr

SeleccTipoMinuta_MVI = ""

swEsCopia = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpDateTime1_Change(Index As Integer)

On Error GoTo Man_Error

If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
Exit Sub

Select Case Index

Case 0
    
    MoverVector
    Label1(8).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(0).text, "yyyymmdd"), 1), 1, 4)

Case 1
    
    'Label1(9).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(1).text, "yyyymmdd"), 1), 1, 4)

Case 2
    
    Label1(10).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(2).text, "yyyymmdd"), 1), 1, 4)

Case 3
    
    Label1(11).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(3).text, "yyyymmdd"), 1), 1, 4)

End Select
vaSpread1.MaxRows = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
    
On Error GoTo Man_Error

    Dim RS As New ADODB.Recordset
    Dim Sql As String
    
    If Index = 1 Then
        'REGIMEN ORIGEN
        
        swUp_REG = True   'MVA - MVI
        
        'esta seleccion es selecc. SEGMENTO (TABLA B_...)
        Sql = "select isnull(reg_codigo,0) as reg_codigo, isnull(reg_nombre,'') as reg_nombre "
        Sql = Sql & " From a_regimen with (nolock) "
        Sql = Sql & " where reg_activo = '1' and reg_indppr = '1' "
        Sql = Sql & " and reg_codigo = " & Val(fpLongInteger1(1).Value) & ""
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute(Sql)
        
        If RS.EOF Then
            
            RS.Close
            Set RS = Nothing: fpayuda(1).Caption = ""
            Call MoverDatos
            Exit Sub
        
        End If
        
        fpayuda(1).Caption = Trim(RS!reg_nombre)
        RS.Close
        Set RS = Nothing
        Call MoverDatos
        
    ElseIf Index = 0 Then
    
        swUp_REG = False 'MVA - MVI
        
        'REGIMEN DESTINO
        Sql = "select isnull(reg_codigo,0) as reg_codigo, isnull(reg_nombre,'') as reg_nombre "
        Sql = Sql & " From a_regimen with (nolock) "
        Sql = Sql & " where reg_activo = '1' and reg_indppr = '1' "
        Sql = Sql & " and reg_codigo = " & Val(fpLongInteger1(0).Value) & ""
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        Set RS = vg_db.Execute(Sql)
        
        If RS.EOF Then
            
            RS.Close
            Set RS = Nothing: fpayuda(4).Caption = ""
            Call MoverDatos
            Exit Sub
        
        End If
        fpayuda(4).Caption = Trim(RS!reg_nombre)
        RS.Close
        Set RS = Nothing
        Call MoverDatos
    
    ElseIf Index = 2 Then
    'SERVICIO ORIGEN
    
            swUp_SERV = True 'MVA - MVI
        
            Sql = "select isnull(ser_codigo,0) as ser_codigo, isnull(ser_nombre,'') as ser_nombre "
            Sql = Sql & " From a_servicio with (nolock) "
            Sql = Sql & " where ser_activo = '1' and ser_indppr = '1' "
            Sql = Sql & " and ser_codigo = " & Val(fpLongInteger1(2).Value) & ""
            Sql = Sql & " order by ser_nombre"
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute(Sql)
        
        If RS.EOF Then
            
            RS.Close
            Set RS = Nothing
            fpayuda(2).Caption = ""
            Call MoverDatos
            Exit Sub
        
        End If
        fpayuda(2).Caption = Trim(RS!ser_nombre)
        RS.Close
        Set RS = Nothing
        Call MoverDatos
    
    ElseIf Index = 3 Then
    
    'SERVICIO DESTINO
        
        swUp_SERV = False 'MVA - MVI
        Sql = "select isnull(ser_codigo,0) as ser_codigo, isnull(ser_nombre,'') as ser_nombre "
        Sql = Sql & " From a_servicio with (nolock) "
        Sql = Sql & " where ser_activo = '1' and ser_indppr = '1' "
        Sql = Sql & " and ser_codigo = " & Val(fpLongInteger1(3).Value) & ""
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute(Sql)
        
        If RS.EOF Then
            
            RS.Close
            Set RS = Nothing: fpayuda(3).Caption = ""
            Call MoverDatos
            Exit Sub
        
        End If
        
        fpayuda(3).Caption = Trim(RS!ser_nombre)
        RS.Close
        Set RS = Nothing
        Call MoverDatos
    
    End If
    
    vaSpread1.MaxRows = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 120
    
    If Index = 0 Then Image1_Click 0
    If Index = 1 Then Image1_Click 1
    If Index = 2 Then Image1_Click 2
    If Index = 3 Then Image1_Click 3
    If Index = 4 Then Image1_Click 4
    If Index = 5 Then Image1_Click 5

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub MoverDatos()
    
On Error GoTo Man_Error

    Dim indblo As Boolean, i As Long, j As Long, IndCol As Long, Sql1 As String

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpText_Change()
    
On Error GoTo Man_Error

    swUp_CECO = True 'MVA - MVI
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    If SeleccTipoMinuta_MVI = "Segmento" Then
        
        Sql = " SELECT DISTINCT isnull(a.sub_codigo,0) as sub_codigo, isnull(a.sub_nombre,'') as Cli_nombre"
        Sql = Sql & " FROM a_subsegmento a with (nolock) "
        Sql = Sql & " LEFT outer JOIN b_detlistaprecio b with (nolock) on a.sub_codigo = b.dlp_codigo "
        Sql = Sql & " WHERE sub_activo = 1 and sub_indppr = '1' "
        Sql = Sql & " AND a.sub_codigo = '" & fpText.text & "'"
    
        Set RS = vg_db.Execute(Sql)
            
    Else
    
        Set RS = vg_db.Execute("sgpadm_s_cliente_V02 45, '" & LimpiaDato(fpText.text) & "', ''")
    
    End If
    
    If RS.EOF Then
        
        RS.Close
        Set RS = Nothing
        fpayuda(0).Caption = ""
        Exit Sub
    
    End If
    
    fpayuda(0).Caption = Trim(RS!Cli_nombre)
    RS.Close
    Set RS = Nothing
    Call MoverDatos
    vaSpread1.MaxRows = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpText1_Change()

On Error GoTo Man_Error

    swUp_CECO = False 'MVA - MVI
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_s_cliente_V02 45, '" & LimpiaDato(fpText1.text) & "', ''")
    If RS.EOF Then
        
        RS.Close
        Set RS = Nothing
        fpayuda(5).Caption = ""
        Exit Sub
    
    End If
    fpayuda(5).Caption = Trim(RS!Cli_nombre)
    RS.Close
    Set RS = Nothing
    Call MoverDatos
    
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpText1_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0

   'CARGA CCOSTO o CLIENTES ORIGEN
    
    swUp_CECO = True  'MVA - MVI
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = ""
    vg_codigo = ""
    Call B_TabEst.LlenaDatos("b_clientes", "cli_", IIf(SeleccTipoMinuta_MVI = "Segmento", "Subsegmentos", "Bloque"), IIf(SeleccTipoMinuta_MVI = "Segmento", "Cliente_SitioRemoto", "Cliente_CopiaMinutaBloque"))
    Call B_TabEst.Show(1)
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.text = vg_codigo
    fpayuda(0).Caption = vg_nombre
    fpText.SetFocus

Case 1
    
    'REGIMEN ORIGEN
    
    swUp_REG = True   'MVA - MVI
    
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = ""
    vg_codigo = ""
    Call B_TabEst.LlenaDatos("Cas_a_regimen", fpText.text, "Regimen", "Regimen_SitioRemoto_block")
    'FIN MVA - CREACION MINUTA DESDE CENTRAL - 2013-01-10
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    fpLongInteger1(1).SetFocus

Case 2
    
    swUp_SERV = True 'MVA - MVI
    
    'SERVICIO ORIGEN
    vg_left = fpayuda(2).Left + 2300
    vg_nombre = ""
    vg_codigo = ""

    'MVA - CREACION MINUTA DESDE CENTRAL - 2013-01-10
    Call B_TabEst.LlenaDatos("Cas_a_servicio", fpText.text, "Servicio", "Servicio_SitioRemoto_block")
    'FIN MVA - CREACION MINUTA DESDE CENTRAL - 2013-01-10

    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre
    fpLongInteger1(2).SetFocus

Case 3
    
    'SERVICIO ORIGEN

    swUp_SERV = False 'MVA - MVI

    vg_left = fpayuda(2).Left + 2300
    vg_nombre = ""
    vg_codigo = ""

    'MVA - CREACION MINUTA DESDE CENTRAL - 2013-01-10
     Call B_TabEst.LlenaDatos("Cas_a_servicio", fpText.text, "Servicio", "Servicio_SitioRemoto_block")
    'FIN MVA - CREACION MINUTA DESDE CENTRAL - 2013-01-10
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(3).Value = Val(vg_codigo)
    fpayuda(3).Caption = vg_nombre
    fpDateTime1(0).SetFocus

Case 4
    
    swUp_REG = False 'MVA - MVI
    
    'REGIMEN DESTINO
    'MVA - CREACION MINUTA DESDE CENTRAL - 2013-01-10
    
    vg_left = fpayuda(4).Left + 2300
    vg_nombre = ""
    vg_codigo = ""
    Call B_TabEst.LlenaDatos("Cas_a_regimen", fpText.text, "Regimen", "Regimen_SitioRemoto_block")
    B_TabEst.Show 1
    
    On Error Resume Next
    RS.Close: Set RS = Nothing
    'FIN MVA - CREACION MINUTA DESDE CENTRAL - 2013-01-10
   
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(4).Caption = vg_nombre
    fpLongInteger1(3).SetFocus

Case 5

    swUp_CECO = False 'MVA - MVI

    'CARGA CCOSTO o CLIENTES DESTINO
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = ""
    vg_codigo = ""
    Call B_TabEst.LlenaDatos("b_clientes", "cli_", "Clientes", "Cliente_CopiaMinutaBloque")
    Call B_TabEst.Show(1)
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText1.text = vg_codigo
    fpayuda(5).Caption = vg_nombre
    fpLongInteger1(1).SetFocus

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Function LeeCodigo(ByVal cadena As String) As String
    
On Error GoTo Man_Error

    LeeCodigo = Mid(cadena, 1, InStr(cadena, " "))
    
Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Function

Private Function LeeDescrip(ByVal cadena As String) As String
    
On Error GoTo Man_Error

    LeeDescrip = Mid(cadena, InStr(cadena, " "))
    
Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Function

Private Function ValidaGrilla(ByVal Spread As vaSpread, ByVal opcion As Integer) As Boolean

On Error GoTo Man_Error

ValidaGrilla = False
Dim estado As String
Dim cont As Integer

cont = 0

If opcion = 1 Then

    For i = 1 To Spread.MaxRows
    
        Spread.Row = i
        Spread.Col = 1
        
        Spread.Col = 5
        estado = Spread.text
        Spread.Col = 1
        If Spread.text = "1" And estado = "" Then
            
            ValidaGrilla = True
            Exit Function
            
        End If
    
    Next

ElseIf opcion = 2 Then

    For i = 1 To Spread.MaxRows
    
        Spread.Row = i
        Spread.Col = 1
        
        If Spread.text <> "1" Then
            
            cont = cont + 1
            
        End If
    
    Next

    If cont = Spread.MaxRows Then ValidaGrilla = True

End If

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Function

Private Function ValidaControles() As Boolean

On Error GoTo Man_Error

ValidaControles = False

If cboTipoMinuta = "" Then
    
    ValidaControles = True
    Exit Function

End If

'ORIGEN
If fpText = "" Or Trim(fpayuda(0).Caption) = "" Then
    
    ValidaControles = True
    Exit Function

End If

If fpLongInteger1(1) = "" Or Trim(fpayuda(1).Caption) = "" Then
    
    ValidaControles = True
    Exit Function

End If

If (fpLongInteger1(2) = "" Or Trim(fpayuda(2).Caption) = "") And Option1(0).Value = True Then
    
    ValidaControles = True
    Exit Function

End If

If fpDateTime1(0) = "" Then
    
    ValidaControles = True
    Exit Function

End If

'DESTINO
If fpText1 = "" Or Trim(fpayuda(5).Caption) = "" Then
    
    ValidaControles = True
    Exit Function

End If

If fpLongInteger1(0) = "" Or Trim(fpayuda(4).Caption) = "" Then
    
    ValidaControles = True
    Exit Function

End If

If (fpLongInteger1(3) = "" Or Trim(fpayuda(3).Caption) = "") And Option1(0).Value = True Then
    
    ValidaControles = True
    Exit Function

End If

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Function

Private Function ValidaCodigos() As Boolean

On Error GoTo Man_Error

ValidaCodigos = False

If CDate(fpDateTime1(0)) > CDate(fpDateTime1(1)) Then
    
    MsgBox "La fecha origen no puede ser mayor a la fecha destino", vbCritical, Me.Caption
    ValidaCodigos = True: Exit Function

End If

'validacion de origen
If cboTipoMinuta = "Bloque" Then
   
   'selecc de un codigo de CECO
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
                
   Set RS = vg_db.Execute("SELECT a.* " & _
                          "FROM b_clientes as a with (nolock) " & _
                          "inner join b_tipominuta as b with (nolock) on b.tip_codigo = a.cli_tipominuta " & _
                          "                                          and b.activo = '1' " & _
                          "WHERE a.cli_codigo = '" & Trim(fpText.text) & "' " & _
                          "AND   a.cli_tipo   = 0 " & _
                          "AND   a.cli_tipominuta in (3,4)")
'                          "AND   cli_tipominuta in (1,3)")
                                     
   If RS.EOF Then
                        
      MsgBox "El código de centro de costo origen es inválido", vbCritical, Me.Caption
      ValidaCodigos = True: Exit Function
                
   End If
                                     
   'para un codigo de regimen
        
   Sql = " SELECT DISTINCT"
   Sql = Sql & " reg_codigo, reg_nombre  "
   Sql = Sql & " FROM a_regimen With(NoLock)  "
   Sql = Sql & " WHERE  reg_activo = '1' and reg_indppr = '1' "
   Sql = Sql & " and upper(reg_codigo) = '" & fpLongInteger1(1) & "'"
'  sql = sql & " and reg_cecori = '" & m_copia_min_seg.fpText & "'"
   Sql = Sql & " ORDER BY reg_nombre"
                
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute(Sql)
                     
   If RS.EOF Then
                        
      MsgBox "El código de régimen origen es inválido", vbCritical, Me.Caption
      ValidaCodigos = True: Exit Function
                
   End If
                     
   'para un codigo de servicio
   If Option1(0).Value = True Then
                
                   Sql = " SELECT DISTINCT  "
                   Sql = Sql & " ser_codigo, ser_nombre   "
                   Sql = Sql & " FROM a_servicio Reg With(NoLock)  "
                   Sql = Sql & " where  ser_activo = '1' and ser_indppr = '1' "
                   Sql = Sql & " and ser_codigo = '" & UCase(LimpiaDato(fpLongInteger1(2))) & "'"
                   Sql = Sql & " ORDER BY ser_nombre "
                
                   If RS.State = 1 Then RS.Close
                   RS.CursorLocation = adUseClient
                   vg_db.CursorLocation = adUseClient
                   
                   Set RS = vg_db.Execute(Sql)
        
                   If RS.EOF Then
                      
                      MsgBox "El código de servicio origen es inválido", vbCritical, Me.Caption
                      ValidaCodigos = True: Exit Function
                   
                   End If
                
                End If
        Else

'validaciones para  SEGMENTO

'para la selecc de un subsemento

                Sql = " SELECT DISTINCT a.sub_codigo, a.sub_nombre"
                Sql = Sql & " FROM a_subsegmento a with (nolock) "
                Sql = Sql & " LEFT outer JOIN b_detlistaprecio b with (nolock) on a.sub_codigo = b.dlp_codigo "
                Sql = Sql & " WHERE sub_activo = 1"
                Sql = Sql & " AND a.sub_codigo = '" & UCase(LimpiaDato(fpText.text)) & "'"
                Sql = Sql & " ORDER BY a.sub_nombre"
                
                If RS.State = 1 Then RS.Close
                RS.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                Set RS = vg_db.Execute(Sql)

                If RS.EOF Then
                        
                    MsgBox "El código de sub segmento origen es inválido", vbCritical, Me.Caption
                        
                    ValidaCodigos = True: Exit Function
                
                End If
 

' para la seleccion de un regimen

                Sql = " SELECT distinct a_regimen.reg_codigo, a_regimen.reg_nombre "
                Sql = Sql & " FROM a_regimen with (nolock) INNER JOIN"
                Sql = Sql & "  b_minuta with (nolock) ON a_regimen.reg_codigo = b_minuta.min_codreg"
                Sql = Sql & " WHERE     (a_regimen.reg_activo = '1') AND (a_regimen.reg_indppr = '3') OR"
                Sql = Sql & " (a_regimen.reg_activo = '1') AND ('3' = '3') "
                Sql = Sql & " AND b_minuta.min_subseg = '" & m_copia_min_seg.fpText & "'"
                Sql = Sql & " AND a_regimen.reg_codigo = '" & UCase(LimpiaDato(fpLongInteger1(1))) & "'"
                Sql = Sql & " ORDER BY a_regimen.reg_nombre "
            
                If RS.State = 1 Then RS.Close
                RS.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS = vg_db.Execute(Sql)

                If RS.EOF Then
                        
                   MsgBox "El código de régimen origen es inválido", vbCritical, Me.Caption
                        
                   ValidaCodigos = True: Exit Function
                
                End If

' para la seleccion de un servicio
                 If Option1(0).Value = True Then
                    
                    Sql = " SELECT distinct a_servicio.ser_codigo, a_servicio.ser_nombre "
                    Sql = Sql & " FROM a_servicio with (nolock) INNER JOIN"
                    Sql = Sql & " b_minuta with (nolock) ON a_servicio.ser_codigo = b_minuta.min_codser"
                    Sql = Sql & " WHERE (a_servicio.ser_indppr = '3') OR"
                    Sql = Sql & " ('3' = '3') "
                    Sql = Sql & " AND (b_minuta.min_subseg = '" & m_copia_min_seg.fpText & "')"
                    Sql = Sql & " AND (a_servicio.ser_codigo = '" & UCase(LimpiaDato(fpLongInteger1(2))) & "')"
                    Sql = Sql & " ORDER BY a_servicio.ser_nombre"
            
                    If RS.State = 1 Then RS.Close
                    RS.CursorLocation = adUseClient
                    vg_db.CursorLocation = adUseClient
                    Set RS = vg_db.Execute(Sql)

                    If RS.EOF Then
                       
                       MsgBox "El código de servicio origen es inválido", vbCritical, Me.Caption
                       ValidaCodigos = True: Exit Function
                    
                    End If
                    
                End If
End If


'validaciones para BLOQUE DESTINO

'selecc de un codigo de CECO
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT a.* " & _
                       "FROM b_clientes as a with (nolock) " & _
                       "inner join b_tipominuta as b with (nolock) on b.tip_codigo = a.cli_tipominuta " & _
                       "                                          and b.activo = '1' " & _
                       "WHERE a.cli_codigo = '" & Trim(fpText1.text) & "' " & _
                       "AND   a.cli_tipo   = 0 " & _
                       "AND   a.cli_tipominuta in (3,4)")
'                      "AND   cli_tipominuta in (1,3)")
                     
If RS.EOF Then
        
   MsgBox "El código de centro de costo origen es inválido", vbCritical, Me.Caption
   ValidaCodigos = True: Exit Function

End If
                     
'para un codigo de regimen

Sql = " SELECT DISTINCT"
Sql = Sql & " reg_codigo, reg_nombre  "
Sql = Sql & " FROM a_regimen With(NoLock)  "
Sql = Sql & " WHERE  reg_activo = '1' and reg_indppr = '1' "
Sql = Sql & " and upper(reg_codigo) = '" & fpLongInteger1(0) & "'"
Sql = Sql & " ORDER BY reg_nombre"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute(Sql)
     
If RS.EOF Then

   MsgBox "El código de régimen origen es inválido", vbCritical, Me.Caption
   ValidaCodigos = True
   Exit Function

End If
     
'para un codigo de servicio
If Option1(0).Value = True Then
   
   Sql = " SELECT DISTINCT  "
   Sql = Sql & " ser_codigo, ser_nombre   "
   Sql = Sql & " FROM a_servicio With(NoLock)  "
   Sql = Sql & " where  ser_activo = '1' and ser_indppr = '1' "
   Sql = Sql & " and ser_codigo = '" & UCase(LimpiaDato(fpLongInteger1(3))) & "'"
   'sql = sql & " and ser_cecori = '" & fpText1 & "'"
   Sql = Sql & " ORDER BY ser_nombre "

   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   Set RS = vg_db.Execute(Sql)

   If RS.EOF Then
      
      MsgBox "El código de servicio origen es inválido", vbCritical, Me.Caption
      ValidaCodigos = True: Exit Function
   
   End If

End If

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Function


Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

vaSpread1.MaxRows = 0
Select Case Index

Case 0
     
     Frame2.Caption = "Seleccionar estructura servicio origen && destino"
     vaSpread1.Col = 3
     vaSpread1.Row = 0
     vaSpread1.text = "Descripción Estructura Origen"
     vaSpread1.Col = 5
     vaSpread1.text = "Descripción Estructura Origen"
     vaSpread1.Row = -1
     vaSpread1.Col = 5
     vaSpread1.CellType = CellTypeComboBox
     '-------> Desbloquear servicios y ayuda
     fpLongInteger1(2).Enabled = True
     fpLongInteger1(3).Enabled = True
     Image1(2).Enabled = True
     Image1(3).Enabled = True
     btnProcesarEstructuras.Caption = "Procesar Estructura"

Case 1
     
     Frame2.Caption = "Seleccionar servicios"
     vaSpread1.Col = 3
     vaSpread1.Row = 0
     vaSpread1.text = "Descripción Servicio"
     vaSpread1.Col = 5
     vaSpread1.text = "Observación"
     vaSpread1.Row = -1
     vaSpread1.Col = 5
     vaSpread1.CellType = CellTypeStaticText

     '-------> Bloquear Servicios y ayuda
     fpLongInteger1(2).Value = ""
     fpLongInteger1(2).Enabled = False
     fpLongInteger1(3).Value = ""
     fpLongInteger1(3).Enabled = False
     Image1(2).Enabled = False
     Image1(3).Enabled = False
     btnProcesarEstructuras.Caption = "Procesar Servicio"

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index

Case 2
    
    If Option1(0).Value = True Then
       
       MinutaBloqueunServicio
    
    Else
       
       MinutaMultipleServicio
    
    End If

Case 4
    
    Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Sub MinutaBloqueunServicio()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim descrip As String
Dim FecFin As Date
Dim FecIni As Date
Dim FecIniOri As Long
Dim FecFinOri As Long
Dim i As Long
Dim OpReemplazaMinuta As String
Dim Sql As String

If ValidaCodigos = True Then Exit Sub

'-------> Validar largo de días bloque minuta
If Val(fpLongInteger1(4).text) <= 0 Then
   
   MsgBox "Debe largo de días, proceso cancelado", vbCritical, Me.Caption
   Exit Sub

End If

'-------> Validar largo de días sea mayor 93
If Val(fpLongInteger1(4).text) > 98 Then
   
   MsgBox "Maximo de días son 98, proceso cancelado", vbCritical, Me.Caption
   Exit Sub

End If

'-------> Validar largo de origen vs entrada
If DateDiff("d", fpDateTime1(0).text, fpDateTime1(1).text) + 1 > Val(fpLongInteger1(4).text) Then

   MsgBox "El largo de día seleccionado origen es mayor al largo de días seleccionado, proceso cancelado", vbCritical, Me.Caption
   Exit Sub

End If

'------->Validar Minuta Bloque
Dim Ceco As String
FecIniOri = 0
FecFinOri = 0
If SeleccTipoMinuta_MVI = "Bloque" Then
   
   Ceco = LimpiaDato(Trim(fpText.text))

Else
   
   Ceco = LimpiaDato(Trim(fpText1.text))

End If

'-------> Validar diías dentro de una minuta
Sql = ""
Sql = Trim(LimpiaDato(fpText1.text))
Sql = Sql & ", " & fpLongInteger1(0).Value & ", " & fpLongInteger1(3).Value & ", " & Format(fpDateTime1(2).text, "yyyymmdd") & ", " & Val(fpLongInteger1(4).text) & ""

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_ValidarMinutaBloqueLargoDias " & Sql & "")

If Not RS.EOF Then
   
   If (RS(0) = "S" And RS(1) = "N") Or (RS(0) = "N" And RS(1) = "S") Then
      
      MsgBox "Existen minuta bloque para este periodo, proceso cancelado" & VgLinea & VgLinea & "( Bloque = " & RS(2) & " del Periodo " & RS(4) & " Hasta " & RS(5) & ")", vbCritical, Me.Caption
      RS.Close
      Set RS = Nothing
      Exit Sub
   
   ElseIf RS(2) <> RS(3) And RS(2) > 0 And RS(3) > 0 Then
      
      MsgBox "Existen minuta bloque para este periodo, proceso cancelado" & VgLinea & VgLinea & "( Bloque = " & RS(3) & " del Periodo " & RS(4) & " Hasta " & RS(5) & ")", vbCritical, Me.Caption
      RS.Close
      Set RS = Nothing
      Exit Sub
   
   End If

End If
RS.Close
Set RS = Nothing

'-------> Validar largo de día
Sql = ""
Sql = Trim(LimpiaDato(fpText1.text))
Sql = Sql & ", " & fpLongInteger1(0).Value & ", " & fpLongInteger1(3).Value & ", " & Format(fpDateTime1(2).text, "yyyymmdd") & ", " & Val(fpLongInteger1(4).text) & ""

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_ValidarLargoDiasMinutaBloque " & Sql & "")
If Not RS.EOF Then
   
   If (RS(0) = "S" And RS(1) = "S") Then

      MsgBox "El largo de días no corresponde con el original, proceso cancelado " & " Bloque =  " & " " & RS(2) & " su largo corresponde = " & RS(3)
      RS.Close
      Set RS = Nothing
      Exit Sub
   
   End If

End If
RS.Close
Set RS = Nothing

Sql = ""
Sql = Ceco
Sql = Sql & ", " & fpLongInteger1(1).Value & ", " & fpLongInteger1(2).Value & ", " & Format(fpDateTime1(0).text, "yyyymmdd") & ", " & Format(fpDateTime1(1).text, "yyyymmdd")

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_ValidarMinutaCopia_V02 " & Sql & "")
If Not RS.EOF Then
   
   FecIniOri = RS!fechadesde
   FecFinOri = RS!fechahasta

End If
RS.Close
Set RS = Nothing

If Ceco = LimpiaDato(Trim(fpText1.text)) And fpLongInteger1(1).Value = fpLongInteger1(0).Value And _
   fpLongInteger1(2).Value = fpLongInteger1(3).Value And Val(Format(fpDateTime1(2).text, "yyyymmdd")) >= FecIniOri And _
   Val(Format(fpDateTime1(2).text, "yyyymmdd")) <= FecFinOri Then
    
   MsgBox "Los dato destino coincide con los dato origen, proceso cancelado", vbCritical, Me.Caption
   Exit Sub

End If
''--------> Validar que ceco sea de tipo bloque no puede crear minuta
FecFin = fpDateTime1(2).text
FecIni = fpDateTime1(2).text

For i = 1 To Format(fpDateTime1(1), "YYYYMMDD") - Format(fpDateTime1(0), "YYYYMMDD")
    
    FecFin = FecFin + 1

Next i

'-------> validacion en relacion al estado de la minuta
Sql = " sgpadm_p_copia_minValidaMinuta_MVI_V02 "
Sql = Sql & " '" & LimpiaDato(fpText1) & "'" 'ceco destino
Sql = Sql & ", " & fpLongInteger1(0) 'regimen destino
Sql = Sql & ", " & fpLongInteger1(3) 'servicio destino
Sql = Sql & ", " & Format(FecIni, "YYYYMMDD") ' fecha desde
Sql = Sql & ", " & Format(FecFin, "YYYYMMDD") ' fecha hasta

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute(Sql)
OpReemplazaMinuta = "S"
If Not RS.EOF Then
    
    If RS!MIN_INDBLO <> 11 Then
        
        MsgBox "Minuta bloqueada, no esta disponible para copia", vbCritical, Me.Caption
        Exit Sub
    
    End If
    
    If RS!MIN_INDBLO = 11 Then
        
        If MsgBox("Esta minuta ya posee datos copiados, ¿ (Si = Será sobreescrita - (No) Anexar estructura) ? ", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
           
           OpReemplazaMinuta = "N"
'            Exit Sub
        End If
    
    End If

End If

'debe haber ingresado/llenado valores en pantalla
If ValidaControles = True Then
    
   MsgBox "Falta(n) ingresar/seleccionar valores en pantalla", vbExclamation, Me.Caption
   Exit Sub

End If
If vaSpread1.MaxRows = 0 Then Exit Sub

'debe tener algun elemento de su izq. seleccionado
If ValidaGrilla(vaSpread1, 2) = True Then

   MsgBox "Debe tickear al menos un elemento a la izquierda", vbExclamation, Me.Caption
   Exit Sub
    
End If

'debe tener algo seleccionado a la derecha si esta tickeado a la izq.
If ValidaGrilla(vaSpread1, 1) = True Then

   MsgBox "Debe seleccionar un elemento a la derecha", vbExclamation, Me.Caption
   Exit Sub
    
End If

'luego aca va el proced. almac que copia los valores a las tablas destino
Call LlenaTablaPasoEst(OpReemplazaMinuta)

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub MinutaMultipleServicio()

On Error GoTo Man_Error

Dim RS         As New ADODB.Recordset
Dim descrip    As String
Dim FecFin     As Date
Dim FecIni     As Date
Dim FecIniOri  As Long
Dim FecFinOri  As Long
Dim i          As Long
Dim EstCopiado As Boolean
Dim Sql        As String

'-------> Validar días bloque minuta
If Val(fpLongInteger1(4).text) <= 0 Then
   
   MsgBox "Debe largo de días, proceso cancelado", vbCritical, Me.Caption
   Exit Sub

End If

'-------> Validar largo de días sea mayor 98
If Val(fpLongInteger1(4).text) > 98 Then
   
   MsgBox "Maximo de días son 98, proceso cancelado", vbCritical, Me.Caption
   Exit Sub

End If

'-------> Validar largo de origen vs entrada
If DateDiff("d", fpDateTime1(0).text, fpDateTime1(1).text) + 1 > Val(fpLongInteger1(4).text) Then

   MsgBox "El largo de día seleccionado origen es mayor al largo de días seleccionado, proceso cancelado", vbCritical, Me.Caption
   Exit Sub

End If

If ValidaCodigos = True Then Exit Sub

'debe haber ingresado/llenado valores en pantalla

If ValidaControles = True Then
    
    MsgBox "Falta(n) ingresar/seleccionar valores en pantalla", vbExclamation, Me.Caption
    Exit Sub

End If

If vaSpread1.MaxRows = 0 Then Exit Sub

'debe tener algun elemento de su izq. seleccionado
If ValidaGrilla(vaSpread1, 2) = True Then

    MsgBox "Debe tickear al menos un elemento a la izquierda", vbExclamation, Me.Caption
    Exit Sub
    
End If

Dim CodServicio As Long
CodServicio = 0

For i = 1 To vaSpread1.MaxRows
    
    vaSpread1.Row = i
    vaSpread1.Col = 1
    
    If vaSpread1.text = "1" Then
       
       EstCopiado = True
       vaSpread1.Col = 2
       CodServicio = vaSpread1.text
       
       vaSpread1.Col = 5
       vaSpread1.text = ""
       
       vaSpread1.Row = i: vaSpread1.Col = -1
       vaSpread1.BackColor = &HC0FFFF
         
       Sql = " sgpadm_Sel_ValidarMinutaBloquexServicioSegmentoTramoDias_V01 "
       Sql = Sql & " '" & fpText1 & "'" 'ceco destino
       Sql = Sql & ", " & fpLongInteger1(0) 'regimen destino
       Sql = Sql & ", " & CodServicio 'servicio destino
       Sql = Sql & ", " & Format(fpDateTime1(0), "YYYYMMDD") ' fecha desde
       Sql = Sql & ", " & Format(fpDateTime1(1), "YYYYMMDD") ' fecha hasta
       Sql = Sql & ", " & Format(fpDateTime1(2), "YYYYMMDD") ' fecha destino
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
            
       Set RS = vg_db.Execute(Sql)
       If Not RS.EOF Then
                
          If RS(0) = "S" Then

             vaSpread1.Row = i: vaSpread1.Col = -1
             vaSpread1.BackColor = &H8080FF
               
             vaSpread1.Col = 5
             vaSpread1.text = "Las fechas indicada origen sobre pasa el bloque " & RS(2) & " " & RS(3)
             EstCopiado = False
              
          End If
              
       End If
       RS.Close
       Set RS = Nothing

       If EstCopiado Then
        
            If SeleccTipoMinuta_MVI = "Bloque" Then
                
                Sql = ""
                Sql = " sgpadm_Ins_CopiaMinutaServicioBloqueaBloque_V03 "
                Sql = Sql & " '" & fpText & "'" 'ceco origen
                Sql = Sql & " ," & fpLongInteger1(1)  'reg origen
                Sql = Sql & " ," & CodServicio  'serv origen
        
                Sql = Sql & " ,'" & fpText1 & "'" 'ceco destino
                Sql = Sql & " ," & fpLongInteger1(0) & "" 'reg destino
                Sql = Sql & " ," & CodServicio & "" 'serv destino
        
                Sql = Sql & " ," & Format(fpDateTime1(0), "YYYYMMDD")  'fecha desde origen
                Sql = Sql & " ," & Format(fpDateTime1(1), "YYYYMMDD")  'fecha hasta origen
                Sql = Sql & " ,'" & Format(fpDateTime1(2), "YYYYMMDD") & "'"  'fecha desde destino
                Sql = Sql & " ," & Val(fpLongInteger1(4).text) & ""
            
            Else
            
                Sql = ""
                Sql = " sgpadm_Pro_CopiaMinutaServicioSegmentoaBloque_V01 "
                Sql = Sql & " '" & fpText & "'" 'ceco origen
                Sql = Sql & " ," & fpLongInteger1(1)  'reg origen
                Sql = Sql & " ," & CodServicio  'serv origen
        
                Sql = Sql & " ,'" & fpText1 & "'" 'ceco destino
                Sql = Sql & " ," & fpLongInteger1(0) & "" 'reg destino
                Sql = Sql & " ," & CodServicio & "" 'serv destino
        
                Sql = Sql & " ," & Format(fpDateTime1(0), "YYYYMMDD")  'fecha desde origen
                Sql = Sql & " ," & Format(fpDateTime1(1), "YYYYMMDD")  'fecha hasta origen
                Sql = Sql & " ,'" & Format(fpDateTime1(2), "YYYYMMDD") & "'" 'fecha desde destino
                Sql = Sql & " ," & Val(fpLongInteger1(4).text) & ""
            
            End If
        
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            Set RS = vg_db.Execute(Sql)
            
            If Not RS.EOF Then
               
               vaSpread1.Col = 5
               
               If RS(0) > 0 Then
                  
                  vaSpread1.text = RS(0) & " " & RS(1)
                  MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
               
               Else
                  
                  vaSpread1.text = "Servicio Procesado [OK]"
               
               End If
            
            End If
            RS.Close
            Set RS = Nothing

        End If
        
    End If

Next i

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub LlenaTablaPasoEst(OpReemplazaMinuta As String)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim codest2 As Variant
Dim cSpi As Long
Dim codest1 As Long

    '-------> Borrar tabla de paso estructura servicio
    vg_db.Execute "DELETE paso_estservicio WHERE ess_spid = @@spid and ess_usr = '" & vg_NUsr & "'"
    '-------> Buscar spid
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS = vg_db.Execute("SELECT @@spid spid")
    If Not RS.EOF Then cSpi = RS!spid
    RS.Close
    Set RS = Nothing
    '-------> Grabar tabla de paso estructura servicio
    For i = 1 To vaSpread1.MaxRows
       
       vaSpread1.Row = i
       vaSpread1.Col = 1
       vaSpread1.Col = 2
       codest1 = Val(vaSpread1.text)
       vaSpread1.Col = 4
       codest2 = Val(vaSpread1.text)
       vaSpread1.Col = 1
       
       If vaSpread1.text = "1" And vaSpread1.Row > 0 And codest1 > 0 And codest2 > 0 Then
          
          vaSpread1.Col = 2
          codest1 = Val(vaSpread1.text)
          vaSpread1.Col = 4
          codest2 = Val(vaSpread1.text)
          vaSpread1.Col = 5
          vg_db.Execute ("INSERT INTO paso_estservicio (ess_spid, ess_usr, ess_codess1, ess_codess2, ess_desest2) VALUES(" & cSpi & ", '" & vg_NUsr & "', " & codest1 & ", " & codest2 & ", '" & Trim((Trim(vaSpread1.text))) & "')")
       
       End If
    
    Next i
    
    If SeleccTipoMinuta_MVI = "Bloque" Then
        
        Sql = ""
'        Sql = " sgpadm_Pro_CopiaMinutaBloqueaBloque "
        Sql = " sgpadm_Ins_CopiaMinutaBloqueaBloque_V03 "
        Sql = Sql & " '" & LimpiaDato(Trim(fpText)) & "'" 'ceco origen
        Sql = Sql & " ," & fpLongInteger1(1)  'reg origen
        Sql = Sql & " ," & fpLongInteger1(2)  'serv origen

        Sql = Sql & " ,'" & LimpiaDato(Trim(fpText1)) & "'" 'ceco destino
        Sql = Sql & " ," & fpLongInteger1(0) & "" 'reg destino
        Sql = Sql & " ," & fpLongInteger1(3) & "" 'serv destino

        Sql = Sql & " ," & Format(fpDateTime1(0), "YYYYMMDD")  'fecha desde origen
        Sql = Sql & " ," & Format(fpDateTime1(1), "YYYYMMDD")  'fecha hasta origen
        Sql = Sql & " ,'" & Format(fpDateTime1(2), "YYYYMMDD") & "'"  'fecha desde destino

        Sql = Sql & " ,0"
        
        Sql = Sql & " ," & cSpi
        Sql = Sql & " ,'" & vg_NUsr & "'"
        Sql = Sql & " ,'" & OpReemplazaMinuta & "'"
        Sql = Sql & " ," & Val(fpLongInteger1(4).text) & ""

    Else
    
        Sql = ""
        Sql = " sgpadm_Pro_CopiaMinutaSegmentoaBloque_V01 "
        Sql = Sql & " '" & fpText & "'" 'ceco origen
        Sql = Sql & " ," & fpLongInteger1(1)  'reg origen
        Sql = Sql & " ," & fpLongInteger1(2)  'serv origen

        Sql = Sql & " ,'" & LimpiaDato(Trim(fpText1)) & "'" 'ceco destino
        Sql = Sql & " ," & fpLongInteger1(0) & "" 'reg destino
        Sql = Sql & " ," & fpLongInteger1(3) & "" 'serv destino

        Sql = Sql & " ," & Format(fpDateTime1(0), "YYYYMMDD")  'fecha desde origen
        Sql = Sql & " ," & Format(fpDateTime1(1), "YYYYMMDD")  'fecha hasta origen
        Sql = Sql & " ,'" & Format(fpDateTime1(2), "YYYYMMDD") & "'" 'fecha desde destino

        Sql = Sql & " ,0"
        
        Sql = Sql & " ," & cSpi
        Sql = Sql & " ,'" & vg_NUsr & "'"
        Sql = Sql & " ,'" & OpReemplazaMinuta & "'"
        
        Sql = Sql & " ," & Val(fpLongInteger1(4).text) & ""
        
    End If
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute(Sql)
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
       
       Else
          
          MsgBox "Proceso Finalizado", vbInformation, Me.Caption
       
       End If
    
    End If
    RS.Close
    Set RS = Nothing

    Dim spid As Long
    
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub MoverVector()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i  As Long

On Error Resume Next

If Trim(fpText.text) = "" Or Trim(fpLongInteger1(1).text) = "" Or Trim(fpLongInteger1(2).text) = "" Or Trim(fpText1) = "" Or Trim(fpLongInteger1(0).text) = "" Or Trim(fpLongInteger1(3).text) = "" Then Exit Sub

'Mover estructura servicio Origen
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_CopiaMinuta_V01 " & Val(fpLongInteger1(5).Value) & ", " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Val(Format(fpDateTime1(0).text, "yyyymm")) & " ")
vaSpread1.MaxRows = 0
If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 2
      vaSpread1.text = RS!mid_estser
      vaSpread1.Col = 3
      vaSpread1.text = IIf(Trim(RS!mid_desest) <> "", Trim(RS!mid_desest), Trim(RS!ess_nombre))
      RS.MoveNext
   
   Loop

End If
RS.Close
Set RS = Nothing
Dim codest As Long
Dim codaux As Long

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open "SELECT DISTINCT isnull(ess_codigo,0) as ess_codigo, isnull(ess_nombre,'') as ess_nombre, isnull(ess_orden,0) as ess_orden FROM a_estservicio With(NoLock) WHERE ess_codser = " & Val(fpLongInteger1(3).Value) & " ORDER BY ess_orden", vg_db, adOpenForwardOnly ', adOpenStatic
If Not RS.EOF Then
   
   For i = 1 To vaSpread1.MaxRows
       
       vaSpread1.Row = i
       vaSpread1.Col = 5
       vaSpread1.Col = 2
       codest = vaSpread1.text
       lisnom = ""
       liscod = ""
       
       Do While Not RS.EOF
          
          vaSpread1.Col = 4
          liscod = liscod & IIf(liscod <> "", Chr$(9), "") & RS!ess_codigo
          vaSpread1.Col = 5
          lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(RS!ess_nombre)
          vaSpread1.Col = 4
          vaSpread1.TypeComboBoxList = liscod
          vaSpread1.Col = 5
          vaSpread1.TypeComboBoxList = lisnom
          
          RS.MoveNext
       
       Loop
       RS.MoveFirst
       If fpLongInteger1(1).Value = fpLongInteger1(3).Value Then
          
          vaSpread1.Col = 4
          vaSpread1.TypeComboBoxList = liscod
          
          For z = 0 To vaSpread1.TypeComboBoxCount
              
              vaSpread1.TypeComboBoxCurSel = z
              If vaSpread1.text = codest Then codaux = z: Exit For
              codaux = -1
          
          Next z
          
          vaSpread1.Col = 5
          vaSpread1.TypeComboBoxCurSel = codaux
       
       End If
   
   Next i

End If
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1
    
    Dim i As Long
    vaSpread1.Col = 1
'    For i = BlockRow To BlockRow2
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 5
        
        If Trim(vaSpread1.text) <> "Minuta Esta Bloqueada" Or Trim(vaSpread1.text) = "Existen minuta bloque para este periodo" Then
           
           vaSpread1.Col = 1
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

If Est Or vaSpread1.MaxRows < 0 Then Exit Sub
vaSpread1.Row = Row
vaSpread1.Col = 5
If Trim(vaSpread1.text) = "Minuta Esta Bloqueada" Or Trim(vaSpread1.text) = "Existen minuta bloque para este periodo" Then
   
   vaSpread1.Col = 1
   Est = True
   vaSpread1.Value = "0"
   Est = False

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Select Case Col

Case 5
    
    Dim indice As Long
    vaSpread1.Row = Row
    vaSpread1.Col = 5
    indice = vaSpread1.TypeComboBoxCurSel
    vaSpread1.Col = 4
    vaSpread1.TypeComboBoxCurSel = indice

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Dim IndCol As Long
IndCol = vaSpread1.ActiveCol
Select Case KeyCode

Case 46 And IndCol = 5
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 4
    vaSpread1.TypeComboBoxCurSel = -1
    vaSpread1.Col = 5
    vaSpread1.TypeComboBoxCurSel = -1

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub
