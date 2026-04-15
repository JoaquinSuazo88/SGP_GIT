VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_Plami1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5025
   ClientLeft      =   3090
   ClientTop       =   3060
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Tahoma"
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5025
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   90
      Width           =   7215
      Begin VB.Frame Frame2 
         Height          =   2625
         Left            =   1215
         TabIndex        =   10
         Top             =   2160
         Width           =   5775
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   2325
            Left            =   90
            TabIndex        =   4
            Top             =   180
            Width           =   5595
            _Version        =   393216
            _ExtentX        =   9869
            _ExtentY        =   4101
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
            MaxCols         =   7
            MaxRows         =   6
            ScrollBars      =   0
            SelectBlockOptions=   0
            SpreadDesigner  =   "M_Plami1.frx":0000
            UserResize      =   0
         End
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   1
         Top             =   760
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
         Left            =   1380
         TabIndex        =   2
         Top             =   1100
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
         Left            =   1380
         TabIndex        =   0
         Top             =   435
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
         Left            =   1395
         TabIndex        =   3
         Top             =   1440
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1676
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
         Text            =   "08/2023"
         DateCalcMethod  =   4
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
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1395
         TabIndex        =   18
         Top             =   1800
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
         AlignTextH      =   0
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
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   200
         MultiLine       =   0   'False
         PasswordChar    =   "*"
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   1
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1905
         Width           =   975
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
         Left            =   120
         TabIndex        =   17
         Top             =   1530
         Width           =   1110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3120
         TabIndex        =   15
         Top             =   1095
         Width           =   3795
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3120
         TabIndex        =   13
         Top             =   765
         Width           =   3795
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3120
         TabIndex        =   11
         Top             =   435
         Width           =   3795
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
         Left            =   105
         TabIndex        =   9
         Top             =   1200
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
         Left            =   105
         TabIndex        =   8
         Top             =   870
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
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   540
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2610
         Picture         =   "M_Plami1.frx":048C
         Top             =   345
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2610
         Picture         =   "M_Plami1.frx":0796
         Top             =   690
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2610
         Picture         =   "M_Plami1.frx":0AA0
         Top             =   1030
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
         Left            =   3165
         TabIndex        =   12
         Top             =   480
         Width           =   3795
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
         Left            =   3165
         TabIndex        =   14
         Top             =   810
         Width           =   3795
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
         Left            =   3165
         TabIndex        =   16
         Top             =   1140
         Width           =   3795
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5025
      Left            =   7425
      TabIndex        =   5
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   8864
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_Plami1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim NomFor As String
Public lc_Aux As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_carga ""
Me.HelpContextID = vg_OpcM
Me.Height = 5535
Me.Width = 8175
fg_centra Me
'-------> 20131025
Label2(1).Visible = False
fpText1.Visible = False

If lc_Aux = "PlaTeo" Then
'    If ValidarMinutaTeorica Then
'       Label2(1).Visible = True
'       fpText1.Visible = True
'    Else
'       Label2(1).Visible = False
'       fpText1.Visible = False
'    End If
    MsgTitulo = "Planificación Teórica"
    Me.Caption = "Planificación Teórica"
    NomFor = "MINTEO"
ElseIf lc_Aux = "PlaRea" Then
'    If ValidarMinutaReal Then
'       Label2(1).Visible = True
'       fpText1.Visible = True
'    Else
'       Label2(1).Visible = False
'       fpText1.Visible = False
'    End If
    MsgTitulo = "Planificación Real"
    Me.Caption = "Planificación Real"
    NomFor = "MINREA"
End If
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
'Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar ": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 2) = "00", False, True)
Set BtnX = Toolbar1.Buttons.Add(, "A_Planificación", , tbrDefault, "A_Planificacion"): BtnX.Visible = True: BtnX.ToolTipText = IIf(lc_Aux = "PlaTeo", "Planificación Teórica", "Planificación Real"): BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 2) = "00", False, True)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_EstFijaDía", , tbrDefault, "A_EstFijaDía"): BtnX.Visible = True: BtnX.ToolTipText = IIf(lc_Aux = "PlaTeo", "Estructura Fija Día Teórica", "Estructura Fija Día Real"): BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 2) = "00", False, True)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico Planificación"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpDateTime1.text = Format(Date, "mm/yyyy")
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
vaSpread1.Visible = False: ArmarCalendario: MoverDatos: vaSpread1.Visible = True
fg_descarga
End Sub

Private Sub Form_Unload(Cancel As Integer)

'MVA - MVI - BLOQUEO BOTON TOOLBAR ACTUALIZAR RECETA - 2013-01-18
vg_Block_Botton_Actua_Receta_MVI = False
vg_Clave_MVI = ""
'FIN MVA - MVI - BLOQUEO BOTON TOOLBAR ACTUALIZAR RECETA - 2013-01-18

End Sub

Private Sub fpDateTime1_Change()
vaSpread1.Visible = False: ArmarCalendario: MoverDatos: vaSpread1.Visible = True
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Select Case Index
  Case 1
    RS.Open RutinaLectura.Regimen(2, Val(fpLongInteger1(1).Value), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": vaSpread1.Visible = False: ArmarCalendario: MoverDatos: vaSpread1.Visible = True: Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    vaSpread1.Visible = False: ArmarCalendario: MoverDatos: vaSpread1.Visible = True
  Case 2
    RS.Open RutinaLectura.Servicio(2, Val(fpLongInteger1(2).Value), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "":  vaSpread1.Visible = False: ArmarCalendario: MoverDatos: vaSpread1.Visible = True: Exit Sub
    fpayuda(2).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    vaSpread1.Visible = False: ArmarCalendario: MoverDatos: vaSpread1.Visible = True
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
End Select
End Sub

Private Sub fpText_Change()
RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": fpLongInteger1(1).Value = "": fpayuda(1).Caption = "": fpLongInteger1(2).Value = "": fpayuda(2).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
vaSpread1.Visible = False: ArmarCalendario: MoverDatos: vaSpread1.Visible = True
fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
fpLongInteger1(2).Value = "": fpayuda(2).Caption = ""
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
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.text = vg_codigo
    fpayuda(0).Caption = vg_nombre
    fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(2).Value = "": fpayuda(2).Caption = ""
    fpLongInteger1(1).SetFocus
Case 1
'    vg_opayuda = 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    fpLongInteger1(2).SetFocus
Case 2
    vg_left = fpayuda(2).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Ser"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre
    fpDateTime1.SetFocus
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RS As New ADODB.Recordset
Dim sql1 As String
vg_tipmin = False
Select Case Button.Index
Case 2, 4 'Acceso planificación teorica o real
    
    vg_Clave_MVI = Me.fpText1 'MVA - MVI - BLOQUEO BOTON TOOLBAR ACTUALIZAR RECETA - 2013-01-18
    
    vg_fecha = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)
    sql1 = IIf(vg_tipbase = "1", " val(mid(min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),min_fecmin),1,6)) ")
    '-------> validar si sitio es simap sacar del sistema
    If ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1) Then
       RS.Open "SELECT DISTINCT min_codigo FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND min_codreg = " & Val(fpLongInteger1(1).Value) & " AND min_codser = " & Val(fpLongInteger1(2).Value) & " AND " & sql1 & " = " & Val(vg_fecha) & " AND min_indblo IN (2,11)", vg_db, adOpenStatic
       If Not RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Minuta corresponde bloque minuta", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       RS.Close: Set RS = Nothing
    End If
    
    RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda(0).Caption = "": fpLongInteger1(1).Value = "": fpayuda(1).Caption = "": fpLongInteger1(2).Value = "": fpayuda(2).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fpayuda(0).Caption = RS!cli_nombre: RS.Close: Set RS = Nothing
    RS.Open RutinaLectura.Regimen(2, Val(fpLongInteger1(1).Value), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Regimen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fpayuda(1).Caption = RS!reg_nombre: RS.Close: Set RS = Nothing
    RS.Open RutinaLectura.Servicio(2, Val(fpLongInteger1(2).Value), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Servicio", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fpayuda(2).Caption = RS!ser_nombre: RS.Close: Set RS = Nothing
    'Si es la opción de planificación validar estructura
    vg_codcasino = LimpiaDato(Trim(fpText.text))
    vg_codregimen = Val(fpLongInteger1(1).Value)
    vg_codservicio = Val(fpLongInteger1(2).Value)
    vg_fecha = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)
    Let Vg_FechaDesde = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)
    Let Vg_FechaHasta = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)
    
    If Button.Index = 2 Then
       RS.Open RutinaLectura.EstServicio(1, Val(fpLongInteger1(2).Value), 0), vg_db, adOpenStatic
       If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe estructura de servicio", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       RS.Close: Set RS = Nothing
       '------> Validar que exista minuta
       RS.Open RutinaLectura.Minutas(5, vg_codregimen, vg_codservicio, Val(vg_fecha), "1"), vg_db, adOpenStatic
       If RS.EOF Then
          RS.Close: Set RS = Nothing
          '-------> validar si sitio es simap sacar del sistema
          If Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1) Then
             MsgBox "No puedes crear minuta concepto Simap, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo
             Exit Sub
          End If
       Else
          RS.Close: Set RS = Nothing
       End If
    ElseIf Button.Index = 4 Then
       '-------> validar si sitio es simap sacar del sistema
       If Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1) Then
          MsgBox "No puedes crear minuta estructura fija concepto Simap, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
       End If
    End If
    vg_codcasino = LimpiaDato(Trim(fpText.text))
    vg_codregimen = Val(fpLongInteger1(1).Value)
    vg_codservicio = Val(fpLongInteger1(2).Value)
    vg_fecha = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)
    Let Vg_FechaDesde = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)
    Let Vg_FechaHasta = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)
    
    Dim tipmin As String
    If NomFor = "MINTEO" Then
        tipmin = "1"
    Else
        tipmin = "2"
    End If
    'Si es la opción de planificación validar estructura
    If Button.Index = 2 Then
       RS.Open RutinaLectura.Minutas(1, Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(vg_fecha), tipmin), vg_db, adOpenStatic
       If Not RS.EOF Then
          RS.Close: Set RS = Nothing
       ElseIf RS.EOF Then
          RS.Close: Set RS = Nothing
          If (Mid(ValidarUsuario(Me), 1, 1)) = "0" Then MsgBox "No esta autorizado crear planificación, conctatece con el administrador de sistema ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       End If
       Unload B_Receta
       If NomFor = "MINTEO" Then Unload M_Plami2: M_Plami2.Show 1
       If NomFor = "MINREA" Then
          RS.Open RutinaLectura.Minutas(2, 0, 0, Val(vg_fecha), "1"), vg_db, adOpenStatic
          If RS!nreg = 0 Then RS.Close: Set RS = Nothing: MsgBox "Debe realizar la planificación teórica de este mes...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
          RS.Close: Set RS = Nothing
          RS.Open RutinaLectura.Minutas(2, 0, 0, Val(vg_fecha), "2"), vg_db, adOpenStatic
          If RS!nreg = 0 Then RS.Close: Set RS = Nothing: MsgBox "Debe realizar el pedido para la planificación teórica de este mes...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
          RS.Close: Set RS = Nothing
          Unload M_MinRea: M_MinRea.Show 1
       End If
     ElseIf NomFor = "MINTEO" Then
        If Not formAbierto("EstTeo") Then
           Dim EstTeo As New M_EstFDi
           EstTeo.lc_Aux1 = "EstTeo"
           EstTeo.Tag = "EstTeo"
           EstTeo.Show 1 ', Partida
           Set EstTeo = Nothing
        End If
     ElseIf NomFor = "MINREA" Then
        If Not formAbierto("EstRea") Then
           Dim EstRea As New M_EstFDi
           EstRea.lc_Aux1 = "EstRea"
           EstRea.Tag = "EstRea"
           EstRea.Show 1 ', Partida
           Set EstRea = Nothing
        End If
     End If

Case 6 'Historico Planificación
    RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda(0).Caption = "": fpLongInteger1(1).Value = "": fpayuda(1).Caption = "": fpLongInteger1(2).Value = "": fpayuda(2).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    vg_codigo = ""
    If NomFor = "MINTEO" Then B_HistPm.LlenarHistPlan "Histórico Planificación Teórica", fpText.text, 1, 1
    If NomFor = "MINREA" Then B_HistPm.LlenarHistPlan "Histórico Planificación Real", fpText.text, 2, 1
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = vg_codregimen
    fpLongInteger1(2).Value = vg_codservicio
    fpDateTime1.text = vg_fecha
    Me.Refresh
Case 8 'salir
'    Unload M_Plami2
    Unload B_Receta
    Me.Hide
    Unload Me
    Unload M_Plami1
End Select
End Sub

'Sub Partidas(tfor As String, NFor As String)
'    Me.Caption = tfor
'    Msgtitulo = tfor
'    NomFor = NFor
'End Sub

Sub ArmarCalendario()
'------- Armar calendario
Dim i As Long, nrosem As Long, diafin As Long
With vaSpread1
    .Row = -1: .Col = -1:
    .BackColor = &H8000000F '&HFFC0C0   '&H80000018
    diafin = fg_mes(Format(fpDateTime1.text, "mm") & Format(fpDateTime1.text, "yyyy"))
    nrosem = 1
    For i = 1 To 6
        For j = 1 To 7
            .Row = i
            .Col = j
            .text = ""
        Next j
    Next i
    For i = 1 To diafin
        Select Case fg_Dia(Format(fpDateTime1.text, "yyyymm") & fg_pone_cero(i, 2))
        Case 1
            .Row = nrosem
            .Col = 7
            .BackColor = &HC0FFFF
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
            nrosem = nrosem + 1
        Case 2
            .Row = nrosem
            .Col = 1
            .BackColor = &HC0FFFF
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
        Case 3
            .Row = nrosem
            .Col = 2
            .BackColor = &HC0FFFF
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
        Case 4
            .Row = nrosem
            .Col = 3
            .BackColor = &HC0FFFF
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
        Case 5
            .Row = nrosem
            .Col = 4
            .BackColor = &HC0FFFF
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
        Case 6
            .Row = nrosem
            .Col = 5
            .BackColor = &HC0FFFF
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
        Case 7
            .Row = nrosem
            .Col = 6
            .BackColor = &HC0FFFF
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
        End Select
    Next i
    .RetainSelBlock = False
End With
End Sub

Sub MoverDatos()
Dim indblo As Boolean, i As Long, j As Long, indcol As Long, sql1 As String
indblo = False
If NomFor = "MINTEO" Then
   RS.Open RutinaLectura.Minutas(3, 0, 0, Val(Format(fpDateTime1.text, "yyyymm")), ""), vg_db, adOpenStatic
   If Not RS.EOF Then indblo = True
   RS.Close: Set RS = Nothing
End If
RS.Open RutinaLectura.Minutas(4, Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1.text, "yyyymm")), IIf(NomFor = "MINTEO", 1, 2)), vg_db, adOpenForwardOnly
If Not RS.EOF Then
   Do While Not RS.EOF
      j = Val(Mid(RS!min_fecmin, 7, 2))
      Select Case fg_Dia(RS!min_fecmin)
      Case 1
          indcol = 7
      Case 2
          indcol = 1
      Case 3
          indcol = 2
      Case 4
          indcol = 3
      Case 5
          indcol = 4
      Case 6
          indcol = 5
      Case 7
          indcol = 6
      End Select
      With vaSpread1
          For i = 1 To 6
              .Row = i
              .Col = indcol
              If Val(.text) = j Then
                 .Col = indcol
    '            .BackColor = IIf(indblo = False, &HC0FFC0, &H8080FF)
                 .BackColor = IIf(indblo = False, &HC0FFFF, &H8080FF)
                 If (CDate(Val(.text) & "/" & Format(fpDateTime1.text, "mm/yyyy")) < Format(Date - IIf((fg_Dia(Format(CDate(Val(vaSpread1.text) & "/" & Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Val(vaSpread1.text) & "/" & Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Or CDate(Val(vaSpread1.text) & "/" & Format(fpDateTime1.text, "mm/yyyy")) < CDate(vg_ciedia)) And NomFor <> "MINTEO" Then
                    .BackColor = &H8080FF
                 End If
                 Exit For
              End If
          Next i
      End With
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing: fg_descarga
End Sub
