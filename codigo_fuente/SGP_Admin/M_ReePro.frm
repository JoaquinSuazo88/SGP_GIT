VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_ReePro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reemplazar ingredientes en recetas"
   ClientHeight    =   6060
   ClientLeft      =   3180
   ClientTop       =   3015
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6060
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   30
      Left            =   5010
      TabIndex        =   14
      Top             =   4770
      Width           =   30
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1035
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   0
      Width           =   10335
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   1920
         _Version        =   196608
         _ExtentX        =   3387
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
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
         MaxLength       =   20
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
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Top             =   585
         Width           =   1920
         _Version        =   196608
         _ExtentX        =   3387
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
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
         MaxLength       =   20
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
      Begin EditLib.fpDoubleSingle fpDouble1 
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   1245
         _Version        =   196608
         _ExtentX        =   2196
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
         BackColor       =   -2147483628
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
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "1"
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
         Index           =   2
         Left            =   4680
         TabIndex        =   3
         Top             =   930
         Width           =   1245
         _Version        =   196608
         _ExtentX        =   2196
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
         BackColor       =   -2147483628
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
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "1"
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
         Index           =   3
         Left            =   7560
         TabIndex        =   4
         Top             =   900
         Width           =   1245
         _Version        =   196608
         _ExtentX        =   2196
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
         BackColor       =   -2147483628
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
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "1"
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
         Caption         =   "% Aprov. Nut."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   6240
         TabIndex        =   20
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "% Aprovechamiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   19
         Top             =   1005
         Width           =   1710
      End
      Begin VB.Label Label3 
         Caption         =   "% Cocción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   3480
         TabIndex        =   18
         Top             =   990
         Width           =   990
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   4290
         TabIndex        =   12
         Top             =   570
         Width           =   5730
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   4290
         TabIndex        =   10
         Top             =   240
         Width           =   5730
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingrediente Origen"
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
         Left            =   180
         TabIndex        =   9
         Top             =   315
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingrediente Destino"
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
         Left            =   180
         TabIndex        =   8
         Top             =   645
         Width           =   1680
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3840
         Picture         =   "M_ReePro.frx":0000
         Top             =   165
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3840
         Picture         =   "M_ReePro.frx":030A
         Top             =   480
         Width           =   480
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   4335
         TabIndex        =   11
         Top             =   285
         Width           =   5730
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   4335
         TabIndex        =   13
         Top             =   615
         Width           =   5730
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4575
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   10980
      _Version        =   393216
      _ExtentX        =   19368
      _ExtentY        =   8070
      _StockProps     =   64
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
      MaxCols         =   13
      MaxRows         =   20
      ProcessTab      =   -1  'True
      RestrictRows    =   -1  'True
      SpreadDesigner  =   "M_ReePro.frx":0614
      UserResize      =   2
      VisibleCols     =   5
      VisibleRows     =   20
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6060
      Left            =   11175
      TabIndex        =   6
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   10689
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   165
      Left            =   120
      TabIndex        =   15
      Top             =   5790
      Visible         =   0   'False
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H80000018&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   7305
      Top             =   5790
      Width           =   300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Recetas Vigentes"
      Height          =   195
      Index           =   0
      Left            =   7665
      TabIndex        =   17
      Top             =   5760
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00D9D9FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   5160
      Top             =   5790
      Width           =   300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Recetas No Vigentes"
      Height          =   195
      Index           =   1
      Left            =   5520
      TabIndex        =   16
      Top             =   5760
      Width           =   1515
   End
End
Attribute VB_Name = "M_ReePro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codproducto As String
Dim CodReceta As Long, nroite As Long
Dim i As Integer, indsel As Integer
Dim canpro As Double, pctapr As Double, pctcoc As Double, pctnut As Double
Dim var1 As Double
Public tipopc As String
Dim MsgTitulo As String

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.Height = 6480
Me.Width = 11805
fg_centra Me
fg_carga ""
Me.HelpContextID = vg_OpcM
Me.HelpContextID = 1093000
vaSpread1.MaxRows = 0: indsel = 0
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 2, 1) = "0", False, True): BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = IIf(tipopc = "Ree%in", False, True): BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 3, 1) = "0", False, True): BtnX.ToolTipText = "Borrar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

If tipopc = "ReeIng" Then
   
   Label3(3).Visible = False
   Label3(0).Visible = False
   Label3(8).Visible = False
   fpDouble1(1).Visible = False
   fpDouble1(2).Visible = False
   fpDouble1(3).Visible = False

ElseIf tipopc = "Ree%in" Then
   
   Label1(0).Visible = False
   fpText1(1).Visible = False
   Image1(1).Visible = False
   fpayuda(1).Visible = False
   lblSOMBRA(1).Visible = False
   Label3(3).Visible = True: Label3(3).Top = 645
   Label3(0).Visible = True: Label3(0).Top = 654
   Label3(8).Visible = True: Label3(8).Top = 654
   fpDouble1(1).Visible = True: fpDouble1(1).Enabled = True: fpDouble1(1).Top = 585
   fpDouble1(2).Visible = True: fpDouble1(2).Enabled = True: fpDouble1(2).Top = 585
   fpDouble1(3).Visible = True: fpDouble1(3).Enabled = True: fpDouble1(3).Top = 585

Else

End If
fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub fpDouble1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub fpText1_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Select Case Index
Case 0
    
    vaSpread1.MaxRows = 0
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
'    Set RS = vg_db.Execute("SELECT DISTINCT ing_nombre FROM b_ingrediente WHERE ing_codigo = '" & fpText1(0).text & "' AND (ing_Indppr = '" & vg_Indppr & "' OR '" & vg_Indppr & "' <> '1')")
    Set RS = vg_db.Execute("sgpadm_Sel_IngredienteRealPropuesta '" & fpText1(0).text & "', " & vg_Indppr & "")
    If RS.EOF Then
       
       RS.Close
       Set RS = Nothing
       fpayuda(0).Caption = ""
       vg_codigo = ""
       Exit Sub
       
    End If
    fpayuda(0).Caption = Trim(RS!ing_nombre)
    RS.Close
    Set RS = Nothing
    MoverDatosGrilla

Case 1
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
'    Set RS = vg_db.Execute("SELECT DISTINCT ing_nombre FROM  b_ingrediente WHERE ing_codigo = '" & fpText1(1).text & "' AND (ing_Indppr = '" & vg_Indppr & "' OR '" & vg_Indppr & "' <> '1')")
    Set RS = vg_db.Execute("sgpadm_Sel_IngredienteRealPropuesta '" & fpText1(1).text & "', " & vg_Indppr & "")
    If RS.EOF Then
       
       RS.Close
       Set RS = Nothing
       fpayuda(1).Caption = ""
       vg_codigo = ""
       Exit Sub
       
    End If
    fpayuda(1).Caption = Trim(RS!ing_nombre)
    RS.Close
    Set RS = Nothing
    If vaSpread1.MaxRows < 1 Then MoverDatosGrilla

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub fpText1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 120
    
    If Index = 0 Then Image1_Click 0
    If Index = 1 Then Image1_Click 1

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Sub MoverDatosGrilla()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim arr As Variant
Dim i As Long
fg_carga ""
vaSpread1.Visible = False
DoEvents
vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.Lock = IIf(Mid(ValidarUsuario(Me), 2, 1) = "0", True, False)
vaSpread1.BackColor = Shape1(0).FillColor

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_s_receta_V07 26, '" & fpText1(0).text & "', '" & IIf(M_Receta.Check2.Value = 1, "x", "") & "', " & vg_filcatdie & ", " & vg_filtippla & ", 0, '" & vg_NUsr & "'")

If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   vaSpread1.Visible = True
   Exit Sub
   
End If

arr = RS.GetRows
RS.Close
Set RS = Nothing
vaSpread1.MaxRows = UBound(arr, 2) + 1
For i = 0 To UBound(arr, 2)
   
   vaSpread1.Row = i + 1
   vaSpread1.Col = -1

   If arr(2, i) <= Val(Format(Date, "yyyymmdd")) And arr(2, i) > 0 Then vaSpread1.BackColor = Shape1(1).FillColor
   
   vaSpread1.Col = 1
   vaSpread1.CellType = CellTypeCheckBox
   vaSpread1.TypeCheckText = " "
   vaSpread1.TypeCheckCenter = True
   vaSpread1.text = "" ' checked
   
   vaSpread1.Col = 2
   vaSpread1.text = "(" & arr(1, i) & ") " & Trim(arr(0, i))
   
   vaSpread1.Col = 3
   vaSpread1.TypeNumberDecPlaces = vg_RDCa
   vaSpread1.text = Format(arr(5, i), fg_Pict(6, vg_RDCa))
   vaSpread1.ForeColor = &HFF0000
   vaSpread1.Lock = IIf(tipopc = "ReeIng", False, True)
   
   vaSpread1.Col = 4
   vaSpread1.text = arr(1, i)
   
   vaSpread1.Col = 5
   vaSpread1.text = arr(4, i)
   
   vaSpread1.Col = 6
   vaSpread1.text = Format(arr(6, i), fg_Pict(6, vg_RDCa))
   vaSpread1.ForeColor = &HFF0000
   vaSpread1.Lock = IIf(tipopc = "ReeIng", False, True)
   
   vaSpread1.Col = 7
   vaSpread1.text = Format(arr(7, i), fg_Pict(6, vg_RDCa))
   vaSpread1.ForeColor = &HFF0000
   vaSpread1.Lock = IIf(tipopc = "ReeIng", False, True)
   
   vaSpread1.Col = 8
   vaSpread1.TypeNumberDecPlaces = vg_RDCa
   vaSpread1.TypeHAlign = TypeHAlignRight
   vaSpread1.text = Format(((((arr(5, i) * arr(6, i)) / 100) * arr(7, i)) / 100), fg_Pict(6, vg_RDCa))
   
   vaSpread1.Col = 9
   vaSpread1.text = arr(8, i)
   vaSpread1.ForeColor = &HFF0000
   vaSpread1.Lock = IIf(tipopc = "ReeIng", False, True)
   
   vaSpread1.Col = 10
   vaSpread1.TypeNumberDecPlaces = vg_RDCa
   vaSpread1.TypeHAlign = TypeHAlignRight
   vaSpread1.text = Format(((arr(8, i) / 100) * arr(5, i)), fg_Pict(6, vg_RDCa))
   
   vaSpread1.Col = 11
   vaSpread1.text = arr(4, i)
   
   vaSpread1.Col = 12
   vaSpread1.text = arr(9, i)

   vaSpread1.Col = 13
   vaSpread1.TypeNumberDecPlaces = vg_RDCa
   vaSpread1.TypeHAlign = TypeHAlignRight
   vaSpread1.text = Format((((arr(5, i) * arr(6, i)) / 100)), fg_Pict(6, vg_RDCa))

Next i

'-------> Traer % de ingrediente
If tipopc = "Ree%in" Then

    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS = vg_db.Execute("select ing_pctapr, ing_pctcoc, ing_pctnut from b_ingrediente With(NoLock) where ing_codigo='" & fpText1(0).text & "'")
    
    If Not RS.EOF Then
       
       fpDouble1(1).text = IIf(IsNull(RS!ing_pctapr), 0, RS!ing_pctapr)
       fpDouble1(2).text = IIf(IsNull(RS!ing_pctcoc), 0, RS!ing_pctcoc)
       fpDouble1(3).text = IIf(IsNull(RS!ing_pctnut), 0, RS!ing_pctnut)
    
    End If
    RS.Close
    Set RS = Nothing

End If
vaSpread1.Visible = True
fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0
    
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(0).Left + 1770
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "AgregarIng"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    vaSpread1.MaxRows = 0
    fpText1(0).text = vg_codigo
    fpayuda(0).Caption = vg_nombre

Case 1
    
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(1).Left + 1770
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "AgregarIng"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    fpText1(1).text = vg_codigo
    fpayuda(1).Caption = vg_nombre

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim CodIng As String, codpra As String, codpro As String, codume As String
Dim caning  As Double, pctapr As Double, pctcoc As Double, pctnut As Double, canser As Double, cannet As Double, valuni As Double
Dim IndentificadorIngSumaTablaGramaje As String

Select Case Button.Index

Case 1, 3
    
    If vaSpread1.MaxRows < 1 Then Exit Sub
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
'    RS.Open "SELECT ing_nombre FROM b_ingrediente With(NoLock) WHERE ing_codigo='" & fpText1(0).text & "'", vg_db, adOpenStatic
    Set RS = vg_db.Execute("sgpadm_Sel_Ingrediente '" & fpText1(0).text & "'")
    If RS.EOF Then
       
       RS.Close
       Set RS = Nothing
       vaSpread1.MaxRows = 0
       MsgBox "No Existe Ingredientes", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente"
       Exit Sub
       
    End If
    RS.Close
    Set RS = Nothing
    
    If fpText1(1).text <> "" Then
       
       If vaSpread1.MaxRows < 1 Then Exit Sub
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
'       RS.Open "SELECT ing_nombre FROM b_ingrediente With(NoLock) WHERE ing_codigo='" & fpText1(1).text & "'", vg_db, adOpenStatic
       Set RS = vg_db.Execute("sgpadm_Sel_Ingrediente '" & fpText1(1).text & "'")
       If RS.EOF Then
       
          RS.Close
          Set RS = Nothing
          vaSpread1.MaxRows = 0
          MsgBox "No Existe Ingredientes a Reemplazar", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente"
          Exit Sub
          
       End If
       RS.Close
       Set RS = Nothing
    
    End If
    
    '-------> Validar que exista los ingredientes origenes
'    RS.Open "SELECT DISTINCT a.rec_nombre, a.rec_codigo, b.red_codpro, b.red_nroite, b.red_canpro, " & _
'            "b.red_pctapr, b.red_pctcoc, b.red_pctnut " & _
'            "FROM  b_receta a With(NoLock), b_recetadet b With(NoLock) " & _
'            "WHERE b.red_codigo = a.rec_codigo " & _
'            "AND   b.red_codpro = '" & fpText1(0).text & "' " & _
'            "AND  (a.rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
'            "AND  (a.rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) " & _
'            "AND   a.rec_tiprec = '0'", vg_db, adOpenStatic
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS = vg_db.Execute("sgpadm_Sel_ValidarSiExistenIngOrigen '" & fpText1(0).text & "', " & vg_filcatdie & "," & vg_filtippla & "")
    If RS.EOF Then
       
       RS.Close
       Set RS = Nothing
       vaSpread1.MaxRows = 0
       MsgBox "No Existe Ingredientes Origen en Recetario", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente"
       Exit Sub
       
    End If
    RS.Close
    Set RS = Nothing
    
    '-------> Validar que este seleccionar a lo menos una receta de la lista
    indsel = 0
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i: vaSpread1.Col = 1
        If vaSpread1.text = "1" Then indsel = 1: Exit For
    
    Next i
    
    If Button.Index = 1 Then
       
       If indsel = 0 Then MsgBox "Seleccione Uno o Más Recetas a Reemplazar", vbCritical + vbOKOnly, "Cambio Ingrediente": Exit Sub
       If fpText1(1).text <> "" Then
          
          msg = " Esta Seguro Reemplazar " & "(" & Trim(fpayuda(0).Caption) & ")" & " Por " & "(" & Trim(fpayuda(1).Caption) & ")" & " En Las Recetas Seleccionadas ?"
       
       Else
          
          msg = " Esta Seguro Remplazar Datos en " & "(" & Trim(fpayuda(0).Caption) & ")" & " "
       
       End If
    
    ElseIf Button.Index = 3 Then
       
       If indsel = 0 Then MsgBox "Seleccione Uno o Más Recetas a Eliminar", vbCritical + vbOKOnly, "Eliminar Ingrediente": Exit Sub
       msg = " Esta Seguro Eliminar " & "(" & Trim(fpayuda(0).Caption) & ")" & " En Las Recetas Seleccionadas ?"
    
    End If
    If MsgBox("Esta Seguro ?", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    fg_carga ""
    Bar1.Visible = True: Bar1.Value = 0
    
    For i = 1 To vaSpread1.MaxRows
        Bar1.Value = Val((i / vaSpread1.MaxRows) * 100)
        vaSpread1.Row = i
        vaSpread1.Col = 1
        
        If vaSpread1.text = "1" Then
           
           M_ReePro.Refresh
           DoEvents
           vaSpread1.Col = 3
           canpro = 0
           canpro = vaSpread1.text
           caning = vaSpread1.text
           
           vaSpread1.Col = 4
           CodReceta = 0
           CodReceta = vaSpread1.text
           vaSpread1.Col = 6
           pctapr = 0
           pctapr = IIf(tipopc = "ReeIng", vaSpread1.text, fpDouble1(1).text)
           
           vaSpread1.Col = 7
           pctcoc = 0
           pctcoc = IIf(tipopc = "ReeIng", vaSpread1.text, fpDouble1(2).text)
           
           vaSpread1.Col = 9
           pctnut = 0
           pctnut = IIf(tipopc = "ReeIng", vaSpread1.text, fpDouble1(3).text)
           
           vaSpread1.Col = 11
           nroite = 0
           nroite = vaSpread1.text
           
           vaSpread1.Col = 12
           IndentificadorIngSumaTablaGramaje = "0"
           IndentificadorIngSumaTablaGramaje = vaSpread1.text
           canser = 0: cannut = 0
           If Button.Index = 1 Then
              
              If fpText1(1).text <> "" Then
                 
                 codproducto = fpText1(1).text
              
              Else
                 
                 codproducto = fpText1(0).text
              
              End If
              
              vg_db.Execute "sgpadm_iu_recetadet 'M1' , " & CodReceta & ", " & nroite & ", '" & fpText1(0).text & "', " & canpro & ", 0, " & _
                            "" & pctapr & ", " & pctcoc & ", " & pctnut & ", '" & codproducto & "', '" & IndentificadorIngSumaTablaGramaje & "'"
              
              If tipopc = "ReeIng" Then
                 
                 vg_db.Execute "sgpadm_Upd_TablaGramajeReceta " & CodReceta & ", '" & fpText1(0).text & "', '" & codproducto & "'"
              
              End If

           ElseIf Button.Index = 3 Then
              
              vg_db.Execute "DELETE b_recetadet FROM b_recetadet " & _
                            "WHERE red_codigo = " & CodReceta & " " & _
                            "AND   red_codpro = '" & Trim(fpText1(0).text) & "' " & _
                            "AND   red_nroite = " & nroite & ""
              
              '------->
              vg_db.Execute "sgpadm_Del_TablaGramajeIngReceta " & CodReceta & ", '" & Trim(fpText1(0).text) & "'"
           
           End If
        
        End If
    
    Next i
    
    Bar1.Visible = False
    fg_descarga
    
    If Button.Index = 1 Then
       
       MsgBox "Actualización [OK]", vbInformation + vbOKOnly, MsgTitulo
    
    ElseIf Button.Index = 3 Then
       
       MsgBox "Eliminación de ingrediente finalizo sin problema", vbInformation + vbOKOnly, "Eliminar ingrediente en receta"
    
    End If
    
    indsel = 0
    MoverDatosGrilla

Case 5
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
If Col = 1 And Row = 0 Then
   
   If indsel = 0 Then
      
      For i = 1 To vaSpread1.MaxRows
          
          vaSpread1.Row = i
          vaSpread1.Col = 1
          vaSpread1.CellType = CellTypeCheckBox
          vaSpread1.TypeCheckText = ""
          vaSpread1.TypeCheckCenter = True
          vaSpread1.Value = "1" ' checked
      
      Next i
      indsel = 1
   
   Else
      
      For i = 1 To vaSpread1.MaxRows
          
          vaSpread1.Row = i
          vaSpread1.Col = 1
          vaSpread1.CellType = CellTypeCheckBox
          vaSpread1.TypeCheckText = " "
          vaSpread1.TypeCheckCenter = True
          vaSpread1.Value = "" ' checked
      
      Next i
      indsel = 0
   
   End If

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

On Error GoTo Man_Error

Select Case Col

Case 3, 6, 7, 9
    
    vaSpread1.Row = Row
    vaSpread1.Col = Col
    If ChangeMade = False Then var1 = Val(vaSpread1.Value) Else If Val(vaSpread1.Value) <= 0 Then vaSpread1.text = var1

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
'------- Calcular Gramaje Neto
pctnut = 0
canpro = 0
pctapr = 0
pctcoc = 0
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 3
canpro = vaSpread1.text

vaSpread1.Col = 9
pctnut = vaSpread1.text

vaSpread1.Col = 10
vaSpread1.CellType = CellTypeStaticText
vaSpread1.TypeHAlign = TypeHAlignRight
vaSpread1.text = Format(CCur((pctnut / 100) * canpro), fg_Pict(6, vg_RDCa))

'------- Calcular % Limpieza & Cocción
vaSpread1.Col = 6
pctapr = vaSpread1.text

'cantservida = CCur((paporv / 100) * canpro)
vaSpread1.Col = 7
pctcoc = vaSpread1.text

'cantservida = CCur((pcoccion / 100) * cantservida)
vaSpread1.Col = 8
vaSpread1.CellType = CellTypeStaticText
vaSpread1.TypeHAlign = TypeHAlignRight
vaSpread1.text = Format(CCur(((pctapr / 100) * canpro) * (pctcoc / 100)), fg_Pict(6, vg_RDCa))

vaSpread1.Col = 13
vaSpread1.CellType = CellTypeStaticText
vaSpread1.TypeHAlign = TypeHAlignRight
vaSpread1.text = Format(CCur((pctapr / 100) * canpro), fg_Pict(6, vg_RDCa))

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Sub MoverDatosIniciales(op As String)

On Error GoTo Man_Error

tipopc = op
Me.Caption = IIf(op = "ReeIng", "Reemplazar Ingrediente Receta", "Reemplazar % Ingrediente")
MsgTitulo = IIf(op = "ReeIng", "Reemplazar Ingrediente Receta", "Reemplazar % Ingrediente")

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub
