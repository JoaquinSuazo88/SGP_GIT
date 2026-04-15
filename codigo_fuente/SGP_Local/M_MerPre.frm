VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_MerPre 
   Caption         =   "Raciones no Vendidas"
   ClientHeight    =   8625
   ClientLeft      =   2445
   ClientTop       =   2190
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Mermas Kilos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   27
      Top             =   7320
      Width           =   12255
      Begin EditLib.fpDoubleSingle Desconche 
         Height          =   375
         Left            =   5400
         TabIndex        =   7
         Top             =   360
         Width           =   1815
         _Version        =   196608
         _ExtentX        =   3201
         _ExtentY        =   661
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
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
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
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "0"
         DecimalPlaces   =   3
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "0"
         NegFormat       =   0
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDoubleSingle Pan 
         Height          =   375
         Left            =   9240
         TabIndex        =   8
         Top             =   360
         Width           =   1815
         _Version        =   196608
         _ExtentX        =   3201
         _ExtentY        =   661
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
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
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
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "0"
         DecimalPlaces   =   3
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "0"
         NegFormat       =   0
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDoubleSingle Produccion 
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1815
         _Version        =   196608
         _ExtentX        =   3201
         _ExtentY        =   661
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
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
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
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "0"
         DecimalPlaces   =   3
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "0"
         NegFormat       =   0
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pan"
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
         Left            =   8280
         TabIndex        =   30
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desconche"
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
         Left            =   4080
         TabIndex        =   29
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Produción "
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
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5415
      Left            =   150
      TabIndex        =   20
      Top             =   1680
      Width           =   12255
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4605
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   12015
         _Version        =   393216
         _ExtentX        =   21193
         _ExtentY        =   8123
         _StockProps     =   64
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
         MaxCols         =   10
         MaxRows         =   30
         SpreadDesigner  =   "M_MerPre.frx":0000
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Día Bloqueado"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   26
         Top             =   5040
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   120
         Top             =   5070
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Día Habilitado"
         Height          =   195
         Index           =   0
         Left            =   2295
         TabIndex        =   25
         Top             =   5040
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   1935
         Top             =   5070
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   3
         Left            =   11640
         TabIndex        =   24
         Top             =   5040
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   2
         Left            =   8640
         TabIndex        =   23
         Top             =   5040
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Totales"
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
         Left            =   6720
         TabIndex        =   22
         Top             =   5040
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   30
      TabIndex        =   9
      Top             =   480
      Width           =   12345
      Begin VB.CheckBox ChcMerma 
         Caption         =   "No considera Mermas"
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
         Left            =   8880
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   6735
         TabIndex        =   1
         Top             =   150
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
         Index           =   1
         Left            =   1215
         TabIndex        =   2
         Top             =   600
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
         Left            =   1215
         TabIndex        =   0
         Top             =   180
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
         Index           =   0
         Left            =   6735
         TabIndex        =   3
         Top             =   600
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
         Text            =   "17/08/2023"
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   11760
         TabIndex        =   31
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cargar Información"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2610
         TabIndex        =   16
         Top             =   180
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Left            =   5940
         TabIndex        =   15
         Top             =   675
         Width           =   540
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
         Left            =   60
         TabIndex        =   14
         Top             =   675
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
         Left            =   5940
         TabIndex        =   13
         Top             =   225
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
         Left            =   60
         TabIndex        =   12
         Top             =   255
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2085
         Picture         =   "M_MerPre.frx":070A
         Top             =   90
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   7605
         Picture         =   "M_MerPre.frx":0A14
         Top             =   60
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2085
         Picture         =   "M_MerPre.frx":0D1E
         Top             =   510
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   8130
         TabIndex        =   11
         Top             =   150
         Width           =   3135
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2610
         TabIndex        =   10
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   8175
         TabIndex        =   18
         Top             =   195
         Width           =   3135
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   2655
         TabIndex        =   17
         Top             =   225
         Width           =   3135
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   2655
         TabIndex        =   19
         Top             =   645
         Width           =   3135
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   12540
      _ExtentX        =   22119
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
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
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_MerPre.frx":1028
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_MerPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim est As Boolean
Dim modo As String

Private Sub ChcMerma_Click()

On Error GoTo Man_Error

Dim i       As Long
Dim mensaje As String

If est Then Exit Sub

mensaje = IIf(vaSpread1.MaxRows > 0, "Esta seguro activar esta opción, movera a cero los registro ingresado...", "Esta seguro activar esta opción")

If ChcMerma.Value = 1 Then

    
    If MsgBox(mensaje, vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
    
       ChcMerma.Value = 0
       
       Exit Sub
    
    End If
    
    For i = 1 To vaSpread1.MaxRows
    
        vaSpread1.Row = i
        
        vaSpread1.Col = 6
        vaSpread1.text = ""
        vaSpread1.Lock = True
        
        vaSpread1.Col = 7
        vaSpread1.text = ""
        vaSpread1.Lock = True
        
        vaSpread1.Col = 8
        vaSpread1.text = ""
        
        vaSpread1.Col = 10
        vaSpread1.text = ""
        
        
    Next i
    
    Desconche.Value = 0
    Desconche.Enabled = False
    
    Pan.Value = 0
    Pan.Enabled = False
    
    Produccion.Value = 0
    Produccion.Enabled = False

    If modo = "" Then
    
       modo = "M"
       
    End If
    
    Gl_Ac_Botones Me, 1, 0, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    
    fpLongInteger1(0).Enabled = False
    fpayuda(1).Enabled = False
    
    fpLongInteger1(1).Enabled = False
    fpayuda(2).Enabled = False
    
    fpDateTime1(0).Enabled = False

Else

    For i = 1 To vaSpread1.MaxRows
    
        vaSpread1.Row = i
        
        vaSpread1.Col = 6
'        vaSpread1.text = ""
        vaSpread1.Lock = False
        
        vaSpread1.Col = 7
'        vaSpread1.text = ""
        vaSpread1.Lock = False
        
        vaSpread1.Col = 8
'        vaSpread1.text = ""
        
        
    Next i
    
    Desconche.Value = 0
    Desconche.Enabled = True
    
    Pan.Value = 0
    Pan.Enabled = True
    
    Produccion.Value = 0
    Produccion.Enabled = True

    If modo = "" Then
    
       modo = "M"
       
    End If
    
    Gl_Ac_Botones Me, 1, 0, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    
    fpLongInteger1(0).Enabled = True
    fpayuda(1).Enabled = True
    
    fpLongInteger1(1).Enabled = True
    fpayuda(2).Enabled = True
    
    fpDateTime1(0).Enabled = True

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Desconche_Change()

On Error GoTo Man_Error
    
If est Then Exit Sub
    
    If modo = "" Then
    
       modo = "M"
       
    End If
    
    Gl_Ac_Botones Me, 1, 0, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Desconche_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If est Then Exit Sub

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga
TraerFechaCierre

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_carga ""

est = True

MsgTitulo = "Raciones no Vendidas"
Me.HelpContextID = vg_OpcM
Me.Height = 9210
Me.Width = 12780
fg_centra Me
modo = ""

Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 3, modo
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = False

Let fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")

fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
vaSpread1.MaxRows = 0
ChcMerma.Enabled = False

est = False
fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpDateTime1_Change(Index As Integer)

On Error GoTo Man_Error

If est Then Exit Sub
If IsDate(fpDateTime1(0).text) = False Then Exit Sub
'Mover_Datos
'est = False
vaSpread1.MaxRows = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If est Then Exit Sub
Select Case Index

Case 0
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    
    If RS.EOF Then
       
       RS.Close
       Set RS = Nothing
       fpayuda(1).Caption = ""
    
    Else
       
       fpayuda(1).Caption = Trim(RS!reg_nombre)
       RS.Close
       Set RS = Nothing
    
    End If
'    Mover_Datos
    vaSpread1.MaxRows = 0
Case 1
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(1).Value) & " AND ser_activo = '1'", vg_db, adOpenStatic
    If RS.EOF Then
       RS.Close: Set RS = Nothing
       fpayuda(2).Caption = ""
    Else
       fpayuda(2).Caption = Trim(RS!ser_nombre)
       RS.Close: Set RS = Nothing
    End If
'    Mover_Datos
    vaSpread1.MaxRows = 0
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpText_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If est Then Exit Sub
vaSpread1(1).MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & fpText.text & "' AND cli_tipo = 0", vg_db, adOpenStatic

If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fpayuda(0).Caption = ""
   fpLongInteger1(0).Value = "": fpayuda(1).Caption = ""
   fpLongInteger1(1).Value = "": fpayuda(2).Caption = ""

Else
   
   fpayuda(0).Caption = Trim(RS!cli_nombre)
   RS.Close
   Set RS = Nothing

End If

fpLongInteger1(0).Value = "": fpayuda(1).Caption = ""
fpLongInteger1(1).Value = "": fpayuda(2).Caption = ""
Mover_Datos

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0
    
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.text = vg_codigo
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
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Ser"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre
    fpDateTime1(0).SetFocus

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Sub Mover_Datos()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim vTotCos As Double, vTotMer As Double, cosali As Double, CosDes As Double

modo = ""
Gl_Ac_Botones Me, 1, 3, modo

If Trim(LimpiaDato(fpText.text)) = "" Or Val(fpLongInteger1(0).Value) = 0 Or Val(fpLongInteger1(1).Value) = 0 Or Trim(fpDateTime1(0).text) = "" Then

   Exit Sub
   
End If

fg_carga ""

'-------> Validar si existe minuta real
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_ValidarMinutaMermaPorPreparacion '" & LimpiaDato(fpText.text) & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Format(fpDateTime1(0).text, "yyyymmdd") & "")

If RS.EOF Then

   fg_descarga
   MsgBox "No existe minuta real, para ese dia", vbInformation + vbOKOnly, MsgTitulo
   Exit Sub

Else

'   ChcMerma.Enabled = True

End If
RS.Close
Set RS = Nothing


With vaSpread1
    
    .Visible = False
    '-------> Validar dia bloqueado
    .Row = -1
    .Col = -1
    .BackColor = IIf(CDate(fpDateTime1(0).text) < CDate(vg_ciedia), Shape1(1).FillColor, Shape1(0).FillColor)
    .MaxRows = 0
    vTotCos = 0
    vTotMer = 0
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS = vg_db.Execute("sgp_Sel_MermaPorPreparacion '" & LimpiaDato(fpText.text) & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Format(fpDateTime1(0).text, "yyyymmdd") & ", " & vg_codbod & "")

    If Not RS.EOF Then
       
       est = True
       
       ChcMerma.Value = IIf(IsNull(RS!Considera_Merma) Or RS!Considera_Merma = "0", 0, 1)
       ChcMerma.Enabled = IIf(CDate(fpDateTime1(0).text) < CDate(vg_ciedia), False, True)
       
       Desconche.Value = RS!Merma_Desconche
       Desconche.Enabled = IIf(ChcMerma.Value = 1 Or CDate(fpDateTime1(0).text) < CDate(vg_ciedia), False, True)
       
       Pan.Value = RS!Merma_Pan
       Pan.Enabled = IIf(ChcMerma.Value = 1 Or CDate(fpDateTime1(0).text) < CDate(vg_ciedia), False, True)
       
       Produccion.Value = RS!Merma_produccion
       Produccion.Enabled = IIf(ChcMerma.Value = 1 Or CDate(fpDateTime1(0).text) < CDate(vg_ciedia), False, True)
       
       est = False
       
       'Aviso que no ha realizado salida producción.
       If Trim(RS!tov_rutcli) = "" Then
          
          MsgBox "No ha realizado la salida de producción ...", vbInformation + vbOKOnly, MsgTitulo
       
       End If
       
       Do While Not RS.EOF
          
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          
          .Col = 1
          .TypeHAlign = TypeHAlignRight
          .text = IIf(IsNull(RS!rec_codigo), "", RS!rec_codigo)
          
          .Col = 2
          .TypeHAlign = TypeHAlignLeft
          .text = IIf(IsNull(RS!rec_nombre), "", RS!rec_nombre)
          
          .Col = 3
          .TypeHAlign = TypeHAlignRight
          .text = IIf(IsNull(RS!mid_numrac), "", RS!mid_numrac)
          
          cosali = IIf(IsNull(RS!mid_cosrec), 0, Format(RS!mid_cosrec, fg_Pict(6, 2)))
          CosDes = IIf(IsNull(RS!mid_cosdes), 0, Format(RS!mid_cosdes, fg_Pict(6, 2)))
          
          .Col = 4
          .TypeHAlign = TypeHAlignRight
          .text = Format((cosali + CosDes), fg_Pict(6, 2))
    
          .Col = 5
          .TypeHAlign = TypeHAlignRight
          .text = Format(IIf(IsNull(RS!mid_numrac), 0, ((cosali + CosDes) * RS!mid_numrac)), fg_Pict(6, 0))
    
          vTotCos = Round(vTotCos + ((cosali + CosDes) * RS!mid_numrac), 0)
          
          .Col = 6
          
          If CDec(RS!mid_nummer) > 0 Then
          
             .CellType = CellTypeNumber
             .TypeNumberSeparator = vg_CSep
             .TypeNumberDecimal = vg_CDec
             .TypeNumberDecPlaces = vg_DCa
          
             .TypeHAlign = TypeHAlignRight
      '      .Lock = IIf(CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")) < Format(Date - IIf((fg_Dia(Format(CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy"), True, False)
             .Lock = IIf(CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Or ChcMerma.Value = 1, True, False)
             .text = IIf(IsNull(RS!mid_nummer) Or RS!mid_nummer = 0, "", RS!mid_nummer)
          
          Else
          
             .CellType = CellTypeNumber
'            .TypeNumberSeparator = vg_CSep
'             .TypeNumberDecimal = vg_CDec
             .TypeNumberDecPlaces = 0
             .TypeHAlign = TypeHAlignRight
             .TypeNumberMin = 0
             .TypeNumberMax = IIf(IsNull(RS!mid_numrac) Or Trim(RS!mid_numrac) = "" Or RS!mid_numrac = 0, 999999, RS!mid_numrac)
             .TypeSpin = False
             .TypeIntegerSpinInc = 1
             .TypeIntegerSpinWrap = False
 
      '      .Lock = IIf(CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")) < Format(Date - IIf((fg_Dia(Format(CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy"), True, False)
             .Lock = IIf(CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Or ChcMerma.Value = 1, True, False)
             .text = Format(IIf(IsNull(RS!mid_nummer) Or RS!mid_nummer = 0, "", RS!mid_nummer), fg_Pict(6, 0))
                   
          End If
          
'          .TypeHAlign = TypeHAlignRight
'    '      .Lock = IIf(CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")) < Format(Date - IIf((fg_Dia(Format(CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy"), True, False)
'          .Lock = IIf(CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Or ChcMerma.Value = 1, True, False)
'          .text = IIf(IsNull(RS!mid_nummer) Or RS!mid_nummer = 0, "", RS!mid_nummer)
          
          .Col = 7
          .CellType = CellTypeNumber
          .TypeNumberSeparator = vg_CSep
          .TypeNumberDecimal = vg_CDec
          .TypeNumberDecPlaces = 4 '3
          .TypeHAlign = TypeHAlignRight
          .TypeNumberMin = 0
          .TypeNumberMax = 999999
          .TypeSpin = False
          .TypeIntegerSpinInc = 1
          .TypeIntegerSpinWrap = False
          
          .TypeHAlign = TypeHAlignRight
          .Lock = IIf(CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Or ChcMerma.Value = 1, True, False)
          .text = IIf(IsNull(RS!CantServida) Or RS!CantServida = 0, "", RS!CantServida)
          
          .Col = 8
          .CellType = CellTypeNumber
          .Lock = True
          .TypeHAlign = TypeHAlignRight
          .text = ""
          
          If (cosali > 0 Or CosDes) And RS!mid_nummer > 0 Then
             
             .text = Format(IIf(IsNull(RS!mid_nummer), "", ((cosali + CosDes) * RS!mid_nummer)), fg_Pict(6, 4)) '0))
          
          End If
          
          vTotMer = Round(vTotMer + ((cosali + CosDes) * IIf(IsNull(RS!mid_nummer), 0, RS!mid_nummer)), 4) '0)
          
          .Col = 9
          .text = RS!mid_numlin
          
          .Col = 10
          .CellType = CellTypeNumber
          .TypeNumberSeparator = vg_CSep
          .TypeNumberDecimal = vg_CDec
          .TypeNumberDecPlaces = 4 '3
          .TypeHAlign = TypeHAlignRight
          .TypeNumberMin = 0
          .TypeNumberMax = 999999
          .TypeSpin = False
          .TypeIntegerSpinInc = 1
          .TypeIntegerSpinWrap = False
          
          .TypeHAlign = TypeHAlignRight
          .Lock = IIf(CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Or ChcMerma.Value = 1, True, False)
          .text = IIf(IsNull(RS!CantBruta) Or RS!CantBruta = 0, "", RS!CantBruta)
          
          RS.MoveNext
       
       Loop
       
       modo = ""
       Gl_Ac_Botones Me, 1, IIf(CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")) < Format(Date - IIf((fg_Dia(Format(CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Format(fpDateTime1(0).text, "dd/mm/yyyy")), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy"), 6, 1), modo
       Toolbar1.Buttons(1).Visible = False
       Toolbar1.Buttons(2).Visible = False
    
    Else
       
       Gl_Ac_Botones Me, 1, 6, modo
       Toolbar1.Buttons(1).Visible = False
       Toolbar1.Buttons(2).Visible = False
    
       est = True
       
       ChcMerma.Enabled = False
       
       Produccion.Enabled = False
       Pan.Enabled = False
       Desconche.Enabled = False
       
       est = False
       
       MsgBox "No tiene registrado comensales del dia en la planificación real...", vbInformation + vbOKOnly, MsgTitulo
    
    End If
    
    Label1(2).Caption = Format(vTotCos, fg_Pict(6, 4)) '0))
    Label1(3).Caption = Format(vTotMer, fg_Pict(6, 4)) '0))
    RS.Close
    Set RS = Nothing
    
    .Visible = True

End With

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Pan_Change()

On Error GoTo Man_Error
    
If est Then Exit Sub
    
    If modo = "" Then
    
       modo = "M"
       
    End If
    
    Gl_Ac_Botones Me, 1, 0, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Pan_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If est Then Exit Sub

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Produccion_Change()

On Error GoTo Man_Error
    
If est Then Exit Sub
    
    If modo = "" Then
    
       modo = "M"
       
    End If
    
    Gl_Ac_Botones Me, 1, 0, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Produccion_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If est Then Exit Sub

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim MyBuffer    As String
Dim RS          As New ADODB.Recordset
Dim vCodReceta  As Long
Dim vNumMerma   As Double
Dim vMermaxOtro As Double
Dim vMermaxSer  As Double
Dim vNumLin     As Long
Dim i           As Long

Select Case Button.Index

Case 3 'modificar
    
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False

Case 5 'eliminar
    
    If Not est < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<GrabaMerma>"
    
    For i = 1 To vaSpread1.MaxRows
    
        vaSpread1.Row = i
        
        vaSpread1.Col = 1
        vCodReceta = vaSpread1.text
    
        vNumMerma = 0
        vMermaxOtro = 0
        vMermaxSer = 0
              
        vaSpread1.Col = 9
        vNumLin = vaSpread1.text
        
        MyBuffer = MyBuffer & " <Merma"
        MyBuffer = MyBuffer & " CR = " & Chr(34) & vCodReceta & Chr(34)
        MyBuffer = MyBuffer & " NM = " & Chr(34) & vNumMerma & Chr(34)
        MyBuffer = MyBuffer & " MO = " & Chr(34) & vMermaxOtro & Chr(34)
        MyBuffer = MyBuffer & " MS = " & Chr(34) & vMermaxSer & Chr(34)
        MyBuffer = MyBuffer & " NL = " & Chr(34) & vNumLin & Chr(34)
      
        MyBuffer = MyBuffer & "/>"
        
    Next i
    
    MyBuffer = MyBuffer & "</GrabaMerma>"
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgp_Upd_XmlMermaPreparacion '" & MyBuffer & "', '" & LimpiaDato(fpText.text) & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Format(fpDateTime1(0).text, "yyyymmdd") & ", '" & IIf(ChcMerma.Value = 1, 1, 0) & "', " & Desconche.Value & " , " & Pan.Value & ", " & Produccion.Value & ", '" & vg_NUsr & "'")
    
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
       
       Else
       
          MsgBox "Proceso Finalizo Correctamente", vbInformation + vbOKOnly, MsgTitulo
       
       End If
    
    End If
    RS.Close
    Set RS = Nothing

    Mover_Datos
    modo = ""
    Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Frame1.Enabled = True

Case 7 'actualizar
    
    Mover_Datos

Case 10
    
    If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
    
       Exit Sub
    
    End If
    
    Mover_Datos
    modo = "": Gl_Ac_Botones Me, 1, 3, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    
    fpLongInteger1(0).Enabled = True
    fpayuda(1).Enabled = True
    
    fpLongInteger1(1).Enabled = True
    fpayuda(2).Enabled = True
    
    fpDateTime1(0).Enabled = True
    
    Frame1.Enabled = True

Case 12
       
    With vaSpread1
        
        '-------> Confirmar cantidades ingresadas
        
        For i = 1 To .MaxRows
        
            .Row = i
        
            .Col = 6
            vNumMerma = 0
            vNumMerma = IIf(Trim(.text) = "", 0, .text)
            
            .Col = 7
            vMermaxOtro = 0
            vMermaxOtro = IIf(Trim(.text) = "", 0, .text)
        
            If vNumMerma = 0 And vMermaxOtro > 0 Then
                          
               MsgBox "Merma x kilo no fue confirmada, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo
               
               .Row = i
               .Col = 6
               .SetActiveCell 7, i
               .SetFocus
               
               Exit Sub
               
            End If
            
            If vNumMerma > 0 And vMermaxOtro = 0 Then
            
               MsgBox "Merma x Raciones no fue confirmada, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo
            
               .Row = i
               .Col = 7
               .SetActiveCell 6, i
               .SetFocus
               
               Exit Sub
               
            End If
            
        Next i
        
        Let MyBuffer = ""
        Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        Let MyBuffer = MyBuffer & "<GrabaMerma>"
        
        For i = 1 To .MaxRows
            
            .Row = i
            
            .Col = 1
            vCodReceta = .text
    
            .Col = 6
            vNumMerma = IIf(Trim(.text) = "", 0, .text)
            
            .Col = 7
            vMermaxSer = IIf(Trim(.text) = "", 0, .text)
              
            .Col = 9
            vNumLin = .text
        
            .Col = 10
            vMermaxOtro = IIf(Trim(.text) = "", 0, .text)
            
            MyBuffer = MyBuffer & " <Merma"
            MyBuffer = MyBuffer & " CR = " & Chr(34) & vCodReceta & Chr(34)
            MyBuffer = MyBuffer & " NM = " & Chr(34) & vNumMerma & Chr(34)
            MyBuffer = MyBuffer & " MO = " & Chr(34) & vMermaxOtro & Chr(34)
            MyBuffer = MyBuffer & " MS = " & Chr(34) & vMermaxSer & Chr(34)
            MyBuffer = MyBuffer & " NL = " & Chr(34) & vNumLin & Chr(34)
      
            MyBuffer = MyBuffer & "/>"
        
        Next i
    
        MyBuffer = MyBuffer & "</GrabaMerma>"
    
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Set RS = vg_db.Execute("sgp_Upd_XmlMermaPreparacion '" & MyBuffer & "', '" & LimpiaDato(fpText.text) & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Format(fpDateTime1(0).text, "yyyymmdd") & ", '" & IIf(ChcMerma.Value = 1, 1, 0) & "', " & Desconche.Value & " , " & Pan.Value & ", " & Produccion.Value & ", '" & vg_NUsr & "'")
    
        If Not RS.EOF Then
       
           If RS(0) > 0 Then
          
              MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
           
           Else
       
              MsgBox "Proceso Finalizo Correctamente", vbInformation + vbOKOnly, MsgTitulo
           
           End If
    
        End If
        RS.Close
        Set RS = Nothing
    
    End With
    
    modo = ""
    Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    
    fpLongInteger1(0).Enabled = True
    fpayuda(1).Enabled = True
    
    fpLongInteger1(1).Enabled = True
    fpayuda(2).Enabled = True
    
    fpDateTime1(0).Enabled = True
    
    Frame1.Enabled = True
    
'    Desconche.Enabled = True
'    Pan.Enabled = True
'    Produccion.Enabled = True

Case 15
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("SELECT DISTINCT b.mid_codigo FROM b_minuta a with (nolock) " & _
            "inner join b_minutadet b with (nolock) on a.min_codigo = b.mid_codigo " & _
            "WHERE a.min_cencos = '" & Trim(LimpiaDato(fpText.text)) & "' " & _
            "AND   a.min_codreg = " & Val(fpLongInteger1(0).Value) & " " & _
            "AND   a.min_codser = " & Val(fpLongInteger1(1).Value) & " " & _
            "AND   a.min_fecmin = " & Format(fpDateTime1(0).text, "yyyymmdd") & " " & _
            "AND   b.mid_nummer > 0 AND b.mid_tipmin = '2'")
    
    If RS.EOF Then
       RS.Close
       Set RS = Nothing
       
       MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    RS.Close
    Set RS = Nothing
    
    I_MermaPreparacion Trim(LimpiaDato(fpText.text)), fpLongInteger1(0).Value & ",", fpLongInteger1(1).Value & ",", Val(Format(fpDateTime1(0).Value, "yyyymmdd")), Val(Format(fpDateTime1(0).Value, "yyyymmdd")), 1

Case 18
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
If Err = -2147467259 Or 2147217900 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

If Trim(fpayuda(0).Caption) = "" Then

   vaSpread1.MaxRows = 0
   MsgBox "Debe ingresar ceco... proceso cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Sub

End If

If Trim(fpayuda(1).Caption) = "" Then

   vaSpread1.MaxRows = 0
   MsgBox "Debe ingresar regimen... proceso cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Sub

End If

If Trim(fpayuda(2).Caption) = "" Then

   vaSpread1.MaxRows = 0
   MsgBox "Debe ingresar servicio... proceso cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Sub

End If

If IsDate(fpDateTime1(0).text) = False Or Trim(fpDateTime1(0).text) = "" Then

   vaSpread1.MaxRows = 0
   MsgBox "Debe ingresar fecha... proceso cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Sub

End If

Mover_Datos

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim codreceta As Long
Dim GramajexKilo As Double
Dim RacionesProgramada As Long

With vaSpread1
    
    If .MaxRows < 1 Or ChangeMade = False Then
    
       Exit Sub
       
    End If
    
    If modo = "" Then
    
       modo = "M"
       
    End If
    
    If ChangeMade = True And modo = "M" And Toolbar1.Buttons(12).Visible = False Then
    
       Gl_Ac_Botones Me, 1, 0, modo
       Toolbar1.Buttons(1).Visible = False
       Toolbar1.Buttons(2).Visible = False
    
       fpLongInteger1(0).Enabled = False
       fpayuda(1).Enabled = False
    
       fpLongInteger1(1).Enabled = False
       fpayuda(2).Enabled = False
    
       fpDateTime1(0).Enabled = False
       
       'Frame1.Enabled = False
    
    End If
    
    Dim vCosRec As Double, vNumMer As Double, vNumPro As Long
    
    .Row = Row
    .Col = 3
    vNumPro = .Value
    
    .Col = 4
    vCosRec = .Value
    
    Select Case Col
        
        Case 6
        
            .Row = Row
            .Col = Col
            .TypeNumberDecimal = vg_CDec
            .TypeNumberDecPlaces = 0
            
            vNumMer = .Value

            
            If vCosRec = 0 Or vNumMer = 0 Then
               
               .Row = Row
               
               .Col = Col
               .text = ""
               
               .Col = 7
               .text = ""
               
               .Col = 8
               .text = ""
            
            Else
                          
                .Col = 3
                RacionesProgramada = .text
                
                If vNumMer > RacionesProgramada Then
                
                   MsgBox "Mermas x Raciones digitada es mayor que la programada", vbCritical + vbOKOnly, MsgTitulo
                   
                   .Row = Row
               
                   .Col = Col
                   .text = ""
                      
                   .Col = 7
                   .text = ""
                      
                   .Col = 8
                   .text = ""
                      
                   .Col = Col
                
                    Exit Sub
                
                End If
                
                .Col = 1
                codreceta = vaSpread1.text
 
               .Col = 8
               .TypeHAlign = TypeHAlignRight
               .text = Format(vCosRec * vNumMer, fg_Pict(6, 4))
                
                If RS.State = 1 Then RS.Close
                RS.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient

                Set RS = vg_db.Execute("sgp_Sel_MermaPorPreparacionReceta '" & LimpiaDato(fpText.text) & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Format(fpDateTime1(0).text, "yyyymmdd") & ", " & codreceta & "")

                If Not RS.EOF Then
                
                   .Col = 7
                   .CellType = CellTypeNumber
                   .TypeNumberSeparator = vg_CSep
                   .TypeNumberDecimal = vg_CDec
                   .TypeNumberDecPlaces = 4 'vg_DCa
                   
                   .text = Format(((RS!CantServida * vNumMer) / RS!Gramajeracionesnovendidas), fg_Pict(6, 4)) '3))
                   
                   .Col = 10
                   .CellType = CellTypeNumber
                   .TypeNumberSeparator = vg_CSep
                   .TypeNumberDecimal = vg_CDec
                   .TypeNumberDecPlaces = 4 'vg_DCa
                   
                   .text = Format(((RS!CantBruta * vNumMer) / RS!Gramajeracionesnovendidas), fg_Pict(6, 4)) '3))
                
                End If
                RS.Close
                Set RS = Nothing
            
            End If
    
        Case 7
        
            .Col = Col
                       
            GramajexKilo = 0
            
            GramajexKilo = IIf(Trim(.text) = "", 0, .text)
            If vCosRec = 0 Or GramajexKilo = 0 Then
               
               .Col = Col
               .text = ""
               
               .Col = 6
               .text = ""
               
               .Col = 8
               .text = ""
            
            Else
               
                .Row = Row
                
                .Col = 1
                codreceta = .text
                
                .Col = 3
                RacionesProgramada = .text
                
                If RS.State = 1 Then RS.Close
                RS.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient

                Set RS = vg_db.Execute("sgp_Sel_MermaPorPreparacionReceta '" & LimpiaDato(fpText.text) & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Format(fpDateTime1(0).text, "yyyymmdd") & ", " & codreceta & "")

                If Not RS.EOF Then
                
                   If Format(((GramajexKilo) / RS!CantServida) * RS!Gramajeracionesnovendidas, fg_Pict(6, 0)) > RacionesProgramada Then
                      
                      RS.Close
                      Set RS = Nothing
                   
                      MsgBox "Mermas x Kilos digitada es mayor que la programada", vbCritical + vbOKOnly, MsgTitulo
                   
                      .Row = Row
                
                      .Col = Col
                      .text = ""
                      
                      .Col = 6
                      .text = ""
                      
                      .Col = 8
                      .text = ""
                      
                      .Col = 10
                      .text = ""
                      
                      .Col = Col
                      
                      Exit Sub
                      
                   End If
                   
                    If Format(((GramajexKilo) / RS!CantServida) * RS!Gramajeracionesnovendidas, fg_Pict(6, 2)) <= 0 Then
                    
                      .Row = Row
                      
                      .Col = Col
                      .text = ""
                      
                      .Col = 6
                      .text = ""
                      
                      .Col = 8
                      .text = ""
                      
                      .Col = 10
                      .text = ""
                      
                      .Col = Col
                    
                    Else
                    
                       .Col = 6
                       .CellType = CellTypeNumber
                       .TypeNumberSeparator = vg_CSep
                       .TypeNumberDecimal = vg_CDec
                       .TypeNumberDecPlaces = vg_DCa
         
                       .text = Format(((GramajexKilo) / RS!CantServida) * RS!Gramajeracionesnovendidas, fg_Pict(6, 2))
                   
                       vNumMer = .Value
                   
                       .Col = 8
                       .TypeHAlign = TypeHAlignRight
                       .text = Format(vCosRec * vNumMer, fg_Pict(6, 0))
                       
                       .Col = 10
                       .CellType = CellTypeNumber
                       .TypeNumberSeparator = vg_CSep
                       .TypeNumberDecimal = vg_CDec
                       .TypeNumberDecPlaces = 4 'vg_DCa
                   
                       .text = Format(((RS!CantBruta * vNumMer) / RS!Gramajeracionesnovendidas), fg_Pict(6, 4)) '3))
                
                  End If
                  
                End If
                RS.Close
                Set RS = Nothing

            End If

    End Select

End With
calmertot


Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo Man_Error

With vaSpread1
    
    If .MaxRows < 1 Then Exit Sub
    
    .Row = NewRow
    .Col = 3
    vNumPro = .Value
    
'    .Col = 6
'    .TypeNumberMax = IIf(IsNull(vNumPro) Or Trim(vNumPro) = "", 999999, vNumPro)
    
    Select Case NewCol
    
        Case 6
        
'           .TypeNumberDecimal = vg_CDec
'           .TypeNumberDecPlaces = 0
            
    
    End Select

End With

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Sub calmertot()

On Error GoTo Man_Error

Dim vTotCos As Double

vTotCos = 0
With vaSpread1
    
    For i = 1 To .MaxRows
        
        .Row = i
        .Col = 8
        
        If Trim(.text) <> "" Then
        
           vTotCos = Round(vTotCos + .Value, 4) '0)
        
        End If
        
    Next i

End With
Label1(3).Caption = Format(vTotCos, fg_Pict(6, 4)) '0))


Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub
