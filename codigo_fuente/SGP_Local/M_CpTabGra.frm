VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_CpTabGra 
   Caption         =   "Copiar Tabla de Gramaje Destino"
   ClientHeight    =   3945
   ClientLeft      =   2355
   ClientTop       =   3030
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   11010
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
      Height          =   1845
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   10290
      Begin VB.Frame Frame2 
         Caption         =   "Zonas Destino con Tabla Gramaje"
         Height          =   1650
         Left            =   7200
         TabIndex        =   24
         Top             =   120
         Width           =   3015
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   1335
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   2355
            _Version        =   393217
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
            Enabled         =   0   'False
         End
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   5
         Left            =   1440
         TabIndex        =   9
         Top             =   520
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
         Index           =   4
         Left            =   1440
         TabIndex        =   7
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
         Index           =   6
         Left            =   1440
         TabIndex        =   11
         Top             =   860
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
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   2040
         TabIndex        =   31
         Top             =   1440
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   6
         Left            =   2385
         Picture         =   "M_CpTabGra.frx":0000
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Zona"
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
         Left            =   210
         TabIndex        =   27
         Top             =   920
         Width           =   900
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   2880
         TabIndex        =   12
         Top             =   860
         Width           =   4110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   2880
         TabIndex        =   10
         Top             =   520
         Width           =   4110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   2880
         TabIndex        =   8
         Top             =   180
         Width           =   4110
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
         Left            =   210
         TabIndex        =   19
         Top             =   575
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Subsegmento"
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
         Left            =   210
         TabIndex        =   18
         Top             =   220
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   2385
         Picture         =   "M_CpTabGra.frx":030A
         Top             =   80
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   2385
         Picture         =   "M_CpTabGra.frx":0614
         Top             =   425
         Width           =   480
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   2925
         TabIndex        =   21
         Top             =   570
         Width           =   4110
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2925
         TabIndex        =   20
         Top             =   225
         Width           =   4110
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   2925
         TabIndex        =   28
         Top             =   900
         Width           =   4110
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
      Height          =   1845
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10290
      Begin VB.Frame Frame3 
         Caption         =   "Zonas Origen con Tabla Gramaje"
         Height          =   1650
         Left            =   7200
         TabIndex        =   22
         Top             =   120
         Width           =   3015
         Begin MSComctlLib.TreeView TvwZon 
            Height          =   1335
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   2355
            _Version        =   393217
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
            Enabled         =   0   'False
         End
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   2
         Left            =   1440
         TabIndex        =   3
         Top             =   520
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
         TabIndex        =   1
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
         TabIndex        =   5
         Top             =   860
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2415
         Picture         =   "M_CpTabGra.frx":091E
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Zona"
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
         Left            =   240
         TabIndex        =   29
         Top             =   920
         Width           =   900
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   2880
         TabIndex        =   6
         Top             =   860
         Width           =   4110
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2385
         Picture         =   "M_CpTabGra.frx":0C28
         Top             =   425
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2385
         Picture         =   "M_CpTabGra.frx":0F32
         Top             =   80
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Subsegmento"
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
         Top             =   220
         Width           =   1155
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
         TabIndex        =   13
         Top             =   575
         Width           =   750
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2880
         TabIndex        =   2
         Top             =   180
         Width           =   4110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2880
         TabIndex        =   4
         Top             =   520
         Width           =   4110
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2925
         TabIndex        =   15
         Top             =   220
         Width           =   4110
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2925
         TabIndex        =   16
         Top             =   570
         Width           =   4110
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   2925
         TabIndex        =   30
         Top             =   900
         Width           =   4110
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   3945
      Left            =   10470
      TabIndex        =   26
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   6959
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_CpTabGra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim i, j As Integer
Dim Msgtitulo As String

Private Sub Form_Load()
Msgtitulo = "Copiar Tabla de Gramaje"
fg_centra Me

Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

'Llena Nodo con Información solamente
TvwZon.Nodes.Clear
Set RS = vg_db.Execute("sgpadm_s_zona 6, 0,''")
Do While Not RS.EOF
   Set rootNode = TvwZon.Nodes.Add(, , "H" & RS!zon_codigo, Trim(RS!Zon_nombre))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing

TreeView1.Nodes.Clear
Set RS = vg_db.Execute("sgpadm_s_zona 6, 0,''")
Do While Not RS.EOF
   Set rootNode = TreeView1.Nodes.Add(, , "H" & RS!zon_codigo, Trim(RS!Zon_nombre))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing

Dim CodzonOri As Variant
ReDim CodzonOri(M_TabGra.TvwZon.Nodes.count)
 For j = 1 To M_TabGra.TvwZon.Nodes.count
       If M_TabGra.TvwZon.Nodes.Item(j).Checked = True Then
       CodzonOri(j) = M_TabGra.TvwZon.Nodes.Item(j).text
       End If
 Next j
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Select Case Index
Case 1
    RS.Open "SELECT * FROM a_subsegmento WHERE sub_codigo=" & Val(fpLongInteger1(Index).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(Index).Caption = "": Exit Sub ':  fpayuda(Index).Caption = "": fpLongInteger1(Index).Value = "": fpayuda(Index).Caption = "" ': Exit Sub
    fpayuda(Index).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing
    CargaCodzon 1
Case 2
    If Val(fpLongInteger1(Index).Value) < 1 Then fpayuda(Index).Caption = "": Exit Sub
    RS.Open "SELECT * FROM a_regimen WHERE reg_codigo=" & Val(fpLongInteger1(Index).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(Index).Caption = "": Exit Sub
    fpayuda(Index).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    CargaCodzon 1
Case 3
    Set RS = vg_db.Execute("sgpadm_s_zona 1, " & Val(fpLongInteger1(Index).Value) & ", ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(Index).Caption = "": Exit Sub
    fpayuda(Index).Caption = Trim(RS!Zon_nombre)
    RS.Close: Set RS = Nothing
Case 4
    RS.Open "SELECT * FROM a_subsegmento WHERE sub_codigo=" & Val(fpLongInteger1(Index).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(Index).Caption = "": Exit Sub 'fpLongInteger1(Index).Value = "": fpayuda(Index).Caption = "": fpLongInteger1(Index).Value = "": fpayuda(Index).Caption = "": Exit Sub
    fpayuda(Index).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing
    CargaCodzon 2
Case 5
    If Val(fpLongInteger1(Index).Value) < 1 Then fpayuda(Index).Caption = "": Exit Sub
    RS.Open "SELECT * FROM a_regimen WHERE reg_codigo=" & Val(fpLongInteger1(Index).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(Index).Caption = "": Exit Sub
    fpayuda(Index).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    CargaCodzon 2
Case 6
    Set RS = vg_db.Execute("sgpadm_s_zona 1, " & Val(fpLongInteger1(Index).Value) & ", ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(Index).Caption = "": Exit Sub
    fpayuda(Index).Caption = Trim(RS!Zon_nombre)
    RS.Close: Set RS = Nothing
    
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 1
    vg_left = fpayuda(Index).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_subsegmento", "sub_", "Subsegmento", "Sub"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(Index).Value = Val(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    fpLongInteger1(Index).SetFocus
Case 2
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(Index).Value = Val(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    fpLongInteger1(Index).SetFocus
Case 3
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_zona", "zon_", "Zona", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(Index).Value = Val(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
Case 4
    vg_left = fpayuda(Index).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_subsegmento", "sub_", "Subsegmento", "Sub"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(Index).Value = Val(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    fpLongInteger1(Index).SetFocus
Case 5
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(Index).Value = Val(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    fpLongInteger1(Index).SetFocus
    
Case 6
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_zona", "zon_", "Zona", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(Index).Value = Val(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
End Select
End Sub

Sub LimpiaZonas(opcion As Integer)
If opcion = 1 Then
  For j = 1 To TvwZon.Nodes.count
        TvwZon.Nodes.Item(j).Checked = False
  Next j
Else
  For j = 1 To TreeView1.Nodes.count
        TreeView1.Nodes.Item(j).Checked = False
  Next j
End If
End Sub
Sub CargaCodzon(opcion As Integer)
Dim codzon As Integer: codzon = 0
Dim subseg As Integer, codreg As Integer

Dim Zon_nombre As String
  If opcion = 1 Then
    If fpLongInteger1(1).Value = "" Then Exit Sub
    If fpLongInteger1(2).Value = "" Then Exit Sub
    LimpiaZonas opcion
    subseg = fpLongInteger1(1).Value: codreg = fpLongInteger1(2).Value
    
    RS.Open "select a.tgr_codzon,b.zon_nombre from b_tablagramaje a, a_zona b Where a.tgr_codzon = b.zon_codigo AND a.tgr_subseg=" & subseg & " and a.tgr_codreg=" & codreg & "   group by a.tgr_codzon,b.zon_nombre  order by a.tgr_codzon", vg_db, adOpenStatic
    If Not RS.EOF Then
      While Not RS.EOF
         codzon = RS!tgr_codzon
         Zon_nombre = RS!Zon_nombre
         For j = 1 To TvwZon.Nodes.count
           If TvwZon.Nodes.Item(j).text = Zon_nombre Then
           TvwZon.Nodes.Item(j).Checked = True
           End If
         Next j
      RS.MoveNext
      Wend
      RS.Close: Set RS = Nothing
    Else
      RS.Close: Set RS = Nothing
      For j = 1 To TreeView1.Nodes.count
        TvwZon.Nodes.Item(j).Checked = False
      Next j
    End If
  End If
  If opcion = 2 Then
    If fpLongInteger1(4).Value = "" Then Exit Sub
    If fpLongInteger1(5).Value = "" Then Exit Sub
    LimpiaZonas opcion
    subseg = fpLongInteger1(4).Value: codreg = fpLongInteger1(5).Value
    RS.Open "select a.tgr_codzon,b.zon_nombre from b_tablagramaje a, a_zona b Where a.tgr_codzon = b.zon_codigo AND a.tgr_subseg=" & subseg & " and a.tgr_codreg=" & codreg & "   group by a.tgr_codzon,b.zon_nombre  order by a.tgr_codzon", vg_db, adOpenStatic
    If Not RS.EOF Then
      While Not RS.EOF
         codzon = RS!tgr_codzon
         Zon_nombre = RS!Zon_nombre
         For j = 1 To TreeView1.Nodes.count
           If TreeView1.Nodes.Item(j).text = Zon_nombre Then
           TreeView1.Nodes.Item(j).Checked = True
           End If
         Next j
      RS.MoveNext
      Wend
      RS.Close: Set RS = Nothing
    Else
      RS.Close: Set RS = Nothing
      For j = 1 To TreeView1.Nodes.count
        TreeView1.Nodes.Item(j).Checked = False
      Next j
    End If
  End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim SubsegO, SubsegD, CodregO, CodregD, CodzonO, CodzonD As Integer
Dim Subs, reg, rec, ing, zon, ins, gr As Integer
Dim CantReg As Integer
Select Case Button.Index
Case 1
If Trim(fpayuda(1).Caption) = "" Or Trim(fpayuda(2).Caption) = "" Or Trim(fpayuda(3).Caption) = "" _
    Or Trim(fpayuda(4).Caption) = "" Or Trim(fpayuda(5).Caption) = "" Or Trim(fpayuda(6).Caption) = "" Then MsgBox "Faltan Datos.", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
SubsegO = fpLongInteger1(1).Value: CodregO = fpLongInteger1(2).Value: CodzonO = fpLongInteger1(3).Value:
SubsegD = fpLongInteger1(4).Value: CodregD = fpLongInteger1(5).Value: CodzonD = fpLongInteger1(6).Value:

'Validar Existencia Datos Origen
RS.Open "select * from b_tablagramaje where tgr_subseg=" & SubsegO & " and tgr_codreg= " & CodregO & " AND tgr_codzon=" & CodzonO & " order by tgr_codzon", vg_db, adOpenStatic
If RS.EOF Then
  MsgBox "No Existe Datos de Origen", vbExclamation + vbOKOnly, Msgtitulo: RS.Close: Set RS = Nothing: Exit Sub
End If

RS.Close: Set RS = Nothing

RS1.Open "select * from b_tablagramaje where tgr_subseg=" & SubsegD & " and tgr_codreg= " & CodregD & " AND tgr_codzon=" & CodzonD & " order by tgr_codzon", vg_db, adOpenStatic
'Validar Existencia de Datos de Destino
If Not RS1.EOF Then If MsgBox("Existe información en Tabla Gramaje destino. se borrara la información existente ...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then RS1.Close: Set RS1 = Nothing: Exit Sub

'Realizaremos la eliminación de tabla gramaje existente
If Not RS1.EOF Then
  vg_db.Execute "DELETE b_tablagramaje WHERE tgr_subseg=" & SubsegD & " and tgr_codreg= " & CodregD & " AND tgr_codzon=" & CodzonD & ""
End If
Set RS1 = Nothing
' Ahora recorremos datos de Origen y insertamos en Destino
  vg_db.Execute "Insert into b_tablagramaje (tgr_subseg, tgr_codreg, tgr_codrec, tgr_coding, tgr_codzon, tgr_codins, tgr_cantgr)" & _
                " Select " & SubsegD & " , " & CodregD & ", tgr_codrec, tgr_coding, " & CodzonD & ", tgr_codins, tgr_cantgr from b_tablagramaje where " & _
                " tgr_subseg= " & SubsegO & " and tgr_codreg= " & CodregO & " AND tgr_codzon=" & CodzonO & " "
MsgBox "Proceso terminado exitosamente."
CargaCodzon 1
CargaCodzon 2
Case 3
Me.Hide: Unload Me: M_TabGra.WindowState = 0
End Select
End Sub
