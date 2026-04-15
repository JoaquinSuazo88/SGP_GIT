VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form I_MenTeo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Planificación Teórica"
   ClientHeight    =   4905
   ClientLeft      =   2190
   ClientTop       =   1770
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4545
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   7395
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   255
         Left            =   4440
         TabIndex        =   30
         Top             =   4080
         Visible         =   0   'False
         Width           =   615
         _Version        =   393216
         _ExtentX        =   1085
         _ExtentY        =   450
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
         MaxCols         =   3
         MaxRows         =   0
         SpreadDesigner  =   "I_MenTeo.frx":0000
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Aportes"
         Enabled         =   0   'False
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
         Height          =   735
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   3000
         Width           =   6975
         Begin VB.OptionButton Option1 
            Caption         =   "Peso Neto"
            Enabled         =   0   'False
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
            Index           =   6
            Left            =   3840
            TabIndex        =   29
            Top             =   360
            Width           =   1260
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Peso Bruto"
            Enabled         =   0   'False
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
            Index           =   4
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Peso Servido"
            Enabled         =   0   'False
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
            Index           =   5
            Left            =   1800
            TabIndex        =   27
            Top             =   360
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Ambos"
            Enabled         =   0   'False
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
            Index           =   7
            Left            =   5520
            TabIndex        =   26
            Top             =   360
            Width           =   1260
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Nutrientes"
         Enabled         =   0   'False
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
         Height          =   615
         Index           =   1
         Left            =   4200
         TabIndex        =   22
         Top             =   2160
         Width           =   2895
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
            Enabled         =   0   'False
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
            Index           =   3
            Left            =   1560
            TabIndex        =   24
            Top             =   300
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   3
            Left            =   2280
            Picture         =   "I_MenTeo.frx":020C
            Top             =   160
            Width           =   480
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6975
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "I_MenTeo.frx":0516
            Left            =   1680
            List            =   "I_MenTeo.frx":0518
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Informes"
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
            Left            =   600
            TabIndex        =   13
            Top             =   300
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   2895
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
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
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
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
            Left            =   1560
            TabIndex        =   5
            Top             =   300
            Width           =   735
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   2
            Left            =   2280
            Picture         =   "I_MenTeo.frx":051A
            Top             =   160
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Recetas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   3885
         Begin VB.OptionButton Option1 
            Caption         =   "Nombre Fantasia"
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
            Index           =   8
            Left            =   120
            TabIndex        =   6
            Top             =   300
            Value           =   -1  'True
            Width           =   1785
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Nombre Receta"
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
            Index           =   9
            Left            =   2070
            TabIndex        =   7
            Top             =   300
            Width           =   1665
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   135
         Left            =   5880
         TabIndex        =   14
         Top             =   4080
         Visible         =   0   'False
         Width           =   615
         _Version        =   393216
         _ExtentX        =   1085
         _ExtentY        =   238
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
         MaxCols         =   4
         MaxRows         =   100
         SpreadDesigner  =   "I_MenTeo.frx":0824
      End
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   1
         Left            =   3135
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1410
         Width           =   3765
         _Version        =   196608
         _ExtentX        =   6641
         _ExtentY        =   556
         Enabled         =   0   'False
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
         BackColor       =   -2147483638
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
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
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
         ControlType     =   3
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
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   0
         Left            =   3135
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3765
         _Version        =   196608
         _ExtentX        =   6641
         _ExtentY        =   556
         Enabled         =   0   'False
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
         BackColor       =   -2147483638
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
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
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
         ControlType     =   3
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
         Left            =   1455
         TabIndex        =   1
         Top             =   1410
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
         Left            =   1440
         TabIndex        =   0
         Top             =   1080
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
         Index           =   0
         Left            =   1455
         TabIndex        =   2
         Top             =   1740
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   4350
         TabIndex        =   3
         Top             =   1740
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
      Begin VB.Label Label2 
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
         Index           =   2
         Left            =   3225
         TabIndex        =   21
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label2 
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
         Index           =   1
         Left            =   285
         TabIndex        =   20
         Top             =   1800
         Width           =   1110
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2685
         Picture         =   "I_MenTeo.frx":0ED8
         Top             =   1290
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2685
         Picture         =   "I_MenTeo.frx":11E2
         Top             =   990
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Casino"
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
         Left            =   285
         TabIndex        =   18
         Top             =   1155
         Width           =   585
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
         Index           =   6
         Left            =   285
         TabIndex        =   17
         Top             =   1470
         Width           =   750
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_MenTeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim i As Integer, isel As Integer
Dim accion As Boolean
Dim Msgtitulo As String, tipmin As String

Private Sub Combo1_Click(Index As Integer)
Select Case Val(fg_codigocbo(Combo1, 0, 1, ""))
Case 0, 1, 4, 5
    Frame3(1).Enabled = False: Frame3(4).Enabled = False
    Option1(2).Enabled = False: Option1(3).Enabled = False
    Option1(4).Enabled = False: Option1(5).Enabled = False
    Option1(6).Enabled = False: Option1(7).Enabled = False
Case 2
    Frame3(1).Enabled = True: Frame3(4).Enabled = True
    Option1(2).Enabled = True: Option1(3).Enabled = True
    Option1(4).Enabled = True: Option1(5).Enabled = True
    Option1(6).Enabled = True: Option1(7).Enabled = True
Case 3
    Frame3(1).Enabled = True: Frame3(4).Enabled = False
    Option1(2).Enabled = True: Option1(3).Enabled = True
    Option1(4).Enabled = False: Option1(5).Enabled = False
    Option1(6).Enabled = False: Option1(7).Enabled = False
End Select
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 5385
Me.Width = 7485
fg_centra Me
'Dim btnX As Button
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): btnX.Visible = True: btnX.ToolTipText = "Vista Previa": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): btnX.Visible = True: btnX.ToolTipText = "Historico Planificacón Teórica"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
fpDateTime1(0).Text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).Text = Format(Date, "dd/mm/yyyy")
accion = True
Combo1(0).Clear
Combo1(0).AddItem "Menú Mecano" & Space(150) & "(0)"
Combo1(0).AddItem "Menú Mensual" & Space(150) & "(1)"
Combo1(0).AddItem "Aporte Nutricionales Detallado" & Space(150) & "(2)"
Combo1(0).AddItem "Aporte Nutricionales Resumido" & Space(150) & "(3)"
Combo1(0).AddItem "Costo Detallado" & Space(150) & "(4)"
Combo1(0).AddItem "Costo Resumido" & Space(150) & "(5)"
Combo1(0).AddItem "Ingredientes Valor Cero en Planificación" & Space(150) & "(6)"
Combo1(0).ListIndex = -1

' *** Llenar Tabla Nutrienetes *** '
RS.Open "select nut_codigo, nut_nombre, nut_indpri, nut_secnro from a_nutriente order by nut_secnro", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No existe maestro nutrientes", vbExclamation + vbOKOnly, Msgtitulo: Me.Hide: Unload Me
vaSpread2.MaxRows = 0
Do While Not RS.EOF
   vaSpread2.MaxRows = vaSpread2.MaxRows + 1
   vaSpread2.Row = vaSpread2.MaxRows
   vaSpread2.Col = 2: vaSpread2.Text = RS!nut_codigo
   vaSpread2.Col = 3: vaSpread2.Text = Trim(RS!nut_nombre)
   If RS!nut_indpri = 1 Then
      vaSpread2.Col = 1
      vaSpread2.CellType = 10
      vaSpread2.TypeCheckText = ""
      vaSpread2.TypeCheckCenter = True
      vaSpread2.Text = "1" ' checked
   Else
      vaSpread2.Col = 1
      vaSpread2.CellType = 10
      vaSpread2.TypeCheckText = ""
      vaSpread2.TypeCheckCenter = True
      vaSpread2.Text = " " ' checked
   End If
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.Text = MuestraCasino(1)
fpayuda(0).Text = MuestraCasino(2)
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpDateTime1_LostFocus(Index As Integer)
If fpDateTime1(Index).Text = "" Then Exit Sub
MoverDatoGrilla
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    If Index = 1 Then Image1_Click 1
End Select
End Sub

Private Sub fpLongInteger1_LostFocus(Index As Integer)
Select Case Index
Case 0
    If Val(fpLongInteger1(0).Value) < 1 Then fpayuda(1).Text = "": Exit Sub
    RS.Open "select * from a_regimen where reg_codigo=" & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Text = "": Exit Sub
    fpayuda(1).Text = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    MoverDatoGrilla
End Select
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

Private Sub fpText_LostFocus()
If fpText.Text = "" Then fpayuda(0).Text = "": Exit Sub
RS.Open "select * from b_clientes where cli_codigo='" & fpText.Text & "' and cli_tipo=0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Text = "": Exit Sub
fpayuda(0).Text = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
MoverDatoGrilla
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casinos", "Casino"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.Text = vg_codigo
    fpayuda(0).Text = vg_nombre
    MoverDatoGrilla
    fpLongInteger1(0).SetFocus
Case 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(1).Text = vg_nombre
    fpDateTime1(0).SetFocus
Case 2
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpText.Text = "" Or fpLongInteger1(0).Value = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Servicio", I_MenTeo.vaSpread1, fpText.Text, fpLongInteger1(0).Value, Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "1"
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
Case 3
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpText.Text = "" Or fpLongInteger1(0).Value = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Nutrientes", I_MenTeo.vaSpread2, fpText.Text, fpLongInteger1(0).Value, Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "2"
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    Image1(2).Enabled = False
Case 1
    Image1(2).Enabled = True
Case 2
    Image1(3).Enabled = False
Case 3
    Image1(3).Enabled = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim opnomrec As Boolean
Select Case Button.Index
Case 1
    If vaSpread1.MaxRows < 1 Then MsgBox "No existe Información", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Open "select cli_nombre from b_clientes where cli_codigo='" & LimpiaDato(Trim(fpText.Text)) & "' and cli_tipo=0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.Text = "": fpayuda(0).Text = "": MsgBox "No existe casino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    RS.Open "select * from a_regimen where reg_codigo=" & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Text = "": MsgBox "No existe regimen", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    If fpDateTime1(0).Value > fpDateTime1(1).Value Then MsgBox "Fecha origen Mayor destino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Mid(fpDateTime1(0).Text, 4, 2) <> Mid(fpDateTime1(1).Text, 4, 2) Then MsgBox "Mes origen mayor destino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Mid(fpDateTime1(0).Text, 7, 4) <> Mid(fpDateTime1(1).Text, 7, 4) Then MsgBox "Año origen mayor destino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Combo1(0).ListIndex = -1 Or Combo1(0).Text = "" Then Exit Sub
    If Option1(0).Value = True Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = 1
           vaSpread1.CellType = 10
           vaSpread1.TypeCheckText = ""
           vaSpread1.TypeCheckCenter = True
           vaSpread1.Text = "1" ' checked
       Next i
    End If
    isel = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.Text = "1" Then isel = 1: Exit For
    Next i
    If isel = 0 Then fg_descarga: MsgBox "Servicio debe ser informado", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    opnomrec = True
    If Option1(9).Value = True Then opnomrec = False
    Select Case Val(fg_codigocbo(Combo1, 0, 1, ""))
    Case 0
       If tipmin = "PLATEO" Then I_MenuPlanMecano fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "1", opnomrec
       If tipmin = "PLAREA" Then I_MenuPlanMecano fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "2", opnomrec
    Case 1
       If tipmin = "PLATEO" Then I_MenuPlanMensual fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "1", opnomrec
       If tipmin = "PLAREA" Then I_MenuPlanMensual fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "2", opnomrec
    Case 2
       If tipmin = "PLATEO" Then I_AportePlanDetallado fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "1", opnomrec
       If tipmin = "PLAREA" Then I_AportePlanDetallado fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "2", opnomrec
    Case 3
       If tipmin = "PLATEO" Then I_AportePlanResumido fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "1", opnomrec
       If tipmin = "PLAREA" Then I_AportePlanResumido fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "2", opnomrec
    Case 4
       If tipmin = "PLATEO" Then I_CostoPlanDetallado fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "1", opnomrec
       If tipmin = "PLAREA" Then I_CostoPlanDetallado fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "2", opnomrec
    Case 5
       If tipmin = "PLATEO" Then I_CostoPlanResumido fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "1", opnomrec
       If tipmin = "PLAREA" Then I_CostoPlanResumido fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "2", opnomrec
    Case 6
       If tipmin = "PLATEO" Then I_IngValCeroPlan fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "1", opnomrec
       If tipmin = "PLAREA" Then I_IngValCeroPlan fpText.Text, Val(fpLongInteger1(0).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "2", opnomrec
    End Select
Case 3
    ' *** Historico Planificación *** '
    RS.Open "select cli_nombre from b_clientes where cli_codigo='" & LimpiaDato(Trim(fpText.Text)) & "' and cli_tipo=0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.Text = "": fpayuda(0).Text = "": MsgBox "No existe casino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    vg_codigo = ""
    If tipmin = "PLATEO" Then B_HistPm.LlenarHistPlan "Histórico Planificación Teórica", fpText.Text, 1, 2
    If tipmin = "PLAREA" Then B_HistPm.LlenarHistPlan "Histórico Planificación Real", fpText.Text, 2, 2
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    RS.Open "select reg_nombre from a_regimen where reg_codigo=" & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Text = "": Exit Sub
    fpayuda(1).Text = Trim(RS!reg_nombre): RS.Close: Set RS = Nothing
    fpDateTime1(0).Text = "01/" & vg_auxfecha: fpDateTime1(1).Text = dEoM("01/" & vg_auxfecha)
    MoverDatoGrilla
    Option1(0).SetFocus
    Me.Refresh
Case 5
    Me.Hide
    Unload Me
End Select
End Sub

Sub MoverDatoGrilla()
If LimpiaDato(Trim(fpText.Text)) = "" And Val(fpLongInteger1(0).Value) = 0 And fpDateTime1(0).Text = "" And fpDateTime1(1).Text = "" Then Exit Sub
fg_carga ""
vaSpread1.MaxRows = 0
RS.Open "select distinct b_minuta.min_codser, a_servicio.ser_nombre, a_servicio.ser_orden " & _
        "from  a_servicio, b_minuta, b_minutadet " & _
        "where b_minuta.min_codigo=b_minutadet.mid_codigo " & _
        "and   b_minuta.min_codser=a_servicio.ser_codigo " & _
        "and   b_minuta.min_cencos='" & LimpiaDato(Trim(fpText.Text)) & "' " & _
        "and   b_minuta.min_codreg=" & Val(fpLongInteger1(0).Value) & " " & _
        "and   b_minuta.min_fecmin>=" & Val(Format(fpDateTime1(0).Text, "yyyymmdd")) & " " & _
        "and   b_minuta.min_fecmin<=" & Val(Format(fpDateTime1(1).Text, "yyyymmdd")) & " " & _
        "and   b_minutadet.mid_tipmin='1' " & _
        "order by a_servicio.ser_orden, a_servicio.ser_nombre", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: Exit Sub
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 2: vaSpread1.Text = RS!min_codser
   vaSpread1.Col = 3: vaSpread1.Text = Trim(RS!ser_nombre)
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing: fg_descarga
End Sub

Sub Inicio(tfor As String, TipM As String)
Me.Caption = tfor
Msgtitulo = tfor
tipmin = TipM
End Sub

