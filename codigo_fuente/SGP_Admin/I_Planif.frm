VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_Planif 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Planificación"
   ClientHeight    =   7230
   ClientLeft      =   4335
   ClientTop       =   2520
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6825
      Index           =   0
      Left            =   60
      TabIndex        =   17
      Top             =   360
      Width           =   7875
      Begin FPSpread.vaSpread vaSpread3 
         Height          =   615
         Left            =   4800
         TabIndex        =   52
         Top             =   5640
         Visible         =   0   'False
         Width           =   1335
         _Version        =   393216
         _ExtentX        =   2355
         _ExtentY        =   1085
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
         SpreadDesigner  =   "I_Planif.frx":0000
      End
      Begin VB.Frame Frame7 
         Caption         =   "Costo x Estructura"
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
         Height          =   615
         Left            =   120
         TabIndex        =   47
         Top             =   5880
         Visible         =   0   'False
         Width           =   3855
         Begin VB.OptionButton Option2 
            Caption         =   "Sumatoria"
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
            Left            =   1920
            TabIndex        =   49
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Ponderado"
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
            TabIndex        =   48
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Estructuras"
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
         TabIndex        =   44
         Top             =   5160
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
            Index           =   11
            Left            =   120
            TabIndex        =   46
            Top             =   300
            Value           =   -1  'True
            Width           =   1785
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Nombre Estructura"
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
            Index           =   10
            Left            =   1920
            TabIndex        =   45
            Top             =   300
            Width           =   1905
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Semana Cerrada"
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
         Left            =   4680
         TabIndex        =   43
         Top             =   4920
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         ItemData        =   "I_Planif.frx":01FE
         Left            =   1575
         List            =   "I_Planif.frx":0200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1840
         Width           =   1500
      End
      Begin VB.Frame Frame5 
         Caption         =   "Planificación"
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
         Left            =   4680
         TabIndex        =   38
         Top             =   3960
         Width           =   3015
         Begin VB.CheckBox Check1 
            Caption         =   "Costo"
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
            Left            =   1920
            TabIndex        =   40
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Raciones"
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
            TabIndex        =   39
            Top             =   240
            Width           =   1215
         End
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   2
         Left            =   1575
         TabIndex        =   3
         Top             =   1520
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
         Height          =   1095
         Left            =   120
         TabIndex        =   24
         Top             =   3960
         Width           =   3885
         Begin VB.CheckBox Check3 
            Caption         =   "No Incluye Parentesis"
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
            Left            =   960
            TabIndex        =   50
            Top             =   600
            Width           =   2175
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
            TabIndex        =   16
            Top             =   300
            Width           =   1665
         End
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
            TabIndex        =   15
            Top             =   300
            Value           =   -1  'True
            Width           =   1785
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
         TabIndex        =   23
         Top             =   2520
         Width           =   2895
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
            TabIndex        =   8
            Top             =   300
            Width           =   735
         End
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
            TabIndex        =   7
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   2
            Left            =   2280
            Picture         =   "I_Planif.frx":0202
            Top             =   160
            Width           =   480
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   7575
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "I_Planif.frx":050C
            Left            =   1320
            List            =   "I_Planif.frx":050E
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   5775
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
            Left            =   480
            TabIndex        =   22
            Top             =   300
            Width           =   735
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
         Left            =   4800
         TabIndex        =   20
         Top             =   2520
         Width           =   2895
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
            TabIndex        =   9
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
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
            TabIndex        =   10
            Top             =   300
            Width           =   735
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   3
            Left            =   2280
            Picture         =   "I_Planif.frx":0510
            Top             =   160
            Width           =   480
         End
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
         TabIndex        =   19
         Top             =   3120
         Width           =   7575
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
            Index           =   12
            Left            =   1560
            TabIndex        =   53
            Top             =   360
            Width           =   1260
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
            Index           =   7
            Left            =   6600
            TabIndex        =   14
            Top             =   360
            Width           =   900
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
            Left            =   3000
            TabIndex        =   12
            Top             =   360
            Width           =   1500
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
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Peso Neto Nut."
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
            Left            =   4680
            TabIndex        =   13
            Top             =   360
            Width           =   1740
         End
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1575
         TabIndex        =   2
         Top             =   1200
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
         Left            =   1575
         TabIndex        =   5
         Top             =   2200
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
         Left            =   6390
         TabIndex        =   6
         Top             =   2205
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1575
         TabIndex        =   1
         Top             =   900
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
         Left            =   120
         TabIndex        =   41
         Top             =   6600
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   495
         Left            =   4320
         TabIndex        =   18
         Top             =   4320
         Visible         =   0   'False
         Width           =   1815
         _Version        =   393216
         _ExtentX        =   3201
         _ExtentY        =   873
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
         MaxRows         =   0
         SpreadDesigner  =   "I_Planif.frx":081A
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2415
         Left            =   5280
         TabIndex        =   25
         Top             =   3960
         Visible         =   0   'False
         Width           =   2415
         _Version        =   393216
         _ExtentX        =   4260
         _ExtentY        =   4260
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
         SpreadDesigner  =   "I_Planif.frx":0A26
      End
      Begin MSComctlLib.TreeView TvwDir 
         Height          =   1725
         Left            =   4320
         TabIndex        =   51
         Top             =   4920
         Visible         =   0   'False
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   3043
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         Left            =   240
         TabIndex        =   42
         Top             =   1905
         Width           =   390
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   2475
         Picture         =   "I_Planif.frx":10DA
         Top             =   1430
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
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
         Left            =   255
         TabIndex        =   36
         Top             =   1550
         Width           =   450
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
         Index           =   2
         Left            =   2910
         TabIndex        =   35
         Top             =   1515
         Width           =   4800
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
         Left            =   2910
         TabIndex        =   34
         Top             =   1200
         Width           =   4800
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
         Left            =   2910
         TabIndex        =   32
         Top             =   900
         Width           =   4800
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
         Left            =   255
         TabIndex        =   29
         Top             =   1250
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sub-Segmento"
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
         Left            =   255
         TabIndex        =   28
         Top             =   935
         Width           =   1245
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2475
         Picture         =   "I_Planif.frx":13E4
         Top             =   770
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2475
         Picture         =   "I_Planif.frx":16EE
         Top             =   1080
         Width           =   480
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
         Left            =   240
         TabIndex        =   27
         Top             =   2280
         Width           =   1110
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
         Left            =   5265
         TabIndex        =   26
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2880
         TabIndex        =   33
         Top             =   915
         Width           =   4800
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2880
         TabIndex        =   31
         Top             =   1215
         Width           =   4800
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2880
         TabIndex        =   37
         Top             =   1545
         Width           =   4800
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_Planif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, isel As Integer
Dim accion As Boolean
Public lc_Aux As String
Dim MsgTitulo As String, TipMin As String
Dim rootNode As Node

Private Sub Check3_Click()

On Error GoTo Man_Error

If Check3.Value = 1 Then
   
   Check3.Caption = "Incluye Parentesis"

Else
   
   Check3.Caption = "No Incluye Parentesis"

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Combo1_Change(Index As Integer)

On Error GoTo Man_Error

Select Case Index
    
    Case 0
    
        MoverDatoGrilla
        
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Combo1_Click(Index As Integer)

On Error GoTo Man_Error

If Val(fg_codigocbo(Combo1, 0, 2, "")) <> 5 Then

   Check2.Visible = True

Else

   Check2.Visible = False

End If

Frame2.Enabled = True
Frame6.Enabled = True
Check2.Visible = True
Frame7.Visible = False
Frame7.Enabled = False
Frame7.Caption = "Costo x Estructura"
Option2(0).Caption = "Ponderado"
Option2(0).Value = True
Option2(1).Caption = "Ponderado Sumatoria"
Option2(1).Value = False

Select Case Val(fg_codigocbo(Combo1, 0, 2, ""))
    
    Case 0, 1, 5, 6
        
        Frame3(1).Enabled = False: Frame3(4).Enabled = False: Frame5.Enabled = IIf(Val(fg_codigocbo(Combo1, 0, 2, "")) = 1, True, False)
        Option1(2).Enabled = False: Option1(3).Enabled = False
        Option1(4).Enabled = False: Option1(5).Enabled = False
        Option1(6).Enabled = False: Option1(7).Enabled = False
        Option1(12).Enabled = False
        If Val(fg_codigocbo(Combo1, 0, 2, "")) <> 1 Then Frame7.Enabled = False: Check1(0).Value = 0: Check1(1).Value = 0 Else Frame7.Enabled = True
    
    Case 2, 7
        
        Frame5.Enabled = False
        Frame3(1).Enabled = True: Frame3(4).Enabled = True
        Option1(2).Enabled = True: Option1(3).Enabled = True
        Option1(4).Enabled = True: Option1(5).Enabled = True
        Option1(6).Enabled = True: Option1(7).Enabled = True
        Option1(12).Enabled = True
        
    Case 3, 4, 8, 9
        
        Frame5.Enabled = False
        Frame3(1).Enabled = True: Frame3(4).Enabled = False
        Option1(2).Enabled = True: Option1(3).Enabled = True
        Option1(4).Enabled = False: Option1(5).Enabled = False
        Option1(6).Enabled = False: Option1(7).Enabled = False
        Option1(12).Enabled = False
        If Val(fg_codigocbo(Combo1, 0, 2, "")) = 9 Or Val(fg_codigocbo(Combo1, 0, 2, "")) = 4 Then Frame2.Enabled = False
    
    Case 10, 11
        
        Frame2.Enabled = False
        Frame6.Enabled = False
        Frame7.Visible = True
        Frame7.Enabled = True
        Frame7.Caption = "Gramaje"
        Option2(0).Caption = "Bruto"
        Option2(0).Value = True
        Option2(1).Caption = "Cant. Servida"
        Option2(1).Value = False
        Check2.Visible = False
        Frame5.Enabled = False

End Select

Select Case Index
    
    Case 1
        
        MoverDatoGrilla

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

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

Dim RS As New ADODB.Recordset

Frame7.Visible = IIf(lc_Aux = "Planif", False, True)
fg_centra Me
EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico Planificacón Teórica"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Combo1(0).Clear
TvwDir.Nodes.Clear

If lc_Aux = "Planif" Then
   
   MsgTitulo = "Informe Planificación"
   Me.Caption = "Informe Planificación"
   Combo1(0).AddItem "Menú Mecano" & Space(150) & "(00)"
   Combo1(0).AddItem "Menú Mensual" & Space(150) & "(01)"
   Combo1(0).AddItem "Menú Mensual Servicios" & Space(150) & "(05)"
   Combo1(0).AddItem "Aporte Nutricionales Detallado" & Space(150) & "(02)"
   Combo1(0).AddItem "Aporte Nutricionales Resumido" & Space(150) & "(03)"
   Combo1(0).AddItem "Aporte Nutricionales por Estructura" & Space(150) & "(04)"
   Combo1(0).AddItem "Menú Mensual (Formato Comercial)" & Space(150) & "(06)"
   Combo1(0).AddItem "Aporte Nutricionales Detallado (Formato Comercial)" & Space(150) & "(07)"
   Combo1(0).AddItem "Aporte Nutricionales Resumido (Formato Comercial)" & Space(150) & "(08)"
   Combo1(0).AddItem "Aporte Nutricionales por Estructura (Formato Comercial)" & Space(150) & "(09)"
   Combo1(0).AddItem "Solo Tabla Gramaje (Formato Comercial)" & Space(150) & "(10)"
   Combo1(0).AddItem "Tabla Gramaje y Frecuencia (Formato Comercial)" & Space(150) & "(11)"
   Combo1(0).ListIndex = 0

Else
   
   MsgTitulo = "Informe Costo Minuta"
   Me.Caption = "Informe Costo Minuta"
   Combo1(0).AddItem "Costo Minuta Detallado" & Space(150) & "(0)"
   Combo1(0).AddItem "Costo Minuta por Estructura" & Space(150) & "(1)"
   Combo1(0).ListIndex = 0

End If

fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
accion = True

'-------> Llenar tabla nutrientes
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_s_nutriente 1, 0, ''")
If RS.EOF Then

   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe maestro nutrientes", vbExclamation + vbOKOnly, MsgTitulo
   Me.Hide
   Unload Me
   
End If

vaSpread2.MaxRows = 0
Do While Not RS.EOF
   
   vaSpread2.MaxRows = vaSpread2.MaxRows + 1
   vaSpread2.Row = vaSpread2.MaxRows
   vaSpread2.Col = 2
   vaSpread2.text = RS!nut_codigo
   
   vaSpread2.Col = 3
   vaSpread2.text = Trim(RS!nut_nombre) & " " & Trim(RS!nut_nomuni)
   
   If RS!nut_indpri = 1 Then
      
      vaSpread2.Col = 1
      vaSpread2.CellType = 10
      vaSpread2.TypeCheckText = ""
      vaSpread2.TypeCheckCenter = True
      vaSpread2.text = "1" ' checked
   
   Else
      
      vaSpread2.Col = 1
      vaSpread2.CellType = 10
      vaSpread2.TypeCheckText = ""
      vaSpread2.TypeCheckCenter = True
      vaSpread2.text = " " ' checked
   
   End If
   
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing

OpUsuario = vg_Indppr

If IsNull(OpUsuario) Or Trim(OpUsuario) = "" Then
    
    MsgBox "Contactese con el Administrador del Sistema...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub

Else
    
    Select Case OpUsuario
    
        Case "1"
            
            Combo1(1).Clear
            Combo1(1).AddItem "Real" & Space(150) & "(1)"
            Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
            vg_IndpprSelec = fg_buscacbo(Combo1, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
        
        Case "2"
            
            Combo1(1).Clear
            Combo1(1).AddItem "Propuesta" & Space(150) & "(2)"
            Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
            vg_IndpprSelec = fg_buscacbo(Combo1, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
        
        Case "3"
            
            Combo1(1).Clear
            Combo1(1).AddItem "Real" & Space(150) & "(1)"
            Combo1(1).AddItem "Propuesta" & Space(150) & "(2)"
            Combo1(1).ListIndex = 0
            vg_IndpprSelec = 1
    
    End Select
    
End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpDateTime1_Change(Index As Integer)

On Error GoTo Man_Error

If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
MoverDatoGrilla

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error


Dim RS As New ADODB.Recordset

Select Case Index
    
    Case 0
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        If vg_Indppr = 1 Or vg_Indppr = 2 Then
        
           Set RS = vg_db.Execute("sgpadm_s_subsegmento 10, " & Val(fpLongInteger1(0).Value) & ", '', '" & vg_Indppr & "'")
        
        Else
           
           Set RS = vg_db.Execute("sgpadm_s_subsegmento 1, " & Val(fpLongInteger1(0).Value) & ", '', ''")
        
        End If
        
        If RS.EOF Then
        
           RS.Close
           Set RS = Nothing
           fpayuda(0).Caption = ""
           Exit Sub
        
        End If
        
        fpayuda(0).Caption = Trim(RS!sub_nombre)
        RS.Close
        Set RS = Nothing
        MoverDatoGrilla
    
    Case 1
    
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        If vg_Indppr = 1 Or vg_Indppr = 2 Then
          
           Set RS = vg_db.Execute("sgpadm_Sel_RegimenRealPropuesta " & Val(fpLongInteger1(1).Value) & ", '" & vg_Indppr & "'")
        
        Else
          
           Set RS = vg_db.Execute("sgpadm_Sel_Regimen " & Val(fpLongInteger1(1).Value) & "")
        
        End If
        
        If RS.EOF Then
        
           RS.Close
           Set RS = Nothing
           fpayuda(1).Caption = ""
           Exit Sub
           
        End If
        
        fpayuda(1).Caption = Trim(RS!reg_nombre)
        RS.Close
        Set RS = Nothing
        MoverDatoGrilla
        
    Case 2
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        Set RS = vg_db.Execute("sgpadm_s_zona 1, " & Val(fpLongInteger1(2).Value) & ", ''")
        If RS.EOF Then
        
           RS.Close
           Set RS = Nothing
           fpayuda(2).Caption = ""
           Exit Sub
           
        End If
        fpayuda(2).Caption = Trim(RS!Zon_nombre)
        RS.Close
        Set RS = Nothing
        
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

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
    
End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Dim OpcionLectura As String

Select Case Index

    Case 0
        
        vg_left = fpayuda(0).Left + 2300
        vg_nombre = ""
        vg_codigo = ""
        B_TabEst.LlenaDatos "a_subsegmento", "sub_", "Subsegmento", "Sub"
        B_TabEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        fpLongInteger1(0).Value = Val(vg_codigo)
        fpayuda(0).Caption = vg_nombre
        MoverDatoGrilla
        fpLongInteger1(1).SetFocus
        
    Case 1
        
        vg_left = fpayuda(1).Left + 2300
        vg_nombre = ""
        vg_codigo = ""
        B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
        B_TabEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        fpLongInteger1(1).Value = Val(vg_codigo)
        fpayuda(1).Caption = vg_nombre
        fpLongInteger1(2).SetFocus
        
    Case 2
        
        vg_left = fpayuda(1).Left + 2300
        vg_nombre = ""
        vg_codigo = ""
        
        If fpLongInteger1(0).Value = "" Or fpLongInteger1(1).Value = "" Then
        
           Exit Sub
        
        End If
        
        Select Case Val(fg_codigocbo(Combo1, 0, 2, ""))
        
            Case 0, 1, 2, 3, 4, 5
                
                OpcionLectura = "1"
                B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", Val(fg_codigocbo(Combo1, 1, 1, ""))
                B_MTaEst.Show 1
            
            Case 6, 7, 8, 9, 10, 11
               
               OpcionLectura = "2"
               B_TabSel.LlenaDatos Me.TvwDir, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), Val(fg_codigocbo(Combo1, 1, 1, "")), OpcionLectura
               B_TabSel.Show 1
        
        End Select
        
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        
    Case 3
        
        vg_left = fpayuda(1).Left + 2300
        vg_nombre = ""
        vg_codigo = ""
        
        If fpLongInteger1(0).Value = "" Or fpLongInteger1(1).Value = "" Or fpLongInteger1(2).Value = "" Then
        
           Exit Sub
           
        End If
        
        B_MTaEst.LlenaDatos "Nutrientes", Me.vaSpread2, fpLongInteger1(0).Value, fpLongInteger1(1).Value, fpLongInteger1(2).Value, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "2", Val(fg_codigocbo(Combo1, 1, 1, ""))
        B_MTaEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        
    Case 4
        
        vg_left = fpayuda(1).Left + 2300
        vg_nombre = ""
        vg_codigo = ""
        B_TabEst.LlenaDatos "a_zona", "zon_", "Zona", "Zon"
        B_TabEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        fpLongInteger1(2).Value = Val(vg_codigo)
        fpayuda(2).Caption = vg_nombre
        fpDateTime1(0).SetFocus

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

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

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error


Dim RS       As New ADODB.Recordset
Dim opnomrec As Boolean
Dim spid     As Long

Select Case Button.Index

    Case 1
    
        If vaSpread1.MaxRows < 1 Then
        
           MsgBox "No existe Información Servicio", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
           
        End If
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        Set RS = vg_db.Execute("sgpadm_s_subsegmento 1, " & Val(fpLongInteger1(0).Value) & ", '', ''")
        If RS.EOF Then
        
           RS.Close
           Set RS = Nothing
           fpLongInteger1(0).Value = ""
           fpayuda(0).Caption = ""
           MsgBox "No existe subsegmento", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
           
        End If
        RS.Close
        Set RS = Nothing
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        Set RS = vg_db.Execute("sgpadm_s_regimen 1, " & Val(fpLongInteger1(1).Value) & ", ''")
        If RS.EOF Then
           
           RS.Close
           Set RS = Nothing
           fpayuda(1).Caption = ""
           MsgBox "No existe regimen", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
           
        End If
        RS.Close
        Set RS = Nothing
        
        Select Case Val(fg_codigocbo(Combo1, 0, 2, ""))
        
            Case 2, 3, 4, 7, 8, 9, 10, 11
               
               Set RS = vg_db.Execute("sgpadm_s_zona 1, " & Val(fpLongInteger1(2).Value) & ", ''")
               If RS.EOF Then
                  
                  RS.Close
                  Set RS = Nothing
                  fpayuda(2).Caption = ""
                  MsgBox "No existe zona", vbExclamation + vbOKOnly, MsgTitulo
                  Exit Sub
                  
               End If
               RS.Close
               Set RS = Nothing
        
        End Select
        If fpDateTime1(0).Value > fpDateTime1(1).Value Then
           
           MsgBox "Fecha origen Mayor destino", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
           
        End If
        
        If Mid(fpDateTime1(0).text, 4, 2) > Mid(fpDateTime1(1).text, 4, 2) Then
           
           MsgBox "Mes origen mayor destino", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If Mid(fpDateTime1(0).text, 7, 4) > Mid(fpDateTime1(1).text, 7, 4) Then
           
           MsgBox "Año origen mayor destino", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If Combo1(0).ListIndex = -1 Or Combo1(0).text = "" Then Exit Sub
        
        If Option1(0).Value = True Then
           
           For i = 1 To vaSpread1.MaxRows
               
               vaSpread1.Row = i
               vaSpread1.Col = 1
               vaSpread1.CellType = 10
               vaSpread1.TypeCheckText = ""
               vaSpread1.TypeCheckCenter = True
               vaSpread1.text = "1" ' checked
           
           Next i
           
           For i = 1 To TvwDir.Nodes.count
               
               TvwDir.Nodes.item(i).Checked = True
           
           Next i
        
        End If
        
        '-------> Borrar tabla paso servicio
        vg_db.Execute "DELETE paso_servicio WHERE ser_spid = @@spid and ser_usr = '" & vg_NUsr & "'"
        isel = 0
        
        '-------> Buscar spid
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        Set RS = vg_db.Execute("SELECT @@spid spid")
        If Not RS.EOF Then
           
           spid = RS!spid
           
        End If
        RS.Close
        Set RS = Nothing
        
        Select Case Val(fg_codigocbo(Combo1, 0, 2, ""))
            
            Case 0, 1, 2, 3, 4, 5
                
                For i = 1 To vaSpread1.MaxRows
                    vaSpread1.Row = i
                    vaSpread1.Col = 1
                    If vaSpread1.text = "1" Then
                        isel = 1
                        vaSpread1.Col = 2
                        vg_db.Execute "INSERT INTO paso_servicio (ser_spid, ser_usr, ser_codigo) VALUES (" & spid & ", '" & vg_NUsr & "', " & Val(vaSpread1.text) & ")"
                    End If
                Next i
                If isel = 0 Then fg_descarga: MsgBox "Servicio debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
            
            Case 6, 7, 8, 9, 10, 11
                
                isel = 0
                For i = 1 To TvwDir.Nodes.count
                    
                    If TvwDir.Nodes.item(i).Checked = True And InStr(TvwDir.Nodes.item(i).key, "EstServicio") <> 0 Then
                       
                       isel = 1
                       Exit For
                    
                    End If
                    
                Next
                
                If isel = 0 Then
                   
                   MsgBox "No ha seleccionado estructura de servicio...", vbExclamation + vbOKOnly, MsgTitulo
                   Exit Sub
                   
                End If
                isel = 0
                If Frame3(1).Enabled = True Then
                   
                   For i = 1 To vaSpread2.MaxRows
                       
                       vaSpread2.Row = i
                       vaSpread2.Col = 1
                       
                       If vaSpread2.text = "1" Then
                         
                          isel = 1
                      
                       End If
                   
                   Next i
                   
                   If isel = 0 Then
                      
                      fg_descarga
                      MsgBox "Nutriente debe ser informado", vbExclamation + vbOKOnly, MsgTitulo
                      Exit Sub
                   
                   End If
                   
                End If
        
        End Select
        opnomrec = True
        
        If Option1(9).Value = True Then
           
           opnomrec = False
           
        End If
        Toolbar1.Enabled = False
        Frame1(0).Enabled = False
        If lc_Aux = "Planif" Then
        vg_CallForm = Me.Name
        vg_CallFormDato = Combo1(0).text
        Select Case Val(fg_codigocbo(Combo1, 0, 2, ""))
           
           Case 0
              
              I_MenuPlanMecano Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, Val(fg_codigocbo(Combo1, 1, 1, "")), Option1(11), Check2.Value, Check3.Value
           
           Case 1
            
            If Check2.Value = 1 Then
                
                I_MenuPlanMensualSemanaCerradaok Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, Val(fg_codigocbo(Combo1, 1, 1, "")), Check1(0).Value, Check1(1).Value, Option1(11), Check2.Value, Check3.Value
            
            Else
                
                I_MenuPlanMensual Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, Val(fg_codigocbo(Combo1, 1, 1, "")), Check1(0).Value, Check1(1).Value, Option1(11), Check2.Value, Check3.Value
            
            End If
           
           Case 2
              
              I_AportePlanDetallado Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, spid, Val(fg_codigocbo(Combo1, 1, 1, "")), Check2.Value, Check3.Value
           
           Case 3
              
              I_AportePlanRes Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, spid, Val(fg_codigocbo(Combo1, 1, 1, "")), Check2.Value, Check3.Value
           
           Case 4
              
              I_AportePlanEstrRes Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, spid, Val(fg_codigocbo(Combo1, 1, 1, "")), Option1(11), Check2.Value
           
           Case 5
              
              I_MenuPlanMensualServicioOk Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, Val(fg_codigocbo(Combo1, 1, 1, "")), Check1(0).Value, Check1(1).Value, Option1(11), Check2.Value, spid, Check3.Value
           
           Case 6
              
              ExportarExcelMenuMensualMKT opnomrec, Check1(0).Value, Check1(1).Value, Option1(11), Check2.Value, Check3.Value
           
           Case 7
              
              ExportaExcelPlanDetalladoResumidoMKT "1", opnomrec, spid, Val(fg_codigocbo(Combo1, 1, 1, "")), Check2.Value, Check3.Value, 1
           
           Case 8
              
              ExportaExcelPlanDetalladoResumidoMKT "1", opnomrec, spid, Val(fg_codigocbo(Combo1, 1, 1, "")), Check2.Value, Check3.Value, 2
           
           Case 9
              
              ExportarExcelAportePlanEstrResMKT "1", opnomrec, spid, Val(fg_codigocbo(Combo1, 1, 1, "")), Option1(11), Check2.Value
           
           Case 10
              
              ExportarExcelSoloTabkaGramajeMKT
           
           Case 11
              
              ExportarExcelTabkaGramajeFrecuenciaMKT
        
        End Select
        Else
           Select Case Val(fg_codigocbo(Combo1, 0, 2, ""))
               
               Case 0
                  
                  I_CostoDetMinuta Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, spid, vg_NUsr, Val(fg_codigocbo(Combo1, 1, 1, "")), Option1(11), Check2.Value
               
               Case 1
                  
                  I_CostoPlanEstrRes Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, spid, vg_NUsr, Val(fg_codigocbo(Combo1, 1, 1, "")), Option1(11), Check2.Value, IIf(Option2(0).Value = True, True, False)
           
           End Select
        
        End If
        
        vg_db.Execute "DELETE paso_servicio WHERE ser_spid= " & spid & " AND ser_usr= '" & vg_NUsr & "'"
        Toolbar1.Enabled = True
        Frame1(0).Enabled = True
        
    Case 3
        
        'Historico Planificación
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        Set RS = vg_db.Execute("sgpadm_s_subsegmento 1, " & Val(fpLongInteger1(0).Value) & ", '', ''")
        If RS.EOF Then
           
           RS.Close
           Set RS = Nothing
           fpLongInteger1(0).Value = ""
           fpayuda(0).Caption = ""
           MsgBox "No existe subsegmento", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        RS.Close
        Set RS = Nothing
        vg_codigo = ""
        B_HistPm.LlenarHistPlan "Histórico Planificación Minuta", Val(fpLongInteger1(0).Value), 1, 1
        B_HistPm.Show 1
        If vg_codigo = "" Then Exit Sub
        fpLongInteger1(1).Value = vg_codregimen
        fpDateTime1(0).text = "01/" & vg_fecha: fpDateTime1(1).text = dEoM("01/" & vg_fecha)
        MoverDatoGrilla
        Option1(0).SetFocus
        Me.Refresh
        
    Case 5
    
        Me.Hide
        Unload Me
        
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Sub MoverDatoGrilla()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If Val(fpLongInteger1(0).Value) = 0 And Val(fpLongInteger1(1).Value) = 0 And fpDateTime1(0).text = "" And fpDateTime1(1).text = "" Then
   
   Exit Sub

End If

fg_carga ""
vaSpread1.MaxRows = 0
Set RS = vg_db.Execute("sgpadm_s_planifminuta 7, " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", 0,0, 0, " & Val(Format(fpDateTime1(0).text, "yyyymmdd")) & ", " & Val(Format(fpDateTime1(1).text, "yyyymmdd")) & "," & Val(fg_codigocbo(Combo1, 1, 1, "")) & "")
If RS.EOF Then

   RS.Close
   Set RS = Nothing
   TvwDir.Nodes.Clear
   fg_descarga
   Exit Sub
   
End If

Do While Not RS.EOF
   
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 2: vaSpread1.text = RS!min_codser
   vaSpread1.Col = 3: vaSpread1.text = Trim(RS!ser_nombre)
   
   If IsNull(RS!ser_nombre) Or RS!ser_nombre = "" Then
      
      RS.Close: Set RS = Nothing: fg_descarga
      MsgBox "Descripciones servicio con valor null o bien en blanco", vbExclamation + vbOKOnly, MsgTitulo
      vaSpread1.MaxRows = 0
      Exit Sub
   
   End If
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing
fg_descarga

If Val(fg_codigocbo(Combo1, 1, 1, "")) = 4 Then
' Llena Estructura minuta
End If

Dim AuxCodigoServicio    As Long
Dim AuxCodigoEstServicio As Long
Dim pcodser              As String
Set RS = vg_db.Execute("sgpadm_Sel_ServicioEstructura " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Val(Format(fpDateTime1(0).text, "yyyymmdd")) & ", " & Val(Format(fpDateTime1(1).text, "yyyymmdd")) & "")
TvwDir.Nodes.Clear
Do While Not RS.EOF
   
   If RS(0) <> AuxCodigoServicio Then
      
      Set rootNode = TvwDir.Nodes.Add(, , "N" & fg_pone_espacio(RS(0), 5), RS(0) & " - " & Trim(RS(1)))
      pcodser = "": pcodser = "N" & fg_pone_espacio(RS(0), 5): AuxCodigoServicio = RS(0)
   
   End If
   
   If RS!ess_codigo <> AuxCodigoEstServicio Then
      
      Set rootNode = TvwDir.Nodes.Add(pcodser, tvwChild, pcodser & "EstServicio" & fg_pone_espacio(RS!ess_codigo, 10), Trim(RS!ess_codigo) & " - " & Trim(RS!ess_nombre))
      AuxCodigoEstServicio = RS!ess_codigo
   
   End If
   
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing
fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Sub ExportarExcelMenuMensualMKT(opnomrec As Boolean, Raciones As Integer, Costo As Integer, opnomest As Boolean, opSemCerrada As Boolean, opparentesis As Boolean)

On Local Error GoTo Error

Dim RS As New ADODB.Recordset

Dim xj As Long
Dim X As Long
Dim j As Long
Dim p As Long
Dim i As Long
Dim ii As Long
Dim xx As Long

Dim VecDia(7) As Long

Dim CodigoServicio As Long
Dim CodigoEstServicio  As Long
Dim NombreServicio As String
Dim NombreServicioAux As String
Dim TipoMinuta As String

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object

Dim numdia As Long
Dim diafin As Long
Dim DiaMayor As Long
Dim AuxDiaMayor As Long
Dim AuxFechaCerrada As Long
Dim auxser As Long
Dim FecMin As String
Dim FecAux As Long
Dim NumLin As Long
Dim NumLinIni As Long
Dim MyBuffer  As String

'-------> Mover estado a la spread1
ii = 1
For i = 1 To TvwDir.Nodes.count
    
    If InStr(TvwDir.Nodes.item(i).key, "EstServicio") = 0 Then
       
       NombreServicio = CStr(Mid(TvwDir.Nodes.item(i).key, 2, Len(TvwDir.Nodes.item(i).key)))
       ii = vaSpread1.SearchCol(2, 0, vaSpread1.MaxRows, NombreServicio, SearchFlagsEqual)
       
       If ii > 0 Then
          
          vaSpread1.Row = ii
          vaSpread1.Row = ii: vaSpread1.Col = 1
          vaSpread1.text = "0"
          vaSpread1.text = IIf(TvwDir.Nodes.item(i).Checked = True, "1", "0")
       
       End If

    End If

Next i

CodigoServicio = 0
NombreServicio = ""
TipoMinuta = "1"
fg_carga ""

'-------> Start a new workbook in Excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add

    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        NumLin = 9
        NumLinIni = 9
        
        If vaSpread1.text = "1" Then
           
           vaSpread1.Col = 2
           CodigoServicio = vaSpread1.text
           vaSpread1.Col = 3
           NombreServicio = vaSpread1.text

            '-------> Buscar Nº días
            If opSemCerrada = False Then
               
               VecDia(7) = 1: VecDia(1) = 2: VecDia(2) = 3: VecDia(3) = 4: VecDia(4) = 5: VecDia(5) = 6: VecDia(6) = 7
            
            Else
               
               VecDia(7) = 7: VecDia(1) = 1: VecDia(2) = 2: VecDia(3) = 3: VecDia(4) = 4: VecDia(5) = 5: VecDia(6) = 6
            
            End If
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            Set RS = vg_db.Execute("SELECT DISTINCT b_minuta.min_fecmin FROM b_minuta, b_minutadet WHERE b_minuta.min_codigo = b_minutadet.mid_codigo " & _
                                   "AND b_minuta.min_subseg = " & Val(fpLongInteger1(0).Value) & " AND b_minuta.min_codreg = " & Val(fpLongInteger1(1).Value) & " AND b_minuta.min_codser = " & CodigoServicio & " AND b_minuta.min_Indppr = '" & Val(fg_codigocbo(Combo1, 1, 1, "")) & "' " & _
                                   "AND SUBSTRING(convert(char(8),b_minuta.min_fecmin),1 ,6) = " & Val(Format(fpDateTime1(0).text, "yyyymm")) & " AND b_minutadet.mid_tipmin = '" & TipoMinuta & "' ORDER BY b_minuta.min_fecmin")
            
            If RS.EOF Then
               
               fg_descarga
               RS.Close
               Set RS = Nothing
               Exit Sub
               
            End If
            Do While Not RS.EOF
               
               If Not opSemCerrada Then
                    
                    Select Case fg_Dia(RS!min_fecmin)
                    
                        Case 1
                        
                            numdia = 7
                        
                        Case 2
                            
                            numdia = 1
                        
                        Case 3
                            
                            numdia = 2
                        
                        Case 4
                            
                            numdia = 3
                        
                        Case 5
                            
                            numdia = 4
                        
                        Case 6
                            
                            numdia = 5
                        
                        Case 7
                            
                            numdia = 6
                            
                    End Select
                    
                    If numdia > DiaMayor Then
                       
                       DiaMayor = numdia
                       
                    End If
                    
               Else
               
                  If Val(Mid(RS!min_fecmin, 7, 2)) >= 1 And Val(Mid(RS!min_fecmin, 7, 2)) <= 7 Then
                     
                     icol = Val(Mid(RS!min_fecmin, 7, 2))
                   
                   ElseIf Val(Mid(RS!min_fecmin, 7, 2)) >= 8 And Val(Mid(RS!min_fecmin, 7, 2)) <= 14 Then
                     
                     icol = Val(Mid(RS!min_fecmin, 7, 2)) - 7
                   
                   ElseIf Val(Mid(RS!min_fecmin, 7, 2)) >= 15 And Val(Mid(RS!min_fecmin, 7, 2)) <= 21 Then
                    
                    icol = Val(Mid(RS!min_fecmin, 7, 2)) - 14
                   
                   ElseIf Val(Mid(RS!min_fecmin, 7, 2)) >= 22 And Val(Mid(RS!min_fecmin, 7, 2)) <= 28 Then
                    
                    icol = Val(Mid(RS!min_fecmin, 7, 2)) = -21
                   
                   ElseIf Val(Mid(RS!min_fecmin, 7, 2)) >= 29 And Val(Mid(RS!min_fecmin, 7, 2)) <= 35 Then
                    
                    icol = Val(Mid(RS!min_fecmin, 7, 2)) = -28
                  
                  End If
                  
                  If icol > DiaMayor Then DiaMayor = icol
                
                End If
                RS.MoveNext
            
            Loop
            RS.Close
            Set RS = Nothing
            numdia = DiaMayor
           
           '-------> Add data to cells of the first worksheet in the new workbook
           Set oSheet = oBook.Worksheets.Add
           NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
           
           If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
              
              NombreServicioAux = Mid(CodigoServicio & NombreServicioAux, 1, 31)
           
           End If
           oSheet.Name = NombreServicioAux
           
           MoverDatosExcel oExcel, oSheet, "A", "A", 2, 2, "S e t  de  M i n u t a"
           '-------> Imprimir Sub_Segmento
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           Set RS = vg_db.Execute("SELECT sub_nombre FROM a_subsegmento WHERE sub_codigo = " & Val(fpLongInteger1(0).Value) & "")
           If RS.EOF Then
               
               fg_descarga
               RS.Close
               Set RS = Nothing
               Exit Sub
               
           End If
           MoverDatosExcel oExcel, oSheet, "A", "A", 3, 3, Trim(RS!sub_nombre)
           RS.Close
           Set RS = Nothing
           '-------> Imprimir Servicio
           MoverDatosExcel oExcel, oSheet, "A", "A", 4, 4, NombreServicio
           '-------> Formatear celda
           PonerFontBold oExcel, oSheet, "A", "A", 2, 4

           PonerCombinarCentrar oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), 2, 2
           PonerCombinarCentrar oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), 3, 3
           PonerCombinarCentrar oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), 4, 4
           
           PonerTipoLetraTamaño oExcel, oSheet, "A", "A", 2, 4, 14
           
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Let MyBuffer = ""
           Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
           Let MyBuffer = MyBuffer & "<EstServicio>"

           For xx = 1 To TvwDir.Nodes.count
              
              If TvwDir.Nodes.item(xx).Checked = True And InStr(TvwDir.Nodes.item(xx).key, "EstServicio") <> 0 And CodigoServicio = LCase(Trim(Mid(TvwDir.Nodes.item(xx).key, 2, 5))) Then
                 
                 Let MyBuffer = MyBuffer & " <EstServicioDet"
                 CodigoEstServicio = LCase(Trim(Mid(TvwDir.Nodes.item(xx).key, 18, 10)))
                 Let MyBuffer = MyBuffer & " CodigoEstServicio = " & Chr(34) & CodigoEstServicio & Chr(34)
                 Let MyBuffer = MyBuffer & "/>"
              
              End If
           
           Next xx
           
           Let MyBuffer = MyBuffer & "</EstServicio>"
           Set RS = vg_db.Execute("sgpadm_Sel_MinutaMenuMensualxEstservicio '" & MyBuffer & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & CodigoServicio & ", " & Val(Format(fpDateTime1(0).text, "yyyymmdd")) & ", " & Val(Format(fpDateTime1(1).text, "yyyymmdd")) & ", " & Val(fg_codigocbo(Combo1, 1, 1, "")) & ", '" & TipoMinuta & "'")
           AuxFec = 0
           auxser = 0
           filmay = 0
           essnom = False
           xx = 1
           Dim AuxNumLinea As Long
           AuxNumLinea = 0
           
           If Not RS.EOF Then
              
              AuxNumLinea = RS!mid_numlin
              FecMin = RS!min_fecmin
              FecAux = RS!min_fecmin
              Est = True: dianva = False
              IniIncDia = 1
              estmes = False
              Do While Est
                 
                 If Mid(FecMin, 7, 2) <> "00" Then
                    
                    If fg_Dia(FecMin) = 2 Then Est = False: diaini = Mid(FecMin, 7, 2): Exit Do
                 
                 End If
                 
                 If Mid(FecMin, 7, 2) = "00" Then
                    
                    dianva = True
                    FecMin = Bom("01/" & Mid(FecMin, 5, 2) & "/" & Mid(FecMin, 1, 4))
                    FecMin = Mid(FecMin, 7, 4) & Mid(FecMin, 4, 2) & Mid(FecMin, 1, 2)
                 
                 Else
                    
                    FecMin = (FecMin - 1)
                 
                 End If
              Loop
              
              If Mid(RS!min_fecmin, 5, 2) <> Mid(FecMin, 5, 2) Then estmes = True
              diafin = Mid(dEoM("01/" & Mid(RS!min_fecmin, 5, 2) & "/" & Mid(RS!min_fecmin, 1, 4)), 1, 2)
              diafan = Val(Mid(Bom("01/" & Mid(RS!min_fecmin, 5, 2) & "/" & Mid(RS!min_fecmin, 1, 4)), 1, 2))
              
              For X = 1 To numdia + 1
                  
                  If X = 1 Then
                      
                      '-------> Imprimir Estructura
                      MoverDatosExcel oExcel, oSheet, "A", "A", 7, 7, "Estructura"
                      DibujarLineas oExcel, oSheet, "A", "A", 7, 7
                  
                  Else
                     
                     Select Case X
                        Case 2
                           
                           If opSemCerrada = False Then
                               
                               MoverDatosExcel oExcel, oSheet, "B", "B", 7, 7, "Lunes " & fg_pone_cero(Str(diaini), 2)
                               DibujarLineas oExcel, oSheet, "B", "B", 7, 7
                           
                           Else
                              
                              MoverDatosExcel oExcel, oSheet, "B", "B", 7, 7, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                              DibujarLineas oExcel, oSheet, "B", "B", 7, 7
                              IniIncDia = IniIncDia + 1
                           
                           End If
                        
                        Case 3
                           
                           If opSemCerrada = False Then
                              
                              MoverDatosExcel oExcel, oSheet, "C", "C", 7, 7, "Martes " & fg_pone_cero(Str(diaini), 2)
                              DibujarLineas oExcel, oSheet, "C", "C", 7, 7
                            
                            Else
                               
                               MoverDatosExcel oExcel, oSheet, "C", "C", 7, 7, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                               DibujarLineas oExcel, oSheet, "C", "C", 7, 7
                               IniIncDia = IniIncDia + 1
                           
                           End If
                        
                        Case 4
                           
                           If opSemCerrada = False Then
                               
                               MoverDatosExcel oExcel, oSheet, "D", "D", 7, 7, "Miércoles " & fg_pone_cero(Str(diaini), 2)
                               DibujarLineas oExcel, oSheet, "D", "D", 7, 7
                           
                           Else
                               
                               MoverDatosExcel oExcel, oSheet, "D", "D", 7, 7, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                               DibujarLineas oExcel, oSheet, "D", "D", 7, 7
                               IniIncDia = IniIncDia + 1
                           
                           End If
                        
                        Case 5
                           
                           If opSemCerrada = False Then
                              
                              MoverDatosExcel oExcel, oSheet, "E", "E", 7, 7, "Jueves " & fg_pone_cero(Str(diaini), 2)
                              DibujarLineas oExcel, oSheet, "E", "E", 7, 7
                           
                           Else
                              
                              MoverDatosExcel oExcel, oSheet, "E", "E", 7, 7, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                              DibujarLineas oExcel, oSheet, "E", "E", 7, 7
                              IniIncDia = IniIncDia + 1
                           
                           End If
                        
                        Case 6
                           
                           If opSemCerrada = False Then
                              
                              MoverDatosExcel oExcel, oSheet, "F", "F", 7, 7, "Viernes " & fg_pone_cero(Str(diaini), 2)
                              DibujarLineas oExcel, oSheet, "F", "F", 7, 7
                           
                           Else
                              
                              MoverDatosExcel oExcel, oSheet, "F", "F", 7, 7, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                              DibujarLineas oExcel, oSheet, "F", "F", 7, 7
                              IniIncDia = IniIncDia + 1
                           
                           End If
                        
                        Case 7
                           
                           If opSemCerrada = False Then
                              
                              MoverDatosExcel oExcel, oSheet, "G", "G", 7, 7, "Sábado " & fg_pone_cero(Str(diaini), 2)
                              DibujarLineas oExcel, oSheet, "G", "G", 7, 7
                           
                           Else
                              
                              MoverDatosExcel oExcel, oSheet, "G", "G", 7, 7, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                              DibujarLineas oExcel, oSheet, "G", "G", 7, 7
                              IniIncDia = IniIncDia + 1
                           
                           End If
                        
                        Case 8
                           
                           If opSemCerrada = False Then
                              
                              MoverDatosExcel oExcel, oSheet, "H", "H", 7, 7, "Domingo " & fg_pone_cero(Str(diaini), 2)
                              DibujarLineas oExcel, oSheet, "H", "H", 7, 7
                           
                           Else
                              
                              MoverDatosExcel oExcel, oSheet, "H", "H", 7, 7, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                              DibujarLineas oExcel, oSheet, "H", "H", 7, 7
                              IniIncDia = IniIncDia + 1
                           
                           End If
                     
                     End Select
                     
                     diaini = diaini + 1
                     If estmes Then
                        
                        If diaini > diafan And diafan > 0 Then diaini = 1: diafan = 0: dianva = False
                     
                     Else
                        
                        If diaini > diafin And diafin > 0 Then diaini = 1: diafan = 0: dianva = False
                     
                     End If
                  
                  End If
              
              Next X
              
              Bar1(0).Visible = True
              Bar1(0).Value = 0
              AuxFechaCerrada = 0
              AuxDiaMayor = 1
              ii = 1
              
              Do While Not RS.EOF
                 
                 Bar1(0).Value = Val((ii / RS.RecordCount) * 100)
                 
                 If RS!min_fecmin = 20131108 Then
                    
                    a = aa
                 
                 End If
                 
                 If RS!min_fecmin <> AuxFechaCerrada And opSemCerrada = True Then
                    
                    If AuxDiaMayor >= DiaMayor Then
                       
                       If AuxNumLinea > filmay Then filmay = AuxNumLinea
                        
                       NumLin = filmay
                       
                       PonerColorInterior oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLinIni - 2
                       PonerColorFont oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLinIni - 2
                       PonerNegrilla oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLinIni - 2
                       PonerTipoLetraTamaño oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLin, 12
                       PonerCentrado oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLin
                       PonerNegrilla oExcel, oSheet, "A", "A", NumLinIni - 2, NumLin

                       '-------> DibujarRectangulo

                       'numlin = numlin + 1
                       NumLin = filmay + 1 'numlin + 1
                       For xj = 1 To DiaMayor + 1
                           DibujarLineas oExcel, oSheet, Chr(xj + 64), Chr(xj + 64), NumLinIni - 1, NumLin
                       Next xj
                       NumLin = NumLin + 3
                       
                       filmay = 0
                       Idia = Mid(RS!min_fecmin, 7, 2): dianva = False
                       If DatePart("w", fg_Ctod1(RS!min_fecmin), 2) <> fg_Dia(RS!min_fecmin) Then Idia = Idia - fg_Dia(RS!min_fecmin) + 2
                       For X = 1 To numdia + 1
                           If X = 1 Then
                              MoverDatosExcel oExcel, oSheet, "A", "A", NumLin, NumLin, "Estructura"
                              DibujarLineas oExcel, oSheet, "A", "A", NumLin, NumLin

                           Else
                              Select Case X
                              Case 2
                                  MoverDatosExcel oExcel, oSheet, "B", "B", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                  DibujarLineas oExcel, oSheet, "B", "B", NumLin, NumLin
                                  IniIncDia = IniIncDia + 1
                              Case 3
                                   MoverDatosExcel oExcel, oSheet, "C", "C", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                   DibujarLineas oExcel, oSheet, "C", "C", NumLin, NumLin
                                   IniIncDia = IniIncDia + 1
                              Case 4
                                   MoverDatosExcel oExcel, oSheet, "D", "D", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                   DibujarLineas oExcel, oSheet, "D", "D", NumLin, NumLin
                                   IniIncDia = IniIncDia + 1
                              Case 5
                                   MoverDatosExcel oExcel, oSheet, "E", "E", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                   DibujarLineas oExcel, oSheet, "E", "E", NumLin, NumLin
                                   IniIncDia = IniIncDia + 1
                              Case 6
                                   MoverDatosExcel oExcel, oSheet, "F", "F", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                   DibujarLineas oExcel, oSheet, "F", "F", NumLin, NumLin
                                   IniIncDia = IniIncDia + 1
                              Case 7
                                   MoverDatosExcel oExcel, oSheet, "G", "G", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                   DibujarLineas oExcel, oSheet, "G", "G", NumLin, NumLin
                                   IniIncDia = IniIncDia + 1
                              Case 8
                                   MoverDatosExcel oExcel, oSheet, "H", "H", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                   DibujarLineas oExcel, oSheet, "H", "H", NumLin, NumLin
                                   IniIncDia = IniIncDia + 1
                              End Select
                              If Idia >= diafin Then Idia = 0: dianva = True
                              Idia = Idia + 1
                           End If
                       Next X
                       NumLin = NumLin + 2
                       NumLinIni = NumLin
                       AuxNumLinea = RS!mid_numlin
                       AuxDiaMayor = 1 '2
                    Else
                    If Val(Mid(RS!min_fecmin, 7, 2)) >= 1 And Val(Mid(RS!min_fecmin, 7, 2)) <= 7 Then
                       AuxDiaMayor = Val(Mid(RS!min_fecmin, 7, 2))
                    ElseIf Val(Mid(RS!min_fecmin, 7, 2)) >= 8 And Val(Mid(RS!min_fecmin, 7, 2)) <= 14 Then
                       AuxDiaMayor = Val(Mid(RS!min_fecmin, 7, 2)) - 7
                    ElseIf Val(Mid(RS!min_fecmin, 7, 2)) >= 15 And Val(Mid(RS!min_fecmin, 7, 2)) <= 21 Then
                       AuxDiaMayor = Val(Mid(RS!min_fecmin, 7, 2)) - 14
                    ElseIf Val(Mid(RS!min_fecmin, 7, 2)) >= 22 And Val(Mid(RS!min_fecmin, 7, 2)) <= 28 Then
                      AuxDiaMayor = Val(Mid(RS!min_fecmin, 7, 2)) - 21
                    ElseIf Val(Mid(RS!min_fecmin, 7, 2)) >= 29 And Val(Mid(RS!min_fecmin, 7, 2)) <= 35 Then
                      AuxDiaMayor = Val(Mid(RS!min_fecmin, 7, 2)) - 28
                    End If
                       
                    End If
                    AuxFechaCerrada = RS!min_fecmin
                    auxser = 0
                 ElseIf DatePart("ww", fg_Ctod1(RS!min_fecmin), 2) <> AuxFec And opSemCerrada = False Then
                    If AuxFec > 0 Then
                       NumLin = filmay
                       
                       PonerColorInterior oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLinIni - 2
                       PonerColorFont oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLinIni - 2
                       PonerNegrilla oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLinIni - 2
                       PonerTipoLetraTamaño oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLin, 12
                       PonerCentrado oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLin
                       PonerNegrilla oExcel, oSheet, "A", "A", NumLinIni - 2, NumLin

                       '-------> DibujarRectangulo
                       NumLin = filmay + 1 'numlin + 1
                       For xj = 1 To DiaMayor + 1
                           DibujarLineas oExcel, oSheet, Chr(xj + 64), Chr(xj + 64), NumLinIni - 1, NumLin
                       Next xj
                       filmay = 0
                       NumLin = NumLin + 3
                       
                       filmay = 0
                       Idia = Mid(RS!min_fecmin, 7, 2): dianva = False
                       If DatePart("w", fg_Ctod1(RS!min_fecmin), 2) <> fg_Dia(RS!min_fecmin) Then Idia = Idia - fg_Dia(RS!min_fecmin) + 2
                       For X = 1 To numdia + 1
                           If X = 1 Then
                              MoverDatosExcel oExcel, oSheet, "A", "A", NumLin, NumLin, "Estructura"
                              DibujarLineas oExcel, oSheet, "A", "A", NumLin, NumLin

                           Else
                              Select Case X
                              Case 2
                                If opSemCerrada = False Then
                                   MoverDatosExcel oExcel, oSheet, "B", "B", NumLin, NumLin, "Lunes " & IIf(Idia > diafin, "", fg_pone_cero(Str(Idia), 2))
                                   DibujarLineas oExcel, oSheet, "B", "B", NumLin, NumLin
                                Else
                                   MoverDatosExcel oExcel, oSheet, "B", "B", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                   DibujarLineas oExcel, oSheet, "B", "B", NumLin, NumLin
                                   IniIncDia = IniIncDia + 1
                                End If
                              Case 3
                                If opSemCerrada = False Then
                                   MoverDatosExcel oExcel, oSheet, "C", "C", NumLin, NumLin, "Martes " & IIf(Idia > diafin, "", fg_pone_cero(Str(Idia), 2))
                                   DibujarLineas oExcel, oSheet, "C", "C", NumLin, NumLin
                                Else
                                   MoverDatosExcel oExcel, oSheet, "C", "C", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                   DibujarLineas oExcel, oSheet, "C", "C", NumLin, NumLin
                                   IniIncDia = IniIncDia + 1
                                End If
                              Case 4
                                If opSemCerrada = False Then
                                   MoverDatosExcel oExcel, oSheet, "D", "D", NumLin, NumLin, "Miércoles " & IIf(Idia > diafin, "", fg_pone_cero(Str(Idia), 2))
                                   DibujarLineas oExcel, oSheet, "D", "D", NumLin, NumLin
                                Else
                                   MoverDatosExcel oExcel, oSheet, "D", "D", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                   DibujarLineas oExcel, oSheet, "D", "D", NumLin, NumLin
                                   IniIncDia = IniIncDia + 1
                                End If
                              Case 5
                                If opSemCerrada = False Then
                                   MoverDatosExcel oExcel, oSheet, "E", "E", NumLin, NumLin, "Jueves " & IIf(Idia > diafin, "", fg_pone_cero(Str(Idia), 2))
                                   DibujarLineas oExcel, oSheet, "E", "E", NumLin, NumLin
                                Else
                                   MoverDatosExcel oExcel, oSheet, "E", "E", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                   DibujarLineas oExcel, oSheet, "E", "E", NumLin, NumLin
                                   IniIncDia = IniIncDia + 1
                                End If
                              Case 6
                                If opSemCerrada = False Then
                                   MoverDatosExcel oExcel, oSheet, "F", "F", NumLin, NumLin, "Viernes " & IIf(Idia > diafin, "", fg_pone_cero(Str(Idia), 2))
                                   DibujarLineas oExcel, oSheet, "F", "F", NumLin, NumLin
                                Else
                                   MoverDatosExcel oExcel, oSheet, "F", "F", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                   DibujarLineas oExcel, oSheet, "F", "F", NumLin, NumLin
                                   IniIncDia = IniIncDia + 1
                                End If
                              Case 7
                                If opSemCerrada = False Then
                                   MoverDatosExcel oExcel, oSheet, "G", "G", NumLin, NumLin, "Sábado " & IIf(Idia > diafin, "", fg_pone_cero(Str(Idia), 2))
                                   DibujarLineas oExcel, oSheet, "G", "G", NumLin, NumLin
                                Else
                                   MoverDatosExcel oExcel, oSheet, "G", "G", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                   DibujarLineas oExcel, oSheet, "G", "G", NumLin, NumLin
                                   IniIncDia = IniIncDia + 1
                                End If
                              Case 8
                                If opSemCerrada = False Then
                                   MoverDatosExcel oExcel, oSheet, "H", "H", NumLin, NumLin, "Domingo " & IIf(Idia > diafin, "", fg_pone_cero(Str(Idia), 2))
                                   DibujarLineas oExcel, oSheet, "H", "H", NumLin, NumLin
                                Else
                                   MoverDatosExcel oExcel, oSheet, "H", "H", NumLin, NumLin, "DIA " & fg_pone_cero(Str(IniIncDia), 2)
                                   DibujarLineas oExcel, oSheet, "H", "H", NumLin, NumLin
                                   IniIncDia = IniIncDia + 1
                                End If
                              End Select
                              If Idia >= diafin Then Idia = 0: dianva = True
                              Idia = Idia + 1
                           End If
                       Next X
                       NumLin = NumLin + 2
                       NumLinIni = NumLin
                    End If
                    auxser = 0
                    AuxFec = DatePart("ww", fg_Ctod1(RS!min_fecmin), 2)
                    AuxNumLinea = RS!mid_numlin
                 End If
              
                 NumLin = NumLin + (RS!mid_numlin - AuxNumLinea)
                 If NumLin > filmay Then filmay = NumLin
                 essnom = False
                 If RS!mid_estser <> auxser Then
                    '------- Revisar si existe una estructura fija
                    essnom = False
                    If NumLin < NumLinIni Then NumLin = NumLinIni
                    If essnom = False Then
                       MoverDatosExcel oExcel, oSheet, "A", "A", NumLin, NumLin, IIf(opnomest = True, IIf(IIf(IsNull(RS!mid_desest) = True, "", RS!mid_desest) <> "", RS!mid_desest, RS!ess_nombre), RS!ess_nombre)
                    End If
                    '------- Fin revisar si existe una estructura fija
                    auxser = RS!mid_estser
                 End If
                 AuxNumLinea = RS!mid_numlin
                 '-------> Mover detalle minuta
                 If opSemCerrada = False Then
                    For X = 1 To 7
                        If VecDia(X) = fg_Dia(RS!min_fecmin) Then
                           MoverDatosExcel oExcel, oSheet, Chr(X + 65), Chr(X + 65), IIf(essnom, p, NumLin), IIf(essnom, p, NumLin), IIf(opnomrec = True, IIf(Not opparentesis, ExtraeParentesis(IIf(IsNull(RS!rec_nomfan) = True, "", RS!rec_nomfan)), IIf(IsNull(RS!rec_nomfan) = True, "", RS!rec_nomfan)) & " " & IIf(Raciones = 1, "( " & RS!mid_numrac & " raciones)", "") & " " & IIf(Costo = 1, " - Costo uni. $ " & Format((RS!mid_cosrec), fg_Pict(6, 2)), ""), IIf(Not opparentesis, ExtraeParentesis(IIf(IsNull(RS!rec_nombre) = True, "", RS!rec_nombre)), IIf(IsNull(RS!rec_nombre) = True, "", RS!rec_nombre)) & " " & IIf(Raciones = 1, "( " & RS!mid_numrac & " raciones)", "") & " " & IIf(Costo = 1, " - Costo uni. $ " & Format((RS!mid_cosrec), fg_Pict(6, 2)), ""))
                        End If
                    Next X
                 Else
                    If Val(Mid(RS!min_fecmin, 7, 2)) >= 1 And Val(Mid(RS!min_fecmin, 7, 2)) <= 7 Then
                       X = Val(Mid(RS!min_fecmin, 7, 2))
                    ElseIf Val(Mid(RS!min_fecmin, 7, 2)) >= 8 And Val(Mid(RS!min_fecmin, 7, 2)) <= 14 Then
                       X = Val(Mid(RS!min_fecmin, 7, 2)) - 7
                    ElseIf Val(Mid(RS!min_fecmin, 7, 2)) >= 15 And Val(Mid(RS!min_fecmin, 7, 2)) <= 21 Then
                       X = Val(Mid(RS!min_fecmin, 7, 2)) - 14
                    ElseIf Val(Mid(RS!min_fecmin, 7, 2)) >= 22 And Val(Mid(RS!min_fecmin, 7, 2)) <= 28 Then
                      X = Val(Mid(RS!min_fecmin, 7, 2)) - 21
                    ElseIf Val(Mid(RS!min_fecmin, 7, 2)) >= 29 And Val(Mid(RS!min_fecmin, 7, 2)) <= 35 Then
                      X = Val(Mid(RS!min_fecmin, 7, 2)) - 28
                    End If
                    MoverDatosExcel oExcel, oSheet, Chr(X + 65), Chr(X + 65), IIf(essnom, p, NumLin), IIf(essnom, p, NumLin), IIf(opnomrec = True, IIf(Not opparentesis, ExtraeParentesis(IIf(IsNull(RS!rec_nomfan) = True, "", RS!rec_nomfan)), IIf(IsNull(RS!rec_nomfan) = True, "", RS!rec_nomfan)) & " " & IIf(Raciones = 1, "( " & RS!mid_numrac & " raciones)", "") & " " & IIf(Costo = 1, " - Costo uni. $ " & Format((RS!mid_cosrec), fg_Pict(6, 2)), ""), IIf(Not opparentesis, ExtraeParentesis(IIf(IsNull(RS!rec_nombre) = True, "", RS!rec_nombre)), IIf(IsNull(RS!rec_nombre) = True, "", RS!rec_nombre)) & " " & IIf(Raciones = 1, "( " & RS!mid_numrac & " raciones)", "") & " " & IIf(Costo = 1, " - Costo uni. $ " & Format((RS!mid_cosrec), fg_Pict(6, 2)), ""))
                 End If
                 
                 RS.MoveNext
                 ii = ii + 1
              Loop
              NumLin = filmay
              
              PonerColorInterior oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLinIni - 2
              PonerColorFont oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLinIni - 2
              PonerNegrilla oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLinIni - 2
              PonerTipoLetraTamaño oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLin, 12
              PonerCentrado oExcel, oSheet, "A", Chr(64 + DiaMayor + 1), NumLinIni - 2, NumLin
              PonerNegrilla oExcel, oSheet, "A", "A", NumLinIni - 2, NumLin
              
              '-------> DibujarRectangulo
              'numlin = numlin + 1
              NumLin = filmay + 1
              For xj = 1 To DiaMayor + 1
                 DibujarLineas oExcel, oSheet, Chr(xj + 64), Chr(xj + 64), NumLinIni - 1, NumLin
              Next xj
              NumLin = NumLin + 3
              '--------> Determinar ancho de columna
              oSheet.Cells.Select
              oExcel.Selection.ColumnWidth = 40#
              '-------> Ajustar Texto
              oSheet.Cells.Select
              With oExcel.Selection
                   .WrapText = True
                   .Orientation = 0
                   .AddIndent = False
                   .ShrinkToFit = False
                   .ReadingOrder = xlContext
              End With
              '-------> Sacar salto de pagina y linea divisora
              oExcel.ActiveWindow.DisplayGridLines = False
              VistaPreliminarExcel oExcel, oSheet, True
           End If
           RS.Close
           Set RS = Nothing
           Bar1(0).Visible = False
           Bar1(0).Value = 0
        
        End If
    Next i
    oExcel.Visible = True '------->Visualizar
    Set oSheet = Nothing
    Set oExcel = Nothing
    Set oBook = Nothing
    fg_descarga
    
    Exit Sub
Error:
    Bar1(0).Visible = False
    Bar1(0).Value = 0
    fg_descarga
    oExcel.DisplayAlerts = False
    oExcel.Quit
    oExcel.DisplayAlerts = True
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Exit Sub

End Sub

Sub ExportarExcelSoloTabkaGramajeMKT()

On Local Error GoTo Error

Dim RS As New ADODB.Recordset

Dim ii As Long
Dim i As Long
Dim xx As Long
Dim NumLinExcel As Long

Dim CodigoServicio As Long
Dim CodigoEstServicio  As Long
Dim NombreServicio As String
Dim NombreServicioAux As String
Dim TipoMinuta As String

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object

Dim AuxPriTipoPlato As String
Dim AuxSegTipoPlato As String
Dim AuxSegTipoPlato1 As String
Dim CantBrutaServida As String

Dim MyBuffer    As String

'-------> Mover estado a la spread1
ii = 1
For i = 1 To TvwDir.Nodes.count
    If InStr(TvwDir.Nodes.item(i).key, "EstServicio") = 0 Then
           NombreServicio = CStr(Mid(TvwDir.Nodes.item(i).key, 2, Len(TvwDir.Nodes.item(i).key)))
           ii = vaSpread1.SearchCol(2, 0, vaSpread1.MaxRows, NombreServicio, SearchFlagsEqual)
           If ii > 0 Then
              vaSpread1.Row = ii
              vaSpread1.Row = ii: vaSpread1.Col = 1
              vaSpread1.text = "0"
              vaSpread1.text = IIf(TvwDir.Nodes.item(i).Checked = True, "1", "0")
           End If
    End If
Next i

NombreServicio = ""
TipoMinuta = "1"
fg_carga ""

'-------> Start a new workbook in Excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add

    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        NumLin = 9
        NumLinIni = 9
        If vaSpread1.text = "1" Then
           vaSpread1.Col = 2: CodigoServicio = vaSpread1.text
           vaSpread1.Col = 3: NombreServicio = vaSpread1.text

           '-------> Add data to cells of the first worksheet in the new workbook
           Set oSheet = oBook.Worksheets.Add
          
           NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
           If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
              NombreServicioAux = Mid(CodigoServicio & NombreServicioAux, 1, 31)
           End If
           oSheet.Name = NombreServicioAux 'Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31) 'NombreServicio
           
           MoverDatosExcel oExcel, oSheet, "A", "A", 2, 2, "TABLA   DE  GRAMAJES"
           '-------> Imprimir Sub_Segmento
           Set RS = vg_db.Execute("SELECT sub_nombre FROM a_subsegmento WHERE sub_codigo = " & Val(fpLongInteger1(0).Value) & "")
           If RS.EOF Then fg_descarga: RS.Close: Set RS = Nothing: Exit Sub
           MoverDatosExcel oExcel, oSheet, "A", "A", 3, 3, Trim(RS!sub_nombre)
           RS.Close: Set RS = Nothing
           '-------> Imprimir Servicio
           MoverDatosExcel oExcel, oSheet, "A", "A", 4, 4, "Servicio " & NombreServicio
           '-------> Formatear celda
           PonerFontBold oExcel, oSheet, "A", "A", 2, 4
           PonerCombinarCentrar oExcel, oSheet, "A", "F", 2, 2
           PonerCombinarCentrar oExcel, oSheet, "A", "F", 3, 3
           PonerCombinarCentrar oExcel, oSheet, "A", "F", 4, 4
           PonerTipoLetraTamaño oExcel, oSheet, "A", "A", 2, 4, 14
           
           Let MyBuffer = ""
           Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
           Let MyBuffer = MyBuffer & "<EstServicio>"

           For xx = 1 To TvwDir.Nodes.count
              If TvwDir.Nodes.item(xx).Checked = True And InStr(TvwDir.Nodes.item(xx).key, "EstServicio") <> 0 And CodigoServicio = LCase(Trim(Mid(TvwDir.Nodes.item(xx).key, 2, 5))) Then
                 Let MyBuffer = MyBuffer & " <EstServicioDet"
                 CodigoEstServicio = LCase(Trim(Mid(TvwDir.Nodes.item(xx).key, 18, 10)))
                 Let MyBuffer = MyBuffer & " CodigoEstServicio = " & Chr(34) & CodigoEstServicio & Chr(34)
                 Let MyBuffer = MyBuffer & "/>"
              End If
           Next xx
           
           Let MyBuffer = MyBuffer & "</EstServicio>"
           
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           Set RS = vg_db.Execute("sgpadm_Sel_MinutaTablaGramajeFrecuencia '" & MyBuffer & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & CodigoServicio & ", " & Val(fpLongInteger1(2).Value) & ", " & Val(Format(fpDateTime1(0).text, "yyyymmdd")) & ", " & Val(Format(fpDateTime1(1).text, "yyyymmdd")) & ", '" & Val(fg_codigocbo(Combo1, 1, 1, "")) & "', '" & TipoMinuta & "', " & IIf(Option2(0).Value = True, "1", "2") & "")
           AuxPriTipoPlato = ""
           AuxSegTipoPlato = ""
           CantBrutaServida = ""
           NumLinExcel = 9
           Bar1(0).Visible = True
           Bar1(0).Value = 0
           ii = 1
           If Not RS.EOF Then
              Do While Not RS.EOF
                 Bar1(0).Value = Val((ii / RS.RecordCount) * 100)
                 If RS!PrimeraCategoria <> AuxPriTipoPlato Then
                    If AuxSegTipoPlato <> "" Then
                       If Trim(CantBrutaServida) <> "" Then
                          CantBrutaServida = Mid(CantBrutaServida, 1, Len(CantBrutaServida) - 3)
                       End If
                       MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, " " & CantBrutaServida
                       DibujarLineas oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel
                       PonerTipoLetraTamaño oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, 12
                       PonerCentrado oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel
                       CantBrutaServida = ""
                       NumLinExcel = NumLinExcel + 2
                    End If
                    MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Trim(RS!PrimeraCategoria)
                    PonerCombinarLeft oExcel, oSheet, "B", "D", NumLinExcel, NumLinExcel
                    PonerFontBold oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel
                    PonerTipoLetraTamaño oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, 12
                    If ii = 1 Then
                       MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, "Gramos"
                       PonerFontBold oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel
                       PonerTipoLetraTamaño oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, 12
                       NumLinExcel = NumLinExcel + 1
                       
                       MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, IIf(Option2(0).Value = True, "Peso Bruto", "Servido")
                       PonerFontBold oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel
                       PonerTipoLetraTamaño oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, 12
                       NumLinExcel = NumLinExcel + 1
                    End If
                    AuxPriTipoPlato = RS!PrimeraCategoria
                    AuxSegTipoPlato = ""
                    CantBrutaServida = ""
                    NumLinExcel = NumLinExcel + 1
                 End If
                 
                 If InStr(RS!SegundaCategoria, "\") = 0 Then
                    Sql = RS!SegundaCategoria
                 Else
                    Sql = Mid(RS!SegundaCategoria, 1, InStr(RS!SegundaCategoria, "\") - 1)
                 End If
                 
                 If RS!SegundaCategoria <> AuxSegTipoPlato Then
                    If AuxSegTipoPlato <> "" Then
                       If Trim(CantBrutaServida) <> "" Then
                          CantBrutaServida = Mid(CantBrutaServida, 1, Len(CantBrutaServida) - 3)
                       End If
                       MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, " " & CantBrutaServida
                       DibujarLineas oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel
                       CantBrutaServida = ""
                       If Sql <> AuxSegTipoPlato1 Then
                          NumLinExcel = NumLinExcel + IIf(AuxSegTipoPlato1 <> "", 2, 1)
                          AuxSegTipoPlato1 = Sql
                       Else
                          NumLinExcel = NumLinExcel + 1
                       End If

                    End If
                    MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, "° " & Trim(RS!SegundaCategoria)
                    PonerCombinarLeft oExcel, oSheet, "B", "D", NumLinExcel, NumLinExcel
                    PonerTipoLetraTamaño oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, 12
                    AuxSegTipoPlato = RS!SegundaCategoria
                 End If
                 CantBrutaServida = CantBrutaServida & IIf(RS!valor = 0, "", RS!valor & "  -  ")

                 RS.MoveNext: ii = ii + 1
              Loop
              If Trim(CantBrutaServida) <> "" Then
                 CantBrutaServida = Mid(CantBrutaServida, 1, Len(CantBrutaServida) - 3)
              End If
              MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, " " & CantBrutaServida
              DibujarLineas oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel
              PonerTipoLetraTamaño oExcel, oSheet, "E", "E", 9, NumLinExcel, 12
              PonerCentrado oExcel, oSheet, "E", "E", 9, NumLinExcel
              CantBrutaServida = ""
              NumLinExcel = NumLinExcel + 1
              '-------> Dibujar Ancho Columna Frecuencia
              oSheet.Columns("E:E").Select
              oExcel.Selection.ColumnWidth = 20
              '-------> Ajustar Texto
              oSheet.Cells.Select
              With oExcel.Selection
                   .WrapText = True
                   .Orientation = 0
                   .AddIndent = False
                   .ShrinkToFit = False
                   .ReadingOrder = xlContext
              End With
              '-------> Sacar salto de pagina y linea divisora
              oExcel.ActiveWindow.DisplayGridLines = False
              VistaPreliminarExcel oExcel, oSheet, False
           End If
           RS.Close
           Set RS = Nothing
           Bar1(0).Visible = False
           Bar1(0).Value = 0
        End If
    Next i
    oExcel.Visible = True '------->Visualizar
    Set oSheet = Nothing
    Set oExcel = Nothing
    Set oBook = Nothing
    fg_descarga
    
    Exit Sub
Error:
    Bar1(0).Visible = False
    Bar1(0).Value = 0
    fg_descarga
    oExcel.DisplayAlerts = False
    oExcel.Quit
    oExcel.DisplayAlerts = True
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Exit Sub

End Sub

Sub ExportarExcelTabkaGramajeFrecuenciaMKT()

On Local Error GoTo Error

Dim RS As New ADODB.Recordset

Dim ii As Long
Dim i As Long
Dim xx As Long
Dim NumLinExcel As Long

Dim CodigoServicio As Long
Dim CodigoEstServicio  As Long
Dim NombreServicio As String
Dim NombreServicioAux As String
Dim TipoMinuta As String

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object

Dim AuxPriTipoPlato As String
Dim AuxSegTipoPlato As String
Dim CantBrutaServida As String
Dim CantBrutaServida1 As String
Dim AuxFrecuencia As String

Dim MyBuffer    As String

'-------> Mover estado a la spread1
ii = 1
For i = 1 To TvwDir.Nodes.count
    If InStr(TvwDir.Nodes.item(i).key, "EstServicio") = 0 Then
           NombreServicio = CStr(Mid(TvwDir.Nodes.item(i).key, 2, Len(TvwDir.Nodes.item(i).key)))
           ii = vaSpread1.SearchCol(2, 0, vaSpread1.MaxRows, NombreServicio, SearchFlagsEqual)
           If ii > 0 Then
              vaSpread1.Row = ii
              vaSpread1.Row = ii: vaSpread1.Col = 1
              vaSpread1.text = "0"
              vaSpread1.text = IIf(TvwDir.Nodes.item(i).Checked = True, "1", "0")
           End If
    
    End If
Next i

NombreServicio = ""
TipoMinuta = "1"
fg_carga ""

'-------> Start a new workbook in Excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add

    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        NumLin = 9
        NumLinIni = 9
        If vaSpread1.text = "1" Then
           vaSpread1.Col = 2: CodigoServicio = vaSpread1.text
           vaSpread1.Col = 3: NombreServicio = vaSpread1.text

           '-------> Add data to cells of the first worksheet in the new workbook
           Set oSheet = oBook.Worksheets.Add
           NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
           If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
              NombreServicioAux = Mid(CodigoServicio & NombreServicioAux, 1, 31)
           End If
           oSheet.Name = NombreServicioAux 'Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31) 'NombreServicio
           
           MoverDatosExcel oExcel, oSheet, "A", "A", 2, 2, "TABLA   DE  GRAMAJES"
           '-------> Imprimir Sub_Segmento
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           Set RS = vg_db.Execute("SELECT sub_nombre FROM a_subsegmento WHERE sub_codigo = " & Val(fpLongInteger1(0).Value) & "")
           If RS.EOF Then
              
              fg_descarga
              RS.Close
              Set RS = Nothing
              Exit Sub
              
           End If
           MoverDatosExcel oExcel, oSheet, "A", "A", 3, 3, Trim(RS!sub_nombre)
           RS.Close: Set RS = Nothing
           '-------> Imprimir Servicio
           MoverDatosExcel oExcel, oSheet, "A", "A", 4, 4, "Servicio " & NombreServicio
           '-------> Formatear celda
           PonerFontBold oExcel, oSheet, "A", "A", 2, 4
           PonerCombinarCentrar oExcel, oSheet, "A", "F", 2, 2
           PonerCombinarCentrar oExcel, oSheet, "A", "F", 3, 3
           PonerCombinarCentrar oExcel, oSheet, "A", "F", 4, 4
           PonerTipoLetraTamaño oExcel, oSheet, "A", "A", 2, 4, 14
           
           Let MyBuffer = ""
           Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
           Let MyBuffer = MyBuffer & "<EstServicio>"

           For xx = 1 To TvwDir.Nodes.count
              If TvwDir.Nodes.item(xx).Checked = True And InStr(TvwDir.Nodes.item(xx).key, "EstServicio") <> 0 And CodigoServicio = LCase(Trim(Mid(TvwDir.Nodes.item(xx).key, 2, 5))) Then
                 Let MyBuffer = MyBuffer & " <EstServicioDet"
                 CodigoEstServicio = LCase(Trim(Mid(TvwDir.Nodes.item(xx).key, 18, 10)))
                 Let MyBuffer = MyBuffer & " CodigoEstServicio = " & Chr(34) & CodigoEstServicio & Chr(34)
                 Let MyBuffer = MyBuffer & "/>"
              End If
           Next xx
           Let MyBuffer = MyBuffer & "</EstServicio>"
           
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           Set RS = vg_db.Execute("sgpadm_Sel_MinutaTablaGramajeFrecuencia '" & MyBuffer & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & CodigoServicio & ", " & Val(fpLongInteger1(2).Value) & ", " & Val(Format(fpDateTime1(0).text, "yyyymmdd")) & ", " & Val(Format(fpDateTime1(1).text, "yyyymmdd")) & ", '" & Val(fg_codigocbo(Combo1, 1, 1, "")) & "', '" & TipoMinuta & "', " & IIf(Option2(0).Value = True, "1", "2") & "")
           AuxPriTipoPlato = ""
           AuxSegTipoPlato = ""
           CantBrutaServida = ""
           NumLinExcel = 9
           Bar1(0).Visible = True
           Bar1(0).Value = 0
           ii = 1
           If Not RS.EOF Then
              Do While Not RS.EOF
                 Bar1(0).Value = Val((ii / RS.RecordCount) * 100)
                 If RS!PrimeraCategoria <> AuxPriTipoPlato Then
                    If AuxSegTipoPlato <> "" Then
                       If Trim(CantBrutaServida) <> "" Then
                          CantBrutaServida = Mid(CantBrutaServida, 1, Len(CantBrutaServida) - 3)
                       End If
                       MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, " " & CantBrutaServida
                       DibujarLineas oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel
                       MoverDatosExcel oExcel, oSheet, "F", "F", NumLinExcel, NumLinExcel, AuxFrecuencia
                       DibujarLineas oExcel, oSheet, "F", "F", NumLinExcel, NumLinExcel
                       
                       PonerTipoLetraTamaño oExcel, oSheet, "E", "F", NumLinExcel, NumLinExcel, 12
                       PonerCentrado oExcel, oSheet, "E", "F", NumLinExcel, NumLinExcel
                       CantBrutaServida = ""
                       NumLinExcel = NumLinExcel + 2
                    Else
                       MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, "ESTRUCTURA"
                       DibujarLineas oExcel, oSheet, "B", "D", NumLinExcel, NumLinExcel
                       
                       MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, IIf(Option2(0).Value = True, "Gramos Peso Bruto", "Servido")
                       DibujarLineas oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel
                            
                       MoverDatosExcel oExcel, oSheet, "F", "F", NumLinExcel, NumLinExcel, "Frecuencia " & " Día"
                       DibujarLineas oExcel, oSheet, "F", "F", NumLinExcel, NumLinExcel

                       PonerCombinarLeft oExcel, oSheet, "B", "D", NumLinExcel, NumLinExcel
                       PonerFontBold oExcel, oSheet, "B", "F", NumLinExcel, NumLinExcel
                       PonerTipoLetraTamaño oExcel, oSheet, "B", "F", NumLinExcel, NumLinExcel, 12

                       NumLinExcel = NumLinExcel + 2
                    End If
                    MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Trim(RS!PrimeraCategoria)
                    PonerCombinarLeft oExcel, oSheet, "B", "D", NumLinExcel, NumLinExcel
                    PonerFontBold oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel
                    PonerTipoLetraTamaño oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, 12
                    AuxPriTipoPlato = RS!PrimeraCategoria
                    AuxSegTipoPlato = ""
                    CantBrutaServida = ""
                    NumLinExcel = NumLinExcel + 1
                 End If
                 
                 If InStr(RS!SegundaCategoria, "\") = 0 Then
                    Sql = RS!SegundaCategoria
                 Else
                    Sql = Mid(RS!SegundaCategoria, 1, InStr(RS!SegundaCategoria, "\") - 1)
                 End If
                 
                 If RS!SegundaCategoria <> AuxSegTipoPlato Then
                    If AuxSegTipoPlato <> "" Then
                       If Trim(CantBrutaServida) <> "" Then
                          CantBrutaServida = Mid(CantBrutaServida, 1, Len(CantBrutaServida) - 3)
                       End If
                       MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, " " & CantBrutaServida
                       DibujarLineas oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel
                       MoverDatosExcel oExcel, oSheet, "F", "F", NumLinExcel, NumLinExcel, AuxFrecuencia
                       DibujarLineas oExcel, oSheet, "F", "F", NumLinExcel, NumLinExcel
                       CantBrutaServida = ""
                       If Sql <> AuxSegTipoPlato1 Then
                          NumLinExcel = NumLinExcel + IIf(AuxSegTipoPlato1 <> "", 2, 1)
                          AuxSegTipoPlato1 = Sql
                       Else
                          NumLinExcel = NumLinExcel + 1
                       End If

'                       NumLinExcel = NumLinExcel + 1
                    End If
                    MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, "° " & Trim(RS!SegundaCategoria)
                    PonerCombinarLeft oExcel, oSheet, "B", "D", NumLinExcel, NumLinExcel
                    PonerTipoLetraTamaño oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, 12
                    AuxSegTipoPlato = RS!SegundaCategoria
                 End If
                 CantBrutaServida = CantBrutaServida & IIf(RS!valor = 0, "", RS!valor & " - ")
                 AuxFrecuencia = IIf(IsNull(RS!Frecuencia) Or RS!Frecuencia = 0, "", RS!Frecuencia)
                 RS.MoveNext: ii = ii + 1
              Loop
              If Trim(CantBrutaServida) <> "" Then
                 CantBrutaServida = Mid(CantBrutaServida, 1, Len(CantBrutaServida) - 3)
              End If
              MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, " " & CantBrutaServida
              DibujarLineas oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel
              MoverDatosExcel oExcel, oSheet, "F", "F", NumLinExcel, NumLinExcel, AuxFrecuencia
              DibujarLineas oExcel, oSheet, "F", "F", NumLinExcel, NumLinExcel
              PonerTipoLetraTamaño oExcel, oSheet, "F", "E", 9, NumLinExcel, 12
              PonerCentrado oExcel, oSheet, "E", "F", 9, NumLinExcel
              CantBrutaServida = ""
              NumLinExcel = NumLinExcel + 1
              '-------> Dibujar Ancho Columna Frecuencia
              oSheet.Columns("E:E").Select
              oExcel.Selection.ColumnWidth = 20
              '-------> Ajustar Texto
              oSheet.Cells.Select
              With oExcel.Selection
                   .WrapText = True
                   .Orientation = 0
                   .AddIndent = False
                   .ShrinkToFit = False
                   .ReadingOrder = xlContext
              End With
              '-------> Sacar salto de pagina y linea divisora
              oExcel.ActiveWindow.DisplayGridLines = False
'              PonerBordeColor oExcel, oSheet, "A", "J", 1, NumLinExcel
              VistaPreliminarExcel oExcel, oSheet, False
           End If
           RS.Close: Set RS = Nothing
           Bar1(0).Visible = False
           Bar1(0).Value = 0
        End If
    Next i
    oExcel.Visible = True '------->Visualizar
    Set oSheet = Nothing
    Set oExcel = Nothing
    Set oBook = Nothing
    fg_descarga
    
    Exit Sub
Error:
    Bar1(0).Visible = False
    Bar1(0).Value = 0
    fg_descarga
    oExcel.DisplayAlerts = False
    oExcel.Quit
    oExcel.DisplayAlerts = True
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Exit Sub
    
End Sub

Sub ExportaExcelPlanDetalladoResumidoMKT(TipMin As String, opnomrec As Boolean, spid As Long, Indppr As String, opSemCerrada As Boolean, opparentesis As Boolean, EstadoPresentacion As String)

On Local Error GoTo Error

Dim RS As New ADODB.Recordset
Dim VecDie() As Long
Dim vecrec() As Double
Dim VecDia() As Double

Dim i As Long
Dim p As Long
Dim X As Long
Dim xx As Long
Dim j As Long
Dim ii As Long
Dim ind_par As Long
Dim ind_ini As Long
Dim NumLinExcel As Long
Dim NumLinExcelIni As Long
Dim ColumnaExcel As String

Dim NumAsc As Long
Dim CantNut As Long
Dim NumCol As Long
Dim CodigoServicio As Long
Dim NombreServicio As String
Dim NombreServicioAux As String
Dim IniIncDia  As Long
Dim auxrec As Long
Dim AuxFec As Long
Dim codapo As Long
Dim codpro As String
Dim TotalAportexDia As Double

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object

fg_carga ""
    
'-------> Start a new workbook in Excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
    
    NumCol = 5
    ReDim Preserve VecDie(0)
    ReDim Preserve vecrec(0)
    ReDim Preserve VecDia(0)
    If Option1(4).Value = True Then NumCol = 2
    If Option1(5).Value = True Then NumCol = 2
    If Option1(6).Value = True Then NumCol = 2
   
    '-------> Mover estado a la spread1
    ii = 1
    For i = 1 To TvwDir.Nodes.count
        
        If InStr(TvwDir.Nodes.item(i).key, "EstServicio") = 0 Then
           
           NombreServicio = CStr(Mid(TvwDir.Nodes.item(i).key, 2, Len(TvwDir.Nodes.item(i).key)))
           ii = vaSpread1.SearchCol(2, 0, vaSpread1.MaxRows, NombreServicio, SearchFlagsEqual)
           
           If ii > 0 Then
              
              vaSpread1.Row = ii
              vaSpread1.Row = ii: vaSpread1.Col = 1
              vaSpread1.text = "0"
              vaSpread1.text = IIf(TvwDir.Nodes.item(i).Checked = True, "1", "0")
           
           End If
        
        End If
    
    Next i
    
    NombreServicio = ""
    CantNut = 0
    
    For i = 1 To vaSpread2.MaxRows
        
        vaSpread2.Row = i
        vaSpread2.Col = 1
        
        If vaSpread2.text = "1" Then
           
           NumCol = NumCol + 1: CantNut = CantNut + 1
           ReDim Preserve VecDie(CantNut)
           vaSpread2.Col = 2
           VecDie(CantNut) = vaSpread2.text
        
        End If
    
    Next i
  
    ReDim Preserve vecrec(CantNut)
    ReDim Preserve VecDia(CantNut)
    
    Dim IndCol  As Long
    Dim IndColA As Long
    Dim oColA   As String
    
    Dim VecColExcel() As Variant
    ReDim VecColExcel(CantNut, 2)
          
    '-------> Setear vector
    For i = 1 To UBound(VecColExcel)
              
        VecColExcel(i, 1) = 0
        VecColExcel(i, 2) = "" 'descripción
          
    Next i
    
    IndCol = 1
    IndColA = 65
    oColA = ""
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        
        If vaSpread1.text = "1" Then
           
           vaSpread1.Col = 2
           CodigoServicio = Val(vaSpread1.text)
           
           vaSpread1.Col = 3
           NombreServicio = vaSpread1.text
           
           '-------> Crear Nueva Hoja Excel
           Set oSheet = oBook.Worksheets.Add
           NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
           
           If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
              
              NombreServicioAux = Mid(CodigoServicio & NombreServicioAux, 1, 31)
           
           End If
           oSheet.Name = NombreServicioAux
           
           NumAsc = IIf(EstadoPresentacion = "1", 66, 65)
           
           Select Case EstadoPresentacion
           
           Case "1" 'Aporte detallado
            
            If Option1(4).Value = True Then
            
            ElseIf Option1(5).Value = True Then
            
            ElseIf Option1(6).Value = True Then
            
            ElseIf Option1(12).Value = True Then
            
            ElseIf Option1(7).Value = True Then
                
                NumAsc = 69
                X = 5
            
            End If
           
           End Select
           '-------> Mover aportes nutricionales
           xx = 1
           IndCol = NumAsc + 1
           IndColA = 65
           oColA = ""
           
           For j = 1 To vaSpread2.MaxRows
               
               vaSpread2.Row = j
               vaSpread2.Col = 1
               
               If vaSpread2.text = "1" Then
                  
                  If Chr(IndCol) = "[" Then
                      
                      oColA = Chr(IndColA)
                      IndColA = IndColA + 1
                      IndCol = 65
                   
                  End If

                  ColumnaExcel = oColA & Chr(IndCol)
                  VecColExcel(xx, 1) = xx
                  VecColExcel(xx, 2) = oColA & Chr(IndCol)
                  
                  IndCol = IndCol + 1
                  xx = xx + 1
                  
               End If
           
           Next j
           
           '-------> Impresión titulo informe
           MoverDatosExcel oExcel, oSheet, "A", "A", 2, 2, "Aporte Nutricional " & IIf(EstadoPresentacion = "1", "Detallado ", "Resumido")
           
           '-------> Imprimir Sub_Segmento
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           Set RS = vg_db.Execute("SELECT sub_nombre FROM a_subsegmento WHERE sub_codigo = " & Val(fpLongInteger1(0).Value) & "")
           If RS.EOF Then fg_descarga: RS.Close: Set RS = Nothing: Exit Sub
           MoverDatosExcel oExcel, oSheet, "A", "A", 3, 3, Trim(RS!sub_nombre)
           RS.Close
           Set RS = Nothing
           '-------> Imprimir Servicio
           MoverDatosExcel oExcel, oSheet, "A", "A", 4, 4, "Servicio " & NombreServicio
           '-------> Formatear celda
           PonerFontBold oExcel, oSheet, "A", "A", 2, 4
           PonerCombinarCentrar oExcel, oSheet, "A", ColumnaExcel, 2, 2
           PonerCombinarCentrar oExcel, oSheet, "A", ColumnaExcel, 3, 3
           PonerCombinarCentrar oExcel, oSheet, "A", ColumnaExcel, 4, 4
           PonerTipoLetraTamaño oExcel, oSheet, "A", "A", 2, 4, 14
           
           '-------> Mover aportes nutriconales
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           Set RS = vg_db.Execute("sgpadm_s_aporteplanifminuta " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & CodigoServicio & ", " & Val(Format(fpDateTime1(0).text, "yyyymmdd")) & ", " & Val(Format(fpDateTime1(1).text, "yyyymmdd")) & ", " & Val(fpLongInteger1(2).Value) & ", '" & vg_NUsr & "','" & Indppr & "'")
           If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
           vaSpread3.MaxRows = 0
           vaSpread3.maxcols = 3
           Do While Not RS.EOF
              
              vaSpread3.MaxRows = vaSpread3.MaxRows + 1
              vaSpread3.Row = vaSpread3.MaxRows
              vaSpread3.Col = 1
              vaSpread3.text = RS!pnu_codpro
              
              vaSpread3.Col = 2
              vaSpread3.text = RS!pnu_codapo
              
              vaSpread3.Col = 3
              vaSpread3.text = RS!pnu_canapo
              
              RS.MoveNext
           
           Loop
           RS.Close
           Set RS = Nothing
           
           Let MyBuffer = ""
           Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
           Let MyBuffer = MyBuffer & "<EstServicio>"

           For xx = 1 To TvwDir.Nodes.count
              
              If TvwDir.Nodes.item(xx).Checked = True And InStr(TvwDir.Nodes.item(xx).key, "EstServicio") <> 0 And CodigoServicio = LCase(Trim(Mid(TvwDir.Nodes.item(xx).key, 2, 5))) Then
                 
                 Let MyBuffer = MyBuffer & " <EstServicioDet"
                 CodigoEstServicio = LCase(Trim(Mid(TvwDir.Nodes.item(xx).key, 18, 10)))
                 Let MyBuffer = MyBuffer & " CodigoEstServicio = " & Chr(34) & CodigoEstServicio & Chr(34)
                 Let MyBuffer = MyBuffer & "/>"
              
              End If
           
           Next xx
           
           Let MyBuffer = MyBuffer & "</EstServicio>"
           
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           Set RS = vg_db.Execute("sgpadm_Sel_MinutaAporteDetxEstServicio '" & MyBuffer & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & CodigoServicio & ", " & Val(fpLongInteger1(2).Value) & ", " & Val(Format(fpDateTime1(0).text, "yyyymmdd")) & ", " & Val(Format(fpDateTime1(1).text, "yyyymmdd")) & ", '" & TipMin & "'," & Indppr & "")
           If RS.EOF Then RS.Close: Set RS = Nothing: Exit For
           
           IniIncDia = 1
           If opSemCerrada = False Then
              
              MoverDatosExcel oExcel, oSheet, "A", "A", 8, 8, "Fecha " & Mid(RS!min_fecmin, 7, 2) & "/" & Mid(RS!min_fecmin, 5, 2) & "/" & Mid(RS!min_fecmin, 1, 4)
           
           Else
              
              MoverDatosExcel oExcel, oSheet, "A", "A", 8, 8, "Día " & IniIncDia
              IniIncDia = IniIncDia + 1
           
           End If
           PonerNegrilla oExcel, oSheet, "A", "A", 8, 8
           PonerTipoLetraTamaño oExcel, oSheet, "A", "A", 8, 8, 12
           
           '--- Se imprimen las columnas del informe --
           MoverDatosExcel oExcel, oSheet, "A", "A", 9, 9, "Preparaciones"
           DibujarLineas oExcel, oSheet, "A", "A", 9, 9
           NumAsc = IIf(EstadoPresentacion = "1", 66, 65)
           Select Case EstadoPresentacion
           
           Case "1" 'Aporte detallado
            
            If Option1(4).Value = True Then
                
                MoverDatosExcel oExcel, oSheet, "B", "B", 9, 9, "C.Bruta"
                DibujarLineas oExcel, oSheet, "B", "B", 9, 9
            
            ElseIf Option1(5).Value = True Then
                
                MoverDatosExcel oExcel, oSheet, "B", "B", 9, 9, "C.Servida"
                DibujarLineas oExcel, oSheet, "B", "B", 9, 9
            
            ElseIf Option1(6).Value = True Then
                
                MoverDatosExcel oExcel, oSheet, "B", "B", 9, 9, "C.Neta Nut."
                DibujarLineas oExcel, oSheet, "B", "B", 9, 9
            
            ElseIf Option1(12).Value = True Then
                
                MoverDatosExcel oExcel, oSheet, "B", "B", 9, 9, "C.Neta"
                DibujarLineas oExcel, oSheet, "B", "B", 9, 9
            
            ElseIf Option1(7).Value = True Then
            
                MoverDatosExcel oExcel, oSheet, "B", "B", 9, 9, "C.Bruta"
                DibujarLineas oExcel, oSheet, "B", "B", 9, 9
                MoverDatosExcel oExcel, oSheet, "C", "C", 9, 9, "C.Neta"
                DibujarLineas oExcel, oSheet, "C", "C", 9, 9
                MoverDatosExcel oExcel, oSheet, "D", "D", 9, 9, "C.Servida"
                DibujarLineas oExcel, oSheet, "D", "D", 9, 9
                MoverDatosExcel oExcel, oSheet, "E", "E", 9, 9, "C.Neta Nut."
                DibujarLineas oExcel, oSheet, "E", "E", 9, 9
                NumAsc = 69
                X = 5
            
            End If
           
           End Select
           '-------> Mover aportes nutricionales
           xx = 1
           For j = 1 To vaSpread2.MaxRows
               
               vaSpread2.Row = j
               vaSpread2.Col = 1
               
               If vaSpread2.text = "1" Then
                  
                  vaSpread2.Col = 3
                  
                  MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(xx, 2)), CStr(VecColExcel(xx, 2)), 9, 9, vaSpread2.text
                  DibujarLineas oExcel, oSheet, CStr(VecColExcel(xx, 2)), CStr(VecColExcel(xx, 2)), 9, 9
                  ColumnaExcel = VecColExcel(xx, 2)

                     xx = xx + 1
               
               End If
           
           Next j
           
           PonerColorInterior oExcel, oSheet, "A", CStr(VecColExcel(CantNut, 2)), 9, 9
           PonerColorFont oExcel, oSheet, "A", CStr(VecColExcel(CantNut, 2)), 9, 9
           PonerNegrilla oExcel, oSheet, "A", CStr(VecColExcel(CantNut, 2)), 9, 9
           PonerTipoLetraTamaño oExcel, oSheet, "A", CStr(VecColExcel(CantNut, 2)), 9, 9, 12
           PonerCentrado oExcel, oSheet, "B", CStr(VecColExcel(CantNut, 2)), 9, 9
           
           auxrec = 0
           AuxFec = 0
           j = 1
           ii = 1
           Bar1(0).Visible = True
           Bar1(0).Value = 0
           NumLinExcel = 10
           
           Do While Not RS.EOF
              
              DoEvents
              Bar1(0).Value = Val((ii / RS.RecordCount) * 100)
              
              If RS!min_fecmin <> AuxFec Then
                 
                 If AuxFec > 0 Then
                    
                    '-------> Salto de pagina x nuevo dìa
                    PonerTodosLosBorde oExcel, oSheet, "A", ColumnaExcel, NumLinExcelIni, IIf(EstadoPresentacion = "2", NumLinExcel - 1, NumLinExcel)
                    If EstadoPresentacion = "2" Then 'Aporte Resumido
                       
                       For j = 1 To CantNut
                           
                           MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), NumLinExcel, NumLinExcel, Format(vecrec(j), fg_Pict(6, 2))
                           
                           vecrec(j) = 0 '-------> Mover valor cero
                       
                       Next j
                       NumLinExcel = NumLinExcel + 1
                    
                    ElseIf EstadoPresentacion = "1" Then 'Aporte x día
                       
                       For j = 1 To CantNut
                           
                           MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, "Total Aporte"
                           PonerNegrilla oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel
                           MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), NumLinExcel, NumLinExcel, Format(vecrec(j), fg_Pict(6, 2))
                           vecrec(j) = 0 '-------> Mover valor cero
                       
                       Next j
                       
                       NumLinExcel = NumLinExcel + 1
                    
                    End If
                    
                    NumLinExcel = NumLinExcel + 1
                    If opSemCerrada = False Then
                       
                       MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "Fecha " & Mid(RS!min_fecmin, 7, 2) & "/" & Mid(RS!min_fecmin, 5, 2) & "/" & Mid(RS!min_fecmin, 1, 4)
                    
                    Else
                       
                       MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "Día " & IniIncDia
                       IniIncDia = IniIncDia + 1
                    
                    End If
                    
                    PonerNegrilla oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel
                    
                    NumLinExcel = NumLinExcel + 1
                    MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "Preparaciones"
                    DibujarLineas oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel
                    NumAsc = IIf(EstadoPresentacion = "1", 66, 65)
                    Select Case EstadoPresentacion
                    
                    Case "1" 'Aporte detallado
                        
                        If Option1(4).Value = True Then
                            
                            MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, "C.Bruta"
                            DibujarLineas oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel
                        
                        ElseIf Option1(5).Value = True Then
                            
                            MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, "C.Servida"
                            DibujarLineas oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel
                        
                        ElseIf Option1(6).Value = True Then
                            
                            MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, "C.Neta Nut."
                            DibujarLineas oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel
                        
                        ElseIf Option1(12).Value = True Then
                            
                            MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, "C.Neta"
                            DibujarLineas oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel
                        
                        ElseIf Option1(7).Value = True Then
                            
                            MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, "C.Bruta"
                            DibujarLineas oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel
                            MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, "C.Neta"
                            DibujarLineas oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel
                            MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, "C.Servida"
                            DibujarLineas oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel
                            MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, "C.Neta Nut."
                            DibujarLineas oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel
                            NumAsc = 69
                            X = 5
                        
                        End If
                    
                    End Select
                    '-------> Mover aportes nutricionales
                    
                    xx = 1
                    For j = 1 To vaSpread2.MaxRows
                        
                        vaSpread2.Row = j
                        vaSpread2.Col = 1
                        
                        If vaSpread2.text = "1" Then
                           
                           vaSpread2.Col = 3
                           MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(xx, 2)), CStr(VecColExcel(xx, 2)), NumLinExcel, NumLinExcel, vaSpread2.text
                           DibujarLineas oExcel, oSheet, CStr(VecColExcel(xx, 2)), CStr(VecColExcel(xx, 2)), NumLinExcel, NumLinExcel
                           
                           xx = xx + 1
                        
                        End If
                    
                    Next j
                    

                    PonerColorInterior oExcel, oSheet, "A", CStr(VecColExcel(IIf((xx - 1) < 1, 1, (xx - 1)), 2)), NumLinExcel, NumLinExcel
                    PonerColorFont oExcel, oSheet, "A", CStr(VecColExcel(IIf((xx - 1) < 1, 1, (xx - 1)), 2)), NumLinExcel, NumLinExcel
                    PonerNegrilla oExcel, oSheet, "A", CStr(VecColExcel(IIf((xx - 1) < 1, 1, (xx - 1)), 2)), NumLinExcel, NumLinExcel
                    
                    NumLinExcel = NumLinExcel + 1
                 
                 End If
                 AuxFec = RS!min_fecmin
                 auxrec = 0
              End If
              '-------> Corte control Recetas
              If RS!mid_codrec <> auxrec Then
                 
                 If auxrec > 0 Then
                    
                    PonerTodosLosBorde oExcel, oSheet, "A", ColumnaExcel, NumLinExcelIni, IIf(EstadoPresentacion = "2", NumLinExcel - 1, NumLinExcel)
                    
                    If EstadoPresentacion = "2" Then 'Aporte resumido
                       
                       For j = 1 To CantNut
                          
                          MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), NumLinExcel, NumLinExcel, Format(vecrec(j), fg_Pict(6, 2))
                          
                          vecrec(j) = 0 '-------> Mover valor cero
                       
                       Next j
                    
                    ElseIf EstadoPresentacion = "1" Then 'Aporte x día
                       
                       For j = 1 To CantNut
                           
                           MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, "Total Aporte"
                           PonerNegrilla oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel
                           MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), NumLinExcel, NumLinExcel, Format(vecrec(j), fg_Pict(6, 2))
                           
                           vecrec(j) = 0 '-------> Mover valor cero
                       
                       Next j
                       
                       NumLinExcel = NumLinExcel + 1
                    
                    End If
                    
                    NumLinExcel = NumLinExcel + 1
                 
                 End If
                 
                 '-------> Nombre recetas
                 If opnomrec = True Then
                    
                    MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, IIf(Not opparentesis, ExtraeParentesis(IIf(IsNull(RS!rec_nomfan) = True, " ", RS!rec_nomfan)), IIf(IsNull(RS!rec_nomfan) = True, " ", RS!rec_nomfan))
                 
                 Else
                    
                    MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, IIf(Not opparentesis, ExtraeParentesis(IIf(IsNull(RS!rec_nombre) = True, " ", RS!rec_nombre)), IIf(IsNull(RS!rec_nombre) = True, " ", RS!rec_nombre))
                 
                 End If
                 
                 PonerNegrilla oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel
                 auxrec = RS!mid_codrec
                 NumLinExcelIni = NumLinExcel
                 If EstadoPresentacion = "1" Then 'Aporte detallado
                    
                    NumLinExcel = NumLinExcel + 1
                 
                 End If
              
              End If
              '-------> Nombre ingredientes
              If EstadoPresentacion = "1" Then 'Aporte detallado
                 
                 MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, IIf(IsNull(RS!ing_nombre), "No Existe Descripción", Trim(RS!ing_nombre))
                 
                 If Option1(4).Value = True Then
                    
                    MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(RS!canpro, fg_Pict(6, vg_RDCa))
                 
                 ElseIf Option1(5).Value = True Then
                    
                    MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(CCur((IIf(RS!red_codpro <> RS!ori_codpro, RS!ing_pctnut, RS!red_pctnut) / 100) * RS!canpro), fg_Pict(6, vg_RDCa))
                 
                 ElseIf Option1(6).Value = True Then
                    
                    MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(CCur((RS!canpro * (IIf(RS!red_codpro <> RS!ori_codpro, RS!ing_pctcoc, RS!red_pctcoc) / 100)) * (IIf(RS!red_codpro <> RS!ori_codpro, RS!ing_pctapr, RS!red_pctapr) / 100)), fg_Pict(6, vg_RDCa))
                 
                 ElseIf Option1(12).Value = True Then
                    
                    MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(CCur((IIf(RS!red_codpro <> RS!ori_codpro, RS!ing_pctapr, RS!red_pctapr) / 100) * RS!canpro), fg_Pict(6, vg_RDCa))
                 
                 ElseIf Option1(7).Value = True Then
                    
                    MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, Format(RS!canpro, fg_Pict(6, vg_RDCa))
                    MoverDatosExcel oExcel, oSheet, "C", "C", NumLinExcel, NumLinExcel, Format(CCur((IIf(RS!red_codpro <> RS!ori_codpro, RS!ing_pctapr, RS!red_pctapr) / 100) * RS!canpro), fg_Pict(6, vg_RDCa))
                    MoverDatosExcel oExcel, oSheet, "D", "D", NumLinExcel, NumLinExcel, Format(CCur((RS!canpro * (IIf(RS!red_codpro <> RS!ori_codpro, RS!ing_pctcoc, RS!red_pctcoc) / 100)) * (IIf(RS!red_codpro <> RS!ori_codpro, RS!ing_pctapr, RS!red_pctapr) / 100)), fg_Pict(6, vg_RDCa))
                    MoverDatosExcel oExcel, oSheet, "E", "E", NumLinExcel, NumLinExcel, Format(CCur((IIf(RS!red_codpro <> RS!ori_codpro, RS!ing_pctnut, RS!red_pctnut) / 100) * RS!canpro), fg_Pict(6, vg_RDCa))
                    Y = 5
                 
                 End If
                 
                 xx = 1
                 
                 For j = 1 To CantNut
                     
                     MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), IIf(EstadoPresentacion = "1", NumLinExcel, NumLinExcel - 1), IIf(EstadoPresentacion = "1", NumLinExcel, NumLinExcel - 1), 0
                     
                     xx = xx + 1
                 
                 Next j
              
              End If
              
              Trim (CStr(RS!red_codpro))
              ind_ini = vaSpread3.SearchCol(1, -1, vaSpread3.MaxRows, RS!red_codpro, SearchFlagsEqual)
              codpro = ""
              For ind_par = ind_ini To vaSpread3.MaxRows
                  
                  vaSpread3.Row = ind_par
                  vaSpread3.Col = 1
                  
                  If vaSpread3.text <> RS!red_codpro Then Exit For
                  
                  vaSpread3.Col = 2
                  codapo = vaSpread3.text
                  
                  vaSpread3.Col = 3
                  canapo = vaSpread3.text
                  
                  For j = 1 To CantNut
                      
                      If VecDie(j) = codapo Then
                         
                         TotalAportexDia = 0
                         
                         If EstadoPresentacion = "1" Then 'Aporte Detallado
                            
                            TotalAportexDia = Format(((((RS!red_pctnut / 100) * (canapo * (RS!canpro))) / RS!ing_facnut)), IIf(codapo = 2, fg_Pict(6, 0), fg_Pict(6, 2)))
                            MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), NumLinExcel, NumLinExcel, CStr(TotalAportexDia)
                         
                         End If
                         
                         If codapo = 2 Then
                            
                            vecrec(j) = Format(CCur(vecrec(j) + ((((RS!red_pctnut / 100) * (canapo * (RS!canpro))) / RS!ing_facnut))), fg_Pict(2, 0))
                         
                         Else
                            
                            vecrec(j) = CCur(vecrec(j) + ((((RS!red_pctnut / 100) * (canapo * (RS!canpro))) / RS!ing_facnut)))
                         
                         End If
                         
                         Exit For
                      
                      End If
                  
                  Next j
              
              Next ind_par
              
              If EstadoPresentacion = "1" Then 'Aporte detallado
                 
                 NumLinExcel = NumLinExcel + 1
              
              End If
              
              RS.MoveNext
              
              ii = ii + 1
           
           Loop
           
           RS.Close: Set RS = Nothing
           PonerTodosLosBorde oExcel, oSheet, "A", ColumnaExcel, NumLinExcelIni, IIf(EstadoPresentacion = "2", NumLinExcel - 1, NumLinExcel)
           If EstadoPresentacion = "2" Then '-------> Aportes resumido
              
              For j = 1 To CantNut
                  
                  MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), NumLinExcel, NumLinExcel, Format(vecrec(j), fg_Pict(6, 2))
                  vecrec(j) = 0 '-------> Mover valor cero
              
              Next j
              PonerTipoLetraTamaño oExcel, oSheet, "A", ColumnaExcel, NumLinExcel, NumLinExcel, 12
              PonerCentrado oExcel, oSheet, "B", ColumnaExcel, NumLinExcel, NumLinExcel
           
           ElseIf EstadoPresentacion = "1" Then 'Aporte x día
              
              For j = 1 To CantNut
                  
                  MoverDatosExcel oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel, "Total Aporte"
                  PonerNegrilla oExcel, oSheet, "B", "B", NumLinExcel, NumLinExcel
                  MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), NumLinExcel, NumLinExcel, Format(vecrec(j), fg_Pict(6, 2))
                  vecrec(j) = 0 '-------> Mover valor cero
              
              Next j
              
              NumLinExcel = NumLinExcel + 1
           
           End If
           PonerCentrado oExcel, oSheet, "B", ColumnaExcel, 8, NumLinExcel
           PonerTipoLetraTamaño oExcel, oSheet, "A", ColumnaExcel, 8, NumLinExcel, 12
           
           PonerAnchoColumna oExcel, oSheet, "A", "A", 1, 1, 60
           PonerAnchoColumna oExcel, oSheet, "B", ColumnaExcel, 1, 1, 20
           
           '-------> Ajustar Texto
           oSheet.Cells.Select
           With oExcel.Selection
                
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .ReadingOrder = xlContext
           
           End With
           '--------> Determinar ancho de columna
           
           '-------> Sacar salto de pagina y linea divisora
           oExcel.ActiveWindow.DisplayGridLines = False
           VistaPreliminarExcel oExcel, oSheet, False
        
        End If
        cLin = ""
        Bar1(0).Value = 0
        Bar1(0).Visible = False
    
    Next i
    
    oExcel.Visible = True '------->Visualizar
    Set oSheet = Nothing
    Set oExcel = Nothing
    Set oBook = Nothing
    Bar1(0).Value = 0
    Bar1(0).Visible = False
    fg_descarga

Exit Sub
Error:
    Bar1(0).Value = 0
    Bar1(0).Visible = False
    fg_descarga
    oExcel.DisplayAlerts = False
    oExcel.Quit
    oExcel.DisplayAlerts = True
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Exit Sub

End Sub

Sub ExportarExcelAportePlanEstrResMKT(TipMin As String, opnomrec As Boolean, spid As Long, Indppr As String, opnomest As Boolean, opSemCerrada As Boolean)

On Local Error GoTo Error

Dim RS As New ADODB.Recordset
Dim VecDie() As Long
Dim vecrec() As Double
Dim VecDia() As Double

Dim g As Long
Dim i As Long
Dim ii As Long
Dim j As Long
Dim X As Long
Dim xx As Long
Dim NumCol As Long
Dim CantNut As Long
Dim IniIncDia As Long
Dim contador As Long
Dim NumAsc As String
Dim ind_par As Long
Dim ind_ini As Long
Dim NumLinExcel As Long
Dim NumLinExcelIni As Long
Dim ColumnaExcel As String

Dim AuxFec As Long
Dim auxrec As Long
Dim AuxEstr As Long

Dim NombreServicio As String
Dim NombreServicioAux As String
Dim CodigoServicio As Long
Dim codapo As Long
Dim codpro As String

Dim canapo As Double

Dim NombreAnt As String

contador = 0
Dim Col, UltCol, TotalCol As Integer
Dim Valc As Double

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object

fg_carga ""
    
'-------> Mover estado a la spread1
ii = 1
For i = 1 To TvwDir.Nodes.count
    If InStr(TvwDir.Nodes.item(i).key, "EstServicio") = 0 Then
           NombreServicio = CStr(Mid(TvwDir.Nodes.item(i).key, 2, Len(TvwDir.Nodes.item(i).key)))
           ii = vaSpread1.SearchCol(2, 0, vaSpread1.MaxRows, NombreServicio, SearchFlagsEqual)
           If ii > 0 Then
              vaSpread1.Row = ii
              vaSpread1.Row = ii: vaSpread1.Col = 1
              vaSpread1.text = "0"
              vaSpread1.text = IIf(TvwDir.Nodes.item(i).Checked = True, "1", "0")
           End If
    End If
Next i
    
NombreServicio = ""
'-------> Start a new workbook in Excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
    
    NumCol = 4
    ReDim Preserve VecDie(0)
    ReDim Preserve vecrec(0)
    ReDim Preserve VecDia(0)
    If Option1(4).Value = True Then NumCol = 2
    If Option1(5).Value = True Then NumCol = 2
    If Option1(6).Value = True Then NumCol = 2
    CantNut = 0
    For i = 1 To vaSpread2.MaxRows
        
        vaSpread2.Row = i
        vaSpread2.Col = 1
        
        If vaSpread2.text = "1" Then
           
           NumCol = NumCol + 1: CantNut = CantNut + 1
           ReDim Preserve VecDie(CantNut)
           vaSpread2.Col = 2
           VecDie(CantNut) = vaSpread2.text
        
        End If
    
    Next i
    ReDim Preserve vecrec(CantNut)
    ReDim Preserve VecDia(CantNut)
    Dim nmayvec As Long
    nmayvec = 0
    contador = 0
    
    Dim IndCol  As Long
    Dim IndColA As Long
    Dim oColA   As String
    
    Dim VecColExcel() As Variant
    ReDim VecColExcel(CantNut, 2)
          
    '-------> Setear vector
    For i = 1 To UBound(VecColExcel)
              
        VecColExcel(i, 1) = 0
        VecColExcel(i, 2) = "" 'descripción
          
    Next i
    
    IndCol = 1
    IndColA = 65
    oColA = ""
    oColA = Trim(Chr(IndColA))
'    IndCol = 65
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        
        If vaSpread1.text = "1" Then
           
           vaSpread1.Col = 2
           CodigoServicio = vaSpread1.text
           vaSpread1.Col = 3
           NombreServicio = vaSpread1.text

           '-------> Crear Nueva Hoja Excel
           Set oSheet = oBook.Worksheets.Add
           NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
           If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
              
              NombreServicioAux = Mid(CodigoServicio & NombreServicioAux, 1, 31)
           
           End If
           oSheet.Name = NombreServicioAux 'Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31) 'NombreServicio
           
           xx = 1
           NumAsc = 65
           
           For j = 1 To vaSpread2.MaxRows
               
               vaSpread2.Row = j
               vaSpread2.Col = 1
               
               If vaSpread2.text = "1" Then
                  
                  If Chr(IndCol) = "[" Then
                      
                      oColA = Chr(IndColA)
                      IndColA = IndColA + 1
                      IndCol = 65
                   
                  End If

                  ColumnaExcel = Chr(xx + NumAsc) '& Chr(IndCol)
                  VecColExcel(xx, 1) = xx
                  VecColExcel(xx, 2) = Chr(xx + NumAsc) '& Chr(IndCol)
                  
                  IndCol = IndCol + 1
                  xx = xx + 1
                  
               End If
           
           Next j
           
           '-------> Impresión titulo informe
           MoverDatosExcel oExcel, oSheet, "A", "A", 2, 2, "Aporte Nutricional "
           
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           Set RS = vg_db.Execute("sgpadm_s_subsegmento 1, " & Val(fpLongInteger1(0).Value) & ", '', ''")
           If Not RS.EOF Then MoverDatosExcel oExcel, oSheet, "A", "A", 3, 3, Trim(RS!sub_nombre)
           RS.Close: Set RS = Nothing

           MoverDatosExcel oExcel, oSheet, "A", "A", 4, 4, "Servicio " & NombreServicio
           
           '-------> Formatear celda
           PonerFontBold oExcel, oSheet, "A", "A", 2, 4
           PonerCombinarCentrar oExcel, oSheet, "A", ColumnaExcel, 2, 2
           PonerCombinarCentrar oExcel, oSheet, "A", ColumnaExcel, 3, 3
           PonerCombinarCentrar oExcel, oSheet, "A", ColumnaExcel, 4, 4
           PonerTipoLetraTamaño oExcel, oSheet, "A", "A", 2, 4, 14
           
           '-------> Definir largo del vector
           Dim vecResumen() As Variant
           ReDim Preserve vecResumen(1000, CantNut + 3)
           
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           Set RS = vg_db.Execute("SELECT COUNT(c.ess_codigo) AS nreg FROM b_minuta a, b_minutadet b, a_estservicio c WHERE a.min_codigo = b.mid_codigo AND b.mid_estser = c.ess_codigo AND a.min_subseg = " & Val(fpLongInteger1(0).Value) & " AND a.min_codreg = " & Val(fpLongInteger1(1).Value) & " AND a.min_codser = " & CodigoServicio & " AND a.min_fecmin >= " & Val(Format(fpDateTime1(0).text, "yyyymmdd")) & " AND a.min_fecmin <= " & Val(Format(fpDateTime1(1).text, "yyyymmdd")) & "")
           If Not RS.EOF Then

'              If RS!nreg > nmayvec Then ReDim Preserve vecResumen(RS!nreg, inut + 3): nmayvec = RS!nreg
           
           End If
           RS.Close: Set RS = Nothing
           
           '-------> Setear vector
           For g = 1 To UBound(vecResumen)
               
               vecResumen(g, 1) = Val(0) 'codigo
               vecResumen(g, 2) = "" 'descripción
               vecResumen(g, 3) = Val(0)
               
               For j = 4 To CantNut + 3 'UBound(vecResumen) 'contador día
                   
                   vecResumen(g, j) = Val(0)
               
               Next j
           
           Next g
           '-------> Se ingresa código estructura
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           Set RS = vg_db.Execute("SELECT DISTINCT c.ess_codigo, c.ess_nombre, b.mid_numlin, b.mid_desest FROM b_minuta a, b_minutadet b, a_estservicio c WHERE a.min_codigo = b.mid_codigo AND b.mid_estser = c.ess_codigo AND a.min_subseg = " & Val(fpLongInteger1(0).Value) & " AND a.min_codreg = " & Val(fpLongInteger1(1).Value) & " AND a.min_codser = " & CodigoServicio & " AND a.min_fecmin >= " & Val(Format(fpDateTime1(0).text, "yyyymmdd")) & " AND a.min_fecmin <= " & Val(Format(fpDateTime1(1).text, "yyyymmdd")) & " ORDER BY b.mid_numlin")
           Do While Not RS.EOF
              
              For g = 1 To UBound(vecResumen)
                  
                  If Val(vecResumen(g, 1)) = RS!ess_codigo Then
                     
                     Exit For
                  
                  ElseIf vecResumen(g, 1) = 0 Then
                     
                     vecResumen(g, 1) = Trim(RS!ess_codigo)
                     vecResumen(g, 2) = IIf(opnomest = True, IIf(IIf(IsNull(RS!mid_desest) = True, "", RS!mid_desest) <> "", RS!mid_desest, RS!ess_nombre), RS!ess_nombre)
                     Exit For
                  
                  End If
              
              Next g
              
              RS.MoveNext
           
           Loop
           RS.Close: Set RS = Nothing
           '-------> Cargar numero estructura en el peridio
           
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           Set RS = vg_db.Execute("sgpadm_s_BuscarNumeroEstructura " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & CodigoServicio & ", " & Val(Format(fpDateTime1(0).text, "yyyymmdd")) & ", " & Val(Format(fpDateTime1(1).text, "yyyymmdd")) & "")
           Do While Not RS.EOF
              
              For g = 1 To UBound(vecResumen)
                  
                  If Val(vecResumen(g, 1)) = RS!mid_estser Then
                     
                     vecResumen(g, 3) = Trim(RS!nReg)
                     Exit For
                  
                  End If
              
              Next g
              
              RS.MoveNext
           
           Loop
           RS.Close: Set RS = Nothing
           
           '-------> Mover aportes nutricionales
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           Set RS = vg_db.Execute("sgpadm_s_aporteplanifminuta " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & CodigoServicio & ", " & Val(Format(fpDateTime1(0).text, "yyyymmdd")) & ", " & Val(Format(fpDateTime1(1).text, "yyyymmdd")) & ", " & Val(fpLongInteger1(2).Value) & ", '" & vg_NUsr & "'," & Indppr & "")
           If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
           vaSpread3.MaxRows = 0
           vaSpread3.maxcols = 3
           Do While Not RS.EOF
              
              vaSpread3.MaxRows = vaSpread3.MaxRows + 1
              vaSpread3.Row = vaSpread3.MaxRows
              vaSpread3.Col = 1: vaSpread3.text = RS!pnu_codpro
              vaSpread3.Col = 2: vaSpread3.text = RS!pnu_codapo
              vaSpread3.Col = 3: vaSpread3.text = RS!pnu_canapo
              RS.MoveNext
           
           Loop
           RS.Close: Set RS = Nothing
           
           Let MyBuffer = ""
           Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
           Let MyBuffer = MyBuffer & "<EstServicio>"

           For xx = 1 To TvwDir.Nodes.count
              
              If TvwDir.Nodes.item(xx).Checked = True And InStr(TvwDir.Nodes.item(xx).key, "EstServicio") <> 0 And CodigoServicio = LCase(Trim(Mid(TvwDir.Nodes.item(xx).key, 2, 5))) Then
                 
                 Let MyBuffer = MyBuffer & " <EstServicioDet"
                 CodigoEstServicio = LCase(Trim(Mid(TvwDir.Nodes.item(xx).key, 18, 10)))
                 Let MyBuffer = MyBuffer & " CodigoEstServicio = " & Chr(34) & CodigoEstServicio & Chr(34)
                 Let MyBuffer = MyBuffer & "/>"
              
              End If
           
           Next xx
           
           Let MyBuffer = MyBuffer & "</EstServicio>"
           
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           Set RS = vg_db.Execute("sgpadm_Sel_MinutaAporteEstxEstServicio '" & MyBuffer & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & CodigoServicio & ", " & Val(fpLongInteger1(2).Value) & ", " & Val(Format(fpDateTime1(0).text, "yyyymmdd")) & ", " & Val(Format(fpDateTime1(1).text, "yyyymmdd")) & ", '" & TipMin & "'," & Indppr & "")
           IniIncDia = 1
           If RS.EOF Then RS.Close: Set RS = Nothing: Exit For
           
           If opSemCerrada = False Then
              
              MoverDatosExcel oExcel, oSheet, "A", "A", 8, 8, "Fecha " & Mid(RS!min_fecmin, 7, 2) & "/" & Mid(RS!min_fecmin, 5, 2) & "/" & Mid(RS!min_fecmin, 1, 4)
           
           Else
              
              MoverDatosExcel oExcel, oSheet, "A", "A", 8, 8, "Día " & IniIncDia
              IniIncDia = IniIncDia + 1
           
           End If
           
           PonerNegrilla oExcel, oSheet, "A", "A", 8, 8
           PonerTipoLetraTamaño oExcel, oSheet, "A", "A", 8, 8, 12
           
           '--- Se imprimen las columnas del informe --
           MoverDatosExcel oExcel, oSheet, "A", "A", 9, 9, "Estructura"
           DibujarLineas oExcel, oSheet, "A", "A", 9, 9
           
           xx = 1
           NumAsc = 65
           For j = 1 To vaSpread2.MaxRows
               
               vaSpread2.Row = j
               vaSpread2.Col = 1
               
               If vaSpread2.text = "1" Then
                  
                  vaSpread2.Col = 3
                  MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(xx, 2)), CStr(VecColExcel(xx, 2)), 9, 9, vaSpread2.text
                  DibujarLineas oExcel, oSheet, CStr(VecColExcel(xx, 2)), CStr(VecColExcel(xx, 2)), 9, 9
                  ColumnaExcel = CStr(VecColExcel(xx, 2))
                  xx = xx + 1
               
               End If
           
           Next j
           PonerColorInterior oExcel, oSheet, "A", CStr(VecColExcel(IIf((xx - 1) < 1, 1, (xx - 1)), 2)), 9, 9
           PonerColorFont oExcel, oSheet, "A", CStr(VecColExcel(IIf((xx - 1) < 1, 1, (xx - 1)), 2)), 9, 9
           PonerNegrilla oExcel, oSheet, "A", CStr(VecColExcel(IIf((xx - 1) < 1, 1, (xx - 1)), 2)), 9, 9
           PonerTipoLetraTamaño oExcel, oSheet, "A", CStr(VecColExcel(IIf((xx - 1) < 1, 1, (xx - 1)), 2)), 9, 9, 12
           PonerCentrado oExcel, oSheet, "B", CStr(VecColExcel(IIf((xx - 1) < 1, 1, (xx - 1)), 2)), 9, 9
           
           NumLinExcel = 10
           auxrec = 0
           AuxFec = 0
           j = 1
           ii = 1
           contador = 1
           Bar1(0).Visible = True
           Bar1(0).Value = 0
           
           Do While Not RS.EOF
              
              DoEvents
              Bar1(0).Value = Val((ii / RS.RecordCount) * 100)
              
              If RS!min_fecmin <> AuxFec Then
                 
                 If AuxFec > 0 Then
                    
                    PonerTodosLosBorde oExcel, oSheet, "A", ColumnaExcel, NumLinExcelIni, NumLinExcel - 1
                    
                    For j = 1 To CantNut
                        
                        Valc = 0: Col = 0: UltCol = 0
                        
                        If contador > 0 Then
                           
                           If VecDie(j) = 2 Then
                              
                              MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), NumLinExcel, NumLinExcel, Format((vecrec(j) / contador), fg_Pict(6, 0))
                           
                           Else
                              
                              MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), NumLinExcel, NumLinExcel, Format((vecrec(j) / contador), fg_Pict(6, 2))
                           
                           End If
                        
                        Else
                           
                           MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), NumLinExcel, NumLinExcel, Format(0, fg_Pict(6, 2))
                        
                        End If
                    
                    Next j
                    
                    For j = 1 To CantNut
                        
                        vecrec(j) = 0
                    
                    Next j
                    
                    NumLinExcel = NumLinExcel + 2
                    
                    '--- Se imprimen las columnas del informe --
                    If opSemCerrada = False Then
                       
                       MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "Fecha " & Mid(RS!min_fecmin, 7, 2) & "/" & Mid(RS!min_fecmin, 5, 2) & "/" & Mid(RS!min_fecmin, 1, 4)
                    
                    Else
                       
                       MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "Día " & IniIncDia
                       IniIncDia = IniIncDia + 1
                    
                    End If
                    PonerNegrilla oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel
                    PonerTipoLetraTamaño oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, 12
                    
                    NumLinExcel = NumLinExcel + 1
                    MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, "Estructura"
                    DibujarLineas oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel
           
                    xx = 1
                    NumAsc = 65
                    For j = 1 To vaSpread2.MaxRows
                        
                        vaSpread2.Row = j
                        vaSpread2.Col = 1
                        
                        If vaSpread2.text = "1" Then
                            
                            vaSpread2.Col = 3
                            MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(xx, 2)), CStr(VecColExcel(xx, 2)), NumLinExcel, NumLinExcel, vaSpread2.text
                            DibujarLineas oExcel, oSheet, CStr(VecColExcel(xx, 2)), CStr(VecColExcel(xx, 2)), NumLinExcel, NumLinExcel
                            ColumnaExcel = CStr(VecColExcel(xx, 2))
                            xx = xx + 1
                        
                        End If
                    
                    Next j
                    
                    PonerColorInterior oExcel, oSheet, "A", CStr(VecColExcel(IIf((xx - 1) < 1, 1, (xx - 1)), 2)), NumLinExcel, NumLinExcel
                    PonerColorFont oExcel, oSheet, "A", CStr(VecColExcel(IIf((xx - 1) < 1, 1, (xx - 1)), 2)), NumLinExcel, NumLinExcel
                    PonerNegrilla oExcel, oSheet, "A", CStr(VecColExcel(IIf((xx - 1) < 1, 1, (xx - 1)), 2)), NumLinExcel, NumLinExcel
                    PonerTipoLetraTamaño oExcel, oSheet, "A", CStr(VecColExcel(IIf((xx - 1) < 1, 1, (xx - 1)), 2)), NumLinExcel, NumLinExcel, 12
                    PonerCentrado oExcel, oSheet, "B", CStr(VecColExcel(IIf((xx - 1) < 1, 1, (xx - 1)), 2)), NumLinExcel, NumLinExcel
                    NumLinExcel = NumLinExcel + 1
                 
                 End If
                 AuxFec = RS!min_fecmin
                 auxrec = 0
                 AuxEstr = 0
                 contador = 0
              End If
              
              If RS!mid_estser <> AuxEstr Then
                 
                 Dim valor As Double
                 
                 If AuxEstr > 0 Then
                    
                    PonerTodosLosBorde oExcel, oSheet, "A", ColumnaExcel, NumLinExcelIni, NumLinExcel - 1
                    
                    For j = 1 To CantNut
                        
                        Valc = 0: Col = 0: UltCol = 0
                        
                        If contador > 0 Then
                           
                           If VecDie(j) = 2 Then
                              
                              MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), NumLinExcel, NumLinExcel, Format((vecrec(j) / contador), fg_Pict(6, 0))
                           
                           Else
                              
                              MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), NumLinExcel, NumLinExcel, Format((vecrec(j) / contador), fg_Pict(6, 2))
                           
                           End If
                        
                        Else
                           
                           MoverDatosExcel oExcel, oSheet, CStr(VecColExcel(j, 2)), CStr(VecColExcel(j, 2)), NumLinExcel, NumLinExcel, Format(0, fg_Pict(6, 2))
                        
                        End If
                    
                    Next j
                    
                    For j = 1 To CantNut
                        
                        vecrec(j) = 0
                    
                    Next j
                    
                    NumLinExcel = NumLinExcel + 1
                    
                 End If
                 
                 NumLinExcelIni = NumLinExcel
                 NombreAnt = IIf(opnomest = True, IIf(IIf(IsNull(RS!mid_desest) = True, "", RS!mid_desest) <> "", RS!mid_desest, RS!ess_nombre), RS!ess_nombre)
                 MoverDatosExcel oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel, IIf(opnomest = True, IIf(IIf(IsNull(RS!mid_desest) = True, "", RS!mid_desest) <> "", RS!mid_desest, RS!ess_nombre), RS!ess_nombre)
                 PonerNegrilla oExcel, oSheet, "A", "A", NumLinExcel, NumLinExcel
                 
                 AuxEstr = RS!mid_estser
                 contador = 0
              
              End If
              
              '-------> Corte control Estructura
              If RS!mid_codrec <> auxrec Or RS!mid_estser <> AuxEstr Then
                
                contador = contador + 1
                auxrec = RS!mid_codrec
              
              End If

              Trim (CStr(RS!red_codpro))
              ind_ini = vaSpread3.SearchCol(1, -1, vaSpread3.MaxRows, RS!red_codpro, SearchFlagsEqual)
              codpro = ""
              For ind_par = ind_ini To vaSpread3.MaxRows
                  
                  vaSpread3.Row = ind_par
                  vaSpread3.Col = 1
                  
                  If vaSpread3.text <> RS!red_codpro Then Exit For
                  
                  vaSpread3.Col = 2
                  codapo = vaSpread3.text
                  vaSpread3.Col = 3
                  canapo = vaSpread3.text
                  
                  For j = 1 To CantNut
                      
                      If VecDie(j) = codapo Then
                         
                         If codapo = 2 Then
                            
                            vecrec(j) = Format(CCur(vecrec(j) + ((((RS!red_pctnut / 100) * (canapo * (RS!canpro))) / RS!ing_facnut))), fg_Pict(6, 0))
                         
                         Else
                            
                            vecrec(j) = CCur(vecrec(j) + ((((RS!red_pctnut / 100) * (canapo * (RS!canpro))) / RS!ing_facnut)))
                         End If
                         
                         Exit For
                      
                      End If
                  
                  Next j
              
              Next ind_par
              
              
              RS.MoveNext
              X = X + 1
              ii = ii + 1
              
           Loop
           RS.Close: Set RS = Nothing
           PonerTodosLosBorde oExcel, oSheet, "A", ColumnaExcel, NumLinExcelIni, NumLinExcel - 1
           
           For j = 1 To CantNut
               
               Valc = 0: Col = 0: UltCol = 0
               
               If contador > 0 Then
                  
                  If VecDie(j) = 2 Then
                     
                     MoverDatosExcel oExcel, oSheet, Chr(j + NumAsc), Chr(j + NumAsc), NumLinExcel, NumLinExcel, Format((vecrec(j) / contador), fg_Pict(6, 0))
                  
                  Else
                     
                     MoverDatosExcel oExcel, oSheet, Chr(j + NumAsc), Chr(j + NumAsc), NumLinExcel, NumLinExcel, Format((vecrec(j) / contador), fg_Pict(6, 2))
                  
                  End If
               
               Else
                  
                  MoverDatosExcel oExcel, oSheet, Chr(j + NumAsc), Chr(j + NumAsc), NumLinExcel, NumLinExcel, Format(0, fg_Pict(6, 2))
               
               End If
           
           Next j
           PonerCentrado oExcel, oSheet, "B", ColumnaExcel, 8, NumLinExcel
           PonerTipoLetraTamaño oExcel, oSheet, "A", ColumnaExcel, 8, NumLinExcel, 12
                    
           For j = 1 To CantNut
               
               vecrec(j) = 0
           
           Next j
           
           PonerAnchoColumna oExcel, oSheet, "A", "A", 1, 1, 60
           PonerAnchoColumna oExcel, oSheet, "B", ColumnaExcel, 1, 1, 20
           
           '-------> Ajustar Texto
           oSheet.Cells.Select
           With oExcel.Selection
                
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .ReadingOrder = xlContext
           
           End With
           '--------> Determinar ancho de columna
           
           '-------> Sacar salto de pagina y linea divisora
           oExcel.ActiveWindow.DisplayGridLines = False
           VistaPreliminarExcel oExcel, oSheet, False
        
        End If
        contador = 0
        Bar1(0).Value = 0
    
    Next i
    oExcel.Visible = True '------->Visualizar
    Set oSheet = Nothing
    Set oExcel = Nothing
    Set oBook = Nothing
    Bar1(0).Value = 0
    Bar1(0).Visible = False
    fg_descarga
    Exit Sub

Error:
    Bar1(0).Value = 0
    Bar1(0).Visible = False
    fg_descarga
    oExcel.DisplayAlerts = False
    oExcel.Quit
    oExcel.DisplayAlerts = True
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Exit Sub
End Sub
