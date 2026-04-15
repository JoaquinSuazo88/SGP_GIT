VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form I_Planif 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Planificación"
   ClientHeight    =   7065
   ClientLeft      =   4080
   ClientTop       =   2340
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6585
      Index           =   0
      Left            =   60
      TabIndex        =   17
      Top             =   360
      Width           =   7395
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
            Width           =   1215
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
         Left            =   4200
         TabIndex        =   43
         Top             =   4920
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         ItemData        =   "I_Planif.frx":0000
         Left            =   1575
         List            =   "I_Planif.frx":0002
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
         Left            =   4080
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
            Picture         =   "I_Planif.frx":0004
            Top             =   160
            Width           =   480
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   6975
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "I_Planif.frx":030E
            Left            =   1680
            List            =   "I_Planif.frx":0310
            Style           =   2  'Dropdown List
            TabIndex        =   0
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
         Left            =   4200
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
            Picture         =   "I_Planif.frx":0312
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
         Width           =   6975
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
            Left            =   5520
            TabIndex        =   14
            Top             =   360
            Width           =   1260
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
            TabIndex        =   13
            Top             =   360
            Width           =   1260
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
         Left            =   4350
         TabIndex        =   6
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
         Top             =   5880
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
         SpreadDesigner  =   "I_Planif.frx":061C
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3015
         Left            =   5280
         TabIndex        =   25
         Top             =   2040
         Visible         =   0   'False
         Width           =   2415
         _Version        =   393216
         _ExtentX        =   4260
         _ExtentY        =   5318
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
         SpreadDesigner  =   "I_Planif.frx":0828
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
         Picture         =   "I_Planif.frx":0EDC
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
         Top             =   1520
         Width           =   4320
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
         Width           =   4320
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
         Width           =   4320
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
         Picture         =   "I_Planif.frx":11E6
         Top             =   770
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2475
         Picture         =   "I_Planif.frx":14F0
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
         Left            =   3225
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
         Width           =   4320
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2880
         TabIndex        =   31
         Top             =   1220
         Width           =   4320
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2880
         TabIndex        =   37
         Top             =   1540
         Width           =   4320
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   7605
      _ExtentX        =   13414
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
Dim RS As New ADODB.Recordset
Dim i As Integer, isel As Integer
Dim accion As Boolean
Public lc_Aux As String
Dim Msgtitulo As String, TipMin As String

Private Sub Check3_Click()
If Check3.Value = 1 Then
   Check3.Caption = "Incluye Parentesis"
Else
   Check3.Caption = "No Incluye Parentesis"
End If
End Sub

Private Sub Combo1_Click(Index As Integer)
If Val(fg_codigocbo(Combo1, 0, 1, "")) <> 5 Then Check2.Visible = True Else Check2.Visible = False
Select Case Val(fg_codigocbo(Combo1, 0, 1, ""))
Case 0, 1, 5
    Frame3(1).Enabled = False: Frame3(4).Enabled = False: Frame5.Enabled = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 1, True, False)
    Option1(2).Enabled = False: Option1(3).Enabled = False
    Option1(4).Enabled = False: Option1(5).Enabled = False
    Option1(6).Enabled = False: Option1(7).Enabled = False
    If Val(fg_codigocbo(Combo1, 0, 1, "")) <> 1 Then Frame7.Enabled = False: Check1(0).Value = 0: Check1(1).Value = 0 Else Frame7.Enabled = True

Case 2
    Frame5.Enabled = False
    Frame3(1).Enabled = True: Frame3(4).Enabled = True
    Option1(2).Enabled = True: Option1(3).Enabled = True
    Option1(4).Enabled = True: Option1(5).Enabled = True
    Option1(6).Enabled = True: Option1(7).Enabled = True
Case 3, 4
    Frame5.Enabled = False
    Frame3(1).Enabled = True: Frame3(4).Enabled = False
    Option1(2).Enabled = True: Option1(3).Enabled = True
    Option1(4).Enabled = False: Option1(5).Enabled = False
    Option1(6).Enabled = False: Option1(7).Enabled = False

End Select
Select Case Index
Case 1
    MoverDatoGrilla
End Select
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = IIf(lc_Aux = "Planif", 7725, 7725)
Me.Width = 7575
Bar1(0).Top = IIf(lc_Aux = "Planif", 5880, 6120)
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
If lc_Aux = "Planif" Then
   Msgtitulo = "Informe Planificación"
   Me.Caption = "Informe Planificación"
   Combo1(0).AddItem "Menú Mecano" & Space(150) & "(0)"
   Combo1(0).AddItem "Menú Mensual" & Space(150) & "(1)"
   Combo1(0).AddItem "Menú Mensual Servicios" & Space(150) & "(5)"
   Combo1(0).AddItem "Aporte Nutricionales Detallado" & Space(150) & "(2)"
   Combo1(0).AddItem "Aporte Nutricionales Resumido" & Space(150) & "(3)"
   Combo1(0).AddItem "Aporte Nutricionales por Estructura" & Space(150) & "(4)"
   Combo1(0).ListIndex = -1
Else
   Msgtitulo = "Informe Costo Minuta"
   Me.Caption = "Informe Costo Minuta"
   Combo1(0).AddItem "Costo Minuta Detallado" & Space(150) & "(0)"
   Combo1(0).AddItem "Costo Minuta por Estructura" & Space(150) & "(1)"
   Combo1(0).ListIndex = -1
End If

fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
accion = True
'-------> Llenar tabla nutrientes
Set RS = vg_db.Execute("sgpadm_s_nutriente 1, 0, ''")
If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No existe maestro nutrientes", vbExclamation + vbOKOnly, Msgtitulo: Me.Hide: Unload Me
vaSpread2.MaxRows = 0
Do While Not RS.EOF
   vaSpread2.MaxRows = vaSpread2.MaxRows + 1
   vaSpread2.Row = vaSpread2.MaxRows
   vaSpread2.Col = 2: vaSpread2.text = RS!nut_codigo
   vaSpread2.Col = 3: vaSpread2.text = Trim(RS!nut_nombre) & " " & Trim(RS!nut_nomuni)
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
RS.Close: Set RS = Nothing

OpUsuario = vg_Indppr
If IsNull(OpUsuario) Or Trim(OpUsuario) = "" Then
    MsgBox "Contactese con el Administrador del Sistema...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
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
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
MoverDatoGrilla
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Select Case Index
Case 0
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
       Set RS = vg_db.Execute("sgpadm_s_subsegmento 10, " & Val(fpLongInteger1(0).Value) & ", '', '" & vg_Indppr & "'")
    Else
       Set RS = vg_db.Execute("sgpadm_s_subsegmento 1, " & Val(fpLongInteger1(0).Value) & ", '', ''")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing
    MoverDatoGrilla
Case 1
'    Set RS = vg_db.Execute("sgpadm_s_regimen 1, " & Val(fpLongInteger1(1).Value) & ", ''")
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & " AND reg_indppr = '" & vg_Indppr & "'")
    Else
      Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & "")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    MoverDatoGrilla
Case 2
    Set RS = vg_db.Execute("sgpadm_s_zona 1, " & Val(fpLongInteger1(2).Value) & ", ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "": Exit Sub
    fpayuda(2).Caption = Trim(RS!Zon_nombre)
    RS.Close: Set RS = Nothing
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    If Index = 0 Then Image1_Click 0
    If Index = 1 Then Image1_Click 1
End Select
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
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
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    fpLongInteger1(2).SetFocus
Case 2
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpLongInteger1(0).Value = "" Or fpLongInteger1(1).Value = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", Val(fg_codigocbo(Combo1, 1, 1, ""))
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
Case 3
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpLongInteger1(0).Value = "" Or fpLongInteger1(1).Value = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Nutrientes", Me.vaSpread2, fpLongInteger1(0).Value, fpLongInteger1(1).Value, fpLongInteger1(2).Value, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "2", Val(fg_codigocbo(Combo1, 1, 1, ""))
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
Case 4
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_zona", "zon_", "Zona", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre
    fpDateTime1(0).SetFocus
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
Dim opnomrec As Boolean, spid As Long
Select Case Button.Index
Case 1
    If vaSpread1.MaxRows < 1 Then MsgBox "No existe Información", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Set RS = vg_db.Execute("sgpadm_s_subsegmento 1, " & Val(fpLongInteger1(0).Value) & ", '', ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(0).Value = "": fpayuda(0).Caption = "": MsgBox "No existe subsegmento", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    Set RS = vg_db.Execute("sgpadm_s_regimen 1, " & Val(fpLongInteger1(1).Value) & ", ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": MsgBox "No existe regimen", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    If Val(fg_codigocbo(Combo1, 0, 1, "")) = 2 Or Val(fg_codigocbo(Combo1, 0, 1, "")) = 3 Then
       Set RS = vg_db.Execute("sgpadm_s_zona 1, " & Val(fpLongInteger1(2).Value) & ", ''")
       If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "": MsgBox "No existe zona", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
       RS.Close: Set RS = Nothing
    End If
    If fpDateTime1(0).Value > fpDateTime1(1).Value Then MsgBox "Fecha origen Mayor destino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Mid(fpDateTime1(0).text, 4, 2) > Mid(fpDateTime1(1).text, 4, 2) Then MsgBox "Mes origen mayor destino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Mid(fpDateTime1(0).text, 7, 4) > Mid(fpDateTime1(1).text, 7, 4) Then MsgBox "Año origen mayor destino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
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
    End If
    '-------> Borrar tabla paso servicio
    vg_db.Execute "DELETE paso_servicio WHERE ser_spid = @@spid and ser_usr = '" & vg_NUsr & "'"
    isel = 0
    '-------> Buscar spid
    Set RS = vg_db.Execute("SELECT @@spid spid")
    If Not RS.EOF Then spid = RS!spid
    RS.Close: Set RS = Nothing
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then
           isel = 1
           vaSpread1.Col = 2
           vg_db.Execute "INSERT INTO paso_servicio (ser_spid, ser_usr, ser_codigo) VALUES (" & spid & ", '" & vg_NUsr & "', " & Val(vaSpread1.text) & ")"
        End If
    Next i
    If isel = 0 Then fg_descarga: MsgBox "Servicio debe ser informado", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    opnomrec = True
    If Option1(9).Value = True Then opnomrec = False
    Toolbar1.Enabled = False
    Frame1(0).Enabled = False
    If lc_Aux = "Planif" Then
    vg_CallForm = Me.Name
    vg_CallFormDato = Combo1(0).text
       Select Case Val(fg_codigocbo(Combo1, 0, 1, ""))
       Case 0
          I_MenuPlanMecano Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, Val(fg_codigocbo(Combo1, 1, 1, "")), Option1(11), Check2.Value, Check3.Value
       Case 1
        If Check2.Value = 1 Then
'            I_MenuPlanMensual2 Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "1", opnomrec, Val(fg_codigocbo(Combo1, 1, 1, "")), Check1(0).Value, Check1(1).Value, Option1(11), Check2.Value
            I_MenuPlanMensualSemanaCerradaok Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, Val(fg_codigocbo(Combo1, 1, 1, "")), Check1(0).Value, Check1(1).Value, Option1(11), Check2.Value, Check3.Value
        Else
            I_MenuPlanMensual Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, Val(fg_codigocbo(Combo1, 1, 1, "")), Check1(0).Value, Check1(1).Value, Option1(11), Check2.Value, Check3.Value
        End If
       Case 2
          I_AportePlanDetallado Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, spid, Val(fg_codigocbo(Combo1, 1, 1, "")), Check2.Value, Check3.Value
       Case 3
          I_AportePlanRes Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, spid, Val(fg_codigocbo(Combo1, 1, 1, "")), Check2.Value, Check3.Value
'          I_AportePlanResumido Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec
       Case 4
          I_AportePlanEstrRes Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, spid, Val(fg_codigocbo(Combo1, 1, 1, "")), Option1(11), Check2.Value
       Case 5
'          I_MenuPlanMensualServicio Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(Format(fpDateTime1(0).Text, "yyyymmdd")), Val(Format(fpDateTime1(1).Text, "yyyymmdd")), "1", opnomrec, Val(fg_codigocbo(Combo1, 1, 1, "")), Check1(0).Value, Check1(1).Value, Option1(11), Check2.Value
          I_MenuPlanMensualServicioOk Me, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, Val(fg_codigocbo(Combo1, 1, 1, "")), Check1(0).Value, Check1(1).Value, Option1(11), Check2.Value, spid, Check3.Value
       End Select
    Else
       Select Case Val(fg_codigocbo(Combo1, 0, 1, ""))
       Case 0
          I_CostoDetMinuta Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, spid, vg_NUsr, Val(fg_codigocbo(Combo1, 1, 1, "")), Option1(11), Check2.Value
       Case 1
          'I_CostoPlanEstrRes Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, spid, vg_NUsr, Val(fg_codigocbo(Combo1, 1, 1, ""))
          I_CostoPlanEstrRes Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, spid, vg_NUsr, Val(fg_codigocbo(Combo1, 1, 1, "")), Option1(11), Check2.Value, IIf(Option2(0).Value = True, True, False)
       End Select
    End If
    vg_db.Execute "DELETE paso_servicio WHERE ser_spid= " & spid & " AND ser_usr= '" & vg_NUsr & "'"
    Toolbar1.Enabled = True
    Frame1(0).Enabled = True
Case 3
    'Historico Planificación
    Set RS = vg_db.Execute("sgpadm_s_subsegmento 1, " & Val(fpLongInteger1(0).Value) & ", '', ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(0).Value = "": fpayuda(0).Caption = "": MsgBox "No existe subsegmento", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
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
End Sub

Sub MoverDatoGrilla()
If Val(fpLongInteger1(0).Value) = 0 And Val(fpLongInteger1(1).Value) = 0 And fpDateTime1(0).text = "" And fpDateTime1(1).text = "" Then Exit Sub
fg_carga ""
vaSpread1.MaxRows = 0
'RS.Open "SELECT DISTINCT b_minuta.min_codser, a_servicio.ser_nombre, a_servicio.ser_orden " & _
'        "FROM  a_servicio, b_minuta, b_minutadet " & _
'        "WHERE b_minuta.min_codigo=b_minutadet.mid_codigo " & _
'        "AND   b_minuta.min_codser=a_servicio.ser_codigo " & _
'        "AND   b_minuta.min_subseg=" & Val(fpLongInteger1(0).Value) & " " & _
'        "AND   b_minuta.min_codreg=" & Val(fpLongInteger1(1).Value) & " " & _
'        "AND   b_minuta.min_fecmin>=" & Val(Format(fpDateTime1(0).Text, "yyyymmdd")) & " " & _
'        "AND   b_minuta.min_fecmin<=" & Val(Format(fpDateTime1(1).Text, "yyyymmdd")) & " " & _
'        "AND   b_minutadet.mid_tipmin='1' " & _
'        "ORDER BY a_servicio.ser_orden, a_servicio.ser_nombre", vg_db, adOpenStatic
Set RS = vg_db.Execute("sgpadm_s_planifminuta 7, " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", 0,0, 0, " & Val(Format(fpDateTime1(0).text, "yyyymmdd")) & ", " & Val(Format(fpDateTime1(1).text, "yyyymmdd")) & ",'" & Val(fg_codigocbo(Combo1, 1, 1, "")) & "'")
If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: Exit Sub
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 2: vaSpread1.text = RS!min_codser
   vaSpread1.Col = 3: vaSpread1.text = Trim(RS!ser_nombre)
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing: fg_descarga
If Val(fg_codigocbo(Combo1, 1, 1, "")) = 4 Then
' Llena Estructura minuta

End If

End Sub


