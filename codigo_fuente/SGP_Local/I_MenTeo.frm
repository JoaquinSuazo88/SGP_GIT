VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_MenTeo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Planificación Teórica"
   ClientHeight    =   5805
   ClientLeft      =   2685
   ClientTop       =   2745
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   5265
      Index           =   0
      Left            =   60
      TabIndex        =   7
      Top             =   480
      Width           =   7395
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Ponderación"
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
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   3840
         Width           =   7095
         Begin VB.CheckBox Check1 
            Caption         =   "Impirmir Recetas Sin Ponderación"
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
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   200
            Width           =   3375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Ponderación"
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
            Index           =   1
            Left            =   4800
            TabIndex        =   34
            Top             =   195
            Width           =   2175
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   135
         Index           =   1
         Left            =   0
         TabIndex        =   32
         Top             =   0
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
         SpreadDesigner  =   "I_MenTeo.frx":0000
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   3375
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
            Index           =   11
            Left            =   2040
            TabIndex        =   31
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
            Index           =   10
            Left            =   120
            TabIndex        =   30
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   4
            Left            =   2760
            Picture         =   "I_MenTeo.frx":06B4
            Top             =   165
            Width           =   480
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   26
         Top             =   4800
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
         SpreadDesigner  =   "I_MenTeo.frx":09BE
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
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   3120
         Width           =   7095
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
            TabIndex        =   25
            Top             =   240
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
            TabIndex        =   24
            Top             =   240
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
            TabIndex        =   23
            Top             =   240
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
            TabIndex        =   22
            Top             =   240
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
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   3375
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
            Left            =   2040
            TabIndex        =   20
            Top             =   300
            Width           =   855
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
            TabIndex        =   19
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   3
            Left            =   2760
            Picture         =   "I_MenTeo.frx":0BCA
            Top             =   165
            Width           =   480
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   7095
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "I_MenTeo.frx":0ED4
            Left            =   1680
            List            =   "I_MenTeo.frx":0ED6
            Style           =   2  'Dropdown List
            TabIndex        =   11
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
            TabIndex        =   12
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
         Left            =   3840
         TabIndex        =   9
         Top             =   1680
         Width           =   3375
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
            TabIndex        =   3
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
            Left            =   2160
            TabIndex        =   4
            Top             =   300
            Width           =   735
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   2
            Left            =   2880
            Picture         =   "I_MenTeo.frx":0ED8
            Top             =   165
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
         TabIndex        =   8
         Top             =   4560
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
            TabIndex        =   5
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
            TabIndex        =   6
            Top             =   300
            Width           =   1665
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   135
         Index           =   0
         Left            =   5880
         TabIndex        =   13
         Top             =   4800
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
         SpreadDesigner  =   "I_MenTeo.frx":11E2
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   960
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
         TabIndex        =   1
         Top             =   1305
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
         Left            =   4950
         TabIndex        =   2
         Top             =   1305
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3120
         TabIndex        =   27
         Top             =   960
         Width           =   4095
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
         Left            =   3825
         TabIndex        =   17
         Top             =   1380
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
         TabIndex        =   16
         Top             =   1380
         Width           =   1110
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2685
         Picture         =   "I_MenTeo.frx":1896
         Top             =   870
         Width           =   480
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
         Index           =   7
         Left            =   285
         TabIndex        =   14
         Top             =   1035
         Width           =   735
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3165
         TabIndex        =   28
         Top             =   1005
         Width           =   4095
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
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
Dim MsgTitulo As String, est As Boolean
Public lc_Aux As String

Private Sub Combo1_Click(Index As Integer)
Frame3(3).Enabled = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 0 Or Val(fg_codigocbo(Combo1, 0, 1, "")) = 1, True, False)
Check1(0).Enabled = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 0 Or Val(fg_codigocbo(Combo1, 0, 1, "")) = 1, True, False)
Check1(1).Enabled = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 0 Or Val(fg_codigocbo(Combo1, 0, 1, "")) = 1, True, False)
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
Case 8
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
Me.Height = 6285
Me.Width = 7665
est = True
fg_centra Me
If lc_Aux = "PlaTei" Then
   Me.Caption = "Informe Planificación Teórica"
   MsgTitulo = "Informe Planificación Teórica"
Else
   Me.Caption = "Informe Planificación Real"
   MsgTitulo = "Informe Planificación Real"
End If
EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)
'Dim btnX As Button
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico Planificacón Teórica"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
Combo1(0).Clear
Combo1(0).AddItem "Menú Mecano" & Space(150) & "(0)"
Combo1(0).AddItem "Menú Mensual" & Space(150) & "(1)"
Combo1(0).AddItem "Aporte Nutricionales Detallado" & Space(150) & "(2)"
Combo1(0).AddItem "Aporte Nutricionales Resumido" & Space(150) & "(3)"
Combo1(0).AddItem "Costo Detallado" & Space(150) & "(4)"
Combo1(0).AddItem "Costo Resumido" & Space(150) & "(5)"
Combo1(0).AddItem "Ingredientes Valor Cero en Planificación" & Space(150) & "(6)"
Combo1(0).AddItem "Menú Mensual Servicios" & Space(150) & "(7)"
Combo1(0).AddItem "Menú Día Aporte Nutricionales" & Space(150) & "(8)"
Combo1(0).ListIndex = 0

'------- Llenar Tabla Nutrientes
RS.Open RutinaLectura.Nutriente(1, 0, ""), vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No existe maestro nutrientes", vbExclamation + vbOKOnly, MsgTitulo: Me.Hide: Unload Me
With vaSpread1(2)
    .MaxRows = 0
    Do While Not RS.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 2: .text = RS!nut_codigo
       .Col = 3: .text = Trim(RS!nut_nombre)
       If RS!nut_indpri = 1 Then
          .Col = 1
          .CellType = 10
          .TypeCheckText = ""
          .TypeCheckCenter = True
          .text = "1" ' checked
       Else
          .Col = 1
          .CellType = 10
          .TypeCheckText = ""
          .TypeCheckCenter = True
          .text = " " ' checked
       End If
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
End With
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
codreg = "": est = False
MoverDatoGrilla
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
MoverDatoGrilla
codreg = ""
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
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
If fpText.text = "" Then fpayuda(0).Caption = "": Exit Sub
RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
MoverDatoGrilla
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
Case 2
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpText.text = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, fpText.text, "", Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "0", lc_Aux, 1, IIf(lc_Aux = "PlaTei", "'1'", "'2'")
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
Case 3
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpText.text = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Nutrientes", Me.vaSpread1, fpText.text, "", Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "2", lc_Aux, 2, IIf(lc_Aux = "PlaTei", "'1'", "'2'")
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
Case 4
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpText.text = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Regimen", Me.vaSpread1, fpText.text, "", Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "0", lc_Aux, 0, IIf(lc_Aux = "PlaTei", "'1'", "'2'")
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
Case 10
    Image1(4).Enabled = False
Case 11
    Image1(4).Enabled = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim opnomrec As Boolean, codser As String, codreg As String, opponderacion As Boolean, opracion As Boolean
Select Case Button.Index
Case 1
    If vaSpread1(1).MaxRows < 1 Then MsgBox "No existe Información", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda(0).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    If fpDateTime1(0).Value > fpDateTime1(1).Value Then MsgBox "Fecha origen Mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Mid(fpDateTime1(0).text, 4, 2) <> Mid(fpDateTime1(1).text, 4, 2) Then MsgBox "Mes origen mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Mid(fpDateTime1(0).text, 7, 4) <> Mid(fpDateTime1(1).text, 7, 4) Then MsgBox "Año origen mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Combo1(0).ListIndex = -1 Or Combo1(0).text = "" Then Exit Sub
    codreg = "": codser = ""
    With vaSpread1(0)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .text = "1" Then .Col = 2: codreg = codreg & "" & .text & ","
        Next i
    End With
    With vaSpread1(1)
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .text = "1" Then .Col = 2: codser = codser & "" & .text & ","
        Next i
    End With
    If Trim(codreg) = "" Then fg_descarga: MsgBox "Regimen debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Trim(codser) = "" Then fg_descarga: MsgBox "Servicio debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    opnomrec = True
    If Option1(9).Value = True Then opnomrec = False
    Frame1(0).Enabled = False
    Toolbar1.Enabled = False
    opponderacion = IIf(Check1(0).Value = 1, True, False)
    opracion = IIf(Check1(1).Value = 1, True, False)
    Select Case Val(fg_codigocbo(Combo1, 0, 1, ""))
    Case 0
       If lc_Aux = "PlaTei" Then I_MenuPlanMecano fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, opponderacion, opracion
       If lc_Aux = "PlaRei" Then I_MenuPlanMecano fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "2", opnomrec, opponderacion, opracion
    Case 1
       If lc_Aux = "PlaTei" Then I_MenuPlanMensual fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, opponderacion, opracion
       If lc_Aux = "PlaRei" Then I_MenuPlanMensual fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "2", opnomrec, opponderacion, opracion
    Case 2
       If lc_Aux = "PlaTei" Then I_AportePlanDetallado fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, Me
       If lc_Aux = "PlaRei" Then I_AportePlanDetallado fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "2", opnomrec, Me
    Case 3
       If lc_Aux = "PlaTei" Then I_AportePlanResumido fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, Me
       If lc_Aux = "PlaRei" Then I_AportePlanResumido fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "2", opnomrec, Me
    Case 4
       If lc_Aux = "PlaTei" Then I_CostoPlanDetallado fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec
       If lc_Aux = "PlaRei" Then I_CostoPlanDetallado fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "2", opnomrec
    Case 5
       If lc_Aux = "PlaTei" Then I_CostoPlanResumido fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec
       If lc_Aux = "PlaRei" Then I_CostoPlanResumido fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "2", opnomrec
    Case 6
       If lc_Aux = "PlaTei" Then I_IngValCeroPlan fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec
       If lc_Aux = "PlaRei" Then I_IngValCeroPlan fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "2", opnomrec
    Case 7
       I_MenuPlanMensualServicio Me, fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), IIf(lc_Aux = "PlaTei", "1", "2"), opnomrec
    Case 8
       If lc_Aux = "PlaTei" Then I_MenuDiaAporteNutricional fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "1", opnomrec, Me
       If lc_Aux = "PlaRei" Then I_MenuDiaAporteNutricional fpText.text, codreg, codser, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), "2", opnomrec, Me
    End Select
    Frame1(0).Enabled = True
    Toolbar1.Enabled = True
Case 3
    'Historico planificación
    RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda(0).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    vg_codigo = ""
    If lc_Aux = "PlaTei" Then B_HistPm.LlenarHistPlan "Histórico Planificación Teórica", fpText.text, 1, 2
    If lc_Aux = "PlaRei" Then B_HistPm.LlenarHistPlan "Histórico Planificación Real", fpText.text, 2, 2
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    fpDateTime1(0).text = "01/" & vg_auxfecha: fpDateTime1(1).text = dEoM("01/" & vg_auxfecha)
    MoverDatoGrilla
    Option1(0).SetFocus
    Me.Refresh
Case 5
    Me.Hide
    Unload Me
End Select
End Sub

Sub MoverDatoGrilla()
fg_carga ""
With vaSpread1(0)
    .MaxRows = 0
    RS.Open RutinaLectura.Regimen(1, 0, ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: Exit Sub
    Do While Not RS.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1: .text = "1"
       .Col = 2: .text = RS!reg_codigo
       .Col = 3: .text = Trim(RS!reg_nombre)
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
End With
With vaSpread1(1)
    .MaxRows = 0
    RS.Open RutinaLectura.Servicio(1, 0, ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: Exit Sub
    Do While Not RS.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1: .text = "1"
       .Col = 2: .text = RS!ser_codigo
       .Col = 3: .text = Trim(RS!ser_nombre)
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
End With
fg_descarga
End Sub
