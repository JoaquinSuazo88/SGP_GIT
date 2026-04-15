VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_SetPlaSansis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planificación Minuta Sansis"
   ClientHeight    =   8025
   ClientLeft      =   3315
   ClientTop       =   2415
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8025
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   7425
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   450
      Width           =   8060
      Begin VB.Frame Frame5 
         Caption         =   "Tipo Minuta"
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
         TabIndex        =   32
         Top             =   1440
         Width           =   3375
         Begin VB.OptionButton Option5 
            Caption         =   "Teórica"
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
            Left            =   240
            TabIndex        =   34
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Real"
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
            Left            =   2160
            TabIndex        =   33
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Opción Impresión"
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
         Left            =   3720
         TabIndex        =   31
         Top             =   1440
         Width           =   4215
         Begin VB.OptionButton Option3 
            Caption         =   "Oficio"
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
            Left            =   1920
            TabIndex        =   12
            Top             =   320
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Carta"
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
            TabIndex        =   11
            Top             =   320
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3795
         Left            =   150
         TabIndex        =   30
         Top             =   3510
         Width           =   7755
         Begin VB.TextBox Text1 
            Height          =   3135
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   13
            Top             =   240
            Width           =   7455
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Incorpora Inserto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   14
            Top             =   3450
            Width           =   2205
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   3720
         TabIndex        =   23
         Top             =   2880
         Width           =   4215
         Begin VB.OptionButton Option1 
            Caption         =   "Con Fecha"
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
            Left            =   2160
            TabIndex        =   10
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Sin Fecha"
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
            Left            =   240
            TabIndex        =   9
            Top             =   300
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   3375
         Begin VB.OptionButton Option1 
            Caption         =   "Sin Código"
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
            Left            =   120
            TabIndex        =   7
            Top             =   300
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Con Código"
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
            Left            =   1920
            TabIndex        =   8
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   3720
         TabIndex        =   20
         Top             =   2160
         Width           =   4215
         Begin VB.OptionButton Option2 
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
            Index           =   1
            Left            =   2160
            TabIndex        =   6
            Top             =   300
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
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
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   300
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   2160
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
            Index           =   1
            Left            =   1920
            TabIndex        =   4
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
            TabIndex        =   3
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   2790
            Picture         =   "I_SetPlaSansis.frx":0000
            Top             =   150
            Width           =   480
         End
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   135
         Index           =   0
         Left            =   3600
         TabIndex        =   17
         Top             =   3000
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
         SpreadDesigner  =   "I_SetPlaSansis.frx":030A
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1695
         TabIndex        =   0
         Top             =   375
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
         BackColor       =   16777215
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
         AlignTextV      =   2
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
         NullColor       =   -2147483624
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   5
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
      Begin EditLib.fpLongInteger Regimen 
         Height          =   315
         Left            =   1695
         TabIndex        =   1
         Top             =   720
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
         BackColor       =   16777215
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
         NullColor       =   16777215
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "0"
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
         Left            =   1695
         TabIndex        =   2
         Top             =   1065
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1658
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
         BackColor       =   16777215
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
         NullColor       =   -2147483624
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "10/2021"
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
         AutoMenu        =   0   'False
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
         Index           =   1
         Left            =   3000
         TabIndex        =   28
         Top             =   720
         Width           =   4845
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3000
         TabIndex        =   26
         Top             =   375
         Width           =   4845
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2550
         Picture         =   "I_SetPlaSansis.frx":4770
         Top             =   600
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2550
         Picture         =   "I_SetPlaSansis.frx":4A7A
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha(mm/aa)"
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
         Left            =   435
         TabIndex        =   24
         Top             =   1110
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ceco"
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
         Left            =   435
         TabIndex        =   21
         Top             =   450
         Width           =   450
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
         Index           =   7
         Left            =   435
         TabIndex        =   16
         Top             =   795
         Width           =   750
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3045
         TabIndex        =   27
         Top             =   405
         Width           =   4845
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3045
         TabIndex        =   29
         Top             =   765
         Width           =   4845
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   615
      Index           =   0
      Left            =   2400
      TabIndex        =   18
      Top             =   7080
      Visible         =   0   'False
      Width           =   1365
      _Version        =   393216
      _ExtentX        =   2408
      _ExtentY        =   1085
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
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
      MaxRows         =   13
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "I_SetPlaSansis.frx":4D84
      StartingColNumber=   6
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_SetPlaSansis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgTitulo As String
Public lc_Aux As String

Private Sub Check1_Click()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If Check1.Value = 1 Then
   
   Text1.Enabled = True
   RS.Open "sgp_Sel_ParametroSansis 9999", vg_db, adOpenDynamic
   If Not RS.EOF Then
      
      Text1.text = RS(0)
   
   End If
   RS.Close
   Set RS = Nothing

Else
   
   Text1.text = ""
   Text1.Enabled = False

End If

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Check2_Click()

On Error GoTo Man_Error

If Check2.Value = 1 Then
   
   Combo1.Enabled = True

Else
   
   Combo1.ListIndex = -1
   Combo1.Enabled = False

End If

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.Height = 8400
Me.Width = 8415
Me.HelpContextID = vg_OpcM
fg_centra Me
MsgTitulo = "Planificación Minutas Sansis"

Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico Planificacón Teórica"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

fpText1.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText1.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)

fpDateTime1.text = Format(Date, "mm/yyyy")
vaSpread1(0).MaxRows = 0
vaSpread2(0).MaxRows = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpDateTime1_Change()

On Error GoTo Man_Error

If IsDate(fpDateTime1.text) = False Then Exit Sub
MoverDatosVector

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub


Private Sub Regimen_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
    
If Val(Regimen.Value) < 1 Then
       
   fpayuda(1).Caption = ""
   Exit Sub
    
End If
    
Set RS = vg_db.Execute("sgp_Sel_RegimenxCodigo " & Regimen.Value & "")
  
If RS.EOF Then
       
   RS.Close
   Set RS = Nothing
   fpayuda(1).Caption = ""
   Exit Sub
    
End If
    
fpayuda(1).Caption = Trim(RS!reg_nombre)
RS.Close
Set RS = Nothing
MoverDatosVector

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Regimen_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Set RS = vg_db.Execute("sgp_Sel_clientes 1, '" & LimpiaDato(fpText1.text) & "'")
If RS.EOF Then

   RS.Close
   Set RS = Nothing
   fpayuda(0).Caption = ""
   Regimen.text = ""
   fpayuda(1).Caption = ""
   fpDateTime1.Enabled = True
   Exit Sub

End If

fpayuda(0).Caption = Trim(RS!cli_nombre)
fpText1.text = RS!cli_codigo
RS.Close
Set RS = Nothing
 
fpDateTime1.Enabled = True

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0
    
    Option1(0).Value = True
    Option1(1).Value = False
    Image1(2).Enabled = False

Case 1
    
    Option1(0).Value = False
    Option1(1).Value = True
    Image1(2).Enabled = True

End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Option1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 120
    
    Image1_Click 2

End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0

    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    Call B_TabEst.LlenaDatos("b_clientes", "cli_", "Clientes", "Cliente_SitioRemoto")
    Call B_TabEst.Show(1)
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText1.text = vg_codigo
    fpayuda(0).Caption = vg_nombre
    Regimen.Value = ""
    Let fpayuda(1).Caption = ""
    Regimen.SetFocus

Case 1
    
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = ""
    vg_codigo = ""
    Call B_TabEst.LlenaDatos("a_regimen", "reg_", "Regimen", "RegBlo")
    Call B_TabEst.Show(1)
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    Regimen.Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    fpDateTime1.SetFocus
    
Case 2

    OpcionLectura = "6"
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = ""
    vg_codigo = ""
    vg_codigo = Trim(LimpiaDato(fpText1.text))
    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, vg_codigo, Regimen.text & ",", Format(fpDateTime1.text, "yyyymmdd"), Format(fpDateTime1.text, "yyyymmdd"), "1", "", 0, IIf(Option5(0).Value = True, 1, 2)
'    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, 0, Val(Regimen.Value), 0, Format(fpDateTime1.text, "yyyymm"), 0, "6", 0
    B_MTaEst.Show 1
    Me.Refresh
    
    If vg_codigo = "" Then
       
       Exit Sub
    
    End If

End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index

Case 1
    
    If Not ValidarDatos Then Exit Sub
    
    vaSpread2(0).MaxRows = 0
    vaSpread2(0).MaxCols = 0
    
    Toolbar1.Enabled = False
    Frame1(0).Enabled = False
    vg_opimp = 0
    vg_opimp = 9999
    
    I_SetMinutaSansis Me
    
    vg_opimp = 0
    Toolbar1.Enabled = True
    Frame1(0).Enabled = True
    
Case 3
    
    Set RS = vg_db.Execute("sgp_Sel_Clientes 1, '" & Trim(LimpiaDato(fpText1.text)) & "'")
    If RS.EOF Then
       
       RS.Close
       Set RS = Nothing
       MsgBox "No existe ceco planificado", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    vg_codigo = ""
    
'    B_HistPm.LlenarHistPlan "Histórico Minuta", 0, Trim(LimpiaDato(fpText1.text)), 5
    B_HistPm.LlenarHistPlan "Histórico Minuta", Trim(LimpiaDato(fpText1.text)), 2, 1
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    Regimen.Value = vg_codregimen
    fpDateTime1.text = vg_fecha
    Me.Refresh

Case 5
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub MoverDatosVector()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If LimpiaDato(Trim(fpText1.text)) = "" Or fpDateTime1.text = "" Or Val(Regimen.Value) = 0 Then

    Exit Sub

End If

fg_carga ""
Set RS = vg_db.Execute("sgp_Sel_ServicioMinutaMes '" & LimpiaDato(Trim(fpText1.text)) & "', " & Regimen.Value & ", " & Format(fpDateTime1.text, "yyyymm") & "")
vaSpread1(0).MaxRows = 0

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
      vaSpread1(0).Row = vaSpread1(0).MaxRows
      
      vaSpread1(0).Col = 2
      vaSpread1(0).Value = RS(0)
      
      vaSpread1(0).Col = 3
      vaSpread1(0).Value = Trim(RS(1))
      
      RS.MoveNext
   
   Loop

End If
RS.Close
Set RS = Nothing
fg_descarga

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub ActivarGrillaTodos()

On Error GoTo Man_Error

For i = 1 To vaSpread1(0).MaxRows
    
    vaSpread1(0).Row = i
    vaSpread1(0).Col = 1
    vaSpread1(0).CellType = 10
    vaSpread1(0).TypeCheckText = ""
    vaSpread1(0).TypeCheckCenter = True
    vaSpread1(0).text = "1" ' checked

Next i

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Function ValidarDatos() As Boolean

On Error GoTo Man_Error

Dim seleccion As Integer
Dim i As Long

ValidarDatos = True

'-------> Validar ceco
If Trim(fpayuda(0).Caption) = "" Then

   MsgBox "Debe registrar ceco...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

'-------> Validar regimen
If Trim(fpayuda(1).Caption) = "" Then

   MsgBox "Debe registrar regimen...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

'-------> Validar fechas
If Trim(fpDateTime1.text) = "" Then
   
   MsgBox "Fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

'If Combo1.ListIndex = -1 And Check2.Value = 1 Then
'
'   MsgBox "Seleccione Calcular Costo x", vbExclamation + vbOKOnly, Msgtitulo
'   ValidarDatos = False
'   Exit Function
'
'End If
'
If Option1(0).Value = True Then
    
   ActivarGrillaTodos
    
End If

'-------> Validar servicios
seleccion = 0
For i = 1 To vaSpread1(0).MaxRows
        
    vaSpread1(0).Row = i
    vaSpread1(0).Col = 1
    
    If vaSpread1(0).text = "1" Then
       
       seleccion = 1
       Exit For
    
    End If
    
Next i
    
If seleccion = 0 Then
   
   MsgBox "Servicio debe ser selecionado", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function
   
End If
    
Exit Function
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
   
End Function
