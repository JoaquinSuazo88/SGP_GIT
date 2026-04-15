VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form I_SetPla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Planificación Minuta"
   ClientHeight    =   8025
   ClientLeft      =   3315
   ClientTop       =   2415
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8025
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   7425
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   450
      Width           =   8060
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
         Left            =   120
         TabIndex        =   37
         Top             =   3360
         Width           =   3375
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
            TabIndex        =   39
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
            TabIndex        =   38
            Top             =   320
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   120
         TabIndex        =   33
         Top             =   3990
         Width           =   7785
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "I_SetPla.frx":0000
            Left            =   1905
            List            =   "I_SetPla.frx":000D
            TabIndex        =   35
            Top             =   450
            Width           =   1935
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Incorpora Food Cost"
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
            Left            =   120
            TabIndex        =   34
            Top             =   180
            Width           =   2415
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Calcular Costo x"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   390
            TabIndex        =   32
            Top             =   510
            Width           =   1320
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2355
         Left            =   150
         TabIndex        =   30
         Top             =   4950
         Width           =   7755
         Begin VB.TextBox Text1 
            Height          =   1700
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   36
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
            TabIndex        =   31
            Top             =   2010
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
         TabIndex        =   19
         Top             =   2640
         Width           =   4215
         Begin VB.OptionButton Option1 
            Caption         =   "Con Fecha"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   2160
            TabIndex        =   21
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Sin Fecha"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   20
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
         TabIndex        =   16
         Top             =   2640
         Width           =   3375
         Begin VB.OptionButton Option1 
            Caption         =   "Sin Código"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   18
            Top             =   300
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Con Código"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   17
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   3720
         TabIndex        =   14
         Top             =   1920
         Width           =   4215
         Begin VB.OptionButton Option2 
            Caption         =   "Nombre Fantasia"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   7
            Top             =   300
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Nombre Receta"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   300
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Servicio"
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
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   3375
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   5
            Top             =   300
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
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
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   2790
            Picture         =   "I_SetPla.frx":002D
            Top             =   150
            Width           =   480
         End
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   135
         Left            =   3600
         TabIndex        =   11
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
         SpreadDesigner  =   "I_SetPla.frx":0337
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1695
         TabIndex        =   3
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
         BackColor       =   -2147483624
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483624
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
      Begin EditLib.fpText fpText 
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
         BackColor       =   -2147483624
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
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
         BackColor       =   -2147483624
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
         NullColor       =   -2147483624
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
         BackColor       =   -2147483624
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
         Text            =   "03/2011"
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
         Index           =   2
         Left            =   3000
         TabIndex        =   28
         Top             =   1410
         Width           =   4845
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3000
         TabIndex        =   26
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
         TabIndex        =   24
         Top             =   375
         Width           =   4845
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2550
         Picture         =   "I_SetPla.frx":479D
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2550
         Picture         =   "I_SetPla.frx":4AA7
         Top             =   600
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2550
         Picture         =   "I_SetPla.frx":4DB1
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha(mm/aa)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   435
         TabIndex        =   22
         Top             =   1110
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Segmento"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   435
         TabIndex        =   15
         Top             =   450
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Punto Venta"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   435
         TabIndex        =   10
         Top             =   1455
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Casino"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   435
         TabIndex        =   9
         Top             =   795
         Width           =   555
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3045
         TabIndex        =   25
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
         TabIndex        =   27
         Top             =   765
         Width           =   4845
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3045
         TabIndex        =   29
         Top             =   1455
         Width           =   4845
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   615
      Left            =   2400
      TabIndex        =   12
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
      SpreadDesigner  =   "I_SetPla.frx":50BB
      StartingColNumber=   6
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_SetPla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim i As Integer, iselecc As Integer, Msgtitulo As String

Private Sub Check1_Click()
On Error GoTo Test
Dim RS1 As New ADODB.Recordset
If Check1.Value = 1 Then
   Text1.Enabled = True
   RS1.Open "select Parametro_Glosa txt from Sdx_Parametro where Parametro_Num=9999", vg_db, adOpenDynamic
   If Not RS1.EOF Then
        Text1.Text = RS1!txt
   End If
   RS1.Close: Set RS1 = Nothing
Else
   Text1.Text = "": Text1.Enabled = False
End If
Exit Sub

Test:
If (Err.Number = 94) Then Resume Next Else MsgBox Err.Description
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   Combo1.Enabled = True
Else
   Combo1.ListIndex = -1
   Combo1.Enabled = False
End If
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
On Error GoTo Man_Error
Me.Height = 7575
Me.Width = 8175
Me.HelpContextID = vg_OpcM
fg_centra Me
Msgtitulo = "Planificación Minutas"
Toolbar1.ImageList = partida.IL1
Toolbar1.Buttons.Clear
'Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar ": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set btnX = Toolbar1.Buttons.Add(, "Vista Previa", , tbrDefault, "Vista Previa"): btnX.Visible = True: btnX.ToolTipText = "Vista Previa ": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): btnX.Visible = True: btnX.ToolTipText = "Historico Minutas"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
fpDateTime1.Text = Format(Date, "mm/yyyy")
vaSpread1.MaxRows = 0: vaSpread2.MaxRows = 0
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub

Private Sub Form_Unload(Cancel As Integer)
'registrar Log sistema salir
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "")
End Sub

Private Sub fpDateTime1_Change()
MoverDatosVector
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Select Case Index
Case 0
    Set RS = vg_db.Execute("min_s_segmento 4, " & Val(fpLongInteger1(0).Value) & ", ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!Unit_Dfnd_Desc)
    RS.Close: Set RS = Nothing
    fpText.Text = "": fpayuda(1).Caption = "": fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": vaSpread1.MaxRows = 0
Case 1
    Set RS = vg_db.Execute("min_s_puntoventa 7, " & Val(fpLongInteger1(1).Value) & ", ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "": Exit Sub
    fpayuda(2).Caption = Trim(RS!Sls_Locn_Name)
    RS.Close: Set RS = Nothing
    MoverDatosVector
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    If Index = 0 Then image1_Click 0
    If Index = 1 Then image1_Click 2
End Select
End Sub

Private Sub fpText_Change()
Set RS = vg_db.Execute("min_s_casino 6, " & Val(fpLongInteger1(0).Value) & ", '" & "00000" & LimpiaDato(Trim(fpText.Text)) & "'")
If RS.EOF Then fpayuda(1).Caption = "": RS.Close: Set RS = Nothing: Exit Sub
fpayuda(1).Caption = Trim(RS!Nombre_Casino)
RS.Close: Set RS = Nothing
fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": vaSpread1.MaxRows = 0
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    image1_Click 1
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    Option1(0).Value = True
    Option1(1).Value = False
    Image1(3).Enabled = False
Case 1
    Option1(0).Value = False
    Option1(1).Value = True
    Image1(3).Enabled = True
End Select
End Sub

Private Sub Option1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    image1_Click 2
End Select
End Sub

Private Sub image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_codigo = ""
    vg_left = fpayuda(0).Left
    B_TabEst.LlenaDatos "Segemento", 0, 2
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Then Exit Sub
    fpLongInteger1(0).Value = vg_codigo
    fpText.Text = "": fpayuda(1).Caption = "": fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": vaSpread1.MaxRows = 0
Case 1
    vg_codigo = ""
    vg_left = fpayuda(1).Left
    B_TabEst.LlenaDatos "Casino", Val(fpLongInteger1(0).Value), 1
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Then Exit Sub
    fpText.Text = vg_codigo
    fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": vaSpread1.MaxRows = 0
Case 2
    vg_codigo = ""
    vg_left = fpayuda(2).Left
    B_TabEst.LlenaDatos "Punto Venta", 0, 3
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Then Exit Sub
    fpLongInteger1(1).Value = vg_codigo
Case 3
    B_MServi.LlenaDatosSer Me, Trim(fpText.Text), Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Format(fpDateTime1.Text, "yyyy"), Format(fpDateTime1.Text, "mm")
    B_MServi.Show 1
    Me.Refresh
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If Combo1.ListIndex = -1 And Check2.Value = 1 Then MsgBox "Seleccione Calcular Costo x", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Option1(0).Value = True Then ActivarGrillaTodos
    Set RS = vg_db.Execute("min_s_casino 6, " & Val(fpLongInteger1(0).Value) & ", '" & "00000" & LimpiaDato(Trim(fpText.Text)) & "'")
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Casino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Set RS = vg_db.Execute("min_s_puntoventa 7, " & Val(fpLongInteger1(1).Value) & ", ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Punto Venta", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If fpDateTime1.Text = "" Then MsgBox "Seleccione Fecha", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If fpDateTime1.Text = "" Then Exit Sub
    ValidarOpcion
Case 3
    Set RS = vg_db.Execute("min_s_casino 6, " & Val(fpLongInteger1(0).Value) & ", '" & "00000" & LimpiaDato(Trim(fpText.Text)) & "'")
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Casino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    vg_codigo = ""
    B_HistPm.LlenarDatos "00000" & LimpiaDato(Trim(fpText.Text)), Val(fpLongInteger1(0).Value)
    B_HistPm.Show 1
    If Trim(vg_codigo) = "" Then Exit Sub
    Dim StrImp As String, StrImpb As String
    StrImp = Trim(vg_codigo): i = 1
    Do While InStr(StrImp, ";") <> 0
       If i = 1 Then
          fpLongInteger1(1).Value = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
       ElseIf i = 3 Then
          fpDateTime1.Text = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
       End If
       StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), ""): i = i + 1
    Loop
    Me.Refresh
Case 5
    Me.Hide
    Unload Me
End Select
End Sub

Sub ValidarOpcion()
On Error GoTo Man_Error
fg_carga ""
iselecc = 0: vg_opimp = 0
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1
    If vaSpread1.Text = "1" Then iselecc = 1: Exit For
Next i
If iselecc = 0 Then fg_descarga: MsgBox "Servicio debe ser informado", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
Set RS = vg_db.Execute("min_s_minutas 3, 0, '" & "00000" & LimpiaDato(Trim(fpText.Text)) & "', " & Val(fpLongInteger1(0).Value) & ", 0, " & Val(fpLongInteger1(1).Value) & ", 0, '" & Format(fpDateTime1.Text, "yyyy") & "', '" & Format(fpDateTime1.Text, "mm") & "'")
If Not RS.EOF Then
   Do While Not RS.EOF
      Exit Do
   Loop
   RS.Close: Set RS = Nothing
   vaSpread2.MaxRows = 0: vaSpread2.MaxCols = 0
   Toolbar1.Enabled = False
   Frame1(0).Enabled = False
   vg_opimp = 0: vg_opimp = 9999
   I_SetMinuta Me
   vg_opimp = 0
   Toolbar1.Enabled = True
   Frame1(0).Enabled = True
Else
   RS.Close: Set RS = Nothing: fg_descarga
   MsgBox "No Existe Información Para Procesar", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
End If
vg_opimp = 0
'registrar Log sistema confirma
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Aceptar"), Me.HelpContextID, "", "")
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub

Sub MoverDatosVector()
If LimpiaDato(Trim(fpText.Text)) <> "" And Format(fpDateTime1.Text, "yyyy") <> "" And Format(fpDateTime1.Text, "mm") <> "" And Val(fpLongInteger1(1).Value) > 0 And Val(fpLongInteger1(1).Value) > 0 Then
   fg_carga ""
   Set RS = vg_db.Execute("min_s_minutas 3, 0,  '" & "00000" & LimpiaDato(Trim(fpText.Text)) & "', " & Val(fpLongInteger1(0).Value) & ", 0, " & Val(fpLongInteger1(1).Value) & ", 0, '" & Format(fpDateTime1.Text, "yyyy") & "', '" & Format(fpDateTime1.Text, "mm") & "'")
   vaSpread1.MaxRows = 0
   If Not RS.EOF Then
      Do While Not RS.EOF
         vaSpread1.MaxRows = vaSpread1.MaxRows + 1
         vaSpread1.Row = vaSpread1.MaxRows
         vaSpread1.Col = 2
         vaSpread1.Value = RS!Serv_No
         vaSpread1.Col = 3
         vaSpread1.Value = Trim(RS!Serv_Name)
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing: fg_descarga
End If
End Sub

Sub ActivarGrillaTodos()
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1
    vaSpread1.CellType = 10
    vaSpread1.TypeCheckText = ""
    vaSpread1.TypeCheckCenter = True
    vaSpread1.Text = "1" ' checked
Next i
End Sub
