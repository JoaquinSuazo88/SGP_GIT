VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form C_ConIngMinBlo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumo Ingrediente Minuta Bloque"
   ClientHeight    =   8910
   ClientLeft      =   3630
   ClientTop       =   2220
   ClientWidth     =   14445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   14445
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
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
      Left            =   12960
      TabIndex        =   13
      Top             =   8430
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar XLS"
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
      Left            =   11400
      TabIndex        =   12
      Top             =   8430
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ingrediente Tabla Gramaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   6975
      Width           =   14175
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
         Height          =   225
         Index           =   1
         Left            =   3405
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Uno"
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
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1395
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   0
         Left            =   1245
         TabIndex        =   10
         Top             =   750
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483637
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
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
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3015
         TabIndex        =   11
         Top             =   750
         Width           =   6825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ingrediente"
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
         Left            =   165
         TabIndex        =   19
         Top             =   810
         Width           =   975
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   480
         Index           =   1
         Left            =   2520
         Picture         =   "C_ConIngMinBlo.frx":0000
         Top             =   660
         Width           =   480
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3060
         TabIndex        =   20
         Top             =   780
         Width           =   6810
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Regimen - Servicios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   14175
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   9720
         TabIndex        =   35
         Top             =   3840
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   6840
         TabIndex        =   34
         Top             =   3840
         Width           =   2835
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   5760
         TabIndex        =   33
         Top             =   3840
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   32
         Top             =   3840
         Width           =   2835
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   31
         Top             =   3840
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   10920
         TabIndex        =   30
         Top             =   3840
         Width           =   2715
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   13695
         _Version        =   393216
         _ExtentX        =   24156
         _ExtentY        =   5741
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
         MaxCols         =   8
         SpreadDesigner  =   "C_ConIngMinBlo.frx":030A
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2280
      Left            =   120
      TabIndex        =   14
      Top             =   15
      Width           =   14175
      Begin VB.CheckBox Check1 
         Caption         =   "Sólo Permitir Nulos"
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
         Left            =   11295
         TabIndex        =   29
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Frame Frame4 
         Height          =   1485
         Left            =   105
         TabIndex        =   21
         Top             =   210
         Width           =   13845
         Begin VB.OptionButton Option2 
            Caption         =   "Org. de Compra x Ceco"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   11175
            TabIndex        =   28
            Top             =   960
            Width           =   2400
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Org. de Compras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   11175
            TabIndex        =   0
            Top             =   210
            Value           =   -1  'True
            Width           =   1800
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Centro de Costo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   11175
            TabIndex        =   1
            Top             =   570
            Width           =   1800
         End
         Begin EditLib.fpText fpText 
            Height          =   315
            Left            =   1680
            TabIndex        =   3
            Top             =   930
            Width           =   1275
            _Version        =   196608
            _ExtentX        =   2249
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin EditLib.fpText fpText2 
            Height          =   315
            Left            =   1680
            TabIndex        =   2
            Top             =   390
            Width           =   1275
            _Version        =   196608
            _ExtentX        =   2249
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   3525
            TabIndex        =   26
            Top             =   390
            Width           =   6015
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   2940
            Picture         =   "C_ConIngMinBlo.frx":1DA2
            Top             =   285
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Org. de Compras"
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
            TabIndex        =   25
            Top             =   435
            Width           =   1425
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
            Left            =   210
            TabIndex        =   23
            Top             =   975
            Width           =   735
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   3555
            TabIndex        =   22
            Top             =   915
            Width           =   6000
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   0
            Left            =   2940
            Picture         =   "C_ConIngMinBlo.frx":20AC
            Top             =   825
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
            Left            =   3600
            TabIndex        =   24
            Top             =   960
            Width           =   6000
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   5
            Left            =   3540
            TabIndex        =   27
            Top             =   405
            Width           =   6045
         End
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   1395
         TabIndex        =   4
         Top             =   1770
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Text            =   "01/09/2013"
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
      Begin EditLib.fpDateTime FpFecHasta 
         Height          =   315
         Left            =   7140
         TabIndex        =   5
         Top             =   1770
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Text            =   "28/09/2013"
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8520
         Top             =   2130
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
               Picture         =   "C_ConIngMinBlo.frx":23B6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   13560
         TabIndex        =   6
         Top             =   1770
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha hasta"
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
         Left            =   5835
         TabIndex        =   16
         Top             =   1860
         Width           =   1065
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
         TabIndex        =   15
         Top             =   1860
         Width           =   1110
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   165
      Index           =   0
      Left            =   120
      TabIndex        =   36
      Top             =   8520
      Visible         =   0   'False
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "C_ConIngMinBlo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim MsgTitulo As String
Public lc_Aux As String
Dim Est As Boolean

Private Sub Command1_Click()

On Local Error GoTo Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Est As Boolean
Dim i As Long
Dim j As Long
Dim MyBuffer As String
Dim codCeco As String
Dim CodRegimen As Long
Dim CodServicio As Long

Dim NomArchivoExcel As String
Dim Extension       As String
'ValidarDatos
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object

Dim wvarTipoReporte As Integer

'If Not ValidarDatos Then Exit Sub

fg_carga ""
Screen.MousePointer = 11
    
'-------> Validar fecha desde - hasta
If CDate(FpFecDesde.text) > CDate(FpFecHasta.text) Then
   
   Call MsgBox("Fecha Desde No Puede Ser Mayor a Fecha Hasta", vbInformation, MsgTitulo)
   Let FpFecDesde.text = Format(Now, "dd/mm/yyyy")
   Call FpFecDesde.SetFocus
   Exit Sub

End If
    
If CDate(FpFecHasta.text) < CDate(FpFecDesde.text) Then
   
   Call MsgBox("Fecha Hasta No Puede Ser Mayor a Fecha Desde", vbInformation, MsgTitulo)
   Let FpFecHasta.text = Format(Now, "dd/mm/yyyy")
   Call FpFecHasta.SetFocus
   Exit Sub

End If



If Option2(2).Value = True Then
  
  Dim Conta As Integer
  Dim seleccion As Integer
  Conta = 0
  
  For i = 1 To vaSpread1.MaxRows
          
      vaSpread1.Row = i
      
      vaSpread1.Col = 1 'Seleccion
      seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
      
      If seleccion = 1 Then
        
        Conta = Conta + 1
      
      End If
      
  Next i
    
  If Conta = 0 Then
        
     MsgBox "Debe haber selecionado al menos un dato de la grilla.", vbExclamation + vbOKOnly, Me.Caption
     i = 0
     fg_descarga
     Exit Sub
    
  End If
  
End If


If Option2(1).Value = True Then
   '-------> Validar Ceco
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT isnull(a.cli_codigo, 0) as cli_codigo, isnull(a.cli_nombre, '') as cli_nombre, isnull(a.cli_tipominuta,0) as cli_tipominuta " & _
            "FROM b_clientes as a WITH (NOLOCK) " & _
            "inner join b_tipominuta as b with (nolock) on b.tip_codigo = a.cli_tipominuta " & _
            "                                          and b.activo = '1' " & _
            "WHERE a.cli_codigo = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
            "AND   a.cli_tipo   = 0 ")
'             "AND   cli_tipominuta in ('3')")
    
    If RS.EOF Then
       
       RS.Close
       Set RS = Nothing
       MsgBox "No existe ceco...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       fpayuda(0).Caption = ""
       vaSpread1.MaxRows = 0
       Exit Sub
    
    Else
       
       If RS!cli_tipominuta <> "3" Then
          
          MsgBox "Ceco debe ser Simap...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
          fpayuda(0).Caption = ""
          vaSpread1.MaxRows = 0
          RS.Close
          Set RS = Nothing
          Exit Sub
       
       End If
    
    End If
    RS.Close
    Set RS = Nothing
    
    '-------> validar minuta bloque
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("SELECT  DISTINCT cbm.min_cecori " & _
                            "FROM    dbo.cas_b_minuta AS cbm WITH (NOLOCK) " & _
                            "INNER JOIN dbo.cas_b_minutadet AS cbm2 WITH ( NOLOCK ) ON cbm2.mid_cecori = cbm.min_cecori " & _
                                                                                  " AND cbm2.mid_codigo = cbm.min_codigo " & _
                            "WHERE   cbm.min_cecori = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                            "AND cbm.min_fecmin >= " & Format(FpFecDesde.text, "yyyymmdd") & " " & _
                            "AND cbm.min_fecmin <= " & Format(FpFecHasta.text, "yyyymmdd") & "")
    If RS.EOF Then
       
       MsgBox "No existe Minuta...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       fpayuda(0).Caption = ""
       vaSpread1.MaxRows = 0
       RS.Close
       Set RS = Nothing
       Exit Sub
    
    End If
    RS.Close
    Set RS = Nothing
    
    '-------> validar seleción regimen de la lista
    Est = False
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        
        If vaSpread1.text = "1" Then
           
           Est = True
        
        End If
    
    Next i
    
    If Not Est Then
       
       If vaSpread1.MaxRows > 0 Then
          
          MsgBox "Regimen debe ser informado de la lista", vbExclamation + vbOKOnly, MsgTitulo
       
       Else
          
          MsgBox "Para la visualizar lista de regimen, debe seleccionar icono de proceso", vbExclamation + vbOKOnly, MsgTitulo
       
       End If
       
       Exit Sub
    
    End If

ElseIf Option1(0).Value = True Then
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("select distinct id_orgcompra from i_org_ceco WITH (NOLOCK) where id_orgcompra = '" & LimpiaDato(Trim(fpText2.text)) & "' and ISNULL(borrado,'') <> 'X'")
    If RS.EOF Then
        
        RS.Close
        Set RS = Nothing
        MsgBox "No existe organización de compras...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        fpayuda(4).Caption = ""
        Exit Sub
    
    End If

End If

'-------> Validar ingrediente
If Option1(0).Value = True Then
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT ing_codigo, ing_nombre " & _
               "FROM b_ingrediente WITH (NOLOCK) " & _
               "WHERE ing_codigo = '" & LimpiaDato(Trim(fpText1(0).text)) & "' " & _
               "AND   ing_indppr   = 1 " & _
               "AND   ing_activo = '1'")
   
   If RS.EOF Then
      
      RS.Close
      Set RS = Nothing
      fpayuda(1).Caption = ""
      MsgBox "No existe Ingrediente", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      Exit Sub
   
   End If
   
   fpayuda(1).Caption = Trim(RS!ing_nombre)
   fpText1(0).text = RS!ing_codigo
   RS.Close
   Set RS = Nothing

End If

If Option2(1).Value = True Then
    
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<RegSer>"

    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        
        If Trim(vaSpread1.text) = "1" Then
           
           Let MyBuffer = MyBuffer & " <DetCodSer"
           vaSpread1.Col = 2
           codCeco = vaSpread1.text
           vaSpread1.Col = 4
           CodRegimen = vaSpread1.text
           vaSpread1.Col = 6
           CodServicio = vaSpread1.text
           Let MyBuffer = MyBuffer & " CodCeco = " & Chr(34) & codCeco & Chr(34)
           Let MyBuffer = MyBuffer & " CodRegimen = " & Chr(34) & CodRegimen & Chr(34)
           Let MyBuffer = MyBuffer & " CodServicio = " & Chr(34) & CodServicio & Chr(34)
           Let MyBuffer = MyBuffer & "/>"
        
        End If
    
    Next i
    
    Let MyBuffer = MyBuffer & "</RegSer>"

End If

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim RowSheet As Long

'-------> Traer consumo y exportar excel

Sql = ""

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Option2(1).Value = True Then
    '--> Centro de Costo
    wvarTipoReporte = 0
    
    Sql = Sql & LimpiaDato(Trim(fpText.text))
    Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
    Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
    Sql = Sql & "," & IIf(Option1(0).Value = True, LimpiaDato(Trim(fpText1(0).text)), 0)
    Sql = Sql & "," & Check1.Value
    Set RS = vg_db.Execute("sgpadm_Sel_ConsumoIngMinBloque_Ceco_V04 '" & MyBuffer & "', " & Sql & "")

ElseIf Option2(0).Value = True Then
    '--> Org. de Compras
    wvarTipoReporte = 1
    
    Sql = Sql & "'" & LimpiaDato(Trim(fpText2.text)) & "'"
    Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
    Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
    Sql = Sql & "," & IIf(Option1(0).Value = True, LimpiaDato(Trim(fpText1(0).text)), 0)
    Sql = Sql & "," & Check1.Value
    Set RS = vg_db.Execute("sgpadm_Sel_ConsumoIngMinBloque_OrgCompra_V04 " & Sql & "")
    
ElseIf Option2(2).Value = True Then
'--> Org. de Compras x Ceco
'    wvarTipoReporte = 2
'
'    Sql = Sql & "'" & LimpiaDato(Trim(fpText2.text)) & "'"
'    Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
'    Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
'    Sql = Sql & "," & IIf(Option1(0).Value = True, LimpiaDato(Trim(fpText1(0).text)), 0)
'    Sql = Sql & "," & Check1.Value
'    Set RS = vg_db.Execute("sgpadm_Sel_ConsumoIngMinBloque_OrgCompraCeco_v02 '" & MyBuffer & "', " & Sql & "")

End If

If Option2(1).Value = True Or Option2(0).Value = True Then

    If Not RS.EOF Then
      
       If RS.RecordCount > 1020000 Then
          
          RS.Close
          Set RS = Nothing
          
          MsgBox "El resultado sobrepasa máximo de fila en Excel, Debe seleccionar menos datos.", vbCritical
          Exit Sub
       
       End If
      
    End If

End If


Dim Ceco As String
Dim NomCeco As String
Dim auxceco As String

Ceco = ""
NomCeco = ""
auxceco = ""
Bar1(0).Value = 0
Bar1(0).Visible = True

If Option2(2).Value = True Then

    For i = 1 To vaSpread1.MaxRows
        
        DoEvents
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
        
        'If vaSpread1.text = "1" Then
        If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
           
           vaSpread1.SetActiveCell 2, vaSpread1.Row
           
           vaSpread1.Col = 2
           Ceco = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
           
           vaSpread1.Col = 3
           NomCeco = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
           
           
           If Ceco <> auxceco Then
           
              If Trim(auxceco) <> "" Then
              
              Let MyBuffer = MyBuffer & "</RegSer>"
              'MsgBox "Genera Excel " & Ceco & " - " & NomCeco
              
                wvarTipoReporte = 2
                Sql = ""
'                Sql = Sql & "'" & LimpiaDato(Trim(fpText2.text)) & "'"
'                Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
'                Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
'                Sql = Sql & "," & IIf(Option1(0).Value = True, LimpiaDato(Trim(fpText1(0).text)), 0)
'                Sql = Sql & "," & Check1.Value
'                Set RS = vg_db.Execute("sgpadm_Sel_ConsumoIngMinBloque_OrgCompraCeco_v02 '" & MyBuffer & "', " & Sql & "")

                Sql = Sql & LimpiaDato(Trim(auxceco))
                Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
                Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
                Sql = Sql & "," & IIf(Option1(0).Value = True, LimpiaDato(Trim(fpText1(0).text)), 0)
                Sql = Sql & "," & Check1.Value
                Set RS = vg_db.Execute("sgpadm_Sel_ConsumoIngMinBloque_Ceco_V04 '" & MyBuffer & "', " & Sql & "")


                If Not RS.EOF Then
              
                    'Genera EXCEL
                    'Format(FpFecDesde.text, "yyyymmdd") & " " & Format(FpFecHasta.text, "yyyymmdd") & " " & Replace(Date, "/", "") & " " & Replace(Time, ":", "") & ".xlsx"
                    CD.FileName = ""
                    '-------> Guardar nombre archivo excel
                    NomArchivoExcel = dir_trabajo & "ExcelMinutaSGP" & "\" & "FiltroIngrediente " & LimpiaDato(Trim(fpText2.text)) & "-" & auxceco & " " & Replace(Date, "/", "-") & " " & Format(Time, "hhmmss") & ".xlsx"
                    CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
                    CD.Filter = "Todos los archivos *.xls,*.xlsx"
                    On Error Resume Next
                    'CD.ShowSave
                    CD.FileName = NomArchivoExcel
                               
                    '-------> JPAZ Permite controlar Boton Cancelar
                    If Err.Number = 32755 Then
                       
                       MsgBox "Proceso cancelado"
                       Exit Sub
                    
                    End If
                                
                    If CD.FileName = "" Then
                       
                       MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
                       Exit Sub
                    
                    Else
                       
                       Extension = ""
                       Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
                       
                       If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
                          MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
                          Exit Sub
                       End If
                       
                       NomArchivoExcel = CD.FileName
                    
                    End If
                              
                    fg_carga ""
                      
                    '-------> Create an instance of Excel and add a workbook
                    Set xlApp = CreateObject("Excel.Application")
                    Set xlWb = xlApp.Workbooks.Add
                    Set xlWs = xlWb.Worksheets("Hoja1")
                      
                    '-------> Display Excel and give user control of Excel's lifetime
                    xlApp.UserControl = True
                        
                    '-------> Check version of Excel
                    Call encabezado(RS, xlWs, wvarTipoReporte, auxceco)
                              
                    xlWs.Cells(7, 1).CopyFromRecordset RS
                    
                    '-------> Auto-fit the column widths and row heights
                    xlApp.Selection.CurrentRegion.Columns.AutoFit
                    xlApp.Selection.CurrentRegion.Rows.AutoFit
                        
                    'xlApp.Columns("A:A").Select
                    'xlApp.Selection.Delete Shift:=xlToLeft
                      
                    xlWb.Close True, NomArchivoExcel
                    
                    Dim XL As New excel.Application 'Crea el objeto excel
                    'XL.Workbooks.Open NomArchivoExcel, , False 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
                    XL.Visible = False
                    'XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
                        
                    '-------> Close ADO objects
                    RS.Close
                    Set RS = Nothing
                        
                    ' -- Cerrar Excel
                    xlApp.Quit
                    '-------> Release Excel references
                    Set xlWs = Nothing
                    Set xlWb = Nothing
                    Set xlApp = Nothing
                  
                
                
                End If
              
                 
              End If
              
            Let MyBuffer = ""
            Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
            Let MyBuffer = MyBuffer & "<RegSer>"
    
    
              
              
           End If
           
           
            Let MyBuffer = MyBuffer & " <DetCodSer"
            vaSpread1.Col = 2
            codCeco = vaSpread1.text
            vaSpread1.Col = 4
            CodRegimen = vaSpread1.text
            vaSpread1.Col = 6
            CodServicio = vaSpread1.text
            Let MyBuffer = MyBuffer & " CodCeco = " & Chr(34) & codCeco & Chr(34)
            Let MyBuffer = MyBuffer & " CodRegimen = " & Chr(34) & CodRegimen & Chr(34)
            Let MyBuffer = MyBuffer & " CodServicio = " & Chr(34) & CodServicio & Chr(34)
            Let MyBuffer = MyBuffer & "/>"
           
        End If
        
        auxceco = Ceco
        
    Next i

End If







If Trim(auxceco) <> "" Then

    If Option2(2).Value = True Then
        wvarTipoReporte = 2
        Let MyBuffer = MyBuffer & "</RegSer>"
        Sql = ""
'        Sql = Sql & "'" & LimpiaDato(Trim(fpText2.text)) & "'"
'        Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
'        Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
'        Sql = Sql & "," & IIf(Option1(0).Value = True, LimpiaDato(Trim(fpText1(0).text)), 0)
'        Sql = Sql & "," & Check1.Value
'        Set RS = vg_db.Execute("sgpadm_Sel_ConsumoIngMinBloque_OrgCompraCeco_v02 '" & MyBuffer & "', " & Sql & "")
        
        Sql = Sql & LimpiaDato(Trim(auxceco))
        Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
        Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
        Sql = Sql & "," & IIf(Option1(0).Value = True, LimpiaDato(Trim(fpText1(0).text)), 0)
        Sql = Sql & "," & Check1.Value
        Set RS = vg_db.Execute("sgpadm_Sel_ConsumoIngMinBloque_Ceco_V04 '" & MyBuffer & "', " & Sql & "")
        

        
    End If
End If





If Not RS.EOF Then
    
    '-------> Guardar nombre archivo excel
    NomArchivoExcel = ""
    CD.FileName = ""
    '-------> Guardar nombre archivo excel
    
    If Option2(0).Value = True Then
        '--> Org. de Compra
        NomArchivoExcel = dir_trabajo & "ExcelMinutaSGP" & "\" & "FiltroIngrediente " & LimpiaDato(Trim(fpText2.text)) & " " & Replace(Date, "/", "-") & " " & Format(Time, "hhmmss") & ".xlsx"
    ElseIf Option2(1).Value = True Then
        '--> Centro de Costo
        NomArchivoExcel = dir_trabajo & "ExcelMinutaSGP" & "\" & "FiltroIngrediente " & LimpiaDato(Trim(fpText.text)) & " " & Replace(Date, "/", "-") & " " & Format(Time, "hhmmss") & ".xlsx"
    ElseIf Option2(2).Value = True Then
        '--> Org. de Compra x Ceco
        NomArchivoExcel = dir_trabajo & "ExcelMinutaSGP" & "\" & "FiltroIngrediente " & LimpiaDato(Trim(fpText2.text)) & "-" & auxceco & " " & Replace(Date, "/", "-") & " " & Format(Time, "hhmmss") & ".xlsx"
    End If
    
    
    CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
    CD.Filter = "Todos los archivos *.xls,*.xlsx"
    On Error Resume Next
    'CD.ShowSave
    CD.FileName = NomArchivoExcel
    
               
    '-------> JPAZ Permite controlar Boton Cancelar
    If Err.Number = 32755 Then
       
       MsgBox "Proceso cancelado"
       Exit Sub
    
    End If
                
    If CD.FileName = "" Then
       
       MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
       Exit Sub
    
    Else
       
       Extension = ""
       Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
       
       If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
          MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
          Exit Sub
       End If
       
       NomArchivoExcel = CD.FileName
    
    End If
              
    fg_carga ""
      
    '-------> Create an instance of Excel and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets("Hoja1")
      
    '-------> Display Excel and give user control of Excel's lifetime
    xlApp.UserControl = True
        
    '-------> Check version of Excel
    Call encabezado(RS, xlWs, wvarTipoReporte, auxceco)
              
    xlWs.Cells(7, 1).CopyFromRecordset RS
    
    '-------> Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
        
    'xlApp.Columns("A:A").Select
    'xlApp.Selection.Delete Shift:=xlToLeft
      
    xlWb.Close True, NomArchivoExcel
    
    'Dim XL As New Excel.Application 'Crea el objeto excel
    'XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
    XL.Visible = False
    'XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
        
    '-------> Close ADO objects
    RS.Close
    Set RS = Nothing
        
    ' -- Cerrar Excel
    xlApp.Quit
    '-------> Release Excel references
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
      
    fg_descarga
    'MsgBox "Proceso Finalizado", vbInformation, Me.Caption
    Bar1(0).Value = 0
    Bar1(0).Visible = False
    
    If MsgBox("Proceso Finalizado." & VgLinea & "¿Desea ingresar al directorio de trabajo?", vbQuestion + vbYesNo, MsgTitulo) = vbYes Then
           Shell "explorer " & dir_trabajo & "ExcelMinutaSGP" & "\", vbNormalFocus
           Exit Sub
    End If

Else

    MsgBox "Proceso finalizado.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fg_descarga
    
End If



Exit Sub

Error:
    'Resume
    fg_descarga
    oExcel.DisplayAlerts = False
    oExcel.Quit
    oExcel.DisplayAlerts = True
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Exit Sub

End Sub

Private Sub Command2_Click()
    
    Me.Hide 'Salir
    Unload Me

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

fg_centra Me
MsgTitulo = "Consumo Ingrediente Minuta Bloque"
FpFecHasta.text = Format(Date, "dd/mm/yyyy")
FpFecDesde.text = Format(Date, "dd/mm/yyyy")
vaSpread1.MaxRows = 0
Est = True

End Sub

Private Sub FpFecDesde_Change()

If IsDate(FpFecDesde.text) = False Then Exit Sub

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub FpFecHasta_Change()

If IsDate(FpFecHasta.text) = False Then Exit Sub

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub fpText_Change()

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_cliente_V02 47, '" & LimpiaDato(fpText.text) & "', ''")

If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fpayuda(0).Caption = ""
   vaSpread1.MaxRows = 0
   Exit Sub

End If
fpayuda(0).Caption = Trim(RS!Cli_nombre)
fpText.text = RS!Cli_codigo
vaSpread1.MaxRows = 0
RS.Close
Set RS = Nothing

End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub fpText1_Change(Index As Integer)

Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT ing_codigo, ing_nombre " & _
            "FROM b_ingrediente WITH (NOLOCK) " & _
            "WHERE ing_codigo = '" & LimpiaDato(Trim(fpText1(0).text)) & "' " & _
            "AND   ing_indppr   = 1 " & _
            "AND   ing_activo = '1'")

If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fpayuda(1).Caption = ""
   Exit Sub

End If

fpayuda(1).Caption = Trim(RS!ing_nombre)
fpText1(0).text = RS!ing_codigo
RS.Close
Set RS = Nothing

End Sub

Private Sub fpText2_Change()

If Not Est Then Exit Sub
Dim RS As New ADODB.Recordset
Dim Sql As String

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = ""
Sql = "sgpadm_Sel_OrgCompras_V02 "
Sql = Sql & " '" & LimpiaDato(Trim(fpText2.text)) & "' "
Set RS = vg_db.Execute(Sql)
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fpayuda(4).Caption = ""
   Exit Sub

End If

fpayuda(4).Caption = RS!ID_Orgcompra
fpText2.text = RS!ID_Orgcompra
RS.Close
Set RS = Nothing

End Sub

Private Sub Image1_Click(Index As Integer)

Select Case Index

Case 0
    
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "Clientesimap"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.text = vg_codigo: fpayuda(0).Caption = vg_nombre
    If Me.Visible Then FpFecDesde.SetFocus

Case 1
    
    vg_left = fpayuda(1).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "INgrediente", "IngReal"
    B_TabEst.Show 1
    Me.Refresh
    Screen.MousePointer = 0
    If vg_codigo = "" Then Exit Sub
    fpText1(0).text = vg_codigo
    fpayuda(1).Caption = vg_nombre

Case 4
    
    vg_left = fpayuda(4).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_centrologisticoceco_sap", "", "Organización de Compras", "Celo"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText2.text = vg_codigo
'    fpayuda(4).Caption = Trim(vg_nombre)

End Select

End Sub

Private Sub Option1_Click(Index As Integer)

fpText1(0).Enabled = IIf(Index = 0, True, False)
Image1(1).Enabled = IIf(Index = 0, True, False)
'-------> limpiar variable

If Index = 1 Then
   
   fpText1(0).text = ""
   fpayuda(1).Caption = ""

End If

End Sub

Private Sub Option2_Click(Index As Integer)

Est = False
fpText2.text = ""
Est = True
fpayuda(4).Caption = ""
fpText.text = ""
fpayuda(0).Caption = ""
vaSpread1.MaxRows = 0

Select Case Index

Case 0
    
    fpText.Enabled = False
    Image1(0).Enabled = False
    fpText2.Enabled = True
    Image1(4).Enabled = True

Case 1
    
    fpText2.Enabled = False
    Image1(4).Enabled = False
    fpText.Enabled = True
    Image1(0).Enabled = True
    
Case 2

    fpText.Enabled = False
    Image1(0).Enabled = False
    fpText2.Enabled = True
    Image1(4).Enabled = True


End Select

End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Local Error GoTo Error

Dim RS As New ADODB.Recordset
Dim Sql As String
    
    Select Case Button.Index
    
    Case 1 'Mostrar datos en la grilla
        
        If Option2(1).Value = True Or Option2(2).Value = True Then
            
            If Option2(1).Value = True Then
                If RS.State = 1 Then RS.Close
                RS.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
              
                Set RS = vg_db.Execute("SELECT isnull(a.cli_codigo,0) as cli_codigo, isnull(a.cli_nombre,'') as cli_nombre, isnull(a.cli_tipominuta,0) as cli_tipominuta " & _
                         "FROM b_clientes as a WITH (NOLOCK) " & _
                         "inner join b_tipominuta as b with (nolock) on b.tip_codigo = a.cli_tipominuta " & _
                         "                                          and b.activo = '1' " & _
                         "WHERE a.cli_codigo = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                         "AND   a.cli_tipo   = 0 ")
    '                    "AND   cli_tipominuta in ('3')")
                
                If RS.EOF Then
                   
                   RS.Close
                   Set RS = Nothing
                   MsgBox "No existe ceco...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
                   fpayuda(0).Caption = ""
                   vaSpread1.MaxRows = 0
                   Exit Sub
                
                Else
                   
                   If RS!cli_tipominuta <> "3" Then
                      
                      MsgBox "Ceco debe ser Simap...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
                      fpayuda(0).Caption = ""
                      vaSpread1.MaxRows = 0
                      RS.Close
                      Set RS = Nothing
                      Exit Sub
                   
                   End If
                
                End If
                RS.Close
                Set RS = Nothing
            
            End If
        
            If CDate(FpFecDesde.text) > CDate(FpFecHasta.text) Then
                
                Call MsgBox("Fecha Desde No Puede Ser Mayor a Fecha Hasta", vbInformation, Me.Caption)
                Let FpFecDesde.text = Format(Now, "dd/mm/yyyy")
                Call FpFecDesde.SetFocus
                Exit Sub
            
            End If
    
            If CDate(FpFecHasta.text) < CDate(FpFecDesde.text) Then
                
                Call MsgBox("Fecha Hasta No Puede Ser Mayor a Fecha Desde", vbInformation, Me.Caption)
                Let FpFecHasta.text = Format(Now, "dd/mm/yyyy")
                Call FpFecHasta.SetFocus
                Exit Sub
            
            End If
        
            vaSpread1.Visible = False
            vaSpread1.MaxRows = 0
            Sql = ""
            Sql = Sql & "'" & LimpiaDato(Trim(fpText.text)) & "'"
            Sql = Sql & "," & Format(FpFecDesde.text, "yyyymmdd")
            Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
            Sql = Sql & "," & "'" & LimpiaDato(Trim(fpText2.text)) & "'"
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute("sgpadm_Sel_RegSerMinutaBloque_v02  " & Sql & "")
            Do While Not RS.EOF
               
               vaSpread1.MaxRows = vaSpread1.MaxRows + 1
               vaSpread1.Row = vaSpread1.MaxRows
               vaSpread1.Col = 1
               vaSpread1.text = "0"
               
               vaSpread1.Col = 2
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS!Cli_codigo
               vaSpread1.Col = 3
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS!Cli_nombre
               
               vaSpread1.Col = 4
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS!Reg_Codigo
               vaSpread1.Col = 5
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS!reg_nombre
               vaSpread1.Col = 6
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS!Ser_codigo
               vaSpread1.Col = 7
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS!ser_nombre
               RS.MoveNext
            
            Loop
            RS.Close
            Set RS = Nothing
            vaSpread1.Visible = True
        
        End If
    
    End Select

Exit Sub

Error:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Exit Sub

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

Dim i As Long
vaSpread1.Col = 1

For i = BlockRow To BlockRow2
    
    vaSpread1.Row = i
    vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")

Next

End Sub

Function ValidarDatos() As Boolean

On Error GoTo Man_Error

ValidarDatos = True

'-------> Validar regimen
'If Trim(fpayuda(0).Caption) = "" Then
'
'   MsgBox "Debe registrar ceco...", vbExclamation + vbOKOnly, MsgTitulo
'   ValidarDatos = False
'   Exit Function
'
'End If

''-------> Validar regimen
'If Trim(fpayuda(1).Caption) = "" Then
'
'   MsgBox "Debe registrar regimen...", vbExclamation + vbOKOnly, Msgtitulo
'   ValidarDatos = False
'   Exit Function
'
'End If





'-------> Validar fechas
If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
   
   MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

'If FpFecDesde.text > FpFecHasta.text Then
If CDate(FpFecDesde.text) > CDate(FpFecHasta.text) Then
   
   MsgBox "Fecha Origen No Puede Ser Mayor Que Fecha Destino", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

Exit Function
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo

End Function


Sub encabezado(ByRef RS As ADODB.Recordset, ByRef xlWs As Object, ByRef TipoReporte As Integer, varCeco As String)

On Error GoTo Man_Error

Dim fldCount As Long
Dim icol As Long

'-------> Copy field names to the first row of the worksheet
fldCount = RS.Fields.count
xlWs.Cells(1, 1).Value = "Consumo Ingrediente Minuta Bloque"
xlWs.Cells(1, 1).Range("A1:E1").Merge

If TipoReporte = 0 Then
    '--> Centro de Costo
    xlWs.Cells(3, 1).Value = "Centro de Costo: "
    xlWs.Cells(3, 3).Value = LimpiaDato(Trim(fpText.text))
    xlWs.Cells(3, 1).Range("A3:B3").Merge
ElseIf TipoReporte = 1 Then
    '--> Org. de Compra
    xlWs.Cells(3, 1).Value = "Organización de Compra: "
    xlWs.Cells(3, 3).Value = LimpiaDato(Trim(fpText2.text))
    xlWs.Cells(3, 1).Range("A3:B3").Merge
ElseIf TipoReporte = 2 Then
    '--> Org. de Compra
    xlWs.Cells(3, 1).Value = "Organización de Compra: "
    xlWs.Cells(3, 3).Value = LimpiaDato(Trim(fpText2.text))
    xlWs.Cells(3, 1).Range("A3:B3").Merge
End If

xlWs.Cells(4, 1).Value = "Periodo: "
xlWs.Cells(4, 3).Value = FpFecDesde.text & " - " & FpFecHasta.text
xlWs.Range("A4:B4").Merge

For icol = 1 To fldCount
    xlWs.Cells(6, icol).Value = RS.Fields(icol - 1).Name
Next

Exit Sub
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo

End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i             As Long
Dim X             As Long
Dim indactivo     As Integer
Dim TexBus        As String
Dim EstBuq        As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(Text1(Index).text, ",")

If Index = 2 Then
   
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(6).text = ""
   Text1(7).text = ""

ElseIf Index = 3 Then
   
   Text1(2).text = ""
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(6).text = ""
   Text1(7).text = ""

ElseIf Index = 4 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(5).text = ""
   Text1(6).text = ""
   Text1(7).text = ""

ElseIf Index = 5 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(6).text = ""
   Text1(7).text = ""

ElseIf Index = 6 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(7).text = ""

ElseIf Index = 7 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(6).text = ""

End If

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 8
    vaSpread1.text = 0

Next

Select Case Index

Case 2, 3, 4, 5, 6, 7
    
    vaSpread1.Visible = False
    
    If Trim(Text1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread1.Col = 2
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 8
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 8
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 8
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 8
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 8
                 vaSpread1.text = 0
                 
              End If
           
           End If
        
        Next i
        
        vaSpread1.SetActiveCell Index + 1, 1
        vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread1.ColUserSortIndicator(IIf(Trim(wvarArr8020(X)) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread1.SortKey(1) = IIf(Trim(wvarArr8020(X)) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(Text1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 8
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
'Resume
End Sub
