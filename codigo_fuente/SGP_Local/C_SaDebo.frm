VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form C_SaDebo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Consulta Salida o Devolución a Bodega"
   ClientHeight    =   7110
   ClientLeft      =   1875
   ClientTop       =   2415
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5625
      Left            =   495
      TabIndex        =   16
      Top             =   1500
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   9922
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Resumen"
      TabPicture(0)   =   "C_SaDebo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "C_SaDebo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3(6)"
      Tab(1).Control(1)=   "Label3(7)"
      Tab(1).Control(2)=   "Frame2(1)"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         Height          =   4845
         Index           =   1
         Left            =   -74790
         TabIndex        =   23
         Top             =   690
         Width           =   9615
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   2685
            Index           =   1
            Left            =   90
            TabIndex        =   24
            Top             =   1770
            Width           =   9435
            _Version        =   393216
            _ExtentX        =   16642
            _ExtentY        =   4736
            _StockProps     =   64
            DisplayRowHeaders=   0   'False
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
            MaxRows         =   20
            SpreadDesigner  =   "C_SaDebo.frx":0038
         End
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   1455
            Index           =   2
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   9435
            _Version        =   393216
            _ExtentX        =   16642
            _ExtentY        =   2566
            _StockProps     =   64
            DisplayRowHeaders=   0   'False
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
            MaxRows         =   5
            ScrollBars      =   2
            SpreadDesigner  =   "C_SaDebo.frx":058F
         End
         Begin VB.Label Label3 
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
            Index           =   5
            Left            =   9105
            TabIndex        =   28
            Top             =   4500
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Totales "
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
            Left            =   6300
            TabIndex        =   27
            Top             =   4500
            Width           =   705
         End
         Begin VB.Label Label5 
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
            Index           =   3
            Left            =   510
            TabIndex        =   26
            Top             =   4560
            Width           =   975
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00C0FFC0&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   3
            Left            =   120
            Top             =   4590
            Width           =   300
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H80000018&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   2
            Left            =   1740
            Top             =   4590
            Width           =   300
         End
         Begin VB.Label Label5 
            Caption         =   "Producto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   2100
            TabIndex        =   25
            Top             =   4560
            Width           =   915
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4605
         Index           =   0
         Left            =   1320
         TabIndex        =   17
         Top             =   690
         Width           =   7395
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3885
            Index           =   0
            Left            =   90
            TabIndex        =   18
            Top             =   210
            Width           =   7125
            _Version        =   393216
            _ExtentX        =   12568
            _ExtentY        =   6853
            _StockProps     =   64
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   7
            MaxRows         =   20
            SpreadDesigner  =   "C_SaDebo.frx":08BA
         End
         Begin VB.Label Label3 
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
            Left            =   6840
            TabIndex        =   22
            Top             =   4200
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label Label3 
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
            Left            =   5160
            TabIndex        =   21
            Top             =   4170
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label Label3 
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
            Index           =   1
            Left            =   3570
            TabIndex        =   20
            Top             =   4170
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Totales "
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
            TabIndex        =   19
            Top             =   4200
            Width           =   705
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Totales "
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
         Left            =   -70470
         TabIndex        =   30
         Top             =   460
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha : "
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
         Left            =   -71310
         TabIndex        =   29
         Top             =   460
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   20
      TabIndex        =   4
      Top             =   390
      Width           =   11145
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   6495
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
         Left            =   855
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
         Left            =   855
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
         Left            =   6495
         TabIndex        =   3
         Top             =   570
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
         Text            =   "12/2017"
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
         Index           =   2
         Left            =   2250
         TabIndex        =   14
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   7890
         TabIndex        =   12
         Top             =   150
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   1725
         Picture         =   "C_SaDebo.frx":0E7C
         Top             =   510
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   7365
         Picture         =   "C_SaDebo.frx":1186
         Top             =   60
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   1725
         Picture         =   "C_SaDebo.frx":1490
         Top             =   90
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
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   255
         Width           =   735
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
         Left            =   5580
         TabIndex        =   8
         Top             =   225
         Width           =   750
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
         TabIndex        =   7
         Top             =   675
         Width           =   705
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
         Left            =   5580
         TabIndex        =   6
         Top             =   630
         Width           =   540
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2250
         TabIndex        =   10
         Top             =   180
         Width           =   3135
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   2295
         TabIndex        =   11
         Top             =   225
         Width           =   3135
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   7935
         TabIndex        =   13
         Top             =   195
         Width           =   3135
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   2295
         TabIndex        =   15
         Top             =   645
         Width           =   3135
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "C_SaDebo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim est As Boolean

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_carga ""
est = True
Msgtitulo = "Informe Consulta Salida o Devolución a Bodega"
Me.HelpContextID = vg_OpcM
Me.Height = 7620
Me.Width = 11280
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = IIf(Val(Mid(ValidarUsuario(Me), 4, 1)) = 1, True, False): BtnX.ToolTipText = "Vista Previa"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
If vg_tipser Then
   label2(2).Visible = False
   label2(3).Visible = False
   fpLongInteger1(0).Visible = False
   fpLongInteger1(1).Visible = False
   Image1(1).Visible = False
   Image1(2).Visible = False
   fpayuda(1).Visible = False
   fpayuda(2).Visible = False
   fpayuda(4).Visible = False
   fpayuda(5).Visible = False
   Label1(0).Top = 225
   fpDateTime1.Top = 150
End If

fpDateTime1.text = Format(Date, "mm/yyyy")
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
vaSpread1(0).MaxRows = 0
vaSpread1(1).MaxRows = 0
vaSpread1(2).MaxRows = 0
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
est = False
fg_descarga
End Sub

Private Sub fpDateTime1_Change()
If est Then Exit Sub
Mover_Datos
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    RS.Open RutinaLectura.Regimen(2, Val(fpLongInteger1(0).Value), ""), vg_db, adOpenStatic
    If RS.EOF Then
       RS.Close: Set RS = Nothing
       fpayuda(1).Caption = ""
    Else
       fpayuda(1).Caption = Trim(RS!reg_nombre)
       RS.Close: Set RS = Nothing
    End If
    Mover_Datos
Case 1
    RS.Open RutinaLectura.Servicio(8, Val(fpLongInteger1(1).Value), ""), vg_db, adOpenStatic
    If RS.EOF Then
       RS.Close: Set RS = Nothing
       fpayuda(2).Caption = ""
    Else
       fpayuda(2).Caption = Trim(RS!ser_nombre)
       RS.Close: Set RS = Nothing
    End If
    Mover_Datos
End Select
End Sub

Private Sub fpLongInteger1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    If Index = 1 Then Image1_Click 1
    If Index = 2 Then Image1_Click 2
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_Change()
If est Then Exit Sub
vaSpread1(1).MaxRows = 0
RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
If RS.EOF Then
   RS.Close: Set RS = Nothing: fpayuda(0).Caption = ""
   fpLongInteger1(0).Value = "": fpayuda(1).Caption = ""
   fpLongInteger1(1).Value = "": fpayuda(2).Caption = ""
Else
   fpayuda(0).Caption = Trim(RS!cli_nombre)
   RS.Close: Set RS = Nothing
End If
fpLongInteger1(0).Value = "": fpayuda(1).Caption = ""
fpLongInteger1(1).Value = "": fpayuda(2).Caption = ""
Mover_Datos
End Sub

Private Sub fpText_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    Image1_Click 0
End Select
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Image1_Click(Index As Integer)
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
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre
    fpDateTime1.SetFocus
End Select
End Sub

Sub Mover_Datos()
Dim salbod As Double, devbod As Double, totbod As Double, codsec As Long, ressec As Boolean, numdoc As Long, auxnumd As Long
Dim sql1 As String, sql2 As String, sql3 As String
fg_carga ""
vaSpread1(0).Visible = False
salbod = 0: devbod = 0: totbod = 0: codsec = 0: auxnumd = 0
Label3(1).Caption = 0: Label3(2).Caption = 0: Label3(3).Caption = 0
Label3(1).Visible = False: Label3(2).Visible = False: Label3(3).Visible = False
vaSpread1(0).MaxRows = 0
SSTab1.Tab = 0
sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2) & "01") & "') ", " '" & Format(fg_Ctod1(Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2) & "01"), "yyyymmdd") & "' ")
sql2 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(Format(dEoM("01/" & fpDateTime1.text), "yyyymmdd")) & "') ", " '" & Format(fg_Ctod1(Format(dEoM("01/" & fpDateTime1.text), "yyyymmdd")), "yyyymmdd") & "' ")
sql3 = IIf(vg_tipbase = "1", " iif(isnull(b.dev_codsec),0,1) ", " CASE WHEN (b.dev_codsec) IS NULL THEN 0 ELSE 1 END ")
If Not vg_tipser Then
   RS.Open "SELECT a.tov_fecpro, a.tov_tipdoc, a.tov_numdoc, " & sql3 & " AS dev_codsec, sum(b.dev_ptotal) AS ptotal " & _
           "FROM  b_totventas a, b_detventas b " & _
           "WHERE a.tov_rutcli = b.dev_rutcli " & _
           "AND   a.tov_tipdoc = b.dev_tipdoc " & _
           "AND   a.tov_numdoc = b.dev_numdoc " & _
           "AND   a.tov_rutcli = '" & fpText.text & "' " & _
           "AND   a.tov_codreg = " & Val(fpLongInteger1(0).Value) & " " & _
           "AND   a.tov_codser = " & Val(fpLongInteger1(1).Value) & " " & _
           "AND  (a.tov_tipdoc = 'SP' OR a.tov_tipdoc = 'DP') " & _
           "AND   b.dev_canmer <> 0 AND a.tov_codbod = " & vg_codbod & " " & _
           "AND   a.tov_estdoc <> 'A' AND a.tov_estdoc <> 'P' " & _
           "AND   a.tov_fecpro >= " & sql1 & " " & _
           "AND   a.tov_fecpro <= " & sql2 & " " & _
           "GROUP BY a.tov_fecpro, a.tov_tipdoc, a.tov_numdoc, " & sql3 & " ORDER BY a.tov_fecpro, a.tov_numdoc", vg_db, adOpenStatic
Else
   RS.Open "SELECT a.tov_fecpro, a.tov_tipdoc, a.tov_numdoc, " & sql3 & " as dev_codsec, sum(b.dev_ptotal) AS ptotal " & _
           "FROM  b_totventas a, b_detventas b " & _
           "WHERE a.tov_rutcli = b.dev_rutcli " & _
           "AND   a.tov_tipdoc = b.dev_tipdoc " & _
           "AND   a.tov_numdoc = b.dev_numdoc " & _
           "AND   a.tov_rutcli = '" & fpText.text & "' " & _
           "AND   a.tov_codreg = 0 " & _
           "AND   a.tov_codser = 0 " & _
           "AND  (a.tov_tipdoc = 'SP' OR a.tov_tipdoc = 'DP') " & _
           "AND   b.dev_canmer <> 0 AND a.tov_codbod = " & vg_codbod & " " & _
           "AND   a.tov_estdoc <> 'A' AND a.tov_estdoc<>'P' " & _
           "AND   a.tov_fecpro >= " & sql1 & " " & _
           "AND   a.tov_fecpro <= " & sql2 & " " & _
           "GROUP BY a.tov_fecpro, a.tov_tipdoc, a.tov_numdoc, " & sql3 & " ORDER BY a.tov_fecpro, a.tov_numdoc", vg_db, adOpenStatic
End If
If Not RS.EOF Then
   Do While Not RS.EOF
      If vaSpread1(0).MaxRows > 0 Then
         If vaSpread1(0).SearchCol(1, 0, vaSpread1(0).MaxRows, Trim(RS!tov_fecpro), SearchFlagsNone) <> -1 And codsec <> RS!dev_codsec Then
            If auxnumd <> RS!tov_numdoc Then
               vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
               vaSpread1(0).Row = vaSpread1(0).MaxRows
            Else
               vaSpread1(0).Row = vaSpread1(0).SearchCol(1, 0, vaSpread1(0).MaxRows, Trim(RS!tov_fecpro), SearchFlagsNone)
               codsec = RS!dev_codsec
            End If
            auxnumd = RS!tov_numdoc
         Else
            vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
            vaSpread1(0).Row = vaSpread1(0).MaxRows
            auxnumd = RS!tov_numdoc
         End If
      Else
         vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
         vaSpread1(0).Row = vaSpread1(0).MaxRows
      End If
      vaSpread1(0).Col = 1
      vaSpread1(0).CellType = CellTypeStaticText
      vaSpread1(0).TypeHAlign = TypeHAlignCenter
      vaSpread1(0).text = RS!tov_fecpro
      
      vaSpread1(0).Col = 2
      vaSpread1(0).CellType = CellTypeStaticText
      vaSpread1(0).TypeHAlign = TypeHAlignLeft
      vaSpread1(0).text = IIf(RS!dev_codsec > 0, "Sector", "Resumen")

      If RS!tov_tipdoc = "SP" Then
         vaSpread1(0).Col = 3
         vaSpread1(0).CellType = CellTypeStaticText
         vaSpread1(0).TypeHAlign = TypeHAlignRight
         vaSpread1(0).text = Format(RS!ptotal, fg_Pict(6, 0))
         salbod = Round(salbod + RS!ptotal)
         
         vaSpread1(0).Col = 6
         vaSpread1(0).text = RS!tov_numdoc
      ElseIf RS!tov_tipdoc = "DP" Then
         vaSpread1(0).Col = 4
         vaSpread1(0).CellType = CellTypeStaticText
         vaSpread1(0).TypeHAlign = TypeHAlignRight
         vaSpread1(0).text = Format(RS!ptotal, fg_Pict(6, 0))
         devbod = Round(devbod + RS!ptotal)
         vaSpread1(0).Col = 7
         vaSpread1(0).text = RS!tov_numdoc
      End If
      vaSpread1(0).Col = 5
      vaSpread1(0).CellType = CellTypeStaticText
      vaSpread1(0).TypeHAlign = TypeHAlignRight
      vaSpread1(0).text = Format(IIf(RS!tov_tipdoc = "SP", (Round(Val(vaSpread1(0).Value) + RS!ptotal)), Round(((Val(vaSpread1(0).Value) - RS!ptotal)))), fg_Pict(6, 0))
      totbod = IIf(RS!tov_tipdoc = "SP", Round(totbod + RS!ptotal), Round(totbod - RS!ptotal))
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
Label3(1).Caption = Format(salbod, fg_Pict(6, 0)): Label3(1).Visible = IIf(salbod > 0, True, False)
Label3(2).Caption = Format(devbod, fg_Pict(6, 0)): Label3(2).Visible = IIf(devbod > 0, True, False)
Label3(3).Caption = Format(totbod, fg_Pict(6, 0)): Label3(3).Visible = IIf(totbod > 0, True, False)
vaSpread1(0).Visible = True
If vaSpread1(0).MaxRows > 0 Then
   vaSpread1(0).Row = -1: vaSpread1(0).Col = -1
   vaSpread1(0).BackColor = &H80000018
   vaSpread1(0).SetActiveCell 3, 1
   vaSpread1(0).Row = 1
   vaSpread1(0).Col = 2: ressec = IIf(vaSpread1(0).text = "Resumen", True, False)
   vaSpread1(0).Col = 6: numdoc = Val(vaSpread1(0).text)
   vaSpread1(0).Col = 1
   MoverDatosDetalle vaSpread1(0).text, "SP", ressec, numdoc
   Label3(7).Caption = vaSpread1(0).text
   SSTab1.TabCaption(1) = "Detalle Salida"
   SSTab1.TabEnabled(1) = IIf(numdoc = 0, False, True)
Else
   SSTab1.TabEnabled(1) = False
   vaSpread1(1).MaxRows = 0
End If
fg_descarga
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If vaSpread1(0).MaxRows < 1 And SSTab1.Tab = 0 Then MsgBox "No Existe Resumen a Visualizar", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If vaSpread1(1).MaxRows < 1 And SSTab1.Tab = 1 Then MsgBox "No Existe Detalle a Visualizar", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Select Case SSTab1.Tab
    Case 0
        I_ResSalDevBod Me
    Case 1
        Dim Tipo As String, ressec As Boolean, numdoc As Long
        Tipo = ""
        vaSpread1(0).Row = vaSpread1(0).ActiveRow
        vaSpread1(0).Col = vaSpread1(0).ActiveCol
        Select Case vaSpread1(0).Col
        Case 3
            Tipo = "SP"
            vaSpread1(0).Col = 6: numdoc = vaSpread1(0).text
        Case 4
            Tipo = "DP"
            vaSpread1(0).Col = 7: numdoc = vaSpread1(0).text
        End Select
        vaSpread1(0).Col = 2: ressec = IIf(vaSpread1(0).text = "Resumen", True, False)

        vaSpread1(0).Col = 1
        I_DetSalDevBod Me, Tipo, vaSpread1(0).text, ressec, numdoc
    End Select
Case 3
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub vaSpread1_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim ressec As Boolean, numdoc As Long
If Index = 2 Then
   If vaSpread1(2).MaxRows < 1 Or Row < 1 Then Exit Sub
   Dim codsec As String, esthidden As Boolean, j As Long
   esthidden = True
   vaSpread1(2).Row = vaSpread1(2).ActiveRow
   vaSpread1(2).Col = 1: codsec = vaSpread1(2).text
   vaSpread1(1).Visible = False
   For i = 1 To vaSpread1(1).MaxRows
       vaSpread1(1).Row = i
       vaSpread1(1).Col = 8
       If codsec = vaSpread1(1).text Then
          If j = 0 Then vaSpread1(1).SetActiveCell 1, i: j = 1
          vaSpread1(1).Col = 5
          If Trim(vaSpread1(1).text) = "" Then vaSpread1(1).RowHidden = False Else vaSpread1(1).RowHidden = False
       Else
          vaSpread1(1).RowHidden = True
       End If
   Next i
   vaSpread1(1).Visible = True
End If

If Index = 1 Or Index = 2 Then Exit Sub
If vaSpread1(0).MaxRows < 1 Or Index = 1 Or Row < 1 Then SSTab1.TabCaption(1) = "Detalle ": SSTab1.TabEnabled(1) = False: Exit Sub
vaSpread1(0).Row = Row
Select Case Col
Case 1, 2, 5
     SSTab1.TabEnabled(1) = False
Case 3
    vaSpread1(0).Col = Col
    If Val(vaSpread1(0).Value) < 1 Then SSTab1.TabCaption(1) = "Detalle Salida": SSTab1.TabEnabled(1) = False: Exit Sub
    vaSpread1(0).Col = 2: ressec = IIf(vaSpread1(0).text = "Resumen", True, False)
    vaSpread1(0).Col = 6: numdoc = vaSpread1(0).text
    vaSpread1(0).Col = 1
    MoverDatosDetalle vaSpread1(0).text, "SP", ressec, numdoc
    Label3(7).Caption = vaSpread1(0).text
    SSTab1.TabCaption(1) = "Detalle Salida"
    SSTab1.TabEnabled(1) = True
Case 4
    vaSpread1(0).Col = Col
    If Val(vaSpread1(0).Value) < 1 Then SSTab1.TabCaption(1) = "Detalle Devolución": SSTab1.TabEnabled(1) = False: Exit Sub
    vaSpread1(0).Col = 2: ressec = IIf(vaSpread1(0).text = "Resumen", True, False)
    vaSpread1(0).Col = 7: numdoc = vaSpread1(0).text
    vaSpread1(0).Col = 1
    MoverDatosDetalle vaSpread1(0).text, "DP", ressec, numdoc
    Label3(7).Caption = vaSpread1(0).text
    SSTab1.TabCaption(1) = "Detalle Devolución"
    SSTab1.TabEnabled(1) = True
End Select
End Sub

Sub MoverDatosDetalle(fecini As Date, Op As String, ressec As Boolean, numdoc As Long)
Dim aAp As String, codsec As String, coding As String, vNumRac As Long
Dim totgrl As Double, vTotSec As Double
fg_carga ""
codsec = "0": coding = "": vNumRac = 0
vaSpread1(1).MaxRows = 0: totgrl = 0
vaSpread1(2).MaxRows = 0
vaSpread1(2).Row = -1: vaSpread1(2).Col = -1
vaSpread1(2).BackColor = &H80000018
Label3(5).Caption = 0
sql1 = IIf(vg_tipbase = "1", " cdate('" & fecini & "') ", " '" & Format(fecini, "yyyymmdd") & "' ")
If ressec Then
   sql2 = IIf(vg_tipbase = "1", " ORDER BY dev.dev_numlin ", "")
   RS.Open "SELECT ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 as sec_codigo, '' as sec_nombre, 0 as sec_orden, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
            "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
            "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_unidad uni " & _
            "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
            "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_coding = ing.ing_codigo " & _
            "AND   ing.ing_unimed = unm.unm_codigo AND dev.dev_codmer = pro.pro_codigo " & _
            "AND   pro.pro_coduni = uni.uni_codigo " & _
            "AND   tov.tov_rutcli = '" & LimpiaDato(Trim(fpText.text)) & "' AND tov.tov_numdoc = " & numdoc & " " & _
            "AND   tov.tov_fecpro = " & sql1 & " " & _
            "AND   tov.tov_codreg = " & Val(fpLongInteger1(0).Value) & " AND tov.tov_codser = " & Val(fpLongInteger1(1).Value) & " " & _
            "AND   tov.tov_tipdoc = '" & Op & "' AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' AND tov.tov_codbod = " & vg_codbod & " " & _
            "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin " & sql2 & " " & _
            "UNION ALL " & _
            "SELECT '' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' as unm_nomcor, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 AS sec_codigo, '' AS sec_nombre, 0 AS sec_orden, 0 AS canmin,  dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
            "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
            "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
            "AND   tov.tov_numdoc = dev.dev_numdoc AND  dev.dev_codmer = pro.pro_codigo " & _
            "AND   pro.pro_coduni = uni.uni_codigo AND  tov.tov_rutcli = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
            "AND   tov.tov_fecpro = " & sql1 & " " & _
            "AND   tov.tov_codreg = " & Val(fpLongInteger1(0).Value) & " AND tov.tov_codser = " & Val(fpLongInteger1(1).Value) & " " & _
            "AND   dev.dev_numdoc = " & numdoc & " AND  tov.tov_tipdoc = '" & Op & "'  " & _
            "AND   tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' AND tov.tov_codbod = " & vg_codbod & " AND (dev.dev_coding = '' OR (dev.dev_coding) IS NULL OR dev.dev_codsec=-1) ORDER BY dev.dev_numlin", vg_db, adOpenStatic
Else
   sql2 = IIf(vg_tipbase = "1", " ORDER BY sec.sec_orden, dev.dev_numlin ", "")
   RS.Open "SELECT ing.ing_codigo, ing.ing_nombre,unm.unm_nomcor, sec.sec_codigo, sec.sec_nombre, sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, SUM(dev.dev_canmin * pro.pro_facing) AS canmin, " & _
            "SUM(dev.dev_canmer) AS dev_canmer, SUM(dev.dev_ptotal) AS dev_ptotal " & _
            "FROM  b_totventas tov, b_detventas dev, b_ingrediente ing, b_productos pro, a_unidadmed unm, a_sector sec, a_unidad uni " & _
            "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
            "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_coding = ing.ing_codigo " & _
            "AND   ing.ing_unimed = unm.unm_codigo AND dev.dev_codmer = pro.pro_codigo " & _
            "AND   dev.dev_codsec = sec.sec_codigo AND pro.pro_coduni = uni.uni_codigo " & _
            "AND   tov.tov_rutcli = '" & LimpiaDato(Trim(fpText.text)) & "' AND tov.tov_numdoc = " & numdoc & " " & _
            "AND   tov.tov_fecpro = " & sql1 & " " & _
            "AND   tov.tov_codreg = " & Val(fpLongInteger1(0).Value) & " AND tov.tov_codser = " & Val(fpLongInteger1(1).Value) & " " & _
            "AND   tov.tov_tipdoc = '" & Op & "' AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' AND tov.tov_codbod = " & vg_codbod & " " & _
            "GROUP BY ing.ing_codigo, ing.ing_nombre, unm.unm_nomcor, sec.sec_codigo, sec.sec_nombre,  sec.sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin " & sql2 & " " & _
            "UNION ALL " & _
            "SELECT '' AS ing_codigo, 'Estructura Fija' AS ing_nombre, '' as unm_nomcor, -1 AS sec_codigo, 'Estructura Fija' AS sec_nombre, 999999999 AS sec_orden, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, dev.dev_canmin, dev.dev_predoc, dev.dev_numlin, 0 AS canmin,  dev.dev_canmer AS dev_canmer, dev.dev_ptotal AS dev_ptotal " & _
            "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni " & _
            "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
            "AND   tov.tov_numdoc = dev.dev_numdoc AND dev.dev_codmer = pro.pro_codigo " & _
            "AND   pro.pro_coduni = uni.uni_codigo AND  tov.tov_rutcli = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
            "AND   tov.tov_fecpro = " & sql1 & " " & _
            "AND   tov.tov_codreg = " & Val(fpLongInteger1(0).Value) & " AND tov.tov_codser = " & Val(fpLongInteger1(1).Value) & " " & _
            "AND   dev.dev_numdoc = " & numdoc & " AND  tov.tov_tipdoc = '" & Op & "'  " & _
            "AND   tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' AND tov.tov_codbod = " & vg_codbod & " AND (dev.dev_coding = '' OR (dev.dev_coding) IS NULL OR dev.dev_codsec = -1) ORDER BY sec_orden, dev.dev_numlin", vg_db, adOpenStatic
       '------- Traer raciones
       RS1.Open "SELECT mir_fecmin, SUM(mir_nrorac) AS nrorac FROM b_minutaraciones " & _
                "WHERE mir_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                "AND   mir_codreg = " & Val(fpLongInteger1(0).Value) & " " & _
                "AND   mir_codser = " & Val(fpLongInteger1(1).Value) & " " & _
                "AND  (mir_rutcli = 'PRODUCIDAS') " & _
                "AND   mir_fecmin = " & Format(fecini, "yyyymmdd") & " GROUP BY mir_fecmin", vg_db, adOpenStatic
       If Not RS1.EOF Then vNumRac = RS1!nrorac
       RS1.Close: Set RS1 = Nothing
End If
 vaSpread1(1).Visible = False
 vaSpread1(2).Visible = False
 i = 0
 Do While Not RS.EOF
    If codsec <> RS!sec_codigo And Not ressec Then
       If vaSpread1(2).MaxRows > 0 And vNumRac > 0 And vTotSec > 0 And Not ressec Then
          vaSpread1(2).Col = 4
          vaSpread1(2).TypeHAlign = TypeHAlignRight
          vaSpread1(2).text = Format((vTotSec / vNumRac), fg_Pict(6, vg_DPr))
       End If
       vTotSec = 0
       vaSpread1(2).MaxRows = vaSpread1(2).MaxRows + 1
       vaSpread1(2).Row = vaSpread1(2).MaxRows
       vaSpread1(2).Col = 1: vaSpread1(2).Value = IIf(RS!sec_codigo = -1, "estfij", RS!sec_codigo)
       vaSpread1(2).Col = 2: vaSpread1(2).Value = Trim(RS!sec_nombre) & " (Nº Raciones :  " & vNumRac & ")"
       codsec = RS!sec_codigo
       coding = 0
    End If
    '------- Ingrediente
    If coding <> RS!ing_codigo Then
       i = i + 1: vaSpread1(1).MaxRows = i
       vaSpread1(1).Row = i
       vaSpread1(1).RowHidden = IIf(vaSpread1(2).Row = 1 Or ressec, False, True)
       vaSpread1(1).Col = 1: vaSpread1(1).TypeHAlign = TypeHAlignLeft: vaSpread1(1).text = RS!ing_codigo
       vaSpread1(1).Col = 2: vaSpread1(1).TypeHAlign = TypeHAlignLeft: vaSpread1(1).text = RS!ing_nombre
       vaSpread1(1).Col = 3: vaSpread1(1).TypeHAlign = TypeHAlignLeft: vaSpread1(1).text = RS!unm_nomcor
       vaSpread1(1).Col = 4: vaSpread1(1).TypeHAlign = TypeHAlignRight: vaSpread1(1).text = IIf(RS!ing_codigo = "", "", Format(RS!canmin, fg_Pict(9, vg_DCa)))
       vaSpread1(1).Col = 8: vaSpread1(1).text = IIf(RS!sec_codigo = -1, "estfij", RS!sec_codigo)
       vaSpread1(1).Col = -1
       vaSpread1(1).FontBold = True
       vaSpread1(1).BackColor = Shape1(3).FillColor
       coding = RS!ing_codigo
    End If
    '------- Productos
    i = i + 1: vaSpread1(1).MaxRows = i
    vaSpread1(1).Row = i
    vaSpread1(1).RowHidden = IIf(vaSpread1(2).Row = 1 Or ressec, False, True)
    vaSpread1(1).Col = 1: vaSpread1(1).TypeHAlign = TypeHAlignLeft: vaSpread1(1).text = RS!pro_codigo
    vaSpread1(1).Col = 2: vaSpread1(1).TypeHAlign = TypeHAlignLeft: vaSpread1(1).text = RS!pro_nombre
    vaSpread1(1).Col = 3: vaSpread1(1).TypeHAlign = TypeHAlignLeft: vaSpread1(1).text = RS!uni_nomcor
    vaSpread1(1).Col = 4: vaSpread1(1).TypeHAlign = TypeHAlignRight: vaSpread1(1).text = Format(RS!dev_canmin, fg_Pict(9, vg_DCa))
    vaSpread1(1).Col = 5: vaSpread1(1).TypeHAlign = TypeHAlignRight: vaSpread1(1).text = Format(RS!dev_canmer, fg_Pict(9, vg_DCa))
    vaSpread1(1).Col = 6: vaSpread1(1).TypeHAlign = TypeHAlignRight: vaSpread1(1).text = Format(RS!dev_predoc, fg_Pict(9, vg_DPr))
    vaSpread1(1).Col = 7: vaSpread1(1).TypeHAlign = TypeHAlignRight: vaSpread1(1).text = Format(RS!dev_ptotal, fg_Pict(9, vg_DPr))
    vaSpread1(1).Col = 8: vaSpread1(1).text = IIf(RS!sec_codigo = -1, "estfij", RS!sec_codigo)
    '------- Mover sectores totales
    If vaSpread1(2).MaxRows > 0 Then
       vaSpread1(2).Col = 3
       vaSpread1(2).TypeHAlign = TypeHAlignRight
       vaSpread1(2).text = Format(IIf(Trim(vaSpread1(2).text) = "", 0, vaSpread1(2).Value) + (RS!dev_ptotal), fg_Pict(9, vg_DPr))
    End If
    vTotSec = Round(vTotSec + RS!dev_ptotal)
    totgrl = Round(totgrl + RS!dev_ptotal)
    vaSpread1(1).Col = -1: vaSpread1(1).BackColor = Shape1(2).FillColor
    RS.MoveNext
Loop
RS.Close: Set RS = Nothing
If vaSpread1(2).MaxRows > 0 And vNumRac > 0 And vTotSec > 0 And Not ressec Then
   vaSpread1(2).Col = 4
   vaSpread1(2).TypeHAlign = TypeHAlignRight
   vaSpread1(2).text = Format((vTotSec / vNumRac), fg_Pict(6, vg_DPr))
End If
vTotSec = 0
Me.MousePointer = 0
vaSpread1(2).Visible = IIf(ressec, False, True)
If ressec Then
   vaSpread1(1).Top = vaSpread1(2).Top
   vaSpread1(1).Height = 2685 + vaSpread1(2).Height
Else
   vaSpread1(1).Top = 1770
   vaSpread1(1).Height = 2685
End If
vaSpread1(1).Visible = True
Label3(5).Caption = Format(totgrl, fg_Pict(6, 0)): Label3(5).Visible = IIf(totgrl > 0, True, False)
fg_descarga
End Sub
