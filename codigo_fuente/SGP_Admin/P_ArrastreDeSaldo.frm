VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form P_ArrastreDeSaldo 
   Caption         =   "Arrastre de Saldo"
   ClientHeight    =   10200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   7335
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   14535
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   3
         Left            =   2520
         TabIndex        =   25
         Top             =   6000
         Width           =   4860
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   26
            Top             =   135
            Width           =   4755
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   2
         Left            =   1440
         TabIndex        =   23
         Top             =   6000
         Width           =   1020
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   24
            Top             =   135
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   0
         Left            =   8640
         TabIndex        =   19
         Top             =   6000
         Width           =   4380
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   45
            TabIndex        =   8
            Top             =   135
            Width           =   4275
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   1
         Left            =   7560
         TabIndex        =   18
         Top             =   6000
         Width           =   1020
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   7
            Top             =   135
            Width           =   915
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir Arrastre Saldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9840
         TabIndex        =   9
         Top             =   6720
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   11280
         TabIndex        =   10
         Top             =   6720
         Width           =   1275
      End
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
         Height          =   420
         Left            =   12645
         TabIndex        =   11
         Top             =   6720
         Width           =   1275
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5535
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   14175
         _Version        =   393216
         _ExtentX        =   25003
         _ExtentY        =   9763
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
         MaxCols         =   13
         MaxRows         =   25
         SpreadDesigner  =   "P_ArrastreDeSaldo.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   2520
      TabIndex        =   12
      Top             =   120
      Width           =   9855
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1875
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2085
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   1875
         TabIndex        =   3
         Top             =   1920
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
      Begin EditLib.fpDateTime FpFecHasta 
         Height          =   315
         Left            =   7815
         TabIndex        =   4
         Top             =   1920
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   9240
         TabIndex        =   5
         Top             =   1920
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   1875
         TabIndex        =   1
         Top             =   675
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
      Begin EditLib.fpText fpOrg 
         Height          =   315
         Left            =   1875
         TabIndex        =   22
         Top             =   1080
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   1875
         TabIndex        =   28
         Top             =   1510
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
         AutoCase        =   1
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
         Index           =   1
         Left            =   3615
         TabIndex        =   30
         Top             =   1500
         Width           =   5415
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
         Index           =   1
         Left            =   3660
         TabIndex        =   29
         Top             =   1510
         Width           =   5415
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3165
         Picture         =   "P_ArrastreDeSaldo.frx":0853
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
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
         Index           =   4
         Left            =   480
         TabIndex        =   27
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Org. Compras"
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
         Left            =   480
         TabIndex        =   21
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Pedido"
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
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   1305
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3165
         Picture         =   "P_ArrastreDeSaldo.frx":0B5D
         Top             =   600
         Width           =   480
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
         Left            =   3615
         TabIndex        =   2
         Top             =   675
         Width           =   5415
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
         Left            =   480
         TabIndex        =   16
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
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
         Left            =   480
         TabIndex        =   15
         Top             =   1995
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin"
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
         Left            =   6720
         TabIndex        =   14
         Top             =   1995
         Width           =   855
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
         Left            =   3660
         TabIndex        =   17
         Top             =   720
         Width           =   5415
      End
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
            Picture         =   "P_ArrastreDeSaldo.frx":0E67
            Key             =   "IMG1"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "P_ArrastreDeSaldo"
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

On Error GoTo Man_Error

Dim i            As Long
Dim ISeleccion   As Boolean
Dim MyBuffer     As String
Dim RS           As New ADODB.Recordset
Dim idpedido     As Double
Dim IdIng        As String
Dim Ceco         As String
Dim tipopedido   As String

ISeleccion = False

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 1
    
    If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
    
       ISeleccion = True
       Exit For
       
    End If

Next i

If Not ISeleccion Then

   MsgBox "Debe haber por lo menos un ítem seleccionado de la lista, para actualizar...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub

End If

If MsgBox("Esta seguro realizar mover cero arrastre de saldo...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<UpdateSaldo>"

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 1
    
    Ceco = ""
    IdIng = ""
    
    If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
    
       vaSpread1.Col = 2
       Ceco = vaSpread1.text
       
       vaSpread1.Col = 4
       IdIng = vaSpread1.text
       
       MyBuffer = MyBuffer & " <DetSaldo"
       MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
       MyBuffer = MyBuffer & " IdIng = " & Chr(34) & IdIng & Chr(34)
       MyBuffer = MyBuffer & "/>"

    End If
    
Next i

MyBuffer = MyBuffer & "</UpdateSaldo>"
      
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

tipopedido = IIf(Combo1(0).ListIndex = -1, "Null", Val(fg_codigocbo(Combo1, 0, 1, "")))

If tipopedido <> 2 Then

   Set RS = vg_db.Execute("sgpadm_Upd_XmlSaldoCeroCDPAP_V02 '" & MyBuffer & "', '" & Format(FpFecHasta.text, "yyyymmdd") & "', '" & tipopedido & "'")

Else

   Set RS = vg_db.Execute("sgpadm_Upd_XmlSaldoCeroProyectado_V02 '" & MyBuffer & "', '" & Format(FpFecHasta.text, "yyyymmdd") & "', , '" & tipopedido & "'")

End If

If Not RS.EOF Then

   If RS(0) > 0 Or RS(0) < 0 Then
        
     fg_descarga
      
     Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), Me.HelpContextID, "", "", "")

     MsgBox RS(1) & VgLinea, vbCritical, MsgTitulo
          
     RS.Close
     Set RS = Nothing
                 
     Exit Sub
              
   Else
        
      If RS(2) > 0 Then
      
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")

         MsgBox "Proceso finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
      
      Else
      
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_No_Encontraron_Datos_Actualizar"), Me.HelpContextID, "", "", "")
         
         MsgBox "Proceso finalizado, no se encontraron datos que actualizar...", vbInformation + vbOKOnly, Me.Caption
         
      End If
              
   End If

End If

RS.Close
Set RS = Nothing

fg_descarga

Toolbar2_ButtonClick Toolbar2.Buttons(1)

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Command2_Click()

On Error GoTo Man_Error

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "", "")

Me.Hide
Unload Me

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub


Private Sub Command3_Click()

On Error GoTo Man_Error
        
    'registrar Log sistema imprimir
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Informe"), Me.HelpContextID, "", "", "")
        
    Call I_ExcelArrastreSaldo.Show(1)

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
        
End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Dim Sql As String
Dim RS  As New ADODB.Recordset

fg_centra Me

MsgTitulo = "Arrastre de Saldo"
FpFecHasta.text = Format(Date, "dd/mm/yyyy")
FpFecDesde.text = Format(Date, "dd/mm/yyyy")
vaSpread1.MaxRows = 0
Est = True
Me.HelpContextID = vg_OpcM

Command1.Enabled = False
Command3.Enabled = False

If Mid(ValidarUsuarioAcceso(Me.HelpContextID, vg_NUsr), 3, 1) = "1" Then

   Command1.Enabled = True

End If

If Mid(ValidarUsuarioAcceso(Me.HelpContextID, vg_NUsr), 5, 1) = "1" Then

   Command3.Enabled = True

End If

'-------> Llenar combo Tipo Pedido
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = ""
Sql = " sgpadm_Sel_TipoPedido "
Set RS = vg_db.Execute(Sql)
Combo1(0).Clear
'Combo1(0).AddItem "Todos" & Space(150) & "(0)"
Do While Not RS.EOF
      
   Combo1(0).AddItem Trim(RS(1)) & Space(150) & "(" & Trim(RS(0)) & ")"
   
   RS.MoveNext

Loop

RS.Close
Set RS = Nothing
Combo1(0).ListIndex = 0

'-------> Llenar combo Tipo Pedido

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

If IsDate(FpFecHasta.text) = False Then Exit Sub

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpOrg_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo


End Sub

Private Sub fpText_Change(Index As Integer)
 
 On Error GoTo Man_Error
 
    Dim RS As New ADODB.Recordset
    Dim Sql As String
    
    Select Case Index
    
    Case 0
    
        If fpText(0).text = "" Then fpayuda(1).Caption = "": Exit Sub
        
        Set RS = vg_db.Execute("SELECT ing_codigo, ing_nombre " & _
                    "FROM b_ingrediente WITH (NOLOCK) " & _
                    "WHERE ing_codigo = '" & LimpiaDato(Trim(fpText(0).text)) & "' " & _
                    "AND   ing_indppr   = 1 " & _
                    "AND   ing_activo = '1'")
        If RS.EOF Then
           
           RS.Close
           Set RS = Nothing
           fpayuda(1).Caption = ""
           Exit Sub
        
        End If
        fpayuda(1).Caption = Trim(RS!ing_nombre)
        fpText(0).text = RS!ing_codigo
        RS.Close
        Set RS = Nothing
    
    Case 1
    
        If fpText(1).text = "" Then fpayuda(0).Caption = "": Exit Sub
    
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Sql = Trim(LimpiaDato(fpText(1).text))
        Set RS = vg_db.Execute("sgpadm_s_cliente_V02 29, '" & Sql & "', ''")
        If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
        fpayuda(0).Caption = Trim(RS!Cli_nombre)
        RS.Close
        Set RS = Nothing
        
    End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)
 
 On Error GoTo Man_Error
    
Select Case Index

Case 0

    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "Clientesimap"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(1).text = vg_codigo: fpayuda(0).Caption = vg_nombre
    If Me.Visible Then fpOrg.SetFocus
    
Case 1
    
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "INgrediente", "IngReal"
    B_TabEst.Show 1
    Me.Refresh
    Screen.MousePointer = 0
    If vg_codigo = "" Then Exit Sub
    fpText(0).text = vg_codigo
    fpayuda(1).Caption = vg_nombre
    If Me.Visible Then FpFecDesde.SetFocus

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub TextDet2_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(TextDet2(Index).text, ",")

If Index = 2 Then
   
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""
   
ElseIf Index = 3 Then
   
   TextDet2(2).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""

ElseIf Index = 4 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(5).text = ""

ElseIf Index = 5 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""

End If

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 13
    vaSpread1.text = 0
    
Next

Select Case Index

Case 2, 3, 4, 5
    
    vaSpread1.Visible = False
    
    If Trim(TextDet2(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
'           indactivo = UCase(Trim(vaSpread1.Value)) Like IIf(Index = 2 Or Index = 3, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
           indactivo = UCase(Trim(vaSpread1.Value)) Like IIf(Index = 2 Or Index = 4, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
           vaSpread1.Col = Index
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 13
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 13
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 13
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 13
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 13
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
    
    If Trim(TextDet2(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 13
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextDet2(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS         As New ADODB.Recordset
Dim tipopedido As String

Select Case Button.Index

Case 1
  
  fg_carga ""
  
  vaSpread1.MaxRows = 0
  
  TextDet2(2).text = ""
  TextDet2(3).text = ""

  Est = True

  If Not ValidarDatos Then Exit Sub
  
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
    
  tipopedido = IIf(Combo1(0).ListIndex = -1, "Null", Val(fg_codigocbo(Combo1, 0, 1, "")))
  
  Set RS = vg_db.Execute("sgpadm_Sel_ConsultarArrastredeSaldo_V03 '" & Trim(LimpiaDato(fpText(1).text)) & "', '" & tipopedido & "','" & Format(FpFecDesde, ("YYYYmmdd")) & "', '" & Format(FpFecHasta, ("YYYYmmdd")) & "', '" & fpOrg.text & "', '" & Trim(LimpiaDato(fpText(0).text)) & "'")
  
  If Not RS.EOF Then
     
   If RS.RecordCount > 50000 Then
      
      RS.Close
      Set RS = Nothing
      
      fg_descarga
 
      Est = False
      MsgBox "El resultado sobrepasa maximo de fila en grilla, Debera seleccionar filtro por ingrediente...", vbCritical
      Exit Sub
   
   End If
     
     
     Do While Not RS.EOF
      
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
      
        vaSpread1.Col = 1
        vaSpread1.text = "0"
      
        vaSpread1.Col = 2
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = RS(0)
        
        vaSpread1.Col = 3
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = RS(1)
        
        vaSpread1.Col = 4
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = RS(2)
      
        vaSpread1.Col = 5
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = Trim(RS(3))
        
        vaSpread1.Col = 6
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = Trim(RS(4))
        
        vaSpread1.Col = 7
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = RS(5)
      
        vaSpread1.Col = 8
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignRight
        vaSpread1.text = RS(6)
      
'        vaSpread1.Col = 7
'        vaSpread1.CellType = CellTypeStaticText
'        vaSpread1.TypeHAlign = TypeHAlignRight
'        vaSpread1.text = RS(6)
      
'        vaSpread1.Col = 8
'        vaSpread1.CellType = CellTypeStaticText
'        vaSpread1.TypeHAlign = TypeHAlignRight
'        vaSpread1.text = RS(7)
      
'        vaSpread1.Col = 9
'        vaSpread1.CellType = CellTypeStaticText
'        vaSpread1.TypeHAlign = TypeHAlignRight
'        vaSpread1.text = RS(8)
        
'        vaSpread1.Col = 10
'        vaSpread1.CellType = CellTypeStaticText
'        vaSpread1.TypeHAlign = TypeHAlignRight
'        vaSpread1.text = RS(2)
        
        vaSpread1.Col = 13
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = 0
             
        RS.MoveNext
        
     Loop
     
  Else
     
     vaSpread1.MaxRows = 0
     MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo
  
  End If
  RS.Close
  Set RS = Nothing
  
  fg_descarga
 
  Est = False

End Select

Exit Sub
Man_Error:
fg_descarga
Est = True
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Function ValidarDatos() As Boolean

On Error GoTo Man_Error

Dim seleccion As Integer
Dim i As Long

ValidarDatos = True

'-------> Validar fechas
If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
   
   MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If
    
If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
   
   MsgBox "La fecha hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

If Trim(fpayuda(0).Caption) = "" And Trim(fpOrg.text) = "" Then
   
   MsgBox "Centro de costo o bien Org. Compras, debe ser informado...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

If Combo1(0).ListIndex = -1 Then

   MsgBox "Tipo pedido debe ser informado...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

Exit Function
Man_Error:
fg_descarga
Est = True
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Function

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Est = True

Select Case BlockCol

Case 1
    
    Dim i As Long
    vaSpread1.Col = 1
       
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
                
        vaSpread1.Col = 1
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    If BlockRow = -1 Then Exit Sub

    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
              
        vaSpread1.Col = 1
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    
        End If
        
    Next
    
End Select

Est = False

Exit Sub
Man_Error:
    Est = False
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

'Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
'
'On Error GoTo Man_Error
'
'Dim EstSel As Boolean
'
'If Est Or ButtonDown = 0 Or vaSpread1.MaxRows < 1 Then Exit Sub
'
'Dim i As Long
'
'For i = 1 To vaSpread1.MaxRows
'
'    vaSpread1.Row = i
'
'    vaSpread1.Col = 1
'
'       If vaSpread1.text = "1" Then
'
'          Est = True
'          vaSpread1.text = "0"
'          Est = False
'
'        End If
'
'Next i
'
'Exit Sub
'Man_Error:
'    fg_descarga
'    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
'
'End Sub
