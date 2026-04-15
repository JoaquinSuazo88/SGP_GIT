VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_EtiquetaReceta 
   Caption         =   "Impresión de Etiquetado de Recetas"
   ClientHeight    =   9705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17655
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9705
   ScaleWidth      =   17655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   9375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16695
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   3015
         Left            =   12240
         TabIndex        =   26
         Top             =   240
         Width           =   4335
         _Version        =   393216
         _ExtentX        =   7646
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
         MaxCols         =   3
         MaxRows         =   10
         SpreadDesigner  =   "I_EtiquetaReceta.frx":0000
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
         Index           =   1
         Left            =   3600
         TabIndex        =   23
         Top             =   2280
         Value           =   -1  'True
         Width           =   1935
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
         Index           =   0
         Left            =   1200
         TabIndex        =   22
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Selección Recetas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Width           =   16455
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   2040
            TabIndex        =   18
            Top             =   5520
            Width           =   4815
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   17
            Top             =   5520
            Width           =   855
         End
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   4935
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   16215
            _Version        =   393216
            _ExtentX        =   28601
            _ExtentY        =   8705
            _StockProps     =   64
            AllowMultiBlocks=   -1  'True
            BackColorStyle  =   1
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
            MaxRows         =   20
            SpreadDesigner  =   "I_EtiquetaReceta.frx":03D7
            ScrollBarTrack  =   3
         End
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Top             =   240
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1676
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
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1676
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
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   8
         Top             =   1440
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
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   360
         Left            =   9240
         TabIndex        =   20
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList1"
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
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Index           =   1
         Left            =   7080
         TabIndex        =   25
         Top             =   1440
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
      Begin VB.Label Label1 
         Caption         =   "Fecha Etiquetado"
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
         Index           =   6
         Left            =   6120
         TabIndex        =   24
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Opción de Impresión "
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
         Left            =   1200
         TabIndex        =   21
         Top             =   1920
         Width           =   1935
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
         Index           =   2
         Left            =   2850
         TabIndex        =   14
         Top             =   930
         Width           =   6735
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
         Left            =   2850
         TabIndex        =   13
         Top             =   555
         Width           =   6735
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
         Left            =   2850
         TabIndex        =   12
         Top             =   210
         Width           =   6735
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
         Index           =   2
         Left            =   2880
         TabIndex        =   11
         Top             =   960
         Width           =   6735
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
         Left            =   2880
         TabIndex        =   10
         Top             =   600
         Width           =   6735
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
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   6735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2400
         Picture         =   "I_EtiquetaReceta.frx":0C81
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2400
         Picture         =   "I_EtiquetaReceta.frx":0F8B
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2400
         Picture         =   "I_EtiquetaReceta.frx":1295
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Minuta"
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
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Régimen"
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
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   9705
      Left            =   17025
      TabIndex        =   19
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   17119
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
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
            Picture         =   "I_EtiquetaReceta.frx":159F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "I_EtiquetaReceta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgTitulo As String
Public lc_Aux As String
Dim Azucares  As String
Dim Calorias  As String
Dim Grasas    As String
Dim Sodio     As String
Dim Logo      As String

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i  As Long

Me.HelpContextID = vg_OpcM

Toolbar1.ImageList = Partida.IL1
'--------------------------- Crea Botones de la toolbar
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
'--------------------------- Da dimensiones al formulario para que no se descentre
Me.Height = 10290
Me.Width = 17895
MsgTitulo = "Impresión Etiquedado de Receta"
fg_centra Me

FpFecDesde(0).text = Format(Date, "dd/mm/yyyy")
FpFecDesde(1).text = Format(Date, "dd/mm/yyyy")

fpText(0).Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText(0).text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)

vaSpread1.MaxRows = 0

'-------
'------- Cargar Nutrientes
'-------

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_NutrienteIndPrincipalSello")

i = 1
vaSpread2.MaxRows = 0
If Not RS.EOF Then

   Do While Not RS.EOF

      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      
      vaSpread2.Col = 1
      vaSpread2.text = IIf(RS(2) = True, "1", "0")
      
      If vaSpread2.text = "1" Then
      
         vaSpread2.Lock = True
         
      End If

      vaSpread2.Col = 2
      vaSpread2.text = RS(0) & " - " & RS(1)

      vaSpread2.Col = 3
      vaSpread2.text = RS(0)

      RS.MoveNext
      
   Loop
  
Else
          
   fg_descarga
   
   MsgBox "No existe información nutrientes, se desactivará el botón impresión...", vbExclamation + vbOKOnly, MsgTitulo
         
   Toolbar1.Buttons(1).Enabled = False
   
End If
RS.Close
Set RS = Nothing

'-------
'------- Validar dirección y resolucion del sitio
'-------

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_ValidarDireccionResolucion '" & MuestraCasino(1) & "'")

If RS.EOF Then
         
   fg_descarga
   
   MsgBox "No existe información [Dirección] o bien [Resolución] en el maestro de contratos. " & "Es importante registrar esos datos, para etiquetado recetas, se desactivará el botón impresión...", vbExclamation + vbOKOnly, MsgTitulo
   
   Toolbar1.Buttons(1).Enabled = False
   
'   RS.Close
'   Set RS = Nothing
'   Exit Sub
   
End If
RS.Close
Set RS = Nothing

'-------
'------- Traer Nombre Logo Etiquetado
'-------

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_NomEtiquetadoNutricional")

If Not RS.EOF Then

   Do While Not RS.EOF

      Azucares = RS(0)
      Calorias = RS(1)
      Grasas = RS(2)
      Sodio = RS(3)
      Logo = RS(4)
          
      RS.MoveNext

   Loop

Else
          
   fg_descarga
   MsgBox "No existe información nombre sello etiquetado en la tabla a_param, se desactivará el botón impresión...", vbExclamation + vbOKOnly, MsgTitulo
   
   Toolbar1.Buttons(1).Enabled = False
'   RS.Close
'   Set RS = Nothing
'   Exit Sub
   
End If
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_Change(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0

     vaSpread1.MaxRows = 0
     
End Select

If IsDate(FpFecDesde(Index).text) = False Then Exit Sub

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Select Case Index
    
    Case 0
        
        vaSpread1.MaxRows = 0
        Set RS = vg_db.Execute("sgp_Sel_RegimenxCodigo " & IIf(Val(fpLongInteger1(0).Value) = 0, -1, Val(fpLongInteger1(0).Value)) & "")
        
        If RS.EOF = True Then
            
            RS.Close
            Set RS = Nothing
            fpayuda(1).Caption = ""
            Exit Sub
        
        End If
        fpayuda(1).Caption = Trim(RS!reg_nombre)
        RS.Close
        Set RS = Nothing
    
    Case 1
        
        vaSpread1.MaxRows = 0
        Set RS = vg_db.Execute("sgp_Sel_ServicioxCodigo " & IIf(Val(fpLongInteger1(1).Value) = 0, -1, Val(fpLongInteger1(1).Value)) & "")
        If RS.EOF Then
            
            RS.Close
            Set RS = Nothing
            fpayuda(2).Caption = ""
            Exit Sub
        
        End If
        fpayuda(2).Caption = Trim(RS!ser_nombre)
        RS.Close
        Set RS = Nothing

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
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
            fpText(0).text = vg_codigo
            fpayuda(0).Caption = vg_nombre
            fpLongInteger1(0).Value = ""
            Let fpayuda(1).Caption = ""
            fpLongInteger1(1).Value = ""
            Let fpayuda(2).Caption = ""
            fpLongInteger1(0).SetFocus
        
        Case 1
            
            vg_left = fpayuda(1).Left + 2300
            vg_nombre = "": vg_codigo = ""
            Call B_TabEst.LlenaDatos("a_regimen", "reg_", "Regimen", "RegBlo")
            Call B_TabEst.Show(1)
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(0).Value = Val(vg_codigo)
            fpLongInteger1(0).SetFocus
            fpayuda(1).Caption = vg_nombre
            fpLongInteger1(1).SetFocus
        
        Case 2
            
            Let vg_left = fpayuda(2).Left + 2300
            Let vg_nombre = ""
            Let vg_codigo = ""
            Call B_TabEst.LlenaDatos("a_servicio", "ser_", "Servicio", "SerBlo")
            Call B_TabEst.Show(1)
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(1).Value = "0"
            fpLongInteger1(1).Value = Val(vg_codigo)
            fpayuda(2).Caption = vg_nombre
            fpLongInteger1(1).SetFocus
            FpFecDesde(0).Enabled = True
            Call FpFecDesde(0).SetFocus
           
    End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub MoverListadoRecetaGrilla(op As Integer)

On Error GoTo Man_Error

Dim RS        As New ADODB.Recordset
Dim vReceta() As Variant
Dim i         As Long

Dim z         As Long
Dim codaux    As Long

Dim lisnom    As String
Dim liscod    As String
Dim cParam    As String
Dim encuentra As Boolean
Dim EstVect   As Boolean

EstVect = False

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

fg_carga ""
vaSpread1.MaxRows = 0

   If op = 0 Then
   
      Set RS = vg_db.Execute("sgp_Sel_EtiquedadoRecetaxMinutaReal " & fpLongInteger1(0).Value & ", " & fpLongInteger1(1).Value & ", '" & Format(FpFecDesde(0).text, "yyyymmdd") & "', '" & Trim(LimpiaDato(fpText(0).text)) & "'")
   
   End If

If Not RS.EOF Then

    '---> Carga vector receta
    EstVect = True
    
    ReDim vReceta(RS.RecordCount, 2)
    i = 1
    Do While Not RS.EOF
      
       vReceta(i, 1) = RS!rec_codigo
       vReceta(i, 2) = IIf(Option1(0).Value = True, RS!rec_nombre, RS!rec_nomfan)
       
       RS.MoveNext
       i = i + 1
       
    Loop
   
    RS.MoveFirst
    
    Do While Not RS.EOF
    
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       vaSpread1.Row = vaSpread1.MaxRows
       
       vaSpread1.Col = 1
       vaSpread1.text = "0"
       
       vaSpread1.Col = 2
       vaSpread1.text = RS!rec_codigo
       
       vaSpread1.Col = 3
       vaSpread1.text = IIf(Option1(0).Value = True, RS!rec_nombre, RS!rec_nomfan)
       
       vaSpread1.Col = 4
       vaSpread1.text = RS!rec_nomfan
    
       vaSpread1.Col = 5
       vaSpread1.text = RS!CatDietetica
    
       vaSpread1.Col = 6
       vaSpread1.text = RS!tippla
       
       vaSpread1.Col = 7
       vaSpread1.text = RS!red_tiprec
       
       vaSpread1.Col = 8
       vaSpread1.text = RS!red_cencos
       
       vaSpread1.Col = 9
       vaSpread1.text = RS!canservida
       
       vaSpread1.Col = 10
       vaSpread1.text = RS!rec_basrac
       
'       vaSpread1.Col = 11
'       vaSpread1.text = RS!rec_codigo
       
       '-------> Mover receta asociada
       If EstVect Then
         
          lisnom = ""
          liscod = ""
          For z = 1 To UBound(vReceta)
             
              lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vReceta(z, 2))
             
              liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vReceta(z, 1)
             
              vaSpread1.Col = 11
              vaSpread1.TypeComboBoxList = lisnom
             
              vaSpread1.Col = 12
              vaSpread1.TypeComboBoxList = liscod
         
          Next z

'          vaSpread1.Col = 10
'          codaux = -1
'          For z = 0 To vaSpread1.TypeComboBoxCount
'
'              vaSpread1.TypeComboBoxCurSel = z
'              If vaSpread1.text = IIf(IsNull(RS1!IdServicio), 0, RS1!IdServicio) Then codaux = z: Exit For
'              codaux = -1
'
'          Next z
'
'          vaSpread1.Col = 9
'          vaSpread1.TypeComboBoxCurSel = codaux
'
'          vaSpread1.Col = 11
'          vaSpread1.text = IIf(IsNull(RS1!ser_NombreFantasia), "", Trim(RS1!ser_NombreFantasia))
      
       End If

       vaSpread1.Col = 13
       vaSpread1.text = 0

       RS.MoveNext
    
    Loop

Else

   fg_descarga
   MsgBox "No existe Información en la minuta real " & VgLinea & "o bien no esta definido los sellos en las recetas...", vbExclamation + vbOKOnly, MsgTitulo

End If
RS.Close
Set RS = Nothing

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(Text1(Index).text, ",")

If Index = 2 Then
   
   Text1(3).text = ""

ElseIf Index = 3 Then
   
   Text1(2).text = ""

End If

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 13
    vaSpread1.text = 0
    
Next

Select Case Index

Case 2, 3
    
    vaSpread1.Visible = False
    
    If Trim(Text1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like IIf(Index = 2, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
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
        vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(Text1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 13
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
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim i           As Long
Dim j           As Long
Dim iselecc     As Long
Dim Sql         As String
Dim MyBuffer    As String
Dim MyBufferNut As String
Dim CodRec      As Long
Dim nomrec      As String
Dim tiprec      As Long
Dim Ceco        As String
Dim CecoPar     As String
Dim ContSele    As Long
Dim CSer        As Double
Dim Por         As Long
Dim RecL        As Long
Dim Nut         As Long
Dim CecRec      As String

Select Case Button.Index
    
    Case 1
        
        If Not ValidarOpciones Then
        
           Exit Sub
           
        End If
        
        If vaSpread1.MaxRows < 1 Then
        
           MsgBox "No existen datos Imprimir...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        '-------
        '------- Validar seleccion de receta
        '-------
        iselecc = 0
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
            
               iselecc = 1
               Exit For
            
            End If
            
        Next i
        
        If iselecc = 0 Then
        
           MsgBox "Debe seleccionar a lo menos una receta", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        '-------
        '------- Validar union de receta este seleccionado
        '-------
        Dim CodigoReceta   As Long
        Dim LineaSeleccion As String
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
            
               vaSpread1.Col = 12
               CodigoReceta = IIf(Trim(vaSpread1.text) = "" Or Val(vaSpread1.text) = 0, 0, vaSpread1.text)
               
               If CodigoReceta > 0 Then
               
                  For j = 1 To vaSpread1.MaxRows
                   
                      vaSpread1.Row = j
                   
                      vaSpread1.Col = 1
                      LineaSeleccion = vaSpread1.text
                   
                      vaSpread1.Col = 2
                   
                      If i <> j And CodigoReceta = Val(vaSpread1.text) And Val(vaSpread1.text) > 0 And LineaSeleccion = "0" Then
                   
                         MsgBox "Debe seleccionar la receta origen asociada, proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo
                         Exit Sub
                   
                      End If
               
                  Next j
                  
               End If
            End If
            
        Next i
        
        '-------
        '------- Traer Parametro solido receta
        '-------
        Dim ParametroSolido As Long
        
        ParametroSolido = GetParametro("ParSoliRec")
        
        If ParametroSolido = 0 Then
        
           MsgBox "No existe código parametro solido receta, en tabla a_param" & VgLinea & "intentelo dentro de una hora...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        '-------
        '------- Traer Parametro liquido receta
        '-------
        Dim ParametroLiquido As Long
        
        ParametroLiquido = GetParametro("ParLiquRec")
        
        If ParametroLiquido = 0 Then
        
           MsgBox "No existe código parametro liquido receta, en tabla a_param" & VgLinea & "intentelo dentro de una hora...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        '-------
        '------- Traer Parametro nutrientes gramos totales
        '-------
        Dim ParametroNutGra As String
        
        ParametroNutGra = GetParametro("ParNutTGra")
        
        If Trim(ParametroNutGra) = "" Then
        
           MsgBox "No existe códigos parametro nutriente gramps totales receta, en tabla a_param" & VgLinea & "intentelo dentro de una hora...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        '-------
        '------- Traer Parametro colesterol
        '-------
        Dim ParColeste As String
        
        ParColeste = GetParametro("ParColeste")
        
        If Trim(ParColeste) = "" Then
        
           MsgBox "No existe códigos parametro nutriente colesterol, en tabla a_param" & VgLinea & "intentelo dentro de una hora...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        '-------
        '------- Traer Parametro % colesterol
        '-------
        Dim ParColPor As String
        
        ParColPor = GetParametro("ParColPor")
        
        If Trim(ParColPor) = "" Then
        
           MsgBox "No existe parametro % colesterol, en tabla a_param" & VgLinea & "intentelo dentro de una hora...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        '-------
        '------- Traer Parametro de porcion receta
        '-------
        Dim MaxEtiquetadoReceta As Long
        
        MaxEtiquetadoReceta = GetParametro("ParMaxEtRe")
        
        If MaxEtiquetadoReceta = 0 Then
        
           MsgBox "No existe código parametro maximo receta, en tabla a_param" & VgLinea & "intentelo dentro de una hora...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        '-------
        '------- Validar maximo de union de receta
        '-------
        Dim CodigoUnionReceta As Long
        Dim ContReceta        As Long
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
            
               vaSpread1.Col = 2
               CodigoUnionReceta = IIf(Trim(vaSpread1.text) = "" Or Val(vaSpread1.text) = 0, 0, vaSpread1.text)
               ContReceta = 0
               
               For j = 1 To vaSpread1.MaxRows
                   
                   vaSpread1.Row = j
                   vaSpread1.Col = 1
                   
                   If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
                   
                      vaSpread1.Col = 12
                   
                      If i <> j And CodigoUnionReceta = Val(vaSpread1.text) And Val(vaSpread1.text) > 0 And vaSpread1.TypeComboBoxCurSel <> -1 Then
                   
                         ContReceta = ContReceta + 1
                   
                      End If
                   
                   End If
                   
               Next j
                           
               If ContReceta > MaxEtiquetadoReceta Then
               
                  MsgBox "Debe seleccionar un maximo " & MaxEtiquetadoReceta & " receta asociada, proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo
                  Exit Sub
               
               End If
               
            End If
            
        Next i
        
        '-------
        '------- Traer Parametro de porcion receta
        '-------
        Dim MaxPorcionReceta As Long
        
        MaxPorcionReceta = GetParametro("ParMaxPorR")
        
        If MaxPorcionReceta = 0 Then
        
           MsgBox "No existe código parametro maximo porción receta, en tabla a_param" & VgLinea & "intentelo dentro de una hora...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        '-------
        '------- Armar Xml Receta
        '-------
        Let MyBuffer = ""
        Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        Let MyBuffer = MyBuffer & "<Receta>"
         
        ContSele = 0
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
            
               vaSpread1.Col = 2
               CodRec = vaSpread1.text
               
               vaSpread1.Col = 3
               nomrec = Trim(vaSpread1.text)
               
               nomrec = Replace(Trim(nomrec), Chr(34), "&quot;")
               nomrec = Replace(Trim(nomrec), Chr(38), "&amp;")
               nomrec = Replace(Trim(nomrec), Chr(39), "&apos;")
               nomrec = Replace(Trim(nomrec), Chr(60), "&lt;")
               nomrec = Replace(Trim(nomrec), Chr(62), "&gt;")
               
               vaSpread1.Col = 7
               tiprec = vaSpread1.text
               
               vaSpread1.Col = 8
               CecRec = Trim(vaSpread1.text)
               
               vaSpread1.Col = 9
               CSer = vaSpread1.text
               
               If CSer <= 0 Then
               
                  MsgBox "Existe cantidad servida con valor cero en la grilla, proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo
                  Exit Sub
               
               End If
               
               vaSpread1.Col = 10
               Por = vaSpread1.text
               
               If Por < 1 Then
               
                  MsgBox "Existe porción con valor cero en la grilla, proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo
                  Exit Sub
               
               End If
               
               If Por > MaxPorcionReceta Then
               
                  MsgBox "Debe seleccionar un maximo " & MaxPorcionReceta & " porción de receta, proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo
                  Exit Sub
               
               End If
               
               vaSpread1.Col = 12
               RecL = IIf(Trim(vaSpread1.text) = "", CodRec, vaSpread1.text) 'vaSpread1.text
               
               If RecL < 1 Then
               
                  MsgBox "Existe código de receta unión con valor cero en la grilla, proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo
                  Exit Sub
               
               End If
               
               MyBuffer = MyBuffer & " <Rec"
               MyBuffer = MyBuffer & " Rec = " & Chr(34) & CodRec & Chr(34)
               MyBuffer = MyBuffer & " NRec = " & Chr(34) & nomrec & Chr(34)
               MyBuffer = MyBuffer & " Tip = " & Chr(34) & tiprec & Chr(34)
               MyBuffer = MyBuffer & " CSer = " & Chr(34) & CSer & Chr(34)
               MyBuffer = MyBuffer & " Por = " & Chr(34) & Por & Chr(34)
               MyBuffer = MyBuffer & " Recl = " & Chr(34) & RecL & Chr(34)
               MyBuffer = MyBuffer & " Ceco = " & Chr(34) & CecRec & Chr(34)
               MyBuffer = MyBuffer & "/>"
               
               ContSele = ContSele + 1
                          
               If ContSele > 1001 Then
               
                  MsgBox "Debe seleccionar un maximo mil receta, proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo
                  Exit Sub
                  
               End If
               
            End If
        
        Next i
         
        MyBuffer = MyBuffer & "</Receta>"
        Sql = ""
        Sql = MyBuffer
               
        '-------
        '------- Armar Xml Nutriente
        '-------
        
        Let MyBufferNut = ""
        Let MyBufferNut = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        Let MyBufferNut = MyBufferNut & "<Nutriente>"
        
        Dim ContNutrientes As Long
        Dim MaxNutrientes  As Long
        ContNutrientes = 0
        
        '-------
        '------- Traer Parametro de nutriente sello
        '-------
        MaxNutrientes = GetParametro("ParMaxNutr")
        
        If MaxNutrientes = 0 Then
        
           MsgBox "No existe código parametro maximo nutriente, en tabla a_param" & VgLinea & "intentelo dentro de una hora...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        For i = 1 To vaSpread2.MaxRows
            
            vaSpread2.Row = i
            vaSpread2.Col = 1
            
            If vaSpread2.text = "1" And vaSpread2.RowHidden = False Then
               
               vaSpread2.Col = 3
               Nut = vaSpread2.text

               MyBufferNut = MyBufferNut & " <Nut"
               MyBufferNut = MyBufferNut & " Nut = " & Chr(34) & Nut & Chr(34)
               MyBufferNut = MyBufferNut & "/>"
                              
               ContNutrientes = ContNutrientes + 1
            
            End If
        
            If ContNutrientes > MaxNutrientes Then
            
               MsgBox "Debe seleccionar un maximo " & MaxNutrientes & " nutrientes, proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo
               Exit Sub
            
            End If
            
        Next i
         
        MyBufferNut = MyBufferNut & "</Nutriente>"
        
        CecoPar = fpText(0).text
        Ceco = fpText(0).text
        
        I_Etiquetado_Receta Sql, MyBufferNut, Ceco, CecoPar, Azucares, Calorias, Grasas, Sodio, Format(FpFecDesde(1).text, "dd/mm/yyyy"), Logo
        
    Case 3
        
        Me.Hide
        Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

If Not ValidarOpciones Then

   Exit Sub
   
End If

MoverListadoRecetaGrilla (0)

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Function ValidarOpciones() As Boolean

On Error GoTo Man_Error

ValidarOpciones = True

If FpFecDesde(0).text = "" Then

   MsgBox "Fecha esta nula o en blanco...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarOpciones = False
   Exit Function

End If

If Trim(fpayuda(0).Caption) = "" Or Trim(fpText(0).text) = "" Then

   MsgBox "Contrato no definido...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarOpciones = False
   Exit Function

End If

If Trim(fpayuda(1).Caption) = "" Or Trim(fpLongInteger1(0).Value) = "" Then

   MsgBox "Régimen no definido...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarOpciones = False
   Exit Function

End If

If Trim(fpayuda(2).Caption) = "" Or Trim(fpLongInteger1(1).Value) = "" Then

   MsgBox "Servicio no definido...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarOpciones = False
   Exit Function

End If

'validar carpeta y archivo de sellos
If Dir(dir_trabajo_Inf & "\" & "Etiquetado", vbDirectory) = "" Then

   MsgBox "No existe la carpeta Etiquetado...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarOpciones = False
   Exit Function
  
End If

Dim fso As Object
'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
Set fso = CreateObject("Scripting.FileSystemObject")

'validar si existe Calorias
If Not fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Calorias) Then

   MsgBox "No existe archivo " & Calorias & " o bien fue borrado...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarOpciones = False
   Exit Function

End If

'validar si existe Azucares
If Not fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Azucares) Then

   MsgBox "No existe archivo " & Azucares & " o bien fue borrado...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarOpciones = False
   Exit Function

End If

'validar si existe Grasas
If Not fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Grasas) Then

   MsgBox "No existe archivo " & Grasas & " o bien fue borrado...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarOpciones = False
   Exit Function

End If

'validar si existe Sodio
If Not fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Sodio) Then

   MsgBox "No existe archivo " & Sodio & " o bien fue borrado...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarOpciones = False
   Exit Function

End If

'validar si existe Logo
If Not fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Logo) Then

   MsgBox "No existe archivo " & Logo & " o bien fue borrado...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarOpciones = False
   Exit Function

End If

Set fso = Nothing

Exit Function
Man_Error:

Set fso = Nothing
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Function

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

est = True

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

est = False

Exit Sub
Man_Error:
    est = False
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim indice As Long

Select Case Col

    Case 11
    
        vaSpread1.Row = Row
        vaSpread1.Col = 11
        indice = vaSpread1.TypeComboBoxCurSel
        
        vaSpread1.Col = 12
        vaSpread1.TypeComboBoxCurSel = indice
    
End Select

Exit Sub
Man_Error:
    est = False
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub

Select Case KeyCode

Case 46
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    
    If vaSpread1.Col <> 11 Then Exit Sub
    vaSpread1.text = ""
    vaSpread1.TypeComboBoxCurSel = -1
    
    vaSpread1.Col = 12
    vaSpread1.text = ""
    vaSpread1.TypeComboBoxCurSel = -1
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

