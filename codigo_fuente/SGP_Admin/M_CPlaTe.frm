VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_CPlaTe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Planificación Teórica"
   ClientHeight    =   7680
   ClientLeft      =   2880
   ClientTop       =   1875
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7680
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos Origen"
      ForeColor       =   &H80000008&
      Height          =   1965
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   0
      Width           =   7290
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   1
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   175
         Width           =   1800
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Top             =   925
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Top             =   1270
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
         Left            =   1440
         TabIndex        =   4
         Top             =   1620
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
         DateCalcMethod  =   3
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
         Left            =   5250
         TabIndex        =   5
         Top             =   1620
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
         DateCalcMethod  =   3
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
         Index           =   5
         Left            =   1440
         TabIndex        =   1
         Top             =   580
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2385
         Picture         =   "M_CPlaTe.frx":0000
         Top             =   1170
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2385
         Picture         =   "M_CPlaTe.frx":030A
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2385
         Picture         =   "M_CPlaTe.frx":0614
         Top             =   475
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Subsegmento"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   44
         Top             =   655
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   43
         Top             =   995
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   42
         Top             =   1360
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   41
         Top             =   1650
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Index           =   6
         Left            =   4065
         TabIndex        =   40
         Top             =   1650
         Width           =   1005
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   2880
         TabIndex        =   39
         Top             =   580
         Width           =   4110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   2880
         TabIndex        =   38
         Top             =   925
         Width           =   4110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   2880
         TabIndex        =   37
         Top             =   1270
         Width           =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   8
         Left            =   2880
         TabIndex        =   36
         Top             =   1650
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   9
         Left            =   6645
         TabIndex        =   35
         Top             =   1650
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Minuta"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   34
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2925
         TabIndex        =   47
         Top             =   1315
         Width           =   4110
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2925
         TabIndex        =   46
         Top             =   970
         Width           =   4110
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2925
         TabIndex        =   45
         Top             =   630
         Width           =   4110
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos Destino"
      ForeColor       =   &H80000008&
      Height          =   1965
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   2000
      Width           =   7290
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   175
         Width           =   1800
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   2
         Left            =   1440
         TabIndex        =   8
         Top             =   925
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   3
         Left            =   1440
         TabIndex        =   9
         Top             =   1270
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
         Index           =   2
         Left            =   1440
         TabIndex        =   10
         Top             =   1620
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
         DateCalcMethod  =   3
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
         Index           =   3
         Left            =   5250
         TabIndex        =   11
         Top             =   1620
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
         DateCalcMethod  =   3
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
         Index           =   4
         Left            =   1440
         TabIndex        =   7
         Top             =   580
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   2385
         Picture         =   "M_CPlaTe.frx":091E
         Top             =   1190
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   2385
         Picture         =   "M_CPlaTe.frx":0C28
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2385
         Picture         =   "M_CPlaTe.frx":0F32
         Top             =   475
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Subsegmento"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   29
         Top             =   655
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   28
         Top             =   995
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   27
         Top             =   1360
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   26
         Top             =   1650
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Index           =   7
         Left            =   4065
         TabIndex        =   25
         Top             =   1650
         Width           =   1005
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   2880
         TabIndex        =   24
         Top             =   580
         Width           =   4110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   2880
         TabIndex        =   23
         Top             =   925
         Width           =   4110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   2880
         TabIndex        =   22
         Top             =   1270
         Width           =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   10
         Left            =   2880
         TabIndex        =   21
         Top             =   1650
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   11
         Left            =   6650
         TabIndex        =   20
         Top             =   1650
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Minuta"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   19
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   2925
         TabIndex        =   32
         Top             =   1315
         Width           =   4110
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   2925
         TabIndex        =   31
         Top             =   970
         Width           =   4110
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   2925
         TabIndex        =   30
         Top             =   630
         Width           =   4110
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estructura Servicio Origen && Destino"
      Height          =   3585
      Left            =   30
      TabIndex        =   17
      Top             =   3960
      Width           =   7305
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2985
         Left            =   180
         TabIndex        =   12
         Top             =   360
         Width           =   6915
         _Version        =   393216
         _ExtentX        =   12197
         _ExtentY        =   5265
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   10
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_CPlaTe.frx":123C
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   1560
      ScaleHeight     =   975
      ScaleWidth      =   4785
      TabIndex        =   14
      Top             =   7320
      Visible         =   0   'False
      Width           =   4845
      Begin MSComctlLib.ProgressBar gauge 
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Procesando Necedidad De Insumos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   4515
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7680
      Left            =   7485
      TabIndex        =   13
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   13547
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_CPlaTe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim indsel As Long
Dim vg_AuxIndppr As String

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

Me.HelpContextID = vg_OpcM
fg_centra Me
fpDateTime1(0).CalFirstDay (1)
EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)
EspFecha fpDateTime1(2)
EspFecha fpDateTime1(3)
MsgTitulo = "Copiar Planificación Teórica"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
Label1(8).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(0).text, "yyyymmdd"), 1), 1, 4)
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
Label1(9).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(1).text, "yyyymmdd"), 1), 1, 4)
fpDateTime1(2).text = Format(Date, "dd/mm/yyyy")
Label1(10).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(2).text, "yyyymmdd"), 1), 1, 4)
fpDateTime1(3).text = Format(Date, "dd/mm/yyyy")
Label1(11).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(3).text, "yyyymmdd"), 1), 1, 4)
vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &HC0FFFF
vg_AuxIndppr = vg_Indppr
OpUsuario = vg_IndpprSelec
If IsNull(OpUsuario) Or Trim(OpUsuario) = "" Then
    MsgBox "Contactese con el Administrador del Sistema...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
Else
    Me.HelpContextID = 1110010
    If Mid(ValidaPerfil(Me), 1, 1) = "1" Then
       vg_Indppr = 3
       Combo2(0).Clear
       Combo2(0).AddItem "Real" & Space(150) & "(1)"
       Combo2(0).AddItem "Propuesta" & Space(150) & "(2)"
       Combo2(0).ListIndex = 0
       
       Combo2(1).Clear
       Combo2(1).AddItem "Real" & Space(150) & "(1)"
       Combo2(1).AddItem "Propuesta" & Space(150) & "(2)"
       Combo2(1).ListIndex = 0
    Else
        Select Case OpUsuario
        Case "1"
            Combo2(1).Clear
            Combo2(1).AddItem "Real" & Space(150) & "(1)"
            Combo2(1).ListIndex = fg_buscacbo(Combo2, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
            Combo2(0).AddItem "Real" & Space(150) & "(1)"
            Combo2(0).AddItem "Propuesta" & Space(150) & "(2)"
            Combo2(0).ListIndex = fg_buscacbo(Combo2, 0, 1, fg_pone_cero(Str(OpUsuario), 1))
        Case "2"
            Combo2(1).Clear
            Combo2(1).AddItem "Propuesta" & Space(150) & "(2)"
            Combo2(1).ListIndex = fg_buscacbo(Combo2, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
            Combo2(0).AddItem "Propuesta" & Space(150) & "(2)"
            Combo2(0).ListIndex = fg_buscacbo(Combo2, 0, 1, fg_pone_cero(Str(OpUsuario), 1))
        End Select
    End If
End If
Me.HelpContextID = vg_OpcM
End Sub

Private Sub Form_Unload(Cancel As Integer)
vg_Indppr = vg_AuxIndppr
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
Select Case Index
Case 0
    MoverVector
    Label1(8).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(0).text, "yyyymmdd"), 1), 1, 4)
Case 1
    Label1(9).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(1).text, "yyyymmdd"), 1), 1, 4)
Case 2
    Label1(10).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(2).text, "yyyymmdd"), 1), 1, 4)
Case 3
    Label1(11).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(3).text, "yyyymmdd"), 1), 1, 4)
End Select
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Dim RS As New ADODB.Recordset
Select Case Index
Case 5
    RS.Open "SELECT * FROM a_subsegmento With(NoLock) WHERE sub_codigo=" & Val(fpLongInteger1(5).Value) & " AND sub_indppr = '" & fg_codigocbo(Combo2, 1, 1, "") & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": fpLongInteger1(0).Value = "": fpayuda(1).Caption = "": fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing
    fpLongInteger1(0).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(1).Value = "": fpayuda(2).Caption = ""
    MoverVector
Case 0
    If Val(fpLongInteger1(0).Value) < 1 Then fpayuda(1).Caption = "": Exit Sub
    RS.Open "SELECT * FROM a_regimen With(NoLock) WHERE reg_codigo=" & Val(fpLongInteger1(0).Value) & " AND reg_indppr = '" & fg_codigocbo(Combo2, 1, 1, "") & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    MoverVector
Case 1
    If Val(fpLongInteger1(1).Value) < 1 Then fpayuda(2).Caption = "": Exit Sub
    RS.Open "SELECT * FROM a_servicio With(NoLock) WHERE ser_codigo=" & Val(fpLongInteger1(1).Value) & " AND ser_indppr = '" & fg_codigocbo(Combo2, 1, 1, "") & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "": Exit Sub
    fpayuda(2).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    MoverVector
Case 2
    If Val(fpLongInteger1(2).Value) < 1 Then fpayuda(4).Caption = "": Exit Sub
    RS.Open "SELECT * FROM a_regimen With(NoLock) WHERE reg_codigo=" & Val(fpLongInteger1(2).Value) & " AND reg_indppr = '" & fg_codigocbo(Combo2, 0, 1, "") & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(4).Caption = "": Exit Sub
    fpayuda(4).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
Case 3
    If Val(fpLongInteger1(3).Value) < 1 Then fpayuda(5).Caption = "": Exit Sub
    RS.Open "SELECT * FROM a_servicio With(NoLock) WHERE ser_codigo=" & Val(fpLongInteger1(3).Value) & " AND ser_indppr = '" & fg_codigocbo(Combo2, 0, 1, "") & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(5).Caption = "": Exit Sub
    fpayuda(5).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    MoverVector
Case 4
    If fpLongInteger1(4).Value = "" Then fpayuda(3).Caption = "": Exit Sub
    RS.Open "SELECT * FROM a_subsegmento With(NoLock) WHERE sub_codigo=" & Val(fpLongInteger1(4).Value) & " AND sub_indppr = '" & fg_codigocbo(Combo2, 0, 1, "") & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(3).Caption = "": fpLongInteger1(2).Value = "": fpayuda(4).Caption = "": fpLongInteger1(3).Value = "": fpayuda(5).Caption = "": Exit Sub
    fpayuda(3).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing
    fpLongInteger1(2).Value = "": fpayuda(4).Caption = ""
    fpLongInteger1(3).Value = "": fpayuda(5).Caption = ""
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
    If Index = 2 Then Image1_Click 2
    If Index = 3 Then Image1_Click 3
    If Index = 4 Then Image1_Click 4
    If Index = 5 Then Image1_Click 5
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
    fpLongInteger1(5).Value = Val(vg_codigo)
'    fpayuda(0).Caption = vg_nombre
    fpLongInteger1(0).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(1).Value = "": fpayuda(2).Caption = ""
    fpLongInteger1(0).SetFocus
Case 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
'    fpayuda(1).Caption = vg_nombre
    fpLongInteger1(1).SetFocus
Case 2
    vg_left = fpayuda(2).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Ser"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
'    fpayuda(2).Caption = vg_nombre
    fpDateTime1(0).SetFocus
Case 3
    vg_left = fpayuda(3).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_subsegmento", "sub_", "Subsegmento", "Sub"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(4).Value = Val(vg_codigo)
'    fpayuda(3).Caption = vg_nombre
    fpLongInteger1(2).Value = "": fpayuda(4).Caption = ""
    fpLongInteger1(3).Value = "": fpayuda(5).Caption = ""
    fpLongInteger1(2).SetFocus
Case 4
    vg_left = fpayuda(4).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
'    fpayuda(4).Caption = vg_nombre
    fpLongInteger1(3).SetFocus
Case 5
    vg_left = fpayuda(5).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Ser"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(3).Value = Val(vg_codigo)
'    fpayuda(5).Caption = vg_nombre
    fpDateTime1(2).SetFocus
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim fecori1 As Long, fecori2 As Long, fecdes1 As Long, fecdes2 As Long, vdia As Long, indice As Long, tiprec As Long, codest1 As Long, codest2 As Long
Dim auxfeco As String, auxfecd As String, vaux1 As Long, vaux2 As Long, diatop As Long, Est As Boolean, cSpi As Long
Dim tipsus As String, tipreg As String, tipser As String
Dim Resp As String
On Error GoTo Man_Error
Select Case Button.Index
Case 2
    If vaSpread1.MaxRows < 1 Then MsgBox "No existe concepto estructuras en datos origen, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Me.HelpContextID = 1110010
    If Mid(ValidaPerfil(Me), 1, 1) <> "1" Then
       Me.HelpContextID = vg_OpcM
       If Replace(Combo2(1).text, " ", "", , , vbTextCompare) = "Propuesta(2)" And Replace(Combo2(0).text, " ", "", , , vbTextCompare) = "Real(1)" Then MsgBox "No es posible realizar copia de Propuesta a Real, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    End If
    Me.HelpContextID = vg_OpcM
    Est = False
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 4: estado = vaSpread1.TypeComboBoxCurSel
        vaSpread1.Col = 1: If vaSpread1.text = "1" And estado = -1 Then MsgBox "Falta seleccionar un concepto estructuras Destino...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub 'Est = True
    Next i
'    If Est Then MsgBox "Falta seleccionar concepto estructuras...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub ' Comenté esta linea, para evitar que termine el proceso si no han seleccionado todos los item Samuel 02/09/09
'20110902
    Est = False
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1: If vaSpread1.text = "1" Then Est = True: Exit For
    Next i
    If Not Est Then MsgBox "Falta seleccionar concepto estructuras...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub ' Comenté esta linea, para evitar que termine el proceso si no han seleccionado todos los item Samuel 02/09/09
'20110902
    '-------> Validar datos origen
    RS.Open "SELECT * FROM a_subsegmento With(NoLock) WHERE sub_codigo=" & Val(fpLongInteger1(5).Value) & " AND sub_indppr = '" & fg_codigocbo(Combo2, 1, 1, "") & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(5).Value = "": fpayuda(0).Caption = "": fpLongInteger1(0).Value = "": fpayuda(1).Caption = "": fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": MsgBox "No existe sub-segmento o bien sub-segmento no corresponde tipo origen ", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    tipsus = RS!sub_indppr
    RS.Close: Set RS = Nothing
    RS.Open "SELECT * FROM a_regimen With(NoLock) WHERE reg_codigo=" & Val(fpLongInteger1(0).Value) & " AND reg_indppr = '" & fg_codigocbo(Combo2, 1, 1, "") & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(0).Value = "": fpayuda(1).Caption = "": MsgBox "No existe regimen o bien regimen no corresponde tipo origen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    tipreg = RS!reg_indppr
    RS.Close: Set RS = Nothing
    RS.Open "SELECT * FROM a_servicio With(NoLock) WHERE ser_codigo=" & Val(fpLongInteger1(1).Value) & " AND ser_indppr = '" & fg_codigocbo(Combo2, 1, 1, "") & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set ConSql = Nothing: fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": MsgBox "No existe servicio o bien servicio no corresponde tipo origen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    tipser = RS!ser_indppr
    RS.Close: Set RS = Nothing
    '-------> validar tipo planificación origen
    If fg_codigocbo(Combo2, 1, 1, "") <> tipsus Or fg_codigocbo(Combo2, 1, 1, "") <> tipreg Or fg_codigocbo(Combo2, 1, 1, "") <> tipser Then MsgBox "Tipo planificación origen, no coincide con los códigos Sub-Segmento, Regimen o Servicio ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Val(fpLongInteger1(5).Value) = Val(fpLongInteger1(4).Value) And Val(fpLongInteger1(0).Value) = Val(fpLongInteger1(2).Value) And Val(fpLongInteger1(1).Value) = Val(fpLongInteger1(3).Value) And fpDateTime1(0).text = fpDateTime1(2).text And fpDateTime1(1).text = fpDateTime1(3).text And Val(fg_codigocbo(Combo2, 1, 1, "")) = Val(fg_codigocbo(Combo2, 0, 1, "")) Then: MsgBox "Datos origen, beben ser distinto datos destino", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    If fpDateTime1(0).text = "" Or fpDateTime1(1).text = "" Or fpDateTime1(2).text = "" Or fpDateTime1(3).text = "" Then MsgBox "Fecha no definida", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If (Val(Mid(fpDateTime1(1).text, 1, 2)) - Val(Mid(fpDateTime1(0).text, 1, 2))) > (Val(Mid(fpDateTime1(3).text, 1, 2)) - Val(Mid(fpDateTime1(2).text, 1, 2))) Then MsgBox "Fecha origen supera nº dìas", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Val(Format(fpDateTime1(0).text, "ddmmyyyy")) > Val(Format(fpDateTime1(1).text, "ddmmyyyy")) Or Val(Format(fpDateTime1(2).text, "ddmmyyyy")) > Val(Format(fpDateTime1(3).text, "ddmmyyyy")) Then MsgBox "Fecha no coincide", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Val(Format(fpDateTime1(0).text, "ddmmyyyy")) > Val(Format(fpDateTime1(1).text, "ddmmyyyy")) Or Val(Format(fpDateTime1(2).text, "ddmmyyyy")) > Val(Format(fpDateTime1(3).text, "ddmmyyyy")) Then MsgBox "Fecha no coincide", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Val(Format(fpDateTime1(0).text, "mm")) <> Val(Format(fpDateTime1(1).text, "mm")) Or Val(Format(fpDateTime1(2).text, "mm")) <> Val(Format(fpDateTime1(3).text, "mm")) Then MsgBox "Fecha no coincide", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If (Val(Mid(fpDateTime1(1).text, 1, 2)) - Val(Mid(fpDateTime1(0).text, 1, 2))) > (Val(Mid(fpDateTime1(3).text, 1, 2)) - Val(Mid(fpDateTime1(2).text, 1, 2))) Then MsgBox "Fecha origen supera nº dìas", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    '-------> Validar datos destino
    RS.Open "SELECT * FROM a_subsegmento With(NoLock) WHERE sub_codigo=" & Val(fpLongInteger1(4).Value) & " AND sub_indppr = '" & fg_codigocbo(Combo2, 0, 1, "") & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(4).Value = "": fpayuda(3).Caption = "": fpLongInteger1(2).Value = "": fpayuda(4).Caption = "": fpLongInteger1(3).Value = "": fpayuda(5).Caption = "": MsgBox "No existe sub-segmento o bien sub-segmento no corresponde tipo destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    tipsus = RS!sub_indppr
    RS.Close: Set RS = Nothing
    RS.Open "SELECT * FROM a_regimen With(NoLock)  WHERE reg_codigo=" & Val(fpLongInteger1(2).Value) & " AND reg_indppr = '" & fg_codigocbo(Combo2, 0, 1, "") & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(2).Value = "": fpayuda(4).Caption = "": MsgBox "No existe regimen o bien regimen no corresponde tipo destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    tipreg = RS!reg_indppr
    RS.Close: Set RS = Nothing
    RS.Open "SELECT * FROM a_servicio With(NoLock)  WHERE ser_codigo=" & Val(fpLongInteger1(3).Value) & " AND ser_indppr = '" & fg_codigocbo(Combo2, 0, 1, "") & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set ConSql = Nothing: fpLongInteger1(3).Value = "": fpayuda(5).Caption = "": MsgBox "No existe servicio o bien servicio no corresponde tipo destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    tipser = RS!ser_indppr
    RS.Close: Set RS = Nothing
    '-------> validar tipo planificación destino
    If fg_codigocbo(Combo2, 0, 1, "") <> tipsus Or fg_codigocbo(Combo2, 0, 1, "") <> tipreg Or fg_codigocbo(Combo2, 0, 1, "") <> tipser Then MsgBox "Tipo planificación destino, no coincide con los códigos Sub-Segmento, Regimen o Servicio ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    '-------> Validar datos destino bloqueado
    fecdes1 = Mid(fpDateTime1(2).text, 7, 4) & Mid(fpDateTime1(2).text, 4, 2)
    RS.Open "SELECT COUNT(a.min_codigo) AS nreg " & _
            "FROM  b_minuta a With(NoLock), b_minutadet b With(NoLock) " & _
            "WHERE a.min_codigo = b.mid_codigo " & _
            "AND   a.min_subseg = " & Val(fpLongInteger1(4).Value) & " " & _
            "AND   a.min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
            "AND   a.min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
            "AND   a.min_indppr = '" & Val(fg_codigocbo(Combo2, 0, 1, "")) & "' " & _
            "AND   substring(convert(char(8),a.min_fecmin),1,6) = " & fecdes1 & " " & _
            "AND   a.min_indblo = 1 " & _
            "AND   b.mid_tipmin = '1'", vg_db, adOpenStatic
    If Not RS.EOF And RS!nReg > 0 Then RS.Close: Set RS = Nothing: MsgBox "Minuta esta bloqueda, proceso cancelado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '-------> Validar si existe datos origen
    fecori1 = Mid(fpDateTime1(0).text, 7, 4) & Mid(fpDateTime1(0).text, 4, 2)
    RS.Open "SELECT DISTINCT min_subseg, min_codreg, min_codser " & _
             "FROM  b_minuta With(NoLock) " & _
             "WHERE min_subseg = " & Val(fpLongInteger1(5).Value) & " " & _
             "AND   min_codreg = " & Val(fpLongInteger1(0).Value) & " " & _
             "AND   min_codser = " & Val(fpLongInteger1(1).Value) & " " & _
             "AND   min_indppr = '" & Val(fg_codigocbo(Combo2, 1, 1, "")) & "' " & _
             "AND   substring(convert(char(8),b_minuta.min_fecmin),1,6) = " & fecori1 & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe datos origen, proceso cancelado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '-------> Grabando plantilla casino origen hacia origen
    vdia = 999999: indice = 0: codest = 0
    fecori1 = Format(fpDateTime1(0).text, "yyyymmdd")
    fecori2 = Format(fpDateTime1(1).text, "yyyymmdd")
    fecdes1 = Format(fpDateTime1(2).text, "yyyymmdd")
    fecdes2 = Format(fpDateTime1(3).text, "yyyymmdd")
    diatop = Format(fpDateTime1(3).text, "yyyymmdd")
    '-------> validar si Existe Datos Destino
    RS.Open "SELECT DISTINCT min_subseg, min_codreg, min_codser " & _
             "FROM  b_minuta With(NoLock) " & _
             "WHERE min_subseg = " & Val(fpLongInteger1(4).Value) & " " & _
             "AND   min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
             "AND   min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
             "AND   min_Indppr = " & Val(fg_codigocbo(Combo2, 0, 1, "")) & " " & _
             "and min_codigo IN (select DISTINCT mid_codigo from b_minutadet With(NoLock) where mid_tipmin = 1) " & _
             "AND   min_fecmin >= " & fecdes1 & " AND min_fecmin <= " & fecdes2 & "", vg_db, adOpenStatic
             Resp = "S"
    If Not RS.EOF Then
            If MsgBox("Existe información casino destino. se borrara la información existente ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
                Resp = "N"
                'RS.Close: Set RS = Nothing
                'Exit Sub

            End If
    End If
    RS.Close: Set RS = Nothing
    '-------> Fin validar si existe datos destino
    RS.Open "SELECT a.*, b.* " & _
            "FROM  b_minuta a With(NoLock), b_minutadet b With(NoLock), b_receta c With(NoLock) " & _
            "WHERE a.min_codigo = b.mid_codigo " & _
            "AND   b.mid_codrec = c.rec_codigo " & _
            "AND  (c.rec_fecvig > " & Format(Date, "yyyymmdd") & " OR c.rec_fecvig <= 0) " & _
            "AND   a.min_subseg = " & Val(fpLongInteger1(5).Value) & " " & _
            "AND   a.min_codreg = " & Val(fpLongInteger1(0).Value) & " " & _
            "AND   a.min_codser = " & Val(fpLongInteger1(1).Value) & " " & _
            "AND   a.min_indppr = '" & Val(fg_codigocbo(Combo2, 1, 1, "")) & "' " & _
            "AND   a.min_fecmin >= " & fecori1 & " " & _
            "AND   a.min_fecmin <= " & fecori2 & " " & _
            "AND   b.mid_tipmin = '1' ORDER BY a.min_fecmin", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    fg_carga ""
    Est = True
    If DatePart("w", fg_Ctod1(RS!min_fecmin), 2) <> DatePart("w", fg_Ctod1(fecdes1), 2) Then
       If MsgBox("No coincide día de la semana. ¿ Desea copiar ? ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then RS.Close: Set RS = Nothing: fg_descarga: Exit Sub Else Est = False
    End If
    
    '-------> Borrar tabla de paso estructura servicio
    vg_db.Execute "DELETE paso_estservicio WHERE ess_spid = @@spid and ess_usr = '" & vg_NUsr & "'"
    '-------> Buscar spid
    Set RS = vg_db.Execute("SELECT @@spid spid")
    If Not RS.EOF Then cSpi = RS!spid
    RS.Close: Set RS = Nothing
    '-------> Grabar tabla de paso estructura servicio
    For i = 1 To vaSpread1.MaxRows
       vaSpread1.Row = i
       vaSpread1.Col = 1
       vaSpread1.Col = 2: codest1 = Val(vaSpread1.text)
       vaSpread1.Col = 4: codest2 = Val(vaSpread1.text)
       vaSpread1.Col = 1
       If vaSpread1.text = "1" And vaSpread1.Row > 0 And codest1 > 0 And codest2 > 0 Then
          vaSpread1.Col = 2: codest1 = Val(vaSpread1.text)
          vaSpread1.Col = 4: codest2 = Val(vaSpread1.text)
          vaSpread1.Col = 5
          vg_db.Execute ("INSERT INTO paso_estservicio (ess_spid, ess_usr, ess_codess1, ess_codess2, ess_desest2) VALUES(" & cSpi & ", '" & vg_NUsr & "', " & codest1 & ", " & codest2 & ", '" & Trim(vaSpread1.text) & "')")
       End If
    Next i
    Toolbar1.Enabled = False
    Frame2.Enabled = False
    Frame1(0).Enabled = False
    Frame1(1).Enabled = False
    vg_db.Execute "sgpadm_p_copiacreaplanif " & Val(fpLongInteger1(5).Value) & ", " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & _
                  "" & Val(fpLongInteger1(4).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & Val(fpLongInteger1(3).Value) & ", " & _
                  "" & fecori1 & " , " & fecori2 & " , " & fecdes1 & ", " & diatop & ", " & IIf(Est, 1, 0) & ", " & cSpi & ", '" & vg_NUsr & "', 0, " & Val(Format(fpDateTime1(3).text, "yyyymm")) & "," & Val(fg_codigocbo(Combo2, 1, 1, "")) & "," & Val(fg_codigocbo(Combo2, 0, 1, "")) & ", '" & Resp & "'"
    
'      For i = 1 To vaSpread1.MaxRows
'       vaSpread1.Row = i
'       vaSpread1.Col = 1
'       If vaSpread1.Text = "1" And vaSpread1.Row > 0 Then
'          vaSpread1.Col = 2: codest1 = vaSpread1.Text
'          vaSpread1.Col = 4: codest2 = vaSpread1.Text
'          vg_db.Execute ("INSERT INTO paso_estservicio VALUES(" & cSpi & ", '" & vg_NUsr & "', " & codest1 & ", " & codest2 & ")")
'       End If
'    Next i
'    vg_db.Execute "sgpadm_iu_CopiMinuEstMod " & Val(fpLongInteger1(5).Value) & ", " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & "," & fecdes1 & "," & diatop & ", " & vaSpread1.Text & ", " & vaSpread1.Text & " "
    Toolbar1.Enabled = True
    Frame2.Enabled = True
    Frame1(0).Enabled = True
    Frame1(1).Enabled = True
    fg_descarga
    Picture1.Visible = False: Label1(5).Visible = False: gauge.Visible = False
    MsgBox "Copia Finalizada Sin Problema", vbInformation + vbOKOnly, MsgTitulo
    If fg_codigocbo(Combo2, 1, 1, "") = "2" And fg_codigocbo(Combo2, 0, 1, "") = "1" Then
       Dim spid As Long
       '-------> Borrar tabla de paso estructura servicio
       vg_db.Execute "DELETE paso_regimen WHERE reg_spid = @@spid and reg_usr = '" & vg_NUsr & "'"
       '-------> Buscar spid
       Set RS = vg_db.Execute("SELECT @@spid spid")
       If Not RS.EOF Then spid = RS!spid
       RS.Close: Set RS = Nothing
       vg_db.Execute "INSERT INTO paso_regimen (reg_spid, reg_usr, reg_codigo) VALUES (" & spid & ", '" & vg_NUsr & "', " & fpLongInteger1(4).Value & ")"
       I_MinutasRealesConRecetasPropuesta Format(fpDateTime1(2).Value, "yyyymmdd"), Format(fpDateTime1(3).Value, "yyyymmdd"), vg_NUsr, spid
    End If
Case 4
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
Toolbar1.Enabled = False
Frame2.Enabled = False
Frame1(0).Enabled = False
Frame1(1).Enabled = False
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then: Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub MoverVector()
Dim RS As New ADODB.Recordset
Dim i As Long
If Trim(fpLongInteger1(5).text) = "" Or Trim(fpLongInteger1(0).text) = "" Or Trim(fpLongInteger1(1).text) = "" Or Trim(fpDateTime1(0).text) = "" Or Trim(fpLongInteger1(3).text) = "" Then Exit Sub
'Borrar este codigo comentado
'RS.Open "SELECT DISTINCT b.mid_estser, c.ess_nombre " & _
'        "FROM a_servicio a, b_minutadet b, a_estservicio c, b_minuta d " & _
'        "WHERE d.min_codigo=b.mid_codigo " & _
'        "AND   d.min_codser=a.ser_codigo " & _
'        "AND   a.ser_codigo=c.ess_codser " & _
'        "AND   c.ess_codigo=b.mid_estser " & _
'        "AND   d.min_subseg=" & Val(fpLongInteger1(5).Value) & " " & _
'        "AND   d.min_codreg=" & Val(fpLongInteger1(0).Value) & " " & _
'        "AND   d.min_codser=" & Val(fpLongInteger1(1).Value) & " " & _
'        "AND substring(convert(char(8),d.min_fecmin),1,6)=" & Val(Format(fpDateTime1(0).Text, "yyyymm")) & "", vg_db, adOpenForwardOnly ', adOpenStatic
'Mover estructura servicio Origen
Set RS = vg_db.Execute("sgpadm_s_CopiaMinuta " & Val(fpLongInteger1(5).Value) & ", " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Val(Format(fpDateTime1(0).text, "yyyymm")) & " ")
vaSpread1.MaxRows = 0
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 2: vaSpread1.text = RS!mid_estser
      vaSpread1.Col = 3: vaSpread1.text = IIf(Trim(RS!mid_desest) <> "", Trim(RS!mid_desest), Trim(RS!ess_nombre))
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
Dim codest As Long, codaux As Long
RS.Open "SELECT DISTINCT ess_codigo, ess_nombre, ess_orden FROM a_estservicio With(NoLock) WHERE ess_codser = " & Val(fpLongInteger1(3).Value) & " ORDER BY ess_orden", vg_db, adOpenForwardOnly ', adOpenStatic
If Not RS.EOF Then
   For i = 1 To vaSpread1.MaxRows
       vaSpread1.Row = i
       vaSpread1.Col = 5
'       If vg_IndpprSelec = 2 Then vaSpread1.TypeComboBoxEditable = True
       vaSpread1.Col = 2: codest = vaSpread1.text
       lisnom = "": liscod = ""
       Do While Not RS.EOF
          vaSpread1.Col = 4: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & RS!ess_codigo
          vaSpread1.Col = 5: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(RS!ess_nombre)
          vaSpread1.Col = 4: vaSpread1.TypeComboBoxList = liscod
          vaSpread1.Col = 5: vaSpread1.TypeComboBoxList = lisnom
          
          RS.MoveNext
       Loop
       RS.MoveFirst
       If fpLongInteger1(1).Value = fpLongInteger1(3).Value Then
          vaSpread1.Col = 4: vaSpread1.TypeComboBoxList = liscod
          For z = 0 To vaSpread1.TypeComboBoxCount
              vaSpread1.TypeComboBoxCurSel = z
              If vaSpread1.text = codest Then codaux = z: Exit For
              codaux = -1
          Next z
          vaSpread1.Col = 5: vaSpread1.TypeComboBoxCurSel = codaux
       End If
   Next i
End If
RS.Close: Set RS = Nothing
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
Select Case BlockCol
Case 1
    Dim i As Long
    vaSpread1.Col = 1
    For i = BlockRow To BlockRow2
        vaSpread1.Row = i
        vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    Next
End Select
End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Select Case Col
Case 5
    Dim indice As Long
    vaSpread1.Row = Row
    vaSpread1.Col = 5: indice = vaSpread1.TypeComboBoxCurSel
    vaSpread1.Col = 4: vaSpread1.TypeComboBoxCurSel = indice
'    MsgBox indice
'    If vg_IndpprSelec = "2" Then vaSpread1.Col = 5: vaSpread1.TypeComboBoxEditable = True
'    If fg_codigocbo(Combo2, 0, 1, "") = "2" Then vaSpread1.Col = 5: vaSpread1.TypeComboBoxEditable = True

End Select
End Sub

'Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
'vaSpread1.Row = Row
'Select Case Col
'Case 5
'If ChangeMade = False Then
'    vaSpread1.Col = 5
'    If Trim(vaSpread1.Text) = "" And indsel > 0 Then
'       vaSpread1.Col = 5: vaSpread1.TypeComboBoxCurSel = indsel
'       Exit Sub
'    End If
'End If
'End Select
'End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim IndCol As Long
IndCol = vaSpread1.ActiveCol
Select Case KeyCode
Case 46 And IndCol = 5
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 4: vaSpread1.TypeComboBoxCurSel = -1
    vaSpread1.Col = 5: vaSpread1.TypeComboBoxCurSel = -1
End Select
End Sub
