VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_LibMsr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liberar Minuta Bloque"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos Origen"
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
      Height          =   1725
      Index           =   0
      Left            =   30
      TabIndex        =   29
      Top             =   0
      Width           =   7650
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Top             =   520
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
         Top             =   870
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
         Top             =   1290
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
         AllowNull       =   -1  'True
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
         Text            =   ""
         DateCalcMethod  =   3
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   5250
         TabIndex        =   5
         Top             =   1290
         Visible         =   0   'False
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
         Left            =   3240
         TabIndex        =   1
         Top             =   1380
         Visible         =   0   'False
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   200
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2865
         Picture         =   "M_LibMsr.frx":0000
         Top             =   810
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2865
         Picture         =   "M_LibMsr.frx":030A
         Top             =   420
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2865
         Picture         =   "M_LibMsr.frx":0614
         Top             =   75
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Casino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   39
         Top             =   225
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   38
         Top             =   570
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   37
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   36
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   4065
         TabIndex        =   35
         Top             =   1320
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3360
         TabIndex        =   34
         Top             =   180
         Width           =   4110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3360
         TabIndex        =   33
         Top             =   525
         Width           =   4110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3360
         TabIndex        =   32
         Top             =   870
         Width           =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   8
         Left            =   2880
         TabIndex        =   31
         Top             =   1320
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   9
         Left            =   6645
         TabIndex        =   30
         Top             =   1320
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
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
         Left            =   3405
         TabIndex        =   40
         Top             =   225
         Width           =   4110
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
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
         Left            =   3405
         TabIndex        =   41
         Top             =   570
         Width           =   4110
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
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
         Left            =   3405
         TabIndex        =   42
         Top             =   915
         Width           =   4110
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos Destino"
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
      Height          =   2205
      Index           =   1
      Left            =   30
      TabIndex        =   14
      Top             =   1725
      Width           =   7650
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1220
         Width           =   1800
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   2
         Left            =   1440
         TabIndex        =   7
         Top             =   520
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
         TabIndex        =   8
         Top             =   870
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
         Visible         =   0   'False
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
         Visible         =   0   'False
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
         TabIndex        =   6
         Top             =   180
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
         Left            =   2865
         Picture         =   "M_LibMsr.frx":091E
         Top             =   810
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   2865
         Picture         =   "M_LibMsr.frx":0C28
         Top             =   420
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2865
         Picture         =   "M_LibMsr.frx":0F32
         Top             =   75
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Subsegmento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   25
         Top             =   285
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   24
         Top             =   630
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   23
         Top             =   1005
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   22
         Top             =   1650
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   4065
         TabIndex        =   21
         Top             =   1650
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3360
         TabIndex        =   20
         Top             =   180
         Width           =   4110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   3360
         TabIndex        =   19
         Top             =   525
         Width           =   4110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   3360
         TabIndex        =   18
         Top             =   870
         Width           =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   10
         Left            =   2880
         TabIndex        =   17
         Top             =   1650
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   11
         Left            =   6645
         TabIndex        =   16
         Top             =   1650
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   15
         Top             =   1305
         Width           =   405
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
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
         Left            =   3405
         TabIndex        =   26
         Top             =   225
         Width           =   4110
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
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
         Left            =   3405
         TabIndex        =   27
         Top             =   570
         Width           =   4110
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
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
         Left            =   3405
         TabIndex        =   28
         Top             =   915
         Width           =   4110
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estructura Servicio Origen && Destino"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3585
      Left            =   30
      TabIndex        =   13
      Top             =   3930
      Width           =   7650
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2985
         Left            =   660
         TabIndex        =   12
         Top             =   360
         Width           =   6555
         _Version        =   393216
         _ExtentX        =   11562
         _ExtentY        =   5265
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         SpreadDesigner  =   "M_LibMsr.frx":123C
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7620
      Left            =   7770
      TabIndex        =   43
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   13441
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_LibMsr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private RS      As New ADODB.Recordset
Private RS1     As New ADODB.Recordset
Private indsel  As Long
Private BtnX    As Variant
Private codest  As Long
Private tipsus  As String
Private tipreg  As String
Private tipser  As String
Private ConSql  As Variant

Private Sub Form_Activate()
    Call fg_descarga
End Sub

Private Sub Form_Load()
Dim OpUsuario As String

    Me.HelpContextID = vg_OpcM
    fg_centra Me
    EspFecha fpDateTime1(0)
    EspFecha fpDateTime1(1)
    EspFecha fpDateTime1(2)
    EspFecha fpDateTime1(3)
    Msgtitulo = "Copiar Planificación Teórica"
    Toolbar1.ImageList = Partida.IL1
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
    fpDateTime1(0).text = Format(Date, "mm/yyyy")
'    fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
'    Label1(8).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(0).text, "yyyymmdd"), 1), 1, 4)
'    fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
'    Label1(9).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(1).text, "yyyymmdd"), 1), 1, 4)
'    fpDateTime1(2).text = Format(Date, "dd/mm/yyyy")
'    Label1(10).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(2).text, "yyyymmdd"), 1), 1, 4)
'    fpDateTime1(3).text = Format(Date, "dd/mm/yyyy")
'    Label1(11).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(3).text, "yyyymmdd"), 1), 1, 4)
    vaSpread1.MaxRows = 0
    vaSpread1.Row = -1: vaSpread1.Col = -1
    vaSpread1.BackColor = &HC0FFFF

    OpUsuario = 1
    If IsNull(OpUsuario) Or Trim(OpUsuario) = "" Then
        MsgBox "Contactese con el Administrador del Sistema...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Else
        Select Case OpUsuario
        Case "1"
    '        Combo2(1).Clear
    '        Combo2(1).AddItem "Real" & Space(150) & "(1)"
    '        Combo2(1).ListIndex = fg_buscacbo(Combo2, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
            Combo2(0).AddItem "Real" & Space(150) & "(1)"
'            Combo2(0).AddItem "Propuesta" & Space(150) & "(2)"
            Combo2(0).ListIndex = fg_buscacbo(Combo2, 0, 1, fg_pone_cero(Str(OpUsuario), 1))
    '    Case "2"
    '        Combo2(1).Clear
    '        Combo2(1).AddItem "Propuesta" & Space(150) & "(2)"
    '        Combo2(1).ListIndex = fg_buscacbo(Combo2, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
    '        Combo2(0).AddItem "Propuesta" & Space(150) & "(2)"
    '        Combo2(0).ListIndex = fg_buscacbo(Combo2, 0, 1, fg_pone_cero(Str(OpUsuario), 1))
        End Select
    End If
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
            RS.Open "SELECT * " & _
                       "FROM b_clientes With(NoLock) " & _
                       "WHERE cli_codigo = '" & Trim(LimpiaDato(FpText.text)) & "' " & _
                       "AND cli_tipo = 0", vg_db, adOpenStatic
               If RS.EOF Then
                   RS.Close
                   Set RS = Nothing
                   fpayuda(0).Caption = ""
                   fpLongInteger1(1).Value = ""
                   fpayuda(1).Caption = ""
                   fpLongInteger1(2).Value = ""
                   fpayuda(2).Caption = ""
                   Exit Sub
               End If
               fpayuda(0).Caption = Trim(RS!cli_nombre)
               RS.Close: Set RS = Nothing
               Call MoverVector
               fpLongInteger1(1).Value = ""
               fpayuda(1).Caption = ""
               
               fpLongInteger1(2).Value = ""
               fpayuda(2).Caption = ""
        
        Case 0
            Set RS = vg_db.Execute("sgpadm_s_casregimen 6, '" & Trim(LimpiaDato(FpText.text)) & "', " & Val(fpLongInteger1(0).Value) & ", ''")
            If Not RS.EOF Then
               fpayuda(1).Caption = RS!reg_nombre
            Else
               fpayuda(1).Caption = ""
            End If
            RS.Close: Set RS = Nothing
            Call MoverVector
        Case 1
            Set RS = vg_db.Execute("sgpadm_s_casservicio 6, '" & Trim(LimpiaDato(FpText.text)) & "', " & Val(fpLongInteger1(1).Value) & ", ''")
            If Not RS.EOF Then
               fpayuda(2).Caption = RS!ser_nombre
            Else
               fpayuda(2).Caption = ""
            End If
            RS.Close: Set RS = Nothing
            Call MoverVector
        Case 2
            If Val(fpLongInteger1(2).Value) < 1 Then fpayuda(4).Caption = "": Exit Sub
            RS.Open "SELECT * FROM a_regimen With(NoLock)  WHERE reg_codigo=" & Val(fpLongInteger1(2).Value) & " AND reg_indppr = 1", vg_db, adOpenStatic
            If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(4).Caption = "": Exit Sub
            fpayuda(4).Caption = Trim(RS!reg_nombre)
            RS.Close: Set RS = Nothing
        Case 3
            If Val(fpLongInteger1(3).Value) < 1 Then fpayuda(5).Caption = "": Exit Sub
            RS.Open "SELECT * FROM a_servicio With(NoLock) WHERE ser_codigo=" & Val(fpLongInteger1(3).Value) & " AND SER_indppr = 1", vg_db, adOpenStatic
            If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(5).Caption = "": Exit Sub
            fpayuda(5).Caption = Trim(RS!ser_nombre)
            RS.Close: Set RS = Nothing
            MoverVector
        Case 4
            If fpLongInteger1(4).Value = "" Then fpayuda(3).Caption = "": Exit Sub
            RS.Open "SELECT * FROM a_subsegmento With(NoLock)  WHERE sub_codigo=" & Val(fpLongInteger1(4).Value) & " AND sub_indppr = 1", vg_db, adOpenStatic
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

Private Sub fpText_Change()
RS.Open "SELECT * " & _
"FROM b_clientes With(NoLock) " & _
"WHERE cli_codigo = '" & Trim(LimpiaDato(FpText.text)) & "' " & _
"AND cli_tipo = 0", vg_db, adOpenStatic
If RS.EOF Then
   RS.Close
   Set RS = Nothing
   fpayuda(0).Caption = ""
   fpLongInteger1(1).Value = ""
   fpayuda(1).Caption = ""
   fpLongInteger1(2).Value = ""
   fpayuda(2).Caption = ""
   Exit Sub
End If
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
Call MoverVector
fpLongInteger1(1).Value = ""
fpayuda(1).Caption = ""
               
fpLongInteger1(2).Value = ""
fpayuda(2).Caption = ""
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Image1_Click(Index As Integer)
Dim Var_IndPpr As String
    Select Case Index
        Case 0
            vg_left = fpayuda(0).Left + 2300
            vg_nombre = "": vg_codigo = ""
            Call B_TabEst.LlenaDatos("b_clientes", "cli_", "Clientes", "Cliente_SitioRemoto")
            B_TabEst.Show 1
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            FpText.text = vg_codigo
            fpayuda(0).Caption = vg_nombre
            fpLongInteger1(0).Value = "": fpayuda(1).Caption = ""
            fpLongInteger1(1).Value = "": fpayuda(2).Caption = ""
            fpLongInteger1(0).SetFocus
        Case 1
            vg_left = fpayuda(1).Left + 2300
            vg_nombre = "": vg_codigo = ""
            Call B_TabEst.LlenaDatos("Cas_a_regimen", Trim(LimpiaDato(FpText.text)), "Regimen", "Regimen_SitioRemoto")
            B_TabEst.Show 1
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(0).Value = Val(vg_codigo)
            fpayuda(1).Caption = vg_nombre
            fpLongInteger1(1).SetFocus
        Case 2
            vg_left = fpayuda(2).Left + 2300
            vg_nombre = "": vg_codigo = ""
            Call B_TabEst.LlenaDatos("Cas_a_servicio", Trim(LimpiaDato(FpText.text)), "Servicio", "Servicio_SitioRemoto")
            B_TabEst.Show 1
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(1).Value = Val(vg_codigo)
            fpayuda(2).Caption = vg_nombre
            fpDateTime1(0).SetFocus
        Case 3
            Var_IndPpr = vg_Indppr
            vg_Indppr = 1
            vg_left = fpayuda(3).Left + 2300
            vg_nombre = "": vg_codigo = ""
            B_TabEst.LlenaDatos "a_subsegmento", "sub_", "Subsegmento", "Sub"
            B_TabEst.Show 1
            Me.Refresh
            vg_Indppr = Var_IndPpr
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(4).Value = Val(vg_codigo)
            fpayuda(3).Caption = vg_nombre
            fpLongInteger1(2).Value = "": fpayuda(4).Caption = ""
            fpLongInteger1(3).Value = "": fpayuda(5).Caption = ""
            fpLongInteger1(2).SetFocus
        Case 4
            Var_IndPpr = vg_Indppr
            vg_Indppr = 1
            vg_left = fpayuda(4).Left + 2300
            vg_nombre = "": vg_codigo = ""
            B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
            B_TabEst.Show 1
            Me.Refresh
            vg_Indppr = Var_IndPpr
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(2).Value = Val(vg_codigo)
            fpayuda(4).Caption = vg_nombre
            fpLongInteger1(3).SetFocus
        Case 5
            Var_IndPpr = vg_Indppr
            vg_Indppr = 1
            vg_left = fpayuda(5).Left + 2300
            vg_nombre = "": vg_codigo = ""
            B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Ser"
            B_TabEst.Show 1
            Me.Refresh
            vg_Indppr = Var_IndPpr
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(3).Value = Val(vg_codigo)
            fpayuda(5).Caption = vg_nombre
            'fpDateTime1(2).SetFocus
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i       As Long
Dim RS      As New ADODB.Recordset
Dim RS1     As New ADODB.Recordset
Dim fecori1 As Long
Dim fecori2 As Long
Dim fecdes1 As Long
Dim fecdes2 As Long
Dim vdia    As Long
Dim indice  As Long
Dim tiprec  As Long
Dim codest1 As Long
Dim codest2 As Long
Dim auxfeco As String
Dim auxfecd As String
Dim vaux1   As Long
Dim vaux2   As Long
Dim diatop  As Long
Dim Est     As Boolean
Dim cSpi    As Long


Dim Resp    As String
Dim estado  As Long


On Error GoTo Man_Error
    Select Case Button.Index
        Case 2
            If IsDate(fpDateTime1(0).text) = False Then
                Call MsgBox("Periodo esta en blanco", vbInformation, Me.Caption)
                Exit Sub
            End If
            
            If vaSpread1.MaxRows < 1 Then
                Call MsgBox("No existe concepto estructuras en datos origen, proceso cancelado", vbExclamation + vbOKOnly, Msgtitulo)
                Exit Sub
            End If
            
'            If Replace(Combo2(1).text, " ", "", , , vbTextCompare) = "Propuesta(2)" And Replace(Combo2(0).text, " ", "", , , vbTextCompare) = "Real(1)" Then
'                Call MsgBox("No es posible realizar copia de Propuesta a Real, proceso cancelado", vbExclamation + vbOKOnly, Msgtitulo)
'                Exit Sub
'            End If
            
            Est = False
            For i = 1 To vaSpread1.MaxRows
                vaSpread1.Row = i
                vaSpread1.Col = 4: estado = vaSpread1.TypeComboBoxCurSel
                vaSpread1.Col = 1
                If vaSpread1.text = "1" And estado = -1 Then
                    Call MsgBox("Falta seleccionar un concepto estructuras Destino...", vbExclamation + vbOKOnly, Msgtitulo)
                    Exit Sub 'Est = True
                End If
            Next i
            'If Est Then MsgBox "Falta seleccionar concepto estructuras...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub ' Comenté esta linea, para evitar que termine el proceso si no han seleccionado todos los item Samuel 02/09/09
            '-------> Validar datos origen
            
            Set RS = vg_db.Execute("sgpadm_s_LiveraMinutaSitRem 2, '" & Trim(LimpiaDato(FpText.text)) & "', 0, 0, 0, '', 0")
            If RS.EOF Then
                RS.Close
                Set RS = Nothing
                FpText.text = ""
                fpayuda(0).Caption = ""
                fpLongInteger1(0).Value = ""
                fpayuda(1).Caption = ""
                fpLongInteger1(1).Value = ""
                fpayuda(2).Caption = ""
                MsgBox "No existe casino", vbExclamation + vbOKOnly, Msgtitulo
                Exit Sub
            End If
            'TipSus = RS!sub_indppr
            RS.Close: Set RS = Nothing
            
            Set RS = vg_db.Execute("sgpadm_s_LiveraMinutaSitRem 3, '', " & fpLongInteger1(0).text & ", 0, 0, '', 0")
            If RS.EOF Then
                RS.Close: Set RS = Nothing
                fpLongInteger1(0).Value = ""
                fpayuda(1).Caption = ""
                Call MsgBox("No Existe Regimen", vbExclamation + vbOKOnly, Msgtitulo)
                Exit Sub
            End If
            'TipReg = RS!reg_indppr
            RS.Close: Set RS = Nothing
            
            Set RS = vg_db.Execute("sgpadm_s_LiveraMinutaSitRem 4, '', " & fpLongInteger1(1).text & ", 0, 0, '', 0")
            If RS.EOF Then
                RS.Close
                Set ConSql = Nothing: fpLongInteger1(1).Value = ""
                fpayuda(2).Caption = ""
                Call MsgBox("No Existe Servicio", vbExclamation + vbOKOnly, Msgtitulo)
                Exit Sub
            End If
            'TipSer = RS!ser_indppr
            RS.Close: Set RS = Nothing
            
            '-------> validar tipo planificación origen
            If ValidaOrigen = False Then Exit Sub
           
            '-------> Validar datos destino
            If ValidaDestino = False Then Exit Sub
            
            '-------> validar tipo planificación destino
'            If fg_codigocbo(Combo2, 0, 1, "") <> TipSus Or fg_codigocbo(Combo2, 0, 1, "") <> TipReg Or fg_codigocbo(Combo2, 0, 1, "") <> TipSer Then
'                Call MsgBox("Tipo planificación destino, no coincide con los códigos Sub-Segmento, Regimen o Servicio ...", vbExclamation + vbOKOnly, Msgtitulo)
'                Exit Sub
'            End If
            '-------> Validar datos destino bloqueado
            '"AND   a.min_cecori = '" & fpLongInteger1(5).text & "' " & _
            fecdes1 = Mid(fpDateTime1(2).text, 7, 4) & Mid(fpDateTime1(2).text, 4, 2)
            
'jpaz 20110310           Let fecdes1 = Mid(fpDateTime1(2).text, 7, 4) & Mid(fpDateTime1(2).text, 4, 2)
            Let fecdes1 = Mid(fpDateTime1(0).text, 4, 4) & Mid(fpDateTime1(0).text, 1, 2)
            
            Set RS = vg_db.Execute("sgpadm_s_LiveraMinutaSitRem 8, '', " & fpLongInteger1(2).text & ", " & fpLongInteger1(3).text & ", " & fecdes1 & ", '" & Val(fg_codigocbo(Combo2, 0, 1, "")) & "', 0")
            If Not RS.EOF And RS!nReg > 0 Then
                RS.Close
                Set RS = Nothing
                Call MsgBox("Minuta esta bloqueda, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo)
                Exit Sub
            End If
            RS.Close: Set RS = Nothing
            '-------> Validar si existe datos origen
            
'jpaz 20110310           fecori1 = Mid(fpDateTime1(0).text, 7, 4) & Mid(fpDateTime1(0).text, 4, 2)
            fecori1 = Mid(fpDateTime1(0).text, 4, 4) & Mid(fpDateTime1(0).text, 1, 2)
            
            Set RS = vg_db.Execute("sgpadm_s_LiveraMinutaSitRem 9, '" & Trim(LimpiaDato(FpText.text)) & "', " & fpLongInteger1(0).text & ", " & fpLongInteger1(1).text & ", " & fecori1 & ", '', 0")
            If RS.EOF Then
                RS.Close
                Set RS = Nothing
                Call MsgBox("No existe datos origen, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo)
                Exit Sub
            End If
            RS.Close: Set RS = Nothing
            '-------> Grabando plantilla casino origen hacia origen
            vdia = 999999: indice = 0: codest = 0
'jpaz 20110310            fecori1 = Format(fpDateTime1(0).text, "yyyymmdd")
'jpaz 20110310            fecori2 = Format(fpDateTime1(1).text, "yyyymmdd")
'jpaz 20110310            fecdes1 = Format(fpDateTime1(2).text, "yyyymmdd")
'jpaz 20110310            fecdes2 = Format(fpDateTime1(3).text, "yyyymmdd")
'jpaz 20110310            diatop = Format(fpDateTime1(3).text, "yyyymmdd")
            fecori1 = Format(dBoM("01/" & fpDateTime1(0).text), "yyyymmdd")
            fecori2 = Format(dEoM("01/" & fpDateTime1(0).text), "yyyymmdd")
            fecdes1 = Format(dBoM("01/" & fpDateTime1(0).text), "yyyymmdd")
            fecdes2 = Format(dEoM("01/" & fpDateTime1(0).text), "yyyymmdd")
            diatop = Format(dEoM("01/" & fpDateTime1(0).text), "yyyymmdd")
            '-------> validar si Existe Datos Destino
            Set RS = vg_db.Execute("sgpadm_s_LiveraMinutaSitRem 10, '" & fpLongInteger1(4).text & "', " & fpLongInteger1(2).text & ", " & fpLongInteger1(3).text & ", " & fecori1 & ", '" & Val(fg_codigocbo(Combo2, 0, 1, "")) & "', " & fecdes2)
            Let Resp = "S"
            If Not RS.EOF Then
                If MsgBox("Existe información casino destino. se borrara la información existente ..." & VgLinea & VgLinea & "              No = No copia Planificación ", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then
                   fg_descarga
                   RS.Close: Set RS = Nothing
                   Exit Sub
'                    Let Resp = "N"
                End If
            End If
            RS.Close: Set RS = Nothing
            '-------> Fin validar si existe datos destino
            Set RS = vg_db.Execute("sgpadm_s_LiveraMinutaSitRem 11, '" & Trim(LimpiaDato(FpText.text)) & "', " & fpLongInteger1(0).text & ", " & fpLongInteger1(1).text & ", " & fecori1 & ", '', " & fecori2)
            If RS.EOF Then
                RS.Close
                Set RS = Nothing
                Call MsgBox("No existe información", vbInformation + vbOKOnly, Msgtitulo)
                Exit Sub
            End If
            fg_carga ""
            Est = True
            If DatePart("w", fg_Ctod1(RS!min_fecmin), 2) <> DatePart("w", fg_Ctod1(fecdes1), 2) Then
                If MsgBox("No coincide día de la semana. ¿ Desea copiar ? ...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then
                    RS.Close
                    Set RS = Nothing
                    fg_descarga
                    Exit Sub
                Else
                    Est = False
                End If
            End If
            
            '-------> Borrar tabla de paso estructura servicio
            vg_db.Execute "DELETE paso_estservicio WHERE ess_spid=@@spid and ess_usr='" & vg_NUsr & "'"
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
                  vg_db.Execute ("INSERT INTO paso_estservicio (ess_spid, ess_usr, ess_codess1, ess_codess2, ess_desest2) VALUES (" & cSpi & ", '" & vg_NUsr & "', " & _
                                    codest1 & ", " & codest2 & ", '" & Trim(vaSpread1.text) & "')")
               End If
            Next i
            Toolbar1.Enabled = False
            Frame2.Enabled = False
            Frame1(0).Enabled = False
            Frame1(1).Enabled = False
            '",'" & Val(fg_codigocbo(Combo2, 1, 1, "")) &
            
'            vg_db.Execute "sgpadm_p_copiacreaplanifSitRem " & Val(fpLongInteger1(5).Value) & ", " & Val(fpLongInteger1(0).Value) & _
'                                                 ", " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(4).Value) & ", " & _
'                                                        Val(fpLongInteger1(2).Value) & ", " & Val(fpLongInteger1(3).Value) & ", " & _
'                                                    "" & fecori1 & " , " & fecori2 & " , " & _
'                                                    fecdes1 & ", " & diatop & ", " & _
'                                                    IIf(Est, 1, 0) & ", " & cSpi & _
'                                                    ", '" & vg_NUsr & "', 0, " & Val(Format(fpDateTime1(3).text, "yyyymm")) & _
'                                                    ",'" & Val(fg_codigocbo(Combo2, 0, 1, "")) & _
'                                                    "', '" & Resp & "'"
            
            vg_db.Execute "sgpadm_p_copiacreaplanifSitRem '" & Trim(LimpiaDato(FpText.text)) & "', " & Val(fpLongInteger1(0).Value) & _
                                                 ", " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(4).Value) & ", " & _
                                                        Val(fpLongInteger1(2).Value) & ", " & Val(fpLongInteger1(3).Value) & ", " & _
                                                    "" & fecori1 & " , " & fecori2 & " , " & _
                                                    fecdes1 & ", " & diatop & ", " & _
                                                    IIf(Est, 1, 0) & ", " & cSpi & _
                                                    ", '" & vg_NUsr & "', 0, " & Val(Format(fpDateTime1(0).text, "yyyymm")) & _
                                                    ",'" & Val(fg_codigocbo(Combo2, 0, 1, "")) & _
                                                    "', '" & Resp & "'"
            Toolbar1.Enabled = True
            Frame2.Enabled = True
            Frame1(0).Enabled = True
            Frame1(1).Enabled = True
            fg_descarga
            'Label1(5).Visible = False
            Call MsgBox("Copia Finalizada Sin Problema", vbInformation + vbOKOnly, Msgtitulo)
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
    If Err = -2147467259 Then
        Call MsgBox("El dato esta asociado a otra tabla...", vbCritical, "Error")
        Exit Sub
    End If
    If Err = 3034 Then: Exit Sub
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Function ValidaOrigen() As Boolean

    Let ValidaOrigen = True
'
'    If fg_codigocbo(Combo2, 1, 1, "") <> TipSus Or _
'        fg_codigocbo(Combo2, 1, 1, "") <> TipReg Or _
'        fg_codigocbo(Combo2, 1, 1, "") <> TipSer Then
'
'        Call MsgBox("Tipo planificación origen, no coincide con los códigos Sub-Segmento, Regimen o Servicio ...", vbExclamation + vbOKOnly, Msgtitulo)
'        Let ValidaOrigen = False
'        Exit Function
'    End If
    
    If Val(fpLongInteger1(5).Value) = Val(fpLongInteger1(4).Value) And Val(fpLongInteger1(0).Value) = Val(fpLongInteger1(2).Value) And _
        Val(fpLongInteger1(1).Value) = Val(fpLongInteger1(3).Value) And fpDateTime1(0).text = fpDateTime1(2).text And _
        fpDateTime1(1).text = fpDateTime1(3).text Then
        
        Call MsgBox("Datos origen, beben ser distinto datos destino", vbCritical + vbOKOnly, Msgtitulo)
        Let ValidaOrigen = False
        Exit Function
    End If
    
    If fpDateTime1(0).text = "" Or fpDateTime1(1).text = "" Or fpDateTime1(2).text = "" Or fpDateTime1(3).text = "" Then
        Call MsgBox("Fecha no definida", vbExclamation + vbOKOnly, Msgtitulo)
        Let ValidaOrigen = False
        Exit Function
    End If
            
'    If (Val(Mid(fpDateTime1(1).text, 1, 2)) - Val(Mid(fpDateTime1(0).text, 1, 2))) > (Val(Mid(fpDateTime1(3).text, 1, 2)) - Val(Mid(fpDateTime1(2).text, 1, 2))) Then
'        Call MsgBox("Fecha origen supera nº dìas", vbExclamation + vbOKOnly, Msgtitulo)
'        Let ValidaOrigen = False
'        Exit Function
'    End If
    
'    If Val(Format(fpDateTime1(0).text, "ddmmyyyy")) > Val(Format(fpDateTime1(1).text, "ddmmyyyy")) Or _
'        Val(Format(fpDateTime1(2).text, "ddmmyyyy")) > Val(Format(fpDateTime1(3).text, "ddmmyyyy")) Then
'
'        Call MsgBox("Fecha no coincide", vbExclamation + vbOKOnly, Msgtitulo)
'        Exit Function
'    End If
    
'    If Val(Format(fpDateTime1(0).text, "ddmmyyyy")) > Val(Format(fpDateTime1(1).text, "ddmmyyyy")) Or Val(Format(fpDateTime1(2).text, "ddmmyyyy")) > Val(Format(fpDateTime1(3).text, "ddmmyyyy")) Then
'        Call MsgBox("Fecha no coincide", vbExclamation + vbOKOnly, Msgtitulo)
'        Let ValidaOrigen = False
'        Exit Function
'    End If
    
'    If Val(Format(fpDateTime1(0).text, "mm")) <> Val(Format(fpDateTime1(1).text, "mm")) Or _
'        Val(Format(fpDateTime1(2).text, "mm")) <> Val(Format(fpDateTime1(3).text, "mm")) Then
'
'        Call MsgBox("Fecha no coincide", vbExclamation + vbOKOnly, Msgtitulo)
'        Let ValidaOrigen = False
'        Exit Function
'    End If
    
'    If (Val(Mid(fpDateTime1(1).text, 1, 2)) - Val(Mid(fpDateTime1(0).text, 1, 2))) > (Val(Mid(fpDateTime1(3).text, 1, 2)) - Val(Mid(fpDateTime1(2).text, 1, 2))) Then
'        Call MsgBox("Fecha origen supera nº dìas", vbExclamation + vbOKOnly, Msgtitulo)
'        Let ValidaOrigen = False
'        Exit Function
'    End If

End Function

Private Function ValidaDestino() As Boolean

    Let ValidaDestino = True
    Set RS = vg_db.Execute("sgpadm_s_LiveraMinutaSitRem 5, '', " & fpLongInteger1(4).text & ", 0, 0,'', 0")
        If RS.EOF Then
            RS.Close
            Set RS = Nothing
            fpLongInteger1(4).Value = ""
            fpayuda(3).Caption = ""
            fpLongInteger1(2).Value = ""
            fpayuda(4).Caption = ""
            fpLongInteger1(3).Value = ""
            fpayuda(5).Caption = ""
            Call MsgBox("No existe casino", vbExclamation + vbOKOnly, Msgtitulo)
            Let ValidaDestino = False
            Exit Function
        End If
        tipsus = RS!sub_indppr
        RS.Close: Set RS = Nothing
            
    Set RS = vg_db.Execute("sgpadm_s_LiveraMinutaSitRem 6, '', " & fpLongInteger1(2).text & ", 0, 0, '', 0")
        If RS.EOF Then
            RS.Close
            Set RS = Nothing
            fpLongInteger1(2).Value = ""
            fpayuda(4).Caption = ""
            Call MsgBox("No Existe Regimen", vbExclamation + vbOKOnly, Msgtitulo)
            Let ValidaDestino = False
            Exit Function
        End If
        'TipReg = RS!reg_indppr
        RS.Close: Set RS = Nothing
            
    Set RS = vg_db.Execute("sgpadm_s_LiveraMinutaSitRem 7, '', 0, " & fpLongInteger1(3).text & ", 0, '', 0")
        If RS.EOF Then
            RS.Close
            Set ConSql = Nothing
            fpLongInteger1(3).Value = ""
            fpayuda(5).Caption = ""
            Call MsgBox("No Existe Servicio", vbExclamation + vbOKOnly, Msgtitulo)
            Let ValidaDestino = False
            Exit Function
        End If
        tipser = RS!ser_indppr
        RS.Close: Set RS = Nothing
    
    Set RS = vg_db.Execute("sgpadm_s_LiveraMinutaSitRem 12, '" & Trim(LimpiaDato(FpText.text)) & "', " & fpLongInteger1(0).text & ", " & fpLongInteger1(1).text & ", " & Format(dBoM("01/" & fpDateTime1(0).text), "yyyymmdd") & ", '', " & Format(dEoM("01/" & fpDateTime1(0).text), "yyyymmdd") & "")
        If RS.EOF Then
            RS.Close
            Set ConSql = Nothing
            Call MsgBox("Para los datos origen, la minuta esta bloqueda", vbExclamation + vbOKOnly, Msgtitulo)
            Let ValidaDestino = False
            Exit Function
        End If
        RS.Close: Set RS = Nothing

End Function

Sub MoverVector()
Dim RS      As New ADODB.Recordset
Dim codest  As Long
Dim i       As Long
Dim codaux  As Long
Dim lisnom  As String
Dim liscod  As String
Dim z       As Long
Dim Anterior As String


    If Trim(FpText.text) = "" Or _
        Trim(fpLongInteger1(0).text) = "" Or _
        Trim(fpLongInteger1(1).text) = "" Or _
        Trim(fpDateTime1(0).text) = "" Or _
        Trim(fpLongInteger1(3).text) = "" Then Exit Sub

'Mover estructura servicio Origen
        Set RS = vg_db.Execute("sgpadm_s_CopiaMinutaSitRem '" & Trim(LimpiaDato(FpText.text)) & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Val(Format(fpDateTime1(0).text, "yyyymm")) & " ")
        vaSpread1.MaxRows = 0
        If Not RS.EOF Then
            Do While Not RS.EOF
                If Val(Anterior) <> RS!mid_estser Then
                    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
                    vaSpread1.Row = vaSpread1.MaxRows
                    vaSpread1.Col = 2: vaSpread1.text = RS!mid_estser
                    'vaSpread1.Col = 3: vaSpread1.text = IIf(Trim(RS!mid_desest) <> "", Trim(RS!mid_desest), Trim(RS!ess_nombre))
                    vaSpread1.Col = 3: vaSpread1.text = Trim(RS!ess_nombre)
                    Let Anterior = RS!mid_estser
                End If
                RS.MoveNext
            Loop
        End If
        RS.Close: Set RS = Nothing

        Set RS = vg_db.Execute("sgpadm_s_LiveraMinutaSitRem 1, '" & Trim(LimpiaDato(FpText.text)) & "', 0," & Val(fpLongInteger1(3).Value) & ", 0, '', 0")
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
                        If Val(vaSpread1.text) = codest Then codaux = z: Exit For
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
Dim indcol As Long
indcol = vaSpread1.ActiveCol
Select Case KeyCode
Case 46 And indcol = 5
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 4: vaSpread1.TypeComboBoxCurSel = -1
    vaSpread1.Col = 5: vaSpread1.TypeComboBoxCurSel = -1
End Select
End Sub
