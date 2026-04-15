VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_CopiaMinBloque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copia Minuta Bloque"
   ClientHeight    =   8670
   ClientLeft      =   4245
   ClientTop       =   2340
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Estructura Servicio Origen && Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4065
      Left            =   360
      TabIndex        =   41
      Top             =   4560
      Width           =   7305
      Begin VB.CommandButton btnProcesarEstructuras 
         Caption         =   "Procesar Estructuras"
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
         Left            =   5160
         TabIndex        =   43
         Top             =   480
         Width           =   1935
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2985
         Left            =   180
         TabIndex        =   42
         Top             =   960
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
         SpreadDesigner  =   "M_CopiaMinBloque.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos Destino"
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
      Height          =   2295
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   2160
      Width           =   7475
      Begin VB.OptionButton Option1 
         Caption         =   "Usar Rac. Origen"
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
         Left            =   165
         TabIndex        =   21
         Top             =   1920
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mantener Rac. en Destino"
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
         Left            =   4680
         TabIndex        =   20
         Top             =   1920
         Width           =   2655
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   2
         Left            =   1440
         TabIndex        =   22
         Top             =   750
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
         TabIndex        =   23
         Top             =   1090
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   1440
         TabIndex        =   24
         Top             =   400
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
         Index           =   2
         Left            =   1440
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
         Left            =   5430
         TabIndex        =   26
         Top             =   1440
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   2685
         Picture         =   "M_CopiaMinBloque.frx":04E3
         Top             =   1020
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   2685
         Picture         =   "M_CopiaMinBloque.frx":07ED
         Top             =   675
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2685
         Picture         =   "M_CopiaMinBloque.frx":0AF7
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   165
         TabIndex        =   36
         Top             =   465
         Width           =   735
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   165
         TabIndex        =   35
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   165
         TabIndex        =   34
         Top             =   1180
         Width           =   750
      End
      Begin VB.Label Label1 
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
         Height          =   315
         Index           =   4
         Left            =   165
         TabIndex        =   33
         Top             =   1530
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Height          =   315
         Index           =   7
         Left            =   4275
         TabIndex        =   32
         Top             =   1530
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   31
         Top             =   1530
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   6840
         TabIndex        =   30
         Top             =   1530
         Width           =   330
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3360
         TabIndex        =   29
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   3360
         TabIndex        =   28
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   3360
         TabIndex        =   27
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3405
         TabIndex        =   37
         Top             =   405
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   3405
         TabIndex        =   38
         Top             =   765
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   3405
         TabIndex        =   39
         Top             =   1125
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos Origen"
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
      Height          =   1815
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   750
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
         TabIndex        =   2
         Top             =   1090
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
         TabIndex        =   3
         Top             =   1440
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   4
         Top             =   400
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
         Index           =   1
         Left            =   5430
         TabIndex        =   5
         Top             =   1440
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2685
         Picture         =   "M_CopiaMinBloque.frx":0E01
         Top             =   1020
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2685
         Picture         =   "M_CopiaMinBloque.frx":110B
         Top             =   675
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2685
         Picture         =   "M_CopiaMinBloque.frx":1415
         Top             =   300
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
         Index           =   3
         Left            =   210
         TabIndex        =   15
         Top             =   465
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
         Index           =   0
         Left            =   210
         TabIndex        =   14
         Top             =   840
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
         Index           =   2
         Left            =   210
         TabIndex        =   13
         Top             =   1180
         Width           =   705
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   210
         TabIndex        =   12
         Top             =   1530
         Width           =   1110
      End
      Begin VB.Label Label1 
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
         Index           =   6
         Left            =   4275
         TabIndex        =   11
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   10
         Top             =   1530
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   6840
         TabIndex        =   9
         Top             =   1530
         Width           =   330
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3240
         TabIndex        =   8
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3240
         TabIndex        =   7
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3285
         TabIndex        =   16
         Top             =   405
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3285
         TabIndex        =   17
         Top             =   765
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3240
         TabIndex        =   6
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3285
         TabIndex        =   18
         Top             =   1125
         Width           =   3975
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8670
      Left            =   8040
      TabIndex        =   40
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   15293
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_CopiaMinBloque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset

Private Sub btnProcesarEstructuras_Click()
Dim RS As New ADODB.Recordset
Dim cadena As String
Dim cadenacodigoestruturadestino As String
Dim Sql As String

' validar si existe información
Sql = ""
Sql = Sql & "SELECT TOP 1 "
Sql = Sql & "        cbm.mid_codigo "
Sql = Sql & "FROM    dbo.b_minutadet AS cbm "
Sql = Sql & "WHERE   cbm.mid_codigo IN ( SELECT  min_codigo "
Sql = Sql & "                            FROM    dbo.b_minuta AS cbm2 "
Sql = Sql & "                            Where cbm2.min_cencos = '" & Trim(LimpiaDato(fpText(0).text)) & "' "
Sql = Sql & "                                    AND cbm2.min_codreg = " & Val(fpLongInteger1(0).Value) & " "
Sql = Sql & "                                    AND cbm2.min_codser = " & Val(fpLongInteger1(1).Value) & " "
Sql = Sql & "                                    AND cbm2.min_fecmin >= " & Format(fpDateTime1(0).text, "yyyymmdd") & " "
Sql = Sql & "                                    AND cbm2.min_fecmin <= " & Format(fpDateTime1(1).text, "yyyymmdd") & " )"
Sql = Sql & ""

Set RS = vg_db.Execute(Sql)
If RS.EOF Then
   RS.Close: Set RS = Nothing
   MsgBox "No existe minuta origen", vbExclamation + vbOKOnly, Me.Caption
   Exit Sub
End If
RS.Close: Set RS = Nothing
'VALIDACION GRAL DE CONTROLES
If ValidaControles = True Then
    MsgBox "Falta(n) ingresar\seleccionar valores en pantalla", vbExclamation, Me.Caption
    Exit Sub
End If

vaSpread1.MaxRows = 0
Dim i As Integer
i = 1

'********************************************************************************************************
    'genera el codigo de la col derecha
'********************************************************************************************************

'********************************************************************************************************
'BLOQUE DESTINO
Sql_MVI = " sgp_Sel_TraeEstructMinutaBloque "
Sql_MVI = Sql_MVI & " 'BloqueDes'"
Sql_MVI = Sql_MVI & " , '" & Trim(fpText(1)) & "'" 'cencoorigen
Sql_MVI = Sql_MVI & " ,'' " 'codigoSubsegmento
Sql_MVI = Sql_MVI & " , 0" 'codigoregimen
Sql_MVI = Sql_MVI & " , " & Trim(fpLongInteger1(3)) 'codigoservicio
Sql_MVI = Sql_MVI & " , 0" 'fechaorigenini
Sql_MVI = Sql_MVI & " , 0" 'fechaorigenfin
'********************************************************************************************************

    Set RS = vg_db.Execute(Sql_MVI)
    cadena = ""
    cadenacodigoestruturadestino = ""
    If RS.EOF = False Then
        While Not RS.EOF
           
            cadena = cadena & Chr(9) & Trim(RS!ess_nombre)
            cadenacodigoestruturadestino = cadenacodigoestruturadestino & Chr(9) & RS!ess_codigo
            
            RS.MoveNext
            i = i + 1
        
        Wend
    End If
    RS.Close: Set RS = Nothing
    If cadena = "" Then
        MsgBox "No hay estructura de servicios para esta selección", vbExclamation, Me.Caption
        Exit Sub
    End If

'********************************************************************************************************
    'genera el codigo de la col izq.
'********************************************************************************************************
    
    
'********************************************************************************************************
'BLOQUE ORIGEN
'********************************************************************************************************
    Sql_MVI = ""
    Sql_MVI = " sgp_Sel_TraeEstructMinutaBloque "
    Sql_MVI = Sql_MVI & " 'BloqueOri'"
    Sql_MVI = Sql_MVI & " , '" & Trim(fpText(0)) & "'" 'cencoorigen
    Sql_MVI = Sql_MVI & " ,'' " 'codigoSubsegmento
    Sql_MVI = Sql_MVI & " , " & Trim(fpLongInteger1(0)) 'codigoregimen
    Sql_MVI = Sql_MVI & " , " & Trim(fpLongInteger1(1)) 'codigoservicio
    Sql_MVI = Sql_MVI & " , '" & Format(fpDateTime1(0), "yyyymmdd") & "'" 'fechaorigenini
    Sql_MVI = Sql_MVI & " , '" & Format(fpDateTime1(1), "yyyymmdd") & "'" 'fechaorigenfin
'********************************************************************************************************
    
Set RS = vg_db.Execute(Sql_MVI)

If RS.EOF = False Then
    While Not RS.EOF
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 2: vaSpread1.text = CStr(RS!ess_codigo)
        vaSpread1.Col = 3: vaSpread1.text = CStr(RS!ess_nombre)
        
        vaSpread1.Col = 4
        
        vaSpread1.TypeComboBoxList = cadenacodigoestruturadestino
        
        vaSpread1.Col = 5
        
        vaSpread1.TypeComboBoxList = cadena
        vaSpread1.Col = 4
        
        'coloca el mismo codigo al lado derecho e izq., siempre y cuando sean el mismo servicio origen y destino
        If fpLongInteger1(1) = fpLongInteger1(3) Then
           vaSpread1.Col = 4
            For i = 0 To vaSpread1.TypeComboBoxCount
                 vaSpread1.TypeComboBoxCurSel = i
                 
                 If RS!ess_codigo = Val(vaSpread1.text) Then
                    vaSpread1.Col = 5
                    vaSpread1.TypeComboBoxCurSel = i
                    Exit For
                 End If
            Next i
        End If
        
        RS.MoveNext
        i = i + 1
    
    Wend
Else
   RS.Close: Set RS = Nothing
   MsgBox "No hay estructura de servicios origen", vbExclamation, Me.Caption
   Exit Sub
End If
RS.Close: Set RS = Nothing
End Sub

Private Function ValidaControles() As Boolean

ValidaControles = False

'ORIGEN
If fpText(0) = "" Then
    ValidaControles = True
    Exit Function
End If

If fpLongInteger1(1) = "" Then
    ValidaControles = True
    Exit Function
End If

If fpLongInteger1(2) = "" Then
    ValidaControles = True
    Exit Function
End If

If fpDateTime1(0) = "" Then
    ValidaControles = True
    Exit Function
End If

'DESTINO
If fpText(1) = "" Then
    ValidaControles = True
    Exit Function
End If

If fpLongInteger1(0) = "" Then
    ValidaControles = True
    Exit Function
End If

If fpLongInteger1(3) = "" Then
    ValidaControles = True
    Exit Function
End If

End Function

Private Function ValidaGrilla(ByVal Spread As vaSpread, ByVal opcion As Integer) As Boolean

ValidaGrilla = False
Dim estado As String
Dim cont As Integer

cont = 0

If opcion = 1 Then

    For i = 1 To Spread.MaxRows
    
        Spread.Row = i
        Spread.Col = 1
        
        Spread.Col = 5: estado = Spread.text
        Spread.Col = 1
        If Spread.text = "1" And estado = "" Then
            
            ValidaGrilla = True
            Exit Function
            
        End If
    
    Next

ElseIf opcion = 2 Then

    For i = 1 To Spread.MaxRows
    
        Spread.Row = i
        Spread.Col = 1
        
        If Spread.text <> "1" Then
            
            cont = cont + 1
            
        End If
    
    Next

    If cont = Spread.MaxRows Then ValidaGrilla = True

End If

End Function

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
fg_centra Me
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
fpText(0).Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText(0).text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
fpText(1).Enabled = ModCasino
Image1(3).Enabled = ModCasino
fpText(1).text = MuestraCasino(1)
fpayuda(3).Caption = MuestraCasino(2)
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
Label1(8).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(0).text, "yyyymmdd"), 1), 1, 4)
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
Label1(9).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(1).text, "yyyymmdd"), 1), 1, 4)
fpDateTime1(2).text = Format(Date, "dd/mm/yyyy")
Label1(10).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(2).text, "yyyymmdd"), 1), 1, 4)
fpDateTime1(3).text = Format(Date, "dd/mm/yyyy")
Label1(11).Caption = Mid(fg_Fecha_Dia(Format(fpDateTime1(3).text, "yyyymmdd"), 1), 1, 4)
vaSpread1.MaxRows = 0
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
Select Case Index
Case 0
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
Select Case Index
Case 0
'    Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(0).Value) & " and reg_codigo > 9999")
    Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(0).Value) & "")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
'    If ((Val(fpLongInteger1(0).Value) < 10000 And Val(fpLongInteger1(1).Value) < 10000) And (Val(fpLongInteger1(2).Value) > 9999 Or Val(fpLongInteger1(3).Value) > 9999)) Then Toolbar1.Buttons(2).Enabled = falso Else Toolbar1.Buttons(2).Enabled = True
Case 1
'    Set RS = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(1).Value) & " and ser_codigo > 9999")
    Set RS = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(1).Value) & "")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "": Exit Sub
    fpayuda(2).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
'    If ((Val(fpLongInteger1(0).Value) < 10000 And Val(fpLongInteger1(1).Value) < 10000) And (Val(fpLongInteger1(2).Value) > 9999 Or Val(fpLongInteger1(3).Value) > 9999)) Then Toolbar1.Buttons(2).Enabled = falso Else Toolbar1.Buttons(2).Enabled = True
Case 2
    Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(2).Value) & " and reg_codigo > 9999")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(4).Caption = "": Exit Sub
    fpayuda(4).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
'    If ((Val(fpLongInteger1(0).Value) < 10000 And Val(fpLongInteger1(1).Value) < 10000) And (Val(fpLongInteger1(2).Value) > 9999 Or Val(fpLongInteger1(3).Value) > 9999)) Then Toolbar1.Buttons(2).Enabled = falso Else Toolbar1.Buttons(2).Enabled = True
Case 3
    Set RS = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(3).Value) & " and ser_codigo > 9999")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(5).Caption = "": Exit Sub
    fpayuda(5).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
'    If ((Val(fpLongInteger1(0).Value) < 10000 And Val(fpLongInteger1(1).Value) < 10000) And (Val(fpLongInteger1(2).Value) > 9999 Or Val(fpLongInteger1(3).Value) > 9999)) Then Toolbar1.Buttons(2).Enabled = falso Else Toolbar1.Buttons(2).Enabled = True
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    If Index = 1 Then Image1_Click 1
    If Index = 2 Then Image1_Click 2
    If Index = 4 Then Image1_Click 4
    If Index = 5 Then Image1_Click 5
End Select
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    If Index = 0 Then Image1_Click 0
    If Index = 3 Then Image1_Click 3
End Select
End Sub

Private Sub fpText_LostFocus(Index As Integer)
Select Case Index
Case 0
    If fpText(0).text = "" Then fpayuda(0).Caption = "": Exit Sub
    Set RS = vg_db.Execute("SELECT * FROM b_clientes WHERE cli_codigo = '" & fpText(0).text & "' AND cli_tipo = 0 and cli_tipominuta = 1")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": fpLongInteger1(0).Value = "": fpayuda(1).Caption = "": fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!cli_nombre)
    RS.Close: Set RS = Nothing
    fpLongInteger1(0).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(1).Value = "": fpayuda(2).Caption = ""
Case 1
    If fpText(1).text = "" Then fpayuda(3).Caption = "": Exit Sub
    Set RS = vg_db.Execute("SELECT * FROM b_clientes WHERE cli_codigo = '" & fpText(1).text & "' AND cli_tipo = 0 and cli_tipominuta = 1")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(3).Caption = "": fpLongInteger1(2).Value = "": fpayuda(4).Caption = "": fpLongInteger1(3).Value = "": fpayuda(5).Caption = "": Exit Sub
    fpayuda(3).Caption = Trim(RS!cli_nombre)
    RS.Close: Set RS = Nothing
    fpLongInteger1(2).Value = "": fpayuda(4).Caption = ""
    fpLongInteger1(3).Value = "": fpayuda(5).Caption = ""
End Select
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
    fpText(0).text = vg_codigo
    fpayuda(0).Caption = vg_nombre
    fpLongInteger1(0).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(1).Value = "": fpayuda(2).Caption = ""
    fpLongInteger1(0).SetFocus
Case 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
'    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "RegBlo"
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
'    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "SerBlo"
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre
    fpDateTime1(0).SetFocus
Case 3
    vg_left = fpayuda(3).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(1).text = vg_codigo
    fpayuda(3).Caption = vg_nombre
    fpLongInteger1(2).Value = "": fpayuda(4).Caption = ""
    fpLongInteger1(3).Value = "": fpayuda(5).Caption = ""
    fpLongInteger1(2).SetFocus
Case 4
    vg_left = fpayuda(4).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "RegBlo"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(4).Caption = vg_nombre
    fpLongInteger1(3).SetFocus
Case 5
    vg_left = fpayuda(5).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "SerBlo"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(3).Value = Val(vg_codigo)
    fpayuda(5).Caption = vg_nombre
    fpDateTime1(2).SetFocus
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim fecori1 As Long, fecori2 As Long, fecdes1 As Long, fecdes2 As Long, vdia As Long, indice As Long, tiprec As Long
Dim auxfeco As String, auxfecd As String, vaux1 As Long, vaux2 As Long, diatop As Long, est As Boolean, enumrac As Long, sql1 As String, sql2 As String
Dim i         As Long
Dim MyBuffer  As String
Dim codest1   As Long
Dim codest2   As Long
Dim descr     As String

On Error GoTo Man_Error
Select Case Button.Index
Case 2

    If vaSpread1.MaxRows = 0 Then Exit Sub

    'debe tener algun elemento de su izq. seleccionado
    If ValidaGrilla(vaSpread1, 2) = True Then

       MsgBox "Debe tickear al menos un elemento a la izquierda", vbExclamation, Me.Caption
         Exit Sub
    
    End If

    'debe tener algo seleccionado a la derecha si esta tickeado a la izq.
    If ValidaGrilla(vaSpread1, 1) = True Then

        MsgBox "Debe seleccionar un elemento a la derecha", vbExclamation, Me.Caption
        Exit Sub
    
    End If

    'luego aca va el proced. almac que copia los valores a las tablas destino
    '-------> Grabar tabla de paso estructura servicio
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<GrabaEstservicio>"

    For i = 1 To vaSpread1.MaxRows
       vaSpread1.Row = i
       vaSpread1.Col = 1
       vaSpread1.Col = 2: codest1 = Val(vaSpread1.text)
       vaSpread1.Col = 4: codest2 = Val(vaSpread1.text)
       vaSpread1.Col = 1
       If vaSpread1.text = "1" And vaSpread1.Row > 0 And codest1 > 0 And codest2 > 0 Then
          vaSpread1.Col = 2: codest1 = Val(vaSpread1.text)
          vaSpread1.Col = 4: codest2 = Val(vaSpread1.text)
          vaSpread1.Col = 5: descr = vaSpread1.text

          MyBuffer = MyBuffer & " <Estservicio"
          descr = Replace(Trim(descr), Chr(34), "&quot;")
          descr = Replace(Trim(descr), Chr(38), "&amp;")
          descr = Replace(Trim(descr), Chr(39), "&apos;")
          descr = Replace(Trim(descr), Chr(60), "&lt;")
          descr = Replace(Trim(descr), Chr(62), "&gt;")
                                
          MyBuffer = MyBuffer & " codest1 = " & Chr(34) & codest1 & Chr(34)
          MyBuffer = MyBuffer & " codest2 = " & Chr(34) & codest2 & Chr(34)
          MyBuffer = MyBuffer & " descr = " & Chr(34) & descr & Chr(34)
          MyBuffer = MyBuffer & "/>"
 
       End If
    Next i
    MyBuffer = MyBuffer & "</GrabaEstservicio>"

    '------- Validar datos origen
    Set RS = vg_db.Execute("SELECT * FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND cli_tipo = 0")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText(0).text = "": fpayuda(0).Caption = "": fpLongInteger1(0).Value = "": fpayuda(1).Caption = "": fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": fg_descarga: MsgBox "No existe contrato", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(0).Value) & "")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(0).Value = "": fpayuda(1).Caption = "": fg_descarga: MsgBox "No Existe Regimen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    Set RS = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(1).Value) & "")
    If RS.EOF Then RS.Close: Set ConSql = Nothing: fpLongInteger1(1).Value = "": fpayuda(2).Caption = "": fg_descarga: MsgBox "No Existe Servicio", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    If fpText(0).text = fpText(1).text And Val(fpLongInteger1(0).Value) = Val(fpLongInteger1(2).Value) And Val(fpLongInteger1(1).Value) = Val(fpLongInteger1(3).Value) And fpDateTime1(0).text = fpDateTime1(2).text And fpDateTime1(1).text = fpDateTime1(3).text Then fg_descarga: MsgBox "Datos origen, beben ser distinto datos destino", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    If fpDateTime1(0).text = "" Or fpDateTime1(1).text = "" Or fpDateTime1(2).text = "" Or fpDateTime1(3).text = "" Then fg_descarga: MsgBox "Fecha no definida", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If (Val(Mid(fpDateTime1(1).text, 1, 2)) - Val(Mid(fpDateTime1(0).text, 1, 2))) > (Val(Mid(fpDateTime1(3).text, 1, 2)) - Val(Mid(fpDateTime1(2).text, 1, 2))) Then fg_descarga: MsgBox "Fecha origen supera nº dìas", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Val(Format(fpDateTime1(0).text, "ddmmyyyy")) > Val(Format(fpDateTime1(1).text, "ddmmyyyy")) Or Val(Format(fpDateTime1(2).text, "ddmmyyyy")) > Val(Format(fpDateTime1(3).text, "ddmmyyyy")) Then fg_descarga: MsgBox "Fecha no coincide", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Val(Format(fpDateTime1(0).text, "ddmmyyyy")) > Val(Format(fpDateTime1(1).text, "ddmmyyyy")) Or Val(Format(fpDateTime1(2).text, "ddmmyyyy")) > Val(Format(fpDateTime1(3).text, "ddmmyyyy")) Then fg_descarga: MsgBox "Fecha no coincide", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Val(Format(fpDateTime1(0).text, "mm")) <> Val(Format(fpDateTime1(1).text, "mm")) Or Val(Format(fpDateTime1(2).text, "mm")) <> Val(Format(fpDateTime1(3).text, "mm")) Then fg_descarga: MsgBox "Fecha no coincide", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If (Val(Mid(fpDateTime1(1).text, 1, 2)) - Val(Mid(fpDateTime1(0).text, 1, 2))) > (Val(Mid(fpDateTime1(3).text, 1, 2)) - Val(Mid(fpDateTime1(2).text, 1, 2))) Then fg_descarga: MsgBox "Fecha origen supera nº dìas", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Val(Format(fpDateTime1(2).text, "yyyymmdd")) > Val(Format(fpDateTime1(3).text, "yyyymmdd")) Then fg_descarga: MsgBox "Fecha destino inicial mayor final", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    '------- Validar datos destino
    Set RS = vg_db.Execute("SELECT * FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText(1).text)) & "' AND cli_tipo = 0")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText(1).text = "": fpayuda(3).Caption = "": fpLongInteger1(2).Value = "": fpayuda(4).Caption = "": fpLongInteger1(3).Value = "": fpayuda(5).Caption = "": fg_descarga: MsgBox "No existe contrato", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(2).Value) & "")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(2).Value = "": fpayuda(4).Caption = "": fg_descarga: MsgBox "No Existe Regimen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    Set RS = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(3).Value) & "")
    If RS.EOF Then RS.Close: Set ConSql = Nothing: fpLongInteger1(3).Value = "": fpayuda(5).Caption = "": fg_descarga: MsgBox "No Existe Servicio", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '------- Validar datos destino bloqueado
    fecdes1 = Format(fpDateTime1(2).text, "yyyymm")
    sql1 = " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) "
    If vg_tipmin Then
       '-------> Validar si existe una planificación teorica el mes a copiar
        Set RS = vg_db.Execute("SELECT COUNT(a.min_codigo) AS nreg " & _
                "FROM  b_minuta a, b_minutadet b " & _
                "WHERE a.min_codigo = b.mid_codigo " & _
                "AND   a.min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
                "AND   a.min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
                "AND   a.min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
                "AND   " & sql1 & " = " & fecdes1 & " " & _
                "AND   a.min_indblo IN (0) " & _
                "AND   b.mid_tipmin = '1'")
        If Not RS.EOF And RS!nreg > 0 Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No es posible copiar, ya que la minuta corresponde minuta normal, proceso cancelado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
        RS.Close: Set RS = Nothing
    End If
    sql2 = IIf(vg_tipmin, 2, 1)
    If vg_tipmin Then
       Set RS = vg_db.Execute("SELECT COUNT(a.min_codigo) AS nreg " & _
               "FROM  b_minuta a, b_minutadet b " & _
               "WHERE a.min_codigo = b.mid_codigo " & _
               "AND   a.min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
               "AND   a.min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
               "AND   a.min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
               "AND   " & sql1 & " = " & fecdes1 & " " & _
               "AND   a.min_indblo In (2,1,99) " & _
               "AND   b.mid_tipmin = '1'")
    Else
       Set RS = vg_db.Execute("SELECT COUNT(a.min_codigo) AS nreg " & _
               "FROM  b_minuta a, b_minutadet b " & _
               "WHERE a.min_codigo = b.mid_codigo " & _
               "AND   a.min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
               "AND   a.min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
               "AND   a.min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
               "AND   " & sql1 & " = " & fecdes1 & " " & _
               "AND   a.min_indblo In (2,1,99) " & _
               "AND   b.mid_tipmin = '1'")
    End If
    If Not RS.EOF And RS!nreg > 0 Then
       RS.Close: Set RS = Nothing
       fg_descarga
       MsgBox "Minuta esta bloqueda, proceso cancelado", vbCritical + vbOKOnly, MsgTitulo
       Exit Sub
    End If
    RS.Close: Set RS = Nothing
    sql2 = IIf(vg_tipmin, 2, 1)
    Set RS = vg_db.Execute("SELECT COUNT(a.min_codigo) AS nreg " & _
            "FROM  b_minuta a, b_minutadet b " & _
            "WHERE a.min_codigo = b.mid_codigo " & _
            "AND   a.min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
            "AND   a.min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
            "AND   a.min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
            "AND   " & sql1 & " = " & fecdes1 & " " & _
            "AND   a.min_indblo IN ( 2,1,99) " & _
            "AND   b.mid_tipmin = '1'")
    If Not RS.EOF And RS!nreg > 0 Then
       RS.Close: Set RS = Nothing
       fg_descarga
       MsgBox "Minuta esta bloqueda, proceso cancelado", vbCritical + vbOKOnly, MsgTitulo
       Exit Sub
    End If
    RS.Close: Set RS = Nothing
    '------- Validar si existe datos origen
    fecori1 = Mid(fpDateTime1(0).text, 7, 4) & Mid(fpDateTime1(0).text, 4, 2)
    sql1 = " convert(int,substring(convert(varchar(8),min_fecmin),1,6)) "
    Set RS = vg_db.Execute("SELECT DISTINCT min_cencos, min_codreg, min_codser " & _
            "FROM  b_minuta " & _
            "WHERE min_cencos = '" & LimpiaDato(Trim(fpText(0).text)) & "' " & _
            "AND   min_codreg = " & Val(fpLongInteger1(0).Value) & " " & _
            "AND   min_codser = " & Val(fpLongInteger1(1).Value) & " " & _
            "AND   " & sql1 & " = " & fecori1 & "")
    If RS.EOF Then
       RS.Close: Set RS = Nothing
       fg_descarga
       MsgBox "No existe datos origen, proceso cancelado", vbCritical + vbOKOnly, MsgTitulo
       Exit Sub
    End If
    RS.Close: Set RS = Nothing
    '------- Grabando plantilla contrato origen hacia origen
    vdia = 999999: indice = 0
    fecori1 = Format(fpDateTime1(0).text, "yyyymmdd")
    fecori2 = Format(fpDateTime1(1).text, "yyyymmdd")
    fecdes1 = Format(fpDateTime1(2).text, "yyyymmdd")
    fecdes2 = Format(fpDateTime1(3).text, "yyyymmdd")
    diatop = Format(fpDateTime1(3).text, "yyyymmdd")
    '------- validar si Existe Datos Destino
    Set RS = vg_db.Execute("SELECT DISTINCT min_cencos, min_codreg, min_codser " & _
             "FROM  b_minuta " & _
             "WHERE min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
             "AND   min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
             "AND   min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
             "AND   min_fecmin >= " & fecdes1 & " and min_fecmin <= " & fecdes2 & "")
    If Not RS.EOF Then
       If MsgBox("Existe información contrato destino. se borrara la información existente ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
          RS.Close: Set RS = Nothing
          fg_descarga
          Exit Sub
       End If
    End If
    RS.Close: Set RS = Nothing
    '------- Fin validar si existe datos destino
    Set RS = vg_db.Execute("SELECT a.*, b.* " & _
            "FROM  b_minuta a, b_minutadet b, b_receta c " & _
            "WHERE a.min_codigo = b.mid_codigo " & _
            "AND   b.mid_codrec = c.rec_codigo " & _
            "AND  (c.rec_fecvig > " & Format(Date, "yyyymmdd") & " OR c.rec_fecvig <= 0 OR (c.rec_fecvig) IS NULL) " & _
            "AND   a.min_cencos = '" & LimpiaDato(Trim(fpText(0).text)) & "' " & _
            "AND   a.min_codreg = " & Val(fpLongInteger1(0).Value) & " " & _
            "AND   a.min_codser = " & Val(fpLongInteger1(1).Value) & " " & _
            "AND   a.min_fecmin >= " & fecori1 & " " & _
            "AND   a.min_fecmin <= " & fecori2 & " " & _
            "AND   b.mid_tipmin = '1' ORDER BY a.min_fecmin")
    If RS.EOF Then
       RS.Close: Set RS = Nothing
       fg_descarga
       MsgBox "No existe información", vbInformation + vbOKOnly, MsgTitulo
       Exit Sub
    End If
    fg_carga ""
    est = True
    If DatePart("w", fg_Ctod1(RS!min_fecmin), 2) <> DatePart("w", fg_Ctod1(fecdes1), 2) Then
       If MsgBox("No coincide día de la semana. ¿ Desea copiar ? ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
          RS.Close: Set RS = Nothing
          fg_descarga
          Exit Sub
       Else
       est = False
       End If
    Else
       RS.Close: Set RS = Nothing
    End If
    Set RS = vg_db.Execute("sgp_Ins_XmlCopiaMinutaBloque '" & MyBuffer & "', '" & LimpiaDato(Trim(fpText(0).text)) & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", '" & LimpiaDato(Trim(fpText(1).text)) & "', " & Val(fpLongInteger1(2).Value) & ", " & Val(fpLongInteger1(3).Value) & ", " & fecori1 & ", " & fecori2 & ", " & fecdes1 & ", " & fecdes2 & ", '" & fecdes1 & "', '" & IIf(Option1(0).Value = True, "1", "2") & "'")
    If Not RS.EOF Then
       If RS(0) > 0 Then
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
       End If
    End If
    RS.Close: Set RS = Nothing
    
'
'    vg_db.BeginTrans
'    Do While Not RS.EOF
'       If RS!min_fecmin <> vdia Then
'          If Not est And fecdes1 > diatop Then GoTo paso
'          If est Then
'             auxfeco = Mid(RS!min_fecmin, 7, 2) & "/" & Mid(RS!min_fecmin, 5, 2) & "/" & Mid(RS!min_fecmin, 1, 4)
'             auxfecd = Mid(fecdes1, 7, 2) & "/" & Mid(fecdes1, 5, 2) & "/" & Mid(fecdes1, 1, 4)
'
'             vaux1 = DatePart("w", auxfeco, 2)
'             vaux2 = DatePart("w", auxfecd, 2)
'
'             Do While (vaux1 <> vaux2)
'                fecdes1 = (fecdes1 + 1)
'                If fecdes1 > diatop Then GoTo paso
'                auxfecd = Mid(fecdes1, 7, 2) & "/" & Mid(fecdes1, 5, 2) & "/" & Mid(fecdes1, 1, 4)
'                vaux2 = DatePart("w", auxfecd, 2)
'             Loop
'          End If
'          indice = 0
'          '------- actualizar nro. raciones totales
'          enumrac = 0
'          If Option1(1).Value = True Then
'             RS1.Open "SELECT sra_serdia, SUM(sra_raciones) AS raciones FROM a_serviciorac WHERE sra_cencos = '" & MuestraCasino(1) & "' AND sra_codser = " & Val(fpLongInteger1(3).Value) & " AND sra_serdia = " & IIf(DatePart("w", fg_Ctod1(fecdes1), 2) = 1, 7, DatePart("w", fg_Ctod1(fecdes1), 2) - 1) & "  GROUP BY sra_serdia", vg_db, adOpenStatic
'             If Not RS1.EOF Then enumrac = RS1!raciones
'             RS1.Close: Set RS1 = Nothing
'          End If
'          sql2 = IIf(vg_tipmin, 11, 0)
'          RS1.Open "SELECT min_codigo FROM b_minuta " & _
'                   "WHERE min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
'                   "AND   min_codreg = " & Val(fpLongInteger1(2).Value) & " " & _
'                   "AND   min_codser = " & Val(fpLongInteger1(3).Value) & " " & _
'                   "AND   min_fecmin = " & fecdes1 & " " & _
'                   "AND   min_indblo = " & sql2 & "", vg_db, adOpenStatic
'          If Not RS1.EOF Then
'             indice = RS1!min_codigo
'             RS1.Close: Set RS1 = Nothing
''             vg_db.Execute "DELETE b_minuta FROM b_minuta " & _
''                           "WHERE mid_cencos='" & vg_codcasino & "' " & _
''                           "AND   mid_codreg=" & vg_codregimen & " " & _
''                           "AND   mid_codser=" & vg_codservicio & " " & _
''                           "AND   min_codigo=" & indiceminutas & " " & _
''                           "AND   min_fecmin=" & Val(wsfecha) & ""
'             vg_db.Execute "DELETE b_minutadet FROM b_minutadet WHERE mid_codigo = " & indice & " AND mid_tipmin = '1'"
'             vg_db.Execute "UPDATE b_minuta SET min_racteo = " & IIf(Option1(0).Value = True, IIf(IsNull(RS!min_racteo), 0, RS!min_racteo), enumrac) & " WHERE min_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' AND min_codreg = " & Val(fpLongInteger1(2).Value) & " AND min_codser = " & Val(fpLongInteger1(3).Value) & " AND min_fecmin = " & fecdes1 & " AND min_codigo = " & indice & ""
'          Else
'             RS1.Close: Set RS1 = Nothing
'             RS1.Open "SELECT min_codigo FROM b_minuta ORDER BY min_codigo DESC", vg_db, adOpenStatic
'             If Not RS1.EOF Then RS1.MoveFirst: indice = RS1!min_codigo + 1 Else indice = 1
'             RS1.Close: Set RS1 = Nothing
'             sql2 = IIf(vg_tipmin, 11, 0)
'             vg_db.Execute "INSERT INTO b_minuta (min_codigo, min_cencos, min_codreg, min_codser, min_fecmin, min_indblo, min_racteo, min_racrea) " & _
'                           "VALUES (" & indice & ", '" & LimpiaDato(Trim(fpText(1).text)) & "', " & Val(fpLongInteger1(2).Value) & ", " & _
'                           "" & Val(fpLongInteger1(3).Value) & ", " & fecdes1 & ", " & sql2 & ", " & IIf(Option1(0).Value = True, IIf(IsNull(RS!min_racteo), 0, RS!min_racteo), enumrac) & ", 0)"
'          End If
'          vdia = RS!min_fecmin
'          If Not est Then
'             fecdes1 = fecdes1 + 1
'          End If
'       End If
'       '------- Traer tipo receta
'       tiprec = 0
'       RS1.Open "SELECT DISTINCT red_tiprec FROM b_recetadet WHERE red_codigo = " & RS!mid_codrec & " AND ((red_tiprec <> 0 AND red_cencos = '" & MuestraCasino(1) & "') OR (red_tiprec = 0 AND red_cencos = '0')) ORDER BY red_tiprec", vg_db, adOpenStatic
'       If Not RS1.EOF Then
'          Do While Not RS1.EOF
'             If RS1!red_tiprec = -1 Then
'                tiprec = IIf((fpLongInteger1(2).Value) < 10000, RS1!red_tiprec, 0)
'             ElseIf RS1!red_tiprec = Val(fpLongInteger1(2).Value) And RS1!red_tiprec = RS!mid_tiprec Then
'                tiprec = RS1!red_tiprec
'                Exit Do
'             ElseIf RS1!red_tiprec <> Val(fpLongInteger1(2).Value) And RS1!red_tiprec = RS!mid_tiprec Then
'                tiprec = RS!mid_tiprec
'                Exit Do
'             End If
'             RS1.MoveNext
'          Loop
'       End If
'       RS1.Close: Set RS1 = Nothing
'       RS1.Open "SELECT * FROM b_minutadet WHERE mid_codigo = " & indice & " AND mid_tipmin = '1' AND mid_numlin = " & RS!mid_numlin & "", vg_db, adOpenStatic
'       If RS1.EOF Then
'          vg_db.Execute "INSERT INTO b_minutadet (mid_codigo, mid_tipmin, mid_numlin, mid_estser, mid_codrec, mid_numrac, mid_descri, mid_cosrec, mid_tiprec, mid_nummer, mid_rec5eta, mid_cosdes, mid_modmina, mid_modminb) " & _
'                        "VALUES (" & indice & ", '1', " & RS!mid_numlin & ", " & RS!mid_estser & ", " & RS!mid_codrec & ", " & IIf(Option1(0).Value = True, IIf(IsNull(RS!mid_numrac), 0, RS!mid_numrac), 0) & ", '" & RS!mid_descri & "', " & fg_CalCtoRecInv(RS!mid_codrec, tiprec, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','"))) & ", " & tiprec & ", 0, '" & IIf((fpLongInteger1(2).Value) < 10000, 0, RS!mid_rec5eta) & "', " & fg_CalCtoRecInv(RS!mid_codrec, tiprec, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))) & ", '0', '0')"
'       Else
'          vg_db.Execute "UPDATE b_minutadet SET mid_modmina= '0', mid_modminb= '0', mid_estser = " & RS!mid_estser & ", mid_codrec = " & RS!mid_codrec & ", mid_numrac = " & RS!mid_numrac & ", mid_descri = '" & RS!mid_descri & "', mid_cosrec = " & fg_CalCtoRecInv(RS!mid_codrec, tiprec, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','"))) & ", mid_tiprec=" & tiprec & ", mid_rec5eta='" & IIf((fpLongInteger1(2).Value) < 10000, 0, 1) & "', mid_cosdes=" & fg_CalCtoRecInv(RS!mid_codrec, tiprec, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))) & " WHERE mid_codigo=" & indice & " AND mid_tipmin='1' AND mid_numlin=" & RS!mid_numlin & ""
'       End If
'       RS1.Close: Set RS1 = Nothing
'       RS.MoveNext
'    Loop
'paso:
'    RS.Close: Set RS = Nothing
'    vg_db.CommitTrans
'    fg_descarga
'    Picture1.Visible = False: Label1(5).Visible = False: gauge.Visible = False
    fg_descarga
    MsgBox "Copia Finalizada Sin Problema", vbInformation + vbOKOnly, MsgTitulo
Case 4
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
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
End Select
End Sub

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
