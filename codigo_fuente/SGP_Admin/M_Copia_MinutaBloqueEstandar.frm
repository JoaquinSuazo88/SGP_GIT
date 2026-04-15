VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_Copia_MinutaBloqueEstandar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copia Minuta Bloque Estandar"
   ClientHeight    =   10020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   16125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
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
      Height          =   4695
      Left            =   105
      TabIndex        =   16
      Top             =   5235
      Width           =   15870
      Begin VB.Frame Frame6 
         Height          =   435
         Left            =   2640
         TabIndex        =   26
         Top             =   3600
         Width           =   6870
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   27
            Top             =   135
            Width           =   6765
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Index           =   0
         Left            =   1560
         TabIndex        =   24
         Top             =   3600
         Width           =   900
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   25
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Procesar Copia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   13090
         TabIndex        =   7
         Top             =   4050
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
         Height          =   540
         Left            =   14440
         TabIndex        =   8
         Top             =   4050
         Width           =   1275
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1485
         TabIndex        =   4
         Top             =   210
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
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   2880
         Left            =   105
         TabIndex        =   5
         Top             =   630
         Width           =   15615
         _Version        =   393216
         _ExtentX        =   27543
         _ExtentY        =   5080
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
         MaxCols         =   5
         SpreadDesigner  =   "M_Copia_MinutaBloqueEstandar.frx":0000
      End
      Begin EditLib.fpDateTime FPFecDestino 
         Height          =   315
         Left            =   1470
         TabIndex        =   6
         Top             =   4185
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   210
         Left            =   5700
         TabIndex        =   20
         Top             =   4380
         Visible         =   0   'False
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   1
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   4440
         TabIndex        =   22
         Top             =   4200
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
            Weight          =   700
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Largo de Días"
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
         Index           =   8
         Left            =   3120
         TabIndex        =   23
         Top             =   4290
         Width           =   1215
      End
      Begin VB.Label lbl_proceso 
         Alignment       =   2  'Center
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   11340
         TabIndex        =   21
         Top             =   3780
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Destino"
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
         Left            =   105
         TabIndex        =   19
         Top             =   4260
         Width           =   1245
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
         Index           =   1
         Left            =   210
         TabIndex        =   17
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Origenes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   1575
      TabIndex        =   9
      Top             =   105
      Width           =   12825
      Begin VB.Frame Frame2 
         Height          =   1170
         Left            =   630
         TabIndex        =   10
         Top             =   210
         Width           =   11775
         Begin EditLib.fpDateTime FpFecDesde 
            Height          =   315
            Left            =   2115
            TabIndex        =   1
            Top             =   705
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
            Left            =   9300
            TabIndex        =   2
            Top             =   705
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Left            =   2130
            TabIndex        =   0
            Top             =   285
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
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   360
            Left            =   10920
            TabIndex        =   18
            Top             =   630
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
            Caption         =   "Fecha Hasta"
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
            Left            =   7995
            TabIndex        =   14
            Top             =   795
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde"
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
            Left            =   840
            TabIndex        =   13
            Top             =   795
            Width           =   1140
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   3420
            Picture         =   "M_Copia_MinutaBloqueEstandar.frx":1919
            Top             =   210
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
            Left            =   3870
            TabIndex        =   12
            Top             =   285
            Width           =   6735
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
            Left            =   855
            TabIndex        =   11
            Top             =   390
            Width           =   735
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
            Left            =   3915
            TabIndex        =   15
            Top             =   330
            Width           =   6735
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3195
         Left            =   315
         TabIndex        =   3
         Top             =   1575
         Width           =   12255
         _Version        =   393216
         _ExtentX        =   21616
         _ExtentY        =   5636
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
         MaxCols         =   10
         SpreadDesigner  =   "M_Copia_MinutaBloqueEstandar.frx":1C23
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
            Picture         =   "M_Copia_MinutaBloqueEstandar.frx":378A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_Copia_minutaBloqueEstandar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lc_Aux As String
Dim MsgTitulo As String

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim RS          As New ADODB.Recordset
Dim seleccion   As String
Dim seleccionx  As String
Dim Fecha       As Date
Dim i           As Long
Dim j           As Long
Dim X           As Long
Dim CecoOrigen  As String
Dim CecoDestino As String
Dim Regimen     As Long
Dim Servicio    As Long
Dim FecIni      As String
Dim FecFin      As String
Dim Bloque      As String
Dim Conta       As Long
Dim Sql         As String
Dim EstCopiado  As Boolean
Dim LargoDia    As Long
Dim FechaDesFin As Date
Dim Id_Bloque   As Long

 '-------> Validar largo días
 If Val(fpLongInteger1(0).Value) < 1 Then
 
     MsgBox "Largo de días debe ser mayor que cero...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
 
 End If
 
 If DateDiff("d", FpFecDesde.text, FpFecHasta) + 1 > Val(fpLongInteger1(0).Value) Then
 
     MsgBox "Largo de días origen es mayor al largo destino...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
 
 End If
 
 If Val(fpLongInteger1(0).Value) > 98 Then
 
     MsgBox "Maximo de días corresponde 98 días...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
 
 End If
 
 '-------> Validar fechas
  If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
     
     MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If
    
  If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
     
     MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If

  '-------> Validar que el día desde coincida con dia despacho
'  If DatePart("w", FpFecDesde, 2) <> DatePart("w", FPFecDestino, 2) Then
'
'     MsgBox "El día de la fecha desde deberá corresponder al día de la fecha destino. " & VgLinea & VgLinea & "                EJ: lunes (F. destino) = lunes (F.desde)", vbExclamation + vbOKOnly, MsgTitulo
'     Exit Sub
'
'  End If
  
  '-------> Validar que exista Ceco
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
  Set RS = vg_db.Execute("sgpadm_s_cliente_V02 45, '" & LimpiaDato(fpText.text) & "', ''")
  If RS.EOF Then
     
     RS.Close
     Set RS = Nothing
     MsgBox "No existe Ceco seleccionado...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If
   
 '-------> Validar fechas
  If Trim(FpFecDestino.text) = "" Then
     
     MsgBox "Fecha destino esta nula...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If
    
'  If vaSpread2.MaxRows > 1 Then
'     vaSpread2.Row = 1
'     vaSpread2.Col = 8
'     Fecha = vaSpread2.text
'  End If
  
  If Format(FpFecDestino, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then

     MsgBox "La fecha destino debe ser mayor que la fecha desde seleccionada en detalle grilla...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub

  End If

  If vaSpread1.MaxRows < 1 Then
     
     MsgBox "Debe seleccionar datos del encabezado...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If
  '-------> Validar que exista un dato seleccionado encabezado
  seleccion = 0
  For i = 1 To vaSpread1.MaxRows
       
       vaSpread1.Row = i
       vaSpread1.Col = 1 'Seleccion
       seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
       If seleccion = 1 And vaSpread1.RowHidden = False Then
          
          Exit For
       
       End If
  
  Next i
  
  If seleccion = 0 Then
     
     MsgBox " Se debe seleccionar un Bloque por lo menos", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If

  seleccion = 0
  Conta = 0
  For i = 1 To vaSpread2.MaxRows
       
       vaSpread2.Row = i
       vaSpread2.Col = 1 'Seleccion
       seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
       If seleccion = 1 And vaSpread2.RowHidden = False Then
          
          Conta = Conta + 1
       
       End If
       
       vaSpread2.Col = 4
       vaSpread2.text = ""
  
  Next i

  ProgressBar1.Scrolling = ccScrollingSmooth
  ProgressBar1.Max = 100
  ProgressBar1.Visible = True
  ProgressBar1.Value = 0
  lbl_proceso.Caption = "0 %"
  lbl_proceso.Visible = True
  
  Toolbar2.Enabled = False
  FpFecDesde.Enabled = False
  FpFecHasta.Enabled = False
  FpFecDestino.Enabled = False
  fpText.Enabled = False
  fpText1.Enabled = False

  fg_carga ""
  EstCopiado = True
  
  CecoOrigen = LimpiaDato(fpText.text)
  
  j = 1
  For i = 1 To vaSpread2.MaxRows
  
      vaSpread2.Row = i
      vaSpread2.Col = 1 'Seleccion
      seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
      If seleccion = 1 And vaSpread2.RowHidden = False Then
          
          Ceco = ""
          vaSpread2.Col = 2
          CecoDestino = vaSpread2.text
          
          For X = 1 To vaSpread1.MaxRows
              
              vaSpread1.Row = X
              vaSpread1.Col = 1 'Seleccion
              seleccionx = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
              If seleccionx = 1 And vaSpread1.RowHidden = False Then
              
'                 Id_Bloque = 0
'                 vaSpread1.Col = 2
'                 Id_Bloque = vaSpread1.text
          
                 Regimen = 0
                 vaSpread1.Col = 7
                 Regimen = vaSpread1.text

                 Servicio = 0
                 vaSpread1.Col = 8
                 Servicio = vaSpread1.text

                 FecIni = ""
                 vaSpread1.Col = 5
                 FecIni = Format(vaSpread1.text, "yyyymmdd")

                 FecFin = ""
                 vaSpread1.Col = 6
                 FecFin = Format(vaSpread1.text, "yyyymmdd")
          
                 '-------> Validar días dentro de una minuta
                 LargoDia = DateDiff("d", CDate(fg_Ctod1(FecIni)), CDate(fg_Ctod1(FecFin))) + 1
                 EstCopiado = True
                 vaSpread2.Col = 4
                 vaSpread2.text = ""
         
                 vaSpread2.Row = i
                 vaSpread2.Col = -1
                 vaSpread2.BackColor = &HC0FFFF
         
                 Sql = ""
                 Sql = CecoDestino
                 Sql = Sql & ", " & Regimen & ", " & Servicio & ", " & Format(FpFecDestino, "YYYYMMDD") & ", " & fpLongInteger1(0).text & ""
                 
                 If RS.State = 1 Then RS.Close
                 RS.CursorLocation = adUseClient
                 vg_db.CursorLocation = adUseClient
                 Set RS = vg_db.Execute("sgpadm_Sel_ValidarMinutaBloqueLargoDias_v01 " & Sql & "")
                 If Not RS.EOF Then
                    
                    If (RS(0) = "S" And RS(1) = "N") Or (RS(0) = "N" And RS(1) = "S") Then
                       
                       vaSpread2.Row = i
                       vaSpread2.Col = -1
                       vaSpread2.BackColor = &H8080FF
                
                       vaSpread2.Col = 4
                       vaSpread2.text = "Existen minuta bloque para este periodo, proceso cancelado " & "( Bloque = " & IIf(RS(2) = 0, RS(3), RS(2)) & " del Periodo " & RS(4) & " Hasta " & RS(5) & ")"
                       EstCopiado = False
                    
                    ElseIf (RS(0) = "S" And RS(1) = "S") Then
               
                       vaSpread2.Row = i
                       vaSpread2.Col = -1
                       vaSpread2.BackColor = &H8080FF
                
                       vaSpread2.Col = 4
                       vaSpread2.text = "El largo de días no corresponde con el original, proceso cancelado " & " Bloque =  " & " " & RS(2) & " su largo corresponde = " & RS(3)
                       EstCopiado = False
            
                    
                    ElseIf RS(2) <> RS(3) And RS(2) > 0 And RS(3) > 0 Then
                       
                       vaSpread2.Row = i
                       vaSpread2.Col = -1
                       vaSpread2.BackColor = &H8080FF
               
                       vaSpread2.Col = 4
                       vaSpread2.text = "Existen minuta bloque para este periodo, proceso cancelado " & "( Bloque = " & RS(3) & " del Periodo " & RS(4) & " Hasta " & RS(5) & ")"
                       EstCopiado = False
                 
                 End If
              
              End If
              RS.Close
              Set RS = Nothing
              
              If EstCopiado Then
                 
                 '-------> validacion en relacion al estado de la minuta
                 EstCopiado = True
                 vaSpread2.Col = 11
                 vaSpread2.text = ""
         
                 vaSpread2.Row = i: vaSpread2.Col = -1
                 vaSpread2.BackColor = &HC0FFFF
                 FechaDesFin = DateAdd("d", fpLongInteger1(0).text - 1, Format(FpFecDestino, "DD/mm/yyyy"))
         
                 Sql = " sgpadm_p_copia_minValidaMinuta_MVI_V02 "
                 Sql = Sql & " '" & CecoDestino & "'" 'ceco destino
                 Sql = Sql & ", " & Regimen 'regimen destino
                 Sql = Sql & ", " & Servicio 'servicio destino
                 Sql = Sql & ", " & Format(FpFecDestino, "YYYYMMDD") ' fecha desde
                 Sql = Sql & ", " & Format(FechaDesFin, "YYYYMMDD") ' fecha hasta

                 If RS.State = 1 Then RS.Close
                 RS.CursorLocation = adUseClient
                 vg_db.CursorLocation = adUseClient
                 Set RS = vg_db.Execute(Sql)
                 If Not RS.EOF Then
                    
                    vaSpread2.Row = i
                    vaSpread2.Col = -1
                    vaSpread2.BackColor = &H8080FF
               
                    vaSpread2.Col = 4
                    vaSpread2.text = "Minuta bloque ya existe"
                    EstCopiado = False
               
                 End If
                 RS.Close
                 Set RS = Nothing
              
              End If
         
              If EstCopiado Then
                 
                 '-------> Proceso de copia
                 Sql = ""
                 Sql = " sgpadm_Ins_CopiaMinutaBloqueEstandar_V03 "
                 Sql = Sql & " '" & CecoOrigen & "' " 'ceco origen
'                 Sql = Sql & " ," & Id_Bloque & " " 'Id_Bloque
                 Sql = Sql & " ," & Regimen & " " 'regimen
                 Sql = Sql & " ," & Servicio & " " 'servicio
                 Sql = Sql & " ,'" & FecIni & "'" 'fecha desde
                 Sql = Sql & " ,'" & FecFin & "'" 'fecha hasta
                 Sql = Sql & " ,'" & CecoDestino & "'"  'Ceco destino
                 Sql = Sql & " ,'" & Format(FpFecDestino, "YYYYMMDD") & "'"  'fecha destino
                 Sql = Sql & " ," & fpLongInteger1(0).Value & ""
                 
                 If RS.State = 1 Then RS.Close
                 RS.CursorLocation = adUseClient
                 vg_db.CursorLocation = adUseClient
                 Set RS = vg_db.Execute(Sql)
                 If Not RS.EOF Then
                    
                    If RS(0) > 0 Then
                       
                       vaSpread2.Col = 4
                       vaSpread2.text = RS(0) & " " & RS(1)
                    
                    Else
                       
                       vaSpread2.Col = 4
                       vaSpread2.text = "Proceso finalizado sin problema"
                    
                    End If
                 
                 End If
                 RS.Close
                 Set RS = Nothing
            
              End If
              
              End If
          
          Next X

         ProgressBar1.Value = ((j / Conta) * 100)
         lbl_proceso.Caption = CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %"
         j = j + 1
         
      End If
  
      DoEvents
       
  Next i

  fg_descarga
  MsgBox "Proceso Finalizado", vbInformation, Me.Caption
                
  ProgressBar1.Visible = False
  lbl_proceso.Visible = False
  
  Toolbar2.Enabled = True
  FpFecDesde.Enabled = True
  FpFecHasta.Enabled = True
  FpFecDestino.Enabled = True
  fpText.Enabled = True
  fpText1.Enabled = True
                
Exit Sub
Man_Error:
    
    fg_descarga
    
    ProgressBar1.Visible = False
    lbl_proceso.Visible = False
    
    Toolbar2.Enabled = True
    FpFecDesde.Enabled = True
    FpFecHasta.Enabled = True
    FpFecDestino.Enabled = True
    fpText.Enabled = True
    fpText1.Enabled = True
    
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub Command2_Click()

On Error GoTo Man_Error

'-------> Salir de la opción
Unload Me

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
fg_centra Me
MsgTitulo = "Copiar Minuta Bloque Estandar"

FpFecHasta.text = Format(Date, "dd/mm/yyyy")
FpFecDesde.text = Format(Date, "dd/mm/yyyy")
FpFecDestino.text = Format(Date, "dd/mm/yyyy")
vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0

vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &HC0FFFF
vaSpread1.MaxRows = 0

vaSpread2.Row = -1: vaSpread2.Col = -1
vaSpread2.BackColor = &HC0FFFF
vaSpread2.MaxRows = 0

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

vaSpread2.MaxRows = 0
TextDet2(2).text = ""
TextDet2(3).text = ""
If IsDate(FpFecDesde.text) = False Then Exit Sub

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FPFecDestino_Change()

On Error GoTo Man_Error

'vaSpread2.MaxRows = 0
If IsDate(FpFecDestino.text) = False Then Exit Sub

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FPFecDestino_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

vaSpread2.MaxRows = 0
If IsDate(FpFecHasta.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpText_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS = vg_db.Execute("sgpadm_s_cliente_V02 45, '" & LimpiaDato(fpText.text) & "', ''")
    If RS.EOF Then
        
        RS.Close
        Set RS = Nothing
        fpayuda(0).Caption = ""
        vaSpread1.MaxRows = 0
        vaSpread2.MaxRows = 0
        Exit Sub
    
    End If
    fpayuda(0).Caption = Trim(IIf(IsNull(RS!Cli_nombre), "", RS!Cli_nombre))
    fpText.text = RS!Cli_codigo
    RS.Close
    Set RS = Nothing
    vaSpread1.MaxRows = 0
    vaSpread2.MaxRows = 0
 
Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpText1_Change()

On Error GoTo Man_Error

LlenarGrillaCeco

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpText1_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

    Select Case Index
        
        Case 0
            
            vg_left = fpayuda(0).Left + 2300
            vg_nombre = "": vg_codigo = ""
            Call B_TabEst.LlenaDatos("b_clientes", "cli_", "Clientes", "Cliente_SitioRemoto")
            Call B_TabEst.Show(1)
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpText.text = vg_codigo
            fpayuda(0).Caption = vg_nombre
            FpFecDesde.SetFocus
    
    End Select
Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub LlenarGrillaCeco()

On Error GoTo Man_Error

Dim RS        As New ADODB.Recordset
Dim OrgCompra As String
Dim Sql       As String

vaSpread2.Row = -1: vaSpread2.Col = -1
vaSpread2.BackColor = &HC0FFFF
vaSpread2.MaxRows = 0

' Control displays text tips aligned to pointer with focus
vaSpread2.TextTip = 2
vaSpread2.TextTipDelay = 250
X = vaSpread2.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)

OrgCompra = LimpiaDato(Trim(fpText1.text))

Sql = ""
Sql = " sgpadm_Sel_OrgComprasxCeco "
Sql = Sql & " '" & OrgCompra & "'"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

Do While Not RS.EOF
   
   vaSpread2.MaxRows = vaSpread2.MaxRows + 1
   vaSpread2.Row = vaSpread2.MaxRows
   
   vaSpread2.Col = 1
   vaSpread2.text = "0"
   
   vaSpread2.Col = 2
   vaSpread2.text = RS!Cli_codigo
   
   vaSpread2.Col = 3
   vaSpread2.text = Trim(IIf(IsNull(RS!Cli_nombre), "", RS!Cli_nombre))
   
   vaSpread2.Col = 5
   vaSpread2.text = 0
   
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing

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

ElseIf Index = 3 Then
   
   TextDet2(2).text = ""

End If

For i = 1 To vaSpread2.MaxRows
           
    vaSpread2.Row = i
    vaSpread2.Col = 5
    vaSpread2.text = 0
    
Next

Select Case Index

Case 2, 3
    
    vaSpread2.Visible = False
    
    If Trim(TextDet2(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread2.MaxRows
           
           vaSpread2.Row = i
           vaSpread2.Col = Index
           indactivo = UCase(Trim(vaSpread2.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread2.Col = 1
           
           If indactivo = -1 And Trim(vaSpread2.text) <> "" Then
              
              vaSpread2.Col = 5
              
              If Val(vaSpread2.Value) <> 1 Then
                              
                 vaSpread2.Col = 1
              
                 If vaSpread2.RowHidden = True Then
                 
                    vaSpread2.RowHidden = False
                    vaSpread2.Col = 5
                    vaSpread2.text = 1
                 
                 Else
                 
                    vaSpread2.Col = 5
                    vaSpread2.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread2.Col = 5
              EstBuq = vaSpread2.Value
              vaSpread2.Col = 1
              
              If vaSpread2.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread2.RowHidden = True
                 
                 vaSpread2.Col = 5
                 vaSpread2.text = 0
                 
              End If
           
           End If
        
        Next i
        
        vaSpread2.SetActiveCell Index + 1, 1
        vaSpread2.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread2.ColUserSortIndicator(IIf(Trim(wvarArr8020(X)) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread2.SortKey(1) = IIf(Trim(wvarArr8020(X)) = "", 0, 0): vaSpread2.SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread2.Sort -1, -1, vaSpread2.maxcols, vaSpread2.MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(TextDet2(Index).text) = "" Then
       
       For i = 1 To vaSpread2.MaxRows
           
           vaSpread2.Row = i
           If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
           
           vaSpread2.Col = 5
           vaSpread2.text = 0
       
       Next
       
       vaSpread2.SetActiveCell Index, vaSpread2.SearchCol(Index, 0, vaSpread2.MaxRows, Trim(TextDet2(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread2.SetActiveCell Index, 1
    
    End If
    
    vaSpread2.Visible = True

End Select

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS        As New ADODB.Recordset
Dim Sql       As String
Dim i         As Long
Dim xmlceco   As String
Dim seleccion As String
Dim codCeco   As String

Select Case Button.Index

Case 1

  '-------> Validar fechas
  If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
     
     MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If
    
  If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
     
     MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If

  '-------> Validar que exista Ceco
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   Set RS = vg_db.Execute("sgpadm_s_cliente_V02 45, '" & LimpiaDato(fpText.text) & "', ''")
   If RS.EOF Then
      
      RS.Close
      Set RS = Nothing
      MsgBox "No existe Ceco seleccionado...", vbExclamation + vbOKOnly, MsgTitulo
      Exit Sub
   
   End If
   
   vaSpread1.MaxRows = 0
   fpText1.text = ""
   vaSpread2.MaxRows = 0
   vaSpread1.Row = -1: vaSpread1.Col = -1
   vaSpread1.BackColor = &HC0FFFF
   
   '-------> Rescata Ceco Seleccionado
   seleccion = 0
   xmlceco = ""
   xmlceco = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
   xmlceco = xmlceco & "<Ce>"
  
   xmlceco = xmlceco & "<C"
   codCeco = LimpiaDato(fpText.text)
   xmlceco = xmlceco & " c = " & Chr(34) & codCeco & Chr(34)
   xmlceco = xmlceco & "/>"
   xmlceco = xmlceco & "</Ce>"
   
   Sql = ""
   Sql = Sql & "'" & xmlceco & "', "
   Sql = Sql & "" & Format(FpFecDesde, ("YYYYMMDD")) & ", "
   Sql = Sql & "" & Format(FpFecHasta, ("YYYYMMDD")) & " "
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   Set RS = vg_db.Execute("sgpadm_Sel_XmlDetalleMinutaBloque_V02 " & Sql & "")
   If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      
      vaSpread1.Col = 1
      vaSpread1.text = "0"
      
'      vaSpread1.Col = 2
'      vaSpread1.text = RS!Id_Bloque
      
      vaSpread1.Col = 3
      vaSpread1.text = RS!Regimen & " - " & Trim(RS!reg_nombre)
      
      vaSpread1.Col = 4
      vaSpread1.text = RS!Servicio & " - " & Trim(RS!ser_nombre)
      
      vaSpread1.Col = 5
      vaSpread1.text = RS!fechadesde
         
      vaSpread1.Col = 6
      vaSpread1.text = RS!fechahasta
         
      vaSpread1.Col = 7
      vaSpread1.text = RS!Regimen
         
      vaSpread1.Col = 8
      vaSpread1.text = RS!Servicio
         
      RS.MoveNext
   
   Loop
   
   Else
      
      vaSpread1.MaxRows = 0
      MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo
   
   End If
   RS.Close
   Set RS = Nothing

End Select

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1
    Dim i As Long
    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
    For i = 1 To vaSpread1.MaxRows 'BlockRow To BlockRow2
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    
        End If
        
    Next
    
End Select

'vaSpread2.MaxRows = 0
'Command1.Enabled = False

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1
    Dim i As Long
    vaSpread2.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread2.Row = i
        
        If vaSpread2.RowHidden = False Then
            
           vaSpread2.Value = IIf(vaSpread2.Value = "1", "0", "1")
        
        End If
    
    Next
    
    For i = 1 To vaSpread2.MaxRows 'BlockRow To BlockRow2
        
        vaSpread2.Row = i
        
        If vaSpread2.RowHidden = False Then
            
           vaSpread2.Value = IIf(vaSpread2.Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    vaSpread2.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread2.Row = i
        If vaSpread2.RowHidden = False Then
            
           vaSpread2.Value = IIf(vaSpread2.Value = "1", "0", "1")
    
        End If
        
    Next
    
End Select

Exit Sub
Man_Error:

    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub
