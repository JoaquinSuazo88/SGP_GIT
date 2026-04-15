VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_CopProCan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar datos de rutas"
   ClientHeight    =   4875
   ClientLeft      =   2325
   ClientTop       =   2595
   ClientWidth     =   8925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   8925
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8055
      Begin VB.Frame Frame2 
         Caption         =   "Opción de copiado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   21
         Top             =   2280
         Width           =   7575
         Begin VB.CheckBox Check1 
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
            Height          =   195
            Index           =   1
            Left            =   2040
            TabIndex        =   27
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Familia"
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
            Left            =   360
            TabIndex        =   26
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todo (Productos, Fechas y Casino)"
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
            Left            =   5280
            TabIndex        =   25
            Top             =   400
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Casino"
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
            Left            =   3600
            TabIndex        =   24
            Top             =   400
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
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
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   23
            Top             =   400
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
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
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   22
            Top             =   400
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   1
         Top             =   780
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
         Index           =   0
         Left            =   1800
         TabIndex        =   0
         Top             =   435
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   2400
         TabIndex        =   2
         Top             =   1320
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   6360
         TabIndex        =   3
         Top             =   1320
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   2
         Left            =   2400
         TabIndex        =   4
         Top             =   1800
         Visible         =   0   'False
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   3
         Left            =   6360
         TabIndex        =   5
         Top             =   1800
         Visible         =   0   'False
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2700
         Picture         =   "M_CopProCan.frx":0000
         Top             =   690
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2700
         Picture         =   "M_CopProCan.frx":030A
         Top             =   360
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3165
         TabIndex        =   17
         Top             =   780
         Width           =   4455
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3165
         TabIndex        =   16
         Top             =   435
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "reemplazados a partir de la definición de la ruta de Origen."
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
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   5010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Recuerde que al seleccionar la opción copiar de fechas o todos los datos existentes serán "
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
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   7800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nota :"
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
         Left            =   120
         TabIndex        =   13
         Top             =   3360
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo Destino Hasta"
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
         Left            =   4320
         TabIndex        =   12
         Top             =   1920
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo Destino Desde"
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
         Left            =   360
         TabIndex        =   11
         Top             =   1920
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo Destino"
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
         Left            =   4320
         TabIndex        =   10
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo Origen"
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
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Destino"
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
         Left            =   360
         TabIndex        =   8
         Top             =   855
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Origen"
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
         Left            =   360
         TabIndex        =   7
         Top             =   525
         Width           =   1035
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3195
         TabIndex        =   18
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3195
         TabIndex        =   19
         Top             =   840
         Width           =   4455
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   4875
      Left            =   8295
      TabIndex        =   20
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   8599
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_CopProCan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim MsgTitulo As String
Dim opcion As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Dim codTippla As Long, nomTippla As String
fg_centra Me
fg_carga ""
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar ": BtnX.Enabled = True
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpDateTime1(0).DateTimeFormat = UserDefined
fpDateTime1(0).UserDefinedFormat = "dd/mm/yyyy"
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).DateTimeFormat = UserDefined
fpDateTime1(1).UserDefinedFormat = "dd/mm/yyyy"
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(2).DateTimeFormat = UserDefined
fpDateTime1(2).UserDefinedFormat = "dd/mm/yyyy"
fpDateTime1(2).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(3).DateTimeFormat = UserDefined
fpDateTime1(3).UserDefinedFormat = "dd/mm/yyyy"
fpDateTime1(3).text = Format(Date, "dd/mm/yyyy")
Label1(4).Caption = IIf(opcion = "ruta", "Ruta Origen", "Origen")
Label1(0).Caption = IIf(opcion = "ruta", "Ruta Destino", "Destino")
Label1(1).Visible = IIf(opcion = "ruta", True, False)
Label1(2).Visible = IIf(opcion = "ruta", True, False)
Label1(8).Caption = IIf(opcion = "ruta", "reemplazados a partir de la definición de la Ruta de Origen.", "reemplazados a partir de la definición de las Reglas de Negocios.")
fpDateTime1(0).Visible = IIf(opcion = "ruta", True, False)
fpDateTime1(1).Visible = IIf(opcion = "ruta", True, False)
Option1(0).Visible = IIf(opcion = "ruta", True, False)
Option1(1).Visible = IIf(opcion = "ruta", True, False)
Option1(2).Visible = IIf(opcion = "ruta", True, False)
Option1(3).Visible = IIf(opcion = "ruta", True, False)
Check1(0).Visible = IIf(opcion = "ruta", False, True)
Check1(1).Visible = IIf(opcion = "ruta", False, True)
Check1(0).Top = 400
Check1(1).Top = 400
fg_descarga
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Select Case Index
Case 0
    If Val(fpLongInteger1(0).Value) < 1 Then fpayuda(0).Caption = "": Exit Sub
    If opcion = "ruta" Then
       Set RS = vg_dbpedweb.Execute("SELECT recorrido, descripcion FROM s_Recorrido WHERE recorrido = " & Val(fpLongInteger1(0).Value) & "")
    Else
       Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegocios 4, " & Val(fpLongInteger1(0).Value) & ", ''")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing:: fpayuda(0).Caption = "": fpLongInteger1(1).Value = "": fpayuda(1).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS(1))
    RS.Close: Set RS = Nothing
Case 1
    If Val(fpLongInteger1(1).Value) < 1 Then fpayuda(1).Caption = "": Exit Sub
    If opcion = "ruta" Then
       Set RS = vg_dbpedweb.Execute("SELECT recorrido, descripcion FROM s_Recorrido WHERE recorrido = " & Val(fpLongInteger1(1).Value) & "")
    Else
       Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegocios 4, " & Val(fpLongInteger1(1).Value) & ", ''")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS(1))
    RS.Close: Set RS = Nothing
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "s_Recorrido", "sub_", "Ruta", IIf(opcion = "ruta", "recorrido", "regneg")
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
    fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(1).SetFocus
Case 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "s_Recorrido", "reg_", "Ruta", IIf(opcion = "ruta", "recorrido", "regneg")
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    If fpDateTime1(0).Enabled = True Then fpDateTime1(0).SetFocus:
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    fpDateTime1(0).Enabled = False
    fpDateTime1(1).Enabled = False
    fpDateTime1(2).Enabled = False
    fpDateTime1(3).Enabled = False
Case 1
    fpDateTime1(0).Enabled = True
    fpDateTime1(1).Enabled = True
    fpDateTime1(2).Enabled = True
    fpDateTime1(3).Enabled = True
Case 2
    fpDateTime1(0).Enabled = True
    fpDateTime1(1).Enabled = True
    fpDateTime1(2).Enabled = False
    fpDateTime1(3).Enabled = False
Case 3
    fpDateTime1(0).Enabled = True
    fpDateTime1(1).Enabled = True
    fpDateTime1(2).Enabled = True
    fpDateTime1(3).Enabled = True
End Select
End Sub

Sub LlenaDatos(tit As String, Op As String)
MsgTitulo = tit
opcion = Op
Me.Caption = tit
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim borpro As String
Select Case Button.Index
Case 1
    If Val(fpLongInteger1(0).Value) = 0 Or Val(fpLongInteger1(1).Value) = 0 Then MsgBox IIf(opcion = "ruta", "Debe seleccionar rutas ...", "Debe seleccionar regla de negocios ..."), vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If opcion = "ruta" Then
       If Option1(0).Value = True Then '-------> copiar ruta producto
          If Val(fpLongInteger1(0).Value) = Val(fpLongInteger1(1).Value) Then MsgBox "Código Origen debe ser distinto Destino ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
          '-------> Validar dato origen
          Set RS = vg_dbpedweb.Execute("pedweb_s_rutaproductos 3, " & Val(fpLongInteger1(0).Value) & ", '', ''")
          If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe datos producto en la ruta origen ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
          RS.Close: Set RS = Nothing
          '-------> validar dato destino
          Set RS = vg_dbpedweb.Execute("pedweb_s_rutaproductos 3, " & Val(fpLongInteger1(1).Value) & ", '', ''")
          If Not RS.EOF Then If MsgBox("Existe información productos destino. se borrara la información existente ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then RS.Close: Set RS = Nothing: Exit Sub
          RS.Close: Set RS = Nothing
          '-------> Proceso de copiado
          Set RS = vg_dbpedweb.Execute("pedweb_p_copiarutaproductos " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & "")
          If Not RS.EOF Then
             MsgBox "Proceso Finalizo [OK] ...", vbExclamation + vbOKOnly, MsgTitulo
          Else
             MsgBox "Proceso Finalizo con problema, reintente mas tarde ...", vbExclamation + vbOKOnly, MsgTitulo
          End If
          RS.Close: Set RS = Nothing
       ElseIf Option1(1).Value = True Then '-------> copiar ruta calendario
          If fpDateTime1(0).text = fpDateTime1(1).text And fpLongInteger1(0).Value = fpLongInteger1(1).Value Then MsgBox "Fecha origen Deben ser diferentes fecha destino ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
          '-------> Validar dato origen
'         Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendario 3, " & Val(fpLongInteger1(0).Value) & ", '', '', '" & Format(fpDateTime1(0).text, "yyyymmdd") & "'")
         Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendario 3, " & Val(fpLongInteger1(0).Value) & ", '', '', '" & fpDateTime1(0).text & "'")
         If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe datos calendario en la ruta origen ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
         RS.Close: Set RS = Nothing
         '-------> validar dato destino
'         Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendario 3, " & Val(fpLongInteger1(1).Value) & ", '', '', '" & Format(fpDateTime1(2).text, "yyyymmdd") & "'")
         Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendario 3, " & Val(fpLongInteger1(1).Value) & ", '', '', '" & fpDateTime1(2).text & "'")
         If Not RS.EOF Then If MsgBox("Existe información calendario destino. se borrara la información existente ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then RS.Close: Set RS = Nothing: Exit Sub
         RS.Close: Set RS = Nothing
         '-------> Proceso de copiado
         Set RS = vg_dbpedweb.Execute("pedweb_p_copiarutacalendario " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", '" & Format(fpDateTime1(0).text, "yyyyMMdd") & "', '" & Format(fpDateTime1(1).text, "yyyyMMdd") & "'")
         If Not RS.EOF Then
            MsgBox "Proceso Finalizo [OK] ...", vbExclamation + vbOKOnly, MsgTitulo
         Else
            MsgBox "Proceso Finalizo con problema, reintente mas tarde ...", vbExclamation + vbOKOnly, MsgTitulo
         End If
         RS.Close: Set RS = Nothing
    
       ElseIf Option1(2).Value = True Then '-------> copiar ruta casino
         If fpDateTime1(0).text = fpDateTime1(1).text And fpLongInteger1(0).Value = fpLongInteger1(1).Value Then MsgBox "Fecha origen Deben ser diferentes fecha destino ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
         '-------> Validar dato origen calendario
'         Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendario 3, " & Val(fpLongInteger1(0).Value) & ", '', '', '" & Format(fpDateTime1(0).text, "yyyymmdd") & "'")
         Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendario 3, " & Val(fpLongInteger1(0).Value) & ", '', '', '" & fpDateTime1(0).text & "'")
         If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe datos calendario en la ruta origen ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
         RS.Close: Set RS = Nothing
         '-------> Validar dato origen casino
         Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendariocasino 1, " & Val(fpLongInteger1(0).Value) & ", '" & Format(fpDateTime1(0).text, "yyyymmdd") & "', ''")
         If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe datos calendario casino origen ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
         RS.Close: Set RS = Nothing
         '-------> Validar dato destino calendario
'         Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendario 3, " & Val(fpLongInteger1(1).Value) & ", '', '', '" & Format(fpDateTime1(1).text, "yyyymmdd") & "'")
         Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendario 3, " & Val(fpLongInteger1(1).Value) & ", '', '', '" & fpDateTime1(1).text & "'")
         If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe datos calendario en la ruta destino ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
         RS.Close: Set RS = Nothing
         '-------> validar dato destino
         Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendariocasino 1, " & Val(fpLongInteger1(1).Value) & ", '" & Format(fpDateTime1(2).text, "yyyymmdd") & "', ''")
         If Not RS.EOF Then If MsgBox("Existe información calendario casino destino. se borrara la información existente ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then RS.Close: Set RS = Nothing: Exit Sub
         RS.Close: Set RS = Nothing
         '-------> Proceso de copiado
         Set RS = vg_dbpedweb.Execute("pedweb_p_copiarutacalendariocasino " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", '" & Format(fpDateTime1(0).text, "yyyyMMdd") & "', '" & Format(fpDateTime1(1).text, "yyyyMMdd") & "'")
         If Not RS.EOF Then
            MsgBox "Proceso Finalizo [OK] ...", vbExclamation + vbOKOnly, MsgTitulo
         Else
            MsgBox "Proceso Finalizo con problema, reintente mas tarde ...", vbExclamation + vbOKOnly, MsgTitulo
         End If
         RS.Close: Set RS = Nothing
    
       ElseIf Option1(3).Value Then '------->copiar las tres alternativas
         If Val(fpLongInteger1(0).Value) = Val(fpLongInteger1(1).Value) Then MsgBox "Código Origen debe ser distinto Destino ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
         If fpDateTime1(0).text = fpDateTime1(1).text And fpLongInteger1(0).Value = fpLongInteger1(1).Value Then MsgBox "Fecha origen Deben ser diferentes fecha destino ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
         '-------> Validar dato origen productos
         Set RS = vg_dbpedweb.Execute("pedweb_s_rutaproductos 3, " & Val(fpLongInteger1(0).Value) & ", '', ''")
         If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe datos producto en la ruta origen ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
         RS.Close: Set RS = Nothing
         '-------> Validar dato origen calendario
'         Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendario 3, " & Val(fpLongInteger1(0).Value) & ", '', '', '" & Format(fpDateTime1(0).text, "yyyymmdd") & "'")
         Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendario 3, " & Val(fpLongInteger1(0).Value) & ", '', '', '" & fpDateTime1(0).text & "'")
         If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe datos calendario en la ruta origen ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
         RS.Close: Set RS = Nothing
         '-------> Validar dato origen casino
         Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendariocasino 1, " & Val(fpLongInteger1(0).Value) & ", '" & Format(fpDateTime1(0).text, "yyyymmdd") & "', ''")
         If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe datos calendario casino origen ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
         RS.Close: Set RS = Nothing
         '-------> Proceso de copiado
         Set RS = vg_dbpedweb.Execute("pedweb_p_copiaproductocalendariocasino " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", '" & Format(fpDateTime1(0).text, "yyyymmdd") & "', '" & Format(fpDateTime1(1).text, "yyyymmdd") & "'")
         If Not RS.EOF Then
            MsgBox "Proceso Finalizo [OK] ...", vbExclamation + vbOKOnly, MsgTitulo
         Else
            MsgBox "Proceso Finalizo con problema, reintente mas tarde ...", vbExclamation + vbOKOnly, MsgTitulo
         End If
         RS.Close: Set RS = Nothing
    
       End If
    Else
       If Val(fpLongInteger1(0).Value) = Val(fpLongInteger1(1).Value) Then MsgBox "Código Origen debe ser distinto Destino ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       If Check1(0).Value = 0 And Check1(1).Value = 0 Then MsgBox "debe seleccionar a lo menos una opción ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       If Check1(0).Value = 1 Then
          '-------> Validar dato origen familia
          Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegociosfamilia 2, " & Val(fpLongInteger1(0).Value) & "")
          If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe datos reglas de negocios familia origen ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
          RS.Close: Set RS = Nothing
          '-------> validar dato destino
          Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegociosfamilia 2, " & Val(fpLongInteger1(1).Value) & "")
          If Not RS.EOF Then If MsgBox("Existe información familias destino. se borrara la información existente ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then RS.Close: Set RS = Nothing: Exit Sub
          RS.Close: Set RS = Nothing
          '-------> Proceso de copiado
          Set RS = vg_dbpedweb.Execute("pedweb_p_copiareglasdenegociosfamilia " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & "")
          If Not RS.EOF Then
             MsgBox "Proceso Familia Finalizo [OK] ...", vbExclamation + vbOKOnly, MsgTitulo
          Else
             MsgBox "Proceso Familia Finalizo con problema, reintente mas tarde ...", vbExclamation + vbOKOnly, MsgTitulo
          End If
          RS.Close: Set RS = Nothing
       End If
       If Check1(1).Value = 1 Then
          '-------> Validar dato origen producto
          Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegociosproducto 2, " & Val(fpLongInteger1(0).Value) & ", '', '', 0")
          If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe datos reglas de negocios familia origen ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
          RS.Close: Set RS = Nothing
          '-------> validar dato destino
          borpro = "S"
          Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegociosproducto 2, " & Val(fpLongInteger1(1).Value) & ", '', '', 0")
          If Not RS.EOF Then
             If MsgBox("Existe información familias destino. se borrara la información existente ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
                borpro = "N"
             Else
                borpro = "S"
             End If
          End If
'             RS.Close: Set RS = Nothing: Exit Sub
          RS.Close: Set RS = Nothing
          '-------> Proceso de copiado
          Set RS = vg_dbpedweb.Execute("pedweb_p_copiareglasdenegociosproducto " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", '" & borpro & "'")
          If Not RS.EOF Then
             MsgBox "Proceso Productos Finalizo [OK] ...", vbExclamation + vbOKOnly, MsgTitulo
          Else
             MsgBox "Proceso Productos Finalizo con problema, reintente mas tarde ...", vbExclamation + vbOKOnly, MsgTitulo
          End If
          RS.Close: Set RS = Nothing
       End If
    End If
Case 3
    Me.Hide
    Unload Me
End Select
End Sub
