VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form M_Minu05 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Minutas"
   ClientHeight    =   3855
   ClientLeft      =   1935
   ClientTop       =   1875
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3855
   ScaleWidth      =   7500
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Parametros Casino Destino"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   1
      Left            =   0
      TabIndex        =   20
      Top             =   2280
      Width           =   6855
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   6
         Left            =   2565
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   750
         Width           =   4125
         _Version        =   196608
         _ExtentX        =   7276
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
         BackColor       =   -2147483638
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
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   3
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
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   5
         Left            =   2565
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   435
         Width           =   4125
         _Version        =   196608
         _ExtentX        =   7276
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
         BackColor       =   -2147483638
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
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   3
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   4
         Left            =   1200
         TabIndex        =   4
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
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
         Index           =   5
         Left            =   1200
         TabIndex        =   5
         Top             =   1065
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
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
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   7
         Left            =   2565
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1065
         Width           =   4125
         _Version        =   196608
         _ExtentX        =   7276
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
         BackColor       =   -2147483638
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
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   3
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   3
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
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
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   26
         Top             =   1065
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
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
         Left            =   210
         TabIndex        =   25
         Top             =   750
         Width           =   705
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   210
         TabIndex        =   24
         Top             =   435
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2080
         Picture         =   "M_Minu05.frx":0000
         Top             =   340
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   2080
         Picture         =   "M_Minu05.frx":030A
         Top             =   680
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   2080
         Picture         =   "M_Minu05.frx":0614
         Top             =   980
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Parametro Casino Origen"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6855
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Postres"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   5280
         TabIndex        =   13
         Top             =   1560
         Width           =   1005
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Salad Bar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   2760
         TabIndex        =   12
         Top             =   1560
         Width           =   1485
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Estructura Fija"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   11
         Top             =   1560
         Width           =   1485
      End
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   2
         Left            =   2565
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   750
         Width           =   4125
         _Version        =   196608
         _ExtentX        =   7276
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
         BackColor       =   -2147483638
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
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   3
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
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   1
         Left            =   2565
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   435
         Width           =   4125
         _Version        =   196608
         _ExtentX        =   7276
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
         BackColor       =   -2147483638
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
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   3
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1200
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
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
         Index           =   2
         Left            =   1200
         TabIndex        =   2
         Top             =   1065
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
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
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   3
         Left            =   2565
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1065
         Width           =   4125
         _Version        =   196608
         _ExtentX        =   7276
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
         BackColor       =   -2147483638
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
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   3
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   1200
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
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
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   19
         Top             =   1065
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
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
         Left            =   210
         TabIndex        =   18
         Top             =   750
         Width           =   705
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
         Index           =   3
         Left            =   210
         TabIndex        =   17
         Top             =   435
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2080
         Picture         =   "M_Minu05.frx":091E
         Top             =   340
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2080
         Picture         =   "M_Minu05.frx":0C28
         Top             =   680
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2080
         Picture         =   "M_Minu05.frx":0F32
         Top             =   980
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1035
      Left            =   1560
      ScaleHeight     =   975
      ScaleWidth      =   4785
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   4845
      Begin MSComctlLib.ProgressBar gauge 
         Height          =   330
         Left            =   120
         TabIndex        =   8
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
         ForeColor       =   &H80000018&
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   4515
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   3855
      Left            =   6870
      TabIndex        =   6
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   6800
      ButtonWidth     =   1138
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_Minu05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConSql As ADODB.Recordset, Consql1 As ADODB.Recordset
Dim vdia As Long, ind_minuta As Long, ind_datadj As Long
Private Sub Form_Activate()
fg_descarga
End Sub
Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me
Mover_Botones

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cargando Tablas Anexa Copiar Plantilla Casino"
End Sub
Sub CopiarPlantillaCasino()

On Error GoTo Man_Error

fg_carga ""

Set ConSql = vg_db.Execute("select  * " & _
             "From Sdx_BloqueoMinutas " & _
             "where codigo_casino='" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "' " & _
             "and   codigo_segmento=0 " & _
             "and   codigo_pventa=" & Val(fpLongInteger1(4).Value) & " " & _
             "and   codigo_servicio=" & Val(fpLongInteger1(5).Value) & " " & _
             "and   fecha=0", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_bloqueominutas 1, '" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "', 0, " & Val(fpLongInteger1(4).Value) & ", " & Val(fpLongInteger1(5).Value) & ", 0, '', ''", , adCmdStoredProc)
If Not ConSql.EOF Then
   If ConSql!estado = 1 Then
      fg_descarga
      MsgBox "Minuta Esta bloqueda, Proceso Cancelado", vbCritical + vbOKOnly, "Copiar Minutas"
      Exit Sub
   ElseIf ConSql!estado = 2 Then
      fg_descarga
      MsgBox "Minuta Esta Liberada, Proceso Cancelado", vbCritical + vbOKOnly, "Copiar Minutas"
      Exit Sub
   End If
End If
ConSql.Close: Set ConSql = Nothing

If fpText(0).Text = fpText(1).Text And Val(fpLongInteger1(1).Value) = Val(fpLongInteger1(4).Value) And Val(fpLongInteger1(2).Value) = Val(fpLongInteger1(5).Value) Then
   fg_descarga
   MsgBox "Datos Del Casino Origen, Deben Ser Distinto Casino Destino", vbCritical + vbOKOnly, "Copiar Minutas"
   Exit Sub
End If

' *** Validar si Existe Datos De Origen ***
Set ConSql = vg_db.Execute("select distinct Sdx_EncMinutas.cod_casino, " & _
             "Sdx_EncMinutas.cod_regimen , Sdx_EncMinutas.cod_servicio " & _
             "From Sdx_EncMinutas " & _
             "where Sdx_EncMinutas.cod_casino='" & "00000" & LimpiaDato(Trim(fpText(0).Text)) & "' " & _
             "and  Sdx_EncMinutas.cod_regimen=" & Val(fpLongInteger1(1).Value) & " " & _
             "and  Sdx_EncMinutas.cod_servicio=" & Val(fpLongInteger1(2).Value) & " " & _
             "and  Sdx_EncMinutas.ind_borrado=0", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_minutas 13, 0, '" & "00000" & LimpiaDato(Trim(fpText(0).Text)) & "', 0, 0, " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", 0, 0, '', ''", , adCmdStoredProc)
If ConSql.EOF Then
   fg_descarga
   MsgBox "No Existe Información Casino Origen", vbCritical + vbOKOnly, "Copiar Minutas"
   ConSql.Close: Set ConSql = Nothing
   Exit Sub
End If
ConSql.Close: Set ConSql = Nothing

' *** Validar si Existe Datos De Destino ***

Set ConSql = vg_db.Execute("select distinct Sdx_EncMinutas.cod_casino, " & _
             "Sdx_EncMinutas.cod_regimen , Sdx_EncMinutas.cod_servicio " & _
             "From Sdx_EncMinutas " & _
             "where Sdx_EncMinutas.cod_casino='" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "' " & _
             "and  Sdx_EncMinutas.cod_regimen=" & Val(fpLongInteger1(4).Value) & " " & _
             "and  Sdx_EncMinutas.cod_servicio=" & Val(fpLongInteger1(5).Value) & " " & _
             "and  Sdx_EncMinutas.ind_borrado=0", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_minutas 13, 0, '" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "', 0, 0, " & Val(fpLongInteger1(4).Value) & ", " & Val(fpLongInteger1(5).Value) & ", 0, 0, '', ''", , adCmdStoredProc)
If Not ConSql.EOF Then
   fg_descarga
   msg = " Existe Información Casino Destino. Se Borrara La Información Existente " & VgLinea & "                          ¿ Esta Seguro De Copiar ?"
   Style = vbYesNo + vbQuestion + vbDefaultButton2
   Help = "DEMO.HLP"
   Ctxt = 1000
   ws_respuesta = MsgBox(msg, Style, TITLE, Help, Ctxt)
   If ws_respuesta = vbNo Then ConSql.Close: Set ConSql = Nothing: Exit Sub
End If
ConSql.Close: Set ConSql = Nothing


' *** Grabando Plantilla Casino Origen Hacia Casino Origen *** '
vdia = 999999: ind_minuta = 0

fg_carga (ss)
gauge.Value = 0
Picture1.Visible = True: Label1(5).Visible = True: gauge.Visible = False
Picture1.Refresh
Label1(5).Caption = "Copiando Planificación Casino Menú Espere Un Momento ..."
vg_db.BeginTrans
  Set ConSql = vg_db.Execute("select Sdx_EncMinutas.dia_minuta, Sdx_DetMinutas.num_linea, " & _
               "Sdx_DetMinutas.tipo_minuta, Sdx_DetMinutas.cod_item, " & _
               "Sdx_DetMinutas.descripcion " & _
               "From Sdx_DetMinutas, Sdx_EncMinutas " & _
               "Where Sdx_EncMinutas.ind_minuta = Sdx_DetMinutas.ind_minuta " & _
               "and ((Sdx_EncMinutas.cod_casino='" & "00000" & LimpiaDato(Trim(fpText(0).Text)) & "') " & _
               "and  (Sdx_EncMinutas.cod_regimen=" & Val(fpLongInteger1(1).Value) & ") " & _
               "and  (Sdx_EncMinutas.cod_servicio=" & Val(fpLongInteger1(2).Value) & ") " & _
               "and  (Sdx_EncMinutas.ind_borrado=0) " & _
               "and  (Sdx_DetMinutas.ind_borrado=0)) " & _
               "order by Sdx_EncMinutas.dia_minuta", , adCmdText)
  If Not ConSql.EOF Then
     Do While Not ConSql.EOF
        If ConSql!dia_minuta <> vdia Then
                 
           ind_minuta = 0
          
           Set Consql1 = vg_db.Execute("select Sdx_EncMinutas.ind_minuta " & _
                         "From  Sdx_EncMinutas " & _
                         "where Sdx_EncMinutas.cod_casino='" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "' " & _
                         "and   Sdx_EncMinutas.cod_regimen=" & Val(fpLongInteger1(4).Value) & " " & _
                         "and   Sdx_EncMinutas.cod_servicio=" & Val(fpLongInteger1(5).Value) & " " & _
                         "and   Sdx_EncMinutas.dia_minuta=" & ConSql!dia_minuta & "", , adCmdText)
           If Not Consql1.EOF Then
             ind_minuta = Consql1!ind_minuta
             Consql1.Close: Set Consql1 = Nothing
             If ind_minuta > 0 Then
                vg_db.Execute "Delete Sdx_EncMinutas " & _
                              "from Sdx_EncMinutas " & _
                              "where cod_casino='" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "' " & _
                              "and   cod_regimen=" & Val(fpLongInteger1(4).Value) & " " & _
                              "and   cod_servicio=" & Val(fpLongInteger1(5).Value) & " " & _
                              "and   ind_minuta=" & ind_minuta & " " & _
                              "and   dia_minuta=" & ConSql!dia_minuta & ""
         
                vg_db.Execute "Delete Sdx_DetMinutas from Sdx_DetMinutas " & _
                              "where ind_minuta=" & ind_minuta & ""
             End If
           Else
              Consql1.Close: Set Consql1 = Nothing
              Set Consql1 = vg_db.Execute("select * from Sdx_Parametro holdlock where Parametro_Num=41", , adCmdText)
              If Not Consql1.EOF Then
                 vg_db.Execute "Update Sdx_Parametro Set Parametro_Val = Parametro_Val + 1 " & _
                               "Where Parametro_Num=41"
              Else
                 vg_db.Execute "insert into Sdx_Parametro (Parametro_Num, Parametro_Desc, Parametro_Val) " & _
                             "values (41, 'Parametro Nueva Planificación Minutas', 1)"
              End If
              Consql1.Close: Set Consql1 = Nothing
   
              Set Consql1 = vg_db.Execute("select Parametro_Val From Sdx_Parametro " & _
                           "Where Parametro_Num=41", , adCmdText)
              If Not Consql1.EOF Then
                 ind_minuta = Consql1!Parametro_Val
              End If
              Consql1.Close: Set Consql1 = Nothing
           End If
           vg_db.Execute "insert into Sdx_EncMinutas (cod_casino, cod_regimen, " & _
                         "cod_servicio, ind_minuta, dia_minuta, fecha_minuta, " & _
                         "op_minuta, ind_borrado) values " & _
                         "('" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "', " & _
                         "" & Val(fpLongInteger1(4).Value) & ", " & _
                         "" & Val(fpLongInteger1(5).Value) & ", " & _
                         "" & ind_minuta & ", " & ConSql!dia_minuta & ", " & _
                         "" & Format(Date, "yyyymm") & ", '0', 0)"
           
           vdia = ConSql!dia_minuta
        End If
        vg_db.Execute "insert into Sdx_DetMinutas (ind_minuta, num_linea, num_dia, " & _
                      "tipo_minuta, cod_item, descripcion, ind_borrado) values " & _
                      "(" & ind_minuta & ", " & ConSql!num_linea & ", " & _
                      "" & ConSql!dia_minuta & ", " & ConSql!tipo_minuta & ", " & _
                      "" & ConSql!cod_item & ", '" & ConSql!descripcion & "', 0)"
        ConSql.MoveNext
     Loop
  End If
  ConSql.Close: Set ConSql = Nothing
  
  ' *** Copiar Estructura Fijas *** '
  vdia = 999999: ind_minuta = 0
  If Check1(0).Value = 1 Then
     Set ConSql = vg_db.Execute("select Sdx_EncEstructuraFija.dia_estfija, Sdx_DetEstructuraFija.num_linea, " & _
                  "Sdx_DetEstructuraFija.tipo_estfija, Sdx_DetEstructuraFija.cod_item, " & _
                  "Sdx_DetEstructuraFija.descripcion " & _
                  "From Sdx_DetEstructuraFija, Sdx_EncEstructuraFija " & _
                  "Where Sdx_EncEstructuraFija.ind_estfija = Sdx_DetEstructuraFija.ind_estfija " & _
                  "and ((Sdx_EncEstructuraFija.cod_casino='" & "00000" & LimpiaDato(Trim(fpText(0).Text)) & "') " & _
                  "and  (Sdx_EncEstructuraFija.cod_regimen=" & Val(fpLongInteger1(1).Value) & ") " & _
                  "and  (Sdx_EncEstructuraFija.cod_servicio=" & Val(fpLongInteger1(2).Value) & ") " & _
                  "and  (Sdx_EncEstructuraFija.ind_borrado=0) " & _
                  "and  (Sdx_DetEstructuraFija.ind_borrado=0)) " & _
                  "order by Sdx_EncEstructuraFija.dia_estfija", , adCmdText)
     If Not ConSql.EOF Then
        Do While Not ConSql.EOF
           If ConSql!dia_estfija <> vdia Then
                 
              ind_minuta = 0
          
              Set Consql1 = vg_db.Execute("select Sdx_EncEstructuraFija.ind_estfija " & _
                            "From  Sdx_EncEstructuraFija " & _
                            "where Sdx_EncEstructuraFija.cod_casino='" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "' " & _
                            "and   Sdx_EncEstructuraFija.cod_regimen=" & Val(fpLongInteger1(4).Value) & " " & _
                            "and   Sdx_EncEstructuraFija.cod_servicio=" & Val(fpLongInteger1(5).Value) & " " & _
                            "and   Sdx_EncEstructuraFija.dia_estfija=" & ConSql!dia_estfija & "", , adCmdText)
              If Not Consql1.EOF Then
                 ind_minuta = Consql1!ind_estfija
                 Consql1.Close: Set Consql1 = Nothing
                 If ind_minuta > 0 Then
                    vg_db.Execute "Delete Sdx_EncEstructuraFija " & _
                                  "from Sdx_EncEstructuraFija " & _
                                  "where cod_casino='" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "' " & _
                                  "and   cod_regimen=" & Val(fpLongInteger1(4).Value) & " " & _
                                  "and   cod_servicio=" & Val(fpLongInteger1(5).Value) & " " & _
                                  "and   ind_estfija=" & ind_minuta & " " & _
                                  "and   dia_estfija=" & ConSql!dia_estfija & ""
         
                    vg_db.Execute "Delete Sdx_DetEstructuraFija from Sdx_DetEstructuraFija " & _
                                  "where ind_estfija=" & ind_minuta & ""
                 End If
              Else
                 Consql1.Close: Set Consql1 = Nothing
                 Set Consql1 = vg_db.Execute("select * from Sdx_Parametro holdlock where Parametro_Num=42", , adCmdText)
                 If Not Consql1.EOF Then
                    vg_db.Execute "Update Sdx_Parametro Set Parametro_Val = Parametro_Val + 1 " & _
                                  "Where Parametro_Num=42"
                 Else
                    vg_db.Execute "insert into Sdx_Parametro (Parametro_Num, Parametro_Desc, Parametro_Val) " & _
                                  "values (42, 'Parametro Estructuras Fijas', 1)"
                 End If
                 Consql1.Close: Set Consql1 = Nothing
   
                 Set Consql1 = vg_db.Execute("select Parametro_Val From Sdx_Parametro " & _
                               "Where Parametro_Num=42", , adCmdText)
                 If Not Consql1.EOF Then
                    ind_minuta = Consql1!Parametro_Val
                 End If
                 Consql1.Close: Set Consql1 = Nothing
              End If
              vg_db.Execute "insert into Sdx_EncEstructuraFija (cod_casino, cod_regimen, " & _
                            "cod_servicio, ind_estfija, dia_estfija, fecha_estfija, " & _
                            "op_estfija, ind_borrado) values " & _
                            "('" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "', " & _
                            "" & Val(fpLongInteger1(4).Value) & ", " & _
                            "" & Val(fpLongInteger1(5).Value) & ", " & _
                            "" & ind_minuta & ", " & ConSql!dia_estfija & ", " & _
                            "" & Format(Date, "yyyymm") & ", '0', 0)"
           
              vdia = ConSql!dia_estfija
           End If
           vg_db.Execute "insert into Sdx_DetEstructuraFija (ind_estfija, num_linea, num_dia, " & _
                         "tipo_estfija, cod_item, descripcion, ind_borrado) values " & _
                         "(" & ind_minuta & ", " & ConSql!num_linea & ", " & _
                         "" & ConSql!dia_estfija & ", " & ConSql!tipo_estfija & ", " & _
                         "" & ConSql!cod_item & ", '" & ConSql!descripcion & "', 0)"
           ConSql.MoveNext
        Loop
     End If
     ConSql.Close: Set ConSql = Nothing
  End If
  
  ' *** Copiar Datos Adjuntos Salad Bar *** '
  vdia = 999999: ind_datadj = 0
  If Check1(1).Value = 1 Then
     Set ConSql = vg_db.Execute("select Sdx_EncDatosAdjuntos.dia_datadj, Sdx_DetDatosAdjuntos.num_linea, " & _
                  "Sdx_DetDatosAdjuntos.tipo_datadj, Sdx_DetDatosAdjuntos.cod_item, " & _
                  "Sdx_DetDatosAdjuntos.descripcion " & _
                  "From Sdx_DetDatosAdjuntos, Sdx_EncDatosAdjuntos " & _
                  "Where Sdx_EncDatosAdjuntos.ind_datadj = Sdx_DetDatosAdjuntos.ind_datadj " & _
                  "and ((Sdx_EncDatosAdjuntos.cod_casino='" & "00000" & LimpiaDato(Trim(fpText(0).Text)) & "') " & _
                  "and  (Sdx_EncDatosAdjuntos.cod_regimen=" & Val(fpLongInteger1(1).Value) & ") " & _
                  "and  (Sdx_EncDatosAdjuntos.cod_servicio=" & Val(fpLongInteger1(2).Value) & ") " & _
                  "and  (Sdx_EncDatosAdjuntos.tipo_datadj='1') " & _
                  "and  (Sdx_EncDatosAdjuntos.ind_borrado=0) " & _
                  "and  (Sdx_DetDatosAdjuntos.ind_borrado=0)) " & _
                  "order by Sdx_EncDatosAdjuntos.dia_datadj", , adCmdText)
     If Not ConSql.EOF Then
        Do While Not ConSql.EOF
           If ConSql!dia_datadj <> vdia Then
                 
              ind_datadj = 0
          
              Set Consql1 = vg_db.Execute("select Sdx_EncDatosAdjuntos.ind_datadj " & _
                            "From  Sdx_EncDatosAdjuntos " & _
                            "where Sdx_EncDatosAdjuntos.cod_casino='" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "' " & _
                            "and   Sdx_EncDatosAdjuntos.cod_regimen=" & Val(fpLongInteger1(4).Value) & " " & _
                            "and   Sdx_EncDatosAdjuntos.cod_servicio=" & Val(fpLongInteger1(5).Value) & " " & _
                            "and   Sdx_EncDatosAdjuntos.dia_datadj=" & ConSql!dia_datadj & " " & _
                            "and   Sdx_EncDatosAdjuntos.tipo_datadj='1'", , adCmdText)
              If Not Consql1.EOF Then
                 ind_datadj = Consql1!ind_datadj
                 Consql1.Close: Set Consql1 = Nothing
                 If ind_datadj > 0 Then
                    vg_db.Execute "Delete Sdx_EncDatosAdjuntos " & _
                                  "from Sdx_EncDatosAdjuntos " & _
                                  "where cod_casino='" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "' " & _
                                  "and   cod_regimen=" & Val(fpLongInteger1(4).Value) & " " & _
                                  "and   cod_servicio=" & Val(fpLongInteger1(5).Value) & " " & _
                                  "and   ind_datadj=" & ind_datadj & " " & _
                                  "and   dia_datadj=" & ConSql!dia_datadj & "" & _
                                  "and   tipo_datadj='1'"
         
                    vg_db.Execute "Delete Sdx_DetDatosAdjuntos from Sdx_DetDatosAdjuntos " & _
                                  "where ind_datadj=" & ind_datadj & ""
                 End If
              Else
                 Consql1.Close: Set Consql1 = Nothing
                 Set Consql1 = vg_db.Execute("select * from Sdx_Parametro holdlock where Parametro_Num=43", , adCmdText)
                 If Not Consql1.EOF Then
                    vg_db.Execute "Update Sdx_Parametro Set Parametro_Val = Parametro_Val + 1 " & _
                                  "Where Parametro_Num=43"
                 Else
                    vg_db.Execute "insert into Sdx_Parametro (Parametro_Num, Parametro_Desc, Parametro_Val) " & _
                                  "values (43, 'Parametro Datos Adjuntos', 1)"
                 End If
                 Consql1.Close: Set Consql1 = Nothing
   
                 Set Consql1 = vg_db.Execute("select Parametro_Val From Sdx_Parametro " & _
                               "Where Parametro_Num=43", , adCmdText)
                 If Not Consql1.EOF Then
                    ind_datadj = Consql1!Parametro_Val
                 End If
                 Consql1.Close: Set Consql1 = Nothing
              End If
              vg_db.Execute "insert into Sdx_EncDatosAdjuntos (cod_casino, cod_regimen, " & _
                            "cod_servicio, ind_datadj, dia_datadj, tipo_datadj, fecha_datadj, " & _
                            "op_datadj, ind_borrado) values " & _
                            "('" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "', " & _
                            "" & Val(fpLongInteger1(4).Value) & ", " & _
                            "" & Val(fpLongInteger1(5).Value) & ", " & _
                            "" & ind_datadj & ", " & ConSql!dia_datadj & ", '" & 1 & "', " & _
                            "" & Format(Date, "yyyymm") & ", '0', 0)"
           
              vdia = ConSql!dia_datadj
           End If
           vg_db.Execute "insert into Sdx_DetDatosAdjuntos (ind_datadj, num_linea, num_dia, " & _
                         "tipo_datadj, cod_item, descripcion, ind_borrado) values " & _
                         "(" & ind_datadj & ", " & ConSql!num_linea & ", " & _
                         "" & ConSql!dia_datadj & ", " & ConSql!tipo_datadj & ", " & _
                         "" & ConSql!cod_item & ", '" & ConSql!descripcion & "', 0)"
           ConSql.MoveNext
        Loop
     End If
     ConSql.Close: Set ConSql = Nothing
  End If
  
  ' *** Copiar Datos Adjuntos Postres *** '
  vdia = 999999
  If Check1(2).Value = 1 Then
     Set ConSql = vg_db.Execute("select Sdx_EncDatosAdjuntos.dia_datadj, Sdx_DetDatosAdjuntos.num_linea, " & _
                  "Sdx_DetDatosAdjuntos.tipo_datadj, Sdx_DetDatosAdjuntos.cod_item, " & _
                  "Sdx_DetDatosAdjuntos.descripcion " & _
                  "From Sdx_DetDatosAdjuntos, Sdx_EncDatosAdjuntos " & _
                  "Where Sdx_EncDatosAdjuntos.ind_datadj = Sdx_DetDatosAdjuntos.ind_datadj " & _
                  "and ((Sdx_EncDatosAdjuntos.cod_casino='" & "00000" & LimpiaDato(Trim(fpText(0).Text)) & "') " & _
                  "and  (Sdx_EncDatosAdjuntos.cod_regimen=" & Val(fpLongInteger1(1).Value) & ") " & _
                  "and  (Sdx_EncDatosAdjuntos.cod_servicio=" & Val(fpLongInteger1(2).Value) & ") " & _
                  "and  (Sdx_EncDatosAdjuntos.tipo_datadj='2') " & _
                  "and  (Sdx_EncDatosAdjuntos.ind_borrado=0) " & _
                  "and  (Sdx_DetDatosAdjuntos.ind_borrado=0)) " & _
                  "order by Sdx_EncDatosAdjuntos.dia_datadj", , adCmdText)
     If Not ConSql.EOF Then
        Do While Not ConSql.EOF
           If ConSql!dia_datadj <> vdia Then
                 
              ind_datadj = 0
          
              Set Consql1 = vg_db.Execute("select Sdx_EncDatosAdjuntos.ind_datadj " & _
                            "From  Sdx_EncDatosAdjuntos " & _
                            "where Sdx_EncDatosAdjuntos.cod_casino='" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "' " & _
                            "and   Sdx_EncDatosAdjuntos.cod_regimen=" & Val(fpLongInteger1(4).Value) & " " & _
                            "and   Sdx_EncDatosAdjuntos.cod_servicio=" & Val(fpLongInteger1(5).Value) & " " & _
                            "and   Sdx_EncDatosAdjuntos.dia_datadj=" & ConSql!dia_datadj & " " & _
                            "and   Sdx_EncDatosAdjuntos.tipo_datadj='2'", , adCmdText)
              If Not Consql1.EOF Then
                 ind_datadj = Consql1!ind_datadj
                 Consql1.Close: Set Consql1 = Nothing
                 If ind_datadj > 0 Then
                    vg_db.Execute "Delete Sdx_EncDatosAdjuntos " & _
                                  "from Sdx_EncDatosAdjuntos " & _
                                  "where cod_casino='" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "' " & _
                                  "and   cod_regimen=" & Val(fpLongInteger1(4).Value) & " " & _
                                  "and   cod_servicio=" & Val(fpLongInteger1(5).Value) & " " & _
                                  "and   ind_datadj=" & ind_datadj & " " & _
                                  "and   dia_datadj=" & ConSql!dia_datadj & "" & _
                                  "and   tipo_datadj='2'"
         
                    vg_db.Execute "Delete Sdx_DetDatosAdjuntos from Sdx_DetDatosAdjuntos " & _
                                  "where ind_datadj=" & ind_datadj & ""
                 End If
              Else
                 Consql1.Close: Set Consql1 = Nothing
                 Set Consql1 = vg_db.Execute("select * from Sdx_Parametro holdlock where Parametro_Num=43", , adCmdText)
                 If Not Consql1.EOF Then
                    vg_db.Execute "Update Sdx_Parametro Set Parametro_Val = Parametro_Val + 1 " & _
                                  "Where Parametro_Num=43"
                 Else
                    vg_db.Execute "insert into Sdx_Parametro (Parametro_Num, Parametro_Desc, Parametro_Val) " & _
                                  "values (43, 'Parametro Datos Adjuntos', 1)"
                 End If
                 Consql1.Close: Set Consql1 = Nothing
   
                 Set Consql1 = vg_db.Execute("select Parametro_Val From Sdx_Parametro " & _
                               "Where Parametro_Num=43", , adCmdText)
                 If Not Consql1.EOF Then
                    ind_datadj = Consql1!Parametro_Val
                 End If
                 Consql1.Close: Set Consql1 = Nothing
              End If
              vg_db.Execute "insert into Sdx_EncDatosAdjuntos (cod_casino, cod_regimen, " & _
                            "cod_servicio, ind_datadj, dia_datadj, tipo_datadj, fecha_datadj, " & _
                            "op_datadj, ind_borrado) values " & _
                            "('" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "', " & _
                            "" & Val(fpLongInteger1(4).Value) & ", " & _
                            "" & Val(fpLongInteger1(5).Value) & ", " & _
                            "" & ind_datadj & ", " & ConSql!dia_datadj & ", '2', " & _
                            "" & Format(Date, "yyyymm") & ", '0', 0)"
           
              vdia = ConSql!dia_datadj
           End If
           vg_db.Execute "insert into Sdx_DetDatosAdjuntos (ind_datadj, num_linea, num_dia, " & _
                         "tipo_datadj, cod_item, descripcion, ind_borrado) values " & _
                         "(" & ind_datadj & ", " & ConSql!num_linea & ", " & _
                         "" & ConSql!dia_datadj & ", " & ConSql!tipo_datadj & ", " & _
                         "" & ConSql!cod_item & ", '" & ConSql!descripcion & "', 0)"
           ConSql.MoveNext
        Loop
     End If
     ConSql.Close: Set ConSql = Nothing
  End If

vg_db.CommitTrans
fg_descarga
Picture1.Visible = False: Label1(5).Visible = False: gauge.Visible = False
MsgBox "Copia Finalizada Sin Problema", vbInformation + vbOKOnly, "Copia Casino"

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
' *** Graba Plantilla Casino Origen Hacia Casino Origen *** '
'fg_carga (ss)
'gauge.Value = 0
'Picture1.Visible = True: Label1(5).Visible = True: gauge.Visible = False
'Picture1.Refresh
'Label1(5).Caption = "Copiando Planificación Casino Menú Espere Un Momento ..."
'vg_db.Execute "sod_p_copiaplanificacionminuta '" & "00000" & LimpiaDato(Trim(fpText(0).Text)) & "', " & _
'"" & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", '" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "', " & _
'"" & Val(fpLongInteger1(4).Value) & ", " & Val(fpLongInteger1(5).Value) & ", " & _
'"' " & vg_NUsr & "', 1, '', ' " & "M_Minu05" & " ', ''"

' *** Graba Estructura Fija Casino Origen Hacia Casino Origen *** '

'If Check1(0).Value = 1 Then
'   vg_db.Execute "sod_p_copiaestructurafija '" & "00000" & LimpiaDato(Trim(fpText(0).Text)) & "', " & _
'   "" & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", '" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "', " & _
'   "" & Val(fpLongInteger1(4).Value) & ", " & Val(fpLongInteger1(5).Value) & ", " & _
'   "' " & vg_NUsr & "', 1, '', ' " & "M_Minu05" & " ', ''"
'End If

' *** Graba Datos Adjuntos Salad Bar Casino Origen Hacia Casino Origen *** '

'If Check1(1).Value = 1 Then
'   vg_db.Execute "sod_p_copiadatosadjuntos '" & "00000" & LimpiaDato(Trim(fpText(0).Text)) & "', " & _
'   "" & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", '" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "', " & _
'   "" & Val(fpLongInteger1(4).Value) & ", " & Val(fpLongInteger1(5).Value) & ", " & _
'   "'" & vg_NUsr & "', 1, '', ' " & "M_Minu05" & " ', '', '1'"
'End If

'' *** Graba Datos Adjuntos Postres Casino Origen Hacia Casino Origen *** '

'If Check1(2).Value = 1 Then
'   vg_db.Execute "sod_p_copiadatosadjuntos '" & "00000" & LimpiaDato(Trim(fpText(0).Text)) & "', " & _
'   "" & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", '" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "', " & _
'   "" & Val(fpLongInteger1(4).Value) & ", " & Val(fpLongInteger1(5).Value) & ", " & _
'   "'" & vg_NUsr & "', 1, '', ' " & "M_Minu05" & " ', '', '2'"
'End If

'fg_descarga
'Picture1.Visible = False: Label1(5).Visible = False: gauge.Visible = False
'MsgBox "Copia Finalizada Sin Problema", vbInformation + vbOKOnly, "Copia Casino"
End Sub
Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
Select Case Index
  Case 1
    vg_codigo = Val(fpLongInteger1(1).Value)
    vg_auxcodpventa = Val(fpLongInteger1(1).Value)
    Set ConSql = vg_db.Execute("select Sls_Locn_No, Sls_Locn_Name " & _
                 "From PB00367 " & _
                 "where Sls_Locn_No=" & vg_codigo & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_puntoventa 7, " & vg_codigo & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: fpayuda(2).Text = "": vg_codigo = 0: vg_auxcodpventa = 0: fpLongInteger1(1).SetFocus: Exit Sub
    fpayuda(2).Text = ConSql!Sls_Locn_Name
    ConSql.Close: Set ConSql = Nothing
  Case 2
    vg_codigo = Val(fpLongInteger1(2).Value)
    vg_auxcodservicio = Val(fpLongInteger1(2).Value)
    Set ConSql = vg_db.Execute("select Serv_No, Serv_Name, Serv_Orig_Type " & _
                 "From PB00331 " & _
                 "where Serv_No=" & vg_codigo & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_servicio 7, " & vg_codigo & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: fpayuda(3).Text = "": vg_codigo = 0: vg_auxcodservicio = 0: fpLongInteger1(2).SetFocus: Exit Sub
    fpayuda(3).Text = ConSql!Serv_Name
    ConSql.Close: Set ConSql = Nothing
  Case 4
    vg_codigo = Val(fpLongInteger1(4).Value)
    vg_auxcodpventa = Val(fpLongInteger1(4).Value)
    Set ConSql = vg_db.Execute("select Sls_Locn_No, Sls_Locn_Name " & _
                 "From PB00367 " & _
                 "where Sls_Locn_No=" & vg_codigo & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_puntoventa 7, " & vg_codigo & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: fpayuda(6).Text = "": vg_codigo = 0: vg_auxcodpventa = 0: fpLongInteger1(4).SetFocus: Exit Sub
    fpayuda(6).Text = ConSql!Sls_Locn_Name
    ConSql.Close: Set ConSql = Nothing
  Case 5
    vg_codigo = Val(fpLongInteger1(5).Value)
    vg_auxcodservicio = Val(fpLongInteger1(5).Value)
    Set ConSql = vg_db.Execute("select Serv_No, Serv_Name, Serv_Orig_Type " & _
                 "From PB00331 " & _
                 "where Serv_No=" & vg_codigo & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_servicio 7, " & vg_codigo & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: fpayuda(7).Text = "": vg_codigo = 0: vg_auxcodservicio = 0:  fpLongInteger1(5).SetFocus: Exit Sub
    fpayuda(7).Text = ConSql!Serv_Name
    ConSql.Close: Set ConSql = Nothing
End Select
End Sub
Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    If Index = 1 Then image1_Click 1
    If Index = 2 Then image1_Click 2
    If Index = 4 Then image1_Click 4
    If Index = 5 Then image1_Click 5
End Select
End Sub
Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
Select Case Index
  Case 0
    vg_auxcodcasino = ""
    vg_auxcodcasino = fpText(0).Text
    Set ConSql = vg_db.Execute("select * " & _
                 "From Sdx_Casino " & _
                 "where Codigo_Casino='" & "00000" & vg_auxcodcasino & "'", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_casino 8, '" & "00000" & vg_auxcodcasino & "', ''", , adCmdStoredProc)
    If ConSql.EOF Then fpText(0).Text = "": fpayuda(1).Text = "": vg_auxcodcasino = "": ConSql.Close: Set ConSql = Nothing: Exit Sub
    fpayuda(1).Text = ConSql!Nombre_Casino
    ConSql.Close: Set ConSql = Nothing
  Case 1
    vg_auxcodcasino = ""
    vg_auxcodcasino = fpText(1).Text
    Set ConSql = vg_db.Execute("select * " & _
                 "From Sdx_Casino " & _
                 "where Codigo_Casino='" & "00000" & vg_auxcodcasino & "'", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_casino 8, '" & "00000" & vg_auxcodcasino & "', ''", , adCmdStoredProc)
    If ConSql.EOF Then fpText(1).Text = "": fpayuda(5).Text = "": vg_auxcodcasino = "": ConSql.Close: Set ConSql = Nothing: Exit Sub
    fpayuda(5).Text = ConSql!Nombre_Casino
    ConSql.Close: Set ConSql = Nothing
End Select
End Sub
Private Sub fpText_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    If Index = 0 Then image1_Click 0
    If Index = 1 Then image1_Click 3
End Select
End Sub
Private Sub image1_Click(Index As Integer)
Select Case Index
  Case 0
    vg_auxcodcasino = ""
    vg_left = fpayuda(1).Left + 2250
    B_Casino.Show 1
    M_Minu05.Refresh
    If vg_auxcodcasino = "" Then Exit Sub
    fpText(0).Text = vg_auxcodcasino
    Set ConSql = vg_db.Execute("select * " & _
                 "From Sdx_Casino " & _
                 "where Codigo_Casino='" & "00000" & vg_auxcodcasino & "'", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_casino 8, '" & "00000" & vg_auxcodcasino & "', ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(1).Text = ConSql!Nombre_Casino
       fpLongInteger1(1).SetFocus
    Else
       fpayuda(1).Text = ""
       vg_auxcodcasino = ""
       MsgBox "Casino No Existe", vbExclamation + vbOKOnly, "Copiar Minutas"
    End If
    ConSql.Close: Set ConSql = Nothing
  Case 1
    vg_opayuda = 1
    vg_codigo = 0
    vg_left = fpayuda(2).Left + 2250
    B_PtoVta.Show 1
    M_Minu05.Refresh
    If vg_codigo = 0 Then Exit Sub
    fpLongInteger1(1).Value = vg_codigo
    vg_auxcodpventa = vg_codigo
    Set ConSql = vg_db.Execute("select Sls_Locn_No, Sls_Locn_Name " & _
                 "From PB00367 " & _
                 "where Sls_Locn_No=" & vg_codigo & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_puntoventa 7, " & vg_codigo & ", ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(2).Text = ConSql!Sls_Locn_Name
       fpLongInteger1(2).SetFocus
    Else
       fpayuda(2).Text = ""
       vg_codigo = 0
       vg_auxcodpventa = 0
       MsgBox "Punto Venta No Existe", vbExclamation + vbOKOnly, "Copiar Minutas"
    End If
    ConSql.Close: Set ConSql = Nothing
  Case 2
    vg_codigo = 0
    vg_left = fpayuda(3).Left + 2250
    B_Servic.Show 1
    M_Minu05.Refresh
    If vg_codigo = 0 Then Exit Sub
    fpLongInteger1(2).Value = vg_codigo
    vg_auxcodservicio = vg_codigo
    Set ConSql = vg_db.Execute("select Serv_No, Serv_Name, Serv_Orig_Type " & _
                 "From PB00331 " & _
                 "where Serv_No=" & vg_codigo & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_servicio 7, " & vg_codigo & ", ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(3).Text = ConSql!Serv_Name
    Else
       fpayuda(3).Text = ""
       vg_codigo = 0
       vg_auxcodservicio = 0
       MsgBox "Servicio No Existe", vbExclamation + vbOKOnly, "Copiar Minutas"
    End If
    ConSql.Close: Set ConSql = Nothing
  Case 3
    vg_auxcodcasino = ""
    vg_left = fpayuda(5).Left + 2250
    B_Casino.Show 1
    M_Minu05.Refresh
    If vg_auxcodcasino = "" Then Exit Sub
    fpText(1).Text = vg_auxcodcasino
    Set ConSql = vg_db.Execute("select * " & _
                 "From Sdx_Casino " & _
                 "where Codigo_Casino='" & "00000" & vg_auxcodcasino & "'", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_casino 8, '" & "00000" & vg_auxcodcasino & "', ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(5).Text = ConSql!Nombre_Casino
       fpLongInteger1(4).SetFocus
    Else
       fpayuda(5).Text = ""
       vg_auxcodcasino = ""
       MsgBox "Casino No Existe", vbExclamation + vbOKOnly, "Copiar Minutas"
    End If
    ConSql.Close: Set ConSql = Nothing
  Case 4
    vg_opayuda = 1
    vg_codigo = 0
    vg_left = fpayuda(6).Left + 2250
    B_PtoVta.Show 1
    M_Minu05.Refresh
    If vg_codigo = 0 Then Exit Sub
    fpLongInteger1(4).Value = vg_codigo
    vg_auxcodpventa = vg_codigo
    Set ConSql = vg_db.Execute("select Sls_Locn_No, Sls_Locn_Name " & _
                 "From PB00367 " & _
                 "where Sls_Locn_No=" & vg_codigo & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_puntoventa 7, " & vg_codigo & ", ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(6).Text = ConSql!Sls_Locn_Name
       fpLongInteger1(5).SetFocus
    Else
       fpayuda(6).Text = ""
       vg_codigo = 0
       vg_auxcodpventa = 0
       MsgBox "Punto Venta No Existe", vbExclamation + vbOKOnly, "Copiar Minutas"
    End If
    ConSql.Close: Set ConSql = Nothing
  Case 5
    vg_codigo = 0
    vg_left = fpayuda(7).Left + 2250
    B_Servic.Show 1
    M_Minu05.Refresh
    If vg_codigo = 0 Then Exit Sub
    fpLongInteger1(5).Value = vg_codigo
    vg_auxcodservicio = vg_codigo
    Set ConSql = vg_db.Execute("select Serv_No, Serv_Name, Serv_Orig_Type " & _
                 "From PB00331 " & _
                 "where Serv_No=" & vg_codigo & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_servicio 7, " & vg_codigo & ", ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(7).Text = ConSql!Serv_Name
    Else
       fpayuda(7).Text = ""
       vg_codigo = 0
       vg_auxcodservicio = 0
       MsgBox "Servicio No Existe", vbExclamation + vbOKOnly, "Copiar Minutas"
    End If
    ConSql.Close: Set ConSql = Nothing
End Select
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 2
'   *** Validar Datos Origen *** '
    Set ConSql = vg_db.Execute("select * " & _
                 "From Sdx_Casino " & _
                 "where Codigo_Casino='" & "00000" & LimpiaDato(Trim(fpText(0).Text)) & "'", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_casino 8, '" & "00000" & LimpiaDato(Trim(fpText(0).Text)) & "', ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: MsgBox "No Existe Casino Origen", vbExclamation + vbOKOnly, "Copiar Minutas": Exit Sub
    Set ConSql = vg_db.Execute("select Sls_Locn_No, Sls_Locn_Name " & _
                 "From PB00367 " & _
                 "where Sls_Locn_No=" & Val(fpLongInteger1(1).Value) & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_puntoventa 7, " & Val(fpLongInteger1(1).Value) & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: MsgBox "No Existe P. Venta Origen", vbExclamation + vbOKOnly, "Copiar Minutas": Exit Sub
    Set ConSql = vg_db.Execute("select Serv_No, Serv_Name, Serv_Orig_Type " & _
                 "From PB00331 " & _
                 "where Serv_No=" & Val(fpLongInteger1(2).Value) & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_servicio 7, " & Val(fpLongInteger1(2).Value) & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: MsgBox "No Existe Servicio Origen", vbExclamation + vbOKOnly, "Copiar Minutas": Exit Sub
'   *** Validar Datos Destino *** '
    Set ConSql = vg_db.Execute("select * " & _
                 "From Sdx_Casino " & _
                 "where Codigo_Casino='" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "'", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_casino 8, '" & "00000" & LimpiaDato(Trim(fpText(1).Text)) & "', ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: MsgBox "No Existe Casino Destino", vbExclamation + vbOKOnly, "Copiar Minutas": Exit Sub
    Set ConSql = vg_db.Execute("select Sls_Locn_No, Sls_Locn_Name " & _
                 "From PB00367 " & _
                 "where Sls_Locn_No=" & Val(fpLongInteger1(4).Value) & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_puntoventa 7, " & Val(fpLongInteger1(4).Value) & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: MsgBox "No Existe P. Venta Destino", vbExclamation + vbOKOnly, "Copiar Minutas": Exit Sub
    Set ConSql = vg_db.Execute("select Serv_No, Serv_Name, Serv_Orig_Type " & _
                 "From PB00331 " & _
                 "where Serv_No=" & Val(fpLongInteger1(5).Value) & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_servicio 7, " & Val(fpLongInteger1(5).Value) & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: MsgBox "No Existe Servicio Destino", vbExclamation + vbOKOnly, "Copiar Minutas": Exit Sub
    CopiarPlantillaCasino
  Case 4
    Me.Hide
    Unload Me
End Select
End Sub
Sub Mover_Botones()

   Toolbar1.ImageList = partida.IL1
   Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
   Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar "
   Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
   Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"

End Sub

