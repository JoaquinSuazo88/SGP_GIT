VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form M_Minu01 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planificación Minutas"
   ClientHeight    =   2340
   ClientLeft      =   1860
   ClientTop       =   2760
   ClientWidth     =   7365
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2340
   ScaleWidth      =   7365
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6735
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   2
         Left            =   2685
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   990
         Width           =   3885
         _Version        =   196608
         _ExtentX        =   6853
         _ExtentY        =   556
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         ControlType     =   2
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
         Left            =   2685
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   675
         Width           =   3885
         _Version        =   196608
         _ExtentX        =   6853
         _ExtentY        =   556
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         ControlType     =   2
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
         Left            =   1335
         TabIndex        =   1
         Top             =   990
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
         Index           =   2
         Left            =   1335
         TabIndex        =   2
         Top             =   1305
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
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   3
         Left            =   2685
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1305
         Width           =   3885
         _Version        =   196608
         _ExtentX        =   6853
         _ExtentY        =   556
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         ControlType     =   2
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
         Left            =   1335
         TabIndex        =   0
         Top             =   675
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
         Index           =   3
         Left            =   165
         TabIndex        =   10
         Top             =   1410
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
         Index           =   2
         Left            =   165
         TabIndex        =   9
         Top             =   1095
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
         Index           =   0
         Left            =   165
         TabIndex        =   8
         Top             =   780
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2200
         Picture         =   "M_Minu01.frx":0000
         Top             =   580
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2200
         Picture         =   "M_Minu01.frx":030A
         Top             =   900
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2200
         Picture         =   "M_Minu01.frx":0614
         Top             =   1220
         Width           =   480
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   2340
      Left            =   6735
      TabIndex        =   3
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   4128
      ButtonWidth     =   1138
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Histórico (Planificación Minutas)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Minu01.frx":091E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Minu01.frx":0C38
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Minu01.frx":0F52
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "M_Minu01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConSql As ADODB.Recordset, Consql1 As ADODB.Recordset
Private Sub Form_Activate()
fg_descarga
End Sub
Private Sub Form_Load()

On Error GoTo Man_Error

fg_carga ""
Me.Height = 2715
Me.Width = 7455
fg_centra Me
vg_codcasino = "": vg_auxcodcasino = ""
vg_codsegmento = 0: vg_auxcodsegmento = 0
vg_codregimen = 0: vg_auxcodregimen = 0
vg_codpventa = 0: vg_auxcodpventa = 0
vg_codservicio = 0: vg_auxcodservicio = 0
vg_auxcategoria1 = 0: vg_auxcategoria2 = 0
vg_auxcategoria3 = 0: vg_auxcategoria4 = 0
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cargando Tabla Planificación Minutas"
End Sub
Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub
Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    If Index = 1 Then image1_Click 1
    If Index = 2 Then image1_Click 2
End Select
End Sub
Private Sub fpLongInteger1_LostFocus(Index As Integer)
Select Case Index
  Case 1
    If Val(fpLongInteger1(1).Value) < 1 Then fpayuda(2).Text = "": Exit Sub
    vg_codigo = Val(fpLongInteger1(1).Value)
    vg_auxcodpventa = Val(fpLongInteger1(1).Value)
    Set ConSql = vg_db.Execute("select Sls_Locn_No, Sls_Locn_Name " & _
                 "From PB00367 " & _
                 "where Sls_Locn_No=" & vg_codigo & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_puntoventa 7, " & vg_codigo & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: fpayuda(2).Text = "": vg_codigo = 0: vg_auxcodpventa = 0: vg_codpventa = 0: Exit Sub
    fpayuda(2).Text = Trim(ConSql!Sls_Locn_Name)
    ConSql.Close: Set ConSql = Nothing
  Case 2
    If Val(fpLongInteger1(2).Value) < 1 Then fpayuda(3).Text = "": Exit Sub
    vg_codigo = Val(fpLongInteger1(2).Value)
    vg_auxcodservicio = Val(fpLongInteger1(2).Value)
    Set ConSql = vg_db.Execute("select Serv_No, Serv_Name, Serv_Orig_Type " & _
                 "From PB00331 " & _
                 "where Serv_No=" & vg_codigo & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_servicio 7, " & vg_codigo & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: fpayuda(3).Text = "": vg_codigo = 0: vg_auxcodservicio = 0: vg_codservicio = 0: Exit Sub
    fpayuda(3).Text = Trim(ConSql!Serv_Name)
    ConSql.Close: Set ConSql = Nothing
End Select
End Sub
Private Sub fpText_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub
Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    image1_Click 0
End Select
End Sub
Private Sub fpText_LostFocus()
If fpText.Text = "" Then fpayuda(1).Text = "": Exit Sub
vg_auxcodcasino = ""
vg_auxcodcasino = fpText.Text
Set ConSql = vg_db.Execute("select * " & _
             "From Sdx_Casino " & _
             "where Codigo_Casino='" & "00000" & vg_auxcodcasino & "'", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_casino 8, '" & "00000" & vg_auxcodcasino & "', ''", , adCmdStoredProc)
If ConSql.EOF Then fpayuda(1).Text = "": vg_auxcodcasino = "": vg_codcasino = "": fpLongInteger1(1).Value = "": fpayuda(2).Text = "": fpLongInteger1(2).Value = "": fpayuda(3).Text = "": ConSql.Close: Set ConSql = Nothing: Exit Sub
fpayuda(1).Text = Trim(ConSql!Nombre_Casino)
ConSql.Close: Set ConSql = Nothing
fpLongInteger1(1).Value = "": fpayuda(2).Text = ""
fpLongInteger1(2).Value = "": fpayuda(3).Text = ""
End Sub
Private Sub image1_Click(Index As Integer)
Select Case Index
  Case 0
    vg_auxcodcasino = ""
    vg_left = fpayuda(1).Left + 2300
    B_Casino.Show 1
    M_Minu01.Refresh
    If vg_auxcodcasino = "" Then Exit Sub
    fpText.Text = vg_auxcodcasino
    Set ConSql = vg_db.Execute("select * " & _
                 "From Sdx_Casino " & _
                 "where Codigo_Casino='" & "00000" & vg_auxcodcasino & "'", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_casino 8, '" & "00000" & vg_auxcodcasino & "', ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(1).Text = Trim(ConSql!Nombre_Casino)
       fpLongInteger1(1).SetFocus
    Else
       fpayuda(1).Text = ""
       vg_auxcodcasino = ""
       MsgBox "Casino No Existe", vbExclamation + vbOKOnly, "Planificación Minutas"
    End If
    ConSql.Close: Set ConSql = Nothing
    fpLongInteger1(1).Value = "": fpayuda(2).Text = ""
    fpLongInteger1(2).Value = "": fpayuda(3).Text = ""
  Case 1
    vg_codigo = 0
    vg_opayuda = 1
    vg_left = fpayuda(1).Left + 2300
    B_PtoVta.Show 1
    M_Minu01.Refresh
    If vg_codigo = 0 Then Exit Sub
    fpLongInteger1(1).Value = vg_codigo
    vg_auxcodpventa = vg_codigo
    Set ConSql = vg_db.Execute("select Sls_Locn_No, Sls_Locn_Name " & _
                 "From PB00367 " & _
                 "where Sls_Locn_No=" & vg_codigo & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_puntoventa 7, " & vg_codigo & ", ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(2).Text = Trim(ConSql!Sls_Locn_Name)
       fpLongInteger1(2).SetFocus
    Else
       fpayuda(2).Text = ""
       vg_codigo = 0
       vg_auxcodpventa = 0
       MsgBox "Punto Venta No Existe", vbExclamation + vbOKOnly, "Planificación Minutas"
    End If
    ConSql.Close: Set ConSql = Nothing
  Case 2
    vg_codigo = 0
    vg_left = fpayuda(3).Left + 2300
    B_Servic.Show 1
    M_Minu01.Refresh
    If vg_codigo = 0 Then Exit Sub
    fpLongInteger1(2).Value = vg_codigo
    vg_auxcodservicio = vg_codigo
    Set ConSql = vg_db.Execute("select Serv_No, Serv_Name, Serv_Orig_Type " & _
                 "From PB00331 " & _
                 "where Serv_No=" & vg_codigo & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_servicio 7, " & vg_codigo & ", ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(3).Text = Trim(ConSql!Serv_Name)
    Else
       fpayuda(3).Text = ""
       vg_codigo = 0
       vg_auxcodservicio = 0
       MsgBox "Servicio No Existe", vbExclamation + vbOKOnly, "Planificación Minutas"
    End If
    ConSql.Close: Set ConSql = Nothing
End Select
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 2
    Set ConSql = vg_db.Execute("select * " & _
                 "From Sdx_Casino " & _
                 "where Codigo_Casino='" & "00000" & LimpiaDato(Trim(fpText.Text)) & "'", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_casino 8, '" & "00000" & LimpiaDato(Trim(fpText.Text)) & "', ''", , adCmdStoredProc)
    If ConSql.EOF Then fpText.Text = "": fpayuda(1).Text = "": fpLongInteger1(1).Value = "": fpayuda(2).Text = "": fpLongInteger1(2).Value = "": fpayuda(3).Text = "": ConSql.Close: Set ConSql = Nothing: MsgBox "No Existe Casino", vbExclamation + vbOKOnly, "Planificación Minutas": Exit Sub
    Set ConSql = vg_db.Execute("select Sls_Locn_No, Sls_Locn_Name " & _
                 "From PB00367 " & _
                 "where Sls_Locn_No=" & Val(fpLongInteger1(1).Value) & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_puntoventa 7, " & Val(fpLongInteger1(1).Value) & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: MsgBox "No Existe Regimen", vbExclamation + vbOKOnly, "Planificación Minutas": Exit Sub
    Set ConSql = vg_db.Execute("select Serv_No, Serv_Name, Serv_Orig_Type " & _
                 "From PB00331 " & _
                 "where Serv_No=" & Val(fpLongInteger1(2).Value) & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_servicio 7, " & Val(fpLongInteger1(2).Value) & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: MsgBox "No Existe Servicio", vbExclamation + vbOKOnly, "Planificación Minutas": Exit Sub
    vg_codcasino = "00000" & LimpiaDato(Trim(fpText.Text))
    vg_codpventa = Val(fpLongInteger1(1).Value)
    vg_codservicio = Val(fpLongInteger1(2).Value)
    Unload M_Minu03
    Unload M_Minu02
'    M_Minu02.Label4.Caption = "* " & Combo1(3).Text & " : " & Combo1(4).Text & " * "
    M_Minu02.Show 0, partida
  Case 4 ' And Toolbar1.Buttons(4).Image = 3
    ' *** Historico Planificación *** '
    
    Set ConSql = vg_db.Execute("select * " & _
                 "From Sdx_Casino " & _
                 "where Codigo_Casino='" & "00000" & LimpiaDato(Trim(fpText.Text)) & "'", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_casino 8, '" & "00000" & LimpiaDato(Trim(fpText.Text)) & "', ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: MsgBox "No Existe Casino", vbExclamation + vbOKOnly, "Planificación Minutas": Exit Sub
    vg_codcasino = "00000" & LimpiaDato(Trim(fpText.Text))
    vg_bcodcasino = "00000" & LimpiaDato(Trim(fpText.Text))
    vg_bopcion = 0
    B_HistPm.Show 1
    If vg_bopcion = 0 Then Exit Sub
    fpLongInteger1(1).Value = vg_bcodpventa
    
    Set ConSql = vg_db.Execute("select Sls_Locn_No, Sls_Locn_Name " & _
                 "From PB00367 " & _
                 "where Sls_Locn_No=" & Val(fpLongInteger1(1).Value) & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_puntoventa 7, " & Val(fpLongInteger1(1).Value) & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: MsgBox "No Existe P. Venta", vbExclamation + vbOKOnly, "Planificación Minutas": Exit Sub
    fpayuda(2).Text = Trim(ConSql!Sls_Locn_Name)
    ConSql.Close: Set ConSql = Nothing
    fpLongInteger1(2).Value = vg_bcodservicio
    
    Set ConSql = vg_db.Execute("select Serv_No, Serv_Name, Serv_Orig_Type " & _
                 "From PB00331 " & _
                 "where Serv_No=" & Val(fpLongInteger1(2).Value) & " " & _
                 "and   Dlt_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_servicio 7, " & Val(fpLongInteger1(2).Value) & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: MsgBox "No Existe Servicio", vbExclamation + vbOKOnly, "Planificación Minutas": Exit Sub
    fpayuda(3).Text = Trim(ConSql!Serv_Name)
    ConSql.Close: Set ConSql = Nothing
    M_Minu01.Refresh
  Case 6
    Unload M_Minu03
    Me.Hide
    Unload Me
End Select
End Sub
