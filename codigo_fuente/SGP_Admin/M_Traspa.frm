VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_Traspa 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspasos"
   ClientHeight    =   7665
   ClientLeft      =   1080
   ClientTop       =   2415
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7665
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   1050
      TabIndex        =   9
      Top             =   375
      Width           =   8580
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   2085
         TabIndex        =   0
         Top             =   495
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
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
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483643
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
         MinValue        =   "-2147483648"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000009&
         Caption         =   "Recibido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   2400
         TabIndex        =   3
         Top             =   1590
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000009&
         Caption         =   "Entregado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   3735
         TabIndex        =   4
         Top             =   1590
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1155
         Width           =   3195
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   0
         Left            =   2085
         TabIndex        =   8
         Top             =   165
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
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
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
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
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   2085
         TabIndex        =   1
         Top             =   825
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         DateCalcMethod  =   0
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   2085
         TabIndex        =   5
         Top             =   1890
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
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
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
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
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   6915
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   495
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
         _ExtentY        =   556
         Enabled         =   -1  'True
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
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
         ThreeDTextHighlightColor=   -2147483633
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
         ControlType     =   2
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Folio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   6195
         TabIndex        =   34
         Top             =   525
         Width           =   495
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3915
         TabIndex        =   28
         Top             =   1890
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3465
         Picture         =   "M_Traspa.frx":0000
         Top             =   1785
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Casino 2"
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
         TabIndex        =   27
         Top             =   1905
         Width           =   750
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2085
         TabIndex        =   26
         Top             =   1545
         Width           =   3180
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2130
         TabIndex        =   25
         Top             =   1605
         Width           =   3180
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   2145
         TabIndex        =   17
         Top             =   1215
         Width           =   3180
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   3615
         TabIndex        =   19
         Top             =   555
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bodega"
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
         Left            =   360
         TabIndex        =   16
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Index           =   6
         Left            =   360
         TabIndex        =   15
         Top             =   195
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3465
         Picture         =   "M_Traspa.frx":030A
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
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
         Left            =   360
         TabIndex        =   14
         Top             =   525
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Emisión"
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
         Left            =   360
         TabIndex        =   13
         Top             =   855
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Traspaso"
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
         Index           =   10
         Left            =   360
         TabIndex        =   12
         Top             =   1545
         Width           =   1230
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3915
         TabIndex        =   10
         Top             =   165
         Width           =   4335
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   3945
         TabIndex        =   11
         Top             =   195
         Width           =   4335
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3945
         TabIndex        =   29
         Top             =   1920
         Width           =   4335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Frame Frame3 
      Height          =   4170
      Left            =   225
      TabIndex        =   24
      Top             =   2655
      Width           =   10365
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3405
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   10140
         _Version        =   393216
         _ExtentX        =   17886
         _ExtentY        =   6006
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
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
         FormulaSync     =   0   'False
         MaxCols         =   9
         MaxRows         =   20
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_Traspa.frx":0614
         ClipboardOptions=   0
      End
      Begin VB.Frame Frame4 
         Height          =   450
         Left            =   8340
         TabIndex        =   30
         Top             =   3615
         Width           =   1695
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   45
            TabIndex        =   31
            Top             =   135
            Width           =   1590
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Left            =   7845
         TabIndex        =   32
         Top             =   3780
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   2175
      TabIndex        =   20
      Top             =   6825
      Width           =   6585
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   330
         TabIndex        =   7
         Top             =   240
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   688
         ButtonWidth     =   2963
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar Producto"
               Description     =   "Agregar Productos"
               Object.ToolTipText     =   "Agregar Producto"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar Producto "
               Description     =   "Eliminar Producto "
               Object.ToolTipText     =   "Eliminar Producto "
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1635
         Top             =   135
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Traspa.frx":0D9B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Traspa.frx":10B5
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad sobrepasa Stock actual"
         Height          =   450
         Index           =   1
         Left            =   4560
         TabIndex        =   21
         Top             =   225
         Width           =   1440
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008484FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   4170
         Top             =   345
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   450
         Top             =   315
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label5 
         Caption         =   "Productos igresados por el Usuario"
         Height          =   450
         Index           =   0
         Left            =   840
         TabIndex        =   23
         Top             =   195
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   735
         Top             =   300
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label5 
         Caption         =   "Productos traidos de la Minuta Real"
         Height          =   450
         Index           =   2
         Left            =   1095
         TabIndex        =   22
         Top             =   180
         Visible         =   0   'False
         Width           =   1425
      End
   End
End
Attribute VB_Name = "M_Traspa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim Msgtitulo As String, Est As Boolean
'Dim btnX As Button

Private Sub Combo1_Click(Index As Integer)
Dim i As Long
If Est Then Exit Sub
Select Case Index
Case 1
    If vaSpread1.MaxRows = 0 Then Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & Trim(LimpiaDato(vaSpread1.Text)) & "'", vg_db, adOpenStatic
        vaSpread1.Col = 9
        If Not RS1.EOF Then vaSpread1.Text = RS1!bod_canmer Else vaSpread1.Text = 0
        RS1.Close: Set RS1 = Nothing
        'REvisa color
        Dim canrea As Double, canbod As Double
        vaSpread1.Col = 4: canrea = Format(vaSpread1.Text, fg_Pict(9, vg_DCa))
        vaSpread1.Col = 9: canbod = Format(vaSpread1.Text, fg_Pict(9, vg_DCa))
        If canbod - canrea < 0 And Option1(0).Value = True Then
            vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
            vaSpread1.Col = 8: vaSpread1.Text = "S"  'Bloqueado
        Else
            vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(2).FillColor
            vaSpread1.Col = 8: vaSpread1.Text = "N" 'No Bloqueado
        End If
    Next i
End Select
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 8010
Me.Width = 10890
fg_centra Me
Est = False
Me.HelpContextID = vg_OpcM
Msgtitulo = "Salida Producción"
Dim X As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
Gl_Mo_Botones Me, 4
vaSpread1.Row = -1
vaSpread1.Col = 4: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 6: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
vaSpread1.Col = 7: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
RS1.Open "select bod_nombre, bod_codigo from a_bodega order by bod_nombre", vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
        Combo1(1).AddItem RS1!bod_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!bod_codigo), 10) & ")"
        RS1.MoveNext
    Loop
End If
RS1.Close: Set RS1 = Nothing
Limpia
End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then
    Frame3.Left = (Me.Width \ 2) - (Frame3.Width \ 2)
    Frame1.Left = (Me.Width \ 2) - (Frame1.Width \ 2)
    Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)
ElseIf Me.WindowState = 0 Then
    Frame3.Left = 255
    Frame3.Width = 10365
    Frame1.Left = 1050
    Frame2.Left = 2175
    vaSpread1.Width = 10140
    Me.Refresh
End If
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_LostFocus(Index As Integer)
If Trim(fpLongInteger1(0).Text) <> "" Then
    BuscaDoc fpLongInteger1(0).Value
End If
End Sub

Private Sub fpText1_Change(Index As Integer)
fpayuda(Index).Caption = ""
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_LostFocus(Index As Integer)
If fpText1(Index).Text = "" Then Exit Sub
Select Case Index
Case 0
    RS1.Open "select cli_nombre from b_clientes where cli_codigo='" & fpText1(0).Text & "' and cli_tipo=0", vg_db, adOpenStatic
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            fpayuda(Index).Caption = RS1!cli_nombre
            Gl_Ac_Botones Me, 4, 2, ""
            fpText1(0).Enabled = False
            RS1.MoveNext
        Loop
    Else
        RS1.Close: Set RS1 = Nothing
        MsgBox "Casino no existe...", vbExclamation + vbOKOnly, Msgtitulo
        Limpia
        If fpText1(0).Enabled = True Then fpText1(0).SetFocus
        Exit Sub
    End If
    RS1.Close: Set RS1 = Nothing
Case 1
    RS1.Open "select cli_nombre from b_clientes where cli_codigo='" & fpText1(1).Text & "' and cli_tipo=2", vg_db, adOpenStatic
    If Not RS1.EOF Then
        fpayuda(Index).Caption = RS1!cli_nombre
        RS1.Close: Set RS1 = Nothing
        Exit Sub
    Else
        RS1.Close: Set RS1 = Nothing
        MsgBox "Casino traspaso no existe...", vbExclamation + vbOKOnly, Msgtitulo
        fpText1(1) = ""
        Exit Sub
    End If
    RS1.Close: Set RS1 = Nothing
End Select
If Trim(fpText1(0).Text) = Trim(fpText1(1).Text) Then
    MsgBox "No se puede realizar transferencia en el mismo casino...", vbExclamation + vbOKOnly, Msgtitulo
    If Index = 0 Then Limpia Else fpText1(Index).Text = ""
    Exit Sub
End If
End Sub

Private Sub Image1_Click(Index As Integer)
vg_codigo = 0
vg_left = fpayuda(Index).Left + 1920
B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", IIf(Index = 1, "Traspaso", "Casino")
B_TabEst.Show 1
Me.Refresh
If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
fpText1(Index) = Trim(vg_codigo)
fpayuda(Index).Caption = vg_nombre
Select Case Index
Case 0
    If Trim(vg_codigo) <> fpText1(Index) Then Limpia
    fpText1_LostFocus 0
    If Trim(fpText1(Index).Text) = "" Then Exit Sub
    If fpDateTime1(Index).Enabled = True Then fpDateTime1(Index).SetFocus
    Gl_Ac_Botones Me, 4, 2, ""
Case 1
    If fpText1(1).Enabled = True Then fpText1(1).SetFocus
    fpText1_LostFocus 1
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
Dim i As Long, cantidad As Double
Label3(0).Caption = IIf(Index = 1, "Casino Origen", "Casino Destino")
Select Case Index
Case 0
    vaSpread1.Col = 5: vaSpread1.Row = -1
    vaSpread1.ColHidden = True
    vaSpread1.Col = 6: vaSpread1.Row = 0
    vaSpread1.Text = "P.M.P."
    vaSpread1.Row = -1:  vaSpread1.Lock = True
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        RS1.Open "SELECT pro_propon FROM b_productos WHERE pro_codigo='" & Trim(LimpiaDato(vaSpread1.Text)) & "'", vg_db, adOpenStatic
        If Not RS1.EOF Then
            vaSpread1.Col = 4: cantidad = vaSpread1.Text
            vaSpread1.Col = 6: vaSpread1.Text = RS1!pro_propon
            vaSpread1.Col = 7: vaSpread1.Text = Format(Format(cantidad, fg_Pict(9, 0)) * RS1!pro_propon, fg_Pict(9, 2))
        End If
        RS1.Close: Set RS1 = Nothing
    Next
    Combo1_Click 1
    If vaSpread1.MaxRows > 0 And vaSpread1.Enabled = True And Est = False Then
        vaSpread1.SetFocus: vaSpread1.SetActiveCell 4, vaSpread1.ActiveRow
    End If
    vaSpread1.ColWidth(2) = 33.25
Case 1
    vaSpread1.Col = 5: vaSpread1.Row = -1
    vaSpread1.ColHidden = False
    vaSpread1.Col = 6: vaSpread1.Row = 0
    vaSpread1.Text = "Precio"
    vaSpread1.Row = -1: vaSpread1.Lock = False
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 6: vaSpread1.Text = 0
        vaSpread1.Col = 7: vaSpread1.Text = 0
    Next
    'REvisa color
    vaSpread1.Col = -1: vaSpread1.Row = -1: vaSpread1.BackColor = Shape1(2).FillColor
    vaSpread1.Col = 8: vaSpread1.Text = "N"
    If vaSpread1.MaxRows > 0 And vaSpread1.Enabled = True And Est = False Then
        vaSpread1.SetFocus: vaSpread1.SetActiveCell 4, vaSpread1.ActiveRow
    End If
    vaSpread1.ColWidth(2) = 23.5
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim folio As Long, rutcli As String, codcas As String, TipDoc As String, numdoc As Long, CodBod As Long, codser As Long, i As Long, canact As Long
Dim numlin As Long, CodMer As String, canmer As Double, canmin As Double, predoc As Double, ptotal As Double, descri As String, total As Double, diablq As Date
Dim coding As String
On Error GoTo Man_Error
If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
Select Case Button.Index
Case 1 'Nuevo
    Limpia
    On Error Resume Next: fpLongInteger1(0).SetFocus
Case 3 'Graba
    If Trim(fpText1(0).Text) = "" Or Trim(fpLongInteger1(0).Text) = "" Or Trim(fpText1(1).Text) = "" _
    Or Trim(Combo1(1).Text) = "" Or Trim(fpDateTime1(0).Text) = "" Then MsgBox "Debe ingresar dato importante...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    'If Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(fpDateTime1(0).Text), "mm/yyyy") < Format(Now, "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If (Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(fpDateTime1(0).Text), "mm/yyyy") < Format(Now, "mm/yyyy")) Or Format(CDate(fpDateTime1(0).Text), "mm/yyyy") < Format(Month(Now) - 1 & "/" & Year(Now), "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Trim(fpText1(0).Text) = Trim(fpText1(1).Text) Then Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 8
        If Left(vaSpread1.Text, 1) = "S" Then MsgBox "Existe una cantidad que exede el Stock...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Next i
    If Format(Label2.Caption, fg_Pict(9, vg_DPr)) = Format(0, fg_Pict(9, vg_DPr)) Then MsgBox "El total del documento debe ser mayor a 0...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    vg_db.BeginTrans
    rutcli = Trim(LimpiaDato(fpText1(0).Text))
    codcas = Trim(LimpiaDato(fpText1(1).Text))
    TipDoc = "TR"
    numdoc = Trim(fpLongInteger1(0).Text)
    CodBod = Val(fg_codigocbo(Combo1, 1, 10, ""))
    codser = IIf(Option1(1).Value = True, 1, 0)
    total = Format(Label2.Caption, fg_Pict(9, vg_DPr))
    fpLongInteger1(1).Text = MuestraFolio
    folio = Trim(fpLongInteger1(1).Text)
    'Encabezado
    vg_db.Execute "insert into b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_totdoc, tov_numinf) " & _
                  "values ('" & rutcli & "', '" & TipDoc & "', " & numdoc & ", " & CodBod & ", CDate('" & _
                  Format(fpDateTime1(0).Text, "dd/mm/yyyy") & "'), 0, " & codser & ", 0, '', '" & codcas & "', " & total & ", " & folio & ")"
    'Detalle
    numlin = 1
    For i = 1 To vaSpread1.MaxRows
        canmin = 0: canmer = 0
        vaSpread1.Row = i
        vaSpread1.Col = 1: CodMer = Trim(LimpiaDato(vaSpread1.Text))
        vaSpread1.Col = 2: descri = Trim(LimpiaDato(vaSpread1.Text))
        vaSpread1.Col = 4: If Option1(1).Value = True Then canmin = LimpiaDato(vaSpread1.Text) Else canmer = LimpiaDato(vaSpread1.Text)
        vaSpread1.Col = 5: If Option1(1).Value = True Then canmer = LimpiaDato(vaSpread1.Text) Else canmin = 0
        vaSpread1.Col = 6: predoc = LimpiaDato(vaSpread1.Text)
        vaSpread1.Col = 7: ptotal = LimpiaDato(vaSpread1.Text)
        If canmer > 0 Then
            ValidaBod CodBod, Trim(LimpiaDato(CodMer))
            'Actualiza Precio Promedio Ponderado si es Traspaso Recibido
            If Option1(1) = True Then
                Dim PMP As Double, auxCanmer As Double, auxPropon As Double
                RS2.Open "Select pro_facing From b_productos Where pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
'                'PMP Ingrediente
'                If Not RS2.EOF Then
'                    coding = IIf(IsNull(RS2!pro_coding), "", RS2!pro_coding)
'                    auxCanmer = 0: auxPropon = 0
'                    RS1.Open "Select Sum(bod.bod_canmer) As canmer From b_productos pro, b_bodegas bod " & _
'                             "Where bod.bod_codpro=pro.pro_codigo And pro.pro_coding='" & coding & "'", vg_db, adOpenStatic
'                    If Not RS1.EOF Then auxCanmer = IIf(IsNull(RS1!canmer), 0, RS1!canmer)
'                    RS1.Close: Set RS1 = Nothing
'                    RS1.Open "Select Sum((pro.pro_propon/pro.pro_facing)*bod_canmer) as propon From b_productos pro, b_bodegas bod " & _
'                             "Where pro.pro_codigo=bod.bod_codpro And pro.pro_coding='" & coding & "'", vg_db, adOpenStatic
'                    If Not RS1.EOF Then auxPropon = IIf(IsNull(RS1!propon), 0, RS1!propon)
'                    RS1.Close: Set RS1 = Nothing
'                    PMP = Val((auxPropon + ((predoc / RS2!pro_facing) * canmer)) / (auxCanmer + canmer))
'                    vg_db.Execute "Update b_ingrediente Set ing_precos=" & PMP & " Where ing_codigo='" & coding & "'"
'                End If
'                RS2.Close: Set RS2 = Nothing
'                'PMP Producto
'                RS1.Open "Select Sum(bod.bod_canmer) As canmer From b_productos pro, b_bodegas bod " & _
'                         "Where bod.bod_codpro=pro.pro_codigo And pro.pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
'                If Not RS1.EOF Then auxCanmer = IIf(IsNull(RS1!canmer), 0, RS1!canmer)
'                RS1.Close: Set RS1 = Nothing
'                RS1.Open "Select pro_propon As propon From b_productos Where pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
'                If Not RS1.EOF Then auxPropon = IIf(IsNull(RS1!propon), 0, RS1!propon)
'                RS1.Close: Set RS1 = Nothing
'                PMP = Round(((auxPropon * auxCanmer) + (predoc * canmer)) / (auxCanmer + canmer), vg_DPr)
'                vg_db.Execute "Update b_productos Set pro_propon=" & PMP & " Where pro_codigo='" & CodMer & "'"
'                'Actuliza codigo compra y pedido de ultimo producto para ingrediente
'                vg_db.Execute "Update b_ingrediente Set ing_codped='" & CodMer & "', ing_codcom='" & CodMer & "' Where ing_codigo='" & coding & "'"
                'PMP Ingrediente
                If Not RS2.EOF Then
                    auxCanmer = 0: auxPropon = 0
                    RS1.Open "Select Sum(bod.bod_canmer) As canmer From b_productos pro, b_bodegas bod " & _
                             "Where bod.bod_codpro=pro.pro_codigo And pro.pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
                    If Not RS1.EOF Then auxCanmer = IIf(IsNull(RS1!canmer), 0, RS1!canmer)
                    RS1.Close: Set RS1 = Nothing
                    RS1.Open "Select Sum((pro.pro_propon/pro.pro_facing)*bod_canmer) as propon From b_productos pro, b_bodegas bod " & _
                             "Where pro.pro_codigo=bod.bod_codpro And pro.pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
                    If Not RS1.EOF Then auxPropon = IIf(IsNull(RS1!propon), 0, RS1!propon)
                    RS1.Close: Set RS1 = Nothing
                    PMP = Val((auxPropon + ((predoc / RS2!pro_facing) * canmer)) / (auxCanmer + canmer))
                    vg_db.Execute "Update b_ingrediente ing, b_productosing pri Set ing.ing_precos=" & PMP & " " & _
                                  "Where pri.pri_coding=ing.ing_codigo And pri.pri_codpro='" & CodMer & "'"
                    'PMP Producto
                    RS1.Open "Select Sum(bod.bod_canmer) As canmer From b_productos pro, b_bodegas bod " & _
                             "Where bod.bod_codpro=pro.pro_codigo And pro.pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
                    If Not RS1.EOF Then auxCanmer = IIf(IsNull(RS1!canmer), 0, RS1!canmer)
                    RS1.Close: Set RS1 = Nothing
                    RS1.Open "Select pro_propon As propon From b_productos Where pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
                    If Not RS1.EOF Then auxPropon = IIf(IsNull(RS1!propon), 0, RS1!propon)
                    RS1.Close: Set RS1 = Nothing
                    PMP = Val(((auxPropon * auxCanmer) + (predoc * canmer)) / (auxCanmer + canmer))
                    vg_db.Execute "Update b_productos Set pro_propon=" & PMP & " Where pro_codigo='" & CodMer & "'"
                    'Actuliza codigo compra y pedido de ultimo producto para ingrediente
                    vg_db.Execute "Update b_ingrediente ing, b_productosing pri Set ing_codped='" & CodMer & "', ing_codcom='" & CodMer & "' " & _
                                  "Where pri.pri_coding=ing.ing_codigo And pri.pri_codpro='" & CodMer & "'"
                End If
                RS2.Close: Set RS2 = Nothing
            End If
            'Graba Detalle
            vg_db.Execute "insert into b_detventas (dev_rutcli, dev_tipdoc, dev_numdoc, dev_numlin, dev_codmer, dev_canmin, dev_canmer, dev_predoc, dev_ptotal, dev_descri, dev_mueinv, dev_porcen, dev_precos, dev_coding) " & _
                          "values ('" & rutcli & "', '" & TipDoc & "', " & numdoc & ", " & numlin & ", '" & CodMer & "', " & canmin & ", " & canmer & ", " & predoc & ", " & ptotal & ", '" & descri & "', 'S', 0, " & predoc & ", '')"
            'Control de Stock
            canact = 0
            RS1.Open "select bod_canmer from b_bodegas where bod_codpro='" & Trim(LimpiaDato(CodMer)) & "' and bod_codbod=" & CodBod, vg_db, adOpenStatic
            If Not RS1.EOF Then
                Do While Not RS1.EOF
                    If Option1(1) = True Then canact = RS1!bod_canmer + canmer ' Recibido
                    If Option1(0) = True Then canact = RS1!bod_canmer - canmer ' Entregado
                    RS1.MoveNext
                Loop
                vg_db.Execute "update b_bodegas set bod_canmer=" & canact & " " & _
                              "where bod_codpro='" & Trim(LimpiaDato(CodMer)) & "' and bod_codbod=" & CodBod
            End If
            RS1.Close: Set RS1 = Nothing: numlin = numlin + 1
            'Actualiza Stock en columna oculta
            vaSpread1.Row = i: vaSpread1.Col = 1
            RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                     "WHERE bod.bod_codbod=" & CodBod & " and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
            vaSpread1.Col = 9
            If Not RS1.EOF Then vaSpread1.Text = RS1!bod_canmer Else vaSpread1.Text = 0
            RS1.Close: Set RS1 = Nothing
        End If
    Next i
    vg_db.CommitTrans
    Gl_Ac_Botones Me, 4, 3, ""
    Frame1.Enabled = False
    Frame2.Enabled = False
    vaSpread1.Col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    I_Traspaso Me
Case 5 'Anular
    'If Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(fpDateTime1(0).Text), "mm/yyyy") < Format(Now, "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If (Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(fpDateTime1(0).Text), "mm/yyyy") < Format(Now, "mm/yyyy")) Or Format(CDate(fpDateTime1(0).Text), "mm/yyyy") < Format(Month(Now) - 1 & "/" & Year(Now), "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Anula documento...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    CodBod = Val(fg_codigocbo(Combo1, 1, 10, ""))
    vg_db.BeginTrans
    'Encabezado
    vg_db.Execute "update b_totventas set tov_estdoc='A' where tov_rutcli='" & Trim(LimpiaDato(fpText1(0).Text)) & "' " & _
                  "and tov_tipdoc='TR' and tov_numdoc=" & fpLongInteger1(0).Value
    Label1.Caption = "ANULADA"
    'Detalle
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: numlin = i
        vaSpread1.Col = 1: CodMer = Trim(LimpiaDato(vaSpread1.Text))
        vaSpread1.Col = 4: If Option1(1).Value = True Then canmin = LimpiaDato(vaSpread1.Text) Else canmer = LimpiaDato(vaSpread1.Text)
        vaSpread1.Col = 5: If Option1(1).Value = True Then canmer = LimpiaDato(vaSpread1.Text) Else canmin = 0
        'Control de Stock
        canact = 0
        RS1.Open "select bod_canmer from b_bodegas where bod_codpro='" & CodMer & "' and bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")), vg_db, adOpenStatic
        If Not RS1.EOF Then
            Do While Not RS1.EOF
                    If Option1(1) = True Then canact = RS1!bod_canmer - canmer ' Anula Recibido
                    If Option1(0) = True Then canact = RS1!bod_canmer + canmer ' Anula Entregado
                RS1.MoveNext
            Loop
            vg_db.Execute "update b_bodegas set bod_canmer=" & canact & " " & _
                          "where bod_codpro='" & Trim(LimpiaDato(CodMer)) & "' and bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, ""))
        End If
        RS1.Close: Set RS1 = Nothing
        'Actualiza Stock en columna oculta
        vaSpread1.Row = i: vaSpread1.Col = 1
        RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod=" & CodBod & " and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
        vaSpread1.Col = 9
        If Not RS1.EOF Then vaSpread1.Text = RS1!bod_canmer Else vaSpread1.Text = 0
        RS1.Close: Set RS1 = Nothing
    Next i
    vg_db.CommitTrans
    Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""
Case 8 'Busqueda
    If Trim(fpText1(0).Text) = "" Then MsgBox "Debe seleccionar casino...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    vg_codigo = Trim(fpText1(0).Text)
    vg_nombre = "TR"
    B_SalBod.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    BuscaDoc Val(vg_codigo)
Case 9 'Imprimir
    I_Traspaso Me
Case 12 'Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
vg_swpegreceta = 0
If Err = 3034 Then Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub BuscaDoc(Codigo As Long)
'Encabezado
RS2.Open "select tov.tov_numinf, tov.tov_totdoc, tov.tov_numdoc, tov.tov_codbod, tov.tov_fecemi, " & _
         "tov.tov_codser, tov.tov_estdoc, tov.tov_codcas " & _
         "from b_totventas tov, b_clientes cli " & _
         "where tov.tov_rutcli='" & LimpiaDato(Trim(fpText1(0).Text)) & "' " & _
         "and tov.tov_tipdoc='TR' " & _
         "and tov.tov_numdoc=" & Val(Codigo) & " " & _
         "and tov.tov_rutcli=cli.cli_codigo", vg_db, adOpenStatic
If Not RS2.EOF Then
    Frame1.Enabled = False
    Frame2.Enabled = False
    vaSpread1.Col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    vaSpread1.MaxRows = 0
    Do While Not RS2.EOF
        Est = True
        fpLongInteger1(0).Text = RS2!tov_numdoc
        fpLongInteger1(1).Text = RS2!tov_numinf
        Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 10, fg_pone_cero(Str(RS2!tov_codbod), 10))
        fpDateTime1(0).Text = RS2!tov_fecemi
        Option1(RS2!tov_codser) = True
        Label1.Caption = IIf(RS2!tov_estdoc = "", "", "ANULADA")
        fpText1(1).Text = RS2!tov_codcas
        fpText1_LostFocus 1
        Label2.Caption = Format(RS2!tov_totdoc, fg_Pict(9, vg_DPr))
        Est = False
        RS2.MoveNext
    Loop
Else
    Gl_Ac_Botones Me, 4, 2, ""
    fpLongInteger1(0).Enabled = False
    RS2.Close: Set RS2 = Nothing
    Exit Sub
End If
RS2.Close: Set RS2 = Nothing
'Detalle
RS1.Open "select dev.dev_codmer, dev.dev_canmin, dev.dev_canmer, dev.dev_predoc, " & _
         "dev.dev_ptotal, dev.dev_descri, uni.uni_nombre " & _
         "from b_detventas dev, b_productos pro ,a_unidad uni " & _
         "where dev.dev_rutcli='" & LimpiaDato(Trim(fpText1(0).Text)) & "'" & _
         "and dev.dev_tipdoc='TR' " & _
         "and dev.dev_numdoc=" & Val(Codigo) & " " & _
         "and dev.dev_codmer=pro.pro_codigo " & _
         "and pro.pro_coduni=uni.uni_codigo order by dev.dev_numlin", vg_db, adOpenStatic
If Not RS1.EOF Then
    i = 1
    Do While Not RS1.EOF
        vaSpread1.MaxRows = i
        vaSpread1.Row = i
        vaSpread1.Col = 1: vaSpread1.Text = RS1!dev_codmer
        vaSpread1.Col = 2: vaSpread1.Text = RS1!dev_descri
        vaSpread1.Col = 3: vaSpread1.Text = RS1!uni_nombre
        vaSpread1.Col = 4: vaSpread1.Text = IIf(Option1(1).Value = True, RS1!dev_canmin, RS1!dev_canmer)
        vaSpread1.Col = 5: vaSpread1.Text = IIf(Option1(1).Value = True, RS1!dev_canmer, 0)
        vaSpread1.Col = 6: vaSpread1.Text = RS1!dev_predoc
        vaSpread1.Col = 7: vaSpread1.Text = RS1!dev_ptotal
        'Trae Stock
        RS2.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & Trim(RS1!dev_codmer) & "'", vg_db, adOpenStatic
        vaSpread1.Col = 9
        If Not RS2.EOF Then vaSpread1.Text = RS2!bod_canmer Else vaSpread1.Text = 0
        RS2.Close: Set RS2 = Nothing
        RS1.MoveNext: i = i + 1
    Loop
End If
RS1.Close: Set RS1 = Nothing
vg_codigo = ""
Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""
End Sub

Sub Limpia()
Est = True
Label1.Caption = ""
Frame1.Enabled = True
Frame2.Enabled = True
fpDateTime1(0).Text = Format(Date, "dd/mm/yyyy")
fpLongInteger1(0).Enabled = True
fpLongInteger1(0).Text = ""
fpLongInteger1(1).Text = MuestraFolio
fpText1(1).Text = ""
fpayuda(1).Caption = ""
Combo1(1).ListIndex = IIf(Combo1(1).ListCount = 1, 0, -1)
vaSpread1.MaxRows = 0
vaSpread1.Col = -1: vaSpread1.Row = -1: vaSpread1.Lock = True
vaSpread1.Col = 4: vaSpread1.Row = -1: vaSpread1.Lock = False
vaSpread1.Col = 5: vaSpread1.Row = -1: vaSpread1.Lock = False
fpText1(0).Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText1(0).Text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
Option1(1).Value = True
Label2.Caption = Format(0, fg_Pict(9, vg_DPr))
Gl_Ac_Botones Me, 4, 2, ""

Est = False
End Sub

Private Function MuestraFolio() As Long
Dim RS As New ADODB.Recordset
MuestraFolio = 0
RS.Open "select max(inf_numero) as folio from a_infcfcfofi where inf_tipo='T' and inf_feccie=0 and isnull(inf_usuario)", vg_db, adOpenStatic
If Not RS.EOF Then MuestraFolio = TipoDato(RS!folio, 0)
RS.Close: Set RS = Nothing
End Function
Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long
Select Case Button.Index
Case 1
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Gen"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Col = 1: vaSpread1.Row = i
        If Trim(vaSpread1.Text) = Trim(vg_codigo) Then MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Next i
    vaSpread1.Row = vaSpread1.ActiveRow
    RS1.Open "SELECT pro.pro_codigo, pro.pro_propon, pro.pro_nombre, uni.uni_nombre " & _
             "FROM b_productos AS pro, a_unidad AS uni " & _
             "WHERE pro.pro_coduni=uni.uni_codigo and pro.pro_codigo='" & vg_codigo & "'", vg_db, adOpenStatic
    If Not RS1.EOF Then
        i = vaSpread1.MaxRows + 1
        Do While Not RS1.EOF
            vaSpread1.MaxRows = i
            vaSpread1.Row = vaSpread1.MaxRows
            vaSpread1.Col = 1: vaSpread1.Text = RS1!pro_codigo
            vaSpread1.Col = 2: vaSpread1.Text = RS1!pro_nombre
            vaSpread1.Col = 3: vaSpread1.Text = RS1!uni_nombre
            vaSpread1.Col = 4: vaSpread1.Text = 0
            vaSpread1.Col = 5: vaSpread1.Text = 0
            vaSpread1.Col = 6: vaSpread1.Text = IIf(Option1(1).Value = True, 0, RS1!pro_propon)
            vaSpread1.Col = 7: vaSpread1.Text = 0
            vaSpread1.Col = 8: vaSpread1.Text = "N" 'No bloquedo
            'Trae Stock
            RS2.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                     "WHERE bod.bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                     "and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & Trim(RS1!pro_codigo) & "'", vg_db, adOpenStatic
            vaSpread1.Col = 9
            If Not RS2.EOF Then vaSpread1.Text = RS2!bod_canmer Else vaSpread1.Text = 0
            RS2.Close: Set RS2 = Nothing
            RS1.MoveNext
            i = i + 1
        Loop
    End If
    RS1.Close: Set RS1 = Nothing
    If vaSpread1.MaxRows = 1 Then Gl_Ac_Botones Me, 4, 2, ""
    vaSpread1.Col = 4: vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.SetActiveCell 4, vaSpread1.MaxRows
    If Option1(1) = True Then
        vaSpread1.Col = 6: vaSpread1.Row = 0
        vaSpread1.Text = "Precio"
        vaSpread1.Row = -1: vaSpread1.Lock = False
    Else
        vaSpread1.Col = 6: vaSpread1.Row = 0
        vaSpread1.Text = "P.M.P."
        vaSpread1.Row = -1: vaSpread1.Lock = True
    End If
    If vaSpread1.Enabled = True Then vaSpread1.SetFocus
Case 2
    If vaSpread1.MaxRows = 0 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    If vaSpread1.MaxRows = 0 Then Gl_Ac_Botones Me, 4, 5, ""
    If vaSpread1.Enabled = True Then vaSpread1.SetFocus
End Select
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
Dim canrea As Double, propon As Double, CodMer As String, i As Long, subtot As Double
Select Case Col
Case 4, 6
    vaSpread1.Row = Row
    vaSpread1.Col = 1: CodMer = vaSpread1.Text
    vaSpread1.Col = 4: canrea = Format(vaSpread1.Text, fg_Pict(9, 2))
    vaSpread1.Col = 6: propon = Format(vaSpread1.Text, fg_Pict(9, 2))
    vaSpread1.Col = 7: vaSpread1.Text = Format(canrea * propon, fg_Pict(9, 2))
    vaSpread1.Col = 9: canbod = Format(vaSpread1.Text, fg_Pict(9, 2))
    '------- Total General ---------
    subtot = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 7: subtot = subtot + Format(vaSpread1.Text, fg_Pict(9, vg_DPr))
    Next
    Label2.Caption = Format(subtot, fg_Pict(9, vg_DPr))
    '-------------------------------
    vaSpread1.Row = Row
    If canbod - canrea >= 0 Or Option1(1).Value = True Then
        vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(2).FillColor
        vaSpread1.Col = 8: vaSpread1.Text = "N"  'No Bloqueado
        Exit Sub
    End If
    vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
    vaSpread1.Col = 8: vaSpread1.Text = "S"  'Bloqueado
End Select
End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Or vaSpread1.MaxRows = 0 Then Exit Sub
If Option1(1).Value = True Then
    If vaSpread1.ActiveCol = 4 Then vaSpread1.SetActiveCell 5, vaSpread1.ActiveRow - 1: Exit Sub
    If vaSpread1.ActiveCol = 5 Then vaSpread1.SetActiveCell 6, vaSpread1.ActiveRow - 1: Exit Sub
    If vaSpread1.ActiveCol = 6 And vaSpread1.ActiveRow - 1 <> vaSpread1.MaxRows Then vaSpread1.SetActiveCell 4, vaSpread1.ActiveRow
End If
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If Row = 0 Then Exit Sub
Dim Stock As String, Nombre As String
TipWidth = 4000
ShowTip = True
MultiLine = 2
vaSpread1.Row = Row: vaSpread1.Col = 9: Stock = vaSpread1.Text
vaSpread1.Row = Row: vaSpread1.Col = 2: Nombre = vaSpread1.Text
TipText = "Bodega   : " & Trim(Left(Combo1(1).Text, 50)) & vbCrLf & _
          "Producto : " & Trim(Nombre) & vbCrLf & _
          "Stock       : " & Format(Trim(Stock), fg_Pict(9, vg_DCa))
End Sub


