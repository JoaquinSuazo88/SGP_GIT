VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_Traspa 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspasos"
   ClientHeight    =   7935
   ClientLeft      =   2295
   ClientTop       =   2040
   ClientWidth     =   17535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   17535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   2970
      TabIndex        =   9
      Top             =   375
      Width           =   11220
      Begin VB.CommandButton Cmd_ImportarGuiaCD 
         Caption         =   "Importar Guía CD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8640
         TabIndex        =   37
         Top             =   1320
         Width           =   2175
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   2085
         TabIndex        =   1
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
         ThreeDInsideHighlightColor=   -2147483637
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
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
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
         Text            =   "0"
         MaxValue        =   "2147483647"
         MinValue        =   "0"
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
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "Entrada"
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
         TabIndex        =   4
         Top             =   1590
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "Salida"
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
         TabIndex        =   6
         Top             =   1590
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1155
         Width           =   3195
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   0
         Left            =   2085
         TabIndex        =   0
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
         ThreeDInsideHighlightColor=   -2147483637
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
         ThreeDTextHighlightColor=   -2147483637
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
         Left            =   2085
         TabIndex        =   2
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
         ThreeDInsideHighlightColor=   -2147483637
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
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   1
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
         Left            =   9555
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   495
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
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
         NoSpecialKeys   =   0
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
      Begin EditLib.fpDoubleSingle FDCLogistico 
         Height          =   315
         Left            =   7080
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1215
         _Version        =   196608
         _ExtentX        =   2143
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
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
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
         Text            =   "0"
         DecimalPlaces   =   0
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "0"
         NegFormat       =   0
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Costo Logistico"
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
         Left            =   5520
         TabIndex        =   38
         Top             =   1605
         Width           =   1320
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   3480
         Picture         =   "M_Traspa.frx":0000
         Top             =   720
         Visible         =   0   'False
         Width           =   480
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
         Left            =   8715
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
         Width           =   6855
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3465
         Picture         =   "M_Traspa.frx":030A
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
         Left            =   240
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
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label Label3 
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
         Index           =   6
         Left            =   240
         TabIndex        =   15
         Top             =   195
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3465
         Picture         =   "M_Traspa.frx":0614
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
         Left            =   240
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
         Left            =   240
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
         Left            =   240
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
         Width           =   6975
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
         Width           =   6975
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
         Width           =   6855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   17535
      _ExtentX        =   30930
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Frame Frame3 
      Height          =   4290
      Left            =   225
      TabIndex        =   24
      Top             =   2655
      Width           =   17205
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   3840
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   3840
         Visible         =   0   'False
         Width           =   8295
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3405
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   16785
         _Version        =   393216
         _ExtentX        =   29607
         _ExtentY        =   6006
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
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
         MaxCols         =   21
         MaxRows         =   20
         ProcessTab      =   -1  'True
         SelectBlockOptions=   1
         SpreadDesigner  =   "M_Traspa.frx":091E
         ClipboardOptions=   0
      End
      Begin VB.Frame Frame4 
         Height          =   450
         Left            =   14880
         TabIndex        =   30
         Top             =   3615
         Width           =   1665
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   45
            TabIndex        =   31
            Top             =   135
            Width           =   1590
         End
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   2520
         Picture         =   "M_Traspa.frx":15D8
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Documento"
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
         Left            =   13125
         TabIndex        =   32
         Top             =   3870
         Width           =   1710
      End
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   9870
      TabIndex        =   20
      Top             =   7065
      Width           =   7425
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   330
         TabIndex        =   8
         Top             =   240
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   688
         ButtonWidth     =   2302
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agr. Prod."
               Description     =   "Agregar Productos"
               Object.ToolTipText     =   "Agregar Producto"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Elim. Prod. "
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
               Picture         =   "M_Traspa.frx":18E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Traspa.frx":1BFC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad sobrepasa Stock actual"
         Height          =   450
         Index           =   1
         Left            =   5040
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
         Left            =   4650
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
Dim RS5 As New ADODB.Recordset
Dim RS6 As New ADODB.Recordset
Dim MsgTitulo As String, est As Boolean, Eststo As Boolean
Dim Row_Activo  As Long
Dim CodigoProd  As String
Dim ant_canrec As Double


Private Sub Combo1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset
Dim i As Long
If est Then Exit Sub
Select Case Index

    Case 1
        
        If vaSpread1.MaxRows = 0 Then Exit Sub
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If RS1.State = 1 Then RS1.Close
            RS1.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient

            RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                     "WHERE bod.bod_codpro = pro.pro_codigo " & _
                     "AND   bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                     "AND   pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "' " & _
                     "AND   pro.pro_ctrsto = 1", vg_db, adOpenStatic
            vaSpread1.Col = 9
            If Not RS1.EOF Then vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
            RS1.Close
            Set RS1 = Nothing
            
            'REvisa color
            Dim canrea As Double, canbod As Double
            vaSpread1.Col = 4
            If Trim(vaSpread1.text) <> "" Then canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
            
            vaSpread1.Col = 9
            If Trim(vaSpread1.text) <> "" Then canbod = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
            
            If canbod - canrea < 0 And Option1(0).Value = True Then
                
                vaSpread1.Col = -1
                vaSpread1.BackColor = Shape1(1).FillColor
                vaSpread1.Col = 8
                vaSpread1.text = "S"  'Bloqueado
            
            Else
                
                vaSpread1.Col = -1
                vaSpread1.BackColor = Shape1(2).FillColor
                vaSpread1.Col = 8
                vaSpread1.text = "N" 'No Bloqueado
            
            End If
        
        Next i

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Cmd_ImportarGuiaCD_Click()

On Error GoTo Man_Error

P_ExportarArchivos.Show 0, M_Traspa

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FDCLogistico_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga
'fpLongInteger1(1).text = MuestraFolio

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Partida.Timer1.Enabled = False

Me.Height = 8415
Me.Width = 17625

fg_centra Me

est = False
Eststo = False
vg_GuiaCD = ""
vg_FechaEmision_GGD = "0000:00:00"

EspFecha fpDateTime1(0)
Me.HelpContextID = vg_OpcM
MsgTitulo = "Traspaso"

Dim X As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
Gl_Mo_Botones Me, 4
vaSpread1.Row = -1

vaSpread1.Col = 4
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = vg_DCa

vaSpread1.Col = 5
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = 2 'vg_DCa

vaSpread1.Col = 6
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = vg_DPr

vaSpread1.Col = 7
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = vg_DCa

vaSpread1.Col = 9
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = vg_DCa

vaSpread1.Col = 13
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = vg_DCa

vaSpread1.Col = 18
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = vg_DCa

FDCLogistico.UseSeparator = True
FDCLogistico.DecimalPoint = vg_CDec
FDCLogistico.Separator = vg_CSep
FDCLogistico.DecimalPlaces = vg_DPr

'-------> Cargar Combo Bodega
CargarDatoCombo Combo1, 1, "b_clientes", "cli_", "CliBod", "N"
Limpia 2
TraerFechaCierre

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Resize()

On Error GoTo Man_Error

If Me.WindowState = 2 Then
    
'   Frame3.Left = (Me.Width \ 2) - (Frame3.Width \ 2)
'   Frame1.Left = (Me.Width \ 2) - (Frame1.Width \ 2)
'   Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)

ElseIf Me.WindowState = 0 Then
    
'   Frame3.Left = 255
'   Frame3.Width = 13365
'   Frame1.Left = 2370
'   Frame2.Left = 6990
'   vaSpread1.Width = 13060
   Me.Refresh

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Unload(Cancel As Integer)

Partida.Timer1.Enabled = True

End Sub

Private Sub fpDateTime1_Change(Index As Integer)

On Error GoTo Man_Error

If est Then Exit Sub

If Option1(0).Value = True Then
   
   Option1_Click 0

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

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

Private Sub fpLongInteger1_LostFocus(Index As Integer)

On Error GoTo Man_Error

Select Case Index

    Case 0
        
        If Val(fpLongInteger1(0).Value) > 0 Then
           
           
           BuscaDoc fpLongInteger1(0).Value
        
        End If
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_Change(Index As Integer)

On Error GoTo Man_Error

fpayuda(Index).Caption = ""

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_LostFocus(Index As Integer)

On Error GoTo Man_Error

If fpText1(Index).text = "" Then Exit Sub
Dim RS1 As New ADODB.Recordset

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Select Case Index

    Case 0
        
        RS1.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo = '" & fpText1(0).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
        If Not RS1.EOF Then
            
            Do While Not RS1.EOF
               
               fpayuda(Index).Caption = RS1!cli_nombre
               Gl_Ac_Botones Me, 4, 2, ""
               fpText1(0).Enabled = False
               RS1.MoveNext
            
            Loop
        
        Else
            
            RS1.Close
            Set RS1 = Nothing
            MsgBox "Contrato no existe...", vbExclamation + vbOKOnly, MsgTitulo
            Limpia 2
            If fpText1(0).Enabled = True Then fpText1(0).SetFocus
            Exit Sub
        
        End If
        RS1.Close
        Set RS1 = Nothing
    
    Case 1
        
        RS1.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText1(1).text)) & "' AND cli_codigo <> '" & LimpiaDato(Trim(fpText1(0).text)) & "' AND (cli_tipo=2 OR cli_tipo = 0)", vg_db, adOpenStatic
        
        If Not RS1.EOF Then
            
            fpayuda(Index).Caption = RS1!cli_nombre
            
            If Trim(fpText1(1).text) = "CD" And Option1(0).Value = False Then
               
               RS1.Close
               Set RS1 = Nothing
               ChequearOCompras
               
               If Val(fpLongInteger1(0).Value) > 0 Then
                  
                  BuscaDoc fpLongInteger1(0).Value
               
               End If
               Exit Sub
            
            End If
        
        Else
            
            RS1.Close
            Set RS1 = Nothing
            MsgBox "Contrato traspaso no existe...", vbExclamation + vbOKOnly, MsgTitulo
            fpText1(1) = ""
            Exit Sub
        
        End If
        RS1.Close
        Set RS1 = Nothing

End Select

If Trim(fpText1(0).text) = Trim(fpText1(1).text) Then
    
    MsgBox "No se puede realizar transferencia en el mismo contrato...", vbExclamation + vbOKOnly, MsgTitulo
    If Index = 0 Then Limpia 2 Else fpText1(Index).text = ""
    Exit Sub

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

    Case 0, 1
        
        vg_codigo = 0
        vg_left = fpayuda(Index).Left + 1920
        B_TabEst.LlenaDatos "b_clientes", "cli_", "Contrato", IIf(Index = 1, "Traspaso" & LimpiaDato(Trim(fpText1(0).text)), "Contrato")
        B_TabEst.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
        fpText1(Index) = Trim(vg_codigo)
        fpayuda(Index).Caption = vg_nombre
        
        Select Case Index
            
            Case 0
                
                If Trim(vg_codigo) <> fpText1(Index) Then Limpia 2
                fpText1_LostFocus 0
                If Trim(fpText1(Index).text) = "" Then Exit Sub
                If fpDateTime1(Index).Enabled = True Then fpDateTime1(Index).SetFocus
                Gl_Ac_Botones Me, 4, 2, ""
            
            Case 1
                
                If fpText1(1).Enabled = True Then fpText1(1).SetFocus
                fpText1_LostFocus 1
        
        End Select
    
    Case 2
        
        vg_FDC = "OC"
        vg_RDC = ""
        B_Guias.Cargar_DoctoGrilla Me, "OC", "Ordenes de Compras", Trim(UCase(fpText1(0).text)), "traspa", 0
        B_Guias.Show 1
        If Trim(vg_Guias) <> "" Then
           
           '------- Total General ---------
           vaSpread1.Visible = False
           If Option1(1).Value = True Then
              
              vaSpread1.ColWidth(2) = 23.5
              vaSpread1.Col = 13
              vaSpread1.ColHidden = False
              
              vaSpread1.Col = 14
              vaSpread1.ColHidden = False
              
              vaSpread1.Col = 15
              vaSpread1.ColHidden = False
           
           End If
           
           subtot = 0
           For i = 1 To vaSpread1.MaxRows
               
               vaSpread1.Row = i
               vaSpread1.Col = 1
               If Trim(vaSpread1.text) <> "" Then
                  
                  vaSpread1.Col = 6
                  subtot = subtot + Format(IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0), fg_Pict(9, vg_DPr))
           
               End If
               
           Next
           Label2.Caption = Format(subtot, fg_Pict(9, vg_DPr))
           '-------------------------------
           vaSpread1.Row = 1
           Text2(0).Visible = True
           Text2(1).Visible = True
           vaSpread1.Visible = True
        
        End If
    
    Case 5 '-------> Mostrar para colombia los productos asociados a formato de compras
        
        If vg_pais = "CO" Then Exit Sub 'mientras tantos que va estar en procesos
        vg_left = fpayuda(1).Left + 1290
        Me.Refresh
        vg_codigo = ""
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        vg_codigo = vaSpread1.text
        
        vaSpread1.Col = 16
        vg_nombre = Trim(vaSpread1.text)
        
        B_TabEst.LlenaDatos vg_nombre, vg_codigo, "Productos SAC", "CamPSAC"
        B_TabEst.Show 1, Me
        Me.Refresh
        If Trim(vg_codigo) = "" Or Val(vg_codigo) = 0 Then Exit Sub
        Text2(0).text = Trim(vg_codigo)
        Text2(1).text = Trim(vg_nombre)
        vaSpread1.Row = vaSpread1.ActiveRow
        
        vaSpread1.Col = 16
        vaSpread1.text = vg_codigo
        
        vaSpread1.Col = 17
        vaSpread1.text = vg_nombre

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset
Dim i As Long, Cantidad As Double
Label3(0).Caption = IIf(Index = 1, "Contrato Origen", "Contrato Destino")

Select Case Index
    
    Case 0
        
        Cmd_ImportarGuiaCD.Enabled = False
        
        FDCLogistico.Enabled = False
        FDCLogistico.text = ""
        
        vaSpread1.Visible = False
        vaSpread1.Col = 7
        vaSpread1.Row = -1
        vaSpread1.ColHidden = True
        
        vaSpread1.Col = 5
        vaSpread1.Row = 0
        vaSpread1.text = "P.M.P."
    '    Label4.Left = 7845: Frame4.Left = 8340
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            If Trim(LimpiaDato(vaSpread1.text)) <> "" And IsDate(fpDateTime1(0).text) Then
            
                If RS1.State = 1 Then RS1.Close
                RS1.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
    
                RS1.Open "SELECT TOP 1 b.ppd_propon, Max(b.ppd_fecdia) AS ppd_fecdia FROM b_productos a, b_productospmpdia b " & _
                         "WHERE a.pro_codigo = b.ppd_codpro " & _
                         "AND   b.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                         "AND   b.ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND b.ppd_fecdia <= " & Format(CDate(fpDateTime1(0).text), "yyyymmdd") & " " & _
                         "AND   a.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "' " & _
                         "AND   a.pro_ctrsto = 1 GROUP BY b.ppd_propon HAVING (b.ppd_propon)>0 ORDER BY Max(b.ppd_fecdia) DESC", vg_db, adOpenStatic
                
                If Not RS1.EOF Then
                    
                    vaSpread1.Col = 4
                    vaSpread1.Lock = False
                    Cantidad = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
                    
                    vaSpread1.Col = 5
                    vaSpread1.Lock = True
                    vaSpread1.text = Format(RS1!ppd_propon, fg_Pict(9, 2))
                    
                    vaSpread1.Col = 6
                    vaSpread1.Lock = True
                    vaSpread1.text = Format(Format(Cantidad, fg_Pict(9, vg_DCa)) * RS1!ppd_propon, fg_Pict(9, vg_DPr))
                
                End If
                RS1.Close
                Set RS1 = Nothing
                
            End If
        
        Next
        
        Combo1_Click 1
        
        If vaSpread1.MaxRows > 0 And vaSpread1.Enabled = True And est = False Then
           
           vaSpread1.Visible = True
           vaSpread1.SetFocus
           vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow
        
        End If
        
        vaSpread1.ColWidth(2) = 87.88 '57.55  '33.25
        
        vaSpread1.Col = 13
        vaSpread1.ColHidden = True
        
        vaSpread1.Col = 14
        vaSpread1.ColHidden = True
        
        vaSpread1.Col = 15
        vaSpread1.ColHidden = True
        vaSpread1.Visible = True
        
        vaSpread1.Col = 18
        vaSpread1.ColHidden = True
        vaSpread1.Visible = True
        
        vaSpread1.Col = 19
        vaSpread1.ColHidden = True
        vaSpread1.Visible = True
        
        vaSpread1.Col = 20
        vaSpread1.ColHidden = True
        vaSpread1.Visible = True
        
        Text2(0).Visible = False
        Text2(1).Visible = False
        Image1(2).Visible = False
        Image1(5).Visible = False
    
    Case 1
        
        Cmd_ImportarGuiaCD.Enabled = True
        
        FDCLogistico.Enabled = True
        FDCLogistico.text = ""
        
        vaSpread1.Visible = False
        vaSpread1.Col = 7
        vaSpread1.Row = -1
        vaSpread1.ColHidden = False
        
        vaSpread1.Col = 5
        vaSpread1.Row = 0
        vaSpread1.text = "Precio Documento"
    '    Label4.Left = 6705: Frame4.Left = 7200
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            
            vaSpread1.Col = 5
            vaSpread1.Lock = False
            vaSpread1.text = 0
            
            vaSpread1.Col = 6
            vaSpread1.text = 0
            
            vaSpread1.Col = 7
            vaSpread1.Lock = False
        
        Next
        
        'REvisa color
        vaSpread1.Col = -1
        vaSpread1.Row = -1
        vaSpread1.BackColor = Shape1(2).FillColor
        
        vaSpread1.Col = 8
        vaSpread1.text = "N"
        
        If vaSpread1.MaxRows > 0 And vaSpread1.Enabled = True And est = False Then
           
           vaSpread1.Visible = True
           vaSpread1.SetFocus
           vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow
        
        End If
        
        If vg_FDC = "OC" Then
           
           vaSpread1.ColWidth(2) = 33.5
           vaSpread1.Col = 13
           vaSpread1.ColHidden = False
           
           vaSpread1.Col = 14
           vaSpread1.ColHidden = False
           
           vaSpread1.Col = 15
           vaSpread1.ColHidden = False
           
           Text2(0).Visible = True
           Text2(1).Visible = True
           
           vaSpread1.Row = vaSpread1.ActiveRow
           vaSpread1.Col = 16
           Text2(0).text = Trim(vaSpread1.text)
           
           vaSpread1.Col = 17
           Text2(1).text = Trim(vaSpread1.text)
           Image1(5).Visible = True
           
        Else
           
           vaSpread1.ColWidth(2) = 77.88 '47.88
           vaSpread1.Col = 13
           vaSpread1.ColHidden = True
           
           vaSpread1.Col = 14
           vaSpread1.ColHidden = True
           
           vaSpread1.Col = 15
           vaSpread1.ColHidden = True
           
           vaSpread1.Col = 18
           vaSpread1.ColHidden = True
           
           vaSpread1.Col = 19
           vaSpread1.ColHidden = True
           
           vaSpread1.Col = 20
           vaSpread1.ColHidden = True
           
           Text2(0).Visible = False
           Text2(1).Visible = False
           Image1(5).Visible = False
        
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

Dim RS1     As New ADODB.Recordset
Dim RS2     As New ADODB.Recordset
Dim Folio   As Long
Dim rutcli  As String
Dim codcas  As String
Dim tipdoc  As String
Dim NumDoc  As Long
Dim codbod  As Long
Dim codser  As Long
Dim i       As Long
Dim canact  As Double
Dim acepre  As String
Dim numlin  As Long
Dim codmer  As String
Dim fecpro  As Date
Dim canmer  As Double
Dim candoc  As Double
Dim canrec  As Double
Dim canmin  As Double
Dim predoc  As Double
Dim ptotal  As Double
Dim descri  As String
Dim total   As Double
Dim diablq  As Date
Dim coding  As String
Dim cancer  As String
Dim movinv  As String
Dim codsac  As String
Dim fecoc   As Date
Dim canoc   As Double
Dim preoc   As Double
Dim estocs  As Boolean
Dim pmp     As Double
Dim EstGCd  As Boolean
Dim IngGuia As Boolean
Dim NombreArchivoExcel As String
Dim Trans   As Boolean

Trans = False
fecpro = Format(fpDateTime1(0).Value, "dd/mm/yyyy")
codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
TraerFechaCierre

Select Case Button.Index
    
    Case 1, 6 '-------> Nuevo
    
    '   Limpia
        If Button.Index = 6 And vaSpread1.MaxRows > 0 Then If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        Limpia IIf(Button.Index = 1, 6, 2)
        On Error Resume Next: fpLongInteger1(0).TabStop = False: fpLongInteger1(0).SetFocus
    
    Case 8 '-------> Graba
        
        '-------> validar costo logistico
        If CDbl(FDCLogistico.text) < 0 And Option1(1) = True Then
        
           MsgBox "Debe ingresar costo logistico, con valor mayor o igual cero...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If Trim(fpText1(0).text) = "" Or Val(fpLongInteger1(0).text) = 0 Or Trim(fpLongInteger1(0).text) = "" Or Trim(fpText1(1).text) = "" _
        Or Trim(Combo1(1).text) = "" Or Trim(fpDateTime1(0).text) = "" Then MsgBox "Debe ingresar dato importante...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        '-------> Validar si el contrato tiene asignado inventario rotativo
        If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
           Format(fpDateTime1(0).Value, "dd/mm/yyyy") > CDate(vg_ciedia) Then MsgBox "Tiene que realizar cierre diario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then
        
           MsgBox "Documento no corresponde al periodo : " & VgLinea & VgLinea & CierreFecha, vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then
        
           MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        'Validar inventario calendarizado 20201001
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 38) Then
        
           MsgBox "Se esta realizando la toma de inventario en estos momento...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
        End If
        
        'Validar ingreso documento inventario calendarizado 20201001
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 40) Then
        
           MsgBox "No puede ingresar documento, antes de un inventario calendarizado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
        End If
        
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 8) Then
        
           MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        If CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Then
        
           MsgBox "Día se encuentra cerrado, no es posible ingresar...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        '-------> Validar Numero folio

        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        If vg_tipbase = "1" Then
           
           RS1.Open "SELECT DISTINCT tov_numinf, format(tov_fecemi, 'mm/yyyy') AS fecemi FROM b_totventas WHERE tov_numinf = " & Val(fpLongInteger1(1).text) & " AND tov_codbod = " & vg_codbod & "", vg_db, adOpenStatic
        
        Else
           
           RS1.Open "SELECT DISTINCT tov_numinf, CASE WHEN convert(varchar(2),datepart(mm,tov_fecemi)) < 10 THEN '0' + convert(varchar(2),datepart(mm,tov_fecemi)) ELSE convert(varchar(2),datepart(mm,tov_fecemi)) END + '/' + convert(varchar(4),datepart(year, tov_fecemi)) AS fecemi FROM b_totventas WHERE tov_numinf = " & Val(fpLongInteger1(1).text) & " AND tov_codbod = " & vg_codbod & "", vg_db, adOpenStatic
        
        End If
        
        If Not RS1.EOF Then
        
           If RS1!fecemi <> Format(fpDateTime1(0).text, "mm/yyyy") Then
           
              MsgBox "Nº folio corresponde al periodo : " & RS1!fecemi & " " & VgLinea & VgLinea & "Tiene que generar un nuevo folio", vbExclamation + vbOKOnly, MsgTitulo
              RS1.Close
              Set RS1 = Nothing
              Exit Sub
        
           End If
           
        End If
        
        RS1.Close
        Set RS1 = Nothing
        
        If Trim(fpText1(0).text) = Trim(fpText1(1).text) Then Exit Sub
        If Format(Label2.Caption, fg_Pict(9, vg_DPr)) = Format(0, fg_Pict(9, vg_DPr)) Then
        
           MsgBox "El total del documento debe ser mayor a 0...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
           
        End If
        
        cancer = ""
        
        '-------> Borrar lineas en blanco
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If Trim(vaSpread1.text) = "" Then
               
               vaSpread1.DeleteRows vaSpread1.Row, 1
               vaSpread1.MaxRows = vaSpread1.MaxRows - 1
            
            End If
        
        Next i
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 8
            
            If Left(vaSpread1.text, 1) = "S" Then
            
               MsgBox "Existe una cantidad que excende el Stock...", vbExclamation + vbOKOnly, MsgTitulo
               Exit Sub
               
            End If
            
            vaSpread1.Col = 5
            If Val(vaSpread1.text) = 0 Then
            
               MsgBox "Existen precio en cero... ", vbExclamation
               vaSpread1.SetActiveCell 5, vaSpread1.Row
               Exit Sub
            
            End If
            
            EstGCd = True
            
            vaSpread1.Col = 21
            EstGCd = IIf(vaSpread1.text = "GuiaCD", True, False)
            
            vaSpread1.Col = 4
            candoc = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
            If Val(vaSpread1.text) = 0 And (vg_FDC <> "OC" Or EstGCd) And Trim(vg_GuiaCD) <> "1" Then
            
               MsgBox "Existen cantidades documento en cero... ", vbExclamation
               vaSpread1.SetActiveCell 4, vaSpread1.Row
               Exit Sub
            
            End If
            
            vaSpread1.Col = 7
            canrec = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
            If Val(vaSpread1.text) = 0 And Option1(1) = True And vg_FDC <> "OC" And Trim(vg_GuiaCD) <> "1" Then
               
               cancer = "Existen cantidades recibidas en cero. "
               Exit For
'               MsgBox "Existen cantidades recibidas en cero... ", vbExclamation
'               vaSpread1.SetActiveCell 7, vaSpread1.Row
'               Exit Sub
        
            ElseIf Val(vaSpread1.text) = 0 And Option1(1) = True And Not EstGCd Then
            
               MsgBox "Existen cantidades recibidas en cero... ", vbExclamation
               vaSpread1.SetActiveCell 7, vaSpread1.Row
               Exit Sub
            
            End If
            
            vaSpread1.Col = 19
            If canrec <> candoc And EstGCd And Trim(vaSpread1.text) = "" Then
            
               MsgBox "Debe ingresar la descripción del motivo... ", vbExclamation
               vaSpread1.SetActiveCell 19, vaSpread1.Row
               Exit Sub
               
            End If
            
        Next i
        
        '-------> Validar precio
        Dim Precio As Double
        Dim prerea As Double
        Dim porpre As Double
        Dim porpar As Double
        Dim CosLog As Double
        Dim estpre As Boolean
        Dim IdMotivo As Long
        
        porpar = 0
        estpre = False
        
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        RS1.Open "SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'porprepro'", vg_db, adOpenStatic
        If Not RS1.EOF Then porpar = RS1!par_valor
        RS1.Close
        Set RS1 = Nothing
        
        If porpar > 0 And Option1(1).Value = True Then
                  
           For i = 1 To vaSpread1.MaxRows
               
               vaSpread1.Row = i
               vaSpread1.Col = 1
               vaSpread1.Col = 5
               Precio = vaSpread1.text
               
               vaSpread1.Col = 11
               prerea = vaSpread1.text
               
               vaSpread1.Col = 12
               vaSpread1.text = "N"
               
               vaSpread1.Col = 10
               
               If prerea > 0 And vaSpread1.text = "S" Then
                  
                  porpre = (IIf(Precio > prerea, Round(Precio / prerea, 1), Round(prerea / Precio, 1)) * porpar)
                  
                  If porpre > porpar Then
                  
                    estpre = True
                    vaSpread1.Col = 12
                    vaSpread1.text = "S"
               
                  End If
                  
               End If
           
           Next i
        
        End If
        
        If MsgBox(IIf(estpre = True, "Existen precios ingresados, que excede al ultimo precio registrado" & VgLinea & VgLinea & cancer & "                 Desea grabar...", cancer & "Desea grabar..."), vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        
        vg_db.BeginTrans
        
        Trans = True
     
        Toolbar1.Enabled = False
        Image1(5).Visible = False
        
        rutcli = Trim(LimpiaDato(fpText1(0).text))
        codcas = Trim(LimpiaDato(fpText1(1).text))
        tipdoc = "TR"
        NumDoc = Trim(fpLongInteger1(0).text)
        codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
        codser = IIf(Option1(1).Value = True, 1, 0)
        total = Format(Label2.Caption, fg_Pict(9, vg_DPr))
        fpLongInteger1(1).text = MuestraFolio
        CosLog = FDCLogistico.text
        Folio = Trim(fpLongInteger1(1).text)
        
'        If vg_FechaEmision_GGD = "0000:00:00" Then
        
'           vg_FechaEmision_GGD = Format(fpDateTime1(0).text, "dd/mm/yyyy")
           
'        End If

        '-------> Encabezado
        If vg_tipbase = "1" Then
           
           vg_db.Execute "INSERT INTO b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_totdoc, tov_numinf) " & _
                         "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & codbod & ", CDate('" & _
                      Format(fpDateTime1(0).text, "dd/mm/yyyy") & "'), 0, " & codser & ", 0, '', '" & codcas & "', " & total & ", " & Folio & ")"
        
        Else
           
            If vg_FechaEmision_GGD = "0000:00:00" Or Trim(vg_FechaEmision_GGD) = "" Then
           
               vg_db.Execute "INSERT INTO b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_totdoc, tov_numinf, tov_costologistico, tov_origen) " & _
                             "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & codbod & ", '" & _
                             Format(fpDateTime1(0).text, "yyyymmdd") & "', 0, " & codser & ", 0, '', '" & codcas & "', " & total & ", " & Folio & ", " & CosLog & ", '" & IIf(Trim(vg_GuiaCD) = "1", "KeyLogistic", "") & "')"
        
            Else
               
               vg_db.Execute "INSERT INTO b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_totdoc, tov_numinf, tov_costologistico, tov_origen, tov_FechaEmision_GGD) " & _
                             "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & codbod & ", '" & _
                             Format(fpDateTime1(0).text, "yyyymmdd") & "', 0, " & codser & ", 0, '', '" & codcas & "', " & total & ", " & Folio & ", " & CosLog & ", '" & IIf(Trim(vg_GuiaCD) = "1", "KeyLogistic", "") & "', '" & Format(vg_FechaEmision_GGD, "yyyymmdd") & "')"
            
            End If
        End If
        
        '-------> Detalle
        numlin = 1
        For i = 1 To vaSpread1.MaxRows
            
            canmin = 0
            canmer = 0
            
            vaSpread1.Row = i
            
            vaSpread1.Col = 1
            codmer = Trim(LimpiaDato(vaSpread1.text))
            
            vaSpread1.Col = 2
            descri = Trim(LimpiaDato(vaSpread1.text))
            
            vaSpread1.Col = 4
            If Option1(1).Value = True Then
            
               canmin = LimpiaDato(vaSpread1.text)
            
            Else
            
               canmer = LimpiaDato(vaSpread1.text)
            
            End If
            
            vaSpread1.Col = 5
            predoc = LimpiaDato(vaSpread1.text)
            
            vaSpread1.Col = 6
            ptotal = LimpiaDato(vaSpread1.text)
            
            vaSpread1.Col = 7
            If Option1(1).Value = True Then
            
               canmer = LimpiaDato(vaSpread1.text)
               
            Else
            
               canmin = 0
            
            End If
            
            vaSpread1.Col = 12
            acepre = Trim(vaSpread1.text)
            
            vaSpread1.Col = 16
            codsac = Trim(vaSpread1.text)
            
            vaSpread1.Col = 13
            canoc = IIf(Trim(vaSpread1.text) = "" Or Trim(vaSpread1.text) = "0", 0, LimpiaDato(vaSpread1.text))
            
            vaSpread1.Col = 14
            preoc = IIf(Trim(vaSpread1.text) = "" Or Trim(vaSpread1.text) = "0", 0, LimpiaDato(vaSpread1.text))
            
            vaSpread1.Col = 15
            fecoc = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
            
            vaSpread1.Col = 20
            IdMotivo = Val(vaSpread1.text) 'vaSpread1.TypeComboBoxCurSel
            
            vaSpread1.Col = 21
            IngGuia = IIf(vaSpread1.text = "GuiaCD", True, False)
            
            ValidaBod codbod, Trim(LimpiaDato(codmer))
            
            '-------> Actualiza Precio Promedio Ponderado si es Traspaso Recibido
            estocs = True
            If vg_FDC = "OC" And canmin = 0 Then
               
               estocs = False
            
            End If
            
            If Option1(1) = True And estocs Then
               
               pmp = 0
               pmp = Cal_PMP(MuestraCasino(1), codbod, codmer, Format(fpDateTime1(0).text, "dd/mm/yyyy"), predoc, canmer)
               
               If RS1.State = 1 Then RS1.Close
               RS1.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient

               RS1.Open "SELECT DISTINCT a.ppd_propon FROM b_productospmpdia a, b_productos b WHERE a.ppd_codpro = b.pro_codigo AND a.ppd_cencos = '" & MuestraCasino(1) & "' AND a.ppd_codpro = '" & codmer & "' AND a.ppd_fecdia = " & Format(fpDateTime1(0).text, "yyyymmdd") & "", vg_db, adOpenStatic
               If RS1.EOF Then
                  
                  RS1.Close
                  Set RS1 = Nothing
                  
                  If vg_tipbase = "1" Then
                     
                     vg_db.Execute "INSERT INTO b_productospmpdia VALUES ('" & MuestraCasino(1) & "', '" & codmer & "', " & Format(fpDateTime1(0).text, "yyyymmdd") & ", " & pmp & ", 0, " & predoc & ", cdate('" & Format(fpDateTime1(0).text, "dd/mm/yyyy") & "'))"
                  
                  Else
                     
                     vg_db.Execute "INSERT INTO b_productospmpdia VALUES ('" & MuestraCasino(1) & "', '" & codmer & "', " & Format(fpDateTime1(0).text, "yyyymmdd") & ", " & pmp & ", 0, " & predoc & ", '" & Format(fpDateTime1(0).text, "yyyymmdd") & "')"
                  
                  End If
                  
                  '------------Actuliza codigo de ultimo producto de compra  compra---------
                  
                  If RS2.State = 1 Then RS2.Close
                  RS2.CursorLocation = adUseClient
                  vg_db.CursorLocation = adUseClient
                  
                  RS2.Open "SELECT DISTINCT pri_coding FROM b_productosing WHERE pri_codpro = '" & codmer & "'", vg_db, adOpenStatic
                  
                  If Not RS2.EOF Then
                     
                     vg_db.Execute "UPDATE b_contlistpreing SET cpi_codcom = '" & codmer & "' WHERE cpi_coding = '" & RS2!pri_coding & "' AND cpi_cencos = '" & MuestraCasino(1) & "'"
                  
                  End If
                  
                  RS2.Close
                  Set RS2 = Nothing
               
               Else
                  
                  '-------> Actualizar pmp si es menor que cero
                  RS1.Close
                  Set RS1 = Nothing
                  
                  If RS1.State = 1 Then RS1.Close
                  RS1.CursorLocation = adUseClient
                  vg_db.CursorLocation = adUseClient
                  
                  RS1.Open "SELECT DISTINCT a.ppd_propon FROM b_productospmpdia a, b_productos b WHERE a.ppd_codpro = b.pro_codigo AND b.pro_ctrsto = 1 AND a.ppd_cencos='" & MuestraCasino(1) & "' AND a.ppd_codpro = '" & codmer & "' AND a.ppd_fecdia = " & Format(fpDateTime1(0).text, "yyyymmdd") & "", vg_db, adOpenStatic
                  
                  If Not RS1.EOF Then
                     
                     vg_db.Execute "UPDATE b_productospmpdia SET ppd_propon = " & pmp & " WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_codpro = '" & codmer & "' AND ppd_fecdia = " & Format(fpDateTime1(0).text, "yyyymmdd") & ""
                     
                     '------------Actuliza codigo de ultimo producto de compra  compra---------
                     If RS2.State = 1 Then RS2.Close
                     RS2.CursorLocation = adUseClient
                     vg_db.CursorLocation = adUseClient
                     
                     RS2.Open "SELECT DISTINCT pri_coding FROM b_productosing WHERE pri_codpro = '" & codmer & "'", vg_db, adOpenStatic
                     
                     If Not RS2.EOF Then
                        
                        vg_db.Execute "UPDATE b_contlistpreing SET cpi_codcom = '" & codmer & "' WHERE cpi_coding = '" & RS2!pri_coding & "' AND cpi_cencos = '" & MuestraCasino(1) & "'"
                     
                     End If
                     RS2.Close
                     Set RS2 = Nothing
                  
                  End If
                  RS1.Close
                  Set RS1 = Nothing
                  
               End If
            
            End If
            
            '-------> Control de Stock
            movinv = ""

            If RS1.State = 1 Then RS1.Close
            RS1.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            RS1.Open "SELECT * FROM b_productos WHERE pro_codigo = '" & codmer & "' AND pro_ctrsto = 1", vg_db, adOpenStatic
            
            If Not RS1.EOF Then
                
                If Option1(1) = True Then
                   
                   vg_db.Execute "UPDATE b_bodegas SET bod_canmer = bod_canmer+" & canmer & " WHERE bod_codpro = '" & codmer & "' AND bod_codbod = " & codbod & ""
                
                Else
                   
                   vg_db.Execute "UPDATE b_bodegas SET bod_canmer = bod_canmer-" & canmer & " WHERE bod_codpro = '" & codmer & "' AND bod_codbod = " & codbod & ""
                
                End If
                movinv = "S"
            
            Else
                
                movinv = "N"
            
            End If
            RS1.Close
            Set RS1 = Nothing
            
            '-------> Graba Detalle
            If vg_FDC = "OC" And (canmin > 0 Or canmer > 0) Then
               
               vg_db.Execute "INSERT INTO b_detventas (dev_rutcli, dev_tipdoc, dev_numdoc, dev_numlin, dev_codmer, dev_canmin, dev_canmer, dev_predoc, dev_ptotal, dev_descri, dev_mueinv, dev_porcen, dev_precos, dev_coding, dev_acepre, dev_IdMotivo) " & _
                             "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & numlin & ", '" & codmer & "', " & canmin & ", " & canmer & ", " & predoc & ", " & ptotal & ", '" & descri & "', '" & movinv & "', 0, " & predoc & ", '', '" & acepre & "', " & IdMotivo & ")"
            
            ElseIf vg_FDC <> "OC" Then
               
               vg_db.Execute "INSERT INTO b_detventas (dev_rutcli, dev_tipdoc, dev_numdoc, dev_numlin, dev_codmer, dev_canmin, dev_canmer, dev_predoc, dev_ptotal, dev_descri, dev_mueinv, dev_porcen, dev_precos, dev_coding, dev_acepre, dev_IdMotivo) " & _
                             "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & numlin & ", '" & codmer & "', " & canmin & ", " & canmer & ", " & predoc & ", " & ptotal & ", '" & descri & "', '" & movinv & "', 0, " & predoc & ", '', '" & acepre & "', " & IdMotivo & ")"
            
            End If
            
            If vg_FDC = "OC" Or vg_GuiaCD = "1" Then
               
               '-------> detalle orden de compras sac recibido
               If vg_FDC = "OC" Or vg_GuiaCD = "" Then
               
                  vg_db.Execute "INSERT INTO b_ocsacrecibido (ocr_rutpro, ocr_tipdoc, ocr_numdoc, ocr_numlin, ocr_codprodsgp, ocr_codprodsac, ocr_cancom, ocr_precom, ocr_canrec, ocr_fecoc, ocr_canoc, ocr_preoc) " & _
                                "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & numlin & ", '" & codmer & "', '" & codsac & "', " & canmin & ", " & predoc & ", " & canmer & ", '" & IIf(vg_tipbase = "1", fecoc, Format(fecoc, "yyyymmdd")) & "', " & canoc & ", " & preoc & ")"
    
               ElseIf vg_FDC = "" Or vg_GuiaCD = "1" Then
               
                  vg_db.Execute "INSERT INTO b_ocsacrecibido (ocr_rutpro, ocr_tipdoc, ocr_numdoc, ocr_numlin, ocr_codprodsgp, ocr_codprodsac, ocr_cancom, ocr_precom, ocr_canrec, ocr_fecoc, ocr_canoc, ocr_preoc) " & _
                                "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & numlin & ", '" & codmer & "', '" & codsac & "', " & IIf(IngGuia, canmin, 0) & ", " & predoc & ", " & canmer & ", '" & IIf(vg_tipbase = "1", fecoc, Format(fecoc, "yyyymmdd")) & "', " & canoc & ", " & preoc & ")"
               
               End If
               
               If Format(CDate(fecoc), "dd/mm/yyyy") = Format(CDate("00000000"), "dd/mm/yyyy") Then
                  
                  If vg_tipbase = "1" Then
                     
                     vg_db.Execute "UPDATE b_ocsacrecibido SET ocr_fecoc = null WHERE ocr_rutpro = '" & rutcli & "' AND ocr_tipdoc = '" & tipdoc & "' AND ocr_numdoc = " & NumDoc & " AND ocr_numlin = " & numlin & " AND ocr_codprodsgp = '" & codmer & "'"
                  
                  Else
                     
                     vg_db.Execute "UPDATE b_ocsacrecibido SET ocr_fecoc = Null WHERE ocr_rutpro = '" & rutcli & "' AND ocr_tipdoc = '" & tipdoc & "' AND ocr_numdoc = " & NumDoc & " AND ocr_numlin = " & numlin & " AND ocr_codprodsgp = '" & codmer & "'"
                  
                  End If
               
               End If
            
            End If
            
            '-------> Actualiza Stock en columna oculta
            vaSpread1.Row = i
            vaSpread1.Col = 1

            If RS1.State = 1 Then RS1.Close
            RS1.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                     "WHERE bod.bod_codpro = pro.pro_codigo AND bod.bod_codbod = " & vg_codbod & " AND pro.pro_codigo = '" & codmer & "' AND pro.pro_ctrsto = 1", vg_db, adOpenStatic
            
            vaSpread1.Col = 9
            If Not RS1.EOF Then
               
               vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa))
            
            Else
               
               vaSpread1.text = 0
            
            End If
            RS1.Close
            Set RS1 = Nothing
            
            numlin = numlin + 1
        
        Next i
        
        vg_db.CommitTrans
        
        Trans = False
        
        Gl_Ac_Botones Me, 4, 3, ""
        Frame1.Enabled = False
        Frame2.Enabled = False
        vaSpread1.Col = -1
        vaSpread1.Row = -1
        vaSpread1.Lock = True
        
        If Trim(vg_GuiaCD) = "1" Then
        
           MsgBox "Carga Exitosa...", vbInformation, MsgTitulo

          '-------> Crear directorio guias Logistico
          If Dir(dir_trabajo_Inf & "\" & "GuiaLogistico", vbDirectory) = "" Then
       
             MkDir dir_trabajo_Inf & "\" & "GuiaLogistico"
          
          End If
          '-------> Fin crear directorio guias Logistico
      
           NombreArchivoExcel = "GuiaCDLogists_" & rutcli & "_" & Format(Date, "yyyymmdd") & "_" & Format(Time, "HHMMSS")
           Generar_ArchivoExcel "sgp_Sel_ImportacionGuiaCdxUno '" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & codbod & ", '" & Format(fpDateTime1(0).text, "yyyymmdd") & "'", dir_trabajo_Inf & "GuiaLogistico\", NombreArchivoExcel
        
        Else
        
           I_Traspaso Me
        
        End If
        vg_FDC = ""
'        vg_GuiaCD = ""
        Toolbar1.Enabled = True
        Toolbar2.Enabled = True
        
    Case 3 '-------> Eliminar
        
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then MsgBox "Periodo esta cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then MsgBox "No puede eliminar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Then MsgBox "No puede elimnar documento, día esta cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If MsgBox("Elimina documento...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
        
        vg_db.BeginTrans
        
        Trans = True
        
        '-------> Stock
        Dim codpro As String
        codpro = ""
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            numlin = i
            
            vaSpread1.Col = 1
            codmer = Trim(LimpiaDato(vaSpread1.text))
            
            vaSpread1.Col = 4
            If Option1(1).Value = True Then canmin = LimpiaDato(vaSpread1.text) Else canmer = LimpiaDato(vaSpread1.text)
            
            vaSpread1.Col = 7
            If Option1(1).Value = True Then canmer = LimpiaDato(vaSpread1.text) Else canmin = 0
            
            '-------> Control de Stock
            If RS1.State = 1 Then RS1.Close
            RS1.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            RS1.Open "SELECT * FROM b_productos WHERE pro_codigo = '" & codmer & "' AND pro_ctrsto = 1", vg_db, adOpenStatic
            If Not RS1.EOF Then
                
                If Option1(1) = True Then
                    

                    If RS2.State = 1 Then RS2.Close
                    RS2.CursorLocation = adUseClient
                    vg_db.CursorLocation = adUseClient
                    RS2.Open "SELECT bod_canmer FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " AND bod_codpro = '" & codmer & "'", vg_db, adOpenStatic
                    
                    If Not RS2.EOF Then
                        
                        If (Round(RS2!bod_canmer, vg_DCa) - canmer) < 0 Then
                            
                            RS1.Close
                            Set RS1 = Nothing
                            
                            RS2.Close
                            Set RS2 = Nothing
                            
                            vg_db.RollbackTrans
                            MsgBox "Documento no puede ser eliminado. No hay stock suficiente...", vbExclamation + vbOKOnly, MsgTitulo
                            Exit Sub
                        
                        End If
                    
                    End If
                    RS2.Close
                    Set RS2 = Nothing
                    
                    vg_db.Execute "UPDATE b_bodegas SET bod_canmer = bod_canmer-" & canmer & " WHERE bod_codpro = '" & codmer & "' AND bod_codbod = " & vg_codbod
                    codpro = codpro & "'" & codmer & "',"
                
                Else
                    
                    vg_db.Execute "UPDATE b_bodegas SET bod_canmer = bod_canmer+" & canmer & " WHERE bod_codpro = '" & codmer & "' AND bod_codbod = " & vg_codbod
                
                End If
            
            End If
            
            RS1.Close
            Set RS1 = Nothing
        
        Next i
        
        vg_db.Execute "DELETE b_detventas FROM b_detventas WHERE dev_rutcli = '" & Trim(LimpiaDato(fpText1(0).text)) & "' " & _
                      "AND dev_tipdoc = 'TR' AND dev_numdoc = " & fpLongInteger1(0).Value
        vg_db.Execute "DELETE b_detventasimp FROM b_detventasimp WHERE imd_rutdoc = '" & Trim(LimpiaDato(fpText1(0).text)) & "' " & _
                      "AND imd_tipdoc = 'TR' AND imd_numdoc = " & fpLongInteger1(0).Value
        vg_db.Execute "DELETE FROM b_ocsacrecibido WHERE ocr_rutpro = '" & Trim(LimpiaDato(fpText1(0).text)) & "' " & _
                      "AND ocr_tipdoc = 'TR' AND ocr_numdoc = " & fpLongInteger1(0).Value
        vg_db.Execute "DELETE b_totventas FROM b_totventas WHERE tov_rutcli = '" & Trim(LimpiaDato(fpText1(0).text)) & "' " & _
                      "AND tov_tipdoc = 'TR' AND tov_numdoc = " & fpLongInteger1(0).Value & " AND tov_codbod = " & vg_codbod & ""
        
        vg_db.CommitTrans

        Trans = False
        
        Limpia 2
    
    Case 11 '------- Busqueda
        
        If Trim(fpText1(0).text) = "" Then MsgBox "Debe seleccionar contrato...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        vg_codigo = Trim(fpText1(0).text)
        vg_nombre = "TR"
        B_SalBod.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
        BuscaDoc Val(vg_codigo)
    
    Case 12 '------- Imprimir
        
        If Trim(vg_GuiaCD) = "1" Then
                   
          '-------> Crear directorio guias Logistico
          If Dir(dir_trabajo_Inf & "\" & "GuiaLogistico", vbDirectory) = "" Then
       
             MkDir dir_trabajo_Inf & "\" & "GuiaLogistico"
          
          End If
          '-------> Fin crear directorio guias Logistico
           
           NombreArchivoExcel = "GuiaCDLogists_" & Trim(LimpiaDato(fpText1(0).text)) & "_" & Format(Date, "yyyymmdd") & "_" & Format(Time, "HHMMSS")
           Generar_ArchivoExcel "sgp_Sel_ImportacionGuiaCdxUno '" & Trim(LimpiaDato(fpText1(0).text)) & "', 'TR', " & Val(fpLongInteger1(0).text) & ", " & Val(fg_codigocbo(Combo1, 1, 10, "")) & ", '" & Format(fpDateTime1(0).text, "yyyymmdd") & "'", dir_trabajo_Inf & "GuiaLogistico\", NombreArchivoExcel
      
        Else
           
           I_Traspaso Me
           
        End If
    
    Case 15 '------- Salir
        
        Me.Hide
        Unload Me

End Select

Exit Sub
Man_Error:
Toolbar1.Enabled = True
Toolbar2.Enabled = True
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub

If Trans Then

   vg_db.RollbackTrans

End If

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Sub BuscaDoc(codigo As Long)

On Error GoTo Man_Error

Dim RS1    As New ADODB.Recordset
Dim RS2    As New ADODB.Recordset
'Dim RS3 As New ADODB.Recordset
Dim estoc  As Boolean
Dim sql1   As String
Dim sql2   As String
Dim cParam As String
Dim z      As Long
Dim j      As Long
Dim codaux As Long
Dim i      As Long

vg_RDC = ""
estoc = False
'Encabezado

'mod. 20130401 sacar bodega         "AND   tov.tov_numdoc = " & Val(codigo) & " AND tov.tov_codbod = " & vg_codbod & " " & _

'If RS2.State = 1 Then RS2.Close
'RS2.CursorLocation = adUseClient
'vg_db.CursorLocation = adUseClient
'
'RS2.Open "SELECT tov.tov_numinf, tov.tov_totdoc, tov.tov_numdoc, tov.tov_codbod, tov.tov_fecemi, " & _
'         "tov.tov_codser, tov.tov_estdoc, tov.tov_codcas " & _
'         "FROM b_totventas tov, b_clientes cli " & _
'         "WHERE tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(0).text)) & "' " & _
'         "AND   tov.tov_tipdoc = 'TR' " & _
'         "AND   tov.tov_numdoc = " & Val(codigo) & " " & _
'         "AND   tov.tov_rutcli = cli.cli_codigo", vg_db, adOpenStatic
'
'If Not RS2.EOF Then
'
'    vaSpread1.Visible = False
'    Frame1.Enabled = False
'    Frame2.Enabled = False
'    vaSpread1.MaxRows = 0
'
'    Do While Not RS2.EOF
'
'        est = True
'        fpLongInteger1(0).text = RS2!tov_numdoc
'        fpLongInteger1(1).text = RS2!tov_numinf
'        Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 10, fg_pone_cero(Str(RS2!tov_codbod), 10))
'        fpDateTime1(0).text = RS2!tov_fecemi
'        Option1(RS2!tov_codser) = True
'        Label1.Caption = IIf(RS2!tov_estdoc = "", "", "ANULADA")
'        fpText1(1).text = RS2!tov_codcas
'
'        '-------> Leer casino
'        If RS3.State = 1 Then RS3.Close
'        RS3.CursorLocation = adUseClient
'        vg_db.CursorLocation = adUseClient
'
'        RS3.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText1(1).text)) & "' AND (cli_tipo=2 OR cli_tipo = 0)", vg_db, adOpenStatic
'        If Not RS3.EOF Then
'
'           fpayuda(1).Caption = RS3!cli_nombre
'
'        End If
'        RS3.Close
'        Set RS3 = Nothing
''        fpText1_LostFocus 1
'
'        Label2.Caption = Format(RS2!tov_totdoc, fg_Pict(9, vg_DPr))
'        est = False
'        RS2.MoveNext
'
'    Loop
'    vaSpread1.Col = -1
'    vaSpread1.Row = -1
'    vaSpread1.Lock = True
'
'Else
'
'    RS2.Close
'    Set RS2 = Nothing
'
'    If Trim(fpText1(1).text) = "CD" And Option1(0).Value = False Then
'
'       '-------> Validar si proveedor puede ingresar documento
'       If RS2.State = 1 Then RS2.Close
'       RS2.CursorLocation = adUseClient
'       vg_db.CursorLocation = adUseClient
'
'       Set RS2 = vg_db.Execute("select isnull(prv_permiteingdoc,0) AS prv_permiteingdoc FROM b_proveedor where prv_codigo = '" & Trim(fpText1(1).text) & "'")
'       If Not RS2.EOF Then
'
'          If RS2(0) = False Then
'
'             RS2.Close
'             Set RS2 = Nothing
'             MsgBox "Proveedor esta bloqueado para el ingreso documento...", vbCritical, MsgTitulo
'             Limpia 2
'             Exit Sub
'
'          End If
'
'       End If
'
'       RS2.Close
'       Set RS2 = Nothing
'       Exit Sub
'
'    End If
'
'    Gl_Ac_Botones Me, 4, 6, ""
'    fpLongInteger1(0).Enabled = False
'
'    Exit Sub
'
'End If
'RS2.Close
'Set RS2 = Nothing
'
''-------> Detalle
'If RS2.State = 1 Then RS2.Close
'RS2.CursorLocation = adUseClient
'vg_db.CursorLocation = adUseClient
'
'RS2.Open "SELECT DISTINCT ocr_rutpro FROM b_ocsacrecibido WHERE ocr_rutpro = '" & LimpiaDato(Trim(fpText1(0).text)) & "' AND ocr_tipdoc = 'TR' AND ocr_numdoc = " & Val(codigo) & "", vg_db, adOpenStatic
'If Not RS2.EOF Then
'
'   If RS1.State = 1 Then RS1.Close
'   RS1.CursorLocation = adUseClient
'   vg_db.CursorLocation = adUseClient
'
'   sql1 = "SELECT DISTINCT dev.dev_codmer, dev.dev_canmin, dev.dev_canmer, dev.dev_predoc, " & _
'          "dev.dev_ptotal, dev.dev_descri, dev.dev_numlin, uni.uni_nombre, " & _
'          "(SELECT ocr_canoc FROM b_ocsacrecibido WHERE ocr_rutpro = dev.dev_rutcli AND ocr_tipdoc = dev.dev_tipdoc AND ocr_numdoc = dev.dev_numdoc AND ocr_numlin = dev.dev_numlin AND ocr_codprodsgp = dev.dev_codmer) AS ocr_canoc, " & _
'          "(SELECT ocr_preoc FROM b_ocsacrecibido WHERE ocr_rutpro = dev.dev_rutcli AND ocr_tipdoc = dev.dev_tipdoc AND ocr_numdoc = dev.dev_numdoc AND ocr_numlin = dev.dev_numlin AND ocr_codprodsgp = dev.dev_codmer) AS ocr_preoc, " & _
'          "(SELECT ocr_fecoc FROM b_ocsacrecibido WHERE ocr_rutpro = dev.dev_rutcli AND ocr_tipdoc = dev.dev_tipdoc AND ocr_numdoc = dev.dev_numdoc AND ocr_numlin = dev.dev_numlin AND ocr_codprodsgp = dev.dev_codmer) AS ocr_fecoc, " & _
'          "(SELECT ocr_codprodsac FROM b_ocsacrecibido WHERE ocr_rutpro = dev.dev_rutcli AND ocr_tipdoc = dev.dev_tipdoc AND ocr_numdoc = dev.dev_numdoc AND ocr_numlin = dev.dev_numlin AND ocr_codprodsgp = dev.dev_codmer) AS ocr_codprodsac, " & _
'          "(SELECT DISTINCT e.foc_nomsac FROM b_formatocompras e, b_ocsacrecibido WHERE ocr_rutpro = dev.dev_rutcli AND ocr_tipdoc = dev.dev_tipdoc AND ocr_numdoc = dev.dev_numdoc AND ocr_numlin = dev.dev_numlin AND ocr_codprodsgp = dev.dev_codmer AND ocr_codprodsac = e.foc_codsac) AS foc_nomsac " & _
'          "FROM b_detventas dev, b_productos pro, a_unidad uni " & _
'          "WHERE dev.dev_rutcli   = '" & LimpiaDato(Trim(fpText1(0).text)) & "' " & _
'          "AND   dev.dev_tipdoc   = 'TR' " & _
'          "AND   dev.dev_numdoc   = " & Val(codigo) & " " & _
'          "AND   dev.dev_codmer   = pro.pro_codigo " & _
'          "AND   pro.pro_coduni   = uni.uni_codigo ORDER BY dev.dev_numlin"
'   RS1.Open sql1, vg_db, adOpenStatic
'   estoc = True
'
'Else
'
'   If RS1.State = 1 Then RS1.Close
'   RS1.CursorLocation = adUseClient
'   vg_db.CursorLocation = adUseClient
'
'   RS1.Open "SELECT dev.dev_codmer, dev.dev_canmin, dev.dev_canmer, dev.dev_predoc, " & _
'            "dev.dev_ptotal, dev.dev_descri, uni.uni_nombre " & _
'            "FROM b_detventas dev, b_productos pro ,a_unidad uni " & _
'            "WHERE dev.dev_rutcli = '" & LimpiaDato(Trim(fpText1(0).text)) & "'" & _
'            "AND   dev.dev_tipdoc = 'TR' " & _
'            "AND   dev.dev_numdoc = " & Val(codigo) & " " & _
'            "AND   dev.dev_codmer = pro.pro_codigo " & _
'            "AND   pro.pro_coduni = uni.uni_codigo ORDER BY dev.dev_numlin", vg_db, adOpenStatic
'   estoc = False
'
'End If
'
'RS2.Close
'Set RS2 = Nothing

Dim vMotivo() As Variant
Dim lisnom    As String
Dim liscod    As String

i = 1
  
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
      
'-------> Cargar vector
Set RS1 = vg_db.Execute("sgp_sel_motivoGuiaCD ")
      
If Not RS1.EOF Then
      
   ReDim vMotivo(RS1.RecordCount, 2)
         
   Do While Not RS1.EOF
         
      vMotivo(i, 1) = RS1![IdMotivo]
      vMotivo(i, 2) = RS1![Descripcion Motivo]
            
      i = i + 1
        
      RS1.MoveNext
            
   Loop
      
End If
RS1.Close
Set RS1 = Nothing

Dim ExisteEncabezado As Boolean

ExisteEncabezado = False

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgp_Sel_EncabezadoTrapaso '" & LimpiaDato(Trim(fpText1(0).text)) & "', " & Val(codigo) & "")

If Not RS1.EOF Then

    ExisteEncabezado = True
    vaSpread1.Visible = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    vaSpread1.MaxRows = 0

    est = True
    fpLongInteger1(0).text = RS1!tov_numdoc
    fpLongInteger1(1).text = RS1!tov_numinf
    Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 10, fg_pone_cero(Str(RS1!tov_codbod), 10))
    fpDateTime1(0).text = RS1!tov_fecemi
    Option1(RS1!tov_codser) = True
    Label1.Caption = IIf(RS1!tov_estdoc = "", "", "ANULADA")
    fpText1(1).text = RS1!tov_codcas

End If
RS1.Close
Set RS1 = Nothing

vg_GuiaCD = ""
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgp_Sel_MostrarTrapasoEntradaGrilla '" & LimpiaDato(Trim(fpText1(0).text)) & "', " & Val(codigo) & "")
If Not RS1.EOF Then
    
   ExisteEncabezado = True
   vaSpread1.Visible = False
   Frame1.Enabled = False
   Frame2.Enabled = False
   vaSpread1.MaxRows = 0
    
   est = True
   fpLongInteger1(0).text = RS1!tov_numdoc
   fpLongInteger1(1).text = RS1!tov_numinf
   Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 10, fg_pone_cero(Str(RS1!tov_codbod), 10))
   fpDateTime1(0).text = RS1!tov_fecemi
   Option1(RS1!tov_codser) = True
   Label1.Caption = IIf(RS1!tov_estdoc = "", "", "ANULADA")
   fpText1(1).text = RS1!tov_codcas
   fpayuda(1).Caption = RS1!cli_nombre
   Label2.Caption = Format(RS1!tov_totdoc, fg_Pict(9, vg_DPr))
   FDCLogistico.text = Format(RS1!tov_costologistico, fg_Pict(9, vg_DPr))
   vg_GuiaCD = IIf(IsNull(RS1!tov_origen), "", IIf(Trim(RS1!tov_origen) = "KeyLogistic", "1", ""))
   
   est = False
    
   vaSpread1.Col = -1
   vaSpread1.Row = -1
   vaSpread1.Lock = True
    
   If Trim(RS1!ocr_rutpro) <> "" Then
   
      estoc = True
   
   Else
   
      estoc = False
      
   End If
   i = 1
    
    Do While Not RS1.EOF
       
       vaSpread1.MaxRows = i
       vaSpread1.Row = i
       
       vaSpread1.Col = 1
       vaSpread1.text = RS1!dev_codmer
       
       vaSpread1.Col = 2
       vaSpread1.text = RS1!dev_descri
       
       vaSpread1.Col = 3
       vaSpread1.text = RS1!uni_nombre
       
       vaSpread1.Col = 4
       vaSpread1.text = IIf(Option1(1).Value = True, RS1!dev_canmin, RS1!dev_canmer)
       
       vaSpread1.Col = 5
       vaSpread1.text = RS1!dev_predoc
       
       vaSpread1.Col = 6
       vaSpread1.text = RS1!dev_ptotal
       
       vaSpread1.Col = 7
       vaSpread1.text = IIf(Option1(1).Value = True, RS1!dev_canmer, 0)
       
       If estoc Then
          
          vaSpread1.Col = 13
          vaSpread1.text = Format(IIf(IsNull(RS1!ocr_canoc), 0, RS1!ocr_canoc), fg_Pict(9, vg_DCa))
          
          vaSpread1.Col = 14
          vaSpread1.text = Format(IIf(IsNull(RS1!ocr_preoc), 0, RS1!ocr_preoc), fg_Pict(9, 2))
          
          vaSpread1.Col = 15
          If IsNull(RS1!ocr_fecoc) Then
             
             vaSpread1.text = ""
          
          Else
             
             vaSpread1.text = Format(IIf(IsNull(RS1!ocr_fecoc), 0, RS1!ocr_fecoc), "dd/mm/yyyy")
          
          End If
          
          vaSpread1.Col = 16
          vaSpread1.text = IIf(IsNull(RS1!ocr_codprodsac), "", RS1!ocr_codprodsac)
          
          vaSpread1.Col = 17
          vaSpread1.text = IIf(IsNull(RS1!fcs_DenMaterial), "", Trim(RS1!fcs_DenMaterial))
          
          If RS1!tov_origen = "KeyLogistic" Then
          
             lisnom = ""
             liscod = ""
             cParam = ""
             encuentra = False
      
             For j = 1 To UBound(vMotivo)
          
                 If vMotivo(j, 1) <> "" Then
             
                    lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vMotivo(j, 2))
                    liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vMotivo(j, 1)
          
                 End If
      
             Next j
      
             vaSpread1.Col = 19
             vaSpread1.TypeComboBoxList = lisnom
      
             vaSpread1.Col = 20
             vaSpread1.TypeComboBoxList = liscod
    
             vaSpread1.Col = 20
             codaux = -1
      
             For z = 0 To vaSpread1.TypeComboBoxCount
          
                 vaSpread1.TypeComboBoxCurSel = z
                 If vaSpread1.text = RS1!IdMotivo Then
                 
                    codaux = z
                    Exit For
                 
                 End If
                 codaux = -1
      
             Next z
      
             vaSpread1.Col = 19
             vaSpread1.TypeComboBoxCurSel = codaux
          
          End If
          
          If vaSpread1.Row = 1 And estoc Then
             
             vaSpread1.Row = 1
             vaSpread1.Col = 16
             Text2(0).text = Trim(vaSpread1.text)
             
             vaSpread1.Row = 1
             vaSpread1.Col = 17
             Text2(1).text = Trim(vaSpread1.text)
          
          End If
       
       End If
       
'       '-------> Trae Stock
'       If RS2.State = 1 Then RS2.Close
'       RS2.CursorLocation = adUseClient
'       vg_db.CursorLocation = adUseClient
'       RS2.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
'                "WHERE bod.bod_codpro = pro.pro_codigo " & _
'                "AND   bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " AND pro.pro_codigo = '" & Trim(RS1!dev_codmer) & "' AND pro.pro_ctrsto = 1", vg_db, adOpenStatic
       vaSpread1.Col = 9
'       If Not RS2.EOF Then
       vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa))
'       RS2.Close
'       Set RS2 = Nothing
       
       RS1.MoveNext
       i = i + 1
    
    Loop
    
    vaSpread1.Visible = True

ElseIf Not ExisteEncabezado Then

   RS1.Close
   Set RS1 = Nothing

   If Trim(fpText1(1).text) = "CD" And Option1(0).Value = False Then

      '-------> Validar si proveedor puede ingresar documento
      If RS2.State = 1 Then RS2.Close
      RS2.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient

      Set RS2 = vg_db.Execute("select isnull(prv_permiteingdoc,0) AS prv_permiteingdoc FROM b_proveedor where prv_codigo = '" & Trim(fpText1(1).text) & "'")
      If Not RS2.EOF Then

         If RS2(0) = False Then

            RS2.Close
            Set RS2 = Nothing
            MsgBox "Proveedor esta bloqueado para el ingreso documento...", vbCritical, MsgTitulo
            Limpia 2
            Exit Sub

         End If

      End If

      RS2.Close
      Set RS2 = Nothing
      Exit Sub

   End If

   Gl_Ac_Botones Me, 4, 6, ""
   fpLongInteger1(0).Enabled = False

   Exit Sub

End If
RS1.Close
Set RS1 = Nothing

If Option1(0).Value = True Then
    
    vaSpread1.Visible = False
    
    vaSpread1.Col = 7
    vaSpread1.Row = -1
    vaSpread1.ColHidden = True
    
    vaSpread1.Col = 5
    vaSpread1.Row = 0
    vaSpread1.text = "P.M.P."
    vaSpread1.ColWidth(2) = 87.88 '57.55  '33.25
    
    vaSpread1.Col = 13
    vaSpread1.ColHidden = True
    
    vaSpread1.Col = 14
    vaSpread1.ColHidden = True
    
    vaSpread1.Col = 15
    vaSpread1.ColHidden = True
    vaSpread1.Visible = True

    vaSpread1.Col = 16
    vaSpread1.ColHidden = True
    vaSpread1.Visible = True
    
    vaSpread1.Col = 19
    vaSpread1.ColHidden = True

Else
   
   If estoc Then
      
      vaSpread1.ColWidth(2) = 39.88 '47.88
'      vaSpread1.ColWidth(2) = 77.88 '23.5

      vaSpread1.Col = 13
      vaSpread1.ColHidden = True
      
      vaSpread1.Col = 14
      vaSpread1.ColHidden = True
      
      vaSpread1.Col = 15
      vaSpread1.ColHidden = True
   
      vaSpread1.Col = 16
      vaSpread1.ColHidden = False
      
      vaSpread1.Col = 19
      vaSpread1.ColHidden = False
   
      Text2(0).Visible = True
      Text2(1).Visible = True
   
   Else
      
      vaSpread1.ColWidth(2) = 77.88 '47.88
      vaSpread1.Col = 13
      vaSpread1.ColHidden = True
      
      vaSpread1.Col = 14
      vaSpread1.ColHidden = True
      
      vaSpread1.Col = 15
      vaSpread1.ColHidden = True
   
      vaSpread1.Col = 16
      vaSpread1.ColHidden = True
   
      vaSpread1.Col = 19
      vaSpread1.ColHidden = True
      
      Text2(0).Visible = False
      Text2(1).Visible = False
   
   End If

End If

vg_codigo = ""
Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub Limpia(op As Integer)

On Error GoTo Man_Error

est = True
Toolbar2.Enabled = True
Row_Activo = 0
CodigoProd = ""
vg_FDC = ""
vg_RDC = ""
vg_GuiaCD = ""
vg_FechaEmision_GGD = "0000:00:00"
Label1.Caption = ""
fpText1(1).text = ""
fpayuda(1).Caption = ""
Option1(0).Enabled = True
Option1(1).Enabled = True
Combo1(1).ListIndex = IIf(Combo1(1).listcount = 1, 0, -1)
vaSpread1.MaxRows = 0
vaSpread1.MaxRows = 1

vaSpread1.Col = -1
vaSpread1.Row = -1
vaSpread1.Lock = True

vaSpread1.Col = 1
vaSpread1.Row = 1
vaSpread1.Lock = False

fpText1(0).Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText1(0).text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
fpLongInteger1(0).Enabled = True
fpLongInteger1(0).text = 0
fpLongInteger1(1).text = MuestraFolio
FDCLogistico.text = 0
'fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")

fpLongInteger1(0).Enabled = True
fpDateTime1(0).Enabled = True
fpText1(1).Enabled = True

fpDateTime1(0).text = ""
Option1(1).Value = True
Label2.Caption = Format(0, fg_Pict(9, vg_DPr))

Text2(0).Visible = False
Text2(1).Visible = False
Text2(0).text = ""
Text2(1).text = ""

Image1(2).Visible = False
Frame1.Enabled = True
Frame2.Enabled = True
Image1(5).Visible = False

Gl_Ac_Botones Me, 4, op, ""

'Me.HelpContextID = 2041000
'Cmd_ImportarGuiaCD.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
Cmd_ImportarGuiaCD.Enabled = True
Me.HelpContextID = vg_OpcM
est = False

If Option1(1).Value = True Then
    
    vaSpread1.Visible = False
    vaSpread1.ColWidth(2) = 77.88 '47.88
    
    vaSpread1.Col = 13
    vaSpread1.ColHidden = True
    
    vaSpread1.Col = 14
    vaSpread1.ColHidden = True
    
    vaSpread1.Col = 15
    vaSpread1.ColHidden = True
    vaSpread1.Visible = True

    vaSpread1.Col = 16
    vaSpread1.ColHidden = True
    vaSpread1.Visible = True
    
    vaSpread1.Col = 18
    vaSpread1.ColHidden = True
    vaSpread1.Visible = True

    vaSpread1.Col = 19
    vaSpread1.ColHidden = True
    vaSpread1.Visible = True

    vaSpread1.Col = 20
    vaSpread1.ColHidden = True
    vaSpread1.Visible = True

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Function MuestraFolio() As Long

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
MuestraFolio = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
RS.Open "SELECT MAX(inf_numero) AS folio FROM a_infcfcfofi WHERE inf_cencos = '" & Trim(fpText1(0).text) & "' AND inf_tipo = 'T' AND inf_feccie = 0 AND (inf_usuario) IS NULL", vg_db, adOpenStatic
If Not RS.EOF Then MuestraFolio = IIf(TipoDato(RS!Folio, 0) = 0, 1, TipoDato(RS!Folio, 0))
RS.Close
Set RS = Nothing

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Function

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim i As Long, propon As Double, iRow As Long
Dim est15 As Boolean
Dim Fecha As String, sql1 As String

Select Case Button.Index
    
    Case 1
        
        Toolbar2.Enabled = False
        vg_nombre = ""
        vg_codigo = ""
        vg_bodega = 0
        vg_bodega = Val(fg_codigocbo(Combo1, 1, 10, ""))
        vg_left = fpayuda(1).Left + 1920
        
        B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", IIf(Option1(1).Value = True, "ProVig", "ProInvNoStock")
        B_TabEst.Show 1
        If vg_codigo = "" Then
        
           Toolbar2.Enabled = True
           Exit Sub
        
        End If
        
        Toolbar2.Enabled = True
        iRow = vaSpread1.MaxRows
        
        If vg_pais <> "CL" Or vg_FDC <> "OC" Then
           
           For i = 1 To vaSpread1.MaxRows
               
               vaSpread1.Col = 1
               vaSpread1.Row = i
               If Trim(vaSpread1.text) = Trim(vg_codigo) Then
               
                  MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo
                  Exit Sub
           
               End If
               
           Next i
        
        End If
        
        vaSpread1.Row = IIf(vaSpread1.MaxRows = vaSpread1.ActiveRow, vaSpread1.MaxRows, vaSpread1.ActiveRow)
'        sql1 = IIf(vg_tipbase = "1", " AND CDATE(x.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), x.foc_vigfin,101) >  '" & Date & "'")
'

'        RS1.Open "SELECT a.pro_codigo, a.pro_nombre, b.uni_nombre, a.pro_ctrsto, " & _
'                 "(SELECT TOP 1 x.foc_codsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_codsac, " & _
'                 "(SELECT TOP 1 x.foc_nomsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_nomsac, " & _
'                 "(SELECT TOP 1 x.foc_unisac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_unisac, " & _
'                 "(SELECT TOP 1 x.foc_faccon FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_faccon " & _
'                 "FROM b_productos a, a_unidad b " & _
'                 "WHERE a.pro_coduni = b.uni_codigo " & _
'                 "AND   a.pro_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic

        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        Set RS1 = vg_db.Execute("sgp_Sel_ProductoTraspasoSalEnt '" & MuestraCasino(1) & "', '" & vg_codigo & "'")
        If Not RS1.EOF Then
            
            Do While Not RS1.EOF
                
                vaSpread1.Col = 1
                If vaSpread1.Lock = True Then
                
                   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
                
                End If
                
                vaSpread1.Row = vaSpread1.MaxRows
                
                For i = 4 To IIf(Option1(0).Value = True, 5, 7)
                    
                    vaSpread1.Col = i
                    vaSpread1.Lock = IIf(i <> 6 And Option1(1).Value = True, False, IIf(i <> 5 And Option1(0).Value = True, False, True))
                
                Next i
                
                vaSpread1.Col = 1
                vaSpread1.text = RS1!pro_codigo
                
                vaSpread1.Col = 2
                vaSpread1.text = RS1!pro_nombre
                
                vaSpread1.Col = 3
                vaSpread1.text = RS1!uni_nombre
                
                vaSpread1.Col = 4
                If vg_GuiaCD = "1" Then
                
                   vaSpread1.Lock = True
                
                ElseIf Trim(vg_GuiaCD) = "" Then
                   
                   vaSpread1.Lock = False
                
                End If
                
                If Trim(vaSpread1.text) = "" Then vaSpread1.text = 0
                
'                '-------> Trae pmp
'                If RS2.State = 1 Then RS2.Close
'                RS2.CursorLocation = adUseClient
'                vg_db.CursorLocation = adUseClient
'
'                RS2.Open "SELECT TOP 1 ppd_cencos, ppd_codpro, ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
'                         "FROM b_productospmpdia " & _
'                         "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
'                         "AND   ppd_codpro = '" & RS1!pro_codigo & "' " & _
'                         "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
'                         "GROUP BY ppd_cencos, ppd_codpro, ppd_propon " & _
'                         "HAVING (ppd_propon) > 0 ORDER BY Max(ppd_fecdia) DESC", vg_db, adOpenStatic
'                If Not RS2.EOF Then propon = RS2!ppd_propon
'                RS2.Close
'                Set RS2 = Nothing
                
                propon = RS1!ppd_propon
                vaSpread1.Col = 5
                vaSpread1.text = IIf(Option1(1).Value = True, 0, propon)
                
                vaSpread1.Col = 6
                If Trim(vaSpread1.text) = "" Then vaSpread1.text = 0
                
                vaSpread1.Col = 7
                If Trim(vaSpread1.text) = "" Then vaSpread1.text = 0
                
                vaSpread1.Col = 8
                vaSpread1.text = "N" 'No bloquedo
                
                vaSpread1.Col = 10
                vaSpread1.text = IIf(IsNull(RS1!pro_ctrsto), "N", IIf(RS1!pro_ctrsto = 1, "S", "N"))
                
                vaSpread1.Col = 11
                vaSpread1.text = propon
                
                vaSpread1.Col = 12
                vaSpread1.text = "N"
                
                vaSpread1.Col = 16
                vaSpread1.text = IIf(IsNull(RS1!fcs_CodMaterial), "", RS1!fcs_CodMaterial)
                
                vaSpread1.Col = 17
                vaSpread1.text = IIf(IsNull(RS1!fcs_DenMaterial), "", RS1!fcs_DenMaterial)
                
                vaSpread1.Col = 16
                Text2(0).text = Trim(vaSpread1.text)
                
                vaSpread1.Col = 17
                Text2(1).text = Trim(vaSpread1.text)
                
                vaSpread1.Col = 4
                If Trim(vaSpread1.text) <> "" Then
                
                   canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
                   
                   vaSpread1.Col = 7
                   vaSpread1.text = Format(canrea, fg_Pict(9, vg_DCa))
                
                End If
                
                vaSpread1.Col = 5
                If Trim(vaSpread1.text) <> "" Then
                
                   propon = Format(vaSpread1.text, fg_Pict(9, 2))
                
                End If
                
                vaSpread1.Col = 6
                If Trim(vaSpread1.text) <> "" Then
                
                   vaSpread1.text = Format(canrea * propon, fg_Pict(9, 2))
                
                End If
                
'                '-------> Trae Stock
'                If RS2.State = 1 Then RS2.Close
'                RS2.CursorLocation = adUseClient
'                vg_db.CursorLocation = adUseClient
'
'                RS2.Open "SELECT b.bod_canmer FROM b_productos a, b_bodegas b " & _
'                         "WHERE b.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
'                         "AND   b.bod_codpro = a.pro_codigo AND a.pro_codigo = '" & Trim(RS1!pro_codigo) & "'", vg_db, adOpenStatic
'                vaSpread1.Col = 9
'                If Not RS2.EOF Then vaSpread1.text = Format(RS2!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
'                RS2.Close
'                Set RS2 = Nothing
                
                vaSpread1.Col = 9
                vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa))
                
                RS1.MoveNext
                
                i = i + 1
                
            Loop
            
        End If
        RS1.Close
        Set RS1 = Nothing
        If vaSpread1.MaxRows = 1 Then Gl_Ac_Botones Me, 4, 6, ""
        vaSpread1.Col = 4
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.SetActiveCell 4, vaSpread1.MaxRows
        
        If Option1(1) = True Then
            
            vaSpread1.Col = 5
            vaSpread1.Row = 0
            vaSpread1.text = "Precio Documento"
        
        Else
            
            vaSpread1.Col = 5
            vaSpread1.Row = 0
            vaSpread1.text = "P.M.P."
        
        End If
        If vaSpread1.Enabled = True Then vaSpread1.SetFocus
    
    Case 2
        
        '-------> Validar exportación Guia CD
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 21
        If vaSpread1.text = "GuiaCD" Then
        
           MsgBox "No permite eliminar producto exportado desde guía CD. Si no desea utilizar el producto puede dejar la columna Cantidad Recibida con valor cero...", vbExclamation + vbError, MsgTitulo
           Exit Sub
        
        End If
        
        est15 = False
        If vaSpread1.MaxRows = 0 Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 15
        est15 = vaSpread1.ColHidden
        Fecha = Trim(vaSpread1.text)
        vaSpread1.Col = 1
        If vaSpread1.Lock = False Or Trim(Fecha) <> "" Then Exit Sub
        If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        If vaSpread1.MaxRows = 0 Then
        
           vaSpread1.MaxRows = vaSpread1.MaxRows + 1
           vaSpread1.Row = vaSpread1.MaxRows
           vaSpread1.Col = 1
           vaSpread1.Lock = False
           
        End If
        
        If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
        
        '------- Total General ---------
        Dim subtot As Double
        subtot = 0
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            If Trim(vaSpread1.text) <> "" Then
               
               vaSpread1.Col = 6
               subtot = subtot + Format(vaSpread1.text, fg_Pict(9, vg_DPr))
        
            End If
            
        Next
        Label2.Caption = Format(subtot, fg_Pict(9, vg_DPr))
        '-------------------------------
        
        If vaSpread1.MaxRows = 0 Then Gl_Ac_Botones Me, 4, 6, ""
        If vaSpread1.Enabled = True Then vaSpread1.SetFocus
        
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim indice As Long

Select Case Col

Case 19
    
    vaSpread1.Row = Row
    
    vaSpread1.Col = 19
    indice = vaSpread1.TypeComboBoxCurSel
    
    vaSpread1.Col = 20
    vaSpread1.TypeComboBoxCurSel = indice

    Toolbar2.Enabled = True
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub

Select Case KeyCode

Case 46
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    
    If vaSpread1.Col <> 19 Then Exit Sub
    vaSpread1.text = ""
    vaSpread1.TypeComboBoxCurSel = -1

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_Advance(ByVal AdvanceNext As Boolean)

On Error GoTo Man_Error

Dim codigo As String
Dim Nombre As String
Dim i As Long

If vaSpread1.MaxRows < 1 Or Frame1.Enabled = False Then

   Exit Sub
   
End If

Dim RS1       As New ADODB.Recordset
Dim canrea    As Double
Dim canrec    As Double
Dim candoc    As Double
Dim EstGuiaCd As Boolean
Dim Motivo    As String

vaSpread1.Row = vaSpread1.ActiveRow

vaSpread1.Col = 4
candoc = IIf(Trim(vaSpread1.text) = 0, 0, Val(vaSpread1.text))

vaSpread1.Col = 21
EstGuiaCd = IIf(Trim(vaSpread1.text) = "", False, True)

EstGuiaCd = IIf(Trim(vg_GuiaCD) = "", False, True)

If EstGuiaCd And candoc > 0 Then

   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = 4
        
   If Trim(vaSpread1.text) <> "" Then
        
      canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
        
   End If
        
   vaSpread1.Col = 7
   If Trim(vaSpread1.text) <> "" Then
        
      canrec = Format(vaSpread1.Value, fg_Pict(9, vg_DCa))
        
   End If
        
   vaSpread1.Col = 19
   Motivo = Trim(vaSpread1.text)
    
   vaSpread1.Col = 7
   
   If canrec > canrea And Trim(Motivo) = "" Then
    
      Toolbar2.Enabled = False
      MsgBox "La cantidad recibida excede de la cantidad es menor...", vbCritical, MsgTitulo
           
   End If
        
   If canrec = canrea Then
        
      vaSpread1.Col = 19
      vaSpread1.Lock = True
      vaSpread1.text = ""
      
      Toolbar2.Enabled = True
        
   End If
        
   vaSpread1.Col = 19
        
   If canrec <> canrea And Trim(vaSpread1.text) = "" Then
            
      Toolbar2.Enabled = False
   
      vaSpread1.Col = 19
      vaSpread1.Lock = False
      vaSpread1.SetActiveCell 19, vaSpread1.Row
      vaSpread1.SetFocus
            
         '   MsgBox "La cantidad recibida es distinta a cantidad documento, debera seleccionar la columna Descripción Motivo ...", vbCritical + vbOKOnly, MsgTitulo
        
      vaSpread1.Col = 19
      Exit Sub
            'vaSpread1.SetActiveCell 19, vaSpread1.Row
            'vaSpread1.SetFocus
        
   End If

End If


vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = Trim(vaSpread1.text)

vaSpread1.Col = 2
Nombre = Trim(vaSpread1.text)

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If AdvanceNext = True And codigo <> "" And Nombre <> "" Then
   
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = 1
   
'   If Option1(1).Value = True Then
'
'      RS1.Open "SELECT a.pro_codigo, c.ppd_propon, a.pro_nombre, b.uni_nombre, a.pro_ctrsto " & _
'               "FROM b_productos a, a_unidad b, b_productospmpdia c " & _
'               "WHERE a.pro_codigo = c.ppd_codpro AND a.pro_coduni = b.uni_codigo AND a.pro_codigo = '" & LimpiaDato(Trim(vaSpread1.text)) & "' " & _
'               "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven = 0) AND c.ppd_cencos = '" & MuestraCasino(1) & "' AND c.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & "", vg_db, adOpenStatic
'   Else
'
'      If Trim(fpDateTime1(0).text) = "" Then MsgBox "Debe ingresar fecha emisión...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
'
'      RS1.Open "SELECT a.pro_codigo, c.ppd_propon, a.pro_nombre, b.uni_nombre, a.pro_ctrsto " & _
'               "FROM b_productos a, a_unidad b, b_productospmpdia c " & _
'               "WHERE a.pro_codigo = c.ppd_codpro AND a.pro_coduni = b.uni_codigo AND a.pro_codigo = '" & LimpiaDato(Trim(vaSpread1.text)) & "' " & _
'               "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven = 0) AND c.ppd_cencos = '" & MuestraCasino(1) & "' AND c.ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND c.ppd_fecdia <= " & Format(CDate(fpDateTime1(0).text), "yyyymmdd") & " " & _
'               "UNION SELECT DISTINCT pro_codigo, 0 AS ppd_propon, pro_nombre, '' AS uni_nombre, 0 AS pro_ctrsto FROM b_productos pro, b_bodegas bod WHERE pro.pro_codigo = bod.bod_codpro AND bod.bod_codbod = " & vg_codbod & " AND bod.bod_canmer > 0 AND pro.pro_codigo = '" & LimpiaDato(Trim(vaSpread1.text)) & "'", vg_db, adOpenStatic
'
'   End If
   
   Set RS1 = vg_db.Execute("sgp_Sel_ProductoTraspasoSalEnt '" & MuestraCasino(1) & "', '" & LimpiaDato(Trim(vaSpread1.text)) & "'")
   If RS1.EOF Then
      
      RS1.Close
      Set RS1 = Nothing
      vaSpread1.Row = vaSpread1.ActiveRow
      For i = 4 To 7
      
         vaSpread1.Col = i
         vaSpread1.Lock = True
         
      Next i
      
      vaSpread1.Row = vaSpread1.ActiveRow
      vaSpread1.Col = 1
      vaSpread1.text = ""
      
      vaSpread1.Col = 2
      vaSpread1.text = ""
      
      vaSpread1.Col = 3
      vaSpread1.text = ""
      
      MsgBox "producto no existe...", vbExclamation + vbOKOnly, MsgTitulo
      vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow
      Exit Sub
   
   End If
   
   RS1.Close
   Set RS1 = Nothing
   
   vaSpread1.Col = 1
   vaSpread1.Lock = True
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.SetActiveCell 1, vaSpread1.MaxRows
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   vaSpread1.Lock = False
   
   For i = 4 To 7
   
       vaSpread1.Col = i
       vaSpread1.Lock = True
       
   Next i

   If Trim(vg_GuiaCD) = "1" Then
   
       vaSpread1.Col = 19
       vaSpread1.Lock = True
   
   End If
   
End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

'On Error GoTo Man_Error

Dim codsgp As String, codsac As String
Dim i As Long

If Trim(CodigoProd) <> "" Then
           For i = 1 To vaSpread1.MaxRows

               vaSpread1.Col = 1
               vaSpread1.Row = i
               If Trim(vaSpread1.text) = Trim(CodigoProd) And Row_Activo <> i And Trim(vaSpread1.text) <> "" Then

                  MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo

                  vaSpread1.Row = Row_Activo
                  vaSpread1.Col = 1
                  vaSpread1.text = ""

                  vaSpread1.Col = 2
                  vaSpread1.text = ""

                  vaSpread1.Col = 3
                  vaSpread1.text = ""

                  vaSpread1.Col = 16
                  vaSpread1.text = ""

                  vaSpread1.Col = 17
                  vaSpread1.text = ""
                  
                  CodigoProd = ""

                  vaSpread1.SetActiveCell 1, Row_Activo
                  Exit Sub

               End If

           Next i
End If

vaSpread1.Row = Row

vaSpread1.Col = Col
vaSpread1.Row = vaSpread1.ActiveRow

vaSpread1.Col = 1
codsgp = Trim(vaSpread1.text)

vaSpread1.Col = 16
Text2(0).text = Trim(vaSpread1.text)

vaSpread1.Col = 16
codsac = Trim(vaSpread1.text)

vaSpread1.Col = 17
Text2(1).text = Trim(vaSpread1.text)

If vg_pais = "CL" And vg_FDC = "OC" And Trim(codsgp) <> "" And Option1(1).Value = True Then
   
   Image1(5).Visible = IIf(ValidarProductosSgpSac(Trim(codsac), Trim(codsgp)), True, False)

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim fecoc  As String
Dim codsac As String
Dim cPro   As String
Dim sql1   As String
Dim i      As Long
Dim propon As Double
Dim RS1    As New ADODB.Recordset
Dim RS2    As New ADODB.Recordset

vaSpread1.Row = Row
vaSpread1.Col = 15
fecoc = ""
fecoc = Trim(vaSpread1.text)

vaSpread1.Col = 16
codsac = ""
codsac = Trim(vaSpread1.text)

If vaSpread1.Lock = True Or fecoc <> "" Then
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
'   RS1.Open "SELECT COUNT(b.fcs_codsgp) AS nreg " & _
'             "FROM  b_formatocompras a, b_formatocomprassgp b " & _
'             "WHERE a.foc_codsac = b.fcs_codsac " & _
'             "AND   b.fcs_codsac = '" & codsac & "'", vg_db, adOpenStatic
   Set RS1 = vg_db.Execute("sgp_Sel_FormatoComprasSapTraspasoSalEnt '" & codsac & "'")
   If Not RS1.EOF And Not IsNull(RS1!nreg) And RS1!nreg > 1 Then
      
      RS1.Close
      Set RS1 = Nothing
      
      vg_nombre = ""
      vg_codigo = ""
      vg_codigo = codsac
      vg_left = fpayuda(0).Left + 1920
      B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "ProVigSac"
      B_TabEst.Show 1
      If vg_codigo = "" Then Exit Sub
      
      If vg_pais <> "CL" Or vg_FDC <> "OC" Then
         
         For i = 1 To vaSpread1.MaxRows
             
             vaSpread1.Col = 1
             vaSpread1.Row = i
             If Trim(vaSpread1.text) = Trim(vg_codigo) Then
             
                Frame5.Enabled = False
                MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo
                vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow
                Frame5.Enabled = True
                Exit Sub
             
             End If
             
         Next i
      
      End If
      sql1 = IIf(vg_tipbase = "1", " AND CDATE(x.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), x.foc_vigfin,101) >  '" & Date & "'")
      
      If RS1.State = 1 Then RS1.Close
      RS1.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
'      RS1.Open "SELECT a.pro_codigo, a.pro_nombre, a.pro_ctacon, a.pro_ctrsto, b.uni_nomcor, a.pro_facing, a.pro_facsto, c.ppd_propon, " & _
'               "(SELECT TOP 1 x.foc_codsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_codsac, " & _
'               "(SELECT TOP 1 x.foc_nomsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_nomsac, " & _
'               "(SELECT TOP 1 x.foc_unisac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_unisac, " & _
'               "(SELECT TOP 1 x.foc_faccon FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_faccon " & _
'               "FROM  b_productos a, a_unidad b, b_productospmpdia c " & _
'               "WHERE a.pro_codigo = c.ppd_codpro AND a.pro_coduni = b.uni_codigo " & _
'               "AND   a.pro_codigo = '" & vg_codigo & "' AND c.ppd_cencos = '" & MuestraCasino(1) & "' AND c.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & "", vg_db, adOpenStatic
      Set RS1 = vg_db.Execute("sgp_Sel_ProductoTraspasoSalEnt '" & MuestraCasino(1) & "', '" & vg_codigo & "'")
      If Not RS1.EOF Then
         
         If IsNull(RS1!pro_ctrsto) Then Frame5.Enabled = False: MsgBox "Producto no tiene asignado, el Movimiento...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Frame5.Enabled = True: Exit Sub
         If RS1!pro_facing = 0 Or RS1!pro_facsto = 0 Then RS1.Close: Set RS1 = Nothing: MsgBox "Factor del producto en cero...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.acitverow: Exit Sub
         
         cPro = RS1!pro_codigo
         vaSpread1.Row = Row
         vaSpread1.Col = 1
         vaSpread1.Row = Row
         vaSpread1.SetActiveCell 1, Row
         
         '-------> desbloquear celda
         vaSpread1.Col = 1
         vaSpread1.text = RS1!pro_codigo
         
         vaSpread1.Col = 2
         vaSpread1.text = RS1!pro_nombre
         
         vaSpread1.Col = 3
         vaSpread1.text = RS1!uni_nombre
         
         vaSpread1.Col = 4
         If Trim(vaSpread1.text) = "" Then vaSpread1.text = 0
         
'         '-------> Trae pmp
'         If RS2.State = 1 Then RS2.Close
'         RS2.CursorLocation = adUseClient
'         vg_db.CursorLocation = adUseClient
'
'         RS2.Open "SELECT ppd_cencos, ppd_codpro, ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
'                  "FROM b_productospmpdia " & _
'                  "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
'                  "AND   ppd_codpro = '" & RS1!pro_codigo & "' " & _
'                  "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
'                  "GROUP BY ppd_cencos, ppd_codpro, ppd_propon " & _
'                  "HAVING (ppd_propon) > 0", vg_db, adOpenStatic
'         If Not RS2.EOF Then propon = RS2!ppd_propon
'         RS2.Close
'         Set RS2 = Nothing
         
         propon = RS1!ppd_propon
         
         vaSpread1.Col = 5
         vaSpread1.text = IIf(Option1(1).Value = True, 0, propon)
         
         vaSpread1.Col = 6
         If Trim(vaSpread1.text) = "" Then vaSpread1.text = 0
         
         vaSpread1.Col = 7
         If Trim(vaSpread1.text) = "" Then vaSpread1.text = 0
         
         vaSpread1.Col = 8
         vaSpread1.text = "N" 'No bloquedo
         
         vaSpread1.Col = 10
         vaSpread1.text = IIf(IsNull(RS1!pro_ctrsto), "N", IIf(RS1!pro_ctrsto = 1, "S", "N"))
         
         vaSpread1.Col = 11
         vaSpread1.text = propon
         
         vaSpread1.Col = 12
         vaSpread1.text = "N"
         
         vaSpread1.Col = 16
         vaSpread1.text = IIf(IsNull(RS1!fcs_CodMaterial), "", RS1!fcs_CodMaterial)
         
         vaSpread1.Col = 17
         vaSpread1.text = IIf(IsNull(RS1!fcs_DenMaterial), "", RS1!fcs_DenMaterial)
         
         vaSpread1.Col = 16
         Text2(0).text = Trim(vaSpread1.text)
         
         vaSpread1.Col = 17
         Text2(1).text = Trim(vaSpread1.text)
         
'         '-------> Trae Stock
'         If RS2.State = 1 Then RS2.Close
'         RS2.CursorLocation = adUseClient
'         vg_db.CursorLocation = adUseClient
'
'         RS2.Open "SELECT b.bod_canmer FROM b_productos a, b_bodegas b " & _
'                  "WHERE b.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
'                  "AND   b.bod_codpro = a.pro_codigo AND a.pro_codigo = '" & Trim(RS1!pro_codigo) & "'", vg_db, adOpenStatic
'         vaSpread1.Col = 9
'         If Not RS2.EOF Then vaSpread1.text = Format(RS2!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
'         RS2.Close
'         Set RS2 = Nothing
        
        vaSpread1.Col = 9
        vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa))
        
     Else
        
        '-------> bloquear celda
        vaSpread1.Row = vaSpread1.ActiveRow
        
        vaSpread1.Col = 1
        vaSpread1.text = ""
        
        vaSpread1.Col = 2
        vaSpread1.text = ""
        
        vaSpread1.Col = 3
        vaSpread1.text = ""
        
        vaSpread1.Col = 16
        vaSpread1.text = ""
        
        vaSpread1.Col = 17
        vaSpread1.text = ""
        
        vaSpread1.Col = 16
        Text2(0).text = Trim(vaSpread1.text)
        
        vaSpread1.Col = 17
        Text2(1).text = Trim(vaSpread1.text)
        
        For i = 4 To 10
        
            vaSpread1.Col = i
            vaSpread1.Lock = True
        
        Next i
     
     End If
     RS1.Close
     Set RS1 = Nothing
     vaSpread1.Refresh
     
     If vaSpread1.Enabled = True Then
     
        On Error Resume Next
        vaSpread1.SetFocus
   
     End If
     
   Else
      
      RS1.Close
      Set RS1 = Nothing
   
   End If

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim RS1     As New ADODB.Recordset
Dim RS2     As New ADODB.Recordset
Dim canrea  As Double
Dim propon  As Double
Dim subtot  As Double
Dim canbod  As Double
Dim canrec  As Double
Dim i       As Long
Dim codmer  As String
Dim EstGuia As String
vaSpread1.Row = Row

Select Case Col

    Case 1
        
        vaSpread1.Col = 1
        If LimpiaDato(Trim(vaSpread1.text)) <> "" And vaSpread1.Lock = False Then
'           sql1 = IIf(vg_tipbase = "1", " AND CDATE(x.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), x.foc_vigfin,101) >  '" & Date & "'")
'
'           If RS1.State = 1 Then RS1.Close
'           RS1.CursorLocation = adUseClient
'           vg_db.CursorLocation = adUseClient
'           Set RS1 = vg_db.Execute("SELECT a.pro_codigo, a.pro_nombre, b.uni_nombre, a.pro_ctrsto, " & _
'                     "(SELECT TOP 1 x.foc_codsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_codsac, " & _
'                     "(SELECT TOP 1 x.foc_nomsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_nomsac, " & _
'                     "(SELECT TOP 1 x.foc_unisac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_unisac, " & _
'                     "(SELECT TOP 1 x.foc_faccon FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_faccon " & _
'                     "FROM  b_productos a, a_unidad b " & _
'                     "WHERE a.pro_coduni = b.uni_codigo " & _
'                     "AND   a.pro_codigo = '" & LimpiaDato(Trim(vaSpread1.text)) & "'")
           
           If RS1.State = 1 Then RS1.Close
           RS1.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient

           vg_codigo = LimpiaDato(Trim(vaSpread1.text))
           CodigoProd = LimpiaDato(Trim(vaSpread1.text))
           Row_Activo = Row
           Set RS1 = vg_db.Execute("sgp_Sel_ProductoTraspasoSalEnt '" & MuestraCasino(1) & "', '" & vg_codigo & "'")

           If Not RS1.EOF Then
              
              Do While Not RS1.EOF
                 
                 vaSpread1.Col = 1
                 vaSpread1.Row = vaSpread1.MaxRows
                 
                 For i = 4 To IIf(Option1(0).Value = True, 5, 7)
                     
                     vaSpread1.Col = i
                     vaSpread1.Lock = IIf(i <> 6 And Option1(1).Value = True, False, IIf(i <> 5 And Option1(0).Value = True, False, True))
                 
                 Next i
                 
                 vaSpread1.Row = Row
                 vaSpread1.Col = 2
                 vaSpread1.text = RS1!pro_nombre
                 
                 vaSpread1.Col = 3
                 vaSpread1.text = RS1!uni_nombre
                 
                 vaSpread1.Col = 4
                 If vg_GuiaCD = "1" Then
                
                    vaSpread1.Lock = True
                
                 ElseIf Trim(vg_GuiaCD) = "" Then
                   
                    vaSpread1.Lock = False
                
                 End If
                 
                 If Trim(vaSpread1.text) = "" Then
                 
                    vaSpread1.text = 0
                    
                 End If
                 
'                 '-------> Trae pmp
                 propon = 0
                 
'                 If RS2.State = 1 Then RS2.Close
'                 RS2.CursorLocation = adUseClient
'                 vg_db.CursorLocation = adUseClient
'                 Set RS2 = vg_db.Execute("SELECT TOP 1 ppd_cencos, ppd_codpro, isnull(ppd_propon,0), Max(ppd_fecdia) AS ppd_fecdia " & _
'                          "FROM b_productospmpdia " & _
'                          "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
'                          "AND   ppd_codpro = '" & RS1!pro_codigo & "' " & _
'                          "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
'                          "GROUP BY ppd_cencos, ppd_codpro, ppd_propon " & _
'                          "HAVING (ppd_propon) > 0 ORDER BY Max(ppd_fecdia) DESC")
'                 If Not RS2.EOF Then propon = RS2!ppd_propon
'                 RS2.Close
'                 Set RS2 = Nothing
                 
                 propon = RS1!ppd_propon
                 
                 vaSpread1.Col = 5
                 vaSpread1.text = IIf(Option1(1).Value = True, 0, propon)
                 
                 vaSpread1.Col = 6
                 If Trim(vaSpread1.text) = "" Then vaSpread1.text = 0
                 
                 vaSpread1.Col = 7
                 If Trim(vaSpread1.text) = "" Then vaSpread1.text = 0
                 
                 vaSpread1.Col = 8
                 vaSpread1.text = "N" 'No bloquedo
                 
                 vaSpread1.Col = 10
                 vaSpread1.text = IIf(IsNull(RS1!pro_ctrsto), "N", IIf(RS1!pro_ctrsto = 1, "S", "N"))
                 
                 vaSpread1.Col = 11
                 vaSpread1.text = propon
                 
                 vaSpread1.Col = 12
                 vaSpread1.text = "N"
                 
                 vaSpread1.Col = 16
                 vaSpread1.text = IIf(IsNull(RS1!fcs_CodMaterial), "", RS1!fcs_CodMaterial)
                 
                 vaSpread1.Col = 17
                 vaSpread1.text = IIf(IsNull(RS1!fcs_DenMaterial), "", RS1!fcs_DenMaterial)
                 
                 vaSpread1.Col = 16
                 Text2(0).text = Trim(vaSpread1.text)
                 
                 vaSpread1.Col = 17
                 Text2(1).text = Trim(vaSpread1.text)
                 
'                 '-------> Trae Stock
'                 If RS2.State = 1 Then RS2.Close
'                 RS2.CursorLocation = adUseClient
'                 vg_db.CursorLocation = adUseClient
'                 Set RS2 = vg_db.Execute("SELECT b.bod_canmer FROM b_productos a, b_bodegas b " & _
'                          "WHERE b.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
'                          "AND   b.bod_codpro = a.pro_codigo AND a.pro_codigo = '" & Trim(RS1!pro_codigo) & "'")
'                 vaSpread1.Col = 9
'                 If Not RS2.EOF Then vaSpread1.text = Format(RS2!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
'                 RS2.Close
'                 Set RS2 = Nothing
                 
                 vaSpread1.Col = 9
                 vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa))
                 
                 RS1.MoveNext
                 i = i + 1
                 
              Loop
              
           End If
           RS1.Close
           Set RS1 = Nothing
           
           If Option1(1) = True Then
              
              vaSpread1.Col = 5
              vaSpread1.Row = 0
              vaSpread1.text = "Precio Documento"
           
           Else
              
              vaSpread1.Col = 5
              vaSpread1.Row = 0
              vaSpread1.text = "P.M.P."
           
           End If
        
        End If
        
        vaSpread1.Row = Row
        vaSpread1.Col = 4
        
        If vg_GuiaCD = "1" Then
                
           vaSpread1.Lock = True
                
        ElseIf Trim(vg_GuiaCD) = "" Then
                   
           vaSpread1.Lock = False
               
        End If
        
        If Trim(vaSpread1.text) <> "" Then
        
           canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
           
           vaSpread1.Col = 7
           vaSpread1.text = Format(canrea, fg_Pict(9, vg_DCa))
        
        End If
        
        vaSpread1.Col = 5
        If Trim(vaSpread1.text) <> "" Then
        
           propon = Format(vaSpread1.text, fg_Pict(9, 2))
        
        End If
        
        vaSpread1.Col = 6
        If Trim(vaSpread1.text) <> "" Then
        
           vaSpread1.text = Format(canrea * propon, fg_Pict(9, 2))
        
        End If
        
        vaSpread1.Col = 9
        If Trim(vaSpread1.text) <> "" Then
        
           canbod = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
        
        End If
        
        '------- Total General ---------
        subtot = 0
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            If Trim(vaSpread1.text) <> "" Then
            
               vaSpread1.Col = 6
               subtot = subtot + Format(IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0), fg_Pict(9, vg_DPr))
               
            End If
        
        Next
        Label2.Caption = Format(subtot, fg_Pict(9, vg_DPr))
        
        '-------------------------------
        vaSpread1.Row = Row
        If (canbod - canrea) >= 0 Or Option1(1).Value = True Then
            
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(2).FillColor
            
            vaSpread1.Col = 8
            vaSpread1.text = "N"  'No Bloqueado
            
            Exit Sub
        
        End If
        
        If Option1(0).Value = True And Not Eststo Then Exit Sub
        vaSpread1.Col = -1
        vaSpread1.BackColor = Shape1(1).FillColor
        vaSpread1.Col = 8
        vaSpread1.text = "S"  'Bloqueado
    
    Case 4, 5
    
        vaSpread1.Row = Row
        vaSpread1.Col = 1
        codmer = vaSpread1.text
        
        '------ Producto no esta estoqueable activar columna 5, para ingresar precio del producto
        If Option1(0).Value = True Then
           
           Eststo = True
           If RS1.State = 1 Then RS1.Close
           RS1.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           RS1.Open "SELECT * FROM b_productos WHERE pro_codigo = '" & codmer & "' AND (pro_ctrsto <> 1 OR (pro_ctrsto) IS NULL)", vg_db, adOpenStatic
           If Not RS1.EOF Then
              
              vaSpread1.Col = 5
              vaSpread1.Lock = False
              Eststo = False
           
           End If
           RS1.Close
           Set RS1 = Nothing
           '------ Fin producto no esta estoqueable activar columna 5, para ingresar precio del producto
        
        End If
        
        vaSpread1.Col = 4
        
        If vg_GuiaCD = "1" Then
                
           vaSpread1.Lock = True
                
        ElseIf Trim(vg_GuiaCD) = "" Then
                   
           vaSpread1.Lock = False
                
        End If
        
        If Trim(vaSpread1.text) <> "" And vg_GuiaCD <> "1" Then
           
           canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
           
           vaSpread1.Col = 7
           vaSpread1.text = Format(canrea, fg_Pict(9, vg_DCa))
        
        End If
        
        vaSpread1.Col = 5
        If Trim(vaSpread1.text) <> "" Then
        
           propon = Format(vaSpread1.text, fg_Pict(9, 2))
           
        End If
        
        vaSpread1.Col = 6
        If Trim(vaSpread1.text) <> "" Then
        
           vaSpread1.text = Format(canrea * propon, fg_Pict(9, 2))
           
        End If
        
        EstGCd = True
        vaSpread1.Col = 21
        EstGCd = IIf(vaSpread1.text = "GuiaCD", True, False)
        
        vaSpread1.Col = 7
        If Trim(vaSpread1.text) <> "" And vg_GuiaCD = "1" And Not EstGCd Then
        
           canrec = vaSpread1.text
           vaSpread1.Col = 6
           vaSpread1.text = Format(canrec * propon, fg_Pict(9, 2))
        
        End If
        
        vaSpread1.Col = 9
        If Trim(vaSpread1.text) <> "" Then
        
           canbod = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
        
        End If
        
        '------- Total General ---------
        subtot = 0
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            If Trim(vaSpread1.text) <> "" Then
            
               vaSpread1.Col = 6
               subtot = subtot + Format(IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0), fg_Pict(9, vg_DPr))
            
            End If
            
        Next
        Label2.Caption = Format(subtot, fg_Pict(9, vg_DPr))
        '-------------------------------
        vaSpread1.Row = Row
        If (canbod - canrea) >= 0 Or Option1(1).Value = True Then
            
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(2).FillColor
            
            vaSpread1.Col = 8
            vaSpread1.text = "N"  'No Bloqueado
            Exit Sub
        
        End If
        If Option1(0).Value = True And Not Eststo Then Exit Sub
        vaSpread1.Col = -1
        vaSpread1.BackColor = Shape1(1).FillColor
        
        vaSpread1.Col = 8
        vaSpread1.text = "S"  'Bloqueado
    
    Case 7 And vg_GuiaCD = "1"
        
        propon = 0
        canrec = 0
        
        vaSpread1.Col = 5 'precio
        If Trim(vaSpread1.text) <> "" Then
        
           propon = Format(vaSpread1.text, fg_Pict(9, 2))
           
        End If
        
        EstGCd = True
            
        vaSpread1.Col = 21
        EstGCd = IIf(vaSpread1.text = "GuiaCD", True, False)
        
        vaSpread1.Col = 7 'cantidad recibida
        If Trim(vaSpread1.text) <> "" And vg_GuiaCD = "1" And Not EstGCd Then
        
           canrec = vaSpread1.text
           
           vaSpread1.Col = 6
           vaSpread1.text = Format(canrec * propon, fg_Pict(9, 2))
        
           '------- Total General ---------
           subtot = 0
           For i = 1 To vaSpread1.MaxRows
            
              vaSpread1.Row = i
              vaSpread1.Col = 1
              If Trim(vaSpread1.text) <> "" Then
            
                 vaSpread1.Col = 6
                 subtot = subtot + Format(IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0), fg_Pict(9, vg_DPr))
            
              End If
            
          Next
          Label2.Caption = Format(subtot, fg_Pict(9, vg_DPr))
          '-------------------------------
        
        End If
        
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

On Error GoTo Man_Error

Dim EstGuiaCd As Boolean
Dim canrea    As Double
Dim canrec    As Double
Dim Motivo    As String

vaSpread1.Row = Row
vaSpread1.Col = 21
EstGuiaCd = IIf(Trim(vaSpread1.text) = "", False, True)

Select Case Col

'    Case 7 And ChangeMade = False And EstGuiaCd
'
'        vaSpread1.Col = 7
'        If Trim(vaSpread1.text) <> "" Then
'
'           ant_canrec = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
'
'        End If
        
    Case 7 And ChangeMade = True And EstGuiaCd

        vaSpread1.Col = 4
        
        If Trim(vaSpread1.text) <> "" Then
        
           canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
        
        End If
        
        vaSpread1.Col = 7
        If Trim(vaSpread1.text) <> "" Then
        
           canrec = Format(vaSpread1.Value, fg_Pict(9, vg_DCa))
        
        End If
        
'        If ant_canrec > canrea And canrec < canrea Then
'           vaSpread1.Col = 19
'           vaSpread1.text = ""
'        ElseIf ant_canrec < canrea And canrec > canrea Then
'           vaSpread1.Col = 19
'           vaSpread1.text = ""
'
'        End If
        
        vaSpread1.Col = 19
        Motivo = Trim(vaSpread1.text)

        vaSpread1.Col = 7
        
       
        If canrec > canrea And Trim(Motivo) = "" Then
           
           Toolbar2.Enabled = False
           MsgBox "La cantidad recibida excede de la cantidad es menor...", vbCritical, MsgTitulo
     
        End If
        
        If canrec = canrea Then
        
            vaSpread1.Col = 19
            vaSpread1.Lock = True
            vaSpread1.text = ""
            Toolbar2.Enabled = True
        
        End If
        
        vaSpread1.Col = 19
        
        If canrec <> canrea And Trim(vaSpread1.text) = "" Then
        
            Toolbar2.Enabled = False
            vaSpread1.Col = 19
            vaSpread1.Lock = False
            vaSpread1.SetActiveCell 19, vaSpread1.Row ': vaSpread1.SetFocus
            
            MsgBox "La cantidad recibida es distinta a cantidad documento, debera seleccionar la columna Descripción Motivo ...", vbCritical + vbOKOnly, MsgTitulo
            
        End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)

'On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Or Frame1.Enabled = False Then Exit Sub

Select Case KeyCode

    Case 46 And Shift = 1
        
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        If vaSpread1.Lock = False Then Exit Sub
        If MsgBox("Elimina producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        If vaSpread1.MaxRows = 0 Then
           
           vaSpread1.MaxRows = vaSpread1.MaxRows + 1
           vaSpread1.Row = vaSpread1.MaxRows
           vaSpread1.Col = 1
           vaSpread1.Lock = False
        
        End If
        
        '------- Total General ---------
        Dim subtot As Double
        subtot = 0
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If Trim(vaSpread1.text) <> "" Then
               
               vaSpread1.Col = 6
               subtot = subtot + Format(vaSpread1.text, fg_Pict(9, vg_DPr))
            
            End If
            
        Next
        Label2.Caption = Format(subtot, fg_Pict(9, vg_DPr))
        '-------------------------------
        If vaSpread1.MaxRows = 0 Then Gl_Ac_Botones Me, 4, 6, ""
        If vaSpread1.Enabled = True Then vaSpread1.SetFocus

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub

Dim RS1       As New ADODB.Recordset
Dim RS2       As New ADODB.Recordset
Dim sql1      As String
Dim codsgp    As String
Dim codsac    As String
Dim EstGuiaCd As Boolean
Dim candoc    As Double
Dim Motivo    As String
Dim ivec      As Long
Dim iRow      As Long
Dim EstMotivo As Boolean
Dim GloMotivo As String

vaSpread1.Row = NewRow
vaSpread1.Col = 1
codsgp = Trim(vaSpread1.text)

vaSpread1.Col = 16
Text2(0).text = Trim(vaSpread1.text)

vaSpread1.Col = 16
codsac = Trim(vaSpread1.text)

vaSpread1.Col = 17
Text2(1).text = Trim(vaSpread1.text)

vaSpread1.Col = 4
candoc = IIf(Trim(vaSpread1.text) = 0, 0, Val(vaSpread1.text))

EstGuiaCd = IIf(Trim(Trim(vg_GuiaCD)) = "", False, True)

vaSpread1.Col = 21
If Trim(vaSpread1.text) <> "" Then

   EstGuiaCd = True 'IIf(Trim(vaSpread1.text) = "", False, True)

ElseIf Trim(vaSpread1.text) = "" Then

   EstGuiaCd = False

End If

If Trim(vg_GuiaCD) = "1" And Trim(vaSpread1.text) = "" And NewRow > 0 Then

   For ivec = 1 To vaSpread1.MaxRows

      vaSpread1.Row = ivec
      vaSpread1.Col = 19
      If vaSpread1.Lock = False Then

         EstGuiaCd = True
         Exit For

      End If

   Next ivec

End If

If vg_pais = "CL" And vg_FDC = "OC" And Trim(codsgp) <> "" And Option1(1).Value = True Then
   
   Image1(5).Visible = IIf(ValidarProductosSgpSac(Trim(codsac), Trim(codsgp)), True, False)

End If

Dim canrea As Double
Dim canrec As Double

If NewRow < 1 Or vaSpread1.MaxRows < 1 Or (NewCol > 2 And Not EstGuiaCd) Then Exit Sub

If EstGuiaCd Then

   
   vaSpread1.Row = Row
   
   vaSpread1.Col = 4
   candoc = IIf(Trim(vaSpread1.text) = 0, 0, Val(vaSpread1.text))
   
   If candoc > 0 Then
   
   vaSpread1.Col = 4

   If Trim(vaSpread1.text) <> "" Then
        
      canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
        
   End If
        
   vaSpread1.Col = 7
   If Trim(vaSpread1.text) <> "" Then
        
      canrec = Format(vaSpread1.Value, fg_Pict(9, vg_DCa))
        
   End If
        
   vaSpread1.Col = 19
   Motivo = Trim(vaSpread1.text)

   vaSpread1.Col = 7
   
   If canrec > canrea And Trim(Motivo) = "" Then
      Toolbar2.Enabled = False
'      MsgBox "La cantidad recibida excede de la cantidad es menor...", vbCritical, MsgTitulo
           
   End If
        
   If canrec = canrea Then
        
      vaSpread1.Col = 19
      vaSpread1.Lock = True
      vaSpread1.text = ""
      Toolbar2.Enabled = True
        
   End If
        
   vaSpread1.Col = 19
        
   If canrec <> canrea And Trim(vaSpread1.text) = "" And Row <> NewRow Then
        
      Toolbar2.Enabled = False
      
      vaSpread1.Col = 19
      vaSpread1.Lock = False
      vaSpread1.SetActiveCell 19, vaSpread1.Row
      vaSpread1.SetFocus
            
         '   MsgBox "La cantidad recibida es distinta a cantidad documento, debera seleccionar la columna Descripción Motivo ...", vbCritical + vbOKOnly, MsgTitulo
        
      vaSpread1.Col = 19
      Exit Sub
            'vaSpread1.SetActiveCell 19, vaSpread1.Row
            'vaSpread1.SetFocus
        
   End If
   
   End If
   
End If

vaSpread1.Row = NewRow

If Row <> NewRow Then
   
   vaSpread1.Row = Row
   vaSpread1.Col = 1
   
   If Trim(vaSpread1.text) = "" Then
      
      vaSpread1.MaxRows = vaSpread1.MaxRows - 1
      vaSpread1.SetActiveCell 1, vaSpread1.MaxRows
      Exit Sub
   
   ElseIf Col <> 1 Then
      
      vaSpread1.Col = 1
      vaSpread1.Lock = True
   
   End If

End If

Select Case Col

    Case 1 And Not EstGuiaCd
    
        Dim codigo As String
        Dim propon As Double
        vaSpread1.Row = Row
        vaSpread1.Col = Col
        Row_Activo = Row
        codigo = vaSpread1.text
        CodigoProd = vaSpread1.text
        If vaSpread1.Lock = True Then Exit Sub
        
        If Trim(vaSpread1.text) = "" Then
            
            vg_nombre = ""
            vg_codigo = ""
            vg_bodega = 0
            vg_bodega = Val(fg_codigocbo(Combo1, 1, 10, ""))
            vg_left = fpayuda(1).Left + 1920
            B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", IIf(Option1(1).Value = True, "ProVig", "ProInvNoStock")
            B_TabEst.Show 1
            vaSpread1.Refresh
            If vg_codigo = "" Then vaSpread1.Refresh: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
            codigo = vg_codigo
        
        End If
        
        If vg_pais <> "CL" Or vg_FDC <> "OC" Then
           
           For i = 1 To vaSpread1.MaxRows
               
               vaSpread1.Col = 1
               vaSpread1.Row = i
               If Trim(vaSpread1.text) = Trim(codigo) And Row <> i And Trim(vaSpread1.text) <> "" Then
               
                  MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo
                  
                  vaSpread1.Row = vaSpread1.ActiveRow
                  vaSpread1.Col = 1
                  vaSpread1.text = ""
                  
                  vaSpread1.Col = 2
                  vaSpread1.text = ""
                  
                  vaSpread1.Col = 3
                  vaSpread1.text = ""

                  vaSpread1.Col = 16
                  vaSpread1.text = ""
                
                  vaSpread1.Col = 17
                  vaSpread1.text = ""
                  
                  CodigoProd = ""
                  
                  vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow
                  Exit Sub
           
               End If
               
           Next i
        
        End If
        
        CodigoProd = ""
        vaSpread1.Row = vaSpread1.ActiveRow
'        sql1 = IIf(vg_tipbase = "1", " AND CDATE(x.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), x.foc_vigfin,101) >  '" & Date & "'")
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
'        If Option1(1).Value = True Then
'
'           RS1.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, b.uni_nombre, a.pro_ctrsto, " & _
'                    "(SELECT TOP 1 x.foc_codsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_codsac, " & _
'                    "(SELECT TOP 1 x.foc_nomsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_nomsac, " & _
'                    "(SELECT TOP 1 x.foc_unisac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_unisac, " & _
'                    "(SELECT TOP 1 x.foc_faccon FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_faccon " & _
'                    "FROM b_productos a, a_unidad b, a_tiposervicio d, b_clientes e " & _
'                    "WHERE (d.tis_codigo = e.cli_codtis OR a.pro_maepro < 1) AND e.cli_codigo = '" & MuestraCasino(1) & "' AND (d.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) AND a.pro_coduni = b.uni_codigo AND a.pro_codigo = '" & codigo & "' " & _
'                    "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven = 0)", vg_db, adOpenStatic
'
'        Else
'
'           RS1.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, b.uni_nombre, a.pro_ctrsto, " & _
'                    "'' AS foc_codsac, '' AS foc_nomsac, '' AS foc_unisac, 1 AS foc_faccon " & _
'                    "FROM b_productos a, a_unidad b, a_tiposervicio d, b_clientes e " & _
'                    "WHERE (d.tis_codigo = e.cli_codtis OR a.pro_maepro < 1) AND e.cli_codigo = '" & MuestraCasino(1) & "' AND (d.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) AND a.pro_coduni = b.uni_codigo AND a.pro_codigo = '" & codigo & "' " & _
'                    "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven = 0) " & _
'                    "UNION SELECT DISTINCT pro.pro_codigo, pro.pro_nombre, c.uni_nombre, pro.pro_ctrsto, '' AS foc_codsac, '' AS foc_nomsac, '' AS foc_unisac, 1 AS foc_faccon FROM b_productos pro, b_bodegas bod, a_unidad c WHERE pro.pro_codigo = bod.bod_codpro AND pro.pro_coduni = c.uni_codigo AND bod.bod_codbod = " & vg_codbod & " AND bod.bod_canmer > 0 AND pro.pro_codigo = '" & codigo & "'", vg_db, adOpenStatic
'
'        End If
        
        Set RS1 = vg_db.Execute("sgp_Sel_ProductoTraspasoSalEnt '" & MuestraCasino(1) & "', '" & codigo & "'")
        
        If Not RS1.EOF Then
            
            Do While Not RS1.EOF
                
                vaSpread1.MaxRows = vaSpread1.ActiveRow
                vaSpread1.Row = vaSpread1.ActiveRow
                
                For i = 4 To IIf(Option1(0).Value = True, 5, 7)
                    
                    vaSpread1.Col = i
                    vaSpread1.Lock = IIf(i <> 6 And Option1(1).Value = True, False, IIf(i <> 5 And Option1(0).Value = True, False, True))
                
                Next i
                
                vaSpread1.Row = vaSpread1.ActiveRow
                vaSpread1.Col = 1
                vaSpread1.text = RS1!pro_codigo
                
                vaSpread1.Col = 2
                vaSpread1.text = RS1!pro_nombre
                
                vaSpread1.Col = 3
                vaSpread1.text = RS1!uni_nombre
                
                vaSpread1.Col = 4
                If vg_GuiaCD = "1" Then
                
                   vaSpread1.Lock = True
                
                ElseIf Trim(vg_GuiaCD) = "" Then
                   
                   vaSpread1.Lock = False
                
                End If
                
                If Trim(vaSpread1.text) = "" Then vaSpread1.text = 0
                
                '-------> Trae pmp
                propon = 0
                
'                If RS2.State = 1 Then RS2.Close
'                RS2.CursorLocation = adUseClient
'                vg_db.CursorLocation = adUseClient
'                RS2.Open "SELECT TOP 1 ppd_cencos, ppd_codpro, ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
'                         "FROM b_productospmpdia " & _
'                         "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
'                         "AND   ppd_codpro = '" & RS1!pro_codigo & "' " & _
'                         "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
'                         "GROUP BY ppd_cencos, ppd_codpro, ppd_propon " & _
'                         "HAVING (ppd_propon) > 0 ORDER BY Max(ppd_fecdia) DESC", vg_db, adOpenStatic
'                If Not RS2.EOF Then propon = RS2!ppd_propon
'                RS2.Close
'                Set RS2 = Nothing
                
                propon = RS1!ppd_propon
                
                vaSpread1.Col = 5
                vaSpread1.text = IIf(Option1(1).Value = True, 0, propon)
                
                vaSpread1.Col = 6
                If Trim(vaSpread1.text) = "" Then vaSpread1.text = 0
                
                vaSpread1.Col = 7
                If Trim(vaSpread1.text) = "" Then vaSpread1.text = 0
                
                vaSpread1.Col = 8
                vaSpread1.text = "N" 'No bloquedo
                
                vaSpread1.Col = 10
                vaSpread1.text = IIf(IsNull(RS1!pro_ctrsto), "N", IIf(RS1!pro_ctrsto = 1, "S", "N"))
                
                vaSpread1.Col = 11
                vaSpread1.text = propon
                
                vaSpread1.Col = 12
                vaSpread1.text = "N"
                
                vaSpread1.Col = 16
                vaSpread1.text = IIf(IsNull(RS1!fcs_CodMaterial), "", RS1!fcs_CodMaterial)
                
                vaSpread1.Col = 17
                vaSpread1.text = IIf(IsNull(RS1!fcs_DenMaterial), "", RS1!fcs_DenMaterial)
                
                vaSpread1.Col = 16
                Text2(0).text = Trim(vaSpread1.text)
                
                vaSpread1.Col = 17
                Text2(1).text = Trim(vaSpread1.text)
                
                vaSpread1.Col = 4
                If Trim(vaSpread1.text) <> "" And CStr(vaSpread1.text) > 0 Then
                
                   canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
                   
                   vaSpread1.Col = 7
                   vaSpread1.text = Format(canrea, fg_Pict(9, vg_DCa))
                
                Else
                    
                   vaSpread1.Col = 7
                   canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
                    
                End If
                
                vaSpread1.Col = 5
                If Trim(vaSpread1.text) <> "" Then propon = Format(vaSpread1.text, fg_Pict(9, 2))
                
                vaSpread1.Col = 6
                If Trim(vaSpread1.text) <> "" Then vaSpread1.text = Format(canrea * propon, fg_Pict(9, 2))
                
'                '-------> Trae Stock
'                If RS2.State = 1 Then RS2.Close
'                RS2.CursorLocation = adUseClient
'                vg_db.CursorLocation = adUseClient
'                RS2.Open "SELECT b.bod_canmer FROM b_productos a, b_bodegas b " & _
'                         "WHERE b.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
'                         "AND   b.bod_codpro = a.pro_codigo AND a.pro_codigo = '" & Trim(RS1!pro_codigo) & "'", vg_db, adOpenStatic
'                vaSpread1.Col = 9
'                If Not RS2.EOF Then vaSpread1.text = Format(RS2!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
'                RS2.Close
'                Set RS2 = Nothing

                vaSpread1.Col = 9
                vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa))
                
                RS1.MoveNext
                
                i = i + 1
                i = 4
            
            Loop
            
            If NewRow <> Row Then vaSpread1.Col = 1: vaSpread1.Lock = True
        
        Else
           
           vaSpread1.Row = vaSpread1.ActiveRow
           
           For i = 4 To 7
           
               vaSpread1.Col = i
               vaSpread1.Lock = True
           
           Next i
           
           vaSpread1.Col = 1
           vaSpread1.text = ""
           
           vaSpread1.Col = 2
           vaSpread1.text = ""
           
           vaSpread1.Col = 3
           vaSpread1.text = ""
           
           vaSpread1.Col = 16
           vaSpread1.text = ""
           
           vaSpread1.Col = 17
           vaSpread1.text = ""
           
           vaSpread1.Col = 16
           Text2(0).text = Trim(vaSpread1.text)
           
           vaSpread1.Col = 17
           Text2(1).text = Trim(vaSpread1.text)
           
           i = 1
           MsgBox "producto no existe...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow
        
        End If
        RS1.Close
        Set RS1 = Nothing
            
        If Option1(1) = True Then
           
           vaSpread1.Col = 5
           vaSpread1.Row = 0
           vaSpread1.text = "Precio"
        
        Else
           
           vaSpread1.Col = 5
           vaSpread1.Row = 0
           vaSpread1.text = "P.M.P."
        
        End If
        
        If vaSpread1.MaxRows = 1 Then Gl_Ac_Botones Me, 4, 6, ""
        vaSpread1.Col = 4
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.SetActiveCell i, vaSpread1.MaxRows
        If vaSpread1.Enabled = True Then vaSpread1.SetFocus

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

On Error GoTo Man_Error

If Row = 0 Then Exit Sub
Dim Stock  As String
Dim Nombre As String
TipWidth = 4000
ShowTip = True

MultiLine = 2
vaSpread1.Row = Row
vaSpread1.Col = 9
Stock = Format(vaSpread1.text, fg_Pict(9, vg_DCa))

vaSpread1.Row = Row
vaSpread1.Col = 2
Nombre = vaSpread1.text

TipText = "Bodega   : " & Trim(Left(Combo1(1).text, 50)) & vbCrLf & _
          "Producto : " & Trim(Nombre) & vbCrLf & _
          "Stock       : " & Format(Trim(Stock), fg_Pict(9, vg_DCa))

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub ChequearOCompras()

On Error GoTo Man_Error

Dim sql1 As String, sql2 As String, sql3 As String
Dim RS5 As New ADODB.Recordset
Dim RS6 As New ADODB.Recordset
'-------> Activar o desactivar ordenes de compras
Image1(2).Enabled = True
sql1 = IIf(vg_tipbase = "1", " val(format(a.solite_dtent, 'yyyymm')) ", " substring(CONVERT(varchar(10), a.solite_dtent,112),1,6) ")
sql2 = IIf(vg_tipbase = "1", " '" & Format(fpDateTime1(0), "yyyymm") & "' ", " '" & Format(fpDateTime1(0), "yyyymm") & "' ")
sql3 = IIf(vg_tipbase = "1", " SUM(IIF(a.tipsol_idsol = 4,(-1 * a.pedite_qtcpa), a.pedite_qtcpa)) AS difer ", " SUM(CASE WHEN a.tipsol_idsol = 4 THEN (-1 * a.pedite_qtcpa) ELSE  a.pedite_qtcpa END) AS difer ")


If RS5.State = 1 Then RS5.Close
RS5.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
RS5.Open "SELECT " & sql3 & " " & _
         "FROM b_ocsac a " & _
         "WHERE a.cadfil_cdfil = '" & MuestraCasino(1) & "' " & _
         "AND   " & sql1 & "   = " & sql2 & " AND a.pedite_flafo = 0" & _
         "", vg_db, adOpenStatic
If Not RS5.EOF And Not IsNull(RS5!difer) Then
   
   sql3 = IIf(vg_tipbase = "1", " SUM(b.ocr_cancom - (IIF(a.tipsol_idsol = 4,(-1 * a.pedite_qtcpa),a.pedite_qtcpa)) ) AS difer ", " SUM(b.ocr_cancom - (CASE WHEN a.tipsol_idsol = 4 THEN (-1 * a.pedite_qtcpa) ELSE a.pedite_qtcpa END)) AS difer ")

   If RS6.State = 1 Then RS6.Close
   RS6.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   RS6.Open "SELECT " & sql3 & " " & _
            "FROM   b_ocsac a, b_ocsacrecibido b " & _
            "WHERE  a.cadfor_nrcgc = b.ocr_rutpro " & _
            "AND    a.solite_dtent = b.ocr_fecoc  " & _
            "AND    a.cadfil_cdfil = '" & MuestraCasino(1) & "' " & _
            "AND    " & sql1 & "   = " & sql2 & " " & _
            "AND    a.cpopro_cdpro = b.ocr_codprodsac AND a.pedite_flafo = 0", vg_db, adOpenStatic
   Image1(2).Visible = IIf(RS6.EOF Or RS6!difer <> 0 Or IsNull(RS6!difer) Or RS5!difer > 0, True, False)
   RS6.Close
   Set RS6 = Nothing

End If
RS5.Close
Set RS5 = Nothing

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub
