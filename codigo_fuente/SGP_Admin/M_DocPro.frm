VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_DocPro 
   Caption         =   "Documento de Proveedor"
   ClientHeight    =   7305
   ClientLeft      =   210
   ClientTop       =   1905
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   45
      TabIndex        =   27
      Top             =   360
      Width           =   11460
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   8115
         TabIndex        =   13
         Top             =   855
         Width           =   3105
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "CFC"
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
            Left            =   120
            TabIndex        =   7
            Top             =   285
            Width           =   1110
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "FOFI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   135
            TabIndex        =   8
            Top             =   585
            Width           =   1095
         End
         Begin EditLib.fpDoubleSingle Double1 
            Height          =   330
            Index           =   6
            Left            =   1770
            TabIndex        =   41
            Top             =   450
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
            _ExtentY        =   582
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
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
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
            Caption         =   "Folio Nº"
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
            Left            =   1890
            TabIndex        =   35
            Top             =   210
            Width           =   690
         End
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         ItemData        =   "M_DocPro.frx":0000
         Left            =   3420
         List            =   "M_DocPro.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   2670
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "M_DocPro.frx":0004
         Left            =   6435
         List            =   "M_DocPro.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   465
         Width           =   2205
      End
      Begin EditLib.fpText fpText 
         Height          =   330
         Index           =   0
         Left            =   285
         TabIndex        =   0
         Top             =   465
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
         _ExtentY        =   582
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
         AlignTextH      =   0
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
         MaxLength       =   20
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   5
         Left            =   9285
         TabIndex        =   2
         Top             =   465
         Width           =   1200
         _Version        =   196608
         _ExtentX        =   2117
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
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   0
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
      Begin EditLib.fpDateTime Date1 
         Height          =   345
         Index           =   1
         Left            =   1860
         TabIndex        =   4
         Top             =   1320
         Width           =   1440
         _Version        =   196608
         _ExtentX        =   2540
         _ExtentY        =   609
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
      Begin EditLib.fpDateTime Date1 
         Height          =   345
         Index           =   0
         Left            =   285
         TabIndex        =   3
         Top             =   1320
         Width           =   1470
         _Version        =   196608
         _ExtentX        =   2593
         _ExtentY        =   609
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
      Begin EditLib.fpText fpText 
         Height          =   330
         Index           =   1
         Left            =   6270
         TabIndex        =   6
         Top             =   1320
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
         _ExtentY        =   582
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
         AlignTextH      =   0
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
         MaxLength       =   20
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden de Compra"
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
         Left            =   6255
         TabIndex        =   46
         Top             =   1035
         Width           =   1485
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   3495
         TabIndex        =   43
         Top             =   1365
         Width           =   2640
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   6510
         TabIndex        =   42
         Top             =   510
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   10530
         Picture         =   "M_DocPro.frx":0008
         Top             =   315
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   1755
         Picture         =   "M_DocPro.frx":0312
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de emisión"
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
         Index           =   13
         Left            =   285
         TabIndex        =   34
         Top             =   1035
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de vecto."
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
         Index           =   14
         Left            =   1890
         TabIndex        =   33
         Top             =   1035
         Width           =   1410
      End
      Begin VB.Label Label1 
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
         Index           =   16
         Left            =   3435
         TabIndex        =   32
         Top             =   1035
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Documento"
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
         Index           =   11
         Left            =   9270
         TabIndex        =   31
         Top             =   225
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
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
         Left            =   6450
         TabIndex        =   29
         Top             =   225
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rut"
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
         Left            =   330
         TabIndex        =   28
         Top             =   225
         Width           =   315
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   2190
         TabIndex        =   12
         Top             =   465
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   2235
         TabIndex        =   30
         Top             =   510
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   45
      TabIndex        =   20
      Top             =   6360
      Width           =   11475
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   15
         Top             =   465
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
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   2
         Left            =   2910
         TabIndex        =   16
         Top             =   465
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
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   3
         Left            =   4275
         TabIndex        =   17
         Top             =   465
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
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   4
         Left            =   5670
         TabIndex        =   18
         Top             =   465
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
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   0
         Left            =   195
         TabIndex        =   14
         Top             =   465
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
         DecimalPoint    =   "."
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ","
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   7
         Left            =   6975
         TabIndex        =   36
         Top             =   225
         Visible         =   0   'False
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
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   8
         Left            =   8160
         TabIndex        =   37
         Top             =   225
         Visible         =   0   'False
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
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   9
         Left            =   6990
         TabIndex        =   38
         Top             =   480
         Visible         =   0   'False
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
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   10
         Left            =   8160
         TabIndex        =   39
         Top             =   480
         Visible         =   0   'False
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
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   11
         Left            =   9375
         TabIndex        =   40
         Top             =   210
         Visible         =   0   'False
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
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Index           =   21
         Left            =   5685
         TabIndex        =   25
         Top             =   225
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Otr. Imp."
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
         Index           =   20
         Left            =   4260
         TabIndex        =   24
         Top             =   225
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "I.V.A"
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
         Index           =   19
         Left            =   2925
         TabIndex        =   23
         Top             =   225
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Neto"
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
         Index           =   18
         Left            =   1560
         TabIndex        =   22
         Top             =   225
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Exento"
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
         Index           =   17
         Left            =   195
         TabIndex        =   21
         Top             =   225
         Width           =   600
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Frame Frame4 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   45
      TabIndex        =   26
      Top             =   2190
      Width           =   11475
      Begin VB.Frame Frame6 
         Height          =   705
         Left            =   7755
         TabIndex        =   45
         Top             =   3360
         Width           =   3585
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   390
            Left            =   120
            TabIndex        =   11
            Top             =   210
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
      End
      Begin VB.Frame Frame3 
         Caption         =   "Glosa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   135
         TabIndex        =   44
         Top             =   3360
         Width           =   7590
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   420
            Index           =   0
            Left            =   75
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   195
            Width           =   7410
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3120
         Left            =   135
         TabIndex        =   9
         Top             =   225
         Width           =   11190
         _Version        =   393216
         _ExtentX        =   19738
         _ExtentY        =   5503
         _StockProps     =   64
         AutoClipboard   =   0   'False
         ColsFrozen      =   2
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
         MaxCols         =   16
         MaxRows         =   50
         SpreadDesigner  =   "M_DocPro.frx":061C
         VisibleCols     =   11
         ScrollBarTrack  =   3
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   105
         Top             =   3585
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
               Picture         =   "M_DocPro.frx":122C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_DocPro.frx":1546
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "M_DocPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset
Dim ibusca As Long, i As Long
Dim itab As Integer, itexto As Integer, numero As Integer
Dim modo As String, Codigo As String, rut As String, contador As Integer
Dim vecdatos(7) As String
Dim Encontrado As Boolean 'Variable para saber si encontro o no el registro
Dim Impuestos() As Variant
Dim indice, CantReg As Long, FolioSn As Long
Dim incluir As String, alterar As String, eliminar As String, imprimir
Dim IndiceOpt As Double, Est As Boolean
'Dim ContLin  As Long
Private Function Fg_Puntocoma(ByVal Parentesis As String) As String
Dim X%
Dim ValLcntH$
ValLcntH = ""

For X = 1 To Len(Parentesis)
    If Asc(Mid(Parentesis, X, 1)) <> 59 Then
       ValLcntH = ValLcntH + Mid(Parentesis, X, 1)
    End If
Next X
Fg_Puntocoma = ValLcntH
End Function
Private Function Valida_DatosGrilla() As Boolean
Valida_DatosGrilla = False
With vaSpread1
    For i = 1 To .MaxRows
        .Row = i: .Col = 4:
        If Val(.Text) = 0 Then Valida_DatosGrilla = True: Exit For
        .Row = i: .Col = 5
        If Val(.Text) = 0 Then Valida_DatosGrilla = True: Exit For
    Next i
End With
End Function

Private Function Nuevo_Registro()
On Error GoTo Nuevo
fpText(0).Text = "": fpayuda(0).Caption = ""
Double1(0).Text = "": Double1(1).Text = "": Double1(2).Text = "": Double1(3).Text = "": Double1(4).Text = "": Double1(5).Text = "": Double1(6).Text = ""
Date1(0).Text = "": Date1(1).Text = "": Combo2(0).ListIndex = -1: Text1(0).Text = ""
Date1(0).Text = Date: Date1(1).Text = Date
If Combo2(1).ListCount > 0 Then Combo2(1).ListIndex = 0
vaSpread1.MaxRows = 0
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Frame6.Enabled = True
vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.Lock = False
Option1(0).Value = False: Option1(1).Value = False
Image1(1).Visible = False: Encontrado = False
vaSpread1.SetActiveCell 1, 1
Exit Function
Nuevo:
    MsgBox Err.Description, vbOKOnly, Msgtitulo
End Function

Private Function SumaDiferencias(FilaDif As Long, CantDif As Double, PreDif As Double, PorDes As Double)
Dim i As Long, TotE As Double, TotN As Double, TotI As Double, TotO As Double, StrImp As String, StrImpb As String
Dim codi As Long, PctI As Double, CosI As Long, aPos As Long, Cant As Double
Dim PreC As Double, MonD As Double, MonI As Double, Cdif As Long
On Error GoTo Error_Suma
Dim RS1 As New ADODB.Recordset
For i = 1 To UBound(Impuestos)
    Impuestos(i, 4) = 0
Next i
'TotE = 0: TotN = 0: TotI = 0: TotO = 0:
For i = FilaDif To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1: CodPro = Trim(vaSpread1.Value)
    vaSpread1.Col = 4: Cant = Val(vaSpread1.Value)
    vaSpread1.Col = 16: Cdif = Val(vaSpread1.Value)
    If CodPro <> "" And Cdif > 0 Then
        PreC = PreDif
        MonD = (CantDif * PreDif) * (PorDes / 100)
        MonPro = (CantDif * PreDif) - MonD

        vaSpread1.Col = 14
        StrImp = Trim(vaSpread1.Text): MonI = 0
        If Len(StrImp) <> 0 Then
            Do While InStr(StrImp, ";") <> 0
                StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
                StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
                codi = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
                PctI = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
                CosI = Val(Mid(StrImpb, 1))
                aPos = fg_BuscaArr(Impuestos, codi, 2, 1)
                If aPos <> 0 Then Impuestos(aPos, 4) = Impuestos(aPos, 4) + MonPro
                If CosI = 1 Then MonI = MonI + Round((PreC - MonD) * (PctI / 100), vg_DPr)
            Loop
        Else
            TotE = TotE + MonPro
        End If
        vaSpread1.Col = 15: vaSpread1.Value = (PreC - MonD) + MonI
    End If
Next i
For i = 1 To UBound(Impuestos)
    If Impuestos(i, 1) = 1 Then
        TotN = TotN + Impuestos(i, 4)
        TotI = TotI + Round(Impuestos(i, 4) * (Impuestos(i, 3) / 100), 0)
    Else
        TotO = TotO + Round(Impuestos(i, 4) * (Impuestos(i, 3) / 100), 0)
    End If
Next i
If fg_codigocbo(Combo2, 0, 2, "") <> "GD" Then
    Double1(7).Value = TotE
    Double1(8).Value = TotN
    Double1(9).Value = TotI
    Double1(10).Value = TotO
    Double1(11).Value = TotE + TotN + TotI + TotO
Else
    Double1(7).Value = 0
    Double1(8).Value = 0
    Double1(9).Value = 0
    Double1(10).Value = 0
    Double1(11).Value = TotE + TotN + TotI + TotO
End If
Exit Function
Error_Suma:
MsgBox "Error : " & Err.Number & " - " & Err.Description, vbExclamation, Msgtitulo
Resume Next

End Function
Public Function SumarTotales()
Dim i As Long, TotE As Double, TotN As Double, TotI As Double, TotO As Double, StrImp As String, StrImpb As String
Dim codi As Long, PctI As Double, CosI As Long, aPos As Long, Cant As Double, DifUni As Double
Dim PreC As Double, MonD As Double, MonI As Double
On Error GoTo Error_Suma
Dim RS1 As New ADODB.Recordset
For i = 1 To UBound(Impuestos)
    Impuestos(i, 4) = 0
Next i
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1: CodPro = Trim(vaSpread1.Value)
    If CodPro <> "" Or IsNull(CodPro) = False Then
        vaSpread1.Col = 4: Cant = Val(vaSpread1.Value)
        vaSpread1.Col = 5: PreC = Val(vaSpread1.Value)
        vaSpread1.Col = 7: MonD = Val(vaSpread1.Value)
        vaSpread1.Col = 8: MonPro = Val(vaSpread1.Value)
        vaSpread1.Col = 14
        DifUni = 0
        If Val(Cant) > 0 And Val(MonD) > 0 Then DifUni = MonD / Cant
        StrImp = Trim(vaSpread1.Text): MonI = 0
'--- calcula otros impuestos
        If Len(StrImp) <> 0 Then
            Do While InStr(StrImp, ";") <> 0
                StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
                StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
                codi = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
                PctI = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
                CosI = Val(Mid(StrImpb, 1))
                aPos = fg_BuscaArr(Impuestos, codi, 2, 1)
                If aPos <> 0 Then Impuestos(aPos, 4) = Impuestos(aPos, 4) + MonPro
                If CosI = 1 Then MonI = MonI + Round((PreC - DifUni) * (PctI / 100), vg_DPr)
            Loop
        Else
            TotE = TotE + MonPro
        End If
        vaSpread1.Col = 15: vaSpread1.Value = (PreC - DifUni) + MonI
'--fin calcula otros impuestos
    End If
Next i
For i = 1 To UBound(Impuestos)
    If Impuestos(i, 1) = 1 Then
        TotN = TotN + Impuestos(i, 4)
        TotI = TotI + Round(Impuestos(i, 4) * (Impuestos(i, 3) / 100), 0)
    Else
        TotO = TotO + Round(Impuestos(i, 4) * (Impuestos(i, 3) / 100), 0)
    End If
Next i
If fg_codigocbo(Combo2, 0, 2, "") <> "GD" Then
    Double1(0).Value = TotE
    Double1(1).Value = TotN
    Double1(2).Value = TotI
    Double1(3).Value = TotO
    Double1(4).Value = TotE + TotN + TotI + TotO
Else
    Double1(0).Value = 0
    Double1(1).Value = 0
    Double1(2).Value = 0
    Double1(3).Value = 0
    Double1(4).Value = TotE + TotN + TotI + TotO
End If
Exit Function
Error_Suma:
MsgBox "Error : " & Err.Number & " " & Err.Description, vbExclamation, Msgtitulo
Resume Next
End Function

Private Sub Combo2_Click(Index As Integer)
Double1_LostFocus 5
Image1(1).Visible = IIf(fg_codigocbo(Combo2, 0, 2, "") = "FA" Or fg_codigocbo(Combo2, 0, 2, "") = "NC", True, False)
Select Case Combo2(0).ListIndex
Case Is = 0, 2, 3
    Double1(0).Enabled = True: Double1(1).Enabled = True: Double1(2).Enabled = True: Double1(3).Enabled = True: Double1(4).Enabled = True
Case Is = 1, 4
    Double1(0).Enabled = False: Double1(1).Enabled = False: Double1(2).Enabled = False: Double1(3).Enabled = False: Double1(4).Enabled = False
Case Is = 5
    Double1(0).Enabled = False: Double1(1).Enabled = False: Double1(2).Enabled = False: Double1(3).Enabled = True: Double1(4).Enabled = True
End Select
End Sub

Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Date1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Date1_LostFocus(Index As Integer)
'valida primer campo de fecha
If Date1(0).Text = "" Then
   Date1(0).Text = Date
'valida segundo campo fecha
End If
If Date1(1).Text = "" Then
    Date1(1).Text = Date
End If
End Sub

Private Sub Double1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim RS3 As New ADODB.Recordset
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Double1_LostFocus(Index As Integer)
If Est Then Exit Sub
If Index <> 5 Then Exit Sub
If Index = 5 And Val(Double1(5)) = 0 Then Exit Sub
Est = True
RS1.Open "SELECT * from  b_totcompras where toc_rutpro = '" & fg_DespintaRut(fpText(0).Text) & "'" & _
          " and toc_tipdoc = '" & fg_codigocbo(Combo2, 0, 2, "") & "'  and toc_numdoc = " & Val(Double1(5).Value), vg_db, adOpenStatic
If RS1.EOF Then 'Si no existe el documento
    If modo = "" Or modo = "N" Then modo = "A": 'Ac_Botones
'    If fg_codigocbo(Combo2, 0, 2, "") = "FA" Then
'        RS3.Open "select count(*) as mayor from b_totcompras where toc_tipdoc = '" & fg_codigocbo(Combo2, 0, 2, "") & "' and toc_rutpro ='" & fg_DespintaRut(fpText(0).Text) & "'", vg_db, adOpenStatic
'        If RS3!mayor = 0 Then
'        Else
'        End If
'        RS3.Close: Set RS3 = Nothing
'    End If
    vg_RDC = fg_DespintaRut(fpText(0).Text)
Else
    Encontrado = True
    modo = "M":
    Gl_Ac_Botones Me, 7, 2, ""
    '----Deshabilitar Botones
    Date1(0).Text = RS1!toc_fecemi
    Date1(1).Text = RS1!toc_fecven
    Double1(0).Text = RS1!toc_exedoc
    Double1(1).Text = RS1!toc_netdoc
    Double1(2).Text = RS1!toc_ivadoc
    Double1(3).Text = RS1!toc_otrimp
    Double1(4).Text = RS1!toc_totdoc
    
    If RS1!toc_tipinf = "C" Then
        Double1(6).Text = RS1!toc_numinf
        Option1(1).Value = True
    ElseIf RS1!toc_tipinf = "F" Then
        Double1(6).Text = RS1!toc_numinf
        Option1(0).Value = True
    End If
    Combo2(1).ListIndex = fg_buscacbo(Combo2, 1, 10, fg_pone_cero(Str(RS1!toc_codbod), 10))
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame6.Enabled = False
    vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.Lock = True
    '******* Detalle de Documento
    RS2.Open "SELECT a.*, b.*, c.uni_nomcor from b_detcompras a, b_productos b, a_unidad c where a.dec_codmer=b.pro_codigo and b.pro_coduni=c.uni_codigo and a.dec_rutpro = '" & fg_DespintaRut(fpText(0).Text) & "' and a.dec_tipdoc = '" & fg_codigocbo(Combo2, 0, 2, "") & "' and a.dec_numdoc = " & Val(Double1(5).Value) & " order by a.dec_numlin", vg_db, adOpenStatic
    With vaSpread1
        .MaxRows = 0
        Do While Not RS2.EOF
            .MaxRows = .MaxRows + 1
            .Col = 1: .Row = .MaxRows: .Value = RS2!dec_codmer
            .Col = 2: .Value = Trim(RS2!pro_nombre)
            .Col = 3: .Value = RS2!uni_nomcor
            .Col = 4: .Value = RS2!dec_canmer
            .Col = 5: .Value = RS2!dec_precom
            .Col = 6: .Value = RS2!dec_pctdes
            .Col = 7: .Value = RS2!dec_valdes
            .Col = 8: .Value = RS2!dec_ptotal
            .Col = 9: .Value = RS2!dec_canrec
            .Col = 10: .Value = RS2!dec_prerec
            .Col = 11: .Value = RS2!dec_descri
            .Col = 13: .Value = RS2!dec_mueinv
            RS2.MoveNext
        Loop
    End With
    RS2.Close: Set RS2 = Nothing
    '******Fin detalle de Documento
End If
RS1.Close: Set RS1 = Nothing
Est = False
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
'*** Creador : MSP
'*** Fecha   : 05-08-2004
Me.Width = 11670
Me.Height = 7860
Me.HelpContextID = vg_OpcM
Msgtitulo = "Documento de proveedores"
Est = False
fg_centra Me
vaSpread1.Row = -1
vaSpread1.Col = 4: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 6: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 7: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 8: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 9: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 10: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
Double1(0).UseSeparator = True
Double1(1).UseSeparator = True
Double1(2).UseSeparator = True
Double1(3).UseSeparator = True
Double1(4).UseSeparator = True
Double1(0).DecimalPoint = vg_CDec: Double1(0).Separator = vg_CSep: Double1(0).DecimalPlaces = vg_DPr
Double1(1).DecimalPoint = vg_CDec: Double1(1).Separator = vg_CSep: Double1(1).DecimalPlaces = vg_DPr
Double1(2).DecimalPoint = vg_CDec: Double1(2).Separator = vg_CSep: Double1(2).DecimalPlaces = vg_DPr
Double1(3).DecimalPoint = vg_CDec: Double1(3).Separator = vg_CSep: Double1(3).DecimalPlaces = vg_DPr
Double1(4).DecimalPoint = vg_CDec: Double1(4).Separator = vg_CSep: Double1(4).DecimalPlaces = vg_DPr
Combo2(0).Clear
Combo2(0).AddItem "FACTURA" & Space(150) & "(FA)"
Combo2(0).AddItem "GUIA DE DEPACHO" & Space(150) & "(GD)"
Combo2(0).AddItem "NOTA DE CREDITO" & Space(150) & "(NC)"
Combo2(0).AddItem "NOTA DE DEBITO" & Space(150) & "(ND)"
Combo2(0).AddItem "BOLETA" & Space(150) & "(BO)"
Combo2(0).AddItem "BOLETA DE HONORARIOS" & Space(150) & "(BH)"
Combo2(0).ListIndex = -1
Combo2(1).Clear
RS1.Open "select * from a_bodega order by bod_nombre", vg_db, adOpenStatic
Do While Not RS1.EOF
    Combo2(1).AddItem RS1!bod_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!bod_codigo), 10) & ")"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
Combo2(1).ListIndex = -1
vaSpread1.Refresh
Gl_Mo_Botones Me, 7
Gl_Ac_Botones Me, 7, 1, ""
'-----Trae todos los impuestos disponibles
RS1.Open "select count(imp_codigo) as mayor from a_impuesto", vg_db, adOpenStatic
indice = RS1!mayor
RS1.Close: Set RS1 = Nothing
ReDim Preserve Impuestos(indice, 5)
indice = 1
RS1.Open "select * from a_impuesto", vg_db, adOpenStatic
Do While Not RS1.EOF
    'ReDim Preserve Impuestos(indice, 5)
    Impuestos(indice, 1) = RS1!imp_codigo
    Impuestos(indice, 2) = RS1!imp_nombre
    Impuestos(indice, 3) = RS1!imp_pctimp
    Impuestos(indice, 4) = 0
    Impuestos(indice, 5) = RS1!imp_inccos
    RS1.MoveNext: indice = indice + 1
Loop
RS1.Close: Set RS1 = Nothing
vg_Guias = ""
Nuevo_Registro
modo = "N"
'Ac_Botones
modo = "A"
End Sub
Private Sub Form_Resize()
If Me.WindowState = 2 Then
    Frame1.Left = (Me.Width \ 2) - (Frame1.Width \ 2)
    Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)
    Frame4.Left = (Me.Width \ 2) - (Frame4.Width \ 2)
Else
    Frame1.Left = 45
    Frame2.Left = 45
    Frame4.Left = 45
End If
End Sub

Private Sub fpText_Change(Index As Integer)
Select Case Index
Case 0
    fpayuda(Index).Caption = ""
End Select
End Sub

Private Sub fpText_GotFocus(Index As Integer)
Select Case Index
Case 0
    If Trim(fpText(0).Text) = "" Or vg_Dig = "N" Then Exit Sub
    fpText(0).Text = fg_DespintaRut(fpText(0).Text)
    fpText(0).Text = Mid(fpText(0).Text, 1, Len(Trim(fpText(0).Text)) - 1)
End Select
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Index = 1 Then Option1(1).SetFocus
SendKeys "{Tab}"
End Sub

Private Sub fpText_LostFocus(Index As Integer)
Select Case Index
Case 0
    If fpText(0).Text = "" Then Exit Sub
    fpText(0).Text = fg_RutDig(Trim(fpText(0).Text))
    RS4.Open "select * from b_proveedor where prv_codigo= '" & Trim(fpText(0).Text) & "'", vg_db, adOpenStatic
    If Not RS4.EOF Then
        fpText(0).Text = fg_PintaRut(fpText(0).Text)
        fpayuda(0).Caption = RS4!prv_nombre
        Double1_LostFocus 5
    Else
        fpText(0).Text = "": fpayuda(0).Caption = "" ':  Combo2(1).ListIndex = -1: Combo2(1).Clear
        RS4.Close: Set RS4 = Nothing: Exit Sub
    End If
    RS4.Close: Set RS4 = Nothing
End Select
End Sub

Private Sub Image1_Click(Index As Integer)
On Error Resume Next
vg_codigo = 0
Select Case Index
Case 0
    vg_left = fpayuda(Index).Left + 1920
    B_TabEst.LlenaDatos "b_proveedor", "prv_", "Proveedor", "Gen"
    B_TabEst.Show 1, Me
    Me.Refresh
    If Trim(vg_codigo) = "" Or Val(vg_codigo) = 0 Then Exit Sub
    fpText(Index).Text = fg_PintaRut(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    Double1_LostFocus 5
    If Combo2(0).Enabled = True Then Combo2(0).SetFocus
Case 1
    If Trim(fg_codigocbo(Combo2, 0, 2, "")) = "FA" Then
        vg_FDC = "GD"
        B_Guias.Cargar_DoctoGrilla "GD", "Guía de Despacho", fg_DespintaRut(fpText(0).Text)
    ElseIf Trim(fg_codigocbo(Combo2, 0, 2, "")) = "NC" Then
        vg_FDC = "SN"
        B_Guias.Cargar_DoctoGrilla "SN", "Solicitud NC", fg_DespintaRut(fpText(0).Text)
    End If
    vg_RDC = fg_DespintaRut(fpText(0).Text)
    B_Guias.Show 1
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
If Encontrado = False Then
    RS1.Open "select * from a_infcfcfofi where inf_tipo='" & IIf(Option1(0).Value = True, "F", "C") & "' and inf_feccie=0", vg_db, adOpenStatic
    
    If Not RS1.EOF Then
        Double1(6).Value = RS1!inf_numero
    Else
        RS2.Open "select max(inf_numero) as Mayor from a_infcfcfofi where inf_tipo='" & IIf(Option1(0).Value = True, "F", "C") & "'", vg_db, adOpenStatic
        Double1(6).Value = TipoDato(RS2!mayor, 0) + 1
        vg_db.BeginTrans
        vg_db.Execute "insert into a_infcfcfofi values ('" & IIf(Option1(0).Value = True, "F", "C") & "', " & Val(Double1(6).Value) & ", 0, '')"
        vg_db.CommitTrans
        RS2.Close: Set RS2 = Nothing
    End If
    IndiceOpt = Index
    RS1.Close: Set RS1 = Nothing
End If
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
IndiceOpt = Index
SendKeys "{Tab}"
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii <> 13 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 11
NumEnter = Len(Text1(0).Text) - InStr(1, Text1(0).Text, Chr(13)) + 1
If Text1(0).Text = "" Then Text1(0) = "  "
glosa = Text1(0).Text ' Mid$(Text1(0).Text, 1, Len(Text1(0).Text) - NumEnter)
vaSpread1.Text = glosa
vaSpread1.SetFocus: vaSpread1.SetActiveCell 4, vaSpread1.ActiveRow
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim diablq As Date
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim v_rut As String, v_tipo As String, v_bodega As Long, v_CtaCon As String, v_Rebaja As String, v_can As Double, v_precio As Double
Dim v_pctdes As Double, v_valdes As Double, v_total As Double, v_descrip As String, v_canrec As Double, v_prerec As Double
Dim StoA As Double, PPPA As Double, PreC As Double, Cdif As Long
On Error GoTo Man_Error
If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
Select Case Button.Index
Case 1
    Gl_Ac_Botones Me, 7, 1, ""
    Nuevo_Registro
    modo = "A"
Case 3 'Grabar
    'If Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(Date1(0).Text), "mm/yyyy") < Format(Now, "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If (Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(Date1(0).Text), "mm/yyyy") < Format(Now, "mm/yyyy")) Or Format(CDate(Date1(0).Text), "mm/yyyy") < Format(Month(Now) - 1 & "/" & Year(Now), "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    'Actualiza_Datos
    If modo <> "A" Then MsgBox "Insuficiencia de datos. Documento no grabado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    v_rut = fg_DespintaRut(fpText(0).Text)
    v_tipo = fg_codigocbo(Combo2, 0, 2, "")
    v_bodega = fg_codigocbo(Combo2, 1, 1, 0)
    v_fecemi = Format(Date1(0).Text, "dd/mm/yyyy")
    v_fecven = Format(Date1(1).Text, "dd/mm/yyyy")
    If Trim(fpText(0).Text) = "" Or Combo2(0).Text = "" Then MsgBox "No hay datos proveedor...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Val(Double1(5).Value) = 0 Then MsgBox "Debe ingresar Nº de documento...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Val(Double1(4).Value) = 0 And v_tipo <> "GD" Then MsgBox "Total documento no puede ser cero...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Date1(0).Text = "" Or Date1(1).Text = "" Then MsgBox "Debe seleccionar fechas...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Option1(IndiceOpt).Value = False Then MsgBox "Tipo de documento no valido...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Not fg_Check_Rut(fg_DespintaRut(fpText(0).Text)) Then MsgBox "El rut no es valido...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If vaSpread1.MaxRows = 0 Then MsgBox "Documento sin detalle...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Valida_DatosGrilla Then MsgBox "La cantidad o el precio de un producto es cero...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Graba documento...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    'suma los descuentos
    v_totdesc = 0
    For cont = 1 To vaSpread1.MaxRows
        vaSpread1.Row = cont: vaSpread1.Col = 7: v_totdesc = v_totdesc + Val(vaSpread1.Value)
    Next cont
    'Obtiene el parametro si el documento es FIFO o CFC
    Opcion = IIf(Option1(0).Value = True, "F", "C")
    '---Comienza Transaccion
    vg_db.BeginTrans
    vg_db.Execute "insert into b_totcompras (toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom) " & _
              "values ('" & v_rut & "', '" & v_tipo & "'," & Val(Double1(5).Text) & "," & v_bodega & ",'" & v_fecemi & "','" & v_fecven & "'," & v_totdesc & "," & Double1(1).Value & "," & Double1(0).Value & "," & Double1(2).Value & "," & Double1(3).Value & "," & Double1(4).Value & "," & Double1(4).Value & ",'" & Opcion & "'," & Double1(6).Value & ", '" & Trim(vg_Guias) & "', '" & Trim(fpText(1).Text) & "')"
    For cont = 1 To vaSpread1.MaxRows
        vaSpread1.Row = cont: vaSpread1.Col = 1
        If Trim(vaSpread1.Text) <> "" And IsNull((vaSpread1.Text)) = False Then
            vaSpread1.Col = 1: codigo_pro = Trim(vaSpread1.Text)
            vaSpread1.Col = 4: v_can = Val(vaSpread1.Value)
            vaSpread1.Col = 5: v_precio = Val(vaSpread1.Value)
            vaSpread1.Col = 6: v_pctdes = Val(vaSpread1.Value)
            vaSpread1.Col = 7: v_valdes = Val(vaSpread1.Value)
            vaSpread1.Col = 8: v_total = Val(vaSpread1.Value)
            vaSpread1.Col = 9: v_canrec = Val(vaSpread1.Value)
            vaSpread1.Col = 10: v_prerec = Val(vaSpread1.Value)
            vaSpread1.Col = 11: v_descrip = Trim(vaSpread1.Text)
            vaSpread1.Col = 12: v_CtaCon = Trim(vaSpread1.Text)
            vaSpread1.Col = 13: v_Rebaja = Trim(vaSpread1.Text)
            If (v_tipo = "FA" And Len(Trim(vg_Guias)) > 0) Or v_tipo = "NC" Then v_Rebaja = "N"
            vaSpread1.Col = 15: PreC = Val(vaSpread1.Text)
            vg_db.Execute "insert into b_detcompras (dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv) " & _
                          "values ('" & v_rut & "', '" & v_tipo & "', " & Double1(5).Value & ", " & cont & ", '" & codigo_pro & "', " & v_can & ", " & v_precio & ", " & v_pctdes & ", " & v_valdes & ", " & v_total & ", '" & TipoDato(v_descrip, "") & "', " & v_canrec & ", " & v_prerec & ", '" & v_Rebaja & "')"
            If v_Rebaja = "S" And Len(Trim(vg_Guias)) = 0 And v_tipo <> "NC" Then 'si el producto rebaja stock
                ValidaBod v_bodega, codigo_pro
    '------------PMP  ---------------------------------------------------
'                'Ingrediente
'                Dim PMP As Double, auxCanmer As Double, auxPropon As Double, coding As String, feccos As Long
'                RS2.Open "SELECT pro_coding, pro_facing FROM b_productos WHERE pro_codigo='" & codigo_pro & "'", vg_db, adOpenStatic
'                coding = ""
'                If Not RS2.EOF Then
'                    coding = IIf(IsNull(RS2!pro_coding), "", RS2!pro_coding)
'                    auxCanmer = 0: auxPropon = 0
'                    RS1.Open "SELECT sum(bod.bod_canmer) as canmer FROM b_productos pro, b_bodegas bod " & _
'                             "WHERE bod.bod_codpro=pro.pro_codigo and pro.pro_coding='" & coding & "'", vg_db, adOpenStatic
'                    If Not RS1.EOF Then auxCanmer = IIf(IsNull(RS1!canmer), 0, RS1!canmer)
'                    RS1.Close: Set RS1 = Nothing
'                    'RS1.Open "SELECT sum(ing_precos) as propon FROM b_ingrediente WHERE ing_codigo='" & coding & "'", vg_db, adOpenStatic
'                    'RS1.Open "SELECT sum(pro_propon/pro_facing) as propon FROM b_productos WHERE pro_coding='" & coding & "'", vg_db, adOpenStatic
'                    RS1.Open "SELECT sum((pro.pro_propon/pro.pro_facing)*bod_canmer) as propon FROM b_productos pro, b_bodegas bod " & _
'                             "WHERE pro.pro_codigo=bod.bod_codpro and pro.pro_coding='" & coding & "'", vg_db, adOpenStatic
'                    If Not RS1.EOF Then auxPropon = IIf(IsNull(RS1!propon), 0, RS1!propon)
'                    RS1.Close: Set RS1 = Nothing
'                    'PMP = Val(((auxPropon * auxCanmer) + ((PreC / RS2!pro_facing) * v_canrec)) / (auxCanmer + v_canrec))
'                    PMP = Val((auxPropon + ((PreC / RS2!pro_facing) * v_canrec)) / (auxCanmer + v_canrec))
'                    feccos = Val(Mid(v_fecemi, 7, 4) & Mid(v_fecemi, 4, 2) & Mid(v_fecemi, 1, 2))
'                    vg_db.Execute "update b_ingrediente set ing_feccos=" & feccos & ", ing_precos=" & PMP & " where ing_codigo='" & coding & "'"
'                End If
'                RS2.Close: Set RS2 = Nothing
'                'Producto
'                StoA = 0: PPPA = 0
'                RS1.Open "select * from b_productos where pro_codigo='" & codigo_pro & "'", vg_db, adOpenStatic
'                If Not RS1.EOF Then PPPA = RS1!pro_propon
'                RS1.Close: Set RS1 = Nothing
'                RS1.Open "select sum(bod_canmer) as stoact from b_bodegas where bod_codpro='" & codigo_pro & "' and bod_codbod=" & v_bodega, vg_db, adOpenStatic
'                StoA = TipoDato(RS1!stoact, 0)
'                RS1.Close: Set RS1 = Nothing
'                vg_db.Execute "update b_productos set pro_upreco=" & PreC & ", pro_fecuco=cdate('" & v_fecemi & "'), pro_propon=" & Round(((StoA * PPPA) + (v_canrec * PreC)) / (StoA + v_canrec), vg_DPr) & " where pro_codigo='" & codigo_pro & "'"
                
                
                
                Dim PMP As Double, auxCanmer As Double, auxPropon As Double, feccos As Long
                RS2.Open "Select pro_facing From b_productos Where pro_codigo='" & codigo_pro & "'", vg_db, adOpenStatic
                'PMP Ingrediente
                If Not RS2.EOF Then
                    auxCanmer = 0: auxPropon = 0
                    RS1.Open "Select Sum(bod.bod_canmer) As canmer From b_productos pro, b_bodegas bod " & _
                             "Where bod.bod_codpro=pro.pro_codigo And pro.pro_codigo='" & codigo_pro & "'", vg_db, adOpenStatic
                    If Not RS1.EOF Then auxCanmer = IIf(IsNull(RS1!canmer), 0, RS1!canmer)
                    RS1.Close: Set RS1 = Nothing
                    RS1.Open "Select Sum((pro.pro_propon/pro.pro_facing)*bod_canmer) as propon From b_productos pro, b_bodegas bod " & _
                             "Where pro.pro_codigo=bod.bod_codpro And pro.pro_codigo='" & codigo_pro & "'", vg_db, adOpenStatic
                    If Not RS1.EOF Then auxPropon = IIf(IsNull(RS1!propon), 0, RS1!propon)
                    RS1.Close: Set RS1 = Nothing
                    
                    PMP = Val((auxPropon + ((PreC / RS2!pro_facing) * v_canrec)) / (auxCanmer + v_canrec))
                    feccos = Val(Mid(v_fecemi, 7, 4) & Mid(v_fecemi, 4, 2) & Mid(v_fecemi, 1, 2))
                    vg_db.Execute "Update b_ingrediente ing, b_productosing pri Set ing_feccos=" & feccos & ", ing.ing_precos=" & PMP & " " & _
                                  "Where pri.pri_coding=ing.ing_codigo And pri.pri_codpro='" & codigo_pro & "'"
                    'Actuliza codigo compra de ultimo producto para ingrediente
                    vg_db.Execute "Update b_ingrediente ing, b_productosing pri Set ing.ing_codcom='" & codigo_pro & "' " & _
                                  "Where pri.pri_coding=ing.ing_codigo And pri.pri_codpro='" & codigo_pro & "'"
                    'PMP Producto
                    RS1.Open "Select Sum(bod.bod_canmer) As canmer From b_productos pro, b_bodegas bod " & _
                             "Where bod.bod_codpro=pro.pro_codigo And pro.pro_codigo='" & codigo_pro & "'", vg_db, adOpenStatic
                    If Not RS1.EOF Then auxCanmer = IIf(IsNull(RS1!canmer), 0, RS1!canmer)
                    RS1.Close: Set RS1 = Nothing
                    RS1.Open "Select pro_propon As propon From b_productos Where pro_codigo='" & codigo_pro & "'", vg_db, adOpenStatic
                    If Not RS1.EOF Then auxPropon = IIf(IsNull(RS1!propon), 0, RS1!propon)
                    RS1.Close: Set RS1 = Nothing
                    PMP = Round(((auxPropon * auxCanmer) + (PreC * v_canrec)) / (auxCanmer + v_canrec), vg_DPr)
                    vg_db.Execute "Update b_productos Set pro_propon=" & PMP & ", pro_upreco=" & PreC & ", pro_fecuco=cdate('" & v_fecemi & "') Where pro_codigo='" & codigo_pro & "'"
                    'vg_db.Execute "Update b_productos set pro_upreco=" & PreC & ", pro_fecuco=cdate('" & v_fecemi & "'), pro_propon=" & Round(((StoA * PPPA) + (v_canrec * PreC)) / (StoA + v_canrec), vg_DPr) & " where pro_codigo='" & codigo_pro & "'"
                
                End If
                RS2.Close: Set RS2 = Nothing
                
                
    '------------Fin PMP ------------------------------------------------
    '------------Actuliza Stock de bodega---------------------------
                vg_db.Execute "update b_bodegas set bod_canmer=bod_canmer+" & v_canrec & " where bod_codbod=" & v_bodega & " and bod_codpro='" & codigo_pro & "'"
    '------------Fin Actualiza Stock -----------------------------------------
    '------------Actuliza codigo de ultimo producto de compra  compra---------
                vg_db.Execute "update b_ingrediente set ing_codcom='" & codigo_pro & "' where ing_codigo='" & coding & "'"
    '------------Fin Actualiza -----------------------------------------------
            End If
         End If
    Next cont
    If Len(Trim(vg_Guias)) > 0 And (v_tipo = "FA" Or v_tipo = "NC") Then
        Dim StrImp As String, StrImpb As String
        StrImp = Trim(vg_Guias)
        Do While InStr(StrImp, ";") <> 0
            StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
            StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
            If v_tipo = "FA" Then
                vg_db.Execute "update b_totcompras set toc_docaso=" & Str(Double1(5).Value) & " where toc_tipdoc='GD' and toc_numdoc=" & Val(StrImpb)
            Else
                vg_db.Execute "update b_totcompras set toc_docaso=" & Str(Double1(5).Value) & " where toc_tipdoc='SN' and toc_numdoc=" & Val(StrImpb)
            End If
        Loop
    End If
    vg_db.Execute "insert into b_detcomprasimp (imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp)  " & _
            "select distinct a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, b.ipr_codimp, c.imp_pctimp, (a.dec_ptotal*(c.imp_pctimp/100)) " & _
            "from b_detcompras a, b_productosimp b, a_impuesto c where a.dec_codmer=b.ipr_codpro and b.ipr_codimp=c.imp_codigo and a.dec_rutpro='" & v_rut & "' And a.dec_tipdoc='" & v_tipo & "' and a.dec_numdoc=" & Val(Double1(5).Value)
    '----Chequeo de Diferencias---
    If v_tipo = "FA" Then
        Cdif = 0
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Row = i
            vaSpread1.Col = 4: v_can = Val(vaSpread1.Value)
            vaSpread1.Col = 5: v_precio = Val(vaSpread1.Value)
            vaSpread1.Col = 6: v_pctdes = Val(vaSpread1.Value)
            vaSpread1.Col = 9: v_canrec = Val(vaSpread1.Value)
            vaSpread1.Col = 10: v_prerec = Val(vaSpread1.Value)
            If v_can > v_canrec Or v_precio > v_prerec Then
                vaSpread1.Col = 16: vaSpread1.Value = 1
                vaSpread1.Col = 7: v_totdesc = v_totdesc + Val(vaSpread1.Text)
                'Sumo solamente las lineas con diferencias (cant =10 ;recib= 9 ;cantcal =1)
                SumaDiferencias i, (v_can - v_canrec), v_prerec, v_pctdes
                Cdif = Cdif + 1
            End If
        Next i
        If Cdif > 0 Then 'Si existen diferencias
            MsgBox "Documento con diferencias. Se emitira solicitud de nota de crédito... ", vbInformation + vbOKOnly, Msgtitulo
        Else
            modo = "A": 'Ac_Botones
            Gl_Ac_Botones Me, 7, 2, ""
            Frame1.Enabled = False
            Frame2.Enabled = False
            Frame3.Enabled = False
            Frame6.Enabled = False
            vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.Lock = True
            vg_RDC = fg_DespintaRut(fpText(0).Text)
            vg_TDC = fg_codigocbo(Combo2, 0, 2, "")
            vg_NDC = Val(Double1(5).Value)
            vg_NSOL = TipoDato(FolioSn, 0)
            fg_carga ""
            vg_db.CommitTrans
            I_DocProvee
            Exit Sub
        End If
        RS1.Open "select toc_numdoc from b_totcompras where toc_tipdoc='SN' order by toc_numdoc desc", vg_db, adOpenStatic
        If Not RS1.EOF Then
            FolioSn = RS1!toc_numdoc + 1
        Else
            FolioSn = 1
        End If
        RS1.Close: Set RS1 = Nothing
        
        vg_db.Execute "insert into b_totcompras (toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc,toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom) " & _
                      "values (" & "'" & v_rut & "', 'SN'," & FolioSn & "," & v_bodega & ",'" & v_fecemi & "','" & v_fecven & "'," & v_totdesc & "," & Double1(8).Value & "," & Double1(7).Value & "," & Double1(9).Value & "," & Double1(10).Value & "," & Double1(11).Value & "," & Double1(11).Value & ",'" & Opcion & "'," & Str(Double1(6).Value) & ", " & Str(Double1(5).Value) & ", '" & Trim(fpText(1).Text) & "')"
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Row = i
            vaSpread1.Col = 1: codigo_pro = Trim(vaSpread1.Text)
            vaSpread1.Col = 4: v_can = Val(vaSpread1.Value)
            vaSpread1.Col = 5: v_precio = Val(vaSpread1.Value)
            vaSpread1.Col = 6: v_pctdes = Val(vaSpread1.Value)
            vaSpread1.Col = 7: v_valdes = Val(vaSpread1.Value)
            vaSpread1.Col = 8: v_total = Val(vaSpread1.Value)
            vaSpread1.Col = 9: v_canrec = Val(vaSpread1.Value)
            vaSpread1.Col = 10: v_prerec = Val(vaSpread1.Value)
            vaSpread1.Col = 11: v_descrip = Trim(vaSpread1.Text)
            vaSpread1.Col = 12: v_CtaCon = Trim(vaSpread1.Text)
            vaSpread1.Col = 15: PreC = Val(vaSpread1.Text)
            v_Rebaja = "S"
            If v_can > v_canrec Or v_precio > v_prerec Then
                vg_db.Execute "insert into b_detcompras (dec_rutpro,dec_tipdoc,dec_numdoc,dec_numlin,dec_codmer,dec_canmer,dec_precom,dec_pctdes,dec_valdes,dec_ptotal,dec_descri,dec_canrec,dec_prerec, dec_mueinv)   values ('" & v_rut & "','" & "SN" & "'," & FolioSn & "," & i & ",'" & codigo_pro & "'," & v_can & "," & v_precio & "," & v_pctdes & "," & v_valdes & "," & v_total & ",'" & TipoDato(v_descrip, "") & "'," & v_canrec & "," & v_prerec & ", '" & v_Rebaja & "')"
            End If
        Next i
    End If
    vg_db.CommitTrans
    '----Fin Chequeo de Diferencias---
    'MsgBox "Grabado ok...", vbInformation + vbOKOnly, MsgTitulo
    modo = "A": 'Ac_Botones
    Gl_Ac_Botones Me, 7, 2, ""
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame6.Enabled = False
    vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.Lock = True
    vg_RDC = fg_DespintaRut(fpText(0).Text)
    vg_TDC = fg_codigocbo(Combo2, 0, 2, "")
    vg_NDC = Val(Double1(5).Value)
    vg_NSOL = TipoDato(FolioSn, 0)
    fg_carga ""
    I_DocProvee
Case 5 'Borrar
    'If Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(Date1(0).Text), "mm/yyyy") < Format(Now, "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If (Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(Date1(0).Text), "mm/yyyy") < Format(Now, "mm/yyyy")) Or Format(CDate(Date1(0).Text), "mm/yyyy") < Format(Month(Now) - 1 & "/" & Year(Now), "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Borra_Datos
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame6.Enabled = True
    vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.Lock = False
    Nuevo_Registro
    Gl_Ac_Botones Me, 7, 1, ""
Case 8
    If Trim(fpText(0).Text) = "" Or Combo2(0).ListIndex < 0 Or Val(Double1(5).Value) = 0 Then MsgBox "No existe documento...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    vg_RDC = fg_DespintaRut(fpText(0).Text)
    vg_TDC = fg_codigocbo(Combo2, 0, 2, "")
    vg_NDC = Val(Double1(5).Value)
    vg_NSOL = TipoDato(FolioSn, 0)
    fg_carga ""
    I_DocProvee
Case 11
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
Resume Next
If Err.Number = -2147467259 Then MsgBox "Documento ya existe...", vbExclamation + vbOKOnly, Msgtitulo: Resume Next: Exit Sub
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
vg_db.RollbackTrans
fg_descarga
End Sub

Private Sub Borra_Datos()
On Error GoTo Man_Error
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim Codigo As String, v_bodega As Long, v_cant As Long, actbod As Boolean, tipaux As String
'--Obtiene rut, tipo de docto y  Nº de Docto.
rut = fg_DespintaRut(fpText(0).Text)
Codigo = fg_codigocbo(Combo2, 0, 2, "")
num = Val(Double1(5).Value)
If num = 0 Then Exit Sub
'----Fin Obtiene Documento
If MsgBox("¿Desea Eliminar Documento?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    v_bodega = fg_codigocbo(Combo2, 1, 1, 0)
    actbod = True
    vg_db.BeginTrans
    RS1.Open "select toc_docaso from b_totcompras where toc_rutpro='" & rut & "' and toc_tipdoc='" & Codigo & "' and toc_numdoc=" & num, vg_db, adOpenStatic
    If Not RS1.EOF Then
        tipaux = ""
        If Codigo = "FA" Then tipaux = "GD"
        If Codigo = "NC" Then tipaux = "SN"
        vg_db.Execute "update b_totcompras set toc_docaso=''" & " where toc_rutpro='" & rut & "' and toc_tipdoc='" & tipaux & "' and toc_docaso='" & Trim(Str(num)) & "'"
        actbod = IIf(Len(Trim(RS1!toc_docaso)) > 0, False, True)
    End If
    RS1.Close: Set RS1 = Nothing
    vg_db.Execute "delete from b_detcomprasimp where imd_rutdoc='" & rut & "' and imd_tipdoc='" & Codigo & "' and imd_numdoc= " & num
    vg_db.Execute "delete from b_detcompras where dec_rutpro='" & rut & "' and dec_tipdoc='" & Codigo & "' and dec_numdoc=" & num
    vg_db.Execute "delete from b_totcompras where toc_rutpro='" & rut & "' and toc_tipdoc='" & Codigo & "' and toc_numdoc=" & num
    If actbod Then
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Row = i: vaSpread1.Col = 13
            If Trim(vaSpread1.Text) = "S" Then
                vaSpread1.Row = i
                vaSpread1.Col = 1: Codigo = Trim(vaSpread1.Text)
                vaSpread1.Col = 9: v_cant = Val(vaSpread1.Text)
                If Codigo <> "NC" Then
                    vg_db.Execute "update b_bodegas set bod_canmer=bod_canmer-" & v_cant & " where bod_codbod=" & v_bodega & " and bod_codpro='" & Codigo & "'"
                    vg_db.Execute "update b_totcompras set toc_docaso=''" & " where toc_rutpro='" & rut & "' and toc_tipdoc='" & Codigo & "' and toc_docaso='" & Trim(Str(num)) & "'"
                'Else
                '    vg_db.Execute "update b_bodegas set bod_canmer=bod_canmer+" & v_cant & " where bod_codbod=" & v_bodega & " and bod_codpro='" & codigo & "'"
                End If
            End If
        Next i
    End If
    vg_db.CommitTrans
    
    fpText(0).Text = "": fpayuda(0).Caption = ""
    Double1(0).Text = "": Double1(1).Text = "": Double1(2).Text = "": Double1(3).Text = "": Double1(4).Text = "": Double1(5).Text = "": Double1(6).Text = ""
    Date1(0).Text = "":    Date1(1).Text = ""
    vaSpread1.MaxRows = 0: vaSpread1.Col = 1: vaSpread1.Row = 1
    modo = "N": 'Ac_Botones
    'fpText(0).SetFocus
Exit Sub
Man_Error:
    If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, Msgtitulo: Exit Sub
    MsgBox "¡Datos no Eliminados!", vbCritical, Msgtitulo
    vg_db.RollbackTrans
    fg_descarga
End Sub

Function Ac_Botones()
Dim RS As New ADODB.Recordset
Dim incluir As String, alterar As String, eliminar As String, imprimir
'-----------------------------VALIDAR USUARIO-----------------
RS.Open "SELECT dpe.dpe_deragr, dpe.dpe_dermod, dpe.dpe_dereli, dpe.dpe_derimp " & _
         "FROM (a_perfil per INNER JOIN a_derechosperfil dpe ON per.per_codigo = dpe.dpe_codper) " & _
         "INNER JOIN a_usuarios usu ON per.per_codigo = usu.usu_perfil " & _
         "WHERE usu.usu_codigo='" & vg_NUsr & "' and dpe.dpe_codopc=" & Me.HelpContextID, vg_db, adOpenStatic
If Not RS.EOF Then
    Do While Not RS.EOF
        incluir = RS!dpe_deragr
        alterar = RS!dpe_dermod
        eliminar = RS!dpe_dereli
        imprimir = RS!dpe_derimp
        RS.MoveNext
    Loop
End If
RS.Close: Set RS = Nothing
'--------------------------------------------------------------
If (modo = "A") Then
    Toolbar1.Buttons(1).Visible = IIf(incluir = 1, True, False): Toolbar1.Buttons(2).Visible = IIf(incluir = 1, False, True)
    Toolbar1.Buttons(3).Visible = IIf(eliminar = 1, True, False): Toolbar1.Buttons(4).Visible = IIf(incluir = 1, False, True)
    Toolbar1.Buttons(6).Visible = IIf(incluir = 1, True, False): Toolbar1.Buttons(7).Visible = IIf(incluir = 1, False, True)
    Toolbar1.Buttons(8).Visible = IIf(incluir = 1, True, False): Toolbar1.Buttons(9).Visible = IIf(incluir = 1, False, True)
    Toolbar1.Buttons(11).Visible = IIf(imprimir = 1, True, False): Toolbar1.Buttons(12).Visible = IIf(imprimir = 1, False, True)
ElseIf modo = "" Then 'Incluir, Borrar
    Toolbar1.Buttons(1).Visible = IIf(incluir = 1, True, False): Toolbar1.Buttons(2).Visible = IIf(incluir = 1, False, True)
    Toolbar1.Buttons(3).Visible = IIf(eliminar = 1, True, False): Toolbar1.Buttons(4).Visible = IIf(incluir = 1, False, True)
    Toolbar1.Buttons(6).Visible = IIf(incluir = 1, False, True): Toolbar1.Buttons(7).Visible = IIf(incluir = 1, True, False)
    Toolbar1.Buttons(8).Visible = IIf(incluir = 1, False, True): Toolbar1.Buttons(9).Visible = IIf(incluir = 1, True, False)
    Toolbar1.Buttons(11).Visible = IIf(imprimir = 1, True, False): Toolbar1.Buttons(12).Visible = IIf(imprimir = 1, False, True)
    Nuevo_Registro
ElseIf modo = "N" Then
    Toolbar1.Buttons(1).Visible = IIf(incluir = 1, False, True): Toolbar1.Buttons(2).Visible = IIf(incluir = 1, True, False)
    Toolbar1.Buttons(3).Visible = IIf(eliminar = 1, True, False): Toolbar1.Buttons(4).Visible = IIf(incluir = 1, False, True)
    Toolbar1.Buttons(6).Visible = IIf(incluir = 1, True, False): Toolbar1.Buttons(7).Visible = IIf(incluir = 1, False, True)
    Toolbar1.Buttons(8).Visible = IIf(incluir = 1, True, False): Toolbar1.Buttons(9).Visible = IIf(incluir = 1, False, True)
    Toolbar1.Buttons(11).Visible = IIf(imprimir = 1, True, False): Toolbar1.Buttons(12).Visible = IIf(imprimir = 1, False, True)
End If
End Function

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long, cCta As String, cPro As String
Select Case Button.Index
Case 1
    vg_nombre = "": vg_codigo = ""
    vg_left = fpayuda(0).Left + 1920
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Gen"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Col = 1: vaSpread1.Row = i
        If Trim(vaSpread1.Text) = Trim(vg_codigo) Then MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Next i
    'obtiene el codigo de producto y el nombre
    'If Trim(vaSpread1.Text) = "" Then
    '    vaSpread1.Row = Row: vaSpread1.col = -1: vaSpread1.Text = ""
    '    Exit Sub
    'End If
    RS1.Open "select pro.pro_codigo, pro.pro_nombre, pro.pro_ctacon, uni.uni_nomcor, pro.pro_facing, pro.pro_facsto from b_productos pro, a_unidad uni " & _
             "where pro.pro_coduni=uni.uni_codigo and (pro.pro_codigo='" & vg_codigo & "' " & _
             ")", vg_db, adOpenStatic
    If Not RS1.EOF Then
        If RS1!pro_facing = 0 Or RS1!pro_facsto = 0 Then
            MsgBox "Factor del producto en cero...", vbExclamation + vbOKOnly, Msgtitulo
            RS1.Close: Set RS1 = Nothing
            Exit Sub
        End If
        cPro = RS1!pro_codigo
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.SetActiveCell 1, vaSpread1.MaxRows
        vaSpread1.Col = 1: vaSpread1.Text = RS1!pro_codigo
        vaSpread1.Col = 2: vaSpread1.Text = RS1!pro_nombre
        vaSpread1.Col = 3: vaSpread1.Text = RS1!uni_nomcor
        vaSpread1.Col = 12: vaSpread1.Text = RS1!pro_ctacon
        For i = 4 To 10: vaSpread1.Col = i: vaSpread1.Text = Format(0, fg_Pict(9, vg_DCa)): Next i
        If Trim(RS1!pro_ctacon) = "" Then MsgBox "El producto no tiene asosiada una cuenta contable...", vbExclamation + vbOKOnly, Msgtitulo: RS1.Close: Set RS1 = Nothing: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
        cCta = RS1!pro_ctacon
        encuentra = False
        RS2.Open "select par_codigo, par_valor from a_param where par_codigo in ('ctagastos', 'ctainsumo', 'ctalimdes','ctamovil')", vg_db, adOpenStatic
        If Not RS2.EOF Then
            Do While Not RS2.EOF
                v_inicio = Val(Mid(Trim(Fg_Puntocoma(RS2!par_valor)), 1, 6))
                v_final = Val(Mid(Trim(Fg_Puntocoma(RS2!par_valor)), Len(Trim(Fg_Puntocoma(RS2!par_valor))) - 5, 6))
                If v_inicio < v_final Then
                    For i = v_inicio To v_final
                        If i = Val(Trim(cCta)) Then
                            encuentra = True
                        End If
                    Next i
                Else
                    For i = v_final To v_inicio
                        If i = Val(Trim(cCta)) Then
                            encuentra = True
                        End If
                    Next i
                End If
                If encuentra = True Then Exit Do
                RS2.MoveNext
            Loop
            If encuentra = False Then MsgBox "No está identificado el tipo de cuenta contable...", vbExclamation + vbOKOnly, Msgtitulo: RS1.Close: Set RS1 = Nothing: RS2.Close: Set RS2 = Nothing: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
            vaSpread1.Col = 13: vaSpread1.Text = IIf(RS2!par_codigo = "ctagastos", "N", "S")
        End If
        RS2.Close: Set RS2 = Nothing
    End If
    RS1.Close: Set RS1 = Nothing
    vaSpread1.Col = 14: vaSpread1.Text = ""
    'Clipboard.Clear
    'Clipboard.SetText "select a.*, b.* from b_productosimp a, a_impuesto b where a.ipr_codimp=b.imp_codigo and a.ipr_codpro='" & cPro & "'"
    RS1.Open "select a.*, b.* from b_productosimp a, a_impuesto b where a.ipr_codimp=b.imp_codigo and a.ipr_codpro='" & cPro & "'", vg_db, adOpenStatic
    Do While Not RS1.EOF
        vaSpread1.Col = vaSpread1.ActiveCol: vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 14: vaSpread1.Text = vaSpread1.Text & Trim(Str(RS1!ipr_codimp)) & "&"
        vaSpread1.Col = 14: vaSpread1.Text = vaSpread1.Text & Trim(Str(RS1!imp_pctimp)) & "&"
        vaSpread1.Col = 14: vaSpread1.Text = vaSpread1.Text & Trim(Str(RS1!imp_inccos)) & ";"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    vaSpread1.SetActiveCell 4, vaSpread1.MaxRows
    If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
Case 2
    If vaSpread1.MaxRows = 0 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
End Select
End Sub
Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If Encontrado = True Then
    vaSpread1.Col = 11: vaSpread1.Row = vaSpread1.ActiveRow
    Text1(0).Text = vaSpread1.Text
End If
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
Dim Precio As Double, cantidad As Double, Descto As Double, Producto As String, DesctoFinal As Double, subtot As Double
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim Codigo As String, Switch As Boolean, encuentra As Boolean
On Error GoTo Error_Celda
With vaSpread1
    .Row = Row: .Col = Col
    'If .Lock = True Then Exit Sub
    .Col = 4: cantidad = Val(vaSpread1.Value)
    .Col = 5: Precio = Val(vaSpread1.Value)
    .Col = 6: Descto = Val(vaSpread1.Value)
    .Col = 7: DesctoFinal = Val(vaSpread1.Value)
    subtot = Precio * cantidad
    Select Case Col
    Case 4
        .Col = 4: canfac = Val(vaSpread1.Value): vaSpread1.Col = 9: vaSpread1.Value = canfac
    Case 5
        .Col = 5: valfac = Val(vaSpread1.Value): vaSpread1.Col = 10: vaSpread1.Value = valfac
    Case 6
        .Col = 7: vaSpread1.Value = Round(subtot * (Descto / 100), vg_DCa)
        .Col = 8: vaSpread1.Value = Round(subtot - ((cantidad * Precio) * (Descto / 100)), vg_DCa)
    Case 7
        .Col = 6: vaSpread1.Value = Round((DesctoFinal * 100) / subtot, vg_DCa)
        .Col = 6: Descto = Val(vaSpread1.Value)
        If Descto >= 99.99 Then
            Descto = 99.99
            .Col = 6: vaSpread1.Value = Descto
            .Col = 7: vaSpread1.Value = Round(subtot * (Descto / 100), vg_DCa)
        End If
        .Col = 8: vaSpread1.Value = Round(subtot - (subtot * (Descto / 100)), vg_DCa)
    End Select
    If Col = 4 Or Col = 5 Then
        .Col = 7: vaSpread1.Value = Round(subtot * (Descto / 100), vg_DCa)
        .Col = 8: vaSpread1.Value = Round(subtot - (subtot * (Descto / 100)), vg_DCa)
    End If
    SumarTotales
End With
Exit Sub
Error_Celda:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbExclamation, Msgtitulo

End Sub
