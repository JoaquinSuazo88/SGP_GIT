VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_Nutric 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nutricionista"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabMaxWidth     =   4
      OLEDropMode     =   1
      TabCaption(0)   =   "Nutricionistas"
      TabPicture(0)   =   "M_Nutric.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_Nutric.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(16)"
      Tab(1).Control(1)=   "Label1(14)"
      Tab(1).Control(2)=   "Label1(13)"
      Tab(1).Control(3)=   "Label1(2)"
      Tab(1).Control(4)=   "Label1(15)"
      Tab(1).Control(5)=   "Label1(3)"
      Tab(1).Control(6)=   "Label1(4)"
      Tab(1).Control(7)=   "Label1(5)"
      Tab(1).Control(8)=   "fpText(5)"
      Tab(1).Control(9)=   "fpText(4)"
      Tab(1).Control(10)=   "fpText(3)"
      Tab(1).Control(11)=   "fpText(2)"
      Tab(1).Control(12)=   "fpText(0)"
      Tab(1).Control(13)=   "fpText(1)"
      Tab(1).Control(14)=   "Option1(1)"
      Tab(1).Control(15)=   "Option1(0)"
      Tab(1).Control(16)=   "vaSpread2"
      Tab(1).ControlCount=   17
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   4575
         Left            =   -68160
         TabIndex        =   11
         Top             =   1060
         Width           =   4095
         _Version        =   393216
         _ExtentX        =   7223
         _ExtentY        =   8070
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         ScrollBars      =   2
         SpreadDesigner  =   "M_Nutric.frx":0038
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   200
         TabIndex        =   18
         Top             =   1800
         Width           =   10905
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3615
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   10605
            _Version        =   393216
            _ExtentX        =   18706
            _ExtentY        =   6376
            _StockProps     =   64
            ButtonDrawMode  =   1
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   6
            MaxRows         =   20
            SpreadDesigner  =   "M_Nutric.frx":1912
            VisibleCols     =   2
            VisibleRows     =   15
            ScrollBarTrack  =   1
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   2160
         TabIndex        =   14
         Top             =   600
         Width           =   6615
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "M_Nutric.frx":1E0F
            Left            =   1680
            List            =   "M_Nutric.frx":1E19
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   240
            Width           =   2865
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Left            =   1680
            TabIndex        =   0
            Top             =   600
            Width           =   2895
            _Version        =   196608
            _ExtentX        =   5106
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
            NoSpecialKeys   =   3
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
         Begin VB.Label Label1 
            Caption         =   "Buscar Columna"
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
            Index           =   11
            Left            =   180
            TabIndex        =   17
            Top             =   345
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "B"
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
            Left            =   4680
            TabIndex        =   16
            Top             =   675
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label Label1 
            Caption         =   "Buscar Texto"
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
            Index           =   0
            Left            =   180
            TabIndex        =   15
            Top             =   675
            Width           =   1410
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Activo"
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
         Index           =   0
         Left            =   -72240
         TabIndex        =   4
         Top             =   960
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bloqueado"
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
         Index           =   1
         Left            =   -69840
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   -74760
         TabIndex        =   6
         Top             =   1800
         Width           =   3045
         _Version        =   196608
         _ExtentX        =   5371
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
         MaxLength       =   30
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   -74760
         TabIndex        =   3
         Top             =   960
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
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
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   1
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   2
         Left            =   -71640
         TabIndex        =   7
         Top             =   1800
         Width           =   3045
         _Version        =   196608
         _ExtentX        =   5371
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   3
         Left            =   -74760
         TabIndex        =   8
         Top             =   2640
         Width           =   3045
         _Version        =   196608
         _ExtentX        =   5371
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   4
         Left            =   -71640
         TabIndex        =   9
         Top             =   2640
         Width           =   3045
         _Version        =   196608
         _ExtentX        =   5371
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
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   1
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   5
         Left            =   -74760
         TabIndex        =   10
         Top             =   3360
         Width           =   3045
         _Version        =   196608
         _ExtentX        =   5371
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
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   1
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
         PasswordChar    =   "*"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Departamento Asociado __________________"
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
         Left            =   -68160
         TabIndex        =   26
         Top             =   720
         Width           =   3990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password ____________________"
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
         Left            =   -74760
         TabIndex        =   25
         Top             =   3120
         Width           =   2985
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario PDA _________________"
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
         Left            =   -71640
         TabIndex        =   24
         Top             =   2400
         Width           =   2940
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código __________"
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
         Index           =   15
         Left            =   -74760
         TabIndex        =   23
         Top             =   720
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre _____________________"
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
         Left            =   -74760
         TabIndex        =   22
         Top             =   1560
         Width           =   2925
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado ____________________________"
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
         Left            =   -72240
         TabIndex        =   21
         Top             =   720
         Width           =   3600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Paterno ______________"
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
         Left            =   -71640
         TabIndex        =   20
         Top             =   1560
         Width           =   2940
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno ______________"
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
         Left            =   -74760
         TabIndex        =   19
         Top             =   2400
         Width           =   2970
      End
   End
End
Attribute VB_Name = "M_Nutric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim modo As String, codigo As String, v_rut As String
Dim Msgtitulo As String, est As Boolean

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 6855
Me.Width = 11640
est = True
Msgtitulo = "Nutricionista"
fg_centra Me
SSTab1.Tab = 0
Combo1(0).ListIndex = 1
modo = ""
Toolbar1.ImageList = Partida.IL1
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
MoverNutricionistas
est = False
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then SSTab1.Move 0, Toolbar1.Height, ScaleWidth, ScaleHeight - Toolbar1.Height
End Sub

Private Sub fpText_Change(Index As Integer)
If est Or Toolbar1.Buttons(12).Visible = True Then Exit Sub
SSTab1.TabEnabled(0) = False
modo = "M"
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpTnombre_Change()
If LimpiaDato(Trim(fptnombre.text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1(0).ItemData(Combo1(0).ListIndex) = 0 Then
   RS.Open "SELECT * FROM b_nutricionistas WHERE UCASE(nut_codigo) LIKE '%" & UCase(LimpiaDato(fptnombre.text)) & "%' ORDER BY nut_codigo", vg_db, adOpenStatic
ElseIf Combo1(0).ItemData(Combo1(0).ListIndex) = 1 Then
   RS.Open "SELECT * FROM b_nutricionistas WHERE UCASE(nut_nombre) LIKE '%" & UCase(LimpiaDato(fptnombre.text)) & "%' ORDER BY nut_nombre", vg_db, adOpenStatic
End If
i = 1: vaSpread1.MaxRows = RS.RecordCount
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.Row = i
      vaSpread1.Col = 1: vaSpread1.text = IIf(IsNull(RS!nut_codigo), "", Trim(RS!nut_codigo))
      vaSpread1.Col = 2: vaSpread1.Lock = True: vaSpread1.text = IIf(IsNull(RS!nut_estado), "", RS!nut_estado)
      vaSpread1.Col = 3: vaSpread1.text = IIf(IsNull(RS!nut_usuario), "", Trim(RS!nut_usuario))
      vaSpread1.Col = 4: vaSpread1.text = IIf(IsNull(RS!nut_appaterno), "", Trim(RS!nut_appaterno))
      vaSpread1.Col = 5: vaSpread1.text = IIf(IsNull(RS!nut_apmaterno), "", Trim(RS!nut_apmaterno))
      vaSpread1.Col = 6: vaSpread1.text = IIf(IsNull(RS!nut_nombre), "", Trim(RS!nut_nombre))
      RS.MoveNext: i = i + 1
   Loop
   vaSpread1.Row = 1: vaSpread1.Col = 1: codigo = Trim(vaSpread1.text)
   MoverDetNutricionistas codigo
   SSTab1.TabEnabled(1) = True
   modo = ""
   Gl_Ac_Botones Me, 1, 1, modo
Else
   SSTab1.TabEnabled(1) = False
End If
RS.Close: Set RS = Nothing
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
End Sub

Private Sub Option1_Click(Index As Integer)
If est Or Toolbar1.Buttons(12).Visible = True Then Exit Sub
SSTab1.TabEnabled(0) = False
modo = "M"
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    SSTab1.TabEnabled(IIf(SSTab1.Tab = 0, 1, 0)) = False
    est = True: modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 1
    For i = 0 To 5
        fpText(i).Enabled = True: fpText(i).text = ""
    Next i
    For i = 1 To vaSpread2.MaxRows
        vaSpread2.Row = i
        vaSpread2.Col = 1
        vaSpread2.text = "0"
    Next i
    fpText(0).Enabled = False
    Option1(0).Value = True
    est = False
Case 3
    If vaSpread1.MaxRows < 1 Then Exit Sub
    fpText(0).Enabled = False
    SSTab1.TabEnabled(0) = False
    SSTab1.Tab = 1: SSTab1.TabEnabled(1) = True
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
Case 5 'Borrar lista
    Borra_Dato
Case 7 'Actualizar lista
    modo = ""
    SSTab1.Tab = 0
    MoverNutricionistas
Case 10 'Cancelar
    Cancela_Datos
Case 12 'Grabar datos
    Actualiza_Datos
Case 15 'Imprimir
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, , Msgtitulo: Exit Sub
    I_Nutricionistas
Case 18 'Salir
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub MoverNutricionistas()
On Error GoTo Man_Error
fg_carga ""
vaSpread2.MaxRows = 0
RS.Open "SELECT * FROM a_departamento WHERE dep_estado='0' ORDER BY dep_codigo", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      vaSpread2.Col = 1: vaSpread2.text = "0"
      vaSpread2.Col = 2: vaSpread2.text = RS!dep_codigo
      vaSpread2.Col = 3: vaSpread2.text = Trim(RS!dep_nombre)
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
vaSpread1.MaxRows = 0
RS.Open "SELECT * FROM b_nutricionistas ORDER BY nut_nombre", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 1: vaSpread1.text = IIf(IsNull(RS!nut_codigo), "", Trim(RS!nut_codigo))
      vaSpread1.Col = 2: vaSpread1.Lock = True: vaSpread1.text = IIf(IsNull(RS!nut_estado), "", RS!nut_estado)
      vaSpread1.Col = 3: vaSpread1.text = IIf(IsNull(RS!nut_usuario), "", Trim(RS!nut_usuario))
      vaSpread1.Col = 4: vaSpread1.text = IIf(IsNull(RS!nut_appaterno), "", Trim(RS!nut_appaterno))
      vaSpread1.Col = 5: vaSpread1.text = IIf(IsNull(RS!nut_apmaterno), "", Trim(RS!nut_apmaterno))
      vaSpread1.Col = 6: vaSpread1.text = IIf(IsNull(RS!nut_nombre), "", Trim(RS!nut_nombre))
      RS.MoveNext
   Loop
   vaSpread1.Row = 1: vaSpread1.Col = 1: codigo = Trim(vaSpread1.text)
   MoverDetNutricionistas codigo
   Gl_Ac_Botones Me, 1, 1, modo
   SSTab1.TabEnabled(1) = True
Else
   SSTab1.Tab = 0
   SSTab1.TabEnabled(1) = False
   modo = "NE"
   Gl_Ac_Botones Me, 1, 2, modo
End If
RS.Close: Set RS = Nothing
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
fptnombre.text = ""
fg_descarga
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"
End Sub

Private Sub MoverDetNutricionistas(codigo As String)
fg_carga ""
Dim i As Long
est = True
RS1.Open "SELECT * FROM b_nutricionistas WHERE nut_codigo=" & codigo & "", vg_db, adOpenStatic
If Not RS1.EOF Then
   fpText(0).text = Trim(RS1!nut_codigo)
   fpText(1).text = IIf(IsNull(RS1!nut_nombre), "", Trim(RS1!nut_nombre))
   fpText(2).text = IIf(IsNull(RS1!nut_appaterno), "", Trim(RS1!nut_appaterno))
   fpText(3).text = IIf(IsNull(RS1!nut_apmaterno), "", Trim(RS1!nut_apmaterno))
   fpText(4).text = IIf(IsNull(RS1!nut_usuario), "", Trim(RS1!nut_usuario))
   fpText(5).text = IIf(IsNull(RS1!nut_password), "", fg_Desencripta(Trim(RS1!nut_password)))
   Option1(IIf(IsNull(RS1!nut_estado) Or RS1!nut_estado = "1", 1, 0)).Value = True
End If
RS1.Close: Set RS1 = Nothing
RS1.Open "SELECT b.* " & _
         "FROM a_departamento a, b_nutricionistasdeptos b " & _
         "WHERE a.dep_codigo=b.nud_coddep " & _
         "AND   b.nud_codnut=" & codigo & "", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      For i = 1 To vaSpread2.MaxRows
          vaSpread2.Row = i
          vaSpread2.Col = 2
          If vaSpread2.text = RS1!nud_coddep Then
             vaSpread2.Col = 1: vaSpread2.text = "1"
             Exit For
          End If
      Next i
      RS1.MoveNext
   Loop
Else
   For i = 1 To vaSpread2.MaxRows
       vaSpread2.Row = i
       vaSpread2.Col = 2
       vaSpread2.Col = 1: vaSpread2.text = "0"
   Next i
End If
RS1.Close: Set RS1 = Nothing
est = False
fg_descarga
End Sub

Private Sub Borra_Dato()
On Error GoTo Man_Error
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: codigo = vaSpread1.text
If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
vg_db.BeginTrans
vg_db.Execute "DELETE * FROM b_nutricionistasdeptos WHERE nud_codnut=" & codigo & ""
vg_db.Execute "DELETE * FROM b_nutricionistas WHERE nut_codigo=" & codigo & ""
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.DeleteRows vaSpread1.Row, 1
vaSpread1.MaxRows = vaSpread1.MaxRows - 1
vaSpread1.Row = vaSpread1.MaxRows
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
If vaSpread1.MaxRows < 1 Then
   SSTab1.TabEnabled(1) = False
   SSTab1.Tab = 0
   modo = "NE"
Else
   modo = ""
   SSTab1.TabEnabled(1) = True
   SSTab1.Tab = 0
End If
vg_db.CommitTrans
Gl_Ac_Botones Me, 1, 1, modo

Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Cancela_Datos()
If MsgBox("Cancela registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
SSTab1.TabEnabled(0) = True
RS.Open "SELECT COUNT(*) AS nreg FROM b_nutricionistas WHERE UCASE(nut_codigo) LIKE '%" & UCase(("")) & "%'", vg_db, adOpenStatic
If RS.EOF Or RS!NReg = 0 Then RS.Close: Set RS = Nothing: SSTab1.TabEnabled(1) = False: modo = "NE": SSTab1.Tab = 0: Gl_Ac_Botones Me, 1, 2, modo: Exit Sub
RS.Close: Set RS = Nothing
If vaSpread1.MaxRows > 0 Then
   SSTab1.TabEnabled(0) = True
Else
   SSTab1.TabEnabled(1) = False
End If
SSTab1.Tab = 0
modo = ""
Gl_Ac_Botones Me, 1, 1, modo
End Sub

Private Sub Actualiza_Datos()
On Error GoTo Man_Error
v_rut = fg_DespintaRut(fpText(0).text)
If modo = "A" Then
    If Trim(fpText(1).text) = "" Or Trim(fpText(2).text) = "" Or Trim(fpText(3).text) = "" Or Trim(fpText(4).text) = "" Or Trim(fpText(4).text) = "" Then MsgBox "Faltan datos importantes para identificar el cliente...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Dim cod As Long
    RS.Open "SELECT nut_codigo FROM b_nutricionistas ORDER BY nut_codigo DESC", vg_db, adOpenStatic
    If Not RS.EOF Then RS.MoveFirst: cod = RS!nut_codigo + 1 Else cod = 1
    RS.Close: Set RS = Nothing
    vg_db.BeginTrans
    vg_db.Execute "INSERT INTO b_nutricionistas (nut_codigo, nut_nombre, nut_appaterno, " & _
                  "nut_apmaterno, nut_usuario, nut_password, nut_estado) VALUES (" & cod & ", " & _
                  "'" & LimpiaDato(Trim(fpText(1).text)) & "', '" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                  "'" & LimpiaDato(Trim(fpText(3).text)) & "', '" & LimpiaDato(Trim(fpText(4).text)) & "', " & _
                  "'" & fg_Encripta(LimpiaDato(Trim(fpText(5).text))) & "', '" & IIf(Option1(0).Value = True, 0, 1) & "')"
    vg_db.Execute "DELETE * FROM b_nutricionistasdeptos WHERE nud_codnut=" & cod & ""
    For i = 1 To vaSpread2.MaxRows
        vaSpread2.Row = i
        vaSpread2.Col = 1
        If vaSpread2.text = "1" Then
           vaSpread2.Col = 2
           vg_db.Execute "INSERT INTO b_nutricionistasdeptos (nud_codnut, nud_coddep) VALUES (" & cod & ", " & Val(vaSpread2.text) & ")"
        End If
    Next i
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
    codigo = cod
    vg_db.CommitTrans
Else
    If Trim(fpText(0).text) = "" Or Trim(fpText(1).text) = "" Or Trim(fpText(2).text) = "" Or Trim(fpText(3).text) = "" Or Trim(fpText(4).text) = "" Or Trim(fpText(4).text) = "" Then MsgBox "Faltan datos importantes para identificar el cliente...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    vg_db.BeginTrans
    vg_db.Execute "UPDATE b_nutricionistas SET nut_nombre='" & LimpiaDato(Trim(fpText(1).text)) & "', nut_appaterno='" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                  "nut_apmaterno='" & LimpiaDato(Trim(fpText(3).text)) & "', nut_usuario='" & LimpiaDato(Trim(fpText(4).text)) & "', " & _
                  "nut_password='" & fg_Encripta(LimpiaDato(Trim(fpText(5).text))) & "', nut_estado='" & IIf(Option1(0).Value = True, 0, 1) & "' " & _
                  "WHERE nut_codigo=" & codigo & ""
    vg_db.Execute "DELETE * FROM b_nutricionistasdeptos WHERE nud_codnut=" & codigo & ""
    For i = 1 To vaSpread2.MaxRows
        vaSpread2.Row = i
        vaSpread2.Col = 1
        If vaSpread2.text = "1" Then
           vaSpread2.Col = 2
           vg_db.Execute "INSERT INTO b_nutricionistasdeptos (nud_codnut, nud_coddep) VALUES (" & codigo & ", " & Val(vaSpread2.text) & ")"
        End If
    Next i
    vg_db.CommitTrans
End If
RS.Open "SELECT * FROM b_nutricionistas WHERE nut_codigo=" & codigo & "", vg_db, adOpenStatic
If Not RS.EOF Then
   vaSpread1.Col = 1: vaSpread1.text = IIf(IsNull(RS!nut_codigo), "", Trim(RS!nut_codigo))
   vaSpread1.Col = 2: vaSpread1.Lock = True: vaSpread1.text = IIf(IsNull(RS!nut_estado), "", Trim(RS!nut_estado))
   vaSpread1.Col = 3: vaSpread1.text = IIf(IsNull(RS!nut_usuario), "", Trim(RS!nut_usuario))
   vaSpread1.Col = 4: vaSpread1.text = IIf(IsNull(RS!nut_appaterno), "", Trim(RS!nut_appaterno))
   vaSpread1.Col = 5: vaSpread1.text = IIf(IsNull(RS!nut_apmaterno), "", Trim(RS!nut_apmaterno))
   vaSpread1.Col = 6: vaSpread1.text = IIf(IsNull(RS!nut_nombre), "", Trim(RS!nut_nombre))
End If
RS.Close: Set RS = Nothing
If SSTab1.Tab = 1 Then
   vaSpread1.SortKey(1) = 2
   vaSpread1.SortKeyOrder(1) = 1
   vaSpread1.Sort 1, 1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
   SSTab1.TabEnabled(0) = True
   If vaSpread1.MaxRows < 1 Then
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(2) = False
   Else
      SSTab1.TabEnabled(1) = True
      SSTab1.Tab = 0
   End If
   Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
Else
   If vaSpread1.MaxRows < 1 Then
      SSTab1.TabEnabled(1) = False
   Else
      SSTab1.TabEnabled(0) = True
      SSTab1.TabEnabled(1) = True
   End If
End If
est = False
modo = ""
Gl_Ac_Botones Me, 1, 1, modo
Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = Row: vaSpread1.Col = 1: codigo = Trim(vaSpread1.text)
MoverDetNutricionistas codigo
End Sub

Private Sub vaSpread2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
End Sub
