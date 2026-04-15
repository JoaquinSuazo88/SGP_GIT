VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_Pacien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paciente"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   17
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   18
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
      TabCaption(0)   =   "Pacientes"
      TabPicture(0)   =   "M_Pacien.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_Pacien.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Option1(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Option1(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fpText(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fpText(5)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fpText(4)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "fpText(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "fpText(6)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "fpText(2)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "fpText(3)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "fpLongInteger1(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "fpLongInteger1(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Date1(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Date1(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "fpayuda(2)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Image1(0)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "fpayuda(0)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1(18)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Image1(1)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "fpayuda(1)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "fpayuda(4)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label1(17)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label1(16)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label1(14)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label1(13)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label1(12)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label1(10)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label1(9)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label1(2)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Label1(3)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Label1(4)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Label1(5)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Label1(6)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Label1(15)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "sombra(1)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "sombra(0)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).ControlCount=   36
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   -65400
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1785
         Width           =   1485
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
         Left            =   -69960
         TabIndex        =   5
         Top             =   960
         Width           =   1335
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
         Left            =   -72360
         TabIndex        =   4
         Top             =   960
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   2160
         TabIndex        =   20
         Top             =   600
         Width           =   6615
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "M_Pacien.frx":0038
            Left            =   1680
            List            =   "M_Pacien.frx":0042
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
            TabIndex        =   23
            Top             =   675
            Width           =   1410
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
            TabIndex        =   22
            Top             =   675
            Visible         =   0   'False
            Width           =   120
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
            TabIndex        =   21
            Top             =   345
            Width           =   1425
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   200
         TabIndex        =   19
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
            MaxCols         =   14
            MaxRows         =   20
            SpreadDesigner  =   "M_Pacien.frx":0056
            VisibleCols     =   2
            VisibleRows     =   15
            ScrollBarTrack  =   1
         End
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
         Index           =   5
         Left            =   -74805
         TabIndex        =   13
         Top             =   3480
         Width           =   10980
         _Version        =   196608
         _ExtentX        =   19368
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   4
         Left            =   -69960
         TabIndex        =   11
         Top             =   2640
         Width           =   1380
         _Version        =   196608
         _ExtentX        =   2434
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
         MaxLength       =   10
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
         Index           =   6
         Left            =   -74760
         TabIndex        =   14
         Top             =   4320
         Width           =   10905
         _Version        =   196608
         _ExtentX        =   19235
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
         MaxLength       =   15
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
         Index           =   3
         Left            =   -68520
         TabIndex        =   8
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   -74760
         TabIndex        =   10
         Top             =   2640
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
         Left            =   -68400
         TabIndex        =   12
         Top             =   2640
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
      Begin EditLib.fpDateTime Date1 
         Height          =   345
         Index           =   0
         Left            =   -74760
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   5160
         Width           =   1710
         _Version        =   196608
         _ExtentX        =   3016
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
         Index           =   1
         Left            =   -72480
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   5160
         Width           =   2670
         _Version        =   196608
         _ExtentX        =   4710
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
         DateTimeFormat  =   2
         UserDefinedFormat=   ""
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   2
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
      Begin VB.Label fpayuda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   -69480
         TabIndex        =   43
         Top             =   5160
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   -67530
         Picture         =   "M_Pacien.frx":0750
         Top             =   2565
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   -67020
         TabIndex        =   41
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Régimen"
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
         Left            =   -68400
         TabIndex        =   40
         Top             =   2400
         Width           =   750
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   -73890
         Picture         =   "M_Pacien.frx":0A5A
         Top             =   2565
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   -73380
         TabIndex        =   38
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   -65340
         TabIndex        =   37
         Top             =   1830
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sexo"
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
         Left            =   -65400
         TabIndex        =   36
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno"
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
         Left            =   -68520
         TabIndex        =   35
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Paterno"
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
         TabIndex        =   34
         Top             =   1560
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
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
         Left            =   -72360
         TabIndex        =   33
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Origen de Dato"
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
         Index           =   12
         Left            =   -69480
         TabIndex        =   32
         Top             =   4920
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha y Hora de Alta"
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
         Left            =   -72480
         TabIndex        =   31
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ingreso"
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
         Index           =   9
         Left            =   -74760
         TabIndex        =   30
         Top             =   4920
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
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
         TabIndex        =   29
         Top             =   1560
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Paciente"
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
         Left            =   -74760
         TabIndex        =   28
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cama"
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
         Left            =   -69960
         TabIndex        =   27
         Top             =   2400
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prescripción Dietética"
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
         Left            =   -74805
         TabIndex        =   26
         Top             =   3240
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios / Observaciones"
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
         Left            =   -74760
         TabIndex        =   25
         Top             =   4080
         Width           =   2520
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
         Index           =   15
         Left            =   -74760
         TabIndex        =   24
         Top             =   720
         Width           =   315
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   -73335
         TabIndex        =   39
         Top             =   2685
         Width           =   3135
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   -66975
         TabIndex        =   42
         Top             =   2685
         Width           =   3135
      End
   End
End
Attribute VB_Name = "M_Pacien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim modo As String, codigo As String, v_rut As String
Dim Msgtitulo As String, est As Boolean

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Date1_Change(Index As Integer)
If est Or Toolbar1.Buttons(12).Visible = True Then Exit Sub
SSTab1.TabEnabled(0) = False
modo = "M"
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub Date1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 6855
Me.Width = 11640
est = True
Msgtitulo = "Paciente"
fg_centra Me
SSTab1.Tab = 0
modo = ""
Toolbar1.ImageList = Partida.IL1
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Date1(0).text = Format(Date, "dd/mm/yyyy")
'Date1(1).text = Format(Date, "dd/mm/yyyy")
Date1(1).text = "": Date1(1).Enabled = False
Combo1(0).ListIndex = 1
With Combo1(1)
    .Clear
    .AddItem "Masculino" & Space(150) & "(M)"
    .AddItem "Femenino" & Space(150) & "(F)"
End With
MoverPacientes
est = False
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then SSTab1.Move 0, Toolbar1.Height, ScaleWidth, ScaleHeight - Toolbar1.Height
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    RS.Open "SELECT * FROM a_regimen WHERE reg_codigo=" & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    fpayuda(0).Caption = ""
    If Not RS.EOF Then fpayuda(0).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
Case 1
    RS.Open "SELECT * FROM a_grupopaciente WHERE grp_codigo=" & Val(fpLongInteger1(1).Value) & "", vg_db, adOpenStatic
    fpayuda(1).Caption = ""
    If Not RS.EOF Then fpayuda(1).Caption = Trim(RS!grp_nombre)
    RS.Close: Set RS = Nothing
End Select
If Toolbar1.Buttons(12).Visible = True Then Exit Sub
SSTab1.TabEnabled(0) = False
modo = "M"
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_Change(Index As Integer)
If est Or Toolbar1.Buttons(12).Visible = True Then Exit Sub
SSTab1.TabEnabled(0) = False
modo = "M"
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub fpText_GotFocus(Index As Integer)
Select Case Index
Case 0
    If Trim(fpText(0).text) = "" Or vg_Dig = "N" Then Exit Sub
    fpText(0).text = fg_DespintaRut(fpText(0).text)
    fpText(0).text = Mid(fpText(0).text, 1, Len(Trim(fpText(0).text)) - 1)
End Select
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_LostFocus(Index As Integer)
Select Case Index
Case 0
    fpText(Index).text = UCase(fpText(Index).text)
    If Trim(fpText(0).text) = "" Or vg_Dig = "N" Then Exit Sub
    fpText(0).text = fg_RutDig(Trim(fpText(0).text))
    fpText(0).text = fg_PintaRut(fpText(0).text)
End Select
End Sub

Private Sub fpTnombre_Change()
If LimpiaDato(Trim(fpTnombre.text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1(0).ItemData(Combo1(0).ListIndex) = 0 Then
   RS.Open "SELECT b_pacientes.*, a_grupopaciente.grp_nombre, a_regimen.reg_nombre " & _
           "FROM (a_grupopaciente RIGHT JOIN b_pacientes ON a_grupopaciente.grp_codigo=b_pacientes.pac_codgrp) " & _
           "LEFT JOIN a_regimen ON b_pacientes.pac_codreg=a_regimen.reg_codigo WHERE UCASE(pac_codigo) LIKE '%" & UCase(LimpiaDato(fpTnombre.text)) & "%' ORDER BY pac_codigo", vg_db, adOpenStatic
ElseIf Combo1(0).ItemData(Combo1(0).ListIndex) = 1 Then
   RS.Open "SELECT b_pacientes.*, a_grupopaciente.grp_nombre, a_regimen.reg_nombre " & _
           "FROM (a_grupopaciente RIGHT JOIN b_pacientes ON a_grupopaciente.grp_codigo=b_pacientes.pac_codgrp) " & _
           "LEFT JOIN a_regimen ON b_pacientes.pac_codreg=a_regimen.reg_codigo WHERE UCASE(b_pacientes.pac_nombre) LIKE '%" & UCase(LimpiaDato(fpTnombre.text)) & "%' ORDER BY b_pacientes.pac_nombre", vg_db, adOpenStatic
End If
With vaSpread1
    i = 1: .MaxRows = RS.RecordCount
    If Not RS.EOF Then
       Do While Not RS.EOF
          .Row = i
          .Col = 1: .text = IIf(IsNull(RS!grp_nombre), "", Trim(RS!grp_nombre))
          .Col = 2: .text = IIf(IsNull(RS!pac_nrocam), "", Trim(RS!pac_nrocam))
          .Col = 3: .text = IIf(IsNull(RS!pac_codigo), "", fg_PintaRut(RS!pac_codigo))
          .Col = 4: .Lock = True: .text = IIf(IsNull(RS!pac_estado), "", Trim(RS!pac_estado))
          .Col = 5: .text = IIf(IsNull(RS!pac_appaterno), "", Trim(RS!pac_appaterno))
          .Col = 6: .text = IIf(IsNull(RS!pac_apmaterno), "", Trim(RS!pac_apmaterno))
          .Col = 7: .text = IIf(IsNull(RS!pac_nombre), "", Trim(RS!pac_nombre))
          .Col = 8: .text = IIf(Trim(RS!pac_sexo) = "M", "Masculino", "Femenino")
          .Col = 9: .text = IIf(IsNull(RS!reg_nombre), "", Trim(RS!reg_nombre))
          .Col = 10: .text = IIf(IsNull(RS!pac_fecing), "", Trim(RS!pac_fecing))
          .Col = 11: .text = IIf(IsNull(RS!pac_fecalt), "", Trim(RS!pac_fecalt))
          .Col = 12: .text = IIf(IsNull(RS!pac_presdiet), "", Trim(RS!pac_presdiet))
          .Col = 13: .text = IIf(IsNull(RS!pac_comentario), "", Trim(RS!pac_comentario))
          .Col = 14: .text = IIf(IsNull(RS!pac_origen), "", Trim(RS!pac_origen))
          RS.MoveNext: i = i + 1
       Loop
       SSTab1.TabEnabled(1) = True
       .Row = 1: .Col = 3: codigo = fg_DespintaRut(.text)
       MoverDetPacientes codigo
       modo = ""
       Gl_Ac_Botones Me, 1, 1, modo
    Else
       SSTab1.TabEnabled(1) = False
    End If
    RS.Close: Set RS = Nothing
    Label1(1).Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registros"
End With
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
    fpText(5).SetFocus
Case 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_grupopaciente", "grp_", "Grupo Paciente", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    fpLongInteger1(0).SetFocus
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
If est Or Toolbar1.Buttons(12).Visible = True Then Exit Sub
SSTab1.TabEnabled(0) = False
modo = "M"
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
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
    For i = 0 To 6
        If i < 11 Then fpText(i).Enabled = True: fpText(i).text = ""
    Next i
    Combo1(1).ListIndex = -1
    Option1(0).Value = True
    fpLongInteger1(0).Value = "": fpayuda(0).Caption = "": fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
    Date1(0).text = "": Date1(1).text = "": Date1(1).Enabled = False
    fpayuda(2).Caption = "SGP"
    fpText(0).SetFocus
    est = False
Case 3
    If vaSpread1.MaxRows < 1 Then Exit Sub
    est = True
    Date1(1).Enabled = True: Date1(1).text = ""
    SSTab1.TabEnabled(0) = False
    SSTab1.Tab = 1: SSTab1.TabEnabled(1) = True
    est = False: modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
Case 5 'Borrar lista
    Borra_DatoPacientes
Case 7 'Actualizar lista
    modo = ""
    SSTab1.Tab = 0
    MoverPacientes
Case 10 'Cancelar
    Cancela_Datos
Case 12 'Grabar datos
    Actualiza_Datos
Case 15 'Imprimir
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, , Msgtitulo: Exit Sub
    I_Pacientes
Case 18 'Salir
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub MoverPacientes()
On Error GoTo Man_Error
fg_carga ""
With vaSpread1
    .MaxRows = 0
    RS.Open "SELECT b_pacientes.*, a_grupopaciente.grp_nombre, a_regimen.reg_nombre " & _
            "FROM (a_grupopaciente RIGHT JOIN b_pacientes ON a_grupopaciente.grp_codigo=b_pacientes.pac_codgrp) " & _
            "LEFT JOIN a_regimen ON b_pacientes.pac_codreg=a_regimen.reg_codigo", vg_db, adOpenStatic
    If Not RS.EOF Then
       Do While Not RS.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1: .text = IIf(IsNull(RS!grp_nombre), "", Trim(RS!grp_nombre))
          .Col = 2: .text = IIf(IsNull(RS!pac_nrocam), "", Trim(RS!pac_nrocam))
          .Col = 3: .text = IIf(IsNull(RS!pac_codigo), "", fg_PintaRut(RS!pac_codigo))
          .Col = 4: .Lock = True: .text = IIf(IsNull(RS!pac_estado), "", Trim(RS!pac_estado))
          .Col = 5: .text = IIf(IsNull(RS!pac_appaterno), "", Trim(RS!pac_appaterno))
          .Col = 6: .text = IIf(IsNull(RS!pac_apmaterno), "", Trim(RS!pac_apmaterno))
          .Col = 7: .text = IIf(IsNull(RS!pac_nombre), "", Trim(RS!pac_nombre))
          .Col = 8: .text = IIf(Trim(RS!pac_sexo) = "M", "Masculino", "Femenino")
          .Col = 9: .text = IIf(IsNull(RS!reg_nombre), "", Trim(RS!reg_nombre))
          .Col = 10: .text = IIf(IsNull(RS!pac_fecing), "", Trim(RS!pac_fecing))
          .Col = 11: .text = IIf(IsNull(RS!pac_fecalt), "", Trim(RS!pac_fecalt))
          .Col = 12: .text = IIf(IsNull(RS!pac_presdiet), "", Trim(RS!pac_presdiet))
          .Col = 13: .text = IIf(IsNull(RS!pac_comentario), "", Trim(RS!pac_comentario))
          .Col = 14: .text = IIf(IsNull(RS!pac_origen), "", Trim(RS!pac_origen))
          RS.MoveNext
       Loop
       .Row = 1: .Col = 3: codigo = fg_DespintaRut(.text)
       MoverDetPacientes codigo
       Gl_Ac_Botones Me, 1, 1, modo
       SSTab1.TabEnabled(1) = True
    Else
       SSTab1.Tab = 0
       SSTab1.TabEnabled(1) = False
       modo = "NE"
       Gl_Ac_Botones Me, 1, 2, modo
    End If
    RS.Close: Set RS = Nothing
    Label1(1).Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registros"
End With
fpTnombre.text = ""
fg_descarga
Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"
End Sub

Private Sub MoverDetPacientes(codigo As String)
fg_carga ""
est = True
RS1.Open "SELECT b_pacientes.*, a_grupopaciente.grp_nombre, a_regimen.reg_nombre " & _
        "FROM (a_grupopaciente RIGHT JOIN b_pacientes ON a_grupopaciente.grp_codigo=b_pacientes.pac_codgrp) " & _
        "LEFT JOIN a_regimen ON b_pacientes.pac_codreg=a_regimen.reg_codigo WHERE b_pacientes.pac_codigo='" & codigo & "'", vg_db, adOpenStatic
If Not RS1.EOF Then
   fpText(0).text = fg_PintaRut(RS1!pac_codigo)
   fpText(1).text = IIf(IsNull(RS1!pac_nombre), "", Trim(RS1!pac_nombre))
   fpText(2).text = IIf(IsNull(RS1!pac_appaterno), "", Trim(RS1!pac_appaterno))
   fpText(3).text = IIf(IsNull(RS1!pac_apmaterno), "", Trim(RS1!pac_apmaterno))
   Combo1(1).ListIndex = IIf(IsNull(RS1!pac_sexo) Or Trim(RS1!pac_sexo) = "", -1, fg_buscacbostring(Combo1, 1, 1, ((RS1!pac_sexo))))
   fpLongInteger1(1).Value = IIf(IsNull(RS1!pac_codgrp), "", Trim(RS1!pac_codgrp))
   fpayuda(1).Caption = IIf(IsNull(RS1!grp_nombre), "", Trim(RS1!grp_nombre))
   fpText(4).text = IIf(IsNull(RS1!pac_nrocam), "", Trim(RS1!pac_nrocam))
   fpLongInteger1(0).Value = IIf(IsNull(RS1!pac_codreg), "", Trim(RS1!pac_codreg))
   fpayuda(0).Caption = IIf(IsNull(RS1!reg_nombre), "", Trim(RS1!reg_nombre))
   fpText(5).text = IIf(IsNull(RS1!pac_presdiet), "", Trim(RS1!pac_presdiet))
   fpText(6).text = IIf(IsNull(RS1!pac_comentario), "", Trim(RS1!pac_comentario))
   Date1(0).text = IIf(IsNull(RS1!pac_fecing), "", Trim(RS1!pac_fecing))
   Date1(1).text = IIf(IsNull(RS1!pac_fecalt), "", Trim(RS1!pac_fecalt))
   fpayuda(2).Caption = IIf(IsNull(RS1!pac_origen), "", Trim(RS1!pac_origen))
   Option1(IIf(IsNull(RS1!pac_estado) Or RS1!pac_estado = "1", 1, 0)).Value = True
End If
RS1.Close: Set RS1 = Nothing
fpText(0).Enabled = False
est = False
fg_descarga
End Sub

Private Sub Borra_DatoPacientes()
On Error GoTo Man_Error
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: codigo = vaSpread1.text
If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
vg_db.BeginTrans
vg_db.Execute "DELETE b_pacientes FROM b_pacientes WHERE pac_codigo='" & codigo & "'"
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
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Cancela_Datos()
If MsgBox("Cancela registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
SSTab1.TabEnabled(0) = True
RS.Open "SELECT COUNT(*) AS nreg FROM b_pacientes WHERE UCASE(pac_codigo) LIKE '%" & UCase(("")) & "%'", vg_db, adOpenStatic
If RS.EOF Or RS!nreg = 0 Then RS.Close: Set RS = Nothing: SSTab1.TabEnabled(1) = False: modo = "NE": SSTab1.Tab = 0: Gl_Ac_Botones Me, 1, 2, modo: Exit Sub
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
With vaSpread1
    If modo = "A" Then
        If Trim(fpText(0).text) = "" Or Trim(fpText(1).text) = "" Or Trim(fpText(2).text) = "" Or Trim(fpText(3).text) = "" Or Trim(fpText(4).text) = "" Or Combo1(1).ListIndex = -1 Or Trim(Date1(0).text) = "" Then MsgBox "Faltan datos importantes para identificar el cliente...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        If Not fg_Check_Rut(v_rut) Then MsgBox "El rut no es valido...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        RS.Open "SELECT * FROM b_pacientes WHERE pac_codigo='" & v_rut & "'", vg_db, adOpenStatic
        If Not RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Paciente existe", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        RS.Close: Set RS = Nothing
        vg_db.BeginTrans
        If Trim(Date1(1).text) = "" Then
           vg_db.Execute "INSERT INTO b_pacientes (pac_codigo, pac_nombre, pac_appaterno, " & _
                         "pac_apmaterno, pac_sexo, pac_codgrp, pac_nrocam, pac_codreg, pac_presdiet, " & _
                         "pac_comentario, pac_fecing, pac_fecalt, pac_origen, pac_estado) VALUES ('" & v_rut & "', " & _
                         "'" & LimpiaDato(Trim(fpText(1).text)) & "', '" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                         "'" & LimpiaDato(Trim(fpText(3).text)) & "', '" & fg_codigocbo(Combo1, 1, 1, "") & "', " & _
                         "" & IIf(Val(fpLongInteger1(1).Value) = 0, "Null", Val(fpLongInteger1(1).Value)) & ", '" & LimpiaDato(Trim(fpText(4).text)) & "', " & _
                         "" & IIf(Val(fpLongInteger1(0).Value) = 0, "Null", Val(fpLongInteger1(0).Value)) & ", '" & LimpiaDato(Trim(fpText(5).text)) & "', " & _
                         "'" & LimpiaDato(Trim(fpText(6).text)) & "', '" & Date1(0).text & "', " & IIf(Trim(Date1(1).text) = "", "Null", Date1(1).text) & ", 'SGP', '" & IIf(Option1(0).Value = True, "0", "1") & "')"
        Else
           vg_db.Execute "INSERT INTO b_pacientes (pac_codigo, pac_nombre, pac_appaterno, " & _
                         "pac_apmaterno, pac_sexo, pac_codgrp, pac_nrocam, pac_codreg, pac_presdiet, " & _
                         "pac_comentario, pac_fecing, pac_fecalt, pac_origen, pac_estado) VALUES ('" & v_rut & "', " & _
                         "'" & LimpiaDato(Trim(fpText(1).text)) & "', '" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                         "'" & LimpiaDato(Trim(fpText(3).text)) & "', '" & fg_codigocbo(Combo1, 1, 1, "") & "', " & _
                         "" & IIf(Val(fpLongInteger1(1).Value) = 0, "Null", Val(fpLongInteger1(1).Value)) & ", '" & LimpiaDato(Trim(fpText(4).text)) & "', " & _
                         "" & IIf(Val(fpLongInteger1(0).Value) = 0, "Null", Val(fpLongInteger1(0).Value)) & ", '" & LimpiaDato(Trim(fpText(5).text)) & "', " & _
                         "'" & LimpiaDato(Trim(fpText(6).text)) & "', '" & Date1(0).text & "', '" & Date1(1).text & "', 'SGP', '" & IIf(Option1(0).Value = True, "0", "1") & "')"
        End If
        .MaxRows = .MaxRows + 1: .Row = .MaxRows
        vg_db.CommitTrans
    Else
        If Trim(fpText(0).text) = "" Or Trim(fpText(1).text) = "" Or Trim(fpText(2).text) = "" Or Trim(fpText(3).text) = "" Or Trim(fpText(4).text) = "" Or Combo1(1).ListIndex = -1 Or Trim(Date1(0).text) = "" Then MsgBox "Faltan datos importantes para identificar el cliente...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        vg_db.BeginTrans
        If Trim(Date1(1).text) = "" Then
           vg_db.Execute "UPDATE b_pacientes SET pac_nombre='" & LimpiaDato(Trim(fpText(1).text)) & "', pac_appaterno='" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                         "pac_apmaterno='" & LimpiaDato(Trim(fpText(3).text)) & "', pac_sexo='" & fg_codigocbo(Combo1, 1, 1, "") & "', pac_codgrp=" & IIf(Val(fpLongInteger1(1).Value) = 0, "Null", Val(fpLongInteger1(1).Value)) & ", " & _
                         "pac_nrocam='" & LimpiaDato(Trim(fpText(4).text)) & "', pac_codreg=" & IIf(Val(fpLongInteger1(0).Value) = 0, "Null", Val(fpLongInteger1(0).Value)) & ", pac_presdiet='" & LimpiaDato(Trim(fpText(5).text)) & "', " & _
                         "pac_comentario='" & LimpiaDato(Trim(fpText(6).text)) & "', pac_fecing='" & Date1(0).text & "', pac_fecalt=" & IIf(Trim(Date1(1).text) = "", "Null", Date1(1).text) & ", pac_estado='" & IIf(Option1(0).Value = True, "0", "1") & "' " & _
                         "WHERE pac_codigo='" & v_rut & "'"
        Else
           vg_db.Execute "UPDATE b_pacientes SET pac_nombre='" & LimpiaDato(Trim(fpText(1).text)) & "', pac_appaterno='" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                         "pac_apmaterno='" & LimpiaDato(Trim(fpText(3).text)) & "', pac_sexo='" & fg_codigocbo(Combo1, 1, 1, "") & "', pac_codgrp=" & IIf(Val(fpLongInteger1(1).Value) = 0, "Null", Val(fpLongInteger1(1).Value)) & ", " & _
                         "pac_nrocam='" & LimpiaDato(Trim(fpText(4).text)) & "', pac_codreg=" & IIf(Val(fpLongInteger1(0).Value) = 0, "Null", Val(fpLongInteger1(0).Value)) & ", pac_presdiet='" & LimpiaDato(Trim(fpText(5).text)) & "', " & _
                         "pac_comentario='" & LimpiaDato(Trim(fpText(6).text)) & "', pac_fecing='" & Date1(0).text & "', pac_fecalt='" & Date1(1).text & "', pac_estado='" & IIf(Option1(0).Value = True, "1", "0") & "' " & _
                         "WHERE pac_codigo='" & v_rut & "'"
        End If
        .Row = .ActiveRow
        vg_db.CommitTrans
    End If
    RS.Open "SELECT b_pacientes.*, a_grupopaciente.grp_nombre, a_regimen.reg_nombre " & _
            "FROM (a_grupopaciente RIGHT JOIN b_pacientes ON a_grupopaciente.grp_codigo=b_pacientes.pac_codgrp) " & _
            "LEFT JOIN a_regimen ON b_pacientes.pac_codreg=a_regimen.reg_codigo WHERE b_pacientes.pac_codigo='" & v_rut & "'", vg_db, adOpenStatic
    If Not RS.EOF Then
       .Col = 1: .text = IIf(IsNull(RS!grp_nombre), "", Trim(RS!grp_nombre))
       .Col = 2: .text = IIf(IsNull(RS!pac_nrocam), "", Trim(RS!pac_nrocam))
       .Col = 3: .text = IIf(IsNull(RS!pac_codigo), "", fg_PintaRut(RS!pac_codigo))
       .Col = 4: .Lock = True: .text = IIf(IsNull(RS!pac_estado), "", Trim(RS!pac_estado))
       .Col = 5: .text = IIf(IsNull(RS!pac_appaterno), "", Trim(RS!pac_appaterno))
       .Col = 6: .text = IIf(IsNull(RS!pac_apmaterno), "", Trim(RS!pac_apmaterno))
       .Col = 7: .text = IIf(IsNull(RS!pac_nombre), "", Trim(RS!pac_nombre))
       .Col = 8: .text = IIf(Trim(RS!pac_sexo) = "M", "Masculino", "Femenino")
       .Col = 9: .text = IIf(IsNull(RS!reg_nombre), "", Trim(RS!reg_nombre))
       .Col = 10: .text = IIf(IsNull(RS!pac_fecing), "", Trim(RS!pac_fecing))
       .Col = 11: .text = IIf(IsNull(RS!pac_fecalt), "", Trim(RS!pac_fecalt))
       .Col = 12: .text = IIf(IsNull(RS!pac_presdiet), "", Trim(RS!pac_presdiet))
       .Col = 13: .text = IIf(IsNull(RS!pac_comentario), "", Trim(RS!pac_comentario))
       .Col = 14: .text = IIf(IsNull(RS!pac_origen), "", Trim(RS!pac_origen))
    End If
    RS.Close: Set RS = Nothing
    If SSTab1.Tab = 1 Then
       .SortKey(1) = 2
       .SortKeyOrder(1) = 1
       .Sort 1, 1, .MaxCols, .MaxRows, SortByRow
       SSTab1.TabEnabled(0) = True
       If .MaxRows < 1 Then
          SSTab1.TabEnabled(1) = False
          SSTab1.TabEnabled(2) = False
       Else
          SSTab1.TabEnabled(1) = True
          SSTab1.Tab = 0
       End If
       Label1(1).Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registros"
    Else
       If .MaxRows < 1 Then
          SSTab1.TabEnabled(1) = False
       Else
          SSTab1.TabEnabled(0) = True
          SSTab1.TabEnabled(1) = True
       End If
    End If
    est = False
    modo = ""
    Gl_Ac_Botones Me, 1, 1, modo
End With
Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = Row: vaSpread1.Col = 3: codigo = fg_DespintaRut(vaSpread1.text)
MoverDetPacientes codigo
End Sub
