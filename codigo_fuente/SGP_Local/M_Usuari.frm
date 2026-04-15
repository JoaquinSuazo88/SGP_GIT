VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form M_Usuari 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Usuario"
   ClientHeight    =   7755
   ClientLeft      =   2055
   ClientTop       =   1785
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   14
      Top             =   360
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   12726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   4
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Usuario"
      TabPicture(0)   =   "M_Usuari.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Shape1(0)"
      Tab(0).Control(3)=   "Shape1(1)"
      Tab(0).Control(4)=   "Label5(1)"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_Usuari.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(12)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(9)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(5)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(15)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(8)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "fpText(5)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "fpText(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "fpText(3)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "fpText(4)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "fpText(2)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "fpText(1)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Combo2(0)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "fpText(6)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Frame3"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Check3"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Text2"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Check1"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).ControlCount=   22
      Begin VB.CheckBox Check1 
         Caption         =   "Visualizar Password"
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
         Left            =   2640
         TabIndex        =   33
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   1695
         Left            =   480
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   5400
         Width           =   6045
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Desbloqueado"
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
         Left            =   4200
         TabIndex        =   1
         Top             =   960
         Width           =   1815
      End
      Begin VB.Frame Frame3 
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
         Height          =   4815
         Left            =   6720
         TabIndex        =   30
         Top             =   720
         Width           =   3855
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   4335
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   3615
            _Version        =   393216
            _ExtentX        =   6376
            _ExtentY        =   7646
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
            SpreadDesigner  =   "M_Usuari.frx":0038
         End
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   6
         Left            =   480
         TabIndex        =   7
         Top             =   4005
         Width           =   6045
         _Version        =   196608
         _ExtentX        =   10663
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
         MaxLength       =   50
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
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "M_Usuari.frx":18FE
         Left            =   480
         List            =   "M_Usuari.frx":1900
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4680
         Width           =   5295
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   -73200
         TabIndex        =   20
         Top             =   1800
         Width           =   7185
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   4455
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   6765
            _Version        =   393216
            _ExtentX        =   11933
            _ExtentY        =   7858
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
            MaxCols         =   2
            MaxRows         =   20
            OperationMode   =   3
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "M_Usuari.frx":1902
            VisibleCols     =   2
            VisibleRows     =   15
            ScrollBarTrack  =   1
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -72960
         TabIndex        =   15
         Top             =   600
         Width           =   6615
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "M_Usuari.frx":1D04
            Left            =   1680
            List            =   "M_Usuari.frx":1D0E
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   240
            Width           =   2865
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Left            =   1680
            TabIndex        =   11
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
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   11
            Left            =   210
            TabIndex        =   19
            Top             =   345
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "B"
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
            Left            =   4680
            TabIndex        =   18
            Top             =   675
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   17
            Top             =   675
            Width           =   1140
         End
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   1560
         Width           =   6045
         _Version        =   196608
         _ExtentX        =   10663
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
         MaxLength       =   50
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
         Left            =   480
         TabIndex        =   3
         Top             =   2160
         Width           =   1845
         _Version        =   196608
         _ExtentX        =   3254
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
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   3
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   4
         Left            =   3555
         TabIndex        =   5
         Top             =   2760
         Width           =   2940
         _Version        =   196608
         _ExtentX        =   5186
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
         MaxLength       =   50
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
         Left            =   480
         TabIndex        =   4
         Top             =   2760
         Width           =   2940
         _Version        =   196608
         _ExtentX        =   5186
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
         MaxLength       =   50
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
         Left            =   480
         TabIndex        =   0
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
         Left            =   480
         TabIndex        =   6
         Top             =   3360
         Width           =   1905
         _Version        =   196608
         _ExtentX        =   3351
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
         Caption         =   "Registrar Ticket"
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
         Left            =   480
         TabIndex        =   32
         Top             =   5160
         Width           =   1380
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H80000018&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   -70815
         Top             =   6750
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00D9D9FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   -72960
         Top             =   6750
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Bloqueado"
         Height          =   195
         Index           =   1
         Left            =   -72600
         TabIndex        =   31
         Top             =   6720
         Width           =   765
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000010&
         ForeColor       =   &H80000011&
         Height          =   300
         Left            =   525
         TabIndex        =   29
         Top             =   4755
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         Left            =   480
         TabIndex        =   28
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telefono"
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
         Left            =   480
         TabIndex        =   27
         Top             =   3120
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
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
         Left            =   3555
         TabIndex        =   26
         Top             =   2520
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Oficina"
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
         Left            =   480
         TabIndex        =   25
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
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
         Left            =   480
         TabIndex        =   24
         Top             =   1920
         Width           =   825
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
         Left            =   480
         TabIndex        =   23
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Perfil"
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
         Left            =   480
         TabIndex        =   22
         Top             =   4440
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Email"
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
         Left            =   480
         TabIndex        =   21
         Top             =   3720
         Width           =   465
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "M_Usuari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim RS2 As New ADODB.Recordset
Dim i As Long, est As Boolean
Dim itab As Integer, itexto As Integer
Dim modo As String, codigo As String
Dim vecdatos(8) As String
Dim estmod As Boolean
Public lc_Aux As String

Private Sub Check1_Click()

If Check1.Value = 1 Then

   fpText(2).PasswordChar = ""
   
Else

   fpText(2).PasswordChar = "*"

End If

End Sub

Private Sub Check3_Click()

On Error GoTo Man_Error

If estmod Then Exit Sub

If Check3.Value = 1 Then
 
   Check3.Caption = "Desbloqueado"
   
ElseIf Check3.Value = 0 Then

   Check3.Caption = "Bloqueado"

End If

itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
itab = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Combo2_Click(Index As Integer)

If Combo2(Index).ListIndex > -1 And itexto = 0 Then


   If Val(fg_codigocbo(Combo2, 0, 10, "")) = 0 Then
   
      MsgBox "No puede seleccionar el perfil usuario Administrador a este usuario", vbInformation, Me.Caption
      Combo2(Index).ListIndex = -1
      Exit Sub
      
   End If
   
   itab = 1
   SSTab1.Tab = 1
   SSTab1.TabEnabled(1) = True
   Gl_Ac_Botones Me, 1, 0, modo
   itab = 0

End If

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

Dim RS1 As New ADODB.Recordset

Me.HelpContextID = vg_OpcM
Me.Height = 8190
Me.Width = 10935

estmod = True

fg_centra Me

If lc_Aux = "UsuGPa" Then
    
    MsgTitulo = "Usuario Grupo Paciente"
    Me.Caption = "Usuario Grupo Paciente"

Else
    
    MsgTitulo = "Mantenedor Usuario"
    Me.Caption = "Mantenedor Usuario"

End If

est = True
Combo1.ListIndex = 1
SSTab1.Tab = 0

Frame3.Caption = IIf(vg_modpac, "Grupo Paciente", "Contratos")

vaSpread2.MaxRows = 0
vaSpread2.Row = 0
vaSpread2.Col = 3
vaSpread2.text = IIf(vg_modpac, "Grupo Paciente", "Contratos")
'Frame3.Visible = IIf(vg_modpac, True, False)

modo = ""
Combo2(0).Clear

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT * FROM a_perfil ORDER BY per_nombre", vg_db, adOpenStatic
Do While Not RS1.EOF
   
   Combo2(0).AddItem RS1!per_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!per_codigo), 10) & ")"
   RS1.MoveNext

Loop
RS1.Close
Set RS1 = Nothing

Combo2(0).ListIndex = -1
Gl_Mo_Botones Me, 21
Gl_Ac_Botones Me, 1, 1, modo

MoverDatosGrilla

est = False
estmod = False

End Sub

Private Sub Form_Resize()

If Me.WindowState <> 1 Then SSTab1.Move 0, Toolbar1.Height, ScaleWidth, ScaleHeight - Toolbar1.Height

End Sub

Private Sub fpText_Change(Index As Integer)

If fpText(Index).text <> vecdatos(Index) And itexto = 0 Then
   
   itab = 1
   SSTab1.Tab = 1
   SSTab1.TabEnabled(1) = True
   Gl_Ac_Botones Me, 1, 0, modo
   itab = 0
   
   If Trim(LimpiaDato(fpText(0).text)) <> "" Then
      
      If Asc(Trim(Mid(fpText(0).text, 1, 1))) >= 48 And Asc(Trim(Mid(fpText(0).text, 1, 1))) <= 57 Then
         
         fpText(0).text = ""
   
      End If
      
   End If

End If

End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)

If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub fpText_LostFocus(Index As Integer)

Select Case Index

Case 0
    
    fpText(Index).text = UCase(fpText(Index).text)
    If Trim(fpText(0).text) = "" Then Exit Sub

End Select

End Sub

Private Sub fpTnombre_Change()

Dim RS1 As New ADODB.Recordset
Dim sql1 As String


If LimpiaDato(Trim(fpTnombre.text)) & Chr(KeyAscii) = "" Then Exit Sub

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    
    sql1 = IIf(vg_tipbase = "1", " UCASE(usu_codigo) ", " UPPER(usu_codigo) ")
    RS1.Open "SELECT isnull(usu_codigo,0) as usu_codigo, isnull(usu_nombre,'') as usu_nombre, isnull(usu_activo,'') as usu_activo FROM a_usuarios WHERE " & sql1 & " LIKE '%" & UCase(LimpiaDato(fpTnombre.text)) & "%' ORDER BY usu_codigo", vg_db, adOpenStatic

ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    
    sql1 = IIf(vg_tipbase = "1", " UCASE(usu_nombre) ", " UPPER(usu_nombre) ")
    RS1.Open "SELECT isnull(usu_codigo, 0) as usu_codigo, isnull(usu_nombre,'') as usu_nombre, isnull(usu_activo,'') as usu_activo FROM a_usuarios WHERE " & sql1 & " LIKE '%" & UCase(LimpiaDato(fpTnombre.text)) & "%' ORDER BY usu_nombre", vg_db, adOpenStatic

End If

vaSpread1.MaxRows = RS1.RecordCount
i = 1

If Not RS1.EOF Then
   
   Do While Not RS1.EOF
      
      vaSpread1.Row = i
      i = i + 1
      
      vaSpread1.Col = -1
      vaSpread1.BackColor = IIf(IsNull(RS1!usu_activo) Or RS1!usu_activo = 0, Shape1(1).FillColor, Shape1(0).FillColor)

      vaSpread1.Col = 1
      vaSpread1.text = RS1!usu_codigo
      
      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = Trim(RS1!usu_Nombre)
      
      RS1.MoveNext
   
   Loop
   
   SSTab1.TabEnabled(1) = True: modo = ""

Else
   
   SSTab1.TabEnabled(1) = False

End If
RS1.Close
Set RS1 = Nothing
Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows < 1, 2, 1), modo
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Select Case SSTab1.Tab

Case 0

   modo = ""
   Check1.Value = 0
   Gl_Ac_Botones Me, 1, 1, modo

Case 1
    
    If vaSpread1.MaxRows > 0 And itab = 0 Then
       
       Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), CStr(Me.HelpContextID), "", "", "")
       
       est = True
       estmod = True
       modo = "M"
       SSTab1.TabEnabled(0) = True
       SSTab1.Tab = 1
       SSTab1.TabEnabled(1) = True
       itexto = 1
       
       MoverDatos
       
       itexto = 0
       est = False
       estmod = False


   End If
   
End Select

End Sub

Private Sub Text2_Change()

If Text2.text <> vecdatos(8) And itexto = 0 Then
   
   itab = 1
   SSTab1.Tab = 1
   SSTab1.TabEnabled(1) = True
   Gl_Ac_Botones Me, 1, 0, modo
   itab = 0
   
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim RS1 As New ADODB.Recordset

Select Case Button.Index

Case 1
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Agregar"), CStr(Me.HelpContextID), "", "", "")
    
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False: itab = 1
    SSTab1.Tab = 1
    SSTab1.TabEnabled(1) = True
    itexto = 1
    estmod = True
    
    Check3.Enabled = False
    
    For i = 0 To 6
        
        If i < 7 Then fpText(i).Enabled = True: fpText(i).text = ""
        vecdatos(i) = ""
    
    Next i
    
    Text2.text = ""
    
    Text2.Enabled = True
    Combo2(0).Enabled = True
    
    Check1.Value = 0
    est = True
    
    For i = 1 To vaSpread2.MaxRows
        
        vaSpread2.Row = i
        vaSpread2.Col = 1
        vaSpread2.text = "0"
    
    Next i
    
    est = False
    estmod = False
    Combo2(0).ListIndex = -1
    itexto = 0
    itab = 0

Case 3
    
    If vaSpread1.MaxRows < 1 Then Exit Sub
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), CStr(Me.HelpContextID), "", "", "")

    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
    itab = 1
    SSTab1.Tab = 1
    SSTab1.TabEnabled(1) = True
    itexto = 1
    estmod = True
    
    MoverDatos
    
    itexto = 0
    itab = 0
    estmod = False
    
Case 5
    
    If vaSpread1.MaxRows < 1 Then Exit Sub

    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Eliminar"), CStr(Me.HelpContextID), "", "", "")
    
    Borra_Datos

Case 7
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Listar_Lista"), CStr(Me.HelpContextID), "", "", "")

    est = True
    modo = ""
    SSTab1.Tab = 0
    
    MoverDatosGrilla
    
    est = False

Case 10
    
    If MsgBox("Cancela registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Cancelar"), CStr(Me.HelpContextID), "", "", "")

    SSTab1.TabEnabled(0) = True
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    sql1 = IIf(vg_tipbase = "1", " UCASE(usu_codigo) ", " UPPER(usu_codigo) ")
    
    RS1.Open "SELECT COUNT(*) AS nreg FROM a_usuarios WHERE " & sql1 & " LIKE '%" & UCase(("")) & "%'", vg_db, adOpenStatic
    If RS1.EOF Or RS1!nreg = 0 Then RS1.Close: Set RS1 = Nothing: SSTab1.TabEnabled(1) = False: modo = "NE": SSTab1.Tab = 0: Gl_Ac_Botones Me, 1, 2, modo: Exit Sub
    RS1.Close
    Set RS1 = Nothing
    
    SSTab1.TabEnabled(1) = IIf(vaSpread1.MaxRows > 0, True, False)
    SSTab1.Tab = 0
    modo = ""
    Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows < 1, IIf(lc_Aux = "UsuGPa", 3, 2), IIf(lc_Aux = "UsuGPa", 5, 1)), modo

Case 12
    
    Actualiza_Datos

'Case 15
    
'    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, , MsgTitulo: Exit Sub
'    I_Usuari

Case 18
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), CStr(Me.HelpContextID), "", "", "")
    
    Me.Hide
    Unload Me

End Select

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

On Error GoTo Man_Error

Dim RS         As New ADODB.Recordset
Dim xlApp      As Object
Dim xlWb       As Object
Dim xlWs       As Object

Select Case ButtonMenu

    Case "Imprimir Usuario Perfil"
    
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Imprimir"), CStr(Me.HelpContextID), "", "", "")
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("sgp_Sel_Imprimir_Encabezado_Usuario")
        If Not RS.EOF Then
            
           If RS.RecordCount > 1020000 Then
      
              RS.Close
              Set RS = Nothing
              MsgBox "El resultado sobrepasa maximo de fila en excel, Debera seleccionar menos datos", vbCritical
              Exit Sub
   
           End If
           
           'Abrimos el Commondialog con ShowOpen
           CD.DialogTitle = "Seleccione un archivo excel"
           CD.Filter = "Archivos xls|*.xls|Archivos xlsx|*.xlsx"
           CD.DefaultExt = "*.xls|*.xlsx"
           CD.FilterIndex = 2
           CD.Flags = cdlOFNFileMustExist
           CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
           CD.FileName = ""
           CD.ShowSave

           'Si seleccionamos un archivo mostramos la ruta
           If CD.FileName <> "" Then

              '-------> Create an instance of Excel and add a workbook
              Set xlApp = CreateObject("Excel.Application")
              Set xlWb = xlApp.Workbooks.Add
              Set xlWs = xlWb.Worksheets("Hoja1")
  
              '-------> Display Excel and give user control of Excel's lifetime
              xlApp.UserControl = True
    
              '-------> Check version of Excel
              Call encabezado(RS, xlWs)
          
              xlWs.Cells(2, 1).CopyFromRecordset RS

              '-------> Auto-fit the column widths and row heights
              xlApp.Selection.CurrentRegion.Columns.AutoFit
              xlApp.Selection.CurrentRegion.Rows.AutoFit
    
              xlWb.Close True, CD.FileName

              Dim XL As New Excel.Application 'Crea el objeto excel
              XL.Workbooks.Open CD.FileName, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
              XL.Visible = True
              XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
              '-------> Close ADO objects
              RS.Close
              Set RS = Nothing
    
              '-- Cerrar Excel
              xlApp.Quit
              '-------> Release Excel references
              Set xlWs = Nothing
              Set xlWb = Nothing
              Set xlApp = Nothing
  
              fg_descarga
              MsgBox "Proceso Finalizado", vbInformation, Me.Caption
                
           Else
              'Si no mostramos un texto de advertencia de que no se seleccionó _
               ninguno, ya que FileName devuelve una cadena vacía
               
               MsgBox "No seleccionó ningún archivo", vbCritical

           End If

        Else
        
            fg_descarga
            MsgBox "No existe información...", vbCritical
            RS.Close
            Set RS = Nothing
        
        
        End If
        fg_descarga
    
    Case "Transacciones Usuarios"
        
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Imprimir"), CStr(Me.HelpContextID), "", "", "")
        
        I_TransUsuario.Show 1

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub MoverDatosGrilla()

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset

fg_carga ""
itab = 0
vaSpread2.MaxRows = 0

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If vg_modpac = True Then
   
   RS1.Open "SELECT * FROM a_grupopaciente ORDER BY grp_codigo", vg_db, adOpenStatic
   
   If Not RS1.EOF Then
      
      Do While Not RS1.EOF
         
         vaSpread2.MaxRows = vaSpread2.MaxRows + 1
         vaSpread2.Row = vaSpread2.MaxRows
                
         vaSpread2.Col = 1
         vaSpread2.text = "0"
         
         vaSpread2.Col = 2
         vaSpread2.text = RS1!grp_codigo
       
         vaSpread2.Col = 3
         vaSpread2.TypeHAlign = TypeHAlignLeft
         vaSpread2.text = RS1!grp_codigo & " - " & IIf(IsNull(RS1!grp_nombre), "", Trim(RS1!grp_nombre))
         
         RS1.MoveNext
      Loop
   
   End If
   RS1.Close
   Set RS1 = Nothing

Else
   
   RS1.Open "SELECT DISTINCT a.* FROM b_clientes a, a_bodega b WHERE a.cli_codbod = b.bod_codigo AND cli_tipo = 0 ORDER BY cli_nombre", vg_db, adOpenStatic
   
   If Not RS1.EOF Then
      
      Do While Not RS1.EOF
         
         vaSpread2.MaxRows = vaSpread2.MaxRows + 1
         vaSpread2.Row = vaSpread2.MaxRows
                
         vaSpread2.Col = 1
         vaSpread2.text = "0"
         
         vaSpread2.Col = 2
         vaSpread2.text = Trim(RS1!cli_codigo)
       
         vaSpread2.Col = 3
         vaSpread2.TypeHAlign = TypeHAlignLeft
         vaSpread2.text = Trim(RS1!cli_codigo) & " - " & IIf(IsNull(RS1!cli_nombre), "", Trim(RS1!cli_nombre))
         
         RS1.MoveNext
      
      Loop
   
   End If
   RS1.Close
   Set RS1 = Nothing

End If

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = Shape1(0).FillColor
vaSpread1.MaxRows = 0

RS1.Open "SELECT * FROM a_usuarios ORDER BY usu_nombre", vg_db, adOpenStatic
If Not RS1.EOF Then
   
   Do While Not RS1.EOF
      
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
                
      vaSpread1.Col = -1
      vaSpread1.BackColor = IIf(IsNull(RS1!usu_activo) Or RS1!usu_activo = 0, Shape1(1).FillColor, Shape1(0).FillColor)
      
      vaSpread1.Col = 1
      vaSpread1.text = RS1!usu_codigo
       
      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = IIf(IsNull(RS1!usu_Nombre), "", (Trim(RS1!usu_Nombre)))
             
      RS1.MoveNext
   
   Loop
   
   Gl_Ac_Botones Me, 1, IIf(lc_Aux = "UsuGPa", 5, 1), modo
   SSTab1.TabEnabled(1) = True

Else
    
    SSTab1.Tab = 0
    SSTab1.TabEnabled(1) = False
    modo = "NE"
    Gl_Ac_Botones Me, 1, IIf(lc_Aux = "UsuGPa", 3, 2), modo

End If
RS1.Close
Set RS1 = Nothing
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
fpTnombre.text = ""
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, "Usuario"
End Sub

Private Sub MoverDatos()

Dim RS1 As New ADODB.Recordset

fg_carga ""

For i = 0 To 6
    
    If i < 7 Then
    
       fpText(i).text = ""
       vecdatos(i) = ""
    
    End If
    
Next i

Text2.text = ""
Check1.Value = 0

fpText(0).Enabled = False
itexto = 1
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = vaSpread1.text

estmod = True

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT * FROM a_usuarios WHERE usu_codigo = '" & codigo & "'", vg_db, adOpenStatic

If Not RS1.EOF Then
    
    fpText(0).text = Trim(RS1!usu_codigo)
    vecdatos(0) = Trim(RS1!usu_codigo)
    
    fpText(1).text = IIf(IsNull(RS1!usu_Nombre), "", Trim(RS1!usu_Nombre))
    vecdatos(1) = IIf(IsNull(RS1!usu_Nombre), "", Trim(RS1!usu_Nombre))
    
    fpText(2).text = IIf(IsNull(RS1!usu_password), "", fg_Desencripta(Trim(RS1!usu_password)))
    vecdatos(2) = IIf(IsNull(RS1!usu_password), "", fg_Desencripta(Trim(RS1!usu_password)))
    
    fpText(3).text = IIf(IsNull(RS1!usu_oficina), "", Trim(RS1!usu_oficina))
    vecdatos(3) = IIf(IsNull(RS1!usu_oficina), "", Trim(RS1!usu_oficina))
    
    fpText(4).text = IIf(IsNull(RS1!usu_depart), "", Trim(RS1!usu_depart))
    vecdatos(4) = IIf(IsNull(RS1!usu_depart), "", Trim(RS1!usu_depart))
    
    fpText(5).text = IIf(IsNull(RS1!usu_telefono), "", Trim(RS1!usu_telefono))
    vecdatos(5) = IIf(IsNull(RS1!usu_telefono), "", Trim(RS1!usu_telefono))
    
    fpText(6).text = IIf(IsNull(RS1!usu_email), "", Trim(RS1!usu_email))
    vecdatos(6) = IIf(IsNull(RS1!usu_email), "", Trim(RS1!usu_email))
    
    Combo2(0).ListIndex = fg_buscacbo(Combo2, 0, 10, fg_pone_cero(Str(RS1!usu_perfil), 10))

    Check3.Value = IIf(IsNull(RS1!usu_activo) Or RS1!usu_activo = 0, 0, 1)
    Check3.Caption = IIf(IsNull(RS1!usu_activo) Or RS1!usu_activo = 0, "Bloqueado", "Desbloqueado")
    Check3.Enabled = True
    
    Text2.text = IIf(IsNull(RS1!Ticket), "", RS1!Ticket)
    vecdatos(8) = IIf(IsNull(RS1!Ticket), "", Trim(RS1!Ticket))

End If
RS1.Close
Set RS1 = Nothing

'bloquear modificación de datos para usuario Admninistrador Sgp local
Dim UsuarioAdm As String
UsuarioAdm = GetParametro_Seguridad("CAdmsgpLoc")
    
fpText(0).Enabled = IIf(UCase(codigo) = UCase(UsuarioAdm), False, True)
fpText(1).Enabled = IIf(UCase(codigo) = UCase(UsuarioAdm), False, True)
fpText(2).Enabled = IIf(UCase(codigo) = UCase(UsuarioAdm), False, True)
fpText(3).Enabled = IIf(UCase(codigo) = UCase(UsuarioAdm), False, True)
fpText(4).Enabled = IIf(UCase(codigo) = UCase(UsuarioAdm), False, True)
fpText(5).Enabled = IIf(UCase(codigo) = UCase(UsuarioAdm), False, True)
fpText(6).Enabled = IIf(UCase(codigo) = UCase(UsuarioAdm), False, True)
Combo2(0).Enabled = IIf(UCase(codigo) = UCase(UsuarioAdm), False, True)
Check3.Enabled = IIf(UCase(codigo) = UCase(UsuarioAdm), False, True)
Text2.Enabled = IIf(UCase(codigo) = UCase(UsuarioAdm), False, True)

est = True

For i = 1 To vaSpread2.MaxRows
    
    vaSpread2.Row = i
    vaSpread2.Col = 1
    vaSpread2.text = "0"

Next i

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If vg_modpac Then
   
   RS1.Open "SELECT * FROM b_usuariogrupopac WHERE ugp_codusu = '" & codigo & "'", vg_db, adOpenStatic
   
   If Not RS1.EOF Then
      
      Do While Not RS1.EOF
         
         For i = 1 To vaSpread2.MaxRows
             
             vaSpread2.Row = i
             vaSpread2.Col = 2
             
             If vaSpread2.text = RS1!ugp_codgrp Then
                
                vaSpread2.Col = 1
                vaSpread2.text = "1"
                Exit For
             
             End If
         
         Next i
         
         RS1.MoveNext
      
      Loop
   
   End If
   
   RS1.Close: Set RS1 = Nothing

Else
   
   RS1.Open "SELECT * FROM b_usuariocontratos WHERE uco_codusu = '" & codigo & "'", vg_db, adOpenStatic
   
   If Not RS1.EOF Then
      
      Do While Not RS1.EOF
         
         For i = 1 To vaSpread2.MaxRows
             
             vaSpread2.Row = i
             vaSpread2.Col = 2
             
             If Trim(vaSpread2.text) = Trim(RS1!uco_codcon) Then
                
                vaSpread2.Col = 1
                vaSpread2.text = "1"
                Exit For
             
             End If
         
         Next i
         
         RS1.MoveNext
      
      Loop
   
   End If
   
   RS1.Close
   Set RS1 = Nothing

End If
est = False
estmod = False
fg_descarga

End Sub

Private Sub Borra_Datos()

On Error GoTo Man_Error

Dim RS     As New ADODB.Recordset
Dim codigo As String

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = vaSpread1.text

Dim UsuarioAdm As String
UsuarioAdm = GetParametro_Seguridad("CAdmsgpLoc")

If UCase(codigo) = UCase(UsuarioAdm) Then

   MsgBox "Usuario no puede ser eliminado, es administrador de sistema", vbInformation + vbOKOnly, MsgTitulo
   Exit Sub

End If

B_CelEdi.Caption = "Registrar número ticket"
B_CelEdi.Label1.Caption = "Número Ticket"
G_Proc.Txt = ""
B_CelEdi.Show 1

If Trim(G_Proc.Txt) = "" Then

    MsgBox "Debe regitrar numero ticket..", vbCritical + vbOKOnly, MsgTitulo
    Exit Sub

End If

If Not IsNumeric(G_Proc.Txt) Then

    MsgBox "Ticket debe ser númerico.. (" & G_Proc.Txt & ")", vbCritical + vbOKOnly, MsgTitulo
    Exit Sub

End If

If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then

    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Cancelar"), CStr(Me.HelpContextID), "", "", codigo & ";" & fpText(1).text)

    Exit Sub

End If

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminar"), CStr(Me.HelpContextID), "", "", codigo & ";" & fpText(1).text)

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Del_UsuarioContratoPaciente '" & codigo & "', '" & G_Proc.Txt & "'")

If Not RS.EOF Then
   
   If RS(0) > 0 Then

      'registrar Log sistema error Eliminacion
      Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, RS(0) & RS(1), "", codigo & ";" & fpText(1).text)
                         
      MsgBox "Registro finalizo con error " & RS(0), vbInformation + vbOKOnly, MsgTitulo
      
   Else
   
      'registrar Log sistema Eliminar
      Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_EliminacionUsuarioManual"), Me.HelpContextID, "", "", codigo & ";" & fpText(1).text)
      
      vaSpread1.Row = vaSpread1.ActiveRow
      vaSpread1.DeleteRows vaSpread1.Row, 1
      vaSpread1.MaxRows = vaSpread1.MaxRows - 1
      vaSpread1.Row = vaSpread1.MaxRows
      Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
      MsgBox "Registro eliminado exitosamente", vbInformation + vbOKOnly, MsgTitulo
      
   End If

End If
RS.Close
Set RS = Nothing

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.DeleteRows vaSpread1.Row, 1
vaSpread1.MaxRows = vaSpread1.MaxRows - 1
vaSpread1.Row = vaSpread1.MaxRows
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"

If vaSpread1.MaxRows < 1 Then
    
    SSTab1.TabEnabled(1) = False: SSTab1.Tab = 0: modo = "NE"

Else
    
    modo = "": SSTab1.TabEnabled(1) = True: SSTab1.Tab = 0

End If

Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows < 1, 2, 1), modo

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub Actualiza_Datos()

On Error GoTo Man_Error

Dim codusu As String, codgrp As Long, codcli As String
Dim RS1 As New ADODB.Recordset
Dim EstModReg As Boolean
Dim EstModAgr As Boolean
Dim Bloqueo   As String
Dim Ticket    As String


EstModReg = False
EstModAgr = False

Dim UsuarioAdm As String
UsuarioAdm = GetParametro_Seguridad("CAdmsgpLoc")
    
If modo = "A" Then
    
    If Trim(fpText(0).text) = "" Or Trim(fpText(1).text) = "" Or Trim(fpText(2).text) = "" Then
    
       MsgBox "Faltan datos importantes para identificar el Usuario...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If Val(fg_codigocbo(Combo2, 0, 10, "")) = 0 Then
    
       MsgBox "Debe seleccionar perfil del usuario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If Trim(Text2.text) = "" Then
    
       MsgBox "Debe ingresar su ticket...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
       
    End If
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "SELECT * FROM a_usuarios WHERE usu_codigo = '" & LimpiaDato(Trim(fpText(0).text)) & "'", vg_db, adOpenStatic
    If Not RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "Usuario existe", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS1.Close
    Set RS1 = Nothing
    
    'Validar password
    If Not fg_ValidaPassword(LimpiaDato(Trim(fpText(0).text)), LimpiaDato(Trim(fpText(2).text)), MsgTitulo) Then
       
       Exit Sub

    End If
    
    vg_db.BeginTrans
    
    EstModAgr = True
    
    vg_db.Execute "INSERT INTO a_usuarios (usu_codigo, usu_nombre, usu_password, usu_oficina,usu_depart, usu_telefono, usu_email, usu_perfil, usu_activo, Ticket ) " & _
                  "VALUES ('" & LimpiaDato(Trim(fpText(0).text)) & "','" & LimpiaDato(Trim(fpText(1).text)) & "', '" & fg_Encripta(LimpiaDato(Trim(fpText(2).text))) & "', '" & LimpiaDato(Trim(fpText(3).text)) & "', '" & LimpiaDato(Trim(fpText(4).text)) & "', '" & LimpiaDato(Trim(fpText(5).text)) & "', '" & LimpiaDato(Trim(fpText(6).text)) & "', " & Val(fg_codigocbo(Combo2, 0, 10, "")) & ", '1', '" & LimpiaDato(Trim(Text2.text)) & "')"
    
    If vg_modpac Then
       
       '------- grabar usuario grupo paciente
       For i = 1 To vaSpread2.MaxRows
           
           vaSpread2.Row = i
           vaSpread2.Col = 1
           
           If vaSpread2.text = "1" Then
              
              vaSpread2.Col = 2: codgrp = Val(vaSpread2.text)
              codusu = LimpiaDato(Trim(fpText(0).text))
              vg_db.Execute "INSERT INTO b_usuariogrupopac (ugp_codgrp, ugp_codusu) VALUES (" & codgrp & ", '" & codusu & "')"
           
           End If
       
       Next i
    
    Else
       
       '------- grabar usuario vs cencos
       
       For i = 1 To vaSpread2.MaxRows
           
           vaSpread2.Row = i
           vaSpread2.Col = 1
           
           If vaSpread2.text = "1" Then
              
              vaSpread2.Col = 2: codcli = vaSpread2.text
              codusu = LimpiaDato(Trim(fpText(0).text))
              vg_db.Execute "INSERT INTO b_usuariocontratos (uco_codusu, uco_codcon) VALUES ('" & codusu & "', '" & codcli & "')"
           
           End If
       
       Next i
    
    End If
    
    vg_db.CommitTrans
    EstModAgr = False

    vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = 1: vaSpread1.Value = LimpiaDato(Trim(fpText(0).text))
    vaSpread1.Col = 2: vaSpread1.Value = LimpiaDato(Trim(fpText(1).text))
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Agregado"), CStr(Me.HelpContextID), "", "", fpText(0).text & ";" & fpText(1).text)

Else

    'Si usuario administrador no validar datos y grabar
    If UCase(Trim(fpText(0).text)) <> UCase(UsuarioAdm) Then
    
        If Trim(fpText(1).text) = "" Then
           
           MsgBox "Faltan datos importantes para identificar al Usuario...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If Val(fg_codigocbo(Combo2, 0, 10, "")) = 0 Then
        
           MsgBox "Debe seleccionar perfil del usuario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        If Trim(Text2.text) = "" Then
        
           MsgBox "Debe ingresar su ticket...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
           
        End If
        
        Dim pswAnt As String
        pswAnt = ""
        Ticket = ""
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS1.Open "SELECT * FROM a_usuarios WHERE usu_codigo = '" & LimpiaDato(Trim(fpText(0).text)) & "'", vg_db, adOpenStatic
        If Not RS1.EOF Then
        
            pswAnt = IIf(IsNull(RS1!usu_password), "", RS1!usu_password)
            Bloqueo = IIf(IsNull(RS1!usu_activo), "0", RS1!usu_activo)
            Ticket = IIf(IsNull(RS1!Ticket), "", RS1!Ticket)
        
        Else
        
           MsgBox "Usuario no existe...", vbExclamation + vbOKOnly, MsgTitulo
           RS1.Close
           Set RS1 = Nothing
           Exit Sub
            
        End If
        
        RS1.Close
        Set RS1 = Nothing
        
        If Trim(Text2.text) = Trim(Ticket) Then
        
           MsgBox "Debe ingresar su ticket...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
           
        End If
        
        If pswAnt <> fg_Encripta(fpText(2).text) Then
            
            If Not fg_ValidaPassword(Trim(fpText(0).text), Trim(fpText(2).text), MsgTitulo) Then Exit Sub
        
        End If
    
        vg_db.BeginTrans
        
        EstModReg = True
    
        vg_db.Execute "UPDATE a_usuarios SET usu_nombre = '" & LimpiaDato(Trim(fpText(1).text)) & "', usu_password = '" & fg_Encripta(LimpiaDato(Trim(fpText(2).text))) & "', " & _
                      "usu_oficina = '" & LimpiaDato(Trim(fpText(3).text)) & "', usu_depart = '" & LimpiaDato(Trim(fpText(4).text)) & "', usu_telefono = '" & LimpiaDato(Trim(fpText(5).text)) & "', " & _
                      "usu_email = '" & LimpiaDato(Trim(fpText(6).text)) & "',usu_perfil = " & Val(fg_codigocbo(Combo2, 0, 10, "")) & ", usu_activo = '" & IIf(Check3.Value = 1, "1", "0") & "', " & _
                      "Ticket = '" & LimpiaDato(Trim(Text2.text)) & "', Fecha_modificacion = '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "' WHERE usu_codigo = '" & LimpiaDato(Trim(fpText(0).text)) & "'"
        
        If pswAnt <> fg_Encripta(fpText(2).text) Then
            
            'INSERTA MODIFICACIÓN DE PASSWORD
            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_CambiaPass"), CStr(Me.HelpContextID), fg_Encripta(fpText(2).text), pswAnt, fpText(0).text & ";" & fpText(1).text)
        
        End If
        
        If Bloqueo <> Check3.Value Then
        
           'INSERTA MODIFICACIÓN DE BLOQUEO
           Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto(IIf(Check3.Value = 0, "vg_logsis_BloquearUsuarioManual", "vg_logsis_Desbloquear")), CStr(Me.HelpContextID), "", "", fpText(0).text & ";" & fpText(1).text)
           
        End If
    
    
    Else
    
    
        vg_db.BeginTrans
    
    End If
    
    If vg_modpac Then
       
       '------- borrar usuario grupo paciente
       vg_db.Execute "DELETE b_usuariogrupopac FROM b_usuariogrupopac WHERE ugp_codusu = '" & LimpiaDato(Trim(fpText(0).text)) & "'"
       '------- grabar usuario grupo paciente
       
       For i = 1 To vaSpread2.MaxRows
           
           vaSpread2.Row = i
           vaSpread2.Col = 1
           
           If vaSpread2.text = "1" Then
              
              vaSpread2.Col = 2: codgrp = Val(vaSpread2.text)
              codusu = LimpiaDato(Trim(fpText(0).text))
              vg_db.Execute "INSERT INTO b_usuariogrupopac (ugp_codgrp, ugp_codusu) VALUES (" & codgrp & ", '" & codusu & "')"
           
           End If
       
       Next i
    
    Else
       
       '------- borrar usuario vs cencos
       vg_db.Execute "DELETE b_usuariocontratos FROM b_usuariocontratos WHERE uco_codusu = '" & LimpiaDato(Trim(fpText(0).text)) & "'"
       '------- grabar usuario vs cencos
       
       For i = 1 To vaSpread2.MaxRows
           
           vaSpread2.Row = i
           vaSpread2.Col = 1
           
           If vaSpread2.text = "1" Then
              
              vaSpread2.Col = 2: codcli = vaSpread2.text
              codusu = LimpiaDato(Trim(fpText(0).text))
              vg_db.Execute "INSERT INTO b_usuariocontratos (uco_codusu, uco_codcon) VALUES ('" & codusu & "', '" & codcli & "')"
           
           End If
       
       Next i
       
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
   
       RS1.Open "SELECT COUNT(*) AS nreg FROM b_usuariocontratos WHERE uco_codusu = '" & LimpiaDato(Trim(fpText(0).text)) & "'", vg_db, adOpenStatic
       
       If Not RS1.EOF And RS1!nreg = 0 Then
          
          vg_db.RollbackTrans: RS1.Close: Set RS1 = Nothing: MsgBox "Debe haber al menos un contrato asociado este usuario, proceso cancelado....", vbExclamation + vbOKOnly, MsgTitulo: est = False: Exit Sub
       
       End If
       RS1.Close: Set RS1 = Nothing
    
    End If
    
    vg_db.CommitTrans
    
    EstModReg = False
 
    vaSpread1.Col = -1
    vaSpread1.BackColor = IIf(Check3.Value = 0, Shape1(1).FillColor, Shape1(0).FillColor)
    
    vaSpread1.Col = 2
    vaSpread1.Value = LimpiaDato(Trim(fpText(1).text))
    XX = Val(fg_codigocbo(Combo2, 0, 10, ""))
    zz = LimpiaDato(Trim(fpText(0).text))

    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), CStr(Me.HelpContextID), "", "", fpText(0).text & ";" & fpText(1).text)

End If

vaSpread1.SortKey(1) = 2
vaSpread1.SortKeyOrder(1) = 1
vaSpread1.Sort 1, 1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
   
SSTab1.TabEnabled(0) = True

If vaSpread1.MaxRows < 1 Then SSTab1.TabEnabled(1) = False Else SSTab1.TabEnabled(1) = True: SSTab1.Tab = 0
itexto = 1
modo = ""
Gl_Ac_Botones Me, 1, IIf(lc_Aux = "UsuGPa", 5, 1), modo
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"

Exit Sub
Man_Error:
        
If EstModReg Then
          
   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), CStr(Me.HelpContextID), Err & ":  " & error$(Err), "", fpText(0).text & ";" & fpText(1).text)

End If

If EstModAgr Then
   
   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Agregado"), CStr(Me.HelpContextID), Err & ":  " & error$(Err), "", fpText(0).text & ";" & fpText(1).text)

End If

If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = Row: vaSpread1.Col = 1: codigo = vaSpread1.text

End Sub

Private Sub vaSpread2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

Dim RS1 As New ADODB.Recordset

If est Then Exit Sub

If Not vg_modpac And ButtonDown = 0 Then
   
   Dim cencos As String
   vaSpread2.Row = Row: vaSpread2.Col = 2: cencos = Trim(vaSpread2.text)
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   RS1.Open "SELECT COUNT(uco_codusu) AS nreg FROM b_usuariocontratos WHERE uco_codusu = '" & codigo & "'", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg = 1 Then est = True: vaSpread2.Row = Row: vaSpread2.Col = 1: vaSpread2.text = "1": RS1.Close: Set RS1 = Nothing: MsgBox "Existe un solo usuario asociado a este contrato, no se desactivara....", vbExclamation + vbOKOnly, MsgTitulo: est = False: Exit Sub
   RS1.Close
   Set RS1 = Nothing

End If
itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
itab = 0

End Sub
