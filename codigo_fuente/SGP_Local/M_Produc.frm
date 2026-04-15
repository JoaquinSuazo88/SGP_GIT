VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_Produc 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producto"
   ClientHeight    =   7920
   ClientLeft      =   2520
   ClientTop       =   1905
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7470
      Left            =   0
      TabIndex        =   25
      Top             =   375
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   13176
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   4
      BackColor       =   -2147483638
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Producto"
      TabPicture(0)   =   "M_Produc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_Produc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(1)=   "fpText1(1)"
      Tab(1).Control(2)=   "fpText1(4)"
      Tab(1).Control(3)=   "fpText1(3)"
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(5)=   "Label3(12)"
      Tab(1).Control(6)=   "Label3(4)"
      Tab(1).Control(7)=   "Label3(2)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Ingrediente"
      TabPicture(2)   =   "M_Produc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(1)=   "Frame1(1)"
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(3)=   "lblNomPro(0)"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame8 
         Height          =   1275
         Left            =   -74760
         TabIndex        =   88
         Top             =   510
         Width           =   7695
         Begin FPSpread.vaSpread vaSpread4 
            Height          =   1005
            Left            =   570
            TabIndex        =   89
            Top             =   180
            Width           =   6555
            _Version        =   393216
            _ExtentX        =   11562
            _ExtentY        =   1773
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
            MaxRows         =   3
            ScrollBars      =   2
            SpreadDesigner  =   "M_Produc.frx":0054
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Impuestos del Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   2520
         Left            =   -74775
         TabIndex        =   73
         Top             =   4845
         Width           =   7680
         Begin FPSpread.vaSpread vaSpread3 
            Height          =   2205
            Left            =   150
            TabIndex        =   74
            Top             =   225
            Width           =   7380
            _Version        =   393216
            _ExtentX        =   13017
            _ExtentY        =   3889
            _StockProps     =   64
            AutoClipboard   =   0   'False
            DisplayRowHeaders=   0   'False
            EditEnterAction =   2
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
            MaxRows         =   20
            SpreadDesigner  =   "M_Produc.frx":033B
            ScrollBarTrack  =   3
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Nutrientes del Ingrediente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   2550
         Index           =   1
         Left            =   -74775
         TabIndex        =   72
         Top             =   4830
         Width           =   7680
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   2205
            Left            =   150
            TabIndex        =   20
            Top             =   255
            Width           =   7380
            _Version        =   393216
            _ExtentX        =   13017
            _ExtentY        =   3889
            _StockProps     =   64
            AutoClipboard   =   0   'False
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
            MaxCols         =   4
            SpreadDesigner  =   "M_Produc.frx":079E
            ScrollBarTrack  =   3
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   0
         Left            =   960
         TabIndex        =   28
         Top             =   480
         Width           =   6015
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   315
            ItemData        =   "M_Produc.frx":20C7
            Left            =   1680
            List            =   "M_Produc.frx":20D4
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   240
            Width           =   2500
         End
         Begin EditLib.fpText fpText 
            Height          =   315
            Left            =   1680
            TabIndex        =   0
            Top             =   585
            Width           =   2505
            _Version        =   196608
            _ExtentX        =   4419
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
            BackColor       =   -2147483628
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
            AutoAdvance     =   -1  'True
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
         Begin VB.Label Label1 
            Caption         =   " Buscar Columna"
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
            Left            =   150
            TabIndex        =   32
            Top             =   270
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   " Buscar Texto"
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
            Left            =   150
            TabIndex        =   31
            Top             =   630
            Width           =   1275
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Label2"
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
            Left            =   4260
            TabIndex        =   30
            Top             =   645
            Width           =   585
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   5730
         Left            =   240
         TabIndex        =   26
         Top             =   1575
         Width           =   7665
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   5355
            Left            =   240
            TabIndex        =   27
            Top             =   225
            Width           =   7245
            _Version        =   393216
            _ExtentX        =   12779
            _ExtentY        =   9446
            _StockProps     =   64
            AutoClipboard   =   0   'False
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
            MaxCols         =   5
            MaxRows         =   20
            OperationMode   =   3
            SelectBlockOptions=   0
            SpreadDesigner  =   "M_Produc.frx":20F1
            VisibleCols     =   3
            VisibleRows     =   15
            ScrollBarTrack  =   3
         End
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   -69510
         TabIndex        =   21
         Top             =   6180
         Visible         =   0   'False
         Width           =   2160
         _Version        =   196608
         _ExtentX        =   3810
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   4
         Left            =   -69465
         TabIndex        =   22
         Top             =   6300
         Visible         =   0   'False
         Width           =   2160
         _Version        =   196608
         _ExtentX        =   3810
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   3
         Left            =   -69450
         TabIndex        =   23
         Top             =   6435
         Visible         =   0   'False
         Width           =   2160
         _Version        =   196608
         _ExtentX        =   3810
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
      Begin VB.Frame Frame3 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   35
         Top             =   1770
         Width           =   7680
         Begin VB.Frame Frame6 
            Height          =   30
            Left            =   30
            TabIndex        =   70
            Top             =   2535
            Width           =   7620
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "P.A.V.B"
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
            Left            =   4095
            TabIndex        =   17
            Top             =   1905
            Width           =   2220
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Ind. Gramos Verdura"
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
            Index           =   1
            Left            =   4095
            TabIndex        =   19
            Top             =   2220
            Width           =   2220
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   1
            Left            =   2115
            TabIndex        =   14
            Top             =   1530
            Width           =   1245
            _Version        =   196608
            _ExtentX        =   2196
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
            BackColor       =   -2147483628
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
            ControlType     =   0
            Text            =   ""
            DecimalPlaces   =   2
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
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   2
            Left            =   2115
            TabIndex        =   16
            Top             =   1860
            Width           =   1245
            _Version        =   196608
            _ExtentX        =   2196
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
            BackColor       =   -2147483628
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
            ControlType     =   0
            Text            =   ""
            DecimalPlaces   =   2
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
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   3
            Left            =   2115
            TabIndex        =   18
            Top             =   2190
            Width           =   1245
            _Version        =   196608
            _ExtentX        =   2196
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
            BackColor       =   -2147483628
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
            ControlType     =   0
            Text            =   ""
            DecimalPlaces   =   2
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
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   4
            Left            =   6105
            TabIndex        =   15
            Top             =   1530
            Width           =   1245
            _Version        =   196608
            _ExtentX        =   2196
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
            BackColor       =   -2147483628
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
            ControlType     =   0
            Text            =   ""
            DecimalPlaces   =   2
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   5
            Left            =   2115
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   210
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
            BackColor       =   -2147483628
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   6
            Left            =   2115
            TabIndex        =   11
            Top             =   540
            Width           =   5220
            _Version        =   196608
            _ExtentX        =   9208
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
            BackColor       =   -2147483628
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   7
            Left            =   2115
            TabIndex        =   12
            Top             =   870
            Width           =   5220
            _Version        =   196608
            _ExtentX        =   9208
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
            BackColor       =   -2147483628
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   4
            Left            =   2115
            TabIndex        =   13
            Top             =   1200
            Width           =   1245
            _Version        =   196608
            _ExtentX        =   2196
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
            BackColor       =   -2147483628
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
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   360
            Left            =   4365
            TabIndex        =   71
            Top             =   150
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   635
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BorderStyle     =   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Ult. Compra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   345
            TabIndex        =   69
            Top             =   2640
            Width           =   1665
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Precio Prom. Pond."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   4140
            TabIndex        =   68
            Top             =   2640
            Width           =   1650
         End
         Begin VB.Label fpAyDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2115
            TabIndex        =   65
            Top             =   2610
            Width           =   1200
         End
         Begin VB.Label fpAyDouble1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   6105
            TabIndex        =   64
            Top             =   2610
            Width           =   1200
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Unidad de Medida"
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
            Left            =   360
            TabIndex        =   47
            Top             =   1230
            Width           =   1635
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   3765
            TabIndex        =   45
            Top             =   1200
            Width           =   3525
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   3300
            Picture         =   "M_Produc.frx":2607
            Top             =   1120
            Width           =   480
         End
         Begin VB.Label Label1 
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
            Height          =   225
            Index           =   7
            Left            =   360
            TabIndex        =   44
            Top             =   585
            Width           =   1635
         End
         Begin VB.Label Label3 
            Caption         =   "Nombre Fantasía"
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
            Index           =   14
            Left            =   360
            TabIndex        =   43
            Top             =   915
            Width           =   1560
         End
         Begin VB.Label Label3 
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
            Height          =   225
            Index           =   13
            Left            =   360
            TabIndex        =   42
            Top             =   255
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Factor Nutricional"
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
            Index           =   9
            Left            =   4125
            TabIndex        =   39
            Top             =   1575
            Width           =   1890
         End
         Begin VB.Label Label3 
            Caption         =   "% Cocción"
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
            Left            =   315
            TabIndex        =   38
            Top             =   1905
            Width           =   1710
         End
         Begin VB.Label Label3 
            Caption         =   "% Aprovechamiento"
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
            Index           =   3
            Left            =   315
            TabIndex        =   37
            Top             =   1575
            Width           =   1710
         End
         Begin VB.Label Label3 
            Caption         =   "% Aprov. Nut."
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
            Index           =   8
            Left            =   315
            TabIndex        =   36
            Top             =   2220
            Width           =   1710
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   3810
            TabIndex        =   46
            Top             =   1245
            Width           =   3525
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   8
            Left            =   2160
            TabIndex        =   67
            Top             =   2655
            Width           =   1200
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   9
            Left            =   6150
            TabIndex        =   66
            Top             =   2655
            Width           =   1200
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4425
         Left            =   -74775
         TabIndex        =   48
         Top             =   375
         Width           =   7680
         Begin VB.ComboBox Combo2 
            Height          =   315
            Index           =   0
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   3480
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Controla Stock"
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
            Index           =   3
            Left            =   165
            TabIndex        =   92
            Top             =   3220
            Width           =   2250
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Fecha Vigencia"
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
            Index           =   2
            Left            =   3435
            TabIndex        =   90
            Top             =   3220
            Width           =   2220
         End
         Begin VB.Frame Frame4 
            Height          =   75
            Left            =   30
            TabIndex        =   84
            Top             =   3855
            Width           =   7620
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   0
            Left            =   2265
            TabIndex        =   8
            Top             =   2490
            Width           =   750
            _Version        =   196608
            _ExtentX        =   1323
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
            BackColor       =   -2147483628
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
            ControlType     =   0
            Text            =   ""
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000"
            MinValue        =   "-9000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   2
            Left            =   2265
            TabIndex        =   2
            Top             =   510
            Width           =   5340
            _Version        =   196608
            _ExtentX        =   9419
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
            BackColor       =   -2147483628
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   1
            Left            =   2265
            TabIndex        =   4
            Top             =   1170
            Width           =   750
            _Version        =   196608
            _ExtentX        =   1323
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
            BackColor       =   -2147483628
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   2265
            TabIndex        =   1
            Top             =   180
            Width           =   2160
            _Version        =   196608
            _ExtentX        =   3810
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
            BackColor       =   -2147483628
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   0
            Left            =   2265
            TabIndex        =   3
            Top             =   840
            Width           =   750
            _Version        =   196608
            _ExtentX        =   1323
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
            BackColor       =   -2147483628
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
            Index           =   2
            Left            =   2265
            TabIndex        =   7
            Top             =   2160
            Width           =   750
            _Version        =   196608
            _ExtentX        =   1323
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
            BackColor       =   -2147483628
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
            Index           =   3
            Left            =   2265
            TabIndex        =   9
            Top             =   2820
            Width           =   1125
            _Version        =   196608
            _ExtentX        =   1984
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
            BackColor       =   -2147483628
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
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   5
            Left            =   2265
            TabIndex        =   6
            Top             =   1830
            Width           =   1245
            _Version        =   196608
            _ExtentX        =   2196
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
            BackColor       =   -2147483628
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
            ControlType     =   0
            Text            =   ""
            DecimalPlaces   =   2
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
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   6
            Left            =   2265
            TabIndex        =   5
            Top             =   1500
            Width           =   1245
            _Version        =   196608
            _ExtentX        =   2196
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
            BackColor       =   -2147483628
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
            ControlType     =   0
            Text            =   ""
            DecimalPlaces   =   2
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
            Left            =   6075
            TabIndex        =   91
            Top             =   3180
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            ButtonStyle     =   3
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
            AutoAdvance     =   -1  'True
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
            Text            =   "12/10/2004"
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
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   -1  'True
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   10
            Left            =   2330
            TabIndex        =   95
            Top             =   3600
            Width           =   2550
         End
         Begin VB.Label Label3 
            Caption         =   "Disponible en Contrato"
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
            Index           =   18
            Left            =   165
            TabIndex        =   93
            Top             =   3600
            Width           =   1950
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fact. Conv. Ingrediente"
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
            Left            =   165
            TabIndex        =   86
            Top             =   1905
            Width           =   2025
         End
         Begin VB.Label fpAyDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   3555
            TabIndex        =   76
            Top             =   4035
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Ult. Compra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   390
            Index           =   6
            Left            =   2580
            TabIndex        =   83
            Top             =   3960
            Width           =   930
         End
         Begin VB.Label Label1 
            Caption         =   "Precio Prom. Ponderado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   390
            Index           =   2
            Left            =   5070
            TabIndex        =   82
            Top             =   3960
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Ult. Precio Compra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   390
            Index           =   5
            Left            =   135
            TabIndex        =   81
            Top             =   3960
            Width           =   975
         End
         Begin VB.Label fpAyDouble1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1125
            TabIndex        =   77
            Top             =   4035
            Width           =   1200
         End
         Begin VB.Label fpAyDouble1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   6285
            TabIndex        =   75
            Top             =   4035
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Cantidad x Unidad"
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
            Index           =   4
            Left            =   165
            TabIndex        =   63
            Top             =   2565
            Width           =   1830
         End
         Begin VB.Label Label1 
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
            Height          =   225
            Index           =   3
            Left            =   165
            TabIndex        =   62
            Top             =   570
            Width           =   1830
         End
         Begin VB.Label Label3 
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
            Height          =   225
            Index           =   1
            Left            =   165
            TabIndex        =   61
            Top             =   225
            Width           =   1830
         End
         Begin VB.Label Label3 
            Caption         =   "Familia Producto"
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
            Index           =   5
            Left            =   165
            TabIndex        =   60
            Top             =   900
            Width           =   1830
         End
         Begin VB.Label Label3 
            Caption         =   "Unidad Stock"
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
            Index           =   6
            Left            =   165
            TabIndex        =   59
            Top             =   1245
            Width           =   1830
         End
         Begin VB.Label Label3 
            Caption         =   "Unidad Embalaje"
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
            Index           =   7
            Left            =   165
            TabIndex        =   58
            Top             =   2250
            Width           =   1830
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3435
            TabIndex        =   56
            Top             =   840
            Width           =   4110
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   3435
            TabIndex        =   54
            Top             =   1185
            Width           =   4110
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   3435
            TabIndex        =   52
            Top             =   2160
            Width           =   4110
         End
         Begin VB.Label Label3 
            Caption         =   "Cuenta Contable"
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
            Index           =   11
            Left            =   165
            TabIndex        =   51
            Top             =   2895
            Width           =   1830
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   3795
            TabIndex        =   49
            Top             =   2820
            Width           =   3750
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   3000
            Picture         =   "M_Produc.frx":2911
            Top             =   735
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   3000
            Picture         =   "M_Produc.frx":2C1B
            Top             =   1080
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   3000
            Picture         =   "M_Produc.frx":2F25
            Top             =   2070
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   3360
            Picture         =   "M_Produc.frx":322F
            Top             =   2745
            Width           =   480
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3480
            TabIndex        =   57
            Top             =   900
            Width           =   4110
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   3480
            TabIndex        =   55
            Top             =   1230
            Width           =   4110
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   3480
            TabIndex        =   53
            Top             =   2205
            Width           =   4110
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   3840
            TabIndex        =   50
            Top             =   2865
            Width           =   3750
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   1170
            TabIndex        =   80
            Top             =   4080
            Width           =   1200
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   3600
            TabIndex        =   79
            Top             =   4080
            Width           =   1200
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   6330
            TabIndex        =   78
            Top             =   4080
            Width           =   1200
         End
         Begin VB.Label Label3 
            Caption         =   "Fact. Conv. Stock"
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
            Index           =   17
            Left            =   165
            TabIndex        =   87
            Top             =   1575
            Width           =   1830
         End
      End
      Begin VB.Label lblNomPro 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   195
         Index           =   0
         Left            =   -72195
         TabIndex        =   85
         Top             =   345
         Width           =   1800
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código Compras"
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
         Index           =   12
         Left            =   -70875
         TabIndex        =   41
         Top             =   6330
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre Fantasía"
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
         Index           =   4
         Left            =   -71010
         TabIndex        =   34
         Top             =   6495
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código Barra"
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
         Left            =   -70620
         TabIndex        =   33
         Top             =   6210
         Visible         =   0   'False
         Width           =   1050
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Unidad Embalaje"
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
      Index           =   10
      Left            =   0
      TabIndex        =   40
      Top             =   45
      Width           =   1890
   End
End
Attribute VB_Name = "M_Produc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim ibusca As Long, codtip As Long
Dim i As Integer, est As Boolean
Dim modo As String, modo2 As String
Dim codigo As String, codfam As Long, aAp As String

Private Sub Check1_Click(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0, 1
    If modo2 = "" Then modo2 = "M"
    Gl_Ac_Botones Me, 8, 0, modo2
Case 2, 3
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
    If Check1(2).Value = 1 Then fpDateTime1(0).Enabled = True: fpDateTime1(0).text = Format(Date, "dd/mm/yyyy") Else fpDateTime1(0).Enabled = False: fpDateTime1(0).text = "  /  /    "
End Select
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If Index = 1 And vaSpread2.MaxRows > 0 Then vaSpread2.SetActiveCell 3, 1
End Sub

Private Sub Combo1_Click()
If est Then Exit Sub
If Combo1.ListIndex = 2 Then
    vg_left = Frame1(0).Left + Combo1.Left + 1920
'    B_TabEst.LlenaDatos "a_tipopro", "tip_", "Familia del Producto", "Gen"
'    B_TabEst.Show 1
    B_ArbEst.MoverDatosTvwDir "a_tipopro", "tip_", "Familia del Producto"
    B_ArbEst.Show 1
    Me.Refresh
    fpText.Enabled = False
    If Val(vg_codigo) = 0 Then Exit Sub
    codtip = Val(vg_codigo)
    fpText.text = vg_nombre
Else
    fpText.Enabled = True
    fpText.text = ""
End If
End Sub

Private Sub Form_Activate()
fg_descarga
TraerFechaCierre
End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.Height = 8265
Me.Width = 8280
Me.HelpContextID = vg_OpcM
fg_centra Me
EspFecha fpDateTime1(0)
MsgTitulo = "Maestro de Productos"
SSTab1.Tab = 0
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Gl_Mo_Botones Me, 8
If vg_modprod Then Gl_Ac_Botones Me, 8, 1, modo2 Else Gl_Ac_Botones Me, 8, 3, modo2
est = True
fpDateTime1(0).text = "  /  /    "
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &H80000018

vaSpread2.Row = -1: vaSpread2.Col = -1
vaSpread2.BackColor = &H80000018

vaSpread3.Row = -1: vaSpread3.Col = -1
vaSpread3.BackColor = &H80000018

vaSpread4Row = -1: vaSpread4.Col = -1
vaSpread4.BackColor = &H80000018

Combo1.ListIndex = 1
MoverDatosGrilla

'-------> Mover tipo productos
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

With Combo2(0)
    RS1.Open RutinaLectura.TipoServicio(2, 0, ""), vg_db, adOpenStatic
    .Clear
    .AddItem "Ambos" & Space(150) & "(0)"
    Do While Not RS1.EOF
       .AddItem Trim(RS1!tis_nombre) & Space(150) & "(" & RS1!tis_codigo & ")"
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    .ListIndex = -1
End With

'-------> Mover nutrientes
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

With vaSpread2
    
    RS1.Open RutinaLectura.Nutriente(3, 0, ""), vg_db, adOpenStatic
    .MaxRows = 0
    
    Do While Not RS1.EOF
       
       .MaxRows = .MaxRows + 1: .Row = .MaxRows
       .Col = 1: .Value = RS1!nut_codigo
       .Col = 2: .Value = Trim(RS1!nut_nombre)
       .Col = 3: .Value = 0
       .Col = 4: .Value = Trim(RS1!nut_nomuni)
       RS1.MoveNext
    
    Loop
    
    RS1.Close: Set RS1 = Nothing

End With

'-------> Mover impuesto
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

With vaSpread3
    
    RS1.Open RutinaLectura.Impuesto(3, 0, ""), vg_db, adOpenStatic
    
    .MaxRows = 0
    .Row = -1: .Col = -1
    .Lock = False 'True
    
    Do While Not RS1.EOF
       
       .Col = 3: .Enabled = False
       .MaxRows = .MaxRows + 1: .Row = .MaxRows
       .Col = 1: .Value = RS1!imp_codigo
       .Col = 2: .Value = Trim(RS1!imp_nombre)
       .Col = 3: .Value = 1
       .Enabled = True
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing

End With

modo = ""

MoverDatos

est = False
Exit Sub
Man_Error:
Resume Next
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Form_Terminate()
'-------> Borrar tablas temporales
If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
'-------> Borrar tablas temporales
If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
End Sub

Private Sub fpDouble1_Change(Index As Integer)
If est Then Exit Sub
'If SSTab1.Tab = 1 And Index > 4 Then Exit Sub
'If SSTab1.Tab = 2 And Index > 4 Then Exit Sub
If Index >= 1 And Index <= 4 And Val(fpDouble1(Index).Value) = 0 Then fpDouble1(Index).Value = 100
If Index = 0 Or Index = 5 Or Index = 6 Then
   If modo = "" Then modo = "M"
   Gl_Ac_Botones Me, 1, 0, modo
   SSTab1.TabEnabled(0) = False
Else
   If modo2 = "" Then modo2 = "M"
   Gl_Ac_Botones Me, 8, 0, modo2
End If
End Sub

Private Sub fpDouble1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 Or Index = 2 Or Index = 5 Or Index = 6 Then fpDouble1(Index).MaxValue = 9000000# Else fpDouble1(Index).MaxValue = 100
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpDouble1_LostFocus(Index As Integer)
If LimpiaDato(Trim(fpText1(5).text)) <> "" And Index >= 1 And Index <= 4 And Val(fpDouble1(Index).Value) = 0 Then _
   fpDouble1(Index).Value = 100
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
If est Then Exit Sub
fpayuda(Index).Caption = ""
If Index = 4 Then
   If modo2 = "" Then modo2 = "M"
   Gl_Ac_Botones Me, 8, 0, modo2
   Exit Sub
End If
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_LostFocus(Index As Integer)
Dim codi As Long, Bd As String, Ul As String
On Error GoTo Man_Error
If fpLongInteger1(Index).Value = "" Then fpayuda(Index).Caption = "": codi = 0: Exit Sub
If Index = 0 Then
   fpayuda(0).Caption = fg_BuscaenArbol(fpLongInteger1(Index).Value, "a_tipopro", "tip_codigo")
   If Trim(fpayuda(0).Caption) = "" Then
      MsgBox "No existe codigo en la tabla..."
      fpayuda(Index).Caption = ""
      fpLongInteger1(Index).Value = ""
      codi = 0
      On Error Resume Next: fpLongInteger1(Index).SetFocus
   End If
   Exit Sub
End If
codi = fpLongInteger1(Index).Value
Bd = IIf(Index = 0, "a_tipopro", IIf(Index = 1, "a_unidad", IIf(Index = 2, "a_embalaje", IIf(Index = 3, "a_ctacontable", "a_unidadmed"))))
Ul = IIf(Bd = "a_unidadmed", "unm", Mid(Bd, 3, 3))
RS1.Open "select " & Ul & "_nombre from " & Bd & " where " & Ul & "_codigo=" & IIf(Ul = "cta", "'" & codi & "'", codi), vg_db, adOpenStatic
If Not RS1.EOF Then
   fpayuda(Index).Caption = RS1(0)
   codi = 0
Else
   MsgBox "No existe codigo en la tabla..."
   fpayuda(Index).Caption = ""
   fpLongInteger1(Index).Value = ""
   codi = 0
   On Error Resume Next: fpLongInteger1(Index).SetFocus
End If
RS1.Close: Set RS1 = Nothing
Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub fpText1_Change(Index As Integer)
If est Then Exit Sub
If Index = 6 Or Index = 7 Then
   If modo2 = "" Then modo2 = "M"
   Gl_Ac_Botones Me, 8, 0, modo2
   Exit Sub
End If
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
End Sub

Private Sub Limpia(Index As Integer)
est = True
Select Case Index
Case 1
    fpText1(0).text = ""
    fpText1(1).text = ""
    fpText1(2).text = ""
    fpText1(4).text = ""
    fpLongInteger1(0).text = ""
    fpLongInteger1(1).text = ""
    fpLongInteger1(2).text = ""
    fpLongInteger1(3).text = ""
    fpayuda(0).Caption = ""
    fpayuda(1).Caption = ""
    fpayuda(2).Caption = ""
    fpayuda(3).Caption = ""
    fpDouble1(0).Value = 0
    fpDouble1(5).Value = 0
    fpDouble1(6).Value = 0
    fpAyDouble1(1).Caption = ""
    fpAyDouble1(2).Caption = ""
    fpText1(5).text = ""
    fpAyDate(7).Caption = ""
    If vg_modprod = False Then Frame5.Enabled = False
    vaSpread3.Col = 3
    For i = 1 To vaSpread3.MaxRows
        vaSpread3.Row = i
        vaSpread3.text = "0"
    Next i
    Check1(2).Value = 0
    Check1(3).Value = 0
Case 2
    'Gl_Ac_Botones Me, 8, 1, modo2
    fpText1(5).ControlType = ControlTypeNormal
    fpText1(5).text = ""
    fpText1(6).text = ""
    fpText1(7).text = ""
    fpLongInteger1(4).text = ""
    fpayuda(4).Caption = ""
    fpDouble1(1).text = ""
    fpDouble1(2).text = ""
    fpDouble1(3).text = ""
    fpDouble1(4).text = ""
    Check1(0).Value = 0
    Check1(1).Value = 0
    fpAyDouble1(0).Caption = ""
    fpAyDate(0).Caption = ""
    If vg_modprod = False Then
'       Frame3.Enabled = False
       fpText1(5).Enabled = False
       fpText1(6).Enabled = False
       fpText1(7).Enabled = False
       fpLongInteger1(4).Enabled = False
       fpDouble1(1).Enabled = False
       fpDouble1(2).Enabled = False
       fpDouble1(3).Enabled = False
       fpDouble1(4).Enabled = False
       Check1(0).Enabled = False
       Check1(1).Enabled = False
       Image1(4).Enabled = False
       vaSpread2.Col = 1: vaSpread2.Col2 = 3: vaSpread2.Row = 1: vaSpread2.Row2 = vaSpread2.MaxRows
       vaSpread2.BlockMode = True
       ' Lock cells
       vaSpread2.Lock = True
       ' Protect the cells from being edited
       vaSpread2.Protect = True
    End If
    vaSpread2.Col = 3
    For i = 1 To vaSpread2.MaxRows
        vaSpread2.Row = i
        vaSpread2.text = Format(0, fg_Pict(9, 4))
    Next i
End Select
est = False
End Sub

Private Sub fpText1_GotFocus(Index As Integer)
If Index = 5 And modo2 <> "A" Then Limpia 2
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_Change()

Dim RS2 As New ADODB.Recordset
If est Then Exit Sub
est = True

On Error GoTo Man_Error
    'If LimpiaDato(Trim(fpText.text)) & Chr(KeyAscii) = "" Then Exit Sub

If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

With vaSpread1
        If Combo1.ItemData(Combo1.ListIndex) = 0 Then
       
       If vg_tipbase = "1" Then
          
          RS2.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, (SELECT TOP 1 ppd_propon FROM " & aAp & " WHERE a.pro_codigo = ppd_codpro AND ppd_cencos = '" & MuestraCasino(1) & "') AS ppd_propon, a.pro_codtip, a.pro_maepro, c.tis_nombre " & _
                   "FROM b_productos a,  a_tiposervicio c, b_clientes d " & _
                   "WHERE (c.tis_codigo = d.cli_codtis OR a.pro_maepro < 1) AND d.cli_codigo = '" & MuestraCasino(1) & "' AND (c.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) " & _
                   "AND   (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) AND a.pro_codigo LIKE '%" & UCase(LimpiaDato(fpText.text)) & "%' ORDER BY a.pro_nombre", vg_db, adOpenStatic
          ibusca = RS2.RecordCount: .MaxRows = RS2.RecordCount
       
       Else
          
'          RS2.Open "SELECT COUNT(a.pro_codigo) AS nreg " & _
'                   "FROM b_productos a, a_tiposervicio c, b_clientes d " & _
'                   "WHERE (c.tis_codigo = d.cli_codtis OR a.pro_maepro < 1) AND d.cli_codigo = '" & MuestraCasino(1) & "' AND (c.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) " & _
'                   "AND   (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) AND a.pro_codigo LIKE '%" & UCase(LimpiaDato(fpText.text)) & "%'", vg_db, adOpenStatic
'          RS2.Close: Set RS2 = Nothing
          
'          RS2.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, (SELECT TOP 1 ppd_propon FROM b_productospmpdia WHERE ppd_codpro = a.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia <= " & Format(Date, "yyyymmdd") & " ORDER BY ppd_fecdia DESC) AS ppd_propon, a.pro_codtip, a.pro_maepro, c.tis_nombre " & _
'                   "FROM b_productos a, a_tiposervicio c, b_clientes d " & _
'                   "WHERE (c.tis_codigo = d.cli_codtis OR a.pro_maepro < 1) AND d.cli_codigo = '" & MuestraCasino(1) & "' AND (c.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) " & _
'                   "AND   (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) AND a.pro_codigo LIKE '%" & UCase(LimpiaDato(fpText.text)) & "%' ORDER BY a.pro_nombre", vg_db, adOpenStatic
          
          Set RS2 = vg_db.Execute("sgp_Sel_BusquedacodProveedor '" & MuestraCasino(1) & "', '%" & UCase(LimpiaDato(fpText.text)) & "%'")
          If RS2.EOF Then .MaxRows = 0: ibusca = 0 Else .MaxRows = RS2.RecordCount: ibusca = RS2.RecordCount
       End If

    ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
       
       If vg_tipbase = "1" Then
          
          RS2.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, (SELECT TOP 1 ppd_propon FROM " & aAp & " WHERE a.pro_codigo = ppd_codpro AND ppd_cencos = '" & MuestraCasino(1) & "') as ppd_propon, a.pro_codtip, a.pro_maepro, c.tis_nombre " & _
                   "FROM b_productos a, a_tiposervicio c, b_clientes d " & _
                   "WHERE (c.tis_codigo = d.cli_codtis OR a.pro_maepro < 1) AND d.cli_codigo = '" & MuestraCasino(1) & "' AND (c.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) " & _
                   "AND    (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) AND UCASE(a.pro_nombre) LIKE '%" & UCase(LimpiaDato(fpText.text)) & "%' ORDER BY a.pro_nombre", vg_db, adOpenStatic
          ibusca = RS2.RecordCount: .MaxRows = RS2.RecordCount
       
       Else
          
'          RS2.Open "SELECT COUNT(a.pro_codigo) as nreg " & _
'                   "FROM b_productos a, a_tiposervicio c, b_clientes d " & _
'                   "WHERE (c.tis_codigo = d.cli_codtis OR a.pro_maepro < 1) AND d.cli_codigo = '" & MuestraCasino(1) & " ' AND (c.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) " & _
'                   "AND   UPPER(a.pro_nombre) LIKE '%" & UCase(LimpiaDato(fpText.text)) & "%'", vg_db, adOpenStatic
'          If RS2.EOF Then .MaxRows = 0: ibusca = 0 Else .MaxRows = RS2!nreg: ibusca = RS2!nreg
'          RS2.Close: Set RS2 = Nothing
'
'          RS2.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, (SELECT TOP 1 ppd_propon FROM b_productospmpdia WHERE ppd_codpro = a.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia <= " & Format(Date, "yyyymmdd") & " ORDER BY ppd_fecdia DESC) AS ppd_propon, a.pro_codtip, a.pro_maepro, c.tis_nombre " & _
'                   "FROM b_productos a, a_tiposervicio c, b_clientes d " & _
'                   "WHERE (c.tis_codigo = d.cli_codtis OR a.pro_maepro < 1) AND d.cli_codigo = '" & MuestraCasino(1) & " ' AND (c.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) " & _
'                   "AND   UPPER(a.pro_nombre) LIKE '%" & UCase(LimpiaDato(fpText.text)) & "%' ORDER BY a.pro_nombre", vg_db, adOpenStatic
       
          Set RS2 = vg_db.Execute("sgp_Sel_BusquedaNomProveedor '" & MuestraCasino(1) & "', '%" & UCase(LimpiaDato(fpText.text)) & "%'")
          If RS2.EOF Then .MaxRows = 0: ibusca = 0 Else .MaxRows = RS2.RecordCount: ibusca = RS2.RecordCount
       
       End If

    ElseIf Combo1.ItemData(Combo1.ListIndex) = 2 Then
       
       If vg_tipbase = "1" Then
          
          RS2.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, b.ppd_propon, a.pro_codtip, a.pro_maepro, c.tis_nombre " & _
                   "FROM b_productos a, " & aAp & " b, a_tiposervicio c, b_clientes d " & _
                   "WHERE (c.tis_codigo = d.cli_codtis OR a.pro_maepro < 1) AND d.cli_codigo = '" & MuestraCasino(1) & "' AND (c.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) " & _
                   "AND    a.pro_codigo = b.ppd_codpro AND b.ppd_cencos = '" & MuestraCasino(1) & "' AND (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0)AND a.pro_codtip = " & codtip & " ORDER BY a.pro_nombre", vg_db, adOpenStatic
          ibusca = RS2.RecordCount: .MaxRows = RS2.RecordCount
       
       Else
          
'          RS2.Open "SELECT COUNT(a.pro_codigo) as nreg " & _
'                   "FROM b_productos a, a_tiposervicio c, b_clientes d " & _
'                   "WHERE (c.tis_codigo = d.cli_codtis OR a.pro_maepro < 1) AND d.cli_codigo = '" & MuestraCasino(1) & "' AND (c.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) " & _
'                   "AND   (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) AND a.pro_codtip = " & codtip & "", vg_db, adOpenStatic
'          If RS2.EOF Then .MaxRows = 0: ibusca = 0 Else .MaxRows = RS2!nreg: ibusca = RS2!nreg
'          RS2.Close: Set RS2 = Nothing
'          RS2.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, (SELECT TOP 1 ppd_propon FROM b_productospmpdia WHERE ppd_codpro = a.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia <= " & Format(Date, "yyyymmdd") & " ORDER BY ppd_fecdia DESC) AS ppd_propon, a.pro_codtip, a.pro_maepro, c.tis_nombre " & _
'                   "FROM b_productos a, a_tiposervicio c, b_clientes d " & _
'                   "WHERE (c.tis_codigo = d.cli_codtis OR a.pro_maepro < 1) AND d.cli_codigo = '" & MuestraCasino(1) & "' AND (c.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) " & _
'                   "AND   (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) AND a.pro_codtip = " & codtip & " ORDER BY a.pro_nombre", vg_db, adOpenStatic
       
          Set RS2 = vg_db.Execute("sgp_Sel_BusquedaFamProveedor '" & MuestraCasino(1) & "', " & codtip & "")
          If RS2.EOF Then .MaxRows = 0: ibusca = 0 Else .MaxRows = RS2.RecordCount: ibusca = RS2.RecordCount
       
       End If

    End If
    i = 1
    If Not RS2.EOF Then
        
        Do While Not RS2.EOF
           
           .Row = i
           i = i + 1
           .Col = 1
           .TypeHAlign = 0
           .Value = RS2!pro_codigo
            
           .Col = 2
           .TypeHAlign = 0
           .Value = Trim(IIf(IsNull(RS2!pro_nombre), "", (RS2!pro_nombre)))
            
           .Col = 3
           .TypeHAlign = TypeHAlignLeft
           .Value = IIf(RS2!pro_maepro = 0, "Ambos", Trim(RS2!tis_nombre))
           
           .Col = 4
           .TypeHAlign = 1
           .Value = Format(IIf(IsNull(RS2!ppd_propon), 0, RS2!ppd_propon), fg_Pict(9, vg_DPr))
            
           .Col = 5
           .text = RS2!pro_codtip
            
           RS2.MoveNext
        
        Loop
        
        .SortKey(1) = 1
        .SortKeyOrder(1) = 1
        .Sort -1, -1, .MaxCols, .MaxRows, SortByRow
        
        .SetActiveCell 1, 1
        vaSpread1_Click 1, 1
        If vg_modprod = True Then Gl_Ac_Botones Me, 1, 1, modo Else Gl_Ac_Botones Me, 1, 5, modo
    
    Else
        
        modo = "": MoverDatos
    
    End If
    
    RS2.Close: Set RS2 = Nothing
    
    est = False
    
    If fpText.text = "" Then
        
        Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registro"
    
    Else
        
        Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
    
    End If
    
    If .MaxRows = 0 Then
        
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
    
    ElseIf .MaxRows > 0 Then
        
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
    
    End If

End With

Exit Sub
Man_Error:
'If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
'vg_db.RollbackTrans
Resume Next
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub fpText1_LostFocus(Index As Integer)
Dim codigo As String, i As Long
codigo = LimpiaDato(Trim(fpText1(5).text))
Select Case Index
Case 5
    If codigo = "" Then MsgBox "Debe ingresar ingrediente...", vbExclamation, MsgTitulo: Exit Sub
    For i = 1 To 4
        If codigo <> "" And Val(fpDouble1(i).Value) = 0 Then _
           fpDouble1(i).Value = 100
    Next i
'jp
    RS1.Open RutinaLectura.Ingrediente(2, codigo, ""), vg_db, adOpenStatic
    If Not RS1.EOF Then fpText1(5).text = "": MsgBox "Ingrediente existe...", vbExclamation, MsgTitulo: SendKeys "+{Tab}"
    RS1.Close: Set RS1 = Nothing: Exit Sub
'    modo2 = "": MoverDatos2 codigo
End Select
End Sub

Private Sub MoverDatos2(codpro As String)
Dim RS1 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset

With vaSpread4
    .MaxRows = 0
    RS1.Open RutinaLectura.ProductoIng(1, codpro, ""), vg_db, adOpenStatic
    If Not RS1.EOF Then
        Do While Not RS1.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1: .text = RS1!ing_codigo
          .Col = 2: .text = RS1!ing_nombre
          RS1.MoveNext
       Loop
    End If
    RS1.Close: Set RS1 = Nothing
    If .MaxRows < 1 Then Exit Sub
    .Row = 1: .Col = 1
End With
RS1.Open RutinaLectura.Ingrediente(3, vaSpread4.text, ""), vg_db, adOpenStatic
If Not RS1.EOF Then
    est = True
    fpText1(5).ControlType = ControlTypeStatic
    fpText1(5).text = RS1!ing_codigo
    fpText1(6).text = RS1!ing_nombre
    fpText1(7).text = RS1!ing_nomfan
    fpLongInteger1(4).text = RS1!ing_unimed
    fpayuda(4).Caption = RS1!unm_nombre
    fpDouble1(1).Value = IIf(IsNull(RS1!ing_pctapr), 0, RS1!ing_pctapr)
    fpDouble1(2).Value = IIf(IsNull(RS1!ing_pctcoc), 0, RS1!ing_pctcoc)
    fpDouble1(3).Value = IIf(IsNull(RS1!ing_pctnut), 0, RS1!ing_pctnut)
    fpDouble1(4).Value = IIf(IsNull(RS1!ing_facnut), 0, RS1!ing_facnut)
    Check1(0).Value = IIf(IsNull(RS1!ing_indpav), 0, RS1!ing_indpav)
    Check1(1).Value = IIf(IsNull(RS1!ing_indgrv), 0, RS1!ing_indgrv)
    fpAyDouble1(0).Caption = IIf(RS1!cpi_precos = 0, "", Format(RS1!cpi_precos, fg_Pict(9, 4)))
    fpAyDate(0).Caption = IIf(RS1!cpi_feccos = 0, "", fg_Ctod1(RS1!cpi_feccos))
    With vaSpread2
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1: CodN = Val(.Value)
            RS3.Open RutinaLectura.ProductoNut(1, vaSpread4.text, CodN, ""), vg_db, adOpenStatic
            If Not RS3.EOF Then
                .Col = 3: .Value = RS3!pnu_canapo
            Else
                .Col = 3: .Value = 0
            End If
            RS3.Close: Set RS3 = Nothing
        Next i
'        RS3.Open RutinaLectura.ProductoNut(2, vaSpread4.text, 0, ""), vg_db, adOpenStatic
'        If Not RS3.EOF Then
'           Do While Not RS3.EOF
'              .Row = .SearchCol(1, -1, .MaxRows, Trim(CStr(RS3!pnu_codapo)), SearchFlagsEqual)
'              .Col = 3: .Value = RS3!pnu_canapo
'              RS3.MoveNext
'           Loop
'        End If
'        RS3.Close: Set RS3 = Nothing
'
    End With
    est = False
Else
    fpText1(5).ControlType = ControlTypeNormal
    'Limpia 2
    'Gl_Ac_Botones Me, 8, 2, modo2
End If
RS1.Close: Set RS1 = Nothing
End Sub

Private Sub MoverDatos3(coding As String)
Dim RS1 As New ADODB.Recordset

RS1.Open RutinaLectura.Ingrediente(3, coding, ""), vg_db, adOpenStatic
If Not RS1.EOF Then
    est = True
    fpText1(5).ControlType = ControlTypeStatic
    fpText1(5).text = RS1!ing_codigo
    fpText1(6).text = RS1!ing_nombre
    fpText1(7).text = RS1!ing_nomfan
    fpLongInteger1(4).text = RS1!ing_unimed
    fpayuda(4).Caption = RS1!unm_nombre
    fpDouble1(1).Value = RS1!ing_pctapr
    fpDouble1(2).Value = RS1!ing_pctcoc
    fpDouble1(3).Value = RS1!ing_pctnut
    fpDouble1(4).Value = RS1!ing_facnut
    Check1(0).Value = IIf(IsNull(RS1!ing_indpav), 0, RS1!ing_indpav)
    Check1(1).Value = IIf(IsNull(RS1!ing_indgrv), 0, RS1!ing_indgrv)
    fpAyDouble1(0).Caption = IIf(RS1!cpi_precos = 0, "", Format(RS1!cpi_precos, fg_Pict(9, 4)))
    fpAyDate(0).Caption = IIf(RS1!cpi_feccos = 0, "", fg_Ctod1(RS1!cpi_feccos))
    With vaSpread4
        If .MaxRows > 0 Then
           For i = 1 To .MaxRows
               .Row = i
               .Col = 1
               If Trim(RS1!ing_codigo) = Trim(.text) Then
                  Exit For
               ElseIf Trim(RS1!ing_codigo) <> Trim(.text) And i = .MaxRows Then
                  .MaxRows = .MaxRows + 1
                  .Row = .MaxRows
                  .Col = 1: .text = Trim(RS1!ing_codigo)
                  .Col = 2: .text = Trim(RS1!ing_nombre)
               End If
           Next i
        Else
           .MaxRows = .MaxRows + 1
           .Row = .MaxRows
           .Col = 1: .text = Trim(RS1!ing_codigo)
           .Col = 2: .text = Trim(RS1!ing_nombre)
        End If
    End With
    With vaSpread2
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1: CodN = Val(.Value)
            RS3.Open RutinaLectura.ProductoNut(1, coding, CodN, ""), vg_db, adOpenStatic
            
            If Not RS3.EOF Then
                
                .Col = 3: .Value = RS3!pnu_canapo
            
            Else
                
                .Col = 3: .Value = 0
            
            End If
            RS3.Close: Set RS3 = Nothing
        Next i
    
    End With
    est = False
Else
    fpText1(5).ControlType = ControlTypeNormal
End If
RS1.Close: Set RS1 = Nothing
End Sub

Private Sub Image1_Click(Index As Integer)
vg_codigo = 0
est = True
Select Case Index
Case 0
    vg_left = fpayuda(1).Left + 1920
    B_ArbEst.MoverDatosTvwDir "a_tipopro", "tip_", "Familia del Producto"
    B_ArbEst.Show 1
    Me.Refresh
    If Val(vg_codigo) = 0 Then Exit Sub
    fpayuda(Index).Caption = vg_nombre
    fpLongInteger1(Index) = Val(vg_codigo)
    On Error Resume Next: fpLongInteger1(1).SetFocus
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
Case 1
    vg_left = fpayuda(1).Left + 1920
    B_TabEst.LlenaDatos "a_unidad", "uni_", "Unidad de Envase", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If Val(vg_codigo) = 0 Then Exit Sub
    fpayuda(Index).Caption = vg_nombre
    fpLongInteger1(Index) = Val(vg_codigo)
    On Error Resume Next: fpDouble1(6).SetFocus
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
Case 2
    vg_left = fpayuda(1).Left + 1920
    B_TabEst.LlenaDatos "a_embalaje", "emb_", "Unidad de Embalaje", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If Val(vg_codigo) = 0 Then Exit Sub
    fpayuda(Index).Caption = vg_nombre
    fpLongInteger1(Index) = Val(vg_codigo)
    On Error Resume Next: fpDouble1(0).SetFocus
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
Case 3
    vg_left = fpayuda(1).Left + 1920
    B_TabEst.LlenaDatos "a_ctacontable", "cta_", "Cuenta Contable", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Then Exit Sub
    fpayuda(Index).Caption = vg_nombre
    fpLongInteger1(Index) = vg_codigo
    On Error Resume Next: fpLongInteger1(Index).SetFocus
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
Case 4
    vg_left = fpText1(5).Left + 1920
    B_TabEst.LlenaDatos "a_unidadmed", "unm_", "Unidad de Medida", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If Val(vg_codigo) = 0 Then Exit Sub
    fpayuda(Index).Caption = vg_nombre
    fpLongInteger1(Index) = Val(vg_codigo)
    On Error Resume Next: fpDouble1(1).SetFocus
    If modo2 = "" Then modo2 = "M"
    Gl_Ac_Botones Me, 8, 0, modo2
End Select
est = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long, fecven As Long, prosto As Integer, maepro As Long
Dim codigo As String, CodBar As String, codcom As String, Nombre As String, codfam As Long, coduni As Long, codemb As Long
Dim uniemb As Double, upreco As Double, fecuco As String, propon As Double, ctacon As String, coding As String
Dim facing As Double, facsto As Double
On Error GoTo Man_Error
Select Case Button.Index
Case 1, 3 '-------> Agregar o Modificar
    modo = "A"
    If Button.Index = 3 Then
        modo = "M"
        If vaSpread1.MaxRows < 1 Then Exit Sub
    End If
    If modo = "A" Then
        vaSpread2.Enabled = False: vaSpread3.Enabled = False
        For i = 1 To vaSpread2.MaxRows: vaSpread2.Col = 3: vaSpread2.Row = i: vaSpread2.Value = 0: Next i
        For i = 1 To vaSpread3.MaxRows: vaSpread3.Col = 3: vaSpread3.Row = i: vaSpread3.Value = 0: Next i
        vaSpread2.Enabled = True: vaSpread3.Enabled = True
        vaSpread4.MaxRows = 0
        lblNomPro(0).Caption = ""
    End If
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.Tab = 1
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
    Me.Refresh
    If modo = "M" Then On Error Resume Next: fpText1(1).SetFocus
    If modo = "A" Then On Error Resume Next: MoverDatos: fpText1(0).SetFocus 'Solo efectos de limpieza
Case 5 '-------> Eliminar
    If vaSpread1.MaxRows < 1 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codigo = vaSpread1.Value
    TITLE = "Eliminar Dato"
    Resp_Delete (TITLE)
    If respuesta = vbYes Then
        '-------> Borrando Tabla Productos y relaciones
        vg_db.BeginTrans
        vg_db.Execute "DELETE b_productosing FROM b_productosing WHERE pri_codpro = '" & codigo & "'"
        vg_db.Execute "DELETE b_productospmpdia FROM b_productospmpdia WHERE ppd_codpro = '" & codigo & "'"
        vg_db.Execute "DELETE b_productosimp FROM b_productosimp WHERE ipr_codpro = '" & codigo & "'"
        vg_db.Execute "DELETE b_productos FROM b_productos WHERE pro_codigo = '" & codigo & "'"
        vg_db.CommitTrans
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        vaSpread1.Row = vaSpread1.MaxRows
        If fpText.text = "" Then
            Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
        Else
            Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Encontrado"
        End If
        SSTab1.Tab = 0
    End If
    On Error Resume Next: fpText.SetFocus
    modo = "": MoverDatos
    Gl_Ac_Botones Me, 1, 1, modo
Case 7 '-------> Actualiza Grilla
    MoverDatosGrilla
    modo = "": MoverDatos
Case 10 '-------> Cancelar
    If MsgBox("Cancelar Operación", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If LimpiaDato(Trim(fpText1(5).text)) = "" And modo <> "A" Then MsgBox "Debe relacionar ingrediente...", vbCritical, MsgTitulo: Exit Sub
    modo = "": MoverDatos
    If vg_modprod = True Then Gl_Ac_Botones Me, 1, 1, modo: Gl_Ac_Botones Me, 8, 1, modo2 Else Gl_Ac_Botones Me, 1, 5, modo: Gl_Ac_Botones Me, 8, 3, modo2
    SSTab1.TabEnabled(0) = True
    If vaSpread1.MaxRows = 0 Then
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.Tab = 0
    ElseIf vaSpread1.MaxRows > 0 Then
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
    End If
Case 12 '-------> Confirmar
    Dim indice As String
    Dim CodN As Long, CanN As Double, CodS
    If modo = "A" Or modo = "M" Then
        If LimpiaDato(Trim(fpText1(0).text)) = "" Or _
           LimpiaDato(Trim(fpText1(2).text)) = "" Or _
           LimpiaDato(Trim(fpLongInteger1(0).text)) = "" Or LimpiaDato(Trim(fpLongInteger1(1).text)) = "" Or _
           LimpiaDato(Trim(fpLongInteger1(2).text)) = "" Or LimpiaDato(Trim(fpLongInteger1(3).text)) = "" Or _
           fpDouble1(0).Value = 0 Or fpDouble1(5).Value = 0 Or fpDouble1(6).Value = 0 Or Combo2(0).ListIndex = -1 _
           Then MsgBox "Debe ingresar información...", vbCritical, MsgTitulo: Exit Sub
        If LimpiaDato(Trim(fpText1(5).text)) = "" Then MsgBox "Debe relacionar ingrediente...", vbCritical, MsgTitulo: Exit Sub
        If MsgBox("** Importante **" & VgLinea & "Revise que la definición de impuestos y cuenta contable del producto sea correcta" & VgLinea & "Desea grabar ?...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        est = True: If vg_modprod = True Then Toolbar2_ButtonClick Toolbar2.Buttons(8): est = False
        vg_db.BeginTrans
        codigo = LimpiaDato(Trim(fpText1(0).text))
        CodBar = LimpiaDato(Trim(fpText1(1).text))
        Nombre = LimpiaDato(Trim(fpText1(2).text))
        codcom = LimpiaDato(Trim(fpText1(4).text))
        codfam = fpLongInteger1(0).Value
        coduni = fpLongInteger1(1).Value
        codemb = fpLongInteger1(2).Value
        ctacon = Trim(fpLongInteger1(3).text)
        uniemb = fpDouble1(0).Value
        facing = fpDouble1(5).Value
        facsto = fpDouble1(6).Value
        upreco = 0
        propon = 0
        fecven = IIf(Check1(2).Value = 0, 0, Mid(fpDateTime1(0).text, 7, 4) & Mid(fpDateTime1(0).text, 4, 2) & Mid(fpDateTime1(0).text, 1, 2))
        prosto = IIf(Check1(3).Value = 0, 0, 1)
        fecuco = "Null"
        maepro = Val(fg_codigocbo(Combo2, 0, 1, ""))
        If modo = "A" Then
            vg_db.Execute "INSERT INTO b_productos (pro_codigo , pro_codbar, pro_codcom, pro_codtip, pro_nombre, pro_coduni, pro_facing, pro_facsto, pro_codemb, pro_uniemb, pro_upreco, pro_fecuco, pro_propon, pro_ctacon, pro_fecven, pro_ctrsto, pro_maepro) " & _
                          "VALUES ('" & codigo & "', '" & CodBar & "', '" & codcom & "', " & codfam & ", '" & Nombre & "', " & coduni & ", " & facing & ", " & facsto & ", " & codemb & ", " & uniemb & ", " & upreco & ", " & fecuco & ", " & propon & ", '" & ctacon & "', " & fecven & ", " & prosto & ", " & maepro & ")"
            
            For i = 1 To vaSpread4.MaxRows
                vaSpread4.Row = i: vaSpread4.Col = 1: coding = "": coding = vaSpread4.text
                vg_db.Execute "INSERT INTO b_productosing (pri_codpro, pri_coding) VALUES ('" & codigo & "', '" & coding & "')"
            Next i
            '-------> Insertar datos tabla b_productospmpdia
            RS1.Open "SELECT DISTINCT uco_codcon FROM b_usuariocontratos", vg_db, adOpenStatic
            If Not RS1.EOF Then
               Do While Not RS1.EOF
                  vg_db.Execute "INSERT INTO b_productospmpdia (ppd_cencos, ppd_codpro, ppd_fecdia, ppd_propon, ppd_saldo, ppd_upreco, ppd_fecuco) VALUES ('" & RS1!uco_codcon & "', '" & codigo & "', " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ",0 ,0 ,0 ,NULL)"
                  RS1.MoveNext
               Loop
            End If
            RS1.Close: Set RS1 = Nothing
            vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
            vaSpread1.SetActiveCell 1, vaSpread1.Row
        Else
            vg_db.Execute "UPDATE b_productos SET pro_codbar='" & CodBar & "', " & _
                          "pro_codcom='" & codcom & "', pro_codtip=" & codfam & ", pro_nombre='" & Nombre & "', " & _
                          "pro_coduni=" & coduni & ", pro_facing=" & facing & ", pro_facsto=" & facsto & ", " & _
                          "pro_codemb=" & codemb & ", pro_uniemb=" & uniemb & ", " & _
                          "pro_ctacon='" & ctacon & "', " & _
                          "pro_fecven=" & fecven & ", pro_ctrsto=" & prosto & ", pro_maepro=" & maepro & "  WHERE pro_codigo='" & codigo & "'"
            vg_db.Execute "DELETE b_productosing FROM b_productosing WHERE pri_codpro='" & codigo & "'"
            For i = 1 To vaSpread4.MaxRows
                vaSpread4.Row = i: vaSpread4.Col = 1: coding = "": coding = vaSpread4.text
                vg_db.Execute "INSERT INTO b_productosing (pri_codpro, pri_coding) VALUES ('" & codigo & "', '" & coding & "')"
            Next i
        
        End If
        vg_db.Execute "DELETE FROM b_productosimp WHERE ipr_codpro='" & codigo & "'"
        With vaSpread3
            For Fila = 1 To .MaxRows
                .Row = Fila
                .Col = 1: CodN = Val(.Value)
                .Col = 3: CanN = Val(.Value)
                If CanN <> 0 Then vg_db.Execute "INSERT into b_productosimp VALUES ('" & codigo & "', " & CodN & ")"
            Next Fila
        End With
        '-------> Actuliza codigo compra y pedido de ultimo producto para ingrediente
        vg_db.Execute "UPDATE b_ingrediente SET ing_codped='" & codigo & "', ing_codcom='" & codigo & "' WHERE ing_codigo='" & coding & "'"
        vg_db.CommitTrans
        With vaSpread1
            .Row = .ActiveRow
            .Col = 1: .TypeHAlign = TypeHAlignLeft: .Value = LimpiaDato(Trim(fpText1(0).text))
            .Col = 2: .TypeHAlign = TypeHAlignLeft: .Value = LimpiaDato(Trim(fpText1(2).text))
            .Col = 3: .TypeHAlign = TypeHAlignLeft: .Value = Trim(Mid(Combo2(0).text, 1, 150))
            .Col = 4: .TypeHAlign = TypeHAlignRight: .Value = IIf(fpAyDouble1(2).Caption <> "", Format(fpAyDouble1(2).Caption, fg_Pict(6, 2)), Format(0, 0))
            .Col = 5: .TypeHAlign = 0: .Value = codfam
            .SortKey(1) = 1
            .SortKeyOrder(1) = 1
            .Sort -1, -1, .MaxCols, .MaxRows, SortByRow
            Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registros"
        End With
    End If
    modo = "": MoverDatos
    If vg_modprod = True Then Gl_Ac_Botones Me, 1, 1, modo Else Gl_Ac_Botones Me, 1, 5, modo
    SSTab1.TabEnabled(0) = True
    SSTab1.Tab = 0
Case 15 '-------> Imprimir
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, "Producto": Exit Sub
    I_Produc.TraspasoGrilla vaSpread1, "P"
    I_Produc.Show 1
Case 18 '-------> Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub MoverDatosGrilla()

Dim RS1 As New ADODB.Recordset
Dim i As Long
Dim codpro As String, codTippro As Long, nomTippro As String
Dim Ceco As Variant
Dim arr

On Error GoTo Man_Error

fg_carga ""

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'With vaSpread1
'    If vg_tipbase = "1" Then
'       aAp = Trim(vg_NUsr) & "_tmp_ProductoProPMP"
'       fg_CheckTmp aAp
'       vg_db.Execute "SELECT TOP 1 ppd_cencos, ppd_codpro, 0 AS ppd_propon, 0 AS ppd_upreco, null AS ppd_fecuco, Max(ppd_fecdia) AS ppd_fecdia " & _
'                     "INTO " & aAp & " " & _
'                     "FROM b_productospmpdia " & _
'                     "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
'                     "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(Date, "yyyymmdd") & " " & _
'                     "AND   ppd_propon > 0 " & _
'                     "GROUP BY ppd_cencos, ppd_codpro ORDER BY Max(ppd_fecdia) DESC"
'       vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
'       vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon, " & aAp & ".ppd_upreco=b_productospmpdia.ppd_upreco, " & aAp & ".ppd_fecuco=b_productospmpdia.ppd_fecuco"
'       vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_upreco, ppd_fecuco, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_upreco, ppd_fecuco, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
'       RS1.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, (SELECT DISTINCT ppd_propon FROM " & aAp & " WHERE ppd_codpro = a.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' ) AS ppd_propon, pro_maepro, a.pro_codtip, IIF(a.pro_maepro= 0 ,'Ambos', c.tis_nombre) AS tis_nombre " & _
'                "FROM b_productos a, a_tiposervicio c, b_clientes d  " & _
'                "WHERE (c.tis_codigo = d.cli_codtis OR a.pro_maepro < 1) " & _
'                "AND    d.cli_codigo = '" & MuestraCasino(1) & "' " & _
'                "AND   (c.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) " & _
'                "AND   (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) ORDER BY a.pro_nombre", vg_db, adOpenStatic
'
'    Else
'       RS1.Open RutinaLectura.Producto(2, "", ""), vg_db, adOpenStatic
        Ceco = MuestraCasino(1)
        Sql = "sgp_Sel_ListaProductos '" & Ceco & "'"
        Set RS1 = vg_db.Execute(Sql)
'    End If
    codTippro = 0
    nomTippro = ""
    vaSpread1.Visible = False
    vaSpread1.MaxRows = 0
    If Not RS1.EOF Then
       arr = RS1.GetRows
       RS1.Close: Set RS1 = Nothing
    End If
        For i = 0 To UBound(arr, 2)
    
'        Do While Not RS1.EOF
            vaSpread1.MaxRows = vaSpread1.MaxRows + 1
            vaSpread1.Row = vaSpread1.MaxRows

            vaSpread1.Col = 1
            vaSpread1.TypeHAlign = 0
            vaSpread1.Value = arr(0, i) 'RS1!pro_codigo

            vaSpread1.Col = 2
            vaSpread1.TypeHAlign = 0
            vaSpread1.Value = arr(1, i) 'Trim(RS1!pro_nombre)

            vaSpread1.Col = 3
            vaSpread1.TypeHAlign = TypeHAlignLeft
            vaSpread1.Value = IIf(arr(3, i) = 0, "Ambos", Trim(arr(5, i))) 'IIf(RS1!pro_maepro = 0, "Ambos", Trim(RS1!tis_nombre))

            vaSpread1.Col = 4
            vaSpread1.TypeHAlign = 1
            vaSpread1.Value = Format(IIf(IsNull(arr(2, i)), 0, arr(2, i)), fg_Pict(9, vg_DPr))   'Format(IIf(IsNull(RS1!ppd_propon), 0, RS1!ppd_propon), fg_Pict(9, vg_DPr))

            vaSpread1.Col = 5
            vaSpread1.text = arr(4, i) 'RS1!pro_codtip
         Next i
 '           RS1.MoveNext
 '       Loop
'        If vg_modprod = falso Then modo = "M": Gl_Ac_Botones Me, 1, 5, modo Else Gl_Ac_Botones Me, 1, 1, modo
'    Else
'
'    End If
'    .SortKey(1) = 1
'    .SortKeyOrder(1) = SortKeyOrderAscending
'    .Sort -1, -1, .MaxCols, .MaxRows, SortByRow
    
    If RS1.State = 1 Then
       
       RS1.Close: Set RS1 = Nothing
       If vg_modprod = falso Then modo = "M": Gl_Ac_Botones Me, 1, 5, modo Else Gl_Ac_Botones Me, 1, 2, modo
    
    Else
       
       If vg_modprod = falso Then modo = "M": Gl_Ac_Botones Me, 1, 5, modo Else Gl_Ac_Botones Me, 1, 1, modo
    
    End If
    vaSpread1.Visible = True
    SSTab1.Tab = 0
    Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
'End With

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub MoverDatos()

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset
Dim codigo As String, coding As String, i As Long, CodN As Long
fg_carga ""
Limpia 1: Limpia 2
'-------------------------Mueve datos de Detalle (1ª Carpeta)--------------------------

If modo = "A" Then fpText1(0).ControlType = ControlTypeNormal Else fpText1(0).ControlType = ControlTypeStatic

If modo = "" Then
    
    If vaSpread1.MaxRows > 0 Then vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: codigo = vaSpread1.Value
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    If vg_tipbase = "1" Then
       
       RS1.Open "SELECT pro.*, cta.cta_nombre, tip.tip_nombre, tip.tip_codigo, uni.uni_nombre, emb.emb_nombre, f.ppd_upreco, f.ppd_fecuco, f.ppd_propon " & _
                "FROM b_productos pro, a_ctacontable cta, a_tipopro tip, a_unidad uni, a_embalaje emb, " & aAp & " f " & _
                "WHERE pro.pro_codigo = f.ppd_codpro " & _
                "AND   pro.pro_ctacon = cta.cta_codigo " & _
                "AND   pro.pro_codtip = tip.tip_codigo " & _
                "AND   pro.pro_coduni = uni.uni_codigo " & _
                "AND   pro.pro_codemb = emb.emb_codigo " & _
                "AND   pro.pro_codigo = '" & codigo & "' AND f.ppd_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
       
       If RS1.EOF Then
          
          RS1.Close: Set RS1 = Nothing
          RS1.Open "SELECT pro.*, cta.cta_nombre, tip.tip_nombre, tip.tip_codigo, uni.uni_nombre, emb.emb_nombre, 0 AS ppd_upreco, null AS ppd_fecuco, 0 AS ppd_propon " & _
                   "FROM b_productos pro, a_ctacontable cta, a_tipopro tip, a_unidad uni, a_embalaje emb " & _
                   "WHERE pro.pro_ctacon = cta.cta_codigo " & _
                   "AND   pro.pro_codtip = tip.tip_codigo " & _
                   "AND   pro.pro_coduni = uni.uni_codigo " & _
                   "AND   pro.pro_codemb = emb.emb_codigo " & _
                   "AND   pro.pro_codigo = '" & codigo & "'", vg_db, adOpenStatic
       
       End If
    
    Else
       
       RS1.Open "SELECT pro.*, cta.cta_nombre, tip.tip_nombre, tip.tip_codigo, uni.uni_nombre, emb.emb_nombre, f.ppd_upreco, f.ppd_fecuco, f.ppd_propon " & _
                "FROM b_productos pro, a_ctacontable cta, a_tipopro tip, a_unidad uni, a_embalaje emb, b_productospmpdia f " & _
                "WHERE pro.pro_codigo = f.ppd_codpro " & _
                "AND   pro.pro_ctacon = cta.cta_codigo " & _
                "AND   pro.pro_codtip = tip.tip_codigo " & _
                "AND   pro.pro_coduni = uni.uni_codigo " & _
                "AND   pro.pro_codemb = emb.emb_codigo " & _
                "AND   pro.pro_codigo = '" & codigo & "' AND f.ppd_cencos = '" & MuestraCasino(1) & "' AND f.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & "", vg_db, adOpenStatic
      
      If RS1.EOF Then
         
         RS1.Close: Set RS1 = Nothing
         RS1.Open "SELECT pro.*, cta.cta_nombre, tip.tip_nombre, tip.tip_codigo, uni.uni_nombre, emb.emb_nombre, 0 AS ppd_upreco, null AS ppd_fecuco, 0 AS ppd_propon " & _
                  "FROM b_productos pro, a_ctacontable cta, a_tipopro tip, a_unidad uni, a_embalaje emb " & _
                  "WHERE pro.pro_ctacon = cta.cta_codigo " & _
                  "AND   pro.pro_codtip = tip.tip_codigo " & _
                  "AND   pro.pro_coduni = uni.uni_codigo " & _
                  "AND   pro.pro_codemb = emb.emb_codigo " & _
                  "AND   pro.pro_codigo = '" & codigo & "' ", vg_db, adOpenStatic
      
      End If
    
    End If
    est = True
    
    If Not RS1.EOF Then
        
        '------- Detalle
        fpText1(0).text = Trim(RS1!pro_codigo)
        fpText1(1).text = IIf(IsNull(RS1!pro_codbar), "", TipoDato(RS1!pro_codbar, ""))
        fpText1(4).text = IIf(IsNull(RS1!pro_codcom), "", TipoDato(RS1!pro_codcom, ""))
        fpLongInteger1(0).text = IIf(IsNull(RS1!pro_codtip), 0, RS1!pro_codtip)
        
        fpayuda(0).Caption = fg_BuscaenArbol(RS1!tip_codigo, "a_tipopro", "tip_codigo")
        fpText1(2).text = IIf(IsNull(RS1!pro_nombre), "", Trim(RS1!pro_nombre))
        lblNomPro(0).Caption = IIf(IsNull(RS1!pro_nombre), "", Trim(RS1!pro_nombre))
        fpLongInteger1(1).text = IIf(IsNull(RS1!pro_coduni), 0, RS1!pro_coduni)
        fpayuda(1).Caption = IIf(IsNull(RS1!uni_nombre), "", Trim(RS1!uni_nombre))
        fpLongInteger1(2).text = IIf(IsNull(RS1!pro_codemb), 0, RS1!pro_codemb)
        fpayuda(2).Caption = IIf(IsNull(RS1!emb_nombre), "", Trim(RS1!emb_nombre))
        fpLongInteger1(3).text = IIf(IsNull(RS1!pro_ctacon), "", RS1!pro_ctacon)
        fpayuda(3).Caption = IIf(IsNull(RS1!cta_nombre), "", Trim(RS1!cta_nombre))
        fpDouble1(0).text = IIf(IsNull(RS1!pro_uniemb), 0, RS1!pro_uniemb)
        fpDouble1(5).text = IIf(IsNull(RS1!pro_facing), 0, RS1!pro_facing)
        fpDouble1(6).text = IIf(IsNull(RS1!pro_facsto), 0, RS1!pro_facsto)
        fpAyDouble1(1).Caption = IIf(RS1!ppd_upreco = 0 Or IsNull(RS1!ppd_upreco), "", Format(RS1!ppd_upreco, fg_Pict(9, vg_DPr)))
        fpAyDate(7).Caption = IIf(IsNull(RS1!ppd_fecuco), "", RS1!ppd_fecuco)
        fpAyDouble1(2).Caption = IIf(RS1!ppd_propon = 0 Or IsNull(RS1!ppd_propon), "", Format(RS1!ppd_propon, fg_Pict(9, vg_DPr)))
        fpText1(5).text = ""
        fpDateTime1(0).text = IIf(IsNull(RS1!pro_fecven) Or RS1!pro_fecven = 0, "  /  /    ", Mid(RS1!pro_fecven, 7, 2) & "/" & Mid(RS1!pro_fecven, 5, 2) & "/" & Mid(RS1!pro_fecven, 1, 4))
        fpDateTime1(0).Enabled = IIf(IsNull(RS1!pro_fecven) Or RS1!pro_fecven = 0, False, IIf(vg_modprod = False, False, True))
        Check1(2).Value = IIf(IsNull(RS1!pro_fecven) Or RS1!pro_fecven = 0, 0, 1)
        Check1(3).Value = IIf(IsNull(RS1!pro_ctrsto) Or RS1!pro_ctrsto = 0, 0, 1)
        Combo2(0).ListIndex = fg_buscacbo(Combo2, 0, 1, (RS1!pro_maepro))
    
    End If
    
    RS1.Close: Set RS1 = Nothing
    
    est = True
    '-------------------------Mueve datos de Ingrediente y Aporte Nutricional (2ª Carpeta)--------------------------
    'jp
    vaSpread4.MaxRows = 0
    fpText1(5).ControlType = ControlTypeNormal
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "SELECT COUNT(pri_codpro) AS nreg FROM b_productosing WHERE pri_codpro = '" & codigo & "'", vg_db, adOpenStatic
    If RS1.EOF Or IsNull(RS1!nreg) Then
       RS1.Close: Set RS1 = Nothing
       fpText1(5).ControlType = ControlTypeNormal
    Else
       RS1.Close: Set RS1 = Nothing
       modo2 = "": MoverDatos2 codigo
    End If
    '-------------------------Mueve datos de Impuesto (3ª Carpeta)--------------------------
    est = True
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    With vaSpread3
        RS1.Open "SELECT * FROM b_productosimp WHERE ipr_codpro = '" & codigo & "'", vg_db, adOpenStatic
        If Not RS1.EOF Then
           Do While Not RS1.EOF
              .Row = .SearchCol(1, -1, .MaxRows, Trim(CStr(RS1!ipr_codimp)), SearchFlagsEqual)
              If .Row > 0 Then .Col = 3: .Value = 1
              RS1.MoveNext
           Loop
        End If
        RS1.Close: Set RS1 = Nothing
        
'        For i = 1 To .MaxRows
'            .Row = i
'            .Col = 1: CodN = Val(.Value)
'            RS1.Open "SELECT * FROM b_productosimp WHERE ipr_codpro = '" & codigo & "' AND ipr_codimp = " & CodN, vg_db, adOpenStatic
'            .Col = 3: .Enabled = False
'            If Not RS1.EOF Then
'               .Value = 1
'            Else
'               .Value = 0
'            End If
'            .Enabled = True
'            RS1.Close: Set RS1 = Nothing
'        Next i
    End With
End If
est = False
fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim CodN As Long, CanN As Double, i As Long
Dim codigo As String, Nombre As String, NomFan As String, unimed As Long, pctapr As Double, pctcoc As Double
Dim pctnut As Double, facnut As Double, indpav As Long, indgrv As Long, precos As Double, feccos As Long, codpro As String, coding As String
On Error GoTo Man_Error
'INCLUIR(1)-BORRAR(3)-CANCELAR(6)-CONFIRMAR(8)-BUSCAR(11)-IMPRIMIR(12)
Select Case Button.Index
Case 1 '------- Nuevo - Limpia
    modo2 = "A"
    If modo = "" Then modo = "M"
    Limpia 2
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
    Gl_Ac_Botones Me, 8, 0, modo2
Case 3 '------- Borrar
    If MsgBox("Se perderan los valores de nutrientes y deberá vincular otro ingrediente al producto, desea eliminar...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    '------- Borrando tabla ingredientes
    codigo = LimpiaDato(Trim(fpText1(5).text))
    vg_db.BeginTrans
    With vaSpread4
        For i = 1 To .MaxRows
            .Row = i: .Col = 1
            If Trim(.text) = codigo Then
               vg_db.Execute "DELETE b_productosing FROM b_productosing WHERE pri_codpro='" & Trim(fpText1(0).text) & "' AND pri_coding='" & codigo & "'"
               .DeleteRows i, 1
               .MaxRows = .MaxRows - 1
               Exit For
            End If
        Next i
    End With
    vg_db.CommitTrans
    Limpia 2
    vg_db.BeginTrans
'        vg_db.Execute "update b_productos set pro_coding=Null WHERE pro_codigo='" & LimpiaDato(Trim(fpText1(0).Text)) & "'"
    vg_db.Execute "DELETE b_contlistpreing FROM b_contlistpreing WHERE cpi_coding='" & codigo & "'"
    vg_db.Execute "DELETE b_ingrediente FROM b_ingrediente WHERE ing_codigo='" & codigo & "'"
    vg_db.Execute "DELETE b_productonut FROM b_productonut WHERE pnu_codpro='" & codigo & "'"
    vg_db.CommitTrans
    If vaSpread4.MaxRows > 0 And Trim(fpText1(0).text) <> "" Then MoverDatos2 Trim(fpText1(0).text)
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
Case 6 '------- Cancelar
    If MsgBox("Cancelar Operación", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    modo2 = "": MoverDatos3 Trim(fpText1(5).text)
    Gl_Ac_Botones Me, 8, 1, modo2
    If modo = "" Then Gl_Ac_Botones Me, 1, 1, modo
    SSTab1.TabEnabled(0) = True
Case 8 '------- Confirmar - Grabar
    If LimpiaDato(Trim(fpText1(5))) = "" Or LimpiaDato(Trim(fpText1(6).text)) = "" Or LimpiaDato(Trim(fpLongInteger1(4).text)) = "" Then _
       MsgBox "Debe ingresar información...", vbCritical, MsgTitulo: Exit Sub
    codigo = LimpiaDato(Trim(fpText1(5).text))
    Nombre = LimpiaDato(Trim(fpText1(6).text))
    NomFan = LimpiaDato(Trim(fpText1(7).text))
    unimed = fpLongInteger1(4).Value
    pctapr = fpDouble1(1).Value
    pctcoc = fpDouble1(2).Value
    pctnut = fpDouble1(3).Value
    facnut = fpDouble1(4).Value
    indpav = Check1(0).Value
    indgrv = Check1(1).Value
    precos = Val(fpAyDouble1(0).Caption)
    feccos = IIf(Trim(fpAyDate(0).Caption) = "", 0, Val(Mid(fpAyDate(0).Caption, 7, 4) & Mid(fpAyDate(0).Caption, 4, 2) & Mid(fpAyDate(0).Caption, 1, 2)))
    vg_db.BeginTrans
    RS1.Open RutinaLectura.Ingrediente(2, codigo, ""), vg_db, adOpenStatic
    If Not RS1.EOF Then
        vg_db.Execute "UPDATE b_ingrediente SET ing_nombre='" & Nombre & "', " & _
                      "ing_nomfan='" & NomFan & "', ing_unimed=" & unimed & ", ing_pctapr=" & pctapr & ", " & _
                      "ing_pctcoc=" & pctcoc & ", ing_pctnut=" & pctnut & ", ing_facnut=" & facnut & ", " & _
                      "ing_indpav=" & indpav & ", ing_indgrv=" & indgrv & ", ing_precos=" & precos & ", " & _
                      "ing_feccos=" & feccos & " WHERE ing_codigo='" & codigo & "'"
        For i = 1 To vaSpread4.MaxRows
            vaSpread4.Row = i: vaSpread4.Col = 1
            If Trim(vaSpread4.text) = codigo Then vaSpread4.Col = 2: vaSpread4.text = Nombre
        Next i
    Else
        vg_db.Execute "INSERT INTO b_ingrediente (ing_codigo, ing_nombre, ing_nomfan, ing_unimed, ing_pctapr, ing_pctcoc, ing_pctnut, ing_facnut, ing_indpav, ing_indgrv, ing_precos, ing_feccos) " & _
                      "VALUES ('" & codigo & "', '" & Nombre & "', '" & NomFan & "', " & unimed & ", " & pctapr & ", " & pctcoc & ", " & pctnut & ", " & facnut & ", " & indpav & ", " & indgrv & ", " & precos & ", " & feccos & ")"
        '------- Insertar datos tabla lista ingrediente
        RS2.Open "SELECT DISTINCT uco_codcon FROM b_usuariocontratos", vg_db, adOpenStatic
        If Not RS2.EOF Then
           Do While Not RS2.EOF
              vg_db.Execute "INSERT INTO b_contlistpreing (cpi_cencos, cpi_coding, cpi_precos, cpi_feccos, cpi_codcom, cpi_codped) VALUES ('" & RS2!uco_codcon & "', '" & codigo & "', 0, 0, '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & LimpiaDato(Trim(fpText1(0).text)) & "')"
              RS2.MoveNext
           Loop
        End If
        RS2.Close: Set RS2 = Nothing
        With vaSpread4
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1: .text = codigo
            .Col = 4: .text = Nombre
        End With
    End If
    RS1.Close: Set RS1 = Nothing
    '------- Nutrientes
    vg_db.Execute "DELETE FROM b_productonut WHERE pnu_codpro='" & codigo & "'"
    With vaSpread2
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1: CodN = Val(.Value)
            .Col = 3: CanN = Val(.Value)
            If CanN <> 0 Then vg_db.Execute "INSERT INTO b_productonut VALUES ('" & codigo & "', " & CodN & ", " & CanN & ")"
        Next i
    End With
    vg_db.CommitTrans
    fpText1(5).ControlType = ControlTypeStatic
    If Not est Then MsgBox "Datos de Ingrediente fueron grabados con exito...", vbInformation, MsgTitulo
    modo2 = "": Gl_Ac_Botones Me, 8, 1, modo2
Case 11 '------- Buscar
    vg_left = fpText1(5).Left + 1920
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Then Exit Sub
    modo2 = "": MoverDatos3 Trim(vg_codigo)
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
Case 12 '------- Imprimir
    I_Produc.TraspasoGrilla vaSpread3, "I"
    I_Produc.Show 1
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows > 0 Then modo = "": MoverDatos
End Sub

Private Sub vaSpread2_EditChange(ByVal Col As Long, ByVal Row As Long)
If est Then Exit Sub
If modo2 = "" Then modo2 = "M"
Gl_Ac_Botones Me, 8, 0, modo2
End Sub

Private Sub vaSpread3_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
End Sub

Private Sub vaSpread4_Click(ByVal Col As Long, ByVal Row As Long)
With vaSpread4
    If .MaxRows < 1 Then Exit Sub
    .Row = .ActiveRow: .Col = 1
    MoverDatos3 Trim(.text)
End With
End Sub
