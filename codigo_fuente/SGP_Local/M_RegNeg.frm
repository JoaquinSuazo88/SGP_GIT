VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_RegNeg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reglas de Negocios"
   ClientHeight    =   8910
   ClientLeft      =   2430
   ClientTop       =   1350
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   13005
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Listar Reglas Negocios"
      TabPicture(0)   =   "M_RegNeg.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "vaSpread1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Reglas de Negocios"
      TabPicture(1)   =   "M_RegNeg.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Familias"
      TabPicture(2)   =   "M_RegNeg.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Productos"
      TabPicture(3)   =   "M_RegNeg.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Casinos"
      TabPicture(4)   =   "M_RegNeg.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame10"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Bloquear Productos"
      TabPicture(5)   =   "M_RegNeg.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame17"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame17 
         Height          =   7575
         Left            =   -74760
         TabIndex        =   44
         Top             =   480
         Width           =   12255
         Begin VB.Frame Frame18 
            Height          =   1095
            Left            =   720
            TabIndex        =   46
            Top             =   240
            Width           =   10455
            Begin VB.ComboBox Combo3 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "M_RegNeg.frx":00A8
               Left            =   1845
               List            =   "M_RegNeg.frx":00AF
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   240
               Width           =   2500
            End
            Begin EditLib.fpText fpText1 
               Height          =   315
               Index           =   2
               Left            =   1845
               TabIndex        =   48
               Top             =   555
               Width           =   2505
               _Version        =   196608
               _ExtentX        =   4410
               _ExtentY        =   870
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
               Index           =   7
               Left            =   4680
               TabIndex        =   51
               Top             =   600
               Visible         =   0   'False
               Width           =   1140
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
               Index           =   6
               Left            =   360
               TabIndex        =   50
               Top             =   645
               Width           =   1140
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
               Index           =   5
               Left            =   360
               TabIndex        =   49
               Top             =   345
               Width           =   1380
            End
         End
         Begin FPSpread.vaSpread vaSpread6 
            Height          =   5940
            Left            =   960
            TabIndex        =   45
            Top             =   1440
            Width           =   10095
            _Version        =   393216
            _ExtentX        =   17806
            _ExtentY        =   10477
            _StockProps     =   64
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
            MaxRows         =   0
            SpreadDesigner  =   "M_RegNeg.frx":00BB
         End
      End
      Begin VB.Frame Frame10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7695
         Left            =   -74760
         TabIndex        =   31
         Top             =   480
         Width           =   12135
         Begin VB.Frame Frame12 
            Caption         =   "Casino Incluido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7095
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   5775
            Begin VB.Frame Frame16 
               Height          =   435
               Left            =   1410
               TabIndex        =   41
               Top             =   6600
               Width           =   4005
               Begin VB.TextBox TextCai1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   3
                  Left            =   45
                  TabIndex        =   42
                  Top             =   135
                  Width           =   3900
               End
            End
            Begin VB.Frame Frame13 
               Height          =   435
               Left            =   480
               TabIndex        =   39
               Top             =   6600
               Width           =   915
               Begin VB.TextBox TextCai1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   2
                  Left            =   45
                  TabIndex        =   40
                  Top             =   135
                  Width           =   810
               End
            End
            Begin FPSpread.vaSpread vaSpread4 
               Height          =   6255
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   5535
               _Version        =   393216
               _ExtentX        =   9763
               _ExtentY        =   11033
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
               MaxCols         =   4
               SpreadDesigner  =   "M_RegNeg.frx":0678
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Casino No Incluido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7095
            Left            =   6120
            TabIndex        =   32
            Top             =   360
            Width           =   5775
            Begin VB.Frame Frame14 
               Height          =   435
               Left            =   1530
               TabIndex        =   35
               Top             =   6600
               Width           =   4125
               Begin VB.TextBox TextCan1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   3
                  Left            =   45
                  TabIndex        =   36
                  Top             =   135
                  Width           =   4020
               End
            End
            Begin VB.Frame Frame15 
               Height          =   435
               Left            =   600
               TabIndex        =   33
               Top             =   6600
               Width           =   915
               Begin VB.TextBox TextCan1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   2
                  Left            =   45
                  TabIndex        =   34
                  Top             =   135
                  Width           =   810
               End
            End
            Begin FPSpread.vaSpread vaSpread5 
               Height          =   6255
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   5535
               _Version        =   393216
               _ExtentX        =   9763
               _ExtentY        =   11033
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
               MaxCols         =   4
               SpreadDesigner  =   "M_RegNeg.frx":1F93
            End
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7335
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   12255
         Begin VB.Frame Frame9 
            Height          =   435
            Left            =   3360
            TabIndex        =   29
            Top             =   6840
            Width           =   1035
            Begin VB.TextBox Textp1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   3
               Left            =   45
               TabIndex        =   30
               Top             =   135
               Width           =   930
            End
         End
         Begin VB.Frame Frame8 
            Height          =   435
            Left            =   4485
            TabIndex        =   26
            Top             =   6840
            Width           =   4965
            Begin VB.TextBox Textp1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   4
               Left            =   45
               TabIndex        =   27
               Top             =   135
               Width           =   4860
            End
         End
         Begin VB.Frame Frame7 
            Height          =   435
            Left            =   1320
            TabIndex        =   24
            Top             =   6840
            Width           =   1995
            Begin VB.TextBox Textp1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   2
               Left            =   45
               TabIndex        =   25
               Top             =   135
               Width           =   1890
            End
         End
         Begin FPSpread.vaSpread vaSpread3 
            Height          =   6420
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   12015
            _Version        =   393216
            _ExtentX        =   21193
            _ExtentY        =   11324
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   7
            SpreadDesigner  =   "M_RegNeg.frx":38C3
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7335
         Left            =   -74760
         TabIndex        =   17
         Top             =   480
         Width           =   12255
         Begin VB.Frame Frame4 
            Height          =   435
            Left            =   1320
            TabIndex        =   21
            Top             =   6840
            Width           =   1035
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   2
               Left            =   45
               TabIndex        =   22
               Top             =   135
               Width           =   930
            End
         End
         Begin VB.Frame Frame6 
            Height          =   435
            Left            =   2445
            TabIndex        =   19
            Top             =   6840
            Width           =   7005
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   3
               Left            =   45
               TabIndex        =   20
               Top             =   135
               Width           =   6900
            End
         End
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   6420
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   12015
            _Version        =   393216
            _ExtentX        =   21193
            _ExtentY        =   11324
            _StockProps     =   64
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
            SpreadDesigner  =   "M_RegNeg.frx":546E
         End
      End
      Begin VB.Frame Frame2 
         Height          =   7215
         Left            =   -74760
         TabIndex        =   9
         Top             =   600
         Width           =   12135
         Begin VB.ComboBox Combo2 
            Height          =   315
            Index           =   0
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1560
            Width           =   1800
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   4080
            TabIndex        =   10
            Top             =   1200
            Width           =   4485
            _Version        =   196608
            _ExtentX        =   7911
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
            Index           =   0
            Left            =   4080
            TabIndex        =   11
            Top             =   840
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
            BackColor       =   16777215
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
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   4125
            TabIndex        =   16
            Top             =   1650
            Width           =   1800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Restricción de Rutas"
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
            Left            =   2040
            TabIndex        =   14
            Top             =   1680
            Width           =   1800
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
            Index           =   2
            Left            =   2040
            TabIndex        =   13
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
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
            Left            =   2040
            TabIndex        =   12
            Top             =   1320
            Width           =   1020
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   971
         Left            =   -72480
         TabIndex        =   2
         Top             =   480
         Width           =   7335
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "M_RegNeg.frx":6FAE
            Left            =   2010
            List            =   "M_RegNeg.frx":6FB8
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   2500
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   1
            Left            =   2010
            TabIndex        =   4
            Top             =   555
            Width           =   2505
            _Version        =   196608
            _ExtentX        =   4410
            _ExtentY        =   870
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
            Index           =   0
            Left            =   525
            TabIndex        =   7
            Top             =   345
            Width           =   1380
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
            Index           =   1
            Left            =   525
            TabIndex        =   6
            Top             =   645
            Width           =   1140
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
            Left            =   4590
            TabIndex        =   5
            Top             =   645
            Width           =   585
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6285
         Left            =   -72480
         TabIndex        =   8
         Top             =   1560
         Width           =   7365
         _Version        =   393216
         _ExtentX        =   12991
         _ExtentY        =   11086
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
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
         MaxCols         =   3
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_RegNeg.frx":6FCC
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_RegNeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String, codigo As String, Msgtitulo As String, codproblo As String, spid As Long
Dim Est As Boolean

Private Sub Combo2_Click(Index As Integer)
If Est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 14, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(5) = False
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 9210
Me.Width = 13095
Msgtitulo = "Reglas de Negocios"
fg_centra Me
SSTab1.Tab = 0
modo = ""
Est = True
Gl_Mo_Botones Me, 14
Gl_Ac_Botones Me, 14, 1, modo
Combo1.ListIndex = 1
Combo3.ListIndex = 0
Combo2(0).AddItem "Con Rutas" & Space(150) & "(1)"
Combo2(0).AddItem "Sin Rutas" & Space(150) & "(2)"
Combo2(0).AddItem "Con Sin Rutas" & Space(150) & "(3)"
Combo2(0).ListIndex = -1
MoverDatosGrilla
MoverDatosReglasdeNegocios
MoverDatosReglasdeNegociosFamilia
MoverDatosReglasdeNegociosProducto
MoverDatosReglasdeNegociosCasino
Gl_Ac_Botones Me, 14, IIf(vaSpread1.MaxRows > 0, 1, 2), modo
End Sub

Sub MoverDatosGrilla()
fg_carga ""
Dim x As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread1.TextTip = 2
' Control displays text tips after 250 milliseconds
vaSpread1.TextTipDelay = 250
' Text tip displays custom font and colors
' Background is yellow, RGB(255, 255, 0)
' Foreground is dark blue, RGB(0, 0, 128)
x = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
vaSpread1.Row = -1
vaSpread1.Col = -1
vaSpread1.Lock = True
Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegocios 1, '', ''")
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   vaSpread1.text = RS!rn_codigo
   vaSpread1.Col = 2
   vaSpread1.text = Trim(RS!rn_nombre)
   vaSpread1.Col = 3
   vaSpread1.text = IIf(Trim(RS!rn_tipo_ruta) = "1", "Con Rutas", IIf(Trim(RS!rn_tipo_ruta) = "2", "Sin Rutas", "Con y Sin Rutas"))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
If vaSpread1.MaxRows > 0 Then
   vaSpread1.Row = 1
   vaSpread1.Col = 1
   codigo = ""
   codigo = Val(vaSpread1.text)
   vaSpread1.SetActiveCell 1, 1 ': vaSpread1.SetFocus
End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
fg_descarga
End Sub

Sub MoverDatosReglasdeNegociosFamilia()
Dim auxfam As String, estmar As Boolean
auxfam = ""
fg_carga ""
estmar = False
MoverDatosReglasdeNegocios
Est = True
Limpia 2
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = vaSpread1.text
Dim x As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread2.TextTip = 2
vaSpread2.TextTipDelay = 250
x = vaSpread2.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)

vaSpread2.Visible = False
Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegociosfamilia 1, '" & codigo & "'")
Do While Not RS.EOF
   If auxfam <> RS!codigo Then
      If estmar Then
         vaSpread2.Col = 1
         vaSpread2.text = "1"
      End If
      estmar = False
      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      vaSpread2.Col = 2
      vaSpread2.text = RS!codigo
      vaSpread2.Col = 3
      vaSpread2.text = Trim(IIf(IsNull(RS!descripcioncompleta), "", RS!descripcioncompleta))
      auxfam = RS!codigo
   End If
   vaSpread2.Col = 4
   vaSpread2.text = IIf(IsNull(RS!rnf_pn) Or RS!rnf_pn = "N", "0", "1")
   If Not IsNull(RS!rnf_pn) And RS!rnf_pn <> "N" Then estmar = True
   vaSpread2.Col = 5
   vaSpread2.text = IIf(IsNull(RS!rnf_pa) Or RS!rnf_pa = "N", "0", "1")
   If Not IsNull(RS!rnf_pa) And RS!rnf_pa <> "N" Then estmar = True
   vaSpread2.Col = 6
   vaSpread2.text = IIf(IsNull(RS!rnf_a) Or RS!rnf_a = "N", "0", "1")
   If Not IsNull(RS!rnf_a) And RS!rnf_a <> "N" Then estmar = True
   RS.MoveNext
Loop
If estmar Then
   vaSpread2.Col = 1
   vaSpread2.text = "1"
End If
RS.Close: Set RS = Nothing
vaSpread2.Visible = True
Est = False
fg_descarga
End Sub

Sub MoverDatosReglasdeNegociosProducto()
Dim auxpro As String, estmar As Boolean
auxfam = ""
fg_carga ""
estmar = False
MoverDatosReglasdeNegocios
Est = True
Limpia 3
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = vaSpread1.text
Dim x As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread3.TextTip = 2
vaSpread3.TextTipDelay = 250
x = vaSpread3.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)

vaSpread3.Visible = False
Set RS = vg_dbpedweb.Execute("select @@Spid  spid")
If Not RS.EOF Then
   spid = RS!spid
End If
RS.Close: Set RS = Nothing
Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegociosproducto 1, '" & codigo & "', '', '" & vg_NUsr & "', " & spid & "")
Do While Not RS.EOF
   vaSpread3.MaxRows = vaSpread3.MaxRows + 1
   vaSpread3.Row = vaSpread3.MaxRows
   vaSpread3.Col = 2
   vaSpread3.text = IIf(IsNull(RS!NombreFamiliaCompleta), "", RS!NombreFamiliaCompleta)
   vaSpread3.Col = 3
   vaSpread3.text = Trim(IIf(IsNull(RS!codigo), "", RS!codigo))
   vaSpread3.Col = 4
   vaSpread3.text = Trim(IIf(IsNull(RS!descripcion), "", RS!descripcion))
   vaSpread3.Col = 5
   vaSpread3.text = IIf(IsNull(RS!rnp_pn) Or RS!rnp_pn = "N", "0", "1")
   vaSpread3.Col = 6
   vaSpread3.text = IIf(IsNull(RS!rnp_pa) Or RS!rnp_pa = "N", "0", "1")
   vaSpread3.Col = 7
   vaSpread3.text = IIf(IsNull(RS!rnp_a) Or RS!rnp_a = "N", "0", "1")
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread3.Visible = True
Est = False
fg_descarga
End Sub

Sub MoverDatosReglasdeNegociosCasino()
fg_carga ""
MoverDatosReglasdeNegocios
Est = True
Limpia 4
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = vaSpread1.text
Dim x As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread4.TextTip = 2
vaSpread4.TextTipDelay = 250
x = vaSpread4.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)

vaSpread4.Visible = False
Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegocioscasino 1, '" & codigo & "'")
Do While Not RS.EOF
   vaSpread4.MaxRows = vaSpread4.MaxRows + 1
   vaSpread4.Row = vaSpread4.MaxRows
   vaSpread4.Col = 2
   vaSpread4.text = IIf(IsNull(RS!centrocosto), "", RS!centrocosto)
   vaSpread4.Col = 3
   vaSpread4.text = Trim(IIf(IsNull(RS!Nombre), "", RS!Nombre))
   vaSpread4.Col = 4
   vaSpread4.text = Trim(IIf(IsNull(RS!codigo_casino), "", RS!codigo_casino))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread4.Visible = True

' Control displays text tips aligned to pointer with focus
vaSpread5.TextTip = 2
vaSpread5.TextTipDelay = 250
x = vaSpread5.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread5.Visible = False
Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegocioscasinonoinc 1, '" & codigo & "'")
Do While Not RS.EOF
   vaSpread5.MaxRows = vaSpread5.MaxRows + 1
   vaSpread5.Row = vaSpread5.MaxRows
   vaSpread5.Col = 2
   vaSpread5.text = IIf(IsNull(RS!centrocosto), "", RS!centrocosto)
   vaSpread5.Col = 3
   vaSpread5.text = Trim(IIf(IsNull(RS!Nombre), "", RS!Nombre))
   vaSpread5.Col = 4
   vaSpread5.text = Trim(IIf(IsNull(RS!codigo), "", RS!codigo))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread5.Visible = True
Est = False
fg_descarga
End Sub

Sub MoverDatosReglasdeNegocios()
fg_carga ""
Est = True
Limpia 1
Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegocios 4, '" & codigo & "', ''")
If Not RS.EOF Then
   fpLongInteger1(0).Value = RS!rn_codigo
   fpText1(0).text = Trim(RS!rn_nombre)
   Combo2(0).ListIndex = fg_buscacbo(Combo2, 0, 1, (RS!rn_tipo_ruta))
   Frame3.Caption = RS!rn_codigo & " - " & Trim(RS!rn_nombre) & " - " & IIf(fg_buscacbo(Combo2, 0, 1, (RS!rn_tipo_ruta)) = "0", "Con Rutas", IIf(fg_buscacbo(Combo2, 0, 1, (RS!rn_tipo_ruta)) = "1", "Sin Rutas", "Con Sin Rutas"))
   Frame5.Caption = RS!rn_codigo & " - " & Trim(RS!rn_nombre) & " - " & IIf(fg_buscacbo(Combo2, 0, 1, (RS!rn_tipo_ruta)) = "0", "Con Rutas", IIf(fg_buscacbo(Combo2, 0, 1, (RS!rn_tipo_ruta)) = "1", "Sin Rutas", "Con Sin Rutas"))
   Frame10.Caption = RS!rn_codigo & " - " & Trim(RS!rn_nombre) & " - " & IIf(fg_buscacbo(Combo2, 0, 1, (RS!rn_tipo_ruta)) = "0", "Con Rutas", IIf(fg_buscacbo(Combo2, 0, 1, (RS!rn_tipo_ruta)) = "1", "Sin Rutas", "Con Sin Rutas"))
End If
RS.Close: Set RS = Nothing
fg_descarga
Est = False
End Sub

Private Sub fpText1_Change(Index As Integer)
Select Case Index
Case 0
    If Est Then Exit Sub
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 14, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False
Case 1
    If LimpiaDato(Trim(fpText1(1).text)) & Chr(KeyAscii) = "" Then Exit Sub
    vaSpread1.Visible = False
    vaSpread1.Row = -1
    vaSpread1.Col = -1
    vaSpread1.Lock = True
    If Combo1.ItemData(Combo1.ListIndex) = 0 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_reglasdenegocios 2, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%'")
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_reglasdenegocios 3, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%'")
    End If
    If RS2.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS2!nReg
    i = 1
    If Not RS2.EOF Then
       Do While Not RS2.EOF
          vaSpread1.Row = i: i = i + 1
          vaSpread1.Col = 1
          vaSpread1.TypeHAlign = 1
          vaSpread1.text = RS2!rn_codigo
          vaSpread1.Col = 2
          vaSpread1.text = Trim(RS2!rn_nombre)
          vaSpread1.Col = 3
          vaSpread1.text = IIf(Trim(RS2!rn_tipo_ruta) = "1", "Con Rutas", IIf(Trim(RS2!rn_tipo_ruta) = "2", "Sin Rutas", "Con y Sin Rutas"))
          RS2.MoveNext
        Loop
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = True
        SSTab1.TabEnabled(4) = True
        SSTab1.TabEnabled(5) = True
        Gl_Ac_Botones Me, 14, 1, modo
    Else
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
        SSTab1.TabEnabled(5) = False
    End If
    RS2.Close: Set RS2 = Nothing
    vaSpread1.Col = 1: vaSpread1.Col2 = vaSpread1.MaxCols: vaSpread1.Row = 1: vaSpread1.Row2 = vaSpread1.MaxRows
    vaSpread1.SetActiveCell 1, 1
    vaSpread1.Visible = True
    If fpText1(1).text = "" Then Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro" Else Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
End Select
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Select Case Index
Case 2
    Est = True
    If Combo1.ItemData(Combo3.ListIndex) = 0 Then
'       Set RS2 = vg_dbpedweb.Execute("pedweb_s_buscarproductos 1, '%" & UCase(LimpiaDato(fpText1(2).Text)) & "%'")
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_buscarproductos 1, '" & UCase(LimpiaDato(fpText1(2).text)) & "'")
    ElseIf Combo1.ItemData(Combo3.ListIndex) = 1 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_buscarproductos 2, '%" & UCase(LimpiaDato(fpText1(2).text)) & "%'")
    End If
    vaSpread6.Visible = False
    vaSpread6.MaxRows = 0
    Label1(7).Caption = "No existe producto"
    codproblo = ""
    Do While Not RS2.EOF
       If IsNull(RS2!codi) Then
          RS2.Close: Set RS2 = Nothing
          vaSpread6.Visible = True
          Label1(7).Visible = True
          Est = False
          Exit Sub
       End If
       Label1(7).Caption = Trim(RS2!codi) & " - " & Trim(RS2!desc1)
       codproblo = RS2!codi
       vaSpread6.MaxRows = vaSpread6.MaxRows + 1
       vaSpread6.Row = vaSpread6.MaxRows
       
       vaSpread6.Col = 1
       vaSpread6.text = "0"
       
       vaSpread6.Col = 2
       vaSpread6.Lock = True
       vaSpread6.text = RS2!rn_codigo
       
       vaSpread6.Col = 3
       vaSpread6.Lock = True
       vaSpread6.text = Trim(RS2!rn_nombre)
       
       vaSpread6.Col = 4
       vaSpread6.text = IIf(IsNull(RS2!rnp_pn) Or RS2!rnp_pn = "N", "0", "1")
       
       vaSpread6.Col = 5
       vaSpread6.text = IIf(IsNull(RS2!rnp_pa) Or RS2!rnp_pa = "N", "0", "1")
       
       vaSpread6.Col = 6
       vaSpread6.text = IIf(IsNull(RS2!rnp_a) Or RS2!rnp_a = "N", "0", "1")
       
       RS2.MoveNext
    Loop
    RS2.Close: Set RS2 = Nothing
    vaSpread6.Visible = True
    Label1(7).Visible = True
    Est = False
'    SendKeys "{Tab}"
End Select
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If Est Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = vaSpread1.text
Select Case SSTab1.Tab
Case 0
    If Toolbar1.Buttons(5).Visible = True Then Toolbar1.Buttons(5).Enabled = True
Case 1
    MoverDatosReglasdeNegocios
    If Toolbar1.Buttons(5).Visible = True Then Toolbar1.Buttons(5).Enabled = True
Case 2
    MoverDatosReglasdeNegociosFamilia
    If Toolbar1.Buttons(5).Visible = True Then Toolbar1.Buttons(5).Enabled = False
Case 3
    MoverDatosReglasdeNegociosProducto
    If Toolbar1.Buttons(5).Visible = True Then Toolbar1.Buttons(5).Enabled = False
Case 4
    MoverDatosReglasdeNegociosCasino
    If Toolbar1.Buttons(5).Visible = True Then Toolbar1.Buttons(5).Enabled = True
Case 5
    fpText1(2).text = ""
    Label1(7).Caption = ""
    Label1(7).Visible = False
    vaSpread6.MaxRows = 0
    If Toolbar1.Buttons(5).Visible = True Then Toolbar1.Buttons(5).Enabled = False
End Select
End Sub

Sub Limpia(op As Integer)
Select Case op
Case 1
    fpLongInteger1(0).Value = ""
    fpLongInteger1(0).Enabled = False
    fpText1(0).text = ""
    Combo2(0).ListIndex = -1
    Frame3.Caption = ""
    Frame5.Caption = ""
    Frame10.Caption = ""
Case 2
    vaSpread2.MaxRows = 0
Case 3
    vaSpread3.MaxRows = 0
Case 4
    vaSpread4.MaxRows = 0
    vaSpread5.MaxRows = 0
End Select
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 2, 3
    vaSpread2.Visible = False
    If Trim(Text1(Index).text) <> "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           vaSpread2.Col = Index: nom = UCase(Trim(vaSpread2.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread2.Col = 2
           If indactivo = -1 And Trim(vaSpread2.text) <> "" Then
              If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
           Else
              If vaSpread2.RowHidden = False Then vaSpread2.RowHidden = True
           End If
        Next i
        vaSpread2.SetActiveCell Index, 1
    End If
    vaSpread2.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread2.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread2.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread2.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread2.Sort -1, -1, vaSpread2.MaxCols, vaSpread2.MaxRows, SortByRow
    If Trim(Text1(Index).text) = "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
       Next
       vaSpread2.SetActiveCell Index, vaSpread2.SearchCol(Index, 0, vaSpread2.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread2.SetActiveCell Index, 1
    End If
    vaSpread2.Visible = True
End Select
End Sub

Private Sub TextCai1_Change(Index As Integer)
Dim i As Long, nom As String
Select Case Index
Case 2, 3, 4
    vaSpread4.Visible = False
    If Trim(TextCai1(Index).text) <> "" Then
       For i = 1 To vaSpread4.MaxRows
           vaSpread4.Row = i
           vaSpread4.Col = Index: nom = UCase(Trim(vaSpread4.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCai1(Index).text) & "*"
           vaSpread4.Col = 2
           If indactivo = -1 And Trim(vaSpread4.text) <> "" Then
              If vaSpread4.RowHidden = True Then vaSpread4.RowHidden = False
           Else
              If vaSpread4.RowHidden = False Then vaSpread4.RowHidden = True
           End If
        Next i
        vaSpread4.SetActiveCell Index, 1
    End If
    vaSpread4.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread4.ColUserSortIndicator(IIf(Trim(TextCai1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread4.SortKey(1) = IIf(Trim(TextCai1(Index).text) = "", 0, 0): vaSpread4.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread4.Sort -1, -1, vaSpread4.MaxCols, vaSpread4.MaxRows, SortByRow
    If Trim(TextCai1(Index).text) = "" Then
       For i = 1 To vaSpread4.MaxRows
           vaSpread4.Row = i
           If vaSpread4.RowHidden = True Then vaSpread4.RowHidden = False
       Next
       vaSpread4.SetActiveCell Index, vaSpread4.SearchCol(Index, 0, vaSpread4.MaxRows, Trim(TextCai1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread4.SetActiveCell Index, 1
    End If
    vaSpread4.Visible = True
End Select
End Sub

Private Sub TextCan1_Change(Index As Integer)
Dim i As Long, nom As String
Select Case Index
Case 2, 3, 4
    vaSpread5.Visible = False
    If Trim(TextCan1(Index).text) <> "" Then
       For i = 1 To vaSpread5.MaxRows
           vaSpread5.Row = i
           vaSpread5.Col = Index: nom = UCase(Trim(vaSpread5.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCan1(Index).text) & "*"
           vaSpread5.Col = 2
           If indactivo = -1 And Trim(vaSpread5.text) <> "" Then
              If vaSpread5.RowHidden = True Then vaSpread5.RowHidden = False
           Else
              If vaSpread5.RowHidden = False Then vaSpread5.RowHidden = True
           End If
        Next i
        vaSpread5.SetActiveCell Index, 1
    End If
    vaSpread5.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread5.ColUserSortIndicator(IIf(Trim(TextCan1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread5.SortKey(1) = IIf(Trim(TextCan1(Index).text) = "", 0, 0): vaSpread5.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread5.Sort -1, -1, vaSpread5.MaxCols, vaSpread5.MaxRows, SortByRow
    If Trim(TextCan1(Index).text) = "" Then
       For i = 1 To vaSpread5.MaxRows
           vaSpread5.Row = i
           If vaSpread5.RowHidden = True Then vaSpread5.RowHidden = False
       Next
       vaSpread5.SetActiveCell Index, vaSpread5.SearchCol(Index, 0, vaSpread5.MaxRows, Trim(TextCan1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread5.SetActiveCell Index, 1
    End If
    vaSpread5.Visible = True
End Select
End Sub

Private Sub Textp1_Change(Index As Integer)
Dim i As Long, nom As String
Select Case Index
Case 2, 3, 4
    vaSpread3.Visible = False
    If Trim(Textp1(Index).text) <> "" Then
       For i = 1 To vaSpread3.MaxRows
           vaSpread3.Row = i
           vaSpread3.Col = Index: nom = UCase(Trim(vaSpread3.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(Textp1(Index).text) & "*"
           vaSpread3.Col = 2
           If indactivo = -1 And Trim(vaSpread3.text) <> "" Then
              If vaSpread3.RowHidden = True Then vaSpread3.RowHidden = False
           Else
              If vaSpread3.RowHidden = False Then vaSpread3.RowHidden = True
           End If
        Next i
        vaSpread3.SetActiveCell Index, 1
    End If
    vaSpread3.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread3.ColUserSortIndicator(IIf(Trim(Textp1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread3.SortKey(1) = IIf(Trim(Textp1(Index).text) = "", 0, 0): vaSpread3.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread3.Sort -1, -1, vaSpread3.MaxCols, vaSpread3.MaxRows, SortByRow
    If Trim(Textp1(Index).text) = "" Then
       For i = 1 To vaSpread3.MaxRows
           vaSpread3.Row = i
           If vaSpread3.RowHidden = True Then vaSpread3.RowHidden = False
       Next
       vaSpread3.SetActiveCell Index, vaSpread3.SearchCol(Index, 0, vaSpread3.MaxRows, Trim(Textp1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread3.SetActiveCell Index, 1
    End If
    vaSpread3.Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codfam As String, rnf_pn As String, rnf_pa As String, rnf_a As String, codpro As String, codregneg As Long
Dim i As Long
Select Case Button.Index
Case 1 '-------> Incluir nuevos registros
    modo = "A"
    Select Case SSTab1.Tab
    Case 0, 1 '-------> Ruta
        Est = True
        SSTab1.Tab = 1
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
        SSTab1.TabEnabled(5) = False
        '-------> Traer ultimo registro
        Limpia 1
        Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegocios 5, '', ''")
        If Not RS.EOF Then RS.MoveFirst: codigo = RS!rn_codigo + 1 Else codigo = 1
        RS.Close: Set RS = Nothing
        fpLongInteger1(0).text = codigo
        fpText1(0).SetFocus
        vg_codigo = "x"
        Est = False
    Case 2 '-------> Familia
    Case 3, 5 '-------> Producto y bloquear productos
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = IIf(SSTab1.Tab = 3, True, False)
        SSTab1.TabEnabled(4) = False
        SSTab1.TabEnabled(5) = IIf(SSTab1.Tab = 3, False, True)
        vg_codigo = "x"
        vg_codigo = ""
        B_RutPro.LlenaDatos codigo, Frame5.Caption, IIf(SSTab1.Tab = 3, "regpro", "regunpro"), spid
        B_RutPro.Show 1
        If Trim(vg_codigo) = "" Then
           SSTab1.TabEnabled(0) = True
           SSTab1.TabEnabled(1) = True
           SSTab1.TabEnabled(2) = True
           SSTab1.TabEnabled(3) = True
           SSTab1.TabEnabled(4) = True
           SSTab1.TabEnabled(5) = True
           Exit Sub
        ElseIf SSTab1.Tab = 5 Then
           fpText1(2).text = vg_codigo
           SendKeys "{Tab}"
           Label1(7).Caption = ""
           Label1(7).Visible = True
           Label1(7).Caption = vg_nombre
           fpText1_KeyPress 2, 13
           Exit Sub
        End If
    Case 4 '-------> Inlcuir casino a la reglas de negocios
        vg_codigo = ""
        If vaSpread5.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        '-------> Validar que exista un registro ruta producto seleccionado
        For i = 1 To vaSpread5.MaxRows
            vaSpread5.Row = i
            vaSpread5.Col = 1
            If vaSpread5.text = "1" Then estmar = True
        Next i
        If Not estmar Then MsgBox "Debe seleccionar un registro de los casino no incluidos...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        For i = 1 To vaSpread5.MaxRows
            vaSpread5.Row = i
            vaSpread5.Col = 1
            If vaSpread5.text = "1" Then
               vaSpread5.Col = 2
               auxcen = vaSpread5.text
               vaSpread5.Col = 3
               nomcen = vaSpread5.text
               vaSpread5.Col = 4
               cencos = vaSpread5.text
               '-------> Mover datos a vector de casino incluido en la reglas de negocios
               vaSpread4.MaxRows = vaSpread4.MaxRows + 1
               vaSpread4.Row = vaSpread4.MaxRows
               vaSpread4.Col = -1
               vaSpread4.BackColor = &H80000013
               vaSpread4.Col = 1
               vaSpread4.text = "1"
               vaSpread4.Col = 2
               vaSpread4.text = auxcen
               vaSpread4.Col = 3
               vaSpread4.text = nomcen
               vaSpread4.Col = 4
               vaSpread4.text = cencos
               vaSpread4.SetActiveCell 2, vaSpread4.MaxRows
            End If
        Next i
        vg_codigo = "X"
        '-------> Bloquer grilla casino incluidos
        vaSpread4.Row = -1
        vaSpread4.Col = -1
        vaSpread4.Lock = True
        '-------> Bloquear grilla casino no incluido
        vaSpread5.Row = -1
        vaSpread5.Col = -1
        vaSpread5.Lock = True
        '-------> Bloquear hoja
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(5) = False
        SSTab1.TabEnabled(4) = True
    End Select
    If vg_codigo <> "" Then Gl_Ac_Botones Me, 14, 0, modo
Case 3 '-------> Alterar registro
    Select Case SSTab1.Tab
    Case 0, 1 '-------> Reglas de Negocios
        modo = "M"
        Gl_Ac_Botones Me, 14, 0, modo
        SSTab1.Tab = 1
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
        SSTab1.TabEnabled(5) = False
        fpText1(0).SetFocus
    Case 2 '-------> Familia
        modo = "M"
        Gl_Ac_Botones Me, 14, 0, modo
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
        SSTab1.TabEnabled(5) = False
    Case 3
        modo = "M"
        Gl_Ac_Botones Me, 14, 0, modo
        SSTab1.Tab = 3
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = True
        SSTab1.TabEnabled(5) = True
        SSTab1.TabEnabled(4) = False
    End Select
Case 5 '-------> Eliminar Registro y sus relaciones
    Select Case SSTab1.Tab
    Case 0, 1, 2, 3 '-------> Reglas de Negocios
        If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        If MsgBox("Elimina registro y todas sus relaciones...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        vg_dbpedweb.Execute ("pedweb_d_reglasdenegocios '" & codigo & "'")
        SSTab1.Tab = 0
        MoverDatosGrilla
        MoverDatosReglasdeNegocios
    Case 4
        '-------> Validar que exista un registro reglas de negocios seleccionado
        If vaSpread4.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        For i = 1 To vaSpread4.MaxRows
            vaSpread4.Row = i
            vaSpread4.Col = 1
            If vaSpread4.text = "1" Then estmar = True
        Next i
        If Not estmar Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
        '-------> rutina de borrado reglas de negocios casino
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        For i = 1 To vaSpread4.MaxRows
            vaSpread4.Row = i
            vaSpread4.Col = 1
            If vaSpread4.text = "1" Then
               vaSpread4.Col = 4
               cencos = Trim(vaSpread4.text)
               vg_dbpedweb.Execute ("pedweb_d_reglasdenegocioscasino " & codigo & ", '" & cencos & "'")
            End If
        Next i
        MoverDatosReglasdeNegociosCasino
    End Select
    modo = "": Gl_Ac_Botones Me, 14, 1, modo
Case 7 '-------> Actualizar lista
    Select Case SSTab1.Tab
    Case 0
        MoverDatosGrilla
        fpText1(1).text = ""
        Gl_Ac_Botones Me, 14, IIf(vaSpread1.MaxRows > 0, 1, 2), modo
    Case 1
        MoverDatosReglasdeNegocios
    Case 2
        MoverDatosReglasdeNegociosFamilia
    Case 3
        MoverDatosReglasdeNegociosProducto
    Case 4
        MoverDatosReglasdeNegociosCasino
    End Select
Case 10 '-------> Cancelar Información
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    Select Case SSTab1.Tab
    Case 1
        SSTab1.Tab = 1
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        MoverDatosReglasdeNegocios
    Case 2
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        MoverDatosReglasdeNegociosFamilia
    Case 3
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        MoverDatosReglasdeNegociosProducto
    Case 4
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        MoverDatosReglasdeNegociosCasino
    Case 5
       vaSpread6.MaxRows = 0
       fpText1(2).text = ""
       Label1(7).Visible = False
       Label1(7).Caption = ""
    End Select
    '-------> Desbloquear hojas
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    SSTab1.TabEnabled(4) = True
    SSTab1.TabEnabled(5) = True
    modo = "": Gl_Ac_Botones Me, 14, 1, modo
Case 12 '-------> GrabaRegistro
    Dim tipreg As String
    Select Case SSTab1.Tab
    Case 1 '-------> Grabar Reglas de Negocios
         tipreg = fg_codigocbo(Combo2, 0, 1, "")
        If LimpiaDato(Trim(fpText1(0).text)) = "" Then MsgBox "Debe ingresar información...", vbCritical, Msgtitulo: Exit Sub
        If modo = "A" Then
           codigo = 0
           Set RS = vg_dbpedweb.Execute("pedweb_iu_reglasdenegocios 'A', 0, '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & tipreg & "', '" & vg_NUsr & "', '', '', ''")
           If Not RS.EOF Then
              codigo = RS!indice
           End If
           RS.Close: Set RS = Nothing
           fpLongInteger1(0).text = codigo
           vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
           vaSpread1.SetActiveCell 1, vaSpread1.Row
        Else
            vg_dbpedweb.Execute "pedweb_iu_reglasdenegocios 'M', " & codigo & ", '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & tipreg & "', '" & vg_NUsr & "', '', '" & vg_NUsr & "', ''"
        End If
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = LimpiaDato(Trim(fpLongInteger1(0).text))
        vaSpread1.Col = 2: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = LimpiaDato(Trim(fpText1(0).text))
        vaSpread1.Col = 3: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = IIf(Trim(tipreg) = "1", "Con Rutas", IIf(Trim(tipreg) = "2", "Sin Rutas", "Con y Sin Rutas"))
    Case 2 '-------> Grabar familia
        fg_carga ""
        vg_dbpedweb.Execute ("DELETE s_RNFamilias WHERE rn_codigo = " & codigo & "")
        For i = 1 To vaSpread2.MaxRows
            vaSpread2.Row = i
            vaSpread2.Col = 1
            If vaSpread2.text = "1" Then
               vaSpread2.Col = 2
               codfam = vaSpread2.text
               vaSpread2.Col = 4
               rnf_pn = IIf(vaSpread2.text = "1", "S", "N")
               vaSpread2.Col = 5
               rnf_pa = IIf(vaSpread2.text = "1", "S", "N")
               vaSpread2.Col = 6
               rnf_a = IIf(vaSpread2.text = "1", "S", "N")
               vg_dbpedweb.Execute ("INSERT INTO s_RNFamilias VALUES ('" & codfam & "', " & codigo & ", '" & rnf_pn & "', '" & rnf_pa & "', '" & rnf_a & "')")
            End If
        Next i
        fg_descarga
    Case 3 '-------> Grabar producto
        fg_carga ""
        For i = 1 To vaSpread3.MaxRows
            vaSpread3.Row = i
            vaSpread3.Col = 3
            codpro = vaSpread3.text
            vaSpread3.Col = 5
            rnf_pn = IIf(vaSpread3.text = "1", "S", "N")
            vaSpread3.Col = 6
            rnf_pa = IIf(vaSpread3.text = "1", "S", "N")
            vaSpread3.Col = 7
            rnf_a = IIf(vaSpread3.text = "1", "S", "N")
            vaSpread3.Col = 1
            If vaSpread3.text = "1" Then
               vg_dbpedweb.Execute ("DELETE s_RNProductos WHERE codigo_producto = '" & codpro & "' AND rn_codigo = " & codigo & "")
               vg_dbpedweb.Execute ("INSERT INTO s_RNProductos VALUES ('" & codpro & "', " & codigo & ", '" & rnf_pn & "', '" & rnf_pa & "', '" & rnf_a & "')")
            ElseIf Trim(rnf_pn) = "N" And Trim(rnf_pa) = "N" And Trim(rnf_a) = "N" Then
               vg_dbpedweb.Execute ("DELETE s_RNProductos WHERE codigo_producto = '" & codpro & "' AND rn_codigo = " & codigo & "")
            End If
        Next i
        MoverDatosReglasdeNegociosProducto
        fg_descarga
    Case 4 '-------> Grabar casino reglas de negocios
        fg_carga ""
        For i = 1 To vaSpread4.MaxRows
            vaSpread4.Row = i
            vaSpread4.Col = 1
            If vaSpread4.text = "1" Then
               vaSpread4.Col = 4
               cencos = vaSpread4.text
               vg_dbpedweb.Execute ("INSERT INTO s_RNCasino VALUES ('" & cencos & "', " & codigo & ")")
            End If
        Next i
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        MoverDatosReglasdeNegociosCasino
    Case 5 '-------> Grabar productos a bloquear
        fg_carga ""
        For i = 1 To vaSpread6.MaxRows
            vaSpread6.Row = i
            vaSpread6.Col = 2
            codregneg = Val(vaSpread6.text)
            vaSpread6.Col = 4
            rnf_pn = IIf(vaSpread6.text = "1", "S", "N")
            vaSpread6.Col = 5
            rnf_pa = IIf(vaSpread6.text = "1", "S", "N")
            vaSpread6.Col = 6
            rnf_a = IIf(vaSpread6.text = "1", "S", "N")
            vaSpread6.Col = 1
            If vaSpread6.text = "1" Then
               vg_dbpedweb.Execute ("DELETE s_RNProductos WHERE codigo_producto = '" & codproblo & "' AND rn_codigo = " & codregneg & "")
               vg_dbpedweb.Execute ("INSERT INTO s_RNProductos VALUES ('" & codproblo & "', " & codregneg & ", '" & rnf_pn & "', '" & rnf_pa & "', '" & rnf_a & "')")
            ElseIf Trim(rnf_pn) = "N" And Trim(rnf_pa) = "N" And Trim(rnf_a) = "N" Then
               vg_dbpedweb.Execute ("DELETE s_RNProductos WHERE codigo_producto = '" & codproblo & "' AND rn_codigo = " & codregneg & "")
            End If
        Next i
        fpText1(2).text = ""
        Label1(7).Caption = ""
        Label1(7).Visible = False
        vaSpread6.MaxRows = 0
        fg_descarga
    End Select
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    SSTab1.TabEnabled(4) = True
    SSTab1.TabEnabled(5) = True
    modo = "": Gl_Ac_Botones Me, 14, 1, modo
Case 19 '------> impresion
    Select Case SSTab1.Tab
    Case 0, 1 '-------> Reglas de Negocios
        I_ReglasdeNegocios
    Case 2 '-------> Familia
        I_WebRep.LlenaDatos "Impresión Regla de Negocios Familia", "regnegfam"
        I_WebRep.Show 1
        Me.Refresh

'        vaSpread1.Row = vaSpread1.ActiveRow
'        vaSpread1.Col = 1
'        codigo = vaSpread1.Text
'        I_ReglasdeNegociosFamilia CStr(codigo)
    Case 3 '-------> Ruta Calendarios
        vg_opimp = 999999
        I_WebRep.LlenaDatos "Impresión Regla de Negocios Productos", "regnegpro"
        I_WebRep.Show 1
        Me.Refresh
        vg_opimp = 0
'        vaSpread1.Row = vaSpread1.ActiveRow
'        vaSpread1.Col = 1
'        codigo = vaSpread1.Text
'        I_ReglasdeNegociosProducto CStr(codigo)
    Case 4
        I_WebRep.LlenaDatos "Impresión Regla de Negocios Casinos", "regnegcas"
        I_WebRep.Show 1
        Me.Refresh
'        vaSpread1.Row = vaSpread1.ActiveRow
'        vaSpread1.Col = 1
'        codigo = vaSpread1.Text
'        I_ReglasdeNegociosCasino CStr(codigo)
    End Select
Case 22
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu
Case "Copiar Datos"
    vg_codigo = "x"
    vg_codigo = ""
    M_CopProCan.LlenaDatos "Copiar Reglas de Negocios", "regneg"
    M_CopProCan.Show 1
    Me.Refresh
Case "Importar Datos"
    If vaSpread1.MaxRows < 1 Then Exit Sub
    vg_codigo = ""
    P_ImpRut.LlenaDatos "Importar Reglas de Negocios", "regneg"
    P_ImpRut.Show 1
    If vg_codigo <> "" And SSTab1.Tab = 3 Then
       MoverDatosReglasdeNegociosProducto
    End If
End Select
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread1.Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
Select Case Col
Case 1
    vaSpread1.Col = Col
    TipText = "Código : " & vaSpread1.text
Case 2
    vaSpread1.Col = Col
    TipText = "Descripción : " & Trim(vaSpread1.text)
Case 3
    vaSpread1.Col = Col
    TipText = "Restrinción Ruta : " & Trim(vaSpread1.text)
End Select
End Sub

Private Sub vaSpread2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
Dim i As Long
vaSpread2.Col = 1
For i = BlockRow To BlockRow2
    vaSpread2.Row = i
    vaSpread2.Value = IIf(vaSpread2.Value = "1", "0", "1")
Next
End Sub

Private Sub vaSpread2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If Est Then Exit Sub
Dim i As Long, estmar As Boolean
vaSpread2.Row = Row
Select Case Col
Case 1
    vaSpread2.Col = 4
    Est = True
'    vaSpread2.Text = IIf(vaSpread2.Text = "1", "0", "1")
    vaSpread2.text = IIf(ButtonDown = 1, "1", "0")
    Est = False
    vaSpread2.Col = 5
    Est = True
'    vaSpread2.Text = IIf(vaSpread2.Text = "1", "0", "1")
    vaSpread2.text = IIf(ButtonDown = 1, "1", "0")
    Est = False
    vaSpread2.Col = 6
    Est = True
'    vaSpread2.Text = IIf(vaSpread2.Text = "1", "0", "1")
    vaSpread2.text = IIf(ButtonDown = 1, "1", "0")
    Est = False
Case 4, 5, 6
    estmar = False
    For i = 4 To 6
        vaSpread2.Col = i
        Est = True
        If vaSpread2.text = "1" Then estmar = True
        Est = False
    Next i
    vaSpread2.Col = 1
    Est = True
    vaSpread2.text = IIf(estmar, "1", "0")
    Est = False
End Select
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 14, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = True
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(5) = False
End Sub

Private Sub vaSpread2_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If vaSpread2.MaxRows < 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread2.Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
Select Case Col
Case 2
    vaSpread2.Col = Col
    TipText = "Código : " & vaSpread2.text
Case 3
    vaSpread2.Col = Col
    TipText = "Descripción : " & Trim(vaSpread2.text)
End Select
End Sub

Private Sub vaSpread3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
Dim i As Long
vaSpread3.Col = 1
For i = BlockRow To BlockRow2
    vaSpread3.Row = i
    vaSpread3.Value = IIf(vaSpread3.Value = "1", "0", "1")
Next
End Sub

Private Sub vaSpread3_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If Est Then Exit Sub
Dim i As Long, estmar As Boolean
vaSpread3.Row = Row
Select Case Col
Case 1
    vaSpread3.Col = 5
    Est = True
'    vaSpread3.Text = IIf(vaSpread3.Text = "1", "0", "1")
    vaSpread3.text = IIf(ButtonDown = 1, "1", "0")
    Est = False
    vaSpread3.Col = 6
    Est = True
'    vaSpread3.Text = IIf(vaSpread3.Text = "1", "0", "1")
    vaSpread3.text = IIf(ButtonDown = 1, "1", "0")
    Est = False
    vaSpread3.Col = 7
    Est = True
    vaSpread3.text = IIf(ButtonDown = 1, "1", "0")
'    vaSpread3.Text = IIf(vaSpread3.Text = "1", "0", "1")
    Est = False
Case 5, 6, 7
    estmar = False
    For i = 5 To 7
        vaSpread3.Col = i
        Est = True
        If vaSpread3.text = "1" Then estmar = True
        Est = False
    Next i
    vaSpread3.Col = 1
    Est = True
    vaSpread3.text = IIf(estmar, "1", "0")
    Est = False
End Select
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 14, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = True
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(5) = False
End Sub

Private Sub vaSpread3_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If vaSpread3.MaxRows < 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread3.Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
Select Case Col
Case 2
    vaSpread3.Col = Col
    TipText = "Familia : " & vaSpread3.text
Case 3
    vaSpread3.Col = Col
    TipText = "Código : " & Trim(vaSpread3.text)
Case 4
    vaSpread3.Col = Col
    TipText = "Descripción : " & Trim(vaSpread3.text)
End Select
End Sub

Private Sub vaSpread6_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If Est Then Exit Sub
Dim i As Long, estmar As Boolean
vaSpread6.Row = Row
Select Case Col
Case 1
    vaSpread6.Col = 4
    Est = True
    vaSpread6.text = IIf(vaSpread6.text = "1", "0", "1")
    Est = False
    vaSpread6.Col = 5
    Est = True
    vaSpread6.text = IIf(vaSpread6.text = "1", "0", "1")
    Est = False
    vaSpread6.Col = 6
    Est = True
    vaSpread6.text = IIf(vaSpread6.text = "1", "0", "1")
    Est = False
Case 4, 5, 6
    estmar = False
    For i = 4 To 6
        vaSpread6.Col = i
        Est = True
        If vaSpread6.text = "1" Then estmar = True
        Est = False
    Next i
    vaSpread6.Col = 1
    Est = True
    vaSpread6.text = IIf(estmar, "1", "0")
    Est = False
End Select
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 14, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
End Sub
