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
   ClientHeight    =   7605
   ClientLeft      =   2055
   ClientTop       =   1785
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   4
      OLEDropMode     =   1
      TabCaption(0)   =   "Usuario"
      TabPicture(0)   =   "M_Usuari.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
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
      Tab(1).Control(9)=   "Label1(7)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(8)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(13)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "fpText(7)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "fpText(5)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "fpText(0)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "fpText(3)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "fpText(4)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "fpText(2)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "fpText(1)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Combo2(0)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "fpText(6)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Combo3(1)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Check3"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Text2"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Check4"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).ControlCount=   25
      TabCaption(2)   =   "Perfiles"
      TabPicture(2)   =   "M_Usuari.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Asignar Contratos"
      TabPicture(3)   =   "M_Usuari.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "Frame4"
      Tab(3).ControlCount=   2
      Begin VB.CheckBox Check4 
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
         Left            =   1440
         TabIndex        =   53
         Top             =   2190
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   1215
         Left            =   480
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   50
         Top             =   5760
         Width           =   6045
      End
      Begin VB.CheckBox Check3 
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
         Height          =   255
         Left            =   4920
         TabIndex        =   46
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Frame Frame5 
         Caption         =   "Asignar Contratos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   -74280
         TabIndex        =   35
         Top             =   1440
         Width           =   6135
         Begin FPSpread.vaSpread vaSpread4 
            Height          =   4080
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Width           =   5625
            _Version        =   393216
            _ExtentX        =   9922
            _ExtentY        =   7197
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
            SpreadDesigner  =   "M_Usuari.frx":0070
         End
         Begin VB.Frame Frame16 
            Height          =   435
            Left            =   1650
            TabIndex        =   38
            Top             =   4320
            Width           =   4005
            Begin VB.TextBox TextCai1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   3
               Left            =   45
               TabIndex        =   39
               Top             =   135
               Width           =   3900
            End
         End
         Begin VB.Frame Frame13 
            Height          =   435
            Left            =   720
            TabIndex        =   36
            Top             =   4320
            Width           =   915
            Begin VB.TextBox TextCai1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   2
               Left            =   45
               TabIndex        =   37
               Top             =   135
               Width           =   810
            End
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74760
         TabIndex        =   31
         Top             =   600
         Width           =   7095
         Begin VB.CheckBox Check2 
            Caption         =   "Ver Precios"
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
            Left            =   3075
            TabIndex        =   45
            Top             =   420
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Pedidos Express"
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
            Left            =   5040
            TabIndex        =   32
            Top             =   420
            Visible         =   0   'False
            Width           =   1815
         End
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   0
            Left            =   1800
            TabIndex        =   33
            Top             =   360
            Visible         =   0   'False
            Width           =   645
            _Version        =   196608
            _ExtentX        =   1138
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Días Tope S/R"
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
            TabIndex        =   34
            Top             =   420
            Visible         =   0   'False
            Width           =   1320
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
         Height          =   5895
         Left            =   -74760
         TabIndex        =   29
         Top             =   480
         Width           =   7215
         Begin VB.Frame Frame7 
            Height          =   435
            Left            =   1080
            TabIndex        =   43
            Top             =   5400
            Width           =   915
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   2
               Left            =   45
               TabIndex        =   44
               Top             =   135
               Width           =   810
            End
         End
         Begin VB.Frame Frame6 
            Height          =   435
            Left            =   2010
            TabIndex        =   41
            Top             =   5400
            Width           =   4845
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   3
               Left            =   45
               TabIndex        =   42
               Top             =   135
               Width           =   4740
            End
         End
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   4935
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   6975
            _Version        =   393216
            _ExtentX        =   12303
            _ExtentY        =   8705
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
            MaxCols         =   3
            SpreadDesigner  =   "M_Usuari.frx":19B6
         End
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         ItemData        =   "M_Usuari.frx":3259
         Left            =   2760
         List            =   "M_Usuari.frx":325B
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1260
         Width           =   1800
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   6
         Left            =   480
         TabIndex        =   8
         Top             =   4305
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
         ItemData        =   "M_Usuari.frx":325D
         Left            =   480
         List            =   "M_Usuari.frx":325F
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4980
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   -74760
         TabIndex        =   17
         Top             =   1740
         Width           =   8025
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   4335
            Left            =   240
            TabIndex        =   1
            Top             =   240
            Width           =   7665
            _Version        =   393216
            _ExtentX        =   13520
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
            MaxRows         =   20
            OperationMode   =   3
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "M_Usuari.frx":3261
            VisibleCols     =   2
            VisibleRows     =   15
            ScrollBarTrack  =   1
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Bloqueado"
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   48
            Top             =   4800
            Width           =   765
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00D9D9FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   240
            Top             =   4830
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Height          =   195
            Index           =   0
            Left            =   2745
            TabIndex        =   47
            Top             =   4800
            Width           =   45
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H80000018&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   2385
            Top             =   4830
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -74520
         TabIndex        =   12
         Top             =   540
         Width           =   6615
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "M_Usuari.frx":3733
            Left            =   1680
            List            =   "M_Usuari.frx":373D
            Style           =   2  'Dropdown List
            TabIndex        =   13
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   675
            Width           =   1140
         End
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   3
         Top             =   1860
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
         TabIndex        =   4
         Top             =   2460
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
         TabIndex        =   6
         Top             =   3060
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
         TabIndex        =   5
         Top             =   3060
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
         TabIndex        =   2
         Top             =   1260
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
         TabIndex        =   7
         Top             =   3660
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
         Index           =   7
         Left            =   3555
         TabIndex        =   52
         Top             =   2460
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
         Caption         =   "Rut (sin punto y guión)"
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
         Left            =   3555
         TabIndex        =   51
         Top             =   2220
         Width           =   1935
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
         TabIndex        =   49
         Top             =   5520
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Usuario"
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
         Left            =   2760
         TabIndex        =   27
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000010&
         ForeColor       =   &H80000011&
         Height          =   300
         Left            =   525
         TabIndex        =   26
         Top             =   5055
         Visible         =   0   'False
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
         TabIndex        =   25
         Top             =   1020
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
         TabIndex        =   24
         Top             =   3420
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
         TabIndex        =   23
         Top             =   2820
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
         TabIndex        =   22
         Top             =   2820
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
         TabIndex        =   21
         Top             =   2220
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
         TabIndex        =   20
         Top             =   1620
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
         TabIndex        =   19
         Top             =   4740
         Visible         =   0   'False
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
         TabIndex        =   18
         Top             =   4020
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
Dim ibusca As Long, i As Long
Dim itab As Integer
Dim modo As String, codigo As String
Dim Est As Boolean, estmod As Boolean
Dim Estgri As Boolean

Private Sub Check1_Click()

On Error GoTo Man_Error

If estmod Then Exit Sub

itab = 1
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
itab = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Check2_Click()

On Error GoTo Man_Error

If estmod Then Exit Sub
itab = 1
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
itab = 0


Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

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
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Check4_Click()

If Check4.Value = 1 Then

   fpText(2).PasswordChar = ""
   
Else

   fpText(2).PasswordChar = "*"

End If

End Sub

Private Sub Combo2_Click(Index As Integer)

On Error GoTo Man_Error

If estmod Then Exit Sub

itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
itab = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Combo3_Click(Index As Integer)

On Error GoTo Man_Error

If estmod Then Exit Sub

itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
itab = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
Me.Height = 8025
Me.Width = 8820
MsgTitulo = "Maestro de Usuarios"
fg_centra Me

Combo1.ListIndex = 1
Estgri = True
estmod = True
SSTab1.Tab = 0
modo = ""

Combo3(1).Clear
Combo3(1).AddItem "Real" & Space(150) & "(1)"
Combo3(1).AddItem "Propuesta" & Space(150) & "(2)"
Combo3(1).AddItem "Ambos" & Space(150) & "(3)"

Gl_Mo_Botones Me, 21
Gl_Ac_Botones Me, 1, 1, modo

MoverDatosGrilla
MoverDatos
MoverDatosPerfiles

'AbrirBaseWebPed
'If vg_estopen And Trim(vg_SqlBaseW) <> "" Then
   MoverDatosPerfiles
   MoverDatosClientes
   SSTab1.TabVisible(2) = True
   SSTab1.TabVisible(3) = True
'Else
'   SSTab1.TabVisible(3) = False
'End If
Estgri = False
estmod = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Resize()

On Error GoTo Man_Error

If Me.WindowState <> 1 Then SSTab1.Move 0, Toolbar1.Height, ScaleWidth, ScaleHeight - Toolbar1.Height

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

If estmod Then Exit Sub
itab = 1
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
itab = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_Change(Index As Integer)

On Error GoTo Man_Error

If estmod Then Exit Sub
itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
itab = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_LostFocus(Index As Integer)

On Error GoTo Man_Error

Select Case Index

    Case 0
    
        fpText(Index).text = UCase(fpText(Index).text)
        If Trim(fpText(0).text) = "" Then Exit Sub
        
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpTnombre_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If LimpiaDato(Trim(FptNombre.text)) & Chr(KeyAscii) = "" Then Exit Sub

If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio_V02 '%" & UCase(LimpiaDato(FptNombre.text)) & "%'")
    
    If RS.EOF Or RS!nReg = 0 Then
       
       RS.Close
       Set RS = Nothing
       ibusca = 0
       vaSpread1.MaxRows = 0
       SSTab1.TabEnabled(1) = False
       SSTab1.TabEnabled(2) = False
       SSTab1.TabEnabled(3) = False
       modo = "NE"
       Gl_Ac_Botones Me, 1, 2, modo
       Exit Sub
    
    End If
    If ibusca <> RS!nReg Then
    
       ibusca = RS!nReg
       vaSpread1.MaxRows = RS!nReg
    
    End If
    RS.Close
    Set RS = Nothing
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio_V03 '%" & UCase(LimpiaDato(FptNombre.text)) & "%'")

ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio_V04 '%" & UCase(LimpiaDato(FptNombre.text)) & "%'")
    If RS.EOF Or RS!nReg = 0 Then RS.Close: Set RS = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: SSTab1.TabEnabled(1) = False: SSTab1.TabEnabled(2) = False: SSTab1.TabEnabled(3) = False: modo = "NE": Gl_Ac_Botones Me, 1, 2, modo: Exit Sub
    If ibusca <> RS!nReg Then ibusca = RS!nReg: vaSpread1.MaxRows = RS!nReg
    RS.Close: Set RS = Nothing
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio_V05 '%" & UCase(LimpiaDato(FptNombre.text)) & "%'")

End If

i = 1
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = Shape1(0).FillColor

If Not RS.EOF Then
    
    Do While Not RS.EOF
        
        vaSpread1.Row = i
        i = i + 1
        
        vaSpread1.Col = 1
        vaSpread1.text = RS!usu_codigo
        
        vaSpread1.Col = -1
        vaSpread1.BackColor = IIf(IsNull(RS!usu_activo) Or RS!usu_activo = 0, Shape1(1).FillColor, Shape1(0).FillColor)
        
        vaSpread1.Col = 2
        vaSpread1.TypeHAlign = 0
        vaSpread1.text = Trim(RS!usu_Nombre)
        
        vaSpread1.Col = 3
        vaSpread1.TypeHAlign = 0
        
        If IsNull(RS!usu_indppr) = True Then
        
           vaSpread1.text = ""
           
        Else
        
           vaSpread1.text = IIf(IsNull(RS!usu_indppr) Or RS!usu_indppr = 0, "", IIf(RS!usu_indppr = 1, "Real", IIf(RS!usu_indppr = 2, "Propuesta", "Ambos")))
           
        End If
        
        RS.MoveNext
    
    Loop
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    modo = ""
    Gl_Ac_Botones Me, 1, 1, modo
    
Else
    
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False

End If
RS.Close
Set RS = Nothing
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

On Error GoTo Man_Error

Select Case SSTab1.Tab

Case 0

   Check4.Value = 0
   
Case 1
    
    If vaSpread1.MaxRows > 0 And itab = 0 Then
       
       If modo <> "A" Then
          
          SSTab1.TabEnabled(0) = IIf(modo = "M", False, True)
          SSTab1.Tab = 1
          SSTab1.TabEnabled(1) = True
       
       End If
       
       estmod = True
       Estgri = True
       
       If Toolbar1.Buttons(12).Visible <> True Then
          
          MoverDatos
          MoverDatosPerfiles
          
          If SSTab1.TabVisible(3) = True Then
             
             MoverDatosClientes
          
          End If
       
       End If
       
       Estgri = False
       estmod = False
    
    End If

Case 2
    
    If vaSpread1.MaxRows > 0 And itab = 0 Then
       
       If modo <> "A" Then
          
          SSTab1.TabEnabled(0) = IIf(modo = "M", False, True)
          SSTab1.TabEnabled(1) = True
       
       End If
       
       SSTab1.Tab = 2
       estmod = True
       
       If Toolbar1.Buttons(12).Visible <> True Then
          
          MoverDatos
          MoverDatosPerfiles
          
          If SSTab1.TabVisible(3) = True Then
             
             MoverDatosClientes
          
          End If
       
       End If
       
       estmod = False
    
    End If

Case 3
    
    If vaSpread1.MaxRows > 0 And itab = 0 Then
       
       If modo <> "A" Then
          
          SSTab1.TabEnabled(0) = IIf(modo = "M", False, True)
          SSTab1.TabEnabled(1) = True
          SSTab1.TabEnabled(2) = True
          SSTab1.TabEnabled(3) = True
       
       End If
       
       SSTab1.Tab = 3
       estmod = True
       
       If Toolbar1.Buttons(12).Visible <> True Then
          
          MoverDatos
          MoverDatosPerfiles
          MoverDatosClientes
       
       End If
       
       estmod = False
    
    End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Text1_Change(Index As Integer)

On Error GoTo Man_Error

Dim i   As Long
Dim nom As String

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
    vaSpread2.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread4.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread2.Sort -1, -1, vaSpread2.maxcols, vaSpread2.MaxRows, SortByRow
    
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

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub text2_Change()

On Error GoTo Man_Error

If estmod Then Exit Sub
itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
itab = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub TextCai1_Change(Index As Integer)

On Error GoTo Man_Error

Dim i   As Long
Dim nom As String

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
    vaSpread4.Sort -1, -1, vaSpread4.maxcols, vaSpread4.MaxRows, SortByRow
    
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

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Select Case Button.Index

    Case 1
        
       Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Agregar"), CStr(Me.HelpContextID), "", "", "")
        
        modo = "A"
        Gl_Ac_Botones Me, 1, 0, modo
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = True
        itab = 1
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True
        estmod = True
        
        Check3.Value = 1
        Check3.Enabled = False
        Check4.Value = 0

        For i = 0 To 7
            
            If i < 8 Then
               
               fpText(i).Enabled = True
               fpText(i).text = ""
               
            End If
        
        Next i
        
        For i = 1 To vaSpread2.MaxRows
            
            vaSpread2.Row = i
            vaSpread2.Col = 1
            vaSpread2.text = ""
        
        Next i
        
        For i = 1 To vaSpread4.MaxRows
            
            vaSpread4.Row = i
            vaSpread4.Col = 1
            vaSpread4.text = ""
        
        Next i
        estmod = False
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
        estmod = True
        MoverDatos
        estmod = False: itab = 0
    
    Case 5
        
        Borra_Datos
    
    Case 7
        
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Listar_Lista"), CStr(Me.HelpContextID), "", "", "")
        
        modo = ""
        SSTab1.Tab = 0
        MoverDatosGrilla
    
    Case 10
        
        If MsgBox("Cancela registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Cancelar"), CStr(Me.HelpContextID), "", "", "")
        
        Select Case SSTab1.Tab
        
        Case 1, 2, 3
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient

            SSTab1.TabEnabled(0) = True
            Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio_V02 '%" & UCase(LimpiaDato("")) & "%'")
            If RS.EOF Or RS!nReg = 0 Then
            
               RS.Close
               Set RS = Nothing
               SSTab1.TabEnabled(1) = False
               SSTab1.TabEnabled(2) = False
               SSTab1.TabEnabled(3) = False
               modo = "NE"
               SSTab1.Tab = 0
               Gl_Ac_Botones Me, 1, 2, modo
               Exit Sub
               
            End If
            
            RS.Close
            Set RS = Nothing
            
            SSTab1.TabEnabled(1) = IIf(vaSpread1.MaxRows > 0, True, False)
            SSTab1.TabEnabled(2) = IIf(vaSpread1.MaxRows > 0, True, False)
            SSTab1.TabEnabled(3) = IIf(vaSpread1.MaxRows > 0, True, False)
            MoverDatosPerfiles
            If SSTab1.TabVisible(3) = True Then MoverDatosClientes
        
        End Select
        SSTab1.Tab = 0
        modo = ""
        Gl_Ac_Botones Me, 1, 1, modo
    
    Case 12
        
        Actualiza_Datos
    
    Case 15
        
        If vaSpread1.MaxRows < 1 Then
           
           MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, , MsgTitulo
           Exit Sub
           
        End If
        
    Case 18
        
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), CStr(Me.HelpContextID), "", "", "")

        Me.Hide
        Unload Me
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub MoverDatosGrilla()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

fg_carga ""
Dim X As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 250
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)

vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = Shape1(0).FillColor
vaSpread1.MaxRows = 0

itab = 0
Check4.Value = 0
    
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio")

If Not RS.EOF Then
    
    Do While Not RS.EOF
        
        DoEvents
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
                
        vaSpread1.Col = -1
        vaSpread1.BackColor = IIf(IsNull(RS!usu_activo) Or RS!usu_activo = 0, Shape1(1).FillColor, Shape1(0).FillColor)

        
        vaSpread1.Col = 1
        vaSpread1.text = RS!usu_codigo
        
        vaSpread1.Col = 2
        vaSpread1.TypeHAlign = 0
        vaSpread1.text = Trim(RS!usu_Nombre)
        
        vaSpread1.Col = 3
        vaSpread1.text = IIf(IsNull(RS!usu_indppr) Or RS!usu_indppr = 0, "", IIf(RS!usu_indppr = 1, "Real", IIf(RS!usu_indppr = 2, "Propuesta", "Ambos")))
              
        RS.MoveNext
    
    Loop
    
    Gl_Ac_Botones Me, 1, 1, modo
    SSTab1.TabEnabled(1) = True

Else
    
    SSTab1.Tab = 0
    SSTab1.TabEnabled(1) = False
    modo = "NE"
    Gl_Ac_Botones Me, 1, 2, modo

End If

RS.Close
Set RS = Nothing

If vaSpread1.MaxRows > 0 Then
   
   vaSpread1.Row = 1
   vaSpread1.Col = 1
   codigo = ""
   codigo = Val(vaSpread1.text)
   vaSpread1.SetActiveCell 1, 1

End If

Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
FptNombre.text = ""
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Usuario"

End Sub

Private Sub MoverDatos()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

fg_carga ""
estmod = True

If modo = "A" Then
   
   codigo = ""
   fg_descarga
   Exit Sub

Else
   
   fpText(0).Enabled = False
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = 1
   codigo = vaSpread1.text

End If

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio_V01 '" & codigo & "'")
If Not RS.EOF Then

    DoEvents
    
    fpText(0).text = ""
    fpText(0).text = Trim(RS!usu_codigo)
    
    fpText(1).text = ""
    fpText(1).text = Trim(RS!usu_Nombre)
    
    fpText(2).text = ""
    fpText(2).text = fg_Desencripta(Trim(RS!usu_password))
    
    fpText(3).text = ""
    fpText(3).text = Trim(IIf(IsNull(RS!usu_oficina), "", RS!usu_oficina))
    
    fpText(4).text = ""
    fpText(4).text = Trim(IIf(IsNull(RS!usu_depart), "", RS!usu_depart))
    
    fpText(5).text = ""
    fpText(5).text = Trim(IIf(IsNull(RS!usu_telefono), "", RS!usu_telefono))
    
    fpText(6).text = ""
    fpText(6).text = Trim(IIf(IsNull(RS!usu_email), "", RS!usu_email))
    
    fpText(7).text = ""
    fpText(7).text = Trim(IIf(IsNull(RS!rut), "", RS!rut))
    
    If IsNull(RS!usu_indppr) Or RS!usu_indppr = 0 Then
       
       Combo3(1).ListIndex = -1
    
    Else
       
       Combo3(1).ListIndex = fg_buscacbo(Combo3, 1, 1, fg_pone_cero(Str(RS!usu_indppr), 1))
       
    End If

    Text2.text = ""
    Text2.text = Trim(IIf(IsNull(RS!Ticket), "", RS!Ticket))
    
    Check3.Value = IIf(IsNull(RS!usu_activo) Or RS!usu_activo = 0, 0, 1)
    Check3.Caption = IIf(IsNull(RS!usu_activo) Or RS!usu_activo = 0, "Bloqueado", "Desbloqueado")
    Check3.Enabled = True

End If

RS.Close
Set RS = Nothing

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub MoverDatosPerfiles()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

fg_carga ""
estmod = True
Dim X As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread2.TextTip = 2
vaSpread2.TextTipDelay = 250
X = vaSpread2.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
Frame3.Caption = ""
Frame4.Caption = ""

If modo = "A" Then
   
   codigo = ""
   Frame3.Caption = Trim(fpText(0).text) & " - " & Trim(fpText(1).text)
   Frame4.Caption = Trim(fpText(0).text) & " - " & Trim(fpText(1).text)
   fg_descarga
   Exit Sub

Else
   
   fpText(0).Enabled = False
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = 1
   codigo = vaSpread1.text

End If
vaSpread2.MaxRows = 0


If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio_V01 '" & codigo & "'")
If Not RS.EOF Then
   
   DoEvents
   Frame3.Caption = Trim(RS!usu_codigo) & " - " & Trim(RS!usu_Nombre)
   Frame4.Caption = Trim(RS!usu_codigo) & " - " & Trim(RS!usu_Nombre)

End If
RS.Close
Set RS = Nothing

Estgri = True


If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_usuarioperfiles '" & codigo & "'")
Do While Not RS.EOF
   
   DoEvents
   vaSpread2.MaxRows = vaSpread2.MaxRows + 1
   vaSpread2.Row = vaSpread2.MaxRows
   vaSpread2.Col = 1
   vaSpread2.text = IIf(IsNull(RS!usp_codper), "0", "1")
   
   vaSpread2.Col = 2
   vaSpread2.text = RS!per_codigo
   
   vaSpread2.Col = 3
   vaSpread2.text = IIf(IsNull(RS!per_nombre), "", Trim(RS!per_nombre))
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing

Estgri = False
estmod = False
fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub MoverDatosClientes()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

fg_carga ""
estmod = True
Frame3.Caption = ""
Frame4.Caption = ""

If modo = "A" Then
   
   codigo = ""
   Frame3.Caption = Trim(fpText(0).text) & " - " & Trim(fpText(1).text)
   Frame4.Caption = Trim(fpText(0).text) & " - " & Trim(fpText(1).text)

Else
   
   fpText(0).Enabled = False
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = 1
   codigo = vaSpread1.text

End If


If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio_V01 '" & codigo & "'")
If Not RS.EOF Then
   
   DoEvents
   Frame3.Caption = Trim(RS!usu_codigo) & " - " & Trim(RS!usu_Nombre)
   Frame4.Caption = Trim(RS!usu_codigo) & " - " & Trim(RS!usu_Nombre)
   fpLongInteger1(0).Value = IIf(IsNull(RS!usu_diatop), 0, RS!usu_diatop)
   Check1.Value = IIf(IsNull(RS!usu_pedexp) Or RS!usu_pedexp = "0", 0, 1)
   Check2.Value = IIf(IsNull(RS!usu_vispre) Or RS!usu_vispre = "0", 0, 1)

End If
RS.Close
Set RS = Nothing

Dim X As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread4.TextTip = 2
vaSpread4.TextTipDelay = 250
X = vaSpread4.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
Estgri = True
vaSpread4.Visible = False
vaSpread4.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_UsuarioCeco 1, '" & codigo & "', ''")
Do While Not RS.EOF
   
   vaSpread4.MaxRows = vaSpread4.MaxRows + 1
   vaSpread4.Row = vaSpread4.MaxRows
   
   vaSpread4.Col = 1
   vaSpread4.text = IIf(IsNull(RS!Ceco) Or Trim(RS!Ceco) = "", "0", "1")
   
   vaSpread4.Col = 2
   vaSpread4.text = IIf(IsNull(RS!Cli_codigo), "", RS!Cli_codigo)
   
   vaSpread4.Col = 3
   vaSpread4.text = Trim(IIf(IsNull(RS!Cli_nombre), "", RS!Cli_nombre))
   
   vaSpread4.Col = 4
   vaSpread4.text = Trim(IIf(IsNull(RS!Ceco) Or Trim(RS!Ceco) = "", "", RS!Ceco))
   
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing

vaSpread4.Visible = True
Estgri = False
estmod = False
fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Borra_Datos()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If vaSpread1.MaxRows < 1 Then Exit Sub

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = vaSpread1.text

If MsgBox("Elimina registro... (" & codigo & ")", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub

B_CelEdi.Caption = "Registrar número ticket"
B_CelEdi.Label1.Caption = "Número Ticket"
G_Proc.txt = ""
B_CelEdi.Show 1

If Trim(G_Proc.txt) = "" Then

    MsgBox "Debe regitrar numero ticket..", vbCritical + vbOKOnly, MsgTitulo
    Exit Sub

End If

If Not IsNumeric(G_Proc.txt) Then

    MsgBox "Ticket debe ser númerico.. (" & G_Proc.txt & ")", vbCritical + vbOKOnly, MsgTitulo
    Exit Sub

End If

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminar"), CStr(Me.HelpContextID), "", "", codigo & ";" & fpText(1).text)
    
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Del_UsuarioPerfilCliente '" & codigo & "', '" & G_Proc.txt & "'")

If Not RS.EOF Then
   
   If RS(0) > 0 Then

      'registrar Log sistema error Eliminacion
      Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, RS(0) & RS(1), "", codigo & ";" & fpText(1).text)
                         
      MsgBox "Registro finalizo con error " & RS(0), vbInformation + vbOKOnly, MsgTitulo
      
   Else
   
      'registrar Log sistema Eliminar
      Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminado"), Me.HelpContextID, "", "", codigo & ";" & fpText(1).text)
      
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

If vaSpread1.MaxRows < 1 Then
    
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    SSTab1.Tab = 0
    modo = "NE"

Else
    
    modo = ""
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    SSTab1.Tab = 0

End If
Gl_Ac_Botones Me, 1, 1, modo

Exit Sub
Man_Error:
If Err = -2147467259 Then
    
    MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
    Exit Sub

End If

If Err = 3034 Then Exit Sub

fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Actualiza_Datos()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim i               As Long
Dim codper          As Long
Dim Bloqueo         As String
Dim MyBufferCliente As String

If modo = "A" Then
    
    If Trim(fpText(0).text) = "" Or Trim(fpText(1).text) = "" Or Trim(fpText(2).text) = "" Then
    
       MsgBox "Faltan datos importantes para identificar el Usuario...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
       
    End If
    
    If Trim(fpText(7).text) = "" Then
    
       MsgBox "Debe ingresar Rut del usuario...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
       
    End If
    
    If Trim(Text2.text) = "" Then
    
       MsgBox "Debe ingresar su ticket...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
       
    End If
    
    If Val(fg_codigocbo(Combo3, 1, 1, "")) = 0 Then
       
       MsgBox "Debe seleccionar tipo usuario...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
       
    End If
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio_V01 '" & LimpiaDato(Trim(fpText(0).text)) & "'")
    If Not RS.EOF Then
       
       RS.Close
       Set RS = Nothing
       MsgBox "Usuario existe", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
       
    End If
    
    RS.Close
    Set RS = Nothing
    
    'Validar password
    If Not fg_ValidaPassword(LimpiaDato(Trim(fpText(0).text)), LimpiaDato(Trim(fpText(2).text)), MsgTitulo) Then
       
       Exit Sub

    End If
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Ins_Usuario_V01 '" & LimpiaDato(Trim(fpText(0).text)) & "','" & LimpiaDato(Trim(fpText(1).text)) & "', '" & fg_Encripta(LimpiaDato(Trim(fpText(2).text))) & "', '" & LimpiaDato(Trim(fpText(3).text)) & "', '" & LimpiaDato(Trim(fpText(4).text)) & "', '" & LimpiaDato(Trim(fpText(5).text)) & "', '" & LimpiaDato(Trim(fpText(6).text)) & "', null, '" & fg_codigocbo(Combo3, 1, 1, "") & "', '" & Check3.Value & "', '" & LimpiaDato(Trim(Text2.text)) & "','" & LimpiaDato(Trim(fpText(7).text)) & "'")
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Agregado"), CStr(Me.HelpContextID), RS(0) & RS(1), "", fpText(0).text & ";" & fpText(1).text)
          RS.Close
          Set RS = Nothing
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
       
       End If
    
    End If
    RS.Close
    Set RS = Nothing
    
    '-------> Actualizar usuario
    If vg_estopen And Trim(vg_SqlBaseW) <> "" Then
       
       vg_db.Execute "UPDATE a_usuarios SET usu_diatop = " & fpLongInteger1(0).Value & ", usu_pedexp = '" & IIf(Check1.Value = 1, 1, 0) & "', usu_vispre = '" & IIf(Check2.Value = 1, 1, 0) & "' WHERE usu_codigo='" & LimpiaDato(Trim(fpText(0).text)) & "'"
    
    End If
    
    Let MyBufferCliente = ""
    Let MyBufferCliente = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferCliente = MyBufferCliente & "<UsuarioPerfil>"
    
    For i = 1 To vaSpread2.MaxRows
        
        vaSpread2.Row = i
        vaSpread2.Col = 1
        
        If vaSpread2.text = "1" Then
           
           vaSpread2.Col = 2
           codper = vaSpread2.text
           
           MyBufferCliente = MyBufferCliente & " <UsuPer"
           MyBufferCliente = MyBufferCliente & " codper = " & Chr(34) & codper & Chr(34)
           Let MyBufferCliente = MyBufferCliente & "/>"
           
        End If
    
    Next i
    
    Let MyBufferCliente = MyBufferCliente & "</UsuarioPerfil>"
    
    Set RS = vg_db.Execute("sgpadm_InsDel_UsuarioPerfil '" & MyBufferCliente & "', '" & LimpiaDato(Trim(fpText(0).text)) & "'")
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Agregado"), CStr(Me.HelpContextID), RS(0) & RS(1), "", fpText(0).text & ";" & fpText(1).text)
          RS.Close
          Set RS = Nothing
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
       
       End If
    
    End If
    RS.Close
    Set RS = Nothing
    
    '-------> Borrado y Grabado usuario cliente
    Set RS = vg_db.Execute("sgpadm_Del_ClienteUsuarios '" & LimpiaDato(Trim(fpText(0).text)) & "'")
    If Not RS.EOF Then
          
       If RS(0) > 0 Then
             
           Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), CStr(Me.HelpContextID), RS(0) & RS(1), "", fpText(0).text & ";" & fpText(1).text)
           RS.Close
           Set RS = Nothing
           MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
           Exit Sub
             
       End If
    
    End If
       
    RS.Close
    Set RS = Nothing
       
    Let MyBufferCliente = ""
    Let MyBufferCliente = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferCliente = MyBufferCliente & "<Cliente>"
    
    For i = 1 To vaSpread4.MaxRows
           
           vaSpread4.Row = i
           vaSpread4.Col = 1
           
           If vaSpread4.text = "1" Then
              
              vaSpread4.Col = 2
              MyBufferCliente = MyBufferCliente & " <Clientes"
              MyBufferCliente = MyBufferCliente & " CodCliente = " & Chr(34) & Trim(vaSpread4.text) & Chr(34)
              Let MyBufferCliente = MyBufferCliente & "/>"
              
           End If
    
    Next i
    
    Let MyBufferCliente = MyBufferCliente & "</Cliente>"
    
    Set RS = vg_db.Execute("sgpadm_Ins_ClienteUsuarios '" & LimpiaDato(Trim(fpText(0).text)) & "', '" & MyBufferCliente & "'")
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Agregado"), CStr(Me.HelpContextID), RS(0) & RS(1), "", fpText(0).text & ";" & fpText(1).text)
          RS.Close
          Set RS = Nothing
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
       
       End If
    
    End If
    RS.Close
    Set RS = Nothing
       
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    
    vaSpread1.Col = -1
    vaSpread1.BackColor = IIf(Check3.Value = 1, Shape1(0).FillColor, Shape1(1).FillColor)
    
    vaSpread1.Col = 1
    vaSpread1.Value = LimpiaDato(Trim(fpText(0).text))
    
    vaSpread1.Col = 2
    vaSpread1.Value = LimpiaDato(Trim(fpText(1).text))
    vaSpread1.Col = 3
    vaSpread1.Value = IIf(fg_codigocbo(Combo3, 1, 1, "") = "-1", "", IIf(fg_codigocbo(Combo3, 1, 1, "") = "1", "Real", IIf(fg_codigocbo(Combo3, 1, 1, "") = "2", "Propuesta", "Ambos"))) 'LimpiaDato(Trim(Combo3(1).ListIndex))
 
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Agregado"), CStr(Me.HelpContextID), "", "", fpText(0).text & ";" & fpText(1).text)

Else
    
    If Trim(fpText(1).text) = "" Then
       
       MsgBox "Faltan datos importantes para identificar al Usuario...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
       
    End If
    
    If Trim(fpText(7).text) = "" Then
    
       MsgBox "Debe ingresar Rut del usuario...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
       
    End If
    
    If Trim(Text2.text) = "" Then
    
       MsgBox "Debe ingresar su ticket...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
       
    End If
    
    If Val(fg_codigocbo(Combo3, 1, 1, "")) = 0 Then
       
       MsgBox "Debe seleccionar perfil del usuario...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
       
    End If
    
    Dim pswAnt As String
    pswAnt = ""
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio_V01 '" & LimpiaDato(Trim(fpText(0).text)) & "'")
    If Not RS.EOF Then
           
        pswAnt = RS!usu_password
        Bloqueo = RS!usu_activo
    
    Else
    
       MsgBox "Usuario no existe...", vbExclamation + vbOKOnly, MsgTitulo
       RS.Close
       Set RS = Nothing
       Exit Sub
    
    End If
    RS.Close
    Set RS = Nothing
    
    If pswAnt <> fg_Encripta(fpText(2).text) Then
        
        If Not fg_ValidaPassword(Trim(fpText(0).text), Trim(fpText(2).text), MsgTitulo) Then Exit Sub
    
    End If
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Upd_Usuario_V01 '" & LimpiaDato(Trim(fpText(0).text)) & "', '" & LimpiaDato(Trim(fpText(1).text)) & "', '" & fg_Encripta(LimpiaDato(Trim(fpText(2).text))) & "', '" & LimpiaDato(Trim(fpText(3).text)) & "', '" & LimpiaDato(Trim(fpText(4).text)) & "', '" & LimpiaDato(Trim(fpText(5).text)) & "', '" & LimpiaDato(Trim(fpText(6).text)) & "', " & fg_codigocbo(Combo3, 1, 1, "") & ", '" & Check3.Value & "', '" & LimpiaDato(Trim(Text2.text)) & "','" & LimpiaDato(Trim(fpText(7).text)) & "'")
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), CStr(Me.HelpContextID), RS(0) & RS(1), "", fpText(0).text & ";" & fpText(1).text)
          RS.Close
          Set RS = Nothing
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
       
       End If
    
    End If
    RS.Close
    Set RS = Nothing
    
    If pswAnt <> fg_Encripta(fpText(2).text) Then
        
        'INSERTA MODIFICACIÓN DE PASSWORD
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_CambiaPass"), CStr(Me.HelpContextID), fg_Encripta(fpText(2).text), pswAnt, fpText(0).text & ";" & fpText(1).text)
    
    End If
    
    If Bloqueo <> Check3.Value Then
    
       'INSERTA MODIFICACIÓN DE BLOQUEO
       Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto(IIf(Check3.Value = 0, "vg_logsis_Bloquear", "vg_logsis_Desbloquear")), CStr(Me.HelpContextID), "", "", fpText(0).text & ";" & fpText(1).text)
       
    End If
    
    
    '-------> Actualizar usuario
    If vg_estopen And Trim(vg_SqlBaseW) <> "" Then
       
       vg_db.Execute "UPDATE a_usuarios SET usu_diatop = " & fpLongInteger1(0).Value & ", usu_pedexp = '" & IIf(Check1.Value = 1, 1, 0) & "', usu_vispre = '" & IIf(Check2.Value = 1, 1, 0) & "' WHERE usu_codigo = '" & LimpiaDato(Trim(fpText(0).text)) & "'"
    
    End If
    
    Let MyBufferCliente = ""
    Let MyBufferCliente = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferCliente = MyBufferCliente & "<UsuarioPerfil>"
    
    For i = 1 To vaSpread2.MaxRows
        
        vaSpread2.Row = i
        vaSpread2.Col = 1
        
        If vaSpread2.text = "1" Then
           
           vaSpread2.Col = 2
           codper = vaSpread2.text
           
           MyBufferCliente = MyBufferCliente & " <UsuPer"
           MyBufferCliente = MyBufferCliente & " codper = " & Chr(34) & codper & Chr(34)
           Let MyBufferCliente = MyBufferCliente & "/>"
           
        End If
    
    Next i
    
    Let MyBufferCliente = MyBufferCliente & "</UsuarioPerfil>"
    
    Set RS = vg_db.Execute("sgpadm_InsDel_UsuarioPerfil '" & MyBufferCliente & "', '" & LimpiaDato(Trim(fpText(0).text)) & "'")
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), CStr(Me.HelpContextID), RS(0) & RS(1), "", fpText(0).text & ";" & fpText(1).text)
          RS.Close
          Set RS = Nothing
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
       
       End If
    
    End If
    RS.Close
    Set RS = Nothing
    
    '-------> Borrado y Grabado usuario cliente
    Set RS = vg_db.Execute("sgpadm_Del_ClienteUsuarios '" & LimpiaDato(Trim(fpText(0).text)) & "'")
    If Not RS.EOF Then
          
       If RS(0) > 0 Then
             
           Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), CStr(Me.HelpContextID), RS(0) & RS(1), "", fpText(0).text & ";" & fpText(1).text)
           RS.Close
           Set RS = Nothing
           MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
           Exit Sub
             
       End If
    
    End If
       
    RS.Close
    Set RS = Nothing
       
    Let MyBufferCliente = ""
    Let MyBufferCliente = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferCliente = MyBufferCliente & "<Cliente>"
    
    For i = 1 To vaSpread4.MaxRows
           
           vaSpread4.Row = i
           vaSpread4.Col = 1
           
           If vaSpread4.text = "1" Then
              
              vaSpread4.Col = 2
              MyBufferCliente = MyBufferCliente & " <Clientes"
              MyBufferCliente = MyBufferCliente & " CodCliente = " & Chr(34) & Trim(vaSpread4.text) & Chr(34)
              Let MyBufferCliente = MyBufferCliente & "/>"
              
           End If
    
    Next i
    
    Let MyBufferCliente = MyBufferCliente & "</Cliente>"
    
    Set RS = vg_db.Execute("sgpadm_Ins_ClienteUsuarios '" & LimpiaDato(Trim(fpText(0).text)) & "', '" & MyBufferCliente & "'")
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), CStr(Me.HelpContextID), RS(0) & RS(1), "", fpText(0).text & ";" & fpText(1).text)
          RS.Close
          Set RS = Nothing
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
       
       End If
    
    End If
    RS.Close
    Set RS = Nothing

    vaSpread1.Col = -1
    vaSpread1.BackColor = IIf(Check3.Value = 1, Shape1(0).FillColor, Shape1(1).FillColor)
    
    vaSpread1.Col = 1
    vaSpread1.Value = LimpiaDato(Trim(fpText(0).text))
    
    vaSpread1.Col = 2
    vaSpread1.Value = LimpiaDato(Trim(fpText(1).text))
    
    vaSpread1.Col = 3
    vaSpread1.Value = IIf(fg_codigocbo(Combo3, 1, 1, "") = "-1", "", IIf(fg_codigocbo(Combo3, 1, 1, "") = "1", "Real", IIf(fg_codigocbo(Combo3, 1, 1, "") = "2", "Propuesta", "Ambos")))
    
    vaSpread1.Col = 2
    vaSpread1.Value = LimpiaDato(Trim(fpText(1).text))
    xx = Val(fg_codigocbo(Combo2, 0, 10, ""))
    zz = LimpiaDato(Trim(fpText(0).text))

    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), CStr(Me.HelpContextID), "", "", fpText(0).text & ";" & fpText(1).text)

End If

vg_Indppr = fg_codigocbo(Combo3, 1, 1, "")

If fpText(0).text = UCase(vg_NUsr) Then

   Partida.StatusBar1.Panels(6).text = "Tipo Acceso : " & IIf(vg_Indppr = 1, "Real", IIf(vg_Indppr = 2, "Propuesta", "Ambos"))

End If

vaSpread1.SortKey(1) = 2
vaSpread1.SortKeyOrder(1) = 1
vaSpread1.Sort 1, 1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
   
SSTab1.TabEnabled(0) = True
If vaSpread1.MaxRows < 1 Then
    
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False

Else
    
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    SSTab1.Tab = 0

End If
estmod = True
modo = ""
Gl_Ac_Botones Me, 1, 1, modo
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

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
        
        Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio_V06")
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

              Dim XL As New excel.Application 'Crea el objeto excel
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
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Function encabezado(ByRef RS As ADODB.Recordset, ByRef xlWs As Object)

On Error GoTo Man_Error

Dim fldCount As Long
Dim icol As Long

'-------> Copy field names to the first row of the worksheet
fldCount = RS.Fields.count
For icol = 1 To fldCount
    xlWs.Cells(1, icol).Value = RS.Fields(icol - 1).Name
Next

Exit Function
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo
    
End Function

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = Row
vaSpread1.Col = 1
codigo = vaSpread1.text

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

On Error GoTo Man_Error

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
        TipText = "Tipo Usuario : " & Trim(vaSpread1.text)

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Dim i As Long
vaSpread2.Col = 1

For i = BlockRow To BlockRow2
    
    vaSpread2.Row = i
    vaSpread2.Value = IIf(vaSpread2.Value = "1", "0", "1")
    
    If Toolbar1.Buttons(12).Visible = False Then
       
       SSTab1.Tab = 2
       SSTab1.TabEnabled(0) = False
       SSTab1.TabEnabled(1) = True
       SSTab1.TabEnabled(2) = True
       SSTab1.TabEnabled(3) = True
       If modo = "" Then modo = "M"
       Gl_Ac_Botones Me, 1, 0, modo
    
    End If

Next i

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

If Estgri Then Exit Sub
vaSpread2.Row = Row
Select Case Col

    Case 1
    
        If Toolbar1.Buttons(12).Visible = False Then
           
           SSTab1.TabEnabled(0) = False
           SSTab1.TabEnabled(1) = True
           SSTab1.TabEnabled(2) = True
           SSTab1.TabEnabled(3) = True
           If modo = "" Then modo = "M"
           Gl_Ac_Botones Me, 1, 0, modo
        
        End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread2_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

On Error GoTo Man_Error

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

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread4_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Dim i As Long
vaSpread4.Col = 1
For i = BlockRow To BlockRow2
    
    vaSpread4.Row = i
    vaSpread4.Value = IIf(vaSpread4.Value = "1", "0", "1")
    
    If Toolbar1.Buttons(12).Visible = False Then
       
       SSTab1.Tab = 2
       SSTab1.TabEnabled(0) = False
       SSTab1.TabEnabled(1) = True
       SSTab1.TabEnabled(2) = True
       SSTab1.TabEnabled(3) = True
       If modo = "" Then modo = "M"
       Gl_Ac_Botones Me, 1, 0, modo
    
    End If

Next i

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread4_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

If Estgri Then Exit Sub
vaSpread4.Row = Row
Select Case Col

    Case 1
        
        If Toolbar1.Buttons(12).Visible = False Then
           
           SSTab1.TabEnabled(0) = False
           SSTab1.TabEnabled(1) = True
           SSTab1.TabEnabled(2) = True
           SSTab1.TabEnabled(3) = True
           If modo = "" Then modo = "M"
           Gl_Ac_Botones Me, 1, 0, modo
        
        End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread4_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

On Error GoTo Man_Error

If vaSpread4.MaxRows < 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread4.Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
Select Case Col

    Case 2
        
        vaSpread4.Col = Col
        TipText = "Centro de Costo : " & vaSpread4.text
    
    Case 3
        
        vaSpread4.Col = Col
        TipText = "Descripción : " & Trim(vaSpread4.text)

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub
