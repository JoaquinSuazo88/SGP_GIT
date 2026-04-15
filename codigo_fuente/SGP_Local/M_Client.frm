VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_Client 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Clientes"
   ClientHeight    =   7470
   ClientLeft      =   4380
   ClientTop       =   1785
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   60
      TabIndex        =   16
      Top             =   360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   4
      OLEDropMode     =   1
      TabCaption(0)   =   "Clientes"
      TabPicture(0)   =   "M_Client.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_Client.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(12)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(10)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(9)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(8)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(7)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(3)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(4)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(5)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(6)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(15)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "sombra(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "fpText(6)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "fpText(9)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "fpText(8)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "fpText(5)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "fpText(10)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "fpText(0)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "fpText(3)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "fpText(4)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "fpText(7)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "fpText(2)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "fpText(1)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Check1"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Combo2(0)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Frame3"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Check2"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).ControlCount=   27
      TabCaption(2)   =   "Centro Costo"
      TabPicture(2)   =   "M_Client.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vaSpread2"
      Tab(2).Control(1)=   "lblNOMBRE(0)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Dirección Despacho Cliente Sap"
      TabPicture(3)   =   "M_Client.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "vaSpread3"
      Tab(3).Control(1)=   "lblNOMBRE(1)"
      Tab(3).ControlCount=   2
      Begin VB.CheckBox Check2 
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
         Left            =   2400
         TabIndex        =   44
         Top             =   840
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Día Cierre de Mes Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         TabIndex        =   38
         Top             =   480
         Width           =   3615
         Begin VB.ComboBox Combo3 
            Height          =   315
            Index           =   0
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   360
            Width           =   1695
         End
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   0
            Left            =   2400
            TabIndex        =   40
            Top             =   360
            Visible         =   0   'False
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
            MaxValue        =   "29"
            MinValue        =   "1"
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Día"
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
            Left            =   2400
            TabIndex        =   43
            Top             =   120
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   180
            TabIndex        =   42
            Top             =   460
            Width           =   1690
         End
      End
      Begin FPSpread.vaSpread vaSpread3 
         Height          =   5130
         Left            =   -74760
         TabIndex        =   36
         Top             =   1080
         Width           =   6975
         _Version        =   393216
         _ExtentX        =   12303
         _ExtentY        =   9049
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
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
         SpreadDesigner  =   "M_Client.frx":0070
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1500
         Width           =   6855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cliente SAP"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1185
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   -74880
         TabIndex        =   22
         Top             =   1860
         Width           =   7185
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3855
            Left            =   240
            TabIndex        =   1
            Top             =   360
            Width           =   6765
            _Version        =   393216
            _ExtentX        =   11933
            _ExtentY        =   6800
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
            SpreadDesigner  =   "M_Client.frx":18CE
            VisibleCols     =   2
            VisibleRows     =   15
            ScrollBarTrack  =   1
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -74640
         TabIndex        =   17
         Top             =   660
         Width           =   6615
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "M_Client.frx":1CD0
            Left            =   1680
            List            =   "M_Client.frx":1CDA
            Style           =   2  'Dropdown List
            TabIndex        =   18
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
            Top             =   675
            Width           =   1410
         End
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   2220
         Width           =   6885
         _Version        =   196608
         _ExtentX        =   12144
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
         Left            =   360
         TabIndex        =   6
         Top             =   2820
         Width           =   6885
         _Version        =   196608
         _ExtentX        =   12144
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
         Index           =   7
         Left            =   5085
         TabIndex        =   11
         Top             =   4020
         Width           =   2145
         _Version        =   196608
         _ExtentX        =   3784
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
         Index           =   4
         Left            =   3915
         TabIndex        =   8
         Top             =   3420
         Width           =   3300
         _Version        =   196608
         _ExtentX        =   5821
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
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   3420
         Width           =   3300
         _Version        =   196608
         _ExtentX        =   5821
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
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   780
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
         Index           =   10
         Left            =   360
         TabIndex        =   14
         Top             =   5820
         Width           =   6885
         _Version        =   196608
         _ExtentX        =   12144
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
         Index           =   5
         Left            =   360
         TabIndex        =   9
         Top             =   4020
         Width           =   2265
         _Version        =   196608
         _ExtentX        =   3995
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
         Index           =   8
         Left            =   360
         TabIndex        =   12
         Top             =   4620
         Width           =   6885
         _Version        =   196608
         _ExtentX        =   12144
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
         Index           =   9
         Left            =   360
         TabIndex        =   13
         Top             =   5220
         Width           =   6885
         _Version        =   196608
         _ExtentX        =   12144
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
         Index           =   6
         Left            =   2775
         TabIndex        =   10
         Top             =   4020
         Width           =   2145
         _Version        =   196608
         _ExtentX        =   3784
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
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   5100
         Left            =   -74760
         TabIndex        =   34
         Top             =   1020
         Width           =   6885
         _Version        =   393216
         _ExtentX        =   12144
         _ExtentY        =   8996
         _StockProps     =   64
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
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
         ScrollBars      =   2
         SpreadDesigner  =   "M_Client.frx":1CEE
         VisibleCols     =   2
         VisibleRows     =   15
         ScrollBarTrack  =   1
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   435
         TabIndex        =   41
         Top             =   1605
         Width           =   6825
      End
      Begin VB.Label lblNOMBRE 
         Alignment       =   2  'Center
         Caption         =   "fff"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Index           =   1
         Left            =   -74040
         TabIndex        =   37
         Top             =   720
         Width           =   5280
      End
      Begin VB.Label lblNOMBRE 
         Alignment       =   2  'Center
         Caption         =   "fff"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Index           =   0
         Left            =   -74400
         TabIndex        =   35
         Top             =   660
         Width           =   5880
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
         Left            =   360
         TabIndex        =   33
         Top             =   540
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fono Nº 1"
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
         TabIndex        =   32
         Top             =   3780
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
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
         Left            =   3915
         TabIndex        =   31
         Top             =   3180
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comuna"
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
         Left            =   360
         TabIndex        =   30
         Top             =   3180
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
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
         Left            =   360
         TabIndex        =   29
         Top             =   2580
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
         Left            =   360
         TabIndex        =   28
         Top             =   1980
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fono Nº 2"
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
         Left            =   2775
         TabIndex        =   27
         Top             =   3780
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
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
         Left            =   5085
         TabIndex        =   26
         Top             =   3780
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Persona de Contactos"
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
         Left            =   360
         TabIndex        =   25
         Top             =   4380
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Giro"
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
         TabIndex        =   24
         Top             =   4980
         Width           =   360
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
         Left            =   360
         TabIndex        =   23
         Top             =   5580
         Width           =   465
      End
   End
End
Attribute VB_Name = "M_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim i As Long, iRow As Long
Dim itab As Integer
Dim modo As String, codigo As String, v_rut As String
Dim MsgTitulo As String, est As Boolean

Private Sub Check1_Click()
If est Or Toolbar1.Buttons(12).Visible = True Then Exit Sub
If Check1.Value = 1 Then
'   Combo2(0).AddItem fpText(1).text
   Combo2(0).ListIndex = -1
   Combo2(0).Enabled = False
Else
   Combo2(0).Clear
End If
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
modo = "M": itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
itab = 0
End Sub

Private Sub Check2_Click()
If est Or Toolbar1.Buttons(12).Visible = True Then Exit Sub
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
modo = "M": itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
itab = 0
End Sub

Private Sub Combo2_Click(Index As Integer)
If est Or Toolbar1.Buttons(12).Visible = True Then Exit Sub
Check1.Enabled = False
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
modo = "M": itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
itab = 0
End Sub

Private Sub Combo3_Click(Index As Integer)
If est Then Exit Sub
Select Case Combo3(0).ListIndex
Case 0
    fpLongInteger1(0).Visible = False
    Label1(13).Visible = False
    fpLongInteger1(0).text = ""
Case 1
    fpLongInteger1(0).Visible = True
    Label1(13).Visible = True
    fpLongInteger1(0).text = ""
End Select
If Toolbar1.Buttons(12).Visible = True Then Exit Sub
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
modo = "M": itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
itab = 0
End Sub

Private Sub Combo4_Click(Index As Integer)
If est Or Toolbar1.Buttons(12).Visible = True Then Exit Sub
SSTab1.TabEnabled(0) = False
modo = "M": itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
itab = 0
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 7980
Me.Width = 7725
est = True
MsgTitulo = "Clientes"
fg_centra Me
SSTab1.Tab = 0
modo = ""
Combo1.ListIndex = 1
Toolbar1.ImageList = Partida.IL1
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
SSTab1.TabVisible(3) = False
'-------> Cargar combo día cierre ventas
With Combo3(0)
    .Clear
    .AddItem "Fin de Mes" & Space(150) & "(1)"
    .AddItem "Otros" & Space(150) & "(2)"
    .ListIndex = -1
End With
MoverCliente
est = False
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then SSTab1.Move 0, Toolbar1.Height, ScaleWidth, ScaleHeight - Toolbar1.Height
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
If est Or Toolbar1.Buttons(12).Visible = True Then Exit Sub
SSTab1.TabEnabled(0) = False
modo = "M": itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
itab = 0
End Sub

Private Sub fpText_Change(Index As Integer)
If est Or Toolbar1.Buttons(12).Visible = True Then Exit Sub
SSTab1.TabEnabled(0) = False
modo = "M": itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
itab = 0
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
Dim RS As New ADODB.Recordset
Dim sql1 As String
If LimpiaDato(Trim(fpTnombre.text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
   sql1 = IIf(vg_tipbase = "1", " UCASE(cli_codigo) ", " UPPER(cli_codigo) ")
   RS.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE " & sql1 & " LIKE '%" & UCase(LimpiaDato(fpTnombre.text)) & "%' AND cli_tipo = 1 ORDER BY cli_codigo", vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
   sql1 = IIf(vg_tipbase = "1", " UCASE(cli_nombre) ", " UPPER(cli_nombre) ")
   RS.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE " & sql1 & " LIKE '%" & UCase(LimpiaDato(fpTnombre.text)) & "%' AND cli_tipo = 1 ORDER BY cli_nombre", vg_db, adOpenStatic
End If
i = 1: vaSpread1.MaxRows = RS.RecordCount
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.Row = i
      
      vaSpread1.Col = 1
      vaSpread1.text = RS!cli_codigo

      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = Trim(RS!cli_nombre)
      RS.MoveNext: i = i + 1
   Loop
   vaSpread1.Row = 1: vaSpread1.Col = 1: codigo = vaSpread1.text
   MoverDetCliente codigo
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
   modo = ""
   Gl_Ac_Botones Me, 1, 1, modo
Else
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
End If
RS.Close: Set RS = Nothing
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Dim RS As New ADODB.Recordset

Select Case SSTab1.Tab
Case 0
    For i = 0 To 10
        fpText(i).Enabled = False
    Next i
    vaSpread2.Enabled = False
    vaSpread3.Enabled = False
    Check1.Enabled = False
    Combo2(0).Enabled = False
    Combo1.Enabled = True: fpTnombre.Enabled = True
    MoverDetCliente codigo
Case 1, 2, 3

    est = False
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "select top 1 cli_codigo " & _
            "from b_clientes as a with (nolock) " & _
            "inner join b_preciovta as b  with (nolock) on b.prv_rutcli = a.cli_codigo " & _
            "where cli_activo = '1' " & _
            "and   cli_tipo = 1 " & _
            "and   cli_codigo = '" & codigo & "' " & _
            "and   b.prv_SPRS = '1' " & _
            "group by cli_codigo", vg_db, adOpenStatic
    If Not RS.EOF Then
       
        Exit Sub
       
    End If
    RS.Close
    Set RS = Nothing
    
    For i = 1 To 10
        fpText(i).Enabled = True
    Next i
    vaSpread2.Enabled = True
    vaSpread3.Enabled = True
    Check1.Enabled = True
    If Check1.Enabled = 1 Then Combo2(0).Enabled = True
    Combo1.Enabled = False
    Combo3(0).Enabled = True
    fpTnombre.Enabled = False
    Check2.Enabled = True

End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1 '-------> Agregar
    SSTab1.TabEnabled(IIf(SSTab1.Tab = 0, 1, 0)) = False
    est = True: modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    If SSTab1.Tab = 0 Or SSTab1.Tab = 1 Then
       Combo2(0).ListIndex = -1
       Combo3(0).ListIndex = -1
       Check2.Value = 1
       SSTab1.TabEnabled(0) = False
       SSTab1.TabEnabled(1) = True
       SSTab1.TabEnabled(2) = False
       SSTab1.TabEnabled(3) = False
       SSTab1.Tab = 1
       For i = 0 To 11
           If i < 11 Then fpText(i).Enabled = True: fpText(i).text = ""
       Next i
    ElseIf SSTab1.Tab = 2 Then
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = False
        SSTab1.Tab = 2
        vaSpread2.MaxRows = vaSpread2.MaxRows + 1
        iRow = vaSpread2.MaxRows: vaSpread2.Row = vaSpread2.MaxRows: vaSpread2.Col = 1: vaSpread2.SetActiveCell 1, vaSpread2.MaxRows: vaSpread2.SetFocus
    ElseIf SSTab1.Tab = 3 Then
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(2) = True
        SSTab1.Tab = 3
        vaSpread3.MaxRows = vaSpread3.MaxRows + 1
        iRow = vaSpread3.MaxRows: vaSpread3.Row = vaSpread3.MaxRows: vaSpread3.Col = 1: vaSpread3.SetActiveCell 1, vaSpread3.MaxRows: vaSpread3.SetFocus
   End If
   est = False: itab = 0
Case 3 '-------> Modificar
    If vaSpread1.MaxRows < 1 Then Exit Sub
    itab = 1
    If SSTab1.Tab = 0 Or SSTab1.Tab = 1 Then
       SSTab1.TabEnabled(0) = False
       SSTab1.TabEnabled(2) = False
       SSTab1.TabEnabled(3) = False
       SSTab1.Tab = 1: SSTab1.TabEnabled(1) = True
    ElseIf SSTab1.Tab = 2 Then
       If vaSpread2.MaxRows < 1 Then Exit Sub
       iRow = vaSpread2.ActiveRow
       SSTab1.TabEnabled(0) = False
       SSTab1.TabEnabled(1) = False
       SSTab1.TabEnabled(3) = False
       SSTab1.Tab = 2: SSTab1.TabEnabled(2) = True
    ElseIf SSTab1.Tab = 3 Then
       If vaSpread3.MaxRows < 1 Then Exit Sub
       iRow = vaSpread3.ActiveRow
       SSTab1.TabEnabled(0) = False
       SSTab1.TabEnabled(1) = False
       SSTab1.TabEnabled(2) = False
       SSTab1.Tab = 3: SSTab1.TabEnabled(3) = True
    End If
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    itab = 0
Case 5 '-------> Borrar lista
    If SSTab1.Tab = 0 Or SSTab1.Tab = 1 Then
       Borra_DatoCliente
    ElseIf SSTab1.Tab = 2 Then
       Borra_DatoCliCencos
    ElseIf SSTab1.Tab = 3 Then
       Borra_DatoSucCli
    End If
Case 7 '-------> Actualizar lista
    modo = ""
    If SSTab1.Tab = 0 Or SSTab1.Tab = 1 Then
       SSTab1.Tab = 0
       MoverCliente
    ElseIf SSTab1.Tab = 2 And vaSpread1.MaxRows > 0 Then
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = 1: codigo = Trim(vaSpread1.text)
       MoverCliCenCos codigo
    ElseIf SSTab1.Tab = 3 And vaSpread1.MaxRows > 0 Then
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = 1: codigo = Trim(vaSpread1.text)
       MoverSucCli codigo
    End If
Case 10 '-------> Cancelar
    Cancela_Datos
Case 12 '-------> Grabar datos
    Actualiza_Datos
Case 15 '-------> Imprimir
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, , MsgTitulo: Exit Sub
    If SSTab1.Tab = 0 Or SSTab1.Tab = 1 Then
       I_Clientes
    ElseIf SSTab1.Tab = 2 Then
       I_ClienteCencos
    ElseIf SSTab1.Tab = 3 Then
       I_SucursalCliente
    End If
Case 18 '-------> Salir
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub MoverCliente()
On Error GoTo Man_Error
Dim RS As New ADODB.Recordset
fg_carga ""
vaSpread1.MaxRows = 0
itab = 0
Combo2(0).Clear
RS.Open "SELECT * FROM b_clientes WHERE cli_tipo = 1 AND cli_clisap = '1' ORDER BY cli_nombre", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      Combo2(0).AddItem Trim(RS!cli_nombre) & Space(150) & "(" & fg_pone_espacio(RS!cli_codigo, 10) & ")"
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing

RS.Open "SELECT * FROM b_clientes WHERE cli_tipo = 1 ORDER BY cli_nombre", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
              
      vaSpread1.Col = 1
      vaSpread1.text = RS!cli_codigo

      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = Trim(RS!cli_nombre)
             
      RS.MoveNext
   Loop
   vaSpread1.Row = 1: vaSpread1.Col = 1: codigo = Trim(vaSpread1.text)
   MoverDetCliente codigo
   Gl_Ac_Botones Me, 1, 1, modo
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
   SSTab1.TabEnabled(3) = True
Else
   SSTab1.Tab = 0
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
   SSTab1.TabEnabled(3) = False
   modo = "NE"
   Gl_Ac_Botones Me, 1, 2, modo
End If
RS.Close: Set RS = Nothing
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
fpTnombre.text = ""
fg_descarga
Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"
End Sub

Private Sub MoverDetCliente(codigo As String)
Dim RS1 As New ADODB.Recordset
fg_carga ""
est = True
Combo2(0).Clear
RS1.Open "SELECT * FROM b_clientes WHERE cli_tipo = 1 AND cli_clisap = '1' ORDER BY cli_nombre", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      Combo2(0).AddItem Trim(RS1!cli_nombre) & Space(150) & "(" & fg_pone_espacio(RS1!cli_codigo, 10) & ")"
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing

RS1.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & codigo & "' AND cli_tipo = 1", vg_db, adOpenStatic
If Not RS1.EOF Then
   lblNOMBRE(0).Caption = "(" & fg_PintaRut(RS1!cli_codigo) & ") - " & Trim(RS1!cli_nombre)
   lblNOMBRE(1).Caption = "(" & fg_PintaRut(RS1!cli_codigo) & ") - " & Trim(RS1!cli_nombre)
   fpText(0).text = fg_PintaRut(RS1!cli_codigo)
   fpText(1).text = IIf(IsNull(RS1!cli_nombre), "", Trim(RS1!cli_nombre))
   fpText(2).text = IIf(IsNull(RS1!cli_direccion), "", Trim(RS1!cli_direccion))
   fpText(3).text = IIf(IsNull(RS1!cli_comuna), "", Trim(RS1!cli_comuna))
   fpText(4).text = IIf(IsNull(RS1!cli_ciudad), "", Trim(RS1!cli_ciudad))
   fpText(5).text = IIf(IsNull(RS1!cli_fono1), "", Trim(RS1!cli_fono1))
   fpText(6).text = IIf(IsNull(RS1!cli_fono2), "", Trim(RS1!cli_fono2))
   fpText(7).text = IIf(IsNull(RS1!cli_fax), "", Trim(RS1!cli_fax))
   fpText(8).text = IIf(IsNull(RS1!cli_percon), "", Trim(RS1!cli_percon))
   fpText(9).text = IIf(IsNull(RS1!cli_giro), "", Trim(RS1!cli_giro))
   fpText(10).text = IIf(IsNull(RS1!cli_email), "", Trim(RS1!cli_email))
   Check1.Value = IIf(IsNull(RS1!cli_clisap) Or RS1!cli_clisap = "0", 0, 1)
   
   If Not IsNull(RS1!cli_codcli) Or Trim(RS1!cli_codcli) <> "" Then Combo2(0).ListIndex = fg_buscacbo(Combo2, 0, 10, RS1!cli_codcli)
   
   If RS1!cli_clisap = "1" Then Combo2(0).ListIndex = -1 'Combo2(0).ListIndex = fg_buscacbostring(Combo2, 0, 10, (RS1!cli_codigo)) 'AddItem Trim(RS1!cli_nombre): Combo2.ListIndex = 0
   
   Combo2(0).Enabled = IIf(Check1.Value = 0, True, False)
   If RS1!cli_clisap = "1" Then SSTab1.TabVisible(3) = True Else SSTab1.TabVisible(3) = False
   Combo3(0).ListIndex = -1
   fpLongInteger1(0).Visible = False
   Label1(13).Visible = False
   If Not IsNull(RS1!cli_cievta) Or Trim(RS1!cli_cievta) <> "" Then
      Combo3(0).ListIndex = fg_buscacbo(Combo3, 0, 1, RS1!cli_cievta)
      fpLongInteger1(0).Visible = IIf(RS1!cli_cievta = "1", False, True)
      fpLongInteger1(0).text = IIf(IsNull(RS1!cli_ciedia), 0, RS1!cli_ciedia)
      Label1(13).Visible = IIf(RS1!cli_cievta = "1", False, True)
   End If
   Check2.Value = IIf(IsNull(RS1!cli_activo) Or RS1!cli_activo = "0", 0, 1)
End If
RS1.Close: Set RS1 = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
RS.Open "select top 1 cli_codigo " & _
            "from b_clientes as a with (nolock) " & _
            "inner join b_preciovta as b  with (nolock) on b.prv_rutcli = a.cli_codigo " & _
            "where cli_activo = '1' " & _
            "and   cli_tipo = 1 " & _
            "and   cli_codigo = '" & codigo & "' " & _
            "and   b.prv_SPRS = '1' " & _
            "group by cli_codigo", vg_db, adOpenStatic
If Not RS.EOF Then
       
   fpText(0).Enabled = False
   fpText(1).Enabled = False
   fpText(2).Enabled = False
   fpText(3).Enabled = False
   fpText(4).Enabled = False
   fpText(5).Enabled = False
   fpText(6).Enabled = False
   fpText(7).Enabled = False
   fpText(8).Enabled = False
   fpText(9).Enabled = False
   fpText(10).Enabled = False
   Check1.Enabled = False
   Combo2(0).Enabled = False
   Combo3(0).Enabled = False
   fpLongInteger1(0).Enabled = False
   Label1(13).Enabled = False
   Check2.Enabled = False
       
End If
RS.Close
Set RS = Nothing

est = False
MoverCliCenCos codigo
MoverSucCli codigo
fg_descarga
End Sub

Sub MoverCliCenCos(codigo As String)
Dim RS1 As New ADODB.Recordset
vaSpread2.MaxRows = 0
RS1.Open "SELECT * FROM b_clientecencos WHERE clc_codcli = '" & codigo & "' ORDER BY clc_codigo", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      vaSpread2.Col = 1: vaSpread2.Lock = True
      vaSpread2.text = Trim(RS1!clc_codigo)
      vaSpread2.Col = 2: vaSpread2.Lock = False
      vaSpread2.text = Trim(RS1!clc_nombre)
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
End Sub

Sub MoverSucCli(codigo As String)
Dim RS1 As New ADODB.Recordset
vaSpread3.MaxRows = 0
RS1.Open "SELECT * FROM b_sucursalcliente WHERE scl_codcli = '" & codigo & "' ORDER BY scl_codigo", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      vaSpread3.MaxRows = vaSpread3.MaxRows + 1
      vaSpread3.Row = vaSpread3.MaxRows
      vaSpread3.Col = 1: vaSpread3.Lock = True
      vaSpread3.text = Trim(RS1!scl_codigo)
      vaSpread3.Col = 2: vaSpread3.Lock = False
      vaSpread3.text = Trim(RS1!scl_direccion)
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
End Sub

Private Sub Borra_DatoCliente()

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: codigo = vaSpread1.text

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
RS.Open "select top 1 cli_codigo " & _
            "from b_clientes as a with (nolock) " & _
            "inner join b_preciovta as b  with (nolock) on b.prv_rutcli = a.cli_codigo " & _
            "where cli_activo = '1' " & _
            "and   cli_tipo = 1 " & _
            "and   cli_codigo = '" & codigo & "' " & _
            "and   b.prv_SPRS = '1' " & _
            "group by cli_codigo", vg_db, adOpenStatic
If Not RS.EOF Then
       
   RS.Close
   Set RS = Nothing
   MsgBox "No puede borrar datos de un cliente SPRS", vbCritical + vbOKOnly, MsgTitulo
       
   Exit Sub
    
End If
RS.Close
Set RS = Nothing

If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub

vg_db.BeginTrans
vg_db.Execute "DELETE b_clientes FROM b_clientes WHERE cli_codigo = '" & codigo & "' AND cli_tipo = 1"
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.DeleteRows vaSpread1.Row, 1
vaSpread1.MaxRows = vaSpread1.MaxRows - 1
vaSpread1.Row = vaSpread1.MaxRows
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
If vaSpread1.MaxRows < 1 Then
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
   SSTab1.Tab = 0
   modo = "NE"
Else
   modo = ""
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
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
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Borra_DatoCliCencos()
On Error GoTo Man_Error
Dim cencos As String
If vaSpread2.MaxRows < 1 Then Exit Sub
vaSpread2.Row = vaSpread2.ActiveRow
vaSpread2.Col = 1: cencos = vaSpread2.text
If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
vg_db.BeginTrans
vg_db.Execute "DELETE b_clientecencos FROM b_clientecencos WHERE clc_codcli = '" & codigo & "' AND clc_codigo = '" & cencos & "'"
vaSpread2.Row = vaSpread2.ActiveRow
vaSpread2.DeleteRows vaSpread2.Row, 1
vaSpread2.MaxRows = vaSpread2.MaxRows - 1
vaSpread2.Row = vaSpread2.MaxRows
If vaSpread2.MaxRows < 1 Then
   modo = "NE"
Else
   modo = ""
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
End If
vg_db.CommitTrans
Gl_Ac_Botones Me, 1, 1, modo

Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Borra_DatoSucCli()
On Error GoTo Man_Error
Dim codsuc As String
If vaSpread3.MaxRows < 1 Then Exit Sub
vaSpread3.Row = vaSpread3.ActiveRow
vaSpread3.Col = 1: codsuc = vaSpread3.text
If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
vg_db.BeginTrans
vg_db.Execute "DELETE b_sucursalcliente FROM b_sucursalcliente WHERE scl_codcli = '" & codigo & "' AND scl_codigo = '" & codsuc & "'"
vaSpread3.Row = vaSpread3.ActiveRow
vaSpread3.DeleteRows vaSpread3.Row, 1
vaSpread3.MaxRows = vaSpread3.MaxRows - 1
vaSpread3.Row = vaSpread3.MaxRows
If vaSpread3.MaxRows < 1 Then
   modo = "NE"
Else
   modo = ""
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
   SSTab1.TabEnabled(3) = True
End If
vg_db.CommitTrans
Gl_Ac_Botones Me, 1, 1, modo

Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Cancela_Datos()
Dim sql1 As String
Dim RS As New ADODB.Recordset
If MsgBox("Cancela registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
SSTab1.TabEnabled(0) = True
sql1 = IIf(vg_tipbase = "1", " UCASE(cli_codigo) ", " UPPER(cli_codigo) ")
RS.Open "SELECT COUNT(*) AS nreg FROM b_clientes WHERE " & sql1 & " LIKE '%" & UCase(("")) & "%' AND cli_tipo = 1", vg_db, adOpenStatic
If RS.EOF Or RS!nreg = 0 Then RS.Close: Set RS = Nothing: SSTab1.TabEnabled(1) = False: SSTab1.TabEnabled(2) = False: SSTab1.TabEnabled(3) = False: modo = "NE": SSTab1.Tab = 0: Gl_Ac_Botones Me, 1, 2, modo: Exit Sub
RS.Close: Set RS = Nothing
If vaSpread1.MaxRows > 0 Then
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
   SSTab1.TabEnabled(3) = True
Else
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
   SSTab1.TabEnabled(3) = False
End If
SSTab1.Tab = 0
modo = ""
Gl_Ac_Botones Me, 1, 1, modo
End Sub

Private Sub Actualiza_Datos()
Dim cencos As String, Nombre As String, codsuc As String, codcli As String, cievta As String
Dim RS As New ADODB.Recordset
On Error GoTo Man_Error
v_rut = fg_DespintaRut(fpText(0).text)
If modo = "A" Then
   If SSTab1.Tab = 1 Then
      codcli = "": cievta = "": TipoVales = ""
      If Combo2(0).ListIndex > -1 Then codcli = fg_codigocbo(Combo2, 0, 10, "")
      If Combo3(0).ListIndex > -1 Then cievta = fg_codigocbo(Combo3, 0, 1, "")
      If Trim(fpText(0).text) = "" Or Trim(fpText(1).text) = "" Or Trim(cievta) = "" Or (Trim(cievta) = "2" And Val(fpLongInteger1(0).Value) < 1) Then MsgBox "Faltan datos importantes para identificar el cliente...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      If Not fg_Check_Rut(v_rut) Then MsgBox "El rut no es valido...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & v_rut & "'", vg_db, adOpenStatic
      If Not RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Cliente existe", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      RS.Close: Set RS = Nothing
      vg_db.BeginTrans
      vg_db.Execute "INSERT INTO b_clientes (cli_codigo, cli_nombre, cli_direccion, " & _
                    "cli_comuna, cli_ciudad, cli_fono1, cli_fono2, cli_fax, cli_percon, " & _
                    "cli_giro, cli_email, cli_tipo, cli_codbod, cli_codtis, cli_codseg, cli_codcli, cli_clisap, cli_socsap, cli_cievta, cli_ciedia, cli_activo) VALUES ('" & v_rut & "', " & _
                    "'" & LimpiaDato(Trim(fpText(1).text)) & "', '" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                    "'" & LimpiaDato(Trim(fpText(3).text)) & "', '" & LimpiaDato(Trim(fpText(4).text)) & "', " & _
                    "'" & LimpiaDato(Trim(fpText(5).text)) & "', '" & LimpiaDato(Trim(fpText(6).text)) & "', " & _
                    "'" & LimpiaDato(Trim(fpText(7).text)) & "', '" & LimpiaDato(Trim(fpText(8).text)) & "', " & _
                    "'" & LimpiaDato(Trim(fpText(9).text)) & "', '" & LimpiaDato(Trim(fpText(10).text)) & "', " & _
                    "" & 1 & ", 0, null, null, '" & codcli & "', '" & IIf(Check1.Value = 0, "0", "1") & "', null, '" & cievta & "', " & Val(fpLongInteger1(0).Value) & ", '" & IIf(Check2.Value = 0, "0", "1") & "')"
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 1: vaSpread1.TypeHAlign = 0: vaSpread1.Value = v_rut
      vaSpread1.Col = 2: vaSpread1.TypeHAlign = 0: vaSpread1.Value = LimpiaDato(Trim(fpText(1).text))
      vg_db.CommitTrans
   ElseIf SSTab1.Tab = 2 Then
      vaSpread2.Row = iRow
      vaSpread2.Col = 1
      cencos = LimpiaDato(Trim(vaSpread2.text))
      vaSpread2.Col = 2
      Nombre = LimpiaDato(Trim(vaSpread2.text))
      If cencos = "" Or Nombre = "" Then MsgBox "Faltan datos importantes centro costo...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      RS.Open "SELECT * FROM b_clientecencos WHERE clc_codigo = '" & cencos & "' AND clc_codcli = '" & v_rut & "'", vg_db, adOpenStatic
      If Not RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Centro costo existe en lista", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      RS.Close: Set RS = Nothing
      vg_db.BeginTrans
      vg_db.Execute "INSERT INTO b_clientecencos (clc_codigo, clc_codcli, clc_nombre) VALUES ('" & cencos & "', '" & v_rut & "', '" & Nombre & "')"
      vaSpread2.Row = iRow: vaSpread2.Col = 1: vaSpread2.Lock = True
      vg_db.CommitTrans
   ElseIf SSTab1.Tab = 3 Then
      vaSpread3.Row = iRow
      vaSpread3.Col = 1
      codsuc = LimpiaDato(Trim(vaSpread3.text))
      vaSpread3.Col = 2
      Nombre = LimpiaDato(Trim(vaSpread3.text))
      If codsuc = "" Or Nombre = "" Then MsgBox "Faltan datos importantes sucursal SAP...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      RS.Open "SELECT * FROM b_sucursalcliente WHERE scl_codigo = '" & cencos & "' AND scl_codcli = '" & v_rut & "'", vg_db, adOpenStatic
      If Not RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Sucursal SAP existe en lista", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      RS.Close: Set RS = Nothing
      vg_db.BeginTrans
      vg_db.Execute "INSERT INTO b_sucursalcliente (scl_codigo, scl_codcli, scl_direccion) VALUES ('" & codsuc & "', '" & v_rut & "', '" & Nombre & "')"
      vaSpread3.Row = iRow: vaSpread3.Col = 1: vaSpread3.Lock = True
      vg_db.CommitTrans
   End If
Else
   If SSTab1.Tab = 1 Then
      codcli = "": cievta = ""
      Dim clisap As Boolean
      If Combo2(0).ListIndex > -1 Then codcli = Trim(fg_codigocbo(Combo2, 0, 10, ""))
      If Combo3(0).ListIndex > -1 Then cievta = fg_codigocbo(Combo3, 0, 1, "")
      If Trim(fpText(1).text) = "" Or Trim(cievta) = "" Or (Trim(cievta) = "2" And Val(fpLongInteger1(0).Value) < 1) Then MsgBox "Faltan datos importantes para identificar el cliente...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      clisap = IIf(Check1.Value = 0, True, False)
      vg_db.BeginTrans
      If vg_tipbase = "1" Then
         RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & v_rut & "' AND cli_clisap = '1' AND '" & IIf(Check1.Value = 0, True, False) & "'", vg_db, adOpenStatic
      Else
         RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & v_rut & "' AND cli_clisap = '1'", vg_db, adOpenStatic
      End If
      If Not RS.EOF And (Not clisap Or clisap) Then
         vg_db.Execute "UPDATE b_clientes SET cli_codcli = ' ' WHERE cli_codcli = '" & v_rut & "' AND cli_tipo=1"
      End If
      RS.Close: Set RS = Nothing
      
      vg_db.Execute "UPDATE b_clientes SET cli_nombre='" & LimpiaDato(Trim(fpText(1).text)) & "', cli_direccion = '" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                    "cli_comuna = '" & LimpiaDato(Trim(fpText(3).text)) & "', cli_ciudad = '" & LimpiaDato(Trim(fpText(4).text)) & "', cli_fono1 = '" & LimpiaDato(Trim(fpText(5).text)) & "', " & _
                    "cli_fono2 = '" & LimpiaDato(Trim(fpText(6).text)) & "', cli_fax = '" & LimpiaDato(Trim(fpText(7).text)) & "', cli_percon = '" & LimpiaDato(Trim(fpText(8).text)) & "', " & _
                    "cli_giro = '" & LimpiaDato(Trim(fpText(9).text)) & "', cli_email = '" & LimpiaDato(Trim(fpText(10).text)) & "', cli_codcli = '" & codcli & "', cli_clisap = '" & IIf(Check1.Value = 0, "0", "1") & "', " & _
                    "cli_cievta = '" & cievta & "', cli_ciedia = " & Val(fpLongInteger1(0).Value) & ", cli_activo = '" & IIf(Check2.Value = 0, "0", "1") & "' WHERE cli_codigo = '" & v_rut & "'"
      vaSpread1.Col = 2
      vaSpread1.Value = LimpiaDato(Trim(fpText(1).text))
      vg_db.CommitTrans
   ElseIf SSTab1.Tab = 2 Then
      vaSpread2.Row = iRow
      vaSpread2.Col = 1
      cencos = LimpiaDato(Trim(vaSpread2.text))
      vaSpread2.Col = 2
      Nombre = LimpiaDato(Trim(vaSpread2.text))
      If cencos = "" Or Nombre = "" Then MsgBox "Faltan datos importantes para identificar centro costo...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      vg_db.BeginTrans
      vg_db.Execute "UPDATE b_clientecencos SET clc_nombre = '" & Nombre & "' WHERE clc_codigo = '" & cencos & "' AND clc_codcli = '" & v_rut & "'"
      vg_db.CommitTrans
   ElseIf SSTab1.Tab = 3 Then
      vaSpread3.Row = iRow
      vaSpread3.Col = 1
      codsuc = LimpiaDato(Trim(vaSpread3.text))
      vaSpread3.Col = 2
      Nombre = LimpiaDato(Trim(vaSpread3.text))
      If codsuc = "" Or Nombre = "" Then MsgBox "Faltan datos importantes sucuarsal SAP...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      vg_db.BeginTrans
      vg_db.Execute "UPDATE b_sucursalcliente SET scl_direccion = '" & Nombre & "' WHERE scl_codigo = '" & codsuc & "' AND scl_codcli = '" & v_rut & "'"
      vg_db.CommitTrans
   End If
End If
If SSTab1.Tab = 1 Then
   vaSpread1.SortKey(1) = 2
   vaSpread1.SortKeyOrder(1) = 1
   vaSpread1.Sort 1, 1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
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
   Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
Else
   If vaSpread1.MaxRows < 1 Then
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(2) = False
      SSTab1.TabEnabled(3) = False
   Else
      SSTab1.TabEnabled(0) = True
      SSTab1.TabEnabled(1) = True
      SSTab1.TabEnabled(2) = True
      SSTab1.TabEnabled(3) = True
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
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = Row: vaSpread1.Col = 1: codigo = Trim(vaSpread1.text)
MoverDetCliente codigo
End Sub

Private Sub vaSpread2_EditChange(ByVal Col As Long, ByVal Row As Long)
If Toolbar1.Buttons(12).Visible = True Then Exit Sub
iRow = Row
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(3) = False
SSTab1.Tab = 2: SSTab1.TabEnabled(2) = True
modo = "M": Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub vaSpread3_EditChange(ByVal Col As Long, ByVal Row As Long)
If Toolbar1.Buttons(12).Visible = True Then Exit Sub
iRow = Row
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.Tab = 3: SSTab1.TabEnabled(3) = True
modo = "M": Gl_Ac_Botones Me, 1, 0, modo
End Sub
