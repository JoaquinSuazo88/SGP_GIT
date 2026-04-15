VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_Provee 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Proveedores"
   ClientHeight    =   8250
   ClientLeft      =   2460
   ClientTop       =   1815
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   60
      TabIndex        =   14
      Top             =   360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   13573
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   4
      OLEDropMode     =   1
      TabCaption(0)   =   "Proveedores"
      TabPicture(0)   =   "M_Provee.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_Provee.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblSOMBRA(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(12)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(10)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(9)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(8)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(7)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(3)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(4)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(5)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(6)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(15)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label1(13)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(14)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label1(16)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lblSOMBRA(10)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lblSOMBRA(0)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lblSOMBRA(1)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label1(17)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Image1(0)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "fpayuda(0)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "fpLongInteger1(0)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "fpText(6)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "fpText(9)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "fpText(8)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "fpText(5)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "fpText(10)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "fpText(0)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "fpText(3)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "fpText(4)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "fpText(7)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "fpText(2)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "fpText(1)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Check1(0)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Combo2(0)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Combo2(1)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Combo2(2)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Check1(1)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Check1(2)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).ControlCount=   39
      Begin VB.CheckBox Check1 
         Caption         =   "Ingreso Documento SGP Local"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   2
         Left            =   5880
         TabIndex        =   47
         Top             =   6240
         Width           =   1410
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Entrega Documento Electrónico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   1
         Left            =   4200
         TabIndex        =   46
         Top             =   6240
         Width           =   1410
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   2
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   6440
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   1
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   5760
         Width           =   2775
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   5760
         Width           =   3375
      End
      Begin VB.CheckBox Check1 
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
         Height          =   300
         Index           =   0
         Left            =   6120
         TabIndex        =   32
         Top             =   960
         Value           =   1  'Checked
         Width           =   930
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   5775
         Left            =   -74880
         TabIndex        =   20
         Top             =   1800
         Width           =   7545
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   5415
            Left            =   240
            TabIndex        =   1
            Top             =   240
            Width           =   7125
            _Version        =   393216
            _ExtentX        =   12568
            _ExtentY        =   9551
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
            SpreadDesigner  =   "M_Provee.frx":0038
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
         TabIndex        =   15
         Top             =   600
         Width           =   6615
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "M_Provee.frx":051D
            Left            =   1680
            List            =   "M_Provee.frx":0527
            Style           =   2  'Dropdown List
            TabIndex        =   16
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
            TabIndex        =   19
            Top             =   315
            Width           =   1470
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
            TabIndex        =   17
            Top             =   645
            Width           =   1455
         End
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   1560
         Width           =   6405
         _Version        =   196608
         _ExtentX        =   11298
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
         Left            =   600
         TabIndex        =   4
         Top             =   2160
         Width           =   6405
         _Version        =   196608
         _ExtentX        =   11298
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
         Index           =   7
         Left            =   5085
         TabIndex        =   9
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
         MaxLength       =   12
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
         Left            =   4035
         TabIndex        =   6
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
         Index           =   3
         Left            =   600
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
         Index           =   0
         Left            =   600
         TabIndex        =   2
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
         MaxLength       =   12
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
         Left            =   600
         TabIndex        =   12
         Top             =   5160
         Width           =   6405
         _Version        =   196608
         _ExtentX        =   11298
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
         Index           =   5
         Left            =   600
         TabIndex        =   7
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
         MaxLength       =   12
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
         Left            =   600
         TabIndex        =   10
         Top             =   3960
         Width           =   6405
         _Version        =   196608
         _ExtentX        =   11298
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
         Index           =   9
         Left            =   600
         TabIndex        =   11
         Top             =   4560
         Width           =   6405
         _Version        =   196608
         _ExtentX        =   11298
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
         Index           =   6
         Left            =   2655
         TabIndex        =   8
         Top             =   3360
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
         MaxLength       =   12
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
         Left            =   600
         TabIndex        =   43
         Top             =   7185
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   1875
         TabIndex        =   44
         Top             =   7185
         Width           =   5190
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   1320
         Picture         =   "M_Provee.frx":053B
         Top             =   7080
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Municipio"
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
         Left            =   600
         TabIndex        =   42
         Top             =   6885
         Width           =   825
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   675
         TabIndex        =   41
         Top             =   6540
         Width           =   3225
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   4260
         TabIndex        =   40
         Top             =   5865
         Width           =   2760
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   10
         Left            =   675
         TabIndex        =   39
         Top             =   5865
         Width           =   3345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Aplicar Cuota Hortofruticola"
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
         Left            =   600
         TabIndex        =   38
         Top             =   6195
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Autoretenedor"
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
         Left            =   4200
         TabIndex        =   37
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Regimen Impuesto"
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
         Left            =   600
         TabIndex        =   36
         Top             =   5520
         Width           =   1575
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
         Left            =   600
         TabIndex        =   31
         Top             =   720
         Width           =   315
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
         Left            =   600
         TabIndex        =   30
         Top             =   3120
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
         Left            =   4035
         TabIndex        =   29
         Top             =   2520
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
         Left            =   600
         TabIndex        =   28
         Top             =   2520
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
         Left            =   600
         TabIndex        =   27
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
         Left            =   600
         TabIndex        =   26
         Top             =   1320
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
         Left            =   2655
         TabIndex        =   25
         Top             =   3120
         Width           =   1110
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
         TabIndex        =   24
         Top             =   3120
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
         Left            =   600
         TabIndex        =   23
         Top             =   3720
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
         Left            =   600
         TabIndex        =   22
         Top             =   4320
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
         Left            =   600
         TabIndex        =   21
         Top             =   4920
         Width           =   465
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   1920
         TabIndex        =   45
         Top             =   7245
         Width           =   5190
      End
   End
End
Attribute VB_Name = "M_Provee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String, codigo As String, v_rut As String, Est As Boolean

Private Sub Check1_Click(Index As Integer)
If Est Then Exit Sub
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
If Trim(modo) = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub Combo2_Click(Index As Integer)
If Est Then Exit Sub
If Toolbar1.Buttons(12).Visible = False Then
   SSTab1.Tab = 1
   SSTab1.TabEnabled(1) = True
   If Trim(modo) = "" Then modo = "M"
   Gl_Ac_Botones Me, 1, 0, modo
End If
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 8730
Me.Width = 8085
MsgTitulo = "Proveedor"
fg_centra Me
Me.HelpContextID = vg_OpcM
SSTab1.Tab = 0
modo = ""
Combo1.ListIndex = 1
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
'-------> Cargar Regimen Impuesto
Combo2(0).Clear
Combo2(0).AddItem "Regimen Simplificado" & Space(150) & "(1)"
Combo2(0).AddItem "Regimen Común" & Space(150) & "(2)"
Combo2(0).AddItem "Gran Contribuyente" & Space(150) & "(3)"
Combo2(0).ListIndex = -1
'-------> Cargar Autoretenedor
Combo2(1).Clear
Combo2(1).AddItem "SI" & Space(150) & "(S)"
Combo2(1).AddItem "NO" & Space(150) & "(N)"
Combo2(1).ListIndex = -1
'-------> Cargar Cuota Hortofruticola
Combo2(2).Clear
Combo2(2).AddItem "SI" & Space(150) & "(S)"
Combo2(2).AddItem "NO" & Space(150) & "(N)"
Combo2(2).ListIndex = -1
MoverDatosGrilla
MoverDatos
Est = False
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then SSTab1.Move 0, Toolbar1.Height, ScaleWidth, ScaleHeight - Toolbar1.Height
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
If Est Then Exit Sub
Set RS = vg_db.Execute("sgpadm_s_municipio 1, " & Val(fpLongInteger1(0).Value) & ", ''")
fpayuda(0).Caption = ""
If Not RS.EOF Then fpayuda(0).Caption = Trim(RS!mun_nombre)
RS.Close: Set RS = Nothing
If Toolbar1.Buttons(12).Visible = False Then
   SSTab1.Tab = 1
   SSTab1.TabEnabled(1) = True
   If Trim(modo) = "" Then modo = "M"
   Gl_Ac_Botones Me, 1, 0, modo
End If
End Sub

Private Sub fpText_Change(Index As Integer)
If Est Then Exit Sub
If Toolbar1.Buttons(12).Visible = False Then
   SSTab1.Tab = 1
   SSTab1.TabEnabled(1) = True
   If Trim(modo) = "" Then modo = "M"
   Gl_Ac_Botones Me, 1, 0, modo
End If
End Sub

Private Sub fpText_GotFocus(Index As Integer)
'If est Then Exit Sub
Select Case Index
Case 0
    If Trim(fpText(0).text) = "" Or vg_Dig = "N" Then Exit Sub
    fpText(0).text = fg_DespintaRut(fpText(0).text)
    fpText(0).text = Mid(fpText(0).text, 1, Len(Trim(fpText(0).text)) - 1)
End Select
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_LostFocus(Index As Integer)
'If est Then Exit Sub
Select Case Index
Case 0
    fpText(Index).text = UCase(fpText(Index).text)
    If Trim(fpText(0).text) = "" Or vg_Dig = "N" Then Exit Sub
    fpText(0).text = fg_RutDig(Trim(fpText(0).text))
    fpText(0).text = fg_PintaRut(fpText(0).text)
End Select
End Sub

Private Sub fpTnombre_Change()
If LimpiaDato(Trim(FptNombre.text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
   Set RS = vg_db.Execute("sgpadm_s_proveedor 2, '', '%" & UCase(LimpiaDato(Trim(FptNombre.text))) & "%'")
   If RS.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS!nReg
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
   Set RS = vg_db.Execute("sgpadm_s_proveedor 3, '', '%" & UCase(LimpiaDato(Trim(FptNombre.text))) & "%'")
   If RS.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS!nReg
End If
i = 1
If Not RS.EOF Then
    Do While Not RS.EOF
        vaSpread1.Row = i
        i = i + 1
        
        vaSpread1.Col = 1
        vaSpread1.text = IIf(IsNull(RS!prv_codigo), "", Trim(RS!prv_codigo))
        
        vaSpread1.Col = 2
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.text = IIf(IsNull(RS!prv_nombre), "", Trim(RS!prv_nombre))
        
        vaSpread1.Col = 3
        vaSpread1.text = IIf(IsNull(RS!prv_activo) Or Trim(RS!prv_activo) = "0", "1", "0")
        
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    vaSpread1_Click 1, 1
    SSTab1.TabEnabled(1) = True
    modo = ""
    Gl_Ac_Botones Me, 1, 1, modo
Else
    RS.Close: Set RS = Nothing
    SSTab1.TabEnabled(1) = False
End If
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
End Sub

Private Sub Image1_Click(Index As Integer)
vg_left = fpayuda(0).Left + 5300
vg_nombre = "": vg_codigo = ""
B_TabEst.LlenaDatos "a_municipio", "mun_", "Municipio", "Gen"
B_TabEst.Show 1
Me.Refresh
If vg_codigo = "" Then Exit Sub
fpLongInteger1(0).Value = Val(vg_codigo)
fpayuda(0).Caption = Trim(vg_nombre)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.Tab = 1: SSTab1.TabEnabled(1) = True
    Est = True
    Check1(0).Value = 1
    Check1(1).Value = 0
    Check1(2).Value = 0
    For i = 0 To 11
        If i < 11 Then fpText(i).Enabled = True: fpText(i).text = ""
        If i < 3 Then Combo2(i).ListIndex = -1
    Next i
    fpLongInteger1(0).text = ""
    fpayuda(0).Caption = ""
    BloquearOpSistema
    Est = False: Est = True
Case 3
    If vaSpread1.MaxRows < 1 Then Exit Sub
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False: itab = 1
    SSTab1.Tab = 1: SSTab1.TabEnabled(1) = True
    MoverDatos
Case 5
    Borra_Datos
Case 7
    modo = ""
    SSTab1.Tab = 0
    MoverDatosGrilla
Case 10
    Cancela_Datos
Case 12
    Actualiza_Datos
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, , MsgTitulo: Exit Sub
    I_Provee
Case 18
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub MoverDatosGrilla()
On Error GoTo Man_Error
fg_carga ""
vaSpread1.MaxRows = 0
Set RS = vg_db.Execute("sgpadm_s_proveedor 1, '',''")
If Not RS.EOF Then
    Do While Not RS.EOF
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
                
        vaSpread1.Col = 1
        vaSpread1.text = IIf(IsNull(RS!prv_codigo), "", RS!prv_codigo)
        
        vaSpread1.Col = 2
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.text = IIf(IsNull(RS!prv_nombre), "", Trim(RS!prv_nombre))
               
        vaSpread1.Col = 3
        vaSpread1.text = IIf(IsNull(RS!prv_activo) Or Trim(RS!prv_activo) = "0", "1", "0")
        
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
RS.Close: Set RS = Nothing
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
FptNombre.text = ""
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Proveedor"
End Sub

Private Sub MoverDatos()
fg_carga ""
Est = True
For i = 0 To 11
    If i < 11 Then fpText(i).text = "": fpText(i).Enabled = True
Next i
If vaSpread1.MaxRows > 0 Then vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: codigo = vaSpread1.Value
fpText(0).Enabled = False
Combo2(0).ListIndex = -1
Combo2(1).ListIndex = -1
Combo2(2).ListIndex = -1
fpayuda(0).Caption = ""
Set RS1 = vg_db.Execute("sgpadm_s_proveedor 6, '" & codigo & "',''")
If Not RS1.EOF Then
    fpText(0).text = fg_PintaRut(RS1!prv_codigo)
    fpText(1).text = IIf(IsNull(RS1!prv_nombre), "", Trim(RS1!prv_nombre))
    fpText(2).text = IIf(IsNull(RS1!prv_direccion), "", Trim(RS1!prv_direccion))
    fpText(3).text = IIf(IsNull(RS1!prv_comuna), "", Trim(RS1!prv_comuna))
    fpText(4).text = IIf(IsNull(RS1!prv_ciudad), "", Trim(RS1!prv_ciudad))
    fpText(5).text = IIf(IsNull(RS1!prv_fono1), "", Trim(RS1!prv_fono1))
    fpText(6).text = IIf(IsNull(RS1!prv_fono2), "", Trim(RS1!prv_fono2))
    fpText(7).text = IIf(IsNull(RS1!prv_fax), "", Trim(RS1!prv_fax))
    fpText(8).text = IIf(IsNull(RS1!prv_percon), "", Trim(RS1!prv_percon))
    fpText(9).text = IIf(IsNull(RS1!prv_giro), "", Trim(RS1!prv_giro))
    fpText(10).text = IIf(IsNull(RS1!prv_emapro), "", Trim(RS1!prv_emapro))
    Check1(0).Value = IIf(IsNull(RS1!prv_activo) Or Trim(RS1!prv_activo) = "" Or RS1!prv_activo = "0", 1, 0)
    If Not IsNull(RS1!prv_regimp) Or Trim(RS1!prv_regimp) <> "" Then Combo2(0).ListIndex = fg_buscacbo(Combo2, 0, 1, (RS1!prv_regimp))
    If Not IsNull(RS1!prv_autret) Or Trim(RS1!prv_autret) <> "" Then Combo2(1).ListIndex = fg_buscacbostring(Combo2, 1, 1, (RS1!prv_autret))
    If Not IsNull(RS1!prv_cuohor) Or Trim(RS1!prv_cuohor) <> "" Then Combo2(2).ListIndex = fg_buscacbostring(Combo2, 2, 1, (RS1!prv_cuohor))
    fpLongInteger1(0).text = IIf(IsNull(RS1!prv_codmun), "", RS1!prv_codmun)
    fpayuda(0).Caption = IIf(IsNull(RS1!mun_nombre), "", Trim(RS1!mun_nombre))
    Check1(1).Value = IIf(IsNull(RS1!prv_docele) Or Trim(RS1!prv_docele) = "N" Or RS1!prv_docele = "0", 0, 1)
    Check1(2).Value = IIf(IsNull(RS1!prv_permiteingdoc) Or RS1!prv_permiteingdoc = False, 0, 1)
End If
RS1.Close: Set RS1 = Nothing
BloquearOpSistema
Est = False
fg_descarga
End Sub

Private Sub Borra_Datos()
On Error GoTo Man_Error
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: codigo = vaSpread1.text
If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
vg_db.Execute "sgpadm_iu_proveedor 'M2' , '" & codigo & "', '', '', '', '', '', '', '', '', '', '', '2', '" & Format(Date, "mm/dd/yyyy") & "', '0', '', '', '', 0, '',0"
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

Private Sub Cancela_Datos()
If MsgBox("Cancela registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
SSTab1.TabEnabled(0) = True
Set RS2 = vg_db.Execute("sgpadm_s_proveedor 2, '', '%" & UCase(LimpiaDato(Trim(FptNombre.text))) & "%'")
If RS2.EOF Or Not IsNumeric(RS2!nReg) Then RS2.Close: Set RS2 = Nothing: SSTab1.TabEnabled(1) = False: modo = "NE": SSTab1.Tab = 0: Gl_Ac_Botones Me, 1, 2, modo: Exit Sub
RS2.Close: Set RS2 = Nothing
If vaSpread1.MaxRows > 0 Then
   MoverDatos
   SSTab1.TabEnabled(1) = True
Else
   SSTab1.TabEnabled(1) = False
End If
SSTab1.Tab = 0
modo = ""
Gl_Ac_Botones Me, 1, 1, modo
End Sub

Private Sub Actualiza_Datos()
On Error GoTo Man_Error
Dim regimp As String, autret As String, cuohor As String
Dim codmun As Long
v_rut = fg_DespintaRut(fpText(0).text)
regimp = IIf(fg_codigocbo(Combo2, 0, 1, "") = "0", "", fg_codigocbo(Combo2, 0, 1, ""))
autret = IIf(fg_codigocbo(Combo2, 1, 1, "") = "0", "", fg_codigocbo(Combo2, 1, 1, ""))
cuohor = IIf(fg_codigocbo(Combo2, 2, 1, "") = "0", "", fg_codigocbo(Combo2, 2, 1, ""))
codmun = IIf(Trim(fpLongInteger1(0).text) = "", 0, fpLongInteger1(0).Value)
If modo = "A" Then
    If Trim(fpText(0).text) = "" Or Trim(fpText(1).text) = "" Then MsgBox "Faltan datos importantes para identificar el proveedor...", vbExclamation + vbOKOnly, "Maestro de Proveedor": Exit Sub
    If Not fg_Check_Rut(v_rut) Then MsgBox "El rut no es valido...", vbExclamation + vbOKOnly, "Valida rut": Exit Sub
    RS.Open "SELECT * FROM b_proveedor WHERE prv_codigo = '" & v_rut & "'", vg_db, adOpenStatic
    If Not RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Proveedor existe", vbExclamation + vbOKOnly, "Maestro de Proveedor": Exit Sub
    RS.Close: Set RS = Nothing
    vg_db.Execute "sgpadm_iu_proveedor 'A', '" & v_rut & "', '" & LimpiaDato(Trim(fpText(1).text)) & "', '" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                  "'" & LimpiaDato(Trim(fpText(3).text)) & "', '" & LimpiaDato(Trim(fpText(4).text)) & "', " & _
                  "'" & LimpiaDato(Trim(fpText(5).text)) & "', '" & LimpiaDato(Trim(fpText(6).text)) & "', " & _
                  "'" & LimpiaDato(Trim(fpText(7).text)) & "', '" & LimpiaDato(Trim(fpText(8).text)) & "', " & _
                  "'" & LimpiaDato(Trim(fpText(9).text)) & "', '" & LimpiaDato(Trim(fpText(10).text)) & "', " & _
                  "'" & IIf(Check1(0).Value = 1, "0", "1") & "', '" & Format(Date, "mm/dd/yyyy") & "', '1', '" & regimp & "', '" & autret & "', '" & cuohor & "', " & codmun & ", '" & IIf(Check1(1).Value = 1, "S", "N") & "', " & IIf(Check1(2).Value = 1, 1, 0) & ""
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = 1: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = v_rut
    vaSpread1.Col = 2: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = LimpiaDato(Trim(fpText(1).text))
    vaSpread1.Col = 3: vaSpread1.text = IIf(Check1(0).Value = 1, "1", "0")
Else
    If Trim(fpText(1).text) = "" Then MsgBox "Faltan datos importantes para identificar el proveedor...", vbExclamation + vbOKOnly, "Maestro de Proveedor": Exit Sub
    vg_db.Execute "sgpadm_iu_proveedor 'M1' , '" & v_rut & "', '" & LimpiaDato(Trim(fpText(1).text)) & "', '" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                  "'" & LimpiaDato(Trim(fpText(3).text)) & "', '" & LimpiaDato(Trim(fpText(4).text)) & "', '" + LimpiaDato(Trim(fpText(5).text)) & "', " & _
                  "'" & LimpiaDato(Trim(fpText(6).text)) & "', '" & LimpiaDato(Trim(fpText(7).text)) & "', '" & LimpiaDato(Trim(fpText(8).text)) & "', " & _
                  "'" & LimpiaDato(Trim(fpText(9).text)) & "', '" & LimpiaDato(Trim(fpText(10).text)) & "', '" & IIf(Check1(0).Value = 1, "0", "1") & "', " & _
                  "'" & Format(Date, "mm/dd/yyyy") & "', '1', '" & regimp & "', '" & autret & "', '" & cuohor & "', " & codmun & ", '" & IIf(Check1(1).Value = 1, "S", "N") & "', " & IIf(Check1(2).Value = 1, 1, 0) & ""
    vaSpread1.Col = 2: vaSpread1.text = LimpiaDato(Trim(fpText(1).text))
    vaSpread1.Col = 3: vaSpread1.text = IIf(Check1(0).Value = 1, "1", "0")
End If
   
vaSpread1.SortKey(1) = 2
vaSpread1.SortKeyOrder(1) = 1
vaSpread1.Sort 1, 1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
   
SSTab1.TabEnabled(0) = True
If vaSpread1.MaxRows < 1 Then
   SSTab1.TabEnabled(1) = False
Else
   SSTab1.TabEnabled(1) = True
   SSTab1.Tab = 0
End If
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

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows > 0 Then modo = "": MoverDatos
End Sub

Sub BloquearOpSistema()
'-------> bloquear opciones del sistema si el pasi = chile
Label1(13).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Label1(14).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Label1(16).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Label1(17).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Combo2(0).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Combo2(1).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Combo2(2).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Image1(0).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
fpLongInteger1(0).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
End Sub


