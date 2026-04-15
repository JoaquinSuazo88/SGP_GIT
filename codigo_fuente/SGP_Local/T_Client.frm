VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form T_Client 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Casino"
   ClientHeight    =   6135
   ClientLeft      =   1155
   ClientTop       =   1380
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   9510
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   4
      OLEDropMode     =   1
      TabCaption(0)   =   "Clientes"
      TabPicture(0)   =   "T_Client.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "T_Client.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(15)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(5)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(4)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(3)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(7)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(8)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(9)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(10)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(12)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "fpText(6)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "fpText(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "fpText(5)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "fpText(4)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "fpText(3)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "fpText(2)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "fpText(1)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -73560
         TabIndex        =   4
         Top             =   480
         Width           =   6615
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "T_Client.frx":0038
            Left            =   1680
            List            =   "T_Client.frx":0042
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   2865
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Left            =   1680
            TabIndex        =   6
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
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
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
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   255
            TabIndex        =   9
            Top             =   675
            Width           =   1350
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
            TabIndex        =   8
            Top             =   675
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label Label1 
            Caption         =   "Buscar Columna"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   255
            TabIndex        =   7
            Top             =   345
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   -73800
         TabIndex        =   2
         Top             =   1680
         Width           =   7185
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3615
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   6765
            _Version        =   393216
            _ExtentX        =   11933
            _ExtentY        =   6376
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
            MaxCols         =   6
            MaxRows         =   20
            OperationMode   =   3
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "T_Client.frx":0056
            VisibleCols     =   6
            VisibleRows     =   15
         End
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   2955
         TabIndex        =   10
         Top             =   1275
         Width           =   4605
         _Version        =   196608
         _ExtentX        =   8123
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
         Index           =   2
         Left            =   2955
         TabIndex        =   11
         Top             =   1590
         Width           =   4605
         _Version        =   196608
         _ExtentX        =   8123
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
         Index           =   3
         Left            =   2955
         TabIndex        =   12
         Top             =   1905
         Width           =   1860
         _Version        =   196608
         _ExtentX        =   3281
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
         Left            =   3000
         TabIndex        =   13
         Top             =   2280
         Width           =   2100
         _Version        =   196608
         _ExtentX        =   3704
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
         Index           =   5
         Left            =   2955
         TabIndex        =   14
         Top             =   2535
         Width           =   2220
         _Version        =   196608
         _ExtentX        =   3916
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
         Index           =   0
         Left            =   2955
         TabIndex        =   15
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
         CharValidationText=   ""
         MaxLength       =   5
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
         Left            =   3000
         TabIndex        =   25
         Top             =   2880
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
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   1920
         TabIndex        =   27
         Top             =   4320
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Giro"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   1920
         TabIndex        =   26
         Top             =   3960
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Contactos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   1920
         TabIndex        =   24
         Top             =   3600
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   1920
         TabIndex        =   23
         Top             =   3240
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Fono Nº 2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   1920
         TabIndex        =   22
         Top             =   3000
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1920
         TabIndex        =   21
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1920
         TabIndex        =   20
         Top             =   1695
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Comuna"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   1920
         TabIndex        =   19
         Top             =   2010
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Ciudad"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   1920
         TabIndex        =   18
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Fono Nº 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   1920
         TabIndex        =   17
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   1920
         TabIndex        =   16
         Top             =   1065
         Width           =   900
      End
   End
End
Attribute VB_Name = "T_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim ibusca As Long, i As Long
Dim itab As Integer, swvalidar As Integer, itexto As Integer, opboton As Integer
Dim cAccion As String, modo As String, codigo As String
Dim vecdatos(11) As String
Private Sub Form_Activate()

fg_descarga
End Sub
Private Sub Form_Load()

Me.Height = 6510
Me.Width = 9600
fg_centra Me
SSTab1.Tab = 0
SSTab1.TabEnabled(1) = False
modo = ""
Combo2.ListIndex = 1

Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): btnX.Visible = True: btnX.ToolTipText = "Incluir"
Set btnX = Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): btnX.Visible = False: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, "A_Alterar", , tbrDefault, "A_Alterar"): btnX.Visible = True: btnX.ToolTipText = "Alterar"
Set btnX = Toolbar1.Buttons.Add(, "I_Alterar", , tbrDefault, "I_Alterar"): btnX.Visible = False: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): btnX.Visible = True: btnX.ToolTipText = "Borrar "
Set btnX = Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): btnX.Visible = False: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): btnX.Visible = True: btnX.ToolTipText = "Actualizar Lista   "
Set btnX = Toolbar1.Buttons.Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): btnX.Visible = False: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): btnX.Visible = False: btnX.ToolTipText = "Cancelar "
Set btnX = Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): btnX.Visible = True: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = False: btnX.ToolTipText = "Confirmar "
Set btnX = Toolbar1.Buttons.Add(, "I_Conformar ", , tbrDefault, "I_Confirmar "): btnX.Visible = True: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): btnX.Visible = True: btnX.ToolTipText = "Imprimir "
Set btnX = Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): btnX.Visible = False: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"

MoverDatosGrilla
End Sub
Private Sub Form_Resize()

If Me.WindowState <> 1 Then SSTab1.Move 0, Toolbar1.Height, ScaleWidth, ScaleHeight - Toolbar1.Height
End Sub
Private Sub fpText_Change(Index As Integer)

If fpText(Index).Text <> vecdatos(Index) And itexto = 0 Then
   SSTab1.TabEnabled(0) = False
   itab = 1
   SSTab1.Tab = 1
   SSTab1.TabEnabled(1) = True
   Ac_Botones 0
   itab = 0
End If
End Sub
Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub
Private Sub fpText_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 13 And Shift = 2 Then Actualiza_Datos: Exit Sub
Select Case KeyCode
  Case 27 And Toolbar1.Buttons(9).Visible = True And Toolbar1.Buttons(11).Visible = True
    Cancela_Datos
  Case 113 And Toolbar1.Buttons(1).Visible = True
    modo = "A"
    Agrega_Datos
  Case 114 And Toolbar1.Buttons(3).Visible = True
    modo = "M"
    Agrega_Datos
  Case 115 And Toolbar1.Buttons(5).Visible = True
    Borra_Datos
End Select
End Sub
Private Sub fpText1_Change()

If fpText1.Text <> vecdatos(10) And itexto = 0 Then
   SSTab1.TabEnabled(0) = False
   itab = 1
   SSTab1.Tab = 1
   SSTab1.TabEnabled(1) = True
   Ac_Botones 0
   itab = 0
End If
End Sub
Private Sub fpText1_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
  Case 27 And Toolbar1.Buttons(9).Visible = True And Toolbar1.Buttons(11).Visible = True
    Cancela_Datos
  Case 113 And Toolbar1.Buttons(1).Visible = True
    modo = "A"
    Agrega_Datos
  Case 114 And Toolbar1.Buttons(3).Visible = True
    modo = "M"
    Agrega_Datos
  Case 115 And Toolbar1.Buttons(5).Visible = True
    Borra_Datos
End Select
End Sub
Private Sub fpTnombre_Change()
If LimpiaDato(Trim(fptnombre.Text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo2.ItemData(Combo2.ListIndex) = 0 Then
   RS.Open "select count(*) as nreg From b_clientes where ucase(cli_codigo) like '%" & UCase(LimpiaDato(fptnombre.Text)) & "%'", vg_db, adOpenStatic
   If RS.EOF Or RS!NReg = 0 Then RS.Close: Set RS = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: SSTab1.TabEnabled(1) = False: modo = "NE": Ac_Botones: Exit Sub
   If ibusca <> RS!NReg Then ibusca = RS!NReg: vaSpread1.MaxRows = RS!NReg
   RS.Close: Set RS = Nothing
   RS.Open "select cli_codigo, cli_nombre from b_clientes Where ucase(cli_codigo) like '%" & UCase(LimpiaDato(fptnombre.Text)) & "%' order by cli_codigo", vg_db, adOpenStatic
ElseIf Combo2.ItemData(Combo2.ListIndex) = 1 Then
   RS.Open "select count(*) as nreg From b_clientes where Ucase(cli_nombre) like '%" & UCase(LimpiaDato(fptnombre.Text)) & "%'", vg_db, adOpenStatic
   If RS.EOF Or RS!NReg = 0 Then RS.Close: Set RS = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: SSTab1.TabEnabled(1) = False: modo = "NE": Ac_Botones: Exit Sub
   If ibusca <> RS!NReg Then ibusca = RS!NReg: vaSpread1.MaxRows = RS!NReg
   RS.Close: Set RS = Nothing
   RS.Open "select cli_codigo, cli_nombre From b_clientes Where Ucase(cli_nombre) like '%" & UCase(LimpiaDato(fptnombre.Text)) & "%' order by cli_nombre", vg_db, adOpenStatic
End If
i = 1
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.Row = i
      i = i + 1

      vaSpread1.Col = 1
      vaSpread1.Text = RS!cli_codigo

      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = 0
      vaSpread1.Text = Trim(RS!cli_nombre)
      RS.MoveNext
   Loop
   SSTab1.TabEnabled(1) = True
   Ac_Botones 1
Else
   SSTab1.TabEnabled(1) = False
End If
RS.Close: Set RS = Nothing
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)

Select Case SSTab1.Tab
  Case 0
    DeshabilitaDatos
  Case 1
    If vaSpread1.MaxRows > 0 And itab = 0 Then
       modo = "M"
       SSTab1.TabEnabled(0) = True
       SSTab1.Tab = 1
       SSTab1.TabEnabled(1) = True
       itexto = 1
       MoverDatos
       itexto = 0
    End If
End Select
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
  Case 1
    modo = "A"
    Ac_Botones
    SSTab1.TabEnabled(0) = False: itab = 1
    SSTab1.Tab = 1: SSTab1.TabEnabled(1) = True
    itexto = 1
    HabilitaDatos
    fpText(0).Enabled = True
    For i = 0 To 9
        fpText(1).Text = ""
    Next i
    For i = 0 To 11
        vecdatos(i) = ""
    Next i
    itexto = 0: itab = 0
  Case 3
    If vaSpread1.MaxRows < 1 Then Exit Sub
    modo = "M"
    Ac_Botones
    SSTab1.TabEnabled(0) = False: itab = 1
    SSTab1.Tab = 1: SSTab1.TabEnabled(1) = True
    itexto = 1
    MoverDatos
    itexto = 0: itab = 0
  Case 5
    Borra_Datos
  Case 7
    MoverDatosGrilla
  Case 10
    Cancela_Datos
  Case 12
    Actualiza_Datos
  Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, "Proveedor": Exit Sub
'    vg_OpImp = 1
'    Preview.Show 1
  Case 18
    Me.Hide
    Unload Me
End Select
End Sub
Private Sub MoverDatosGrilla()

On Error GoTo Man_Error

fg_carga (ss)
vaSpread1.MaxRows = 0
itab = 0
RS.Open "select * from b_clientes order by cli_nombre", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
              
      vaSpread1.Col = 1
      vaSpread1.Text = RS!cli_codigo

      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = 0
      vaSpread1.Text = Trim(RS!cli_nombre)
             
      RS.MoveNext
   Loop
   modo = "M"
   Ac_Botones
   SSTab1.TabEnabled(1) = True
Else
   SSTab1.Tab = 0
   SSTab1.TabEnabled(1) = False
   modo = "NE"
   Ac_Botones
End If
RS.Close: Set RS = Nothing
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
fptnombre.Text = ""
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Proveedor"
End Sub
Private Sub MoverDatos()

fg_carga (ss)
HabilitaDatos
itexto = 1
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: codigo = vaSpread1.Text
fpText(0).Enabled = False
For i = 0 To 11
    vecdatos(i) = ""
Next i

RS.Open "select * from b_clientes where cli_codigo='" & codigo & "", vg_db, adOpenStatic
If Not RS.EOF Then
   fpText(0).Text = RS!cli_codigo: vecdatos(0) = RS!cli_codigo
   fpText(1).Text = Trim(RS!cli_nombre): vecdatos(1) = Trim(RS!cli_nombre)
   fpText(2).Text = Trim(RS!cli_direccion): vecdatos(2) = Trim(RS!cli_comuna)
   fpText(3).Text = Trim(RS!cli_ciudad): vecdatos(3) = Trim(RS!cli_ciudad)
   fpText(4).Text = Trim(RS!cli_fono1): vecdatos(4) = Trim(RS!cli_fono1)
   fpText(5).Text = Trim(RS!cli_fono2): vecdatos(5) = Trim(RS!cli_fono2)
   fpText(6).Text = Trim(RS!cli_fax): vecdatos(6) = Trim(RS!cli_fax)
   fpText(7).Text = Trim(RS!cli_percon): vecdatos(7) = Trim(RS!cli_percon)
   fpText(8).Text = Trim(RS!cli_giro): vecdatos(8) = Trim(RS!cli_giro)
   fpText(9).Text = Trim(RS!cli_email): vecdatos(9) = Trim(RS!cli_email)
   fpText(10).Text = Trim(RS!cli_tipo): vecdatos(10) = Trim(RS!cli_tipo)
End If
RS.Close: Set RS = Nothing

fg_descarga
End Sub
Private Sub Borra_Datos()

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: codigo = vaSpread1.Text
TITLE = "Eliminar Dato"
Resp_Delete (TITLE)
If respuesta = vbYes Then
   vg_db.BeginTrans
     ' ***      Borrando Casino *** '
     vg_db.Execute "delete b_clientes from b_clientes where cli_codigo='" & codigo & "'"
     vaSpread1.Row = vaSpread1.ActiveRow
     vaSpread1.DeleteRows vaSpread1.Row, 1
     vaSpread1.MaxRows = vaSpread1.MaxRows - 1
     vaSpread1.Row = vaSpread1.MaxRows
     Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
     If vaSpread1.MaxRows < 1 Then
        SSTab1.TabEnabled(1) = False
        SSTab1.Tab = 0
     Else
        SSTab1.TabEnabled(1) = True
        SSTab1.Tab = 0
     End If
  vg_db.CommitTrans
End If
fptnombre.SetFocus

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Private Sub Cancela_Datos()

TITLE = "Casino"
msg = "Cancelar Operación"
Style = vbYesNo + vbQuestion + vbDefaultButton2
Help = "DEMO.HLP"
Ctxt = 1000
ws_respuesta = MsgBox(msg, Style, TITLE, Help, Ctxt)
Select Case ws_respuesta
  Case Is = vbYes
    SSTab1.TabEnabled(0) = True
    RS.Open "select count(*) as nreg from b_clientes where ucase(cli_codigo) like '%" & UCase(("")) & "%'", vg_db, adOpenStatic
    If RS.EOF Or RS!NReg = 0 Then
       RS.Close: Set RS = Nothing
       SSTab1.TabEnabled(1) = False
       modo = "NE"
       SSTab1.Tab = 0
    ElseIf RS!NReg > 0 Then
       RS.Close: Set RS = Nothing
       If vaSpread1.MaxRows > 1 Then
          SSTab1.TabEnabled(1) = True
       Else
          SSTab1.TabEnabled(1) = False
       End If
       SSTab1.Tab = 0
       modo = ""
    End If
    Ac_Botones
  Case Is = vbCancel
    Exit Sub
End Select
End Sub
Private Sub Actualiza_Datos()

On Error GoTo Man_Error

If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
   swvalidar = 0
   ValidarCampos
   If swvalidar = 1 Then Exit Sub
   vg_db.BeginTrans
     If modo = "A" Then
        Set ConSql = vg_db.Execute("select * " & _
                     "From Sdx_Casino " & _
                     "where Codigo_Casino='" & "00000" + LimpiaDato(Trim(fpText(0).Text)) & "'", , adCmdText)
        If Not ConSql.EOF Then
           If ConSql!IndBorrado = 1 Then
              MsgBox "Código Ya Existe, Fue Eliminado Intente con otro Código", vbExclamation + vbOKOnly, "Casino"
           Else
              MsgBox "Código Ya Existe", vbExclamation + vbOKOnly, "Casino"
           End If
           ConSql.Close: Set ConSql = Nothing
           Exit Sub
        Else
           ConSql.Close: Set ConSql = Nothing
           vg_db.Execute "insert into Sdx_Casino (Codigo_Casino, Nombre_Casino, Direccion_Casino, " & _
                         "Ciudad_Casino, Fono_Casino, Fax_Casino, Glosa_Casino, PrimerNombre_Casino, " & _
                         "SegundoNombre_Casino, PrimerApellido_Casino, SegundoApellido_Casino, " & _
                         "FechaCreacion_Casino, Codigo_Segmento, IndBorrado, Codigo_Unidad, Version) " & _
                         "values ('" & "00000" + LimpiaDato(fpText(0).Text) & "', " & _
                         "'" & LimpiaDato(Trim(fpText(1).Text)) & "', '" & LimpiaDato(Trim(fpText(2).Text)) & "', " & _
                         "'" & LimpiaDato(Trim(fpText(3).Text)) & "', '" & LimpiaDato(Trim(fpText(4).Text)) & "', " & _
                         "'" & LimpiaDato(Trim(fpText(5).Text)) & "', '" & LimpiaDato(Trim(fpText1.Text)) & "', " & _
                         "'" & LimpiaDato(Trim(fpText(6).Text)) & "', '" & LimpiaDato(Trim(fpText(7).Text)) & "', " & _
                         "'" & LimpiaDato(Trim(fpText(8).Text)) & "', '" & LimpiaDato(Trim(fpText(9).Text)) & "', " & _
                         "" & Format(Date, "yyyymmdd") & ", " & Val(fpLongInteger1.Value) & ", 0, '0', '0')"
            Set ConSql = vg_db.Execute("select * " & _
                         "From Sdx_Casino " & _
                         "where Codigo_Casino='" & "00000" + LimpiaDato(Trim(fpText(0).Text)) & "' " & _
                         "", , adCmdText)
           If Not ConSql.EOF Then
              vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
              vaSpread1.Col = 1: vaSpread1.TypeHAlign = 0: vaSpread1.Value = LimpiaDato(Trim(fpText(0).Text))
              vaSpread1.Col = 2: vaSpread1.TypeHAlign = 0: vaSpread1.Value = Trim(ConSql!Nombre_Casino)
           End If
           ConSql.Close: Set ConSql = Nothing
        End If
     Else
        vg_db.Execute "update Sdx_Casino set Nombre_Casino='" & LimpiaDato(Trim(fpText(1).Text)) & "', Direccion_Casino='" & LimpiaDato(Trim(fpText(2).Text)) & "', " & _
                      "Ciudad_Casino='" & LimpiaDato(Trim(fpText(3).Text)) & "', Fono_Casino='" & LimpiaDato(Trim(fpText(4).Text)) & "', Fax_Casino='" + LimpiaDato(Trim(fpText(5).Text)) & "', " & _
                      "Glosa_Casino='" & LimpiaDato(Trim(fpText1.Text)) & "', PrimerNombre_Casino='" & LimpiaDato(Trim(fpText(6).Text)) & "', SegundoNombre_Casino='" & LimpiaDato(Trim(fpText(7).Text)) & "', " & _
                      "PrimerApellido_Casino='" & LimpiaDato(Trim(fpText(8).Text)) & "', SegundoApellido_Casino='" & LimpiaDato(Trim(fpText(9).Text)) & "', Codigo_Segmento=" & Val(fpLongInteger1.Text) & " " & _
                      "where Codigo_Casino='" & "00000" & LimpiaDato(Trim(fpText(0).Text)) & "'"
        vaSpread1.Col = 2
        vaSpread1.Value = LimpiaDato(Trim(fpText(1).Text))
     End If
     vaSpread1.SortKey(1) = 2
     vaSpread1.SortKeyOrder(1) = 1
     vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
     
     SSTab1.TabEnabled(0) = True
     If vaSpread1.MaxRows < 1 Then
        SSTab1.TabEnabled(1) = False
     Else
        SSTab1.TabEnabled(1) = True
        SSTab1.Tab = 0
     End If
     itexto = 1
     Ac_Botones 1
     Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
   vg_db.CommitTrans
End If

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Private Sub DeshabilitaDatos()

fptnombre.Enabled = True
vaSpread1.Enabled = True
For i = 0 To 9
    fpText(1).Enabled = False
Next i
End Sub
Private Sub HabilitaDatos()

vaSpread1.Enabled = False
For i = 0 To 9
    fpText(1).Enabled = True
Next i
End Sub
Function Ac_Botones()

If modo = "A" Or modo = "M" Then
    Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(3).Visible = False: Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False: Toolbar1.Buttons(8).Visible = True
    Toolbar1.Buttons(10).Visible = True: Toolbar1.Buttons(11).Visible = False
    Toolbar1.Buttons(12).Visible = True: Toolbar1.Buttons(13).Visible = False
    Toolbar1.Buttons(15).Visible = False: Toolbar1.Buttons(16).Visible = True
'    Combo1.Enabled = False: fpText1.Enabled = False
ElseIf modo = "" Then
    Toolbar1.Buttons(1).Visible = True: Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(3).Visible = True: Toolbar1.Buttons(4).Visible = False
    Toolbar1.Buttons(5).Visible = True: Toolbar1.Buttons(6).Visible = False
    Toolbar1.Buttons(7).Visible = True: Toolbar1.Buttons(8).Visible = False
    Toolbar1.Buttons(10).Visible = False: Toolbar1.Buttons(11).Visible = True
    Toolbar1.Buttons(12).Visible = False: Toolbar1.Buttons(13).Visible = True
    Toolbar1.Buttons(15).Visible = True: Toolbar1.Buttons(16).Visible = False
ElseIf modo = "NE" Then
    Toolbar1.Buttons(1).Visible = True: Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(3).Visible = False: Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False: Toolbar1.Buttons(8).Visible = True
    Toolbar1.Buttons(10).Visible = False: Toolbar1.Buttons(11).Visible = True
    Toolbar1.Buttons(12).Visible = False: Toolbar1.Buttons(13).Visible = True
    Toolbar1.Buttons(15).Visible = False: Toolbar1.Buttons(16).Visible = True

'    Combo1.Enabled = True: fpText1.Enabled = True
End If
End Function
Private Sub ValidarCampos()

If swvalidar = 0 And fpText(0).Text = "" Then swvalidar = 1: MsgBox "Debe ingresar Código", vbExclamation + vbOKOnly, "Casino": fpText(0).SetFocus
If swvalidar = 0 And fpText(1).Text = "" Then swvalidar = 1: MsgBox "Debe ingresar Nombre", vbExclamation + vbOKOnly, "Casino": fpText(1).SetFocus
'If swvalidar = 0 And fpText(2).Text = "" Then swvalidar = 1: MsgBox "Debe ingresar Dirección", vbExclamation + vbOKOnly, "Casino": fpText(2).SetFocus
End Sub
Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = Row
vaSpread1.Col = 1
codigo = vaSpread1.Value
End Sub
Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
  Case 27 And Toolbar1.Buttons(9).Visible = True And Toolbar1.Buttons(11).Visible = True
    Cancela_Datos
  Case 113 And Toolbar1.Buttons(1).Visible = True
    modo = "A"
    Agrega_Datos
  Case 114 And Toolbar1.Buttons(3).Visible = True
    modo = "M"
    If vaSpread1.MaxRows < 1 Then Exit Sub
    Agrega_Datos
  Case 115 And Toolbar1.Buttons(5).Visible = True
    Borra_Datos
End Select
End Sub
