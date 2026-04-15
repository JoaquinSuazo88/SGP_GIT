VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_UsuCtr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantención Usuario & Control Acceso"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   7335
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   4
      TabCaption(0)   =   "Usuario"
      TabPicture(0)   =   "M_UsuCtr.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_UsuCtr.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(5)"
      Tab(1).Control(1)=   "Label1(4)"
      Tab(1).Control(2)=   "Label1(3)"
      Tab(1).Control(3)=   "Label1(2)"
      Tab(1).Control(4)=   "Label1(0)"
      Tab(1).Control(5)=   "Label1(6)"
      Tab(1).Control(6)=   "fpText1(5)"
      Tab(1).Control(7)=   "fpText1(4)"
      Tab(1).Control(8)=   "fpText1(3)"
      Tab(1).Control(9)=   "fpText1(2)"
      Tab(1).Control(10)=   "fpText1(1)"
      Tab(1).Control(11)=   "fpText1(0)"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Control Acceso"
      TabPicture(2)   =   "M_UsuCtr.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(1)=   "TvwDir"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Control Recetario"
      TabPicture(3)   =   "M_UsuCtr.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "TvwDir2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   -74400
         TabIndex        =   3
         Top             =   540
         Width           =   6015
         Begin EditLib.fpText fpText 
            Height          =   315
            Left            =   1680
            TabIndex        =   4
            Top             =   315
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
            NoSpecialKeys   =   3
            AutoAdvance     =   -1  'True
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
            ThreeDFrameColor=   -2147483637
            Appearance      =   1
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label1 
            Caption         =   " Buscar Texto"
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
            Index           =   1
            Left            =   360
            TabIndex        =   6
            Top             =   415
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Label2"
            Height          =   195
            Left            =   4260
            TabIndex        =   5
            Top             =   435
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   -74200
         TabIndex        =   1
         Top             =   1380
         Width           =   5625
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   2895
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   5205
            _Version        =   393216
            _ExtentX        =   9181
            _ExtentY        =   5106
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
            SpreadDesigner  =   "M_UsuCtr.frx":0070
            VisibleCols     =   2
            VisibleRows     =   15
         End
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   0
         Left            =   -72900
         TabIndex        =   7
         Top             =   1500
         Width           =   2385
         _Version        =   196608
         _ExtentX        =   4207
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   -72900
         TabIndex        =   8
         Top             =   1815
         Width           =   4635
         _Version        =   196608
         _ExtentX        =   8176
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   2
         Left            =   -72900
         TabIndex        =   9
         Top             =   2130
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         MaxLength       =   10
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   3
         Left            =   -72900
         TabIndex        =   10
         Top             =   2445
         Width           =   1815
         _Version        =   196608
         _ExtentX        =   3201
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   4
         Left            =   -72900
         TabIndex        =   11
         Top             =   2760
         Width           =   3735
         _Version        =   196608
         _ExtentX        =   6588
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   5
         Left            =   -72900
         TabIndex        =   12
         Top             =   3075
         Width           =   3735
         _Version        =   196608
         _ExtentX        =   6588
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
      Begin MSComctlLib.TreeView TvwDir 
         Height          =   3855
         Left            =   -74760
         TabIndex        =   19
         Top             =   600
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6800
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
      End
      Begin MSComctlLib.TreeView TvwDir2 
         Height          =   3855
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6800
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Label4"
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
         Left            =   240
         TabIndex        =   23
         Top             =   4560
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
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
         Left            =   -74760
         TabIndex        =   20
         Top             =   4560
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Login"
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
         Left            =   -74280
         TabIndex        =   18
         Top             =   1605
         Width           =   1395
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
         Index           =   0
         Left            =   -74280
         TabIndex        =   17
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Password"
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
         Left            =   -74280
         TabIndex        =   16
         Top             =   2235
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Telefono"
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
         Left            =   -74280
         TabIndex        =   15
         Top             =   2550
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Oficina"
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
         Left            =   -74280
         TabIndex        =   14
         Top             =   2865
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento"
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
         Left            =   -74280
         TabIndex        =   13
         Top             =   3180
         Width           =   1395
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":04D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":0B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":0E20
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":113A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":1454
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":176E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":1A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":1DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":20BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":23D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":26F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":2A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":2D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":303E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incluir"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Alterar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Borrar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Actualizar Lista"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   15
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_UsuCtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConSql As ADODB.Recordset, Consql1 As ADODB.Recordset, Consql2 As ADODB.Recordset
Dim modo As String, loginusuario As String, hijo As String, hijo2 As String
Dim vecdatos(5) As String
Dim ivalidar As Integer, itexto As Integer, itab As Integer, fin As Integer
Dim i As Long, j As Long, ibusca As Long, nivel As Long, codhijo As Long, codhijo2 As Long, indindex As Long
Dim dest As Node, sourcenode As Node, nd As Node, rootnode As Node
'Dim codigo As Long, irow As Long
'Dim ivalidar As Integer, vg_incluir As Integer, vg_alterar As Integer, vg_borrar As Integer
'Dim vg_imprimir As Integer, indavan As Integer
'Dim modo As String, Nombre As String
Private Sub Form_Activate()
fg_descarga
End Sub
Private Sub Form_Load()
Me.Height = 5670
Me.Width = 7425
fg_centra Me
SSTab1.Tab = 0
itab = 0
modo = "M"
MoverDatosGrillas
End Sub
Private Sub fpText_Change()
If LimpiaDato(Trim(fpText.Text)) & Chr(KeyAscii) = "" Then Exit Sub
Set ConSql = vg_db.Execute("select count(loginusuario) as nreg " & _
             "From Sdx_Usuario " & _
             "where ucase(nombre) like '%" + UCase(LimpiaDato(fpText.Text)) + "%'", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_usuario 3, '', '%" + UCase(LimpiaDato(fpText.Text)) + "%'", , adCmdStoredProc)
If ConSql.EOF Or ConSql!NReg = 0 Then ConSql.Close: Set ConSql = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros": Exit Sub
If ibusca <> ConSql!NReg Then ibusca = ConSql!NReg: vaSpread1.MaxRows = ConSql!NReg: ConSql.Close: Set ConSql = Nothing
Set ConSql = vg_db.Execute("select loginusuario, nombre " & _
             "From Sdx_Usuario " & _
             "where ucase(nombre) like '%" + UCase(LimpiaDato(fpText.Text)) + "%' " & _
             "order by nombre", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_usuario 4, '', '%" + UCase(LimpiaDato(fpText.Text)) + "%'", , adCmdStoredProc)
i = 1
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
      vaSpread1.Row = i
      i = i + 1
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = 1
      vaSpread1.Text = Trim(ConSql!loginusuario)
      vaSpread1.Col = 2
      vaSpread1.Text = Trim(ConSql!Nombre)
      ConSql.MoveNext
   Loop
End If
ConSql.Close: Set ConSql = Nothing
If fpText.Text = "" Then
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Else
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Encontrado"
End If
End Sub
Private Sub fpText1_Change(Index As Integer)
If fpText1(Index).Text <> vecdatos(Index) And itexto = 0 And modo = "M" Then
   If Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True Then Exit Sub
   Ac_HabDes 4
   Ac_Boton 1
End If
End Sub
Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub
Private Sub fpText1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
'  Case 13 And Toolbar1.Buttons(9).Visible = True And Toolbar1.Buttons(11).Visible = True
'    Actualiza_Datos
  Case 27 And Toolbar1.Buttons(11).Visible = True And Toolbar1.Buttons(13).Visible = True
'    Cancela_Fila
  Case 113 And Toolbar1.Buttons(1).Visible = True
    modo = "A"
    Agrega_Dato
  Case 114 And Toolbar1.Buttons(3).Visible = True
    modo = "M"
    Agrega_Dato
  Case 115 And Toolbar1.Buttons(5).Visible = True
'    Borra_Fila
End Select
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
  Case 0
    If itab = 2 Then itab = 0: Ac_Boton 2
  Case 1
    If vaSpread1.MaxRows > 0 And modo = "M" Then
       If itab = 2 Then itab = 1: Ac_Boton 2
       modo = "M"
       MoverDetalleDatos
'       M_Receta.Refresh
    ElseIf vaSpread1.MaxRows < 1 And modo = "M" Then
       SSTab1.Tab = 0
       Exit Sub
    End If
  Case 2
    If vaSpread1.MaxRows < 1 Then SSTab1.Tab = 0: Exit Sub
    Ac_Boton 4
    itab = 2
    MoverSistema
  Case 3
    If vaSpread1.MaxRows < 1 Then SSTab1.Tab = 0: Exit Sub
    Ac_Boton 4
    itab = 2
    MoverRecetas
End Select
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 1
    Agrega_Dato
  Case 3
    If vaSpread1.MaxRows < 1 Then Exit Sub
    Altera_Dato
  Case 5
    Borra_Fila
  Case 7
'    fpText1.Text = ""
'    MoverDatosGrillas
  Case 10
    Cancela_Dato
  Case 12
    Actualiza_Dato
  Case 15
'    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, "Color": Exit Sub
'    vg_opimp = 14
'    Preview.Show 1
  Case 18
    Me.Hide
    Unload Me
End Select
End Sub
Sub Agrega_Dato()
itexto = 1
modo = "A"
Ac_Boton 1
Ac_HabDes 2
LimpiarVariable
itexto = 0
End Sub
Sub Altera_Dato()
Ac_HabDes 2
Ac_Boton 1
If SSTab1.Tab = 1 Or SSTab1.Tab = 0 Then
   SSTab1.TabEnabled(2) = False
   SSTab1.Tab = 1
'   MoverDetalleDatos
End If
End Sub
Sub Borra_Fila()

On Error GoTo Man_Error

Resp_Delete ("Mantención")
If respuesta = vbYes Then
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = 1: loginusuario = vaSpread1.Value
   vg_db.BeginTrans
     vg_db.Execute "delete Sdx_UsuCtrlAcceso from Sdx_UsuCtrlAcceso where login='" & loginusuario & "'"
     vg_db.Execute "delete Sdx_Usuario from Sdx_Usuario where loginusuario='" & loginusuario & "'"
   vg_db.CommitTrans
   vaSpread1.Action = 5
   vaSpread1.MaxRows = vaSpread1.MaxRows - 1
   Set ConSql = vg_db.Execute("select count(loginusuario) as nreg " & _
                "From Sdx_Usuario " & _
                "where ucase(nombre) like '%" + UCase(("")) + "%'", , adCmdText)
'   Set ConSql = vg_db.Execute("sod_s_usuario 3, '', '%" + UCase(("")) + "%'", , adCmdStoredProc)
   If ConSql.EOF Or ConSql!NReg = 0 Then
      ConSql.Close: Set ConSql = Nothing
      Ac_Boton 3
   ElseIf ConSql!NReg > 0 Then
      ConSql.Close: Set ConSql = Nothing
      Ac_Boton 2
   End If
End If

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Sub Cancela_Dato()
TITLE = "Usuario"
msg = "Cancelar Operación"
Style = vbYesNo + vbQuestion + vbDefaultButton2
Help = "DEMO.HLP"
Ctxt = 1000
ws_respuesta = MsgBox(msg, Style, TITLE, Help, Ctxt)
Select Case ws_respuesta
  Case Is = vbYes
    If SSTab1.Tab = 1 Then
       Set ConSql = vg_db.Execute("select count(loginusuario) as nreg " & _
                    "From Sdx_Usuario " & _
                    "where ucase(nombre) like '%" + UCase(LimpiaDato("")) + "%'", , adCmdText)
'       Set ConSql = vg_db.Execute("sod_s_usuario 3, '', '%" + UCase(LimpiaDato("")) + "%'", , adCmdStoredProc)
       If ConSql.EOF Or ConSql!NReg = 0 Then
          ConSql.Close: Set ConSql = Nothing
          Ac_HabDes 1
          Ac_Boton 3
          SSTab1.Tab = 0
       ElseIf ConSql!NReg > 0 Then
          modo = "M"
          ConSql.Close: Set ConSql = Nothing
          If modo = "A" Then
             SSTab1.Tab = 0
          ElseIf modo = "M" Then
             MoverDetalleDatos
'             SSTab1.TabEnabled(2) = True
          End If
          Ac_HabDes 3
          Ac_Boton 2
       End If
    End If
  Case Is = vbCancel
    Exit Sub
End Select
End Sub
Sub Actualiza_Dato()

On Error GoTo Man_Error

ivalidar = 0
ValidarCampos
If ivalidar = 1 Then Exit Sub
If modo = "A" Then
   vg_db.BeginTrans
     vg_db.Execute "insert into Sdx_Usuario (loginusuario, passwordusuario, " & _
                   "nombre, telefono, oficina, departamento) " & _
                   "values ('" & LimpiaDato(Trim(fpText1(0).Text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText1(2).Text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText1(1).Text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText1(3).Text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText1(4).Text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText1(5).Text)) & "')"
     Set ConSql = vg_db.Execute("select * " & _
                  "From Sdx_Usuario " & _
                  "where loginusuario='" & LimpiaDato(Trim(fpText1(0).Text)) & "'", , adCmdText)
'     Set ConSql = vg_db.Execute("sod_i_usuario '" & xxLimpiaDato(Trim(fpText1(0).Text)) & "', '" & xxLimpiaDato(Trim(fpText1(2).Text)) & "', '" & xxLimpiaDato(Trim(fpText1(1).Text)) & "', '" & xxLimpiaDato(Trim(fpText1(3).Text)) & "', '" & xxLimpiaDato(Trim(fpText1(4).Text)) & "', '" & LimpiaDato(Trim(fpText1(5).Text)) & "'", , adCmdText)  'adCmdStoredProc)
     If Not ConSql.EOF Then
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 1
        vaSpread1.TypeHAlign = 1
        vaSpread1.Text = Trim(ConSql!loginusuario)
        vaSpread1.Col = 2
        vaSpread1.Text = Trim(ConSql!Nombre)
     End If
     ConSql.Close: Set ConSql = Nothing
     modo = "M"
   vg_db.CommitTrans
ElseIf modo = "M" Then
   vg_db.BeginTrans
     vg_db.Execute "update Sdx_Usuario set passwordusuario='" & LimpiaDato(Trim(fpText1(2).Text)) & "', nombre='" & LimpiaDato(Trim(fpText1(1).Text)) & "', telefono='" & LimpiaDato(Trim(fpText1(3).Text)) & "', oficina='" & LimpiaDato(Trim(fpText1(4).Text)) & "', departamento='" & LimpiaDato(Trim(fpText1(5).Text)) & "' WHERE loginusuario='" & LimpiaDato(Trim(fpText1(0).Text)) & "'"
     vaSpread1.Row = vaSpread1.ActiveRow
     vaSpread1.Col = 2: vaSpread1.Text = LimpiaDato(Trim(fpText1(1).Text))
   vg_db.CommitTrans
End If
modo = "M"
Ac_HabDes 3
Ac_Boton 2

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Private Sub ValidarCampos()
If ivalidar = 0 And fpText1(0).Text = "" Then ivalidar = 1: MsgBox "Debe Ingresar Login", vbExclamation + vbOKOnly, "Usuario": fpText1(0).SetFocus
If ivalidar = 0 And fpText1(1).Text = "" Then ivalidar = 1: MsgBox "Debe Ingresar Nombre", vbExclamation + vbOKOnly, "Usuario": fpText1(1).SetFocus
If ivalidar = 0 And fpText1(2).Text = "" Then ivalidar = 1: MsgBox "Debe Ingresar Pasword", vbExclamation + vbOKOnly, "Usuario": fpText1(2).SetFocus
If modo = "A" Then
   Set ConSql = vg_db.Execute("select * " & _
                "From Sdx_Usuario " & _
                "where loginusuario='" & LimpiaDato(Trim(fpText1(0).Text)) & "'", , adCmdText)
'   Set ConSql = vg_db.Execute("sod_s_usuario 2, '" & LimpiaDato(Trim(fpText1(0).Text)) & "', ''", , adCmdStoredProc)
   If Not ConSql.EOF Then ivalidar = 1: MsgBox "Usuario ya existe...", vbExclamation + vbOKOnly, "Mantención de usuarios": ConSql.Close: Set ConSql = Nothing: Exit Sub
End If
End Sub
Sub LimpiarVariable()
For i = 0 To 5
    vecdatos(i) = ""
    fpText1(i).Text = ""
    If modo = "M" And i = 0 Then
       fpText1(i).Enabled = False
    Else
       fpText1(i).Enabled = True
    End If
Next i
End Sub
Sub MoverDetalleDatos()
itexto = 1
LimpiarVariable
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: loginusuario = vaSpread1.Text
Set ConSql = vg_db.Execute("select * " & _
             "From Sdx_Usuario " & _
             "where loginusuario='" & loginusuario & "'", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_usuario 2, '" & loginusuario & "', ''", , adCmdStoredProc)
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
      fpText1(0).Text = Trim(ConSql!loginusuario)
      fpText1(1).Text = Trim(ConSql!Nombre)
      fpText1(2).Text = Trim(ConSql!passwordusuario)
      fpText1(3).Text = Trim(ConSql!telefono)
      fpText1(4).Text = Trim(ConSql!oficina)
      fpText1(5).Text = Trim(ConSql!departamento)
      ConSql.MoveNext
   Loop
Else
   Ac_HabDes 1
   Ac_Boton 3
End If
ConSql.Close: Set ConSql = Nothing
itexto = 0
End Sub
Sub MoverDatosGrillas()
vaSpread1.MaxRows = 0
Set ConSql = vg_db.Execute("select * " & _
             "From Sdx_Usuario " & _
             "order by nombre", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_usuario 1, '', ''", , adCmdStoredProc)
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = 1
      vaSpread1.Text = Trim(ConSql!loginusuario)
      vaSpread1.Col = 2
      vaSpread1.Text = Trim(ConSql!Nombre)
      ConSql.MoveNext
   Loop
   vaSpread1.Row = 1: vaSpread1.Col = 1: vaSpread1.EditMode = True
Else
   Ac_HabDes 1
   Ac_Boton 3
End If
ConSql.Close: Set ConSql = Nothing
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
End Sub
Sub MoverSistema()
fg_carga ""
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: loginusuario = vaSpread1.Text
vaSpread1.Col = 2: Label3.Caption = vaSpread1.Text
TvwDir.Nodes.Clear
Set nd = TvwDir.Nodes.Add(, , "R", "Sistema Propuesta")
nivel = 65: fin = 1: indindex = 1
Set ConSql = vg_db.Execute("SELECT Sdx_Programa.codprograma, Sdx_Programa.descripcion, Sdx_Programa.nomprograma, Sdx_Programa.opprog, Sdx_Programa.estado, Sdx_UsuCtrlAcceso.acceso, Sdx_UsuCtrlAcceso.incluir, Sdx_UsuCtrlAcceso.alterar, Sdx_UsuCtrlAcceso.eliminar, Sdx_UsuCtrlAcceso.imprimir, Sdx_Programa.codprog_anterior, Sdx_UsuCtrlAcceso.login " & _
             "FROM Sdx_UsuCtrlAcceso LEFT JOIN Sdx_Programa ON Sdx_UsuCtrlAcceso.programa = Sdx_Programa.codprograma " & _
             "Where (((Sdx_Programa.codprog_anterior) = 0) And ((Sdx_UsuCtrlAcceso.login) = '" & loginusuario & "')) " & _
             "ORDER BY Sdx_Programa.codprograma", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_usuctrlacceso 1, '" & loginusuario & "', 0, ''", , adCmdStoredProc)
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
      Set rootnode = TvwDir.Nodes.Add(nd, tvwChild, Chr(nivel) & ConSql!codprograma, Trim(ConSql!descripcion))
      indindex = indindex + 1
      If ConSql!acceso = "1" Or ConSql!acceso = "0" Then TvwDir.Nodes.Item(indindex).Checked = True: TvwDir.Nodes.Item(1).Checked = True
'      TvwDir.Nodes.Add nd, tvwChild, Chr(nivel) & ConSql!codprog, Trim(Consql1!descripcion)
'      Set rootnode = TvwDir.Nodes.Add(, , Chr(nivel) & ConSql!codprog, Trim(ConSql!descripcion))
         ' agregar un nodo hijo postizo, si fuera necesario
      If rootnode.Children = 0 And ConSql!estado = "1" Then
         hijo = Chr(nivel)
         codhijo = ConSql!codprograma
        ' la propiedad Texto de los nodos postizos es "***"
'           TvwDir.Nodes.Add rootnode.Index, tvwChild, , "*"
         BuscarHijos1 rootnode
      End If
      nivel = nivel + 1
      ConSql.MoveNext
   Loop
End If
ConSql.Close: Set ConSql = Nothing
fg_descarga
End Sub
Sub MoverRecetas()
fg_carga ""
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: loginusuario = vaSpread1.Text
vaSpread1.Col = 2: Label4.Caption = vaSpread1.Text
TvwDir2.Nodes.Clear
Set nd = TvwDir2.Nodes.Add(, , "R", "Control Recetario")
nivel = 65: fin = 1: indindex = 1
Set ConSql = vg_db.Execute("select Unit_Dfnd_No, Unit_Dfnd_Desc " & _
             "From Sdx_PB00074 " & _
             "order by Unit_Dfnd_No", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_usuctrlrecetas 1, '" & loginusuario & "', 0, ''", , adCmdStoredProc)
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
      Set rootnode = TvwDir2.Nodes.Add(nd, tvwChild, Chr(nivel) & ConSql!Unit_Dfnd_No, Trim(ConSql!Unit_Dfnd_Desc))
      indindex = indindex + 1
      Set Consql1 = vg_db.Execute("select estado " & _
                    "From Sdx_BloqueoRecetas " & _
                    "where cod_recetario=" & ConSql!Unit_Dfnd_No & " " & _
                    "and   loginusuario='" & loginusuario & "'", , adCmdText)
      If Not Consql1.EOF Then
         If Consql1!estado = "1" Then TvwDir2.Nodes.Item(indindex).Checked = True: TvwDir2.Nodes.Item(1).Checked = True
            ' agregar un nodo hijo postizo, si fuera necesario
            ' If rootnode.Children = 0 And ConSql!estado = "1" Then
      End If
      Consql1.Close: Set Consql1 = Nothing
      hijo = Chr(nivel)
      codhijo = ConSql!Unit_Dfnd_No
      ' la propiedad Texto de los nodos postizos es "***"
      BuscarHijosReceta rootnode
      nivel = nivel + 1
      ConSql.MoveNext
   Loop
End If
ConSql.Close: Set ConSql = Nothing
fg_descarga
End Sub
Function Ac_Boton(Boton As Integer)
Select Case Boton
  Case 1
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(8).Visible = True

    Toolbar1.Buttons(10).Visible = True
    Toolbar1.Buttons(11).Visible = False
    Toolbar1.Buttons(12).Visible = True
    Toolbar1.Buttons(13).Visible = False
    fpText.Enabled = False
  Case 2
    Toolbar1.Buttons(1).Visible = True: Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(3).Visible = True: Toolbar1.Buttons(4).Visible = False
    Toolbar1.Buttons(5).Visible = True: Toolbar1.Buttons(6).Visible = False
    
    Toolbar1.Buttons(7).Visible = True
    Toolbar1.Buttons(8).Visible = False

    Toolbar1.Buttons(10).Visible = False
    Toolbar1.Buttons(11).Visible = True
    Toolbar1.Buttons(12).Visible = False
    Toolbar1.Buttons(13).Visible = True
  Case 3
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(8).Visible = True

    Toolbar1.Buttons(10).Visible = False
    Toolbar1.Buttons(11).Visible = True
    Toolbar1.Buttons(12).Visible = False
    Toolbar1.Buttons(13).Visible = True

  Case 4
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(8).Visible = True

    Toolbar1.Buttons(10).Visible = False
    Toolbar1.Buttons(11).Visible = True
    Toolbar1.Buttons(12).Visible = False
    Toolbar1.Buttons(13).Visible = True

End Select
End Function
Function Ac_HabDes(Opcion As Integer)
Select Case Opcion
  Case 1
   fpText.Enabled = False
   SSTab1.TabEnabled(0) = True
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
   SSTab1.TabEnabled(3) = False
  Case 2
   SSTab1.TabEnabled(0) = False
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = False
   SSTab1.TabEnabled(3) = False
   SSTab1.Tab = 1
  Case 3
   fpText1(0).Enabled = False
   fpText.Enabled = True
   SSTab1.TabEnabled(0) = True
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
   SSTab1.TabEnabled(3) = True
  Case 4
   SSTab1.TabEnabled(0) = False
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = False
   SSTab1.TabEnabled(3) = False
End Select
End Function
Private Sub TvwDir_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim cKey As String, lKey As Integer, indkey As Long, indx As Long, indj As Long, indi As Long, lCheck As Boolean, lCheck1 As Boolean
Dim ckey2 As String, programa As String, incluir As String, alterar As String, eliminar As String, imprimir As String, acceso As String
Dim igraba As Integer
Dim indn As Long

On Error GoTo Man_Error

fg_carga ""

vg_db.BeginTrans

incluir = "0": alterar = "0": eliminar = "0": imprimir = "0": acceso = "0"
TvwDir.Nodes.Item(Node.Key).Selected = True
lCheck = TvwDir.Nodes.Item(TvwDir.SelectedItem.Index).Checked
lCheck1 = TvwDir.Nodes.Item(TvwDir.SelectedItem.Index).Checked
ckey2 = TvwDir.Nodes.Item(TvwDir.SelectedItem.Index).Key

If TvwDir.SelectedItem.Children > 0 Then
'   If lCheck = True Then
'      Resp_Habilitar ("Mantención")
'      If respuesta = vbYes Then fg_descarga: TvwDir.Nodes.Item(Node.index).Selected = True: TvwDir.Nodes.Item(Node.index).Checked = False: Exit Sub
'   ElseIf lCheck = False Then
'      Resp_Deshabilitar ("Mantención")
'      If respuesta = vbYes Then fg_descarga: TvwDir.Nodes.Item(Node.index).Selected = True: TvwDir.Nodes.Item(Node.index).Checked = True: Exit Sub
'   End If
   
   igraba = 1
   indi = TvwDir.SelectedItem.Child.Index
   indj = TvwDir.SelectedItem.Child.Index
   indn = TvwDir.SelectedItem.Child.LastSibling.Index
   For indx = 1 To 40
       While indj <> indn
         indj = TvwDir.Nodes(indj).Next.Index
       Wend
       If TvwDir.Nodes.Item(indj).Children > 0 Then
          indn = TvwDir.Nodes.Item(indn).Child.LastSibling.Index
          indj = indj + 1
       End If
   Next indx
   If Node.Index > 1 Then
'      indi = indi - 1
      programa = Mid(TvwDir.Nodes(Node.Index).Key, 2, 20)
      If lCheck1 = True Then
         acceso = 1: incluir = 1: alterar = 1: eliminar = 1: imprimir = 1
      ElseIf lCheck1 = False Then
         acceso = 0: incluir = 0: alterar = 0: eliminar = 0: imprimir = 0
      End If
      Set ConSql = vg_db.Execute("select * " & _
                   "from Sdx_UsuCtrlAcceso " & _
                   "where login='" & loginusuario & "' " & _
                   "and programa=" & programa & "", , adCmdText)
      If Not ConSql.EOF Then
         ConSql.Close: Set ConSql = Nothing
         If acceso = "0" Then
            vg_db.Execute "delete Sdx_UsuCtrlAcceso " & _
                          "from Sdx_UsuCtrlAcceso " & _
                          "where login='" & loginusuario & "' " & _
                          "and programa=" & programa & ""
         Else
            vg_db.Execute "Update Sdx_UsuCtrlAcceso " & _
                          "set acceso='" & acceso & "' " & _
                          "where login='" & loginusuario & "' " & _
                          "and   programa=" & programa & ""
          End If
      Else
         ConSql.Close: Set ConSql = Nothing
         Set ConSql = vg_db.Execute("select * " & _
                      "from Sdx_Programa " & _
                      "where codprograma=" & Val(programa) & " " & _
                      "and nomprograma<>''", , adCmdText)
         If Not ConSql.EOF Then
            If ConSql!opprog = "1" Then
               vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, " & _
                             "acceso, incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                             "values ('" & loginusuario & "', '" & programa & "', " & _
                             "'" & acceso & "', '" & incluir & "', '" & alterar & "', " & _
                             "'" & eliminar & "', '" & imprimir & "', '0')"
            ElseIf ConSql!opprog = "0" Then
               vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, acceso, " & _
                             "incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                             "values ('" & loginusuario & "', '" & programa & "', " & _
                             "'" & acceso & "', '0', '0', '0', '0', '0')"
            End If
         Else
            vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, acceso, " & _
                          "incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                          "values ('" & loginusuario & "', '" & programa & "', '0', " & _
                          "'0', '0', '0', '0', '0')"
         End If
         ConSql.Close: Set ConSql = Nothing
      End If

'      vg_db.Execute "sod_iud_usuctracceso 1, '" & loginusuario & "', '" & programa & "', " & _
'      "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'      "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
   End If
   For indx = indi To indj
     If indx > 1 Then
        TvwDir.Nodes.Item(indx).Checked = lCheck1
        If TvwDir.Nodes(indx).Text <> "Incluir" And TvwDir.Nodes(indx).Text <> "Alterar" And TvwDir.Nodes(indx).Text <> "Eliminar" And TvwDir.Nodes(indx).Text <> "Imprimir" Then
           programa = Mid(TvwDir.Nodes(indx).Key, 2, 20)
           If lCheck1 = True Then
              acceso = 1: incluir = 1: alterar = 1: eliminar = 1: imprimir = 1
           ElseIf lCheck1 = False Then
              acceso = 0: incluir = 0: alterar = 0: eliminar = 0: imprimir = 0
           End If
           Set ConSql = vg_db.Execute("select * " & _
                        "from Sdx_UsuCtrlAcceso " & _
                        "where login='" & loginusuario & "' " & _
                        "and programa=" & programa & "", , adCmdText)
           If Not ConSql.EOF Then
              ConSql.Close: Set ConSql = Nothing
              If acceso = "0" Then
                 vg_db.Execute "delete Sdx_UsuCtrlAcceso " & _
                               "from Sdx_UsuCtrlAcceso " & _
                               "where login='" & loginusuario & "' " & _
                               "and programa=" & programa & ""
              Else
                 vg_db.Execute "Update Sdx_UsuCtrlAcceso " & _
                               "set acceso='" & acceso & "' " & _
                               "where login='" & loginusuario & "' " & _
                               "and   programa=" & programa & ""
              End If
           Else
             ConSql.Close: Set ConSql = Nothing
              Set ConSql = vg_db.Execute("select * " & _
                           "from Sdx_Programa " & _
                           "where codprograma=" & Val(programa) & " " & _
                           "and nomprograma<>''", , adCmdText)
              If Not ConSql.EOF Then
                 If ConSql!opprog = "1" Then
                    vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, " & _
                                  "acceso, incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                                  "values ('" & loginusuario & "', '" & programa & "', " & _
                                  "'" & acceso & "', '" & incluir & "', '" & alterar & "', " & _
                                  "'" & eliminar & "', '" & imprimir & "', '0')"
                 ElseIf ConSql!opprog = "0" Then
                   vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, acceso, " & _
                                 "incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                                 "values ('" & loginusuario & "', '" & programa & "', " & _
                                 "'" & acceso & "', '0', '0', '0', '0', '0')"
                End If
             Else
                vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, acceso, " & _
                              "incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                              "values ('" & loginusuario & "', '" & programa & "', '0', " & _
                              "'0', '0', '0', '0', '0')"
             End If
             ConSql.Close: Set ConSql = Nothing
           End If
'           vg_db.Execute "sod_iud_usuctracceso 1, '" & loginusuario & "', '" & programa & "', " & _
'           "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'           "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
        End If
     End If
   Next indx
End If

If Node.Index > 1 Then
   cKey = TvwDir.Nodes(Node.Index).Parent.Key
   If TvwDir.SelectedItem.Children = 0 Then
      igraba = 0
      If TvwDir.Nodes(Node.Index).Text <> "Incluir" And TvwDir.Nodes(Node.Index).Text <> "Alterar" And TvwDir.Nodes(Node.Index).Text <> "Eliminar" And TvwDir.Nodes(Node.Index).Text <> "Imprimir" Then
         programa = Mid(TvwDir.Nodes(Node.Index).Key, 2, 20)
         If lCheck1 = True Then
            acceso = "1": incluir = "1": alterar = "1": eliminar = "1": imprimir = "1"
         ElseIf lCheck1 = False Then
            acceso = "0": incluir = "0": alterar = "0": eliminar = "0": imprimir = "0"
         End If
         Set ConSql = vg_db.Execute("select * " & _
                      "from Sdx_UsuCtrlAcceso " & _
                      "where login='" & loginusuario & "' " & _
                      "and programa=" & programa & "", , adCmdText)
         If Not ConSql.EOF Then
            ConSql.Close: Set ConSql = Nothing
            If acceso = "0" Then
               vg_db.Execute "delete Sdx_UsuCtrlAcceso " & _
                             "from Sdx_UsuCtrlAcceso " & _
                             "where login='" & loginusuario & "' " & _
                             "and programa=" & programa & ""
            Else
               vg_db.Execute "Update Sdx_UsuCtrlAcceso " & _
                             "set acceso='" & acceso & "' " & _
                             "where login='" & loginusuario & "' " & _
                             "and   programa=" & programa & ""
            End If
         Else
            ConSql.Close: Set ConSql = Nothing
            Set ConSql = vg_db.Execute("select * " & _
                         "from Sdx_Programa " & _
                         "where codprograma=" & Val(programa) & " " & _
                         "and nomprograma<>''", , adCmdText)
            If Not ConSql.EOF Then
               If ConSql!opprog = "1" Then
                  vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, " & _
                                "acceso, incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                                "values ('" & loginusuario & "', '" & programa & "', " & _
                                "'" & acceso & "', '" & incluir & "', '" & alterar & "', " & _
                                "'" & eliminar & "', '" & imprimir & "', '0')"
               ElseIf ConSql!opprog = "0" Then
                  vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, acceso, " & _
                                "incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                                "values ('" & loginusuario & "', '" & programa & "', " & _
                                "'" & acceso & "', '0', '0', '0', '0', '0')"
               End If
            Else
               vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, acceso, " & _
                             "incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                             "values ('" & loginusuario & "', '" & programa & "', '0', " & _
                             "'0', '0', '0', '0', '0')"
            End If
            ConSql.Close: Set ConSql = Nothing
         End If
'         vg_db.Execute "sod_iud_usuctracceso 1, '" & loginusuario & "', '" & programa & "', " & _
'         "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'         "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
      ElseIf TvwDir.Nodes(Node.Index).Text = "Incluir" Then
         If lCheck1 = True Then
            incluir = "1"
         ElseIf lCheck1 = False Then
            incluir = "0"
         End If
         programa = Mid(TvwDir.Nodes(Node.Index).Parent.Key, 2, 20)
         vg_db.Execute "Update Sdx_UsuCtrlAcceso " & _
                       "set incluir='" & incluir & "'" & _
                       "where login='" & loginusuario & "' " & _
                       "and   programa=" & programa & ""
'         vg_db.Execute "sod_iud_usuctracceso 2, '" & loginusuario & "', '" & programa & "', " & _
'         "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'         "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
      ElseIf TvwDir.Nodes(Node.Index).Text = "Alterar" Then
         If lCheck1 = True Then
            alterar = "1"
         ElseIf lCheck1 = False Then
            alterar = "0"
         End If
         programa = Mid(TvwDir.Nodes(Node.Index).Parent.Key, 2, 20)
         Set ConSql = vg_db.Execute("select * " & _
                      "from Sdx_UsuCtrlAcceso " & _
                      "where login='" & loginusuario & "' " & _
                      "and programa=" & programa & "", , adCmdText)
         If Not ConSql.EOF Then
            vg_db.Execute "Update Sdx_UsuCtrlAcceso " & _
                          "set alterar='" & alterar & "' " & _
                          "where login='" & loginusuario & "' " & _
                          "and   programa=" & programa & ""
         Else
            vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, " & _
                          "acceso, incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                          "values ('" & loginusuario & "', '" & programa & "', " & _
                          "'0', '0', '" & alterar & "', '0', '0', '0')"
         End If
'         vg_db.Execute "sod_iud_usuctracceso 3, '" & loginusuario & "', '" & programa & "', " & _
'         "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'         "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
      ElseIf TvwDir.Nodes(Node.Index).Text = "Eliminar" Then
         If lCheck1 = True Then
            eliminar = "1"
         ElseIf lCheck1 = False Then
            eliminar = "0"
         End If
         programa = Mid(TvwDir.Nodes(Node.Index).Parent.Key, 2, 20)
         vg_db.Execute "Update Sdx_UsuCtrlAcceso " & _
                       "set eliminar='" & eliminar & "' " & _
                       "where login='" & loginusuario & "' " & _
                       "and   programa=" & programa & ""
'         vg_db.Execute "sod_iud_usuctracceso 4, '" & loginusuario & "', '" & programa & "', " & _
'         "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'         "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
      ElseIf TvwDir.Nodes(Node.Index).Text = "Imprimir" Then
         If lCheck1 = True Then
            imprimir = "1"
         ElseIf lCheck1 = False Then
            imprimir = "0"
         End If
         programa = Mid(TvwDir.Nodes(Node.Index).Parent.Key, 2, 20)
         vg_db.Execute "Update Sdx_UsuCtrlAcceso " & _
                       "set imprimir='" & imprimir & "' " & _
                       "where login='" & loginusuario & "' " & _
                       "and   programa=" & programa & ""
'         vg_db.Execute "sod_iud_usuctracceso 5, '" & loginusuario & "', '" & programa & "', " & _
'         "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'         "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
      End If
      indn = TvwDir.SelectedItem.FirstSibling.Index: indi = TvwDir.SelectedItem.LastSibling.Index
      For indj = indn To indi
          If TvwDir.SelectedItem.Children = 0 And TvwDir.Nodes.Item(indj).Checked = True Then
             lCheck = True: lCheck1 = True
          End If
      Next indj
      For indj = indn To TvwDir.Nodes.Count
          If Mid(ckey2, 1, 1) = Mid(TvwDir.Nodes(indj).Parent.Key, 1, 1) Then
             indi = indj
          End If
      Next indj
      For indj = 1 To indi
          If TvwDir.SelectedItem.Children = 0 And TvwDir.Nodes.Item(indj).Checked = True Then
             lCheck = True: lCheck1 = True
          End If
      Next indj
   Else
      indn = TvwDir.SelectedItem.FirstSibling.Index
      indi = TvwDir.SelectedItem.LastSibling.Index
      For indj = indn To indi
          If TvwDir.SelectedItem.Children > 0 And TvwDir.Nodes.Item(indj).Checked = True Then
             lCheck = True
          End If
      Next indj
      If cKey <> "R" And TvwDir.Nodes(Node.Index - 1).Children > 1 Then
'         TvwDir.Nodes.Item(Node.index - 1).Checked = lCheck
         TvwDir.Nodes.Item(Node.Index - 1).Checked = lCheck
         
         programa = Mid(TvwDir.Nodes(Node.Index - 1).Key, 2, 20)
         If lCheck = True Then
            acceso = 1: incluir = 1: alterar = 1: eliminar = 1: imprimir = 1
         ElseIf lCheck = False Then
            acceso = 0: incluir = 0: alterar = 0: eliminar = 0: imprimir = 0
         End If
         Set ConSql = vg_db.Execute("select * " & _
                      "from Sdx_UsuCtrlAcceso " & _
                      "where login='" & loginusuario & "' " & _
                      "and programa=" & programa & "", , adCmdText)
         If Not ConSql.EOF Then
            ConSql.Close: Set ConSql = Nothing
            If acceso = "0" Then
               vg_db.Execute "delete Sdx_UsuCtrlAcceso " & _
                             "from Sdx_UsuCtrlAcceso " & _
                             "where login='" & loginusuario & "' " & _
                             "and programa=" & programa & ""
            Else
               vg_db.Execute "Update Sdx_UsuCtrlAcceso " & _
                             "set acceso='" & acceso & "' " & _
                             "where login='" & loginusuario & "' " & _
                             "and   programa=" & programa & ""
            End If
         Else
            ConSql.Close: Set ConSql = Nothing
            Set ConSql = vg_db.Execute("select * " & _
                         "from Sdx_Programa " & _
                         "where codprog=" & Val(programa) & " " & _
                         "and nomprog<>''", , adCmdText)
            If Not ConSql.EOF Then
               If ConSql!opprog = "1" Then
                  vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, " & _
                                "acceso, incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                                "values ('" & loginusuario & "', '" & programa & "', " & _
                                "'" & acceso & "', '" & incluir & "', '" & alterar & "', " & _
                                "'" & eliminar & "', '" & imprimir & "', '0')"
               ElseIf ConSql!opprog = "0" Then
                  vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, acceso, " & _
                                "incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                                "values ('" & loginusuario & "', '" & programa & "', " & _
                                "'" & acceso & "', '0', '0', '0', '0', '0')"
               End If
            Else
               vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, acceso, " & _
                             "incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                             "values ('" & loginusuario & "', '" & programa & "', '0', " & _
                             "'0', '0', '0', '0', '0')"
            End If
            ConSql.Close: Set ConSql = Nothing
         End If
'         vg_db.Execute "sod_iud_usuctracceso 1, '" & loginusuario & "', '" & programa & "', " & _
'         "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'         "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
         
'         cKey = TvwDir.Nodes(Node.index - 1).Parent.Key
         cKey = TvwDir.Nodes(Node.Index).Parent.Key
       End If
      
      indn = indi
      For indj = indn To TvwDir.Nodes.Count
          If TvwDir.SelectedItem.Children > 0 And TvwDir.Nodes.Item(indj).Checked = True Then
             lCheck = True
          End If
      Next indj
   
   End If
   
'   TvwDir.Nodes.Item(Node.Key).Selected = True
'   lCheck = TvwDir.Nodes.Item(TvwDir.SelectedItem.Index).Checked
   indkey = 1
   For indi = 1 To TvwDir.Nodes.Count
       If TvwDir.Nodes.Item(indi).Key = cKey Then
          TvwDir.Nodes.Item(indi).Checked = lCheck
          If cKey = "R" Then Exit For
         
          programa = Mid(TvwDir.Nodes(indi).Key, 2, 20)
          If lCheck = True Then
             acceso = 1 ': incluir = 1: alterar = 1: eliminar = 1: imprimir = 1
          ElseIf lCheck = False Then
             acceso = 0: incluir = 0: alterar = 0: eliminar = 0: imprimir = 0
          End If
           Set ConSql = vg_db.Execute("select * " & _
                       "from Sdx_UsuCtrlAcceso " & _
                       "where login='" & loginusuario & "' " & _
                       "and programa=" & programa & "", , adCmdText)
          If Not ConSql.EOF Then
             ConSql.Close: Set ConSql = Nothing
             If acceso = "0" Then
                vg_db.Execute "delete Sdx_UsuCtrlAcceso " & _
                              "from Sdx_UsuCtrlAcceso " & _
                              "where login='" & loginusuario & "' " & _
                              "and programa=" & programa & ""
             Else
                vg_db.Execute "Update Sdx_UsuCtrlAcceso " & _
                              "set acceso='" & acceso & "' " & _
                              "where login='" & loginusuario & "' " & _
                              "and   programa=" & programa & ""
             End If
          Else
             ConSql.Close: Set ConSql = Nothing
             Set ConSql = vg_db.Execute("select * " & _
                          "from Sdx_Programa " & _
                          "where codprograma=" & Val(programa) & " " & _
                          "and nomprograma<>''", , adCmdText)
             If Not ConSql.EOF Then
                If ConSql!opprog = "1" Then
                   vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, " & _
                                 "acceso, incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                                 "values ('" & loginusuario & "', '" & programa & "', " & _
                                 "'" & acceso & "', '" & incluir & "', '" & alterar & "', " & _
                                 "'" & eliminar & "', '" & imprimir & "', '0')"
                ElseIf ConSql!opprog = "0" Then
                   vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, acceso, " & _
                                 "incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                                 "values ('" & loginusuario & "', '" & programa & "', " & _
                                 "'" & acceso & "', '0', '0', '0', '0', '0')"
                End If
             Else
                vg_db.Execute "insert into Sdx_UsuCtrlAcceso (login, programa, acceso, " & _
                              "incluir, alterar, eliminar, imprimir, otrosacceso) " & _
                              "values ('" & loginusuario & "', '" & programa & "', '0', " & _
                              "'0', '0', '0', '0', '0')"
             End If
             ConSql.Close: Set ConSql = Nothing
          End If
'          vg_db.Execute "sod_iud_usuctracceso 1, '" & loginusuario & "', '" & programa & "', " & _
'          "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'          "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
          
          Set Node = TvwDir.Nodes.Item(indi)
          cKey = TvwDir.Nodes(Node.Index).Parent.Key
          indi = 0
       End If
   Next indi
End If

vg_db.CommitTrans
fg_descarga

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Sub BuscarHijos1(nd As Node)
Set Consql1 = vg_db.Execute("SELECT Sdx_Programa.codprograma, Sdx_Programa.descripcion, Sdx_Programa.nomprograma, Sdx_Programa.opprog, Sdx_Programa.estado, Sdx_UsuCtrlAcceso.acceso, Sdx_UsuCtrlAcceso.incluir, Sdx_UsuCtrlAcceso.alterar, Sdx_UsuCtrlAcceso.eliminar, Sdx_UsuCtrlAcceso.imprimir, Sdx_Programa.codprog_anterior, Sdx_UsuCtrlAcceso.login " & _
              "FROM Sdx_UsuCtrlAcceso LEFT JOIN Sdx_Programa ON Sdx_UsuCtrlAcceso.programa = Sdx_Programa.codprograma " & _
              "Where (((Sdx_Programa.codprog_anterior) = " & codhijo & ") And ((Sdx_UsuCtrlAcceso.login) = '" & loginusuario & "')) " & _
              "ORDER BY Sdx_Programa.codprograma", , adCmdText)
'Set Consql1 = vg_db.Execute("sod_s_usuctrlacceso 2, '" & loginusuario & "', " & codhijo & ", ''", , adCmdStoredProc)
If Not Consql1.EOF Then
   Do While Not Consql1.EOF
'      Set hijo = TvwDir.Nodes.Add("Planificación", tvwChild, , Trim(ConSql1!descripcion))
      TvwDir.Nodes.Add hijo & codhijo, tvwChild, hijo & Consql1!codprograma, Trim(Consql1!descripcion)
      indindex = indindex + 1
      If Consql1!acceso = "1" Or Consql1!acceso = "0" Then TvwDir.Nodes.Item(indindex).Checked = True
      If Consql1!opprog = "1" Then
         TvwDir.Nodes.Add hijo & Consql1!codprograma, tvwChild, hijo & Consql1!codprograma + 99, "Incluir"
         indindex = indindex + 1
         If Consql1!incluir = "1" Then TvwDir.Nodes.Item(indindex).Checked = True
         TvwDir.Nodes.Add hijo & Consql1!codprograma, tvwChild, hijo & Consql1!codprograma + 999, "Alterar"
         indindex = indindex + 1
         If Consql1!alterar = "1" Then TvwDir.Nodes.Item(indindex).Checked = True
         TvwDir.Nodes.Add hijo & Consql1!codprograma, tvwChild, hijo & Consql1!codprograma + 9999, "Eliminar"
         indindex = indindex + 1
         If Consql1!eliminar = "1" Then TvwDir.Nodes.Item(indindex).Checked = True
         TvwDir.Nodes.Add hijo & Consql1!codprograma, tvwChild, hijo & Consql1!codprograma + 99999, "Imprimir"
         indindex = indindex + 1
         If Consql1!imprimir = "1" Then TvwDir.Nodes.Item(indindex).Checked = True
      End If
      If Consql1!estado = 1 Then
         ' la propiedad Texto de los nodos positivos es "***"
         codhijo2 = Consql1!codprograma
         BuscarHijos2 nd, Consql1!codprograma
'         TvwDir.Nodes.Add nd.Index, tvwChild, , "*"
      End If
      Consql1.MoveNext
   Loop
End If
Consql1.Close: Set Consql1 = Nothing
End Sub
Sub BuscarHijos2(nd2 As Node, codprog As Long)
Set Consql2 = vg_db.Execute("SELECT Sdx_Programa.codprograma, Sdx_Programa.descripcion, Sdx_Programa.nomprograma, Sdx_Programa.opprog, Sdx_Programa.estado, Sdx_UsuCtrlAcceso.acceso, Sdx_UsuCtrlAcceso.incluir, Sdx_UsuCtrlAcceso.alterar, Sdx_UsuCtrlAcceso.eliminar, Sdx_UsuCtrlAcceso.imprimir, Sdx_Programa.codprog_anterior, Sdx_UsuCtrlAcceso.login " & _
              "FROM Sdx_UsuCtrlAcceso LEFT JOIN Sdx_Programa ON Sdx_UsuCtrlAcceso.programa = Sdx_Programa.codprograma " & _
              "Where (((Sdx_Programa.codprog_anterior) = " & codhijo2 & ") And ((Sdx_UsuCtrlAcceso.login) = '" & loginusuario & "')) " & _
              "ORDER BY Sdx_Programa.codprograma", , adCmdText)
'Set Consql2 = vg_db.Execute("sod_s_usuctrlacceso 2, '" & loginusuario & "', " & codhijo2 & ", ''", , adCmdStoredProc)
If Not Consql2.EOF Then
   Do While Not Consql2.EOF
      TvwDir.Nodes.Add hijo & codhijo2, tvwChild, hijo & Consql2!codprograma, Trim(Consql2!descripcion)
      indindex = indindex + 1
      If Consql2!acceso = 1 Then TvwDir.Nodes.Item(indindex).Checked = True
      If Consql2!opprog = 1 Then
         TvwDir.Nodes.Add hijo & Consql2!codprograma, tvwChild, hijo & Consql2!codprograma + 99, "Incluir"
         indindex = indindex + 1
         If Consql2!incluir = "1" Then TvwDir.Nodes.Item(indindex).Checked = True
         TvwDir.Nodes.Add hijo & Consql2!codprograma, tvwChild, hijo & Consql2!codprograma + 999, "Alterar"
         indindex = indindex + 1
         If Consql2!alterar = "1" Then TvwDir.Nodes.Item(indindex).Checked = True
         TvwDir.Nodes.Add hijo & Consql2!codprograma, tvwChild, hijo & Consql2!codprograma + 9999, "Eliminar"
         indindex = indindex + 1
         If Consql2!eliminar = "1" Then TvwDir.Nodes.Item(indindex).Checked = True
         TvwDir.Nodes.Add hijo & Consql2!codprograma, tvwChild, hijo & Consql2!codprograma + 99999, "Imprimir"
         indindex = indindex + 1
         If Consql2!imprimir = "1" Then TvwDir.Nodes.Item(indindex).Checked = True
      End If
      If Consql2!estado = 1 Then
      End If
      Consql2.MoveNext
   Loop
End If
Consql2.Close: Set Consql2 = Nothing
End Sub
Sub BuscarHijosReceta(nd As Node)
Set Consql1 = vg_db.Execute("select PB00074.Unit_Dfnd_No, PB00074.Unit_Dfnd_Desc, " & _
              "PB00074.Prev_Unit_Dfnd_No " & _
              "From PB00074, Sdx_PB00074 " & _
              "Where Sdx_PB00074.Unit_Dfnd_No = PB00074.Prev_Unit_Dfnd_No " & _
              "and   Sdx_PB00074.Unit_Dfnd_No=" & codhijo & " " & _
              "order by PB00074.Unit_Dfnd_No", , adCmdText)
'Set Consql1 = vg_db.Execute("sod_s_usuctrlrecetas 2, '" & loginusuario & "', " & codhijo & ", ''", , adCmdStoredProc)
If Not Consql1.EOF Then
   Do While Not Consql1.EOF
'      Set hijo = TvwDir.Nodes.Add("Planificación", tvwChild, , Trim(ConSql1!descripcion))
      TvwDir2.Nodes.Add hijo & codhijo, tvwChild, hijo & hijo & Consql1!Unit_Dfnd_No, Trim(Consql1!Unit_Dfnd_Desc)
      indindex = indindex + 1
      Set Consql2 = vg_db.Execute("select estado " & _
                    "From Sdx_BloqueoRecetas " & _
                    "where cod_recetario=" & Consql1!Prev_Unit_Dfnd_No & " " & _
                    "and   cod_subrecetario=" & Consql1!Unit_Dfnd_No & " " & _
                    "and   loginusuario='" & loginusuario & "'", , adCmdText)
      If Not Consql2.EOF Then
         If Consql2!estado = "1" Then TvwDir2.Nodes.Item(indindex).Checked = True
      End If
      Consql2.Close: Set Consql2 = Nothing
      TvwDir2.Nodes.Add hijo & hijo & Consql1!Unit_Dfnd_No, tvwChild, hijo & Consql1!Unit_Dfnd_No + 99, "Bloquear"
      Consql1.MoveNext
   Loop
End If
Consql1.Close: Set Consql1 = Nothing
End Sub
Private Sub TvwDir2_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim cKey As String, lKey As Integer, indkey As Long, indx As Long, indj As Long, indi As Long, lCheck As Boolean, lCheck1 As Boolean
Dim ckey2 As String, programa As String, estado As String
Dim igraba As Integer
Dim indn As Long

fg_carga ""

TvwDir2.Nodes.Item(Node.Key).Selected = True
lCheck = TvwDir2.Nodes.Item(TvwDir2.SelectedItem.Index).Checked
lCheck1 = TvwDir2.Nodes.Item(TvwDir2.SelectedItem.Index).Checked
ckey2 = TvwDir2.Nodes.Item(TvwDir2.SelectedItem.Index).Key

If TvwDir2.SelectedItem.Children > 0 Then
   igraba = 1
   indi = TvwDir2.SelectedItem.Child.Index
   indj = TvwDir2.SelectedItem.Child.Index
   indn = TvwDir2.SelectedItem.Child.LastSibling.Index
   For indx = 1 To 40
       While indj <> indn
         indj = TvwDir2.Nodes(indj).Next.Index
       Wend
       If TvwDir2.Nodes.Item(indj).Children > 0 Then
          indn = TvwDir2.Nodes.Item(indn).Child.LastSibling.Index
          indj = indj + 1
       End If
   Next indx
   If Node.Index > 1 Then
      programa = Mid(TvwDir2.Nodes(Node.Index).Key, 2, 20)
      If lCheck1 = True Then
         estado = 1
      ElseIf lCheck1 = False Then
         estado = 0
      End If
'      vg_db.Execute "sod_iud_usuctracceso 1, '" & loginusuario & "', '" & programa & "', " & _
'      "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'      "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
   End If
   For indx = indi To indj
     If indx > 1 Then
        TvwDir2.Nodes.Item(indx).Checked = lCheck1
        If TvwDir2.Nodes(indx).Text <> "Bloquear" Then
           programa = Mid(TvwDir2.Nodes(indx).Key, 2, 20)
           If lCheck1 = True Then
              estado = 1
           ElseIf lCheck1 = False Then
              estado = 0
           End If
'           vg_db.Execute "sod_iud_usuctracceso 1, '" & loginusuario & "', '" & programa & "', " & _
'           "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'           "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
        End If
     End If
   Next indx
End If

If Node.Index > 1 Then
   cKey = TvwDir2.Nodes(Node.Index).Parent.Key
   If TvwDir2.SelectedItem.Children = 0 Then
      igraba = 0
      If TvwDir2.Nodes(Node.Index).Text <> "Bloquear" Then
         programa = Mid(TvwDir2.Nodes(Node.Index).Key, 2, 20)
         If lCheck1 = True Then
            estado = "1"
         ElseIf lCheck1 = False Then
            estado = "0"
         End If
'         vg_db.Execute "sod_iud_usuctracceso 1, '" & loginusuario & "', '" & programa & "', " & _
'         "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'         "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
      ElseIf TvwDir2.Nodes(Node.Index).Text = "Incluir" Then
         If lCheck1 = True Then
            incluir = "1"
         ElseIf lCheck1 = False Then
            incluir = "0"
         End If
         programa = Mid(TvwDir2.Nodes(Node.Index).Parent.Key, 2, 20)
'         vg_db.Execute "sod_iud_usuctracceso 2, '" & loginusuario & "', '" & programa & "', " & _
'         "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'         "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
      ElseIf TvwDir2.Nodes(Node.Index).Text = "Alterar" Then
         If lCheck1 = True Then
            alterar = "1"
         ElseIf lCheck1 = False Then
            alterar = "0"
         End If
         programa = Mid(TvwDir2.Nodes(Node.Index).Parent.Key, 2, 20)
'         vg_db.Execute "sod_iud_usuctracceso 3, '" & loginusuario & "', '" & programa & "', " & _
'         "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'         "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
      ElseIf TvwDir2.Nodes(Node.Index).Text = "Eliminar" Then
         If lCheck1 = True Then
            eliminar = "1"
         ElseIf lCheck1 = False Then
            eliminar = "0"
         End If
         programa = Mid(TvwDir2.Nodes(Node.Index).Parent.Key, 2, 20)
'         vg_db.Execute "sod_iud_usuctracceso 4, '" & loginusuario & "', '" & programa & "', " & _
'         "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'         "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
      ElseIf TvwDir2.Nodes(Node.Index).Text = "Imprimir" Then
         If lCheck1 = True Then
            imprimir = "1"
         ElseIf lCheck1 = False Then
            imprimir = "0"
         End If
         programa = Mid(TvwDir2.Nodes(Node.Index).Parent.Key, 2, 20)
'         vg_db.Execute "sod_iud_usuctracceso 5, '" & loginusuario & "', '" & programa & "', " & _
'         "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'         "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
      End If
      indn = TvwDir2.SelectedItem.FirstSibling.Index: indi = TvwDir2.SelectedItem.LastSibling.Index
      For indj = indn To indi
          If TvwDir2.SelectedItem.Children = 0 And TvwDir2.Nodes.Item(indj).Checked = True Then
             lCheck = True: lCheck1 = True
          End If
      Next indj
      For indj = indn To TvwDir.Nodes.Count
          If Mid(ckey2, 1, 1) = Mid(TvwDir2.Nodes(indj).Parent.Key, 1, 1) Then
             indi = indj
          End If
      Next indj
      For indj = 1 To indi
          If TvwDir2.SelectedItem.Children = 0 And TvwDir2.Nodes.Item(indj).Checked = True Then
             lCheck = True: lCheck1 = True
          End If
      Next indj
   Else
      indn = TvwDir2.SelectedItem.FirstSibling.Index
      indi = TvwDir2.SelectedItem.LastSibling.Index
      For indj = indn To indi
          If TvwDir2.SelectedItem.Children > 0 And TvwDir2.Nodes.Item(indj).Checked = True Then
             lCheck = True
          End If
      Next indj
      If cKey <> "R" And TvwDir2.Nodes(Node.Index - 1).Children > 1 Then
         TvwDir2.Nodes.Item(Node.Index - 1).Checked = lCheck
         
         programa = Mid(TvwDir2.Nodes(Node.Index - 1).Key, 2, 20)
         If lCheck = True Then
            acceso = 1: incluir = 1: alterar = 1: eliminar = 1: imprimir = 1
         ElseIf lCheck = False Then
            acceso = 0: incluir = 0: alterar = 0: eliminar = 0: imprimir = 0
         End If
'         vg_db.Execute "sod_iud_usuctracceso 1, '" & loginusuario & "', '" & programa & "', " & _
'         "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'         "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
         cKey = TvwDir2.Nodes(Node.Index).Parent.Key
       End If
      
      indn = indi
      For indj = indn To TvwDir2.Nodes.Count
          If TvwDir2.SelectedItem.Children > 0 And TvwDir2.Nodes.Item(indj).Checked = True Then
             lCheck = True
          End If
      Next indj
   
   End If
   indkey = 1
   For indi = 1 To TvwDir2.Nodes.Count
       If TvwDir2.Nodes.Item(indi).Key = cKey Then
          TvwDir2.Nodes.Item(indi).Checked = lCheck
          If cKey = "R" Then Exit For
         
          programa = Mid(TvwDir2.Nodes(indi).Key, 2, 20)
          If lCheck = True Then
             acceso = 1: incluir = 1: alterar = 1: eliminar = 1: imprimir = 1
          ElseIf lCheck = False Then
             acceso = 0: incluir = 0: alterar = 0: eliminar = 0: imprimir = 0
          End If
'          vg_db.Execute "sod_iud_usuctracceso 1, '" & loginusuario & "', '" & programa & "', " & _
'          "'" & acceso & "', '" & incluir & "', '" & alterar & "', '" & eliminar & "', " & _
'          "'" & imprimir & "', '" & "0" & "', " & Val(programa) & ""
          
          Set Node = TvwDir2.Nodes.Item(indi)
          cKey = TvwDir2.Nodes(Node.Index).Parent.Key
          indi = 0
       End If
   Next indi
End If

fg_descarga
End Sub
