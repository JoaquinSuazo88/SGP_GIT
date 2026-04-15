VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form M_Receta 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receta"
   ClientHeight    =   7590
   ClientLeft      =   1845
   ClientTop       =   1830
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":094E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":0C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":0F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":129C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":15B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":18D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":1BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":1F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":221E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":2538
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":2852
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":2B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":2E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":31A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":34BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":37D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":3AF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":3E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":4124
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":443E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":4758
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":4A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":4D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":50A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":53C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":56DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":59F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":5D0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      TabMaxWidth     =   4
      OLEDropMode     =   1
      TabCaption(0)   =   "Receta"
      TabPicture(0)   =   "M_Receta.frx":6028
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1(1)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle Receta Patrón"
      TabPicture(1)   =   "M_Receta.frx":6044
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame11(0)"
      Tab(1).Control(1)=   "Frame1(0)"
      Tab(1).Control(2)=   "Frame1(3)"
      Tab(1).Control(3)=   "Frame3(0)"
      Tab(1).Control(4)=   "Frame1(2)"
      Tab(1).Control(5)=   "ImageList4"
      Tab(1).Control(6)=   "Toolbar2"
      Tab(1).Control(7)=   "vaSpread1(1)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Detalle Receta Local"
      TabPicture(2)   =   "M_Receta.frx":6060
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame11(1)"
      Tab(2).Control(1)=   "Frame1(6)"
      Tab(2).Control(2)=   "Frame3(1)"
      Tab(2).Control(3)=   "Frame1(5)"
      Tab(2).Control(4)=   "Frame1(4)"
      Tab(2).Control(5)=   "Toolbar4"
      Tab(2).Control(6)=   "vaSpread1(2)"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Detalle Receta x Regimen"
      TabPicture(3)   =   "M_Receta.frx":607C
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "vaSpread1(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Toolbar5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame1(7)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame3(2)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame1(8)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Frame1(9)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Frame6"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Frame11(2)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "Metodos Preparación"
      TabPicture(4)   =   "M_Receta.frx":6098
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).Control(1)=   "Frame5(0)"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Grupo Vulnerable"
      TabPicture(5)   =   "M_Receta.frx":60B4
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame5(1)"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame11 
         Caption         =   "Alergeno"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   2
         Left            =   6960
         TabIndex        =   131
         Top             =   480
         Width           =   2175
         Begin VB.ListBox Alergeno 
            Height          =   1410
            Index           =   2
            ItemData        =   "M_Receta.frx":60D0
            Left            =   120
            List            =   "M_Receta.frx":60D7
            Style           =   1  'Checkbox
            TabIndex        =   132
            Top             =   360
            Width           =   1950
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Alergeno"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   1
         Left            =   -68040
         TabIndex        =   129
         Top             =   360
         Width           =   2175
         Begin VB.ListBox Alergeno 
            Height          =   1410
            Index           =   1
            ItemData        =   "M_Receta.frx":60E5
            Left            =   120
            List            =   "M_Receta.frx":60EC
            Style           =   1  'Checkbox
            TabIndex        =   130
            Top             =   360
            Width           =   1950
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Alergeno"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   0
         Left            =   -67920
         TabIndex        =   127
         Top             =   360
         Width           =   2175
         Begin VB.ListBox Alergeno 
            Height          =   1410
            Index           =   0
            ItemData        =   "M_Receta.frx":60FA
            Left            =   120
            List            =   "M_Receta.frx":6101
            Style           =   1  'Checkbox
            TabIndex        =   128
            Top             =   360
            Width           =   1950
         End
      End
      Begin VB.Frame Frame5 
         Height          =   5175
         Index           =   1
         Left            =   -74520
         TabIndex        =   124
         Top             =   1515
         Width           =   10125
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   4365
            Index           =   1
            Left            =   180
            TabIndex        =   125
            Top             =   330
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   7699
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"M_Receta.frx":610F
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
            Index           =   1
            Left            =   210
            TabIndex        =   126
            Top             =   4830
            Width           =   585
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   60
         TabIndex        =   119
         Top             =   510
         Width           =   6855
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   1
            Left            =   1230
            TabIndex        =   120
            Top             =   270
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
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   2640
            TabIndex        =   122
            Top             =   270
            Width           =   4065
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   6
            Left            =   2130
            Picture         =   "M_Receta.frx":6191
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Regimen"
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
            Index           =   26
            Left            =   360
            TabIndex        =   121
            Top             =   330
            Width           =   750
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   7
            Left            =   2655
            TabIndex        =   123
            Top             =   285
            Width           =   4095
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   9
         Left            =   9330
         TabIndex        =   108
         Top             =   1200
         Width           =   2885
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   8
            Left            =   1605
            TabIndex        =   109
            Top             =   1155
            Width           =   855
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   1323
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
            MaxValue        =   "900000"
            MinValue        =   "-900000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   9
            Left            =   1605
            TabIndex        =   110
            Top             =   810
            Width           =   855
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   1323
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
            BackColor       =   16777215
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
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   10
            Left            =   1575
            TabIndex        =   111
            Top             =   150
            Width           =   855
            _Version        =   196608
            _ExtentX        =   1508
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
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
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
            Text            =   "1"
            DecimalPlaces   =   0
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   11
            Left            =   1590
            TabIndex        =   112
            Top             =   465
            Width           =   855
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   1323
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
            BackColor       =   16777215
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
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
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
            Caption         =   "C. Bruta"
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
            Index           =   27
            Left            =   330
            TabIndex        =   116
            Top             =   570
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "G. Neto"
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
            Index           =   26
            Left            =   330
            TabIndex        =   115
            Top             =   1260
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "C. Servida"
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
            Index           =   25
            Left            =   330
            TabIndex        =   114
            Top             =   900
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Raciones"
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
            Index           =   24
            Left            =   330
            TabIndex        =   113
            Top             =   255
            Width           =   810
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1545
         Index           =   8
         Left            =   90
         TabIndex        =   97
         Top             =   1200
         Width           =   6825
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   4
            Left            =   2040
            TabIndex        =   98
            Top             =   150
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
            MaxLength       =   80
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   5
            Left            =   2040
            TabIndex        =   99
            Top             =   495
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
            MaxLength       =   80
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
            Caption         =   "Nombre Receta"
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
            Index           =   23
            Left            =   300
            TabIndex        =   105
            Top             =   255
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Fantasia"
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
            Index           =   22
            Left            =   300
            TabIndex        =   104
            Top             =   600
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Categoria Dietetica"
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
            Index           =   21
            Left            =   300
            TabIndex        =   103
            Top             =   945
            Width           =   1650
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Plato"
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
            Index           =   20
            Left            =   300
            TabIndex        =   102
            Top             =   1230
            Width           =   885
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   5
            Left            =   1920
            Picture         =   "M_Receta.frx":649B
            Top             =   750
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   1920
            Picture         =   "M_Receta.frx":67A5
            Top             =   1065
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   2370
            TabIndex        =   101
            Top             =   1170
            Width           =   4095
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   2370
            TabIndex        =   100
            Top             =   840
            Width           =   4095
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   2415
            TabIndex        =   107
            Top             =   885
            Width           =   4095
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2415
            TabIndex        =   106
            Top             =   1215
            Width           =   4095
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Aporte Nutricionales"
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
         Height          =   4110
         Index           =   2
         Left            =   9330
         TabIndex        =   95
         Top             =   2775
         Width           =   2885
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   3420
            Index           =   2
            Left            =   165
            TabIndex        =   96
            Top             =   360
            Width           =   2550
            _Version        =   393216
            _ExtentX        =   4498
            _ExtentY        =   6033
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
            SpreadDesigner  =   "M_Receta.frx":6AAF
            ScrollBarTrack  =   3
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   7
         Left            =   210
         TabIndex        =   86
         Top             =   6150
         Width           =   7650
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   25
            Left            =   3255
            TabIndex        =   94
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "P.A.V.B. : "
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
            Index           =   24
            Left            =   2295
            TabIndex        =   93
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Costo : "
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
            Index           =   23
            Left            =   5940
            TabIndex        =   92
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label2 
            Caption         =   "Gr.Net.Verd. : "
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
            Index           =   22
            Left            =   60
            TabIndex        =   91
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   21
            Left            =   6720
            TabIndex        =   90
            Top             =   360
            Width           =   680
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   20
            Left            =   1320
            TabIndex        =   89
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "P.A.V.B. % : "
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
            Index           =   19
            Left            =   4140
            TabIndex        =   88
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   18
            Left            =   5200
            TabIndex        =   87
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         Height          =   5175
         Index           =   0
         Left            =   -74520
         TabIndex        =   83
         Top             =   1515
         Width           =   10125
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   4365
            Index           =   0
            Left            =   180
            TabIndex        =   84
            Top             =   330
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   7699
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"M_Receta.frx":6F15
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
            Index           =   0
            Left            =   210
            TabIndex        =   85
            Top             =   4830
            Width           =   585
         End
      End
      Begin VB.Frame Frame4 
         Height          =   750
         Left            =   -74505
         TabIndex        =   78
         Top             =   780
         Width           =   10095
         Begin VB.ComboBox Combo2 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   675
            TabIndex        =   80
            Text            =   "Combo2"
            Top             =   270
            Width           =   2055
         End
         Begin VB.ComboBox Combo2 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   2790
            TabIndex        =   79
            Text            =   "Combo2"
            Top             =   270
            Width           =   1005
         End
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   360
            Left            =   3915
            TabIndex        =   81
            Top             =   225
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   635
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Appearance      =   1
            Style           =   1
            ImageList       =   "ImageList4"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   10
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Negrilla"
                  ImageIndex      =   8
                  Style           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Italica"
                  ImageIndex      =   9
                  Style           =   1
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Subrayado"
                  ImageIndex      =   10
                  Style           =   1
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Alinear a la Izquierda"
                  ImageIndex      =   11
                  Style           =   1
                  Value           =   1
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Centrar"
                  ImageIndex      =   12
                  Style           =   1
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Alinear a la Derecha"
                  ImageIndex      =   13
                  Style           =   1
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Justificar"
                  ImageIndex      =   14
                  Style           =   1
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Vriñeta"
                  ImageIndex      =   15
                  Style           =   1
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fonts"
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
            Left            =   135
            TabIndex        =   82
            Top             =   315
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   6
         Left            =   -74760
         TabIndex        =   59
         Top             =   5820
         Width           =   7650
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   17
            Left            =   5200
            TabIndex        =   67
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "P.A.V.B. % : "
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
            Index           =   16
            Left            =   4140
            TabIndex        =   66
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   15
            Left            =   1320
            TabIndex        =   65
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   14
            Left            =   6720
            TabIndex        =   64
            Top             =   360
            Width           =   680
         End
         Begin VB.Label Label2 
            Caption         =   "Gr.Net.Verd. : "
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
            Index           =   13
            Left            =   60
            TabIndex        =   63
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Costo : "
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
            Left            =   5940
            TabIndex        =   62
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label2 
            Caption         =   "P.A.V.B. : "
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
            Left            =   2295
            TabIndex        =   61
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   10
            Left            =   3255
            TabIndex        =   60
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Aporte Nutricionales"
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
         Height          =   4110
         Index           =   1
         Left            =   -65640
         TabIndex        =   57
         Top             =   2440
         Width           =   2885
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   3420
            Index           =   1
            Left            =   165
            TabIndex        =   58
            Top             =   360
            Width           =   2550
            _Version        =   393216
            _ExtentX        =   4498
            _ExtentY        =   6033
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
            SpreadDesigner  =   "M_Receta.frx":6F97
            ScrollBarTrack  =   3
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1635
         Index           =   5
         Left            =   -74880
         TabIndex        =   50
         Top             =   780
         Width           =   6825
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   2
            Left            =   2040
            TabIndex        =   51
            Top             =   240
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
            MaxLength       =   80
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   3
            Left            =   2040
            TabIndex        =   52
            Top             =   555
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
            MaxLength       =   80
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
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   2370
            TabIndex        =   72
            Top             =   870
            Width           =   4095
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2370
            TabIndex        =   70
            Top             =   1200
            Width           =   4095
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   1920
            Picture         =   "M_Receta.frx":73FD
            Top             =   1090
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   1920
            Picture         =   "M_Receta.frx":7707
            Top             =   780
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Plato"
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
            Index           =   19
            Left            =   300
            TabIndex        =   56
            Top             =   1260
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Categoria Dietetica"
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
            Index           =   18
            Left            =   300
            TabIndex        =   55
            Top             =   975
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre Fantasia"
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
            Index           =   17
            Left            =   300
            TabIndex        =   54
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre Receta"
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
            Index           =   16
            Left            =   300
            TabIndex        =   53
            Top             =   345
            Width           =   1695
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   2415
            TabIndex        =   71
            Top             =   1245
            Width           =   4095
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   2415
            TabIndex        =   73
            Top             =   915
            Width           =   4095
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1635
         Index           =   4
         Left            =   -65640
         TabIndex        =   41
         Top             =   780
         Width           =   2885
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   4
            Left            =   1605
            TabIndex        =   42
            Top             =   1220
            Width           =   855
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   1323
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
            MaxValue        =   "900000"
            MinValue        =   "-900000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   5
            Left            =   1605
            TabIndex        =   43
            Top             =   900
            Width           =   855
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   1323
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
            BackColor       =   16777215
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
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   6
            Left            =   1575
            TabIndex        =   44
            Top             =   270
            Width           =   855
            _Version        =   196608
            _ExtentX        =   1508
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
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
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
            Text            =   "1"
            DecimalPlaces   =   0
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   7
            Left            =   1590
            TabIndex        =   45
            Top             =   590
            Width           =   855
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   1323
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
            BackColor       =   16777215
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
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
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
            Caption         =   "Raciones"
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
            Left            =   330
            TabIndex        =   49
            Top             =   375
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "C. Servida"
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
            Left            =   330
            TabIndex        =   48
            Top             =   990
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "G. Neto"
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
            Left            =   330
            TabIndex        =   47
            Top             =   1320
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "C. Bruta"
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
            Left            =   330
            TabIndex        =   46
            Top             =   690
            Width           =   705
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   1
         Left            =   -74040
         TabIndex        =   33
         Top             =   540
         Width           =   9375
         Begin EditLib.fpText fpTnombre 
            Height          =   315
            Left            =   2040
            TabIndex        =   34
            Top             =   570
            Visible         =   0   'False
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
            AllowNull       =   -1  'True
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
            OnFocusNoSelect =   -1  'True
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
            Left            =   750
            TabIndex        =   36
            Top             =   570
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label Label1 
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
            Index           =   10
            Left            =   4680
            TabIndex        =   35
            Top             =   600
            Visible         =   0   'False
            Width           =   585
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   5265
         Left            =   -73320
         TabIndex        =   27
         Top             =   1680
         Width           =   7185
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   4155
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   960
            Width           =   6765
            _Version        =   393216
            _ExtentX        =   11933
            _ExtentY        =   7329
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
            SpreadDesigner  =   "M_Receta.frx":7A11
            VisibleCols     =   2
            VisibleRows     =   15
            ScrollBarTrack  =   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Plato"
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
            Left            =   435
            TabIndex        =   32
            Top             =   675
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Categoria Dietetica"
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
            Left            =   435
            TabIndex        =   31
            Top             =   300
            Width           =   1650
         End
         Begin VB.Label Label2 
            Caption         =   "Todos"
            Height          =   255
            Index           =   8
            Left            =   2160
            TabIndex        =   30
            Top             =   300
            Width           =   4455
         End
         Begin VB.Label Label2 
            Caption         =   "Todos"
            Height          =   255
            Index           =   9
            Left            =   2160
            TabIndex        =   29
            Top             =   660
            Width           =   4455
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1635
         Index           =   0
         Left            =   -65640
         TabIndex        =   20
         Top             =   780
         Width           =   2885
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   2
            Left            =   1605
            TabIndex        =   21
            Top             =   1220
            Width           =   855
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   1323
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
            MaxValue        =   "900000"
            MinValue        =   "-900000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   1
            Left            =   1605
            TabIndex        =   22
            Top             =   900
            Width           =   855
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   1323
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
            BackColor       =   16777215
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
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   0
            Left            =   1575
            TabIndex        =   23
            Top             =   270
            Width           =   855
            _Version        =   196608
            _ExtentX        =   1508
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
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
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
            Text            =   "1"
            DecimalPlaces   =   0
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   3
            Left            =   1590
            TabIndex        =   39
            Top             =   590
            Width           =   855
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   1323
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
            BackColor       =   16777215
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
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
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
            Caption         =   "C. Bruta"
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
            Left            =   330
            TabIndex        =   40
            Top             =   690
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "G. Neto"
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
            Left            =   330
            TabIndex        =   26
            Top             =   1320
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "C. Servida"
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
            Left            =   330
            TabIndex        =   25
            Top             =   990
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Raciones"
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
            Left            =   330
            TabIndex        =   24
            Top             =   375
            Width           =   810
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1635
         Index           =   3
         Left            =   -74880
         TabIndex        =   13
         Top             =   780
         Width           =   6825
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   2040
            TabIndex        =   14
            Top             =   240
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
            MaxLength       =   80
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   1
            Left            =   2040
            TabIndex        =   15
            Top             =   555
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
            MaxLength       =   80
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
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   2370
            TabIndex        =   76
            Top             =   870
            Width           =   4095
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   2370
            TabIndex        =   74
            Top             =   1200
            Width           =   4095
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre Receta"
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
            Index           =   5
            Left            =   300
            TabIndex        =   19
            Top             =   345
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre Fantasia"
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
            Index           =   6
            Left            =   300
            TabIndex        =   18
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Categoria Dietetica"
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
            Index           =   8
            Left            =   300
            TabIndex        =   17
            Top             =   975
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Plato"
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
            Index           =   9
            Left            =   300
            TabIndex        =   16
            Top             =   1260
            Width           =   1695
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   1920
            Picture         =   "M_Receta.frx":7E63
            Top             =   780
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   1920
            Picture         =   "M_Receta.frx":816D
            Top             =   1090
            Width           =   480
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   2415
            TabIndex        =   77
            Top             =   915
            Width           =   4095
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   2415
            TabIndex        =   75
            Top             =   1245
            Width           =   4095
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Aporte Nutricionales"
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
         Height          =   4110
         Index           =   0
         Left            =   -65640
         TabIndex        =   11
         Top             =   2440
         Width           =   2885
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   3420
            Index           =   0
            Left            =   165
            TabIndex        =   12
            Top             =   360
            Width           =   2550
            _Version        =   393216
            _ExtentX        =   4498
            _ExtentY        =   6033
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
            SpreadDesigner  =   "M_Receta.frx":8477
            ScrollBarTrack  =   3
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   2
         Left            =   -74760
         TabIndex        =   2
         Top             =   5820
         Width           =   7650
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   5
            Left            =   3255
            TabIndex        =   10
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "P.A.V.B. : "
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
            Index           =   2
            Left            =   2295
            TabIndex        =   9
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Costo : "
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
            Left            =   5940
            TabIndex        =   8
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label2 
            Caption         =   "Gr.Net.Verd. : "
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
            Index           =   1
            Left            =   60
            TabIndex        =   7
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   3
            Left            =   6720
            TabIndex        =   6
            Top             =   360
            Width           =   680
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   4
            Left            =   1320
            TabIndex        =   5
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "P.A.V.B. % : "
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
            Index           =   6
            Left            =   4140
            TabIndex        =   4
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   7
            Left            =   5200
            TabIndex        =   3
            Top             =   360
            Width           =   495
         End
      End
      Begin MSComctlLib.ImageList ImageList4 
         Left            =   -67200
         Top             =   5820
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
               Picture         =   "M_Receta.frx":88DD
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":8BF7
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":8F11
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":922B
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":9545
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":985F
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":9B79
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":9E93
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":A1AD
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":A4C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":A7E1
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":A93B
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":AA95
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":ABEF
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":AD49
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   720
         Left            =   -74760
         TabIndex        =   37
         Top             =   5460
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   1270
         ButtonWidth     =   2858
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         ImageList       =   "ImageList4"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agrega Ing.    "
               Description     =   "Agregar Ingrediente"
               Object.ToolTipText     =   "Agregar Ingrediente"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agrega Linea    "
               Description     =   "Insertar Linea"
               Object.ToolTipText     =   "Insertar Linea"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Borra Linea    "
               Description     =   "Borrar Ingrediente"
               Object.ToolTipText     =   "Borrar Ingrediente"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Mueve Up    "
               Description     =   "Mueve Up"
               Object.ToolTipText     =   "Mueve Up"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Mueve Dn    "
               Description     =   "Mueve Dn"
               Object.ToolTipText     =   "Mueve Dn"
               ImageIndex      =   5
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2880
         Index           =   1
         Left            =   -74880
         TabIndex        =   38
         Top             =   2460
         Width           =   9105
         _Version        =   393216
         _ExtentX        =   16060
         _ExtentY        =   5080
         _StockProps     =   64
         ColsFrozen      =   2
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
         GridShowHoriz   =   0   'False
         MaxCols         =   10
         MaxRows         =   40
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_Receta.frx":AEA3
         VisibleCols     =   10
         VisibleRows     =   40
         ScrollBarTrack  =   3
      End
      Begin MSComctlLib.Toolbar Toolbar4 
         Height          =   720
         Left            =   -74760
         TabIndex        =   68
         Top             =   5460
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   1270
         ButtonWidth     =   2858
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         ImageList       =   "ImageList4"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agrega Ing.    "
               Description     =   "Agregar Ingrediente"
               Object.ToolTipText     =   "Agregar Ingrediente"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agrega Linea    "
               Description     =   "Insertar Linea"
               Object.ToolTipText     =   "Insertar Linea"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Borra Linea    "
               Description     =   "Borrar Ingrediente"
               Object.ToolTipText     =   "Borrar Ingrediente"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Mueve Up    "
               Description     =   "Mueve Up"
               Object.ToolTipText     =   "Mueve Up"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Mueve Dn    "
               Description     =   "Mueve Dn"
               Object.ToolTipText     =   "Mueve Dn"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2880
         Index           =   2
         Left            =   -74880
         TabIndex        =   69
         Top             =   2460
         Width           =   9105
         _Version        =   393216
         _ExtentX        =   16060
         _ExtentY        =   5080
         _StockProps     =   64
         ColsFrozen      =   2
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
         GridShowHoriz   =   0   'False
         MaxCols         =   10
         MaxRows         =   40
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_Receta.frx":B6AB
         VisibleCols     =   10
         VisibleRows     =   40
         ScrollBarTrack  =   3
      End
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   720
         Left            =   210
         TabIndex        =   117
         Top             =   5790
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   1270
         ButtonWidth     =   2858
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         ImageList       =   "ImageList4"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agrega Ing.    "
               Description     =   "Agregar Ingrediente"
               Object.ToolTipText     =   "Agregar Ingrediente"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agrega Linea    "
               Description     =   "Insertar Linea"
               Object.ToolTipText     =   "Insertar Linea"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Borra Linea    "
               Description     =   "Borrar Ingrediente"
               Object.ToolTipText     =   "Borrar Ingrediente"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Mueve Up    "
               Description     =   "Mueve Up"
               Object.ToolTipText     =   "Mueve Up"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Mueve Dn    "
               Description     =   "Mueve Dn"
               Object.ToolTipText     =   "Mueve Dn"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2880
         Index           =   3
         Left            =   90
         TabIndex        =   118
         Top             =   2790
         Width           =   9105
         _Version        =   393216
         _ExtentX        =   16060
         _ExtentY        =   5080
         _StockProps     =   64
         ColsFrozen      =   2
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
         GridShowHoriz   =   0   'False
         MaxCols         =   11
         MaxRows         =   40
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_Receta.frx":BEBE
         VisibleCols     =   10
         VisibleRows     =   40
         ScrollBarTrack  =   3
      End
   End
End
Attribute VB_Name = "M_Receta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset, RS2 As New ADODB.Recordset
Dim itab As Integer, i As Integer, itexto As Integer
Dim ibusca As Long, codigo As Long, codcatdie As Long, codtipplato As Long, inddet As Long, indapo As Long
Dim est As Boolean
Dim MsgTitulo As String, codpro1 As String, codpro2 As String, modo As String, metodoreceta As String, nombusca As String, grupovulnerable As String
Dim canpro1 As Double, canpro2 As Double, pctapr1 As Double, pctapr2 As Double, pctcoc1 As Double, pctcoc2 As Double, pctnut1 As Double, pctnut2 As Double, cospro As Double, candiet As Double, cosrec As Double

Private Sub Combo2_Click(Index As Integer)
Select Case Index
Case 0
    RichTextBox1(1).SelFontName = Combo2(0).text
Case 1
    RichTextBox1(1).SelFontSize = Combo2(1).text
End Select
End Sub

Private Sub Form_Activate()
fg_descarga
TraerFechaCierre
End Sub

Private Sub Form_Load()

Me.Height = 8070
Me.Width = 12735
fg_centra Me
Me.HelpContextID = vg_OpcM
MsgTitulo = "Recetas"
est = True
Dim i As Long
With vaSpread2(i)
    For i = 0 To 3
        .Row = -1: vaSpread1(i).Col = -1
        .BackColor = &H80000018
        If i <> 3 Then
           vaSpread2(i).Row = -1: vaSpread2(i).Col = -1
           vaSpread2(i).BackColor = &H80000018
        End If
    Next i
End With
'------- Llenar palabra
With Combo2(0)
    .AddItem Screen.Fonts(0)
    For i = 1 To Screen.FontCount - 1
        .AddItem Screen.Fonts(i)
    Next
    .ListIndex = 0
End With
'------- Llenar tamaño
With Combo2(1)
    .AddItem 8
    For i = 9 To 72
        .AddItem i
    Next i
    .ListIndex = 0
End With
'------- Mover nutrientes
Dim ii As Long
ii = 1
RS.Open RutinaLectura.Nutriente(1, 0, ""), vg_db, adOpenStatic
If Not RS.EOF Then
   vaSpread2(0).MaxRows = 0
   vaSpread2(1).MaxRows = 0
   vaSpread2(2).MaxRows = 0
   Do While Not RS.EOF
      For i = 0 To 2
          grdAddRow vaSpread2(i)
               
          grdCellTypeStatic vaSpread2(i), 1, ii, 1
          grdSetText vaSpread2(i), 1, ii, IIf(IsNull(RS!nut_codigo), 0, RS!nut_codigo)
          
          grdCellTypeStatic vaSpread2(i), 2, ii, 0
          grdSetText vaSpread2(i), 2, ii, IIf(IsNull(RS!nut_nombre), "", Trim(RS!nut_nombre))
               
          grdCellTypeStatic vaSpread2(i), 3, ii, 1
          grdSetText vaSpread2(i), 3, ii, Format(0, fg_Pict(6, 2))
          grdRowColForeColor vaSpread2(i), ii, ii, 3, 3, &HFF0000
      Next i
      RS.MoveNext: ii = ii + 1
   Loop
End If
RS.Close: Set RS = Nothing

itexto = 1: vg_dbndecimal = 2: ibusca = 0
modo = ""
Gl_Mo_Botones Me, 3
Gl_Ac_Botones Me, 3, 3, modo
Hab_Des 3
vg_filcatdie = 0: vg_filtippla = 0
If vg_newcodrec > 0 Then
   Me.HelpContextID = "1020000"
    modo = "M": Hab_Des 0: Gl_Ac_Botones Me, 3, IIf(vg_newestrec = True, 3, 4), modo
    If vg_tiprec = 0 Then
       inddet = 1: indapo = 0
    ElseIf vg_tiprec = -1 Then
       inddet = 2: indapo = 1
    ElseIf vg_tiprec > 0 Then
       inddet = 3: indapo = 2
       RS.Open RutinaLectura.Regimen(2, vg_tiprec, ""), vg_db, adOpenStatic
       If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(6).Caption = "": Exit Sub
       fpLongInteger1(1).Value = vg_tiprec
       fpayuda(6).Caption = Trim(RS!reg_nombre)
       RS.Close: Set RS = Nothing
       Dim X As Boolean
       vaSpread1(inddet).TextTip = 2
       vaSpread1(inddet).TextTipDelay = 0
       X = vaSpread1(inddet).SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
    End If
    MoverDetalleDatos
    If vg_newestrec = True Or ("S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) And vg_5etapas) Then
        Frame1(0).Enabled = False
        Frame1(3).Enabled = False
'        Frame1(4).Enabled = False
        Frame1(5).Enabled = False
        Toolbar2.Enabled = False
        Toolbar4.Enabled = False
        Toolbar5.Enabled = False
    End If
    Exit Sub
End If

RS.Open "SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'catdefecto'", vg_db, adOpenStatic
If Not RS.EOF Then

   vg_filcatdie = IIf(IsNull(RS!par_valor), 0, RS!par_valor)
   
   If RS!par_valor = "0" Then
   
      Label2(8).Caption = "Todos"
   
   Else
   
      Label2(8).Caption = fg_BuscaenArbol(RS!par_valor, "a_recetacatdie", "car_codigo")

   End If
    
End If

RS.Close
Set RS = Nothing

Mover_ListaReceta
est = False

'MVA - MVI - BLOQUEO BOTON TOOLBAR ACTUALIZAR RECETA - 2013-01-18
If vg_Block_Botton_Actua_Receta_MVI = True Then

    Toolbar1.Buttons(3).Enabled = BloqueaBotonActua(vg_Clave_MVI)

End If
'MVA - MVI - BLOQUEO BOTON TOOLBAR ACTUALIZAR RECETA - 2013-01-18

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
If est Then Exit Sub
Dim RS As New ADODB.Recordset
Select Case Index
Case 1
    RS.Open RutinaLectura.Regimen(2, Val(fpLongInteger1(1).Value), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(6).Caption = "": Exit Sub
    fpayuda(6).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    vg_tiprec = IIf(Val(fpLongInteger1(1).Value) > 0, Val(fpLongInteger1(1).Value), -2): inddet = 3: indapo = 2
    MoverDetalleDatos
    If Not vg_5etapas And Val(fpLongInteger1(1).Value) >= 10000 Then
       Frame1(0).Enabled = False
       Frame1(3).Enabled = False
       Frame1(4).Enabled = False
       Frame1(5).Enabled = False
       Toolbar2.Enabled = False
       Toolbar4.Enabled = False
       Toolbar5.Enabled = False
       vaSpread1(inddet).Col = 1: vaSpread1(inddet).Col2 = 10: vaSpread1(inddet).Row = 1: vaSpread1(inddet).Row2 = 40
       vaSpread1(inddet).BlockMode = True
       ' Lock cells
       vaSpread1(inddet).Lock = True
       ' Protect the cells from being edited
       vaSpread1(inddet).Protect = True
       ' Turn block mode off
       vaSpread1(inddet).BlockMode = False
       '  modo = ""
    Else
       Frame1(0).Enabled = True
       Frame1(3).Enabled = True
       Frame1(4).Enabled = True
       Frame1(5).Enabled = True
       Toolbar2.Enabled = True
       Toolbar4.Enabled = True
       Toolbar5.Enabled = True
    End If
End Select
End Sub

Private Sub fpText1_Change(Index As Integer)
If est Then Exit Sub
If Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpTnombre_Change()
Dim RS As New ADODB.Recordset
Dim sql1 As String
sql1 = IIf(vg_tipbase = "1", " UCASE(rec_nombre) ", " UPPER(rec_nombre) ")
RS.Open "SELECT rec_codigo, rec_nombre FROM b_receta WHERE (rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
        "AND  (rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) " & _
        "AND  (rec_fecvig > " & Format(Date, "yyyymmdd") & " OR rec_fecvig <= 0 OR (rec_fecvig) IS NULL) " & _
        "AND " & sql1 & " LIKE '%" & LimpiaDato(UCase(fpTnombre.text)) & "%' ORDER BY rec_nombre", vg_db, adOpenStatic
If ibusca <> RS.RecordCount Then ibusca = RS.RecordCount: vaSpread1(0).MaxRows = RS.RecordCount
i = 1
If Not RS.EOF Then
   Do While Not RS.EOF
      grdCellTypeStatic vaSpread1(0), 1, i, 1
      grdSetText vaSpread1(0), 1, i, IIf(IsNull(RS!rec_codigo), 0, RS!rec_codigo)
      
      grdCellTypeStatic vaSpread1(0), 2, i, 0
      grdSetText vaSpread1(0), 2, i, IIf(IsNull(RS!rec_nombre), 0, RS!rec_nombre)
      
      RS.MoveNext: i = i + 1
   Loop
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(3) = True
   SSTab1.TabEnabled(4) = True
End If
RS.Close: Set RS = Nothing
vaSpread1(0).SetActiveCell 1, 1
Label1(10).Caption = Format(vaSpread1(0).MaxRows, fg_Pict(7, 0)) & " Reg. Encontrados"
End Sub

Private Sub RichTextBox1_Change(Index As Integer)
'If RichTextBox1(0).TextRTF <> metodoreceta And itexto = 0 And modo = "M" Then
If itexto = 0 And modo = "M" Then
   If Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
End If
End Sub

Private Sub RichTextBox1_Click(Index As Integer)
'------- Bold
If RichTextBox1(Index).SelBold = False Or IsNull(RichTextBox1(Index).SelBold) Then
   Toolbar3.Buttons(1).Value = 1: Toolbar3.Buttons(1).Value = 0
ElseIf RichTextBox1(Index).SelBold = True Then
   Toolbar3.Buttons(1).Value = 0: Toolbar3.Buttons(1).Value = 1
End If
'------- Italic
If RichTextBox1(Index).SelItalic = False Or IsNull(RichTextBox1(Index).SelItalic) Then
   Toolbar3.Buttons(2).Value = 1: Toolbar3.Buttons(2).Value = 0
ElseIf RichTextBox1(Index).SelItalic = True Then
   Toolbar3.Buttons(2).Value = 0: Toolbar3.Buttons(2).Value = 1
End If
'------- Subrayado
If RichTextBox1(Index).SelUnderline = False Or IsNull(RichTextBox1(Index).SelUnderline) Then
   Toolbar3.Buttons(3).Value = 1: Toolbar3.Buttons(3).Value = 0
ElseIf RichTextBox1(Index).SelUnderline = True Then
   Toolbar3.Buttons(3).Value = 0: Toolbar3.Buttons(3).Value = 1
End If
'------- Viñetas
If RichTextBox1(Index).SelBullet = False Or IsNull(RichTextBox1(Index).SelBullet) Then
   Toolbar3.Buttons(10).Value = 1: Toolbar3.Buttons(10).Value = 0
ElseIf RichTextBox1(Index).SelBullet = True Then
   Toolbar3.Buttons(10).Value = 0: Toolbar3.Buttons(10).Value = 1
End If
If RichTextBox1(Index).SelAlignment = 0 Then Toolbar3.Buttons(5).Value = 0: Toolbar3.Buttons(5).Value = 1: Toolbar3.Buttons(6).Value = 1: Toolbar3.Buttons(6).Value = 0: Toolbar3.Buttons(7).Value = 1: Toolbar3.Buttons(7).Value = 0
If RichTextBox1(Index).SelAlignment = 2 Then Toolbar3.Buttons(5).Value = 1: Toolbar3.Buttons(5).Value = 0: Toolbar3.Buttons(6).Value = 0: Toolbar3.Buttons(6).Value = 1: Toolbar3.Buttons(7).Value = 1: Toolbar3.Buttons(7).Value = 0
If RichTextBox1(Index).SelAlignment = 1 Then Toolbar3.Buttons(5).Value = 1: Toolbar3.Buttons(5).Value = 0: Toolbar3.Buttons(6).Value = 1: Toolbar3.Buttons(6).Value = 0: Toolbar3.Buttons(7).Value = 0: Toolbar3.Buttons(7).Value = 1
End Sub

Private Sub RichTextBox1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'------- Bold
If RichTextBox1(Index).SelBold = False Or IsNull(RichTextBox1(Index).SelBold) Then
   Toolbar3.Buttons(1).Value = 1: Toolbar3.Buttons(1).Value = 0
ElseIf RichTextBox1(Index).SelBold = True Then
   Toolbar3.Buttons(1).Value = 0: Toolbar3.Buttons(1).Value = 1
End If
'------- Italic
If RichTextBox1(Index).SelItalic = False Or IsNull(RichTextBox1(Index).SelItalic) Then
   Toolbar3.Buttons(2).Value = 1: Toolbar3.Buttons(2).Value = 0
ElseIf RichTextBox1(Index).SelItalic = True Then
   Toolbar3.Buttons(2).Value = 0: Toolbar3.Buttons(2).Value = 1
End If
'------- Subrayado
If RichTextBox1(Index).SelUnderline = False Or IsNull(RichTextBox1(Index).SelUnderline) Then
   Toolbar3.Buttons(3).Value = 1: Toolbar3.Buttons(3).Value = 0
ElseIf RichTextBox1(Index).SelUnderline = True Then
   Toolbar3.Buttons(3).Value = 0: Toolbar3.Buttons(3).Value = 1
End If
'------- Viñetas
If RichTextBox1(Index).SelBullet = False Or IsNull(RichTextBox1(Index).SelBullet) Then
   Toolbar3.Buttons(10).Value = 1: Toolbar3.Buttons(10).Value = 0
ElseIf RichTextBox1(Index).SelBullet = True Then
   Toolbar3.Buttons(10).Value = 0: Toolbar3.Buttons(10).Value = 1
End If
If RichTextBox1(Index).SelAlignment = 0 Then Toolbar3.Buttons(5).Value = 0: Toolbar3.Buttons(5).Value = 1: Toolbar3.Buttons(6).Value = 1: Toolbar3.Buttons(6).Value = 0: Toolbar3.Buttons(7).Value = 1: Toolbar3.Buttons(7).Value = 0
If RichTextBox1(Index).SelAlignment = 2 Then Toolbar3.Buttons(5).Value = 1: Toolbar3.Buttons(5).Value = 0: Toolbar3.Buttons(6).Value = 0: Toolbar3.Buttons(6).Value = 1: Toolbar3.Buttons(7).Value = 1: Toolbar3.Buttons(7).Value = 0
If RichTextBox1(Index).SelAlignment = 1 Then Toolbar3.Buttons(5).Value = 1: Toolbar3.Buttons(5).Value = 0: Toolbar3.Buttons(6).Value = 1: Toolbar3.Buttons(6).Value = 0: Toolbar3.Buttons(7).Value = 0: Toolbar3.Buttons(7).Value = 1
End Sub

Private Sub Image1_Click(Index As Integer)
vg_codigo = 0
Select Case Index
Case 2
    vg_nombre = ""
    vg_codigo = 0
    vg_left = fpayuda(2).Left + 550
    B_ArbEst.MoverDatosTvwDir "a_recetacatdie", "car_", "Categoria Dietetica"
    B_ArbEst.Show 1
    Me.Refresh
    If vg_codigo = 0 Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
    codcatdie = vg_codigo
    fpayuda(2).Caption = vg_nombre
Case 3
    vg_nombre = ""
    vg_codigo = 0
    vg_left = fpayuda(3).Left + 550
    B_ArbEst.MoverDatosTvwDir "a_recetatippla", "tip_", "Tipo Plato"
    B_ArbEst.Show 1
    Me.Refresh
    If vg_codigo = 0 Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
    codtipplato = vg_codigo
    fpayuda(3).Caption = vg_nombre
Case 6
    vg_left = fpayuda(6).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", IIf(vg_5etapas = False, "No5etapas", "Gen")
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(6).Caption = vg_nombre
End Select
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
Case 1, 2, 3
    est = True
    If vaSpread1(0).MaxRows > 0 And modo = "M" Then
       modo = "M"
       If SSTab1.Tab = 1 Then
          vg_tiprec = 0: inddet = 1: indapo = 0
          MoverDetalleDatos
       ElseIf SSTab1.Tab = 2 Then
          vg_tiprec = -1: inddet = 2: indapo = 1
          MoverDetalleDatos
       ElseIf SSTab1.Tab = 3 Then
          vg_tiprec = IIf(Val(fpLongInteger1(1).Value) > 0, Val(fpLongInteger1(1).Value), -2): inddet = 3: indapo = 2
          MoverDetalleDatos
       End If
       If Not vg_5etapas And Val(fpLongInteger1(1).Value) >= 10000 Then
          Frame1(0).Enabled = False
          Frame1(3).Enabled = False
          Frame1(4).Enabled = False
          Frame1(5).Enabled = False
          Toolbar2.Enabled = False
          Toolbar4.Enabled = False
          Toolbar5.Enabled = False
       Else
          Frame1(0).Enabled = True
'          Frame1(3).Enabled = True
          Frame1(4).Enabled = True
'          Frame1(5).Enabled = True
          Toolbar2.Enabled = True
          Toolbar4.Enabled = True
          Toolbar5.Enabled = True
       End If
    ElseIf vaSpread1(0).MaxRows < 1 And modo = "M" Then
       SSTab1.Tab = 0
       Exit Sub
    End If
Case 4
    If vaSpread1(0).MaxRows < 1 Then SSTab1.Tab = 0: Exit Sub
    izquerda = 0
    CargaMetodoReceta
Case 5
    If vaSpread1(0).MaxRows < 1 Then SSTab1.Tab = 0: Exit Sub
    izquerda = 0
    CargaGrupoVulnerable
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim indice As Long
Dim tiprec As String, sql1 As String
On Error GoTo Man_Error
Select Case Button.Index
Case 1 '----------> Agregar
    modo = "A": itexto = 1: est = True
    inddet = 1
    LimpiarVariable
    Hab_Des 0: SSTab1.Tab = 1
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False
    Toolbar2.Enabled = True
    Gl_Ac_Botones Me, 3, 0, modo
    itexto = 0: est = False
Case 3 '---------> Modificar
    modo = "M"
    Dim estrec As Boolean
    estrec = False
    If SSTab1.Tab = 2 Or SSTab1.Tab = 3 Then
       estrec = True
       For i = 1 To 40
           vaSpread1(inddet).Row = i
           vaSpread1(inddet).Col = 1
           If Trim(vaSpread1(inddet).text) <> "" Then estrec = False: Exit For
       Next i
    End If
    If vaSpread1(0).MaxRows < 1 Or SSTab1.Tab = 1 Or SSTab1.Tab = 4 Or SSTab1.Tab = 5 Or estrec Then Exit Sub
    If (SSTab1.Tab = 1 Or SSTab1.Tab = 2) And vg_modrec = False Then
       Hab_Des 0
    ElseIf (SSTab1.Tab = 1 Or SSTab1.Tab = 4 Or SSTab1.Tab = 5) And vg_modrec = True Then
       Hab_Des 0
    End If
    Gl_Ac_Botones Me, 3, 0, modo
    If SSTab1.Tab = 1 Or SSTab1.Tab = 0 Then
       SSTab1.TabEnabled(4) = False
       SSTab1.TabEnabled(5) = False
       SSTab1.Tab = 1
    ElseIf SSTab1.Tab = 4 Then
       If vg_modrec = True Then Gl_Ac_Botones Me, 3, 0, modo: SSTab1.TabEnabled(1) = False Else Gl_Ac_Botones Me, 3, 4, modo
       CargaMetodoReceta
    ElseIf SSTab1.Tab = 5 Then
       If vg_modrec = True Then Gl_Ac_Botones Me, 3, 0, modo: SSTab1.TabEnabled(1) = False Else Gl_Ac_Botones Me, 3, 4, modo
       CargaGrupoVulnerable
    End If
Case 5 '---------> Borrar
    If vaSpread1(0).MaxRows < 1 Then Exit Sub
    vaSpread1(0).Row = vaSpread1(0).ActiveRow
    vaSpread1(0).Col = 1: codigo = Val(vaSpread1(0).text)
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    '-------> Borrando tabla receta
    vg_db.BeginTrans
    vg_db.Execute "DELETE b_recetadet FROM b_recetadet WHERE red_codigo=" & codigo & ""
    vg_db.Execute "DELETE b_receta FROM b_receta WHERE rec_codigo=" & codigo & ""
    vaSpread1(0).Row = vaSpread1(0).ActiveRow
    vaSpread1(0).DeleteRows vaSpread1(0).Row, 1
    vaSpread1(0).MaxRows = vaSpread1(0).MaxRows - 1
    vaSpread1(0).Row = vaSpread1(0).MaxRows
    Label1(10).Caption = Format(vaSpread1(0).MaxRows, fg_Pict(7, 0)) & " Registros"
    If vaSpread1(0).MaxRows < 1 Then
       Label1(1).Visible = False
       fpTnombre.Visible = False
       Label1(10).Visible = False
       SSTab1.TabEnabled(1) = False
       SSTab1.TabEnabled(2) = False
       SSTab1.TabEnabled(3) = False
       SSTab1.TabEnabled(4) = False
       SSTab1.TabEnabled(5) = False
       SSTab1.Tab = 0
       Gl_Ac_Botones Me, 3, 2, modo
    Else
       SSTab1.TabEnabled(1) = True
       SSTab1.Tab = 0
       fpTnombre.SetFocus
    End If
    vg_db.CommitTrans
Case 7 '---------> Actualizar Lista
    Mover_ListaReceta
Case 10 '---------> Cancelar
    If MsgBox("Cancela registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If SSTab1.Tab = 1 Or SSTab1.Tab = 2 Or SSTab1.Tab = 3 Then
       sql1 = IIf(vg_opbase = "1", " UCASE(rec_nombre) ", " UPPER(rec_nombre) ")
       RS.Open "SELECT COUNT(*) AS nreg FROM b_receta WHERE (rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
               "AND (rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) " & _
               "AND " & sql1 & " LIKE '%" & LimpiaDato(UCase(fpTnombre.text)) & "%'", vg_db, adOpenStatic
       If RS.EOF Or RS!nreg = 0 Then RS.Close: Set RS = Nothing: Hab_Des 2: Gl_Ac_Botones Me, 3, 2, modo: SSTab1.Tab = 0: Exit Sub  'Ac_Boton 2
       modo = "M"
       RS.Close: Set RS = Nothing
       If modo = "A" Then
          SSTab1.Tab = 0
       ElseIf modo = "M" Then
          MoverDetalleDatos
          SSTab1.TabEnabled(2) = True
          SSTab1.TabEnabled(3) = True
          SSTab1.TabEnabled(4) = True
          SSTab1.TabEnabled(5) = True
       End If
       If vg_newcodrec = 0 Then
          If vg_modrec = True Then Gl_Ac_Botones Me, 3, 1, modo Else Gl_Ac_Botones Me, 3, 4, modo
          Hab_Des 1
       ElseIf vg_newcodrec > 0 Then
'          Gl_Ac_Botones Me, 3, 3, modo: Hab_Des 0
          Gl_Ac_Botones Me, 3, 4, modo: Hab_Des 0
          SSTab1.TabEnabled(3) = True
       End If
    ElseIf SSTab1.Tab = 4 Then
       vaSpread1(0).Row = vaSpread1(0).ActiveRow
       vaSpread1(0).Col = 1
       codigo = Val(vaSpread1(0).text)
       RS.Open "SELECT rec_metpre FROM b_receta WHERE rec_codigo = " & codigo & "", vg_db, adOpenStatic
       If RS.EOF Then RS.Close: Set RS = Nothing: Hab_Des 2: Gl_Ac_Botones Me, 3, 2, modo: SSTab1.Tab = 0: Exit Sub 'Ac_Boton 2
       modo = "M"
       RS.Close: Set RS = Nothing
       If modo = "A" Then
          SSTab1.Tab = 0
       ElseIf modo = "M" Then
          CargaMetodoReceta
          SSTab1.TabEnabled(1) = True
       End If
       If vg_modrec = True Then Gl_Ac_Botones Me, 3, 1, modo Else Gl_Ac_Botones Me, 3, 4, modo
       Hab_Des 1
    End If
Case 12 '---------> Confirmar
    If SSTab1.Tab = 1 And (LimpiaDato(Trim(fpText1(0).text)) = "" Or LimpiaDato(Trim(fpText1(1).text)) = "" Or fpayuda(2).Caption = "" Or fpayuda(3).Caption = "" Or Val(fpDouble1(0).Value) = 0) Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo:  Exit Sub
    If SSTab1.Tab = 3 And Val(fpLongInteger1(1).Value) = 0 Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo:  Exit Sub
    If vg_newcodrec > 0 Then
       modo = "M": tipprec = "0"
    Else
       tipprec = "0"
    End If
    If modo = "A" Or modo = "M" Then
       indice = 0
       If modo = "A" Then
          vg_db.BeginTrans
            RS.Open "SELECT rec_codigo FROM b_receta ORDER BY rec_codigo DESC", vg_db, adOpenStatic
            If Not RS.EOF Then RS.MoveFirst: indice = RS!rec_codigo + 1 Else indice = 1
            RS.Close: Set RS = Nothing
            '------- Grabar encabezado recetas
            vg_db.Execute "INSERT INTO b_receta VALUES (" & indice & ", " & codcatdie & ", " & codtipplato & ", " & _
                          "'" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & LimpiaDato(Trim(fpText1(1).text)) & "', ' ', ' ', " & _
                          "' ', 1, " & IIf(vg_tiprec > 0, 1, vg_tiprec) & ", 0, ' ')"
            '------- Grabar detalle recetas
            For i = 1 To vaSpread1(inddet).MaxRows
                vaSpread1(inddet).Row = i
                vaSpread1(inddet).Col = 1
                If Trim(vaSpread1(inddet).text) <> "" Then
                   codpro1 = 0: canpro1 = 0: pctapr1 = 0: pctcoc1 = 0: pctnut1 = 0
                   vaSpread1(inddet).Col = 1: codpro1 = vaSpread1(inddet).text
                   vaSpread1(inddet).Col = 3: canpro1 = vaSpread1(inddet).text
                   vaSpread1(inddet).Col = 5: pctapr1 = vaSpread1(inddet).text
                   vaSpread1(inddet).Col = 6: pctcoc1 = vaSpread1(inddet).text
                   vaSpread1(inddet).Col = 8: pctnut1 = vaSpread1(inddet).text
                   vaSpread1(inddet).Col = 10: cospro = vaSpread1(inddet).text
                   vg_db.Execute "INSERT INTO b_recetadet VALUES (" & indice & ", " & i & ", '" & codpro1 & "', " & canpro1 & ", " & cospro & ", " & pctapr1 & ", " & pctcoc1 & ", " & pctnut1 & ", " & 0 & ", '0')"
                End If
            Next i
            modo = "M"
          vg_db.CommitTrans
          If vg_newcodrec > 0 Then vg_newcodrec = indice: vg_newnomrec = LimpiaDato(Trim(fpText1(0).text)): Exit Sub
          vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1: vaSpread1(0).Row = vaSpread1(0).MaxRows
          vaSpread1(0).Col = 1: vaSpread1(0).TypeHAlign = 1: vaSpread1(0).text = indice
          vaSpread1(0).Col = 2: vaSpread1(0).TypeHAlign = 0: vaSpread1(0).text = LimpiaDato(Trim(fpText1(0).text))
       Else
          vg_db.BeginTrans
            If SSTab1.Tab = 1 Or SSTab1.Tab = 2 Or SSTab1.Tab = 3 Then
               If fpDouble1(1).Value = "" Then fpDouble1(1).Value = 0
               If fpDouble1(2).Value = "" Then fpDouble1(2).Value = 0
               If fpDouble1(3).Value = "" Then fpDouble1(3).Value = 0
               '------- Actualizar encabezado recetas
               vg_db.Execute "UPDATE b_receta SET rec_catdie = " & codcatdie & ", rec_tippla = " & codtipplato & ", " & _
                             "rec_nombre = '" & LimpiaDato(Trim(fpText1(0).text)) & "', rec_nomfan = '" & LimpiaDato(Trim(fpText1(1).text)) & "', " & _
                             "rec_basrac = 1 WHERE rec_codigo = " & codigo & ""
               '------- Grabar detalle recetas
               vg_db.Execute "DELETE b_recetadet FROM b_recetadet WHERE red_codigo = " & codigo & " AND red_tiprec = " & vg_tiprec & " AND red_cencos = '" & IIf(vg_tiprec = 0, 0, MuestraCasino(1)) & "'"
               For i = 1 To vaSpread1(inddet).MaxRows
                   vaSpread1(inddet).Row = i
                   vaSpread1(inddet).Col = 1
                   If Trim(vaSpread1(inddet).text) <> "" Then
                      codpro1 = 0: canpro1 = 0: pctapr1 = 0: pctcoc1 = 0: pctnut1 = 0: cospro = 0
                      vaSpread1(inddet).Col = 1: codpro1 = vaSpread1(inddet).text
                      vaSpread1(inddet).Col = 3: canpro1 = vaSpread1(inddet).text
                      vaSpread1(inddet).Col = 5: pctapr1 = vaSpread1(inddet).text
                      vaSpread1(inddet).Col = 6: pctcoc1 = vaSpread1(inddet).text
                      vaSpread1(inddet).Col = 8: pctnut1 = vaSpread1(inddet).text
                      vaSpread1(inddet).Col = 10: cospro = vaSpread1(inddet).text
                      vg_db.Execute "INSERT INTO b_recetadet VALUES (" & codigo & ", " & i & ", '" & codpro1 & "', " & canpro1 & ", " & cospro & ", " & pctapr1 & ", " & pctcoc1 & ", " & pctnut1 & ", " & vg_tiprec & ", '" & IIf(vg_tiprec = 0, 0, MuestraCasino(1)) & "')"
                   End If
               Next i
               If vg_newcodrec > 0 Then
                  vg_newnomrec = LimpiaDato(Trim(fpText1(0).text))
                  vg_db.CommitTrans
'                  Gl_Ac_Botones Me, 3, 3, modo
                  Gl_Ac_Botones Me, 3, 4, modo
                  SSTab1.TabEnabled(1) = True: SSTab1.TabEnabled(2) = True: SSTab1.TabEnabled(3) = True
                  vg_auxtiprec = vg_tiprec
                  Exit Sub
               End If
               vaSpread1(0).Row = vaSpread1(0).ActiveRow
               vaSpread1(0).Col = 2
               vaSpread1(0).TypeHAlign = 0
               vaSpread1(0).Value = LimpiaDato(Trim(fpText1(0).text))
            ElseIf SSTab1.Tab = 4 And vg_modrec Then
               vg_db.Execute "UPDATE b_receta SET rec_metpre = '" & Trim(LimpiaDato(RichTextBox1(0).TextRTF)) & "' WHERE rec_codigo = " & codigo & ""
            ElseIf SSTab1.Tab = 5 And vg_modrec Then
               vg_db.Execute "UPDATE b_receta SET rec_gruvul = '" & Trim(LimpiaDato(RichTextBox1(1).TextRTF)) & "' WHERE rec_codigo = " & codigo & ""
            End If
          vg_db.CommitTrans
       End If
       vaSpread1(0).SortKey(1) = 2
       vaSpread1(0).SortKeyOrder(1) = 1
       vaSpread1(0).Sort -1, -1, vaSpread1(0).MaxCols, vaSpread1(0).MaxRows, SortByRow
       For i = 1 To vaSpread1(0).MaxRows
           vaSpread1(0).Row = i: vaSpread1(0).Col = 1
           If vaSpread1(0).text = codigo Then vaSpread1(0).OperationMode = 3: vaSpread1(0).SetActiveCell 1, vaSpread1(0).Row
       Next i
       itexto = 1
       If vg_modrec = True Then Gl_Ac_Botones Me, 3, 1, modo Else Gl_Ac_Botones Me, 3, 4, modo
       Hab_Des 1
       Label1(10).Caption = Format(vaSpread1(0).MaxRows, fg_Pict(7, 0)) & " Registros"
       itexto = 0
    End If
Case 15 '------- Copiar Receta Patrón
     vg_swpegreceta = 0: vg_codreceta = codigo
     M_CpoRec.Show 1
     If vg_swpegreceta = 0 Then Exit Sub
     If vg_newcodrec > 0 Then
        RS.Open "SELECT rec_tiprec FROM  b_receta WHERE rec_codigo = " & vg_codreceta & "", vg_db, adOpenStatic
        If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Receta No Existe", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
        'If RS!rec_tiprec < 1 Then
        vg_tiprec = RS!rec_tiprec: vg_auxtiprec = RS!rec_tiprec
        RS.Close: Set RS = Nothing
     Else
        RS.Open "SELECT rec_tiprec FROM b_receta WHERE b_receta.rec_codigo = " & vg_codreceta & "", vg_db, adOpenStatic
        If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Receta No Existe", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
        vg_tiprec = RS!rec_tiprec: vg_auxtiprec = RS!rec_tiprec
        RS.Close: Set RS = Nothing
     End If
     modo = "M"
     If vg_tiprec = -1 Or vg_tiprec = -2 Then
        inddet = 2: indapo = 1: SSTab1.Tab = 2
     Else
        est = True
        RS.Open RutinaLectura.Regimen(2, vg_tiprec, ""), vg_db, adOpenStatic
        If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(6).Caption = "": Exit Sub
        fpLongInteger1(1).Value = vg_tiprec
        fpayuda(6).Caption = Trim(RS!reg_nombre)
        RS.Close: Set RS = Nothing
        est = False: inddet = 3: indapo = 2: SSTab1.Tab = 3
     End If
     MoverDetalleDatos
Case 17 '------- Filtrar
    SSTab1.Tab = 0
    B_DieTip.Show 1
    Label2(8).Caption = "Todos": Label2(9).Caption = "Todos"
    If vg_filnomtippla <> "" Then Label2(9).Caption = vg_filnomtippla
    If vg_filnomcatdie <> "" Then Label2(8).Caption = vg_filnomcatdie
    If vg_opcion = 2 Then Exit Sub
    Mover_ListaReceta
Case 19 '------- Cambiar productos en recetas
'    If vaSpread1(0).MaxRows < 1 Then Exit Sub
'    SSTab1.Tab = 0
'    M_ReePro.Show 1
Case 21 '------- Imprimir
    If vaSpread1(0).MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, "Receta": Exit Sub
    I_Receta.CargarDatos 0
    I_Receta.Show 1
Case 24 '------- Salir
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

Private Sub Mover_ListaReceta()

On Error GoTo Man_Error

Dim i As Long
fg_carga ""
vaSpread1(0).MaxRows = 0
itab = 0
vaSpread1(0).Visible = False
i = 1

'------- Mover encabezado recetas
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open RutinaLectura.Receta(1, 0, vg_filcatdie, vg_filtippla, "", 0), vg_db, adOpenStatic

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      grdAddRow vaSpread1(0)
               
      grdCellTypeStatic vaSpread1(0), 1, i, 1
      grdSetText vaSpread1(0), 1, i, IIf(IsNull(RS!rec_codigo), 0, RS!rec_codigo)

      grdCellTypeStatic vaSpread1(0), 2, i, 0
      grdSetText vaSpread1(0), 2, i, IIf(IsNull(RS!rec_nombre), "", RS!rec_nombre)

      RS.MoveNext: i = i + 1
   
   Loop
   
   modo = "M"
   vaSpread1(0).Row = 1
   vaSpread1(0).Col = 1
   codigo = vaSpread1(0).text
   
   vaSpread1(0).SetActiveCell 1, 1
   
   Label1(1).Visible = True
   fpTnombre.Visible = True
   Label1(10).Visible = True:
   If vg_modrec = True Then Gl_Ac_Botones Me, 3, 1, modo Else Gl_Ac_Botones Me, 3, 4, modo
   
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
   SSTab1.TabEnabled(3) = True
   SSTab1.TabEnabled(4) = True
   SSTab1.TabEnabled(5) = True
   
   '------- Grabar Categoría Dietética por Defecto Recetas
   If vg_filcatdie > -1 Then
   
      vg_db.BeginTrans
      
      vg_db.Execute "UPDATE a_param SET par_valor = '" & vg_filcatdie & "' WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'catdefecto'"
      
      vg_db.CommitTrans
      
   End If
   
Else
   
   Label1(1).Visible = True
   fpTnombre.Visible = True
   Label1(10).Visible = False
   
   SSTab1.Tab = 0
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
   SSTab1.TabEnabled(3) = False
   SSTab1.TabEnabled(4) = False
   
   If vg_modrec = True Then Gl_Ac_Botones Me, 3, 2, modo Else Gl_Ac_Botones Me, 3, 4, modo

End If

RS.Close
Set RS = Nothing
Label1(10).Caption = Format(vaSpread1(0).MaxRows, fg_Pict(7, 0)) & " Registros"
vaSpread1(0).Visible = True
fpTnombre.text = ""
SSTab1.Tab = 0
fg_descarga

Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub LlenaCombo(Index, codreceta As Long)

On Error GoTo Man_Error

Dim RS  As New ADODB.Recordset
Dim Sql As String

'Ini : Carga Alergeno
Alergeno(0).Clear
Alergeno(1).Clear
Alergeno(2).Clear

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = " SELECT  ae.IdAlergeno , isnull(ae.NombreAlergeno, '') as NombreAlergeno, " & _
      "      CASE WHEN ISNULL(bro.IdReceta, 0) = 0 THEN 0 " & _
      "           ELSE 1 " & _
      "      END As Selected " & _
      "FROM    dbo.a_Alergeno AS ae WITH ( NOLOCK ) " & _
            " INNER JOIN dbo.b_recetaAlergeno AS bro WITH ( NOLOCK ) ON ae.IdAlergeno = bro.IdAlergeno " & _
                                                                        " AND bro.IdReceta = " & codreceta & " " & _
                                                                        " and isnull(bro.activo,'') = '1' " & _
                                                                        " AND isnull(ae.Activo,'') = '1' " & _
    "Union All " & _
    "SELECT  ae.IdAlergeno , " & _
    "        isnull(ae.NombreAlergeno, '') as NombreAlergeno, " & _
    "        0 AS selected " & _
    "FROM    dbo.a_Alergeno AS ae WITH ( NOLOCK ) " & _
    "WHERE   isnull(ae.Activo,'') = '1' " & _
    "AND NOT EXISTS ( SELECT 1 " & _
                     "FROM   b_recetaAlergeno AS bro " & _
                     "WHERE  bro.IdReceta = " & codreceta & " " & _
                     "AND    bro.idAlergeno = ae.idAlergeno " & _
                     "AND    isnull(bro.Activo,'') = '1' )"

Set RS = vg_db.Execute(Sql)
contador = 0
         
Do While Not RS.EOF
         
   Alergeno(0).AddItem RS("NombreAlergeno") & Space(150) & RS("IdAlergeno")
   If RS("selected") = 1 Then Alergeno(0).Selected(contador) = True
          
   Alergeno(1).AddItem RS("NombreAlergeno") & Space(150) & RS("IdAlergeno")
   If RS("selected") = 1 Then Alergeno(1).Selected(contador) = True
          
   Alergeno(2).AddItem RS("NombreAlergeno") & Space(150) & RS("IdAlergeno")
   If RS("selected") = 1 Then Alergeno(2).Selected(contador) = True
          
    RS.MoveNext
    contador = contador + 1
    
Loop
RS.Close
Set RS = Nothing
'Fin : Carga Alergeno

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub MoverDetalleDatos()

Dim grnetverd As Double, canpavb As Double

fg_carga ""

est = True
itexto = 1
LimpiarVariable

LlenaCombo 0, codigo

If vg_newcodrec > 0 Then
   codigo = vg_newcodrec
Else
   vaSpread1(0).Row = vaSpread1(0).ActiveRow
   vaSpread1(0).Col = 1
   If vaSpread1(0).Row < 1 Then vaSpread1(0).Row = 1
   codigo = Val(vaSpread1(0).Value)
End If
If vg_newcodrec > 0 Then codigo = vg_newcodrec

'------- Lectura de encabezado recetas
fpDouble1(0).Value = 1
fpDouble1(6).Value = 1
fpDouble1(10).Value = 1


RS.Open RutinaLectura.Receta(2, codigo, 0, 0, "", 0), vg_db, adOpenStatic
Frame1(3).Caption = codigo
Frame1(5).Caption = codigo
Frame1(8).Caption = codigo
If Not RS.EOF Then
   fpText1(0).text = Trim(RS!rec_nombre)
   fpText1(1).text = Trim(RS!rec_nomfan)
   fpText1(2).text = Trim(RS!rec_nombre)
   fpText1(3).text = Trim(RS!rec_nomfan)
   fpText1(4).text = Trim(RS!rec_nombre)
   fpText1(5).text = Trim(RS!rec_nomfan)
   fpayuda(2).Caption = "": fpayuda(2).Caption = fg_BuscaenArbol(RS!rec_catdie, "a_recetacatdie", "car_codigo")
   fpayuda(3).Caption = "": fpayuda(3).Caption = fg_BuscaenArbol(RS!rec_tippla, "a_recetatippla", "tip_codigo")
   fpayuda(1).Caption = "": fpayuda(1).Caption = fg_BuscaenArbol(RS!rec_catdie, "a_recetacatdie", "car_codigo")
   fpayuda(0).Caption = "": fpayuda(0).Caption = fg_BuscaenArbol(RS!rec_tippla, "a_recetatippla", "tip_codigo")
   fpayuda(4).Caption = "": fpayuda(4).Caption = fg_BuscaenArbol(RS!rec_catdie, "a_recetacatdie", "car_codigo")
   fpayuda(5).Caption = "": fpayuda(5).Caption = fg_BuscaenArbol(RS!rec_tippla, "a_recetatippla", "tip_codigo")
   fpDouble1(0).Value = IIf(IsNull(RS!rec_basrac), 1, RS!rec_basrac)
   fpDouble1(6).Value = IIf(IsNull(RS!rec_basrac), 1, RS!rec_basrac)
   fpDouble1(10).Value = IIf(IsNull(RS!rec_basrac), 1, RS!rec_basrac)
   codcatdie = RS!rec_catdie: codtipplato = RS!rec_tippla
   
End If
RS.Close: Set RS = Nothing

'------- Lectura detalle recetas
cosrec = 0: grnetverd = 0: canpavb = 0
If vg_newestrec = False Or vg_auxtiprec <> vg_tiprec Then
   RS.Open "SELECT b.red_codigo, b.red_nroite, b.red_codpro, b.red_canpro, " & _
           "b.red_cospro, b.red_pctapr, b.red_pctcoc, b.red_pctnut, " & _
           "((b.red_pctnut/100)*(b.red_canpro)) AS canneta, " & _
           "((b.red_pctnut/100)*(b.red_canpro)) AS cangvneta, " & _
           "(b.red_canpro*isnull(contlis.cpi_precos,0)) AS precos, " & _
           "(((b.red_pctapr/100)*b.red_canpro)*(b.red_pctcoc/100)) AS canservida, " & _
           "c.ing_nombre, c.ing_indgrv, d.unm_codigo, d.unm_nomcor, isnull(contlis.pro_codtip,0) as pro_codtip " & _
           "FROM  b_receta a " & _
           "inner join b_recetadet b on b.red_codigo = a.rec_codigo " & _
           "inner join b_ingrediente c on b.red_codpro = c.ing_codigo " & _
           "inner join a_unidadmed d on c.ing_unimed = d.unm_codigo " & _
           "left join (select f.cpi_codped, f.cpi_coding, f.cpi_precos, e.pro_codtip from b_contlistpreing f inner join b_productos e on f.cpi_codped = e.pro_codigo and f.cpi_cencos = '" & MuestraCasino(1) & "') as contlis on c.ing_codigo = contlis.cpi_coding " & _
           "WHERE b.red_codigo = " & codigo & " AND b.red_tiprec = " & vg_tiprec & " AND b.red_cencos = '" & IIf(vg_tiprec = 0, 0, MuestraCasino(1)) & "'", vg_db, adOpenStatic

Else
   RS.Open "SELECT b.red_codigo, b.red_nroite, b.red_codpro, b.red_canpro, " & _
           "b.red_cospro, b.red_pctapr, b.red_pctcoc, b.red_pctnut, " & _
           "((b.red_pctnut/100)*(b.red_canpro)) AS canneta, " & _
           "((b.red_pctnut/100)*(b.red_canpro)) AS cangvneta, " & _
           "(b.red_canpro*e.mic_cospro) AS precos, " & _
           "(((b.red_pctapr/100)*b.red_canpro)*(b.red_pctcoc/100)) AS canservida, " & _
           "c.ing_nombre, c.ing_indgrv, d.unm_codigo, d.unm_nomcor, isnull(contlis.pro_codtip,0) as pro_codtip " & _
           "FROM  b_receta a " & _
           "inner join b_recetadet b on b.red_codigo = a.rec_codigo " & _
           "inner join b_ingrediente c on c.ing_codigo = b.red_codpro " & _
           "inner join a_unidadmed d on c.ing_unimed = d.unm_codigo " & _
           "left join b_minutacosto e on b.red_codpro = e.mic_codpro and e.mic_codpro = c.ing_codigo AND   e.mic_cencos = '" & MuestraCasino(1) & "' AND   e.mic_fecval = " & vg_fecval & " AND   e.mic_tipmin = '" & vg_opcion & "' " & _
           "left join (select g.cpi_codped, g.cpi_coding, g.cpi_precos, f.pro_codtip from b_contlistpreing g inner join b_productos f on g.cpi_codped = f.pro_codigo and g.cpi_cencos = '" & MuestraCasino(1) & "') as contlis on c.ing_codigo = contlis.cpi_coding " & _
           "WHERE b.red_codigo = " & codigo & " " & _
           " " & _
           "AND   b.red_tiprec = " & vg_tiprec & " AND b.red_cencos = '" & IIf(vg_tiprec = 0, 0, MuestraCasino(1)) & "' " & _
           " " & _
           "", vg_db, adOpenStatic

End If
vaSpread1(inddet).Visible = False
vaSpread2(indapo).Visible = False
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1(inddet).Row = RS!red_nroite
      If IsNull(RS!ing_nombre) = False Then
         formatearcelda vaSpread1(inddet).Row, RS!red_codpro, RS!ing_nombre, RS!unm_nomcor, RS!red_canpro, RS!red_pctapr, RS!red_pctcoc, RS!red_pctnut, RS!canservida, RS!canneta, IIf(IsNull(RS!precos), 0, RS!precos), inddet, RS!pro_codtip
         If RS!ing_indgrv = 1 Then grnetverd = CCur(grnetverd + RS!cangvneta)
         cosrec = CCur(IIf(IsNull(RS!precos), 0, RS!precos) + cosrec)
      End If
      RS.MoveNext
   Loop
End If
If vaSpread1(inddet).MaxRows > 0 Then vaSpread1(inddet).Row = 1
RS.Close: Set RS = Nothing
'------- Calcular Aporte de alto valor biologigo
If vg_newestrec = False Or vg_auxtiprec <> vg_tiprec Then
   RS.Open "SELECT DISTINCT b.red_codigo, " & _
           "SUM(((b.red_pctnut/100)*(d.pnu_canapo*b.red_canpro)/c.ing_facnut)) AS canpavb " & _
           "FROM  b_receta a, b_recetadet b, b_ingrediente c, b_productonut d, a_nutriente e " & _
           "WHERE b.red_codigo = a.rec_codigo " & _
           "AND   b.red_codpro = c.ing_codigo " & _
           "AND   c.ing_codigo = d.pnu_codpro " & _
           "AND   d.pnu_codapo = e.nut_codigo " & _
           "AND   b.red_codigo = " & codigo & " " & _
           "AND   b.red_tiprec = " & vg_tiprec & " AND b.red_cencos = '" & IIf(vg_tiprec = 0, 0, MuestraCasino(1)) & "' " & _
           "AND   e.nut_codigo = 3 " & _
           "AND   c.ing_indpav = 1 " & _
           "GROUP BY b.red_codigo", vg_db, adOpenStatic
Else
   RS.Open "SELECT DISTINCT b.red_codpro, " & _
           "SUM(((b.red_pctnut/100)*(d.pnu_canapo*b.red_canpro)/c.ing_facnut)) AS canpavb " & _
           "FROM  b_receta a, b_recetadet b, b_ingrediente c, b_productonut d, a_nutriente e, b_minutacosto f " & _
           "WHERE b.red_codigo = a.rec_codigo " & _
           "AND   b.red_codpro = f.mic_codpro " & _
           "AND   f.mic_codpro = c.ing_codigo " & _
           "AND   c.ing_codigo = d.pnu_codpro " & _
           "AND   d.pnu_codapo = e.nut_codigo " & _
           "AND   b.red_codigo = " & codigo & " " & _
           "AND   b.red_tiprec = " & vg_tiprec & " AND b.red_cencos = '" & IIf(vg_tiprec = 0, 0, MuestraCasino(1)) & "' " & _
           "AND   e.nut_codigo = 3 " & _
           "AND   f.mic_cencos = '" & MuestraCasino(1) & "' " & _
           "AND   f.mic_fecval = " & vg_fecval & " " & _
           "AND   f.mic_tipmin = '" & vg_opcion & "' " & _
           "AND   c.ing_indpav = 1 " & _
           "GROUP BY b.red_codpro", vg_db, adOpenStatic
End If
If Not RS.EOF Then canpavb = Format(RS!canpavb, fg_Pict(6, 2))
RS.Close: Set RS = Nothing
'------- Calcular resumen aportes Nutricionales
If vg_newestrec = False Or vg_auxtiprec <> vg_tiprec Then
   RS.Open "SELECT DISTINCT e.nut_nombre, e.nut_codigo, " & _
           "SUM((((b.red_pctnut/100)*(d.pnu_canapo*(b.red_canpro/a.rec_basrac)))/c.ing_facnut)) AS candiet, " & _
           "e.nut_secnro " & _
           "FROM  b_receta a, b_recetadet b, b_ingrediente c, b_productonut d , a_nutriente e " & _
           "WHERE b.red_codigo = a.rec_codigo " & _
           "AND   b.red_codpro = c.ing_codigo " & _
           "AND   c.ing_codigo = d.pnu_codpro " & _
           "AND   d.pnu_codapo = e.nut_codigo " & _
           "AND   b.red_codigo = " & codigo & " " & _
           "AND   b.red_tiprec = " & vg_tiprec & " AND b.red_cencos = '" & IIf(vg_tiprec = 0, 0, MuestraCasino(1)) & "' " & _
           "AND   c.ing_facnut > 0 " & _
           "GROUP BY e.nut_nombre, e.nut_codigo, e.nut_secnro " & _
           "ORDER BY e.nut_secnro", vg_db, adOpenStatic
Else
   RS.Open "SELECT DISTINCT e.nut_nombre, e.nut_codigo, " & _
           "SUM((((b.red_pctnut/100)*(d.pnu_canapo*(b.red_canpro/a.rec_basrac)))/c.ing_facnut)) AS candiet, " & _
           "e.nut_secnro " & _
           "FROM  b_receta a, b_recetadet b, b_ingrediente c, b_productonut d, a_nutriente e, b_minutacosto f " & _
           "WHERE b.red_codigo = a.rec_codigo " & _
           "AND   b.red_codpro = f.mic_codpro " & _
           "AND   f.mic_codpro = c.ing_codigo " & _
           "AND   c.ing_codigo = d.pnu_codpro " & _
           "AND   d.pnu_codapo = e.nut_codigo " & _
           "AND   b.red_codigo = " & codigo & " " & _
           "AND   b.red_tiprec = " & vg_tiprec & " AND b.red_cencos = '" & IIf(vg_tiprec = 0, 0, MuestraCasino(1)) & "' " & _
           "AND   c.ing_facnut > 0 " & _
           "AND   f.mic_cencos = '" & MuestraCasino(1) & "' " & _
           "AND   f.mic_fecval = " & vg_fecval & " " & _
           "AND   f.mic_tipmin = '" & vg_opcion & "' " & _
           "GROUP BY e.nut_nombre, e.nut_codigo, e.nut_secnro " & _
           "ORDER BY e.nut_secnro", vg_db, adOpenStatic
End If
vaSpread2(indapo).Visible = False
For i = 1 To vaSpread2(indapo).MaxRows
    grdCellTypeStatic vaSpread2(indapo), 3, i, 1
    grdSetText vaSpread2(indapo), 3, i, Format(0, fg_Pict(6, 2))
    grdRowColForeColor vaSpread2(indapo), i, i, 3, 3, &HFF0000
Next i
Dim ind_ini As Long
If Not RS.EOF Then
   i = 1
   Do While Not RS.EOF
      ind_ini = vaSpread2(indapo).SearchCol(1, -1, vaSpread2(indpao).MaxRows, Trim(CStr(RS!nut_codigo)), SearchFlagsEqual)
      If ind_ini > 0 Then
         grdCellTypeStatic vaSpread2(indapo), 3, ind_ini, 1
         grdSetText vaSpread2(indapo), 3, ind_ini, Format(RS!candiet, fg_Pict(6, 2))
         grdRowColForeColor vaSpread2(indapo), ind_ini, ind_ini, 3, 3, &HFF0000
      End If
      If RS!nut_codigo = 3 Then candiet = CCur(candiet + RS!candiet)
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
vaSpread1(inddet).Visible = True
vaSpread2(indapo).Visible = True
Label2(4).Caption = Format(grnetverd, fg_Pict(6, 2))
Label2(5).Caption = Format(0, fg_Pict(6, 2))
If fpDouble1(0).Value > 1 Then
   Label2(5).Caption = Format(CCur(canpavb / fpDouble1(0).Value), fg_Pict(6, 2))
End If
Label2(15).Caption = Format(grnetverd, fg_Pict(6, 2))
Label2(10).Caption = Format(0, fg_Pict(6, 2))
If fpDouble1(6).Value > 0 Then
   Label2(10).Caption = Format(CCur(canpavb / fpDouble1(6).Value), fg_Pict(6, 2))
End If
Label2(20).Caption = Format(grnetverd, fg_Pict(6, 2))
Label2(25).Caption = Format(0, fg_Pict(6, 2))
If fpDouble1(10).Value > 0 Then
   Label2(25).Caption = Format(CCur(canpavb / fpDouble1(10).Value), fg_Pict(6, 2))
End If

If candiet > 0 Then
   Label2(7).Caption = Format(CCur(((canpavb / fpDouble1(0).Value) / candiet) * 100), fg_Pict(6, 2))
   Label2(17).Caption = Format(CCur(((canpavb / fpDouble1(6).Value) / candiet) * 100), fg_Pict(6, 2))
   Label2(18).Caption = Format(CCur(((canpavb / fpDouble1(10).Value) / candiet) * 100), fg_Pict(6, 2))
Else
   Label2(7).Caption = Format(0, fg_Pict(6, 2))
   Label2(17).Caption = Format(0, fg_Pict(6, 2))
   Label2(18).Caption = Format(0, fg_Pict(6, 2))
End If
If cosrec > 0 Then
   Label2(3).Caption = Format(0, fg_Pict(6, 2))
   If fpDouble1(0).Value > 0 Then
      Label2(3).Caption = Format(CCur(cosrec / fpDouble1(0).Value), fg_Pict(6, 2))
   End If
   Label2(14).Caption = Format(0, fg_Pict(6, 2))
   If fpDouble1(6).Value > 0 Then
      Label2(14).Caption = Format(CCur(cosrec / fpDouble1(6).Value), fg_Pict(6, 2))
   End If
   Label2(21).Caption = Format(0, fg_Pict(6, 2))
   If fpDouble1(10).Value > 0 Then
      Label2(21).Caption = Format(CCur(cosrec / fpDouble1(10).Value), fg_Pict(6, 2))
   End If
End If
calnetoservido
itexto = 0
If vg_modrec = False Then
   For i = 1 To vaSpread1(inddet).MaxRows
        vaSpread1(inddet).Row = i
        vaSpread1(inddet).Col = 1
        If Trim(vaSpread1(inddet).text) = "" Then
           vaSpread1(inddet).Row2 = 40
           vaSpread1(inddet).Col2 = 10
           vaSpread1(inddet).BlockMode = True
           ' Lock cells
           vaSpread1(inddet).Lock = True
           ' Protect the cells from being edited
           vaSpread1(inddet).Protect = True
           ' Turn block mode off
           vaSpread1(inddet).BlockMode = False
        End If
   Next i
End If
est = False

If vg_newestrec = True Or ("S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) And vg_5etapas) Then
   Frame1(0).Enabled = False
   Frame1(3).Enabled = False
'   Frame1(4).Enabled = False
   Frame1(5).Enabled = False
   Toolbar2.Enabled = False
   Toolbar4.Enabled = False
   Toolbar5.Enabled = False
End If
fg_descarga
End Sub

Private Sub LimpiarVariable()
fpText1(0).text = "": fpText1(1).text = "": fpText1(2).text = "": fpText1(3).text = "": fpText1(4).text = "": fpText1(5).text = ""
fpDouble1(1).Value = "": fpDouble1(2).Value = "": fpDouble1(3).Value = ""
fpDouble1(6).Value = "": fpDouble1(7).Value = "": fpDouble1(5).Value = "": fpDouble1(4).Value = ""
fpDouble1(10).Value = "": fpDouble1(11).Value = "": fpDouble1(9).Value = "": fpDouble1(8).Value = ""
fpayuda(2).Caption = "": fpayuda(3).Caption = "": fpayuda(1).Caption = "": fpayuda(0).Caption = "": fpayuda(4).Caption = "": fpayuda(5).Caption = ""
vaSpread1(inddet).MaxRows = 0: vaSpread1(inddet).MaxRows = 40
Label2(3).Caption = Format(0, fg_Pict(6, 2)): Label2(4).Caption = Format(0, fg_Pict(6, 2))
Label2(5).Caption = Format(0, fg_Pict(6, 2)): Label2(7).Caption = Format(0, fg_Pict(6, 2))
Label2(15).Caption = Format(0, fg_Pict(6, 2)): Label2(10).Caption = Format(0, fg_Pict(6, 2))
Label2(17).Caption = Format(0, fg_Pict(6, 2)): Label2(14).Caption = Format(0, fg_Pict(6, 2))
Label2(20).Caption = Format(0, fg_Pict(6, 2)): Label2(25).Caption = Format(0, fg_Pict(6, 2))
Label2(18).Caption = Format(0, fg_Pict(6, 2)): Label2(21).Caption = Format(0, fg_Pict(6, 2))
With vaSpread2(indapo)
    For i = 1 To .MaxRows
       .Row = i
       .Col = 3
       .CellType = 5
       .TypeHAlign = 1
       .text = Format(0, fg_Pict(6, 2))
       .ForeColor = &HFF0000
    Next i
End With
candiet = 0: codcatdie = 0: codtipplato = 0
If vg_modrec = False Then Frame1(3).Enabled = False: Frame1(5).Enabled = False: Frame1(8).Enabled = False
End Sub

Function Hab_Des(op As Integer)
With SSTab1
    Select Case op
    Case 0
        If modo = "A" Or modo = "M" Then
           .TabEnabled(0) = False
           If .Tab = 1 Or .Tab = 0 Or .Tab = 2 Or .Tab = 3 Then
              If vg_modrec And vg_newcodrec < 1 Then
                 If vg_tiprec = 0 Then
                    .TabEnabled(0) = False
                    .TabEnabled(2) = False
                    .TabEnabled(3) = False
                    .TabEnabled(4) = False
                    .TabEnabled(5) = False
                 ElseIf vg_tiprec = -1 Then
                    .TabEnabled(0) = False
                    .TabEnabled(1) = False
                    .TabEnabled(3) = False
                    .TabEnabled(4) = False
                    .TabEnabled(5) = False
                 ElseIf vg_tiprec > 0 Then
                    .TabEnabled(0) = False
                    .TabEnabled(1) = False
                    .TabEnabled(2) = False
                    .TabEnabled(4) = False
                    .TabEnabled(5) = False
                 End If
                 Exit Function
              End If
              If vg_newcodrec > 0 Then
                 If vg_tiprec = 0 Then
                    .TabEnabled(1) = True: .TabEnabled(2) = True: .TabEnabled(3) = True
                    .Tab = 1
                 ElseIf vg_tiprec = -1 Then
                    If est Then .TabEnabled(3) = True Else .TabEnabled(3) = False
                    .TabEnabled(1) = True: .TabEnabled(2) = True
                    .Tab = 2
                 ElseIf vg_tiprec > 0 Then
                    If est Then .TabEnabled(2) = True Else .TabEnabled(2) = False
                    .TabEnabled(1) = True: .TabEnabled(3) = True
                    .Tab = 3
                 End If
              Else
                 If vg_tiprec = 0 Then
                    .TabEnabled(0) = False
                    .TabEnabled(1) = True
                    .TabEnabled(2) = True
                 ElseIf vg_tiprec = -1 Then
                    .TabEnabled(0) = False
                    .TabEnabled(3) = False
                    .TabEnabled(2) = True
                    .TabEnabled(1) = True
                 ElseIf vg_tiprec > 0 Then
                    .TabEnabled(0) = False
                    .TabEnabled(2) = False
                    .TabEnabled(3) = True
                    .TabEnabled(1) = True
                 End If
              End If
    '          .TabEnabled(4) = False
           ElseIf .Tab = 4 Then
              .TabEnabled(1) = False
              .TabEnabled(2) = False
              .TabEnabled(3) = False
              .TabEnabled(5) = False
           ElseIf .Tab = 5 Then
              .TabEnabled(1) = False
              .TabEnabled(2) = False
              .TabEnabled(3) = False
              .TabEnabled(4) = False
           End If
           fpTnombre.Enabled = False
        End If
    Case 1
        .TabEnabled(0) = True
        If .Tab = 1 Or .Tab = 0 Or .Tab = 2 Or .Tab = 3 Then
           If vg_tiprec = 0 Then .TabEnabled(1) = True: .TabEnabled(2) = True: .TabEnabled(3) = True Else .TabEnabled(1) = True: .TabEnabled(2) = True: .TabEnabled(3) = True
           .TabEnabled(4) = True: .TabEnabled(5) = True
        ElseIf .Tab = 4 Then
           .TabEnabled(1) = True: .TabEnabled(2) = True: .TabEnabled(3) = True: .TabEnabled(5) = True
        ElseIf .Tab = 5 Then
           .TabEnabled(1) = True: .TabEnabled(2) = True: .TabEnabled(3) = True: .TabEnabled(4) = True
        End If
        fpTnombre.Enabled = True
    Case 2
        .TabEnabled(0) = True: fpTnombre.Enabled = True: .TabEnabled(1) = False: .TabEnabled(2) = False
    Case 3
        .Tab = 0: .TabEnabled(0) = True: .TabEnabled(1) = False: .TabEnabled(2) = False: .TabEnabled(3) = False ': .TabEnabled(4) = False
    End Select
End With
End Function

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu
'Case "Copiar Recetas"
'    If vaSpread1(0).MaxRows < 1 Then Exit Sub
'    vg_swpegreceta = 0: vg_codreceta = 0
'    vaSpread1(0).Row = vaSpread1(0).ActiveRow
'    vaSpread1(0).col = 1: vg_codreceta = Val(vaSpread1(0).Value)
'    nombusca = "": vg_swpegreceta = 0
'    M_CpoRec.Show 1
'    Me.Refresh
'    If vg_swpegreceta = 1 Then nombusca = fpTnombre.Text: fpTnombre.Text = "": fpTnombre.Text = nombusca
'Case "Pegar Recetas"
'    If vaSpread1(0).MaxRows < 1 Then Exit Sub
'    vg_swpegreceta = 0: vg_codreceta = 0
'    vaSpread1(0).Row = vaSpread1(0).ActiveRow
'    vaSpread1(0).col = 1: vg_codreceta = Val(vaSpread1(0).Value)
'    M_PegRec.Show 1
'    If SSTab1.TabEnabled(1) = True And vg_swpegreceta = 1 Then MoverDetalleDatos
'Case "Mover Recetas"
'    If vaSpread1(0).MaxRows < 1 Then Exit Sub
'    SSTab1.Tab = 0: vg_swmovreceta = 0
'    M_MovRec.LlenarRecetas Label2(9).Caption, vg_filcatdie, vg_filtippla
'    M_MovRec.Show 1
'    If vg_swmovreceta = 1 And fpTnombre.Text <> "" Then
'       nombusca = fpTnombre.Text: fpTnombre.Text = "": fpTnombre.Text = nombusca
'    ElseIf vg_swmovreceta = 1 And fpTnombre.Text = "" Then
'       fpTnombre.Text = " "
'    End If
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim MaxFila As Long, sql1 As String, sql2 As String, sql3 As String, ind_ini As Long
vg_modrec = IIf(GetParametro("Modrec ") = "1", True, False)
Select Case Button.Index
Case 1
    With vaSpread1(inddet)
        If vg_modrec = False Then
           .Row = .ActiveRow
           .Col = 1
           If Trim(.text) = "" Or .Lock = True Then MsgBox "No esta autorizado, para ingresar ó modificar ingredientes en receta, comuniquese con el administrador", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        End If
        vg_nombre = "": vg_codigo = ""
        vg_left = fpayuda(2).Left + 550
        If vg_modrec = False Then
           Dim fampro As String, codtip As Long, parval As String
           fampro = "": codtip = 0: parval = ""
           .Row = .ActiveRow
           .Col = 1
           sql1 = IIf(vg_tipbase = "1", " trim(str(a.pro_codtip)) ", " ltrim(convert(varchar(20),a.pro_codtip)) ")
           RS1.Open "SELECT par_codigo, par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND " & sql1 & " = 'CamIng'", vg_db, adOpenStatic
           Do While Not RS1.EOF
              sql1 = IIf(vg_tipbase = "1", " mid(par_codigo,1,6) ", " substring(par_codigo,1,6) ")
              RS2.Open "SELECT a.pro_codtip FROM b_productos a, b_productosing b " & _
                       "WHERE  b.pri_codpro = a.pro_codigo " & _
                       "AND  " & sql1 & " IN ('" & fg_CambiaChar(GetParametro(RS1!par_codigo), ";", "','") & "') " & _
                       "AND   b.pri_coding = '" & Trim(.text) & "'", vg_db, adOpenStatic
              If Not RS2.EOF Then
                 parval = RS1!par_valor
                 codtip = RS2!pro_codtip
                 fampro = fg_buscarcodtip(RS1!par_valor & ";", Trim(Str(RS2!pro_codtip)))
                 fampro = fg_CambiaChar(fampro, ";", "','")
                 RS2.Close: Set RS2 = Nothing
                 Exit Do
              End If
              RS2.Close: Set RS2 = Nothing
              RS1.MoveNext
           Loop
           RS1.Close: Set RS1 = Nothing
    '       RS2.Open "SELECT a.pro_codtip FROM b_productos a, b_productosing b " & _
    '                "WHERE  b.pri_codpro=a.pro_codigo " & _
    '                "AND  (trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvfre"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvpre"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvcon"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne1"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne2"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne3"), ";", "','") & "')) " & _
    '                "AND   b.pri_coding='" & Trim(vaSpread1(inddet).text) & "'", vg_db, adOpenStatic
    '       If RS2.EOF Then RS2.Close: Set RS2 = Nothing: Exit Sub
    '       If RS2!pro_codtip = (fg_CambiaChar(GetParametro("fvfre"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("fvpre"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("fvcon"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("fvpre"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("fvfre"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("fvcon"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("fvcon"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("fvpre"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("fvfre"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("carne1"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("carne1"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("carne2"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("carne2"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("carne2"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("carne1"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("carne3"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("carne1"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("carne2"), ";", "','")) & "', '" & (fg_CambiaChar(GetParametro("carne3"), ";", "','"))
    '       End If
    '       RS2.Close: Set RS2 = Nothing
           B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", fampro
        ElseIf vg_modrec = True Then
           B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "Gen"
        End If
        B_TabEst.Show 1
        If vg_codigo = "" Then Exit Sub
        If vg_codigo <> "" And modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
        .Row = .ActiveRow
        .Col = 3
        If Val(.text) > 0 Then
           canpro1 = 0: codpro1 = "": pctnut1 = 0
           .Col = 1: codpro1 = .text
           .Col = 3: canpro1 = .text
           .Col = 8: pctnut1 = .text
           .Col = 10: cospro = .text
           '------- Resta Aporte
           sql2 = IIf(vg_tipbase = "1", " (((((" & pctnut1 & "/100)*(c.pnu_canapo*(" & canpro1 & "/" & Val(fpDouble1(0).Value) & ")))/a.ing_facnut))) as  canneta ", " (((((convert(float," & pctnut1 & ")/100)*(c.pnu_canapo*(convert(float," & canpro1 & ")/ convert(int," & fpDouble1(0).Value & ")))/a.ing_facnut))) AS  canneta ")
           sql3 = IIf(vg_tipbase = "1", " (((" & pctnut1 & "/100)*" & canpro1 & ")) AS cangrverneto  ", " (((convert(float," & pctnut1 & ")/100)*convert(float," & canpro1 & "))) AS cangrverneto  ")
           RS.Open "SELECT DISTINCT b.nut_nombre, b.nut_codigo, a.ing_indgrv, a.ing_indpav, b.nut_secnro, " & _
                   "" & sql2 & ", " & _
                   "" & sql3 & " " & _
                   "FROM  b_ingrediente a, a_nutriente b, b_productonut c " & _
                   "WHERE a.ing_codigo = c.pnu_codpro " & _
                   "AND   c.pnu_codapo = b.nut_codigo " & _
                   "AND   a.ing_codigo = '" & codpro1 & "' " & _
                   "ORDER BY b.nut_secnro", vg_db, adOpenStatic
           vaSpread2(indapo).Visible = False
           If Not RS.EOF Then
              i = 1
              If RS!ing_indgrv = 1 Then Label2(4).Caption = Format(CCur(Label2(4).Caption - RS!cangrverneto), fg_Pict(6, 2))
              Do While Not RS.EOF
                 If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then Label2(5).Caption = Format(CCur(Label2(5).Caption - RS!canneta), fg_Pict(6, 2))
                 ind_ini = vaSpread2(indapo).SearchCol(1, -1, vaSpread2(indpao).MaxRows, Trim(CStr(RS!nut_codigo)), SearchFlagsEqual)
                 If ind_ini > 0 Then
                    vaSpread2(indapo).Row = ind_ini
                    vaSpread2(indapo).Col = 3
                    vaSpread2(indapo).text = Format(CCur(vaSpread2(indapo).text - RS!canneta), fg_Pict(6, 2))
                 End If
                 RS.MoveNext ': i = i + 1
              Loop
              Label2(3).Caption = Format(CCur(Label2(3).Caption - cospro), fg_Pict(6, 2))
              CalTotalPavb
           End If
           RS.Close: Set RS = Nothing
           vaSpread2(indapo).Visible = True
        End If
        RS.Open "SELECT a.*, b.* FROM  b_ingrediente a, a_unidadmed b " & _
                "WHERE a.ing_unimed=b.unm_codigo " & _
                "AND   a.ing_codigo='" & vg_codigo & "'", vg_db, adOpenStatic
        If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1(inddet).text = "": Exit Sub
        formatearcelda .Row, RS!ing_codigo, RS!ing_nombre, RS!unm_nomcor, 0, RS!ing_pctapr, RS!ing_pctcoc, RS!ing_pctnut, 0, 0, 0, inddet, 0
        RS.Close: Set RS = Nothing
        Me.Refresh
        .Row = .ActiveRow
        .SetActiveCell 3, .Row
        .SetFocus
    End With
    calnetoservido
Case 2
    If vg_modrec = False Then Exit Sub
    With vaSpread1(inddet)
        If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = False Or Toolbar1.Buttons(12).Visible = False Then Hab_Des 0
        For i = 1 To .MaxRows
            .Row = i: .Col = 1
            If .text <> "" Then MaxFila = .Row
        Next i
        .Row = .ActiveRow
        .Col = .ActiveCol
        If MaxFila < .MaxRows Then
           MaxFila = MaxFila + 1
        Else
           Exit Sub
        End If
        '------- Insertar columna
        If .Row + 1 < 41 Then
           .MoveRange 1, (.ActiveRow), .MaxCols, (.MaxRows - 1), 1, (.Row + 1)
           .ClearRange 1, .ActiveRow, .MaxCols, .ActiveRow, False
        End If
    End With
Case 3
    If vg_modrec = False Then Exit Sub
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
    '------- Resta aporte
    With vaSpread1(inddet)
        .Row = .ActiveRow
        .Col = 1
        If .text = "" Then .DeleteRows .Row, 1: .MaxRows = .MaxRows - 1: .MaxRows = .MaxRows + 1: wsmaxfilas = wsmaxfilas - 1:  .Row = .ActiveRow: Exit Sub
        canpro1 = 0: codpro1 = "": pctnut1 = 0
        .Col = 1: codpro1 = .text
        .Col = 3: canpro1 = .text
        .Col = 8: pctnut1 = .text
        If canpro1 > 0 Then
           sql2 = IIf(vg_tipbase = "1", " (((((" & pctnut1 & "/100)*(c.pnu_canapo*(" & canpro1 & "/" & Val(fpDouble1(0).Value) & ")))/a.ing_facnut))) as  canneto ", " (((((convert(float," & pctnut1 & ")/100)*(c.pnu_canapo*(convert(float," & canpro1 & ")/convert(int," & fpDouble1(0).Value & "))))/a.ing_facnut))) AS  canneto ")
           sql3 = IIf(vg_tipbase = "1", " (((" & pctnut1 & "/100)*" & canpro1 & ")) as cangrverneto ", " (((convert(float," & pctnut1 & ")/100)*convert(float," & canpro1 & "))) AS cangrverneto ")
           RS.Open "SELECT DISTINCT b.nut_nombre, b.nut_codigo, a.ing_indgrv, a.ing_indpav, b.nut_secnro, " & _
                 "" & sql2 & ", " & _
                 "" & sql3 & " " & _
                 "FROM  b_ingrediente a, a_nutriente b, b_productonut c " & _
                 "WHERE a.ing_codigo = c.pnu_codpro " & _
                 "AND   c.pnu_codapo = b.nut_codigo " & _
                 "AND   a.ing_codigo = '" & codpro1 & "' " & _
                 "ORDER BY b.nut_secnro", vg_db, adOpenStatic
           If Not RS.EOF Then
              i = 1
              If RS!ing_indgrv = 1 Then Label2(4).Caption = Format(CCur(Label2(4).Caption - RS!cangrverneto), fg_Pict(6, 2))
              Do While Not RS.EOF
                 If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then Label2(5).Caption = Format(CCur(Label2(5).Caption - RS!canneto), fg_Pict(6, 2))
                 vaSpread2(indapo).Row = i
                 vaSpread2(indapo).Col = 3
                 vaSpread2(indapo).text = Format(CCur(vaSpread2(indapo).text - RS!canneto), fg_Pict(6, 2))
                 RS.MoveNext: i = i + 1
              Loop
           End If
           RS.Close: Set RS = Nothing
           '------- Calcular aportes pavb
           CalTotalPavb
        End If
        .DeleteRows .Row, 1
        .MaxRows = .MaxRows - 1
        .MaxRows = .MaxRows + 1
        cosrec = 0
        For i = 1 To .MaxRows
            .Row = i: .Col = 10
            If .text <> "" Then cosrec = CCur(cosrec + .text)
        Next i
        Label2(3).Caption = Format((cosrec), fg_Pict(6, 2))
        wsmaxfilas = wsmaxfilas - 1
        .Row = .ActiveRow
    End With
    calnetoservido
Case 4
    If vg_modrec = False Then Exit Sub
    With vaSpread1(inddet)
        .Row = .ActiveRow
        .Col = .ActiveCol
        If .Row > 1 Then
           If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
           '------- Copiar datos ultima fila
           .MaxRows = .MaxRows + 1
           .MoveRange 1, (.ActiveRow - 1), .MaxCols, (.ActiveRow - 1), 1, .MaxRows
           '------- Copiar datos fila seleccionada
           .ClearRange 1, (.ActiveRow + 1), .MaxCols, (.ActiveRow - 1), False
           .MoveRange 1, (.ActiveRow), .MaxCols, (.ActiveRow), 1, (.ActiveRow - 1)
           '------- Devolver datos fila y restar ultima fila
           .ClearRange 1, .ActiveRow, .MaxCols, .ActiveRow, False
           .MoveRange 1, .MaxRows, .MaxCols, .MaxRows, 1, .ActiveRow
           .MaxRows = .MaxRows - 1
           .Row = .ActiveRow - 1
           .Col = 2
           .SetActiveCell .Col, .Row
        End If
    End With
Case 5
    If vg_modrec = False Then Exit Sub
    With vaSpread1(inddet)
        .Row = .ActiveRow
        .Col = .ActiveCol
        If .Row + 1 < 41 Then
           '------- Copiar datos ultima fila
           If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
           .MaxRows = .MaxRows + 1
           .MoveRange 1, (.ActiveRow + 1), .MaxCols, (.ActiveRow + 1), 1, .MaxRows
           '------- Copiar datos fila seleccionada
           .ClearRange 1, (.ActiveRow + 1), .MaxCols, (.ActiveRow + 1), False
           .MoveRange 1, (.ActiveRow), .MaxCols, (.ActiveRow), 1, (.ActiveRow + 1)
           '------- Devolver datos fila y restar ultima fila
           .ClearRange 1, .ActiveRow, .MaxCols, .ActiveRow, False
           .MoveRange 1, .MaxRows, .MaxCols, .MaxRows, 1, .ActiveRow
           .MaxRows = .MaxRows - 1
           .Row = .ActiveRow + 1
           .Col = 2
           .SetActiveCell .Col, .Row
        End If
    End With
End Select
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim izquierda As Integer
izquierda = 0
Select Case Button.Index
Case 1
    Toolbar3.Buttons(1).Value = 1: Toolbar3.Buttons(1).Value = 0
    If RichTextBox1(0).SelBold = True Then
       RichTextBox1(0).SelBold = False
       Toolbar3.Buttons(1).Value = 1
       Toolbar3.Buttons(1).Value = 0
    ElseIf RichTextBox1(0).SelBold = False Or IsNull(RichTextBox1(0).SelBold) Then
       RichTextBox1(0).SelBold = True
       Toolbar3.Buttons(1).Value = 0
       Toolbar3.Buttons(1).Value = 1
    End If
Case 2
    If RichTextBox1(0).SelItalic = True Then
       RichTextBox1(0).SelItalic = False
       Toolbar3.Buttons(2).Value = 1
       Toolbar3.Buttons(2).Value = 0
    ElseIf RichTextBox1(0).SelItalic = False Or IsNull(RichTextBox1(0).SelBold) Then
       RichTextBox1(0).SelItalic = True
       Toolbar3.Buttons(2).Value = 0
       Toolbar3.Buttons(2).Value = 1
    End If
Case 3
    Toolbar3.Buttons(3).Value = 1
    If RichTextBox1(0).SelUnderline = True Then
       RichTextBox1(0).SelUnderline = False
       Toolbar3.Buttons(3).Value = 1
       Toolbar3.Buttons(3).Value = 0
    ElseIf RichTextBox1(0).SelUnderline = False Or IsNull(RichTextBox1(0).SelUnderline) Then
       RichTextBox1(0).SelUnderline = True
       Toolbar3.Buttons(3).Value = 0
       Toolbar3.Buttons(3).Value = 1
    End If
Case 5
    Toolbar3.Buttons(6).Value = 1: Toolbar3.Buttons(6).Value = 0
    Toolbar3.Buttons(5).Value = 1
    Toolbar3.Buttons(7).Value = 1: Toolbar3.Buttons(7).Value = 0
    Toolbar3.Buttons(8).Value = 1: Toolbar3.Buttons(8).Value = 0
    If izquierda = 0 Then
       Toolbar3.Buttons(5).Value = 1
       Toolbar3.Buttons(5).Value = 0
       Toolbar3.Buttons(8).Value = 0
       Toolbar3.Buttons(8).Value = 1
       izquierda = 1
    Else
       Toolbar3.Buttons(5).Value = 0
       Toolbar3.Buttons(5).Value = 1
       Toolbar3.Buttons(8).Value = 1
       Toolbar3.Buttons(8).Value = 0
       izquierda = 0
    End If
    RichTextBox1(0).SelAlignment = 0
Case 6
    izquierda = 1
    Toolbar3.Buttons(5).Value = 1: Toolbar3.Buttons(5).Value = 0
    Toolbar3.Buttons(6).Value = 1
    Toolbar3.Buttons(7).Value = 1: Toolbar3.Buttons(7).Value = 0
    Toolbar3.Buttons(8).Value = 1: Toolbar3.Buttons(8).Value = 0
    If RichTextBox1(0).SelAlignment = 2 Then
       RichTextBox1(0).SelAlignment = 0
       Toolbar3.Buttons(6).Value = 1
       Toolbar3.Buttons(6).Value = 0
       Toolbar3.Buttons(5).Value = 0
       Toolbar3.Buttons(5).Value = 1
    Else
       RichTextBox1(0).SelAlignment = 2
       Toolbar3.Buttons(6).Value = 0
       Toolbar3.Buttons(6).Value = 1
    End If
Case 7
    izquierda = 1
    Toolbar3.Buttons(5).Value = 1: Toolbar3.Buttons(5).Value = 0
    Toolbar3.Buttons(7).Value = 1
    Toolbar3.Buttons(6).Value = 1: Toolbar3.Buttons(6).Value = 0
    Toolbar3.Buttons(8).Value = 1: Toolbar3.Buttons(8).Value = 0
    If RichTextBox1(0).SelAlignment = 1 Then
       RichTextBox1(0).SelAlignment = 0
       Toolbar3.Buttons(7).Value = 1
       Toolbar3.Buttons(7).Value = 0
       Toolbar3.Buttons(5).Value = 0
       Toolbar3.Buttons(5).Value = 1
    Else
       RichTextBox1(0).SelAlignment = 1
       Toolbar3.Buttons(7).Value = 0
       Toolbar3.Buttons(7).Value = 1
    End If
Case 8
    Toolbar3.Buttons(6).Value = 1: Toolbar3.Buttons(6).Value = 0
    Toolbar3.Buttons(8).Value = 1
    Toolbar3.Buttons(7).Value = 1: Toolbar3.Buttons(7).Value = 0
    Toolbar3.Buttons(5).Value = 1: Toolbar3.Buttons(8).Value = 0
    If izquierda = 1 Then
       Toolbar3.Buttons(5).Value = 1
       Toolbar3.Buttons(5).Value = 0
       Toolbar3.Buttons(8).Value = 0
       Toolbar3.Buttons(8).Value = 1
       izquierda = 0
    Else
       Toolbar3.Buttons(5).Value = 0
       Toolbar3.Buttons(5).Value = 1
       Toolbar3.Buttons(8).Value = 1
       Toolbar3.Buttons(8).Value = 0
       izquierda = 1
    End If
    RichTextBox1(0).SelAlignment = 0
Case 10
    Toolbar3.Buttons(10).Value = 1: Toolbar3.Buttons(10).Value = 0
    If RichTextBox1(0).SelBullet = True Then
       RichTextBox1(0).SelBullet = False
       Toolbar3.Buttons(10).Value = 1
       Toolbar3.Buttons(10).Value = 0
    ElseIf RichTextBox1(0).SelBullet = False Then
       RichTextBox1(0).SelBullet = True
       Toolbar3.Buttons(10).Value = 0
       Toolbar3.Buttons(10).Value = 1
    End If
End Select
End Sub

Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim MaxFila As Long, sql1 As String, sql2 As String, sql3 As String
Dim auxmodrec As Boolean
vg_modrec = IIf(GetParametro("Modrec ") = "1", trae, flase)
Select Case Button.Index
Case 1
    auxmodrec = vg_modrec
    If vg_pais = "CO" And Not vg_modrec Then vg_modrec = True
    With vaSpread1(inddet)
        If vg_modrec = False Then
           .Row = .ActiveRow
           .Col = 1
           If Trim(.text) = "" Or .Lock = True Then MsgBox "No esta autorizado, para ingresar ó modificar ingredientes en receta, comuniquese con el administrador", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        End If
        vg_nombre = "": vg_codigo = ""
        vg_left = fpayuda(2).Left + 550
        Toolbar4.Enabled = False
        If vg_modrec = False Then
           Dim fampro As String, codtip As Long, parval As String
           fampro = "": codtip = 0: parval = ""
           .Row = .ActiveRow
           .Col = 1
    '       RS2.Open "SELECT a.pro_codtip FROM b_productos a, b_productosing b " & _
    '                "WHERE  b.pri_codpro=a.pro_codigo " & _
    '                "AND  (trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvfre"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvpre"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvcon"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne1"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne2"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne3"), ";", "','") & "')) " & _
    '                "AND   b.pri_coding='" & Trim(vaSpread1(inddet).text) & "'", vg_db, adOpenStatic
           sql1 = IIf(vg_tipbase = "1", " mid(par_codigo,1,6) ", " substring(par_codigo,1,6) ")
           RS1.Open "SELECT par_codigo, par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND " & sql1 & " = 'CamIng'", vg_db, adOpenStatic
           sql1 = IIf(vg_tipbase = "1", " trim(str(a.pro_codtip)) ", " ltrim(convert(varchar(20),a.pro_codtip)) ")
           Do While Not RS1.EOF
              RS2.Open "SELECT a.pro_codtip FROM b_productos a, b_productosing b " & _
                       "WHERE  b.pri_codpro = a.pro_codigo " & _
                       "AND  " & sql1 & " IN ('" & fg_CambiaChar(GetParametro(RS1!par_codigo), ";", "','") & "') " & _
                       "AND   b.pri_coding = '" & Trim(.text) & "'", vg_db, adOpenStatic
              If Not RS2.EOF Then
                 parval = RS1!par_valor
                 codtip = RS2!pro_codtip
                 fampro = fg_buscarcodtip(RS1!par_valor & ";", Trim(Str(RS2!pro_codtip)))
                 fampro = fg_CambiaChar(fampro, ";", "','")
                 RS2.Close: Set RS2 = Nothing
                 Exit Do
              End If
              RS2.Close: Set RS2 = Nothing
              RS1.MoveNext
           Loop
           RS1.Close: Set RS1 = Nothing
    '       If RS2!pro_codtip = (fg_CambiaChar(GetParametro("fvfre"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("fvpre"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("fvcon"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("fvpre"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("fvfre"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("fvcon"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("fvcon"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("fvpre"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("fvfre"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("carne1"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("carne1"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("carne2"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("carne2"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("carne2"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("carne1"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("carne3"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("carne1"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("carne2"), ";", "','")) & "', '" & (fg_CambiaChar(GetParametro("carne3"), ";", "','"))
    '       End If
    '       RS2.Close: Set RS2 = Nothing
           B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", fampro
        ElseIf vg_modrec = True Then
           B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "Gen"
        End If
        B_TabEst.Show 1
        If vg_codigo = "" Then Toolbar4.Enabled = True: Exit Sub
        If vg_codigo <> "" And modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
        .Row = .ActiveRow
        .Col = 3
        If Val(.text) > 0 Then
           canpro1 = 0: codpro1 = "": pctnut1 = 0
           .Col = 1: codpro1 = .text
           .Col = 3: canpro1 = .text
           .Col = 8: pctnut1 = .text
           .Col = 10: cospro = .text
           '------- Resta Aporte
           sql2 = IIf(vg_tipbase = "1", " (((((" & pctnut1 & "/100)*(c.pnu_canapo*(" & canpro1 & "/" & Val(fpDouble1(0).Value) & ")))/a.ing_facnut))) AS  canneta ", " (((((convert(float," & pctnut1 & ")/100)*(c.pnu_canapo*(convert(float," & canpro1 & ")/convert(int," & fpDouble1(0).Value & "))))/a.ing_facnut))) AS  canneta ")
           sql3 = IIf(vg_tipbase = "1", " (((" & pctnut1 & "/100)*" & canpro1 & ")) AS cangrverneto ", " (((convert(float," & pctnut1 & ")/100)* convert(float," & canpro1 & "))) AS cangrverneto ")
           RS.Open "SELECT DISTINCT b.nut_nombre, b.nut_codigo, a.ing_indgrv, a.ing_indpav, b.nut_secnro, " & _
                   "" & sql2 & ", " & _
                   "" & sql3 & " " & _
                   "FROM  b_ingrediente a, a_nutriente b, b_productonut c " & _
                   "WHERE a.ing_codigo = c.pnu_codpro " & _
                   "AND   c.pnu_codapo = b.nut_codigo " & _
                   "AND   a.ing_codigo = '" & codpro1 & "' " & _
                   "ORDER BY b.nut_secnro", vg_db, adOpenStatic
           If Not RS.EOF Then
              i = 1
              If RS!ing_indgrv = 1 Then Label2(4).Caption = Format(CCur(Label2(4).Caption - RS!cangrverneto), fg_Pict(6, 2))
              Do While Not RS.EOF
                 If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then Label2(5).Caption = Format(CCur(Label2(5).Caption - RS!canneta), fg_Pict(6, 2))
                 vaSpread2(indapo).Row = i
                 vaSpread2(indapo).Col = 3
                 vaSpread2(indapo).text = Format(CCur(vaSpread2(indapo).text - RS!canneta), fg_Pict(6, 2))
                 i = i + 1
                 RS.MoveNext
              Loop
              Label2(3).Caption = Format(CCur(Label2(3).Caption - cospro), fg_Pict(6, 2))
              CalTotalPavb
           End If
           RS.Close: Set RS = Nothing
        End If
        RS.Open "SELECT a.*, b.* FROM  b_ingrediente a, a_unidadmed b " & _
                "WHERE a.ing_unimed = b.unm_codigo " & _
                "AND   a.ing_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic
        If RS.EOF Then RS.Close: Set RS = Nothing: .text = "": Toolbar4.Enabled = True: Exit Sub
        formatearcelda .Row, RS!ing_codigo, RS!ing_nombre, RS!unm_nomcor, 0, RS!ing_pctapr, RS!ing_pctcoc, RS!ing_pctnut, 0, 0, 0, inddet, 0
        RS.Close: Set RS = Nothing
        Me.Refresh
        .Row = .ActiveRow
        .SetActiveCell 3, .Row
        .SetFocus
    End With
    calnetoservido
    vg_modrec = auxmodrec
    Toolbar4.Enabled = True
Case 2
    If vg_modrec = False Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = False Or Toolbar1.Buttons(12).Visible = False Then Hab_Des 0
    With vaSpread1(inddet)
        For i = 1 To .MaxRows
            .Row = i: .Col = 1
            If .text <> "" Then MaxFila = .Row
        Next i
        .Row = .ActiveRow
        .Col = .ActiveCol
        If MaxFila < .MaxRows Then
           MaxFila = MaxFila + 1
        Else
           Exit Sub
        End If
        '------- Insertar columna
        If .Row + 1 < 41 Then
           .MoveRange 1, (.ActiveRow), .MaxCols, (.MaxRows - 1), 1, (.Row + 1)
           .ClearRange 1, .ActiveRow, .MaxCols, .ActiveRow, False
        End If
    End With
Case 3
    If vg_modrec = False Then Exit Sub
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
    '------- Resta aporte
    With vaSpread1(inddet)
        .Row = .ActiveRow
        .Col = 1
        If .text = "" Then .DeleteRows .Row, 1: .MaxRows = .MaxRows - 1: .MaxRows = .MaxRows + 1: wsmaxfilas = wsmaxfilas - 1: .Row = .ActiveRow: Exit Sub
        canpro1 = 0: codpro1 = "": pctnut1 = 0
        .Col = 1: codpro1 = .text
        .Col = 3: canpro1 = .text
        .Col = 8: pctnut1 = .text
        If canpro1 > 0 Then
           sql2 = IIf(vg_tipbase = "1", " (((((" & pctnut1 & "/100)*(c.pnu_canapo*(" & canpro1 & "/" & Val(fpDouble1(0).Value) & ")))/a.ing_facnut))) AS  canneto ", " (((((convert(float," & pctnut1 & ")/100)*(c.pnu_canapo*(convert(float," & canpro1 & ")/convert(int," & fpDouble1(0).Value & "))))/a.ing_facnut))) AS  canneto ")
           sql3 = IIf(vg_tipbase = "1", " (((" & pctnut1 & "/100)*" & canpro1 & ")) AS cangrverneto ", " (((convert(float," & pctnut1 & ")/100)*convert(float," & canpro1 & "))) AS cangrverneto ")
           RS.Open "SELECT DISTINCT b.nut_nombre, b.nut_codigo, a.ing_indgrv, a.ing_indpav, b.nut_secnro, " & _
                 "" & sql2 & ", " & _
                 "" & sql3 & " " & _
                 "FROM  b_ingrediente a, a_nutriente b, b_productonut c " & _
                 "WHERE a.ing_codigo = c.pnu_codpro " & _
                 "AND   c.pnu_codapo = b.nut_codigo " & _
                 "AND   a.ing_codigo = '" & codpro1 & "' " & _
                 "ORDER BY b.nut_secnro", vg_db, adOpenStatic
           If Not RS.EOF Then
              i = 1
              If RS!ing_indgrv = 1 Then Label2(4).Caption = Format(CCur(Label2(4).Caption - RS!cangrverneto), fg_Pict(6, 2))
              Do While Not RS.EOF
                 If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then Label2(5).Caption = Format(CCur(Label2(5).Caption - RS!canneto), fg_Pict(6, 2))
                 vaSpread2(indapo).Row = i
                 vaSpread2(indapo).Col = 3
                 vaSpread2(indapo).text = Format(CCur(vaSpread2(indapo).text - RS!canneto), fg_Pict(6, 2))
                 RS.MoveNext: i = i + 1
              Loop
           End If
           RS.Close: Set RS = Nothing
           '------- Calcular aportes pavb
           CalTotalPavb
        End If
        .DeleteRows .Row, 1
        .MaxRows = .MaxRows - 1
        .MaxRows = .MaxRows + 1
        cosrec = 0
        For i = 1 To .MaxRows
            .Row = i: .Col = 10
            If .text <> "" Then cosrec = CCur(cosrec + .text)
        Next i
        Label2(3).Caption = Format((cosrec), fg_Pict(6, 2))
        wsmaxfilas = wsmaxfilas - 1
        .Row = .ActiveRow
    End With
    calnetoservido
Case 4
    If vg_modrec = False Then Exit Sub
    With vaSpread1(inddet)
        .Row = .ActiveRow
        .Col = .ActiveCol
        If .Row > 1 Then
           If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
           '------- Copiar datos ultima fila
           .MaxRows = .MaxRows + 1
           .MoveRange 1, (.ActiveRow - 1), .MaxCols, (.ActiveRow - 1), 1, .MaxRows
           '------- Copiar datos fila seleccionada
           .ClearRange 1, (.ActiveRow + 1), .MaxCols, (.ActiveRow - 1), False
           .MoveRange 1, (.ActiveRow), .MaxCols, (.ActiveRow), 1, (.ActiveRow - 1)
           '------- Devolver datos fila y restar ultima fila
           .ClearRange 1, .ActiveRow, .MaxCols, .ActiveRow, False
           .MoveRange 1, .MaxRows, .MaxCols, .MaxRows, 1, .ActiveRow
           .MaxRows = .MaxRows - 1
           .Row = .ActiveRow - 1
           .Col = 2
           .SetActiveCell .Col, .Row
        End If
    End With
Case 5
    If vg_modrec = False Then Exit Sub
    With vaSpread1(inddet)
        .Row = .ActiveRow
        .Col = .ActiveCol
        If .Row + 1 < 41 Then
           '------- Copiar datos ultima fila
           If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
           .MaxRows = .MaxRows + 1
           .MoveRange 1, (.ActiveRow + 1), .MaxCols, (.ActiveRow + 1), 1, .MaxRows
           '------- Copiar datos fila seleccionada
           .ClearRange 1, (.ActiveRow + 1), .MaxCols, (.ActiveRow + 1), False
           .MoveRange 1, (.ActiveRow), .MaxCols, (.ActiveRow), 1, (.ActiveRow + 1)
           '------- Devolver datos fila y restar ultima fila
           .ClearRange 1, .ActiveRow, .MaxCols, .ActiveRow, False
           .MoveRange 1, .MaxRows, .MaxCols, .MaxRows, 1, .ActiveRow
           .MaxRows = .MaxRows - 1
           .Row = .ActiveRow + 1
           .Col = 2
           .SetActiveCell .Col, .Row
        End If
    End With
End Select
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim MaxFila As Long, sql1 As String, sql2 As String, sql3 As String
vg_modrec = IIf(GetParametro("Modrec ") = "1", trae, flase)
Select Case Button.Index
Case 1
    With vaSpread1(inddet)
        If vg_modrec = False Then
           .Row = .ActiveRow
           .Col = 1
           If Trim(.text) = "" Or .Lock = True Then MsgBox "No esta autorizado, para ingresar ó modificar ingredientes en receta, comuniquese con el administrador", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        End If
        vg_nombre = "": vg_codigo = ""
        vg_left = fpayuda(2).Left + 550
        Toolbar5.Enabled = False
        If vg_modrec = False Then
           Dim fampro As String, codtip As Long, parval As String
           fampro = "": codtip = 0: parval = ""
           .Row = .ActiveRow
           .Col = 1
           sql1 = IIf(vg_tipbase = "1", " mid(par_codigo,1,6) ", " substring(par_codigo,1,6) ")
           RS1.Open "SELECT par_codigo, par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND " & sql1 & " = 'CamIng'", vg_db, adOpenStatic
           Do While Not RS1.EOF
              sql1 = IIf(vg_tipbase = "1", " trim(str(a.pro_codtip)) ", " ltrim(convert(varchar(20),a.pro_codtip)) ")
              RS2.Open "SELECT a.pro_codtip FROM b_productos a, b_productosing b " & _
                       "WHERE  b.pri_codpro = a.pro_codigo " & _
                       "AND  " & sql1 & " IN ('" & fg_CambiaChar(GetParametro(RS1!par_codigo), ";", "','") & "') " & _
                       "AND   b.pri_coding = '" & Trim(.text) & "'", vg_db, adOpenStatic
              If Not RS2.EOF Then
                 parval = RS1!par_valor
                 codtip = RS2!pro_codtip
                 fampro = fg_buscarcodtip(RS1!par_valor & ";", Trim(Str(RS2!pro_codtip)))
                 fampro = fg_CambiaChar(fampro, ";", "','")
                 RS2.Close: Set RS2 = Nothing
                 Exit Do
              End If
              RS2.Close: Set RS2 = Nothing
              RS1.MoveNext
           Loop
           RS1.Close: Set RS1 = Nothing
    '       RS2.Open "SELECT a.pro_codtip FROM b_productos a, b_productosing b " & _
    '                "WHERE  b.pri_codpro=a.pro_codigo " & _
    '                "AND  (trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvfre"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvpre"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvcon"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne1"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne2"), ";", "','") & "') " & _
    '                "OR    trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne3"), ";", "','") & "')) " & _
    '                "AND   b.pri_coding='" & Trim(vaSpread1(inddet).text) & "'", vg_db, adOpenStatic
    '       If RS2.EOF Then RS2.Close: Set RS2 = Nothing: Exit Sub
    '       If RS2!pro_codtip = (fg_CambiaChar(GetParametro("fvfre"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("fvpre"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("fvcon"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("fvpre"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("fvfre"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("fvcon"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("fvcon"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("fvpre"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("fvfre"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("carne1"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("carne1"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("carne2"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("carne2"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("carne2"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("carne1"), ";", "','"))
    '       ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("carne3"), ";", "','")) Then
    '          fampro = (fg_CambiaChar(GetParametro("carne1"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("carne2"), ";", "','")) & "', '" & (fg_CambiaChar(GetParametro("carne3"), ";", "','"))
    '       End If
    '       RS2.Close: Set RS2 = Nothing
           B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", fampro
        ElseIf vg_modrec = True Then
           B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "Gen"
        End If
        B_TabEst.Show 1
        If vg_codigo = "" Then Toolbar5.Enabled = True: Exit Sub
        If vg_codigo <> "" And modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
        .Row = .ActiveRow
        .Col = 3
        If Val(.text) > 0 Then
           canpro1 = 0: codpro1 = "": pctnut1 = 0
           .Col = 1: codpro1 = .text
           .Col = 3: canpro1 = .text
           .Col = 8: pctnut1 = .text
           .Col = 10: cospro = .text
           '------- Resta Aporte
           sql2 = IIf(vg_tipbase = "1", " (((((" & pctnut1 & "/100)*(c.pnu_canapo*(" & canpro1 & "/" & Val(fpDouble1(0).Value) & ")))/a.ing_facnut))) AS  canneta ", " (((((convert(float," & pctnut1 & ")/100)*(c.pnu_canapo*(convert(float," & canpro1 & ")/convert(int," & fpDouble1(0).Value & "))))/a.ing_facnut))) AS  canneta ")
           sql3 = IIf(vg_tipbase = "1", " (((" & pctnut1 & "/100)*" & canpro1 & ")) AS cangrverneto ", " (((convert(float," & pctnut1 & ")/100)*convert(float," & canpro1 & "))) AS cangrverneto ")
           RS.Open "SELECT DISTINCT b.nut_nombre, b.nut_codigo, a.ing_indgrv, a.ing_indpav, b.nut_secnro, " & _
                   "" & sql2 & ", " & _
                   "" & sql3 & " " & _
                   "FROM  b_ingrediente a, a_nutriente b, b_productonut c " & _
                   "WHERE a.ing_codigo = c.pnu_codpro " & _
                   "AND   c.pnu_codapo = b.nut_codigo " & _
                   "AND   a.ing_codigo = '" & codpro1 & "' " & _
                   "ORDER BY b.nut_secnro", vg_db, adOpenStatic
           If Not RS.EOF Then
              i = 1
              If RS!ing_indgrv = 1 Then Label2(4).Caption = Format(CCur(Label2(4).Caption - RS!cangrverneto), fg_Pict(6, 2))
              Do While Not RS.EOF
                 If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then Label2(5).Caption = Format(CCur(Label2(5).Caption - RS!canneta), fg_Pict(6, 2))
                 vaSpread2(indapo).Row = i
                 vaSpread2(indapo).Col = 3
                 vaSpread2(indapo).text = Format(CCur(vaSpread2(indapo).text - RS!canneta), fg_Pict(6, 2))
                 RS.MoveNext: i = i + 1
              Loop
              Label2(3).Caption = Format(CCur(Label2(3).Caption - cospro), fg_Pict(6, 2))
              CalTotalPavb
           End If
           RS.Close: Set RS = Nothing
        End If
        RS.Open "SELECT a.*, b.* FROM  b_ingrediente a, a_unidadmed b " & _
                "WHERE a.ing_unimed = b.unm_codigo " & _
                "AND   a.ing_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic
        If RS.EOF Then RS.Close: Set RS = Nothing: .text = "": Toolbar5.Enabled = True: Exit Sub
        formatearcelda .Row, RS!ing_codigo, RS!ing_nombre, RS!unm_nomcor, 0, RS!ing_pctapr, RS!ing_pctcoc, RS!ing_pctnut, 0, 0, 0, inddet, 0
        RS.Close: Set RS = Nothing
        Me.Refresh
        .Row = .ActiveRow
        .SetActiveCell 3, .Row
        .SetFocus
    End With
    calnetoservido
    Toolbar5.Enabled = True
Case 2
    If vg_modrec = False Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = False Or Toolbar1.Buttons(12).Visible = False Then Hab_Des 0
    With vaSpread1(inddet)
        For i = 1 To .MaxRows
            .Row = i: .Col = 1
            If .text <> "" Then MaxFila = .Row
        Next i
        .Row = .ActiveRow
        .Col = .ActiveCol
        If MaxFila < .MaxRows Then
           MaxFila = MaxFila + 1
        Else
           Exit Sub
        End If
        '------- Insertar columna
        If .Row + 1 < 41 Then
           .MoveRange 1, (.ActiveRow), .MaxCols, (.MaxRows - 1), 1, (.Row + 1)
           .ClearRange 1, .ActiveRow, .MaxCols, .ActiveRow, False
        End If
    End With
Case 3
    If vg_modrec = False Then Exit Sub
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
    '------- Resta aporte
    With vaSpread1(inddet)
        .Row = .ActiveRow
        .Col = 1
        If .text = "" Then .DeleteRows .Row, 1: .MaxRows = .MaxRows - 1: .MaxRows = .MaxRows + 1: wsmaxfilas = wsmaxfilas - 1: .Row = .ActiveRow: Exit Sub
        canpro1 = 0: codpro1 = "": pctnut1 = 0
        .Col = 1: codpro1 = .text
        .Col = 3: canpro1 = .text
        .Col = 8: pctnut1 = .text
        If canpro1 > 0 Then
           sql2 = IIf(vg_tipbase = "1", " (((((" & pctnut1 & "/100)*(c.pnu_canapo*(" & canpro1 & "/" & Val(fpDouble1(0).Value) & ")))/a.ing_facnut))) AS  canneto ", " (((((convert(float," & pctnut1 & ")/100)*(c.pnu_canapo*(convert(float," & canpro1 & ")/convert(int," & fpDouble1(0).Value & "))))/a.ing_facnut))) AS  canneto ")
           sql3 = IIf(vg_tipbase = "1", " (((" & pctnut1 & "/100)*" & canpro1 & ")) AS cangrverneto ", " (((convert(float," & pctnut1 & ")/100)*convert(float," & canpro1 & "))) AS cangrverneto ")
           RS.Open "SELECT DISTINCT b.nut_nombre, b.nut_codigo, a.ing_indgrv, a.ing_indpav, b.nut_secnro, " & _
                 "" & sql2 & ", " & _
                 "" & sql3 & " " & _
                 "FROM  b_ingrediente a, a_nutriente b, b_productonut c " & _
                 "WHERE a.ing_codigo = c.pnu_codpro " & _
                 "AND   c.pnu_codapo = b.nut_codigo " & _
                 "AND   a.ing_codigo = '" & codpro1 & "' " & _
                 "ORDER BY b.nut_secnro", vg_db, adOpenStatic
           If Not RS.EOF Then
              i = 1
              If RS!ing_indgrv = 1 Then Label2(4).Caption = Format(CCur(Label2(4).Caption - RS!cangrverneto), fg_Pict(6, 2))
              Do While Not RS.EOF
                 If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then Label2(5).Caption = Format(CCur(Label2(5).Caption - RS!canneto), fg_Pict(6, 2))
                 vaSpread2(indapo).Row = i
                 vaSpread2(indapo).Col = 3
                 vaSpread2(indapo).text = Format(CCur(vaSpread2(indapo).text - RS!canneto), fg_Pict(6, 2))
                 RS.MoveNext: i = i + 1
              Loop
           End If
           RS.Close: Set RS = Nothing
           '------- Calcular aportes pavb
           CalTotalPavb
        End If
        .DeleteRows .Row, 1
        .MaxRows = .MaxRows - 1
        .MaxRows = .MaxRows + 1
        cosrec = 0
        For i = 1 To .MaxRows
            .Row = i: .Col = 10
            If .text <> "" Then cosrec = CCur(cosrec + .text)
        Next i
        Label2(3).Caption = Format((cosrec), fg_Pict(6, 2))
        wsmaxfilas = wsmaxfilas - 1
        .Row = .ActiveRow
    End With
    calnetoservido
Case 4
    If vg_modrec = False Then Exit Sub
    With vaSpread1(inddet)
        .Row = .ActiveRow
        .Col = .ActiveCol
        If .Row > 1 Then
           If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
           '------- Copiar datos ultima fila
           .MaxRows = .MaxRows + 1
           .MoveRange 1, (.ActiveRow - 1), .MaxCols, (.ActiveRow - 1), 1, .MaxRows
           '------- Copiar datos fila seleccionada
           .ClearRange 1, (.ActiveRow + 1), .MaxCols, (.ActiveRow - 1), False
           .MoveRange 1, (.ActiveRow), .MaxCols, (.ActiveRow), 1, (.ActiveRow - 1)
           '------- Devolver datos fila y restar ultima fila
           .ClearRange 1, .ActiveRow, .MaxCols, .ActiveRow, False
           .MoveRange 1, .MaxRows, .MaxCols, .MaxRows, 1, .ActiveRow
           .MaxRows = .MaxRows - 1
           .Row = .ActiveRow - 1
           .Col = 2
           .SetActiveCell .Col, .Row
        End If
    End With
Case 5
    If vg_modrec = False Then Exit Sub
    With vaSpread1(inddet)
        .Row = .ActiveRow
        .Col = .ActiveCol
        If .Row + 1 < 41 Then
           '------- Copiar datos ultima fila
           If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
           .MaxRows = .MaxRows + 1
           .MoveRange 1, (.ActiveRow + 1), .MaxCols, (.ActiveRow + 1), 1, .MaxRows
           '------- Copiar datos fila seleccionada
           .ClearRange 1, (.ActiveRow + 1), .MaxCols, (.ActiveRow + 1), False
           .MoveRange 1, (.ActiveRow), .MaxCols, (.ActiveRow), 1, (.ActiveRow + 1)
           '------- Devolver datos fila y restar ultima fila
           .ClearRange 1, .ActiveRow, .MaxCols, .ActiveRow, False
           .MoveRange 1, .MaxRows, .MaxCols, .MaxRows, 1, .ActiveRow
           .MaxRows = .MaxRows - 1
           .Row = .ActiveRow + 1
           .Col = 2
           .SetActiveCell .Col, .Row
        End If
    End With
End Select
End Sub

Private Sub vaSpread1_EditMode(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If est Then Exit Sub
Dim sql1 As String, sql2 As String, sql3 As String, ind_ini As Long
With vaSpread1(inddet)
    Select Case Index
    Case 1, 2, 3
        If modo = "M" And ChangeMade = True Then
           Gl_Ac_Botones Me, 3, 0, modo
           If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
              Hab_Des 0
           Else
              Exit Sub
           End If
        End If
        Select Case Col
        Case 1
            .Row = Row: .Col = 1
    '        If vg_modrec = False Then
    '           .Row = .ActiveRow
    '           .col = 1
    '           If Trim(.Text) = "" Then MsgBox "No esta autorizado, para ingresar nuevos ingredientes en receta, comuniquese con el administrador", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    '        End If
            If .text = "" And vg_modrec = True Then Exit Sub
            If ChangeMade = False Then codpro2 = .text: Exit Sub
            codpro1 = .text
            If vg_modrec = False Then
               Dim fampro As String, codtip As Long, parval As String
               fampro = "": codtip = 0: parval = ""
               sql1 = IIf(vg_tipbase = "1", " mid(par_codigo,1,6) ", " substring(par_codigo,1,6) ")
               RS1.Open "SELECT par_codigo, par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND " & sql1 & " = 'CamIng'", vg_db, adOpenStatic
               Do While Not RS1.EOF
                  sql1 = IIf(vg_tipbase = "1", " trim(str(a.pro_codtip)) ", " ltrim(convert(varchar(20),a.pro_codtip)) ")
                  RS2.Open "SELECT a.pro_codtip FROM b_productos a, b_productosing b " & _
                           "WHERE  b.pri_codpro = a.pro_codigo " & _
                           "AND  " & sql1 & " IN ('" & fg_CambiaChar(GetParametro(RS1!par_codigo), ";", "','") & "') " & _
                           "AND   b.pri_coding = '" & codpro2 & "'", vg_db, adOpenStatic
                  If Not RS2.EOF Then
                     parval = RS1!par_valor
                     codtip = RS2!pro_codtip
                     fampro = fg_buscarcodtip(RS1!par_valor & ";", Trim(Str(RS2!pro_codtip)))
                     fampro = fg_CambiaChar(fampro, ";", "','")
                     RS2.Close: Set RS2 = Nothing
                     Exit Do
                  End If
                  RS2.Close: Set RS2 = Nothing
                  RS1.MoveNext
               Loop
               RS1.Close: Set RS1 = Nothing
    
    '           RS.Open "SELECT a.pro_codtip FROM b_productos a, b_productosing b " & _
    '                    "WHERE  b.pri_codpro=a.pro_codigo " & _
    '                    "AND (trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvfre"), ";", "','") & "') " & _
    '                    "OR   trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvpre"), ";", "','") & "') " & _
    '                    "OR   trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvcon"), ";", "','") & "') " & _
    '                    "OR   trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne1"), ";", "','") & "') " & _
    '                    "OR   trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne2"), ";", "','") & "') " & _
    '                    "OR   trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne3"), ";", "','") & "')) " & _
    '                    "AND  b.pri_coding='" & codpro2 & "'", vg_db, adOpenStatic
    '           If RS.EOF Then RS.Close: Set RS = Nothing: .Col = 1: .text = codpro2: Exit Sub
    '           If RS!pro_codtip = (fg_CambiaChar(GetParametro("fvfre"), ";", "','")) Then
    '              fampro = (fg_CambiaChar(GetParametro("fvpre"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("fvcon"), ";", "','"))
    '           ElseIf RS!pro_codtip = (fg_CambiaChar(GetParametro("fvpre"), ";", "','")) Then
    '              fampro = (fg_CambiaChar(GetParametro("fvfre"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("fvcon"), ";", "','"))
    '           ElseIf RS!pro_codtip = (fg_CambiaChar(GetParametro("fvcon"), ";", "','")) Then
    '              fampro = (fg_CambiaChar(GetParametro("fvpre"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("fvfre"), ";", "','"))
    '           ElseIf RS!pro_codtip = (fg_CambiaChar(GetParametro("carne1"), ";", "','")) Then
    '              fampro = (fg_CambiaChar(GetParametro("carne2"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("carne1"), ";", "','"))
    '           ElseIf RS!pro_codtip = (fg_CambiaChar(GetParametro("carne2"), ";", "','")) Then
    '              fampro = (fg_CambiaChar(GetParametro("carne1"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("carne2"), ";", "','"))
    '           ElseIf RS2!pro_codtip = (fg_CambiaChar(GetParametro("carne3"), ";", "','")) Then
    '              fampro = (fg_CambiaChar(GetParametro("carne1"), ";", "','")) & "','" & (fg_CambiaChar(GetParametro("carne2"), ";", "','")) & "', '" & (fg_CambiaChar(GetParametro("carne3"), ";", "','"))
    '           End If
    '           RS.Close: Set RS2 = Nothing
               sql1 = IIf(vg_tipbase = "1", " trim(str(e.tip_codigo)) ", " ltrim(convert(varchar(20),e.tip_codigo)) ")
               RS.Open "SELECT a.*, b.* FROM  b_ingrediente a, a_unidadmed b, b_productos c, b_productosing d, a_tipopro e " & _
                       "WHERE  a.ing_codigo = d.pri_coding " & _
                       "AND    a.ing_unimed = b.unm_codigo " & _
                       "AND    d.pri_codpro = c.pro_codigo " & _
                       "AND    c.pro_codtip = e.tip_codigo " & _
                       "AND    " & sql1 & " IN ('" & (fampro) & "') " & _
                       "AND    a.ing_codigo = '" & codpro1 & "'", vg_db, adOpenStatic
            ElseIf vg_modrec = True Then
               RS.Open "SELECT a.*, b.* FROM  b_ingrediente a, a_unidadmed b " & _
                       "WHERE a.ing_unimed = b.unm_codigo " & _
                       "AND   a.ing_codigo = '" & codpro1 & "'", vg_db, adOpenStatic
            End If
            If RS.EOF Then RS.Close: Set RS = Nothing: .Col = 1: .text = codpro2: Exit Sub 'codpro2 = "":
            formatearcelda .Row, RS!ing_codigo, RS!ing_nombre, RS!unm_nomcor, 0, RS!ing_pctapr, RS!ing_pctcoc, RS!ing_pctnut, 0, 0, 0, inddet, 0
            RS.Close: Set RS = Nothing
            .Row = .ActiveRow
            .SetActiveCell 2, .Row
            .SetFocus
        Case 3
            .Row = Row
            .Col = Col
            If ChangeMade = False Then canpro2 = .text: Exit Sub
            canpro1 = .text
            '------- traer precio producto
            .Col = 1: codpro1 = .text
            RS.Open "SELECT c.cpi_precos, b.pro_codtip FROM b_ingrediente a, b_productos b, b_contlistpreing c " & _
                    "WHERE c.cpi_codped=b.pro_codigo AND a.ing_codigo=c.cpi_coding AND a.ing_codigo='" & codpro1 & "' AND c.cpi_cencos='" & MuestraCasino(1) & "'", vg_db, adOpenStatic
            If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
            '------- Validar gramo familia producto 5 etapas
            If inddet = 3 And vg_newestrec = True Or ("S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) And vg_5etapas And vg_tiprec > 0) Then
               RS1.Open "SELECT DISTINCT a.gfp_cencos FROM b_gramofamproducto a, b_receta b " & _
                        "WHERE a.gfp_catdie = b.rec_catdie " & _
                        "AND   a.gfp_tiprec = b.rec_tippla " & _
                        "AND   b.rec_codigo = " & codigo & " " & _
                        "AND   a.gfp_cencos = '" & vg_codcasino & "' " & _
                        "AND   a.gfp_codreg = " & vg_tiprec & " " & _
                        "AND   a.gfp_fampro = " & RS!pro_codtip & " " & _
                        "AND  (a.gfp_graini IS NOT NULL OR a.gfp_grafin IS NOT NULL) " & _
                        "AND   a.gfp_graini > 0 AND a.gfp_grafin > 0 AND (" & canpro1 & " >= a.gfp_graini AND " & canpro1 & " <= a.gfp_grafin)", vg_db, adOpenStatic
               If RS1.EOF Then RS.Close: Set RS = Nothing: RS1.Close: Set RS1 = Nothing: .Col = Col: .text = canpro2: MsgBox "Gramaje esta fuera de rango, comuniquese con el administrador", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
               RS1.Close: Set RS1 = Nothing
            End If
            .Col = 10: .text = Format(CCur(RS!cpi_precos * canpro1), fg_Pict(6, 2))
            RS.Close: Set RS = Nothing
            If canpro2 <> canpro1 Then
               .Col = 1: codpro1 = .text
               .Col = 8: pctnut1 = .text
               '------- Calcular gramaje neto
               .Col = 9
               .CellType = CellTypeStaticText
               .TypeHAlign = TypeHAlignRight
               .text = Format(CCur((pctnut1 / 100) * canpro1), fg_Pict(6, 2))
               '------- Calcular % limpieza & cocción
               .Col = 5: pctapr1 = .text
               .Col = 6: pctcoc1 = .text
               .Col = 7
               .CellType = CellTypeStaticText
               .TypeHAlign = TypeHAlignRight
               .text = Format(CCur(((pctapr1 / 100) * canpro1) * pctcoc1 / 100), fg_Pict(6, 2))
               cosrec = 0
               For i = 1 To .MaxRows
                   .Row = i: .Col = 10
                   If .text <> "" Then cosrec = CCur(cosrec + .text)
               Next i
               Label2(3).Caption = Format(CCur(cosrec / Val(fpDouble1(0).Value)), fg_Pict(6, 2))
               '------- Calcular total pavb
               CalTotalPavb
               calnetoservido
               '------- Resta aporte
               If Val(fpDouble1(0).text) < 1 Then Exit Sub
               If canpro2 > 0 Then
                  sql2 = IIf(vg_tipbase = "1", " (((((" & pctnut1 & "/100)*(c.pnu_canapo*(" & canpro2 & "/" & Val(fpDouble1(0).Value) & ")))/a.ing_facnut))) AS  candiet ", " (((((convert(float," & pctnut1 & ")/100)*(c.pnu_canapo*(convert(float," & canpro2 & ")/convert(int," & fpDouble1(0).Value & "))))/a.ing_facnut))) AS  candiet ")
                  sql3 = IIf(vg_tipbase = "1", " (((" & pctnut1 & "/100)*" & canpro2 & ")) AS cangrverneto ", " (((convert(float," & pctnut1 & ")/100)*convert(float," & canpro2 & "))) AS cangrverneto ")
                  RS.Open "SELECT DISTINCT b.nut_nombre, b.nut_codigo, a.ing_indgrv, a.ing_indpav, b.nut_secnro, " & _
                          "" & sql2 & ", " & _
                          "" & sql3 & " " & _
                          "FROM  b_ingrediente a, a_nutriente b, b_productonut c " & _
                          "WHERE a.ing_codigo = c.pnu_codpro " & _
                          "AND   c.pnu_codapo = b.nut_codigo " & _
                          "AND   a.ing_codigo = '" & codpro1 & "' " & _
                          "ORDER BY b.nut_secnro", vg_db, adOpenStatic
                  If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
    '             i = 1
                  vaSpread2(indapo).Visible = False
                  If RS!ing_indgrv = 1 Then Label2(4).Caption = Format(CCur(Label2(4).Caption - RS!cangrverneto), fg_Pict(6, 2))
                  Do While Not RS.EOF
                     If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then Label2(5).Caption = Format(CCur(Label2(5).Caption - RS!candiet), fg_Pict(6, 2))
                     ind_ini = vaSpread2(indapo).SearchCol(1, -1, vaSpread2(indpao).MaxRows, Trim(CStr(RS!nut_codigo)), SearchFlagsEqual)
                     If ind_ini > 0 Then
                        vaSpread2(indapo).Row = ind_ini 'i
                        vaSpread2(indapo).Col = 3
                        vaSpread2(indapo).text = Format(CCur(vaSpread2(indapo).text - RS!candiet), fg_Pict(6, 2))
                     End If
    '                i = i + 1
                     RS.MoveNext
                  Loop
                  RS.Close: Set RS = Nothing
                  vaSpread2(indapo).Visible = True
               End If
               '------- Sumar aporte
               If canpro1 < 0 Then Exit Sub
               sql2 = IIf(vg_tipbase = "1", " (((((" & pctnut1 & "/100)*(c.pnu_canapo*(" & canpro1 & "/" & Val(fpDouble1(0).Value) & ")))/a.ing_facnut))) AS  candiet ", " (((((convert(float," & pctnut1 & ")/100)*(c.pnu_canapo*(convert(float," & canpro1 & ")/convert(int," & fpDouble1(0).Value & "))))/a.ing_facnut))) AS  candiet ")
               sql3 = IIf(vg_tipbase = "1", " (((" & pctnut1 & "/100)*" & canpro1 & ")) AS cangrverneto  ", " (((convert(float," & pctnut1 & ")/100)*convert(float," & canpro1 & "))) AS cangrverneto  ")
               RS.Open "SELECT DISTINCT b.nut_nombre, b.nut_codigo, a.ing_indgrv, a.ing_indpav, b.nut_secnro, " & _
                       "" & sql2 & ", " & _
                       "" & sql3 & " " & _
                       "FROM  b_ingrediente a, a_nutriente b, b_productonut c " & _
                       "WHERE a.ing_codigo = c.pnu_codpro " & _
                       "AND   c.pnu_codapo = b.nut_codigo " & _
                       "AND   a.ing_codigo = '" & codpro1 & "' " & _
                       "ORDER BY b.nut_secnro", vg_db, adOpenStatic
               If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
    '           i = 1
               vaSpread2(indapo).Visible = False
               If RS!ing_indgrv = 1 Then Label2(4).Caption = Format(CCur(Label2(4).Caption + RS!cangrverneto), fg_Pict(6, 2))
               Do While Not RS.EOF
                  If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then Label2(5).Caption = Format(CCur(Label2(5).Caption + RS!candiet), fg_Pict(6, 2))
                  ind_ini = vaSpread2(indapo).SearchCol(1, -1, vaSpread2(indpao).MaxRows, Trim(CStr(RS!nut_codigo)), SearchFlagsEqual)
                  If ind_ini > 0 Then
                     vaSpread2(indapo).Row = ind_ini 'i
                     vaSpread2(indapo).Col = 3
                     vaSpread2(indapo).text = Format(CCur(RS!candiet + vaSpread2(indapo).text), fg_Pict(6, 2))
                  End If
    '              i = i + 1
                  RS.MoveNext
               Loop
               RS.Close: Set RS = Nothing
               vaSpread2(indapo).Visible = True
            End If
        Case 5, 6
            .Row = Row
            .Col = 1
            If .text = "" Then Exit Sub
            If ChangeMade = False Then .Col = 5: pctapr2 = .text: .Col = 6: pctcoc2 = .text: Exit Sub
            .Col = 3: canpro1 = .text
            .Col = 5: pctapr1 = .text
            .Col = 6: pctcoc1 = .text
            If pctapr1 = 0 Then .Col = 5: .text = pctapr2: Exit Sub
            If pctcoc1 = 0 Then .Col = 6: .text = pctcoc2: Exit Sub
            '------- Calcular % limpieza & cocción
            .Col = 7
            .CellType = 5
            .TypeHAlign = 1
            .text = Format(CCur(((pctapr1 / 100) * canpro1) * (pctcoc1 / 100)), fg_Pict(6, 2))
            calnetoservido
        Case 8
            .Row = Row
            .Col = 1
            If .text = "" Then Exit Sub
            If ChangeMade = False Then .Col = 8: pctnut2 = .text: Exit Sub
            .Col = 1: codpro1 = .text
            .Col = 3: canpro1 = .text
            .Col = 8: pctnut1 = .text
            If pctnut1 = 0 Then .Col = 8: .text = pctnut2: Exit Sub
            '------- Resta aporte
            If canpro1 < 0 Then Exit Sub
            sql2 = IIf(vg_tipbase = "1", " (((((" & pctnut2 & "/100)*(c.pnu_canapo*(" & canpro1 & "/" & Val(fpDouble1(0).Value) & ")))/a.ing_facnut))) AS  candiet ", " (((((convert(float," & pctnut2 & ")/100)*(c.pnu_canapo*(convert(float," & canpro1 & ")/convert(int," & fpDouble1(0).Value & "))))/a.ing_facnut))) AS  candiet ")
            sql3 = IIf(vg_tipbase = "1", " (((" & pctnut2 & "/100)*" & canpro1 & ")) AS cangrverneto ", " (((convert(float," & pctnut2 & ")/100)*convert(float," & canpro1 & "))) AS cangrverneto ")
            RS.Open "SELECT DISTINCT b.nut_nombre, b.nut_codigo, a.ing_indgrv, a.ing_indpav, b.nut_secnro, " & _
                    "" & sql2 & ", " & _
                    "" & sql3 & " " & _
                    "FROM  b_ingrediente a, a_nutriente b, b_productonut c " & _
                    "WHERE a.ing_codigo = c.pnu_codpro " & _
                    "AND   c.pnu_codapo = b.nut_codigo " & _
                    "AND   a.ing_codigo = '" & codpro1 & "' " & _
                    "ORDER BY b.nut_secnro", vg_db, adOpenStatic
            If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
            vaSpread2(indapo).Visible = False
    '        i = 1
            If RS!ing_indgrv = 1 Then Label2(4).Caption = Format(CCur(Label2(4).Caption - RS!cangrverneto), fg_Pict(6, 2))
            Do While Not RS.EOF
               If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then Label2(5).Caption = Format(CCur(Label2(5).Caption - RS!candiet), fg_Pict(6, 2))
               ind_ini = vaSpread2(indapo).SearchCol(1, -1, vaSpread2(indpao).MaxRows, Trim(CStr(RS!nut_codigo)), SearchFlagsEqual)
               If ind_ini > 0 Then
                  vaSpread2(indapo).Row = ind_ini 'i
                  vaSpread2(indapo).Col = 3
                  vaSpread2(indapo).text = Format(CCur(vaSpread2(indapo).text - RS!candiet), fg_Pict(6, 2))
               End If
    '           i = i + 1
               RS.MoveNext
            Loop
            RS.Close: Set RS = Nothing
            vaSpread2(indapo).Visible = True
            '------- Sumar aporte
            If canpro1 < 0 Then Exit Sub
            sql2 = IIf(vg_tipbase = "1", " (((((" & pctnut1 & "/100)*(c.pnu_canapo*(" & canpro1 & "/" & Val(fpDouble1(0).Value) & ")))/a.ing_facnut))) AS  candiet ", " (((((convert(float," & pctnut1 & ")/100)*(c.pnu_canapo*(convert(float," & canpro1 & ")/convert(int," & fpDouble1(0).Value & "))))/a.ing_facnut))) AS  candiet ")
            sql3 = IIf(vg_tipbase = "1", " (((" & pctnut1 & "/100)*" & canpro1 & ")) AS cangrverneto ", " (((convert(float," & pctnut1 & ")/100)*convert(float," & canpro1 & "))) AS cangrverneto ")
            RS.Open "SELECT DISTINCT b.nut_nombre, b.nut_codigo, a.ing_indgrv, a.ing_indpav, b.nut_secnro, " & _
                    "" & sql2 & ", " & _
                    "" & sql3 & " " & _
                    "FROM  b_ingrediente a, a_nutriente b, b_productonut c " & _
                    "WHERE a.ing_codigo = c.pnu_codpro " & _
                    "AND   c.pnu_codapo = b.nut_codigo " & _
                    "AND   a.ing_codigo = '" & codpro1 & "' " & _
                    "ORDER BY b.nut_secnro", vg_db, adOpenStatic
            If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
    '        i = 1
            vaSpread2(indapo).Visible = False
            If RS!ing_indgrv = 1 Then Label2(4).Caption = Format(CCur(Label2(4).Caption + RS!cangrverneto), fg_Pict(6, 2))
            Do While Not RS.EOF
               If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then Label2(5).Caption = Format(CCur(Label2(5).Caption + RS!candiet), fg_Pict(6, 2))
               ind_ini = vaSpread2(indapo).SearchCol(1, -1, vaSpread2(indpao).MaxRows, Trim(CStr(RS!nut_codigo)), SearchFlagsEqual)
               If ind_ini > 0 Then
                  vaSpread2(indapo).Row = ind_ini 'i
                  vaSpread2(indapo).Col = 3
                  vaSpread2(indapo).text = Format(CCur(RS!candiet + vaSpread2(indapo).text), fg_Pict(6, 2))
               End If
    '           i = i + 1
               RS.MoveNext
            Loop
            RS.Close: Set RS = Nothing
            vaSpread2(indapo).Visible = True
            '------- Calcular gramaje neto
            .Col = 9
            .CellType = 5
            .TypeHAlign = 1
            .text = Format(CCur((pctnut1 / 100) * canpro1), fg_Pict(6, 2))
            '------- Calcular total pavb
            CalTotalPavb
            calnetoservido
        End Select
    End Select
End With
End Sub

Private Sub vaSpread1_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
With vaSpread1(0)
    Select Case Index
    Case 0
        If .MaxRows < 1 Or NewRow = -1 Then Exit Sub
        .Row = NewRow
        .Col = 1
        codigo = Val(.text)
        modo = "M"
    End Select
End With
End Sub

Private Sub vaSpread1_TextTipFetch(Index As Integer, ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If Row = 0 Or Index = 0 Or Index = 1 Then Exit Sub
With vaSpread1(Index)
    .Row = Row: .Col = 3: If .Lock = True Then Exit Sub
    Dim param As String, Nombre As String
    TipWidth = 4000
    ShowTip = True
    MultiLine = 2
    .Row = Row: .Col = 2: Nombre = .text
    .Row = Row: .Col = 11: param = .text
    TipText = "Ingrediente : " & Trim(Nombre) & vbCrLf & _
              "Gramo Fam. Producto : " & param
End With
End Sub

Sub CargaMetodoReceta()
itexto = 1
RichTextBox1(0).TextRTF = "": metodoreceta = ""
If vg_newcodrec > 0 Then
   codigo = vg_newcodrec
   RS.Open "SELECT rec_nombre FROM b_receta WHERE rec_codigo=" & codigo & "", vg_db, adOpenStatic
   If Not RS.EOF Then Label3(0).Caption = Trim(RS!rec_nombre)
   RS.Close: Set RS = Nothing
Else
   vaSpread1(0).Row = vaSpread1(0).ActiveRow
   vaSpread1(0).Col = 1: codigo = vaSpread1(0).text
   vaSpread1(0).Col = 2: Label3(0).Caption = Trim(vaSpread1(0).text)
End If
If vg_modrec = False Then
   Frame4.Enabled = False: Frame5(0).Enabled = False
End If
modo = "M"
RS.Open "SELECT rec_metpre FROM b_receta WHERE rec_codigo=" & codigo & " AND (rec_metpre) IS NOT NULL", vg_db, adOpenStatic
If Not RS.EOF Then
   RichTextBox1(0).TextRTF = RS!rec_metpre
   metodoreceta = RichTextBox1(0).TextRTF 'fg_bcoenter(RichTextBox1.textRTF) 'LimpiaDato(ConSql!Rcpe_Mthd_Desc)
End If
RS.Close: Set RS = Nothing
itexto = 0
End Sub

Sub CargaGrupoVulnerable()
itexto = 1
RichTextBox1(1).TextRTF = "": grupovulnerable = ""
If vg_newcodrec > 0 Then
   codigo = vg_newcodrec
   RS.Open "SELECT rec_nombre FROM b_receta WHERE rec_codigo = " & codigo & "", vg_db, adOpenStatic
   If Not RS.EOF Then Label3(1).Caption = Trim(RS!rec_nombre)
   RS.Close: Set RS = Nothing
Else
   vaSpread1(0).Row = vaSpread1(0).ActiveRow
   vaSpread1(0).Col = 1: codigo = vaSpread1(0).text
   vaSpread1(0).Col = 2: Label3(1).Caption = Trim(vaSpread1(0).text)
End If
If vg_modrec = False Then
   'Frame4.Enabled = False
   Frame5(1).Enabled = False
End If
'------- Validar si esta la opción de visualizar grupo vulnerable
If Not vg_modrec Then
   RS.Open "SELECT DISTINCT par_codigo FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'opgruvul' AND par_valor = 'S'", vg_db, adOpenStatic
   If RS.EOF Then RS.Close: Set RS = Nothing: itexto = 0: Exit Sub
   RS.Close: Set RS = Nothing
End If
modo = "M"
RS.Open "SELECT rec_gruvul FROM b_receta WHERE rec_codigo = " & codigo & " AND (rec_gruvul) IS NOT NULL", vg_db, adOpenStatic
If Not RS.EOF Then
   RichTextBox1(1).TextRTF = RS!rec_gruvul
   grupovulnerable = RichTextBox1(1).TextRTF 'fg_bcoenter(RichTextBox1.textRTF) 'LimpiaDato(ConSql!Rcpe_Mthd_Desc)
End If
RS.Close: Set RS = Nothing
itexto = 0
End Sub

Sub calnetoservido()
Dim totcservida As Double, totgneto As Double, totcbruta As Double, totcos As Double
With vaSpread1(inddet)
    If .MaxRows < 1 Then Exit Sub
    totcservida = 0: totgneto = 0: totcos = 0
    For i = 1 To .MaxRows
        .Row = i
        .Col = 3
        If Trim(.text) <> "" Then totcbruta = CCur(totcbruta + .text)
        .Col = 7
        If Trim(.text) <> "" Then totcservida = CCur(totcservida + .text)
        .Col = 9
        If Trim(.text) <> "" Then totgneto = CCur(totgneto + .text)
        .Col = 10
        If Trim(.text) <> "" Then totcos = CCur(totcos + .text)
    Next i
End With
'-------> Mover Totales
SetFpDouble fpDouble1, 1, 1, totcservida
SetFpDouble fpDouble1, 2, 1, totgneto
SetFpDouble fpDouble1, 3, 1, totcbruta

SetFpDouble fpDouble1, 4, 1, totgneto
SetFpDouble fpDouble1, 5, 1, totcservida
SetFpDouble fpDouble1, 7, 1, totcbruta

SetFpDouble fpDouble1, 8, 1, totgneto
SetFpDouble fpDouble1, 9, 1, totcservida
SetFpDouble fpDouble1, 11, 1, totcbruta
Label2(3).Caption = Format(totcos, fg_Pict(6, 2)): Label2(14).Caption = Format(totcos, fg_Pict(6, 2)): Label2(21).Caption = Format(totcos, fg_Pict(6, 2))
End Sub

Sub CalTotalPavb()
With vaSpread2(indapo)
    candiet = 0
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .text = 3 Then
           .Col = 3
           candiet = .text
           If candiet > 0 Then
              If inddet = 1 Then
                 Label2(7).Caption = Format(CCur(((Label2(5).Caption / fpDouble1(0).Value) / candiet) * 100), fg_Pict(6, 2))
              ElseIf inddet = 2 Then
                  Label2(17).Caption = Format(CCur(((Label2(5).Caption / fpDouble1(6).Value) / candiet) * 100), fg_Pict(6, 2))
              ElseIf inddet = 3 Then
                  Label2(18).Caption = Format(CCur(((Label2(5).Caption / fpDouble1(10).Value) / candiet) * 100), fg_Pict(6, 2))
              End If
           Else
              Label2(7).Caption = Format(0, fg_Pict(6, 2))
              Label2(17).Caption = Format(0, fg_Pict(6, 2))
              Label2(18).Caption = Format(0, fg_Pict(6, 2))
           End If
           Exit For
        End If
    Next i
End With
End Sub

Sub formatearcelda(Fila As Long, codpro As String, NomPro As String, NomCor As String, canpro As Double, pctapr As Double, pctcoc As Double, pctnut As Double, canservida As Double, canneta As Double, cospro As Double, Index As Long, codtip As Long)
Dim estmod As Boolean, sql1 As String
With vaSpread1(Index)
    estmod = False
    'If vg_modrec = False And (vg_tiprec = -1 Or vg_tiprec < 10000) And ("S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) and Not vg_5etapas) Then
    If vg_modrec = False And (vg_tiprec = -1 Or vg_tiprec < 10000) And ("S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) Or Not vg_5etapas) Then
    '   RS2.Open "SELECT a.pro_codigo FROM b_productos a, b_productosing b " & _
    '            "WHERE  b.pri_codpro=a.pro_codigo " & _
    '            "AND   (trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvfre"), ";", "','") & "') " & _
    '            "OR     trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvpre"), ";", "','") & "') " & _
    '            "OR     trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("fvcon"), ";", "','") & "') " & _
    '            "OR     trim(str(a.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne1"), ";", "','") & "') " & _
    '            "OR     trim(str(A.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne2"), ";", "','") & "') " & _
    '            "OR     trim(str(A.pro_codtip)) IN ('" & fg_CambiaChar(GetParametro("carne3"), ";", "','") & "')) " & _
    '            "AND    b.pri_coding='" & CodPro & "'", vg_db, adOpenStatic
       sql1 = IIf(vg_tipbase = "1", " mid(par_codigo,1,6) ", " substring(par_codigo,1,6) ")
       RS1.Open "SELECT par_codigo, par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND " & sql1 & " = 'CamIng'", vg_db, adOpenStatic
       Do While Not RS1.EOF
          sql1 = IIf(vg_tipbase = "1", " trim(str(a.pro_codtip)) ", " ltrim(convert(varchar(20),a.pro_codtip)) ")
          RS2.Open "SELECT a.pro_codtip FROM b_productos a, b_productosing b " & _
                   "WHERE  b.pri_codpro=a.pro_codigo " & _
                   "AND  " & sql1 & " IN ('" & fg_CambiaChar(GetParametro(RS1!par_codigo), ";", "','") & "') " & _
                   "AND   b.pri_coding = '" & codpro & "'", vg_db, adOpenStatic
          If Not RS2.EOF Then
             If Not RS2.EOF And Not vg_5etapas Then estmod = True
             RS2.Close: Set RS2 = Nothing
             Exit Do
          End If
          RS2.Close: Set RS2 = Nothing
          RS1.MoveNext
       Loop
       RS1.Close: Set RS1 = Nothing
    '   If Not RS2.EOF And Not vg_5etapas Then estmod = True
    '   RS2.Close: Set RS2 = Nothing
    ElseIf vg_newestrec = True Or ("S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) And vg_5etapas And vg_tiprec > 0) Then
       '------- Parametrización gramos familia producto 5 etapa
       RS2.Open "SELECT DISTINCT a.gfp_cencos, a.gfp_graini, a.gfp_grafin FROM b_gramofamproducto a, b_receta b " & _
                "WHERE a.gfp_catdie = b.rec_catdie " & _
                "AND   a.gfp_tiprec = b.rec_tippla " & _
                "AND   b.rec_codigo = " & codigo & " " & _
                "AND   a.gfp_cencos = '" & vg_codcasino & "' " & _
                "AND   a.gfp_codreg = " & vg_tiprec & " " & _
                "AND   a.gfp_fampro = " & codtip & " " & _
                "AND  (a.gfp_graini IS NOT NULL OR a.gfp_grafin IS NOT NULL) " & _
                "AND   a.gfp_graini > 0 AND a.gfp_grafin >0 " & _
                "AND  (" & canpro & " >= a.gfp_graini AND " & canpro & " <= a.gfp_grafin)", vg_db, adOpenStatic
       If Not RS2.EOF Then .Col = 11: .text = "(" & Format(RS2!gfp_graini, fg_Pict(6, 2)) & " - " & Format(RS2!gfp_grafin, fg_Pict(6, 2)) & ")": estmod = True
       RS2.Close: Set RS2 = Nothing
    End If
    .Row = Fila
    .Col = 1: .text = codpro
    .Lock = IIf(vg_tiprec = 0 Or vg_newestrec = True Or ("S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) And vg_5etapas) Or Not estmod, True, False)
    
    .Col = 2
    .text = NomPro
    
    .Col = 3
    .CellType = CellTypeCurrency
    .TypeCurrencyDecPlaces = vg_RDCa
    .TypeCurrencyMin = "0"
    .TypeCurrencyMax = "99999999"
    .TypeFloatMoney = False
    .TypeFloatSeparator = True
    .TypeHAlign = TypeHAlignRight
    .TypeFloatCurrencyChar = Asc("$")
    .TypeFloatDecimalChar = Asc(".")
    .TypeFloatSepChar = Asc(",")
    .TypeCurrencyShowSymbol = False
    .text = Format(canpro, fg_Pict(6, vg_RDCa))
    .ForeColor = &HFF0000
    If vg_modrec Then
       .Lock = False
    ElseIf Not vg_modrec Then
       .Lock = IIf(vg_tiprec = 0 Or vg_newestrec = True Or ("S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) And vg_5etapas And Not estmod), True, False)
    End If
    
    .Col = 4
    .TypeHAlign = TypeHAlignLeft
    .text = Trim(RS!unm_nomcor)
    
    .Col = 5
    .CellType = CellTypeCurrency
    .TypeFloatDecimalPlaces = 2
    .TypeFloatMin = "0"
    .TypeFloatMax = "9999"
    .TypeFloatMoney = False
    .TypeFloatSeparator = True
    .TypeHAlign = TypeHAlignRight
    .TypeFloatCurrencyChar = Asc("$")
    .TypeFloatDecimalChar = Asc(".")
    .TypeFloatSepChar = Asc(",")
    .TypeCurrencyShowSymbol = False
    .text = Format(pctapr, fg_Pict(3, 2))
    .ForeColor = &HFF0000
    If vg_modrec Then
       .Lock = False
    ElseIf Not vg_modrec Then
       .Lock = IIf(vg_tiprec = 0 Or vg_newestrec = True Or ("S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) And vg_5etapas), True, False)
    End If
    
    .Col = 6
    .CellType = CellTypeCurrency
    .TypeFloatDecimalPlaces = 2
    .TypeFloatMin = "0"
    .TypeFloatMax = "999999"
    .TypeFloatMoney = False
    .TypeFloatSeparator = True
    .TypeHAlign = TypeHAlignRight
    .TypeFloatCurrencyChar = Asc("$")
    .TypeFloatDecimalChar = Asc(".")
    .TypeFloatSepChar = Asc(",")
    .TypeCurrencyShowSymbol = False
    .text = Format(pctcoc, fg_Pict(6, 2))
    .ForeColor = &HFF0000
    If vg_modrec Then
       .Lock = False
    ElseIf Not vg_modrec Then
       .Lock = IIf(vg_tiprec = 0 Or vg_newestrec = True Or ("S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) And vg_5etapas), True, False)
    End If
    
    .Col = 7
    .CellType = CellTypeStaticText
    .TypeHAlign = TypeHAlignRight
    If canservida = 0 Then .text = "" Else .text = Format(canservida, fg_Pict(6, vg_RDCa))
    
    .Col = 8
    .CellType = CellTypeNumber
    .TypeIntegerMin = 2
    .TypeIntegerMax = 100
    .TypeHAlign = TypeHAlignRight
    .TypeSpin = False
    .TypeIntegerSpinInc = 1
    .TypeIntegerSpinWrap = False
    .text = Format(pctnut, fg_Pict(3, 0))
    .ForeColor = &HFF0000
    If vg_modrec Then
       .Lock = False
    ElseIf Not vg_modrec Then
        .Lock = IIf(vg_tiprec = 0 Or vg_newestrec = True Or ("S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) And vg_5etapas), True, False)
    End If
    
    .Col = 9
    .CellType = CellTypeStaticText
    .TypeHAlign = TypeHAlignRight
    If canneta = 0 Then .text = "" Else .text = Format(canneta, fg_Pict(6, vg_RDCa))
    
    .Col = 10
    .CellType = CellTypeStaticText
    .TypeHAlign = TypeHAlignRight
    .text = Format(cospro, fg_Pict(7, 2))
End With
End Sub
