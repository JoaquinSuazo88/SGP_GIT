VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form M_SacToWeb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SacToWeb"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10395
   Icon            =   "M_SacToWeb.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   10335
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   10395
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   105
         Picture         =   "M_SacToWeb.frx":08CA
         Top             =   60
         Width           =   480
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "SacToWeb - Actualización de datos Web Pedidos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   765
         TabIndex        =   9
         Top             =   195
         Width           =   4080
      End
   End
   Begin VB.Frame fraLoadData 
      Caption         =   "Información a subir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5160
      Left            =   345
      TabIndex        =   1
      Top             =   1095
      Width           =   2175
      Begin VB.ListBox lstLoadData 
         Height          =   3210
         ItemData        =   "M_SacToWeb.frx":1194
         Left            =   150
         List            =   "M_SacToWeb.frx":11AA
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   345
         Width           =   1890
      End
      Begin MSComctlLib.ProgressBar prbOther 
         Height          =   285
         Left            =   150
         TabIndex        =   3
         Top             =   4680
         Visible         =   0   'False
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComctlLib.ProgressBar prbMain 
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   3960
         Visible         =   0   'False
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblStatusOther 
         AutoSize        =   -1  'True
         Caption         =   "lblStatusOther"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   180
         TabIndex        =   6
         Top             =   4410
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblStatusMain 
         AutoSize        =   -1  'True
         Caption         =   "lblStatusMain"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   180
         TabIndex        =   5
         Top             =   3690
         Visible         =   0   'False
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   12
      Left            =   0
      TabIndex        =   0
      Top             =   1035
      Width           =   13755
   End
   Begin MSComctlLib.ImageList Icons 
      Index           =   0
      Left            =   6015
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_SacToWeb.frx":1204
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_SacToWeb.frx":1ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_SacToWeb.frx":23B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   8055
      Top             =   285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tlbAction 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   660
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "Icons(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cargar Datos"
            Object.Tag             =   "Load"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   "Cancel"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   "Close"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   5070
      Left            =   2595
      TabIndex        =   10
      Top             =   1170
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   8943
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Casinos"
      TabPicture(0)   =   "M_SacToWeb.frx":2752
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTab(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Productos"
      TabPicture(1)   =   "M_SacToWeb.frx":276E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTab(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Calendario"
      TabPicture(2)   =   "M_SacToWeb.frx":278A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraTab(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Solicitudes"
      TabPicture(3)   =   "M_SacToWeb.frx":27A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraTab(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Proveedores"
      TabPicture(4)   =   "M_SacToWeb.frx":27C2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Check List"
      TabPicture(5)   =   "M_SacToWeb.frx":27DE
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "fraTab(5)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.Frame fraTab 
         Height          =   4650
         Index           =   2
         Left            =   -74880
         TabIndex        =   65
         Top             =   360
         Width           =   7455
         Begin VB.Frame Frame1 
            Caption         =   "Período"
            ForeColor       =   &H00800000&
            Height          =   795
            Index           =   10
            Left            =   135
            TabIndex        =   66
            Top             =   225
            Width           =   5220
            Begin EditLib.fpDoubleSingle Double1 
               Height          =   315
               Index           =   2
               Left            =   540
               TabIndex        =   67
               Top             =   315
               Width           =   600
               _Version        =   196608
               _ExtentX        =   1058
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
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
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
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               Text            =   "0"
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle Double1 
               Height          =   315
               Index           =   3
               Left            =   1710
               TabIndex        =   68
               Top             =   315
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
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
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
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               Text            =   "0"
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
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
               Caption         =   "Año"
               Height          =   195
               Index           =   3
               Left            =   1350
               TabIndex        =   70
               Top             =   405
               Width           =   285
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Mes"
               Height          =   195
               Index           =   4
               Left            =   180
               TabIndex        =   69
               Top             =   405
               Width           =   300
            End
         End
         Begin EditLib.fpBoolean chkSelectAll 
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   71
            Top             =   1140
            Width           =   225
            _Version        =   196608
            _ExtentX        =   397
            _ExtentY        =   344
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ThreeDInsideStyle=   0
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   0
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            AutoToggle      =   -1  'True
            BooleanStyle    =   0
            ToggleFalse     =   ""
            TextFalse       =   ""
            BooleanPicture  =   2
            AlignPictureH   =   3
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   0   'False
            AlignTextH      =   0
            AlignTextV      =   1
            ToggleTrue      =   ""
            TextTrue        =   ""
            Value           =   0
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   ""
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin FPSpread.vaSpread spdCentralCompra 
            Height          =   3195
            Index           =   2
            Left            =   150
            TabIndex        =   72
            Top             =   1100
            Width           =   5205
            _Version        =   393216
            _ExtentX        =   9181
            _ExtentY        =   5636
            _StockProps     =   64
            ColHeaderDisplay=   0
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            MaxRows         =   6
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "M_SacToWeb.frx":27FA
            TextTip         =   2
            ScrollBarTrack  =   3
            CellNoteIndicator=   1
         End
      End
      Begin VB.Frame fraTab 
         Height          =   4650
         Index           =   4
         Left            =   -74940
         TabIndex        =   64
         Top             =   315
         Width           =   7455
      End
      Begin VB.Frame fraTab 
         Height          =   4650
         Index           =   0
         Left            =   -74940
         TabIndex        =   61
         Top             =   315
         Width           =   7455
         Begin EditLib.fpBoolean chkSelectAll 
            Height          =   195
            Index           =   0
            Left            =   510
            TabIndex        =   62
            Top             =   305
            Width           =   225
            _Version        =   196608
            _ExtentX        =   397
            _ExtentY        =   344
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ThreeDInsideStyle=   0
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   0
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            AutoToggle      =   -1  'True
            BooleanStyle    =   0
            ToggleFalse     =   ""
            TextFalse       =   ""
            BooleanPicture  =   2
            AlignPictureH   =   3
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   0   'False
            AlignTextH      =   0
            AlignTextV      =   1
            ToggleTrue      =   ""
            TextTrue        =   ""
            Value           =   0
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   ""
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin FPSpread.vaSpread spdCentralCompra 
            Height          =   3195
            Index           =   0
            Left            =   150
            TabIndex        =   63
            Top             =   255
            Width           =   5595
            _Version        =   393216
            _ExtentX        =   9869
            _ExtentY        =   5636
            _StockProps     =   64
            ColHeaderDisplay=   0
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            MaxRows         =   6
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "M_SacToWeb.frx":2B92
            TextTip         =   2
            CellNoteIndicator=   1
         End
      End
      Begin VB.Frame fraTab 
         Height          =   4650
         Index           =   1
         Left            =   -74940
         TabIndex        =   58
         Top             =   315
         Width           =   7455
         Begin EditLib.fpBoolean chkSelectAll 
            Height          =   195
            Index           =   1
            Left            =   510
            TabIndex        =   59
            Top             =   305
            Width           =   225
            _Version        =   196608
            _ExtentX        =   397
            _ExtentY        =   344
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ThreeDInsideStyle=   0
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   0
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            AutoToggle      =   -1  'True
            BooleanStyle    =   0
            ToggleFalse     =   ""
            TextFalse       =   ""
            BooleanPicture  =   2
            AlignPictureH   =   3
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   0   'False
            AlignTextH      =   0
            AlignTextV      =   1
            ToggleTrue      =   ""
            TextTrue        =   ""
            Value           =   0
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   ""
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin FPSpread.vaSpread spdCentralCompra 
            Height          =   3195
            Index           =   1
            Left            =   150
            TabIndex        =   60
            Top             =   255
            Width           =   5595
            _Version        =   393216
            _ExtentX        =   9869
            _ExtentY        =   5636
            _StockProps     =   64
            ColHeaderDisplay=   0
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            MaxRows         =   6
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "M_SacToWeb.frx":2EDD
            TextTip         =   2
            CellNoteIndicator=   1
         End
      End
      Begin VB.Frame fraTab 
         Height          =   4650
         Index           =   3
         Left            =   -75000
         TabIndex        =   32
         Top             =   345
         Width           =   7455
         Begin VB.Frame Frame1 
            Caption         =   "Período"
            ForeColor       =   &H00800000&
            Height          =   945
            Index           =   2
            Left            =   165
            TabIndex        =   51
            Top             =   210
            Width           =   4140
            Begin EditLib.fpDoubleSingle Double1 
               Height          =   315
               Index           =   0
               Left            =   540
               TabIndex        =   52
               Top             =   405
               Width           =   600
               _Version        =   196608
               _ExtentX        =   1058
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
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
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
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               Text            =   "0"
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle Double1 
               Height          =   315
               Index           =   1
               Left            =   1710
               TabIndex        =   53
               Top             =   405
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
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
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
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               Text            =   "0"
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle Double1 
               Height          =   315
               Index           =   6
               Left            =   3240
               TabIndex        =   54
               Top             =   405
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
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
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
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               Text            =   "0"
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
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
               Caption         =   "Semana"
               Height          =   195
               Index           =   8
               Left            =   2610
               TabIndex        =   57
               Top             =   480
               Width           =   585
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Año"
               Height          =   195
               Index           =   1
               Left            =   1350
               TabIndex        =   56
               Top             =   480
               Width           =   285
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Mes"
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   55
               Top             =   480
               Width           =   300
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Casinos"
            ForeColor       =   &H00800000&
            Height          =   3240
            Index           =   4
            Left            =   165
            TabIndex        =   38
            Top             =   1230
            Width           =   7125
            Begin VB.TextBox txtSearchDesc 
               Height          =   315
               Index           =   0
               Left            =   1860
               TabIndex        =   47
               Top             =   2760
               Visible         =   0   'False
               Width           =   5055
            End
            Begin VB.TextBox txtSearchCodigo 
               Height          =   315
               Index           =   0
               Left            =   1035
               TabIndex        =   46
               Top             =   2760
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.ComboBox cboCentralCompra 
               Height          =   315
               Index           =   0
               Left            =   1805
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   240
               Width           =   3210
            End
            Begin VB.CheckBox chkAllCasinos 
               Caption         =   "Todas"
               Height          =   195
               Index           =   0
               Left            =   5120
               TabIndex        =   44
               Top             =   300
               Width           =   1545
            End
            Begin VB.ComboBox cboRegional 
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   1805
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   650
               Width           =   3210
            End
            Begin VB.CheckBox chkAllRegional 
               Caption         =   "Todas"
               Enabled         =   0   'False
               Height          =   195
               Index           =   0
               Left            =   5120
               TabIndex        =   41
               Top             =   710
               Width           =   1545
            End
            Begin VB.OptionButton Option2 
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   720
               Width           =   255
            End
            Begin VB.OptionButton Option1 
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   315
               Value           =   -1  'True
               Width           =   255
            End
            Begin EditLib.fpBoolean chkSelectAll 
               Height          =   195
               Index           =   3
               Left            =   6210
               TabIndex        =   43
               Tag             =   "Casinos"
               Top             =   225
               Visible         =   0   'False
               Width           =   225
               _Version        =   196608
               _ExtentX        =   397
               _ExtentY        =   344
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ThreeDInsideStyle=   0
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   0
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               AutoToggle      =   -1  'True
               BooleanStyle    =   0
               ToggleFalse     =   ""
               TextFalse       =   ""
               BooleanPicture  =   2
               AlignPictureH   =   3
               AlignPictureV   =   1
               GroupId         =   0
               GroupTag        =   0
               GroupSelect     =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               MultiLine       =   0   'False
               AlignTextH      =   0
               AlignTextV      =   1
               ToggleTrue      =   ""
               TextTrue        =   ""
               Value           =   0
               BooleanMode     =   0
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               BorderGrayAreaColor=   -2147483637
               ToggleGrayed    =   ""
               TextGrayed      =   ""
               AllowMnemonic   =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               ThreeDOnFocusInvert=   0   'False
               Caption         =   ""
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               BooleanDataType =   0
               OLEDropMode     =   0
            End
            Begin FPSpread.vaSpread spdCasinos 
               Height          =   1640
               Index           =   0
               Left            =   180
               TabIndex        =   48
               Top             =   1045
               Width           =   6735
               _Version        =   393216
               _ExtentX        =   11880
               _ExtentY        =   2893
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
               ScrollBarExtMode=   -1  'True
               SpreadDesigner  =   "M_SacToWeb.frx":321D
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Central de Compra"
               Height          =   195
               Index           =   2
               Left            =   395
               TabIndex        =   50
               Top             =   315
               Width           =   1305
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Regional"
               Height          =   195
               Index           =   5
               Left            =   395
               TabIndex        =   49
               Top             =   720
               Width           =   630
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Tipo Solicitud"
            ForeColor       =   &H00800000&
            Height          =   945
            Index           =   3
            Left            =   4395
            TabIndex        =   33
            Top             =   210
            Width           =   2895
            Begin EditLib.fpBoolean chkTipoSolicitud 
               Height          =   300
               Index           =   0
               Left            =   195
               TabIndex        =   34
               Tag             =   "0"
               Top             =   255
               Width           =   750
               _Version        =   196608
               _ExtentX        =   1323
               _ExtentY        =   529
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ThreeDInsideStyle=   0
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   0
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               AutoToggle      =   -1  'True
               BooleanStyle    =   0
               ToggleFalse     =   ""
               TextFalse       =   "Todas"
               BooleanPicture  =   2
               AlignPictureH   =   3
               AlignPictureV   =   1
               GroupId         =   0
               GroupTag        =   0
               GroupSelect     =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               MultiLine       =   0   'False
               AlignTextH      =   0
               AlignTextV      =   1
               ToggleTrue      =   ""
               TextTrue        =   "Todas"
               Value           =   0
               BooleanMode     =   0
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               BorderGrayAreaColor=   -2147483637
               ToggleGrayed    =   ""
               TextGrayed      =   ""
               AllowMnemonic   =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               ThreeDOnFocusInvert=   0   'False
               Caption         =   "Todas"
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               BooleanDataType =   0
               OLEDropMode     =   0
            End
            Begin EditLib.fpBoolean chkTipoSolicitud 
               Height          =   300
               Index           =   1
               Left            =   195
               TabIndex        =   35
               Tag             =   "1"
               Top             =   570
               Width           =   825
               _Version        =   196608
               _ExtentX        =   1455
               _ExtentY        =   529
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ThreeDInsideStyle=   0
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   0
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               AutoToggle      =   -1  'True
               BooleanStyle    =   0
               ToggleFalse     =   ""
               TextFalse       =   "Normal"
               BooleanPicture  =   2
               AlignPictureH   =   3
               AlignPictureV   =   1
               GroupId         =   0
               GroupTag        =   0
               GroupSelect     =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               MultiLine       =   0   'False
               AlignTextH      =   0
               AlignTextV      =   1
               ToggleTrue      =   ""
               TextTrue        =   "Normal"
               Value           =   0
               BooleanMode     =   0
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               BorderGrayAreaColor=   -2147483637
               ToggleGrayed    =   ""
               TextGrayed      =   ""
               AllowMnemonic   =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               ThreeDOnFocusInvert=   0   'False
               Caption         =   "Normal"
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               BooleanDataType =   0
               OLEDropMode     =   0
            End
            Begin EditLib.fpBoolean chkTipoSolicitud 
               Height          =   300
               Index           =   2
               Left            =   1380
               TabIndex        =   36
               Tag             =   "3"
               Top             =   255
               Width           =   840
               _Version        =   196608
               _ExtentX        =   1482
               _ExtentY        =   529
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ThreeDInsideStyle=   0
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   0
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               AutoToggle      =   -1  'True
               BooleanStyle    =   0
               ToggleFalse     =   ""
               TextFalse       =   "Extra"
               BooleanPicture  =   2
               AlignPictureH   =   3
               AlignPictureV   =   1
               GroupId         =   0
               GroupTag        =   0
               GroupSelect     =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               MultiLine       =   0   'False
               AlignTextH      =   0
               AlignTextV      =   1
               ToggleTrue      =   ""
               TextTrue        =   "Extra"
               Value           =   0
               BooleanMode     =   0
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               BorderGrayAreaColor=   -2147483637
               ToggleGrayed    =   ""
               TextGrayed      =   ""
               AllowMnemonic   =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               ThreeDOnFocusInvert=   0   'False
               Caption         =   "Extra"
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               BooleanDataType =   0
               OLEDropMode     =   0
            End
            Begin EditLib.fpBoolean chkTipoSolicitud 
               Height          =   300
               Index           =   3
               Left            =   1380
               TabIndex        =   37
               Tag             =   "4"
               Top             =   570
               Width           =   1215
               _Version        =   196608
               _ExtentX        =   2143
               _ExtentY        =   529
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ThreeDInsideStyle=   0
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   0
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               AutoToggle      =   -1  'True
               BooleanStyle    =   0
               ToggleFalse     =   ""
               TextFalse       =   "Cancelación"
               BooleanPicture  =   2
               AlignPictureH   =   3
               AlignPictureV   =   1
               GroupId         =   0
               GroupTag        =   0
               GroupSelect     =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               MultiLine       =   0   'False
               AlignTextH      =   0
               AlignTextV      =   1
               ToggleTrue      =   ""
               TextTrue        =   "Cancelación"
               Value           =   0
               BooleanMode     =   0
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               BorderGrayAreaColor=   -2147483637
               ToggleGrayed    =   ""
               TextGrayed      =   ""
               AllowMnemonic   =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               ThreeDOnFocusInvert=   0   'False
               Caption         =   "Cancelación"
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               BooleanDataType =   0
               OLEDropMode     =   0
            End
         End
      End
      Begin VB.Frame fraTab 
         Height          =   4605
         Index           =   5
         Left            =   60
         TabIndex        =   11
         Top             =   345
         Width           =   7455
         Begin VB.Frame Frame1 
            Caption         =   "Casinos"
            ForeColor       =   &H00800000&
            Height          =   3240
            Index           =   14
            Left            =   165
            TabIndex        =   24
            Top             =   1230
            Width           =   7125
            Begin VB.CheckBox chkAllCasinos 
               Caption         =   "Todas"
               Height          =   195
               Index           =   1
               Left            =   4920
               TabIndex        =   28
               Top             =   300
               Width           =   1545
            End
            Begin VB.ComboBox cboCentralCompra 
               Height          =   315
               Index           =   1
               Left            =   1605
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   240
               Width           =   3210
            End
            Begin VB.TextBox txtSearchCodigo 
               Height          =   315
               Index           =   1
               Left            =   1035
               TabIndex        =   26
               Top             =   2760
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.TextBox txtSearchDesc 
               Height          =   315
               Index           =   1
               Left            =   1860
               TabIndex        =   25
               Top             =   2760
               Visible         =   0   'False
               Width           =   5055
            End
            Begin EditLib.fpBoolean chkSelectAll 
               Height          =   195
               Index           =   4
               Left            =   6345
               TabIndex        =   29
               Tag             =   "Casinos"
               Top             =   270
               Visible         =   0   'False
               Width           =   225
               _Version        =   196608
               _ExtentX        =   397
               _ExtentY        =   344
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ThreeDInsideStyle=   0
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   0
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               AutoToggle      =   -1  'True
               BooleanStyle    =   0
               ToggleFalse     =   ""
               TextFalse       =   ""
               BooleanPicture  =   2
               AlignPictureH   =   3
               AlignPictureV   =   1
               GroupId         =   0
               GroupTag        =   0
               GroupSelect     =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               MultiLine       =   0   'False
               AlignTextH      =   0
               AlignTextV      =   1
               ToggleTrue      =   ""
               TextTrue        =   ""
               Value           =   0
               BooleanMode     =   0
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               BorderGrayAreaColor=   -2147483637
               ToggleGrayed    =   ""
               TextGrayed      =   ""
               AllowMnemonic   =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               ThreeDOnFocusInvert=   0   'False
               Caption         =   ""
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               BooleanDataType =   0
               OLEDropMode     =   0
            End
            Begin FPSpread.vaSpread spdCasinos 
               Height          =   2040
               Index           =   1
               Left            =   180
               TabIndex        =   30
               Top             =   585
               Width           =   6735
               _Version        =   393216
               _ExtentX        =   11880
               _ExtentY        =   3598
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
               ScrollBarExtMode=   -1  'True
               SpreadDesigner  =   "M_SacToWeb.frx":4B0A
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Central de Compra"
               Height          =   195
               Index           =   10
               Left            =   195
               TabIndex        =   31
               Top             =   315
               Width           =   1305
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Tipo Check List"
            ForeColor       =   &H00800000&
            Height          =   945
            Index           =   9
            Left            =   4395
            TabIndex        =   19
            Top             =   210
            Width           =   2895
            Begin EditLib.fpBoolean chkTipoSolicitud 
               Height          =   300
               Index           =   4
               Left            =   195
               TabIndex        =   20
               Tag             =   "0"
               Top             =   255
               Width           =   750
               _Version        =   196608
               _ExtentX        =   1323
               _ExtentY        =   529
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ThreeDInsideStyle=   0
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   0
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               AutoToggle      =   -1  'True
               BooleanStyle    =   0
               ToggleFalse     =   ""
               TextFalse       =   "Todos"
               BooleanPicture  =   2
               AlignPictureH   =   3
               AlignPictureV   =   1
               GroupId         =   0
               GroupTag        =   0
               GroupSelect     =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               MultiLine       =   0   'False
               AlignTextH      =   0
               AlignTextV      =   1
               ToggleTrue      =   ""
               TextTrue        =   "Todos"
               Value           =   0
               BooleanMode     =   0
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               BorderGrayAreaColor=   -2147483637
               ToggleGrayed    =   ""
               TextGrayed      =   ""
               AllowMnemonic   =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               ThreeDOnFocusInvert=   0   'False
               Caption         =   "Todos"
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               BooleanDataType =   0
               OLEDropMode     =   0
            End
            Begin EditLib.fpBoolean chkTipoSolicitud 
               Height          =   300
               Index           =   5
               Left            =   195
               TabIndex        =   21
               Tag             =   "1"
               Top             =   570
               Width           =   825
               _Version        =   196608
               _ExtentX        =   1455
               _ExtentY        =   529
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ThreeDInsideStyle=   0
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   0
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               AutoToggle      =   -1  'True
               BooleanStyle    =   0
               ToggleFalse     =   ""
               TextFalse       =   "Normal"
               BooleanPicture  =   2
               AlignPictureH   =   3
               AlignPictureV   =   1
               GroupId         =   0
               GroupTag        =   0
               GroupSelect     =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               MultiLine       =   0   'False
               AlignTextH      =   0
               AlignTextV      =   1
               ToggleTrue      =   ""
               TextTrue        =   "Normal"
               Value           =   0
               BooleanMode     =   0
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               BorderGrayAreaColor=   -2147483637
               ToggleGrayed    =   ""
               TextGrayed      =   ""
               AllowMnemonic   =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               ThreeDOnFocusInvert=   0   'False
               Caption         =   "Normal"
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               BooleanDataType =   0
               OLEDropMode     =   0
            End
            Begin EditLib.fpBoolean chkTipoSolicitud 
               Height          =   300
               Index           =   6
               Left            =   1380
               TabIndex        =   22
               Tag             =   "3"
               Top             =   255
               Width           =   840
               _Version        =   196608
               _ExtentX        =   1482
               _ExtentY        =   529
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ThreeDInsideStyle=   0
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   0
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               AutoToggle      =   -1  'True
               BooleanStyle    =   0
               ToggleFalse     =   ""
               TextFalse       =   "Extra"
               BooleanPicture  =   2
               AlignPictureH   =   3
               AlignPictureV   =   1
               GroupId         =   0
               GroupTag        =   0
               GroupSelect     =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               MultiLine       =   0   'False
               AlignTextH      =   0
               AlignTextV      =   1
               ToggleTrue      =   ""
               TextTrue        =   "Extra"
               Value           =   0
               BooleanMode     =   0
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               BorderGrayAreaColor=   -2147483637
               ToggleGrayed    =   ""
               TextGrayed      =   ""
               AllowMnemonic   =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               ThreeDOnFocusInvert=   0   'False
               Caption         =   "Extra"
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               BooleanDataType =   0
               OLEDropMode     =   0
            End
            Begin EditLib.fpBoolean chkTipoSolicitud 
               Height          =   300
               Index           =   7
               Left            =   1380
               TabIndex        =   23
               Tag             =   "4"
               Top             =   570
               Width           =   1215
               _Version        =   196608
               _ExtentX        =   2143
               _ExtentY        =   529
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ThreeDInsideStyle=   0
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   0
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               AutoToggle      =   -1  'True
               BooleanStyle    =   0
               ToggleFalse     =   ""
               TextFalse       =   "Cancelación"
               BooleanPicture  =   2
               AlignPictureH   =   3
               AlignPictureV   =   1
               GroupId         =   0
               GroupTag        =   0
               GroupSelect     =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               MultiLine       =   0   'False
               AlignTextH      =   0
               AlignTextV      =   1
               ToggleTrue      =   ""
               TextTrue        =   "Cancelación"
               Value           =   0
               BooleanMode     =   0
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               BorderGrayAreaColor=   -2147483637
               ToggleGrayed    =   ""
               TextGrayed      =   ""
               AllowMnemonic   =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               ThreeDOnFocusInvert=   0   'False
               Caption         =   "Cancelación"
               ThreeDFrameColor=   -2147483633
               Appearance      =   0
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               BooleanDataType =   0
               OLEDropMode     =   0
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Período"
            ForeColor       =   &H00800000&
            Height          =   945
            Index           =   13
            Left            =   165
            TabIndex        =   12
            Top             =   210
            Width           =   4140
            Begin EditLib.fpDoubleSingle Double1 
               Height          =   315
               Index           =   4
               Left            =   540
               TabIndex        =   13
               Top             =   405
               Width           =   600
               _Version        =   196608
               _ExtentX        =   1058
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
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
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
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               Text            =   "0"
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle Double1 
               Height          =   315
               Index           =   5
               Left            =   1710
               TabIndex        =   14
               Top             =   405
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
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
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
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               Text            =   "0"
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle Double1 
               Height          =   315
               Index           =   7
               Left            =   3240
               TabIndex        =   15
               Top             =   405
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
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
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
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               Text            =   "0"
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
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
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
               Caption         =   "Año"
               Height          =   195
               Index           =   6
               Left            =   1350
               TabIndex        =   18
               Top             =   480
               Width           =   285
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Mes"
               Height          =   195
               Index           =   7
               Left            =   180
               TabIndex        =   17
               Top             =   480
               Width           =   300
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Semana"
               Height          =   195
               Index           =   9
               Left            =   2610
               TabIndex        =   16
               Top             =   480
               Width           =   585
            End
         End
      End
   End
End
Attribute VB_Name = "M_SacToWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS1 As ADODB.Recordset
Dim RS2 As ADODB.Recordset
Dim CNN As New ADODB.Connection
Dim CMD As New ADODB.Command
Dim fila_selec As Long
Dim cCompleto As Boolean
Dim tip1 As String
Dim tip2 As String
Dim tip3 As String
Dim tip4 As String
Dim blnSetDataCtrl As Boolean


Private Function blnContinuaProceso(strMessage As String) As Boolean
Dim intCurrentMousePointer As Integer

    intCurrentMousePointer = Screen.MousePointer
    Screen.MousePointer = vbDefault
    DoEvents
    blnContinuaProceso = (MsgBox(strMessage & Chr(13) & "¿ Desea continuar ?", vbQuestion + vbYesNo, App.Title) = vbYes)
    Screen.MousePointer = intCurrentMousePointer
    DoEvents
    
End Function

Private Sub GetCentralDeCompras()
Dim strSQL As String
Dim intIndex As Integer
Dim rsData As New ADODB.Recordset

    Set rsData = vg_dbpedweb.Execute("sac_s_centralcompras")
    cboCentralCompra(0).Clear
    cboCentralCompra(1).Clear
    
    spdCentralCompra(0).MaxRows = 0
    spdCentralCompra(1).MaxRows = 0
    spdCentralCompra(2).MaxRows = 0
    
    While (Not rsData.EOF)
        cboCentralCompra(0).AddItem rsData!TABCEN_CDCEN & " - " & rsData!TABCEN_DSCEN
        cboCentralCompra(1).AddItem rsData!TABCEN_CDCEN & " - " & rsData!TABCEN_DSCEN
        For intIndex = 0 To 2
            spdCentralCompra(intIndex).MaxRows = spdCentralCompra(intIndex).MaxRows + 1
            spdCentralCompra(intIndex).Row = spdCentralCompra(intIndex).MaxRows
            spdCentralCompra(intIndex).Col = 1: spdCentralCompra(intIndex).text = "0"
            spdCentralCompra(intIndex).Col = 2: spdCentralCompra(intIndex).text = rsData("TabCen_CdCen")
            spdCentralCompra(intIndex).Col = 3: spdCentralCompra(intIndex).text = rsData("TabCen_DsCen")
        Next
        rsData.MoveNext
    Wend
    rsData.Close
    Set rsData = Nothing

End Sub

Private Sub GetRegionalCBO()
Dim strSQL As String
Dim intIndex As Integer
Dim rsData As New ADODB.Recordset
    Set rsData = vg_dbpedweb.Execute("sac_s_regional 1, '0'")
    cboRegional(0).Clear
    While (Not rsData.EOF)
        cboRegional(0).AddItem rsData!tabrgi_idrgi & " - " & rsData!TABRGI_DSRGI
        rsData.MoveNext
    Wend
    
    rsData.Close
    Set rsData = Nothing
End Sub

Private Sub InicializaForm()
Dim intIndex As Integer

    blnSetDataCtrl = False
    tabMain.Tab = 0
    fraLoadData.Width = 2220
    DoEvents

    tabMain.Tab = 0
    For intIndex = 0 To tabMain.Tabs - 1
        tabMain.TabEnabled(intIndex) = False
        fraTab(intIndex).Enabled = False
    Next

    For intIndex = 0 To lstLoadData.ListCount - 1
        lstLoadData.Selected(intIndex) = False
    Next

    Double1(0).Value = Month(Date): Double1(1).Value = Year(Date)
    Double1(2).Value = Month(Date): Double1(3).Value = Year(Date)
    Double1(4).Value = Month(Date): Double1(5).Value = Year(Date)
    
    Call GetCentralDeCompras
    Call GetRegionalCBO
    
    chkSelectAll(0).Value = ValueTrue
    chkSelectAll(1).Value = ValueTrue
    chkSelectAll(2).Value = ValueTrue
    chkSelectAll(3).Value = ValueTrue
    chkSelectAll(4).Value = ValueTrue
    
    cboCentralCompra(0).ListIndex = -1
    cboCentralCompra(1).ListIndex = -1
    chkAllCasinos(0).Value = 1
    chkAllCasinos(1).Value = 1
    chkAllRegional(0).Value = 1
    Call cboCentralCompra_Click(0)
    Call cboCentralCompra_Click(1)
    
    Screen.MousePointer = vbDefault
    DoEvents

End Sub

Private Function blnLoadDataCasinos() As Boolean

Const strSQLSAC = "SELECT [:FIELDS] " & _
                  "FROM CadFil " & _
                  "WHERE TabCen_CdCen = '[:COD_CENTRAL_COMPRA]'" & _
                  "AND CadFil_FlBlo = 0 "

Dim strSQL As String
Dim strCentralCompra As String
Dim intRow As Integer
Dim rsData As New ADODB.Recordset
Dim lngCountReg As Long

    blnLoadDataCasinos = False

    Print #1, "DELETE FROM FROM s_Clientes "
    
    Call SetViewStatusProgress(True, True)
    lstLoadData.ListIndex = 0
    DoEvents

    prbMain.Max = spdCentralCompra(0).MaxRows

    For intRow = 1 To spdCentralCompra(0).MaxRows
        
        lblStatusMain.Caption = ""
        lblStatusOther.Caption = ""
        prbOther.Value = 0
        DoEvents
        
        spdCentralCompra(0).Row = intRow
        spdCentralCompra(0).Col = 1
        
        If (spdCentralCompra(0).Value = 1) Then
            
            spdCentralCompra(0).Col = 3
            strCentralCompra = spdCentralCompra(0).text
            lblStatusMain.Caption = "Cargando Casinos: Central de Compra " & strCentralCompra
            prbMain.Value = intRow
            lblStatusOther.Caption = ""
            prbOther.Value = 0
            DoEvents
            
            spdCentralCompra(0).Col = 2
            strSQL = Replace(strSQLSAC, "[:COD_CENTRAL_COMPRA]", spdCentralCompra(0).text)
            rsData.Open Replace(strSQL, "[:FIELDS]", "COUNT(*)"), vg_dbsac, adOpenStatic
            lngCountReg = rsData(0)
            rsData.Close
            
            If (lngCountReg = 0) Then
                SendMsg "Cargando Casinos" & Chr(13) & _
                        "No existen Casinos asociados a Central de Compra " & strCentralCompra, vbExclamation
            End If

            prbOther.Max = lngCountReg
            If (lngCountReg > 0) Then
                rsData.Open Replace(strSQL, "[:FIELDS]", "*"), vg_dbsac, adOpenStatic
                While (Not rsData.EOF)
                    lblStatusOther.Caption = "Casino: " & strCentralCompra
                    prbOther.Value = intRow
                    DoEvents
                    
                    Print #1, "INSERT INTO s_Clientes( Codigo, CentralDeCompra, IdServicio, " & _
                              "IdRegional, IdRegion, TipoAdministracion, Nombre, Supervisor, " & _
                              "Administrador, EntregaDomingo, EntregaLunes, EntregaMartes, " & _
                              "EntregaMiercoles, EntregaJueves, EntregaViernes, EntregaSabado, " & _
                              "Bloqueo, RegimenEspecial, CentroCosto ) " & _
                              "VALUES('" & RS1!CADFIL_IDFIL & "', '" & RS1!TABCEN_CDCEN & "', '" & _
                              RS1!TABSER_IDSER & "', '" & RS1!tabrgi_idrgi & "', '" & _
                              RS1!TABREG_IDREG & "', '" & RS1!TABADM_IDADM & "', '" & _
                              RS1!CADFIL_NMFIL & "', '" & TipoDato(RS1!CADFIL_NMSUP, "") & "', '" & _
                              TipoDato(RS1!CADFIL_NMADM, "") & "', " & _
                              TipoDato(RS1!CADFIL_FLDOM, 0) & ", " & _
                              TipoDato(RS1!CADFIL_FLSEG, 0) & ", " & _
                              TipoDato(RS1!CADFIL_FLTER, 0) & ", " & _
                              TipoDato(RS1!CADFIL_FLQUA, 0) & ", " & _
                              TipoDato(RS1!CADFIL_FLQUI, 0) & ", " & _
                              TipoDato(RS1!CADFIL_FLSEX, 0) & ", " & _
                              TipoDato(RS1!CADFIL_FLSAB, 0) & ", " & _
                              TipoDato(RS1!CADFIL_FLBLO, 0) & ", " & _
                              TipoDato(RS1!CADFIL_FLREG, 0) & ", '" & RS1!CADFIL_CDFIL & "')"
                    
                    rsData.MoveNext
                Wend
            End If  ' If (lngCountReg > 0) Then
            rsData.Close
        
        End If  ' If (spdCentralCompra(0).Value = 1) Then
    
    Next
                       
End Function

Private Sub LoadData()
Dim i As Long, z As Long, k As Long, cCas As String, cCen As String, nCen As String, cCat As String, nCat As String
Dim nReg As Long, nSem As Long, lAct As Double, cFec As Date
Dim cPer As String, nFil As Long, nPed As Long, cArc As String, nIdPe As Long
Dim strSQL As String, nTip As String
Dim pArch As String, nArch As String
Dim spid As Long
Dim Est As Boolean
Dim cdcen As String
        
On Error GoTo ManError

    fraLoadData.Width = 9990
    Screen.MousePointer = vbHourglass
    DoEvents
    
    If (lstLoadData.Selected(0)) Then   ' Casinos
        tabMain.Tab = 0
        prbMain.Min = 1: prbMain.Max = spdCentralCompra(0).MaxRows: prbMain.Visible = True
        Est = True
        For i = 1 To spdCentralCompra(0).MaxRows
            DoEvents
            spdCentralCompra(0).Row = i: spdCentralCompra(0).Col = 1
            If spdCentralCompra(0).Value = 1 Then
                spdCentralCompra(0).Col = 2: cCen = Trim(spdCentralCompra(0).Value)
                lblStatusMain.Visible = True: lblStatusMain.Caption = "Cargando central de compras " & cCen
                
                If Est Then
                   '-----> Borrar tabla paso_importacion casino -----
                   vg_dbpedweb.Execute "DELETE paso_casino WHERE cas_spid = @@spid AND cas_usr = '" & vg_NUsr & "'"
                   Set RS1 = vg_dbpedweb.Execute("SELECT @@spid spid")
                   If Not RS1.EOF Then spid = RS1!spid
                   RS1.Close: Set RS1 = Nothing
                End If
                Est = False
                Set RS1 = vg_dbpedweb.Execute("sac_s_filial 2, '" & cCen & "'")
                nReg = TipoDato(RS1!nReg, 0)
                RS1.Close: Set RS1 = Nothing
                
                If nReg = 0 Then
                    If (Not blnContinuaProceso("No existen Casinos para Central de Compra: " & cCen)) Then GoTo AbortLoadData
                End If
                
                If nReg <> 0 Then
                    Set RS1 = vg_dbpedweb.Execute("sac_s_filial 3, '" & cCen & "'")
                    prbOther.Min = 0: prbOther.Max = nReg: lblStatusOther.Visible = True: prbOther.Visible = True: z = 1
                    Do While Not RS1.EOF
                        lblStatusOther.Caption = "Cargando casinos " & Trim(Str(z)) & "/" & Trim(Str(nReg))
                        DoEvents
                          vg_dbpedweb.Execute ("INSERT INTO paso_casino VALUES ('" & vg_NUsr & "', " & spid & ", '" & RS1!CADFIL_IDFIL & "')")
                        RS1.MoveNext: prbOther.Value = z: z = z + 1
                    Loop
                    RS1.Close: Set RS1 = Nothing
                    prbOther.Visible = False: lblStatusOther.Visible = False
                End If
            End If
            prbMain.Value = i
        Next i
        prbMain.Visible = False: lblStatusMain.Visible = False
        If (lstLoadData.Selected(0)) And Not Est Then   ' Casinos
            vg_dbpedweb.Execute ("pedweb_p_importarcasino '" & vg_NUsr & "', " & spid & "")
        End If
    End If
    
    If (lstLoadData.Selected(1)) Then   ' Productos
        tabMain.Tab = 1
        prbMain.Min = 1: prbMain.Max = spdCentralCompra(1).MaxRows: prbMain.Visible = True
        Set RS1 = vg_dbpedweb.Execute("sac_s_familiaproductos")
        nReg = TipoDato(RS1!nReg, 0)
        RS1.Close: Set RS1 = Nothing
        If nReg <> 0 Then
           lblStatusOther.Caption = "Un Momento Cargando Familias Productos ..."
           vg_dbpedweb.Execute ("pedweb_p_importarfamiliaproductos")
        End If
        Set RS1 = vg_dbpedweb.Execute("sac_s_unidadmedida")
        nReg = TipoDato(RS1!nReg, 0)
        RS1.Close: Set RS1 = Nothing
        If nReg <> 0 Then
           lblStatusOther.Caption = "Un Momento Cargando Unidades Medida...."
           vg_dbpedweb.Execute ("pedweb_p_importarunidadmedida")
        End If
        
        Set RS1 = vg_dbpedweb.Execute("sac_s_productosraiz")
        nReg = TipoDato(RS1!nReg, 0)
        RS1.Close: Set RS1 = Nothing
        If nReg <> 0 Then
           lblStatusOther.Caption = "Un Momento Cargando Productos Raiz...."
           vg_dbpedweb.Execute ("pedweb_p_importarproductosraiz")
        End If
        
        For i = 1 To spdCentralCompra(1).MaxRows
            spdCentralCompra(1).Row = i: spdCentralCompra(1).Col = 1
            If spdCentralCompra(1).Value = 1 Then
                spdCentralCompra(1).Col = 2: cCen = Trim(spdCentralCompra(1).Value)
                spdCentralCompra(1).Col = 3: nCen = Trim(spdCentralCompra(1).Value)
                lblStatusMain.Visible = True: lblStatusMain.Caption = "Cargando Productos Central de Compra " & nCen
                lblStatusOther.Caption = ""
                prbOther.Visible = False
                nReg = 0
                Set RS1 = vg_dbpedweb.Execute("sac_s_productoscentral '" & cCen & "', ''")
                If Not RS1.EOF Then nReg = TipoDato(RS1!nReg, 0)
                RS1.Close: Set RS1 = Nothing
                If nReg <> 0 Then
                    prbOther.Min = 0: prbOther.Max = nReg: lblStatusOther.Visible = True: prbOther.Visible = True: z = 1
                    vg_dbpedweb.Execute ("pedweb_p_importarproductoscentral '" & cCen & "', ''")
                End If
                
                nReg = 0
                
                '*******************************************************************************************
                ' Por uso de Productos en CheckList, se cargan productos incluso cuya vigencia
                ' caducó hace 180 días.
                '*******************************************************************************************
                Set RS1 = vg_dbpedweb.Execute("sac_s_productosload '" & cCen & "'")
                If Not RS1.EOF Then nReg = TipoDato(RS1!nReg, 0)
                RS1.Close: Set RS1 = Nothing
                If nReg <> 0 Then
                   lblStatusOther.Caption = "Cargando productos "
                    vg_dbpedweb.Execute ("pedweb_p_importarproductosload '" & cCen & "'")
                End If
            End If
            prbMain.Value = i
        Next i
        prbMain.Visible = False: lblStatusMain.Visible = False
        prbOther.Visible = False: lblStatusOther.Visible = False
        vg_dbpedweb.Execute ("pedweb_udi_importarproductosload")
    End If
    
    If (lstLoadData.Selected(2)) Then   ' Calendario Compras
        tabMain.Tab = 2
        cPer = fg_pone_cero(Str(Double1(3).Value), 4) & fg_pone_cero(Str(Double1(2).Value), 2)
        prbMain.Min = 1: prbMain.Max = spdCentralCompra(2).MaxRows: prbMain.Visible = True
        For i = 1 To spdCentralCompra(1).MaxRows
            spdCentralCompra(2).Row = i: spdCentralCompra(2).Col = 1
            If spdCentralCompra(2).Value = 1 Then
                spdCentralCompra(2).Col = 2: cCen = Trim(spdCentralCompra(2).Value)
                lblStatusMain.Visible = True: lblStatusMain.Caption = "Cargando calendario de compras central : " & cCen
'                vg_dbpedweb.Execute "DELETE FROM s_calendario WHERE anomes='" & cPer & "' and centralcompra='" & cCen & "'"
'                Set RS1 = vg_dbsac.Execute("SELECT COUNT(*) as nReg FROM CICCOT WHERE CICCPA_DTREF='" & cPer & "' and TABCEN_CDCEN='" & cCen & "'")
                Set RS1 = vg_dbpedweb.Execute("sac_s_calendariocompras '" & cPer & "', '" & cCen & "'")
                nReg = TipoDato(RS1!nReg, 0)
                RS1.Close: Set RS1 = Nothing
                If nReg <> 0 Then
'                    Set RS1 = vg_dbsac.Execute("SELECT * FROM CICCOT WHERE CICCPA_DTREF='" & cPer & "' and TABCEN_CDCEN='" & cCen & "'")
                    prbOther.Min = 0: prbOther.Max = nReg: lblStatusOther.Visible = True: prbOther.Visible = True: z = 1
'                    Do While Not RS1.EOF
                        lblStatusOther.Caption = "Un momento, Cargando registros "
                        DoEvents
                        vg_dbpedweb.Execute ("pedweb_p_importarcalendariocompras '" & cPer & "', '" & cCen & "'")
'                        vg_dbpedweb.Execute "INSERT INTO s_calendario (codigo, centralcompra, anomes, semana, familia, categoria, periododesde, periodohasta, fecha, estado, control) " & _
'                                  "VALUES( '" & RS1!CICCOT_IDCOT & "', '" & RS1!TABCEN_CDCEN & "', '" & RS1!CICCPA_DTREF & "', " & RS1!CICCPA_NRSEM & ", '" & RS1!TABFAM_IDFAM & "', '" & RS1!TABCAT_IDCAT & "', convert(datetime,'" & RS1!CICCOT_DTPDE & "',103), convert(datetime,'" & RS1!CICCOT_DTPAT & "',103), convert(datetime,'" & RS1!CICCOT_DTSTA & "',103), '" & TipoDato(RS1!CICCOT_FLSTA, 0) & "', '" & Abs(TipoDato(RS1!CICCOT_FLATU, 0)) & "' )"
'                        RS1.MoveNext: prbOther.Value = z: z = z + 1
'                    Loop
'                    RS1.Close: Set RS1 = Nothing
                    prbOther.Visible = False: lblStatusOther.Visible = False
                End If
            End If
            prbMain.Value = i
        Next i
        prbMain.Visible = False: lblStatusMain.Visible = False
    End If
    
    If (lstLoadData.Selected(3)) Then   ' Solicitudes Compras
        tabMain.Tab = 3
        cPer = fg_pone_cero(Double1(1).Value, 4) & fg_pone_cero(Double1(0).Value, 2)
        nSem = Val(Double1(6).Value)
        nTip = strGetTipoSolicitudes(0)
        If tip1 = "0" Then tip1 = "9"
        If tip1 = "1" And tip2 = "3" And tip3 = "4" Then tip1 = "0": tip2 = "0": tip3 = "0"
        cdcen = ""
        If chkAllCasinos(0).Value = 0 Then cdcen = Trim(Replace(Mid(cboCentralCompra(0).text, 1, 3), "-", ""))
        Set RS1 = vg_dbpedweb.Execute("sac_s_solicituddecompras '" & cPer & "', '" & nSem & "', '" & tip1 & "', '" & tip2 & "', '" & tip3 & "', '" & cdcen & "'")
        If TipoDato(RS1!nReg, 0) = 0 Then
            If (Not blnContinuaProceso("No existen Solicitudes de Compras")) Then
                RS1.Close: Set RS1 = Nothing
                GoTo AbortLoadData
            End If
        End If
        If TipoDato(RS1!nReg, 0) > 0 Then
            lblStatusMain.Visible = True
            lblStatusOther.Visible = False: prbOther.Visible = True: prbOther.Min = 0: prbOther.Max = TipoDato(RS1!nReg, 0): z = 0
'            If tip1 = "0" Then tip1 = "9"
'            If tip1 = "1" And tip2 = "3" And tip3 = "4" Then tip1 = "0": tip2 = "0": tip3 = "0"
            cdcen = ""
            If chkAllCasinos(0).Value = 0 Then cdcen = Trim(Replace(Mid(cboCentralCompra(0).text, 1, 3), "-", ""))
             DoEvents
             lblStatusMain.Caption = "Un Momento, Procesando Solicitud de Compra ...."
             vg_dbpedweb.Execute ("pedweb_p_importarsolicituddecompras '" & cPer & "', '" & nSem & "', '" & tip1 & "', '" & tip2 & "', '" & tip3 & "', '" & cdcen & "'")
             DoEvents
            prbMain.Visible = False: lblStatusMain.Visible = False
            prbOther.Visible = False: lblStatusOther.Visible = False
        End If
        RS1.Close: Set RS1 = Nothing
        
    End If
    
    If (lstLoadData.Selected(4)) Then   ' Proveedores
        tabMain.Tab = 4
        Set RS1 = vg_dbpedweb.Execute("sac_s_proveedores 1")
        If (TipoDato(RS1!nReg, 0) = 0) Then
            If (Not blnContinuaProceso("No existen Proveedores")) Then
                RS1.Close: Set RS1 = Nothing
                GoTo AbortLoadData
            End If
        Else
           lblStatusMain.Visible = True: prbMain.Visible = True: prbMain.Min = 0: prbMain.Max = TipoDato(RS1!nReg, 0): z = 0
           RS1.Close: Set RS1 = Nothing
           DoEvents
           lblStatusMain.Caption = "Un momento, Cargando Proveedores "
           vg_dbpedweb.Execute ("pedweb_p_importarproveedores")
           prbMain.Visible = False: lblStatusMain.Visible = False
        End If
    End If
    
    If (lstLoadData.Selected(5)) Then   ' Check List
        tabMain.Tab = 5
        cPer = fg_pone_cero(Double1(5).Value, 4) & fg_pone_cero(Double1(4).Value, 2)
        nSem = Val(Double1(7).Value)
        nTip = strGetTipoSolicitudes(4)
        strSQL = "DELETE s_checklist FROM "
        strSQL = strSQL & "s_checklist INNER JOIN "
        strSQL = strSQL & "s_checklistItems ON s_checklist.numeropedido = s_checklistItems.numero INNER JOIN "
        strSQL = strSQL & "s_SolicitudDeCompra ON s_checklistItems.solicitudcompra = s_SolicitudDeCompra.codigo "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "s_checklist.anomes = '" & cPer & "' "
        strSQL = strSQL & "AND s_checklist.semana = " & nSem & " "
        strSQL = strSQL & IIf(chkAllCasinos(1).Value = 0, "AND s_checklist.centrocompra = '" & Trim(Replace(Mid(cboCentralCompra(1).text, 1, 3), "-", "")) & "' ", "")
        strSQL = strSQL & "AND s_SolicitudDeCompra.tipo IN ( " & nTip & " ) "
        lblStatusMain.Caption = "Buscando datos de Check List..."
        lblStatusMain.Visible = True
        DoEvents
        cdcen = ""
        If chkAllCasinos(1).Value = 0 Then cdcen = Trim(Replace(Mid(cboCentralCompra(1).text, 1, 3), "-", ""))
        Set RS1 = vg_dbpedweb.Execute("sac_s_checklist '" & cPer & "', '" & nSem & "', '" & cdcen & "'")
        If (TipoDato(RS1!nReg, 0) = 0) Then
            If (Not blnContinuaProceso("No existen Check List.")) Then
                RS1.Close: Set RS1 = Nothing
                GoTo AbortLoadData
            End If
        Else
            lblStatusMain.Visible = True: lblStatusOther.Visible = True: prbOther.Visible = True: prbOther.Min = 0: prbOther.Max = TipoDato(RS1!nReg, 0): z = 0
            RS1.Close: Set RS1 = Nothing
            cdcen = ""
            If chkAllCasinos(1).Value = 0 Then cdcen = Trim(Replace(Mid(cboCentralCompra(1).text, 1, 3), "-", ""))
            lblStatusMain.Caption = "Un Momento Procesando Chek List....."
            lblStatusOther.Caption = ""
            vg_dbpedweb.Execute ("pedweb_p_importarcheklistload " & cPer & " , " & nSem & ", '" & cdcen & "'")
        End If
        Set RS1 = Nothing
    End If
    prbOther.Visible = False: lblStatusOther.Visible = False
 
    '-------> Proceso Grabado
    '-------> Fin Proceso Grabado
    prbMain.Visible = False: lblStatusMain.Visible = False
    prbOther.Visible = False: lblStatusOther.Visible = False
    Screen.MousePointer = vbDefault
    DoEvents

    Exit Sub
    
AbortLoadData:
    prbMain.Visible = False: lblStatusMain.Visible = False
    prbOther.Visible = False: lblStatusOther.Visible = False
    Screen.MousePointer = vbDefault
'    If Dir(pArch) <> "" Then Kill (pArch)
    Exit Sub
    
ManError:
    Select Case Err
        Case 35764
            DoEvents
            For i = 1 To 1000000
            Next i
            Resume
        Case Else
            Screen.MousePointer = vbDefault
            DoEvents
            MsgBox "Error : " & Err & ", " & Err.Description & "...", vbInformation & vbOKOnly, App.Title
    End Select

End Sub

Private Sub SendMsg(strMsg As String, intButtons As Integer)
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox strMsg, intButtons, App.Title
End Sub

Private Sub SetViewStatusProgress(blnMain As Boolean, Optional blnOther = False)
    lblStatusMain.Caption = ""
    lblStatusMain.Visible = blnMain
    prbMain.Value = 0
    prbMain.Visible = blnMain
    lblStatusOther.Caption = ""
    lblStatusOther.Visible = blnOther
    prbOther.Value = 0
    prbOther.Visible = blnOther
    DoEvents
End Sub

Private Function strGetTipoSolicitudes(intIndexCtrlMain As Integer) As String
Dim strTmp As String

    strTmp = chkTipoSolicitud(intIndexCtrlMain + 1).Tag & ","
    tip1 = chkTipoSolicitud(intIndexCtrlMain + 1).Tag
    strTmp = strTmp & chkTipoSolicitud(intIndexCtrlMain + 2).Tag & ","
    tip2 = chkTipoSolicitud(intIndexCtrlMain + 2).Tag
    strTmp = strTmp & chkTipoSolicitud(intIndexCtrlMain + 3).Tag
    tip3 = chkTipoSolicitud(intIndexCtrlMain + 3).Tag
    If (chkTipoSolicitud(intIndexCtrlMain).Value = ValueFalse) Then
        strTmp = ""
        tip1 = "0": tip2 = "0": tip3 = "0"
        If (chkTipoSolicitud(intIndexCtrlMain + 1).Value = ValueTrue) Then
           strTmp = strTmp & chkTipoSolicitud(intIndexCtrlMain + 1).Tag & ","
           tip1 = chkTipoSolicitud(intIndexCtrlMain + 1).Tag
        End If
        If (chkTipoSolicitud(intIndexCtrlMain + 2).Value = ValueTrue) Then
           strTmp = strTmp & chkTipoSolicitud(intIndexCtrlMain + 2).Tag & ","
           tip2 = chkTipoSolicitud(intIndexCtrlMain + 2).Tag
        End If
        If (chkTipoSolicitud(intIndexCtrlMain + 3).Value = ValueTrue) Then
           strTmp = strTmp & chkTipoSolicitud(intIndexCtrlMain + 3).Tag & ","
           tip3 = chkTipoSolicitud(intIndexCtrlMain + 3).Tag
        End If
        If (Trim(strTmp) <> "") Then strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
    End If
    
    strGetTipoSolicitudes = strTmp

End Function

Private Sub cboRegional_Click(Index As Integer)
Dim rsData As New ADODB.Recordset
Dim strSQL As String
    strSQL = "-1"
    If (chkAllRegional(Index).Value = 0) Then strSQL = Mid(cboRegional(Index).text, 1, 2)
    Set rsData = vg_dbpedweb.Execute("sac_s_filial 1, '" & strSQL & "'")
    spdCasinos(Index).MaxRows = 0
    While (Not rsData.EOF)
        spdCasinos(Index).MaxRows = spdCasinos(Index).MaxRows + 1
        spdCasinos(Index).Row = spdCasinos(Index).MaxRows
        'spdCasinos(Index).Col = 1: spdCasinos(Index).Value = 1
        spdCasinos(Index).Col = 1: spdCasinos(Index).text = CStr(rsData("TABCEN_CDCEN"))
        spdCasinos(Index).Col = 2: spdCasinos(Index).text = CStr(rsData("CADFIL_CDFIL"))
        spdCasinos(Index).Col = 3: spdCasinos(Index).text = CStr(rsData("CADFIL_NMFIL"))
        rsData.MoveNext
    Wend
    rsData.Close
    Set rsData = Nothing

End Sub

Private Sub chkAllCasinos_Click(Index As Integer)

    If (chkAllCasinos(Index).Value = 1) Then
        cboCentralCompra(Index).ListIndex = -1
    End If
    
    spdCasinos(Index).Enabled = (chkAllCasinos(Index).Value = 0)
    cboCentralCompra(Index).Enabled = (chkAllCasinos(Index).Value = 0)
    If (Index = 0) Then
        chkSelectAll(3).Enabled = (chkAllCasinos(Index).Value = 0)
    Else
        chkSelectAll(4).Enabled = (chkAllCasinos(Index).Value = 0)
    End If

End Sub


Private Sub chkAllRegional_Click(Index As Integer)
    If (chkAllRegional(Index).Value = 1) Then
        cboRegional(Index).ListIndex = -1
    End If
    
    spdCasinos(Index).Enabled = (chkAllRegional(Index).Value = 0)
    cboRegional(Index).Enabled = (chkAllRegional(Index).Value = 0)
    If (Index = 0) Then
        chkSelectAll(3).Enabled = (chkAllRegional(Index).Value = 0)
    Else
        chkSelectAll(4).Enabled = (chkAllRegional(Index).Value = 0)
    End If

End Sub

Private Sub chkSelectAll_Change(Index As Integer)

    If (chkSelectAll(Index).Tag = "") Then
        spdCentralCompra(Index).Row = 1
        spdCentralCompra(Index).Col = 1
        spdCentralCompra(Index).Row2 = spdCentralCompra(Index).MaxRows
        spdCentralCompra(Index).Col2 = 1
        spdCentralCompra(Index).BlockMode = True
        spdCentralCompra(Index).Value = IIf(chkSelectAll(Index).Value = ValueTrue, 1, 0)
        spdCentralCompra(Index).BlockMode = False
    Else
        spdCasinos(Index - 3).Row = 1
        spdCasinos(Index - 3).Col = 1
        spdCasinos(Index - 3).Row2 = spdCasinos(Index - 3).MaxRows
        spdCasinos(Index - 3).Col2 = 1
        spdCasinos(Index - 3).BlockMode = True
        spdCasinos(Index - 3).Value = IIf(chkSelectAll(Index).Value = ValueTrue, 1, 0)
        spdCasinos(Index - 3).BlockMode = False
    End If

End Sub

Private Sub chkTipoSolicitud_Change(Index As Integer)
Dim intIndex As Integer

    If (blnSetDataCtrl) Then Exit Sub

    intIndex = 0
    If (Index >= 4) Then intIndex = 4

    blnSetDataCtrl = True
    If (chkTipoSolicitud(Index).Tag = "0") Then
        chkTipoSolicitud(intIndex + 1).Value = chkTipoSolicitud(Index).Value
        chkTipoSolicitud(intIndex + 2).Value = chkTipoSolicitud(Index).Value
        chkTipoSolicitud(intIndex + 3).Value = chkTipoSolicitud(Index).Value
    Else
        chkTipoSolicitud(intIndex).Value = IIf((chkTipoSolicitud(intIndex + 1).Value = ValueTrue) And _
                                               (chkTipoSolicitud(intIndex + 2).Value = ValueTrue) And _
                                               (chkTipoSolicitud(intIndex + 3).Value = ValueTrue), _
                                               ValueTrue, ValueFalse)
    End If
    blnSetDataCtrl = False

End Sub

Private Sub cboCentralCompra_Click(Index As Integer)
Dim rsData As New ADODB.Recordset
Dim strSQL As String
    strSQL = "X"
    If (chkAllCasinos(Index).Value = 0) Then strSQL = Trim(Replace(Mid(cboCentralCompra(Index).text, 1, 3), "-", ""))
    Set rsData = vg_dbpedweb.Execute("sac_s_filtrocentralcompra '" & strSQL & "'")
    spdCasinos(Index).MaxRows = 0
    While (Not rsData.EOF)
        spdCasinos(Index).MaxRows = spdCasinos(Index).MaxRows + 1
        spdCasinos(Index).Row = spdCasinos(Index).MaxRows
        'spdCasinos(Index).Col = 1: spdCasinos(Index).Value = 1
        spdCasinos(Index).Col = 1: spdCasinos(Index).text = CStr(rsData("TABCEN_CDCEN"))
        spdCasinos(Index).Col = 2: spdCasinos(Index).text = CStr(rsData("CADFIL_CDFIL"))
        spdCasinos(Index).Col = 3: spdCasinos(Index).text = CStr(rsData("tabrgi_dsrgi"))
        spdCasinos(Index).Col = 4: spdCasinos(Index).text = CStr(rsData("CADFIL_NMFIL"))
        rsData.MoveNext
    Wend
    rsData.Close
    Set rsData = Nothing
End Sub

Private Function GetRegional(idReg As Long) As String
Dim strSQL As String
Dim rsData As New ADODB.Recordset
    Set rsData = vg_dbpedweb.Execute("sac_s_regional 2, '" & idReg & "'")
    If Not rsData.EOF Then
        GetRegional = rsData("TABRGI_DSRGI")
    Else
        GetRegional = "-"
    End If
    
    rsData.Close
    Set rsData = Nothing

End Function

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()

    App.Title = "SacToWeb"
    fg_centra Me
    
    '-------> Abrir base sac
    'AbrirBaseSac
   
    Call InicializaForm
    
End Sub



Private Sub Inet1_StateChanged(ByVal State As Integer)
Select Case State
Case 5
    'envio de solicitud
    cCompleto = False
Case 9
    'desconexion
Case 12
    'solicitud completada
    cCompleto = True
End Select
End Sub

Private Sub lstLoadData_ItemCheck(Item As Integer)

    tabMain.TabEnabled(Item) = lstLoadData.Selected(Item)
    If (tabMain.TabEnabled(Item)) Then tabMain.Tab = Item
    fraTab(Item).Enabled = lstLoadData.Selected(Item)
'    MsgBox Item

End Sub


Private Sub Option1_Click()
chkAllCasinos(0).Enabled = True
cboRegional(0).Enabled = False
chkAllRegional(0).Enabled = False
    Call GetCentralDeCompras
    Call GetRegionalCBO
'    Call cboCentralCompra_Click(0)
    chkAllCasinos(0).Value = 1
    Call cboCentralCompra_Click(0)
End Sub

Private Sub Option2_Click()
cboCentralCompra(0).Enabled = False
chkAllCasinos(0).Enabled = False
chkAllRegional(0).Enabled = True
    Call GetCentralDeCompras
    Call GetRegionalCBO
    chkAllRegional(0).Value = 1
    Call cboRegional_Click(0)
End Sub

Private Sub tlbAction_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long
Dim estsel As Boolean

    Select Case Button.Tag
        Case "Load"
                    
            If (lstLoadData.Selected(0)) Then
               estsel = False
               For i = 1 To spdCentralCompra(0).MaxRows
                   spdCentralCompra(0).Row = i
                   spdCentralCompra(0).Col = 1
                   If spdCentralCompra(0).text = "1" Then estsel = True
               Next i
               If Not estsel Then tabMain.Tab = 1: MsgBox "Debe seleccionar a lo mes un cantral de compras.", vbCritical, App.Title: Exit Sub
            End If
            
            If (lstLoadData.Selected(2)) Then   ' Calendario Compras
                If Val(Double1(2).Value) = 0 Or Val(Double1(3).Value) = 0 Then tabMain.Tab = 2: MsgBox "Debe ingresar mes y año del calendario.", vbCritical, App.Title: Exit Sub
            End If
            
            If (lstLoadData.Selected(3)) Then   ' Solicitudes Compras
                If Val(Double1(0).Value) = 0 Or Val(Double1(1).Value) = 0 Then tabMain.Tab = 3: MsgBox "Debe ingresar mes y año del calendario.", vbCritical, App.Title: Exit Sub
                If Val(Double1(6).Value) = 0 Then tabMain.Tab = 3: MsgBox "Debe ingresar semana.", vbCritical, App.Title: Exit Sub
                If (strGetTipoSolicitudes(0) = "") Then
                    tabMain.Tab = 3
                    MsgBox "Debe seleccionar al menos un Tipo de Solicitud.", vbCritical, App.Title
                    Exit Sub
                End If
            End If
            
            If (lstLoadData.Selected(5)) Then   ' Check List
                If Val(Double1(4).Value) = 0 Or Val(Double1(5).Value) = 0 Then tabMain.Tab = 5: MsgBox "Debe ingresar mes y año de los Check List.", vbCritical, App.Title: Exit Sub
                If Val(Double1(7).Value) = 0 Then tabMain.Tab = 5: MsgBox "Debe ingresar semana.", vbCritical, App.Title: Exit Sub
                If (strGetTipoSolicitudes(4) = "") Then
                    tabMain.Tab = 5
                    MsgBox "Debe seleccionar al menos un Tipo de Solicitud.", vbCritical, App.Title
                    Exit Sub
                End If
            End If
            
            Call LoadData
            MsgBox "Proceso finalizado OK.", vbInformation, "SacToWeb"
            Call InicializaForm
    
        Case "Cancel": Call InicializaForm
        
        Case "Close": Unload Me
    
    End Select

End Sub

