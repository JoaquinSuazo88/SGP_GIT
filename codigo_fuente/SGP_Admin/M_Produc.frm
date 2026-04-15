VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_Produc 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producto"
   ClientHeight    =   10260
   ClientLeft      =   4065
   ClientTop       =   1230
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9615
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   16960
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
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
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "Shape1(0)"
      Tab(0).Control(3)=   "Label5(0)"
      Tab(0).Control(4)=   "Shape1(1)"
      Tab(0).Control(5)=   "Label5(1)"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_Produc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "fpText1(1)"
      Tab(1).Control(3)=   "fpText1(4)"
      Tab(1).Control(4)=   "fpText1(3)"
      Tab(1).Control(5)=   "Label3(12)"
      Tab(1).Control(6)=   "Label3(4)"
      Tab(1).Control(7)=   "Label3(2)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Ingrediente"
      TabPicture(2)   =   "M_Produc.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblNomPro(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame1(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame8"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Formato Compra xxx"
      TabPicture(3)   =   "M_Produc.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblNomPro(1)"
      Tab(3).Control(1)=   "Frame9"
      Tab(3).Control(2)=   "Frame10"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Formato Compras SAC"
      TabPicture(4)   =   "M_Produc.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame11"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Formato Compras SAP"
      TabPicture(5)   =   "M_Produc.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame14"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Left            =   -74880
         TabIndex        =   142
         Top             =   900
         Width           =   10815
         Begin MSComctlLib.Toolbar Toolbar5 
            Height          =   360
            Left            =   7380
            TabIndex        =   143
            Top             =   240
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   635
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BorderStyle     =   1
         End
         Begin FPSpread.vaSpread vaSpread10 
            Height          =   5895
            Left            =   120
            TabIndex        =   144
            Top             =   720
            Width           =   10545
            _Version        =   393216
            _ExtentX        =   18600
            _ExtentY        =   10398
            _StockProps     =   64
            ButtonDrawMode  =   2
            DisplayRowHeaders=   0   'False
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
            MaxCols         =   8
            MaxRows         =   5
            SpreadDesigner  =   "M_Produc.frx":00A8
         End
         Begin FPSpread.vaSpread vaSpread9 
            Height          =   375
            Left            =   1080
            TabIndex        =   145
            Top             =   5880
            Visible         =   0   'False
            Width           =   615
            _Version        =   393216
            _ExtentX        =   1085
            _ExtentY        =   661
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
            MaxCols         =   1
            MaxRows         =   1
            SpreadDesigner  =   "M_Produc.frx":0881
         End
      End
      Begin VB.Frame Frame10 
         Height          =   1965
         Left            =   -74610
         TabIndex        =   119
         Top             =   690
         Width           =   7725
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Fecha Venc."
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
            Height          =   300
            Index           =   4
            Left            =   180
            TabIndex        =   120
            Top             =   1410
            Width           =   1830
         End
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   360
            Left            =   4350
            TabIndex        =   121
            Top             =   300
            Width           =   3150
            _ExtentX        =   5556
            _ExtentY        =   635
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BorderStyle     =   1
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   7
            Left            =   1800
            TabIndex        =   122
            Top             =   1050
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
            Index           =   8
            Left            =   1800
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   390
            Width           =   1965
            _Version        =   196608
            _ExtentX        =   3466
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
            Index           =   9
            Left            =   1800
            TabIndex        =   124
            Top             =   720
            Width           =   5760
            _Version        =   196608
            _ExtentX        =   10160
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
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   1
            Left            =   2160
            TabIndex        =   125
            Top             =   1380
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
         Begin VB.Label Label3 
            Caption         =   "Unidad x Embalaje"
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
            Index           =   19
            Left            =   180
            TabIndex        =   128
            Top             =   1125
            Width           =   1710
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
            Index           =   18
            Left            =   225
            TabIndex        =   127
            Top             =   435
            Width           =   855
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
            Index           =   8
            Left            =   225
            TabIndex        =   126
            Top             =   765
            Width           =   1635
         End
      End
      Begin VB.Frame Frame9 
         Height          =   4245
         Left            =   -74670
         TabIndex        =   117
         Top             =   2640
         Width           =   7965
         Begin FPSpread.vaSpread vaSpread5 
            Height          =   3885
            Left            =   120
            TabIndex        =   118
            Top             =   210
            Width           =   7785
            _Version        =   393216
            _ExtentX        =   13732
            _ExtentY        =   6853
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
            MaxRows         =   20
            SpreadDesigner  =   "M_Produc.frx":0D4A
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1275
         Left            =   1800
         TabIndex        =   116
         Top             =   510
         Width           =   7695
         Begin FPSpread.vaSpread vaSpread4 
            Height          =   1005
            Left            =   570
            TabIndex        =   34
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
            MaxCols         =   4
            MaxRows         =   3
            ScrollBars      =   2
            SpreadDesigner  =   "M_Produc.frx":11B8
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
         Height          =   2550
         Left            =   -73575
         TabIndex        =   115
         Top             =   7005
         Width           =   8160
         Begin FPSpread.vaSpread vaSpread3 
            Height          =   2205
            Left            =   390
            TabIndex        =   33
            Top             =   255
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
            SpreadDesigner  =   "M_Produc.frx":15D3
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
         Left            =   1800
         TabIndex        =   114
         Top             =   5160
         Width           =   7680
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   2205
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   7380
            _Version        =   393216
            _ExtentX        =   13018
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
            SpreadDesigner  =   "M_Produc.frx":1A83
            ScrollBarTrack  =   3
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   0
         Left            =   -72360
         TabIndex        =   108
         Top             =   480
         Width           =   6015
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   315
            ItemData        =   "M_Produc.frx":33F9
            Left            =   1680
            List            =   "M_Produc.frx":3406
            Style           =   2  'Dropdown List
            TabIndex        =   109
            Top             =   240
            Width           =   2500
         End
         Begin EditLib.fpText fpText 
            Height          =   315
            Left            =   1680
            TabIndex        =   110
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
            TabIndex        =   113
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
            TabIndex        =   112
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
            TabIndex        =   111
            Top             =   645
            Width           =   585
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7350
         Left            =   -73500
         TabIndex        =   106
         Top             =   1695
         Width           =   8025
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   6945
            Left            =   195
            TabIndex        =   107
            Top             =   225
            Width           =   7635
            _Version        =   393216
            _ExtentX        =   13467
            _ExtentY        =   12250
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
            MaxCols         =   6
            OperationMode   =   3
            SelectBlockOptions=   0
            SpreadDesigner  =   "M_Produc.frx":3423
            VisibleCols     =   1
            VisibleRows     =   15
            ScrollBarTrack  =   3
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3375
         Left            =   1800
         TabIndex        =   87
         Top             =   1770
         Width           =   7680
         Begin EditLib.fpDoubleSingle fpHCarbono 
            Height          =   315
            Left            =   6105
            TabIndex        =   48
            Top             =   2520
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   ""
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
         Begin VB.Frame Frame6 
            Height          =   30
            Left            =   0
            TabIndex        =   88
            Top             =   2920
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
            TabIndex        =   45
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
            TabIndex        =   46
            Top             =   2220
            Width           =   2220
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Index           =   2
            Left            =   2115
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   2500
            Width           =   1800
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   1
            Left            =   2115
            TabIndex        =   40
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
            TabIndex        =   41
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
            TabIndex        =   42
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
            TabIndex        =   44
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
            DecimalPlaces   =   8
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
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   210
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
            TabIndex        =   36
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   7
            Left            =   2115
            TabIndex        =   37
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   4
            Left            =   2115
            TabIndex        =   39
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
            Left            =   4155
            TabIndex        =   89
            Top             =   150
            Width           =   3150
            _ExtentX        =   5556
            _ExtentY        =   635
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BorderStyle     =   1
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Huella Carbono"
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
            Index           =   28
            Left            =   4095
            TabIndex        =   149
            Top             =   2595
            Width           =   1320
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
            TabIndex        =   102
            Top             =   3075
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
            TabIndex        =   101
            Top             =   3075
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
            TabIndex        =   47
            Top             =   3030
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
            TabIndex        =   49
            Top             =   3030
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
            TabIndex        =   100
            Top             =   1250
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
            TabIndex        =   99
            Top             =   1200
            Width           =   3525
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   3300
            Picture         =   "M_Produc.frx":4F23
            Top             =   1095
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
            TabIndex        =   98
            Top             =   600
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
            TabIndex        =   97
            Top             =   935
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
            TabIndex        =   96
            Top             =   280
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
            TabIndex        =   95
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
            TabIndex        =   94
            Top             =   1920
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
            TabIndex        =   93
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
            TabIndex        =   92
            Top             =   2240
            Width           =   1710
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Ingrediente"
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
            Index           =   23
            Left            =   315
            TabIndex        =   91
            Top             =   2600
            Width           =   1695
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   14
            Left            =   2160
            TabIndex        =   90
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   3810
            TabIndex        =   103
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
            TabIndex        =   104
            Top             =   3030
            Width           =   1200
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   9
            Left            =   6150
            TabIndex        =   105
            Top             =   3030
            Width           =   1200
         End
      End
      Begin VB.Frame Frame5 
         Height          =   6555
         Left            =   -73575
         TabIndex        =   7
         Top             =   375
         Width           =   8160
         Begin VB.Frame Frame12 
            Caption         =   "Tipo Producto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   141
            Top             =   4440
            Width           =   3855
            Begin VB.OptionButton Option1 
               Caption         =   "Servicio"
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
               Left            =   2400
               TabIndex        =   27
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Insumo"
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
               Left            =   600
               TabIndex        =   26
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Index           =   4
            Left            =   5715
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   4020
            Width           =   2295
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
            Height          =   300
            Index           =   3
            Left            =   165
            TabIndex        =   20
            Top             =   3555
            Width           =   2040
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Fecha Vencimiento"
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
            Height          =   300
            Index           =   2
            Left            =   3675
            TabIndex        =   21
            Top             =   3555
            Width           =   2190
         End
         Begin VB.Frame Frame4 
            Height          =   75
            Left            =   30
            TabIndex        =   9
            Top             =   5745
            Width           =   8100
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Index           =   0
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   4020
            Width           =   2295
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Index           =   3
            Left            =   6195
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   1560
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Index           =   1
            Left            =   6195
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1560
            Width           =   1800
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Afecto a Cuota Horfruticola"
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
            Height          =   420
            Index           =   5
            Left            =   165
            TabIndex        =   25
            Top             =   4500
            Width           =   2040
         End
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   18
            Top             =   2850
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
            Left            =   1920
            TabIndex        =   10
            Top             =   510
            Width           =   6180
            _Version        =   196608
            _ExtentX        =   10901
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   1
            Left            =   1920
            TabIndex        =   12
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
            Left            =   1920
            TabIndex        =   17
            Top             =   180
            Width           =   2160
            _Version        =   196608
            _ExtentX        =   3810
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
            Left            =   1920
            TabIndex        =   11
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
            Left            =   1920
            TabIndex        =   16
            Top             =   2520
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
            Left            =   3225
            TabIndex        =   38
            Top             =   2820
            Visible         =   0   'False
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
            MinValue        =   "1"
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
            Left            =   1920
            TabIndex        =   14
            Top             =   1830
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
            Left            =   1920
            TabIndex        =   13
            Top             =   1500
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
            Left            =   6585
            TabIndex        =   22
            Top             =   3540
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   10
            Left            =   5925
            TabIndex        =   51
            Top             =   180
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
         Begin EditLib.fpDoubleSingle fpDouble1 
            Height          =   315
            Index           =   8
            Left            =   1920
            TabIndex        =   19
            Top             =   3180
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
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9999999999"
            MinValue        =   "1"
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   5
            Left            =   1920
            TabIndex        =   28
            Top             =   5040
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
            Index           =   6
            Left            =   1920
            TabIndex        =   29
            Top             =   5370
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
            Index           =   7
            Left            =   1920
            TabIndex        =   15
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
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   3195
            TabIndex        =   147
            Top             =   2160
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   16
            Left            =   3240
            TabIndex        =   148
            Top             =   2220
            Width           =   4830
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   7
            Left            =   2640
            Picture         =   "M_Produc.frx":522D
            Top             =   2040
            Width           =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Un. F. Conv. Ing."
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
            Left            =   165
            TabIndex        =   146
            Top             =   2240
            Width           =   1485
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   15
            Left            =   5760
            TabIndex        =   140
            Top             =   4035
            Width           =   2310
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo O.C."
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
            Index           =   26
            Left            =   4680
            TabIndex        =   139
            Top             =   4125
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fact. Conv. Ing."
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
            TabIndex        =   77
            Top             =   1905
            Width           =   1395
         End
         Begin VB.Label fpAyDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   3795
            TabIndex        =   31
            Top             =   5925
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
            Left            =   2820
            TabIndex        =   76
            Top             =   5850
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
            Left            =   5550
            TabIndex        =   75
            Top             =   5850
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
            TabIndex        =   74
            Top             =   5850
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
            TabIndex        =   30
            Top             =   5925
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
            Left            =   6765
            TabIndex        =   32
            Top             =   5925
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
            TabIndex        =   73
            Top             =   2925
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
            TabIndex        =   72
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
            TabIndex        =   71
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
            TabIndex        =   70
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
            TabIndex        =   69
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
            TabIndex        =   68
            Top             =   2610
            Width           =   1830
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3195
            TabIndex        =   67
            Top             =   840
            Width           =   4830
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   3195
            TabIndex        =   66
            Top             =   1185
            Width           =   4830
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   3195
            TabIndex        =   65
            Top             =   2520
            Width           =   4830
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
            TabIndex        =   64
            Top             =   3255
            Width           =   1830
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   3435
            TabIndex        =   63
            Top             =   3180
            Width           =   4590
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   2640
            Picture         =   "M_Produc.frx":5537
            Top             =   735
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   2640
            Picture         =   "M_Produc.frx":5841
            Top             =   1080
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   2640
            Picture         =   "M_Produc.frx":5B4B
            Top             =   2430
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   3000
            Picture         =   "M_Produc.frx":5E55
            Top             =   3105
            Width           =   480
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
            TabIndex        =   62
            Top             =   1575
            Width           =   1830
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
            Index           =   20
            Left            =   165
            TabIndex        =   61
            Top             =   4125
            Width           =   1950
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   10
            Left            =   2325
            TabIndex        =   60
            Top             =   4035
            Width           =   2310
         End
         Begin VB.Label Label3 
            Caption         =   "Cód. Compras"
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
            Index           =   21
            Left            =   4560
            TabIndex        =   59
            Top             =   240
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   11
            Left            =   6240
            TabIndex        =   58
            Top             =   1580
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Producto"
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
            Index           =   22
            Left            =   4560
            TabIndex        =   57
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Ret. en la Fuente"
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
            Index           =   24
            Left            =   165
            TabIndex        =   56
            Top             =   5100
            Width           =   1590
         End
         Begin VB.Label Label3 
            Caption         =   "Retención ICA"
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
            Index           =   25
            Left            =   165
            TabIndex        =   55
            Top             =   5415
            Width           =   1590
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   5
            Left            =   2640
            Picture         =   "M_Produc.frx":615F
            Top             =   4920
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   6
            Left            =   2640
            Picture         =   "M_Produc.frx":6469
            Top             =   5280
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   3195
            TabIndex        =   53
            Top             =   5040
            Width           =   4830
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   3195
            TabIndex        =   52
            Top             =   5370
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3240
            TabIndex        =   78
            Top             =   900
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   3240
            TabIndex        =   79
            Top             =   1230
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   3240
            TabIndex        =   80
            Top             =   2565
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   3480
            TabIndex        =   81
            Top             =   3225
            Width           =   4590
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   12
            Left            =   3240
            TabIndex        =   85
            Top             =   5085
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   13
            Left            =   3240
            TabIndex        =   86
            Top             =   5415
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   3840
            TabIndex        =   82
            Top             =   5970
            Width           =   1200
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   6810
            TabIndex        =   84
            Top             =   5970
            Width           =   1200
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   1170
            TabIndex        =   83
            Top             =   5970
            Width           =   1200
         End
      End
      Begin VB.Frame Frame11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Left            =   -74280
         TabIndex        =   3
         Top             =   600
         Width           =   9855
         Begin FPSpread.vaSpread vaSpread7 
            Height          =   375
            Left            =   1080
            TabIndex        =   4
            Top             =   5880
            Visible         =   0   'False
            Width           =   615
            _Version        =   393216
            _ExtentX        =   1085
            _ExtentY        =   661
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
            MaxCols         =   1
            MaxRows         =   1
            SpreadDesigner  =   "M_Produc.frx":6773
         End
         Begin FPSpread.vaSpread vaSpread6 
            Height          =   5895
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   9585
            _Version        =   393216
            _ExtentX        =   16907
            _ExtentY        =   10398
            _StockProps     =   64
            ButtonDrawMode  =   2
            DisplayRowHeaders=   0   'False
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
            MaxCols         =   7
            MaxRows         =   5
            SpreadDesigner  =   "M_Produc.frx":6C3C
         End
         Begin MSComctlLib.Toolbar Toolbar4 
            Height          =   360
            Left            =   6660
            TabIndex        =   6
            Top             =   240
            Width           =   2790
            _ExtentX        =   4921
            _ExtentY        =   635
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BorderStyle     =   1
         End
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   -69510
         TabIndex        =   129
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
         TabIndex        =   130
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
         TabIndex        =   131
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
         Index           =   1
         Left            =   -72630
         TabIndex        =   138
         Top             =   450
         Width           =   1800
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
         Left            =   2805
         TabIndex        =   137
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
         TabIndex        =   136
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
         TabIndex        =   135
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
         TabIndex        =   134
         Top             =   6210
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H80000018&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   -69555
         Top             =   9180
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Productos Vigentes"
         Height          =   195
         Index           =   0
         Left            =   -69195
         TabIndex        =   133
         Top             =   9150
         Width           =   1380
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00D9D9FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   -71700
         Top             =   9180
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Productos No Vigentes"
         Height          =   195
         Index           =   1
         Left            =   -71340
         TabIndex        =   132
         Top             =   9150
         Width           =   1635
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
      TabIndex        =   1
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
Dim RS4 As New ADODB.Recordset
Dim ibusca As Long, codtip As Long
Dim i As Long, Est As Boolean, EstDet As Boolean
Dim modo As String, modo2 As String, modo3 As String
Dim codigo As String, codfam As Long
Dim operr As Integer
Dim ComboValOri As String
Dim ComboValOri1 As String

Private Sub Check1_Click(Index As Integer)

If Est Then Exit Sub

Select Case Index

Case 0, 1
    
    If modo2 = "" Then modo2 = "M"
    Gl_Ac_Botones Me, 8, 0, modo2

Case 2, 3, 5
    
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False
    If Index = 5 Then Exit Sub
    If Check1(2).Value = 1 Then
       
       fpDateTime1(0).Enabled = True
       fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
    
    Else
       
       fpDateTime1(0).Enabled = False
       fpDateTime1(0).text = "  /  /    "
    
    End If

Case 4
    
    If Check1(4).Value = 1 Then fpDateTime1(1).Enabled = True: fpDateTime1(1).text = Format(Date, "dd/mm/yyyy") Else fpDateTime1(1).Enabled = False: fpDateTime1(1).text = "  /  /    "
    If modo3 = "" Then modo3 = "M"
    Gl_Ac_Botones Me, 9, 0, modo3

End Select

End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If Index = 1 And vaSpread2.MaxRows > 0 Then vaSpread2.SetActiveCell 3, 1

End Sub

Private Sub Combo1_Click()

If Est Then Exit Sub

If Combo1.ListIndex = 2 Then
    
    vg_left = Frame1(0).Left + Combo1.Left + 1920
    B_ArbEst.MoverDatosTvwDir "a_tipopro", "tip_", "Familia del Producto"
    B_ArbEst.Show 1
    Me.Refresh
    If Val(vg_codigo) = 0 Then Exit Sub
    codtip = Val(vg_codigo)
    fpText.text = vg_nombre
    fpText.Enabled = False

Else
    
    fpText.Enabled = True
    fpText.text = ""

End If

End Sub

Private Sub Combo2_Click(Index As Integer)

If Est Then Exit Sub
Select Case Index

Case 0, 1, 4
    
    If Est Then Exit Sub
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False

Case 2
    
    If modo2 = "" Then modo2 = "M"
    Gl_Ac_Botones Me, 8, 0, modo2
    Est = False
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False

End Select

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.Height = 10695
Me.Width = 11475
Me.HelpContextID = vg_OpcM
fg_centra Me
MsgTitulo = "Maestro de Productos"
SSTab1.TabVisible(3) = False
SSTab1.TabEnabled(3) = False
SSTab1.Tab = 0
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Me.HelpContextID = "1011000"
Gl_Mo_Botones Me, 8
Gl_Ac_Botones Me, 8, 1, modo2
Me.HelpContextID = vg_OpcM
Gl_Mo_Botones Me, 9
Gl_Ac_Botones Me, 9, 1, modo3
Gl_Mo_Botones Me, 11
Gl_Ac_Botones Me, 11, 1, ""
Gl_Mo_Botones Me, 17
Gl_Ac_Botones Me, 15, 1, ""
Est = True
EstDet = True
fpDateTime1(0).text = "  /  /    "
fpDateTime1(1).text = "  /  /    "
Combo1.ListIndex = 1

'-------> Mover tipo productos
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_s_tiposervicio 5, 0,''")
Combo2(0).Clear
Combo2(0).AddItem "Ambos" & Space(150) & "(0)"

Do While Not RS1.EOF
   
   Combo2(0).AddItem Trim(RS1!tis_nombre) & Space(150) & "(" & RS1!tis_codigo & ")"
   RS1.MoveNext

Loop
RS1.Close
Set RS1 = Nothing
Combo2(0).ListIndex = -1

'-------> Mover tipo productos
Combo2(1).Clear
Combo2(1).AddItem "Real" & Space(150) & "(1)"
Combo2(1).AddItem "Propuesta" & Space(150) & "(2)"

'-------> Mover tipo ingrediente
Combo2(2).Clear
Combo2(2).AddItem "Real" & Space(150) & "(1)"
Combo2(2).AddItem "Propuesta" & Space(150) & "(2)"

'-------> Mover tipo Orden de Compras
Combo2(4).Clear
Combo2(4).AddItem "SAC" & Space(150) & "(1)"
Combo2(4).AddItem "LOCAL" & Space(150) & "(2)"

MoverDatosGrilla

'-------> Mover nutrientes
Dim IndApro As Long

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_s_nutriente 1, 0, ''")
vaSpread2.MaxRows = 0
vaSpread2.MaxRows = RS1.RecordCount
IndApo = 1

Do While Not RS1.EOF
    
'    vaSpread2.MaxRows = vaSpread2.MaxRows + 1
    vaSpread2.Row = IndApo 'vaSpread2.MaxRows
    vaSpread2.Col = 1
    vaSpread2.Value = RS1!nut_codigo
    
    vaSpread2.Col = 2
    vaSpread2.Value = Trim(RS1!nut_nombre)
    
    vaSpread2.Col = 3
    vaSpread2.Value = 0
    
    vaSpread2.Col = 4
    vaSpread2.Value = Trim(RS1!nut_nomuni)
    
    IndApo = IndApo + 1
    RS1.MoveNext

Loop
RS1.Close
Set RS1 = Nothing

'-------> Mover impuesto
Dim IndImp As Long

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_s_impuesto 1, 0, ''")
vaSpread3.MaxRows = 0
vaSpread3.MaxRows = RS1.RecordCount
IndImp = 1

Do While Not RS1.EOF
    
    vaSpread3.Col = 3
    vaSpread3.Enabled = False
    
    'vaSpread3.MaxRows = vaSpread3.MaxRows + 1
    vaSpread3.Row = IndImp 'vaSpread3.MaxRows
    vaSpread3.Col = 1
    vaSpread3.Value = RS1!imp_codigo
    
    vaSpread3.Col = 2
    vaSpread3.Value = Trim(RS1!imp_nombre)
    
    vaSpread3.Col = 3
    vaSpread3.Value = 1
    vaSpread3.Enabled = True
    
    IndImp = IndImp + 1
    
    RS1.MoveNext

Loop
RS1.Close
Set RS1 = Nothing
modo = ""
MoverDatos
Est = False

Exit Sub
Man_Error:

End Sub

Private Sub Form_Terminate()

vg_opimp = 0

End Sub

Private Sub fpDateTime1_Change(Index As Integer)

If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
If Est Then Exit Sub
Select Case Index

Case 0
    
    If modo = "" Then modo = "M"
    If vg_Indppr = 1 Or vg_Indppr = 3 Then
      Gl_Ac_Botones Me, 1, 0, modo
      SSTab1.TabEnabled(0) = False
      SSTab1.TabEnabled(4) = False
      SSTab1.TabEnabled(5) = False
    ElseIf vg_Indppr = 2 And ComboValOri = vg_Indppr Then
      Gl_Ac_Botones Me, 1, 0, modo
      SSTab1.TabEnabled(0) = False
      SSTab1.TabEnabled(4) = False
      SSTab1.TabEnabled(5) = False
    End If

Case 1
    If modo3 = "" Then modo3 = "M"
    Gl_Ac_Botones Me, 9, 0, modo3

End Select

End Sub

Private Sub fpDouble1_Change(Index As Integer)

If Est Then Exit Sub
If Index >= 1 And Index <= 3 And Val(fpDouble1(Index).Value) = 0 Then fpDouble1(Index).Value = 100
If Index = 0 Or Index = 5 Or Index = 6 Then
    
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False

ElseIf Index = 7 Then
    
    If modo3 = "" Then modo3 = "M"
    Gl_Ac_Botones Me, 9, 0, modo3
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False

Else
    
    If modo2 = "" Then modo2 = "M"
    Gl_Ac_Botones Me, 8, 0, modo2
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False

End If

End Sub

Private Sub fpDouble1_KeyPress(Index As Integer, KeyAscii As Integer)

If Index = 0 Or Index = 2 Or Index = 5 Or Index = 6 Or Index = 7 Then fpDouble1(Index).MaxValue = 9000000# Else fpDouble1(Index).MaxValue = 100
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub fpDouble1_LostFocus(Index As Integer)

Select Case Index

Case Is <> 8
    
    If LimpiaDato(Trim(fpText1(5).text)) <> "" And Index >= 1 And Index <= 4 And Val(fpDouble1(Index).Value) = 0 Then _
    fpDouble1(Index).Value = 100

Case 8
    
    codi = fpDouble1(8).Value
    If Trim(codi) = "" Then Exit Sub
    Bd = "a_ctacontable"
    Ul = "cta"
    
    Set RS1 = Nothing
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "SELECT " & Ul & "_nombre FROM " & Bd & " WHERE " & Ul & "_codigo=" & IIf(Ul = "cta", "'" & codi & "'", codi), vg_db, adOpenStatic
    If Not RS1.EOF Then
        fpayuda(3).Caption = IIf(IsNull(RS1(0)) Or Trim(RS1(0) = ""), "", RS1(0))
        codi = 0
    Else
        MsgBox "No existe cuenta corriente"
        fpayuda(3).Caption = ""
        fpDouble1(8).Value = ""
        codi = 0
        On Error Resume Next: fpDouble1(8).SetFocus
    End If
    RS1.Close: Set RS1 = Nothing

End Select

End Sub

Private Sub fpHCarbono_Change()
    
If Est Then Exit Sub

If modo2 = "" Then modo2 = "M"
Gl_Ac_Botones Me, 8, 0, modo2
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(5) = False
Exit Sub

End Sub

Private Sub fpHCarbono_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

If Est Then Exit Sub

fpayuda(Index).Caption = ""

If Index = 4 Then
    
    If modo2 = "" Then modo2 = "M"
    Gl_Ac_Botones Me, 8, 0, modo2
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False
    Exit Sub

End If

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(5) = False

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

If Index = 7 Then
    
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS1 = vg_db.Execute("sgpadm_Sel_UnidadConversionIngredienteCodigo '%" & UCase(LimpiaDato(fpLongInteger1(7).Value)) & "%'")
   
   If Not RS1.EOF Then
    
      fpayuda(7).Caption = IIf(IsNull(RS1(0)) Or Trim(RS1(0) = ""), "", RS1(1))
   
   Else
    
      MsgBox "No existe codigo unidad factor ingrediente..."
      fpayuda(7).Caption = ""
      fpLongInteger1(7).Value = ""
 
      On Error Resume Next: fpLongInteger1(Index).SetFocus

   End If
   RS1.Close: Set RS1 = Nothing
   
   Exit Sub

End If

codi = fpLongInteger1(Index).Value
Bd = IIf(Index = 0, "a_tipopro", IIf(Index = 1, "a_unidad", IIf(Index = 2, "a_embalaje", IIf(Index = 3, "a_ctacontable", IIf(Index = 5, "b_retencionfuente", IIf(Index = 6, "b_retencionica", "a_unidadmed"))))))
Ul = IIf(Bd = "a_unidadmed", "unm", IIf(Bd = "b_retencionfuente", "ref", IIf(Bd = "b_retencionica", "rei", Mid(Bd, 3, 3))))

Set RS1 = Nothing

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT " & Ul & "_nombre FROM " & Bd & " WHERE " & Ul & "_codigo=" & IIf(Ul = "cta", "'" & codi & "'", codi), vg_db, adOpenStatic

If Not RS1.EOF Then
    
    fpayuda(Index).Caption = IIf(IsNull(RS1(0)) Or Trim(RS1(0) = ""), "", RS1(0))
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
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpText1_Change(Index As Integer)

If Est Then Exit Sub
If Index = 6 Or Index = 7 Then
    If modo2 = "" Then modo2 = "M"
    Me.HelpContextID = "1011000"
    Gl_Ac_Botones Me, 8, 0, modo2
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False
    Me.HelpContextID = vg_OpcM
ElseIf Index = 8 Or Index = 9 Then
    If modo3 = "" Then modo3 = "M"
    Gl_Ac_Botones Me, 9, 0, modo3
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False
End If
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(5) = False

End Sub

Private Sub Limpia(Index As Integer)

Est = True

Select Case Index

Case 1
    
    fpText1(0).text = ""
    fpText1(1).text = ""
    fpText1(2).text = ""
    fpText1(4).text = ""
    fpLongInteger1(0).text = ""
    fpLongInteger1(1).text = ""
    fpLongInteger1(2).text = ""
    fpLongInteger1(5).text = ""
    fpLongInteger1(6).text = ""
    fpLongInteger1(7).text = ""
    fpDouble1(8).text = ""
    fpayuda(0).Caption = ""
    fpayuda(1).Caption = ""
    fpayuda(2).Caption = ""
    fpayuda(3).Caption = ""
    fpayuda(5).Caption = ""
    fpayuda(6).Caption = ""
    fpayuda(7).Caption = ""
    fpDouble1(0).Value = 0
    fpDouble1(5).Value = 0
    fpDouble1(6).Value = 0
    fpDouble1(8).Value = ""
    fpAyDouble1(1).Caption = ""
    fpAyDouble1(2).Caption = ""
    fpText1(5).text = ""
    fpText1(10).text = ""
    fpAyDate(7).Caption = ""
    Check1(5).Value = 0
    Option1(0).Value = True
    Option1(1).Value = False
    vaSpread3.Col = 3
    
    For i = 1 To vaSpread3.MaxRows
        
        vaSpread3.Row = i
        vaSpread3.text = "0"
    
    Next i
    
    fpDateTime1(0).Enabled = False: fpDateTime1(0).text = "  /  /    "
    Check1(2).Value = 0: Check1(3).Value = 0
    Combo2(4).ListIndex = -1
    BloquearOpSistema

Case 2
    
    fpText1(5).Enabled = True
    fpText1(5).ControlType = ControlTypeNormal
    fpText1(5).text = ""
    fpText1(6).text = ""
    fpText1(7).text = ""
    fpLongInteger1(4).text = ""
    fpayuda(4).Caption = ""
    fpDouble1(1).Value = 100
    fpDouble1(2).Value = 100
    fpDouble1(3).Value = 100
    fpDouble1(4).Value = 100
    Check1(0).Value = 0
    Check1(1).Value = 0
    fpAyDouble1(0).Caption = ""
    fpAyDate(0).Caption = ""
    vaSpread2.Col = 3
    fpHCarbono.Value = 0
    
    For i = 1 To vaSpread2.MaxRows
        
        vaSpread2.Row = i
        vaSpread2.text = Format(0, fg_Pict(9, 4))
    
    Next i

Case 3
    
    fpText1(8).Enabled = True
    fpText1(8).ControlType = ControlTypeNormal
    fpText1(8).text = ""
    fpText1(9).text = ""
    fpDouble1(7).Value = 0
    fpDateTime1(1).Enabled = False: fpDateTime1(1).text = "  /  /    "
    Check1(4).Value = 0

Case 4
    
    Frame11.Caption = ""
    vaSpread6.MaxRows = 0

Case 5
    
    Frame14.Caption = ""
    vaSpread10.MaxRows = 0

End Select

Est = False

End Sub

Private Sub fpText1_GotFocus(Index As Integer)

If Index = 5 And modo2 <> "A" Then Limpia 2
If Index = 8 And modo3 <> "A" Then Limpia 3

End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub fpText_Change()

If Est Then Exit Sub

On Error GoTo Man_Error

codTippro = 0
nomTippro = ""

'If LimpiaDato(Trim(fpText.Text)) & Chr(KeyAscii) = "" Then Exit Sub

If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    
    Set RS2 = vg_db.Execute("sgpadm_Sel_productos 4, '', '%" & UCase(LimpiaDato(Trim(fpText.text))) & "%', '" & vg_NUsr & "'")
    If RS2.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS2!nReg

ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    
    Set RS2 = vg_db.Execute("sgpadm_Sel_productos 6, '', '%" & UCase(LimpiaDato(Trim(fpText.text))) & "%', '" & vg_NUsr & "'")
    If RS2.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS2!nReg

ElseIf Combo1.ItemData(Combo1.ListIndex) = 2 Then
    
    Set RS2 = vg_db.Execute("sgpadm_Sel_productos 7, " & codtip & ", '', '" & vg_NUsr & "'")
    If RS2.EOF Or RS2!nReg = 0 Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS2!nReg
    RS2.Close: Set RS2 = Nothing
    Set RS2 = vg_db.Execute("sgpadm_Sel_productos 8, " & codtip & ", '', '" & vg_NUsr & "'")

End If

vaSpread1.Visible = False
i = 1
If Not RS2.EOF Then
    
    Do While Not RS2.EOF
        
        vaSpread1.Row = i: i = i + 1
        vaSpread1.Col = -1
        If Val(Format(Date, "yyyymmdd")) > RS2!pro_fecven And RS2!pro_fecven > 0 Then vaSpread1.BackColor = Shape1(1).FillColor Else vaSpread1.BackColor = Shape1(0).FillColor
        
        vaSpread1.Col = 1
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.Value = RS2!pro_codigo
        
        vaSpread1.Col = 2
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.Value = Trim(RS2!pro_nombre)
        
        vaSpread1.Col = 3
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.Value = IIf(RS2!pro_maepro = 0, "Ambos", Trim(RS2!tis_nombre))
        
        vaSpread1.Col = 4
        vaSpread1.TypeHAlign = TypeHAlignRight
        vaSpread1.Value = IIf(RS2!pro_propon < 1, "", Format(RS2!pro_propon, fg_Pict(9, vg_DPr)))
        
        vaSpread1.Col = 5
        vaSpread1.text = RS2!pro_codtip
        
        vaSpread1.Col = 6
        vaSpread1.text = IIf(IsNull(RS2!pro_indppr) Or Trim(RS2!pro_indppr) = "", "", IIf(RS2!pro_indppr = "1", "Real", "Propuesta"))
                
        RS2.MoveNext
        
    Loop
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Combo1.ItemData(Combo1.ListIndex) = 0, 1, 2)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Combo1.ItemData(Combo1.ListIndex) = 0, 1, 2)
    vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
    vaSpread1.SetActiveCell 1, 1
    vaSpread1.Col = 6
    If vg_Indppr = 2 And vaSpread1.text = "Real" Then
       Gl_Ac_Botones Me, 1, 11, modo
    Else
        Gl_Ac_Botones Me, 1, 1, modo
    End If
Else
    Est = True: modo = "": MoverDatos: Est = False
End If
RS2.Close: Set RS2 = Nothing

vaSpread1.Visible = True
If Trim(fpText.text) = "" Then Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro" Else Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
If vaSpread1.MaxRows = 0 Then
   
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
   SSTab1.TabEnabled(4) = False
   SSTab1.TabEnabled(5) = False

ElseIf vaSpread1.MaxRows > 0 Then
   
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
   SSTab1.TabEnabled(4) = True
   SSTab1.TabEnabled(5) = True

End If
EstDet = False
Est = False

Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
Resume Next
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpText1_LostFocus(Index As Integer)

If Est Then Exit Sub
Dim codigo As String, i As Long

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'est = True
Select Case Index

Case 5
    
    codigo = LimpiaDato(Trim(fpText1(5).text))
    If codigo = "" And modo2 <> "A" Then MsgBox "Debe ingresar ingrediente...", vbExclamation, MsgTitulo: Exit Sub
    For i = 1 To 4
        If codigo <> "" And Val(fpDouble1(i).Value) = 0 Then _
           fpDouble1(i).Value = 100
    Next i
    Set RS1 = vg_db.Execute("sgpadm_s_ingrediente_V02 3, '" & codigo & "', ''")
    If RS1.EOF Then
       RS1.Close: Set RS1 = Nothing
       fpText1(5).text = "": MsgBox "Ingrediente no existe...", vbExclamation, MsgTitulo: SendKeys "+{Tab}"
    Else
       Est = False
       codigo = RS1!ing_codigo
       RS1.Close: Set RS1 = Nothing
       modo2 = "": MoverDatos3 codigo
    End If

Case 8
    
    codigo = LimpiaDato(Trim(fpText1(8).text))
    If codigo = "" Then: Exit Sub
    If codigo = "" And modo3 <> "A" Then MsgBox "Debe ingresar producto compra...", vbExclamation, MsgTitulo: Exit Sub
    Set RS1 = vg_db.Execute("SELECT pco_codigo FROM b_productocompra WHERE pco_codigo='" & codigo & "'")
    If Not RS1.EOF Then
       modo3 = "": RS1.Close: Set RS1 = Nothing
       fpText1(8).text = "": MsgBox "Código compras existe...", vbExclamation, MsgTitulo: SendKeys "+{Tab}"
    Else
       RS1.Close: Set RS1 = Nothing
       Est = False: modo3 = "A"
       If vg_Indppr = 1 Or vg_Indppr = 3 Then Gl_Ac_Botones Me, 9, 0, modo3
       If vg_Indppr = 2 Or vg_Indppr = ComboValOri Then Gl_Ac_Botones Me, 9, 0, modo3
    End If

End Select

End Sub

Private Sub MoverDatos2(codpro As String)

Dim RS4    As New ADODB.Recordset
Dim RS3    As New ADODB.Recordset
Dim IndIng As Long

If RS4.State = 1 Then RS4.Close
RS4.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS4 = vg_db.Execute("sgpadm_s_ingrediente_V02 1, '" & codpro & "', ''")
vaSpread4.MaxRows = 0
vaSpread4.MaxRows = RS4.RecordCount
IndIng = 1

If Not RS4.EOF Then
   
   Do While Not RS4.EOF
      
'      vaSpread4.MaxRows = vaSpread4.MaxRows + 1
      vaSpread4.Row = IndIng 'vaSpread4.MaxRows
      
      vaSpread4.Col = 1
      vaSpread4.text = RS4!ing_codigo
      
      vaSpread4.Col = 2
      vaSpread4.text = RS4!ing_nombre
      
      vaSpread4.Col = 3
      vaSpread4.text = IIf(IsNull(RS4!pri_propre) Or RS4!pri_propre = 0, 0, RS4!pri_propre)
      
      vaSpread4.Col = 4
      vaSpread4.text = IIf(IsNull(RS4!ing_indppr) Or RS4!ing_indppr = "", "", IIf(RS4!ing_indppr = "1", "Real", "Propuesta"))
      
      IndIng = IndIng + 1
      RS4.MoveNext
   
   Loop

End If

RS4.Close: Set RS4 = Nothing

If vaSpread4.MaxRows < 1 Then Exit Sub
vaSpread4.Row = 1: vaSpread4.Col = 1

If RS4.State = 1 Then RS4.Close
RS4.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS4 = vg_db.Execute("sgpadm_s_ingrediente_V02 2, '" & vaSpread4.text & "', ''")
If Not RS4.EOF Then
    
    Est = True
    
    fpText1(5).ControlType = ControlTypeStatic
    fpText1(5).text = RS4!ing_codigo
    fpText1(6).text = RS4!ing_nombre
    fpText1(7).text = RS4!ing_nomfan
    fpLongInteger1(4).text = RS4!ing_unimed
    fpayuda(4).Caption = RS4!unm_nombre
    fpDouble1(1).Value = RS4!ing_pctapr
    fpDouble1(2).Value = RS4!ing_pctcoc
    fpDouble1(3).Value = RS4!ing_pctnut
    fpDouble1(4).Value = RS4!ing_facnut
    Check1(0).Value = RS4!ing_indpav
    Check1(1).Value = RS4!ing_indgrv
    If IsNull(RS4!ing_indppr) Or Trim(RS4!ing_indppr) = "" Then
      
      Combo2(2).ListIndex = -1
    
    Else
      
      Combo2(2).ListIndex = fg_buscacbo(Combo2, 2, 1, fg_pone_cero(Str(RS4!ing_indppr), 1))
      If vg_Indppr = 2 Then Gl_Ac_BotonesRealPropuesta Me, 8, 1, modo, vg_Indppr, RS4!ing_indppr
      If vg_Indppr = 2 And vg_Indppr <> RS4!ing_indppr Then ConfiControlesProducto 2, False Else ConfiControlesProducto 2, True
      Combo2(2).Enabled = IIf(vg_Indppr = "1" Or vg_Indppr = "2", False, True)
    
    End If
    fpAyDouble1(0).Caption = IIf(RS4!ing_precos = 0, "", Format(RS4!ing_precos, fg_Pict(9, 4)))
    fpAyDate(0).Caption = IIf(RS4!ing_feccos = 0, "", fg_Ctod1(RS4!ing_feccos))
    
    fpHCarbono.text = IIf(RS4!Huella_Carbono = 0, 0, Format(RS4!Huella_Carbono, fg_Pict(6, 2)))
    
    If RS3.State = 1 Then RS3.Close
    RS3.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS3 = vg_db.Execute("sgpadm_s_productonutriente 1, '" & vaSpread4.text & "', 0")
    
    If Not RS3.EOF Then
       
       Do While Not RS3.EOF
          
          vaSpread2.Row = vaSpread2.SearchCol(1, -1, vaSpread2.MaxRows, Trim(CStr(RS3!pnu_codapo)), SearchFlagsEqual)
          
          vaSpread2.Col = 3
          vaSpread2.Value = RS3!pnu_canapo
          
          RS3.MoveNext
       
       Loop
    
    End If
    RS3.Close: Set RS3 = Nothing
    Est = False

Else
    
    fpText1(5).ControlType = ControlTypeNormal

End If
RS4.Close: Set RS4 = Nothing

End Sub

Private Sub MoverDatos3(CodIng As String)

Dim RS1    As New ADODB.Recordset
Dim RS3    As New ADODB.Recordset

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_s_ingrediente_V02 4, '" & CodIng & "', ''")

If Not RS1.EOF Then
    
    Est = True
    
    If vaSpread4.MaxRows > 0 Then
       
       For i = 1 To vaSpread4.MaxRows
           
           vaSpread4.Row = i
           vaSpread4.Col = 1
           
           If Trim(RS1!ing_codigo) = Trim(vaSpread4.text) Then
              
              Exit For
           
           ElseIf Trim(RS1!ing_codigo) <> Trim(vaSpread4.text) And i = vaSpread4.MaxRows Then
              
              vaSpread4.MaxRows = vaSpread4.MaxRows + 1
              vaSpread4.Row = vaSpread4.MaxRows
              vaSpread4.Col = 1: vaSpread4.text = Trim(RS1!ing_codigo)
              vaSpread4.Col = 2: vaSpread4.text = Trim(RS1!ing_nombre)
              vaSpread4.Col = 4: vaSpread4.text = IIf(RS1!ing_indppr = "1", "Real", "Propuesta")
           
           End If
       
       Next i
    
    Else
       
       vaSpread4.MaxRows = vaSpread4.MaxRows + 1
       vaSpread4.Row = vaSpread4.MaxRows
       vaSpread4.Col = 1: vaSpread4.text = Trim(RS1!ing_codigo)
       vaSpread4.Col = 2: vaSpread4.text = Trim(RS1!ing_nombre)
       vaSpread4.Col = 4: vaSpread4.text = IIf(RS1!ing_indppr = "1", "Real", "Propuesta")
    
    End If
    
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
    Check1(0).Value = RS1!ing_indpav
    Check1(1).Value = RS1!ing_indgrv
    fpHCarbono.Value = RS1!Huella_Carbono
    fpAyDouble1(0).Caption = IIf(RS1!ing_precos = 0, "", Format(RS1!ing_precos, fg_Pict(9, 4)))
    fpAyDate(0).Caption = IIf(RS1!ing_feccos = 0, "", fg_Ctod1(RS1!ing_feccos))
    
    If IsNull(RS1!ing_indppr) Or Trim(RS1!ing_indppr) = "" Then
      
      Combo2(2).ListIndex = -1
    
    Else
      
      Combo2(2).ListIndex = fg_buscacbo(Combo2, 2, 1, fg_pone_cero(Str(RS1!ing_indppr), 1))
    
    End If
    
    For i = 1 To vaSpread2.MaxRows
        
        vaSpread2.Row = i
        vaSpread2.Col = 1: CodN = Val(vaSpread2.Value)
        
        If RS3.State = 1 Then RS3.Close
        RS3.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS3 = vg_db.Execute("sgpadm_s_productonutriente 2, '" & CodIng & "', " & CodN & "")
        If Not RS3.EOF Then
            
            vaSpread2.Col = 3: vaSpread2.Value = RS3!pnu_canapo
        
        Else
            
            vaSpread2.Col = 3: vaSpread2.Value = 0
        
        End If
        RS3.Close: Set RS3 = Nothing
    
    Next i
    
    Est = False

Else
    
    fpText1(5).ControlType = ControlTypeNormal
End If
RS1.Close: Set RS1 = Nothing

End Sub

Private Sub MoverDatos4(CodIng As String)

Dim RS1    As New ADODB.Recordset
Dim RS3    As New ADODB.Recordset

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_s_ingrediente_V02 2, '" & CodIng & "', ''")

If Not RS1.EOF Then
   
   vaSpread4.Enabled = False
   Est = True
   fpText1(5).ControlType = ControlTypeStatic
   fpLongInteger1(4).text = RS1!ing_unimed
   fpayuda(4).Caption = RS1!unm_nombre
   fpDouble1(1).Value = RS1!ing_pctapr
   fpDouble1(2).Value = RS1!ing_pctcoc
   fpDouble1(3).Value = RS1!ing_pctnut
   fpDouble1(4).Value = RS1!ing_facnut
   Check1(0).Value = RS1!ing_indpav
   Check1(1).Value = RS1!ing_indgrv
   fpAyDouble1(0).Caption = IIf(RS1!ing_precos = 0, "", Format(RS1!ing_precos, fg_Pict(9, 4)))
   fpAyDate(0).Caption = IIf(RS1!ing_feccos = 0, "", fg_Ctod1(RS1!ing_feccos))
   
   For i = 1 To vaSpread2.MaxRows
       
       vaSpread2.Row = i
       vaSpread2.Col = 1: CodN = Val(vaSpread2.Value)
       RS3.Open "sgpadm_s_productonutriente 2, '" & CodIng & "', " & CodN & "", vg_db, adOpenStatic
       
       If Not RS3.EOF Then
           
          vaSpread2.Col = 3
          vaSpread2.Value = RS3!pnu_canapo
       
       Else
           
          vaSpread2.Col = 3
          vaSpread2.Value = 0
       
       End If
       
       RS3.Close: Set RS3 = Nothing
   
   Next i
   Est = False

Else
    
    fpText1(5).ControlType = ControlTypeNormal

End If
RS1.Close: Set RS1 = Nothing

End Sub

Private Sub MoverDatos5(procom As String)

Dim RS1    As New ADODB.Recordset

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("SELECT * FROM b_productocompra WHERE pco_codigo='" & procom & "'")

If Not RS1.EOF Then
    
    Est = True
    fpText1(8).ControlType = ControlTypeStatic
    fpText1(8).text = Trim(RS1!pco_codigo)
    fpText1(9).text = Trim(RS1!pco_nombre)
    fpDouble1(7).Value = IIf(IsNull(RS1!pco_undemb) Or RS1!pco_undemb = 0, 0, RS1!pco_undemb)
    fpDateTime1(1).text = IIf(IsNull(RS1!pco_fecven) Or RS1!pco_fecven = 0, "  /  /    ", Mid(RS1!pco_fecven, 7, 2) & "/" & Mid(RS1!pco_fecven, 5, 2) & "/" & Mid(RS1!pco_fecven, 1, 4))
    fpDateTime1(1).Enabled = IIf(IsNull(RS1!pco_fecven) Or RS1!pco_fecven = 0, False, True)
    Check1(4).Value = IIf(IsNull(RS1!pco_fecven) Or RS1!pco_fecven = 0, 0, 1)
    Est = False

Else
    
    fpText1(8).ControlType = ControlTypeNormal

End If
RS1.Close
Set RS1 = Nothing

End Sub

Private Sub MoverDatos6(codpro As String)

Dim RS4    As New ADODB.Recordset
Dim IndPrc As Long

If RS4.State = 1 Then RS4.Close
RS4.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS4 = vg_db.Execute("SELECT * FROM b_productocompra WHERE pco_codpro = '" & codpro & "' ORDER BY pco_nombre")

vaSpread5.MaxRows = 0
vaSpread5.MaxRows = RS4.RecordCount
IndPrc = 1

If Not RS4.EOF Then
   
   Do While Not RS4.EOF
      
'      vaSpread5.MaxRows = vaSpread5.MaxRows + 1
      vaSpread5.Row = IndPrc 'vaSpread5.MaxRows
      
      vaSpread5.Col = 1
      vaSpread5.text = Trim(RS4!pco_codigo)
      
      vaSpread5.Col = 2
      vaSpread5.text = Trim(RS4!pco_nombre)
      
      vaSpread5.Col = 3
      vaSpread5.TypeHAlign = TypeHAlignRight: vaSpread5.text = IIf(IsNull(RS4!pco_undemb) Or RS4!pco_undemb = 0, 0, Format(RS4!pco_undemb, fg_Pict(6, 2)))
      
      vaSpread5.Col = 4
      vaSpread5.TypeHAlign = TypeHAlignCenter: vaSpread5.text = IIf(IsNull(RS4!pco_fecven) Or RS4!pco_fecven = 0, "  /  /    ", Mid(RS4!pco_fecven, 7, 2) & "/" & Mid(RS4!pco_fecven, 5, 2) & "/" & Mid(RS4!pco_fecven, 1, 4))
      
      IndPrc = IndPrc + 1
      RS4.MoveNext
   
   Loop

End If
RS4.Close: Set RS4 = Nothing

If vaSpread5.MaxRows < 1 Then Exit Sub

vaSpread5.Row = 1
vaSpread5.Col = 1

If RS4.State = 1 Then RS4.Close
RS4.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS4 = vg_db.Execute("SELECT * FROM b_productocompra WHERE pco_codpro='" & codpro & "' AND pco_codigo='" & Trim(vaSpread5.text) & "'")

If Not RS4.EOF Then
    
    Est = True
    fpText1(8).ControlType = ControlTypeStatic
    fpText1(8).text = Trim(RS4!pco_codigo)
    fpText1(9).text = Trim(RS4!pco_nombre)
    fpDouble1(7).Value = IIf(IsNull(RS4!pco_undemb) Or RS4!pco_undemb = 0, 0, RS4!pco_undemb)
    fpDateTime1(1).text = IIf(IsNull(RS4!pco_fecven) Or RS4!pco_fecven = 0, "  /  /    ", Mid(RS4!pco_fecven, 7, 2) & "/" & Mid(RS4!pco_fecven, 5, 2) & "/" & Mid(RS4!pco_fecven, 1, 4))
    fpDateTime1(1).Enabled = IIf(IsNull(RS4!pco_fecven) Or RS4!pco_fecven = 0, False, True)
    Check1(4).Value = IIf(IsNull(RS4!pco_fecven) Or RS4!pco_fecven = 0, 0, 1)
    Est = False

Else
    
    fpText1(8).ControlType = ControlTypeNormal

End If
RS4.Close
Set RS4 = Nothing

End Sub

Private Sub Image1_Click(Index As Integer)
vg_codigo = 0
Est = True
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
    fpayuda(3).Caption = vg_nombre
'    fpLongInteger1(Index) = vg_codigo
    fpDouble1(8).Value = Val(vg_codigo)
'    On Error Resume Next: fpLongInteger1(Index).SetFocus
    On Error Resume Next: fpDouble1(8).SetFocus
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

Case 5
    
    vg_left = fpayuda(1).Left + 1920
    B_TabEst.LlenaDatos "b_retencionfuente", "ref_", "Retencion Fuente", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Then Exit Sub
    fpayuda(5).Caption = vg_nombre
    fpLongInteger1(5).Value = Val(vg_codigo)
    On Error Resume Next: fpLongInteger1(6).SetFocus
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False

Case 6
    
    vg_left = fpayuda(1).Left + 1920
    B_TabEst.LlenaDatos "b_retencionica", "rei_", "Retencion Ica", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Then Exit Sub
    fpayuda(6).Caption = vg_nombre
    fpLongInteger1(6).Value = Val(vg_codigo)
'    On Error Resume Next: fpLongInteger1(6).SetFocus
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False

Case 7
    
    vg_left = fpayuda(7).Left + 1920
    B_TabEst.LlenaDatos "a_unidadmed", "unm_", "Unidad de Medida", "GenUFacIng"
    B_TabEst.Show 1
    Me.Refresh
    If Val(vg_codigo) = 0 Then Exit Sub
    fpayuda(7).Caption = vg_nombre
    fpLongInteger1(7) = Val(vg_codigo)
    On Error Resume Next: fpLongInteger1(2).SetFocus
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo

End Select
Est = False

End Sub

Private Sub Option1_Click(Index As Integer)

If Est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(5) = False

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Frame5.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1" Or Mid(ValidarUsuario(Me), 2, 1) = "1", True, False)

If Frame5.Enabled = False Then
   
   vaSpread4.Row = -1: vaSpread4.Col = -1: vaSpread4.Lock = True
   vaSpread3.Row = -1: vaSpread3.Col = -1: vaSpread3.Lock = True
   vaSpread2.Row = -1: vaSpread2.Col = -1: vaSpread2.Lock = True

End If
'If SSTab1.Tab = 1 Then Gl_Ac_Botones Me, 8, 1, modo2: SSTab1.TabEnabled(0) = True: SSTab1.TabEnabled(4) = True
Frame3.Enabled = Frame5.Enabled
If Not EstDet And modo <> "A" Then vaSpread1_Click vaSpread1.ActiveCol, vaSpread1.ActiveRow

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS1    As New ADODB.Recordset
Dim i      As Long, fecven As Long
Dim ctrsto As Integer
Dim codigo As String, codbar As String, codcom As String, Nombre As String, codfam As Long, coduni As Long, codemb As Long, maepro As Long, Indppr As String
Dim uniemb As Double, upreco As Double, fecuco As String, propon As Double, ctacon As String, CodIng As String, noming As String
Dim facing As Double, facsto As Double, propre As Long, codref As Long, codrei As Long
Dim StrFam As String, StrFamb As String, fampr1 As Long, fampr2 As Long, fampr3 As Long, prodact As String, cuohor As String, tipord As String, tippro As String
Dim UnFacIng As Long

Select Case Button.Index

Case 1, 3 '-------> Agregar o Modificar
    
    modo = "A"
    If Button.Index = 3 Then modo = "M": If vaSpread1.MaxRows < 1 Then Exit Sub
    
    If modo = "A" Then
        
        EstDet = True
        vaSpread2.Enabled = False
        vaSpread3.Enabled = False
        For i = 1 To vaSpread2.MaxRows: vaSpread2.Col = 3: vaSpread2.Row = i: vaSpread2.Value = 0: Next i
        For i = 1 To vaSpread3.MaxRows: vaSpread3.Col = 3: vaSpread3.Row = i: vaSpread3.Value = 0: Next i
        
        vaSpread2.Enabled = True
        vaSpread3.Enabled = True
        vaSpread4.MaxRows = 0
        vaSpread5.MaxRows = 0
        
        lblNomPro(0).Caption = ""
        lblNomPro(1).Caption = ""
        
        Gl_Ac_Botones Me, 8, 1, modo
    
    End If
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False
    SSTab1.Tab = 1 ': EstDet = True
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    Me.Refresh
    If modo = "M" Then fpText1(0).Enabled = False: On Error Resume Next: fpText1(10).SetFocus
    If modo = "A" Then
       
       fpText1(0).Enabled = False: On Error Resume Next: fpText1(10).SetFocus: MoverDatos '---------> Solo efectos de limpieza
       '-------> Asignar codigo productos
       
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS1 = vg_db.Execute("sgpadm_Sel_productos 9, '', '', '" & vg_NUsr & "'")
'       If Not RS1.EOF Then RS1.MoveFirst: codigo = RS1!pro_codigo + 1 Else codigo = 1
       If Not RS1.EOF Then RS1.MoveFirst: codigo = RS1!pro_codigo Else codigo = 1
       
       RS1.Close: Set RS1 = Nothing
       Est = True
       fpText1(0).text = codigo
       If vg_Indppr = "2" Then
          
          Combo2(1).ListIndex = 1
          Combo2(1).Enabled = False
          Combo2(2).ListIndex = 1
          Combo2(2).Enabled = False
       
       ElseIf vg_Indppr = "1" Then
          
          Combo2(1).ListIndex = 0
          Combo2(1).Enabled = False
          Combo2(2).ListIndex = 0
          Combo2(2).Enabled = False
       
       Else
          
          Combo2(1).ListIndex = 0
          Combo2(2).ListIndex = 0
       
       End If
       Est = False
    
    End If

Case 5 '-------> Eliminar
    
    If vaSpread1.MaxRows < 1 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codigo = vaSpread1.Value
    
    If MsgBox("Eliminar Dato", vbQuestion + vbYesNo, MsgTitulo) = vbYes Then
        
        '-------> Borrar Productos y productos ingrediente
        operr = 2
        vg_db.BeginTrans
        vg_db.Execute "DELETE b_productosimp FROM b_productosimp WHERE ipr_codpro = '" & codigo & "'"
        vg_db.Execute "DELETE b_productosing FROM b_productosing WHERE pri_codpro = '" & codigo & "'"
        vg_db.Execute "DELETE b_productos FROM b_productos WHERE pro_codigo = '" & codigo & "'"
        vg_db.CommitTrans
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        vaSpread1.Row = vaSpread1.MaxRows
        
        If fpText.text = "" Then
            
            Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
        
        Else
            
            Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
        
        End If
        SSTab1.Tab = 0
    
    End If
    
    On Error Resume Next: fpText.SetFocus
    modo = "": MoverDatos
    Gl_Ac_Botones Me, 1, 1, modo

Case 7 '-------> Actualiza Grilla
    
    fpText.text = ""
    MoverDatosGrilla
    modo = "": MoverDatos

Case 10 '------- Cancelar
    
    If MsgBox("Cancelar Operación", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If LimpiaDato(Trim(fpText1(5).text)) = "" And modo <> "A" Then MsgBox "Debe relacionar ingrediente...", vbCritical, MsgTitulo: Exit Sub
    modo = "": MoverDatos
    EstDet = True
    If vg_Indppr = 2 Then Gl_Ac_BotonesRealPropuesta Me, 1, 1, modo, vg_Indppr, ComboValOri Else Gl_Ac_Botones Me, 1, 1, modo
    Me.HelpContextID = "1011000"
    Gl_Ac_Botones Me, 8, 1, modo2
    Me.HelpContextID = vg_OpcM
    Gl_Ac_Botones Me, 9, 1, modo3
    SSTab1.TabEnabled(0) = True
    
    If vaSpread1.MaxRows = 0 Then
        
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(4) = False
        SSTab1.TabEnabled(5) = False
        SSTab1.Tab = 0
    
    ElseIf vaSpread1.MaxRows > 0 Then
        
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(4) = True
        SSTab1.TabEnabled(5) = True
    
    End If

Case 12 '-------> Confirmar
    
    Dim indice As String
    Dim CodN As Long, CanN As Double, CodS
    
    If modo = "A" Or modo = "M" Then
       
       If LimpiaDato(Trim(fpText1(0).text)) = "" Or Trim(fpayuda(0).Caption) = "" Or _
          LimpiaDato(Trim(fpText1(2).text)) = "" Or _
          LimpiaDato(Trim(fpLongInteger1(0).text)) = "" Or _
          LimpiaDato(Trim(fpDouble1(8).text)) = "" Or _
          fpDouble1(0).Value = 0 Or fpDouble1(5).Value = 0 Or fpDouble1(6).Value = 0 Or Combo2(0).ListIndex = -1 Or _
          Combo2(4).ListIndex = -1 Then MsgBox "Debe ingresar información...", vbCritical, MsgTitulo: Exit Sub
       
       If LimpiaDato(Trim(fpLongInteger1(1).text)) = "" Or LimpiaDato(Trim(fpayuda(1).Caption)) = "" Then
       
           MsgBox "Debe ingresar unidad stock...", vbCritical, MsgTitulo
           Exit Sub
                  
       End If
       
       If LimpiaDato(Trim(fpLongInteger1(2).text)) = "" Or LimpiaDato(Trim(fpayuda(2).Caption)) = "" Then
       
           MsgBox "Debe ingresar unidad embalaje...", vbCritical, MsgTitulo
           Exit Sub
                  
       End If
       
       If LimpiaDato(Trim(fpLongInteger1(7).text)) = "" Or LimpiaDato(Trim(fpayuda(7).Caption)) = "" Then
       
           MsgBox "Debe ingresar unidad factor ingrediente...", vbCritical, MsgTitulo
           Exit Sub
                  
       End If
       
       If LimpiaDato(Trim(fpText1(5).text)) = "" Then MsgBox "Debe relacionar ingrediente...", vbCritical, MsgTitulo: Exit Sub
       
       If Check1(2).Value = 1 And Trim(fpDateTime1(0).text) = "" Then MsgBox "Debe ingresar fecha vencimiento...", vbCritical, MsgTitulo: Exit Sub
        
        '-------> Validar familia productos
        
        If modo = "M" Then
           
           If RS1.State = 1 Then RS1.Close
           RS1.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
    
           RS1.Open "sgpadm_Sel_productos 10, '" & LimpiaDato(Trim(fpText1(0).text)) & "', '', '" & vg_NUsr & "'", vg_db, adOpenStatic
           If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
           
           If fampr3 <> RS1!pro_codtip Then
              
              operr = 1
           
           End If
           RS1.Close: Set RS1 = Nothing
        
        End If
        
        '-------> Fin Validar familia productos
        If MsgBox("** Importante **" & VgLinea & "Revise que la definición de impuestos y cuenta contable del producto sea correcta" & VgLinea & "Desea grabar ?...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        If Val(fg_codigocbo(Combo2, 1, 1, "")) = "-1" Then MsgBox "Debe Seleccionar tipo de producto...", vbCritical, MsgTitulo: Exit Sub
        fecven = 0
        fecven = IIf(Check1(2).Value = 0, 0, Mid(fpDateTime1(0).text, 7, 4) & Mid(fpDateTime1(0).text, 4, 2) & Mid(fpDateTime1(0).text, 1, 2))
        If fecven < Val(Format(Date, "yyyymmdd")) And fecven > 0 Then prodact = "S" Else prodact = "N"
        Est = True: Toolbar2_ButtonClick Toolbar2.Buttons(8): Est = False
        operr = 2
        codigo = LimpiaDato(Trim(fpText1(0).text))
        codbar = LimpiaDato(Trim(fpText1(1).text))
        Nombre = LimpiaDato(Trim(Mid(fpText1(2).text, 1, 50)))
        codcom = LimpiaDato(Trim(fpText1(10).text))
        codfam = fpLongInteger1(0).Value
        coduni = fpLongInteger1(1).Value
        codemb = fpLongInteger1(2).Value
        ctacon = Trim(fpDouble1(8).text)
        UnFacIng = fpLongInteger1(7).Value
        uniemb = fpDouble1(0).Value
        facing = fpDouble1(5).Value
        facsto = fpDouble1(6).Value
        upreco = Val(fpAyDouble1(1).Caption)
        propon = Val(fpAyDouble1(2).Caption)
        fecven = IIf(Check1(2).Value = 0, 30000101, Mid(fpDateTime1(0).text, 7, 4) & Mid(fpDateTime1(0).text, 4, 2) & Mid(fpDateTime1(0).text, 1, 2))
        ctrsto = IIf(Check1(3).Value = 0, 0, 1)
        fecuco = IIf(Trim(fpAyDate(7).Caption) = "", "Null", "cdate('" & fpAyDate(7).Caption & "')")
        maepro = Val(fg_codigocbo(Combo2, 0, 1, ""))
        Indppr = Val(fg_codigocbo(Combo2, 1, 1, ""))
        codref = Val(fpLongInteger1(5).Value)
        codrei = Val(fpLongInteger1(6).Value)
        cuohor = IIf(Check1(5).Value = 0, "N", "S")
        tipord = Val(fg_codigocbo(Combo2, 4, 1, ""))
        tippro = IIf(Option1(0).Value = True, "0", "1")
        Dim unimed As Long
        
        If modo = "A" Then
           
           codigo = 0
           
           If RS1.State = 1 Then RS1.Close
           RS1.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient

           Set RS1 = vg_db.Execute("sgpadm_Ins_Productos '', '" & codbar & "', '" & codcom & "', " & codfam & ", '" & Nombre & "', " & coduni & ", " & facing & ", " & facsto & ", " & codemb & ", " & uniemb & ", " & upreco & ", " & fecuco & ", " & propon & ", '" & ctacon & "', " & fecven & ", " & ctrsto & ", " & maepro & ", " & Indppr & ", " & codref & ", " & codrei & ", '" & cuohor & "', '" & tipord & "', '" & tippro & "', " & UnFacIng & ", '" & vg_NUsr & "'")
           If Not RS1.EOF Then
              
              codigo = RS1!indice
           
           End If
           RS1.Close: Set RS1 = Nothing
           
           For i = 1 To vaSpread4.MaxRows
               
               vaSpread4.Row = i: vaSpread4.Col = 1: CodIng = "": CodIng = vaSpread4.text
               vaSpread4.Col = 2: noming = "": noming = Trim(vaSpread4.text)
               vaSpread4.Col = 3: propre = 0: propre = IIf(Trim(vaSpread4.text) = "", 0, Val(vaSpread4.text))
               vg_db.Execute "sgpadm_iu_productosing 'A', '" & codigo & "', '" & CodIng & "', " & propre & ""
               vg_db.Execute "sgpadm_p_actuaproding '" & CodIng & "'"
               
               '-------> Traer codigo unidad medida del ingrediente
               
               If RS1.State = 1 Then RS1.Close
               RS1.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
    
               RS1.Open "sgpadm_s_ingrediente_V02 5, '" & CodIng & "', ''", vg_db, adOpenStatic
               If RS1.EOF Then RS1.Close: Set RS1 = Nothing: vg_db.RollbackTrans: MsgBox "No existe código unidad medida ingrediente, proceso cancelado...", vbCritical, MsgTitulo: Exit Sub
               unimed = 0: unimed = RS1!ing_unimed
               RS1.Close: Set RS1 = Nothing
               If CodIng <> "720" And CodIng <> "762" And CodIng <> "742" Then
               
               ElseIf CodIng = "720" Or CodIng = "762" Or CodIng = "742" Then
               
               End If
           
           Next i
           
           fpText1(0).text = codigo
           vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
           vaSpread1.SetActiveCell 1, vaSpread1.Row
        
        Else
            
            vg_db.Execute "sgpadm_Upd_Productos '" & codigo & "', '" & codbar & "', '" & codcom & "', " & codfam & ", " & _
                          "'" & Nombre & "', " & coduni & ", " & facing & ", " & facsto & ", " & codemb & ", " & uniemb & ", " & _
                          "" & upreco & ", " & fecuco & ", " & propon & ", '" & ctacon & "', " & fecven & ", " & ctrsto & ", " & _
                          "" & maepro & ", " & Indppr & ", " & codref & ", " & codrei & ", '" & cuohor & "', '" & tipord & "', '" & tippro & "', " & UnFacIng & ", '" & vg_NUsr & "'"
                          
            vg_db.Execute "DELETE b_productosing FROM b_productosing WHERE pri_codpro = '" & codigo & "'"
            For i = 1 To vaSpread4.MaxRows
                
                vaSpread4.Row = i: vaSpread4.Col = 1: CodIng = "": CodIng = vaSpread4.text
                vaSpread4.Col = 2: noming = Trim(vaSpread4.text)
                vaSpread4.Col = 3: propre = 0: propre = IIf(Trim(vaSpread4.text) = "", 0, Val(vaSpread4.text))
                vg_db.Execute "sgpadm_iu_productosing 'A', '" & codigo & "', '" & CodIng & "', " & propre & ""
                vg_db.Execute "sgpadm_p_actuaproding '" & CodIng & "'"
                
                '-------> Traer codigo unidad medida del ingrediente
                
                If RS1.State = 1 Then RS1.Close
                RS1.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
 
                RS1.Open "sgpadm_s_ingrediente_V02 5, '" & CodIng & "', ''", vg_db, adOpenStatic
                If RS1.EOF Then RS1.Close: Set RS1 = Nothing: vg_db.RollbackTrans: MsgBox "No existe código unidad medida ingrediente, proceso cancelado...", vbCritical, MsgTitulo: Exit Sub
                unimed = 0: unimed = RS1!ing_unimed
                RS1.Close: Set RS1 = Nothing
                If CodIng <> "720" And CodIng <> "762" And CodIng <> "742" Then
                ElseIf CodIng = "720" Or CodIng = "762" Or CodIng = "742" Then
                End If
            
            Next i
        
        End If
        
        vg_db.Execute "DELETE FROM b_productosimp WHERE ipr_codpro = '" & codigo & "'"
        
        For Fila = 1 To vaSpread3.MaxRows
            
            vaSpread3.Row = Fila
            vaSpread3.Col = 1: CodN = Val(vaSpread3.Value)
            vaSpread3.Col = 3: CanN = Val(vaSpread3.Value)
            If CanN <> 0 Then vg_db.Execute "sgpadm_iu_productosimp 'A', '" & codigo & "', " & CodN & ""
        
        Next Fila
        
        '-------> Validar si el producto esta vigente
        Dim codpro As String
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        Set RS1 = vg_db.Execute("sgpadm_Sel_productos 11, '" & CodIng & "', '', '" & vg_NUsr & "'")
        If Not RS1.EOF Then
           
           '-------> Actualiza codigo compra y pedido de ultimo producto para ingrediente
           codpro = RS1!pro_codigo
           RS1.Close: Set RS1 = Nothing
           vg_db.Execute "sgpadm_iu_ingrediente_V02 'M2', '" & CodIng & "', '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, '" & codpro & "', '" & codpro & "', ''"
        
        Else
           
           RS1.Close: Set RS1 = Nothing
        
        End If

        'ACTUALIZA GRILLA DESPUES DE MODIFICAR
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = -1
        If fecven < Val(Format(Date, "yyyymmdd")) And fecven > 0 Then vaSpread1.BackColor = Shape1(1).FillColor Else vaSpread1.BackColor = Shape1(0).FillColor
        vaSpread1.Col = 1: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.Value = LimpiaDato(Trim(fpText1(0).text))
        vaSpread1.Col = 2: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.Value = LimpiaDato(Trim(fpText1(2).text))
        vaSpread1.Col = 3: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.Value = Trim(Mid(Combo2(0).text, 1, 150))
        vaSpread1.Col = 4: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.Value = IIf(fpAyDouble1(2).Caption <> "", Format(fpAyDouble1(2).Caption, fg_Pict(6, 2)), Format(fpAyDouble1(2).Caption, 0))
        vaSpread1.Col = 5: vaSpread1.TypeHAlign = 0: vaSpread1.Value = codfam
        vaSpread1.Col = 6: vaSpread1.TypeHAlign = 0: vaSpread1.Value = IIf(Indppr = "1", "Real", "Propuesta")
        vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread1.ColUserSortIndicator(IIf(Combo1.ItemData(Combo1.ListIndex) = 0, 1, 2)) = ColUserSortIndicatorAscending
        vaSpread1.SortKey(1) = IIf(Combo1.ItemData(Combo1.ListIndex) = 0, 1, 2)
        vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
        Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
    
    End If
    modo = ""
    MoverDatos
    Gl_Ac_Botones Me, 1, 1, modo
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    SSTab1.TabEnabled(4) = True
    SSTab1.TabEnabled(5) = True
    SSTab1.Tab = 0
    Est = True: Est = False

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
If Err = -2147467259 Or Err = -2147217900 Then
   If operr = 1 Then
       RS1.Close: Set RS1 = Nothing
       MsgBox "No puede Cambiar Familia Producto Tecfood...", vbCritical, "Error": Exit Sub
   ElseIf operr = 2 Then
      vg_db.RollbackTrans
      MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
   End If

End If
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub MoverDatosGrilla()


On Error GoTo Man_Error

Dim codpro As String, codTippro As Long, nomTippro As String
Dim IndPro As Long

fg_carga ""
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
    
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_Sel_productos 1, '', '', '" & vg_NUsr & "'")
codTippro = 0
nomTippro = ""
IndPro = 1
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = Shape1(0).FillColor
vaSpread1.MaxRows = RS1.RecordCount

If Not RS1.EOF Then

    Do While Not RS1.EOF
       
'       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       vaSpread1.Row = IndPro 'vaSpread1.MaxRows
       
       vaSpread1.Col = 1
       vaSpread1.TypeHAlign = TypeHAlignLeft
       vaSpread1.text = Trim(RS1!pro_codigo)
       vaSpread1.Col = -1
       If RS1!pro_fecven < Val(Format(Date, "yyyymmdd")) And RS1!pro_fecven > 0 Then vaSpread1.BackColor = Shape1(1).FillColor
                
       vaSpread1.Col = 2
       vaSpread1.TypeHAlign = TypeHAlignLeft
       vaSpread1.Value = Trim(RS1!pro_nombre)
                
       vaSpread1.Col = 3
       vaSpread1.TypeHAlign = TypeHAlignLeft
       vaSpread1.Value = IIf(RS1!pro_maepro = 0, "Ambos", Trim(RS1!tis_nombre))
                
       vaSpread1.Col = 4
       vaSpread1.TypeHAlign = TypeHAlignRight
       vaSpread1.Value = IIf(RS1!pro_propon < 1, "", Format(RS1!pro_propon, fg_Pict(9, vg_DPr)))
         
       vaSpread1.Col = 5
       vaSpread1.text = RS1!pro_codtip
       
       vaSpread1.Col = 6
       vaSpread1.text = IIf(IsNull(RS1!pro_indppr) Or Trim(RS1!pro_indppr) = "2", "Propuesta", "Real")

       codTippro = RS1!pro_codtip
       nomTippro = vaSpread1.text
       
       IndPro = IndPro + 1
       
       RS1.MoveNext
    
    Loop
    
    Gl_Ac_Botones Me, 1, 1, modo

Else
    
    Gl_Ac_Botones Me, 1, 2, modo

End If
vaSpread1.SortKey(1) = 1
vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
SSTab1.Tab = 0
RS1.Close
Set RS1 = Nothing
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
vaSpread1.Visible = True
fg_descarga

Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub ConfiControlesProducto(Index As Integer, habilita As Boolean)
  Select Case Index
    
    Case 1
     
     fpText1(2).Enabled = IIf(habilita = True, True, False)
     fpLongInteger1(0).Enabled = IIf(habilita = True, True, False)
     fpLongInteger1(1).Enabled = IIf(habilita = True, True, False)
     fpLongInteger1(7).Enabled = IIf(habilita = True, True, False)
     fpDouble1(6).Enabled = IIf(habilita = True, True, False)
     fpDouble1(5).Enabled = IIf(habilita = True, True, False)
     fpLongInteger1(2).Enabled = IIf(habilita = True, True, False)
     fpDouble1(0).Enabled = IIf(habilita = True, True, False)
     fpDouble1(8).Enabled = IIf(habilita = True, True, False)
     Image1(0).Enabled = IIf(habilita = True, True, False)
     Image1(1).Enabled = IIf(habilita = True, True, False)
     Image1(2).Enabled = IIf(habilita = True, True, False)
     Image1(3).Enabled = IIf(habilita = True, True, False)
     Check1(3).Enabled = IIf(habilita = True, True, False)
     Check1(2).Enabled = IIf(habilita = True, True, False)
     fpDateTime1(0).Enabled = IIf(habilita = True, True, False)
     Frame5.Enabled = IIf(habilita = True, True, False)
     Frame7.Enabled = IIf(habilita = True, True, False)
     Combo2(0).Enabled = IIf(habilita = True, True, False)
     Combo2(1).Enabled = IIf(habilita = True, True, False)
     vaSpread3.Enabled = IIf(habilita = True, True, False)
     Frame11.Enabled = IIf(habilita = True, True, False)
    
    Case 2
     
     Frame8.Enabled = IIf(habilita = True, True, False)
     Frame3.Enabled = IIf(habilita = True, True, False)
     Frame1(1).Enabled = IIf(habilita = True, True, False)
     fpText1(6).Enabled = IIf(habilita = True, True, False)
     fpText1(7).Enabled = IIf(habilita = True, True, False)
     fpLongInteger1(4).Enabled = IIf(habilita = True, True, False)
     fpDouble1(1).Enabled = IIf(habilita = True, True, False)
     fpDouble1(2).Enabled = IIf(habilita = True, True, False)
     fpDouble1(3).Enabled = IIf(habilita = True, True, False)
     Combo2(2).Enabled = IIf(habilita = True, True, False)
     fpDouble1(4).Enabled = IIf(habilita = True, True, False)
     Check1(0).Enabled = IIf(habilita = True, True, False)
     Check1(1).Enabled = IIf(habilita = True, True, False)
     Image1(4).Enabled = IIf(habilita = True, True, False)
  
  End Select

End Sub

Private Sub MoverDatos()

On Error GoTo Man_Error

Dim RS1    As New ADODB.Recordset
Dim codigo As String, CodIng As String, i As Long, CodN As Long
fg_carga ""
Limpia 1: Limpia 2: Limpia 3: Limpia 4: Limpia 5
'-------------------------Mueve datos de Detalle (1ª Carpeta)--------------------------
If modo = "" Then
    
    If vaSpread1.MaxRows > 0 Then vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: codigo = vaSpread1.Value
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("sgpadm_Sel_productos 2, '" & codigo & "', '', '" & vg_NUsr & "'")
    Est = True
    If Not RS1.EOF Then
        
        '-------> Detalle
        fpText1(0).text = Trim(RS1!pro_codigo)
        fpText1(1).text = TipoDato(RS1!pro_codbar, "")
        fpText1(10).text = TipoDato(RS1!pro_codcom, "")
        fpLongInteger1(0).text = RS1!pro_codtip
               
        fpayuda(0).Caption = RS1!FamProducto
        fpText1(2).text = Trim(RS1!pro_nombre)
        Frame11.Caption = Trim(RS1!pro_nombre)
        lblNomPro(0).Caption = Trim(RS1!pro_nombre)
        lblNomPro(1).Caption = Trim(RS1!pro_nombre)
        fpLongInteger1(1).text = RS1!pro_coduni
        fpayuda(1).Caption = Trim(RS1!uni_nombre)
        fpLongInteger1(7).text = RS1!pro_UniFactorIng
        fpayuda(7).Caption = Trim(RS1!unm_nomcor)
        fpLongInteger1(2).text = RS1!pro_codemb
        fpayuda(2).Caption = Trim(IIf(IsNull(RS1!emb_nombre), "", RS1!emb_nombre))
        fpDouble1(8).text = RS1!pro_ctacon
        fpayuda(3).Caption = Trim(RS1!cta_nombre)
        fpDouble1(0).text = RS1!pro_uniemb
        fpDouble1(5).text = RS1!pro_facing
        fpDouble1(6).text = RS1!pro_facsto
        fpAyDouble1(1).Caption = IIf(RS1!pro_upreco = 0, "", Format(RS1!pro_upreco, fg_Pict(9, vg_DPr)))
        fpAyDate(7).Caption = IIf(IsNull(RS1!pro_fecuco), "", RS1!pro_fecuco)
        fpAyDouble1(2).Caption = IIf(RS1!pro_propon = 0, "", Format(RS1!pro_propon, fg_Pict(9, vg_DPr)))
        fpText1(5).text = ""
        fpDateTime1(0).text = IIf(IsNull(RS1!pro_fecven) Or RS1!pro_fecven = 0, "  /  /    ", Mid(RS1!pro_fecven, 7, 2) & "/" & Mid(RS1!pro_fecven, 5, 2) & "/" & Mid(RS1!pro_fecven, 1, 4))
        fpDateTime1(0).Enabled = IIf(IsNull(RS1!pro_fecven) Or RS1!pro_fecven = 0, False, True)
        Check1(2).Value = IIf(IsNull(RS1!pro_fecven) Or RS1!pro_fecven = 0, 0, 1)
        Check1(3).Value = IIf(IsNull(RS1!pro_ctrsto) Or RS1!pro_ctrsto = 0, 0, 1)
        Combo2(0).ListIndex = fg_buscacbo(Combo2, 0, 1, (RS1!pro_maepro))
        fpLongInteger1(5).text = IIf(IsNull(RS1!pro_codref) Or RS1!pro_codref = 0, "", RS1!pro_codref)
        fpayuda(5).Caption = Trim(IIf(IsNull(RS1!ref_nombre), "", RS1!ref_nombre))
        fpLongInteger1(6).text = IIf(IsNull(RS1!pro_codrei) Or RS1!pro_codrei = 0, "", RS1!pro_codrei)
        fpayuda(6).Caption = Trim(IIf(IsNull(RS1!rei_nombre), "", RS1!rei_nombre))
        Check1(5).Value = IIf(IsNull(RS1!pro_cuohor) Or RS1!pro_cuohor = "N", 0, 1)
        If IsNull(RS1!pro_indppr) Or Trim(RS1!pro_indppr) = "" Then
          
          Combo2(1).ListIndex = -1
        
        Else
          
          Combo2(1).ListIndex = fg_buscacbo(Combo2, 1, 1, Trim(RS1!pro_indppr))
          
          If vg_Indppr = "2" And RS1!pro_indppr = "1" Then
             
             Gl_Ac_Botones Me, 1, 11, modo
          
          Else
             
             Gl_Ac_Botones Me, 1, 1, modo
          
          End If
          
          If vg_Indppr = "2" And vg_Indppr <> RS1!pro_indppr Then ConfiControlesProducto 1, False Else ConfiControlesProducto 1, True
        
        End If
        Combo2(1).Enabled = IIf(vg_Indppr = "1" Or vg_Indppr = "2", False, True)
        Combo2(4).ListIndex = fg_buscacbo(Combo2, 4, 1, (RS1!pro_tipord))
        Option1(0).Value = IIf(IsNull(RS1!pro_tippro) Or RS1!pro_tippro = "0", True, False)
        Option1(1).Value = IIf(RS1!pro_tippro = "0", False, True)
    
    End If
    RS1.Close: Set RS1 = Nothing
    Est = True
    '-------------------------Mueve datos de Ingrediente y Aporte Nutricional (2ª Carpeta)--------------------------
    vaSpread4.MaxRows = 0
    fpText1(5).ControlType = ControlTypeNormal
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("sgpadm_s_productoingrediente 1, '" & codigo & "', ''")
    
    If RS1.EOF Or IsNull(RS1!nReg) Then
       
       RS1.Close: Set RS1 = Nothing
       fpText1(5).ControlType = ControlTypeNormal
    
    Else
       
       RS1.Close: Set RS1 = Nothing
       modo2 = "": MoverDatos2 codigo
    
    End If
    '-------------------------Mueve datos formato de compra --------------------------------
    vaSpread5.MaxRows = 0
    fpText1(8).ControlType = ControlTypeNormal
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS1 = vg_db.Execute("SELECT COUNT(pco_codpro) AS nreg FROM b_productocompra WHERE pco_codpro='" & codigo & "'")
    If RS1.EOF Or IsNull(RS1!nReg) Then
       
       RS1.Close: Set RS1 = Nothing
       fpText1(8).ControlType = ControlTypeNormal
    
    Else
       
       RS1.Close: Set RS1 = Nothing
       modo3 = "": MoverDatos6 codigo
    
    End If
    '-------------------------Mueve datos de Impuesto (3ª Carpeta)--------------------------
    Est = True
    vaSpread3.Enabled = False
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS1 = vg_db.Execute("sgpadm_s_productoimpuesto 1, '" & codigo & "', ''")
    If Not RS1.EOF Then
       
       Do While Not RS1.EOF
          
          vaSpread3.Row = vaSpread3.SearchCol(1, -1, vaSpread3.MaxRows, Trim(CStr(RS1!ipr_codimp)), SearchFlagsEqual)
          If vaSpread3.Row > 0 Then vaSpread3.Col = 3: vaSpread3.Value = 1
          RS1.MoveNext
       
       Loop
    
    End If
    RS1.Close: Set RS1 = Nothing
    vaSpread3.Enabled = True
    '-------------------------Muevo datos formato de compras ----------------------------------
    Est = True
    vaSpread6.Visible = False
    vaSpread6.Row = -1: vaSpread6.Col = -1
    vaSpread6.BackColor = Shape1(0).FillColor
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS1 = vg_db.Execute("sgpadm_Sel_productos 17, '" & codigo & "', 'sac', '" & vg_NUsr & "'")
    
    Dim IndSac As Long
    vaSpread6.MaxRows = 0
    vaSpread6.MaxRows = RS1.RecordCount
    IndSac = 1
    
    Do While Not RS1.EOF
       
       Frame11.Caption = IIf(IsNull(RS1!fcs_codsgp), "", RS1!fcs_codsgp) & " - " & IIf(IsNull(RS1!pro_nombre), "", Trim(RS1!pro_nombre))
       
    '   vaSpread6.MaxRows = vaSpread6.MaxRows + 1
       vaSpread6.Row = IndSac 'vaSpread6.MaxRows
       vaSpread6.Col = 1
       
       If RS1!fcs_sgppre = 0 Then
          
          vaSpread6.CellType = CellTypeStaticText
       
       Else
          
          vaSpread7.Row = 1: vaSpread7.Col = 1
          vaSpread6.CellType = CellTypePicture
          vaSpread6.TypePictCenter = True
          vaSpread6.TypePictMaintainScale = True
          vaSpread6.TypePictStretch = True
          vaSpread6.TypePictPicture = vaSpread7.TypePictPicture
       
       End If
       
       vaSpread6.Col = 2: vaSpread6.text = IIf(IsNull(RS1(0)), "", RS1(0)) 'IIf(IsNull(RS1!foc_codsac), "", RS1!foc_codsac)
       vaSpread6.Col = 3: vaSpread6.text = IIf(IsNull(RS1(1)), "", RS1(1)) 'IIf(IsNull(RS1!foc_nomsac), "", RS1!foc_nomsac)
       vaSpread6.Col = 4: vaSpread6.text = IIf(IsNull(RS1(2)), "", RS1(2)) 'IIf(IsNull(RS1!foc_unisac), "", RS1!foc_unisac)
       vaSpread6.Col = 5: vaSpread6.text = IIf(IsNull(RS1(5)), "", RS1(5)) 'IIf(IsNull(RS1!fcs_codsgp), "", RS1!fcs_codsgp)
       vaSpread6.Col = 6: vaSpread6.text = IIf(IsNull(RS1(8)), "", RS1(8)) 'IIf(IsNull(RS1!pro_nombre), "", RS1!pro_nombre)
'       vaSpread6.Col = 7: vaSpread6.TypeCurrencyDecPlaces = vg_RDCa: vaSpread6.text = IIf(IsNull(RS1!foc_faccon), 0, RS1!foc_faccon)
       vaSpread6.Col = 7: vaSpread6.text = IIf(IsNull(RS1(11)), 0, RS1(11)) 'IIf(IsNull(RS1!foc_faccon), 0, RS1!foc_faccon)
       
       IndSac = IndSac + 1
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing
    vaSpread6.Visible = True
    
    '-------------------------Muevo datos formato de compras sap ----------------------------------
    Est = True
    vaSpread10.Visible = False
    vaSpread10.Row = -1: vaSpread10.Col = -1
    vaSpread10.BackColor = Shape1(0).FillColor
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("sgpadm_Sel_productos 17, '" & codigo & "', 'sap', '" & vg_NUsr & "'")
    Dim IndSap As Long
    vaSpread10.MaxRows = 0
    vaSpread10.MaxRows = RS1.RecordCount
    IndSap = 1
    
    Do While Not RS1.EOF
       
       Frame14.Caption = IIf(IsNull(RS1!pro_codigo), "", RS1!pro_codigo) & " - " & IIf(IsNull(RS1!pro_nombre), "", Trim(RS1!pro_nombre)) & " - Cta. Con.(" & IIf(IsNull(RS1!pro_ctacon), "", Trim(RS1!pro_ctacon)) & ")"
    '   vaSpread10.MaxRows = vaSpread10.MaxRows + 1
       vaSpread10.Row = IndSap 'vaSpread10.MaxRows
       vaSpread10.Col = 1
       
       If RS1(6) = 0 Then
          
          vaSpread10.CellType = CellTypeStaticText
       
       Else
          
          vaSpread9.Row = 1: vaSpread9.Col = 1
          vaSpread10.CellType = CellTypePicture
          vaSpread10.TypePictCenter = True
          vaSpread10.TypePictMaintainScale = True
          vaSpread10.TypePictStretch = True
          vaSpread10.TypePictPicture = vaSpread9.TypePictPicture
       
       End If
       
       vaSpread10.Col = 2: vaSpread10.text = IIf(IsNull(RS1(0)), "", RS1(0)) 'IIf(IsNull(RS1!foc_codsac), "", RS1!foc_codsac)
       vaSpread10.Col = 3: vaSpread10.text = IIf(IsNull(RS1(1)), "", RS1(1)) 'IIf(IsNull(RS1!foc_nomsac), "", RS1!foc_nomsac)
       vaSpread10.Col = 4: vaSpread10.text = IIf(IsNull(RS1(2)), "", RS1(2)) 'IIf(IsNull(RS1!foc_unisac), "", RS1!foc_unisac)
       vaSpread10.Col = 5: vaSpread10.text = IIf(IsNull(RS1(5)), "", RS1(5)) 'IIf(IsNull(RS1!fcs_codsgp), "", RS1!fcs_codsgp)
       vaSpread10.Col = 6: vaSpread10.text = IIf(IsNull(RS1(8)), "", RS1(8)) 'IIf(IsNull(RS1!pro_nombre), "", RS1!pro_nombre)
'       vaSpread10.Col = 7: vaSpread6.TypeCurrencyDecPlaces = vg_RDCa: vaSpread6.text = IIf(IsNull(RS1!foc_faccon), 0, RS1!foc_faccon)
       vaSpread10.Col = 7: vaSpread10.text = IIf(IsNull(RS1(13)), 0, RS1(13)) 'IIf(IsNull(RS1!fcs_ctacon), 0, RS1!fcs_ctacon)
       vaSpread10.Col = 8: vaSpread10.text = IIf(IsNull(RS1(11)), 0, RS1(11)) 'IIf(IsNull(RS1!foc_faccon), 0, RS1!foc_faccon)
       
       IndSap = IndSap + 1
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing
    vaSpread10.Visible = True
    
    Est = False

End If
Est = False
fg_descarga

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim CodN As Long, CanN As Double, i As Long, j As Long
Dim codigo As String, Nombre As String, nomfan As String, unimed As Long, pctapr As Double, pctcoc As Double
Dim pctnut As Double, facnut As Double, indpav As Long, indgrv As Long, precos As Double, feccos As Long, codpro As String, CodIng As String, Indppr As String
Dim StrFam As String, StrFamb As String, noming As String
Dim fampr1 As Long, fampr2 As Long, fampr3 As Long, inderror As Long
Dim HuellaCarbono As Double

On Error GoTo Man_Error

'------> INCLUIR(1)-BORRAR(3)-CANCELAR(6)-CONFIRMAR(8)-BUSCAR(11)-IMPRIMIR(12)
Select Case Button.Index

Case 1 '-------> Nuevo - Limpia
    
    modo2 = "A"
    If modo = "" Then modo = "M"
    Limpia 2
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(4) = False
    SSTab1.TabEnabled(5) = False
    Gl_Ac_Botones Me, 8, 0, modo2
    
    If modo2 = "A" Then
       
       ConfiControlesProducto 2, True
       fpText1(5).Enabled = False
       '-------> Asignar correlativo ingrediente
       codigo = ""

       If RS3.State = 1 Then RS3.Close
       RS3.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS3.Open "sgpadm_s_ingrediente_V02 6, '', ''", vg_db, adOpenStatic
       
       If Not RS3.EOF Then
          
          RS3.MoveFirst
          codigo = RS3!ing_codigo + 1
       
       Else
          
          codigo = 1
       
       End If
       
       RS3.Close: Set RS3 = Nothing
       fpText1(5).text = codigo
    
    End If

Case 3 '-------> Borra
    '-------> Validar si existen formato compra
    If vaSpread5.MaxRows > 0 Then MsgBox "Para eliminar ingrediente, Debe eliminar formato de compra...", vbCritical, MsgTitulo: Exit Sub
    '-------> Fin validar si existen formato compra
    If MsgBox("Se perderan los valores de nutrientes y deberá vincular otro ingrediente al producto, desea eliminar...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    '-------> Borrar Tabla Ingredientes
    codigo = LimpiaDato(Trim(fpText1(5).text))
    For i = 1 To vaSpread4.MaxRows
        
        vaSpread4.Row = i: vaSpread4.Col = 1
        
        If Trim(vaSpread4.text) = codigo Then
           
           vg_db.BeginTrans
           inderror = 0
           vg_db.Execute "DELETE b_productosing FROM b_productosing WHERE pri_codpro='" & Trim(fpText1(0).text) & "' AND pri_coding='" & codigo & "'"
           vg_db.CommitTrans
           vaSpread4.DeleteRows i, 1
           vaSpread4.MaxRows = vaSpread4.MaxRows - 1
           Exit For
        
        End If
    
    Next i
    
    If vaSpread4.MaxRows > 0 Then vaSpread4.Row = 1: vaSpread4.Col = 1: Limpia 2: modo2 = "": fpText1(5).text = Trim(vaSpread4.text): MoverDatos3 Trim(vaSpread4.text) Else Limpia 2
    
    '-------> Borrar Generico Tecfood
    inderror = 1
    
    vg_db.BeginTrans
    inderror = 0
    vg_db.Execute "DELETE b_productonut FROM b_productonut WHERE pnu_codpro='" & codigo & "'"
    vg_db.Execute "DELETE b_ingrediente FROM b_ingrediente WHERE ing_codigo='" & codigo & "'"
    vg_db.CommitTrans
    If vaSpread4.MaxRows > 0 And Trim(fpText1(0).text) <> "" Then MoverDatos2 Trim(fpText1(0).text)
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False

Case 6 '------> Cancelar
    If MsgBox("Cancelar Operación", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    Combo2(2).Enabled = True
'jp
    If vaSpread4.MaxRows > 0 And modo2 = "A" Then
       
       vaSpread4.Row = 1: vaSpread4.Col = 1: modo2 = "":  MoverDatos3 Trim(vaSpread4.text)
    
    ElseIf vaSpread4.MaxRows > 0 And modo2 = "M" Then
       
       modo2 = "": MoverDatos3 Trim(fpText1(5).text)
    
    ElseIf vaSpread4.MaxRows < 1 Then
       
       fpText1(5).text = "": Limpia 2
    
    End If
    
    Gl_Ac_Botones Me, 8, 1, modo2
    If modo = "" Then Gl_Ac_Botones Me, 1, 1, modo
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(4) = True
    SSTab1.TabEnabled(5) = True
    vaSpread4.Enabled = True

Case 8 '-------> Confirmar - Grabar
    
    If (LimpiaDato(Trim(fpText1(5))) = "" And modo2 <> "A") Or LimpiaDato(Trim(fpText1(6).text)) = "" Or LimpiaDato(Trim(fpLongInteger1(4).text)) = "" Then _
       MsgBox "Debe ingresar información...", vbCritical, MsgTitulo: Exit Sub
    '-------> Fin Validar familia productos
    If modo2 <> "A" Then codigo = LimpiaDato(Trim(fpText1(5).text))
    Nombre = LimpiaDato(Trim(Mid(fpText1(6).text, 1, 50)))
    nomfan = LimpiaDato(Trim(Mid(fpText1(7).text, 1, 50)))
    unimed = fpLongInteger1(4).Value
    pctapr = fpDouble1(1).Value
    pctcoc = fpDouble1(2).Value
    pctnut = fpDouble1(3).Value
    facnut = fpDouble1(4).Value
    indpav = Check1(0).Value
    indgrv = Check1(1).Value
    precos = Val(fpAyDouble1(0).Caption)
    feccos = IIf(Trim(fpAyDate(0).Caption) = "", 0, Val(Mid(fpAyDate(0).Caption, 7, 4) & Mid(fpAyDate(0).Caption, 4, 2) & Mid(fpAyDate(0).Caption, 1, 2)))
    Indppr = Val(fg_codigocbo(Combo2, 2, 1, ""))
    HuellaCarbono = fpHCarbono.Value
    inderror = 2
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS1 = vg_db.Execute("sgpadm_s_ingrediente_V02 7, '" & codigo & "', ''")
    If Not RS1.EOF Then
       
       RS1.Close: Set RS1 = Nothing
       vg_db.Execute "sgpadm_iu_ingrediente_V02 'M1', '" & codigo & "', '" & Nombre & "', '" & nomfan & "', " & unimed & ", " & pctapr & ", " & _
                      "" & pctcoc & ", " & pctnut & ", " & facnut & ", " & indpav & ", " & indgrv & ", " & precos & ", " & _
                      "" & feccos & ", '', '', " & Indppr & ", " & HuellaCarbono & ""
    
    Else
       
       codigo = ""
       
       If RS3.State = 1 Then RS3.Close
       RS3.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
   
       Set RS3 = vg_db.Execute("sgpadm_iu_ingrediente_V02 'A', '', '" & Nombre & "', '" & nomfan & "', " & unimed & ", " & pctapr & ", " & pctcoc & ", " & pctnut & ", " & facnut & ", " & indpav & ", " & indgrv & ", " & precos & ", " & feccos & ", '', '', " & Indppr & ", " & HuellaCarbono & "")
       If Not RS3.EOF Then
          
          codigo = RS3!indice
       
       End If
       RS3.Close: Set RS3 = Nothing
       fpText1(5).text = codigo
       vaSpread4.MaxRows = vaSpread4.MaxRows + 1
       vaSpread4.Row = vaSpread4.MaxRows
       vaSpread4.Col = 1: vaSpread4.text = codigo
       vaSpread4.Col = 2: vaSpread4.text = Nombre
       vaSpread4.Col = 4: vaSpread4.text = IIf(Indppr = "1", "Real", "Propuesta")
       RS1.Close: Set RS1 = Nothing
    
    End If
    '-------> Nutrientes
    vg_db.Execute "DELETE FROM b_productonut WHERE pnu_codpro = '" & codigo & "'"
    '-------> Validar si existe producto generico, borrar aportes nutricionales
    Dim indtec As String
    indtec = ""
    '-------> Fin validar si existe producto generico, borrar aportes nutricionales
    For i = 1 To vaSpread2.MaxRows
        
        vaSpread2.Row = i
        vaSpread2.Col = 1: CodN = Val(vaSpread2.Value)
        vaSpread2.Col = 3: CanN = Val(vaSpread2.Value)
       
        If CanN <> 0 Then
           
           vg_db.Execute "sgpadm_iu_productonut 'A', '" & codigo & "', " & CodN & ", " & CanN & ""
        
        End If
    
    Next i
    
    fpText1(5).ControlType = ControlTypeStatic
    If Not Est Then MsgBox "Datos de Ingrediente fueron grabados con exito...", vbInformation, MsgTitulo
    Me.HelpContextID = "1011000"
    modo2 = "": Gl_Ac_Botones Me, 8, 1, modo2
    If vaSpread4.Enabled = False Then vaSpread4.Enabled = True: SSTab1.TabEnabled(0) = True: SSTab1.TabEnabled(4) = True: SSTab1.TabEnabled(5) = True Else SSTab1.TabEnabled(0) = True: SSTab1.TabEnabled(4) = True: SSTab1.TabEnabled(5) = True
    Me.HelpContextID = vg_OpcM

Case 11, 12 '-------> Buscar o Copiar Aportes Nutricionales
    
    vg_left = fpText1(5).Left + 1920
    vg_codigo = ""
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", IIf(Button.Index = 11, "Ingredientes", "Aporte Nutricionales"), "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Then Exit Sub
    If Button.Index = 11 Then
       
       modo2 = "": MoverDatos3 Trim(vg_codigo)
       If modo = "" Then modo = "M"
       Gl_Ac_Botones Me, 1, 0, modo
    
    Else
       
       If modo2 = "" Then modo2 = "M"
       Gl_Ac_Botones Me, 8, 0, modo2
       '-------> Traer aportes nutriconales
       MoverDatos4 Trim(vg_codigo)
    
    End If
    SSTab1.TabEnabled(0) = False

Case 13 '-------> Imprimir
    
    I_Produc.TraspasoGrilla vaSpread3, "I"
    I_Produc.Show 1

End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then
    If inderror = 0 Then
       vg_db.RollbackTrans
    
    ElseIf inderror = 1 Then
       
       MsgBox "El dato esta asociado a otra tabla... en tecfood", vbCritical, "Error"
       Exit Sub
    
    ElseIf inderror = 2 Then
       
       RS1.Close: Set RS1 = Nothing
    
    End If
    If vaSpread4.MaxRows < 1 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error" Else MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
    Exit Sub

ElseIf Err = -2147217900 Then
    
    If inderror = 0 Then
       
       vg_db.RollbackTrans: Exit Sub
    
    End If

End If
If Err = 3034 Then
   
   If inderror = 0 Then
   
   ElseIf inderror = 2 Then
   
   End If
   
   Exit Sub

End If

If inderror = 0 Then

ElseIf inderror = 2 Then

End If
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim fecven As Long, i As Long, j As Long
Dim codigo As String, Nombre As String, rutpro As String
Dim auxtec As String, indtec As String, StrFam As String, StrFamb As String
Dim CodIng As String, noming As String
Dim fampr1 As Long, fampr2 As Long, fampr3 As Long, unimed As Long
Dim faconv As Double
On Error GoTo Man_Error
'-------> INCLUIR(1)-BORRAR(3)-CANCELAR(6)-CONFIRMAR(8)-BUSCAR(11)-IMPRIMIR(12)
Select Case Button.Index

Case 1 '-------> Nuevo - Limpia
    
    modo3 = "A"
    Limpia 3
    Gl_Ac_Botones Me, 9, 0, modo3

Case 3 '-------> Borra
    
    If MsgBox("Eliminar Dato", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    codigo = LimpiaDato(Trim(fpText1(8).text))
    vg_db.BeginTrans
    For i = 1 To vaSpread5.MaxRows
        vaSpread5.Row = i: vaSpread5.Col = 1
        If Trim(vaSpread5.text) = codigo Then
           vg_db.Execute "DELETE b_productocompra FROM b_productocompra WHERE pco_codpro='" & Trim(fpText1(0).text) & "' AND pco_codigo='" & codigo & "'"
           vaSpread5.DeleteRows i, 1
           vaSpread5.MaxRows = vaSpread5.MaxRows - 1
           Exit For
        End If
    Next i
    If vaSpread5.MaxRows > 0 Then vaSpread5.Row = 1: vaSpread5.Col = 1: Limpia 3: modo3 = "": fpText1(8).text = Trim(vaSpread5.text): MoverDatos5 Trim(vaSpread5.text) Else Limpia 3
    vg_db.CommitTrans

Case 6 '-------> Cancelar
    
    If MsgBox("Cancelar Operación", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If vaSpread5.MaxRows > 0 And modo3 = "A" Then
       vaSpread5.Row = 1: vaSpread5.Col = 1: modo3 = "": MoverDatos5 Trim(vaSpread5.text)
    ElseIf vaSpread5.MaxRows > 0 And modo3 = "M" Then
       modo3 = "": MoverDatos5 Trim(fpText1(8).text): modo3 = ""
    ElseIf vaSpread5.MaxRows < 1 Then
       modo3 = "": fpText1(8).text = "": Limpia 3
    End If
    Gl_Ac_Botones Me, 9, 1, modo3
    vaSpread5.Enabled = True

Case 8 '-------> Confirmar - Grabar
    
    If (LimpiaDato(Trim(fpText1(8))) = "" And modo3 <> "A") Or LimpiaDato(Trim(fpText1(9).text)) = "" Or LimpiaDato(Trim(fpDouble1(7).text)) = "" Then _
       MsgBox "Debe ingresar información...", vbCritical, MsgTitulo: Exit Sub
    '-------> Validar familia productos
    StrFam = fg_BuscaCodArbol(Val(fpLongInteger1(0).Value), "a_tipopro", "tip_codigo")
    
    If Len(StrFam) <> 0 Then
       
       Do While InStr(StrFam, ";") <> 0
          
          StrFamb = Mid(StrFam, 1, InStr(StrFam, ";") - 1)
          StrFam = IIf(Len(StrFam) > InStr(StrFam, ";"), Mid(StrFam, InStr(StrFam, ";") + 1), "")
          fampr1 = Val(Mid(StrFamb, 1, InStr(StrFamb, "&") - 1)): StrFamb = Mid(StrFamb, InStr(StrFamb, "&") + 1)
          fampr2 = Val(Mid(StrFamb, 1, InStr(StrFamb, "&") - 1)): StrFamb = Mid(StrFamb, InStr(StrFamb, "&") + 1)
          If Val(Mid(StrFamb, 1)) = 0 Then MsgBox "Debe seleccionar un nivel superior, en familia producto...", vbCritical, MsgTitulo: Exit Sub
          fampr3 = Val(Mid(StrFamb, 1))
       
       Loop
    
    End If
    '-------> Fin Validar familia productos
    codigo = LimpiaDato(Trim(fpText1(8).text))
    Nombre = LimpiaDato(Trim(fpText1(9).text))
    faconv = fpDouble1(7).Value
    fecven = IIf(Check1(4).Value = 0, 0, Format(fpDateTime1(1).text, "yyyymmdd"))
    vg_db.BeginTrans

    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("SELECT * FROM b_productocompra WHERE pco_codigo='" & codigo & "'")
    
    If Not RS1.EOF Then
        
        vg_db.Execute "UPDATE b_productocompra SET pco_nombre='" & Nombre & "', " & _
                      "pco_undemb=" & faconv & ", pco_fecven=" & fecven & " WHERE pco_codigo='" & codigo & "'"
        
        For i = 1 To vaSpread5.MaxRows
            
            vaSpread5.Row = i: vaSpread5.Col = 1
            If Trim(vaSpread5.text) = codigo Then
               
               For j = 1 To vaSpread4.MaxRows
                  
                  vaSpread4.Row = j: vaSpread4.Col = 1: CodIng = "": CodIng = vaSpread4.text
                  vaSpread4.Col = 2: noming = "": noming = Trim(vaSpread4.text)
                '-------> Traer codigo unidad medida del ingrediente
                  
                  If RS2.State = 1 Then RS2.Close
                  RS2.CursorLocation = adUseClient
                  vg_db.CursorLocation = adUseClient
              
                  Set RS2 = vg_db.Execute("sgpadm_s_ingrediente_V02 5, '" & CodIng & "', ''")
                  If RS2.EOF Then RS2.Close: Set RS1 = Nothing: vg_db.RollbackTrans: MsgBox "No existe código unidad medida ingrediente, proceso cancelado...", vbCritical, MsgTitulo: Exit Sub
                  unimed = 0: unimed = RS2!ing_unimed
                  RS2.Close: Set RS2 = Nothing
                  '-------> Grabar maestro producto tecfood, fam.prod1 - fam.prod2 - fam.prod3 - cod.ing. - nom.ing. - cod.unmed - cod.prod - nom.prod - fac.conversión - opcion - objeto
               
               Next j
               
               vaSpread5.Col = 2: vaSpread5.text = Trim(Nombre)
               vaSpread5.Col = 3: vaSpread5.TypeHAlign = TypeHAlignRight: vaSpread5.text = Format(faconv, fg_Pict(6, 2))
               vaSpread5.Col = 4: vaSpread5.TypeHAlign = TypeHAlignCenter: vaSpread5.text = IIf(Check1(4).Value = 0, "  /  /    ", fpDateTime1(1).text)
            
            End If
        
        Next i
    
    Else
       If vaSpread4.MaxRows < 0 Then Exit Sub
       codigo = LimpiaDato(Trim(fpText1(8)))
       vg_db.Execute "INSERT INTO b_productocompra (pco_codpro, pco_codigo, pco_nombre, pco_undemb, pco_fecven) " & _
                     "VALUES ('" & Trim(fpText1(0).text) & "', '" & codigo & "', '" & Nombre & "', " & faconv & ", " & fecven & ")"
       For i = 1 To vaSpread4.MaxRows
           
           vaSpread4.Row = i: vaSpread4.Col = 1: CodIng = "": CodIng = vaSpread4.text
           vaSpread4.Col = 2: noming = "": noming = Trim(vaSpread4.text)
           '-------> Traer codigo unidad medida del ingrediente
           
           If RS2.State = 1 Then RS2.Close
           RS2.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS2 = vg_db.Execute("sgpadm_s_ingrediente_V02 5, '" & CodIng & "', ''")
           If RS2.EOF Then RS2.Close: Set RS2 = Nothing: vg_db.RollbackTrans: MsgBox "No existe código unidad medida ingrediente, proceso cancelado...", vbCritical, MsgTitulo: Exit Sub
           unimed = 0: unimed = RS2!ing_unimed
           RS2.Close: Set RS2 = Nothing
           '-------> Grabar maestro producto tecfood, fam.prod1 - fam.prod2 - fam.prod3 - cod.ing. - nom.ing. - cod.unmed - cod.prod - nom.prod - fac.conversión - opcion - objeto
       
       Next i
       
       '-------> Fin agregar Productos al maestro tecfood
       fpText1(8).text = codigo
       vaSpread5.MaxRows = vaSpread5.MaxRows + 1
       vaSpread5.Row = vaSpread5.MaxRows
       vaSpread5.Col = 1: vaSpread5.text = codigo
       vaSpread5.Col = 2: vaSpread5.text = Nombre
       vaSpread5.Col = 3: vaSpread5.TypeHAlign = TypeHAlignRight: vaSpread5.text = Format(faconv, fg_Pict(6, 2))
       vaSpread5.Col = 4: vaSpread5.TypeHAlign = TypeHAlignCenter: vaSpread5.text = IIf(Check1(4).Value = 0, "  /  /    ", fpDateTime1(1).text)
    
    End If
    RS1.Close: Set RS1 = Nothing
'tecfood    vg_dbtec.CommitTrans
    vg_db.CommitTrans
    fpText1(8).ControlType = ControlTypeStatic
    If Not Est Then MsgBox "Datos de compras fueron grabados con exito...", vbInformation, MsgTitulo
    modo3 = "": Gl_Ac_Botones Me, 9, 1, modo3
    If vaSpread5.Enabled = False Then vaSpread5.Enabled = True ': SSTab1.TabEnabled(0) = True

Case 50 '11, 12 '-------> Buscar o Copiar Aportes Nutricionales
    
    vg_left = fpText1(5).Left + 1920
    vg_codigo = ""
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", IIf(Button.Index = 11, "Ingredientes", "Aporte Nutricionales"), "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Then Exit Sub
    If Button.Index = 11 Then
       
       modo3 = "": MoverDatos3 Trim(vg_codigo)
       If modo = "" Then modo = "M"
       Gl_Ac_Botones Me, 1, 0, modo
    
    Else
       
       If modo3 = "" Then modo3 = "M"
       Gl_Ac_Botones Me, 9, 0, modo3
       '-------> Traer aportes nutriconales
       MoverDatos4 Trim(vg_codigo)
    
    End If
    SSTab1.TabEnabled(0) = False

Case 13 'Imprimir
    
    I_Produc.TraspasoGrilla vaSpread3, "C"
    I_Produc.Show 1

End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then
    
    vg_db.RollbackTrans
    If vaSpread4.MaxRows < 1 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
    Exit Sub

End If
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim sgpre As Long, codsac As String, nomsac As String, codsgp As String, unisac As String, estpre As Boolean, i As Long, j As Long
Dim faccon As Double
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset

Select Case Button.Index

Case 1 '-------> Nuevo
    vg_left = 8920
    vg_auxcod = "sac"
    B_TabEst.LlenaDatos "b_formatocompras", "foc_", "Formato Compras", "ForCom"
    B_TabEst.Show 1
    Me.Refresh
    vg_auxcod = ""
    If Trim(vg_codigo) = "" Then Exit Sub
    estexi = False
    
    For i = 1 To vaSpread6.MaxRows
        
        vaSpread6.Row = i
        vaSpread6.Col = 2
        If vg_codigo = vaSpread6.text Then estexi = True: vaSpread6.RowHidden = False: Exit For
    
    Next i
    
    If Not estexi Then
       
       vaSpread6.MaxRows = vaSpread6.MaxRows + 1
       vaSpread6.Row = vaSpread6.MaxRows
       vaSpread6.Col = 1: vaSpread6.CellType = CellTypeStaticText
       vaSpread6.Col = 2: vaSpread6.text = vg_codigo
       vaSpread6.Col = 3: vaSpread6.text = vg_nombre
       vaSpread6.Col = 4: vaSpread6.text = vg_ames
    
    End If
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 11, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(5) = False

Case 3 '-------> Eliminar
    
    If vaSpread6.MaxRows < 1 Or vaSpread6.ActiveRow < 1 Then Exit Sub
    vaSpread6.Row = vaSpread6.ActiveRow
    vaSpread6.Col = vaSpread6.ActiveCol
    vaSpread6.CellType = CellTypeStaticText
    vaSpread6.RowHidden = True
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 11, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(5) = False

Case 5 '-------> Fijar Vinculo
    
    If vaSpread6.MaxRows < 1 Or vaSpread6.ActiveRow < 1 Then Exit Sub
    IRow = vaSpread6.ActiveRow
    vaSpread6.Visible = False
    For i = 1 To vaSpread6.MaxRows
        
        vaSpread6.Row = i
        vaSpread6.Col = 1
        
        If vaSpread6.Row = IRow Then
           
           vaSpread7.Row = 1: vaSpread7.Col = 1
           vaSpread6.Col = 1
           vaSpread6.CellType = CellTypePicture
           vaSpread6.TypePictCenter = True
           vaSpread6.TypePictMaintainScale = True
           vaSpread6.TypePictStretch = True
           vaSpread6.Row = IRow
           vaSpread6.TypePictPicture = vaSpread7.TypePictPicture
           vaSpread6.text = vaSpread7.text
        
        Else
           
           vaSpread6.CellType = CellTypeStaticText
        
        End If
    
    Next i
    vaSpread6.Visible = True
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 11, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(5) = False

Case 7 '-------> Cancelar
    
    Est = True
    vaSpread6.Visible = False
    vaSpread6.MaxRows = 0
    vaSpread6.Row = -1: vaSpread6.Col = -1
    vaSpread6.BackColor = Shape1(0).FillColor
    codigo = LimpiaDato(Trim(fpText1(0).text))
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("sgpadm_Sel_productos 17, '" & codigo & "', 'sac', '" & vg_NUsr & "'")
    Do While Not RS1.EOF
       
       Frame11.Caption = IIf(IsNull(RS1!pro_codigo), "", RS1!pro_codigo) & " - " & IIf(IsNull(RS1!pro_nombre), "", Trim(RS1!pro_nombre)) & " - Cta. Con.(" & IIf(IsNull(RS1!pro_ctacon), "", Trim(RS1!pro_ctacon)) & ")"
       vaSpread6.MaxRows = vaSpread6.MaxRows + 1
       vaSpread6.Row = vaSpread6.MaxRows
       vaSpread6.Col = 1
       
       If RS1!fcs_sgppre = 0 Then
          
          vaSpread6.CellType = CellTypeStaticText
       
       Else
          
          vaSpread7.Row = 1: vaSpread7.Col = 1
          vaSpread6.CellType = CellTypePicture
          vaSpread6.TypePictCenter = True
          vaSpread6.TypePictMaintainScale = True
          vaSpread6.TypePictStretch = True
          vaSpread6.TypePictPicture = vaSpread7.TypePictPicture
       
       End If
       
       vaSpread6.Col = 2: vaSpread6.text = IIf(IsNull(RS1(0)), "", RS1(0)) 'IIf(IsNull(RS1!foc_codsac), "", RS1!foc_codsac)
       vaSpread6.Col = 3: vaSpread6.text = IIf(IsNull(RS1(1)), "", RS1(1)) 'IIf(IsNull(RS1!foc_nomsac), "", RS1!foc_nomsac)
       vaSpread6.Col = 4: vaSpread6.text = IIf(IsNull(RS1(2)), "", RS1(2)) 'IIf(IsNull(RS1!foc_unisac), "", RS1!foc_unisac)
       vaSpread6.Col = 5: vaSpread6.text = IIf(IsNull(RS1(5)), "", RS1(5)) 'IIf(IsNull(RS1!fcs_codsgp), "", RS1!fcs_codsgp)
       vaSpread6.Col = 6: vaSpread6.text = IIf(IsNull(RS1(8)), "", RS1(8)) 'IIf(IsNull(RS1!pro_nombre), "", RS1!pro_nombre)
       vaSpread6.Col = 7: vaSpread6.text = IIf(IsNull(RS1(11)), 0, RS1(11)) 'IIf(IsNull(RS1!foc_faccon), 0, RS1!foc_faccon)
       RS1.MoveNext
    
    Loop
    
    RS1.Close: Set RS1 = Nothing
    vaSpread6.Visible = True
    Est = False
    modo = ""
    Gl_Ac_Botones Me, 11, 1, modo
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(5) = True

Case 9 '-------> Actualizar
    
    codigo = LimpiaDato(Trim(fpText1(0).text))
    '-------> Validar si existe un formato de compra prefijado
    estpre = False
    j = 0
    For i = 1 To vaSpread6.MaxRows
        
        vaSpread6.Row = i
        vaSpread6.Col = 1
        If vaSpread6.RowHidden = False And vaSpread6.CellType = CellTypePicture Then estpre = True
        If vaSpread6.RowHidden = False Then j = j + 1
    
    Next i

    vg_db.Execute "DELETE b_formatocomprassgp FROM b_formatocomprassgp WHERE fcs_codsgp = '" & codigo & "'"
    For i = 1 To vaSpread6.MaxRows
        
        vaSpread6.Row = i
        
        If vaSpread6.RowHidden = False Then
           
           vaSpread6.Col = 1: sgppre = IIf(vaSpread6.CellType = CellTypePicture, 1, 0)
           vaSpread6.Col = 2: codsac = Trim(vaSpread6.text)
           vaSpread6.Col = 3: nomsac = Trim(vaSpread6.text)
           vaSpread6.Col = 4: unisac = Trim(vaSpread6.text)
           vaSpread6.Col = 7: faccon = Val(vaSpread6.text)
           Set RS1 = vg_db.Execute("sgpadm_Sel_productos 27, '" & codigo & "', '" & codsac & "', '" & vg_NUsr & "'")
           If RS1.EOF Then
              
              vg_db.Execute "INSERT INTO b_formatocomprassgp (fcs_codsac, fcs_codsgp, fcs_sgppre) VALUES ('" & codsac & "', '" & codigo & "', " & IIf(Not estpre And j = 1, 1, sgppre) & ")"
           
           Else
              
              vg_db.Execute "UPDATE b_formatocomprassgp SET fcs_codsgp = '" & codigo & "', fcs_sgppre = " & IIf(Not estpre And j = 1, 1, sgppre) & "  WHERE fcs_codsac = '" & codsac & "'"
           
           End If
           RS1.Close: Set RS1 = Nothing
           vg_db.Execute "UPDATE b_formatocompras SET foc_faccon = " & faccon & " WHERE foc_codsac = '" & codsac & "'"
        
        End If
    
    Next i
    modo = ""
    Gl_Ac_Botones Me, 11, 1, modo
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(5) = True

Case 12 '-------> Imprimir
    
    vg_opimp = 1
    I_ForCom.Show 1
    vg_opimp = 0

End Select

End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim sgpre As Long, codsac As String, nomsac As String, codsgp As String, unisac As String, estpre As Boolean, i As Long, j As Long
Dim faccon As Double
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset

Select Case Button.Index

Case 1 '-------> Nuevo
    
    vg_left = 8920
    vg_auxcod = "sap"
    B_TabEst.LlenaDatos "b_formatocompras", "foc_", "Formato Compras", "ForCom"
    B_TabEst.Show 1
    Me.Refresh

    If Trim(vg_codigo) = "" Then vg_auxcod = "": Exit Sub
    estexi = False
    
    For i = 1 To vaSpread6.MaxRows
        
        vaSpread10.Row = i
        vaSpread10.Col = 2
        If vg_codigo = vaSpread10.text Then estexi = True: vaSpread10.RowHidden = False: Exit For
    
    Next i
    
    If Not estexi Then
       
       vaSpread10.MaxRows = vaSpread10.MaxRows + 1
       vaSpread10.Row = vaSpread10.MaxRows
       vaSpread10.Col = 1: vaSpread10.CellType = CellTypeStaticText
       vaSpread10.Col = 2: vaSpread10.text = vg_codigo
       vaSpread10.Col = 3: vaSpread10.text = vg_nombre
       vaSpread10.Col = 4: vaSpread10.text = vg_ames
       vaSpread10.Col = 7: vaSpread10.text = vg_auxcod
    
    End If
    
    If modo = "" Then modo = "M"
    vg_auxcod = ""
    Gl_Ac_Botones Me, 15, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(4) = False

Case 3 '-------> Eliminar
    
    If vaSpread10.MaxRows < 1 Or vaSpread10.ActiveRow < 1 Then Exit Sub
    vaSpread10.Row = vaSpread10.ActiveRow
    vaSpread10.Col = vaSpread10.ActiveCol
    vaSpread10.CellType = CellTypeStaticText
    vaSpread10.RowHidden = True
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 15, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(4) = False

Case 5 '-------> Fijar Vinculo
    
    If vaSpread10.MaxRows < 1 Or vaSpread10.ActiveRow < 1 Then Exit Sub
    IRow = vaSpread10.ActiveRow
    vaSpread10.Visible = False
    
    For i = 1 To vaSpread10.MaxRows
        
        vaSpread10.Row = i
        vaSpread10.Col = 1
        
        If vaSpread10.Row = IRow Then
           
           vaSpread9.Row = 1: vaSpread9.Col = 1
           vaSpread10.Col = 1
           vaSpread10.CellType = CellTypePicture
           vaSpread10.TypePictCenter = True
           vaSpread10.TypePictMaintainScale = True
           vaSpread10.TypePictStretch = True
           vaSpread10.Row = IRow
           vaSpread10.TypePictPicture = vaSpread9.TypePictPicture
           vaSpread10.text = vaSpread9.text
        
        Else
           
           vaSpread10.CellType = CellTypeStaticText
        
        End If
    
    Next i
    
    vaSpread10.Visible = True
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 15, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(4) = False

Case 7 '-------> Cancelar
    
    Est = True
    vaSpread10.Visible = False
    vaSpread10.MaxRows = 0
    vaSpread10.Row = -1: vaSpread10.Col = -1
    vaSpread10.BackColor = Shape1(0).FillColor
    codigo = LimpiaDato(Trim(fpText1(0).text))
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS1 = vg_db.Execute("sgpadm_Sel_productos 17, '" & codigo & "', 'sap', '" & vg_NUsr & "'")
    Do While Not RS1.EOF
       
       Frame14.Caption = IIf(IsNull(RS1!pro_codigo), "", RS1!pro_codigo) & " - " & IIf(IsNull(RS1!pro_nombre), "", Trim(RS1!pro_nombre)) & " - Cta. Con.(" & IIf(IsNull(RS1!pro_ctacon), "", Trim(RS1!pro_ctacon)) & ")"
       vaSpread10.MaxRows = vaSpread10.MaxRows + 1
       vaSpread10.Row = vaSpread10.MaxRows
       vaSpread10.Col = 1
       
       If RS1(6) = 0 Then
          
          vaSpread10.CellType = CellTypeStaticText
       
       Else
          
          vaSpread9.Row = 1: vaSpread9.Col = 1
          vaSpread10.CellType = CellTypePicture
          vaSpread10.TypePictCenter = True
          vaSpread10.TypePictMaintainScale = True
          vaSpread10.TypePictStretch = True
          vaSpread10.TypePictPicture = vaSpread9.TypePictPicture
       
       End If
       
       vaSpread10.Col = 2: vaSpread10.text = IIf(IsNull(RS1(0)), "", RS1(0)) 'IIf(IsNull(RS1!foc_codsac), "", RS1!foc_codsac)
       vaSpread10.Col = 3: vaSpread10.text = IIf(IsNull(RS1(1)), "", RS1(1)) 'IIf(IsNull(RS1!foc_nomsac), "", RS1!foc_nomsac)
       vaSpread10.Col = 4: vaSpread10.text = IIf(IsNull(RS1(2)), "", RS1(2)) 'IIf(IsNull(RS1!foc_unisac), "", RS1!foc_unisac)
       vaSpread10.Col = 5: vaSpread10.text = IIf(IsNull(RS1(5)), "", RS1(5)) 'IIf(IsNull(RS1!fcs_codsgp), "", RS1!fcs_codsgp)
       vaSpread10.Col = 6: vaSpread10.text = IIf(IsNull(RS1(8)), "", RS1(8)) 'IIf(IsNull(RS1!pro_nombre), "", RS1!pro_nombre)
       vaSpread10.Col = 7: vaSpread10.text = IIf(IsNull(RS1(13)), "", RS1(13)) 'IIf(IsNull(RS1!fcs_ctacon), "", RS1!fcs_ctacon)
       vaSpread10.Col = 8: vaSpread10.text = IIf(IsNull(RS1(11)), 0, RS1(11)) 'IIf(IsNull(RS1!foc_faccon), 0, RS1!foc_faccon)
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing
    vaSpread10.Visible = True
    Est = False
    modo = ""
    Gl_Ac_Botones Me, 15, 1, modo
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(4) = True

Case 9 '-------> Actualizar
    
    codigo = LimpiaDato(Trim(fpText1(0).text))
    '-------> Validar si existe un formato de compra prefijado
    estpre = False
    j = 0
    
    For i = 1 To vaSpread10.MaxRows
        
        vaSpread10.Row = i
        vaSpread10.Col = 1
        If vaSpread10.RowHidden = False And vaSpread10.CellType = CellTypePicture Then estpre = True
        If vaSpread10.RowHidden = False Then j = j + 1
    
    Next i

    vg_db.Execute "DELETE b_formatocompras_sap_sgp FROM b_formatocompras_sap_sgp WHERE fss_CodSgp = '" & codigo & "'"
    
    For i = 1 To vaSpread10.MaxRows
        
        vaSpread10.Row = i
        
        If vaSpread10.RowHidden = False Then
           
           vaSpread10.Col = 1: sgppre = IIf(vaSpread10.CellType = CellTypePicture, 1, 0)
           vaSpread10.Col = 2: codsac = Trim(vaSpread10.text)
           vaSpread10.Col = 3: nomsac = Trim(vaSpread10.text)
           vaSpread10.Col = 4: unisac = Trim(vaSpread10.text)
           vaSpread10.Col = 8: faccon = Val(vaSpread10.text)
           
           If RS1.State = 1 Then RS1.Close
           RS1.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
       
           Set RS1 = vg_db.Execute("sgpadm_s_formatocomprassap 3, '" & codigo & "', '" & codsac & "'")
           
           If RS1.EOF Then
              
              vg_db.Execute "INSERT INTO b_formatocompras_sap_sgp (fss_CodMaterial, fss_CodSgp, fss_SgpPre) VALUES ('" & codsac & "', '" & codigo & "', " & IIf(Not estpre And j = 1, 1, sgppre) & ")"
           
           Else
              
              vg_db.Execute "UPDATE b_formatocompras_sap_sgp SET fss_CodSgp = '" & codigo & "', fss_SgpPre = " & IIf(Not estpre And j = 1, 1, sgppre) & "  WHERE fss_CodMaterial = '" & codsac & "'"
           
           End If
           RS1.Close: Set RS1 = Nothing
           vg_db.Execute "UPDATE b_formatocompras_sap SET fcs_faccon = " & faccon & " WHERE fcs_CodMaterial = '" & codsac & "'"
        
        End If
    
    Next i
    modo = ""
    Gl_Ac_Botones Me, 15, 1, modo
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(4) = True

Case 12 '-------> Imprimir
    
    vg_opimp = 1
    I_ForCom.Show 1
    vg_opimp = 0

End Select

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows > 0 Then EstDet = True: modo = "": MoverDatos
End Sub

Private Sub vaSpread1_ScriptBeforeUserSort(ByVal Col As Long, ByVal State As Long, DefaultAction As Variant)

a = State
b = defaulaction

End Sub

Private Sub vaSpread10_EditChange(ByVal Col As Long, ByVal Row As Long)

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 15, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(4) = False

End Sub

Private Sub vaSpread2_EditChange(ByVal Col As Long, ByVal Row As Long)

If Est Then Exit Sub
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(5) = False
If modo2 = "" Then modo2 = "M"
Gl_Ac_Botones Me, 8, 0, modo2

End Sub

Private Sub vaSpread3_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If Est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo: SSTab1.TabEnabled(0) = False: SSTab1.TabEnabled(4) = False: SSTab1.TabEnabled(5) = False

End Sub

Private Sub vaSpread4_Click(ByVal Col As Long, ByVal Row As Long)

If vaSpread4.MaxRows < 1 Then Exit Sub
vaSpread4.Row = vaSpread4.ActiveRow: vaSpread4.Col = 1
MoverDatos3 Trim(vaSpread4.text)

End Sub

Private Sub vaSpread5_Click(ByVal Col As Long, ByVal Row As Long)

If vaSpread5.MaxRows < 1 Then Exit Sub
vaSpread5.Row = vaSpread5.ActiveRow: vaSpread5.Col = 1
MoverDatos5 Trim(vaSpread5.text)

End Sub

Sub BloquearOpSistema()

'-------> bloquear opciones del sistema si el pasi = chile
Label3(24).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Label3(25).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Image1(5).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Image1(6).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
fpLongInteger1(5).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
fpLongInteger1(6).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Check1(5).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)

End Sub

Private Sub vaSpread6_EditChange(ByVal Col As Long, ByVal Row As Long)

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 11, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(5) = False

End Sub
