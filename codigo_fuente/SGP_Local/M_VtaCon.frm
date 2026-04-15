VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_VtaCon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venta Servicio Contado"
   ClientHeight    =   8085
   ClientLeft      =   2895
   ClientTop       =   2115
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   5415
      Left            =   30
      TabIndex        =   10
      Top             =   2520
      Width           =   10755
      Begin TabDlg.SSTab SSTab1 
         Height          =   4695
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8281
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabMaxWidth     =   4
         TabCaption(0)   =   "Venta Servicio"
         TabPicture(0)   =   "M_VtaCon.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1(6)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(5)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Shape1(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label5(0)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Shape1(0)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label5(1)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "vaSpread1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Detalle Centro Costo"
         TabPicture(1)   =   "M_VtaCon.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label1(10)"
         Tab(1).Control(1)=   "vaSpread2"
         Tab(1).ControlCount=   2
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3795
            Left            =   120
            TabIndex        =   27
            Top             =   540
            Width           =   10200
            _Version        =   393216
            _ExtentX        =   17992
            _ExtentY        =   6694
            _StockProps     =   64
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
            MaxCols         =   7
            MaxRows         =   12
            ProcessTab      =   -1  'True
            ScrollBars      =   0
            SpreadDesigner  =   "M_VtaCon.frx":0038
            UserResize      =   1
         End
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   3630
            Left            =   -73800
            TabIndex        =   32
            Top             =   840
            Width           =   8145
            _Version        =   393216
            _ExtentX        =   14367
            _ExtentY        =   6403
            _StockProps     =   64
            AutoClipboard   =   0   'False
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
            MaxCols         =   4
            ProcessTab      =   -1  'True
            ScrollBars      =   2
            SpreadDesigner  =   "M_VtaCon.frx":07D9
         End
         Begin VB.Label Label1 
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
            Index           =   10
            Left            =   -73740
            TabIndex        =   33
            Top             =   480
            Width           =   5280
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Dias Bloqueados"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   31
            Top             =   4410
            Width           =   1200
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H008484FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   4440
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Dias Habilitados"
            Height          =   195
            Index           =   0
            Left            =   2325
            TabIndex        =   30
            Top             =   4410
            Width           =   1140
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H80000018&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   1965
            Top             =   4440
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Mes : "
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
            Left            =   5730
            TabIndex        =   29
            Top             =   4410
            Width           =   1035
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0"
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
            Left            =   8280
            TabIndex        =   28
            Top             =   4410
            Width           =   120
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Diciembre 2004"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   1800
         TabIndex        =   17
         Top             =   320
         Width           =   4275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2145
      Left            =   1110
      TabIndex        =   5
      Top             =   390
      Width           =   8595
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "M_VtaCon.frx":20D2
         Left            =   1380
         List            =   "M_VtaCon.frx":20D4
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1600
         Width           =   2670
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   0
         Top             =   560
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   2
         Left            =   1380
         TabIndex        =   1
         Top             =   900
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
         NoSpecialKeys   =   2
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   6260
         TabIndex        =   4
         Top             =   1605
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1658
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
         ButtonStyle     =   1
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
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   "08/2010"
         DateCalcMethod  =   4
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/yyyy"
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
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   0
         IncHour         =   0
         IncMinute       =   0
         IncSecond       =   0
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   2
         Top             =   1250
         Width           =   1340
         _Version        =   196608
         _ExtentX        =   2364
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
         Index           =   1
         Left            =   1380
         TabIndex        =   23
         Top             =   210
         Width           =   1340
         _Version        =   196608
         _ExtentX        =   2364
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
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   -1  'True
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
         MaxLength       =   10
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
         ButtonAlign     =   1
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2670
         Picture         =   "M_VtaCon.frx":20D6
         Top             =   120
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   3180
         TabIndex        =   24
         Top             =   210
         Width           =   3975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contrato"
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
         Left            =   240
         TabIndex        =   22
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opcional"
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
         Left            =   7335
         TabIndex        =   21
         Top             =   1320
         Width           =   765
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2670
         Picture         =   "M_VtaCon.frx":23E0
         Top             =   1170
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3180
         TabIndex        =   19
         Top             =   1250
         Width           =   3975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Left            =   240
         TabIndex        =   18
         Top             =   1300
         Width           =   600
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   15
         Top             =   1670
         Width           =   2640
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3180
         TabIndex        =   12
         Top             =   900
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3180
         TabIndex        =   11
         Top             =   555
         Width           =   3975
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2670
         Picture         =   "M_VtaCon.frx":26EA
         Top             =   820
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2670
         Picture         =   "M_VtaCon.frx":29F4
         Top             =   460
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   980
         Width           =   705
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   640
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
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
         Left            =   5190
         TabIndex        =   6
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3225
         TabIndex        =   13
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3225
         TabIndex        =   14
         Top             =   945
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3225
         TabIndex        =   20
         Top             =   1290
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3225
         TabIndex        =   25
         Top             =   255
         Width           =   3975
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_VtaCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim Msgtitulo As String, modo As String
Dim est As Boolean, mesblo As Boolean, valcli As Boolean, indrow As Long, indcol As Long, numdia As String
Dim vecmon() As Double

Private Sub Combo2_Click(Index As Integer)
If Combo2(0).ListIndex = -1 Or est Then Exit Sub
Mover_Datos
End Sub

Private Sub Form_Activate()
fg_descarga
TraerFechaCierre
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 8565
Me.Width = 10920
Msgtitulo = "Venta Servicio Contado"
fg_centra Me
est = True: mesblo = False: modo = "": valcli = False: numdia = ""
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 3, modo
Combo2(0).Clear
Combo2(0).AddItem "CONTADO" & Space(150) & "(0)"
Combo2(0).AddItem "CHEQUE" & Space(150) & "(1)"
Combo2(0).AddItem "CHEQUE RESTAURANT" & Space(150) & "(2)"
Combo2(0).AddItem "TARJETA CREDITO" & Space(150) & "(3)"
Combo2(0).AddItem "VALE" & Space(150) & "(4)"
Combo2(0).ListIndex = -1
fpText(1).Enabled = ModCasino
Image1(3).Enabled = ModCasino
fpText(1).text = MuestraCasino(1)
fpayuda(4).Caption = MuestraCasino(2)
fpDateTime1.text = Format(Date, "mm/yyyy")
MoverFormato False
SSTab1.TabVisible(1) = False: SSTab1.Tab = 0
est = False
End Sub

Private Sub fpDateTime1_Change()
If est Then Exit Sub
If Trim(fpDateTime1.text) = "" Then Exit Sub
Mover_Datos
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

If est Then Exit Sub
Select Case Index
Case 1
    RS.Open "SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).text) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    Mover_Datos
Case 2
    RS.Open "SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(2).Value) & " AND ser_activo = '1'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "": Exit Sub
    fpayuda(2).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    Mover_Datos
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Trim(fpLongInteger1(Index).text) = "" Or Val(fpLongInteger1(Index).Value) < 1 Then fpLongInteger1(Index).text = ""
SendKeys "{Tab}"
End Sub

Private Sub fpText_Change(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    fpayuda(Index).Caption = ""
Case 1
    RS1.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & fpText(1).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(4).Caption = "": Exit Sub
    fpayuda(4).Caption = Trim(RS1!cli_nombre)
    RS1.Close: Set RS1 = Nothing
    Mover_Datos
End Select
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_GotFocus(Index As Integer)
Select Case Index
Case 0
    If Trim(fpText(0).text) = "" Or vg_Dig = "N" Then Exit Sub
    fpText(0).text = fg_DespintaRut(fpText(0).text)
    fpText(0).text = Mid(fpText(0).text, 1, Len(Trim(fpText(0).text)) - 1)
End Select
End Sub

Private Sub fpText_LostFocus(Index As Integer)
Select Case Index
Case 0
    If fpText(0).text = "" Then fpayuda(3).Caption = "": Exit Sub
    fpText(0).text = fg_RutDig(Trim(fpText(0).text))
    RS1.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & Trim(fpText(0).text) & "' AND cli_tipo = 1 AND cli_activo = '1'", vg_db, adOpenStatic
    If Not RS1.EOF Then
        fpText(0).text = fg_PintaRut(fpText(0).text)
        fpayuda(3).Caption = RS1!cli_nombre
        RS1.Close: Set RS1 = Nothing
        Mover_Datos
    Else
        RS1.Close: Set RS1 = Nothing: MsgBox "Cliente no existe...", vbCritical, Msgtitulo
        fpText(0).text = "": fpayuda(3).Caption = ""
        Exit Sub
    End If
End Select
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Clientes", "Cliente"
    B_TabEst.Show 1, Me
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(0).text = fg_PintaRut(vg_codigo)
    fpayuda(3).Caption = vg_nombre
    Mover_Datos
'    If Combo2(0).Enabled Then Combo2(0).SetFocus
Case 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    fpLongInteger1(2).SetFocus
Case 2
    vg_left = fpayuda(2).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Ser"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre
    If Combo2(0).Enabled Then fpText(0).SetFocus
Case 3
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(1).text = vg_codigo
    fpayuda(4).Caption = vg_nombre
    fpLongInteger1(1).SetFocus
End Select
End Sub

Sub Mover_Datos()
SSTab1.TabVisible(1) = False
If Trim(fpText(1).text) = "" Or Combo2(0).ListIndex = -1 Or Trim(fpDateTime1.text) = "" Or est Then Exit Sub
fg_carga ""
valcli = False
RS.Open "SELECT DISTINCT clc_codcli FROM b_clientecencos WHERE clc_codcli = '" & Trim(fg_DespintaRut(fpText(0).text)) & "'", vg_db, adOpenStatic
If Not RS.EOF Then valcli = True
RS.Close: Set RS = Nothing

ReDim Preserve vecmon(fg_mes(Mid(fg_pone_cero(fpDateTime1.text, 7), 1, 2) & Mid(fg_pone_cero(fpDateTime1.text, 7), 4, 4)))
For i = 1 To UBound(vecmon)
    vecmon(i) = 0
Next i
Label1(6).Caption = 0
If valcli Then
   If vg_tipbase = "1" Then
      RS.Open "SELECT a.vtc_totmon, a.vtc_fecvta FROM b_ventacontado a, b_ventacontadodet b " & _
              "WHERE a.vtc_codigo = b.vtd_codigo " & _
              "AND   a.vtc_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
              "AND   a.vtc_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
              "AND   a.vtc_codser = " & Val(fpLongInteger1(2).Value) & " " & _
              "AND   a.vtc_forpag = " & Val(fg_codigocbo(Combo2, 0, 1, "")) & " " & _
              "AND val(mid(a.vtc_fecvta,1,6)) = " & Val(Format(fpDateTime1.text, "yyyymm")) & "", vg_db, adOpenStatic
   Else
      RS.Open "SELECT a.vtc_totmon, a.vtc_fecvta FROM b_ventacontado a, b_ventacontadodet b " & _
              "WHERE a.vtc_codigo = b.vtd_codigo " & _
              "AND   a.vtc_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
              "AND   a.vtc_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
              "AND   a.vtc_codser = " & Val(fpLongInteger1(2).Value) & " " & _
              "AND   a.vtc_forpag = " & Val(fg_codigocbo(Combo2, 0, 1, "")) & " " & _
              "AND   convert(int,substring(convert(varchar(8),a.vtc_fecvta),1,6)) = " & Val(Format(fpDateTime1.text, "yyyymm")) & "", vg_db, adOpenStatic
   End If
Else
   If vg_tipbase = "1" Then
      RS.Open "SELECT vtc_totmon, vtc_fecvta FROM b_ventacontado " & _
              "WHERE vtc_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
              "AND   vtc_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
              "AND   vtc_codser = " & Val(fpLongInteger1(2).Value) & " " & _
              "AND   vtc_forpag = " & Val(fg_codigocbo(Combo2, 0, 1, "")) & " " & _
              "AND   val(mid(vtc_fecvta,1,6)) = " & Val(Format(fpDateTime1.text, "yyyymm")) & " AND vtc_opccli = '0'", vg_db, adOpenStatic
   Else
      RS.Open "SELECT vtc_totmon, vtc_fecvta FROM b_ventacontado " & _
              "WHERE vtc_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
              "AND   vtc_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
              "AND   vtc_codser = " & Val(fpLongInteger1(2).Value) & " " & _
              "AND   vtc_forpag = " & Val(fg_codigocbo(Combo2, 0, 1, "")) & " " & _
              "AND   convert(int,substring(convert(varchar(8),vtc_fecvta),1,6)) = " & Val(Format(fpDateTime1.text, "yyyymm")) & " AND vtc_opccli = '0'", vg_db, adOpenStatic
   End If
End If
If Not RS.EOF Then
   Do While Not RS.EOF
      vecmon(Val(Mid(RS!vtc_fecvta, 7, 2))) = RS!vtc_totmon
      Label1(6).Caption = Format((Label1(6).Caption + RS!vtc_totmon), fg_Pict(11, 2))
      RS.MoveNext
   Loop
   MoverFormato True
   Gl_Ac_Botones Me, 1, IIf(mesblo = False, 4, 6), modo
'   If valcli Then SSTab1.TabEnabled(1) = True
Else
   RS.Close: Set RS = Nothing
   MoverFormato False
   SumarTotales
   fg_descarga
   If mesblo = True Then Exit Sub
   If Trim(fpText(1).text) = "" Or fpLongInteger1(1).Value = "" Or fpLongInteger1(2).Value = "" Or fpDateTime1.text = "" Then
      Gl_Ac_Botones Me, 1, 3, modo
   Else
      Gl_Ac_Botones Me, 1, IIf(mesblo = False, 2, 3), modo
   End If
   Exit Sub
End If
RS.Close: Set RS = Nothing
vaSpread1.Row = 2: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, 2
If vaSpread1.Enabled = True Then vaSpread1.SetFocus
fg_descarga
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
Case 0
    If indcol = 0 Or indrow = 0 Then Exit Sub
    vaSpread1.SetActiveCell indcol, indrow ': vaSpread1.SetFocus
Case 1
'    If indcol = 0 Or indrow = 0 Then SSTab1.TabVisible(1) = False: Exit Sub
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim totmon As Double
Dim diavta As Long, codigo As Long, X As Long
Dim graenc As Boolean
On Error GoTo Man_Error
Select Case Button.Index
Case 1 '------- Incluir
    If Trim(fpayuda(4).Caption) = "" Or Trim(fpayuda(1).Caption) = "" Or Trim(fpayuda(2).Caption) = "" Or fpDateTime1.text = "" Or Combo2(0).ListIndex = -1 Then MsgBox "Falta información en el encabezado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    modo = "A": Gl_Ac_Botones Me, 1, 0, modo
    Frame1.Enabled = False
    MoverFormato True
    For i = 1 To vaSpread1.MaxRows Step 2
        vaSpread1.Row = i
        For j = 1 To vaSpread1.MaxCols
            vaSpread1.Col = j
            If Trim(vaSpread1.text) <> "" Then vaSpread1.Row = i + 1: vaSpread1.SetActiveCell j, i + 1: i = vaSpread1.MaxRows: Exit For
        Next j
    Next i
    SumarTotales
'    If valcli Then SSTab1.TabVisible(1) = True
    vaSpread1.SetActiveCell 2, 2: vaSpread1.SetFocus
Case 3 '------- Alterar
    If Trim(fpayuda(4).Caption) = "" Or Trim(fpayuda(1).Caption) = "" Or Trim(fpayuda(2).Caption) = "" Or fpDateTime1.text = "" Or Combo2(0).ListIndex = -1 Then MsgBox "Falta información en el encabezado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Frame1.Enabled = False
    For i = 1 To vaSpread1.MaxRows Step 2
        vaSpread1.Row = i
        For j = 1 To vaSpread1.MaxCols
            vaSpread1.Col = j
            If Trim(vaSpread1.text) <> "" Then vaSpread1.Row = i + 1: vaSpread1.SetActiveCell j, i + 1: i = vaSpread1.MaxRows: Exit For
        Next j
    Next i
    SumarTotales
    modo = "M": Gl_Ac_Botones Me, 1, 0, modo
    Me.Refresh
Case 5 '------- Borrar
    If vaSpread1.ActiveRow < 1 Then MsgBox "No existe información a borrar...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina Documento...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    If vg_tipbase = "1" Then
       vg_db.Execute "DELETE b_ventacontadodet.* FROM b_ventacontado INNER JOIN b_ventacontadodet ON b_ventacontado.vtc_codigo = b_ventacontadodet.vtd_codigo " & _
                     "WHERE b_ventacontado.vtc_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' AND b_ventacontado.vtc_codreg = " & Val(fpLongInteger1(1).Value) & " AND b_ventacontado.vtc_codser = " & Val(fpLongInteger1(2).Value) & " AND val(mid(b_ventacontado.vtc_fecvta,1,6)) = " & Val(Format(fpDateTime1.text, "yyyymm")) & " AND b_ventacontado.vtc_forpag = " & Val(fg_codigocbo(Combo2, 0, 1, "")) & ""
    Else
       vg_db.Execute "DELETE b_ventacontadodet FROM b_ventacontado, b_ventacontadodet WHERE b_ventacontado.vtc_codigo = b_ventacontadodet.vtd_codigo " & _
                     "AND b_ventacontado.vtc_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' AND b_ventacontado.vtc_codreg = " & Val(fpLongInteger1(1).Value) & " AND b_ventacontado.vtc_codser = " & Val(fpLongInteger1(2).Value) & " AND convert(int,substring(convert(varchar(8),b_ventacontado.vtc_fecvta),1,6)) = " & Val(Format(fpDateTime1.text, "yyyymm")) & " AND b_ventacontado.vtc_forpag = " & Val(fg_codigocbo(Combo2, 0, 1, "")) & ""
    End If
    If valcli = False Then
       If vg_tipbase = "1" Then
          vg_db.Execute "DELETE b_ventacontado FROM b_ventacontado " & _
                        "WHERE  vtc_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' AND vtc_codreg = " & Val(fpLongInteger1(1).Value) & " AND vtc_codser = " & Val(fpLongInteger1(2).Value) & " AND val(mid(vtc_fecvta,1,6))=" & Val(Format(fpDateTime1.text, "yyyymm")) & " AND vtc_forpag = " & Val(fg_codigocbo(Combo2, 0, 1, "")) & " AND vtc_opccli = '" & IIf(valcli = True, "1", "0") & "'"
       Else
          vg_db.Execute "DELETE b_ventacontado FROM b_ventacontado " & _
                        "WHERE  vtc_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' AND vtc_codreg = " & Val(fpLongInteger1(1).Value) & " AND vtc_codser = " & Val(fpLongInteger1(2).Value) & " AND convert(int,substring(convert(varchar(8),vtc_fecvta),1,6))=" & Val(Format(fpDateTime1.text, "yyyymm")) & " AND vtc_forpag = " & Val(fg_codigocbo(Combo2, 0, 1, "")) & " AND vtc_opccli = '" & IIf(valcli = True, "1", "0") & "'"
       End If
    Else
       If vg_tipbase = "1" Then
          vg_db.Execute "DELETE b_ventacontado FROM b_ventacontado  " & _
                        "WHERE  b_ventacontado.vtc_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' AND b_ventacontado.vtc_codreg = " & Val(fpLongInteger1(1).Value) & " AND b_ventacontado.vtc_codser = " & Val(fpLongInteger1(2).Value) & " AND val(mid(b_ventacontado.vtc_fecvta,1,6))=" & Val(Format(fpDateTime1.text, "yyyymm")) & " AND b_ventacontado.vtc_forpag=" & Val(fg_codigocbo(Combo2, 0, 1, "")) & " AND b_ventacontado.vtc_opccli = '" & IIf(valcli = True, "1", "0") & "'"
       Else
          vg_db.Execute "DELETE b_ventacontado FROM b_ventacontado  " & _
                        "WHERE  b_ventacontado.vtc_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' AND b_ventacontado.vtc_codreg = " & Val(fpLongInteger1(1).Value) & " AND b_ventacontado.vtc_codser = " & Val(fpLongInteger1(2).Value) & " AND convert(int,substring(convert(varchar(8),b_ventacontado.vtc_fecvta),1,6)) = " & Val(Format(fpDateTime1.text, "yyyymm")) & " AND b_ventacontado.vtc_forpag=" & Val(fg_codigocbo(Combo2, 0, 1, "")) & " AND b_ventacontado.vtc_opccli = '" & IIf(valcli = True, "1", "0") & "'"
       End If
    End If
    vg_db.CommitTrans
    Mover_Datos
    SumarTotales
Case 7 '------- Actualizar
    Mover_Datos
    SumarTotales
Case 10 '------- Cancelar
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    SSTab1.TabVisible(1) = False
    If modo = "A" Then
       MoverFormato False
       modo = "": Gl_Ac_Botones Me, 1, 2, modo
    Else
       Mover_Datos
       modo = "": Gl_Ac_Botones Me, 1, 4, modo
    End If
    Frame1.Enabled = True
    SumarTotales
    SSTab1.TabEnabled(0) = True: SSTab1.Tab = 0
Case 12 '------- Confirmar
    fg_carga ""
    Dim cencos As String, detmon As Double, codcli As String, descripcion As String
    If Trim(fpayuda(4).Caption) = "" Or Trim(fpayuda(1).Caption) = "" Or Trim(fpayuda(2).Caption) = "" Or fpDateTime1.text = "" Or Combo2(0).ListIndex = -1 Then Exit Sub
    If modo = "A" Then
       For i = 1 To vaSpread1.MaxRows Step 2
           vaSpread1.Row = i
           For j = 1 To vaSpread1.MaxCols
               vaSpread1.Row = i
               vaSpread1.Col = j
              If Trim(vaSpread1.text) <> "" Then
                 diavta = Trim(vaSpread1.text)
                 vaSpread1.Row = i + 1
                 totmon = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
                 If totmon > 0 Then
                    codigo = 0
                    RS.Open "SELECT vtc_codigo FROM b_ventacontado ORDER BY vtc_codigo DESC", vg_db, adOpenStatic
                    If Not RS.EOF Then RS.MoveFirst: codigo = RS!vtc_codigo + 1 Else codigo = 1
                    RS.Close: Set RS = Nothing
                     vg_db.BeginTrans
                     vg_db.Execute "INSERT INTO b_ventacontado (vtc_codigo, vtc_cencos, vtc_codreg, vtc_codser, vtc_fecvta, vtc_forpag, vtc_totmon, vtc_opccli) VALUES (" & codigo & ", '" & LimpiaDato(Trim(fpText(1).text)) & "', " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & Mid(fg_pone_cero(fpDateTime1.text, 7), 4, 4) & Mid(fg_pone_cero(fpDateTime1.text, 7), 1, 2) & fg_pone_cero(diavta, 2) & ", " & Val(fg_codigocbo(Combo2, 0, 1, "")) & ", " & totmon & ", '" & IIf(valcli = True, "1", "0") & "')"
                     If valcli And SSTab1.TabVisible(1) = True Then
                        vg_db.Execute "DELETE b_ventacontadodet FROM b_ventacontadodet WHERE vtd_codigo=" & codigo & ""
                        For X = 1 To vaSpread2.MaxRows
                            cencos = "": detmon = 0: codcli = "": descripcion = ""
                            vaSpread2.Row = X
                            vaSpread2.Col = 1: cencos = Trim(vaSpread2.text)
                            vaSpread2.Col = 3: descripcion = Trim(vaSpread2.text)
                            vaSpread2.Col = 4: detmon = IIf(Trim(vaSpread2.text) = "", 0, vaSpread2.text)
                            If detmon > 0 Then
                               vg_db.Execute "INSERT INTO b_ventacontadodet (vtd_codigo, vtd_numlin, vtd_codcli, vtd_codcco, vtd_descripcion, vtd_detmon) VALUES (" & codigo & ", " & X & ", '" & Trim(fg_DespintaRut(fpText(0).text)) & "', '" & cencos & "', '" & descripcion & "', " & detmon & ")"
                            End If
                        Next X
                     End If
                     vg_db.CommitTrans
                 End If
              End If
           Next j
       Next i
       SSTab1.TabEnabled(0) = True
       modo = "M"
    Else
       For i = 1 To vaSpread1.MaxRows Step 2
           vaSpread1.Row = i
           For j = 1 To vaSpread1.MaxCols
               vaSpread1.Row = i
               vaSpread1.Col = j
               If Trim(vaSpread1.text) <> "" Then
                  diavta = Trim(vaSpread1.text)
                  vaSpread1.Row = i + 1
                  totmon = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
                  RS.Open "SELECT DISTINCT vtc_codigo, vtc_fecvta, vtc_totmon " & _
                          "FROM   b_ventacontado " & _
                          "WHERE  vtc_cencos='" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
                          "AND    vtc_codreg=" & Val(fpLongInteger1(1).Value) & " " & _
                          "AND    vtc_codser=" & Val(fpLongInteger1(2).Value) & " " & _
                          "AND    vtc_fecvta=" & Val(Mid(fg_pone_cero(fpDateTime1.text, 7), 4, 4) & Mid(fg_pone_cero(fpDateTime1.text, 7), 1, 2) & fg_pone_cero(Str(Right(diavta, 2)), 2)) & " " & _
                          "AND    vtc_forpag=" & Val(fg_codigocbo(Combo2, 0, 1, "")) & " AND vtc_opccli='" & IIf(valcli = True, "1", "0") & "'", vg_db, adOpenStatic
                  If RS.EOF And totmon > 0 Then
                     codigo = 0
                     RS1.Open "SELECT vtc_codigo FROM b_ventacontado ORDER BY vtc_codigo DESC", vg_db, adOpenStatic
                     If Not RS1.EOF Then RS1.MoveFirst: codigo = RS1!vtc_codigo + 1 Else codigo = 1
                     RS1.Close: Set RS1 = Nothing
                     vg_db.BeginTrans
                     vg_db.Execute "INSERT INTO b_ventacontado (vtc_codigo, vtc_cencos, vtc_codreg, vtc_codser, vtc_fecvta, vtc_forpag, vtc_totmon, vtc_opccli) VALUES (" & codigo & ", '" & LimpiaDato(Trim(fpText(1).text)) & "', " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & Mid(fg_pone_cero(fpDateTime1.text, 7), 4, 4) & Mid(fg_pone_cero(fpDateTime1.text, 7), 1, 2) & fg_pone_cero(diavta, 2) & ", " & Val(fg_codigocbo(Combo2, 0, 1, "")) & ", " & totmon & ", '" & IIf(valcli = True, "1", "0") & "')"
                     If valcli And SSTab1.TabVisible(1) = True And fg_pone_cero(Val(numdia), 2) = fg_pone_cero(diavta, 2) Then
                        vg_db.Execute "DELETE b_ventacontadodet FROM b_ventacontadodet WHERE vtd_codigo=" & codigo & ""
                        For X = 1 To vaSpread2.MaxRows
                            cencos = "": detmon = 0: codcli = "": descripcion = ""
                            vaSpread2.Row = X
                            vaSpread2.Col = 1: cencos = Trim(vaSpread2.text)
                            vaSpread2.Col = 3: descripcion = Trim(vaSpread2.text)
                            vaSpread2.Col = 4: detmon = IIf(Trim(vaSpread2.text) = "", 0, vaSpread2.text)
                            If detmon > 0 Then
                               vg_db.Execute "INSERT INTO b_ventacontadodet (vtd_codigo, vtd_numlin, vtd_codcli, vtd_codcco, vtd_descripcion, vtd_detmon) VALUES (" & codigo & ", " & X & ", '" & Trim(fg_DespintaRut(fpText(0).text)) & "', '" & cencos & "', '" & descripcion & "', " & detmon & ")"
                            End If
                        Next X
                     End If
                     vg_db.CommitTrans
                  ElseIf Not RS.EOF Then
                     If RS!vtc_totmon <> totmon Or valcli Then
                        vg_db.BeginTrans
                        graenc = False
                        vg_db.Execute "UPDATE b_ventacontado SET vtc_totmon=" & totmon & " WHERE vtc_codigo=" & RS!vtc_codigo & " AND vtc_cencos='" & LimpiaDato(Trim(fpText(1).text)) & "' AND vtc_codreg=" & Val(fpLongInteger1(1).Value) & " AND vtc_codser=" & Val(fpLongInteger1(2).Value) & " AND vtc_fecvta=" & Val(Mid(fg_pone_cero(fpDateTime1.text, 7), 4, 4) & Mid(fg_pone_cero(fpDateTime1.text, 7), 1, 2) & fg_pone_cero(Str(Right(diavta, 2)), 2)) & "  AND vtc_forpag=" & Val(fg_codigocbo(Combo2, 0, 1, "")) & " AND vtc_opccli='" & IIf(valcli = True, "1", "0") & "'"
'                        If Trim(NumDia) = "" Then fg_descarga: SSTab1.TabEnabled(1) = False: vg_db.CommitTrans: Exit Sub
                        If valcli And SSTab1.TabVisible(1) = True And fg_pone_cero(Val(numdia), 2) = fg_pone_cero(diavta, 2) Then
                           vg_db.Execute "DELETE b_ventacontadodet FROM b_ventacontadodet WHERE vtd_codigo=" & RS!vtc_codigo & ""
                           For X = 1 To vaSpread2.MaxRows
                               cencos = "": detmon = 0: codcli = "": descripcion = ""
                               vaSpread2.Row = X
                               vaSpread2.Col = 1: cencos = Trim(vaSpread2.text)
                               vaSpread2.Col = 3: descripcion = Trim(vaSpread2.text)
                               vaSpread2.Col = 4: detmon = IIf(Trim(vaSpread2.text) = "", 0, vaSpread2.text)
                               If detmon > 0 Then
                                  vg_db.Execute "INSERT INTO b_ventacontadodet (vtd_codigo, vtd_numlin, vtd_codcli, vtd_codcco, vtd_descripcion, vtd_detmon) VALUES (" & RS!vtc_codigo & ", " & X & ", '" & Trim(fg_DespintaRut(fpText(0).text)) & "', '" & cencos & "', '" & descripcion & "', " & detmon & ")"
                                  graenc = True
                               End If
                           Next X
                           If Not graenc Then
                              vg_db.Execute "DELETE b_ventacontado FROM b_ventacontado " & _
                                            "WHERE vtc_fecvta=" & RS!vtc_fecvta & " " & _
                                            "AND   vtc_codigo=" & RS!vtc_codigo & " " & _
                                            "AND   vtc_cencos='" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
                                            "AND   vtc_codreg=" & Val(fpLongInteger1(1).Value) & " " & _
                                            "AND   vtc_codser=" & Val(fpLongInteger1(2).Value) & " " & _
                                            "AND   vtc_forpag=" & Val(fg_codigocbo(Combo2, 0, 1, "")) & " AND vtc_opccli='" & IIf(valcli = True, "1", "0") & "'"
                           End If
                        End If
                        vg_db.CommitTrans
                     End If
                  End If
                  RS.Close: Set RS = Nothing
               End If
           Next j
       Next i
       modo = "M"
       SSTab1.TabEnabled(0) = True
    End If
    modo = "": Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 4), modo
    Frame1.Enabled = True
    SumarTotales
    fg_descarga
Case 15 ' Impirmir
    If vaSpread1.MaxRows < 1 Then Exit Sub
    I_VentaContado fpText(1).text, Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), fpDateTime1.text, Trim(Mid(Combo2(0).text, 1, 150))
Case 18 'Salir
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub MoverFormato(opval As Boolean)
Dim diafin As Long, nrosem As Long, i As Long, j As Long
'-------> Armar calendario
vaSpread1.Visible = False
vaSpread1.Row = -1: vaSpread1.Col = -1:
vaSpread1.BackColor = &H80000018 'Label1(0).BackColor
Label1(4).Caption = Meses("01/" & Mid(fg_pone_cero(fpDateTime1.text, 7), 1, 2) & "/" & Mid(fg_pone_cero(fpDateTime1.text, 7), 4, 4)) & " " & Mid(fg_pone_cero(fpDateTime1.text, 6), 4, 4)
diafin = fg_mes(Mid(fg_pone_cero(fpDateTime1.text, 7), 1, 2) & Mid(fg_pone_cero(fpDateTime1.text, 7), 4, 4))
nrosem = 1
For i = 1 To vaSpread1.MaxRows Step 2
    For j = 1 To vaSpread1.MaxCols
        vaSpread1.Row = i
        vaSpread1.Col = j
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = ""
        vaSpread1.BackColor = &H8000000F
        vaSpread1.Lock = False
        vaSpread1.Row = i + 1
        vaSpread1.Col = j
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = ""
        vaSpread1.Lock = False
    Next j
Next i
For i = 1 To diafin
    Select Case fg_Dia(Mid(fg_pone_cero(fpDateTime1.text, 7), 4, 4) & Mid(fg_pone_cero(fpDateTime1.text, 7), 1, 2) & fg_pone_cero(i, 2))
    Case 1
        vaSpread1.Row = nrosem
        vaSpread1.Col = 7
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignCenter
        vaSpread1.ForeColor = &H800000
        vaSpread1.text = CStr(i)
        vaSpread1.BackColor = &H8000000F
        vaSpread1.Lock = False
        
        vaSpread1.Row = nrosem + 1
        vaSpread1.Col = 7
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = ""
        If opval And Not valcli Then
           vaSpread1.CellType = CellTypeCurrency
           vaSpread1.TypeVAlign = TypeVAlignCenter
           vaSpread1.TypeCurrencyDecimal = "."
           vaSpread1.TypeCurrencyMin = 0
           vaSpread1.TypeCurrencyMax = 99999999999#
           vaSpread1.TypeCurrencyDecPlaces = 2 '0
           vaSpread1.TypeCurrencyNegStyle = TypeCurrencyNegStyle1
           vaSpread1.TypeCurrencyPosStyle = TypeCurrencyPosStyle1
           vaSpread1.TypeCurrencySeparator = ","
           vaSpread1.TypeCurrencyShowSep = True
           vaSpread1.TypeCurrencyShowSymbol = False
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        ElseIf opval And valcli Then
           vaSpread1.TypeHAlign = TypeHAlignCenter
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        End If
        nrosem = nrosem + 2
    Case 2
        vaSpread1.Row = nrosem
        vaSpread1.Col = 1
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignCenter
        vaSpread1.ForeColor = &H800000
        vaSpread1.text = CStr(i)
        vaSpread1.BackColor = &H8000000F
        vaSpread1.Lock = False
        
        vaSpread1.Row = nrosem + 1
        vaSpread1.Col = 1
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = ""
        If opval And Not valcli Then
           vaSpread1.CellType = CellTypeCurrency
           vaSpread1.TypeVAlign = TypeVAlignCenter
           vaSpread1.TypeCurrencyDecimal = "."
           vaSpread1.TypeCurrencyMin = 0
           vaSpread1.TypeCurrencyMax = 99999999999#
           vaSpread1.TypeCurrencyDecPlaces = 2 '0
           vaSpread1.TypeCurrencyNegStyle = TypeCurrencyNegStyle1
           vaSpread1.TypeCurrencyPosStyle = TypeCurrencyPosStyle1
           vaSpread1.TypeCurrencySeparator = ","
           vaSpread1.TypeCurrencyShowSep = True
           vaSpread1.TypeCurrencyShowSymbol = False
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        ElseIf opval And valcli Then
           vaSpread1.TypeHAlign = TypeHAlignCenter
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        End If
    Case 3
        vaSpread1.Row = nrosem
        vaSpread1.Col = 2
         vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignCenter
        vaSpread1.ForeColor = &H800000
        vaSpread1.text = CStr(i)
        vaSpread1.BackColor = &H8000000F
        vaSpread1.Lock = False
       
        vaSpread1.Row = nrosem + 1
        vaSpread1.Col = 2
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = ""
        If opval And Not valcli Then
           vaSpread1.CellType = CellTypeCurrency
           vaSpread1.TypeVAlign = TypeVAlignCenter
           vaSpread1.TypeCurrencyDecimal = "."
           vaSpread1.TypeCurrencyMin = 0
           vaSpread1.TypeCurrencyMax = 99999999999#
           vaSpread1.TypeCurrencyDecPlaces = 2 '0
           vaSpread1.TypeCurrencyNegStyle = TypeCurrencyNegStyle1
           vaSpread1.TypeCurrencyPosStyle = TypeCurrencyPosStyle1
           vaSpread1.TypeCurrencySeparator = ","
           vaSpread1.TypeCurrencyShowSep = True
           vaSpread1.TypeCurrencyShowSymbol = False
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        ElseIf opval And valcli Then
           vaSpread1.TypeHAlign = TypeHAlignCenter
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        End If
    Case 4
        vaSpread1.Row = nrosem
        vaSpread1.Col = 3
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignCenter
        vaSpread1.ForeColor = &H800000
        vaSpread1.text = CStr(i)
        vaSpread1.BackColor = &H8000000F
        vaSpread1.Lock = False
        
        vaSpread1.Row = nrosem + 1
        vaSpread1.Col = 3
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = ""
        If opval And Not valcli Then
           vaSpread1.CellType = CellTypeCurrency
           vaSpread1.TypeVAlign = TypeVAlignCenter
           vaSpread1.TypeCurrencyDecimal = "."
           vaSpread1.TypeCurrencyMin = 0
           vaSpread1.TypeCurrencyMax = 99999999999#
           vaSpread1.TypeCurrencyDecPlaces = 2 '0
           vaSpread1.TypeCurrencyNegStyle = TypeCurrencyNegStyle1
           vaSpread1.TypeCurrencyPosStyle = TypeCurrencyPosStyle1
           vaSpread1.TypeCurrencySeparator = ","
           vaSpread1.TypeCurrencyShowSep = True
           vaSpread1.TypeCurrencyShowSymbol = False
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        ElseIf opval And valcli Then
           vaSpread1.TypeHAlign = TypeHAlignCenter
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        End If
    Case 5
        vaSpread1.Row = nrosem
        vaSpread1.Col = 4
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignCenter
        vaSpread1.ForeColor = &H800000
        vaSpread1.text = CStr(i)
        vaSpread1.BackColor = &H8000000F
        vaSpread1.Lock = False
        
        vaSpread1.Row = nrosem + 1
        vaSpread1.Col = 4
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = ""
        If opval And Not valcli Then
           vaSpread1.CellType = CellTypeCurrency
           vaSpread1.TypeVAlign = TypeVAlignCenter
           vaSpread1.TypeCurrencyDecimal = "."
           vaSpread1.TypeCurrencyMin = 0
           vaSpread1.TypeCurrencyMax = 99999999999#
           vaSpread1.TypeCurrencyDecPlaces = 2 '0
           vaSpread1.TypeCurrencyNegStyle = TypeCurrencyNegStyle1
           vaSpread1.TypeCurrencyPosStyle = TypeCurrencyPosStyle1
           vaSpread1.TypeCurrencySeparator = ","
           vaSpread1.TypeCurrencyShowSep = True
           vaSpread1.TypeCurrencyShowSymbol = False
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        ElseIf opval And valcli Then
           vaSpread1.TypeHAlign = TypeHAlignCenter
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        End If
    Case 6
        vaSpread1.Row = nrosem
        vaSpread1.Col = 5
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignCenter
        vaSpread1.ForeColor = &H800000
        vaSpread1.text = CStr(i)
        vaSpread1.BackColor = &H8000000F
        vaSpread1.Lock = False
        
        vaSpread1.Row = nrosem + 1
        vaSpread1.Col = 5
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = ""
        If opval And Not valcli Then
           vaSpread1.CellType = CellTypeCurrency
           vaSpread1.TypeVAlign = TypeVAlignCenter
           vaSpread1.TypeCurrencyDecimal = "."
           vaSpread1.TypeCurrencyMin = 0
           vaSpread1.TypeCurrencyMax = 99999999999#
           vaSpread1.TypeCurrencyDecPlaces = 2 '0
           vaSpread1.TypeCurrencyNegStyle = TypeCurrencyNegStyle1
           vaSpread1.TypeCurrencyPosStyle = TypeCurrencyPosStyle1
           vaSpread1.TypeCurrencySeparator = ","
           vaSpread1.TypeCurrencyShowSep = True
           vaSpread1.TypeCurrencyShowSymbol = False
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        ElseIf opval And valcli Then
           vaSpread1.TypeHAlign = TypeHAlignCenter
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        End If
    Case 7
        vaSpread1.Row = nrosem
        vaSpread1.Col = 6
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignCenter
        vaSpread1.ForeColor = &H800000
        vaSpread1.text = CStr(i)
        vaSpread1.BackColor = &H8000000F
        vaSpread1.Lock = False
        
        vaSpread1.Row = nrosem + 1
        vaSpread1.Col = 6
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = ""
        If opval And Not valcli Then
           vaSpread1.CellType = CellTypeCurrency
           vaSpread1.TypeVAlign = TypeVAlignCenter
           vaSpread1.TypeCurrencyDecimal = "."
           vaSpread1.TypeCurrencyMin = 0
           vaSpread1.TypeCurrencyMax = 99999999999#
           vaSpread1.TypeCurrencyDecPlaces = 2 '0
           vaSpread1.TypeCurrencyNegStyle = TypeCurrencyNegStyle1
           vaSpread1.TypeCurrencyPosStyle = TypeCurrencyPosStyle1
           vaSpread1.TypeCurrencySeparator = ","
           vaSpread1.TypeCurrencyShowSep = True
           vaSpread1.TypeCurrencyShowSymbol = False
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        ElseIf opval And valcli Then
           vaSpread1.TypeHAlign = TypeHAlignCenter
           vaSpread1.text = Format(vecmon(i), fg_Pict(11, 2))
           vaSpread1.ForeColor = &HFF0000
        End If
    End Select
Next i
vaSpread1.RetainSelBlock = False
vaSpread1.Visible = True
''-------> Bloquea días de cierre en color rojo
'Dim diablq As Date
'Dim v_columnas As Double
'If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Date, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
'If Format(Date, "dd/mm/yyyy") > diablq Or Format(CDate(fpDateTime1.text), "mm/yyyy") < Format(Month(Date) - 1 & "/" & Year(Date), "mm/yyyy") Then
'    If Month(Date) = 1 Then
'        v_columnas = ((dEoM(Format("01/" & "12/" & Year(Date) - 1, "dd/mm/yyyy")) - CDate(CDate("01/" + Mid(fpDateTime1.text, 1, 8))) + 1) * 2) - 1
'    Else
'        v_columnas = ((dEoM(Format("01/" & (Month(Date) - 1) & "/" & Year(Date), "dd/mm/yyyy")) - CDate(CDate("01/" + Mid(fpDateTime1.text, 1, 8))) + 1) * 2) - 1
'    End If
'Else
'   v_columnas = 0
'End If
'
'If v_columnas > 0 Then
'    For i = 2 To vaSpread1.MaxRows Step 2
'        vaSpread1.Row = i
'        For j = 1 To vaSpread1.MaxCols
'            vaSpread1.Col = j
'            vaSpread1.Lock = True
'            vaSpread1.BackColor = Shape1(0).FillColor
'        Next j
'   Next i
'   vaSpread1.SetActiveCell i, 1
'End If
''-------> Fin Bloqueo de celdas

'-------> Bloquear cierre diario
For i = 1 To vaSpread1.MaxRows Step 2
    vaSpread1.Row = i
    For j = 1 To vaSpread1.MaxCols
        vaSpread1.Col = j
        If Trim(vaSpread1.text) <> "" And Trim(vaSpread1.text) <> "0" Then
           If CDate(fg_pone_cero(vaSpread1.text, 2) & "/" & Trim(fpDateTime1.text)) < CDate(vg_ciedia) Then
              vaSpread1.Row = i + 1
              vaSpread1.Lock = True
              vaSpread1.BackColor = Shape1(0).FillColor
              vaSpread1.Row = i
           End If
        End If
    Next j
Next i
'-------> Fin bloquear cierre diario
vaSpread1.SetActiveCell 1, 1
If vaSpread1.MaxRows < 1 Then Exit Sub
'vaSpread1.Row = 2: vaSpread1.Col = vaSpread1.MaxCols
'If vaSpread1.BackColor = Shape1(0).FillColor Then

If CDate(dEoM(fg_pone_cero(1, 2) & "/" & Trim(fpDateTime1.text))) <= CDate(vg_ciedia) - 1 Then
   mesblo = True: modo = "": Gl_Ac_Botones Me, 1, 3, modo
Else
   mesblo = False
End If
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
TraerFechaCierre
indrow = 0: indcol = 0
If Row = 0 Or Not valcli Or mesblo Then SSTab1.TabVisible(1) = False: Exit Sub
If valcli And Toolbar1.Buttons(1).Visible = False Then SSTab1.TabVisible(1) = True
i = Row
vaSpread1.Row = Row
vaSpread1.Col = Col
indcol = Col
If vaSpread1.BackColor = Shape1(1).FillColor Or vaSpread1.BackColor = Shape1(0).FillColor Then
   vaSpread1.Row = Row - 1: i = Row - 1: indrow = Row
End If
numdia = ""
numdia = fg_pone_cero(vaSpread1.text, 2)
If Trim(numdia) = "" Then SSTab1.TabVisible(1) = False: Exit Sub
fg_carga ""
RS.Open "SELECT b.* FROM b_ventacontado a, b_ventacontadodet b, b_clientecencos c " & _
        "WHERE a.vtc_codigo = b.vtd_codigo " & _
        "AND   b.vtd_codcli = c.clc_codcli " & _
        "AND   b.vtd_codcco = c.clc_codigo " & _
        "AND   a.vtc_cencos = '" & LimpiaDato(Trim(fpText(1).text)) & "' " & _
        "AND   a.vtc_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
        "AND   a.vtc_codser = " & Val(fpLongInteger1(2).Value) & " " & _
        "AND   a.vtc_fecvta = " & Val(Format(fpDateTime1.text, "yyyymm") & numdia) & " " & _
        "AND   a.vtc_forpag = " & Val(fg_codigocbo(Combo2, 0, 1, "")) & "", vg_db, adOpenStatic
vaSpread2.MaxRows = 0
RS1.Open "SELECT * FROM b_clientecencos WHERE clc_codcli = '" & Trim(fg_DespintaRut(fpText(0).text)) & "'", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      vaSpread2.Col = 1: vaSpread2.text = RS1!clc_codigo
      vaSpread2.Col = 2: vaSpread2.text = Trim(RS1!clc_nombre)
      vaSpread2.Col = 3: vaSpread2.text = " "
      vaSpread2.Col = 4: vaSpread2.text = "   "
      If CDate(numdia & "/" & Trim(fpDateTime1.text)) <= CDate(vg_ciedia) - 1 Then
         vaSpread2.Row = -1: vaSpread2.Col = -1
         vaSpread2.BackColor = Shape1(0).FillColor
         vaSpread2.Row = vaSpread2.MaxRows
         vaSpread2.Col = 3: vaSpread2.Lock = True
         vaSpread2.Col = 4: vaSpread2.Lock = True
      Else
         vaSpread2.Row = -1: vaSpread2.Col = -1
         vaSpread2.BackColor = Shape1(1).FillColor
         vaSpread2.Row = vaSpread2.MaxRows
         vaSpread2.Col = 3: vaSpread2.Lock = False
         vaSpread2.Col = 4: vaSpread2.Lock = False
      End If
      If Not RS.EOF Then
         RS.MoveFirst
         Do While Not RS.EOF
            If Trim(RS1!clc_codigo) = Trim(RS!vtd_codcco) Then
               vaSpread2.Col = 3: vaSpread2.text = Trim(RS!vtd_descripcion)
               vaSpread2.Col = 4: vaSpread2.text = Format(RS!vtd_detmon, fg_Pict(11, 2))
               Exit Do
            End If
            RS.MoveNext
         Loop
         RS.MoveFirst
      End If
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
RS.Close: Set RS = Nothing
vaSpread1.Row = 0
Label1(10).Caption = vaSpread1.text
vaSpread1.Row = i
Label1(10).Caption = Label1(10).Caption & "  " & fg_pone_cero(vaSpread1.text, 2)
fg_descarga
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.MaxRows < 1 Or mesblo Then Exit Sub
If modo = "" Then modo = "M"
If ChangeMade = True And modo = "M" Then
   Gl_Ac_Botones Me, 1, 0, modo: Frame1.Enabled = False
End If
SumarTotales
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Select Case NewRow
Case 3, 5, 7, 9, 11
    If NewCol = 1 Then vaSpread1.Row = NewRow + 1: vaSpread1.SetActiveCell NewCol, NewRow + 1
End Select
End Sub

Sub SumarTotales()
Dim totmon As Double
'------- Sumar Totales
totmon = 0
For i = 2 To vaSpread1.MaxRows Step 2
    vaSpread1.Row = i
    For j = 1 To vaSpread1.MaxCols
        vaSpread1.Col = j
        If Trim(vaSpread1.text) <> "" Then
           totmon = Format((totmon + vaSpread1.text), fg_Pict(11, 2))
        End If
    Next j
Next i
Label1(6).Caption = Format(totmon, fg_Pict(11, 2))
End Sub

Private Sub vaSpread2_EditChange(ByVal Col As Long, ByVal Row As Long)
If vaSpread2.MaxRows < 1 Or mesblo Then Exit Sub 'Or Toolbar1.Buttons(12).Visible = True Then Exit Sub
Dim totmon As Double
If modo = "" Then modo = "M"
SSTab1.TabEnabled(0) = False
Gl_Ac_Botones Me, 1, 0, modo
totmon = 0
For i = 1 To vaSpread2.MaxRows
    vaSpread2.Row = i
    vaSpread2.Col = 4
    If Trim(vaSpread2.text) <> "" Then totmon = Format(totmon + IIf(Trim(vaSpread2.text) = "", 0, vaSpread2.text), fg_Pict(11, 2))
Next i
vaSpread1.Row = indrow
vaSpread1.Col = indcol
vaSpread1.text = Format(totmon, fg_Pict(11, 2))
Frame1.Enabled = False
End Sub
