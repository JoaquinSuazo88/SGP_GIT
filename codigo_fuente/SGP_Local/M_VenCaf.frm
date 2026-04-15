VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_VenCaf 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Venta Cafetería"
   ClientHeight    =   6270
   ClientLeft      =   3330
   ClientTop       =   3015
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5910
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   10425
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Venta Cafetería"
      TabPicture(0)   =   "M_VenCaf.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "vaSpread1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Inventario producto"
      TabPicture(1)   =   "M_VenCaf.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "Shape1(2)"
      Tab(1).Control(2)=   "Shape1(1)"
      Tab(1).Control(3)=   "ImageList1"
      Tab(1).Control(4)=   "vaSpread2"
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1350
         Left            =   90
         TabIndex        =   3
         Top             =   315
         Width           =   9555
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   870
            Width           =   3195
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   5
            Top             =   525
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   24
            Top             =   180
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
            ButtonStyle     =   3
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
            Text            =   ""
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
            ButtonColor     =   -2147483633
            AutoMenu        =   -1  'True
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
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
            Left            =   3360
            TabIndex        =   12
            Top             =   525
            Width           =   4545
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            Left            =   360
            TabIndex        =   11
            Top             =   225
            Width           =   540
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   2910
            Picture         =   "M_VenCaf.frx":0038
            Top             =   420
            Width           =   480
         End
         Begin VB.Label Label3 
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
            Index           =   6
            Left            =   360
            TabIndex        =   10
            Top             =   570
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Bodega"
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
            Left            =   360
            TabIndex        =   9
            Top             =   930
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Left            =   3345
            TabIndex        =   8
            Top             =   240
            Width           =   600
         End
         Begin VB.Label label 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   1590
            TabIndex        =   7
            Top             =   930
            Width           =   3180
         End
         Begin VB.Label label 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   6
            Left            =   3405
            TabIndex        =   6
            Top             =   555
            Width           =   4545
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2925
         Left            =   90
         TabIndex        =   2
         Top             =   1740
         Width           =   9555
         _Version        =   393216
         _ExtentX        =   16854
         _ExtentY        =   5159
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         ButtonDrawMode  =   1
         EditEnterAction =   4
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
         MaxCols         =   7
         MaxRows         =   20
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_VenCaf.frx":0342
         TextTip         =   2
         TextTipDelay    =   0
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   4680
         Left            =   -74910
         TabIndex        =   13
         Top             =   405
         Width           =   9075
         _Version        =   393216
         _ExtentX        =   16007
         _ExtentY        =   8255
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   8
         MaxRows         =   10
         SpreadDesigner  =   "M_VenCaf.frx":0A74
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -70530
         Top             =   5070
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_VenCaf.frx":10B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_VenCaf.frx":13D2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Datos del Cliente"
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   90
         TabIndex        =   15
         Top             =   4725
         Width           =   9555
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   2
            Left            =   2175
            TabIndex        =   20
            Top             =   600
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   0
            Left            =   2175
            TabIndex        =   22
            Top             =   1050
            Visible         =   0   'False
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
            MaxValue        =   "2147483647"
            MinValue        =   "-2147483648"
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
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   1
            Left            =   2175
            TabIndex        =   25
            Top             =   1395
            Visible         =   0   'False
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
            ButtonStyle     =   3
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
            Text            =   ""
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
            Index           =   1
            Left            =   2175
            TabIndex        =   26
            Top             =   255
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
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
            Left            =   405
            TabIndex        =   23
            Top             =   1095
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Documento"
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
            Left            =   405
            TabIndex        =   21
            Top             =   1425
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Centro Costo"
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
            Left            =   405
            TabIndex        =   19
            Top             =   675
            Width           =   1110
         End
         Begin VB.Label Label3 
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
            Index           =   0
            Left            =   405
            TabIndex        =   17
            Top             =   345
            Width           =   600
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   3555
            Picture         =   "M_VenCaf.frx":16EC
            Top             =   150
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   4005
            TabIndex        =   16
            Top             =   255
            Width           =   4545
         End
         Begin VB.Label label 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   4050
            TabIndex        =   18
            Top             =   285
            Width           =   4545
         End
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008484FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   -68610
         Top             =   5145
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   -68610
         Top             =   5265
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad sobrepasa Stock actual"
         Height          =   195
         Left            =   -68250
         TabIndex        =   14
         Top             =   5100
         Width           =   2355
      End
   End
End
Attribute VB_Name = "M_VenCaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset
Dim est As Boolean
Dim modo As String

Private Sub Combo1_Click(Index As Integer)
Dim codbod As Long, i As Long
If est Then Exit Sub
est = True
codbod = Val(fg_codigocbo(Combo1, 0, 10, ""))
For i = 1 To vaSpread2.MaxRows
    vaSpread2.Col = 1: vaSpread2.Row = i
    RevisaSobreStock vaSpread2.text, i, True
Next i
est = False
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 6780
Me.Width = 9915
fg_centra Me
SSTab1.Tab = 0
est = True
EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)
Me.HelpContextID = vg_OpcM
MsgTitulo = "Venta Cafetería"
Dim X As Boolean
vaSpread2.TextTip = 2
vaSpread2.TextTipDelay = 0
X = vaSpread2.SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
Gl_Mo_Botones Me, 11
vaSpread1.Row = -1
vaSpread1.Col = 3: vaSpread1.TypeNumberShowSep = True: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 4: vaSpread1.TypeNumberShowSep = True: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
vaSpread2.Col = 3: vaSpread2.TypeNumberShowSep = True: vaSpread2.TypeNumberSeparator = vg_CSep: vaSpread2.TypeNumberDecimal = vg_CDec: vaSpread2.TypeNumberDecPlaces = IIf(vg_pais = "CL", 3, vg_DCa)
vaSpread2.Col = 4: vaSpread2.TypeNumberShowSep = True: vaSpread2.TypeNumberSeparator = vg_CSep: vaSpread2.TypeNumberDecimal = vg_CDec: vaSpread2.TypeNumberDecPlaces = IIf(vg_pais = "CL", 3, vg_DCa)
vaSpread2.Col = 5: vaSpread2.TypeNumberShowSep = True: vaSpread2.TypeNumberSeparator = vg_CSep: vaSpread2.TypeNumberDecimal = vg_CDec: vaSpread2.TypeNumberDecPlaces = vg_DPr
'-------> Cargar Combo Bodega
CargarDatoCombo Combo1, 0, "b_clientes", "cli_", "CliBod", "N"
Limpia 1
est = False
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    est = True
    If Not IsDate(fpDateTime1(0).text) Then Limpia 2: est = False: Exit Sub
    vaSpread1.MaxRows = 0: modo = "": Gl_Ac_Botones Me, 11, IIf(vaSpread1.MaxRows = 0, 3, 1), modo
    MoverDatos
    est = False
End Select
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_Change(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    fpayuda(Index).Caption = ""
Case 1
    fpayuda(Index).Caption = ""
End Select
If modo = "" Then modo = "M": Gl_Ac_Botones Me, 11, 0, modo
End Sub

Private Sub fpText1_GotFocus(Index As Integer)
If est Then Exit Sub
If fpText1(Index).text = "" Then Exit Sub
Select Case Index
Case 1
    est = True
    fpText1(1).text = fg_DespintaRut(Trim(fpText1(1).text))
    fpText1(1).text = Mid(fpText1(1).text, 1, Len(Trim(fpText1(1).text)) - 1)
    est = False
End Select
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_LostFocus(Index As Integer)
If est Then Exit Sub
If fpText1(Index).text = "" Then Exit Sub
Select Case Index
Case 0
    est = True
    RS1.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo = '" & fpText1(0).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          fpayuda(Index).Caption = RS1!cli_nombre
          RS1.MoveNext
       Loop
    Else
       RS1.Close: Set RS1 = Nothing
       fpText1(0).text = ""
       MsgBox "Contrato no existe...", vbExclamation + vbOKOnly, MsgTitulo
       If fpText1(0).Enabled = True Then fpText1(0).SetFocus
       est = False
       Exit Sub
    End If
    RS1.Close: Set RS1 = Nothing
    est = False
Case 1
    est = True
    If InStr(Trim(fpText1(1).text), "-") = 0 And Trim(fpText1(1).text) <> "" Then
        fpText1(1).text = fg_RutDig(Trim(fpText1(1).text))
        fpText1(1).text = fg_PintaRut(Trim(fpText1(1).text))
    End If
    RS1.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & fg_DespintaRut(fpText1(1).text) & "' AND cli_tipo = 1 AND cli_activo = '1'", vg_db, adOpenStatic
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            fpText1(1).text = fg_PintaRut(RS1!cli_codigo)
            fpayuda(Index).Caption = RS1!cli_nombre
            RS1.MoveNext
        Loop
    Else
        RS1.Close: Set RS1 = Nothing
        fpText1(1).text = ""
        MsgBox "Cliente no existe...", vbExclamation + vbOKOnly, MsgTitulo
        If fpText1(1).Enabled = True Then fpText1(1).SetFocus
        est = False
        Exit Sub
    End If
    RS1.Close: Set RS1 = Nothing
    est = False
End Select
End Sub

Private Sub Image1_Click(Index As Integer)
Dim Variable As String
vg_codigo = 0
vg_left = fpayuda(Index).Left + 1920
vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 5:
Variable = IIf(Index = 0, "Contrato", "Cliente")
If Variable = "Cliente" Then
   If vaSpread1.TypeComboBoxCurSel = 1 Then
      B_TabEst.LlenaDatos "b_clientes", "cli_", Variable, Variable
   ElseIf vaSpread1.TypeComboBoxCurSel = 2 Then
      Variable = "CliAlum"
      B_TabEst.LlenaDatos "b_clientes", "cli_", Variable, Variable
   End If
Else
   Variable = "Contrato"
   B_TabEst.LlenaDatos "b_clientes", "cli_", Variable, Variable
End If
B_TabEst.Show 1
Me.Refresh
If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
fpText1(Index) = Trim(vg_codigo)
fpayuda(Index).Caption = vg_nombre
Select Case Index
Case 0
    If Trim(fpText1(Index).text) = "" Then Exit Sub
    If fpDateTime1(Index).Enabled = True Then fpDateTime1(Index).SetFocus
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim cencos As String, codbod As Long, Fecha As Date, rutcli As String, cencli As String
Dim cCAr As String, cCIn As Double, cPVe As Double, cTip As String, cNli As Long
Dim cCPr As String, cCCa As Double, cCDi As Double, cPCo As Double, cAso As String, cCal As Double
Dim RS1 As New ADODB.Recordset
On Error GoTo Man_Error
fecpro = Format(fpDateTime1(0).Value, "dd/mm/yyyy")
codbod = Val(fg_codigocbo(Combo1, 0, 10, ""))
If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
TraerFechaCierre
Select Case Button.Index
Case 1 '-------> Agrega
    '-------> Validar si el contrato tiene asignado inventario rotativo
    If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
       Format(fpDateTime1(0).Value, "dd/mm/yyyy") > CDate(vg_ciedia) Then MsgBox "Tiene que realizar cierre diario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then
    
       MsgBox "Periodo cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then
    
       MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    'Validar inventario calendarizado 20201001
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 38) Then
        
       MsgBox "Se esta realizando la toma de inventario en estos momento...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
    End If
        
    'Validar ingreso documento inventario calendarizado 20201001
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 40) Then
        
       MsgBox "No puede ingresar documento, antes de un inventario calendarizado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
    End If
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 8) Then
       
       MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Then
    
       MsgBox "Día se encuentra cerrado, no es posible ingresar...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    est = True
    SSTab1.Tab = 0
    vg_left = Screen.Width \ 2 - B_TabEst.Width \ 2 'vaSpread1.Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_totpreciocaf", "tpc_", "Artículos", "Tpc"
    B_TabEst.Show 1
    If Val(vg_codigo) = 0 Then est = False: Exit Sub
    
    RS1.Open "SELECT * FROM b_detpreciocaf WHERE dpc_cencos = '" & MuestraCasino(1) & "' AND dpc_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: est = False: MsgBox "El articulo no tiene composición...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS1.Close: Set RS1 = Nothing
    
    RS1.Open "SELECT * FROM b_totpreciocaf WHERE tpc_cencos = '" & MuestraCasino(1) & "' AND tpc_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic
    If Not RS1.EOF Then
        modo = "A"
        Gl_Ac_Botones Me, 11, 0, modo
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 1: vaSpread1.text = Trim(RS1!tpc_codigo)
        vaSpread1.Col = 2: vaSpread1.text = Trim(RS1!tpc_nombre)
        vaSpread1.Col = 3: vaSpread1.text = 0
        vaSpread1.Col = 4: vaSpread1.text = Format(RS1!tpc_precio, fg_Pict(9, vg_DPr))
        vaSpread1.Col = 5: vaSpread1.TypeComboBoxList = "CONTADO" & Chr$(9) & "CREDITO" & Chr$(9) & "CUENTA ABONO CLIENTE"
        fpText1(1).text = ""
        fpayuda(1).Caption = ""
        fpText1(2).text = ""
        SumaCantidades True
        If Me.Visible And vaSpread1.Enabled And vaSpread1.Visible Then vaSpread1.SetFocus
        vaSpread1.SetActiveCell 3, vaSpread1.MaxRows
    Else
        RS1.Close: Set RS1 = Nothing
        MsgBox "Artículo de cafetería no existe", vbCritical + vbOKOnly, MsgTitulo
        vaSpread1.text = "": vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: est = False: Exit Sub
    End If
    RS1.Close: Set RS1 = Nothing
    est = False
    vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.Lock = False
    vaSpread2.Row = -1
    vaSpread2.Col = -1: vaSpread2.Lock = False
    vaSpread2.Col = 3: vaSpread2.Lock = True
    vaSpread2.Col = 5: vaSpread2.Lock = True
Case 3 '-------> Modifica
    modo = "M"
    Gl_Ac_Botones Me, 11, 0, modo
Case 5 '-------> Eliminar
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then MsgBox "Periodo cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Then MsgBox "No puede eliminar documento, día esta cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Elimina artículo" & IIf(vaSpread1.MaxRows = 1, "... junto con el último artículo eliminará también el documento", "") & "...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    est = True
    cencos = Trim(LimpiaDato(fpText1(0).text))
    Fecha = Format(fpDateTime1(0).text, "dd/mm/yyyy")
    vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: cCAr = Trim(vaSpread1.text)
    Toolbar1.Enabled = False
    vg_db.BeginTrans
        '-------> Detalle Articulos
        If vg_tipbase = "1" Then
           vg_db.Execute "DELETE FROM b_detventascaf WHERE dvc_cencos = '" & cencos & "' AND dvc_fecing = cdate('" & Fecha & "') AND dvc_numlin = " & vaSpread1.ActiveRow
           vg_db.Execute "UPDATE b_detventascaf SET dvc_numlin=dvc_numlin-1 WHERE dvc_cencos = '" & cencos & "' AND dvc_fecing = cdate('" & Fecha & "') AND dvc_numlin > " & vaSpread1.ActiveRow
        Else
           vg_db.Execute "DELETE FROM b_detventascaf WHERE dvc_cencos = '" & cencos & "' AND dvc_fecing = '" & Format(Fecha, "yyyymmdd") & "' AND dvc_numlin = " & vaSpread1.ActiveRow
           vg_db.Execute "UPDATE b_detventascaf SET dvc_numlin=dvc_numlin-1 WHERE dvc_cencos = '" & cencos & "' AND dvc_fecing = '" & Format(Fecha, "yyyymmdd") & "' AND dvc_numlin > " & vaSpread1.ActiveRow
        End If
              
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        
        SumaCantidades True
        
        '-------> Detalle Productos
        If vg_tipbase = "1" Then
           vg_db.Execute "DELETE FROM b_detventascafpro WHERE dvp_cencos = '" & cencos & "' AND dvp_fecing = cdate('" & Fecha & "')"
        Else
           vg_db.Execute "DELETE FROM b_detventascafpro WHERE dvp_cencos = '" & cencos & "' AND dvp_fecing = '" & Format(Fecha, "yyyymmdd") & "'"
        End If
        For i = 1 To vaSpread2.MaxRows
            vaSpread2.Row = i
            vaSpread2.Col = 1:  cCPr = Trim(vaSpread2.text)
            vaSpread2.Col = 3:  cCCa = Format(vaSpread2.text, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
            vaSpread2.Col = 4:  cCDi = Format(vaSpread2.text, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
            vaSpread2.Col = 5:  cPCo = Format(vaSpread2.text, fg_Pict(9, vg_DPr))
            vaSpread2.Col = 6:  cAso = Trim(vaSpread2.text)
            If vg_tipbase = "1" Then
               vg_db.Execute "INSERT INTO b_detventascafpro (dvp_cencos, dvp_fecing, dvp_codmer, dvp_cancal, dvp_candig, dvp_precos) VALUES " & _
                             "('" & cencos & "', cdate('" & Fecha & "'), '" & cCPr & "', " & cCCa & ", " & cCDi & ", " & cPCo & ")"
            Else
               vg_db.Execute "INSERT INTO b_detventascafpro (dvp_cencos, dvp_fecing, dvp_codmer, dvp_cancal, dvp_candig, dvp_precos) VALUES " & _
                             "('" & cencos & "', '" & Format(Fecha, "yyyymmdd") & "', '" & cCPr & "', " & cCCa & ", " & cCDi & ", " & cPCo & ")"
            End If
        Next i
        '-------> Encabezado
        If vaSpread1.MaxRows = 0 Then
           If vg_tipbase = "1" Then
              vg_db.Execute "DELETE FROM b_totventascaf  WHERE tvc_cencos = '" & cencos & "' AND tvc_fecing = cdate('" & Fecha & "') AND tvc_codbod = " & vg_codbod & ""
           Else
              vg_db.Execute "DELETE FROM b_totventascaf  WHERE tvc_cencos = '" & cencos & "' AND tvc_fecing = '" & Format(Fecha, "yyyymmdd") & "' AND tvc_codbod = " & vg_codbod & ""
           End If
        End If
    vg_db.CommitTrans
    MoverDatos
    est = False
    modo = "": Gl_Ac_Botones Me, 11, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
    Toolbar1.Enabled = True
Case 7 '-------> Actualiza Grilla
    Toolbar1.Enabled = False
    est = True
    MoverDatos
    est = False
    modo = "": Gl_Ac_Botones Me, 11, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
    Toolbar1.Enabled = True
Case 10 '-------> Cancela
    If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    Toolbar1.Enabled = False
    est = True
    If modo = "A" Then
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    End If
    MoverDatos
    est = False
    modo = "": Gl_Ac_Botones Me, 11, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
    Toolbar1.Enabled = True
Case 12 '-------> Graba
    '-------> Validar si el contrato tiene asignado inventario rotativo
    If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
       Format(fpDateTime1(0).Value, "dd/mm/yyyy") > CDate(vg_ciedia) Then MsgBox "Tiene que realizar cierre diario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then MsgBox "Documento no corresponde al periodo : " & VgLinea & VgLinea & CierreFecha, vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 8) Then MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    est = True
    If GrabaDatos(0) Then est = False: Toolbar1.Enabled = True: Exit Sub
    modo = "": Gl_Ac_Botones Me, 11, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
    MoverDatos
    est = False
    Toolbar1.Enabled = True
Case 15 '-------> Imprimir
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Toolbar1.Enabled = False
    I_VentaCafeteria Me
    Toolbar1.Enabled = True
Case 18 '-------> Cerrar venta
    '-------> Validar si el contrato tiene asignado inventario rotativo
    If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
       Format(fpDateTime1(0).Value, "dd/mm/yyyy") > CDate(vg_ciedia) Then MsgBox "Tiene que realizar cierre diario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then MsgBox "Documento no corresponde al periodo : " & VgLinea & VgLinea & CierreFecha, vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 8) Then MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Toolbar1.Enabled = False
    est = True
    If GrabaDatos(1) Then est = False: Toolbar1.Enabled = True: Exit Sub
    modo = "": Gl_Ac_Botones Me, 11, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
    MoverDatos
    est = False
    Toolbar1.Enabled = True
Case 19 '-------> Reabrir venta
    '-------> Validar si el contrato tiene asignado inventario rotativo
    If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
       Format(fpDateTime1(0).Value, "dd/mm/yyyy") > CDate(vg_ciedia) Then MsgBox "Tiene que realizar cierre diario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then MsgBox "Documento no corresponde al periodo : " & VgLinea & VgLinea & CierreFecha, vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 8) Then MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Toolbar1.Enabled = False
    est = True
    modo = "R"
    If GrabaDatos(2) Then est = False: Toolbar1.Enabled = True: Exit Sub
    modo = "": Gl_Ac_Botones Me, 11, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
    MoverDatos
    est = False
    Toolbar1.Enabled = True
Case 21 '-------> Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Function GrabaDatos(op As Integer) As Boolean
Dim cencos As String, codbod As Long, Fecha As Date, rutcli As String, cencli As String
Dim cCAr As String, cCIn As Double, cPVe As Double, cTip As String, cNli As Long
Dim cCPr As String, cCCa As Double, cCDi As Double, cPCo As Double, cAso As String, cCal As Double, precio As Double
Dim sql1 As String
Dim RS1 As New ADODB.Recordset
On Error GoTo Man_Error
GrabaDatos = True
cencos = Trim(LimpiaDato(fpText1(0).text))
Fecha = Format(fpDateTime1(0).text, "dd/mm/yyyy")
codbod = Val(fg_codigocbo(Combo1, 0, 10, ""))
If InStr(Trim(fpText1(1).text), "-") = 0 And Trim(fpText1(1).text) <> "" Then
    fpText1(1).text = fg_RutDig(Trim(fpText1(1).text))
    fpText1(1).text = fg_PintaRut(Trim(fpText1(1).text))
End If
rutcli = fg_DespintaRut(Trim(LimpiaDato(fpText1(1).text)))
cencli = Trim(LimpiaDato(fpText1(2).text))

If cencos = "" Or codbod = 0 Or Trim(fpDateTime1(0).text) = "" Then MsgBox "Debe ingresar dato en el encabezado...", vbExclamation + vbOKOnly, MsgTitulo: est = False: Exit Function

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: cCAr = Trim(vaSpread1.text)
vaSpread1.Col = 3: cCIn = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
vaSpread1.Col = 4: cPVe = Format(vaSpread1.text, fg_Pict(9, vg_DPr))
vaSpread1.Col = 5: cTip = Left(Trim(vaSpread1.text), 2)

If Val(cCAr) = 0 Then MsgBox "Debe ingresar artículo...", vbExclamation + vbOKOnly, MsgTitulo: SSTab1.Tab = 0: vaSpread1.SetActiveCell 1, vaSpread1.Row: est = False: Exit Function
If cCIn = 0 Then MsgBox "La cantidad debe ser mayor a cero...", vbExclamation + vbOKOnly, MsgTitulo: SSTab1.Tab = 0: vaSpread1.SetActiveCell 3, vaSpread1.Row: est = False: Exit Function
If cPVe = 0 Then MsgBox "El precio debe ser mayor a cero...", vbExclamation + vbOKOnly, MsgTitulo: SSTab1.Tab = 0: vaSpread1.SetActiveCell 4, vaSpread1.Row: est = False: Exit Function
If Trim(cTip) = "" Then MsgBox "Debe seleccionar tipo pago...", vbExclamation + vbOKOnly, MsgTitulo: SSTab1.Tab = 0: vaSpread1.SetActiveCell 5, vaSpread1.Row: est = False: Exit Function
If (Trim(cTip) = "CR" Or Trim(cTip) = "CU") And Trim(rutcli) = "" Then MsgBox "Debe Ingresar cliente...", vbExclamation + vbOKOnly, MsgTitulo: SSTab1.Tab = 0: est = False: Exit Function
If vaSpread1.MaxRows = 0 Then MsgBox "Debe ingresar por lo menos un articulo...", vbExclamation + vbOKOnly, MsgTitulo: SSTab1.Tab = 0: est = False: Exit Function
If vaSpread2.MaxRows = 0 Then MsgBox "No hay productos en la composición...", vbExclamation + vbOKOnly, MsgTitulo: SSTab1.Tab = 0: est = False: Exit Function

For i = 1 To vaSpread2.MaxRows
    vaSpread2.Row = i
    vaSpread2.Col = 3
    If Format(IIf(vaSpread2.text = "", 0, vaSpread2.text), fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa))) = 0 Then MsgBox "La cantidad debe ser mayor a cero...", vbExclamation + vbOKOnly, MsgTitulo: SSTab1.Tab = 1: vaSpread2.SetActiveCell 3, vaSpread2.Row: est = False: Exit Function
    vaSpread2.Col = 5
    If Format(IIf(vaSpread2.text = "", 0, vaSpread2.text), fg_Pict(9, vg_DPr)) = 0 Then MsgBox "El precio debe ser mayor a cero...", vbExclamation + vbOKOnly, MsgTitulo: SSTab1.Tab = 1: vaSpread2.SetActiveCell 5, vaSpread2.Row: est = False: Exit Function
Next i

For i = 1 To vaSpread2.MaxRows
    vaSpread2.Row = i
    vaSpread2.Col = 3
    If Format(IIf(vaSpread2.text = "", 0, vaSpread2.text), fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa))) = 0 Then MsgBox "La cantidad debe ser mayor a cero...", vbExclamation + vbOKOnly, MsgTitulo: SSTab1.Tab = 1: vaSpread2.SetActiveCell 3, vaSpread2.Row: est = False: Exit Function
    vaSpread2.Col = 5
    If Format(IIf(vaSpread2.text = "", 0, vaSpread2.text), fg_Pict(9, vg_DPr)) = 0 Then MsgBox "El precio debe ser mayor a cero...", vbExclamation + vbOKOnly, MsgTitulo: SSTab1.Tab = 1: vaSpread2.SetActiveCell 5, vaSpread2.Row: est = False: Exit Function
    If op = 1 Then
        vaSpread2.Col = 4
        If Format(IIf(vaSpread2.text = "", 0, vaSpread2.text), fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa))) = 0 Then MsgBox "La cantidad debe ser mayor a cero...", vbExclamation + vbOKOnly, MsgTitulo: SSTab1.Tab = 1: vaSpread2.SetActiveCell 4, vaSpread2.Row: est = False: Exit Function
        vaSpread2.Col = 7
        If Trim(vaSpread2.text) = "S" Then MsgBox "Cantidad exede el Stock......", vbExclamation + vbOKOnly, MsgTitulo: SSTab1.Tab = 1: vaSpread2.SetActiveCell 4, vaSpread2.Row: est = False: Exit Function
    End If
Next i
If op = 1 Then
    If MsgBox("Al cerrar ud. no podrá ingresar más ventas para este día. Desea cerrar documento...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then est = False: Exit Function
ElseIf op = 2 Then
    If MsgBox("Desea reabrir...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then est = False: Exit Function
Else
    If MsgBox("Desea grabar...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then est = False: Exit Function
End If
If modo = "A" Then
    'Encabezado
    vg_db.BeginTrans
    If vg_tipbase = "1" Then
       RS1.Open "SELECT 1 FROM b_totventascaf WHERE tvc_cencos = '" & cencos & "' AND tvc_fecing = cdate('" & Fecha & "') AND tvc_codbod=" & vg_codbod & "", vg_db, adOpenStatic
       If RS1.EOF Then
          vg_db.Execute "INSERT INTO b_totventascaf (tvc_cencos, tvc_fecing, tvc_codbod, tvc_estado) VALUES " & _
                        "('" & cencos & "', cdate('" & Fecha & "'), " & codbod & ", '')"
       End If
    Else
       RS1.Open "SELECT 1 FROM b_totventascaf WHERE tvc_cencos = '" & cencos & "' AND tvc_fecing = '" & Format(Fecha, "yyyymmdd") & "' AND tvc_codbod = " & vg_codbod & "", vg_db, adOpenStatic
       If RS1.EOF Then
          vg_db.Execute "INSERT INTO b_totventascaf (tvc_cencos, tvc_fecing, tvc_codbod, tvc_estado) VALUES " & _
                        "('" & cencos & "', '" & Format(Fecha, "yyyymmdd") & "', " & codbod & ", '')"
       End If
    End If
    RS1.Close: Set RS1 = Nothing
    
    cNli = 0
    If vg_tipbase = "1" Then
       RS1.Open "SELECT MAX(dvc_numlin) AS maxlin FROM b_detventascaf WHERE dvc_cencos = '" & cencos & "' AND dvc_fecing = cdate('" & Fecha & "')", vg_db, adOpenStatic
    Else
       RS1.Open "SELECT MAX(dvc_numlin) AS maxlin FROM b_detventascaf WHERE dvc_cencos = '" & cencos & "' AND dvc_fecing = '" & Format(Fecha, "yyyymmdd") & "'", vg_db, adOpenStatic
    End If
    If Not RS1.EOF Then
        cNli = IIf(IsNull(RS1!maxlin), 0, RS1!maxlin) + 1
    Else
        cNli = 1
    End If
    RS1.Close: Set RS1 = Nothing
       
    'Detalle Articulos
    If vg_tipbase = "1" Then
       vg_db.Execute "INSERT INTO b_detventascaf (dvc_cencos, dvc_fecing, dvc_numlin, dvc_articulo, dvc_canart, dvc_precio, dvc_tippag, dvc_rutcli, dvc_cencli, dvc_tipdoc, dvc_numdoc, dvc_fecdoc) VALUES " & _
                     "('" & cencos & "', cdate('" & Fecha & "'), " & cNli & ", '" & cCAr & "', " & cCIn & ", " & cPVe & ", '" & cTip & "', '" & rutcli & "', '" & cencli & "', '', 0, cdate('0'))"
    Else
       vg_db.Execute "INSERT INTO b_detventascaf (dvc_cencos, dvc_fecing, dvc_numlin, dvc_articulo, dvc_canart, dvc_precio, dvc_tippag, dvc_rutcli, dvc_cencli, dvc_tipdoc, dvc_numdoc, dvc_fecdoc) VALUES " & _
                     "('" & cencos & "', '" & Format(Fecha, "yyyymmdd") & "', " & cNli & ", '" & cCAr & "', " & cCIn & ", " & cPVe & ", '" & cTip & "', '" & rutcli & "', '" & cencli & "', '', 0, 0)"
    End If
     'Detalle Productos
    If vg_tipbase = "1" Then
       vg_db.Execute "DELETE FROM b_detventascafpro WHERE dvp_cencos = '" & cencos & "' AND dvp_fecing = cdate('" & Fecha & "')"
    Else
       vg_db.Execute "DELETE FROM b_detventascafpro WHERE dvp_cencos = '" & cencos & "' AND dvp_fecing = '" & Format(Fecha, "yyyymmdd") & "'"
    End If
    For i = 1 To vaSpread2.MaxRows
        vaSpread2.Row = i
        vaSpread2.Col = 1:  cCPr = Trim(vaSpread2.text)
        vaSpread2.Col = 3:  cCCa = Format(vaSpread2.text, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
        vaSpread2.Col = 4:  cCDi = Format(vaSpread2.text, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
        vaSpread2.Col = 5:  cPCo = Format(vaSpread2.text, fg_Pict(9, vg_DPr))
        vaSpread2.Col = 6:  cAso = Trim(vaSpread2.text)
        If vg_tipbase = "1" Then
           vg_db.Execute "INSERT INTO b_detventascafpro (dvp_cencos, dvp_fecing, dvp_codmer, dvp_cancal, dvp_candig, dvp_precos) VALUES " & _
                         "('" & cencos & "', cdate('" & Fecha & "'), '" & cCPr & "', " & cCCa & ", " & cCDi & ", " & cPCo & ")"
        Else
           vg_db.Execute "INSERT INTO b_detventascafpro (dvp_cencos, dvp_fecing, dvp_codmer, dvp_cancal, dvp_candig, dvp_precos) VALUES " & _
                         "('" & cencos & "', '" & Format(Fecha, "yyyymmdd") & "', '" & cCPr & "', " & cCCa & ", " & cCDi & ", " & cPCo & ")"
        End If
    Next i
    vg_db.CommitTrans
ElseIf modo = "M" Then
    vg_db.BeginTrans
    If vg_tipbase = "1" Then
       'Encabezado
       vg_db.Execute "UPDATE b_totventascaf  SET tvc_codbod = " & codbod & ", tvc_estado = '' WHERE tvc_cencos = '" & cencos & "' AND tvc_fecing = cdate('" & Fecha & "' AND tvc_codbod = " & vg_codbod & ")"
       'Detalle Articulos
       vg_db.Execute "UPDATE b_detventascaf SET dvc_articulo = '" & cCAr & "', dvc_canart = " & cCIn & ", " & _
                     "dvc_precio = " & cPVe & ", dvc_tippag = '" & cTip & "', dvc_rutcli = '" & rutcli & "', " & _
                     "dvc_cencli = '" & cencli & "', dvc_tipdoc = '', dvc_numdoc = 0, dvc_fecdoc = cdate('0') WHERE dvc_cencos = '" & cencos & "' AND dvc_fecing = cdate('" & Fecha & "') AND dvc_numlin = " & vaSpread1.ActiveRow
       'Detalle Productos
       vg_db.Execute "DELETE FROM b_detventascafpro WHERE dvp_cencos = '" & cencos & "' AND dvp_fecing = cdate('" & Fecha & "')"
    Else
       'Encabezado
       vg_db.Execute "UPDATE b_totventascaf  SET tvc_codbod = " & codbod & ", tvc_estado = '' WHERE tvc_cencos = '" & cencos & "' AND tvc_fecing = '" & Format(Fecha, "yyyymmdd") & "' AND tvc_codbod = " & vg_codbod & ""
       'Detalle Articulos
       vg_db.Execute "UPDATE b_detventascaf SET dvc_articulo = '" & cCAr & "', dvc_canart = " & cCIn & ", " & _
                     "dvc_precio = " & cPVe & ", dvc_tippag = '" & cTip & "', dvc_rutcli = '" & rutcli & "', " & _
                     "dvc_cencli = '" & cencli & "', dvc_tipdoc = '', dvc_numdoc = 0, dvc_fecdoc = 0 WHERE dvc_cencos = '" & cencos & "' AND dvc_fecing = '" & Format(Fecha, "yyyymmdd") & "' AND dvc_numlin = " & vaSpread1.ActiveRow
       'Detalle Productos
       vg_db.Execute "DELETE FROM b_detventascafpro WHERE dvp_cencos = '" & cencos & "' AND dvp_fecing = '" & Format(Fecha, "yyyymmdd") & "'"
    End If
    For i = 1 To vaSpread2.MaxRows
        vaSpread2.Row = i
        vaSpread2.Col = 1:  cCPr = Trim(vaSpread2.text)
        vaSpread2.Col = 3:  cCCa = Format(vaSpread2.text, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
        vaSpread2.Col = 4:  cCDi = Format(vaSpread2.text, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
        vaSpread2.Col = 5:  cPCo = Format(vaSpread2.text, fg_Pict(9, vg_DPr))
        vaSpread2.Col = 6:  cAso = Trim(vaSpread2.text)
        If vg_tipbase = "1" Then
           vg_db.Execute "INSERT INTO b_detventascafpro (dvp_cencos, dvp_fecing, dvp_codmer, dvp_cancal, dvp_candig, dvp_precos) VALUES " & _
                         "('" & cencos & "', cdate('" & Fecha & "'), '" & cCPr & "', " & cCCa & ", " & cCDi & ", " & cPCo & ")"
        Else
           vg_db.Execute "INSERT INTO b_detventascafpro (dvp_cencos, dvp_fecing, dvp_codmer, dvp_cancal, dvp_candig, dvp_precos) VALUES " & _
                         "('" & cencos & "', '" & Format(Fecha, "yyyymmdd") & "', '" & cCPr & "', " & cCCa & ", " & cCDi & ", " & cPCo & ")"
        End If
    Next i
    vg_db.CommitTrans
End If

If op = 1 Then
    cencos = Trim(LimpiaDato(fpText1(0).text))
    Fecha = Format(fpDateTime1(0).text, "dd/mm/yyyy")
    codbod = Val(fg_codigocbo(Combo1, 0, 10, ""))
    
    vg_db.BeginTrans
        If vg_tipbase = "1" Then
           vg_db.Execute "UPDATE b_totventascaf SET tvc_estado = 'C' WHERE tvc_cencos = '" & cencos & "' AND tvc_fecing = cdate('" & Fecha & "') AND tvc_codbod = " & vg_codbod & ""
        Else
           vg_db.Execute "UPDATE b_totventascaf SET tvc_estado = 'C' WHERE tvc_cencos = '" & cencos & "' AND tvc_fecing = '" & Format(Fecha, "yyyymmdd") & "' AND tvc_codbod = " & vg_codbod & ""
        End If
        'Control de Stock
        For i = 1 To vaSpread2.MaxRows
            cCal = 0
            vaSpread2.Row = i
            vaSpread2.Col = 1:  cCPr = Trim(vaSpread2.text)
            vaSpread2.Col = 4:  cCDi = Format(vaSpread2.text, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
            vaSpread2.Col = 5: precio = vaSpread2.text
            '-------> Actualizar
            sql1 = IIf(vg_tipbase = "1", " cdate('" & Fecha & "') ", "  '" & Format(Fecha, "yyyymmdd") & "' ")
            vg_db.Execute "UPDATE b_detventascafpro SET dvp_cancal = " & cCDi & ", dvp_candig =  " & cCDi & ", dvp_precos = " & precio & " WHERE dvp_cencos = '" & cencos & "' AND dvp_fecing = " & sql1 & " AND dvp_codmer = '" & cCPr & "'"

            RS1.Open "SELECT bod_canmer FROM b_bodegas WHERE bod_codpro = '" & Trim(LimpiaDato(cCPr)) & "' AND bod_codbod = " & vg_codbod, vg_db, adOpenStatic
            If Not RS1.EOF Then
                vg_db.Execute "UPDATE b_bodegas SET bod_canmer = bod_canmer-" & cCDi & " " & _
                              "WHERE bod_codpro = '" & Trim(LimpiaDato(cCPr)) & "' AND bod_codbod = " & vg_codbod
            End If
            RS1.Close: Set RS1 = Nothing
        Next i
        
    vg_db.CommitTrans
ElseIf op = 2 Then
    cencos = Trim(LimpiaDato(fpText1(0).text))
    Fecha = Format(fpDateTime1(0).text, "dd/mm/yyyy")
    codbod = Val(fg_codigocbo(Combo1, 0, 10, ""))
    vg_db.BeginTrans
        If vg_tipbase = "1" Then
           vg_db.Execute "UPDATE b_totventascaf SET tvc_estado = '' WHERE tvc_cencos = '" & cencos & "' AND tvc_fecing = cdate('" & Fecha & "') AND tvc_codbod = " & vg_codbod & ""
        Else
           vg_db.Execute "UPDATE b_totventascaf SET tvc_estado = '' WHERE tvc_cencos = '" & cencos & "' AND tvc_fecing = '" & Format(Fecha, "yyyymmdd") & "' AND tvc_codbod = " & vg_codbod & ""
        End If
        'Control de Stock
        For i = 1 To vaSpread2.MaxRows
            cCal = 0
            vaSpread2.Row = i
            vaSpread2.Col = 1:  cCPr = Trim(vaSpread2.text)
            vaSpread2.Col = 4:  cCDi = Format(vaSpread2.text, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
            RS1.Open "SELECT bod_canmer FROM b_bodegas WHERE bod_codpro = '" & Trim(LimpiaDato(cCPr)) & "' AND bod_codbod = " & vg_codbod, vg_db, adOpenStatic
            If Not RS1.EOF Then
               vg_db.Execute "UPDATE b_bodegas SET bod_canmer = bod_canmer+" & cCDi & " WHERE bod_codpro = '" & Trim(LimpiaDato(cCPr)) & "' AND bod_codbod = " & vg_codbod
            End If
            RS1.Close: Set RS1 = Nothing
        Next i
    vg_db.CommitTrans
End If
GrabaDatos = False
Exit Function
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Function
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Function

Private Sub Limpia(op As Integer)
If op = 1 Then
    fpDateTime1(0).text = ""
    Combo1(0).ListIndex = IIf(Combo1(0).listcount = 1, 0, -1)
End If
If op = 2 Then
    Combo1(0).ListIndex = IIf(Combo1(0).listcount = 1, 0, -1)
End If
Label1.Caption = ""
vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.Lock = False
vaSpread2.Row = -1
vaSpread2.Col = -1: vaSpread2.Lock = False
vaSpread2.Col = 3: vaSpread2.Lock = True
vaSpread2.Col = 5: vaSpread2.Lock = True
Image1(0).Enabled = ModCasino
fpText1(0).Enabled = ModCasino
fpText1(0).text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
fpText1(1).text = ""
fpayuda(1).Caption = ""
fpText1(2).text = ""
fpLongInteger1(0).text = ""
fpDateTime1(1).text = ""
Gl_Ac_Botones Me, 11, 3, modo
Frame2.Enabled = False
End Sub

Private Sub MoverDatos()
Dim cencos As String, Fecha As Date, codbod As Long, cArt As String, cTip As String
Dim RS1 As New ADODB.Recordset, RS2 As New ADODB.Recordset, RS3 As New ADODB.Recordset, i As Long
On Error GoTo Man_Error
Image1(1).Enabled = True
vaSpread1.Visible = False: vaSpread2.Visible = False
vaSpread1.MaxRows = 0: vaSpread2.MaxRows = 0
If Trim(LimpiaDato(fpText1(0).text)) = "" Or Val(fg_codigocbo(Combo1, 0, 10, "")) = 0 Or Trim(fpDateTime1(0).text) = "" Then Exit Sub
cencos = Trim(LimpiaDato(fpText1(0).text))
Fecha = Format(fpDateTime1(0).text, "dd/mm/yyyy")
codbod = Val(fg_codigocbo(Combo1, 0, 10, ""))
If vg_tipbase = "1" Then
   RS1.Open "SELECT * FROM b_totventascaf WHERE tvc_cencos = '" & cencos & "' AND tvc_fecing = cdate('" & Fecha & "') AND tvc_codbod = " & vg_codbod & "", vg_db, adOpenStatic
Else
   RS1.Open "SELECT * FROM b_totventascaf WHERE tvc_cencos = '" & cencos & "' AND tvc_fecing = '" & Format(Fecha, "yyyymmdd") & "' AND tvc_codbod = " & vg_codbod & "", vg_db, adOpenStatic
End If
If Not RS1.EOF Then
    Combo1(0).ListIndex = fg_buscacbo(Combo1, 0, 10, fg_pone_cero(Str(RS1!tvc_codbod), 10))
    If vg_tipbase = "1" Then
       RS2.Open "SELECT a.*, b.tpc_nombre FROM b_detventascaf a, b_totpreciocaf b WHERE a.dvc_articulo = tpc_codigo " & _
                "AND a.dvc_cencos = '" & cencos & "' AND a.dvc_fecing = cdate('" & Fecha & "') AND b.tpc_cencos = '" & MuestraCasino(1) & "' ORDER BY a.dvc_numlin", vg_db, adOpenStatic
    Else
       RS2.Open "SELECT a.*, b.tpc_nombre FROM b_detventascaf a, b_totpreciocaf b WHERE a.dvc_articulo = tpc_codigo " & _
                "AND a.dvc_cencos = '" & cencos & "' AND a.dvc_fecing = '" & Format(Fecha, "yyyymmdd") & "' AND b.tpc_cencos = '" & MuestraCasino(1) & "' ORDER BY a.dvc_numlin", vg_db, adOpenStatic
    End If
    If Not RS2.EOF Then
        i = 1
        Do While Not RS2.EOF
            vaSpread1.MaxRows = i
            vaSpread1.Row = vaSpread1.MaxRows
            vaSpread1.Col = 1: vaSpread1.text = Trim(IIf(IsNull(RS2!dvc_articulo), "", RS2!dvc_articulo))
            vaSpread1.Col = 2: vaSpread1.text = Trim(IIf(IsNull(RS2!tpc_nombre), "", RS2!tpc_nombre))
            vaSpread1.Col = 3: vaSpread1.text = Trim(IIf(IsNull(RS2!dvc_canart), "", RS2!dvc_canart))
            vaSpread1.Col = 4: vaSpread1.text = Trim(IIf(IsNull(RS2!dvc_precio), "", RS2!dvc_precio))
            vaSpread1.Col = 5: vaSpread1.TypeComboBoxList = "CONTADO" & Chr$(9) & "CREDITO" & Chr$(9) & "CUENTA ABONO CLIENTE"
            'mod jpaz vaSpread1.Col = 5: vaSpread1.TypeComboBoxCurSel = Trim(IIf(IsNull(RS2!dvc_tippag), -1, IIf(RS2!dvc_tippag = "CO", 0, 1))) 'cTip = Left(Trim(vaSpread1.Text), 2)
            vaSpread1.Col = 5: vaSpread1.TypeComboBoxCurSel = Trim(IIf(IsNull(RS2!dvc_tippag), -1, IIf(RS2!dvc_tippag = "CO", 0, IIf(RS2!dvc_tippag = "CR", 1, 2)))) 'cTip = Left(Trim(vaSpread1.Text), 2)
            vaSpread1.Col = 6: vaSpread1.text = fg_DespintaRut(Trim(IIf(IsNull(RS2!dvc_rutcli), "", RS2!dvc_rutcli)))
            vaSpread1.Col = 7: vaSpread1.text = Trim(IIf(IsNull(RS2!dvc_cencli), "", RS2!dvc_cencli))
            RS2.MoveNext: i = i + 1
        Loop
    End If
    RS2.Close: Set RS2 = Nothing
    SSTab1.Tab = 0
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Or CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Or _
       CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 8) Or RS1!tvc_estado = "C" Then
        If vg_tipbase = "1" Then
           RS2.Open "SELECT dvp.*, pro.pro_nombre FROM b_detventascafpro dvp, b_productos pro " & _
                    "WHERE pro.pro_codigo = dvp.dvp_codmer AND dvp_cencos = '" & cencos & "' AND dvp_fecing = cdate('" & Fecha & "') ORDER BY dvp_codmer", vg_db, adOpenStatic
        Else
           RS2.Open "SELECT dvp.*, pro.pro_nombre FROM b_detventascafpro dvp, b_productos pro " & _
                    "WHERE pro.pro_codigo = dvp.dvp_codmer AND dvp_cencos = '" & cencos & "' AND dvp_fecing = '" & Format(Fecha, "yyyymmdd") & "' ORDER BY dvp_codmer", vg_db, adOpenStatic
        End If
        vaSpread2.Row = 0
        If Not RS2.EOF Then
            Do While Not RS2.EOF
                vaSpread2.MaxRows = vaSpread2.MaxRows + 1: vaSpread2.Row = vaSpread2.MaxRows
                vaSpread2.Col = 1: vaSpread2.text = RS2!dvp_codmer
                vaSpread2.Col = 2: vaSpread2.text = RS2!pro_nombre
                vaSpread2.Col = 3: vaSpread2.text = Format(RS2!dvp_cancal, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
                vaSpread2.Col = 4: vaSpread2.text = Format(RS2!dvp_candig, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
                vaSpread2.Col = 5: vaSpread2.text = Format(RS2!dvp_precos, fg_Pict(9, vg_DPr))
                vaSpread2.Col = 6: vaSpread2.text = ""
                vaSpread2.Col = 7: vaSpread2.text = ""
                vaSpread2.Col = 8: vaSpread2.text = ""
                RevisaSobreStock Trim(RS2!dvp_codmer), vaSpread2.Row, False
                RS2.MoveNext
            Loop
        End If
        RS2.Close: Set RS2 = Nothing
    Else
       SumaCantidades True
    End If
    
    If vaSpread1.MaxRows > 0 Then
        vaSpread1.SetActiveCell 3, 1
    
        'Muestra datos del cliente si es necesario
        vaSpread1.Row = 1: vaSpread1.Col = 5: cTip = vaSpread1.text: vaSpread1.Col = 6
        RS3.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & Trim(vaSpread1.text) & "' AND cli_tipo = " & IIf(Mid(cTip, 1, 2) = "CR", 1, 3) & "", vg_db, adOpenStatic
        If Not RS3.EOF Then
            fpText1(1).text = fg_PintaRut(Trim(IIf(IsNull(RS3!cli_codigo), "", RS3!cli_codigo)))
            fpayuda(1).Caption = Trim(IIf(IsNull(RS3!cli_nombre), "", RS3!cli_nombre))
        Else
            fpText1(1).text = ""
            fpayuda(1).Caption = ""
        End If
        RS3.Close: Set RS3 = Nothing
        vaSpread1.Col = 7
        fpText1(2).text = Trim(vaSpread1.text)
        vaSpread1.Col = 5
        Frame2.Enabled = (vaSpread1.TypeComboBoxCurSel = 1 And Label1.Caption = "")
    End If
    Gl_Ac_Botones Me, 11, 1, modo
    
    If RS1!tvc_estado = "C" Then
        vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.Lock = True
        vaSpread2.Row = -1: vaSpread2.Col = -1: vaSpread2.Lock = True
        Combo1(0).Enabled = False
        Frame2.Enabled = False
        Label1.Caption = "CERRADA"
        Image1(1).Enabled = False
    Else
        vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.Lock = False
        vaSpread2.Row = -1
        vaSpread2.Col = -1: vaSpread2.Lock = False
        vaSpread2.Col = 3: vaSpread2.Lock = True
        vaSpread2.Col = 5: vaSpread2.Lock = True
        Combo1(0).Enabled = True
        Label1.Caption = ""
        Image1(1).Enabled = True
    End If
    
Else
    Limpia 3
    If Trim(LimpiaDato(fpText1(0).text)) = "" Or Val(fg_codigocbo(Combo1, 0, 10, "")) = 0 Or Trim(fpDateTime1(0).text) = "" Then
        Gl_Ac_Botones Me, 11, 3, modo
    Else
        Gl_Ac_Botones Me, 11, 2, modo
    End If
End If
RS1.Close: Set RS1 = Nothing
vaSpread1.Visible = True: vaSpread2.Visible = True
Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
Toolbar1.Enabled = True
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub


Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
If est Then Exit Sub
If modo = "" Then modo = "M": Gl_Ac_Botones Me, 11, 0, modo
vaSpread1.Col = 5: vaSpread1.Row = Row
Frame2.Enabled = (vaSpread1.TypeComboBoxCurSel = 1 Or vaSpread1.TypeComboBoxCurSel = 2 And Label1.Caption = "")
If vaSpread1.TypeComboBoxCurSel = 0 Then
    fpText1(1).text = ""
    fpayuda(1).Caption = ""
    fpText1(2).text = ""
End If
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If est Then Exit Sub
If modo = "" Then modo = "M": Gl_Ac_Botones Me, 11, 0, modo
Select Case Col
Case 3
    est = True
    SumaCantidades True
    est = False
End Select
End Sub

Private Sub SumaCantidades(ConProductos As Boolean)
Dim cCAc As Double, cCIn As Double, cCPr As Double, precio As Double, cPro As String, cArt As String, cArt2 As String, i As Long, X As Long, propon As Double
Dim cencos As String, Fecha As Date, codbod As Long
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
On Error GoTo Man_Error
cencos = Trim(LimpiaDato(fpText1(0).text))
Fecha = Format(fpDateTime1(0).text, "dd/mm/yyyy")
'Suma cantidad ingresada en la venta
If ConProductos Then
    vaSpread2.MaxRows = 0
    For X = 1 To vaSpread1.MaxRows
        vaSpread1.Row = X: vaSpread1.Col = 1: codigo = Val(vaSpread1.text)
        If vg_tipbase = "1" Then
           RS1.Open "SELECT dpc.*, pro.pro_nombre, (SELECT dvp.dvp_precos FROM b_detventascafpro AS dvp WHERE dvp.dvp_cencos = '" & cencos & "' AND dvp.dvp_fecing = cdate('" & Fecha & "') AND dvp.dvp_codmer = pro.pro_codigo) AS dvp_precos, " & _
                    "(SELECT dvp.dvp_candig FROM b_detventascafpro AS dvp WHERE dvp.dvp_cencos = '" & cencos & "' AND dvp.dvp_fecing = cdate('" & Fecha & "') AND dvp.dvp_codmer = pro.pro_codigo) AS dvp_candig " & _
                    "FROM b_detpreciocaf dpc, b_productos pro WHERE pro.pro_codigo = dpc.dpc_codmer AND dpc.dpc_cencos = '" & MuestraCasino(1) & "' AND dpc.dpc_codigo = '" & codigo & "'", vg_db, adOpenStatic
        Else
           RS1.Open "SELECT dpc.*, pro.pro_nombre, (SELECT dvp.dvp_precos FROM b_detventascafpro AS dvp WHERE dvp.dvp_cencos = '" & cencos & "' AND dvp.dvp_fecing = '" & Format(Fecha, "yyyymmdd") & "' AND dvp.dvp_codmer = pro.pro_codigo) AS dvp_precos, " & _
                    "(SELECT dvp.dvp_candig FROM b_detventascafpro AS dvp WHERE dvp.dvp_cencos = '" & cencos & "' AND dvp.dvp_fecing = '" & Format(Fecha, "yyyymmdd") & "' AND dvp.dvp_codmer = pro.pro_codigo) AS dvp_candig " & _
                    "FROM b_detpreciocaf dpc, b_productos pro WHERE pro.pro_codigo = dpc.dpc_codmer AND dpc.dpc_cencos = '" & MuestraCasino(1) & "' AND dpc.dpc_codigo = '" & codigo & "'", vg_db, adOpenStatic
        End If
        If Not RS1.EOF Then
            Do While Not RS1.EOF
                '-------> Busca si existe el producto, y si no, agrega fila
                vaSpread2.Row = 0
                For i = 1 To vaSpread2.MaxRows
                    vaSpread2.Col = 1: vaSpread2.Row = i
                    If Trim(vaSpread2.text) = Trim(RS1!dpc_codmer) Then Exit For
                    vaSpread2.Row = 0
                Next i
                If vaSpread2.Row = 0 Then vaSpread2.MaxRows = vaSpread2.MaxRows + 1: vaSpread2.Row = vaSpread2.MaxRows
                vaSpread2.Col = 1: vaSpread2.text = RS1!dpc_codmer
                vaSpread2.Col = 2: vaSpread2.text = RS1!pro_nombre
                vaSpread2.Col = 3: vaSpread2.text = 0
                vaSpread2.Col = 4: vaSpread2.text = 0 'IIf(IsNull(RS1!dvp_candig), 0, RS1!dvp_candig)
                propon = 0
                RS2.Open "SELECT TOP 1 ppd_cencos, ppd_codpro, ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                         "FROM b_productospmpdia " & _
                         "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                         "AND   ppd_codpro = '" & RS1!dpc_codmer & "' " & _
                         "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(fpDateTime1(0).text), "yyyymmdd") & " " & _
                         "GROUP BY ppd_cencos, ppd_codpro, ppd_propon " & _
                         "HAVING (ppd_propon)>0 ORDER BY Max(ppd_fecdia) DESC", vg_db, adOpenStatic
                If Not RS2.EOF Then propon = RS2!ppd_propon
                RS2.Close: Set RS2 = Nothing
                vaSpread2.Col = 5: vaSpread2.text = IIf(IsNull(RS1!dvp_precos), IIf(IsNull(propon), 0, propon), RS1!dvp_precos)
                vaSpread2.Col = 6: vaSpread2.text = vaSpread2.text & IIf(ArticuloRepetido(Trim(codigo), vaSpread2.Row), "", Trim(codigo) & "&" & Trim(Str(RS1!dpc_cantidad)) & ";")
                'RevisaSobreStock Trim(RS1!dpc_codmer), vaSpread2.Row, True
                RS1.MoveNext
            Loop
            
        End If
        RS1.Close: Set RS1 = Nothing
    Next X
    
End If
For X = 1 To vaSpread1.MaxRows
    vaSpread1.Row = X
    vaSpread1.Col = 1: cArt = Trim(vaSpread1.text)
    vaSpread1.Col = 3
    If Trim(vaSpread1.text) <> "" Then
        cCIn = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
        For i = 1 To vaSpread2.MaxRows
            vaSpread2.Row = i: vaSpread2.Col = 6
            StrImp = Trim(vaSpread2.text): cCAc = 0
            If Len(StrImp) <> 0 Then
                Do While InStr(StrImp, ";") <> 0
                    StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
                    StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
                    cArt2 = Trim(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
                    cCPr = Format(Mid(StrImpb, 1), fg_Pict(9, 3))
                    If cArt = cArt2 Then
                        cCAc = cCAc + (cCPr * cCIn)
                    End If
                Loop
            End If
            vaSpread2.Col = 1: cPro = Trim(vaSpread2.text)
            vaSpread2.Col = 3: vaSpread2.text = Format(IIf(Trim(vaSpread2.text) = "", 0, vaSpread2.text), fg_Pict(9, 3)) + cCAc
            vaSpread2.Col = 4: vaSpread2.text = Format(IIf(Trim(vaSpread2.text) = "", 0, vaSpread2.text), fg_Pict(9, 3)) + cCAc
            RevisaSobreStock cPro, vaSpread2.Row, True
        Next i
    End If
Next X
vaSpread2.Row = -1:
vaSpread2.Col = 3: vaSpread2.Lock = True
vaSpread2.Col = 5: vaSpread2.Lock = True
vaSpread2.SortKey(1) = 1
vaSpread2.SortKeyOrder(1) = 1
vaSpread2.Sort -1, -1, vaSpread2.MaxCols, vaSpread2.MaxRows, SortByRow
Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Function ArticuloRepetido(cArt As String, Row As Long) As Boolean
Dim cArt2 As String, cCon As Long
ArticuloRepetido = False
vaSpread2.Row = Row: vaSpread2.Col = 6
StrImp = Trim(vaSpread2.text): cCon = 0
If Len(StrImp) <> 0 Then
    Do While InStr(StrImp, ";") <> 0
        StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
        StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
        cArt2 = Trim(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
        cCPr = Format(Mid(StrImpb, 1), fg_Pict(9, 3))
        If cArt = cArt2 Then
            cCon = cCon + 1
        End If
    Loop
End If
If cCon > 0 Then ArticuloRepetido = True
End Function

Private Sub RevisaSobreStock(cPro As String, Row As Long, color As Boolean)
Dim i As Long, cSto As Double, cCRe As Double, codbod As Long
Dim RS1 As New ADODB.Recordset
codbod = Val(fg_codigocbo(Combo1, 0, 10, ""))
vaSpread2.Row = Row
vaSpread2.Col = 4: cCRe = Format(vaSpread2.text, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
ValidaBod codbod, cPro
RS1.Open "SELECT bod_canmer FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " AND bod_codpro = '" & Trim(cPro) & "'", vg_db, adOpenStatic
If Not RS1.EOF Then
    cSto = Format(RS1!bod_canmer, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
    vaSpread2.Col = 8: vaSpread2.text = IIf(IsNull(RS1!bod_canmer), 0, Format(RS1!bod_canmer, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa))))
End If
RS1.Close: Set RS1 = Nothing
If color Then
    vaSpread2.Col = -1: vaSpread2.BackColor = Shape1(IIf((cSto - cCRe) < 0, 1, 2)).FillColor
    vaSpread2.Col = 7: vaSpread2.text = IIf((cSto - cCRe) < 0, "S", "N")
End If
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim codigo As String, canaco As Double, i As Long, X As Long, cTip As String
Dim RS1 As New ADODB.Recordset, RS2 As New ADODB.Recordset
If est Or Row < 1 Then Exit Sub
Select Case Col
Case 1
    est = True
    If Col <> NewCol And (modo = "A") And Toolbar1.Buttons(12).Visible = True Then
        vaSpread1.Col = Col: vaSpread1.Row = Row
        If Trim(vaSpread1.text) = "" Then
            Cancel = True
        Else
            vaSpread1.Col = 1: codigo = Trim(vaSpread1.text)
            
            For i = 1 To vaSpread1.MaxRows
                vaSpread1.Row = i: vaSpread1.Col = 1
                If Trim(codigo) = Trim(vaSpread1.text) And Row <> i And Trim(vaSpread1.text) <> "" Then
                    Cancel = True
                    vaSpread1.Col = 1: vaSpread1.Row = Row: vaSpread1.text = ""
                    MsgBox "Artículo de cafetería ya fué ingresado", vbCritical + vbOKOnly, MsgTitulo
                    est = False
                    Exit Sub
                End If
            Next i
            RS1.Open "SELECT * FROM b_totpreciocaf WHERE tpc_cencos = '" & MuestraCasino(1) & "' AND tpc_codigo = '" & codigo & "'", vg_db, adOpenStatic
            If Not RS1.EOF Then
                vaSpread1.Row = Row
                vaSpread1.Col = 1: vaSpread1.text = Trim(RS1!tpc_codigo)
                vaSpread1.Col = 2: vaSpread1.text = Trim(RS1!tpc_nombre)
                vaSpread1.Col = 4: vaSpread1.text = Format(RS1!tpc_precio, fg_Pict(9, vg_DPr))
                vaSpread1.Col = 5: vaSpread1.TypeComboBoxList = "CONTADO" & Chr$(9) & "CREDITO" & Chr$(9) & "CUENTA ABONO CLIENTE"
                vaSpread1.Col = 6: vaSpread1.text = ""
                vaSpread1.Col = 7: vaSpread1.text = ""
                vaSpread1.SetActiveCell 3, vaSpread1.ActiveRow
                SumaCantidades True
            Else
                Cancel = True
                MsgBox "Artículo de cafetería no existe", vbCritical + vbOKOnly, MsgTitulo
                vaSpread1.text = "": vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow
                RS1.Close: Set RS1 = Nothing
                est = False
                Exit Sub
            End If
            RS1.Close: Set RS1 = Nothing
        End If
    End If
    est = False
End Select
If Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    Cancel = True
ElseIf Row <> NewRow And NewRow > 0 And modo = "" And Toolbar1.Buttons(12).Visible = False Then
    est = True
    vaSpread1.Row = NewRow: vaSpread1.Col = 5
    Frame2.Enabled = (vaSpread1.TypeComboBoxCurSel = 1 Or vaSpread1.TypeComboBoxCurSel = 2 And Label1.Caption = "")
    vaSpread1.Col = 5: cTip = vaSpread1.text
    vaSpread1.Col = 6
    RS1.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & Trim(vaSpread1.text) & "' AND cli_tipo = " & IIf(cTip = "CREDITO", 1, 3) & "", vg_db, adOpenStatic
    If Not RS1.EOF Then
        fpText1(1).text = fg_PintaRut(RS1!cli_codigo)
        fpayuda(1).Caption = RS1!cli_nombre
    Else
        fpText1(1).text = ""
        fpayuda(1).Caption = ""
    End If
    RS1.Close: Set RS1 = Nothing
    vaSpread1.Col = 7
    fpText1(2).text = Trim(vaSpread1.text)
    est = False
End If
End Sub

Private Sub vaSpread2_EditChange(ByVal Col As Long, ByVal Row As Long)
Dim cPro As String
If est Then Exit Sub
If modo = "" Then modo = "M": Gl_Ac_Botones Me, 11, 0, modo
Select Case Col
Case 4
    est = True
    vaSpread2.Col = 1: vaSpread2.Row = Row: cPro = Trim(vaSpread2.text)
    RevisaSobreStock cPro, Row, True
    est = False
End Select
End Sub

Private Sub vaSpread2_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If Row = 0 Then Exit Sub
Dim Stock As String, Nombre As String
TipWidth = 4000
ShowTip = True
MultiLine = 2
vaSpread2.Row = Row: vaSpread2.Col = 8: Stock = Format(vaSpread2.text, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
vaSpread2.Row = Row: vaSpread2.Col = 2: Nombre = vaSpread2.text
TipText = "Bodega   : " & Trim(Left(Combo1(0).text, 50)) & vbCrLf & _
          "Producto : " & Trim(Nombre) & vbCrLf & _
          "Stock       : " & Format(Trim(Stock), fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
End Sub
