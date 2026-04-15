VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_ListPreWeb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Precios Web"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   13005
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Lista de Precios"
      TabPicture(0)   =   "M_ListPreWeb.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "vaSpread1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Lista de  Precios"
      TabPicture(1)   =   "M_ListPreWeb.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Contratos Asignados"
      TabPicture(2)   =   "M_ListPreWeb.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Ver Lista Precios"
      TabPicture(3)   =   "M_ListPreWeb.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame6"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame6 
         Height          =   6255
         Left            =   240
         TabIndex        =   33
         Top             =   1800
         Width           =   12255
         Begin VB.Frame Frame8 
            Height          =   435
            Left            =   240
            TabIndex        =   44
            Top             =   5640
            Width           =   1275
            Begin VB.TextBox TextCai2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   1
               Left            =   45
               TabIndex        =   45
               Top             =   135
               Width           =   1170
            End
         End
         Begin VB.Frame Frame7 
            Height          =   435
            Left            =   1650
            TabIndex        =   42
            Top             =   5640
            Width           =   6885
            Begin VB.TextBox TextCai2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   2
               Left            =   45
               TabIndex        =   43
               Top             =   135
               Width           =   6780
            End
         End
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   5250
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Width           =   11790
            _Version        =   393216
            _ExtentX        =   20796
            _ExtentY        =   9260
            _StockProps     =   64
            AllowCellOverflow=   -1  'True
            AutoCalc        =   0   'False
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
            FormulaSync     =   0   'False
            MaxCols         =   3
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "M_ListPreWeb.frx":0070
            ScrollBarTrack  =   3
            ClipboardOptions=   0
         End
      End
      Begin VB.Frame Frame5 
         Height          =   975
         Left            =   2280
         TabIndex        =   32
         Top             =   720
         Width           =   8055
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   36
            Top             =   600
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   556
            Enabled         =   -1  'True
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
            BackColor       =   -2147483624
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   35
            Top             =   240
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
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   7440
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "M_ListPreWeb.frx":19B6
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   360
            Left            =   3360
            TabIndex        =   41
            Top             =   600
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Cargar Información"
                  ImageIndex      =   1
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   3000
            Picture         =   "M_ListPreWeb.frx":1D50
            Top             =   120
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3465
            TabIndex        =   38
            Top             =   195
            Width           =   4335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Lista Precio"
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
            Left            =   360
            TabIndex        =   37
            Top             =   310
            Width           =   1020
         End
         Begin VB.Label Label1 
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
            Index           =   4
            Left            =   360
            TabIndex        =   34
            Top             =   660
            Width           =   540
         End
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3495
            TabIndex        =   39
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   971
         Left            =   -72480
         TabIndex        =   7
         Top             =   840
         Width           =   7335
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "M_ListPreWeb.frx":205A
            Left            =   2010
            List            =   "M_ListPreWeb.frx":2064
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   240
            Width           =   2500
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   1
            Left            =   2010
            TabIndex        =   9
            Top             =   555
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
            NoSpecialKeys   =   0
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
            Left            =   4590
            TabIndex        =   12
            Top             =   645
            Width           =   585
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
            Index           =   1
            Left            =   525
            TabIndex        =   11
            Top             =   645
            Width           =   1140
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
            Index           =   0
            Left            =   525
            TabIndex        =   10
            Top             =   345
            Width           =   1380
         End
      End
      Begin VB.Frame Frame2 
         Height          =   7215
         Left            =   -74640
         TabIndex        =   2
         Top             =   600
         Width           =   12135
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   4200
            TabIndex        =   3
            Top             =   1440
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
            MaxLength       =   50
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
            Left            =   4200
            TabIndex        =   4
            Top             =   1080
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
            Caption         =   "Descripción"
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
            Left            =   2160
            TabIndex        =   6
            Top             =   1560
            Width           =   1020
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
            Index           =   2
            Left            =   2160
            TabIndex        =   5
            Top             =   1080
            Width           =   600
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
         Height          =   7575
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   12255
         Begin VB.Frame Frame4 
            Height          =   1935
            Left            =   5760
            TabIndex        =   27
            Top             =   2280
            Width           =   855
            Begin VB.CommandButton Command1 
               Caption         =   "<<"
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
               Index           =   3
               Left            =   200
               TabIndex        =   31
               Top             =   1440
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   "<"
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
               Index           =   2
               Left            =   200
               TabIndex        =   30
               Top             =   1080
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   ">>"
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
               Index           =   1
               Left            =   200
               TabIndex        =   29
               Top             =   600
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   ">"
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
               Index           =   0
               Left            =   200
               TabIndex        =   28
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Asignado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7095
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   5415
            Begin VB.Frame Frame16 
               Height          =   435
               Left            =   1050
               TabIndex        =   24
               Top             =   6600
               Width           =   4005
               Begin VB.TextBox TextCai1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   3
                  Left            =   45
                  TabIndex        =   25
                  Top             =   135
                  Width           =   3900
               End
            End
            Begin VB.Frame Frame13 
               Height          =   435
               Left            =   120
               TabIndex        =   22
               Top             =   6600
               Width           =   915
               Begin VB.TextBox TextCai1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   2
                  Left            =   45
                  TabIndex        =   23
                  Top             =   135
                  Width           =   810
               End
            End
            Begin FPSpread.vaSpread vaSpread4 
               Height          =   6255
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   5175
               _Version        =   393216
               _ExtentX        =   9128
               _ExtentY        =   11033
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
               OperationMode   =   4
               SelectBlockOptions=   0
               SpreadDesigner  =   "M_ListPreWeb.frx":2078
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Disponible"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7095
            Left            =   6720
            TabIndex        =   15
            Top             =   240
            Width           =   5415
            Begin VB.Frame Frame14 
               Height          =   435
               Left            =   1050
               TabIndex        =   18
               Top             =   6600
               Width           =   4125
               Begin VB.TextBox TextCan1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   3
                  Left            =   45
                  TabIndex        =   19
                  Top             =   135
                  Width           =   4020
               End
            End
            Begin VB.Frame Frame15 
               Height          =   435
               Left            =   120
               TabIndex        =   16
               Top             =   6600
               Width           =   915
               Begin VB.TextBox TextCan1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   2
                  Left            =   45
                  TabIndex        =   17
                  Top             =   135
                  Width           =   810
               End
            End
            Begin FPSpread.vaSpread vaSpread5 
               Height          =   6255
               Left            =   120
               TabIndex        =   20
               Top             =   240
               Width           =   5175
               _Version        =   393216
               _ExtentX        =   9128
               _ExtentY        =   11033
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
               OperationMode   =   4
               SelectBlockOptions=   0
               SpreadDesigner  =   "M_ListPreWeb.frx":39D8
            End
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6285
         Left            =   -72480
         TabIndex        =   13
         Top             =   1920
         Width           =   7365
         _Version        =   393216
         _ExtentX        =   12991
         _ExtentY        =   11086
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   2
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_ListPreWeb.frx":534D
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_ListPreWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String, codigo As String, Msgtitulo As String
Dim est As Boolean

Private Sub Command1_Click(Index As Integer)
Dim i As Long, codcen As String, nomcen As String, codcco As String, estmar As Boolean
Select Case Index
Case 0
    vg_codigo = "X"
    modo = "M"
    estmar = False
    For i = 1 To vaSpread4.MaxRows
        vaSpread4.Row = i
        If vaSpread4.SelModeSelected = True Then estmar = True
    Next i
    If Not estmar Then MsgBox "Debe seleccionar a lo menos un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    For i = 1 To vaSpread4.MaxRows
        vaSpread4.Row = i
        If vaSpread4.SelModeSelected = True Then
           vaSpread4.Col = 2
           codcen = vaSpread4.text
           vaSpread4.Col = 3
           nomcen = vaSpread4.text
           vaSpread4.Col = 4
           codcco = vaSpread4.text
           vaSpread4.DeleteRows vaSpread4.Row, 1
           vaSpread4.MaxRows = vaSpread4.MaxRows - 1
           
           vaSpread5.MaxRows = vaSpread5.MaxRows + 1
           vaSpread5.Row = vaSpread5.MaxRows
           vaSpread5.Col = 2
           vaSpread5.text = codcen
           vaSpread5.Col = 3
           vaSpread5.text = nomcen
           vaSpread5.Col = 4
           vaSpread5.text = codcco
        End If
    Next i
Case 1
    vg_codigo = "X"
    modo = "M"
    For i = 1 To vaSpread4.MaxRows
        vaSpread4.Row = i
        vaSpread4.Col = 2
        codcen = vaSpread4.text
        vaSpread4.Col = 3
        nomcen = vaSpread4.text
        vaSpread4.Col = 4
        codcco = vaSpread4.text
        vaSpread5.MaxRows = vaSpread5.MaxRows + 1
        vaSpread5.Row = vaSpread5.MaxRows
        vaSpread5.Col = 2
        vaSpread5.text = codcen
        vaSpread5.Col = 3
        vaSpread5.text = nomcen
        vaSpread5.Col = 4
        vaSpread5.text = codcco
    Next i
    vaSpread4.MaxRows = 0
'        vaSpread1.ClearSelection = True
'        If vaSpread4.OperationMode = OperationModeMulti Then
Case 2
    vg_codigo = "X"
    modo = "M"
    estmar = False
    For i = 1 To vaSpread5.MaxRows
        vaSpread5.Row = i
        If vaSpread5.SelModeSelected = True Then estmar = True
    Next i
    If Not estmar Then MsgBox "Debe seleccionar a lo menos un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    For i = 1 To vaSpread5.MaxRows
        vaSpread5.Row = i
        If vaSpread5.SelModeSelected = True Then
           vaSpread5.Col = 2
           codcen = vaSpread5.text
           vaSpread5.Col = 3
           nomcen = vaSpread5.text
           vaSpread5.Col = 4
           codcco = vaSpread5.text
           vaSpread5.DeleteRows vaSpread5.Row, 1
           vaSpread5.MaxRows = vaSpread5.MaxRows - 1
           
           vaSpread4.MaxRows = vaSpread4.MaxRows + 1
           vaSpread4.Row = vaSpread4.MaxRows
           vaSpread4.Col = 2
           vaSpread4.text = codcen
           vaSpread4.Col = 3
           vaSpread4.text = nomcen
           vaSpread4.Col = 4
           vaSpread4.text = codcco
        End If
    Next i
Case 3
    vg_codigo = "X"
    modo = "M"
    For i = 1 To vaSpread5.MaxRows
        vaSpread5.Row = i
        vaSpread5.Col = 2
        codcen = vaSpread5.text
        vaSpread5.Col = 3
        nomcen = vaSpread5.text
        vaSpread5.Col = 4
        codcco = vaSpread5.text
        vaSpread4.MaxRows = vaSpread4.MaxRows + 1
        vaSpread4.Row = vaSpread4.MaxRows
        vaSpread4.Col = 2
        vaSpread4.text = codcen
        vaSpread4.Col = 3
        vaSpread4.text = nomcen
        vaSpread4.Col = 4
        vaSpread4.text = codcco
    Next i
    vaSpread5.MaxRows = 0
End Select
'-------> Bloquear hoja
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = True
If vg_codigo <> "" Then Gl_Ac_Botones Me, 14, 0, modo
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 9390
Me.Width = 13095
Msgtitulo = "Lista de Precios"
fg_centra Me
SSTab1.Tab = 0
modo = ""
est = True
Gl_Mo_Botones Me, 14
Toolbar1.Buttons.item(15).ButtonMenus(1).Visible = False
Gl_Ac_Botones Me, 14, 1, modo
Combo1.ListIndex = 1
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
MoverDatosGrilla
MoverDatosListadePrecios
MoverDatosListadePreciosCasinoAsignados
Gl_Ac_Botones Me, 14, IIf(vaSpread1.MaxRows > 0, 1, 2), modo
End Sub

Sub MoverDatosGrilla()
fg_carga ""
Dim X As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 250
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
vaSpread1.Row = -1
vaSpread1.Col = -1
vaSpread1.Lock = True
Set RS = vg_dbpedweb.Execute("pedweb_s_listaprecios 2, '', '', ''")
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   vaSpread1.text = RS!codigo
   vaSpread1.Col = 2
   vaSpread1.text = Trim(RS!descripcion)
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
If vaSpread1.MaxRows > 0 Then
   vaSpread1.Row = 1
   vaSpread1.Col = 1
   codigo = ""
   codigo = Val(vaSpread1.text)
   vaSpread1.SetActiveCell 1, 1 ': vaSpread1.SetFocus
End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
fg_descarga
End Sub

Sub MoverDatosListadePreciosCasinoAsignados()
fg_carga ""
MoverDatosListadePrecios
est = True
Limpia 2

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = vaSpread1.text
Dim X As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread4.TextTip = 2
vaSpread4.TextTipDelay = 250
X = vaSpread4.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread4.Visible = False
Set RS = vg_dbpedweb.Execute("pedweb_s_listadeprecioscasinoasignados 1, '" & codigo & "'")
Do While Not RS.EOF
   vaSpread4.MaxRows = vaSpread4.MaxRows + 1
   vaSpread4.Row = vaSpread4.MaxRows
   vaSpread4.Col = 2
   vaSpread4.text = IIf(IsNull(RS!centrocosto), "", RS!centrocosto)
   vaSpread4.Col = 3
   vaSpread4.text = Trim(IIf(IsNull(RS!Nombre), "", RS!Nombre))
   vaSpread4.Col = 4
   vaSpread4.text = Trim(IIf(IsNull(RS!CodigoContrato), "", RS!CodigoContrato))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread4.Visible = True

' Control displays text tips aligned to pointer with focus
vaSpread5.TextTip = 2
vaSpread5.TextTipDelay = 250
X = vaSpread5.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread5.Visible = False
Set RS = vg_dbpedweb.Execute("pedweb_s_listadeprecioscasinodisponibles 1, '" & codigo & "'")
Do While Not RS.EOF
   vaSpread5.MaxRows = vaSpread5.MaxRows + 1
   vaSpread5.Row = vaSpread5.MaxRows
   vaSpread5.Col = 2
   vaSpread5.text = IIf(IsNull(RS!centrocosto), "", RS!centrocosto)
   vaSpread5.Col = 3
   vaSpread5.text = Trim(IIf(IsNull(RS!Nombre), "", RS!Nombre))
   vaSpread5.Col = 4
   vaSpread5.text = Trim(IIf(IsNull(RS!codigo), "", RS!codigo))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread5.Visible = True
est = False
fg_descarga
End Sub

Sub MoverDatosListadePrecios()
fg_carga ""
est = True
Limpia 1
Set RS = vg_dbpedweb.Execute("pedweb_s_listaprecios 3, '" & codigo & "', '', ''")
If Not RS.EOF Then
   fpLongInteger1(0).Value = RS!codigo
   fpText1(0).text = Trim(RS!descripcion)
   Frame3.Caption = RS!codigo & " - " & Trim(RS!descripcion)
End If
RS.Close: Set RS = Nothing
fg_descarga
est = False
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Select Case Index
Case 1
    If Val(fpLongInteger1(1).Value) < 1 Then fpayuda(0).Caption = "": Exit Sub
    Set RS = vg_dbpedweb.Execute("pedweb_s_listaprecios 3, '" & fpLongInteger1(1).Value & "', '', ''")
    If RS.EOF Then RS.Close: Set RS = Nothing:: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!descripcion)
    RS.Close: Set RS = Nothing
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 1
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
End Select
End Sub

Private Sub fpText1_Change(Index As Integer)
Select Case Index
Case 0
    If est Then Exit Sub
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 14, 0, modo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(2) = False
Case 1
    If LimpiaDato(Trim(fpText1(1).text)) & Chr(KeyAscii) = "" Then Exit Sub
    vaSpread1.Visible = False
    vaSpread1.Row = -1
    vaSpread1.Col = -1
    vaSpread1.Lock = True
    If Combo1.ItemData(Combo1.ListIndex) = 0 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_listaprecios 4, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%', ''")
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_listaprecios 5, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%', ''")
    End If
    If RS2.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS2!nReg
    i = 1
    If Not RS2.EOF Then
       Do While Not RS2.EOF
          vaSpread1.Row = i: i = i + 1
          vaSpread1.Col = 1
          vaSpread1.TypeHAlign = 1
          vaSpread1.text = RS2!codigo
          vaSpread1.Col = 2
          vaSpread1.text = Trim(RS2!descripcion)
          RS2.MoveNext
        Loop
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
        Gl_Ac_Botones Me, 14, 1, modo
    Else
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
    End If
    RS2.Close: Set RS2 = Nothing
    vaSpread1.Col = 1: vaSpread1.col2 = vaSpread1.maxcols: vaSpread1.Row = 1: vaSpread1.row2 = vaSpread1.MaxRows
    vaSpread1.SetActiveCell 1, 1
    vaSpread1.Visible = True
    If fpText1(1).text = "" Then Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro" Else Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
End Select
End Sub

Private Sub image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "s_Lista_Precios", "sub_", "Lista de Precios", "lispreweb"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
    fpDateTime1(0).SetFocus
End Select
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If est Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = vaSpread1.text
Select Case SSTab1.Tab
Case 0
Case 1
    MoverDatosListadePrecios
Case 2
    MoverDatosListadePreciosCasinoAsignados
Case 3
    fpayuda(0).Caption = ""
    If Val(fpLongInteger1(1).Value) = 0 Then
       fpLongInteger1(1).Value = Val(codigo)
    End If
    Set RS = vg_dbpedweb.Execute("pedweb_s_listaprecios 3, '" & fpLongInteger1(1).Value & "', '', ''")
    If Not RS.EOF Then
       fpLongInteger1(1).Value = RS!codigo
       fpayuda(0).Caption = Trim(RS!descripcion)
    End If
    RS.Close: Set RS = Nothing
    vaSpread2.MaxRows = 0
End Select
End Sub

Sub Limpia(Op As Integer)
Select Case Op
Case 1
    fpLongInteger1(0).Value = ""
    fpLongInteger1(0).Enabled = False
    fpText1(0).text = ""
    Frame3.Caption = ""
Case 2
    vaSpread4.MaxRows = 0
    vaSpread5.MaxRows = 0
End Select
End Sub

Private Sub TextCai1_Change(Index As Integer)
Dim i As Long, nom As String
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
End Sub

Private Sub TextCai2_Change(Index As Integer)
Dim i As Long, nom As String
Select Case Index
Case 1, 2
    vaSpread2.Visible = False
    If Trim(TextCai2(Index).text) <> "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           vaSpread2.Col = Index: nom = UCase(Trim(vaSpread2.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCai2(Index).text) & "*"
           vaSpread2.Col = Index
           If indactivo = -1 And Trim(vaSpread2.text) <> "" Then
              If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
           Else
              If vaSpread2.RowHidden = False Then vaSpread2.RowHidden = True
           End If
        Next i
        vaSpread2.SetActiveCell Index, 1
    End If
    vaSpread2.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread2.ColUserSortIndicator(IIf(Trim(TextCai2(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread2.SortKey(1) = IIf(Trim(TextCai2(Index).text) = "", 0, 0): vaSpread2.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread2.Sort -1, -1, vaSpread2.maxcols, vaSpread2.MaxRows, SortByRow
    If Trim(TextCai2(Index).text) = "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
       Next
       vaSpread2.SetActiveCell Index, vaSpread2.SearchCol(Index, 0, vaSpread2.MaxRows, Trim(TextCai2(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread2.SetActiveCell Index, 1
    End If
    vaSpread2.Visible = True
End Select

End Sub

Private Sub TextCan1_Change(Index As Integer)
Dim i As Long, nom As String
Select Case Index
Case 2, 3, 4
    vaSpread5.Visible = False
    If Trim(TextCan1(Index).text) <> "" Then
       For i = 1 To vaSpread5.MaxRows
           vaSpread5.Row = i
           vaSpread5.Col = Index: nom = UCase(Trim(vaSpread5.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCan1(Index).text) & "*"
           vaSpread5.Col = 2
           If indactivo = -1 And Trim(vaSpread5.text) <> "" Then
              If vaSpread5.RowHidden = True Then vaSpread5.RowHidden = False
           Else
              If vaSpread5.RowHidden = False Then vaSpread5.RowHidden = True
           End If
        Next i
        vaSpread5.SetActiveCell Index, 1
    End If
    vaSpread5.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread5.ColUserSortIndicator(IIf(Trim(TextCan1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread5.SortKey(1) = IIf(Trim(TextCan1(Index).text) = "", 0, 0): vaSpread5.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread5.Sort -1, -1, vaSpread5.maxcols, vaSpread5.MaxRows, SortByRow
    If Trim(TextCan1(Index).text) = "" Then
       For i = 1 To vaSpread5.MaxRows
           vaSpread5.Row = i
           If vaSpread5.RowHidden = True Then vaSpread5.RowHidden = False
       Next
       vaSpread5.SetActiveCell Index, vaSpread5.SearchCol(Index, 0, vaSpread5.MaxRows, Trim(TextCan1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread5.SetActiveCell Index, 1
    End If
    vaSpread5.Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codcco As String
Dim i As Long
Select Case Button.Index
Case 1 '-------> Incluir nuevos registros
    modo = "A"
    Select Case SSTab1.Tab
    Case 0, 1 '-------> lista de precios
        est = True
        SSTab1.Tab = 1
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
        '-------> Traer ultimo registro
        Limpia 1
        Set RS = vg_dbpedweb.Execute("pedweb_s_listaprecios 1, '', '', ''")
        If Not RS.EOF Then RS.MoveFirst: codigo = RS!codigo + 1 Else codigo = 1
        RS.Close: Set RS = Nothing
        fpLongInteger1(0).text = codigo
        fpText1(0).SetFocus
        vg_codigo = "x"
        est = False
    End Select
    If vg_codigo <> "" Then Gl_Ac_Botones Me, 14, 0, modo
Case 3 '-------> Alterar registro
    Select Case SSTab1.Tab
    Case 0, 1 '-------> Lista de Precios
        modo = "M"
        Gl_Ac_Botones Me, 14, 0, modo
        SSTab1.Tab = 1
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
        fpText1(0).SetFocus
    Case 2 '-------> Asignar casinos
        modo = "M"
        Gl_Ac_Botones Me, 14, 0, modo
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
    End Select
Case 5 '-------> Eliminar Registro y sus relaciones
    Select Case SSTab1.Tab
    Case 0, 1 '-------> Lista de Precios
        If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        If MsgBox("Elimina registro y todas sus relaciones...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        vg_dbpedweb.Execute ("pedweb_d_listadeprecios '" & codigo & "'")
        SSTab1.Tab = 0
        MoverDatosGrilla
        MoverDatosListadePrecios
    End Select
    modo = "": Gl_Ac_Botones Me, 14, 1, modo
Case 7 '-------> Actualizar lista
    Select Case SSTab1.Tab
    Case 0
        MoverDatosGrilla
        fpText1(1).text = ""
        Gl_Ac_Botones Me, 14, IIf(vaSpread1.MaxRows > 0, 1, 2), modo
    Case 1
        MoverDatosListadePrecios
    Case 2
        MoverDatosListadePreciosCasinoAsignados
    End Select
Case 10 '-------> Cancelar Información
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    Select Case SSTab1.Tab
    Case 1
        SSTab1.Tab = 1
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        MoverDatosListadePrecios
    Case 2
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        MoverDatosListadePreciosCasinoAsignados
    End Select
    '-------> Desbloquear hojas
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    modo = "": Gl_Ac_Botones Me, 14, 1, modo
Case 12 '-------> Graba Registro
    Select Case SSTab1.Tab
    Case 1 '-------> Grabar Lista de Precios
        If LimpiaDato(Trim(fpText1(0).text)) = "" Then MsgBox "Debe ingresar información...", vbCritical, Msgtitulo: Exit Sub
        If modo = "A" Then
           codigo = 0
           Set RS = vg_dbpedweb.Execute("pedweb_iu_listadeprecios 'A', 0, '" & LimpiaDato(Trim(fpText1(0).text)) & "'")
           If Not RS.EOF Then
              codigo = RS!indice
           End If
           RS.Close: Set RS = Nothing
           fpLongInteger1(0).text = codigo
           vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
           vaSpread1.SetActiveCell 1, vaSpread1.Row
        Else
            vg_dbpedweb.Execute "pedweb_iu_reglasdenegocios 'M', " & codigo & ", '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & tipreg & "', '" & vg_NUsr & "', '', '" & vg_NUsr & "', ''"
        End If
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = LimpiaDato(Trim(fpLongInteger1(0).text))
        vaSpread1.Col = 2: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = LimpiaDato(Trim(fpText1(0).text))
    Case 2 '-------> Grabar asignación casino
        fg_carga ""
        vg_dbpedweb.Execute ("DELETE s_Contrato WHERE CodigoLista = " & codigo & "")
        For i = 1 To vaSpread4.MaxRows
            vaSpread4.Row = i
            vaSpread4.Col = 4
            codcco = vaSpread4.text
            vg_dbpedweb.Execute ("INSERT INTO s_Contrato VALUES (" & codigo & ", '" & codcco & "')")
        Next i
        fg_descarga
    End Select
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    modo = "": Gl_Ac_Botones Me, 14, 1, modo
Case 19 '------> impresion
    Select Case SSTab1.Tab
    Case 0, 1 '-------> Lista de precios
        I_ListadePrecios
    Case 2 '-------> asignacion casino
        I_WebRep.LlenaDatos "Impresión Lista Precios Casinos Asignado", "lisprecas"
        I_WebRep.Show 1
        Me.Refresh
        
'        vaSpread1.Row = vaSpread1.ActiveRow
'        vaSpread1.Col = 1
'        codigo = vaSpread1.Text
'        I_ListadePreciosCasinoAsignados CStr(codigo)
    Case 3
        vg_opimp = 9999999
        I_WebRep.LlenaDatos "Impresión Lista Precios", "lisprecio"
        I_WebRep.Show 1
        Me.Refresh
        vg_opimp = 0
    End Select
Case 22
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu
Case "Importar Datos"
    If vaSpread1.MaxRows < 1 Then Exit Sub
    P_ImpRut.LlenaDatos "Importar Lista de Precios", "lispre"
    P_ImpRut.Show 1
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    MoverListaPrecio
End Select
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
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
End Select
End Sub

Private Sub vaSpread4_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
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
End Sub

Private Sub vaSpread5_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If vaSpread5.MaxRows < 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread5.Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
Select Case Col
Case 1
    vaSpread5.Col = Col
    TipText = "Centro de Costo : " & vaSpread5.text
Case 2
    vaSpread5.Col = Col
    TipText = "Descripción : " & Trim(vaSpread5.text)
End Select
End Sub

Sub MoverListaPrecio()
If Val(fpLongInteger1(1).Value) < 1 Or Trim(fpDateTime1(Index).text) = "" Then fpayuda(0).Caption = "": Exit Sub
Dim X As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread2.TextTip = 2
vaSpread2.TextTipDelay = 250
X = vaSpread2.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread2.Visible = False
vaSpread2.MaxRows = 0
Set RS = vg_dbpedweb.Execute("pedweb_s_listaprecios 6, " & Val(fpLongInteger1(1).Value) & ", '', '" & Format(fpDateTime1(Index).text, "yyyymmdd") & "'")
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      vaSpread2.Col = 1
      vaSpread2.text = IIf(IsNull(RS!codigo), "", RS!codigo)
      vaSpread2.Col = 2
      vaSpread2.text = Trim(IIf(IsNull(RS!descripcion), "", RS!descripcion))
      vaSpread2.Col = 3
      vaSpread2.TypeHAlign = TypeHAlignRight
      vaSpread2.text = Format(IIf(IsNull(RS!precio), 0, RS!precio), fg_Pict(6, 2))
      RS.MoveNext
   Loop
Else
   vaSpread2.Visible = True
   MsgBox "No existe información a consultar...", vbCritical, Msgtitulo
End If
RS.Close: Set RS = Nothing
vaSpread2.Visible = True
End Sub
