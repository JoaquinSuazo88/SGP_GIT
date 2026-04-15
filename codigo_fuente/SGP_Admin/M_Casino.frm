VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4DBFB8CD-9EF9-11D0-8BC4-00AA00B42B7C}#3.0#0"; "Cal32x30.ocx"
Begin VB.Form M_Casino 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Maestro de Casino"
   ClientHeight    =   10230
   ClientLeft      =   2550
   ClientTop       =   765
   ClientWidth     =   15435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   15435
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9705
      Left            =   60
      TabIndex        =   42
      Top             =   450
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   17119
      _Version        =   393216
      Style           =   1
      Tabs            =   13
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   520
      TabMaxWidth     =   4
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Casinos"
      TabPicture(0)   =   "M_Casino.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame18"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_Casino.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Regimen\Servicio"
      TabPicture(2)   =   "M_Casino.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Interfaces"
      TabPicture(3)   =   "M_Casino.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Calendario Días Feriados"
      TabPicture(4)   =   "M_Casino.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Parametro Despachos"
      TabPicture(5)   =   "M_Casino.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame8"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Actividades Diarias"
      TabPicture(6)   =   "M_Casino.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame9"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Parámetros de Stock"
      TabPicture(7)   =   "M_Casino.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame10"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Paramertización Código Barra"
      TabPicture(8)   =   "M_Casino.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame13"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Parametros (Q) Prod. y Vendidas"
      TabPicture(9)   =   "M_Casino.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame19"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Parametros Inv. Calendarizado "
      TabPicture(10)  =   "M_Casino.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame20"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "Parametros Precio Venta Cliente Calendarizado"
      TabPicture(11)  =   "M_Casino.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Frame21"
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "Parametro Categoria Dietetica"
      TabPicture(12)  =   "M_Casino.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Frame26"
      Tab(12).ControlCount=   1
      Begin VB.Frame Frame26 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Left            =   -74940
         TabIndex        =   152
         Top             =   570
         Width           =   15015
         Begin VB.CommandButton Command3 
            Caption         =   "Copiar Parametro C. Dietetica"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   13080
            TabIndex        =   154
            Top             =   8160
            Width           =   1695
         End
         Begin MSComctlLib.TreeView TvwDietetica 
            Height          =   7335
            Index           =   0
            Left            =   240
            TabIndex        =   153
            Top             =   480
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
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
         End
      End
      Begin VB.Frame Frame21 
         Height          =   8775
         Left            =   -74940
         TabIndex        =   141
         Top             =   570
         Width           =   13095
         Begin CalObjXLib.fpCalendar fpCalendar3 
            Height          =   6270
            Left            =   720
            TabIndex        =   142
            Top             =   840
            Width           =   11655
            _Version        =   196608
            _ExtentX        =   20558
            _ExtentY        =   11060
            _StockProps     =   68
            FirstDayOfWeek  =   1
            CurrentDate     =   "20000120"
            DateMin         =   "00000000"
            DateMax         =   "00000000"
            GrayAreaStyle   =   1
            GrayAreaBackColor=   -2147483633
            GrayAreaForeColor=   -2147483632
            HeaderStyle     =   2
            MonthHeaderStyle=   1
            YearHeaderStyle =   1
            BorderGrayAreaColor=   -2147483637
            ElementPictureBackground=   0   'False
            DisplayFormat   =   3
            BorderInnerStyle=   0
            BorderInnerHighlightColor=   -2147483633
            BorderInnerShadowColor=   -2147483642
            BorderInnerWidth=   1
            BorderOuterStyle=   0
            BorderOuterHighlightColor=   -2147483628
            BorderOuterShadowColor=   -2147483632
            BorderOuterWidth=   1
            BorderFrameWidth=   0
            BorderOutlineColor=   -2147483642
            BorderFrameColor=   -2147483633
            BorderOutlineWidth=   1
            BorderOutlineStyle=   1
            SpeedScrollYearIncrement=   1
            SpeedScrollMonthIncrement=   1
            GrayAreaAllowScroll=   0   'False
            WeekNumbers     =   0
            WeekDayHeader   =   3
            ElementTextStyle=   "M_Casino.frx":016C
            DrawFocusRect   =   0
            MultiSelect     =   2
            YearStartMonth  =   1
            YearStartDay    =   1
            HeaderSeparatorWidth=   0
            HeaderSeparatorColor=   0
            YearFormatStyle =   2
            RangeBeginDate  =   "00000000"
            RangeEndDate    =   "00000000"
            GridLineColor   =   0
            GridLineStyle   =   3
            AutoSet         =   -1  'True
            InheritOverride =   1
            CompactFormat   =   ""
            MouseIcon       =   "M_Casino.frx":0415
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8775
         Left            =   -74780
         TabIndex        =   43
         Top             =   570
         Width           =   12855
         Begin CalObjXLib.fpCalendar fpCalendar1 
            Height          =   7215
            Left            =   225
            TabIndex        =   44
            Top             =   480
            Width           =   12255
            _Version        =   196608
            _ExtentX        =   21616
            _ExtentY        =   12726
            _StockProps     =   68
            FirstDayOfWeek  =   1
            CurrentDate     =   "20000120"
            DateMin         =   "00000000"
            DateMax         =   "00000000"
            GrayAreaStyle   =   1
            GrayAreaBackColor=   -2147483633
            GrayAreaForeColor=   -2147483632
            HeaderStyle     =   2
            MonthHeaderStyle=   1
            YearHeaderStyle =   1
            BorderGrayAreaColor=   -2147483637
            ElementPictureBackground=   0   'False
            DisplayFormat   =   3
            BorderInnerStyle=   0
            BorderInnerHighlightColor=   -2147483633
            BorderInnerShadowColor=   -2147483642
            BorderInnerWidth=   1
            BorderOuterStyle=   0
            BorderOuterHighlightColor=   -2147483628
            BorderOuterShadowColor=   -2147483632
            BorderOuterWidth=   1
            BorderFrameWidth=   0
            BorderOutlineColor=   -2147483642
            BorderFrameColor=   -2147483633
            BorderOutlineWidth=   1
            BorderOutlineStyle=   1
            SpeedScrollYearIncrement=   1
            SpeedScrollMonthIncrement=   1
            GrayAreaAllowScroll=   0   'False
            WeekNumbers     =   0
            WeekDayHeader   =   3
            ElementTextStyle=   "M_Casino.frx":0431
            DrawFocusRect   =   0
            MultiSelect     =   2
            YearStartMonth  =   1
            YearStartDay    =   1
            HeaderSeparatorWidth=   0
            HeaderSeparatorColor=   0
            YearFormatStyle =   2
            RangeBeginDate  =   "00000000"
            RangeEndDate    =   "00000000"
            GridLineColor   =   0
            GridLineStyle   =   3
            AutoSet         =   -1  'True
            InheritOverride =   1
            CompactFormat   =   ""
            MouseIcon       =   "M_Casino.frx":06DA
         End
      End
      Begin VB.Frame Frame20 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8775
         Left            =   -74940
         TabIndex        =   134
         Top             =   570
         Width           =   13095
         Begin VB.CommandButton Command2 
            Caption         =   "Copiar Inv. Calendarizado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   10920
            TabIndex        =   136
            Top             =   7080
            Width           =   1455
         End
         Begin CalObjXLib.fpCalendar fpCalendar2 
            Height          =   6270
            Left            =   720
            TabIndex        =   135
            Top             =   480
            Width           =   11655
            _Version        =   196608
            _ExtentX        =   20558
            _ExtentY        =   11060
            _StockProps     =   68
            FirstDayOfWeek  =   1
            CurrentDate     =   "20000120"
            DateMin         =   "00000000"
            DateMax         =   "00000000"
            GrayAreaStyle   =   1
            GrayAreaBackColor=   -2147483633
            GrayAreaForeColor=   -2147483632
            HeaderStyle     =   2
            MonthHeaderStyle=   1
            YearHeaderStyle =   1
            BorderGrayAreaColor=   -2147483637
            ElementPictureBackground=   0   'False
            DisplayFormat   =   3
            BorderInnerStyle=   0
            BorderInnerHighlightColor=   -2147483633
            BorderInnerShadowColor=   -2147483642
            BorderInnerWidth=   1
            BorderOuterStyle=   0
            BorderOuterHighlightColor=   -2147483628
            BorderOuterShadowColor=   -2147483632
            BorderOuterWidth=   1
            BorderFrameWidth=   0
            BorderOutlineColor=   -2147483642
            BorderFrameColor=   -2147483633
            BorderOutlineWidth=   1
            BorderOutlineStyle=   1
            SpeedScrollYearIncrement=   1
            SpeedScrollMonthIncrement=   1
            GrayAreaAllowScroll=   0   'False
            WeekNumbers     =   0
            WeekDayHeader   =   3
            ElementTextStyle=   "M_Casino.frx":06F6
            DrawFocusRect   =   0
            MultiSelect     =   2
            YearStartMonth  =   1
            YearStartDay    =   1
            HeaderSeparatorWidth=   0
            HeaderSeparatorColor=   0
            YearFormatStyle =   2
            RangeBeginDate  =   "00000000"
            RangeEndDate    =   "00000000"
            GridLineColor   =   0
            GridLineStyle   =   3
            AutoSet         =   -1  'True
            InheritOverride =   1
            CompactFormat   =   ""
            MouseIcon       =   "M_Casino.frx":099F
         End
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   8
            Left            =   2760
            TabIndex        =   139
            Top             =   6960
            Width           =   795
            _Version        =   196608
            _ExtentX        =   1402
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Text            =   "0"
            MaxValue        =   "10"
            MinValue        =   "0"
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
            Index           =   9
            Left            =   8400
            TabIndex        =   140
            Top             =   6960
            Width           =   795
            _Version        =   196608
            _ExtentX        =   1402
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Text            =   "0"
            MaxValue        =   "10"
            MinValue        =   "0"
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
         Begin VB.Label Label7 
            Caption         =   "Días Holgura Despues"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6120
            TabIndex        =   138
            Top             =   7080
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "Días Holgura Antes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   137
            Top             =   7080
            Width           =   2055
         End
      End
      Begin VB.Frame Frame19 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8655
         Left            =   -74940
         TabIndex        =   131
         Top             =   690
         Width           =   13095
         Begin VB.CommandButton Command1 
            Caption         =   "Agregar Servicio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   11400
            TabIndex        =   133
            Top             =   6360
            Visible         =   0   'False
            Width           =   1335
         End
         Begin FPSpread.vaSpread vaSpread7 
            Height          =   7575
            Left            =   1800
            TabIndex        =   132
            Top             =   480
            Width           =   9510
            _Version        =   393216
            _ExtentX        =   16775
            _ExtentY        =   13361
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
            MaxCols         =   7
            SpreadDesigner  =   "M_Casino.frx":09BB
         End
      End
      Begin VB.Frame Frame18 
         Height          =   8775
         Left            =   -74780
         TabIndex        =   121
         Top             =   690
         Width           =   14775
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   6825
            Left            =   1680
            TabIndex        =   128
            Top             =   1320
            Width           =   11355
            Begin FPSpread.vaSpread vaSpread1 
               Height          =   6390
               Left            =   210
               TabIndex        =   129
               Top             =   210
               Width           =   10995
               _Version        =   393216
               _ExtentX        =   19394
               _ExtentY        =   11271
               _StockProps     =   64
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
               MaxCols         =   6
               MaxRows         =   10
               OperationMode   =   3
               RowsFrozen      =   1
               ScrollBars      =   2
               SelectBlockOptions=   0
               SpreadDesigner  =   "M_Casino.frx":243C
               VisibleCols     =   2
               VisibleRows     =   10
               ScrollBarTrack  =   1
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   3720
            TabIndex        =   122
            Top             =   120
            Width           =   7215
            Begin VB.ComboBox Combo1 
               Height          =   315
               Index           =   0
               ItemData        =   "M_Casino.frx":2A3C
               Left            =   1680
               List            =   "M_Casino.frx":2A46
               Style           =   2  'Dropdown List
               TabIndex        =   123
               Top             =   240
               Width           =   2865
            End
            Begin EditLib.fpText fptnombre 
               Height          =   315
               Left            =   1680
               TabIndex        =   124
               Top             =   600
               Width           =   2895
               _Version        =   196608
               _ExtentX        =   5106
               _ExtentY        =   556
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Caption         =   "Buscar Texto"
               BeginProperty Font 
                  Name            =   "Tahoma"
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
               TabIndex        =   127
               Top             =   675
               Width           =   1470
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "B"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   4680
               TabIndex        =   126
               Top             =   675
               Visible         =   0   'False
               Width           =   120
            End
            Begin VB.Label Label1 
               Caption         =   "Buscar Columna"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   11
               Left            =   150
               TabIndex        =   125
               Top             =   345
               Width           =   1485
            End
         End
         Begin FPSpread.vaSpread vaSpread3 
            Height          =   5100
            Left            =   1800
            TabIndex        =   130
            Top             =   2640
            Visible         =   0   'False
            Width           =   7305
            _Version        =   393216
            _ExtentX        =   12885
            _ExtentY        =   8996
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
            MaxCols         =   40
            MaxRows         =   20
            ScrollBars      =   2
            SpreadDesigner  =   "M_Casino.frx":2A5A
            VisibleCols     =   2
            VisibleRows     =   15
            ScrollBarTrack  =   1
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H80000018&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   0
            Top             =   9480
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Frame Frame13 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7695
         Left            =   -74280
         TabIndex        =   99
         Top             =   810
         Width           =   11535
         Begin FPSpread.vaSpread vaSpread6 
            Height          =   5055
            Left            =   240
            TabIndex        =   100
            Top             =   1080
            Width           =   11055
            _Version        =   393216
            _ExtentX        =   19500
            _ExtentY        =   8916
            _StockProps     =   64
            ButtonDrawMode  =   1
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
            MaxCols         =   5
            MaxRows         =   3
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "M_Casino.frx":3976
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Parámetros de Stock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -73680
         TabIndex        =   81
         Top             =   2670
         Width           =   8775
         Begin VB.OptionButton Option2 
            Caption         =   "Realiza Inventario Rotativo por Requerimiento Menú"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   90
            Top             =   840
            Width           =   5295
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Realiza Inventario Rotativo por Inventario de Stock"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   89
            Top             =   360
            Value           =   -1  'True
            Width           =   4815
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Genera Ajuste/Implantación automática en el Inventario"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   87
            Top             =   2640
            Width           =   5415
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Realizado Diariamente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   86
            Top             =   2280
            Width           =   2655
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            ItemData        =   "M_Casino.frx":3EB9
            Left            =   3720
            List            =   "M_Casino.frx":3EBB
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   1680
            Width           =   1695
         End
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   4
            Left            =   3720
            TabIndex        =   88
            Top             =   1200
            Width           =   555
            _Version        =   196608
            _ExtentX        =   979
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.Label Label5 
            Caption         =   "Tipo de Inventario Rotativo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   84
            Top             =   1800
            Width           =   3015
         End
         Begin VB.Label Label4 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   83
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Porcentaje para Inventario Rotativo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   82
            Top             =   1320
            Width           =   3135
         End
      End
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7215
         Left            =   -74760
         TabIndex        =   79
         Top             =   750
         Width           =   11415
         Begin FPSpread.vaSpread vaSpread5 
            Height          =   6615
            Left            =   360
            TabIndex        =   80
            Top             =   360
            Width           =   10695
            _Version        =   393216
            _ExtentX        =   18865
            _ExtentY        =   11668
            _StockProps     =   64
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
            MaxRows         =   10
            SpreadDesigner  =   "M_Casino.frx":3EBD
         End
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8415
         Left            =   -74400
         TabIndex        =   77
         Top             =   750
         Width           =   11655
         Begin FPSpread.vaSpread vaSpread4 
            Height          =   7815
            Left            =   240
            TabIndex        =   78
            Top             =   360
            Width           =   11295
            _Version        =   393216
            _ExtentX        =   19923
            _ExtentY        =   13785
            _StockProps     =   64
            ButtonDrawMode  =   1
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   14
            MaxRows         =   20
            ProcessTab      =   -1  'True
            ScrollBars      =   2
            SpreadDesigner  =   "M_Casino.frx":4258
         End
      End
      Begin VB.Frame Frame3 
         Height          =   8415
         Left            =   -74780
         TabIndex        =   74
         Top             =   690
         Width           =   12615
         Begin MSComctlLib.TreeView TvwDir 
            Height          =   7515
            Left            =   180
            TabIndex        =   75
            Top             =   600
            Width           =   12285
            _ExtentX        =   21669
            _ExtentY        =   13256
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   210
            TabIndex        =   76
            Top             =   270
            Width           =   555
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8775
         Left            =   105
         TabIndex        =   47
         Top             =   570
         Width           =   14895
         Begin VB.CheckBox Check5 
            Caption         =   "Integra SPRS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12720
            TabIndex        =   157
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Integra Ceco AMD"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   11400
            TabIndex        =   155
            Top             =   360
            Width           =   2055
         End
         Begin VB.Frame Frame25 
            Caption         =   "Tipo Acceso Minuta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   8880
            TabIndex        =   149
            Top             =   720
            Width           =   3735
            Begin VB.OptionButton Option5 
               Caption         =   "Minuta AMD"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   151
               Top             =   360
               Width           =   1335
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Minuta Normal"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   150
               Top             =   360
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin VB.Frame Frame24 
            Caption         =   "Org. Compras (Zona)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   12720
            TabIndex        =   146
            Top             =   2880
            Width           =   2055
            Begin EditLib.fpText fpText 
               Height          =   315
               Index           =   17
               Left            =   240
               TabIndex        =   147
               Top             =   480
               Width           =   1500
               _Version        =   196608
               _ExtentX        =   2646
               _ExtentY        =   556
               Enabled         =   0   'False
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
         End
         Begin VB.Frame Frame23 
            Caption         =   "Tipo Negocio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4215
            Left            =   12120
            TabIndex        =   144
            Top             =   4440
            Width           =   2655
            Begin VB.ListBox TipoNegocio 
               Height          =   3660
               Index           =   0
               ItemData        =   "M_Casino.frx":9B83
               Left            =   120
               List            =   "M_Casino.frx":9B8A
               Style           =   1  'Checkbox
               TabIndex        =   145
               Top             =   360
               Width           =   2415
            End
         End
         Begin VB.Frame Frame22 
            Caption         =   "Sellos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   11040
            TabIndex        =   143
            Top             =   2880
            Width           =   1575
            Begin VB.ListBox Sello 
               Height          =   960
               Index           =   0
               ItemData        =   "M_Casino.frx":9B9B
               Left            =   120
               List            =   "M_Casino.frx":9BA2
               Style           =   1  'Checkbox
               TabIndex        =   148
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.Frame Frame17 
            Caption         =   "Tipo Ceco"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   11040
            TabIndex        =   118
            Top             =   1680
            Width           =   1575
            Begin VB.OptionButton Option4 
               Caption         =   "Real"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   120
               Top             =   360
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Propuesta"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   119
               Top             =   720
               Width           =   1215
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Formato de Compras"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   8880
            TabIndex        =   110
            Top             =   2880
            Width           =   2055
            Begin VB.ComboBox Combo1 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   3
               ItemData        =   "M_Casino.frx":9BAD
               Left            =   75
               List            =   "M_Casino.frx":9BAF
               Style           =   2  'Dropdown List
               TabIndex        =   111
               Top             =   420
               Width           =   1725
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   11
               Left            =   105
               TabIndex        =   113
               Top             =   465
               Width           =   1725
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Tipo de Minuta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   12720
            TabIndex        =   108
            Top             =   1680
            Width           =   2055
            Begin VB.ComboBox Combo1 
               Appearance      =   0  'Flat
               Height          =   315
               Index           =   2
               ItemData        =   "M_Casino.frx":9BB1
               Left            =   90
               List            =   "M_Casino.frx":9BB3
               Style           =   2  'Dropdown List
               TabIndex        =   109
               Top             =   435
               Width           =   1725
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   12
               Left            =   120
               TabIndex        =   112
               Top             =   480
               Width           =   1725
            End
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Portal Electronico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   11520
            TabIndex        =   107
            Top             =   240
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CheckBox ChkBlockTraFinSemana 
            Caption         =   "Trabaja fin de semana"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7080
            TabIndex        =   104
            Top             =   8280
            Width           =   1575
         End
         Begin VB.CheckBox ChkBlockMinContrato 
            Caption         =   "Bloquear Minuta Contrato"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13080
            TabIndex        =   103
            Top             =   7680
            Width           =   1695
         End
         Begin VB.CheckBox ChkBlockMinTeo 
            Caption         =   "Bloquear Minuta Teorica"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13080
            TabIndex        =   102
            Top             =   7080
            Width           =   1695
         End
         Begin VB.CheckBox ChkBlockMinReal 
            Caption         =   "Bloquear Minuta Real"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12840
            TabIndex        =   101
            Top             =   7680
            Width           =   1695
         End
         Begin VB.CheckBox ChkMInRet 
            Caption         =   "Minuta Sitio Remoto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   11520
            TabIndex        =   25
            Top             =   480
            Width           =   2055
         End
         Begin VB.Frame Frame12 
            Caption         =   "Tipo Operación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   12240
            TabIndex        =   98
            Top             =   6150
            Visible         =   0   'False
            Width           =   2040
            Begin VB.OptionButton Option3 
               Caption         =   "Gravada"
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
               Left            =   120
               TabIndex        =   15
               Top             =   480
               Value           =   -1  'True
               Width           =   1695
            End
            Begin VB.OptionButton Option3 
               Caption         =   "No Gravada"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   16
               Top             =   840
               Width           =   1695
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Envio Hipersensibilidad Alimentaria"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   28
            Top             =   5160
            Width           =   3450
         End
         Begin VB.Frame Frame11 
            Caption         =   "Generar Pedido"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   8880
            TabIndex        =   95
            Top             =   1710
            Visible         =   0   'False
            Width           =   2055
            Begin VB.OptionButton Option3 
               Caption         =   "SGP"
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
               Left            =   120
               TabIndex        =   7
               Top             =   720
               Width           =   1695
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Pagina Web"
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
               Left            =   120
               TabIndex        =   6
               Top             =   360
               Value           =   -1  'True
               Width           =   1695
            End
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            ItemData        =   "M_Casino.frx":9BB5
            Left            =   6600
            List            =   "M_Casino.frx":9BB7
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   3105
            Width           =   1725
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
            Height          =   195
            Index           =   0
            Left            =   10320
            TabIndex        =   2
            Top             =   360
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Envio Grupo Vulnerable"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   4800
            Width           =   2370
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Activa Módulo Pedido Paciente "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   4320
            TabIndex        =   27
            Top             =   4800
            Width           =   3090
         End
         Begin VB.Frame Frame7 
            Caption         =   "Sobreescribre Receta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   48
            Top             =   5520
            Width           =   3735
            Begin VB.OptionButton Option1 
               Caption         =   "Solo Fijos"
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
               Left            =   240
               TabIndex        =   32
               Top             =   360
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Todos"
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
               Left            =   240
               TabIndex        =   33
               Top             =   720
               Width           =   1215
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Ninguno"
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
               Left            =   2520
               TabIndex        =   34
               Top             =   720
               Width           =   1095
            End
         End
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   1
            Left            =   1800
            TabIndex        =   1
            Top             =   330
            Width           =   6555
            _Version        =   196608
            _ExtentX        =   11562
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Left            =   120
            TabIndex        =   3
            Top             =   870
            Width           =   8250
            _Version        =   196608
            _ExtentX        =   14552
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Left            =   5715
            TabIndex        =   10
            Top             =   2010
            Width           =   2670
            _Version        =   196608
            _ExtentX        =   4710
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Index           =   4
            Left            =   4440
            TabIndex        =   5
            Top             =   1425
            Width           =   3990
            _Version        =   196608
            _ExtentX        =   7047
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Left            =   120
            TabIndex        =   4
            Top             =   1425
            Width           =   3840
            _Version        =   196608
            _ExtentX        =   6773
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Left            =   120
            TabIndex        =   0
            Top             =   330
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Left            =   2835
            TabIndex        =   13
            Top             =   3090
            Width           =   3405
            _Version        =   196608
            _ExtentX        =   6006
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Left            =   120
            TabIndex        =   8
            Top             =   2010
            Width           =   2505
            _Version        =   196608
            _ExtentX        =   4419
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Index           =   8
            Left            =   120
            TabIndex        =   11
            Top             =   2550
            Width           =   8250
            _Version        =   196608
            _ExtentX        =   14552
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Left            =   120
            TabIndex        =   12
            Top             =   3090
            Width           =   2445
            _Version        =   196608
            _ExtentX        =   4313
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Left            =   2835
            TabIndex        =   9
            Top             =   2010
            Width           =   2655
            _Version        =   196608
            _ExtentX        =   4683
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   3720
            Width           =   555
            _Version        =   196608
            _ExtentX        =   979
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   12
            Left            =   4125
            TabIndex        =   40
            Top             =   7485
            Width           =   3240
            _Version        =   196608
            _ExtentX        =   5715
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Index           =   11
            Left            =   720
            TabIndex        =   39
            Top             =   7485
            Width           =   3240
            _Version        =   196608
            _ExtentX        =   5715
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   1
            Left            =   4320
            TabIndex        =   19
            Top             =   3720
            Width           =   555
            _Version        =   196608
            _ExtentX        =   979
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Left            =   120
            TabIndex        =   21
            Top             =   4320
            Width           =   555
            _Version        =   196608
            _ExtentX        =   979
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
            Index           =   3
            Left            =   4320
            TabIndex        =   23
            Top             =   4320
            Width           =   555
            _Version        =   196608
            _ExtentX        =   979
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   13
            Left            =   3960
            TabIndex        =   29
            Top             =   5325
            Width           =   1245
            _Version        =   196608
            _ExtentX        =   2196
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            MaxLength       =   4
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
            Index           =   5
            Left            =   4320
            TabIndex        =   35
            Top             =   5940
            Width           =   555
            _Version        =   196608
            _ExtentX        =   979
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   14
            Left            =   5400
            TabIndex        =   30
            Top             =   5325
            Width           =   1245
            _Version        =   196608
            _ExtentX        =   2196
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            MaxLength       =   4
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
            Index           =   6
            Left            =   6840
            TabIndex        =   31
            Top             =   5325
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            MinValue        =   "0"
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
            Index           =   7
            Left            =   4320
            TabIndex        =   37
            Top             =   6585
            Width           =   555
            _Version        =   196608
            _ExtentX        =   979
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   15
            Left            =   720
            TabIndex        =   106
            Top             =   8400
            Width           =   6240
            _Version        =   196608
            _ExtentX        =   11007
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            MaxLength       =   100
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
         Begin VB.Frame Frame16 
            Caption         =   "Ofertas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4215
            Left            =   8880
            TabIndex        =   114
            Top             =   4440
            Width           =   2655
            Begin VB.ListBox List1 
               Height          =   3660
               Index           =   1
               ItemData        =   "M_Casino.frx":9BB9
               Left            =   120
               List            =   "M_Casino.frx":9BBB
               Style           =   1  'Checkbox
               TabIndex        =   115
               Top             =   360
               Width           =   2415
            End
         End
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   16
            Left            =   8520
            TabIndex        =   117
            Top             =   360
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            MaxLength       =   10
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
         Begin MSComDlg.CommonDialog CD 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label9 
            Caption         =   "No existe Inf. T.Gramaje"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12720
            TabIndex        =   158
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código Optimun"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   28
            Left            =   8520
            TabIndex        =   116
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Email Envio Pedidos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   27
            Left            =   2760
            TabIndex        =   105
            Top             =   8160
            Width           =   1860
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000002&
            BorderWidth     =   4
            X1              =   120
            X2              =   8280
            Y1              =   7920
            Y2              =   7920
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   10
            Left            =   5280
            TabIndex        =   38
            Top             =   6585
            Width           =   3045
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   5
            Left            =   4830
            Picture         =   "M_Casino.frx":9BBD
            Top             =   6480
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Región"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   4320
            TabIndex        =   96
            Top             =   6315
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Casino I. SAC"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   26
            Left            =   6840
            TabIndex        =   94
            Top             =   5100
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "C.Compras SAC"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   24
            Left            =   5400
            TabIndex        =   93
            Top             =   5100
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Municipio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   23
            Left            =   4320
            TabIndex        =   91
            Top             =   5685
            Width           =   795
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   4830
            Picture         =   "M_Casino.frx":9EC7
            Top             =   5835
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   9
            Left            =   5280
            TabIndex        =   36
            Top             =   5940
            Width           =   3045
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Datos Ejecutivo Contable"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   16
            Left            =   2760
            TabIndex        =   69
            Top             =   7005
            Width           =   2430
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000002&
            BorderWidth     =   4
            X1              =   120
            X2              =   8280
            Y1              =   6960
            Y2              =   6960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   4125
            TabIndex        =   68
            Top             =   7245
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Persona de Contactos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   720
            TabIndex        =   67
            Top             =   7275
            Width           =   1845
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1050
            TabIndex        =   18
            Top             =   3720
            Width           =   3045
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   600
            Picture         =   "M_Casino.frx":A1D1
            Top             =   3630
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Segmento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   66
            Top             =   3480
            Width           =   1260
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   6660
            TabIndex        =   65
            Top             =   3150
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Opción de Envio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   6540
            TabIndex        =   64
            Top             =   2880
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   2835
            TabIndex        =   63
            Top             =   2880
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Giro"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   62
            Top             =   2880
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Persona de Contactos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   61
            Top             =   2340
            Width           =   1845
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   5715
            TabIndex        =   60
            Top             =   1800
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fono Nº 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   2835
            TabIndex        =   59
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1800
            TabIndex        =   58
            Top             =   120
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   57
            Top             =   660
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comuna"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   56
            Top             =   1215
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ciudad"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   4440
            TabIndex        =   55
            Top             =   1215
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fono Nº 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   54
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   53
            Top             =   120
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Zona"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   4320
            TabIndex        =   52
            Top             =   3495
            Width           =   420
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   4830
            Picture         =   "M_Casino.frx":A4DB
            Top             =   3630
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   5280
            TabIndex        =   20
            Top             =   3720
            Width           =   3045
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Servicio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   51
            Top             =   4080
            Width           =   1080
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   600
            Picture         =   "M_Casino.frx":A7E5
            Top             =   4200
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   1050
            TabIndex        =   22
            Top             =   4320
            Width           =   3045
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Segmento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   4320
            TabIndex        =   50
            Top             =   4080
            Width           =   870
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   4830
            Picture         =   "M_Casino.frx":AAEF
            Top             =   4200
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   5280
            TabIndex        =   24
            Top             =   4320
            Width           =   3045
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sociedad SAP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   22
            Left            =   3960
            TabIndex        =   49
            Top             =   5100
            Width           =   1140
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   1065
            TabIndex        =   70
            Top             =   3765
            Width           =   3075
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   1065
            TabIndex        =   72
            Top             =   4365
            Width           =   3075
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   5295
            TabIndex        =   71
            Top             =   3765
            Width           =   3075
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   8
            Left            =   5295
            TabIndex        =   73
            Top             =   4365
            Width           =   3075
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   5295
            TabIndex        =   92
            Top             =   5985
            Width           =   3075
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   5295
            TabIndex        =   97
            Top             =   6630
            Width           =   3075
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7335
         Left            =   -74400
         TabIndex        =   45
         Top             =   870
         Width           =   11535
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   6615
            Left            =   1200
            TabIndex        =   46
            Top             =   480
            Width           =   9135
            _Version        =   393216
            _ExtentX        =   16113
            _ExtentY        =   11668
            _StockProps     =   64
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
            MaxRows         =   20
            SpreadDesigner  =   "M_Casino.frx":ADF9
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Espere. Buscando Información..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   156
         Top             =   9360
         Width           =   2685
      End
   End
End
Attribute VB_Name = "M_Casino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim ibusca As Long, i As Long, dia As Long, mes As Long, ano As Long, colori As Long
Dim itab As Integer
Dim modo As String, codigo As String, MsgTitulo As String

Dim dest       As Node
Dim sourcenode As Node
Dim nd         As Node
Dim rootNode   As Node
Dim nd2        As Node
Dim nd1        As Node
Dim ndl        As Node

Dim Nivel2     As Long
Dim Nivel3     As Long
Dim Nivel4     As Long
Dim Nivel5     As Long
Dim Nivel6     As Long

Dim Nodx       As Node
Dim Nod2       As Node
Dim Nod3       As Node
Dim Nod4       As Node
Dim Nod5       As Node
Dim Nod6       As Node

Dim Est        As Boolean
Dim estdes     As Boolean
Dim estact     As Boolean
Public lc_Aux  As String

Private KeyAscii As Variant
Private nivel    As Long

Private Sub Check1_Click(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub
itab = 1
SSTab1.TabEnabled(0) = False

Select Case Index

Case 0, 1, 2, 3
    
    DehabilitarOpciones

End Select
Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    
On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Check2_Click(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub

SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(6) = False
SSTab1.TabEnabled(8) = False
SSTab1.Tab = 7
SSTab1.TabEnabled(7) = True
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 13, 0, modo

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Check3_Click()
    
On Error GoTo Man_Error

    If Est Then Exit Sub
    itab = 1
    SSTab1.TabEnabled(0) = False
'    Select Case Index
'    Case 0, 1, 2, 3
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
        SSTab1.TabEnabled(6) = False
        SSTab1.TabEnabled(7) = False
        SSTab1.TabEnabled(8) = False
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True
'    End Select
    Gl_Ac_Botones Me, 13, 0, modo
    itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Check3_KeyPress(KeyAscii As Integer)
    
On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Check4_Click()

On Error GoTo Man_Error

If Est Then Exit Sub
itab = 1
SSTab1.TabEnabled(0) = False
    
    DehabilitarOpciones

Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Check5_Click()

On Error GoTo Man_Error

If Est Then Exit Sub
itab = 1
SSTab1.TabEnabled(0) = False
    
    DehabilitarOpciones

Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub ChkBlockMinContrato_Click()
    
On Error GoTo Man_Error

    If Est Then Exit Sub
    itab = 1
    SSTab1.TabEnabled(0) = False
'    Select Case Index
'    Case 0, 1, 2, 3
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
        SSTab1.TabEnabled(6) = False
        SSTab1.TabEnabled(7) = False
        SSTab1.TabEnabled(8) = False
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True
'    End Select
    Gl_Ac_Botones Me, 13, 0, modo
    itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub ChkBlockMinReal_Click()
    
On Error GoTo Man_Error

    If Est Then Exit Sub
    itab = 1
    SSTab1.TabEnabled(0) = False
'    Select Case Index
'    Case 0, 1, 2, 3
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
        SSTab1.TabEnabled(6) = False
        SSTab1.TabEnabled(7) = False
        SSTab1.TabEnabled(8) = False
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True
'    End Select
    Gl_Ac_Botones Me, 13, 0, modo
    itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub ChkBlockMinTeo_Click()
    
On Error GoTo Man_Error

    If Est Then Exit Sub
    itab = 1
    SSTab1.TabEnabled(0) = False
'    Select Case Index
'    Case 0, 1, 2, 3
    DehabilitarOpciones
'    End Select
    Gl_Ac_Botones Me, 13, 0, modo
    itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub ChkBlockTraFinSemana_Click()
    
On Error GoTo Man_Error

    If Est Then Exit Sub
    itab = 1
    SSTab1.TabEnabled(0) = False
'    Select Case Index
'    Case 0, 1, 2, 3
     DehabilitarOpciones
'    End Select
    Gl_Ac_Botones Me, 13, 0, modo
    itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub ChkMInRet_Click()
    
On Error GoTo Man_Error

    If Est Then Exit Sub
    itab = 1
    SSTab1.TabEnabled(0) = False
'    Select Case Index
'    Case 0, 1, 2, 3
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
        SSTab1.TabEnabled(6) = False
        SSTab1.TabEnabled(7) = False
        SSTab1.TabEnabled(8) = False
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True
'    End Select
    Gl_Ac_Botones Me, 13, 0, modo
    itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub ChkMInRet_KeyPress(KeyAscii As Integer)
    
On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Combo1_Click(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub
If Index = 0 Then FptNombre.text = "": Exit Sub
DehabilitarOpciones
Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Combo2_Click(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub
itab = 1
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(6) = False
SSTab1.TabEnabled(8) = False
SSTab1.Tab = 7
SSTab1.TabEnabled(7) = True
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Command2_Click()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
Set RS = vg_db.Execute("sgpadm_Sel_ValidarParametroInvCalendarizado '" & codigo & "', '" & fpCalendar2.Year & "'")

If Not RS.EOF Then

   RS.Close
   Set RS = Nothing

   M_CopiaInvCalendarizado.LlenarDatos codigo, fpCalendar2.Year
   M_CopiaInvCalendarizado.Show 1
   Me.Refresh
   
Else

   fg_descarga
        
   MsgBox "Ceco no tiene información a copiar... ", vbCritical, MsgTitulo
      
   RS.Close
   Set RS = Nothing

End If

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Command3_Click()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
Set RS = vg_db.Execute("sgpadm_Sel_ValidarParametroCategoriaDietetica_V01 '" & codigo & "'")

If Not RS.EOF Then

   RS.Close
   Set RS = Nothing

   M_CopiaParamCategoriaDietetica.LlenarDatos codigo
   M_CopiaParamCategoriaDietetica.Show 1
   Me.Refresh
   
Else

   fg_descarga
        
   MsgBox "Ceco no tiene información a copiar... ", vbCritical, MsgTitulo
      
   RS.Close
   Set RS = Nothing

End If

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

    Dim RS As New ADODB.Recordset
    
    Me.HelpContextID = vg_OpcM
    Me.Height = 10665
    Me.Width = 15525 '12135
    
    MsgTitulo = IIf(lc_Aux = "MCasino", "Casino", "Parametro Despachos")
    fg_centra Me
    SSTab1.Tab = 0
    modo = ""
    Est = True
    estdes = True
    estact = True
    
    Label8.Visible = False
    Combo1(0).ListIndex = 1
    Combo1(1).Clear
    Combo1(1).AddItem "Ftp" & Space(150) & "(1)"
    Combo1(1).AddItem "Mail" & Space(150) & "(2)"
    Combo1(1).AddItem "Archivo" & Space(150) & "(3)"
    Combo1(1).ListIndex = -1
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
  
    Set RS = vg_db.Execute("sgpadm_Sel_CargarTipoMinuta")
  
    Combo1(2).Clear
    
    If Not RS.EOF Then
   
       Do While Not RS.EOF
    
          Combo1(2).AddItem RS!tip_descripcion & Space(150) & "(" & RS!tip_codigo & ")"
'          Combo1(2).AddItem "Mio" & Space(150) & "(2)"
'          Combo1(2).AddItem "Simap" & Space(150) & "(3)"
'          Combo1(2).ListIndex = -1
          
          RS.MoveNext
          
      Loop
      
    End If
    RS.Close
    Set RS = Nothing
    
    Combo1(2).ListIndex = -1
    
'    Combo1(3).ListIndex = 1
    Combo1(3).Clear
    Combo1(3).AddItem "Grande" & Space(150) & "(1)"
    Combo1(3).AddItem "Chico" & Space(150) & "(2)"
    Combo1(3).ListIndex = -1
    
    '-------> Llenar combo opciòn tipo parametro stock
    Combo2(0).Clear
    Combo2(0).AddItem "Curva ABC" & Space(150) & "(1)"
    Combo2(0).AddItem "Lista General" & Space(150) & "(2)"
    Combo2(0).ListIndex = -1
       
    Gl_Mo_Botones Me, 13
    Frame4.Left = 105 '480
    
    If lc_Aux = "MCasino" Then
       
       SSTab1.TabVisible(5) = False
       SSTab1.TabVisible(9) = False
       SSTab1.TabVisible(10) = False
       SSTab1.TabVisible(11) = False
       SSTab1.TabVisible(12) = False

       Frame11.Visible = IIf(vg_pais = "CL", True, False)
       Frame12.Visible = IIf(vg_pais = "CL", False, True)
       Gl_Ac_Botones Me, 13, 1, modo
    
    ElseIf lc_Aux = "MCasppr" Then
    
       Me.Caption = "Parametrización (Q) Servicios Principales"
       MsgTitulo = "Parametrización Servicios Principales"
       Gl_Ac_Botones Me, 13, 3, modo
       SSTab1.TabVisible(1) = False
       SSTab1.TabVisible(2) = False
       SSTab1.TabVisible(3) = False
       SSTab1.TabVisible(4) = False
       SSTab1.TabVisible(5) = False
       SSTab1.TabVisible(6) = False
       SSTab1.TabVisible(7) = False
       SSTab1.TabVisible(8) = False
    
       SSTab1.TabVisible(9) = True
       SSTab1.TabVisible(10) = False
       SSTab1.TabVisible(11) = False
       SSTab1.TabVisible(12) = False

    
    ElseIf lc_Aux = "MCaspic" Then
    
       Me.Caption = "Parametrización Inventario Calendarizado"
       MsgTitulo = "Parametrización Inventario Calendarizado"
       Gl_Ac_Botones Me, 13, 3, modo
       SSTab1.TabVisible(1) = False
       SSTab1.TabVisible(2) = False
       SSTab1.TabVisible(3) = False
       SSTab1.TabVisible(4) = False
       SSTab1.TabVisible(5) = False
       SSTab1.TabVisible(6) = False
       SSTab1.TabVisible(7) = False
       SSTab1.TabVisible(8) = False
       SSTab1.TabVisible(9) = False
       SSTab1.TabVisible(11) = False
       SSTab1.TabVisible(12) = False
   
       SSTab1.TabVisible(10) = True
    
    ElseIf lc_Aux = "MCaspcc" Then
    
       Me.Caption = "Parametrización Pvta. Cliente Calendarizado"
       MsgTitulo = "Parametrización Pvta. Cliente Calendarizado"
       Gl_Ac_Botones Me, 13, 3, modo
       SSTab1.TabVisible(1) = False
       SSTab1.TabVisible(2) = False
       SSTab1.TabVisible(3) = False
       SSTab1.TabVisible(4) = False
       SSTab1.TabVisible(5) = False
       SSTab1.TabVisible(6) = False
       SSTab1.TabVisible(7) = False
       SSTab1.TabVisible(8) = False
       SSTab1.TabVisible(9) = False
       SSTab1.TabVisible(10) = False
       
       SSTab1.TabVisible(11) = True
       SSTab1.TabVisible(12) = False
    
    ElseIf lc_Aux = "MCaspcd" Then
    
       Me.Caption = "Parametrización Categoría Dietética"
       MsgTitulo = "Parametrización Categoría Dietética"
       Gl_Ac_Botones Me, 13, 3, modo
       SSTab1.TabVisible(1) = False
       SSTab1.TabVisible(2) = False
       SSTab1.TabVisible(3) = False
       SSTab1.TabVisible(4) = False
       SSTab1.TabVisible(5) = False
       SSTab1.TabVisible(6) = False
       SSTab1.TabVisible(7) = False
       SSTab1.TabVisible(8) = False
       SSTab1.TabVisible(9) = False
       SSTab1.TabVisible(10) = False
       
       SSTab1.TabVisible(11) = False
       SSTab1.TabVisible(12) = True
    
    Else
       
       Gl_Ac_Botones Me, 13, 3, modo
       SSTab1.TabVisible(1) = False
       SSTab1.TabVisible(2) = False
       SSTab1.TabVisible(3) = False
       SSTab1.TabVisible(4) = False
       SSTab1.TabVisible(6) = False
       SSTab1.TabVisible(7) = False
       SSTab1.TabVisible(8) = False
       SSTab1.TabVisible(9) = False
       SSTab1.TabVisible(10) = False
       SSTab1.TabVisible(11) = False
       SSTab1.TabVisible(12) = False
    
    End If
    
    MoverDatosGrilla
       
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Form_Resize()

On Error GoTo Man_Error

If Me.WindowState <> 1 Then SSTab1.Move 0, Toolbar1.Height, ScaleWidth, ScaleHeight - Toolbar1.Height

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar1_AfterSelection()

On Error GoTo Man_Error

fpCalendar1.Element = ElementSelection
fpCalendar1.ElementIndex = fg_pone_cero(ano, 4) & fg_pone_cero(mes, 2) & fg_pone_cero(dia, 2)

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar1_BeforeSelection(Cancel As Integer)

On Error GoTo Man_Error

fpCalendar1.Element = ElementSelection
fpCalendar1.ElementIndex = fg_pone_cero(ano, 4) & fg_pone_cero(mes, 2) & fg_pone_cero(dia, 2)
If fpCalendar1.ElementBackColor = &HFF& Then
   
   Cancel = True
   fpCalendar1.ElementBackColor = -2147483633 'colori
   fpCalendar1.ElementForeColor = vbBlack

Else
   
   Cancel = False

End If

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar1_DateChanging(Month As Integer, Day As Integer, Year As Integer, State As Integer, ByVal Shift As Integer, Cancel As Integer)

On Error GoTo Man_Error

dia = Day
mes = Month
ano = Year

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar1_DblClick(CurrentMonth As Integer, CurrentDay As Integer, CurrentYear As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub
vg_nombre = ""
M_Feriado.LlenaDatos codigo, fg_pone_cero(CurrentDay, 2) & "/" & fg_pone_cero(CurrentMonth, 2) & "/" & fg_pone_cero(CurrentYear, 4), 1, ""
M_Feriado.Show 1, M_Casino
If Trim(vg_nombre) = "" Then Exit Sub
fpCalendar1.Element = ElementSpecificDate
fpCalendar1.ElementIndex = fg_pone_cero(CurrentYear, 4) & fg_pone_cero(CurrentMonth, 2) & fg_pone_cero(CurrentDay, 2)
fpCalendar1.ElementBackColor = &HFF&
fpCalendar1.ElementForeColor = vbBlack
fpCalendar1.ElementText = Trim(vg_nombre)
'fpCalendar1.MultiSelect = MultiSelectSimple
'fpCalendar1.DrawFocusRect = 2
If modo = "" Then modo = "M"
itab = 1
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.Tab = 4
SSTab1.TabEnabled(4) = True
SSTab1.TabEnabled(6) = False
SSTab1.TabEnabled(7) = False
SSTab1.TabEnabled(8) = False
Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar1_ViewChange(BeginMonth As Integer, BeginDay As Integer, BeginYear As Integer, EndMonth As Integer, EndDay As Integer, EndYear As Integer, Cancel As Integer)

On Error GoTo Man_Error

Cancel = IIf(Toolbar1.Buttons(12).Visible = True, True, False)
If Cancel = False Then MoverDiasFeriados

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar2_AfterSelection()

On Error GoTo Man_Error

fpCalendar2.Element = ElementSelection
fpCalendar2.ElementIndex = fg_pone_cero(ano, 4) & fg_pone_cero(mes, 2) & fg_pone_cero(dia, 2)

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar2_BeforeSelection(Cancel As Integer)

On Error GoTo Man_Error

fpCalendar2.Element = ElementSelection
fpCalendar2.ElementIndex = fg_pone_cero(ano, 4) & fg_pone_cero(mes, 2) & fg_pone_cero(dia, 2)
If fpCalendar2.ElementBackColor = &HFF& Then

   Cancel = True
   fpCalendar2.ElementBackColor = -2147483633 'colori
   fpCalendar2.ElementForeColor = vbBlack

Else

   Cancel = False

End If

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar2_DateChanging(Month As Integer, Day As Integer, Year As Integer, State As Integer, ByVal Shift As Integer, Cancel As Integer)

On Error GoTo Man_Error

dia = Day
mes = Month
ano = Year

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar2_DblClick(CurrentMonth As Integer, CurrentDay As Integer, CurrentYear As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub
vg_nombre = "dia"
'M_Feriado.LlenaDatos codigo, fg_pone_cero(CurrentDay, 2) & "/" & fg_pone_cero(CurrentMonth, 2) & "/" & fg_pone_cero(CurrentYear, 4), 1, ""
'M_Feriado.Show 1, M_Casino
'If Trim(vg_nombre) = "" Then Exit Sub
fpCalendar2.Element = ElementSpecificDate
fpCalendar2.ElementIndex = fg_pone_cero(CurrentYear, 4) & fg_pone_cero(CurrentMonth, 2) & fg_pone_cero(CurrentDay, 2)
fpCalendar2.ElementBackColor = &HFF&
fpCalendar2.ElementForeColor = vbBlack
fpCalendar2.ElementText = Trim(vg_nombre)
'fpCalendar2.MultiSelect = MultiSelectSimple
'fpCalendar2.DrawFocusRect = 2
If modo = "" Then modo = "M"
itab = 1
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(6) = False
SSTab1.TabEnabled(7) = False
SSTab1.TabEnabled(8) = False
SSTab1.TabEnabled(9) = False
SSTab1.TabEnabled(11) = False
SSTab1.TabVisible(12) = False

SSTab1.Tab = 10
SSTab1.TabEnabled(10) = True

Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub

Select Case KeyCode

    Case 46
    
'        fpCalendar2.Element = ElementSelection
        fpCalendar2.Element = ElementSpecificDate
        
        fpCalendar2.ElementIndex = fg_pone_cero(ano, 4) & fg_pone_cero(mes, 2) & fg_pone_cero(dia, 2)
        If fpCalendar2.ElementBackColor = -2147483635 Or fpCalendar2.ElementBackColor = &HFF& Then
           
           'Cancel = True
           fpCalendar2.ElementBackColor = -2147483633 'colori
           fpCalendar2.ElementText = ""
           fpCalendar2.ElementForeColor = vbBlack
        
           If modo = "" Then modo = "M"
           itab = 1
           SSTab1.TabEnabled(0) = False
           SSTab1.TabEnabled(1) = False
           SSTab1.TabEnabled(2) = False
           SSTab1.TabEnabled(3) = False
           SSTab1.TabEnabled(4) = False
           SSTab1.TabEnabled(6) = False
           SSTab1.TabEnabled(7) = False
           SSTab1.TabEnabled(8) = False
           SSTab1.TabEnabled(9) = False
           SSTab1.TabEnabled(11) = False
           SSTab1.TabVisible(12) = False
           
           SSTab1.Tab = 10
           SSTab1.TabEnabled(10) = True
        
           Gl_Ac_Botones Me, 13, 0, modo
           itab = 0

'        Else
'
'            fpCalendar2.ElementBackColor = -2147483633 'colori
'           fpCalendar2.ElementForeColor = vbBlack
'          'Cancel = False
        
        End If

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar2_ViewChange(BeginMonth As Integer, BeginDay As Integer, BeginYear As Integer, EndMonth As Integer, EndDay As Integer, EndYear As Integer, Cancel As Integer)

On Error GoTo Man_Error

'Cancel = IIf(Toolbar1.Buttons(12).Visible = True, False, True)
Cancel = False
MoverInvCandelarizado
modo = ""
Gl_Ac_Botones Me, 13, IIf(lc_Aux = "MCasino", 1, 3), modo
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar3_AfterSelection()

On Error GoTo Man_Error

fpCalendar3.Element = ElementSelection
fpCalendar3.ElementIndex = fg_pone_cero(ano, 4) & fg_pone_cero(mes, 2) & fg_pone_cero(dia, 2)

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar3_BeforeSelection(Cancel As Integer)

On Error GoTo Man_Error

fpCalendar3.Element = ElementSelection
fpCalendar3.ElementIndex = fg_pone_cero(ano, 4) & fg_pone_cero(mes, 2) & fg_pone_cero(dia, 2)
If fpCalendar3.ElementBackColor = &HFF& Then

   Cancel = True
   fpCalendar3.ElementBackColor = -2147483633 'colori
   fpCalendar3.ElementForeColor = vbBlack

Else

   Cancel = False

End If

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar3_DateChanging(Month As Integer, Day As Integer, Year As Integer, State As Integer, ByVal Shift As Integer, Cancel As Integer)

On Error GoTo Man_Error

dia = Day
mes = Month
ano = Year

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar3_DblClick(CurrentMonth As Integer, CurrentDay As Integer, CurrentYear As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub
vg_nombre = "dia"
'M_Feriado.LlenaDatos codigo, fg_pone_cero(CurrentDay, 2) & "/" & fg_pone_cero(CurrentMonth, 2) & "/" & fg_pone_cero(CurrentYear, 4), 1, ""
'M_Feriado.Show 1, M_Casino
'If Trim(vg_nombre) = "" Then Exit Sub
fpCalendar3.Element = ElementSpecificDate
fpCalendar3.ElementIndex = fg_pone_cero(CurrentYear, 4) & fg_pone_cero(CurrentMonth, 2) & fg_pone_cero(CurrentDay, 2)
fpCalendar3.ElementBackColor = &HFF&
fpCalendar3.ElementForeColor = vbBlack
fpCalendar3.ElementText = Trim(vg_nombre)
'fpCalendar3.MultiSelect = MultiSelectSimple
'fpCalendar3.DrawFocusRect = 2
If modo = "" Then modo = "M"
itab = 1
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(6) = False
SSTab1.TabEnabled(7) = False
SSTab1.TabEnabled(8) = False
SSTab1.TabEnabled(9) = False
SSTab1.TabEnabled(10) = False
SSTab1.Tab = 11
SSTab1.TabEnabled(11) = True
SSTab1.TabVisible(12) = False

Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar3_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub

Select Case KeyCode

    Case 46
    
'        fpCalendar3.Element = ElementSelection
        fpCalendar3.Element = ElementSpecificDate
        
        fpCalendar3.ElementIndex = fg_pone_cero(ano, 4) & fg_pone_cero(mes, 2) & fg_pone_cero(dia, 2)
        If fpCalendar3.ElementBackColor = -2147483635 Or fpCalendar3.ElementBackColor = &HFF& Then
           
           'Cancel = True
           fpCalendar3.ElementBackColor = -2147483633 'colori
           fpCalendar3.ElementText = ""
           fpCalendar3.ElementForeColor = vbBlack
        
           If modo = "" Then modo = "M"
           itab = 1
           SSTab1.TabEnabled(0) = False
           SSTab1.TabEnabled(1) = False
           SSTab1.TabEnabled(2) = False
           SSTab1.TabEnabled(3) = False
           SSTab1.TabEnabled(4) = False
           SSTab1.TabEnabled(6) = False
           SSTab1.TabEnabled(7) = False
           SSTab1.TabEnabled(8) = False
           SSTab1.TabEnabled(9) = False
           SSTab1.TabEnabled(10) = False
           SSTab1.Tab = 11
           SSTab1.TabEnabled(11) = True
           SSTab1.TabVisible(12) = False

           Gl_Ac_Botones Me, 13, 0, modo
           itab = 0
       
        End If

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpCalendar3_ViewChange(BeginMonth As Integer, BeginDay As Integer, BeginYear As Integer, EndMonth As Integer, EndDay As Integer, EndYear As Integer, Cancel As Integer)

On Error GoTo Man_Error

'Cancel = IIf(Toolbar1.Buttons(12).Visible = True, False, True)
Cancel = False
MoverPvtaClienteCalendarizado
modo = ""
Gl_Ac_Botones Me, 13, IIf(lc_Aux = "MCasino", 1, 3), modo
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
    If Est Then Exit Sub
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Select Case Index
        
        Case 0
            
            Set RS = vg_db.Execute("sgpadm_s_subsegmento 1, " & Val(fpLongInteger1(0).Value) & ",'', ''")
            fpayuda(1).Caption = ""
            If Not RS.EOF Then fpayuda(1).Caption = Trim(RS!sub_nombre)
            RS.Close
            Set RS = Nothing
            DehabilitarOpciones
            Gl_Ac_Botones Me, 13, 0, modo
            itab = 0
        
        Case 1
            
            Set RS = vg_db.Execute("sgpadm_s_zona 1, " & Val(fpLongInteger1(1).Value) & ",''")
            fpayuda(0).Caption = ""
            If Not RS.EOF Then fpayuda(0).Caption = Trim(RS!Zon_nombre)
            RS.Close
            Set RS = Nothing
            DehabilitarOpciones
            Gl_Ac_Botones Me, 13, 0, modo
            itab = 0
        
        Case 2
            
            Set RS = vg_db.Execute("sgpadm_s_tiposervicio 1, " & Val(fpLongInteger1(2).Value) & ",''")
            fpayuda(3).Caption = ""
            If Not RS.EOF Then fpayuda(3).Caption = Trim(RS!tis_nombre)
            RS.Close
            Set RS = Nothing
            DehabilitarOpciones
            Gl_Ac_Botones Me, 13, 0, modo
            itab = 0
        
        Case 3
            
            Set RS = vg_db.Execute("sgpadm_s_segmento 1, " & Val(fpLongInteger1(3).Value) & ",''")
            fpayuda(7).Caption = ""
            If Not RS.EOF Then fpayuda(7).Caption = Trim(RS!seg_nombre)
            RS.Close
            Set RS = Nothing
            DehabilitarOpciones
            Gl_Ac_Botones Me, 13, 0, modo
            itab = 0
        
        Case 4
            
            itab = 1
            SSTab1.TabEnabled(0) = False
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.TabEnabled(4) = False
            SSTab1.TabEnabled(6) = False
            SSTab1.TabEnabled(8) = False
            SSTab1.Tab = 7
            SSTab1.TabEnabled(7) = True
            
            If modo = "" Then
            
               modo = "M"
               Gl_Ac_Botones Me, 13, 0, modo
               itab = 0
            
            End If
            
        Case 5
            
            Set RS = vg_db.Execute("sgpadm_s_municipio 1, " & Val(fpLongInteger1(5).Value) & ",''")
            fpayuda(9).Caption = ""
            If Not RS.EOF Then fpayuda(9).Caption = Trim(RS!mun_nombre)
            RS.Close
            Set RS = Nothing
            DehabilitarOpciones
            Gl_Ac_Botones Me, 13, 0, modo
            itab = 0
        
        Case 6
            
            DehabilitarOpciones
            Gl_Ac_Botones Me, 13, 0, modo
            itab = 0
        
        Case 7
            
            Set RS = vg_db.Execute("sgpadm_s_region 1, " & Val(fpLongInteger1(7).Value) & ",''")
            fpayuda(10).Caption = ""
            If Not RS.EOF Then fpayuda(10).Caption = Trim(RS!reg_nombre)
            RS.Close
            Set RS = Nothing
            DehabilitarOpciones
            Gl_Ac_Botones Me, 13, 0, modo
            itab = 0
    
        Case 8
        
            If modo = "" Then modo = "M"
            itab = 1
            SSTab1.TabEnabled(0) = False
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.TabEnabled(4) = False
            SSTab1.TabEnabled(6) = False
            SSTab1.TabEnabled(7) = False
            SSTab1.TabEnabled(8) = False
            SSTab1.TabEnabled(9) = False
            SSTab1.TabEnabled(11) = False
            SSTab1.TabVisible(12) = False
            
            SSTab1.Tab = 10
            SSTab1.TabEnabled(10) = True
            Gl_Ac_Botones Me, 13, 0, modo
            itab = 0
            
        Case 9

            If modo = "" Then modo = "M"
            itab = 1
            SSTab1.TabEnabled(0) = False
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.TabEnabled(4) = False
            SSTab1.TabEnabled(6) = False
            SSTab1.TabEnabled(7) = False
            SSTab1.TabEnabled(8) = False
            SSTab1.TabEnabled(9) = False
            SSTab1.TabEnabled(11) = False
            SSTab1.TabVisible(12) = False
            
            SSTab1.Tab = 10
            SSTab1.TabEnabled(10) = True
            Gl_Ac_Botones Me, 13, 0, modo
            itab = 0
            
    End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpText_Change(Index As Integer)

If Est Then Exit Sub
DehabilitarOpciones
Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpTnombre_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
If LimpiaDato(Trim(FptNombre.text)) & Chr(KeyAscii) = "" Then Exit Sub
vaSpread1.Visible = False
codigo = ""

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Combo1(0).ItemData(Combo1(0).ListIndex) = 0 Then
   
   Sql = IIf(lc_Aux = "MCasino", "sgpadm_s_cliente_V02 3, '', '%" & UCase(LimpiaDato(FptNombre.text)) & "%'", "sgpadm_s_cliente_V02 49, '', '%" & UCase(LimpiaDato(FptNombre.text)) & "%'")
   Set RS = vg_db.Execute(Sql)
   If RS.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS!nReg

ElseIf Combo1(0).ItemData(Combo1(0).ListIndex) = 1 Then
   
   Sql = IIf(lc_Aux = "MCasino", "sgpadm_s_cliente_V02 4, '', '%" & UCase(LimpiaDato(FptNombre.text)) & "%'", "sgpadm_s_cliente_V02 50, '', '%" & UCase(LimpiaDato(FptNombre.text)) & "%'")
   Set RS = vg_db.Execute(Sql)
   If RS.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS!nReg

End If
i = 1

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      vaSpread1.Row = i

      vaSpread1.Row = i: vaSpread1.Col = -1
      vaSpread1.BackColor = Shape1(0).FillColor
      
      vaSpread1.Col = 1
      vaSpread1.text = RS!Cli_codigo
      
      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = 0
      vaSpread1.text = Trim(RS!Cli_nombre)
      
      vaSpread1.Col = 3
      vaSpread1.text = IIf(RS!cli_tipo = 0, "Operador", "Traspaso")
      
      vaSpread1.Col = 4
      
      If Not IsNull(RS!cli_openvio) Then
         
         If RS!cli_openvio = 1 Then
            
            vaSpread1.text = "Ftp"
         
         ElseIf RS!cli_openvio = 2 Then
            
            vaSpread1.text = "Mail"
         
         ElseIf RS!cli_openvio = 3 Then
            
            vaSpread1.text = "Archivo"
         
         End If
      
      Else
         
         vaSpread1.text = "No Especificado"
      
      End If
      
      vaSpread1.Col = 5
      vaSpread1.text = IIf(IsNull(RS!sub_nombre), "", Trim(RS!sub_nombre))
      
      vaSpread1.Col = 6
      vaSpread1.text = IIf(IsNull(RS!CLI_ACTIVO) Or Trim(RS!CLI_ACTIVO) = "0", "", Trim(RS!CLI_ACTIVO))
      
      RS.MoveNext
      i = i + 1
   
   Loop
   
   vaSpread1.Row = 1
   vaSpread1.Col = 1
   codigo = Trim(vaSpread1.text)
   
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
   SSTab1.TabEnabled(3) = True
   SSTab1.TabEnabled(4) = True
   SSTab1.TabEnabled(5) = True
   SSTab1.TabEnabled(8) = True
   
   modo = ""
   Gl_Ac_Botones Me, 13, IIf(lc_Aux = "MCasino", 1, 3), modo

Else
   
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
   SSTab1.TabEnabled(3) = False
   SSTab1.TabEnabled(4) = False
   SSTab1.TabEnabled(5) = False
   SSTab1.TabEnabled(8) = False

End If
RS.Close
Set RS = Nothing
vaSpread1.Visible = True
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

vg_left = fpayuda(1).Left + 2300
vg_nombre = "": vg_codigo = ""

Select Case Index

Case 0
    
    B_TabEst.LlenaDatos "a_subsegmento", "sub_", "Subsegmento", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(1).Caption = Trim(vg_nombre)

Case 1
    
    B_TabEst.LlenaDatos "a_zona", "zon_", "Zona", "Zon"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(0).Caption = Trim(vg_nombre)

Case 2
    
    B_TabEst.LlenaDatos "a_tiposervicio", "tis_", "Tipo de Servicio", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(3).Caption = Trim(vg_nombre)

Case 3
    
    B_TabEst.LlenaDatos "a_segmento", "seg_", "Segmento", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(3).Value = Val(vg_codigo)
    fpayuda(7).Caption = Trim(vg_nombre)

Case 4
    
    B_TabEst.LlenaDatos "a_municipio", "mun_", "Municipio", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(5).Value = Val(vg_codigo)
    fpayuda(9).Caption = Trim(vg_nombre)

Case 5
    
    B_TabEst.LlenaDatos "a_region", "reg_", "Región", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(7).Value = Val(vg_codigo)
    fpayuda(10).Caption = Trim(vg_nombre)

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub List1_Click(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub
DehabilitarOpciones
Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub
DehabilitarOpciones
Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Option2_Click(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(6) = False
SSTab1.TabEnabled(8) = False
SSTab1.Tab = 7
SSTab1.TabEnabled(7) = True
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 13, 0, modo

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Option3_Click(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub
DehabilitarOpciones
Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Option5_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If Est Then Exit Sub

If Option5(0).Value = True And modo <> "A" Then

    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_ValidarNuevoFormatoDetalleMinuta_V01 '" & fpText(0).text & "'")

    If Not RS.EOF Then

       If RS(0) > 0 Then
       
          If MsgBox("Si cambia esta opción, a minuta bloque normal, cuando vaya grabar en la minuta bloque normal, se perderan los items categoria dietetica y tipo de plato" & "  Esta seguro de realizar el cambio...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
          
             Est = True
             Option5(0).Value = False
             Option5(1).Value = True
             RS.Close
             Set RS = Nothing
             itab = 0
             Est = False
             Exit Sub
             
          End If
       
       End If

    End If
    
    RS.Close
    Set RS = Nothing

End If

DehabilitarOpciones
Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Option4_Click(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub
DehabilitarOpciones
Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Select Case SSTab1.Tab
        
        Case 0
            
            ActivarBotones
            
            For i = 0 To 12
                
                fpText(i).Enabled = False
            
            Next i
            
            Combo1(0).Enabled = True
            FptNombre.Enabled = True
        
        Case 1
            
            ActivarBotones
            
            If vaSpread1.MaxRows > 0 And (itab = 0 Or itab = 1) Then  '0 Then
               
               If modo = "A" Then
                  
                  Gl_Ac_Botones Me, 13, 0, modo
               
               Else
                  
                  modo = "M"
                  SSTab1.TabEnabled(0) = True
                  SSTab1.Tab = 1
                  SSTab1.TabEnabled(1) = True
                  MoverDatos
               
               End If
            
            End If
        
        Case 2
            
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = True
            If vaSpread1.MaxRows < 1 Then Exit Sub
            modo = "M"
            MoverDatos2
        
        Case 3
            
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = True
            If vaSpread1.MaxRows < 1 Then Exit Sub
            modo = "M"
            MoverDatos3
        
        Case 4 '-------> Parametro Días Feriados
            
            ActivarBotones
            If vaSpread1.MaxRows < 1 Then Exit Sub
            SSTab1.TabEnabled(0) = True
            MoverDatos4
        
        Case 5 '-------> parametro Despacho
            
            ActivarBotones
            If vaSpread1.MaxRows < 1 Then Exit Sub
            MoverDatos5
        
        Case 6 '-------> actividades diarias
            
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = True
            If vaSpread1.MaxRows < 1 Then Exit Sub
            MoverDatos6
        
        Case 7 '-------> Parametros de Stock
            
            If vaSpread1.MaxRows < 1 Then Exit Sub
            MoverDatos7
        
        Case 8
            
            If vaSpread1.MaxRows < 1 Then Exit Sub
            MoverDatos8
    
        Case 9 'Servicio Preferido
        
            If vaSpread1.MaxRows < 1 Then Exit Sub
            MoverDatos9
        
        Case 10 '-------> Parametro stock
            
            ActivarBotones
            If vaSpread1.MaxRows < 1 Then Exit Sub
            SSTab1.TabEnabled(0) = True
            MoverDatos10
        
        Case 11 '-------> Parametro precio venta cliente calendarizado
            
            ActivarBotones
            If vaSpread1.MaxRows < 1 Then Exit Sub
            SSTab1.TabEnabled(0) = True
            MoverDatos11
        
        Case 12 '-------> Parametro categoria dietetica
        
            ActivarBotones
            If vaSpread1.MaxRows < 1 Then Exit Sub
            SSTab1.TabEnabled(0) = True
            MoverDatos12
        
    End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub TipoNegocio_Click(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub
DehabilitarOpciones
Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub sello_Click(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub
DehabilitarOpciones
Gl_Ac_Botones Me, 13, 0, modo
itab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

    Select Case Button.Index
        
        Case 1
            
            modo = "A"
            Check3.Visible = False
            ChkMInRet.Visible = False
            ChkBlockMinContrato.Visible = False
            ChkBlockMinTeo.Visible = False
            ChkBlockMinReal.Visible = False
            ChkBlockTraFinSemana.Visible = True 'False
            Combo1(2).ListIndex = -1
            Combo1(3).ListIndex = -1
            Gl_Ac_Botones Me, 13, 0, modo
            SSTab1.TabEnabled(0) = False
            itab = 1
            Est = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.TabEnabled(4) = False
            SSTab1.TabEnabled(6) = False
            SSTab1.TabEnabled(7) = False
            SSTab1.TabEnabled(8) = False
            Check1(1).Value = 0
            Check1(2).Value = 0
            Check1(0).Value = 1
            Check1(3).Value = 0
            Check4.Value = 0
            Check5.Value = 0
            Option1(0).Value = True
            Option3(3).Value = True
            Option3(2).Value = False
            fpayuda(0).Caption = ""
            fpayuda(1).Caption = ""
            fpayuda(3).Caption = ""
            fpayuda(7).Caption = ""
            fpayuda(9).Caption = ""
            fpayuda(10).Caption = ""
            
            For i = 0 To 17
                
                If i < 17 Then fpText(i).Enabled = True: fpText(i).text = ""
                If i < 4 Then fpLongInteger1(i).Value = ""
                If i = 5 Then fpLongInteger1(i).Value = ""
                If i = 6 Then fpLongInteger1(i).Value = ""
                If i = 7 Then fpLongInteger1(i).Value = ""
                If i = 17 Then fpText(i).Enabled = False
                
            Next i
            
            Combo1(1).ListIndex = -1
                       
           'INI ARI
     
           'Carga la LisBOX
           List1(1).Clear
           Dim Sql As String
           
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Sql = " sgpadm_Sel_OfertasCasino '" & fpText(0).text & "'"
           Set RS = vg_db.Execute(Sql)
           Dim contador As Long
           contador = 0
           
           '-------> Inicio LLenar grilla
       
           Do While Not RS.EOF
              
              List1(1).AddItem RS("Descripcion") & Space(150) & RS("codigo_oferta")
              If RS("selected") = 1 Then List1(1).Selected(contador) = True
          
              RS.MoveNext
              contador = contador + 1
          
          Loop
    
          RS.Close
          Set RS = Nothing
                
        'FIN ARI
    
        'INI Tipo Negocio
     
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
  
        'Carga la LisBOX
        TipoNegocio(0).Clear
     
        Dim Sql1 As String
        Sql1 = " sgpadm_Sel_CasinoTipoNegocio '" & fpText(0).text & "'"
        Set RS = vg_db.Execute(Sql1)
        Dim contador1 As Long
        contador1 = 0
          
        '-------> Inicio LLenar grilla
       
        Do While Not RS.EOF
         
           TipoNegocio(0).AddItem RS("NombreTipoNegocio") & Space(150) & RS("IdTipoNegocio")
        
           If RS("selected") = 1 Then TipoNegocio(0).Selected(contador1) = True
          
           RS.MoveNext
           contador1 = contador1 + 1
    
        Loop
    
        RS.Close
        Set RS = Nothing
        
        'FIN Tipo Negocio
    
        'INI Sello
     
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
  
        'Carga la LisBOX
        Sello(0).Clear
     
        Dim Sql2 As String
        Sql2 = " sgpadm_Sel_CasinoSello '" & fpText(0).text & "'"
        Set RS = vg_db.Execute(Sql2)
        Dim contador2 As Long
        contador2 = 0
          
        '-------> Inicio LLenar grilla
       
        Do While Not RS.EOF
         
           Sello(0).AddItem RS("NombreSellos") & Space(150) & RS("IdSellos")
        
           If RS("selected") = 1 Then Sello(0).Selected(contador2) = True
          
           RS.MoveNext
           contador2 = contador2 + 1
    
        Loop
    
        RS.Close
        Set RS = Nothing
        
        'FIN Sello
    
        Est = False
        itab = 0
        modo = "A"
        
        Case 3
            
            If vaSpread1.MaxRows < 1 Then Exit Sub
            modo = "M"
            itab = 1
            
            If SSTab1.Tab = 4 Then
               
               vg_nombre = "": vg_codigo = ""
               M_ADiaFe.Show 1, M_Casino
               If Trim(vg_codigo) = "" Then Exit Sub
               SSTab1.TabEnabled(0) = False
               SSTab1.TabEnabled(1) = False
               SSTab1.TabEnabled(2) = False
               SSTab1.TabEnabled(3) = False
               SSTab1.TabEnabled(8) = False
               SSTab1.TabEnabled(4) = True
               PonerDiasFeriados vg_codigo, vg_nombre
            
            ElseIf SSTab1.Tab = 6 Then
               
               SSTab1.TabEnabled(0) = False
               SSTab1.TabEnabled(1) = False
               SSTab1.TabEnabled(2) = False
               SSTab1.TabEnabled(3) = False
               SSTab1.TabEnabled(4) = False
               SSTab1.TabEnabled(5) = False
               SSTab1.TabEnabled(8) = False
               SSTab1.TabEnabled(6) = True
            
            ElseIf SSTab1.Tab = 5 Or SSTab1.TabVisible(5) = True Then
               
               SSTab1.TabEnabled(0) = False
               SSTab1.TabEnabled(1) = False
               SSTab1.TabEnabled(2) = False
               SSTab1.TabEnabled(3) = False
               SSTab1.TabEnabled(4) = False
               SSTab1.TabEnabled(8) = False
               SSTab1.TabEnabled(5) = True
               SSTab1.Tab = 5
            
            ElseIf SSTab1.Tab = 9 Or SSTab1.TabVisible(9) = True Then
               
               SSTab1.TabEnabled(0) = False
               SSTab1.TabEnabled(1) = False
               SSTab1.TabEnabled(2) = False
               SSTab1.TabEnabled(3) = False
               SSTab1.TabEnabled(4) = False
               SSTab1.TabEnabled(5) = False
               SSTab1.TabEnabled(8) = False
               SSTab1.TabEnabled(10) = False
               SSTab1.TabEnabled(11) = False
               SSTab1.TabVisible(12) = False

               SSTab1.TabEnabled(9) = True
               SSTab1.Tab = 9
               
            ElseIf SSTab1.Tab = 10 Or SSTab1.TabVisible(10) = True Then
               
               SSTab1.TabEnabled(0) = False
               SSTab1.TabEnabled(1) = False
               SSTab1.TabEnabled(2) = False
               SSTab1.TabEnabled(3) = False
               SSTab1.TabEnabled(4) = False
               SSTab1.TabEnabled(5) = False
               SSTab1.TabEnabled(8) = False
               SSTab1.TabEnabled(9) = False
               SSTab1.TabEnabled(11) = False
               SSTab1.TabEnabled(10) = True
               SSTab1.TabVisible(12) = False
               
               SSTab1.Tab = 10
            
            ElseIf SSTab1.Tab = 11 Or SSTab1.TabVisible(11) = True Then
               
               SSTab1.TabEnabled(0) = False
               SSTab1.TabEnabled(1) = False
               SSTab1.TabEnabled(2) = False
               SSTab1.TabEnabled(3) = False
               SSTab1.TabEnabled(4) = False
               SSTab1.TabEnabled(5) = False
               SSTab1.TabEnabled(8) = False
               SSTab1.TabEnabled(9) = False
               SSTab1.TabEnabled(10) = False
               SSTab1.TabEnabled(11) = True
               SSTab1.TabVisible(12) = False
               
               SSTab1.Tab = 11
            
            ElseIf SSTab1.Tab = 12 Or SSTab1.TabVisible(12) = True Then
               
               SSTab1.TabEnabled(0) = False
               SSTab1.TabEnabled(1) = False
               SSTab1.TabEnabled(2) = False
               SSTab1.TabEnabled(3) = False
               SSTab1.TabEnabled(4) = False
               SSTab1.TabEnabled(5) = False
               SSTab1.TabEnabled(8) = False
               SSTab1.TabEnabled(9) = False
               SSTab1.TabEnabled(10) = False
               SSTab1.TabEnabled(11) = False
               SSTab1.TabVisible(12) = True
               
               SSTab1.Tab = 12
            
            Else
               
               SSTab1.TabEnabled(0) = False
               SSTab1.TabEnabled(2) = True
               SSTab1.TabEnabled(3) = True
               SSTab1.TabEnabled(4) = True
               SSTab1.Tab = 1
               SSTab1.TabEnabled(1) = True
            
            End If
            Gl_Ac_Botones Me, 13, 0, modo
            itab = 0
        
        Case 5
            
            Borra_Datos
        
        Case 7
            
            modo = ""
            If SSTab1.Tab = 4 Then
               
               MoverDiasFeriados
            
            ElseIf SSTab1.Tab = 6 Then
               
               If vaSpread1.MaxRows < 1 Then Exit Sub
               MoverDatos6
            
            ElseIf SSTab1.Tab = 10 Then
            
               MoverInvCandelarizado
             
            ElseIf SSTab1.Tab = 11 Then
            
               MoverPvtaClienteCalendarizado
            
            Else
               
               SSTab1.Tab = 0
               MoverDatosGrilla
            
            End If
        
        Case 10
        
            Cancela_Datos
        
        Case 12
            
            Actualiza_Datos
        
        Case 15
            
            M_CpaDFe.Show 1, M_Casino
        
        Case 17
            
            If SSTab1.TabVisible(1) = True Then
               
               I_Casinos
            
            ElseIf SSTab1.TabVisible(5) = True Then
               
               I_ParametroDespachos
            
            End If
        
        Case 20
            
            ExportarExcel
        
        Case 22
            
            Me.Hide
            Unload Me
    
    End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Sub ExportarExcel()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim Sql             As String
Dim NomArchivoExcel As String
Dim Extension       As String

Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object

Label8.Visible = True

'-------> Validar cantidad registro se sobre pase hoja excel
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_ExportExcelInfClientes_V04 ")

If Not RS.EOF Then
  
   If RS.RecordCount > 1020000 Then
      
      RS.Close
      Set RS = Nothing
      
      MsgBox "El resultado sobrepasa maximo de fila en excel...", vbCritical
      Exit Sub
   
   End If
  
End If

Label8.Visible = False

'-------> Guardar nombre archivo excel
NomArchivoExcel = ""
CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
CD.Filter = "Todos los archivos *.xls,*.xlsx"
On Error Resume Next
CD.ShowSave
           
'-------> JPAZ Permite controlar Boton Cancelar
If Err.Number = 32755 Then
   
   MsgBox "Proceso cancelado"
   Exit Sub

End If
            
If CD.FileName = "" Then
   
   MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
   Exit Sub

Else
   
   Extension = ""
   Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
   
   If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
      MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
      Exit Sub
   End If
   
   NomArchivoExcel = CD.FileName

End If
          
fg_carga ""
  
'-------> Create an instance of Excel and add a workbook
Set xlApp = CreateObject("Excel.Application")
Set xlWb = xlApp.Workbooks.Add
Set xlWs = xlWb.Worksheets("Hoja1")
  
'-------> Display Excel and give user control of Excel's lifetime
xlApp.UserControl = True
    
'-------> Check version of Excel
Call encabezado(RS, xlWs)
          
xlWs.Cells(2, 1).CopyFromRecordset RS

'-------> Auto-fit the column widths and row heights
xlApp.Selection.CurrentRegion.Columns.AutoFit
xlApp.Selection.CurrentRegion.Rows.AutoFit
    
'xlApp.Columns("A:A").Select
'xlApp.Selection.Delete Shift:=xlToLeft
  
xlWb.Close True, NomArchivoExcel

Dim XL As New excel.Application 'Crea el objeto excel
XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
XL.Visible = True
XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
'-------> Close ADO objects
RS.Close
Set RS = Nothing
    
' -- Cerrar Excel
xlApp.Quit
'-------> Release Excel references
Set xlWs = Nothing
Set xlWb = Nothing
Set xlApp = Nothing


fg_descarga
MsgBox "Proceso Finalizado", vbInformation, Me.Caption
                
Exit Sub
Man_Error:
    Label8.Visible = False
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub MoverDatosGrilla()

On Error GoTo Man_Error

Dim RS     As New ADODB.Recordset
Dim IndCas As Long

fg_carga ""
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = Shape1(0).FillColor
vaSpread1.Row = -1
vaSpread1.Col = -1
vaSpread1.Lock = True
itab = 0
codigo = ""

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If lc_Aux = "MCasino" Then
    
    Set RS = vg_db.Execute("sgpadm_s_cliente_V02 1, '',''")

Else
    
    Set RS = vg_db.Execute("sgpadm_s_cliente_V02 48, '',''")

End If

If Not RS.EOF Then
   
   vaSpread1.MaxRows = RS.RecordCount
   IndCas = 1
   
   Do While Not RS.EOF
      
'      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = IndCas 'vaSpread1.MaxRows
              
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = -1
      vaSpread1.Row = IndCas
      vaSpread1.Col = -1
      vaSpread1.BackColor = Shape1(0).FillColor
      
      vaSpread1.Col = 1
      vaSpread1.text = RS!Cli_codigo

      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = 0
      vaSpread1.text = Trim(RS!Cli_nombre)
      
      vaSpread1.Col = 3
      vaSpread1.text = IIf(RS!cli_tipo = 0, "Operador", "Traspaso")
             
      vaSpread1.Col = 4
      If Not IsNull(RS!cli_openvio) Then
         
         If RS!cli_openvio = 1 Then
            
            vaSpread1.text = "Ftp"
         
         ElseIf RS!cli_openvio = 2 Then
            
            vaSpread1.text = "Mail"
         
         ElseIf RS!cli_openvio = 3 Then
            
            vaSpread1.text = "Archivo"
         
         End If
      
      Else
         
         vaSpread1.text = "No Especificado"
      
      End If
      
      vaSpread1.Col = 5
      vaSpread1.text = IIf(IsNull(RS!sub_nombre), "", Trim(RS!sub_nombre))
     
      vaSpread1.Col = 6
      vaSpread1.text = IIf(IsNull(RS!CLI_ACTIVO) Or Trim(RS!CLI_ACTIVO) = "0", "0", Trim(RS!CLI_ACTIVO))
     
      RS.MoveNext
      
      IndCas = IndCas + 1
   
   Loop
   
   SSTab1.TabEnabled(1) = True
   vaSpread1.Row = 1
   vaSpread1.Col = 1
   codigo = Trim(vaSpread1.text)

Else
   
   SSTab1.Tab = 0
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
   SSTab1.TabEnabled(3) = False
   SSTab1.TabEnabled(4) = False
   SSTab1.TabEnabled(5) = False
   SSTab1.TabEnabled(8) = False
   SSTab1.TabEnabled(9) = False
   SSTab1.TabEnabled(10) = False
   SSTab1.TabEnabled(11) = False
   SSTab1.TabVisible(12) = False

   modo = "NE"
   Gl_Ac_Botones Me, 13, 2, modo

End If
RS.Close
Set RS = Nothing
vaSpread1.Visible = True
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
FptNombre.text = ""
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatosGrillaOculta()

Dim RS As New ADODB.Recordset
Dim codaux As String

On Error GoTo Man_Error
    
    vaSpread3.MaxRows = 0
    vaSpread3.maxcols = 41
    
    itab = 0
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_s_ExportExcelInfClientes_V03")
    If Not RS.EOF Then
        
        Do While Not RS.EOF
            
            With vaSpread3
                
                .MaxRows = vaSpread3.MaxRows + 1
                .Row = vaSpread3.MaxRows
                
                If codaux <> RS!Cli_codigo Then
                    
                    .Col = 1: .text = IIf(IsNull(RS!Cli_codigo), "", RS!Cli_codigo)
                    .Col = 2: .text = IIf(IsNull(RS!Cli_nombre), "", RS!Cli_nombre)
                    .Col = 3: .text = IIf(IsNull(RS!cli_direccion), "", RS!cli_direccion)
                    .Col = 4: .text = IIf(IsNull(RS!cli_comuna), "", RS!cli_comuna)
                    .Col = 5: .text = IIf(IsNull(RS!cli_ciudad), "", RS!cli_ciudad)
                    .Col = 6: .CellType = CellTypeStaticText: .text = IIf(IsNull(RS!cli_fono1), "", RS!cli_fono1)
                    .Col = 7: .text = IIf(IsNull(RS!cli_fono2), "", RS!cli_fono2)
                    .Col = 8: .text = IIf(IsNull(RS!cli_fax), "", RS!cli_fax)
                    .Col = 9: .text = IIf(IsNull(RS!cli_percon), "", RS!cli_percon)
                    .Col = 10: .text = IIf(IsNull(RS!cli_email), "", RS!cli_email)
                    .Col = 11: .text = IIf(IsNull(RS!cli_giro), "", RS!cli_giro)
                    .Col = 12: .text = IIf(RS!cli_tipo = 0, "Operador", "Traspaso")
                    .Col = 13: .text = IIf(RS!cli_openvio = 1, "FTP", IIf(RS!cli_openvio = 2, "Mail", IIf(RS!cli_openvio = 3, "Archivo", "No Especificado")))
                    .Col = 14: .text = IIf(IsNull(RS!cli_subseg), "", RS!cli_subseg)
                    .Col = 15: .text = IIf(IsNull(RS!sub_nombre), "", RS!sub_nombre)
                    .Col = 16: .text = IIf(IsNull(RS!cli_codzon), "", RS!cli_codzon)
                    .Col = 17: .text = IIf(IsNull(RS!Zon_nombre), "", RS!Zon_nombre)
                    .Col = 18: .text = IIf(IsNull(RS!cli_nomcontable), "", RS!cli_nomcontable)
                    .Col = 19: .text = IIf(IsNull(RS!cli_emailcontable), "", RS!cli_emailcontable)
                    .Col = 20: .text = IIf(IsNull(RS!cli_gruvul), "", RS!cli_gruvul)
                    .Col = 21: .text = IIf(IsNull(RS!cli_modpac), "", RS!cli_modpac)
                    .Col = 22: .text = IIf(IsNull(RS!cli_codtis), "", RS!cli_codtis)
                    .Col = 23: .text = IIf(IsNull(RS!tis_nombre), "", RS!tis_nombre)
                    .Col = 24: .text = IIf(IsNull(RS!cli_codseg), "", RS!cli_codseg)
                    .Col = 25: .text = IIf(IsNull(RS!seg_nombre), "", RS!seg_nombre)
                    .Col = 26: .text = IIf(IsNull(RS!cli_socsap), "", RS!cli_socsap)
                    .Col = 27: .text = IIf(IsNull(RS!CLI_ACTIVO), "", RS!CLI_ACTIVO)
                    .Col = 28: .text = IIf(IsNull(RS!cli_sobrec), "", RS!cli_sobrec)
                    .Col = 29: .text = IIf(RS!int_cfc = 1, "X", "")
                    .Col = 30: .text = IIf(RS!int_inventario = 1, "X", "")
                    .Col = 31: .text = IIf(RS!int_guiaventa = 1, "X", "")
                    .Col = 32: .text = IIf(RS!int_cierrediario = 1, "X", "")
'                    .Col = 38: .text = IIf(RS!cli_tipominuta = 1, "Bloque", IIf(RS!cli_tipominuta = 2, "Mio", IIf(RS!cli_tipominuta = 3, "Simap", "Otros")))
                    .Col = 38: .text = RS!cli_tipominuta
                    .Col = 39: .text = IIf(RS!cli_tipoformatocompras = 1, "Grande", IIf(RS!cli_tipoformatocompras = 2, "Chico", "No Especificado"))
                    .Col = 40: .text = IIf(RS!cli_tipoceco = "0", "Real", IIf(RS!cli_tipoceco = "1", "Comercial", "No Especificado"))
                    .Col = 41: .text = IIf(RS!cli_AMD = "2", "Minuta Normal", IIf(RS!cli_AMD = "1", "Minuta AMD", "No Especificado"))
                    
                    codaux = RS!Cli_codigo
                
                End If
                 
                 .Col = 33: .text = IIf(IsNull(RS!Reg_Codigo), "", RS!Reg_Codigo)
                 .Col = 34: .text = IIf(IsNull(RS!reg_nombre), "", RS!reg_nombre)
                 .Col = 35: .text = IIf(RS!reg_activo = 1, "S", "N")
                 .Col = 36: .text = IIf(IsNull(RS!Ser_codigo), "", RS!Ser_codigo)
                 .Col = 37: .text = IIf(IsNull(RS!ser_nombre), "", RS!ser_nombre)
                
                RS.MoveNext
            
            End With
       
       Loop
       
       vaSpread3.Row = 1: vaSpread3.Col = 1
    
    End If
    
    RS.Close: Set RS = Nothing

Exit Sub
Man_Error:
    MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatos()
        
On Error GoTo Man_Error

    Check3.Visible = False
    ChkMInRet.Visible = False
    ChkBlockMinContrato.Visible = False
    ChkBlockMinTeo.Visible = False
    ChkBlockMinReal.Visible = False
    ChkBlockTraFinSemana.Visible = True 'False

    fg_carga ""
    
    Dim RS As New ADODB.Recordset
    Est = True
    Combo1(0).Enabled = False
    FptNombre.Enabled = False
    Option3(0).Value = 0
    Option3(1).Value = 0
    Option3(3).Value = True
    Option3(2).Value = False
    Option4(0).Value = False
    Option4(1).Value = False
    
    For i = 0 To 14
        
        If i < 14 Then fpText(i).text = "": fpText(i).Enabled = True
        If i = 14 Then fpLongInteger1(0).Value = ""
    
    Next i
    
    fpText(17).text = ""
    
    fpLongInteger1(0).Value = ""
    fpLongInteger1(1).Value = ""
    fpLongInteger1(2).Value = ""
    fpLongInteger1(3).Value = ""
    fpLongInteger1(5).Value = ""
    fpLongInteger1(6).Value = ""
    fpLongInteger1(7).Value = ""
    fpayuda(0).Caption = ""
    fpayuda(1).Caption = ""
    fpayuda(3).Caption = ""
    fpayuda(7).Caption = ""
    fpayuda(9).Caption = ""
    fpayuda(10).Caption = ""
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    codigo = vaSpread1.text
    fpText(0).Enabled = False
    
    Label9.Caption = "No existe Inf. T.Gramaje"
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS = vg_db.Execute("sgpadm_Sel_DetalleCliente_V03 '" & codigo & "'")
    If Not RS.EOF Then
        
        fpText(0).text = RS!Cli_codigo
        fpText(1).text = IIf(IsNull(RS!Cli_nombre), "", Trim(RS!Cli_nombre))
        fpText(2).text = IIf(IsNull(RS!cli_direccion), "", Trim(RS!cli_direccion))
        fpText(3).text = IIf(IsNull(RS!cli_comuna), "", Trim(RS!cli_comuna))
        fpText(4).text = IIf(IsNull(RS!cli_ciudad), "", Trim(RS!cli_ciudad))
        fpText(5).text = IIf(IsNull(RS!cli_fono1), "", Trim(RS!cli_fono1))
        fpText(6).text = IIf(IsNull(RS!cli_fono2), "", Trim(RS!cli_fono2))
        fpText(7).text = IIf(IsNull(RS!cli_fax), "", Trim(RS!cli_fax))
        fpText(8).text = IIf(IsNull(RS!cli_percon), "", Trim(RS!cli_percon))
        fpText(9).text = IIf(IsNull(RS!cli_giro), "", Trim(RS!cli_giro))
        fpText(10).text = IIf(IsNull(RS!cli_email), "", Trim(RS!cli_email))
        Check1(0).Value = IIf(RS!cli_tipo = 2, 1, 0)
        Combo1(1).ListIndex = IIf(IsNull(RS!cli_openvio), -1, fg_buscacbo(Combo1, 1, 1, IIf(IsNull(RS!cli_openvio), -1, RS!cli_openvio)))
        fpLongInteger1(0).Value = IIf(IsNull(RS!cli_subseg) Or RS!cli_subseg = 0, "", RS!cli_subseg)
        fpayuda(1).Caption = IIf(IsNull(RS!cli_subseg) Or IsNull(RS!sub_nombre), "", Trim(RS!sub_nombre))
        fpText(11).text = IIf(IsNull(RS!cli_nomcontable), "", Trim(RS!cli_nomcontable))
        fpText(12).text = IIf(IsNull(RS!cli_emailcontable), "", Trim(RS!cli_emailcontable))
        Check1(1).Value = IIf(IsNull(RS!cli_gruvul) Or RS!cli_gruvul = "N", 0, 1)
        fpLongInteger1(1).Value = IIf(IsNull(RS!cli_codzon) Or RS!cli_codzon = 0, "", RS!cli_codzon)
        fpayuda(0).Caption = IIf(IsNull(RS!cli_codzon) Or IsNull(RS!Zon_nombre), "", Trim(RS!Zon_nombre))
        Check1(2).Value = IIf(IsNull(RS!cli_modpac) Or RS!cli_modpac = "N", 0, 1)
        fpLongInteger1(2).Value = IIf(IsNull(RS!cli_codtis) Or RS!cli_codtis = 0, "", RS!cli_codtis)
        fpayuda(3).Caption = IIf(IsNull(RS!cli_codtis) Or IsNull(RS!tis_nombre), "", Trim(RS!tis_nombre))
        fpLongInteger1(3).Value = IIf(IsNull(RS!cli_codseg) Or RS!cli_codseg = 0, "", RS!cli_codseg)
        fpayuda(7).Caption = IIf(IsNull(RS!cli_codseg) Or IsNull(RS!seg_nombre), "", Trim(RS!seg_nombre))
        fpText(13).text = IIf(IsNull(RS!cli_socsap), "", Trim(RS!cli_socsap))
        Check1(0).Value = IIf(IsNull(RS!CLI_ACTIVO) Or Trim(RS!CLI_ACTIVO) = "" Or RS!CLI_ACTIVO = "0", 0, 1)
        Option1(IIf(IsNull(RS!cli_sobrec) Or RS!cli_sobrec = "0", 0, IIf(RS!cli_sobrec = "1", 1, 2))).Value = True
        fpLongInteger1(5).Value = IIf(IsNull(RS!cli_codmun) Or RS!cli_codmun = 0, "", RS!cli_codmun)
        fpayuda(9).Caption = IIf(IsNull(RS!cli_codmun) Or IsNull(RS!mun_nombre), "", Trim(RS!mun_nombre))
        fpLongInteger1(6).Value = IIf(IsNull(RS!cli_ccisac) Or RS!cli_ccisac = 0, "", Trim(RS!cli_ccisac))
        fpText(14).text = IIf(IsNull(RS!cli_cecsac), "", Trim(RS!cli_cecsac))
        Option3(0).Value = IIf(IsNull(RS!cli_opgped) Or RS!cli_opgped = "0", 1, 0)
        Option3(1).Value = IIf(IsNull(RS!cli_opgped) Or RS!cli_opgped = "0", 0, 1)
        Check1(3).Value = IIf(IsNull(RS!cli_hipali) Or RS!cli_hipali = "N", 0, 1)
        fpLongInteger1(7).Value = IIf(IsNull(RS!cli_codreg) Or RS!cli_codreg = 0, "", RS!cli_codreg)
        fpayuda(10).Caption = IIf(IsNull(RS!cli_codreg) Or IsNull(RS!reg_nombre), "", Trim(RS!reg_nombre))
        Option3(3).Value = IIf(IsNull(RS!cli_tipope) Or RS!cli_tipope = "0", True, False)
        Option3(2).Value = IIf(IsNull(RS!cli_tipope) Or RS!cli_tipope = "1", True, False)
        Let ChkMInRet.Value = IIf(IsNull(RS!cli_minsre) Or RS!cli_minsre = "0", 0, 1)
'        Let Check3.Value = IIf(IsNull(RS!cli_portalelectronico) Or RS!cli_portalelectronico = "0", 0, 1)
        Let Check3.Value = IIf(IsNull(RS!cli_portalelectronico) Or RS!cli_portalelectronico = "0", 1, 1)
        Let ChkBlockMinReal.Value = IIf(IsNull(RS!cli_blockminreal) Or RS!cli_blockminreal = "0", 0, 1) 'MVA - MVI - 2013-01-04
        Let ChkBlockMinTeo.Value = IIf(IsNull(RS!cli_blockminteo) Or RS!cli_blockminteo = "0", 0, 1) 'MVA - MVI - 2013-01-04
        Let ChkBlockMinContrato.Value = IIf(IsNull(RS!cli_blockmincontrato) Or RS!cli_blockmincontrato = "0", 0, 1) 'MVA - MVI - 2013-01-04
        Let ChkBlockTraFinSemana.Value = IIf(IsNull(RS!cli_blockmintrabajafinsemana) Or RS!cli_blockmintrabajafinsemana = "0", 0, 1) 'JPA - MVI - 2013-03-08
        Let fpText(15).text = IIf(IsNull(RS!cli_emailenviopedido), "", RS!cli_emailenviopedido) 'JPAZ - MVI - 2013-03-18

        'INI ARI
        'Mover campos nuevos al combo
        Combo1(2).ListIndex = IIf(IsNull(RS!cli_tipominuta), -1, fg_buscacbo(Combo1, 2, 1, IIf(IsNull(RS!cli_tipominuta), -1, RS!cli_tipominuta)))
        Combo1(3).ListIndex = IIf(IsNull(RS!cli_tipoformatocompras), -1, fg_buscacbo(Combo1, 3, 1, IIf(IsNull(RS!cli_tipoformatocompras), -1, RS!cli_tipoformatocompras)))
        'FIN ARI
                
        'Ini Optimun
        fpText(16).text = IIf(IsNull(RS!Cecos_AX) Or Trim(RS!Cecos_AX) = "", "", RS!Cecos_AX)
        'Fin Optimun
    
        'Ini Organización Compras
        fpText(17).text = IIf(IsNull(RS!ID_Orgcompra) Or Trim(RS!ID_Orgcompra) = "", "", RS!ID_Orgcompra)
        'Fin Organización Compras
    
        If Not IsNull(RS!cli_tipoceco) Then
           
           Option4(IIf(IsNull(RS!cli_tipoceco) Or Trim(RS!cli_tipoceco) = "" Or RS!cli_tipoceco = "0", 0, IIf(RS!cli_tipoceco = "1", 1, 2))).Value = True
        
        End If
    
        'acesso tipo minuta
        Option5(0).Value = IIf(IsNull(RS!cli_AMD) Or RS!cli_AMD = "2", 1, 0)
        Option5(1).Value = IIf(IsNull(RS!cli_AMD) Or RS!cli_AMD = "1", 1, 0)

        'Integra AMD
        Check4.Value = IIf(RS!IdIntegraAMD = 1, 1, 0)
        
        'Integra SPRS
        Check5.Value = IIf(RS!cli_IntegraSPRS = 1, 1, 0)
        
    End If
    RS.Close
    Set RS = Nothing
    
     If RS.State = 1 Then RS.Close
     RS.CursorLocation = adUseClient
     vg_db.CursorLocation = adUseClient
    
     Set RS = vg_db.Execute("sgpadm_Sel_ExisteDatosTablaGramaje '" & codigo & "'")
    
     If Not RS.EOF Then
     
        If RS(0) = "1" Then
           
           Label9.Caption = "Existe Inf. T.Gramaje"
        
        Else
           
           Label9.Caption = "No existe Inf. T.Gramaje"
        
        End If
     
     End If
     RS.Close
     Set RS = Nothing
   
    
    'INI ARI
     
     If RS.State = 1 Then RS.Close
     RS.CursorLocation = adUseClient
     vg_db.CursorLocation = adUseClient
     
     'Carga la LisBOX
     List1(1).Clear
     Dim Sql As String
     Sql = " sgpadm_Sel_OfertasCasino '" & fpText(0).text & "'"
     Set RS = vg_db.Execute(Sql)
     Dim contador As Long
     contador = 0
          
     '-------> Inicio LLenar grilla
       
     Do While Not RS.EOF
         
        List1(1).AddItem RS("Descripcion") & Space(150) & RS("codigo_oferta")
       ' List1.ItemData(List1.ListIndex) = RS("codigo_oferta")
        If RS("selected") = 1 Then List1(1).Selected(contador) = True
          
       RS.MoveNext
       contador = contador + 1
    
     Loop
    
     RS.Close
     Set RS = Nothing
         
     'FIN ARI
    
    'INI Tipo Negocio
     
     If RS.State = 1 Then RS.Close
     RS.CursorLocation = adUseClient
     vg_db.CursorLocation = adUseClient
  
     'Carga la LisBOX
     TipoNegocio(0).Clear
     
     Dim Sql1 As String
     Sql1 = " sgpadm_Sel_CasinoTipoNegocio '" & fpText(0).text & "'"
     Set RS = vg_db.Execute(Sql1)
     Dim contador1 As Long
     contador1 = 0
          
     '-------> Inicio LLenar grilla
       
     Do While Not RS.EOF
         
        TipoNegocio(0).AddItem RS("NombreTipoNegocio") & Space(150) & RS("IdTipoNegocio")
        
        If RS("selected") = 1 Then TipoNegocio(0).Selected(contador1) = True
          
        RS.MoveNext
        contador1 = contador1 + 1
    
     Loop
    
     RS.Close
     Set RS = Nothing
        
     'FIN Tipo Negocio
        
     'INI Sello
     
     If RS.State = 1 Then RS.Close
     RS.CursorLocation = adUseClient
     vg_db.CursorLocation = adUseClient
  
     'Carga la LisBOX
     Sello(0).Clear
     
     Dim Sql2 As String
     Sql2 = " sgpadm_Sel_CasinoSello '" & fpText(0).text & "'"
     Set RS = vg_db.Execute(Sql2)
     Dim contador2 As Long
     contador2 = 0
          
     '-------> Inicio LLenar grilla
       
     Do While Not RS.EOF
         
        Sello(0).AddItem RS("NombreSellos") & Space(150) & RS("IdSellos")
        
        If RS("selected") = 1 Then Sello(0).Selected(contador2) = True
          
        RS.MoveNext
        contador2 = contador2 + 1
    
     Loop
    
     RS.Close
     Set RS = Nothing
        
     'FIN Sello
        
    Frame4.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1" Or Mid(ValidarUsuario(Me), 2, 1) = "1", True, False)
    Est = False
    Call fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatos2()

On Error GoTo Man_Error

Dim j      As Long
Dim z      As Long
Dim padre  As String
Dim RS     As New ADODB.Recordset
Dim auxreg As Long, X As Long
    
    fg_carga ""
    auxreg = 0
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    codigo = Trim(vaSpread1.text)
    
    vaSpread1.Col = 2
    Label2.Caption = Trim(vaSpread1.text)
    
    nivel = 65
    i = 1
    j = 1
    X = 1
    z = 1
    
    Me.Refresh
    TvwDir.Nodes.Clear
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS = vg_db.Execute("sgpadm_s_buscarregcasinonivel '" & codigo & "'")
    If Not RS.EOF Then
       
       Do While Not RS.EOF
          
          Set rootNode = TvwDir.Nodes.Add(, , "N" & fg_pone_espacio(RS!Reg_Codigo, 10), RS!Reg_Codigo & " - " & Trim(RS!reg_nombre))
          TvwDir.Nodes.item(i).Checked = IIf(RS!crs_codreg = 1, True, False)
          j = i
          i = i + 1
          TvwDir.Nodes.Add rootNode.Index, tvwChild, , "*"
          i = i + 1
          RS.MoveNext
          
       Loop
    
    End If
    
    RS.Close
    Set RS = Nothing
    
    Call fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatos3()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Est = True

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_cliente_V02 13, '" & codigo & "', ''")
Do While Not RS.EOF
       
   Frame5.Caption = Trim(RS!Cli_codigo) & " - " & Trim(RS!Cli_nombre)
   RS.MoveNext
    
Loop
RS.Close
Set RS = Nothing
    
vaSpread2.Visible = False
vaSpread2.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_casinointerfaz 1, '" & codigo & "'")
Do While Not RS.EOF
       
   vaSpread2.MaxRows = vaSpread2.MaxRows + 1
   vaSpread2.Row = vaSpread2.MaxRows
   vaSpread2.Col = 1: vaSpread2.text = RS!tii_activo
   vaSpread2.Col = 2: vaSpread2.text = RS!tii_codigo
   vaSpread2.Col = 3: vaSpread2.text = Trim(RS!tii_nombre)
   RS.MoveNext
    
Loop
RS.Close
Set RS = Nothing
vaSpread2.Visible = True
Est = False

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatos4()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Est = True
Me.Refresh

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_cliente_V02 13, '" & codigo & "', ''")
Do While Not RS.EOF
       
   Frame6.Caption = Trim(RS!Cli_codigo) & " - " & Trim(RS!Cli_nombre)
   RS.MoveNext
    
Loop
RS.Close
Set RS = Nothing
MoverDiasFeriados
Est = False

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatos5()

On Error GoTo Man_Error

Frame4.Visible = False
Frame8.Visible = False
vaSpread4.Visible = False

Dim RS As New ADODB.Recordset
Dim i As Long, codaux As Long, j As Long
Dim X, ParentId As Long

Me.Refresh

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_cliente_V02 13, '" & codigo & "', ''")
Do While Not RS.EOF
       
   Frame8.Caption = Trim(RS!Cli_codigo) & " - " & Trim(RS!Cli_nombre)
   RS.MoveNext
    
Loop
RS.Close
Set RS = Nothing
   
vaSpread4.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_sel_casino_param_despacho ")
If Not RS.EOF Then
        
   With vaSpread4
            
        Do While Not RS.EOF
                
           .MaxRows = .MaxRows + 1
           .Row = .MaxRows
           .Col = 1
           .text = RS!pas_codigo
           
           .Col = 1
           .Font.Bold = True
           
           .Col = 2
           .text = Trim(RS!pas_nombre)
           
           .Col = 2
           .Font.Bold = True
           
           .Col = 3
           .CellType = CellTypeStaticText
           
           .Col = 4
           .CellType = CellTypeStaticText
           
           .Col = 5
           .CellType = CellTypeStaticText
           
           .Col = 6
           .CellType = CellTypeStaticText
           
           .Col = 7
           .CellType = CellTypeStaticText
           
           .Col = 8
           .CellType = CellTypeStaticText
           
           .Col = 9
           .CellType = CellTypeStaticText
           
           .Col = 10
           .CellType = CellTypeStaticText
           
           .Col = 11
           .CellType = CellTypeStaticText
           
           .Col = 12
           .CellType = CellTypeStaticText
           
           .Col = 13
           .CellType = CellTypeStaticText
           
           .Col = 14
           .CellType = CellTypeStaticText
           
           ParentId = RS(0)
                
           'sacar los hijos
                 
           X = Hijos(0, codigo, ParentId, vaSpread4)
                
           RS.MoveNext
            
        Loop
        
   End With
    
End If
RS.Close
Set RS = Nothing
vaSpread4.SetActiveCell 4, 1
Frame8.Visible = True
    
vaSpread4.Visible = True

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatos10()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Est = True
Me.Refresh

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_cliente_V02 13, '" & codigo & "', ''")
Do While Not RS.EOF
       
   Frame20.Caption = Trim(RS!Cli_codigo) & " - " & Trim(RS!Cli_nombre)
   RS.MoveNext
    
Loop
RS.Close
Set RS = Nothing

MoverInvCandelarizado
Est = False

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatos11()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Est = True
Me.Refresh

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_cliente_V02 13, '" & codigo & "', ''")
Do While Not RS.EOF
       
   Frame21.Caption = Trim(RS!Cli_codigo) & " - " & Trim(RS!Cli_nombre)
   RS.MoveNext
    
Loop
RS.Close
Set RS = Nothing

MoverPvtaClienteCalendarizado
Est = False

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatos12()

On Error GoTo Man_Error

'tvwFirst : Añade el nodo al principio
'tvwLast : Añade el nodo al final
'tvwNext : Lo añade al siguiente nodo indicado
'tvwPrevious : Lo añade al lugar anterior al nodo indicado
'tvwChild : Nuevo nodo Hijo o secundario del nodo indicado

Dim RS   As New ADODB.Recordset
Dim RS1  As New ADODB.Recordset
Dim RS2  As New ADODB.Recordset
Dim Iini As Long

fg_carga ""

Est = True
Me.Refresh

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_cliente_V02 13, '" & codigo & "', ''")
Do While Not RS.EOF
       
   Frame26.Caption = Trim(RS!Cli_codigo) & " - " & Trim(RS!Cli_nombre)
   RS.MoveNext
    
Loop
RS.Close
Set RS = Nothing

' *** Llenar Categoria dietetica ***'
TvwDietetica(0).Nodes.Clear

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS1 = vg_db.Execute("sgpadm_Sel_CasinoParamCategoriaDieteticaPrimerNivel_V01 '" & codigo & "'")
If Not RS1.EOF Then
   
   Do While Not RS1.EOF
           
      Set Nodx = TvwDietetica(0).Nodes.Add(, , "R" & RS1!car_codigo, RS1!car_codigo & " - " & Trim(RS1!car_nombre))
      
      TvwDietetica(0).Nodes.item(Nodx.Index).Checked = IIf(RS1!MARCA = "1", True, False)

      ' agregar un nodo hijo postizo, si fuera necesario

      If Nodx.Children = 0 Then
         
         If RS2.State = 1 Then RS2.Close
         RS2.CursorLocation = adUseClient
         vg_db.CursorLocation = adUseClient
         
         Set RS2 = vg_db.Execute("sgpadm_Sel_CasinoParamCategoriaDieteticaOtrosNiveles_V01 " & RS1!car_codigo & ", '1', '" & codigo & "'")
         
         If Not RS2.EOF Then
            
            ' la propiedad Texto de los nodos postizos es "***"
            TvwDietetica(0).Nodes.item(TvwDietetica(0).Nodes.count).Selected = True
            TvwDietetica(0).Nodes.Add Nodx.Index, tvwChild, , "*"
            
            Set nd1 = TvwDietetica(0).SelectedItem
            TvwDietetica_Expand_2 0, nd1, codigo
         
         End If
         RS2.Close
         Set RS2 = Nothing
      
      End If
      
      RS1.MoveNext
   
   Loop

End If
RS1.Close
Set RS1 = Nothing

'For Iini = 1 To TvwDietetica(0).Nodes.count
'
'    TvwDietetica(0).Nodes.item(Iini).Checked = True
'
'Next

fg_descarga
Est = False

Exit Sub
Est = False

Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Public Function Hijos(id As Long, Ceco As String, ParentId As Long, vaSpread4 As Object) As Long
   
On Error GoTo Man_Error

   If ParentId = 0 Then
      Hijos = id
      Exit Function
   End If
   
   Dim RS As New ADODB.Recordset
   Dim j As Integer
   Dim r  As Long
 
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_sel_casino_param_despacho_detalle_V02 '" & Ceco & "','" & ParentId & "'")
   If Not RS.EOF Then
       
       Do While Not RS.EOF
               
               vaSpread4.MaxRows = vaSpread4.MaxRows + 1
               vaSpread4.Row = vaSpread4.MaxRows
               vaSpread4.Col = 1: vaSpread4.text = RS!tip_codigo
               vaSpread4.Col = 2: vaSpread4.text = vaSpread4.text & vbCrLf & Space(CInt(RS!Level) * 2) & "   " & Trim(RS!tip_nombre)
               vaSpread4.Col = 3: vaSpread4.text = IIf(IsNull(RS!pad_diaseg) Or RS!pad_diaseg = 0, "", RS!pad_diaseg)
               vaSpread4.Col = 4: vaSpread4.TypeComboBoxList = "MENSUAL" & Chr$(9) & "QUINCENAL" & Chr$(9) & "SEMANAL" & Chr$(9) & "DIARIO 10" & Chr$(9) & "DIARIO"
               vaSpread4.Col = 5: vaSpread4.TypeComboBoxList = "M" & Chr$(9) & "Q" & Chr$(9) & "S" & Chr$(9) & "D" & Chr$(9) & "E"
               
               For i = 0 To vaSpread4.TypeComboBoxCount
                    
                    vaSpread4.TypeComboBoxCurSel = i
                    If vaSpread4.text = Mid(RS!pad_tipo, 1, 1) Then j = i: Exit For
                    j = -1
                
                Next i
                vaSpread4.Col = 4: vaSpread4.TypeComboBoxCurSel = j
                vaSpread4.Col = 6
                
                If j = 1 Then
                    
                    vaSpread4.Col = 6: vaSpread4.CellType = CellTypeComboBox: vaSpread4.TypeComboBoxList = "QUINCENAL 1-15" & Chr$(9) & "QUINCENAL 2-16" & Chr$(9) & "QUINCENAL 3-17" & Chr$(9) & "QUINCENAL 4-18"
                    vaSpread4.Col = 7: vaSpread4.CellType = CellTypeComboBox: vaSpread4.TypeComboBoxList = "Q1" & Chr$(9) & "Q2" & Chr$(9) & "Q3" & Chr$(9) & "Q4"
                    
                    For i = 0 To vaSpread4.TypeComboBoxCount
                        
                        vaSpread4.TypeComboBoxCurSel = i
                        If vaSpread4.text = RS!pad_tipo Then j = i: Exit For
                        j = -1
                    
                    Next i
                    
                    vaSpread4.Col = 6: vaSpread4.TypeComboBoxCurSel = j
                
                Else
                    
                    vaSpread4.CellType = CellTypeStaticText
                    vaSpread4.Col = 7: vaSpread4.CellType = CellTypeStaticText
                
                End If
              
                estdes = True
                If RS!pad_tipo <> "E" And RS!pad_tipo <> "S" Then
                    
                    vaSpread4.Col = 8: vaSpread4.CellType = CellTypeStaticText
                    vaSpread4.Col = 9: vaSpread4.CellType = CellTypeStaticText
                    vaSpread4.Col = 10: vaSpread4.CellType = CellTypeStaticText
                    vaSpread4.Col = 11: vaSpread4.CellType = CellTypeStaticText
                    vaSpread4.Col = 12: vaSpread4.CellType = CellTypeStaticText
                    vaSpread4.Col = 13: vaSpread4.CellType = CellTypeStaticText
                    vaSpread4.Col = 14: vaSpread4.CellType = CellTypeStaticText
                
                Else
                    
                  If IsNull(RS!pad_tipo) Then
                  
                    vaSpread4.Col = 8: vaSpread4.CellType = CellTypeStaticText
                    vaSpread4.Col = 9: vaSpread4.CellType = CellTypeStaticText
                    vaSpread4.Col = 10: vaSpread4.CellType = CellTypeStaticText
                    vaSpread4.Col = 11: vaSpread4.CellType = CellTypeStaticText
                    vaSpread4.Col = 12: vaSpread4.CellType = CellTypeStaticText
                    vaSpread4.Col = 13: vaSpread4.CellType = CellTypeStaticText
                    vaSpread4.Col = 14: vaSpread4.CellType = CellTypeStaticText
                  
                  Else
                    
                    vaSpread4.Col = 8: vaSpread4.CellType = CellTypeCheckBox: vaSpread4.TypeHAlign = TypeHAlignCenter: vaSpread4.text = IIf(IsNull(RS!pad_diario), 0, Mid(RS!pad_diario, 1, 1))
                    vaSpread4.Col = 9: vaSpread4.CellType = CellTypeCheckBox: vaSpread4.TypeHAlign = TypeHAlignCenter: vaSpread4.text = IIf(IsNull(RS!pad_diario), 0, Mid(RS!pad_diario, 2, 1))
                    vaSpread4.Col = 10: vaSpread4.CellType = CellTypeCheckBox: vaSpread4.TypeHAlign = TypeHAlignCenter: vaSpread4.text = IIf(IsNull(RS!pad_diario), 0, Mid(RS!pad_diario, 3, 1))
                    vaSpread4.Col = 11: vaSpread4.CellType = CellTypeCheckBox: vaSpread4.TypeHAlign = TypeHAlignCenter: vaSpread4.text = IIf(IsNull(RS!pad_diario), 0, Mid(RS!pad_diario, 4, 1))
                    vaSpread4.Col = 12: vaSpread4.CellType = CellTypeCheckBox: vaSpread4.TypeHAlign = TypeHAlignCenter: vaSpread4.text = IIf(IsNull(RS!pad_diario), 0, Mid(RS!pad_diario, 5, 1))
                    vaSpread4.Col = 13: vaSpread4.CellType = CellTypeCheckBox: vaSpread4.TypeHAlign = TypeHAlignCenter: vaSpread4.text = IIf(IsNull(RS!pad_diario), 0, Mid(RS!pad_diario, 6, 1))
                    vaSpread4.Col = 14: vaSpread4.CellType = CellTypeCheckBox: vaSpread4.TypeHAlign = TypeHAlignCenter: vaSpread4.text = IIf(IsNull(RS!pad_diario), 0, Mid(RS!pad_diario, 7, 1))
                  
                  End If
                    
                End If
                estdes = False
               ParentId = RS(0)
            
            If RS.Fields("tip_previo") = 0 Then
                
                r = RS.Fields("tip_codigo")
            
            Else
                
                r = Hijos(1, Ceco, RS.Fields("tip_codigo"), vaSpread4)  ' Recursivo.
            
            End If
            
            RS.MoveNext
        
        Loop
    
    End If
   
   Hijos = r
   RS.Close
   Set RS = Nothing
   
Exit Function
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Function

Private Sub MoverDatos6()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i As Long, codaux As Long, j As Long

    Me.Refresh
    estact = True
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_s_cliente_V02 13, '" & codigo & "', ''")
    Do While Not RS.EOF
       
       Frame9.Caption = Trim(RS!Cli_codigo) & " - " & Trim(RS!Cli_nombre)
       RS.MoveNext
    
    Loop
    RS.Close: Set RS = Nothing
    
    vaSpread5.Visible = False
    vaSpread5.MaxRows = 0
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_s_casinotipoactividades '" & codigo & "'")
    If Not RS.EOF Then
       
       Do While Not RS.EOF
          
          vaSpread5.MaxRows = vaSpread5.MaxRows + 1
          vaSpread5.Row = vaSpread5.MaxRows
          vaSpread5.Col = 1: vaSpread5.text = RS!tia_codigo
          vaSpread5.Col = 2: vaSpread5.text = IIf(IsNull(RS!tia_codigo), "", Trim(RS!tia_nombre))
          vaSpread5.Col = 3: vaSpread5.text = IIf(IsNull(RS!tia_opcion) Or Trim(RS!tia_opcion) = "", "0", "1")
          RS.MoveNext
       
       Loop
    
    End If
    
    RS.Close
    Set RS = Nothing
    vaSpread5.SetActiveCell 1, 1
    vaSpread5.Visible = True
    estact = False

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatos7()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i As Long, codaux As Long, j As Long

    Me.Refresh
    Est = True
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_s_cliente_V02 13, '" & codigo & "', ''")
    Do While Not RS.EOF
       
       Frame10.Caption = Trim(RS!Cli_codigo) & " - " & Trim(RS!Cli_nombre)
       RS.MoveNext
    
    Loop
    RS.Close: Set RS = Nothing
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_s_casinoparametrostock '" & codigo & "'")
    If Not RS.EOF Then
       
       Option2(0).Value = IIf(RS!cps_invsto = "S", True, False)
       Option2(1).Value = IIf(RS!cps_reqmen = "S", True, False)
       fpLongInteger1(4).Value = IIf(RS!cps_porinv > 0, RS!cps_porinv, "")
       Combo2(0).ListIndex = IIf(IsNull(RS!cps_liscri), -1, fg_buscacbo(Combo2, 0, 1, IIf(IsNull(RS!cps_liscri), -1, RS!cps_liscri)))
       Check2(2).Value = IIf(RS!cps_diario = "S", 1, 0)
       Check2(3).Value = IIf(RS!cps_ajuimp = "S", 1, 0)
    
    Else
       
       Option2(0).Value = False
       Option2(1).Value = False
       fpLongInteger1(4).Value = ""
       Check2(2).Value = 0
       Check2(3).Value = 0
       Combo2(0).ListIndex = -1
    
    End If
    RS.Close: Set RS = Nothing
    Est = False

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatos8()
    
On Error GoTo Man_Error

    Dim RS As New ADODB.Recordset
    Dim i As Long
    Dim vTipoVale() As Variant
    Dim IndTipVal As Long, j As Long, z As Long, CodTipVal As Long
    Dim lisnom As String
    Dim liscod As String

    Me.Refresh
    Est = True
    IndTipVal = 1
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_ParametroCodigoBarraxCodigo '%%'")
    If Not RS.EOF Then
       
       ReDim vTipoVale(RS!nReg + 1, 2)
       IndTipVal = RS!nReg + 1
    
    Else
       
       ReDim vTipoVale(1, 2)
    
    End If
    RS.Close
    Set RS = Nothing
    
    IndTipVal = 1
    vTipoVale(1, 1) = "Seleccione..."
    vTipoVale(1, 2) = 0
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_ParametroCodigoBarra 0")
    i = 2
    Do While Not RS.EOF
       
       vTipoVale(i, 1) = Trim(RS!atr_nombre)
       vTipoVale(i, 2) = Trim(RS!atr_codigo_barra)
       RS.MoveNext: i = i + 1
    
    Loop
    RS.Close
    Set RS = Nothing
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_s_cliente_V02 13, '" & codigo & "', ''")
    Do While Not RS.EOF
       
       Frame13.Caption = Trim(RS!Cli_codigo) & " - " & Trim(RS!Cli_nombre)
       RS.MoveNext
    
    Loop
    RS.Close
    Set RS = Nothing
    
    vaSpread6.Visible = False
    For i = 1 To vaSpread6.MaxRows
        
        vaSpread6.Row = i
        vaSpread6.Col = 1
        vaSpread6.text = ""

        vaSpread6.Col = 2
        vaSpread6.text = ""

        vaSpread6.Col = 3
        vaSpread6.text = ""

        vaSpread6.Col = 4
        vaSpread6.text = ""
        
        vaSpread6.Col = 5
        vaSpread6.text = ""

        '-------> Mover concepto atributo
        If IndTipVal > 0 Then
           
           lisnom = "": liscod = ""
           
           For j = 1 To UBound(vTipoVale)
               
               vaSpread6.Col = 1: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & vTipoVale(j, 1) ' & " " & Trim(vTipoVale(j, 2))
               vaSpread6.Col = 2: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vTipoVale(j, 2)
               vaSpread6.Col = 1: vaSpread6.TypeComboBoxList = lisnom
               vaSpread6.Col = 2: vaSpread6.TypeComboBoxList = liscod
           
           Next j
          
          vaSpread6.Col = 1: vaSpread6.TypeComboBoxCurSel = 0
          vaSpread6.Col = 2: lisnom = vaSpread6.text
        
        Else
           
           vaSpread6.TypeComboBoxClear 1, i
           vaSpread6.TypeComboBoxClear 2, i
        
        End If
    
    Next i
    i = 1
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_s_casinoparametrolecvales '" & codigo & "'")
    Do While Not RS.EOF
       
       vaSpread6.Row = i

       vaSpread6.Col = 3
       vaSpread6.text = IIf(IsNull(RS!cbar_posinicial), "", RS!cbar_posinicial)

       vaSpread6.Col = 4
       vaSpread6.text = IIf(IsNull(RS!cbar_largo), "", RS!cbar_largo)

       If IndTipVal > 0 Then
          
          vaSpread6.Col = 2
          CodTipVal = -1
          
          For z = 0 To vaSpread6.TypeComboBoxCount
              
              vaSpread6.TypeComboBoxCurSel = z
              If vaSpread6.text = RS!atr_codigo_barra Then CodTipVal = z: Exit For
              CodTipVal = -1
          
          Next z
          
          If CodTipVal = -1 Then CodTipVal = 0
          vaSpread6.Col = 1: vaSpread6.TypeComboBoxCurSel = CodTipVal
          vaSpread6.Col = 2: lisnom = vaSpread6.text

          vaSpread6.Col = 5
          vaSpread6.text = lisnom

       End If
       
       RS.MoveNext: i = i + 1
    
    Loop
    
    RS.Close
    Set RS = Nothing
    Est = False
    vaSpread6.Visible = True

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatos9()
    
On Error GoTo Man_Error

    Dim RS As New ADODB.Recordset
    Dim i As Long

    Me.Refresh
    Est = True
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_s_cliente_V02 13, '" & codigo & "', ''")
    Do While Not RS.EOF
       
       Frame19.Caption = Trim(RS!Cli_codigo) & " - " & Trim(RS!Cli_nombre)
       RS.MoveNext
    
    Loop
    RS.Close
    Set RS = Nothing
    
    vaSpread7.MaxRows = 0
    vaSpread7.Visible = False
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_TraerServicioPreferido '" & codigo & "'")
    Do While Not RS.EOF
       
       vaSpread7.MaxRows = vaSpread7.MaxRows + 1
       vaSpread7.Row = vaSpread7.MaxRows

       vaSpread7.Col = 1
       vaSpread7.text = IIf(IsNull(RS!Reg_Codigo), "", RS!Reg_Codigo)

       vaSpread7.Col = 2
       vaSpread7.text = IIf(IsNull(RS!reg_nombre), "", RS!reg_nombre)
       
       vaSpread7.Col = 3
       vaSpread7.text = IIf(IsNull(RS!Ser_codigo), "", RS!Ser_codigo)
       
       vaSpread7.Col = 4
       vaSpread7.text = IIf(IsNull(RS!ser_nombre), "", RS!ser_nombre)
       
       vaSpread7.Col = 5
       vaSpread7.text = IIf(IsNull(RS!Preferido), "", RS!Activo)
       vaSpread7.Lock = True
       
       vaSpread7.Col = 6
       vaSpread7.text = IIf(IsNull(RS!Activo), "", RS!Preferido)
       
       RS.MoveNext
       i = i + 1
    
    Loop
    
    RS.Close
    Set RS = Nothing
    Est = False
    vaSpread7.Visible = True

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverInvCandelarizado()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i As Long, j As Long
    
    With fpCalendar2
        
        '.CurrentDate = Now
        .Visible = False
        .AutoSet = True
        .DisplayFormat = 3
        colori = .ElementBackColor
        colori = .ElementBackColor
        .ShortDayName(1) = "Dom"
        .ShortDayName(2) = "Lun"
        .ShortDayName(3) = "Mar"
        .ShortDayName(4) = "Mie"
        .ShortDayName(5) = "Jue"
        .ShortDayName(6) = "Vie"
        .ShortDayName(7) = "Sab"
        .LongMonthName(1) = "Enero"
        .LongMonthName(2) = "Febrero"
        .LongMonthName(3) = "Marzo"
        .LongMonthName(4) = "Abril"
        .LongMonthName(5) = "Mayo"
        .LongMonthName(6) = "Junio"
        .LongMonthName(7) = "Julio"
        .LongMonthName(8) = "Agosto"
        .LongMonthName(9) = "Septiembre"
        .LongMonthName(10) = "Octubre"
        .LongMonthName(11) = "Noviembre"
        .LongMonthName(12) = "Diciembre"
        
        For i = 1 To 12
            
            For j = 1 To Val(Mid(dEoM("01/" & fg_pone_cero(i, 2) & "/" & .Year), 1, 2))
                
                .Element = ElementSpecificDate
                .ElementIndex = .Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
                .ElementBackColor = -2147483633
                .ElementText = ""
                .ElementForeColor = vbBlack
            
            Next j
        
        Next i
              
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Set RS = vg_db.Execute("sgpadm_Sel_TraerInventariocalendarizado '" & codigo & "'")
        Do While Not RS.EOF
            
            .ElementIndex = Format((RS!Fecha_Inventario), "yyyymmdd")
            .Element = ElementSpecificDate
            .ElementIndex = Format((RS!Fecha_Inventario), "yyyymmdd")
            .ElementText = "Dia"
            .ElementBackColor = &HFF&
            .ElementForeColor = vbBlack
            RS.MoveNext
        
        Loop
        RS.Close
        Set RS = Nothing
        .Visible = True
    
    End With

    '-------> Mover dias de holgura
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    fpLongInteger1(8).Value = 0
    fpLongInteger1(9).Value = 0
    
    Set RS = vg_db.Execute("sgpadm_Sel_TraerDiasHolguraInvCalendarizado '" & codigo & "'")
        
    If Not RS.EOF Then
    
       fpLongInteger1(8).Value = RS!Dia_Holgura_Antes
       fpLongInteger1(9).Value = RS!Dia_Holgura_Despues
    
    End If
    
    RS.Close
    Set RS = Nothing

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverPvtaClienteCalendarizado()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i As Long, j As Long
    
    With fpCalendar3
        
        '.CurrentDate = Now
        .Visible = False
        .AutoSet = True
        .DisplayFormat = 3
        colori = .ElementBackColor
        colori = .ElementBackColor
        .ShortDayName(1) = "Dom"
        .ShortDayName(2) = "Lun"
        .ShortDayName(3) = "Mar"
        .ShortDayName(4) = "Mie"
        .ShortDayName(5) = "Jue"
        .ShortDayName(6) = "Vie"
        .ShortDayName(7) = "Sab"
        .LongMonthName(1) = "Enero"
        .LongMonthName(2) = "Febrero"
        .LongMonthName(3) = "Marzo"
        .LongMonthName(4) = "Abril"
        .LongMonthName(5) = "Mayo"
        .LongMonthName(6) = "Junio"
        .LongMonthName(7) = "Julio"
        .LongMonthName(8) = "Agosto"
        .LongMonthName(9) = "Septiembre"
        .LongMonthName(10) = "Octubre"
        .LongMonthName(11) = "Noviembre"
        .LongMonthName(12) = "Diciembre"
        
        For i = 1 To 12
            
            For j = 1 To Val(Mid(dEoM("01/" & fg_pone_cero(i, 2) & "/" & .Year), 1, 2))
                
                .Element = ElementSpecificDate
                .ElementIndex = .Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
                .ElementBackColor = -2147483633
                .ElementText = ""
                .ElementForeColor = vbBlack
            
            Next j
        
        Next i
              
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Set RS = vg_db.Execute("sgpadm_Sel_TraerPvtaClienteCalendarizado '" & codigo & "'")
        Do While Not RS.EOF
            
            .ElementIndex = Format((RS!Fecha_PvtaCliente), "yyyymmdd")
            .Element = ElementSpecificDate
            .ElementIndex = Format((RS!Fecha_PvtaCliente), "yyyymmdd")
            .ElementText = "Dia"
            .ElementBackColor = &HFF&
            .ElementForeColor = vbBlack
            RS.MoveNext
        
        Loop
        RS.Close
        Set RS = Nothing
        .Visible = True
    
    End With

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDiasFeriados()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i As Long, j As Long
    
    With fpCalendar1
        
        '.CurrentDate = Now
        .Visible = False
        .AutoSet = True
        .DisplayFormat = 3
        colori = .ElementBackColor
        colori = .ElementBackColor
        .ShortDayName(1) = "Dom"
        .ShortDayName(2) = "Lun"
        .ShortDayName(3) = "Mar"
        .ShortDayName(4) = "Mie"
        .ShortDayName(5) = "Jue"
        .ShortDayName(6) = "Vie"
        .ShortDayName(7) = "Sab"
        .LongMonthName(1) = "Enero"
        .LongMonthName(2) = "Febrero"
        .LongMonthName(3) = "Marzo"
        .LongMonthName(4) = "Abril"
        .LongMonthName(5) = "Mayo"
        .LongMonthName(6) = "Junio"
        .LongMonthName(7) = "Julio"
        .LongMonthName(8) = "Agosto"
        .LongMonthName(9) = "Septiembre"
        .LongMonthName(10) = "Octubre"
        .LongMonthName(11) = "Noviembre"
        .LongMonthName(12) = "Diciembre"
        
        For i = 1 To 12
            
            For j = 1 To Val(Mid(dEoM("01/" & fg_pone_cero(i, 2) & "/" & .Year), 1, 2))
                
                .Element = ElementSpecificDate
                .ElementIndex = .Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
                .ElementBackColor = -2147483633
                .ElementText = ""
                .ElementForeColor = vbBlack
            
            Next j
        
        Next i
              
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Set RS = vg_db.Execute("sgpadm_s_diasferiados 1, '" & codigo & "', '', '', ''")
        Do While Not RS.EOF
            
            .ElementIndex = Format(CDate(RS!CFI_Fecha), "yyyymmdd")
            .Element = ElementSpecificDate
            .ElementIndex = Format(CDate(RS!CFI_Fecha), "yyyymmdd")
            .ElementText = IIf(IsNull(RS!CFI_Glosa), "", Trim(RS!CFI_Glosa))
            .ElementBackColor = &HFF&
            .ElementForeColor = vbBlack
            RS.MoveNext
        
        Loop
        RS.Close
        Set RS = Nothing
        .Visible = True
    
    End With

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub PonerDiasFeriados(op As String, dia As String)

On Error GoTo Man_Error

Dim i As Long, j As Long

For i = 1 To 12
    
    For j = 1 To Val(Mid(dEoM("01/" & fg_pone_cero(i, 2) & "/" & fpCalendar1.Year), 1, 2))
        
        fpCalendar1.ElementIndex = fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
        
        If dia = "Ambos" And (fpCalendar1.ShortDayName(fg_Dia(fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2))) = "Sab" Or _
           fpCalendar1.ShortDayName(fg_Dia(fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2))) = "Dom") Then
           
           fpCalendar1.Element = ElementSpecificDate
           fpCalendar1.ElementIndex = fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
           fpCalendar1.ElementBackColor = IIf(op = "Incluir", &HFF&, -2147483633)
           fpCalendar1.ElementForeColor = vbBlack
        
        ElseIf dia = "Domingo" And fpCalendar1.ShortDayName(fg_Dia(fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2))) = "Dom" Then
           
           fpCalendar1.Element = ElementSpecificDate
           fpCalendar1.ElementIndex = fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
           fpCalendar1.ElementBackColor = IIf(op = "Incluir", &HFF&, -2147483633)
           fpCalendar1.ElementForeColor = vbBlack
        
        ElseIf dia = "Sabado" And fpCalendar1.ShortDayName(fg_Dia(fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2))) = "Sab" Then
           
           fpCalendar1.Element = ElementSpecificDate
           fpCalendar1.ElementIndex = fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
           fpCalendar1.ElementBackColor = IIf(op = "Incluir", &HFF&, -2147483633)
           fpCalendar1.ElementForeColor = vbBlack
        
        End If
    
    Next j
Next i

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Borra_Datos()

Dim fecdfe As Variant
Dim RS     As New ADODB.Recordset

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = vaSpread1.text

If MsgBox(IIf(SSTab1.Tab = 4, "Elimina día...", "Elimina registro..."), vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
   
   If SSTab1.Tab = 4 Then
      
      fpCalendar1.NextSelection = ""
      
      For i = 1 To fpCalendar1.SelCount
          
          fecdfe = fpCalendar1.NextSelection
          
          If Val(fecdfe) > 0 Then
             
             vg_db.Execute "DELETE Cas_b_Fecha_Inhabiles FROM Cas_b_Fecha_Inhabiles WHERE CFI_CeCo = '" & codigo & "' AND CFI_Fecha = '" & Format(fg_Ctod1(fecdfe), "mm/dd/yyyy") & "'"
             fpCalendar1.Element = ElementSpecificDate
             fpCalendar1.ElementIndex = fecdfe
             fpCalendar1.ElementBackColor = -2147483633
             fpCalendar1.ElementText = ""
             fpCalendar1.ElementForeColor = vbBlack
          
          End If
      
      Next i
   
   ElseIf SSTab1.Tab = 5 Or SSTab1.TabVisible(5) = True Then '-------> borrar Parametro Despachos
      
      Dim codpad As Long
      vaSpread4.Row = vaSpread4.ActiveRow
      vaSpread4.Col = 1
      codpad = Val(vaSpread4.text)
      
      If Trim(vaSpread4.text) <> "" Then
         
         vg_db.Execute "DELETE b_parametrodespachos FROM b_parametrodespachos WHERE pad_codtip = " & codpad & " AND pad_cencos = '" & codigo & "'"
         MoverDatos5
      
      End If
   
   ElseIf SSTab1.Tab = 9 Or SSTab1.TabVisible(9) = True Then '-------> borra servicio principal
              
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient

      Set RS = vg_db.Execute("sgpadm_Upd_ServicioPrincipal '" & codigo & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

      If Not RS.EOF Then
  
         If RS(0) > 0 Or RS(0) < 0 Then
          
           fg_descarga
        
           MsgBox "Proceso Eliminación " & RS(1), vbCritical, MsgTitulo
      
           RS.Close
           Set RS = Nothing
                   
           Exit Sub
                
         Else
          
            MsgBox "Proceso eliminación finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
          
        End If
                 
      End If
  
      RS.Close
      Set RS = Nothing
              
      MoverDatos9
   
   ElseIf SSTab1.Tab = 10 Or SSTab1.TabVisible(10) = True Then '-------> borra inventario calendarizado
           
         If SSTab1.Tab = 10 Then
          
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient

            Set RS = vg_db.Execute("sgpadm_Upd_InventarioCalendarizado '" & codigo & "', '" & fpCalendar2.Year & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

            If Not RS.EOF Then
  
               If RS(0) > 0 Or RS(0) < 0 Then
          
                  fg_descarga
        
                  MsgBox "Proceso Eliminación " & RS(1), vbCritical, MsgTitulo
      
                  RS.Close
                  Set RS = Nothing
                   
                  Exit Sub
                
               Else
          
                  MsgBox "Proceso eliminación finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
          
               End If
                 
            End If
  
            RS.Close
            Set RS = Nothing
         
         Else
            
            MsgBox "Debe seleccionar la hoja parmetros inv. calendarizado y su año...", vbInformation + vbOKOnly, Me.Caption
            
         End If
         MoverDatos10
   
   ElseIf SSTab1.Tab = 11 Or SSTab1.TabVisible(11) = True Then '-------> borra Precio Venta cliente calendarizado
           
         If SSTab1.Tab = 11 Then
          
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient

            Set RS = vg_db.Execute("sgpadm_Upd_PvtaClienteCalendarizado '" & codigo & "', '" & fpCalendar3.Year & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

            If Not RS.EOF Then
  
               If RS(0) > 0 Or RS(0) < 0 Then
          
                  fg_descarga
        
                  MsgBox "Proceso Eliminación " & RS(1), vbCritical, MsgTitulo
      
                  RS.Close
                  Set RS = Nothing
                   
                  Exit Sub
                
               Else
          
                  MsgBox "Proceso eliminación finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
          
               End If
                 
            End If
  
            RS.Close
            Set RS = Nothing
         
         Else
            
            MsgBox "Debe seleccionar la hoja parmetros inv. calendarizado y su año...", vbInformation + vbOKOnly, Me.Caption
            
         End If
         MoverDatos11
   
   ElseIf SSTab1.Tab = 12 Or SSTab1.TabVisible(12) = True Then '-------> borra Parametro Categoria Dietetica
                    
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient

            Set RS = vg_db.Execute("sgpadm_Del_ClienteParametroCategoriaDietetica_V01 '" & codigo & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

            If Not RS.EOF Then
  
               If RS(0) > 0 Or RS(0) < 0 Then
          
                  fg_descarga
        
                  MsgBox "Proceso Eliminación " & RS(1), vbCritical, MsgTitulo
      
                  RS.Close
                  Set RS = Nothing
                   
                  Exit Sub
                
               Else
          
                  MsgBox "Proceso eliminación finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
          
               End If
                 
            End If
  
            RS.Close
            Set RS = Nothing
         
         If SSTab1.Tab = 12 Then
         
            MoverDatos12
        
         End If
   
   Else
      
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
    
      Set RS = vg_db.Execute("sgpadm_Del_Casino_V02 '" & codigo & "'")
      If Not RS.EOF Then
       
         If RS(0) > 0 Then
          
            MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
            RS.Close
            Set RS = Nothing
            Exit Sub
       
         End If
    
      End If
      RS.Close
      Set RS = Nothing

      vaSpread1.Row = vaSpread1.ActiveRow
      vaSpread1.DeleteRows vaSpread1.Row, 1
      vaSpread1.MaxRows = vaSpread1.MaxRows - 1
      vaSpread1.Row = vaSpread1.MaxRows
      Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
      Limpia
      
      If vaSpread1.MaxRows < 1 Then
         
         SSTab1.TabEnabled(1) = False
         SSTab1.TabEnabled(2) = False
         SSTab1.TabEnabled(3) = False
         SSTab1.TabEnabled(4) = False
         SSTab1.TabEnabled(8) = False
         SSTab1.Tab = 0
         modo = "NE"
      
      Else
         
         modo = ""
         SSTab1.TabEnabled(1) = True
         SSTab1.TabEnabled(2) = True
         SSTab1.TabEnabled(3) = True
         SSTab1.TabEnabled(4) = True
         SSTab1.TabEnabled(8) = False
         SSTab1.Tab = 0
      
      End If
   
   End If

Exit Sub
Man_Error:

If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Cancela_Datos()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
    
    With SSTab1
        
        If MsgBox("Cancela registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Set RS = vg_db.Execute("sgpadm_s_cliente_V02 6, '',''")
        If RS.EOF Or RS!nReg = 0 Then RS.Close: Set RS = Nothing: .TabEnabled(1) = False: .TabEnabled(2) = False: .TabEnabled(3) = False: modo = "NE": .Tab = 0: Gl_Ac_Botones Me, 13, 2, modo: Exit Sub
        RS.Close
        Set RS = Nothing
        
        If .Tab = 6 Then
           
           .TabEnabled(0) = True
           .TabEnabled(1) = True
           .TabEnabled(2) = True
           .TabEnabled(3) = True
           .TabEnabled(4) = True
           .TabEnabled(6) = True
           .TabEnabled(7) = True
           .TabEnabled(8) = True
           If vaSpread1.MaxRows < 1 Then Exit Sub
           MoverDatos6
        
        ElseIf .Tab = 8 Then
           
           .TabEnabled(0) = True
           .TabEnabled(1) = True
           .TabEnabled(2) = True
           .TabEnabled(3) = True
           .TabEnabled(4) = True
           .TabEnabled(6) = True
           .TabEnabled(7) = True
           .TabEnabled(8) = True
           If vaSpread1.MaxRows < 1 Then Exit Sub
           MoverDatos8
        
        Else
           
           Limpia
           
           If vaSpread1.MaxRows > 0 Then
              
              .TabEnabled(1) = True
              .TabEnabled(2) = True
              .TabEnabled(3) = True
              .TabEnabled(4) = True
              .TabEnabled(6) = True
              .TabEnabled(7) = True
              .TabEnabled(8) = True
           
           Else
              
              .TabEnabled(1) = False
              .TabEnabled(2) = False
              .TabEnabled(3) = False
              .TabEnabled(4) = False
              .TabEnabled(6) = False
              .TabEnabled(7) = False
              .TabEnabled(8) = False
           
           End If
           
           .TabEnabled(0) = True
           .Tab = 0
        
        End If
        
        modo = ""
        Gl_Ac_Botones Me, 13, IIf(lc_Aux = "MCasino", 1, 3), modo
    
    End With

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Actualiza_Datos()

On Error GoTo Man_Error

Dim RS               As New ADODB.Recordset
Dim tipo             As Integer
Dim GruVul           As String
Dim hipali           As String
Dim modpac           As String
Dim IdIntegraAMD     As Long
Dim IdIntegraSPRS    As Long
Dim i                As Long
Dim j                As Long
Dim sobrec           As String
Dim X                As Long
Dim codmun           As Long
Dim codrgi           As Long
Dim opgped           As String
Dim tipope           As String
Dim emailenviopedido As String
Dim CodOptimun       As String
Dim socsap           As String
Dim codsgp           As String
Dim tipoceco         As String
Dim AMD              As String

'INI ARI
' Validacion donde debe Elegir Tipo de Propuesta y Formato de Compras

Dim tipoestructura As Integer
Dim formacompras As Integer

tipo = 0: opgped = ""
GruVul = IIf(Check1(1).Value = 1, "S", "N")
hipali = IIf(Check1(3).Value = 1, "S", "N")
modpac = IIf(Check1(2).Value = 1, "S", "N")
IdIntegraAMD = IIf(Check4.Value = 1, 1, 2)
IdIntegraSPRS = IIf(Check5.Value = 1, 1, 2)
sobrec = IIf(Option1(0).Value = True, "0", IIf(Option1(1).Value = True, "1", "2"))
codmun = IIf(fpLongInteger1(5).text = "", 0, fpLongInteger1(5).Value)
codrgi = IIf(fpLongInteger1(7).text = "", 0, fpLongInteger1(7).Value)
opgped = IIf(Option3(0).Value = True, "0", "1")
tipope = IIf(Option3(3).Value = True, "0", "1")
CodOptimun = LimpiaDato(Trim(fpText(16).text))
socsap = LimpiaDato(Trim(fpText(13).text))
tipoceco = IIf(Option4(0).Value = True, "0", IIf(Option4(1).Value = True, "1", ""))
AMD = IIf(Option5(0).Value = True, "2", IIf(Option5(1).Value = True, "1", ""))

If modo = "A" Then
   
   If Trim(fpText(0).text) = "" Or Trim(fpText(1).text) = "" Then MsgBox "Faltan datos importantes para identificar el cliente...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   If Trim(fpText(11).text) = "" Then MsgBox "Faltan descripción del contable...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   If Trim(fpText(12).text) = "" Then MsgBox "Faltan correo del contable...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   If Trim(fpayuda(3).Caption) = "" Then MsgBox "Faltan tipo servicio...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   If Trim(fpayuda(7).Caption) = "" Then MsgBox "Faltan Segmento...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   If Trim(fpText(13).text) = "" Then MsgBox "Faltan Sociedad SAP...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   If socsap <> "" And CodOptimun = "" Then MsgBox "Faltan código optimun...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   If socsap = "" And CodOptimun <> "" Then MsgBox "Faltan sociedad sap...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   If InStr(Trim(fpText(15).text), ";") <> 0 Then MsgBox "La separaciones email envio pedido, debe ser un caracter coma...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   If Trim(tipoceco) = "" Then MsgBox "Faltan definir concepto tipo ceco...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   If Trim(AMD) = "" Then MsgBox "Faltan definir concepto tipo acceso minuta...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   
   If Combo1(2) = "" Then
      
      MsgBox "Debe Seleccionar Tipo de Propuesta", 16
      Combo1(2).SetFocus
      Exit Sub
   
   Else
      
      If Combo1(3) = "" Then
         
         MsgBox "Debe Seleccionar Formato de Compras", 16
         Combo1(3).SetFocus
         Exit Sub
      
      End If
   
   End If
      
   tipoestructura = IIf(Combo1(2).ListIndex = -1, "Null", Val(fg_codigocbo(Combo1, 2, 1, "")))
   formacompras = IIf(Combo1(3).ListIndex = -1, "Null", Val(fg_codigocbo(Combo1, 3, 1, "")))
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
    
   Set RS = vg_db.Execute("sgpadm_s_cliente_V02 1, '" & LimpiaDato(Trim(fpText(0).text)) & "',''")
   If Not RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Cliente existe", vbExclamation + vbOKOnly, "Maestro de Clientes": Exit Sub
   RS.Close: Set RS = Nothing
   
   '--> validar que no exista codigo optimun
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
    
   Set RS = vg_db.Execute("sgpadm_Sel_ValidarCodigoOptimun '" & LimpiaDato(Trim(fpText(0).text)) & "','" & CodOptimun & "'")
   If Not RS.EOF Then
      
      codsgp = RS!Cecos_Sap
      RS.Close
      Set RS = Nothing
      MsgBox "Código optimun esta asociado código sgp : " & codsgp, vbExclamation + vbOKOnly, "Maestro de Clientes"
      Exit Sub
   
   End If
   RS.Close
   Set RS = Nothing
   
   'MVA 2012-01-07 - CAMBIO DE NOMBRE DEL SP
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_iu_clientes_Ver7 'A', '" & LimpiaDato(Trim(fpText(0).text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText(1).text)) & "', '" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText(3).text)) & "', '" & LimpiaDato(Trim(fpText(4).text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText(5).text)) & "', '" & LimpiaDato(Trim(fpText(6).text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText(7).text)) & "', '" & LimpiaDato(Trim(fpText(8).text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText(9).text)) & "', '" & LimpiaDato(Trim(fpText(10).text)) & "', " & _
                   "" & tipo & ", " & IIf(Combo1(1).ListIndex = -1, "Null", Val(fg_codigocbo(Combo1, 1, 1, ""))) & ", " & _
                   "" & Val(fpLongInteger1(0).Value) & ",'" & LimpiaDato(Trim(fpText(11).text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText(12).text)) & "', '" & GruVul & "', " & Val(fpLongInteger1(1).Value) & ", " & _
                   "'" & modpac & "', " & Val(fpLongInteger1(2).Value) & ", " & Val(fpLongInteger1(3).Value) & ", " & _
                   "'" & LimpiaDato(Trim(fpText(13).text)) & "', '" & IIf(Check1(0).Value = 1, "1", "0") & "', " & _
                   "'" & sobrec & "', " & codmun & ", " & Val(fpLongInteger1(6).Value) & ", " & _
                   "'" & LimpiaDato(Trim(fpText(14).text)) & "', '" & opgped & "', '" & hipali & "', " & _
                   "" & codrgi & ", '" & tipope & "', '" & IIf(ChkMInRet.Value = 1, "1", "0") & "', " & _
                   "'" & IIf(Check3.Value = 1, "1", "0") & "','" & IIf(ChkBlockMinReal.Value = 1, "1", "0") & "', " & _
                   "'" & IIf(ChkBlockMinTeo.Value = 1, "1", "0") & "', '" & IIf(ChkBlockMinContrato.Value = 1, "1", "0") & "', " & _
                   "'" & IIf(ChkBlockTraFinSemana.Value = 1, "1", "0") & "', '" & LimpiaDato(Trim(fpText(15).text)) & "', " & _
                   "" & tipoestructura & "," & formacompras & ", '" & CodOptimun & "', '" & tipoceco & "', '" & AMD & "', " & IdIntegraAMD & ", " & IdIntegraSPRS & "")
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
       
       Else
          
          MsgBox "Proceso Finalizado [OK]", vbInformation, Me.Caption
       
       End If
    
    End If
    RS.Close: Set RS = Nothing
   
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = -1
   vaSpread1.BackColor = Shape1(0).FillColor
   vaSpread1.Col = 1
   vaSpread1.TypeHAlign = TypeHAlignLeft
   vaSpread1.Value = LimpiaDato(Trim(fpText(0).text))
   vaSpread1.Col = 2
   vaSpread1.TypeHAlign = TypeHAlignLeft
   vaSpread1.Value = LimpiaDato(Trim(fpText(1).text))
   vaSpread1.Col = 3
   vaSpread1.Value = IIf(tipo = 0, "Operador", "Traspaso")
   vaSpread1.Col = 4
   If Combo1(1).ListIndex <> -1 Then vaSpread1.Value = Trim(Mid(Combo1(1).text, 1, 10)) Else vaSpread1.text = "No Especificado"
   vaSpread1.Col = 5
   vaSpread1.TypeHAlign = TypeHAlignLeft
   vaSpread1.Value = Trim(fpayuda(1).Caption)
   vaSpread1.Col = 6
   vaSpread1.Value = IIf(Check1(0).Value = 1, "1", "0")

Else
   
   If SSTab1.Tab = 1 Then
      
      If Trim(fpText(1).text) = "" Then MsgBox "Faltan datos importantes para identificar el cliente...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      If Trim(fpText(11).text) = "" Then MsgBox "Faltan descripción del contable...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      If Trim(fpText(12).text) = "" Then MsgBox "Faltan correo del contable...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      If Trim(fpayuda(3).Caption) = "" Then MsgBox "Faltan tipo servicio...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      If Trim(fpayuda(7).Caption) = "" Then MsgBox "Faltan Segmento...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      If Trim(fpText(13).text) = "" Then MsgBox "Faltan Sociedad SAP...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      If socsap <> "" And CodOptimun = "" Then MsgBox "Faltan código optimun...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      If socsap = "" And CodOptimun <> "" Then MsgBox "Faltan sociedad sap...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      If InStr(Trim(fpText(15).text), ";") <> 0 Then MsgBox "La separaciones email envio pedido, debe ser un caracter coma...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      If Trim(tipoceco) = "" Then MsgBox "Faltan definir concepto tipo ceco...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
      If Trim(AMD) = "" Then MsgBox "Faltan definir concepto tipo acceso minuta...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   
      If Combo1(2) = "" Then
         
         MsgBox "Debe Seleccionar Tipo de Propuesta", 16
         Combo1(2).SetFocus
         Exit Sub
      
      Else
         
         If Combo1(3) = "" Then
            
            MsgBox "Debe Seleccionar Formato de Compras", 16
            Combo1(3).SetFocus
            Exit Sub
         
         End If
      
      End If
      
      tipoestructura = IIf(Combo1(2).ListIndex = -1, "Null", Val(fg_codigocbo(Combo1, 2, 1, "")))
      formacompras = IIf(Combo1(3).ListIndex = -1, "Null", Val(fg_codigocbo(Combo1, 3, 1, "")))
      
      '--> validar que no exista codigo optimun
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS = vg_db.Execute("sgpadm_Sel_ValidarCodigoOptimun '" & LimpiaDato(Trim(fpText(0).text)) & "','" & CodOptimun & "'")
      If Not RS.EOF Then
         
         codsgp = RS!Cecos_Sap
         RS.Close
         Set RS = Nothing
         MsgBox "Código optimun esta asociado código sgp : " & codsgp, vbExclamation + vbOKOnly, "Maestro de Clientes"
         Exit Sub
      
      End If
      RS.Close
      Set RS = Nothing
   
      'MVA 2012-01-07 - CAMBIO DE NOMBRE DEL SP
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS = vg_db.Execute("sgpadm_iu_clientes_Ver7 'M', '" & LimpiaDato(Trim(fpText(0).text)) & "', " & _
                      "'" & LimpiaDato(Trim(fpText(1).text)) & "', '" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                      "'" & LimpiaDato(Trim(fpText(3).text)) & "', '" & LimpiaDato(Trim(fpText(4).text)) & "', " & _
                      "'" & LimpiaDato(Trim(fpText(5).text)) & "', " & "'" & LimpiaDato(Trim(fpText(6).text)) & "', " & _
                      "'" & LimpiaDato(Trim(fpText(7).text)) & "', '" & LimpiaDato(Trim(fpText(8).text)) & "', " & _
                      "'" & LimpiaDato(Trim(fpText(9).text)) & "', '" & LimpiaDato(Trim(fpText(10).text)) & "', " & _
                      "" & tipo & ", " & IIf(Combo1(1).ListIndex = -1, "Null", Val(fg_codigocbo(Combo1, 1, 1, ""))) & ", " & _
                      "" & Val(fpLongInteger1(0).Value) & ", '" & LimpiaDato(Trim(fpText(11).text)) & "', " & _
                      "'" & LimpiaDato(Trim(fpText(12).text)) & "', '" & GruVul & "', " & Val(fpLongInteger1(1).Value) & ", " & _
                      "'" & modpac & "', " & Val(fpLongInteger1(2).Value) & ", " & Val(fpLongInteger1(3).Value) & ", " & _
                      "'" & LimpiaDato(Trim(fpText(13).text)) & "', '" & IIf(Check1(0).Value = 1, "1", "0") & "', " & _
                      "'" & sobrec & "', " & codmun & ", " & Val(fpLongInteger1(6).Value) & ", " & _
                      "'" & LimpiaDato(Trim(fpText(14).text)) & "', '" & opgped & "', '" & hipali & "', " & _
                      "" & codrgi & ", '" & tipope & "', '" & IIf(ChkMInRet.Value = 1, "1", "0") & "', " & _
                      "'" & IIf(Check3.Value = 1, "1", "0") & "','" & IIf(ChkBlockMinReal.Value = 1, "1", "0") & "', " & _
                      "'" & IIf(ChkBlockMinTeo.Value = 1, "1", "0") & "', '" & IIf(ChkBlockMinContrato.Value = 1, "1", "0") & "', " & _
                      "'" & IIf(ChkBlockTraFinSemana.Value = 1, "1", "0") & "', '" & LimpiaDato(Trim(fpText(15).text)) & "', " & _
                      "" & tipoestructura & "," & formacompras & ", '" & CodOptimun & "', '" & tipoceco & "', '" & AMD & "', " & IdIntegraAMD & ", " & IdIntegraSPRS & "")
                      
      If Not RS.EOF Then
         
         If RS(0) > 0 Then
            
            MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
         
         Else
            
            MsgBox "Proceso Finalizado [OK]", vbInformation, Me.Caption
         
         End If
      
      End If
      RS.Close: Set RS = Nothing
      
      '------- Si sub-segmento es igual cero eliminar relación subsegmento
      If Val(fpLongInteger1(0).Value) = 0 Then
         
         vg_db.Execute "DELETE b_casinoregser FROM b_casinoregser WHERE crs_cencos='" & LimpiaDato(Trim(fpText(0).text)) & "'"
         vg_db.Execute "DELETE b_paramcostopatron FROM b_paramcostopatron WHERE pcp_cencos='" & LimpiaDato(Trim(fpText(0).text)) & "'"
      
      End If
      
      vaSpread1.Col = 2
      vaSpread1.Value = LimpiaDato(Trim(fpText(1).text))
      vaSpread1.Col = 3
      vaSpread1.Value = IIf(tipo = 0, "Operador", "Traspaso")
      vaSpread1.Col = 4
      If Combo1(1).ListIndex <> -1 Then vaSpread1.Value = Trim(Mid(Combo1(1).text, 1, 10)) Else vaSpread1.text = "No Especificado"
      vaSpread1.Col = 5
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.Value = Trim(fpayuda(1).Caption)
      vaSpread1.Col = 6
      vaSpread1.Value = IIf(Check1(0).Value = 1, "1", "0")
   
   ElseIf SSTab1.Tab = 2 Then
      
      Dim codReg As Long, codser As Long
      For i = 1 To TvwDir.Nodes.count
      
      'Nodes.Child.text <> "*"
          
          If TvwDir.Nodes.item(i).Children > 0 Then
             
             '------- Sacar codigo regimen
             codReg = Val(Mid(TvwDir.Nodes.item(i).text, 1, InStr(TvwDir.Nodes.item(i).text, "-") - 1))
             codser = 0
          
          ElseIf TvwDir.Nodes.item(i).text <> "*" Then
             
             '------- Sacar codigo servicio
             codser = Val(Mid(TvwDir.Nodes.item(i).text, 1, InStr(TvwDir.Nodes.item(i).text, "-") - 1))
             codReg = Val(Mid(TvwDir.Nodes.item(i).key, 12, 21))
          
          End If
          
          If TvwDir.Nodes.item(i).Checked = True Then
             
             If codReg > 0 And codser > 0 Then
                
                If RS.State = 1 Then RS.Close
                RS.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                RS.Open "SELECT * FROM b_casinoregser WHERE crs_cencos = '" & codigo & "' AND crs_codreg = " & codReg & " AND crs_codser = " & codser & "", vg_db, adOpenStatic
                If RS.EOF Then vg_db.Execute "INSERT INTO b_casinoregser (crs_cencos, crs_codreg, crs_codser) VALUES ('" & codigo & "', " & codReg & ", " & codser & ")"
                RS.Close
                Set RS = Nothing
             
             End If
          
          ElseIf codReg > 0 And codser > 0 Then
             
             vg_db.Execute "DELETE b_casinoregser FROM b_casinoregser WHERE crs_cencos = '" & codigo & "' AND crs_codreg = " & codReg & " AND crs_codser =  " & codser & ""
          
          End If
      
      Next i

   ElseIf SSTab1.Tab = 3 Then
      
      Dim codens  As Long
      vg_db.Execute "DELETE b_casinointerfaz WHERE cai_cencos = '" & codigo & "'"
      
      For i = 1 To vaSpread2.MaxRows
          
          vaSpread2.Row = i
          vaSpread2.Col = 2
          codens = Val(vaSpread2.text)
          vaSpread2.Col = 1
'          If vaSpread2.text = "1" Then vg_db.Execute "INSERT INTO b_casinointerfaz (cai_cencos, cai_codtii) VALUES ('" & Codigo & "', " & codens & ")"
          If vaSpread2.text = "1" Then vg_db.Execute "sgpadm_iu_casinointerfaz 'A', '" & codigo & "', " & codens & ""
      
      Next i
   
   ElseIf SSTab1.Tab = 4 Then
      
      Dim feccal As Long
      vg_db.Execute "DELETE Cas_b_Fecha_Inhabiles WHERE CFI_CeCo = '" & codigo & "' AND Year(CFI_Fecha) = '" & fpCalendar1.Year & "'"
      
      For i = 1 To 12
          
          For j = 1 To Val(Mid(dEoM("01/" & fg_pone_cero(i, 2) & "/" & fpCalendar1.Year), 1, 2))
              
              fpCalendar1.ElementIndex = fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
              
              If fpCalendar1.ElementBackColor = &HFF& Then
                 
                 feccal = fpCalendar1.ElementIndex
                 vg_db.Execute "INSERT INTO Cas_b_Fecha_Inhabiles (CFI_CeCo, CFI_Fecha, CFI_Glosa) VALUES ('" & codigo & "', '" & Format(fg_Ctod1(feccal), "yyyymmdd") & "', '" & Trim(fpCalendar1.ElementText) & "')"
              
              End If
            
            Next j
      
      Next i
   
   ElseIf SSTab1.Tab = 5 Then
      
      Dim tipdes As String, estext As Boolean, vCodFam As Long, desdia As String, candiaseg As Long, tipdes1 As String
      '-------> Validar datos parametro despachos
      
      For i = 1 To vaSpread4.MaxRows
          
          vaSpread4.Row = i
          vaSpread4.Col = 1
          
          If Trim(vaSpread4.text) <> "" Then
             
             vaSpread4.Col = 5
             tipdes = vaSpread4.text
             
             If tipdes = "E" Or tipdes = "S" Then
                
                estext = False
                
                For j = 8 To vaSpread4.maxcols
                    
                    vaSpread4.Col = j
                    If vaSpread4.text <> "0" And Trim(vaSpread4.text) <> "" Then estext = True
                
                Next j
                
                If Not estext Then MsgBox "No a especificado los días de despachos", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
             
             End If
             
             vaSpread4.Col = 7: tipdes1 = vaSpread4.text
             
             If Trim(tipdes) = "Q" And Trim(tipdes1) = "" Then
                
                MsgBox "No a especificado la segunda quincena de despachos", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
             
             End If
          
          End If
      
      Next i

      vg_db.Execute "DELETE b_parametrodespachos WHERE pad_cencos = '" & codigo & "'"
      For i = 1 To vaSpread4.MaxRows
          
          vaSpread4.Row = i
          vaSpread4.Col = 1
          vCodFam = Val(vaSpread4.text)
          vaSpread4.Col = 5
          tipdes = ""
          tipdes = vaSpread4.text
          
          If Trim(vCodFam) <> "" And tipdes <> "" Then
             
             vaSpread4.Col = 1
             vCodFam = Val(vaSpread4.text)
             vaSpread4.Col = 3
             candiaseg = Val(vaSpread4.text)
             vaSpread4.Col = 5
             tipdes = vaSpread4.text
             
             If tipdes = "Q" Then
                
                vaSpread4.Col = 7: tipdes = vaSpread4.text
             
             End If
             desdia = ""
             
             If tipdes = "E" Or tipdes = "S" Then
                
                X = 1
                
                For j = 8 To vaSpread4.maxcols
                    
                    vaSpread4.Col = j
                    desdia = desdia & IIf(Trim(vaSpread4.text) = "" Or Trim(vaSpread4.text) = "0", "0", X) 'vaSpread4.text)
                    X = X + 1
                
                Next j
             
             End If
             
             vg_db.Execute "INSERT INTO b_parametrodespachos (pad_cencos, pad_codtip, pad_tipo, pad_diaseg, pad_diario) VALUES ('" & codigo & "', " & vCodFam & ", '" & tipdes & "', " & candiaseg & ", '" & desdia & "')"
          
          End If
      
      Next i
   
   ElseIf SSTab1.Tab = 6 Then
      
      Dim tipact As Long
      vg_db.Execute "DELETE b_casinotipoactividades WHERE cta_cencos = '" & codigo & "'"
      
      For i = 1 To vaSpread5.MaxRows
          
          vaSpread5.Row = i
          vaSpread5.Col = 3
          
          If vaSpread5.text = "1" Then
             
             vaSpread5.Col = 1
             tipact = Val(vaSpread5.text)
             vg_db.Execute "INSERT INTO b_casinotipoactividades (cta_cencos, cta_tipact) VALUES ('" & codigo & "', " & tipact & ")"
          
          End If
      
      Next i
   
   ElseIf SSTab1.Tab = 7 Then
      
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
    
      Set RS = vg_db.Execute("sgpadm_s_casinoparametrostock '" & codigo & "'")
      If RS.EOF Then
         
         vg_db.Execute "sgpadm_iu_casinoparametrostock 'A', '" & codigo & "', '" & IIf(Option2(0).Value = True, "S", "N") & "', '" & IIf(Option2(1).Value = True, "S", "N") & "', " & IIf(fpLongInteger1(4).text = "", 0, fpLongInteger1(4).Value) & ",  '" & IIf(Combo2(0).ListIndex = -1, "0", Val(fg_codigocbo(Combo2, 0, 1, ""))) & "', '" & IIf(Check2(2).Value = 1, "S", "N") & "', '" & IIf(Check2(3).Value = 1, "S", "N") & "'"
      
      Else
         
         vg_db.Execute "sgpadm_iu_casinoparametrostock 'M', '" & codigo & "', '" & IIf(Option2(0).Value = True, "S", "N") & "', '" & IIf(Option2(1).Value = True, "S", "N") & "', " & IIf(fpLongInteger1(4).text = "", 0, fpLongInteger1(4).Value) & ",  '" & IIf(Combo2(0).ListIndex = -1, "0", Val(fg_codigocbo(Combo2, 0, 1, ""))) & "', '" & IIf(Check2(2).Value = 1, "S", "N") & "', '" & IIf(Check2(3).Value = 1, "S", "N") & "'"
      
      End If
      RS.Close
      Set RS = Nothing
   
   ElseIf SSTab1.Tab = 8 Then
      
      Dim BackColor As String
      Dim Atributo As String
      Dim Atributo1 As String
      Dim PosInicial As Long
      Dim Largo As Long
      Dim TipVal As String
      Dim TipVal1 As String
      '-------> Validar descripcion no se repita
      For j = 1 To vaSpread6.MaxRows
          
          vaSpread6.Col = 1: vaSpread6.Row = j
          Atributo = vaSpread6.text
          Atributo1 = vaSpread6.text
          vaSpread6.Col = 3
          PosInicial = IIf(vaSpread6.text = "", 0, vaSpread6.text)
          vaSpread6.Col = 4
          Largo = IIf(vaSpread6.text = "", 0, vaSpread6.text)
          vaSpread6.Col = 5
          TipVal = vaSpread6.text
          BackColor = vaSpread6.BackColor
          If Atributo <> "" And PosInicial = 0 Then MsgBox "Favor ingresar posición inicial, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, MsgTitulo: vaSpread6.Row = j: vaSpread6.Col = 3: vaSpread6.SetActiveCell 3, vaSpread6.Row: vaSpread6.SetFocus: Exit Sub
          If Atributo <> "" And Largo = 0 Then MsgBox "Favor ingresar largo, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, MsgTitulo: vaSpread6.Row = j: vaSpread6.Col = 4: vaSpread6.SetActiveCell 4, vaSpread6.Row: vaSpread6.SetFocus: Exit Sub
          If Atributo <> "" And TipVal = "" Or TipVal = "0" Then MsgBox "Favor ingresar Atributo, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, MsgTitulo: vaSpread6.Row = j: vaSpread6.Col = 1: vaSpread6.SetActiveCell 1, vaSpread6.Row: vaSpread6.SetFocus: Exit Sub
          
          If Atributo <> "" Then
             
             For i = 1 To vaSpread6.MaxRows
                 
                 vaSpread6.Col = 1
                 vaSpread6.Row = i
                 Atributo1 = vaSpread6.text
                 If UCase(Trim(Atributo1)) = UCase(Trim(Atributo)) And j <> i And BackColor = vaSpread6.BackColor Then MsgBox "Descripción ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread6.Row = j: vaSpread6.Col = 1: vaSpread6.SetActiveCell 1, vaSpread6.Row: vaSpread6.SetFocus: Exit Sub
                 vaSpread6.Col = 5
                 TipVal1 = vaSpread6.text
                 If UCase(Trim(Atributo1)) = UCase(Trim(Atributo)) And j <> i And BackColor <> vaSpread6.BackColor And TipVal = TipVal1 Then MsgBox "Tipo vale ya existe, para este atributo en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread6.Row = j: vaSpread6.Col = 1: vaSpread6.SetActiveCell 1, vaSpread6.Row: vaSpread6.SetFocus: Exit Sub
             
             Next i
          
          End If
      
      Next j
      '-------< Borrar datos
      vg_db.Execute "DELETE a_par_codigo_barra WHERE cli_codigo = '" & codigo & "'"
      '-------> Insertar datos
      For i = 1 To vaSpread6.MaxRows
          
          vaSpread6.Row = i
          
          vaSpread6.Col = 1
          Atributo = vaSpread6.text
          If Atributo <> "" Then
            
            vaSpread6.Col = 3
            PosInicial = vaSpread6.text

            vaSpread6.Col = 4
            Largo = vaSpread6.text
          
            vaSpread6.Col = 2
            TipVal = vaSpread6.text
          
            vg_db.Execute "INSERT INTO a_par_codigo_barra (atr_codigo_barra, cli_codigo, cbar_posinicial, cbar_largo) VALUES (" & CCur(TipVal) & ", '" & codigo & "', " & PosInicial & ", " & Largo & ")"
        
        End If
      
      Next i
      
      MsgBox "Registro guardo exitosamente", vbInformation + vbOKOnly, MsgTitulo
      
   ElseIf SSTab1.Tab = 9 Then
      
      Dim seleccion As Boolean
      Dim MyBufferS As String
      Dim regimen As Long
      Dim Servicio As Long
      
      seleccion = False
      For i = 1 To vaSpread7.MaxRows
      
          vaSpread7.Row = i
          vaSpread7.Col = 6
          If vaSpread7.text = "1" Then
             
             seleccion = True
             
          End If
      
      Next i
      
      If Not seleccion Then
      
         MsgBox "Debe Seleccionar a lo menos un servicio...", vbInformation + vbOKOnly, MsgTitulo
         Exit Sub
      
      End If
      
      Let MyBufferS = ""
      Let MyBufferS = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
      Let MyBufferS = MyBufferS & "<ServicioPri>"

      For i = 1 To vaSpread7.MaxRows
      
          vaSpread7.Row = i
          vaSpread7.Col = 6
          If vaSpread7.text = "1" Then
             
             vaSpread7.Col = 1
             regimen = vaSpread7.text
             
             vaSpread7.Col = 3
             Servicio = vaSpread7.text
             
             MyBufferS = MyBufferS & " <DetServicioPri"
             MyBufferS = MyBufferS & " Reg = " & Chr(34) & regimen & Chr(34)
             MyBufferS = MyBufferS & " Ser = " & Chr(34) & Servicio & Chr(34)
             MyBufferS = MyBufferS & "/>"
             
          End If
      
      Next i
        
      MyBufferS = MyBufferS & "</ServicioPri>"
      
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient

      Set RS = vg_db.Execute("sgpadm_InsUpd_XmlServicioPrincipal '" & MyBufferS & "', '" & codigo & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

      If Not RS.EOF Then
  
         If RS(0) > 0 Or RS(0) < 0 Then
          
           fg_descarga
        
           MsgBox RS(1), vbCritical, MsgTitulo
      
           RS.Close
           Set RS = Nothing
                   
           Exit Sub
                
         Else
          
            MsgBox "Proceso finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
          
        End If
                 
      End If
  
      RS.Close
      Set RS = Nothing
    
   ElseIf SSTab1.Tab = 10 Then
      
      Dim feccal2  As Long
      Dim MyBuffer As String

      Let MyBuffer = ""
      Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
      Let MyBuffer = MyBuffer & "<InventarioCal>"

      For i = 1 To 12
          
          For j = 1 To Val(Mid(dEoM("01/" & fg_pone_cero(i, 2) & "/" & fpCalendar2.Year), 1, 2))
              fpCalendar2.Element = ElementSpecificDate
              fpCalendar2.ElementIndex = fpCalendar2.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
              
              If fpCalendar2.ElementBackColor = &HFF& Then
                 
                 feccal2 = fpCalendar2.ElementIndex
                 
                 MyBuffer = MyBuffer & " <DetInventarioCal"
                 MyBuffer = MyBuffer & " Fec = " & Chr(34) & feccal2 & Chr(34)
                 MyBuffer = MyBuffer & "/>"
             
              End If
            
            Next j
      
      Next i
   
      MyBuffer = MyBuffer & "</InventarioCal>"
      
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient

      Set RS = vg_db.Execute("sgpadm_UpdIns_XmlInventarioCalendarizado '" & MyBuffer & "', '" & codigo & "', '" & fpCalendar2.Year & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "', " & fpLongInteger1(8).Value & ", " & fpLongInteger1(9).Value & "")

      If Not RS.EOF Then
  
         If RS(0) > 0 Or RS(0) < 0 Then
          
           fg_descarga
        
           MsgBox RS(1), vbCritical, MsgTitulo
      
           RS.Close
           Set RS = Nothing
                   
           Exit Sub
                
         Else
          
            MsgBox "Proceso finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
          
        End If
                 
      End If
  
      RS.Close
      Set RS = Nothing
   
   ElseIf SSTab1.Tab = 11 Then
      
      Dim feccal3  As Long
      Dim MyBuffer1 As String

      Let MyBuffer1 = ""
      Let MyBuffer1 = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
      Let MyBuffer1 = MyBuffer1 & "<PrecioVta>"

      For i = 1 To 12
          
          For j = 1 To Val(Mid(dEoM("01/" & fg_pone_cero(i, 2) & "/" & fpCalendar3.Year), 1, 2))
              fpCalendar3.Element = ElementSpecificDate
              fpCalendar3.ElementIndex = fpCalendar3.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
              
              If fpCalendar3.ElementBackColor = &HFF& Then
                 
                 feccal3 = fpCalendar3.ElementIndex
                 
                 MyBuffer1 = MyBuffer1 & " <DetPvta"
                 MyBuffer1 = MyBuffer1 & " Fec = " & Chr(34) & feccal3 & Chr(34)
                 MyBuffer1 = MyBuffer1 & "/>"
             
              End If
            
            Next j
      
      Next i
   
      MyBuffer1 = MyBuffer1 & "</PrecioVta>"
      
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient

      Set RS = vg_db.Execute("sgpadm_UpdIns_XmlPvtaClienteCalendarizado '" & MyBuffer1 & "', '" & codigo & "', '" & fpCalendar3.Year & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

      If Not RS.EOF Then
  
         If RS(0) > 0 Or RS(0) < 0 Then
          
           fg_descarga
        
           MsgBox RS(1), vbCritical, MsgTitulo
      
           RS.Close
           Set RS = Nothing
                   
           Exit Sub
                
         Else
          
            MsgBox "Proceso finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
          
        End If
                 
      End If
  
      RS.Close
      Set RS = Nothing
      
   ElseIf SSTab1.Tab = 12 Then
      
      Dim XmlDietetica As String
      Dim IndFiltro    As Long
        
      '---------> Armar Xml Categoria Dietetica & Tipo Plato
      Let XmlDietetica = ""
      Let XmlDietetica = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
      Let XmlDietetica = XmlDietetica & "<Dietetica>"
        
      For IndFiltro = 1 To TvwDietetica(0).Nodes.count
        
          If TvwDietetica(0).Nodes.item(IndFiltro).Checked = True And Trim(TvwDietetica(0).Nodes.item(IndFiltro).text) <> "*" Then
           
             XmlDietetica = XmlDietetica & " <DetDietetica"
        
             XmlDietetica = XmlDietetica & " Die = " & Chr(34) & Val(Mid(TvwDietetica(0).Nodes.item(IndFiltro).text, 1, InStr(TvwDietetica(0).Nodes.item(IndFiltro).text, " - ") - 1)) & Chr(34)
             XmlDietetica = XmlDietetica & "/>"
        
          End If
        
      Next IndFiltro
        
      XmlDietetica = XmlDietetica & "</Dietetica>"

      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient

      Set RS = vg_db.Execute("sgpadm_Ins_XmlClienteParametroCategoriaDietetica_V01 '" & XmlDietetica & "', '" & codigo & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

      If Not RS.EOF Then
  
         If RS(0) > 0 Or RS(0) < 0 Then
          
           fg_descarga
        
           MsgBox RS(1), vbCritical, MsgTitulo
      
           RS.Close
           Set RS = Nothing
                   
           Exit Sub
                
         Else
          
            MsgBox "Proceso finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
          
        End If
                 
      End If
  
      RS.Close
      Set RS = Nothing
      
   End If

End If
     
vaSpread1.SortKey(1) = 2
vaSpread1.SortKeyOrder(1) = 1
vaSpread1.Sort 1, 1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
     
SSTab1.TabEnabled(0) = True

If vaSpread1.MaxRows < 1 Then
   
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
   SSTab1.TabEnabled(3) = False
   SSTab1.TabEnabled(4) = False
   SSTab1.TabEnabled(6) = False
   SSTab1.TabEnabled(7) = False
   SSTab1.TabEnabled(8) = False

Else
   
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
   SSTab1.TabEnabled(3) = True
   SSTab1.TabEnabled(4) = True
   SSTab1.TabEnabled(6) = True
   SSTab1.TabEnabled(7) = True
   SSTab1.TabEnabled(8) = True
   SSTab1.Tab = IIf(SSTab1.Tab = 4, 4, IIf(SSTab1.Tab = 8, 8, 0))

End If

' Graba los casinos Actualizados

Dim CodigoReceta As Long
Dim iselecc      As Long
Dim RS1          As New ADODB.Recordset
Dim Sql          As String
Dim desoferta    As String
        
 For i = 0 To List1(1).ListCount - 1
   
   If List1(1).Selected(i) = True Then
       
       iselecc = 1
   
   Else
       
       iselecc = 0
   
   End If
       
   desoferta = ""
   desoferta = List1(1).List(i)
   CodigoReceta = Val(fg_codigolistaNuevo(desoferta, 1, 10, ""))
      
   Sql = " sgpadm_iu_OfertasCasino "
   Sql = Sql & "'" & Trim(fpText(0).text) & "',"
   Sql = Sql & CodigoReceta & ","
   Sql = Sql & iselecc & ","
   Sql = Sql & "'" & UCase(vg_NUsr) & "'"
      
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
      
   Set RS1 = vg_db.Execute(Sql)

 Next i
 
 'Ini : Tipo Negocio
 
  desoferta = ""
  CodigoReceta = 0
  iselecc = 0
  
 For i = 0 To TipoNegocio(0).ListCount - 1
   
   If TipoNegocio(0).Selected(i) = True Then
       
       iselecc = 1
   
   Else
       
       iselecc = 0
   
   End If
       
   desoferta = ""
   desoferta = TipoNegocio(0).List(i)
   CodigoReceta = Val(fg_codigolistaNuevo(desoferta, 1, 10, ""))
      
   Sql = " sgpadm_iu_CasinoTipoNegocio "
   Sql = Sql & "'" & Trim(fpText(0).text) & "',"
   Sql = Sql & CodigoReceta & ","
   Sql = Sql & iselecc & ","
   Sql = Sql & "'" & UCase(vg_NUsr) & "'"
      
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
      
   Set RS1 = vg_db.Execute(Sql)

 Next i
 
 'Fin : Tipo Negocio

 'Ini : Sello
 Dim dessello As String
 Dim codsello As Long
 
 dessello = ""
 codsello = 0
 iselecc = 0
  
 For i = 0 To Sello(0).ListCount - 1
   
   If Sello(0).Selected(i) = True Then
       
       iselecc = 1
   
   Else
       
       iselecc = 0
   
   End If
       
   dessello = ""
   dessello = Sello(0).List(i)
   codsello = Val(fg_codigolistaNuevo(dessello, 1, 10, ""))
      
   Sql = " sgpadm_iu_CasinoSello "
   Sql = Sql & "'" & Trim(fpText(0).text) & "',"
   Sql = Sql & codsello & ","
   Sql = Sql & iselecc & ","
   Sql = Sql & "'" & UCase(vg_NUsr) & "'"
      
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
      
   Set RS1 = vg_db.Execute(Sql)

 Next i
 
 'Fin : Sello

Est = False
modo = ""
Gl_Ac_Botones Me, 13, IIf(lc_Aux = "MCasino", 1, 3), modo
Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub tvwDir_Expand(ByVal Node As MSComctlLib.Node)

On Error GoTo Man_Error

Dim RS    As New ADODB.Recordset
Dim dest1 As Node
Dim i     As Long

Set dest1 = Node
If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
If Node.Child.text <> "*" Then Exit Sub
' eliminar el elemento hijo positivo
TvwDir.Nodes.Remove Node.Child.Index

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_buscarsercasino '" & codigo & "', " & Val(Mid(TvwDir.Nodes(dest1.Index).key, 2, 10)) & "")
If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      Set nd = TvwDir.Nodes.Add(dest1.Index, tvwChild, "H" & fg_pone_espacio(RS!Ser_codigo, 10) & fg_pone_espacio(Val(Mid(TvwDir.Nodes(dest1.Index).key, 2, 10)), 10), RS!Ser_codigo & " - " & Trim(RS!ser_nombre))
      TvwDir.Nodes.item(nd.Index).Checked = IIf(RS!crs_codser = 1, True, False)
      RS.MoveNext: i = i + 1
   
   Loop

End If
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub TvwDir_NodeCheck(ByVal Node As MSComctlLib.Node)

On Error GoTo Man_Error

Dim lCheck As Boolean, lCheck1 As Boolean
Dim itesel As Node
Dim i As Long, j As Long
Dim cKey As String
fg_carga ""
If Mid(ValidarUsuario(Me), 2, 1) <> "1" Then
   
   fg_descarga
   Exit Sub

End If
TvwDir.Nodes.item(Node.key).Selected = True
Set itesel = TvwDir.SelectedItem
tvwDir_Expand itesel
TvwDir.Nodes.item(Node.key).Selected = True
lCheck = TvwDir.Nodes.item(TvwDir.SelectedItem.Index).Checked
lCheck1 = TvwDir.Nodes.item(TvwDir.SelectedItem.Index).Checked
cKey = Trim(TvwDir.Nodes.item(TvwDir.SelectedItem.Index).key)
If TvwDir.SelectedItem.Children > 0 Then
   
   For i = TvwDir.SelectedItem.Index + 1 To TvwDir.Nodes.count

       If Val(Mid(cKey, 2, 11)) = Val(Mid(TvwDir.Nodes.item(i).key, 12, 21)) Then TvwDir.Nodes.item(i).Checked = lCheck1
   
   Next i

Else
   
   For i = 1 To TvwDir.Nodes.count
       
       If TvwDir.Nodes.item(i).Children = 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDir.Nodes.item(i).key, 12, 21)) Then
          
          j = i: Exit For
       
       End If
   
   Next i

   For i = j To TvwDir.Nodes.count
       
       If TvwDir.Nodes.item(i).Children > 0 Then Exit For
       If TvwDir.Nodes.item(i).Checked = True Then lCheck1 = True 'TvwDir.Nodes.Item(i).Checked: Exit For
   
   Next i
   
   For i = (TvwDir.SelectedItem.Index - 1) To 1 Step -1

       If TvwDir.Nodes.item(i).Children > 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDir.Nodes.item(i).key, 2, 11)) Then
          
          TvwDir.Nodes.item(i).Checked = lCheck1
          Exit For
       
       ElseIf TvwDir.Nodes.item(i).Checked = True And TvwDir.Nodes.item(i).Children > 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDir.Nodes.item(i).key, 2, 11)) Then
          
          Exit For
       
       End If
   
   Next i

End If
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(6) = False
SSTab1.TabEnabled(7) = False
SSTab1.TabEnabled(8) = False

Gl_Ac_Botones Me, 13, 0, modo
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = Row: vaSpread1.Col = 1: codigo = vaSpread1.text

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Sub Limpia()

On Error GoTo Man_Error

Dim i As Long
Est = True
For i = 0 To 14
   
   fpText(i).text = ""
   If i < 7 Then fpLongInteger1(i).text = ""
   If i < 9 Then fpayuda(i).Caption = ""
   If i < 3 Then Check1(i).Value = 0

Next i

fpText(17).text = ""

Est = False

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub vaSpread2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

If vaSpread2.MaxRows < 1 Or Est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 13, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(6) = False
SSTab1.TabEnabled(7) = False
SSTab1.TabEnabled(8) = False

SSTab1.Tab = 3
SSTab1.TabEnabled(3) = True

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub vaSpread2_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If vaSpread2.MaxRows < 1 Then Exit Sub
If Col = 1 And Row = 0 Then

  vaSpread2.Row = -1
  vaSpread2.Col = 1
  vaSpread2.text = IIf(vaSpread2.Value = "1", "0", "1")

End If

If Col = 1 Then Exit Sub

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub vaSpread4_Change(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If vaSpread4.MaxRows < 1 Then Exit Sub
If modo = "" Then modo = "M"
If modo = "M" And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = False: SSTab1.TabEnabled(0) = False

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub vaSpread4_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    
On Error GoTo Man_Error

    With vaSpread4
        
        If .MaxRows < 1 Then Exit Sub
        If modo = "" Then modo = "M"
        If modo = "M" And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo: SSTab1.TabEnabled(0) = False: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = False
        Dim tipdes As String
        .Row = Row
        
        Select Case Col
        
        Case 4
            
            .Col = 4: tipdes = .TypeComboBoxCurSel
            .Col = 5: .TypeComboBoxCurSel = tipdes
            .EditEnterAction = EditEnterActionNone
            estdes = True
            
            If .text <> "S" And .text <> "E" Then
               
               .Col = 6: .CellType = CellTypeStaticText: .text = ""
               .Col = 7: .CellType = CellTypeStaticText: .text = ""
               .Col = 8: .CellType = CellTypeStaticText: .text = ""
               .Col = 9: .CellType = CellTypeStaticText: .text = ""
               .Col = 10: .CellType = CellTypeStaticText: .text = ""
               .Col = 11: .CellType = CellTypeStaticText: .text = ""
               .Col = 12: .CellType = CellTypeStaticText: .text = ""
               .Col = 13: .CellType = CellTypeStaticText: .text = ""
               .Col = 14: .CellType = CellTypeStaticText: .text = ""
               '-------> Mover datos segunda quincena
               .Col = 6
               
               If tipdes = "1" Then
                  
                  .Col = 6: .CellType = CellTypeComboBox: .TypeComboBoxList = "QUINCENAL 1-15" & Chr$(9) & "QUINCENAL 2-16" & Chr$(9) & "QUINCENAL 3-17" & Chr$(9) & "QUINCENAL 4-18"
                  .Col = 7: .CellType = CellTypeComboBox: .TypeComboBoxList = "Q1" & Chr$(9) & "Q2" & Chr$(9) & "Q3" & Chr$(9) & "Q4"
               
               End If
            
            Else
               
               .Col = 6: .CellType = CellTypeStaticText: .text = ""
               .Col = 7: .CellType = CellTypeStaticText: .text = ""
               .Col = 8: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
               .Col = 9: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
               .Col = 10: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
               .Col = 11: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
               .Col = 12: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
               .Col = 13: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
               .Col = 14: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
            
            End If
            estdes = False
        
        Case 6
            
            .Col = 6: tipdes = .TypeComboBoxCurSel
            .Col = 7: .TypeComboBoxCurSel = tipdes
            .EditEnterAction = EditEnterActionNone
        
        End Select
    
    End With

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub vaSpread4_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

If (Col <> 8 And Col <> 9 And Col <> 10 And Col <> 11 And Col <> 12 And Col <> 13 And Col <> 14) Or Row = 0 Or estdes Then Exit Sub
vaSpread4.Row = Row
vaSpread4.Col = 4
If vaSpread4.TypeComboBoxCurSel = 2 Then
   
   For i = 8 To 14
       If i <> Col Then estdes = True: vaSpread4.Col = i: vaSpread4.text = "0": estdes = False
   Next i

End If

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub vaSpread5_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
        
On Error GoTo Man_Error

        If (Col <> 3) Or Row = 0 Or estact Then Exit Sub
        If modo = "" Then modo = "M"
        Gl_Ac_Botones Me, 1, 0, modo
    
    With SSTab1
        
        .TabEnabled(0) = False
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .TabEnabled(3) = False
        .TabEnabled(4) = False
        .TabEnabled(5) = False
        .TabEnabled(7) = False
        .TabEnabled(8) = False
        .TabEnabled(9) = False
        .TabEnabled(10) = False
        .TabEnabled(11) = False
        .TabVisible(12) = False
   
    End With

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Sub ActivarBotones()

On Error GoTo Man_Error

If lc_Aux = "MCasino" Then
   
   Gl_Ac_Botones Me, 13, 1, modo

Else
   
   Gl_Ac_Botones Me, 13, 3, modo

End If

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub vaSpread6_Change(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim Atributo As String
Dim PosInicial As Long
Dim Largo As Long
Dim i As Long
If vaSpread6.MaxRows < 1 Then Exit Sub

If modo = "" Then modo = "M"
If modo = "M" And Toolbar1.Buttons(12).Visible = False Then
    
    Gl_Ac_Botones Me, 1, 0, modo
    With SSTab1
        
        .TabEnabled(0) = False
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .TabEnabled(3) = False
        .TabEnabled(4) = False
        .TabEnabled(5) = False
        .TabEnabled(6) = False
        .TabEnabled(7) = False
        .TabEnabled(8) = True
        .TabEnabled(9) = False
        .TabEnabled(10) = False
        .TabEnabled(11) = False
        .TabVisible(12) = False

    End With

End If

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub vaSpread6_ComboSelChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Select Case Col

Case 1
    
    Dim indice As Long
    Dim CodVal As String
    vaSpread6.Row = Row
    vaSpread6.Col = 1
    indice = vaSpread6.TypeComboBoxCurSel
    vaSpread6.Col = 2
    vaSpread6.TypeComboBoxCurSel = indice
    CodVal = ""
    CodVal = vaSpread6.text
    vaSpread6.Col = 5
    If vaSpread6.text <> "Seleccione..." Then vaSpread6.text = CodVal
    Gl_Ac_Botones Me, 1, 0, modo

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub vaSpread6_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If vaSpread6.MaxRows < 1 Then Exit Sub
Select Case KeyCode
Case 46
    
    vaSpread6.Row = vaSpread6.ActiveRow
    vaSpread6.Col = vaSpread6.ActiveCol
    If vaSpread6.Col <> 4 Then Exit Sub
    vaSpread6.text = ""
    vaSpread6.TypeComboBoxCurSel = -1
    vaSpread6.Col = 5
    vaSpread6.text = ""
    If modo = "" Then modo = "M"
    If Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub vaSpread7_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

If vaSpread7.MaxRows < 1 Or Est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 13, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(6) = False
SSTab1.TabEnabled(7) = False

SSTab1.Tab = 9
SSTab1.TabEnabled(9) = True

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub vaSpread7_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If vaSpread7.MaxRows < 1 Then Exit Sub
If Col = 6 And Row = 0 Then

  vaSpread7.Row = -1
  vaSpread7.Col = 6
  vaSpread7.text = IIf(vaSpread7.Value = "1", "0", "1")

End If

If Col = 6 Then Exit Sub

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Sub DehabilitarOpciones()

On Error GoTo Man_Error

itab = 1
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(6) = False
SSTab1.TabEnabled(7) = False
SSTab1.TabEnabled(8) = False
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Private Sub TvwDietetica_Expand_2(Index As Integer, ByVal Node As MSComctlLib.Node, Ceco As String)

On Error GoTo Man_Error

Dim RS1       As New ADODB.Recordset
Dim RS2       As New ADODB.Recordset
Dim estnivel2 As Boolean

estnivel2 = True
Set dest = Node
Nivel3 = 0

Select Case Index

Case 0
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" Then Exit Sub
    
    ' eliminar el elemento hijo positivo
       
       TvwDietetica(0).Nodes.Remove Node.Child.Index
       
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Nivel3 = 0
       Nivel2 = Val(Mid(TvwDietetica(0).Nodes(dest.Index).key, 2, 10))
       
       Set RS1 = vg_db.Execute("sgpadm_Sel_CasinoParamCategoriaDieteticaOtrosNiveles_V01 " & Nivel2 & ", '2', '" & Ceco & "'")
       
       If Not RS1.EOF Then
          
          Do While Not RS1.EOF
                                      
             Set Nod2 = TvwDietetica(0).Nodes.Add(Nodx, tvwChild, "H" & RS1!car_codigo & fg_pone_espacio(Nivel2, 10), RS1!car_codigo & " - " & Trim(RS1!car_nombre))
             
             TvwDietetica(0).Nodes.item(Nod2.Index).Checked = IIf(RS1!MARCA = "1", True, False)
             
             If Nod2.Children = 0 Then
                
                If RS2.State = 1 Then RS2.Close
                RS2.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS2 = vg_db.Execute("sgpadm_Sel_CasinoParamCategoriaDieteticaOtrosNiveles_V01 " & RS1!car_codigo & ", '1', '" & Ceco & "'")
                
                If Not RS2.EOF Then
                   
                Nivel3 = RS1!car_codigo
   
                ' la propiedad Texto de los nodos positivos es "***"
                TvwDietetica(0).Nodes.item(TvwDietetica(0).Nodes.count).Selected = True
                TvwDietetica(0).Nodes.Add Nod2.Index, tvwChild, , "**"
                
                Set nd = TvwDietetica(0).SelectedItem
                   
                Set nd1 = TvwDietetica(0).SelectedItem
                TvwDietetica_Expand_3 0, nd1, Ceco
                estnivel2 = False
                
                End If
                
                RS2.Close
                Set RS2 = Nothing
                
             End If
                       
             RS1.MoveNext
          
          Loop
       
       End If
       
       RS1.Close
       Set RS1 = Nothing

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub TvwDietetica_Expand_3(Index As Integer, ByVal Node As MSComctlLib.Node, Ceco As String)

On Error GoTo Man_Error

Dim RS5       As New ADODB.Recordset
Dim RS6       As New ADODB.Recordset
Dim estnivel3 As Boolean

estnivel3 = True
Nivel4 = 0
Set dest = Node

Select Case Index

Case 0
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" Then Exit Sub

    ' eliminar el elemento hijo positivo
       TvwDietetica(0).Nodes.Remove Node.Child.Index
       
       If RS5.State = 1 Then RS5.Close
       RS5.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS5 = vg_db.Execute("sgpadm_Sel_CasinoParamCategoriaDieteticaOtrosNiveles_V01  " & Nivel3 & ", '1', '" & Ceco & "'")
       
       If Not RS5.EOF Then
          
          Do While Not RS5.EOF
             
             Set Nod3 = TvwDietetica(0).Nodes.Add(Nod2, tvwChild, "H" & RS5!car_codigo & fg_pone_espacio(Nivel3, 10), RS5!car_codigo & " - " & Trim(RS5!car_nombre))
             
             TvwDietetica(0).Nodes.item(Nod3.Index).Checked = IIf(RS5!MARCA = "1", True, False)

             If Nod3.Children = 0 Then
                
                If RS6.State = 1 Then RS6.Close
                RS6.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS6 = vg_db.Execute("sgpadm_Sel_CasinoParamCategoriaDieteticaOtrosNiveles_V01 " & RS5!car_codigo & ", '1', '" & Ceco & "'")
                
                If Not RS6.EOF Then
                   
                   Nivel4 = RS5!car_codigo
                
                   ' la propiedad Texto de los nodos positivos es "***"
                   TvwDietetica(0).Nodes.item(TvwDietetica(0).Nodes.count).Selected = True
                   TvwDietetica(0).Nodes.Add Nod3.Index, tvwChild, , "****"
                
                   Set nd = TvwDietetica(0).SelectedItem
                
                   Set ndl = TvwDietetica(0).SelectedItem
                   TvwDietetica_Expand_4 0, ndl, Ceco 'dest
                   estnivel3 = False
                
                End If
                
                RS6.Close
                Set RS6 = Nothing
                
             End If
             
             RS5.MoveNext
          
          Loop
       
       RS5.Close
       Set RS5 = Nothing
    
    End If

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub TvwDietetica_Expand_4(Index As Integer, ByVal Node As MSComctlLib.Node, Ceco As String)

On Error GoTo Man_Error

Dim RS7       As New ADODB.Recordset
Dim RS8       As New ADODB.Recordset
Dim estnivel4 As Boolean

Set dest = Node
Select Case Index

Case 0
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" And Node.Child.text <> "****" Then Exit Sub

    ' eliminar el elemento hijo positivo
       TvwDietetica(0).Nodes.Remove Node.Child.Index
       
       If RS7.State = 1 Then RS7.Close
       RS7.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS7 = vg_db.Execute("sgpadm_Sel_CasinoParamCategoriaDieteticaOtrosNiveles_V01 " & Nivel4 & ", '1', '" & Ceco & "'")
       
       If Not RS7.EOF Then
          
          Do While Not RS7.EOF
             
             Set Nod4 = TvwDietetica(0).Nodes.Add(Nod3, tvwChild, "H" & RS7!car_codigo & fg_pone_espacio(Val(Nivel4), 10), RS7!car_codigo & " - " & Trim(RS7!car_nombre))
             
             TvwDietetica(0).Nodes.item(Nod4.Index).Checked = IIf(RS7!MARCA = "1", True, False)
             
             If Nod4.Children = 0 Then
                
                If RS8.State = 1 Then RS8.Close
                RS8.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS8 = vg_db.Execute("sgpadm_Sel_CasinoParamCategoriaDieteticaOtrosNiveles_V01 " & RS7!car_codigo & ", '1', '" & Ceco & "'")
                
                If Not RS8.EOF Then
                   
                   Nivel5 = RS7!car_codigo
                   
                   ' la propiedad Texto de los nodos positivos es "*****"
                   TvwDietetica(0).Nodes.item(TvwDietetica(0).Nodes.count).Selected = True
                   TvwDietetica(0).Nodes.Add Nod4.Index, tvwChild, , "*****"
                
                   Set ndl = TvwDietetica(0).SelectedItem
                   TvwDietetica_Expand_5 0, ndl, Ceco 'dest
                   estnivel4 = False
                
                End If
                
                RS8.Close
                Set RS8 = Nothing
                
             End If
             
             RS7.MoveNext
          
          Loop
       
       End If
       
       RS7.Close
       Set RS7 = Nothing
    
End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub TvwDietetica_Expand_5(Index As Integer, ByVal Node As MSComctlLib.Node, Ceco As String)

On Error GoTo Man_Error

Dim RS9        As New ADODB.Recordset
Dim RS10       As New ADODB.Recordset
Dim estnivel5  As Boolean

Set dest = Node
Select Case Index

Case 0
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" And Node.Child.text <> "****" And Node.Child.text <> "*****" Then Exit Sub

    ' eliminar el elemento hijo positivo
       TvwDietetica(0).Nodes.Remove Node.Child.Index
       
       If RS9.State = 1 Then RS9.Close
       RS9.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS9 = vg_db.Execute("sgpadm_Sel_CasinoParamCategoriaDieteticaOtrosNiveles_V01 " & Nivel5 & ", '1', '" & Ceco & "'")
       
       If Not RS9.EOF Then
          
          Do While Not RS9.EOF
             
              Set Nod5 = TvwDietetica(0).Nodes.Add(Nod4, tvwChild, "H" & RS9!car_codigo & fg_pone_espacio(Val(Nivel5), 10), RS9!car_codigo & " - " & Trim(RS9!car_nombre))
             
             TvwDietetica(0).Nodes.item(Nod5.Index).Checked = IIf(RS9!MARCA = "1", True, False)
             
             If Nod5.Children = 0 Then
                
                If RS10.State = 1 Then RS10.Close
                RS10.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS10 = vg_db.Execute("sgpadm_Sel_CasinoParamCategoriaDieteticaOtrosNiveles_V01 " & RS9!car_codigo & ", '1', '" & Ceco & "'")
                
                If Not RS10.EOF Then
                   
                   Nivel6 = RS9!car_codigo
                   
                   ' la propiedad Texto de los nodos positivos es "*****"
                   TvwDietetica(0).Nodes.item(TvwDietetica(0).Nodes.count).Selected = True
                   TvwDietetica(0).Nodes.Add Nod5.Index, tvwChild, , "******"
                
                   Set ndl = TvwDietetica(0).SelectedItem
                   TvwDietetica_Expand_6 0, ndl, Ceco  'dest
                   estnivel5 = False
                
                End If
                
                RS10.Close
                Set RS10 = Nothing
                
             End If
             
             RS9.MoveNext
          
          Loop
       
       End If
       
       RS9.Close
       Set RS9 = Nothing
    
End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub TvwDietetica_Expand_6(Index As Integer, ByVal Node As MSComctlLib.Node, Ceco As String)

On Error GoTo Man_Error

Dim RS11       As New ADODB.Recordset
Dim RS12       As New ADODB.Recordset
Dim estnivel6  As Boolean

Set dest = Node
Select Case Index

Case 0
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" And Node.Child.text <> "****" And Node.Child.text <> "*****" And Node.Child.text <> "******" Then Exit Sub

    ' eliminar el elemento hijo positivo
       TvwDietetica(0).Nodes.Remove Node.Child.Index
       
       If RS11.State = 1 Then RS11.Close
       RS11.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS11 = vg_db.Execute("sgpadm_Sel_CasinoParamCategoriaDieteticaOtrosNiveles_V01 " & Nivel6 & ", '1', '" & Ceco & "'")
       
       If Not RS11.EOF Then
          
          Do While Not RS11.EOF
             
             Set Nod6 = TvwDietetica(0).Nodes.Add(Nod5, tvwChild, "H" & RS11!car_codigo & fg_pone_espacio(Val(Nivel6), 10), RS11!car_codigo & " - " & Trim(RS11!car_nombre))
             
             TvwDietetica(0).Nodes.item(Nod6.Index).Checked = IIf(RS11!MARCA = "1", True, False)
             
             If Nod6.Children = 0 Then
                
                If RS12.State = 1 Then RS12.Close
                RS12.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS12 = vg_db.Execute("sgpadm_Sel_CasinoParamCategoriaDieteticaOtrosNiveles_V01 " & RS11!car_codigo & ", '1', '" & Ceco & "'")
                
                If Not RS12.EOF Then
                   
                   ' la propiedad Texto de los nodos positivos es "******"
                   TvwDietetica(0).Nodes.item(TvwDietetica(0).Nodes.count).Selected = True
                   TvwDietetica(0).Nodes.Add Nod6.Index, tvwChild, , "******"
                
                   Set nd = TvwDietetica(0).SelectedItem
                   TvwDietetica_Expand_6 0, dest, Ceco
                   estnivel6 = False
                
                End If
                
                RS12.Close
                Set RS12 = Nothing
                
             End If
             
             RS11.MoveNext
          
          Loop
       
       End If
       
       RS11.Close
       Set RS11 = Nothing

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub TvwDietetica_NodeCheck(Index As Integer, ByVal Node As MSComctlLib.Node)

On Error GoTo Man_Error

Dim lCheck        As Boolean
Dim lCheck1       As Boolean
Dim itesel        As Node
Dim i             As Long
Dim j             As Long
Dim p             As Long
Dim cKey          As String
Dim cKey2         As String
Dim cKey3         As String
Dim cKey4         As String
Dim cKey5         As String

Dim cKeyFullPath  As String

fg_carga ""

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 13, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(6) = False
SSTab1.TabEnabled(7) = False

SSTab1.Tab = 12
SSTab1.TabEnabled(12) = True

cKey2 = ""
cKey3 = ""
cKey4 = ""
cKey5 = ""
cKey = ""
cKeyFullPath = ""

TvwDietetica(Index).Nodes.item(Node.key).Selected = True
Set itesel = TvwDietetica(Index).SelectedItem
TvwDietetica(Index).Nodes.item(Node.key).Selected = True
lCheck = TvwDietetica(Index).Nodes.item(TvwDietetica(Index).SelectedItem.Index).Checked
lCheck1 = TvwDietetica(Index).Nodes.item(TvwDietetica(Index).SelectedItem.Index).Checked
cKey = Trim(TvwDietetica(Index).Nodes.item(TvwDietetica(Index).SelectedItem.Index).key)
cKeyFullPath = TvwDietetica(Index).Nodes.item(TvwDietetica(Index).SelectedItem.Index).text '.Parent '.Root '.LastSibling '.Children '.FirstSibling '.FullPath

If TvwDietetica(Index).SelectedItem.Children > 0 Then
   
   For i = TvwDietetica(Index).SelectedItem.Index + 1 To TvwDietetica(Index).Nodes.count
      
       If Mid(TvwDietetica(Index).Nodes.item(i).key, 1, 1) = "R" Then
       
          Exit For
       
       End If
       
       If (Val(Mid(cKey, 2, 11)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 12, 21))) Or _
          (ValidarKey(cKey3, Mid(TvwDietetica(Index).Nodes.item(i).key, 12, 21), ",") _
          And Val(Mid(cKey2, 2, 11)) > 0) Then


'       (Val(Mid(cKey2, 2, 10)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).Key, 12, 21)) Or _
'       And Val(Mid(cKey2, 2, 11)) > 0) Then
       
          TvwDietetica(Index).Nodes.item(i).Checked = lCheck1
                   
             If TvwDietetica(Index).Nodes.item(i).Children > 0 Then
               
                cKey2 = Trim(TvwDietetica(Index).Nodes.item(i).key)
                cKey3 = Trim(Mid(Trim(TvwDietetica(Index).Nodes.item(i).key), 2, 10)) & "," & cKey3

                
             End If
       
       ElseIf Val(Mid(cKey, 2, 10)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 11, 22)) And TvwDietetica(Index).Nodes.item(i).Children = 0 Then
       
          TvwDietetica(Index).Nodes.item(i).Checked = lCheck1
          
       ElseIf ValidarKey(TvwDietetica(Index).Nodes.item(i).FullPath, cKeyFullPath, "\") Then
       
          TvwDietetica(Index).Nodes.item(i).Checked = lCheck1
       
       End If
       
   Next i


   For i = 1 To TvwDietetica(Index).Nodes.count
       
       If TvwDietetica(Index).Nodes.item(i).Children = 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 12, 21)) Then
          
          j = i
          Exit For
       
       ElseIf Val(Mid(cKey, 2, 10)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 11, 22)) And TvwDietetica(Index).Nodes.item(i).Children = 0 Then
          
          j = i
          Exit For
       
       End If
   
   Next i

'   lCheck1 = False
   
   If j > 0 Then
      
      For i = j To TvwDietetica(Index).Nodes.count
       
         If TvwDietetica(Index).Nodes.item(i).Checked = True And Val(Mid(cKey, 11, 10)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 11, 10)) Then
          
          lCheck1 = True
          Exit For
   
         End If
     
         If TvwDietetica(Index).Nodes.item(i).Children > 0 Then
       
            Exit For
        
         End If
       
      Next i
   
   End If
   
   Dim lCheck2 As Boolean
   lCheck2 = False
   
   For i = (TvwDietetica(Index).SelectedItem.Index - 1) To 1 Step -1

       cKey2 = Trim(TvwDietetica(Index).Nodes.item(i).key)
       If TvwDietetica(Index).Nodes.item(i).Children > 0 And Val(Mid(cKey, 22, 30)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 2, 11)) Then
          
          TvwDietetica(Index).Nodes.item(i).Checked = lCheck1
''          Exit For
       
       ElseIf Val(Mid(cKey, 11, 22)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 2, 10)) And TvwDietetica(Index).Nodes.item(i).Children > 0 Then
        
          TvwDietetica(Index).Nodes.item(i).Checked = lCheck1
'          Exit For
          
'       ElseIf Index = 1 And (Val(Mid(cKey2, 2, 11)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).Key, 2, 10))) And CStr(Mid(TvwDietetica(Index).Nodes.item(i).Key, 1, 1)) = "R" Then
       ElseIf (Val(Mid(cKey2, 2, 11)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 2, 10))) And CStr(Mid(TvwDietetica(Index).Nodes.item(i).key, 1, 1)) = "R" Then
          
          For p = i + 1 To TvwDietetica(Index).Nodes.count

              If (Val(Mid(cKey, 2, 11)) = Val(Mid(TvwDietetica(Index).Nodes.item(p).key, 12, 21))) Or (Val(Mid(cKey2, 2, 11)) = Val(Mid(TvwDietetica(Index).Nodes.item(p).key, 12, 21)) And Val(Mid(cKey2, 2, 11)) > 0) Then
       
                 If TvwDietetica(Index).Nodes.item(p).Checked <> lCheck1 Then
                     
                        lCheck2 = TvwDietetica(Index).Nodes.item(p).Checked
                        
                 End If
                     
                 If TvwDietetica(Index).Nodes.item(p).Children > 0 Then

                    cKey2 = Trim(TvwDietetica(Index).Nodes.item(p).key)

                 End If

              End If
             
          Next p
             
          If (Val(Mid(cKey, 2, 11)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 2, 11))) Then
             
             TvwDietetica(Index).Nodes.item(i).Checked = IIf(Not lCheck2, lCheck1, lCheck2)
          
          End If
          
          Exit For
          
       ElseIf TvwDietetica(Index).Nodes.item(i).Checked = True And TvwDietetica(Index).Nodes.item(i).Children > 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 2, 11)) Then
          
          Exit For
       
       ElseIf (TvwDietetica(Index).Nodes.item(i).Checked = True Or TvwDietetica(Index).Nodes.item(i).Children > 0) Then
          
       
       End If
   
   Next i

Else
   
   For i = 1 To TvwDietetica(Index).Nodes.count
       
       If TvwDietetica(Index).Nodes.item(i).Children = 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 12, 21)) Then
          
          j = i
          Exit For
       
       ElseIf Val(Mid(cKey, 2, 10)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 11, 22)) And TvwDietetica(Index).Nodes.item(i).Children = 0 Then
          
          j = i
          Exit For
       
       End If
   
   Next i

   For i = j To TvwDietetica(Index).Nodes.count
       
         If TvwDietetica(Index).Nodes.item(i).Checked = True And Val(Mid(cKey, 11, 10)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 11, 10)) Then
          
          lCheck1 = True
          Exit For
   
       End If
     
       If TvwDietetica(Index).Nodes.item(i).Children > 0 Then
       
          Exit For
        
       End If
       
       
   Next i
'   Dim lCheck2 As Boolean
   lCheck2 = False
   
   For i = (TvwDietetica(Index).SelectedItem.Index - 1) To 1 Step -1

       cKey2 = Trim(TvwDietetica(Index).Nodes.item(i).key)
       If TvwDietetica(Index).Nodes.item(i).Children > 0 And Val(Mid(cKey, 22, 30)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 2, 11)) Then
          
          TvwDietetica(Index).Nodes.item(i).Checked = lCheck1
       
       ElseIf Val(Mid(cKey, 11, 22)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 2, 10)) And TvwDietetica(Index).Nodes.item(i).Children > 0 Then
        
          TvwDietetica(Index).Nodes.item(i).Checked = lCheck1
       
       ElseIf (Val(Mid(cKey2, 2, 11)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 2, 10))) And CStr(Mid(TvwDietetica(Index).Nodes.item(i).key, 1, 1)) = "R" Then
          
          For p = i + 1 To TvwDietetica(Index).Nodes.count

              If (Val(Mid(cKey, 2, 11)) = Val(Mid(TvwDietetica(Index).Nodes.item(p).key, 12, 21))) Or (Val(Mid(cKey2, 2, 11)) = Val(Mid(TvwDietetica(Index).Nodes.item(p).key, 12, 21)) And Val(Mid(cKey2, 2, 11)) > 0) Then
       
                 If TvwDietetica(Index).Nodes.item(p).Checked <> lCheck1 Then
                     
                        lCheck2 = TvwDietetica(Index).Nodes.item(p).Checked
                        
                 End If
                     
                 If TvwDietetica(Index).Nodes.item(p).Children > 0 Then

                    cKey2 = Trim(TvwDietetica(Index).Nodes.item(p).key)

                 End If

              End If
             
          Next p
             
          If (Val(Mid(cKey, 2, 11)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 2, 11))) Then
          
              TvwDietetica(Index).Nodes.item(i).Checked = IIf(Not lCheck2, lCheck1, lCheck2)
          
          End If
          
          Exit For
          
       ElseIf TvwDietetica(Index).Nodes.item(i).Checked = True And TvwDietetica(Index).Nodes.item(i).Children > 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDietetica(Index).Nodes.item(i).key, 2, 11)) Then
          
          Exit For
       
       ElseIf (TvwDietetica(Index).Nodes.item(i).Checked = True Or TvwDietetica(Index).Nodes.item(i).Children > 0) Then
          
       
       End If
   
   Next i

End If
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Function ValidarKey(ArregloKey As String, key As String, Caracter As String) As Boolean

Dim Ind           As Long
Dim cKeyArreglo() As String

ValidarKey = False
If Trim(ArregloKey) <> "" Then
cKeyArreglo = Split(ArregloKey, Caracter)

For Ind = 0 To UBound(cKeyArreglo)

    If cKeyArreglo(Ind) = Trim(key) Then
    
       ValidarKey = True
       Exit For
       
    End If

Next

End If
End Function


Private Sub TvwDietetica_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 27
    
    vg_opcion = 2
    Me.Hide

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub TvwDietetica_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)

On Error GoTo Man_Error

Dim CodCatDie As Long
Dim nomcatdie As String

Set dest = Node

Select Case Index

Case 0
     
     CodCatDie = Val(Mid(TvwDietetica(0).Nodes(dest.Index).key, 2, 20))
     nomcatdie = Trim((TvwDietetica(0).Nodes(dest.Index).text))

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub




