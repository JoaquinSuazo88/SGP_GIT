VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4DBFB8CD-9EF9-11D0-8BC4-00AA00B42B7C}#3.0#0"; "Cal32x30.ocx"
Begin VB.Form M_Ruta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor Ruta Despacho"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Listar Ruta"
      TabPicture(0)   =   "M_Ruta.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "vaSpread1"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Ruta"
      TabPicture(1)   =   "M_Ruta.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Productos"
      TabPicture(2)   =   "M_Ruta.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(3)=   "Frame6"
      Tab(2).Control(4)=   "Frame7"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Calendario"
      TabPicture(3)   =   "M_Ruta.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame8"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Casino"
      TabPicture(4)   =   "M_Ruta.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame9"
      Tab(4).ControlCount=   1
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
         Height          =   7695
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   10695
         Begin VB.Frame Frame11 
            Caption         =   "Casino No Incluido"
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
            Left            =   5520
            TabIndex        =   29
            Top             =   360
            Width           =   5055
            Begin VB.Frame Frame15 
               Height          =   435
               Left            =   480
               TabIndex        =   38
               Top             =   6600
               Width           =   915
               Begin VB.TextBox TextCan1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   2
                  Left            =   45
                  TabIndex        =   39
                  Top             =   135
                  Width           =   810
               End
            End
            Begin VB.Frame Frame14 
               Height          =   435
               Left            =   1410
               TabIndex        =   36
               Top             =   6600
               Width           =   3285
               Begin VB.TextBox TextCan1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   3
                  Left            =   45
                  TabIndex        =   37
                  Top             =   135
                  Width           =   3180
               End
            End
            Begin FPSpread.vaSpread vaSpread5 
               Height          =   6255
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   4815
               _Version        =   393216
               _ExtentX        =   8493
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
               SpreadDesigner  =   "M_Ruta.frx":008C
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Casino Incluido"
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
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   5055
            Begin VB.Frame Frame13 
               Height          =   435
               Left            =   480
               TabIndex        =   34
               Top             =   6600
               Width           =   915
               Begin VB.TextBox TextCai1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   2
                  Left            =   45
                  TabIndex        =   35
                  Top             =   135
                  Width           =   810
               End
            End
            Begin VB.Frame Frame12 
               Height          =   435
               Left            =   1410
               TabIndex        =   32
               Top             =   6600
               Width           =   3285
               Begin VB.TextBox TextCai1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   3
                  Left            =   45
                  TabIndex        =   33
                  Top             =   135
                  Width           =   3180
               End
            End
            Begin FPSpread.vaSpread vaSpread4 
               Height          =   6255
               Left            =   120
               TabIndex        =   30
               Top             =   240
               Width           =   4815
               _Version        =   393216
               _ExtentX        =   8493
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
               SpreadDesigner  =   "M_Ruta.frx":1A09
            End
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
         Height          =   7695
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   10695
         Begin FPSpread.vaSpread vaSpread3 
            Height          =   6885
            Left            =   3840
            TabIndex        =   25
            Top             =   600
            Width           =   6720
            _Version        =   393216
            _ExtentX        =   11853
            _ExtentY        =   12144
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
            MaxCols         =   5
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "M_Ruta.frx":3371
            ScrollBarTrack  =   3
            ClipboardOptions=   0
         End
         Begin CalObjXLib.fpCalendar fpCalendar1 
            Height          =   4830
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   3615
            _Version        =   196608
            _ExtentX        =   6376
            _ExtentY        =   8520
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
            DisplayFormat   =   2
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
            ElementTextStyle=   "M_Ruta.frx":4E3B
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
            MouseIcon       =   "M_Ruta.frx":50E4
         End
      End
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   -74880
         TabIndex        =   22
         Top             =   600
         Width           =   10815
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   6330
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   10470
            _Version        =   393216
            _ExtentX        =   18468
            _ExtentY        =   11165
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
            MaxCols         =   6
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "M_Ruta.frx":5100
            ScrollBarTrack  =   3
            ClipboardOptions=   0
         End
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Left            =   -69840
         TabIndex        =   20
         Top             =   7560
         Width           =   4125
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   45
            TabIndex        =   21
            Top             =   135
            Width           =   4020
         End
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   -72795
         TabIndex        =   18
         Top             =   7560
         Width           =   1245
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   19
            Top             =   135
            Width           =   1140
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   -74200
         TabIndex        =   16
         Top             =   7560
         Width           =   1395
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   17
            Top             =   135
            Width           =   1290
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   -71505
         TabIndex        =   14
         Top             =   7560
         Width           =   1605
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   15
            Top             =   135
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Left            =   -74760
         TabIndex        =   9
         Top             =   600
         Width           =   10575
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   3000
            TabIndex        =   12
            Top             =   1080
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
            Left            =   3000
            TabIndex        =   13
            Top             =   720
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
            Left            =   1320
            TabIndex        =   11
            Top             =   1200
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
            Left            =   1320
            TabIndex        =   10
            Top             =   720
            Width           =   600
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   971
         Left            =   -73080
         TabIndex        =   2
         Top             =   720
         Width           =   6015
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "M_Ruta.frx":6BF6
            Left            =   2010
            List            =   "M_Ruta.frx":6C00
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   2500
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   1
            Left            =   2010
            TabIndex        =   4
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
            TabIndex        =   7
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
            TabIndex        =   6
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
            TabIndex        =   5
            Top             =   345
            Width           =   1380
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6285
         Left            =   -73080
         TabIndex        =   8
         Top             =   1800
         Width           =   6030
         _Version        =   393216
         _ExtentX        =   10636
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
         SpreadDesigner  =   "M_Ruta.frx":6C14
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_Ruta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String, codigo As String, Msgtitulo As String
Dim Est As Boolean
Dim indgra As Long
Dim ano As String
Dim mes As String
Dim dia As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 9375
Me.Width = 11415
Msgtitulo = "Ruta"
fg_centra Me
SSTab1.Tab = 0
SSTab1.TabEnabled(4) = False
modo = ""
Est = True
Gl_Mo_Botones Me, 14
Gl_Ac_Botones Me, 14, 1, modo
Combo1.ListIndex = 0
ano = ""
mes = ""
dia = ""
MoverDatosGrilla
MoverDatosRuta
MoverDatosRutaProductos
MoverDatosRutaCalendario
Gl_Ac_Botones Me, 14, IIf(vaSpread1.MaxRows > 0, 1, 2), modo
End Sub

Sub MoverDatosGrilla()
fg_carga ""
Dim x As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 250
x = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)

vaSpread1.Visible = False
vaSpread1.Row = -1
vaSpread1.Col = -1
vaSpread1.Lock = True
vaSpread1.MaxRows = 0
Set RS = vg_dbpedweb.Execute("pedweb_s_ruta 4, '', ''")
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   
   vaSpread1.Col = 1
   vaSpread1.TypeHAlign = TypeHAlignRight
   vaSpread1.text = RS!recorrido
   
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
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
fg_descarga
End Sub

Sub MoverDatosRuta()
fg_carga ""
Dim ano As String
Est = True
Limpia 1
'-------> Cargar Ruta
Set RS = vg_dbpedweb.Execute("pedweb_s_ruta 5, " & codigo & ", ''")
If Not RS.EOF Then
   fpLongInteger1(0).Enabled = False
   fpLongInteger1(0).Value = RS!recorrido
   fpText1(0).text = Trim(RS!descripcion)
   Frame7.Caption = "(" & RS!recorrido & ") - " & Trim(RS!descripcion)
   Frame8.Caption = "(" & RS!recorrido & ") - " & Trim(RS!descripcion)
End If
RS.Close: Set RS = Nothing
Est = False
fg_descarga
End Sub

Sub MoverDatosRutaProductos()
fg_carga ""
Limpia 2
'-------> Ruta Productos
Dim x As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread2.TextTip = 2
    ' Control displays text tips after 250 milliseconds
vaSpread2.TextTipDelay = 250
' Text tip displays custom font and colors
' Background is yellow, RGB(255, 255, 0)
' Foreground is dark blue, RGB(0, 0, 128)
x = vaSpread2.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)

Set RS = vg_dbpedweb.Execute("pedweb_s_rutaproductos 1, " & codigo & ", '', ''")
Do While Not RS.EOF
   vaSpread2.MaxRows = vaSpread2.MaxRows + 1
   vaSpread2.Row = vaSpread2.MaxRows
   vaSpread2.Col = 1
   vaSpread2.text = "0"
   vaSpread2.Col = 2
   vaSpread2.text = IIf(IsNull(RS!producto), "", RS!producto)
   vaSpread2.Col = 3
   vaSpread2.text = IIf(IsNull(RS!pce_codcen), "", RS!pce_codcen)
   vaSpread2.Col = 4
   vaSpread2.text = IIf(IsNull(RS!familia), "", RS!familia)
   vaSpread2.Col = 5
   vaSpread2.text = IIf(IsNull(RS!descripcion), "", RS!descripcion)
   vaSpread2.Col = 6
   vaSpread2.CellType = CellTypeStaticText
   vaSpread2.TypeHAlign = TypeHAlignCenter
   vaSpread2.text = IIf(IsNull(RS!vigencia), "", RS!vigencia)
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
fg_descarga
End Sub

Sub MoverDatosRutaCalendario()
Dim i As Long, j As Long
fg_carga ""
Limpia 3
For i = 1 To 12
    For j = 1 To Val(Mid(dEoM("01/" & fg_pone_cero(i, 2) & "/" & fpCalendar1.Year), 1, 2))
        fpCalendar1.Element = ElementSpecificDate
        fpCalendar1.ElementIndex = fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
        fpCalendar1.ElementBackColor = -2147483633
        fpCalendar1.ElementText = ""
        fpCalendar1.ElementForeColor = vbBlack
        fpCalendar1.MultiSelect = MultiSelectNone
    Next j
Next i
'-------> Ruta Calendario
If Trim(ano) = "" Or Trim(mes) = "" Then
   Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendario 1, " & codigo & ", '', '', ''")
   If Not RS.EOF And Not IsNull(RS!FechaDespacho) Then
      ano = (Format(RS!FechaDespacho, "yyyy"))
      mes = (Format(RS!FechaDespacho, "mm"))
      dia = (Format(RS!FechaDespacho, "dd"))
   Else
      ano = Format(Date, "yyyy")
      mes = Format(Date, "mm")
      dia = Format(Date, "dd")
   End If
   RS.Close: Set RS = Nothing
End If
fpCalendar1.CurrentDate = Format(fg_pone_cero(dia, 2) & "/" & fg_pone_cero(mes, 2) & "/" & fg_pone_cero(ano, 4), "yyyymmdd")
Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendario 2, " & codigo & ", '" & ano & "', '" & mes & "', ''")
If RS.EOF Then SSTab1.TabEnabled(4) = False Else SSTab1.TabEnabled(4) = True
Do While Not RS.EOF
   vaSpread3.MaxRows = vaSpread3.MaxRows + 1
   vaSpread3.Row = vaSpread3.MaxRows
   vaSpread3.Col = 1
   vaSpread3.text = ""
   vaSpread3.Col = 2
   vaSpread3.text = RS!FechaDespacho
   vaSpread3.Col = 3
   vaSpread3.text = RS!fechaTopeIngreso
   vaSpread3.Col = 4
   vaSpread3.text = RS!DiasTopeAdicional
   vaSpread3.Col = 5
   vaSpread3.text = RS!HoraTopeAdicional
   fpCalendar1.Element = ElementSpecificDate
   fpCalendar1.ElementIndex = fg_pone_cero(Mid(RS!FechaDespacho, 7, 4), 4) & fg_pone_cero(Mid(RS!FechaDespacho, 4, 2), 2) & fg_pone_cero(Mid(RS!FechaDespacho, 1, 2), 2)
   fpCalendar1.ElementBackColor = &HDEFEDE
   fpCalendar1.ElementForeColor = vbBlack
   fpCalendar1.ElementText = Trim(vg_nombre)
   fpCalendar1.DrawFocusRect = AroundText '= CAL_DRAWFOCUSRECT_AROUND_TEXT
   
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
fpCalendar1.MultiSelect = MultiSelectExtended
fg_descarga
End Sub

Sub MoverDatosCasinos()
Dim FecDes As String
fg_carga ""
Dim x As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread4.TextTip = 2
' Control displays text tips after 250 milliseconds
vaSpread4.TextTipDelay = 250
' Text tip displays custom font and colors
' Background is yellow, RGB(255, 255, 0)
' Foreground is dark blue, RGB(0, 0, 128)
x = vaSpread4.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)

vaSpread3.Row = vaSpread3.ActiveRow
vaSpread3.Col = 2
FecDes = "": FecDes = Format(vaSpread3.text, "yyyymmdd")
Frame9.Caption = Frame8.Caption & " " & vaSpread3.text
'-------> Mover casino incluido en la ruta
vaSpread4.Visible = False
vaSpread4.MaxRows = 0
'-------> Bloquer grilla casino incluidos
vaSpread4.Row = -1
vaSpread4.Col = -1
vaSpread4.Lock = False
Set RS = vg_dbpedweb.Execute("pedweb_s_casinoincluidoruta 1, " & codigo & ", '" & FecDes & "'")
Do While Not RS.EOF
   vaSpread4.MaxRows = vaSpread4.MaxRows + 1
   vaSpread4.Row = vaSpread4.MaxRows
   vaSpread4.Col = 1
   vaSpread4.text = "0"
   vaSpread4.Col = 2
   vaSpread4.CellType = CellTypeStaticText
   vaSpread4.text = Trim(RS!centrocosto)
   vaSpread4.Col = 3
   vaSpread4.CellType = CellTypeStaticText
   vaSpread4.text = Trim(RS!Nombre)
   vaSpread4.Col = 4
   vaSpread4.CellType = CellTypeStaticText
   vaSpread4.text = Trim(RS!Casino)
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread4.Visible = True

' Control displays text tips aligned to pointer with focus
vaSpread5.TextTip = 2
' Control displays text tips after 250 milliseconds
vaSpread5.TextTipDelay = 250
' Text tip displays custom font and colors
' Background is yellow, RGB(255, 255, 0)
' Foreground is dark blue, RGB(0, 0, 128)
x = vaSpread5.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
'-------> mover casino no incluido en la ruta
vaSpread5.Visible = False
vaSpread5.MaxRows = 0
'-------> Bloquear grilla casino no incluido
vaSpread5.Row = -1
vaSpread5.Col = -1
vaSpread5.Lock = False
Set RS = vg_dbpedweb.Execute("pedweb_s_casinonoincluidoruta 1, " & codigo & ", '" & FecDes & "'")
Do While Not RS.EOF
   vaSpread5.MaxRows = vaSpread5.MaxRows + 1
   vaSpread5.Row = vaSpread5.MaxRows
   vaSpread5.Col = 1
   vaSpread5.text = "0"
   vaSpread5.Col = 2
   vaSpread5.CellType = CellTypeStaticText
   vaSpread5.text = Trim(RS!centrocosto)
   vaSpread5.Col = 3
   vaSpread5.CellType = CellTypeStaticText
   vaSpread5.text = Trim(RS!Nombre)
   vaSpread5.Col = 4
   vaSpread5.CellType = CellTypeStaticText
   vaSpread5.text = Trim(RS!codigo)
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread5.Visible = True
fg_descarga
End Sub

Sub Limpia(op As Integer)
Select Case op
Case 1
    fpLongInteger1(0).Value = ""
    fpLongInteger1(0).Enabled = False
    fpText1(0).text = ""
    Frame7.Caption = ""
    Frame8.Caption = ""
Case 2
    vaSpread2.MaxRows = 0
Case 3
    vaSpread3.MaxRows = 0
End Select
End Sub

Private Sub fpCalendar1_AfterSelection()
Dim i As Long
If vaSpread3.MaxRows < 1 Then Exit Sub
For i = 1 To vaSpread3.MaxRows
    vaSpread3.Row = i
    vaSpread3.Col = 2
    If Trim(vaSpread3.text) = Trim(fg_pone_cero(dia, 2) & "/" & fg_pone_cero(mes, 2) & "/" & fg_pone_cero(ano, 4)) Then
       vaSpread3.SetActiveCell 2, i
       Exit For
    End If
Next i
End Sub

Private Sub fpCalendar1_DateChanging(Month As Integer, Day As Integer, Year As Integer, State As Integer, ByVal Shift As Integer, Cancel As Integer)
ano = fg_pone_cero(Year, 4)
mes = fg_pone_cero(Month, 2)
dia = fg_pone_cero(Day, 2)
End Sub

Private Sub fpCalendar1_ViewChange(BeginMonth As Integer, BeginDay As Integer, BeginYear As Integer, EndMonth As Integer, EndDay As Integer, EndYear As Integer, Cancel As Integer)
ano = fg_pone_cero(EndYear, 4)
mes = fg_pone_cero(EndMonth, 2)
dia = fg_pone_cero(EndDay, 2)
Cancel = IIf(Toolbar1.Buttons(12).Visible = True, True, False)
If Cancel = False Then MoverDatosRutaCalendario
End Sub

Private Sub fpText1_Change(Index As Integer)
Select Case Index
Case 0
    If Est Then Exit Sub
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 14, 0, modo
Case 1
    If LimpiaDato(Trim(fpText1(1).text)) & Chr(KeyAscii) = "" Then Exit Sub
    vaSpread1.Visible = False
    If Combo1.ItemData(Combo1.ListIndex) = 0 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_ruta 6, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%'")
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_ruta 7, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%'")
    End If
    If RS2.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS2!nreg
    i = 1
    If Not RS2.EOF Then
       Do While Not RS2.EOF
          vaSpread1.Row = i: i = i + 1
          vaSpread1.Col = 1
          vaSpread1.TypeHAlign = 1
          vaSpread1.text = RS2!recorrido
          vaSpread1.Col = 2
          vaSpread1.text = Trim(RS2!descripcion)
          RS2.MoveNext
        Loop
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = True
        Gl_Ac_Botones Me, 14, 1, modo
    Else
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
    End If
    RS2.Close: Set RS2 = Nothing
    vaSpread1.Col = 1: vaSpread1.Col2 = vaSpread1.MaxCols: vaSpread1.Row = 1: vaSpread1.Row2 = vaSpread1.MaxRows
    vaSpread1.SetActiveCell 1, 1
    vaSpread1.Visible = True
    If fpText1(1).text = "" Then Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro" Else Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
End Select
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = vaSpread1.text
Select Case SSTab1.Tab
Case 0
    SSTab1.TabEnabled(4) = False
Case 1
    SSTab1.TabEnabled(4) = False
    MoverDatosRuta
Case 2
    SSTab1.TabEnabled(4) = False
    MoverDatosRutaProductos
Case 3
    MoverDatosRutaCalendario
Case 4
    MoverDatosCasinos
End Select
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 2, 3, 4, 5
    vaSpread2.Visible = False
    If Trim(Text1(Index).text) <> "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           vaSpread2.Col = Index: nom = UCase(Trim(vaSpread2.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread2.Col = 1
           If indactivo = -1 And Trim(vaSpread2.text) <> "" Then
              If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
           Else
              If vaSpread2.RowHidden = False Then vaSpread2.RowHidden = True
           End If
        Next i
        vaSpread2.SetActiveCell Index, 1
    End If
    vaSpread2.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread2.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread2.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread2.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread2.Sort -1, -1, vaSpread2.MaxCols, vaSpread2.MaxRows, SortByRow
    If Trim(Text1(Index).text) = "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
       Next
       vaSpread2.SetActiveCell Index, vaSpread2.SearchCol(Index, 0, vaSpread2.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread2.SetActiveCell Index, 1
    End If
    vaSpread2.Visible = True
End Select

End Sub

Private Sub TextCai1_Change(Index As Integer)
Select Case Index
Case 2, 3
    vaSpread4.Visible = False
    If Trim(TextCai1(Index).text) <> "" Then
       For i = 1 To vaSpread4.MaxRows
           vaSpread4.Row = i
           vaSpread4.Col = Index: nom = UCase(Trim(vaSpread4.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCai1(Index).text) & "*"
           vaSpread4.Col = 1
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
    vaSpread4.Sort -1, -1, vaSpread4.MaxCols, vaSpread4.MaxRows, SortByRow
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

Private Sub TextCan1_Change(Index As Integer)
Select Case Index
Case 2, 3
    vaSpread5.Visible = False
    If Trim(TextCan1(Index).text) <> "" Then
       For i = 1 To vaSpread5.MaxRows
           vaSpread5.Row = i
           vaSpread5.Col = Index: nom = UCase(Trim(vaSpread5.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCan1(Index).text) & "*"
           vaSpread5.Col = 1
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
    vaSpread5.Sort -1, -1, vaSpread5.MaxCols, vaSpread5.MaxRows, SortByRow
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
Dim i As Long, codpro As String, codcco As String, fectop As String, diatop As Long, horadi As String, cencos As String, auxcen As String, nomcen As String
Dim estmar As Boolean, FecDes As String
Select Case Button.Index
Case 1 '-------> Incluir nuevos registros
    modo = "A"
    Select Case SSTab1.Tab
    Case 0, 1 '-------> Ruta
        SSTab1.Tab = 1
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
        '-------> Traer ultimo registro
        Limpia 1
        Set RS = vg_dbpedweb.Execute("pedweb_s_ruta 1, '', ''")
        If Not RS.EOF Then RS.MoveFirst: codigo = RS!recorrido + 1 Else codigo = 1
        RS.Close: Set RS = Nothing
        fpLongInteger1(0).text = codigo
        fpText1(0).SetFocus
        vg_codigo = "x"
    Case 2 '-------> Ruta productos
        indgra = 0
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = False
        vg_codigo = ""
        B_RutPro.LlenaDatos codigo, Frame7.Caption, "rutpro", 0
        B_RutPro.Show 1
        If Trim(vg_codigo) = "" Then
           SSTab1.TabEnabled(0) = True
           SSTab1.TabEnabled(1) = True
           SSTab1.TabEnabled(2) = True
           SSTab1.TabEnabled(3) = True
           Exit Sub
        End If
        indgra = Val(vg_codigo)
    Case 3 '-------> Calendario
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = True
        SSTab1.TabEnabled(4) = False
        vg_codigo = "x"
        SetearCelda
    Case 4 '-------> Inlcuir casino a la ruta
        vg_codigo = ""
        If vaSpread5.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        '-------> Validar que exista un registro ruta producto seleccionado
        For i = 1 To vaSpread5.MaxRows
            vaSpread5.Row = i
            vaSpread5.Col = 1
            If vaSpread5.text = "1" Then estmar = True
        Next i
        If Not estmar Then MsgBox "Debe seleccionar un registro de los casino no incluidos...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        For i = 1 To vaSpread5.MaxRows
            vaSpread5.Row = i
            vaSpread5.Col = 1
            If vaSpread5.text = "1" Then
               vaSpread5.Col = 2
               auxcen = vaSpread5.text
               vaSpread5.Col = 3
               nomcen = vaSpread5.text
               vaSpread5.Col = 4
               cencos = vaSpread5.text
               '-------> Mover datos a vector de casino incluido en la ruta
               vaSpread4.MaxRows = vaSpread4.MaxRows + 1
               vaSpread4.Row = vaSpread4.MaxRows
               vaSpread4.Col = -1
               vaSpread4.BackColor = &H80000013
               vaSpread4.Col = 1
               vaSpread4.text = "1"
               vaSpread4.Col = 2
               vaSpread4.text = auxcen
               vaSpread4.Col = 3
               vaSpread4.text = nomcen
               vaSpread4.Col = 4
               vaSpread4.text = cencos
               vaSpread4.SetActiveCell 2, vaSpread4.MaxRows
            End If
        Next i
        vg_codigo = "X"
        '-------> Bloquer grilla casino incluidos
        vaSpread4.Row = -1
        vaSpread4.Col = -1
        vaSpread4.Lock = True
        '-------> Bloquear grilla casino no incluido
        vaSpread5.Row = -1
        vaSpread5.Col = -1
        vaSpread5.Lock = True
        '-------> Bloquear hoja
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        SSTab1.TabEnabled(4) = True
    End Select
    If vg_codigo <> "" Then Gl_Ac_Botones Me, 14, 0, modo
Case 3 '-------> Alterar registro
    Select Case SSTab1.Tab
    Case 0, 1
        modo = "M"
        Gl_Ac_Botones Me, 14, 0, modo
        SSTab1.Tab = 1
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = False
        fpText1(0).SetFocus
    Case 3
        modo = "M"
        SetearCelda
        Gl_Ac_Botones Me, 14, 0, modo
        SSTab1.Tab = 3
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = True
    End Select
Case 5 '-------> Eliminar Registro
    Dim cencom As String, codfam As String
    estmar = False
    Select Case SSTab1.Tab
    Case 0, 1
        If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        '-------> Validar si existe registro relacionado a la ruta productos
        Set RS = vg_dbpedweb.Execute("pedweb_s_ruta 2, " & codigo & ", ''")
        If Not RS.EOF Then
           RS.Close: Set RS = Nothing: MsgBox "Existen datos relacionado, en la ruta productos...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        End If
        RS.Close: Set RS = Nothing
        '-------> Validar si existe registro relacionado a la ruta calendarios
        Set RS = vg_dbpedweb.Execute("pedweb_s_ruta 3, " & codigo & ", ''")
        If Not RS.EOF Then
           RS.Close: Set RS = Nothing: MsgBox "Existen datos relacionado, en la ruta calendarios...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        End If
        RS.Close: Set RS = Nothing
        '-------> borrar ruta
        vg_dbpedweb.Execute ("pedweb_d_ruta " & codigo & "")
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        codigo = 0
        If vaSpread1.MaxRows > 0 Then
           vaSpread1.SetActiveCell 1, vaSpread1.MaxRows
           vaSpread1.Row = vaSpread1.ActiveRow
           vaSpread1.Col = 1
           codigo = vaSpread1.text
        End If
        MoverDatosRuta
        MoverDatosRutaProductos
        MoverDatosRutaCalendario
        modo = "": Gl_Ac_Botones Me, 14, 1, modo
    Case 2
        If vaSpread2.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        '-------> Validar que exista un registro ruta producto seleccionado
        For i = 1 To vaSpread2.MaxRows
            vaSpread2.Row = i
            vaSpread2.Col = 1
            If vaSpread2.text = "1" Then estmar = True
        Next i
        If Not estmar Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
        '-------> rutina de borrado ruta productos
        fg_carga ""
        For i = 1 To vaSpread2.MaxRows
            vaSpread2.Row = i
            vaSpread2.Col = 1
            If vaSpread2.text = "1" Then
               vaSpread2.Col = 2
               codpro = Trim(vaSpread2.text)
               vaSpread2.Col = 3
               cencom = Trim(vaSpread2.text)
               vaSpread2.Col = 4
               codfam = Trim(vaSpread2.text)
               vg_dbpedweb.Execute ("pedweb_d_rutaproductos " & codigo & ", '" & codpro & "', '" & cencom & "'")
            End If
        Next i
        fg_descarga
        MoverDatosRutaProductos
        modo = "": Gl_Ac_Botones Me, 14, 1, modo
    Case 3
        '-------> Validar que exista un registro ruta calendario seleccionado
        If vaSpread3.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        For i = 1 To vaSpread3.MaxRows
            vaSpread3.Row = i
            vaSpread3.Col = 1
            If vaSpread3.text = "1" Then estmar = True
        Next i
        If Not estmar Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
        '-------> rutina de borrado ruta calendario
        fg_carga ""
        For i = 1 To vaSpread3.MaxRows
            vaSpread3.Row = i
            vaSpread3.Col = 1
            If vaSpread3.text = "1" Then
               vaSpread3.Col = 2
               FecDes = Trim(vaSpread3.text)
               vg_dbpedweb.Execute ("pedweb_d_rutacalendario " & codigo & ", '" & Format(FecDes, "yyyymmdd") & "'")
            End If
        Next i
        fg_descarga
        MoverDatosRutaCalendario
        modo = "": Gl_Ac_Botones Me, 14, 1, modo
    Case 4
        '-------> Validar que exista un registro ruta calendario seleccionado
        If vaSpread4.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        For i = 1 To vaSpread4.MaxRows
            vaSpread4.Row = i
            vaSpread4.Col = 1
            If vaSpread4.text = "1" Then estmar = True
        Next i
        If Not estmar Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
        '-------> rutina de borrado ruta calendario
        vaSpread3.Row = vaSpread3.ActiveRow
        vaSpread3.Col = 2
        FecDes = Format(vaSpread3.text, "yyyymmdd")
        For i = 1 To vaSpread4.MaxRows
            vaSpread4.Row = i
            vaSpread4.Col = 1
            If vaSpread4.text = "1" Then
               vaSpread4.Col = 4
               cencos = Trim(vaSpread4.text)
               vg_dbpedweb.Execute ("pedweb_d_casinoincluidoruta " & codigo & ", '" & FecDes & "', '" & cencos & "'")
            End If
        Next i
        MoverDatosCasinos
        modo = "": Gl_Ac_Botones Me, 14, 1, modo
    End Select
Case 7 '-------> Actualizar lista
    Select Case SSTab1.Tab
    Case 0
        MoverDatosGrilla
        MoverDatosRuta
        MoverDatosRutaProductos
        MoverDatosRutaCalendario
        fpText1(1).text = ""
        Gl_Ac_Botones Me, 14, IIf(vaSpread1.MaxRows > 0, 1, 2), modo
    Case 1
        MoverDatosRuta
    Case 2
        MoverDatosRutaProductos
    Case 3
        MoverDatosRutaCalendario
    Case 4
        MoverDatosCasinos
    End Select
Case 10 '-------> Cancelar Información
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    Select Case SSTab1.Tab
    Case 1
        SSTab1.Tab = 1
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        MoverDatosRuta
        MoverDatosRutaProductos
        MoverDatosRutaCalendario
    Case 2
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        MoverDatosRutaProductos
    Case 3
        MoverDatosRutaCalendario
    Case 4
        MoverDatosCasinos
    End Select
    '-------> Desbloquear hojas
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    modo = "": Gl_Ac_Botones Me, 14, 1, modo
Case 12 '-------> grabaRegistro
    Select Case SSTab1.Tab
    Case 1 '-------> Grabar ruta
        If LimpiaDato(Trim(fpText1(0).text)) = "" Then MsgBox "Debe ingresar información...", vbCritical, Msgtitulo: Exit Sub
        If modo = "A" Then
           codigo = 0
           Set RS = vg_dbpedweb.Execute("pedweb_iu_ruta 'A', '', '" & LimpiaDato(Trim(fpText1(0).text)) & "'")
           If Not RS.EOF Then
              codigo = RS!indice
           End If
           RS.Close: Set RS = Nothing
           fpLongInteger1(0).text = codigo
           vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
           vaSpread1.SetActiveCell 1, vaSpread1.Row
        Else
            vg_dbpedweb.Execute "pedweb_iu_ruta 'M', " & codigo & ", '" & LimpiaDato(Trim(fpText1(0).text)) & "'"
        End If
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = LimpiaDato(Trim(fpLongInteger1(0).text))
        vaSpread1.Col = 2: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = LimpiaDato(Trim(fpText1(0).text))
    Case 2 '-------> Grabar ruta producto
        For i = indgra To vaSpread2.MaxRows
            vaSpread2.Row = i
            vaSpread2.Col = 2
            codpro = vaSpread2.text
            vaSpread2.Col = 3
            codcco = vaSpread2.text
            vg_dbpedweb.Execute ("DELETE s_Recorrido_Productos WHERE recorrido = " & codigo & " AND producto = '" & codpro & "' AND CCompra = '" & codcco & "'")
            vg_dbpedweb.Execute ("INSERT INTO s_Recorrido_Productos VALUES (" & codigo & ", '" & codpro & "', '" & codcco & "')")
        Next i
    Case 3 '-------> Grabar calendario
        vaSpread3.Row = vaSpread3.ActiveRow
        vaSpread3.Col = 3
        fectop = vaSpread3.text
        vaSpread3.Col = 4
        diatop = vaSpread3.text
        vaSpread3.Col = 5
        horadi = vaSpread3.text
        If Trim(fectop) = "" Or diatop = 0 Or Trim(horadi) = "" Then MsgBox "Faltan datos en calendario...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        vaSpread3.Col = 5
        horadi = "19000101 " & vaSpread3.text
        If modo = "A" Then
           fpCalendar1.NextSelection = ""
           For i = 1 To fpCalendar1.SelCount
               fecdfe = fpCalendar1.NextSelection
               If Val(fecdfe) > 0 Then
                  fpCalendar1.ElementIndex = fecdfe
                  vg_dbpedweb.Execute ("DELETE s_Recorrido_Despacho WHERE ruta = " & codigo & " AND FechaDespacho = '" & fecdfe & "'")
                  vg_dbpedweb.Execute ("INSERT INTO s_Recorrido_Despacho VALUES (" & codigo & ", '" & fecdfe & "', '" & Format(fectop, "yyyymmdd") & "', " & diatop & ", '" & horadi & "')")
                  fpCalendar1.Element = ElementSpecificDate
                  fpCalendar1.ElementBackColor = -2147483633
                  fpCalendar1.ElementText = ""
                  fpCalendar1.ElementForeColor = vbBlack
               End If
           Next i
           MoverDatosRutaCalendario
        Else
           For i = 1 To vaSpread3.MaxRows
               vaSpread3.Row = i
               vaSpread3.Col = 2
               If vaSpread3.BackColor = &H80000013 Then
                  vaSpread3.Col = 3
                  fectop = vaSpread3.text
                  vaSpread3.Col = 4
                  diatop = vaSpread3.text
                  vaSpread3.Col = 5
                  horadi = vaSpread3.text
                  vaSpread3.Col = 5
                  horadi = "19000101 " & vaSpread3.text
                  vaSpread3.Col = 2
                  fecdfe = Format(vaSpread3.text, "yyyymmdd")
                  vg_dbpedweb.Execute ("UPDATE s_Recorrido_Despacho SET FechaTopeIngreso = '" & Format(fectop, "yyyymmdd") & "', DiasTopeAdicional = " & diatop & ", HoraTopeAdicional = '" & horadi & "' WHERE ruta = " & codigo & " AND FechaDespacho = '" & fecdfe & "'")
                  vaSpread3.Row = i 'vaSpread3.ActiveRow
                  vaSpread3.Col = -1
                  vaSpread3.BackColor = &H80000018
               End If
           Next i
        End If
    Case 4 '-------> Grabar calendario casino
        fg_carga ""
        vaSpread3.Row = vaSpread3.ActiveRow
        vaSpread3.Col = 2
        FecDes = Format(vaSpread3.text, "yyyymmdd")
        For i = 1 To vaSpread4.MaxRows
            vaSpread4.Row = i
            vaSpread4.Col = 1
            If vaSpread4.text = "1" Then
               vaSpread4.Col = 4
               cencos = vaSpread4.text
               vg_dbpedweb.Execute ("INSERT INTO s_Recorrido_Despacho_casino (ruta, FechaDespacho, Casino, fecha2) VALUES (" & codigo & ", convert(datetime, '" & vaSpread3.text & "', 103), '" & cencos & "', null)")
            End If
        Next i
       MoverDatosCasinos
    End Select
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    modo = "": Gl_Ac_Botones Me, 14, 1, modo
Case 19 '------> impresion
    Select Case SSTab1.Tab
    Case 0, 1 '-------> ruta
        I_Ruta
    Case 2 '-------> Ruta productos
        vg_opimp = 99999
        I_WebRep.LlenaDatos "Impresión Rutas Productos", "rutprod"
        I_WebRep.Show 1
        Me.Refresh
        vg_opimp = 0
'        vaSpread1.Row = vaSpread1.ActiveRow
'        vaSpread1.Col = 1
'        codigo = vaSpread1.Text
'        I_RutaProductos CStr(codigo)
    Case 3 '-------> Ruta Calendarios
'        I_WebRep.LlenaDatos "Impresión Rutas Calendarios", "rutcale"
'        I_WebRep.Show 1
'        Me.Refresh
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        I_RutaCalendarios CStr(codigo), fpCalendar1.Year, fg_pone_cero(fpCalendar1.Month, 2)
    Case 4
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        I_RutaCalendarioCasinos CStr(codigo), fpCalendar1.Year, fg_pone_cero(fpCalendar1.Month, 2), fg_pone_cero(fpCalendar1.Day, 2)
    End Select
Case 22
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu
Case "Copiar Datos"
    M_CopProCan.LlenaDatos "Copiar Rutas", "ruta"
    M_CopProCan.Show 1
    Me.Refresh
Case "Importar Datos"
    If vaSpread1.MaxRows < 1 Then Exit Sub
    P_ImpRut.LlenaDatos "Importar Rutas Productos - Calendario - Casino", "ruta"
    P_ImpRut.Show 1
End Select
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = Row
vaSpread1.Col = 1
codigo = "": codigo = Val(vaSpread1.text)
MoverDatosRuta
MoverDatosRutaProductos
MoverDatosRutaCalendario
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

Private Sub vaSpread2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
Dim i As Long
Select Case BlockCol
Case 1
    vaSpread2.Col = 1
    For i = BlockRow To BlockRow2
        vaSpread2.Row = i
        vaSpread2.Value = IIf(vaSpread2.Value = "1", "0", "1")
    Next i
End Select
End Sub

Private Sub vaSpread2_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If vaSpread2.MaxRows < 1 Or Col = 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread2.Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
Select Case Col
Case 2
    vaSpread2.Col = Col
    TipText = "Código Productos : " & vaSpread2.text
Case 3
    vaSpread2.Col = Col
    TipText = "Central de Compras : " & Trim(vaSpread2.text)
Case 4
    vaSpread2.Col = Col
    TipText = "Familia de Producto : " & Trim(vaSpread2.text)
Case 5
    vaSpread2.Col = Col
    TipText = "Descripción : " & Trim(vaSpread2.text)
End Select
End Sub

Private Sub vaSpread3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
Dim i As Long
Select Case BlockCol
Case 1
    vaSpread3.Col = 1
    For i = BlockRow To BlockRow2
        vaSpread3.Row = i
        vaSpread3.Value = IIf(vaSpread3.Value = "1", "0", "1")
    Next i
End Select
End Sub

Sub SetearCelda()
If modo = "A" Then
   vaSpread3.MaxRows = vaSpread3.MaxRows + 1
   vaSpread3.Row = vaSpread3.MaxRows
   vaSpread3.SetActiveCell 3, vaSpread3.MaxRows
Else
   vaSpread3.Row = vaSpread3.ActiveRow
   vaSpread3.Col = -1
   vaSpread3.BackColor = &H80000013
End If

'-------> definir formato fecha
vaSpread3.Col = 3
vaSpread3.CellType = CellTypeDate
vaSpread3.TypeDateFormat = TypeDateFormatDDMMYY
vaSpread3.TypeDateMin = "01011973"
vaSpread3.TypeDateMax = "31125000"
vaSpread3.TypeHAlign = TypeHAlignCenter
vaSpread3.TypeDateCentury = True
If modo = "A" Then vaSpread3.text = Format(Date, "dd/mm/yyyy")

'-------> definir formato numerico
vaSpread3.Col = 4
vaSpread3.CellType = CellTypeNumber
vaSpread3.TypeNumberDecPlaces = 0
vaSpread3.TypeIntegerMin = 1
vaSpread3.TypeIntegerMax = 99
vaSpread3.TypeHAlign = TypeHAlignCenter
vaSpread3.TypeSpin = False
vaSpread3.TypeIntegerSpinInc = 1
vaSpread3.TypeIntegerSpinWrap = False

'-------> definir hora
vaSpread3.Col = 5
If modo = "M" Then hora = Mid(vaSpread3.text, 1, 2) & Mid(vaSpread3.text, 4, 2) & "00" Else hora = Format(Time, "hhmmss") '"145236"
vaSpread3.CellType = CellTypeTime
vaSpread3.TypeTime24Hour = TypeTime24Hour24HourClock
vaSpread3.TypeTimeSeconds = False
vaSpread3.TypeTimeSeparator = Asc(":")
vaSpread3.TypeHAlign = TypeHAlignCenter
vaSpread3.Value = hora
End Sub

Private Sub vaSpread4_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
Dim i As Long
vaSpread4.Col = 1
For i = BlockRow To BlockRow2
    vaSpread4.Row = i
    vaSpread4.Value = IIf(vaSpread4.Value = "1", "0", "1")
Next i
End Sub

Private Sub vaSpread4_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If vaSpread4.MaxRows < 1 Or Col = 1 Then Exit Sub
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

Private Sub vaSpread5_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
Dim i As Long
vaSpread5.Col = 1
For i = BlockRow To BlockRow2
    vaSpread5.Row = i
    vaSpread5.Value = IIf(vaSpread5.Value = "1", "0", "1")
Next i
End Sub

Private Sub vaSpread5_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If vaSpread5.MaxRows < 1 Or Col = 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread5.Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
Select Case Col
Case 2
    vaSpread5.Col = Col
    TipText = "Centro de Costo : " & vaSpread5.text
Case 3
    vaSpread5.Col = Col
    TipText = "Descripción : " & Trim(vaSpread5.text)
End Select
End Sub
