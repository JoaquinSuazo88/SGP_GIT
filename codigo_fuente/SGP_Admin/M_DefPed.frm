VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_DefPed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definir Pedidos"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12990
      _ExtentX        =   22913
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lista de Definición Pedidos"
      TabPicture(0)   =   "M_DefPed.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "vaSpread1"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Definición Pedidos"
      TabPicture(1)   =   "M_DefPed.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Familia Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   240
         TabIndex        =   23
         Top             =   3000
         Width           =   12255
         Begin VB.Frame Frame12 
            Height          =   435
            Left            =   3450
            TabIndex        =   26
            Top             =   4560
            Width           =   6525
            Begin VB.TextBox TextCai1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   3
               Left            =   45
               TabIndex        =   27
               Top             =   135
               Width           =   6420
            End
         End
         Begin VB.Frame Frame13 
            Height          =   435
            Left            =   2520
            TabIndex        =   24
            Top             =   4560
            Width           =   915
            Begin VB.TextBox TextCai1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   2
               Left            =   45
               TabIndex        =   25
               Top             =   135
               Width           =   810
            End
         End
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   4095
            Left            =   2160
            TabIndex        =   28
            Top             =   360
            Width           =   8055
            _Version        =   393216
            _ExtentX        =   14208
            _ExtentY        =   7223
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
            SpreadDesigner  =   "M_DefPed.frx":0038
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   240
         TabIndex        =   17
         Top             =   660
         Width           =   12255
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   2640
            TabIndex        =   6
            Top             =   240
            Width           =   885
            _Version        =   196608
            _ExtentX        =   1561
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
            MaxLength       =   4
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
            Left            =   2640
            TabIndex        =   7
            Top             =   600
            Width           =   525
            _Version        =   196608
            _ExtentX        =   926
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
            MaxValue        =   "99"
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   2
            Left            =   2640
            TabIndex        =   9
            Top             =   1320
            Width           =   5805
            _Version        =   196608
            _ExtentX        =   10239
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
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   1
            Left            =   2640
            TabIndex        =   8
            Top             =   960
            Width           =   945
            _Version        =   196608
            _ExtentX        =   1667
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
            Text            =   "05/2016"
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
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   0
            Left            =   2640
            TabIndex        =   10
            Top             =   1680
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
            Text            =   "13/07/2004"
            DateCalcMethod  =   4
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
            ButtonColor     =   -2147483637
            AutoMenu        =   -1  'True
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   2
            Left            =   6160
            TabIndex        =   12
            Top             =   1680
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
            Text            =   "13/07/2004"
            DateCalcMethod  =   4
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
            ButtonColor     =   -2147483637
            AutoMenu        =   -1  'True
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   3
            Left            =   4080
            TabIndex        =   11
            Top             =   1680
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            ButtonIncrement =   0
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
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   1
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
            Text            =   "10:10"
            DateCalcMethod  =   3
            DateTimeFormat  =   5
            UserDefinedFormat=   "hh:mm"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   2
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   2
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
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   4
            Left            =   7630
            TabIndex        =   13
            Top             =   1680
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   1
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
            Text            =   "10:10"
            DateCalcMethod  =   3
            DateTimeFormat  =   5
            UserDefinedFormat=   "hh:mm"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   2
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   2
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Periodo de Digitación"
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
            Left            =   600
            TabIndex        =   22
            Top             =   1760
            Width           =   1845
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
            Index           =   6
            Left            =   600
            TabIndex        =   21
            Top             =   1450
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
            Left            =   600
            TabIndex        =   20
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "C. Compras"
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
            Left            =   600
            TabIndex        =   19
            Top             =   320
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Semana"
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
            Left            =   600
            TabIndex        =   18
            Top             =   700
            Width           =   690
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   971
         Left            =   -72120
         TabIndex        =   2
         Top             =   540
         Width           =   7335
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "M_DefPed.frx":195B
            Left            =   2010
            List            =   "M_DefPed.frx":196B
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
            TabIndex        =   16
            Top             =   345
            Width           =   1380
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
            TabIndex        =   15
            Top             =   645
            Width           =   1140
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
            TabIndex        =   14
            Top             =   645
            Width           =   585
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6285
         Left            =   -74760
         TabIndex        =   5
         Top             =   1620
         Width           =   12210
         _Version        =   393216
         _ExtentX        =   21537
         _ExtentY        =   11086
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
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
         MaxCols         =   7
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_DefPed.frx":199E
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
   End
End
Attribute VB_Name = "M_DefPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String, codigo As Long, Msgtitulo As String
Dim est As Boolean, estmod As Boolean

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 9405
Me.Width = 13095
Msgtitulo = "Definir Pedidos"
fg_centra Me
SSTab1.Tab = 0
modo = ""
est = True
estmod = False
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 0
MoverDatosGrilla
MoverDatosDefiniciónPedidos
Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows > 0, 1, 2), modo
End Sub

Sub MoverDatosGrilla()
fg_carga ""
Dim x As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 250
x = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
vaSpread1.Row = -1
vaSpread1.Col = -1
vaSpread1.Lock = True
Set RS = vg_dbpedweb.Execute("pedweb_s_definirpedidos 1, '', '', '', '', 0")
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   vaSpread1.text = Trim(RS!IdDefinicion)
   vaSpread1.Col = 2
   vaSpread1.text = Trim(RS!CentralDeCompra)
   vaSpread1.Col = 3
   vaSpread1.text = RS!YearMes 'Trim(Mid(RS!YearMes, 1, 2)) & "/" & Trim(Mid(RS!YearMes, 1, 4))
   vaSpread1.Col = 4
   vaSpread1.text = Trim(RS!Semana)
   vaSpread1.Col = 5
   vaSpread1.text = Trim(RS!descripcion)
   vaSpread1.Col = 6
   vaSpread1.text = Format(RS!IniFchTopeDigitacion, "dd/mm/yyyy") & " " & Format(RS!IniFchTopeDigitacion, "hh:mm")
   vaSpread1.Col = 7
   vaSpread1.text = Format(RS!FinFchTopeDigitacion, "dd/mm/yyyy") & " " & Format(RS!FinFchTopeDigitacion, "hh:mm")
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
If vaSpread1.MaxRows > 0 Then
   vaSpread1.Row = 1
   vaSpread1.Col = 1
   codigo = 0
   codigo = Val(vaSpread1.text)
   vaSpread1.SetActiveCell 1, 1 ': vaSpread1.SetFocus
End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
fg_descarga
End Sub

Sub MoverDatosDefiniciónPedidos()
fg_carga ""
est = True
If modo <> "A" Then
   estmod = False
   Set RS = vg_dbpedweb.Execute("pedweb_s_definirpedidos 9, " & codigo & ", '', '', '', 0")
   If RS.EOF Then estmod = False Else estmod = True
   RS.Close: Set RS = Nothing
   Limpia 1
   Set RS = vg_dbpedweb.Execute("pedweb_s_definirpedidos 2, " & codigo & ", '', '', '', 0")
   If Not RS.EOF Then
      fpText1(0).text = Trim(RS!CentralDeCompra)
      fpLongInteger1(0).Value = IIf(IsNull(RS!Semana), 0, RS!Semana)
      fpDateTime1(1).text = IIf(IsNull(RS!YearMes), "", Mid(RS!YearMes, 5, 2) & "/" & Mid(RS!YearMes, 1, 4))
      fpText1(2).text = Trim(IIf(IsNull(RS!descripcion), "", RS!descripcion))
      fpDateTime1(0).text = IIf(IsNull(RS!IniFchTopeDigitacion), "", Format(RS!IniFchTopeDigitacion, "dd/mm/yyyy"))
      fpDateTime1(3).text = IIf(IsNull(RS!IniFchTopeDigitacion), "", Format(RS!IniFchTopeDigitacion, "hh:mm"))
      fpDateTime1(2).text = IIf(IsNull(RS!FinFchTopeDigitacion), "", Format(RS!FinFchTopeDigitacion, "dd/mm/yyyy"))
      fpDateTime1(4).text = IIf(IsNull(RS!FinFchTopeDigitacion), "", Format(RS!FinFchTopeDigitacion, "hh:mm"))
   End If
   RS.Close: Set RS = Nothing
End If
Limpia 2
If modo <> "A" Then
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = 1
   codigo = vaSpread1.text
End If
Dim x As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread2.TextTip = 2
vaSpread2.TextTipDelay = 250
x = vaSpread2.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread2.Visible = False
Set RS = vg_dbpedweb.Execute("pedweb_s_definirpedidos 7, " & codigo & ", '', '" & fpText1(0).text & "', '" & Format(fpDateTime1(1).text, "yyyymm") & "', " & Val(fpLongInteger1(0).Value) & "")
Do While Not RS.EOF
   vaSpread2.MaxRows = vaSpread2.MaxRows + 1
   vaSpread2.Row = vaSpread2.MaxRows
   vaSpread2.Col = 1
   vaSpread2.Lock = IIf(Not estmod And (IsNull(RS!sel) Or RS!sel = 0), False, IIf(estmod, False, True))
   vaSpread2.text = IIf(IsNull(RS!sel) Or RS!sel = 0, "0", "1")
   vaSpread2.Col = 2
   vaSpread2.text = IIf(IsNull(RS!codigo), "", RS!codigo)
   vaSpread2.Col = 3
   vaSpread2.text = Trim(IIf(IsNull(RS!descripcion), "", RS!descripcion))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread2.Visible = True
est = False
fg_descarga
End Sub

Sub Limpia(op As Integer)
Select Case op
Case 1
   fpText1(0).text = ""
   fpText1(0).Enabled = IIf(modo = "A", True, False)
   fpText1(2).text = ""
   fpText1(2).Enabled = estmod
   fpLongInteger1(0).Value = ""
   fpLongInteger1(0).Enabled = estmod
   fpDateTime1(1).text = Format(Date, "mm/yyyy")
   fpDateTime1(1).Enabled = estmod
   fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
   fpDateTime1(0).Enabled = estmod
   fpDateTime1(2).text = Format(Date, "dd/mm/yyyy")
   fpDateTime1(2).Enabled = estmod
   fpDateTime1(3).text = Format(Time, "hh:mm")
   fpDateTime1(3).Enabled = estmod
   fpDateTime1(4).text = Format(Time, "hh:mm")
   fpDateTime1(4).Enabled = estmod
Case 2
    vaSpread2.MaxRows = 0
    vaSpread2.Col = -1
    vaSpread2.Row = -1
    vaSpread2.Lock = IIf(Not estmod, True, False)
    TextCai1(2).text = ""
    TextCai1(3).text = ""
End Select
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0, 1, 2
    If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
    If Index = 1 And modo = "A" Then
       MoverDatosDefiniciónPedidos
    End If
Case 3, 4
    If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
End Select
If Toolbar1.Buttons(12).Visible = False Then
   SSTab1.TabEnabled(0) = False
   If modo = "" Then modo = "M"
   Gl_Ac_Botones Me, 1, 0, modo
End If
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
If est Then Exit Sub
If Index = 0 And modo = "A" Then
   MoverDatosDefiniciónPedidos
End If

If Toolbar1.Buttons(12).Visible = False Then
   SSTab1.TabEnabled(0) = False
   If modo = "" Then modo = "M"
   Gl_Ac_Botones Me, 1, 0, modo
End If
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
'If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_Change(Index As Integer)
Select Case Index
Case 0, 2
    If est Then Exit Sub
    If Index = 0 And modo = "A" Then
       MoverDatosDefiniciónPedidos
    End If
    If Toolbar1.Buttons(12).Visible = False Then
       SSTab1.TabEnabled(0) = False
       If modo = "" Then modo = "M"
       Gl_Ac_Botones Me, 1, 0, modo
    End If
Case 1
    If LimpiaDato(Trim(fpText1(1).text)) & Chr(KeyAscii) = "" Then Exit Sub
    vaSpread1.Visible = False
    If Combo1.ItemData(Combo1.ListIndex) = 0 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_definirpedidos 3, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%', '', '', 0")
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_definirpedidos 4, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%', '', '', 0")
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 2 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_definirpedidos 5, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%', '', '', 0")
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 3 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_definirpedidos 6, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%', '', '', 0")
    End If
    If RS2.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS2!nReg
    i = 1
    If Not RS2.EOF Then
       Do While Not RS2.EOF
          vaSpread1.Row = i: i = i + 1
          vaSpread1.Col = 1
          vaSpread1.text = Trim(RS2!IdDefinicion)
          vaSpread1.Col = 2
          vaSpread1.text = Trim(RS2!CentralDeCompra)
          vaSpread1.Col = 3
          vaSpread1.text = RS2!YearMes 'Trim(Mid(RS!YearMes, 1, 2)) & "/" & Trim(Mid(RS!YearMes, 1, 4))
          vaSpread1.Col = 4
          vaSpread1.text = Trim(RS2!Semana)
          vaSpread1.Col = 5
          vaSpread1.text = Trim(RS2!descripcion)
          vaSpread1.Col = 6
          vaSpread1.text = Format(RS2!IniFchTopeDigitacion, "dd/mm/yyyy") & " " & Format(RS2!IniFchTopeDigitacion, "hh:mm")
          vaSpread1.Col = 7
          vaSpread1.text = Format(RS2!FinFchTopeDigitacion, "dd/mm/yyyy") & " " & Format(RS2!FinFchTopeDigitacion, "hh:mm")
          RS2.MoveNext
        Loop
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        Gl_Ac_Botones Me, 1, 1, modo
    Else
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
    End If
    RS2.Close: Set RS2 = Nothing
    vaSpread1.Col = 1: vaSpread1.col2 = vaSpread1.maxcols: vaSpread1.Row = 1: vaSpread1.row2 = vaSpread1.MaxRows
    vaSpread1.SetActiveCell 1, 1
    vaSpread1.Visible = True
    If fpText1(1).text = "" Then Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro" Else Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
End Select
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If modo <> "A" Then
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = 1
   codigo = vaSpread1.text
End If
Select Case SSTab1.Tab
Case 1
    MoverDatosDefiniciónPedidos
End Select
End Sub

Private Sub TextCai1_Change(Index As Integer)
Select Case Index
Case 2, 3
    vaSpread2.Visible = False
    If Trim(TextCai1(Index).text) <> "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           vaSpread2.Col = Index: nom = UCase(Trim(vaSpread2.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCai1(Index).text) & "*"
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
    vaSpread2.ColUserSortIndicator(IIf(Trim(TextCai1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread2.SortKey(1) = IIf(Trim(TextCai1(Index).text) = "", 0, 0): vaSpread2.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread2.Sort -1, -1, vaSpread2.maxcols, vaSpread2.MaxRows, SortByRow
    If Trim(TextCai1(Index).text) = "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
       Next
       vaSpread2.SetActiveCell Index, vaSpread2.SearchCol(Index, 0, vaSpread2.MaxRows, Trim(TextCai1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread2.SetActiveCell Index, 1
    End If
    vaSpread2.Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1 '-------> Incluir nuevos registros
    modo = "A"
    codigo = 0
    SSTab1.Tab = 1
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = True
    estmod = True
    Limpia 1
    Limpia 2
    fpText1(0).SetFocus
    vg_codigo = "x"
    If vg_codigo <> "" Then Gl_Ac_Botones Me, 1, 0, modo
Case 3 '-------> Alterar registro
    If vaSpread1.MaxRows < 1 Then Exit Sub
    modo = "M"
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    codigo = vaSpread1.text
    Set RS = vg_dbpedweb.Execute("pedweb_s_definirpedidos 9, " & codigo & ", '', '', '', 0")
    If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
    RS.Close: Set RS = Nothing
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.Tab = 1
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = True
    If fpText1(2).Enabled = True Then fpText1(2).SetFocus
Case 5 '-------> Eliminar Registro
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    codigo = vaSpread1.text
    Set RS = vg_dbpedweb.Execute("pedweb_s_definirpedidos 8, " & codigo & ", '', '', '', 0")
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No es posible eliminar documento. Periodo de digitación no esta vigente...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    If MsgBox("Elimina registro y todas sus relaciones...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    '-------> borrar ruta
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    codigo = vaSpread1.text
    vg_dbpedweb.Execute ("pedweb_d_definirpedidos " & codigo & "")
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
    MoverDatosDefiniciónPedidos
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
Case 7 '-------> Actualizar lista
    Select Case SSTab1.Tab
    Case 0
        fpText1(1).text = ""
        MoverDatosGrilla
    Case 1
        MoverDatosDefiniciónPedidos
    End Select
Case 10 '-------> Cancelar Información
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    SSTab1.Tab = 1
    If vaSpread1.MaxRows > 0 Then modo = "M"
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    codigo = vaSpread1.text
    MoverDatosDefiniciónPedidos
    '-------> Desbloquear hojas
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
Case 12 '-------> grabaRegistro
    Dim i As Long, isel As Boolean, codfam As String
    isel = False
    For i = 1 To vaSpread2.MaxRows
        vaSpread2.Row = i
        vaSpread2.Col = 1
        If vaSpread2.text = "1" Then isel = True: Exit For
    Next i
    If Not isel Then MsgBox "Debe seleccionar a lo menos una familia...", vbCritical, Msgtitulo: Exit Sub
    If LimpiaDato(Trim(fpText1(0).text)) = "" Or LimpiaDato(Trim(fpText1(2).text)) = "" Or fpDateTime1(0).text = "" _
       Or fpDateTime1(1).text = "" Or fpDateTime1(2).text = "" Or fpDateTime1(3).text = "" Or fpDateTime1(4).text = "" _
       Or fpLongInteger1(0).Value = 0 Then MsgBox "Debe ingresar información...", vbCritical, Msgtitulo: Exit Sub
    If modo = "A" Then
       '-------> Validar central-periodo-semana
'       Set RS = vg_dbpedweb.Execute("pedweb_s_definirpedidos 10, 0, '', '" & LimpiaDato(Trim(fpText1(0).Text)) & "', '" & Format(fpDateTime1(1).Text, "yyyymm") & "', " & fpLongInteger1(0).Value & "")
'       If Not RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Registrio ya existe...", vbCritical, Msgtitulo: Exit Sub
'       RS.Close: Set RS = Nothing
       codigo = 0
       Set RS = vg_dbpedweb.Execute("pedweb_iu_definepedidos 'A', 0, '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & Format(fpDateTime1(1).text, "yyyymm") & "', " & fpLongInteger1(0).Value & ", '" & Format(fpDateTime1(0).text, "yyyymmdd") & " " & fpDateTime1(3).text & "', '" & Format(fpDateTime1(2).text, "yyyymmdd") & " " & fpDateTime1(4).text & "', '" & LimpiaDato(Trim(fpText1(2).text)) & "'")
       If Not RS.EOF Then
          codigo = RS!indice
       End If
       RS.Close: Set RS = Nothing
       '-------> Grabar detalle definición pedido
       Set RS = vg_dbpedweb.Execute("pedweb_d_detalledefinepedidos " & codigo & "")
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           vaSpread2.Col = 1
           If vaSpread2.text = "1" Then
              codfam = ""
              vaSpread2.Col = 2: codfam = vaSpread2.text
              Set RS = vg_dbpedweb.Execute("pedweb_iu_detalledefinepedidos 'A', " & codigo & ", '" & codfam & "'")
           End If
       Next i
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
       vaSpread1.SetActiveCell 1, vaSpread1.Row
    Else
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = 1
       codigo = vaSpread1.text
       vg_dbpedweb.Execute ("pedweb_iu_definepedidos 'M', " & codigo & ", '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & Format(fpDateTime1(1).text, "yyyymm") & "', " & fpLongInteger1(0).Value & ", '" & Format(fpDateTime1(0).text, "yyyymmdd") & " " & fpDateTime1(3).text & "', '" & Format(fpDateTime1(2).text, "yyyymmdd") & " " & fpDateTime1(4).text & "', '" & LimpiaDato(Trim(fpText1(2).text)) & "'")
       '-------> Grabar detalle definición pedido
       Set RS = vg_dbpedweb.Execute("pedweb_d_detalledefinepedidos " & codigo & "")
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           vaSpread2.Col = 1
           If vaSpread2.text = "1" Then
              codfam = ""
              vaSpread2.Col = 2: codfam = vaSpread2.text
              Set RS = vg_dbpedweb.Execute("pedweb_iu_detalledefinepedidos 'A', " & codigo & ", '" & codfam & "'")
           End If
       Next i
    
    End If
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = codigo
    vaSpread1.Col = 2: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = LimpiaDato(Trim(fpText1(0).text))
    vaSpread1.Col = 3: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = Format(fpDateTime1(1).text, "yyyymm")
    vaSpread1.Col = 4: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = fpLongInteger1(0).Value
    vaSpread1.Col = 5: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = LimpiaDato(Trim(fpText1(2).text))
    vaSpread1.Col = 6: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = fpDateTime1(0).text & " " & fpDateTime1(3).text
    vaSpread1.Col = 7: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = fpDateTime1(2).text & " " & fpDateTime1(4).text
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
Case 15 '------> impresion
'    I_DefinicionPedido
Case 18
    Me.Hide
    Unload Me
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
Case 2
    vaSpread1.Col = Col
    TipText = "Central de Compras : " & vaSpread1.text
Case 3
    vaSpread1.Col = Col
    TipText = "Año/Mes : " & Trim(vaSpread1.text)
Case 4
    vaSpread1.Col = Col
    TipText = "Semana : " & Trim(vaSpread1.text)
Case 5
    vaSpread1.Col = Col
    TipText = "Descripción : " & Trim(vaSpread1.text)
Case 6
    vaSpread1.Col = Col
    TipText = "Inicio Tope Digitación : " & Trim(vaSpread1.text)
Case 7
    vaSpread1.Col = Col
    TipText = "Term. Tope Digitación : " & Trim(vaSpread1.text)
End Select
End Sub

Private Sub vaSpread2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
If vaSpread2.MaxRows < 1 Or est Then Exit Sub
Dim i As Long
vaSpread2.Col = 1
For i = BlockRow To BlockRow2
    vaSpread2.Row = i
    If vaSpread2.Lock = False Then
       vaSpread2.Value = IIf(vaSpread2.Value = "1", "0", "1")
    Else
       Exit Sub
    End If
Next i
If Toolbar1.Buttons(12).Visible = False Then
   SSTab1.TabEnabled(0) = False
   If modo = "" Then modo = "M"
   Gl_Ac_Botones Me, 1, 0, modo
End If
End Sub

Private Sub vaSpread2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vaSpread2.MaxRows < 1 Or est Then Exit Sub
vaSpread2.Row = Row
Select Case Col
Case 1
    If Toolbar1.Buttons(12).Visible = False Then
       SSTab1.TabEnabled(0) = False
       If modo = "" Then modo = "M"
       Gl_Ac_Botones Me, 1, 0, modo
    End If
End Select
End Sub

Private Sub vaSpread2_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If vaSpread2.MaxRows < 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread2.Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
Select Case Col
Case 2
    vaSpread2.Col = Col
    TipText = "Código : " & vaSpread2.text
Case 3
    vaSpread2.Col = Col
    TipText = "Descripción : " & Trim(vaSpread2.text)
End Select
End Sub
