VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_LecVal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Lectura Vales"
   ClientHeight    =   8175
   ClientLeft      =   1395
   ClientTop       =   1575
   ClientWidth     =   15570
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   15570
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   7530
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   15345
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4575
         Left            =   225
         TabIndex        =   14
         Top             =   2400
         Width           =   15000
         _Version        =   393216
         _ExtentX        =   26458
         _ExtentY        =   8070
         _StockProps     =   64
         AutoClipboard   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   9
         SpreadDesigner  =   "I_LecVal.frx":0000
      End
      Begin VB.Frame Frame2 
         Height          =   2055
         Left            =   2595
         TabIndex        =   1
         Top             =   240
         Width           =   10095
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   5
            Top             =   360
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
            Index           =   1
            Left            =   1920
            TabIndex        =   6
            Top             =   720
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
            Left            =   1920
            TabIndex        =   7
            Top             =   1080
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
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   360
            Left            =   8985
            TabIndex        =   16
            Top             =   1440
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            ImageList       =   "ImageList1"
            DisabledImageList=   "ImageList1"
            HotImageList    =   "ImageList1"
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
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   20
            Top             =   1440
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
            Index           =   1
            Left            =   7305
            TabIndex        =   21
            Top             =   1440
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
            ButtonColor     =   -2147483637
            AutoMenu        =   -1  'True
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Final"
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
            Left            =   6120
            TabIndex        =   23
            Top             =   1500
            Width           =   1005
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio"
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
            TabIndex        =   22
            Top             =   1500
            Width           =   1065
         End
         Begin VB.Label Label3 
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
            Height          =   225
            Index           =   5
            Left            =   9000
            TabIndex        =   19
            Top             =   1155
            Width           =   840
         End
         Begin VB.Label Label3 
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
            Height          =   225
            Index           =   4
            Left            =   9000
            TabIndex        =   18
            Top             =   795
            Width           =   840
         End
         Begin VB.Label Label3 
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
            Height          =   225
            Index           =   3
            Left            =   9000
            TabIndex        =   17
            Top             =   465
            Width           =   840
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   2910
            Picture         =   "I_LecVal.frx":1A0E
            Top             =   960
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   3
            Left            =   3420
            TabIndex        =   12
            Top             =   1050
            Width           =   5235
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   2910
            Picture         =   "I_LecVal.frx":1D18
            Top             =   600
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   2
            Left            =   3420
            TabIndex        =   10
            Top             =   690
            Width           =   5235
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   2910
            Picture         =   "I_LecVal.frx":2022
            Top             =   240
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3420
            TabIndex        =   8
            Top             =   330
            Width           =   5235
         End
         Begin VB.Label Label3 
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
            Height          =   225
            Index           =   2
            Left            =   360
            TabIndex        =   4
            Top             =   1155
            Width           =   1560
         End
         Begin VB.Label Label3 
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
            Height          =   225
            Index           =   1
            Left            =   360
            TabIndex        =   3
            Top             =   795
            Width           =   1560
         End
         Begin VB.Label Label3 
            Caption         =   "Punto de Venta"
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
            Left            =   360
            TabIndex        =   2
            Top             =   465
            Width           =   1560
         End
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
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
            Left            =   3465
            TabIndex        =   9
            Top             =   375
            Width           =   5235
         End
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
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
            Index           =   2
            Left            =   3465
            TabIndex        =   11
            Top             =   735
            Width           =   5235
         End
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
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
            Index           =   3
            Left            =   3465
            TabIndex        =   13
            Top             =   1095
            Width           =   5235
         End
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   210
         TabIndex        =   24
         Top             =   7335
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Un Momento, Procesando Información"
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
         Left            =   465
         TabIndex        =   25
         Top             =   7035
         Visible         =   0   'False
         Width           =   3285
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   15570
      _ExtentX        =   27464
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "I_LecVal.frx":232C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "I_LecVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lc_Aux As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
MsgTitulo = "Reporte Lectura Vales"
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exportar Excel"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

fpDateTime1(0).DateTimeFormat = UserDefined
fpDateTime1(0).UserDefinedFormat = "dd/mm/yyyy"
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")

fpDateTime1(1).DateTimeFormat = UserDefined
fpDateTime1(1).UserDefinedFormat = "dd/mm/yyyy"
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")

EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)

vaSpread1.MaxRows = 0
vaSpread1.Row = -1
vaSpread1.Col = -1
vaSpread1.BackColor = &HC0FFFF     'Amarillo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
On Error GoTo Man_Error

If Trim(fpDateTime1(0).text) = "" Or Trim(fpDateTime1(1).text) = "" Then Exit Sub
If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then Exit Sub

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
vaSpread1.MaxRows = 0
Select Case Index
  Case 0
    RS.Open "select distinct a.ate_codatencion, a.ate_descripcion from a_pto_atencion a inner join b_detallelectura b on a.ate_codatencion = b.ate_codatencion where b.cli_codigo = '" & MuestraCasino(1) & "' and b.ate_codatencion = " & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!ate_descripcion)
    RS.Close: Set RS = Nothing
  Case 1
    RS.Open "select distinct a.reg_codigo, a.reg_nombre from a_regimen a inner join b_detallelectura b on a.reg_codigo = b.reg_codigo where b.cli_codigo = '" & MuestraCasino(1) & "' and b.ate_codatencion = " & Val(fpLongInteger1(0).Value) & " and b.reg_codigo = " & Val(fpLongInteger1(1).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "": Exit Sub
    fpayuda(2).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
  Case 2
    RS.Open "select distinct a.ser_codigo, a.ser_nombre from a_servicio a inner join b_detallelectura b on a.ser_codigo = b.ser_codigo where b.cli_codigo = '" & MuestraCasino(1) & "' and b.ate_codatencion = " & Val(fpLongInteger1(0).Value) & " and b.ser_codigo = " & Val(fpLongInteger1(2).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(3).Caption = "":  Exit Sub
    fpayuda(3).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
End Sub

Private Sub Image1_Click(Index As Integer)
On Error GoTo Man_Error

Select Case Index
Case 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = Trim(MuestraCasino(1))
    B_TabEst.LlenaDatos "a_pto_atencion", "ate_", "Punto Atencion", "PtoAte"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    fpLongInteger1(1).SetFocus
Case 2
    vg_left = fpayuda(2).Left + 2300
    vg_nombre = "": vg_codigo = Trim(MuestraCasino(1)): vg_ptoate = fpLongInteger1(0).Value
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "RegVal"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre
    fpLongInteger1(2).SetFocus
Case 3
    vg_left = fpayuda(2).Left + 2300
    vg_nombre = "": vg_codigo = Trim(MuestraCasino(1)): vg_ptoate = fpLongInteger1(0).Value
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "SerVal"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(3).Caption = vg_nombre
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error

Select Case Button.Index
Case 1
    If Trim(fpDateTime1(0).text) = "" Or Not IsDate(fpDateTime1(0).text) Then MsgBox "Fecha origen esta blanco", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Trim(fpDateTime1(1).text) = "" Or Not IsDate(fpDateTime1(1).text) Then MsgBox "Fecha destino esta blanco", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If fpDateTime1(0).Value > fpDateTime1(1).Value Then MsgBox "Fecha origen Mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If DateDiff("y", fpDateTime1(0).text, fpDateTime1(0).text) > 365 Then
       Call MsgBox("Rango De Fecha No Puede Ser Mayor a 12 Meses", vbInformation, Me.Caption)
       Exit Sub
    End If
    I_DetalleLecturaVales MuestraCasino(1), Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Format(fpDateTime1(0).text, "yyyymmdd"), Format(fpDateTime1(1).text, "yyyymmdd")
Case 3
    If vaSpread1.MaxRows > 0 Then ExportarExcel
Case 5
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i As Long
Dim Sql As String
Dim codate As Long
Dim codreg As Long
Dim cosser As Long
Dim ContadorCodBarra As Double
Dim Fecha As String
Dim CodBarra As String

If Trim(fpDateTime1(0).text) = "" Or Not IsDate(fpDateTime1(0).text) Then MsgBox "Fecha origen esta blanco", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
If Trim(fpDateTime1(1).text) = "" Or Not IsDate(fpDateTime1(1).text) Then MsgBox "Fecha destino esta blanco", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
If fpDateTime1(0).Value > fpDateTime1(1).Value Then MsgBox "Fecha origen Mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
If DateDiff("y", fpDateTime1(0).text, fpDateTime1(0).text) > 365 Then
   Call MsgBox("Rango De Fecha No Puede Ser Mayor a 12 Meses", vbInformation, Me.Caption)
   Exit Sub
End If

codate = 0
codreg = 0
codser = 0
ContadorCodBarra = 0
Fecha = ""
CodBarra = ""
Select Case Button.Index
Case 1
    
    vaSpread1.Visible = False
    vaSpread1.MaxRows = 0
    Sql = ""
    Sql = Sql & "sgp_Sel_ListarLecturaCodigoBarra "
    Sql = Sql & " '" & MuestraCasino(1) & "' "
    Sql = Sql & ", " & CLng(Val(fpLongInteger1(0).Value)) & " "
    Sql = Sql & ", " & CLng(Val(fpLongInteger1(1).Value)) & " "
    Sql = Sql & ", " & CLng(Val(fpLongInteger1(2).Value)) & " "
    Sql = Sql & ", " & Format(fpDateTime1(0).text, "yyyymmdd") & " "
    Sql = Sql & ", " & Format(fpDateTime1(1).text, "yyyymmdd") & " "
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS = vg_db.Execute(Sql)
    If RS.EOF Then
       RS.Close: Set RS = Nothing
       vaSpread1.Visible = True
       MsgBox "No existe información", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    End If
    Label1(0).Visible = True: Label1(0).Caption = "Un Momento, cargando datos": DoEvents
    Bar1(0).Min = 0: Bar1(0).Value = 0:
'    Bar1(0).max = RS.RecordCount
    Bar1(0).Visible = True
    i = 1
    Do While Not RS.EOF
       Bar1(0).Value = Val((i / RS.RecordCount) * 100)
       
       If CodBarra <> RS!cli_codigo Or RS!fechavale <> Fecha Then
          If CodBarra <> "" Then
             vaSpread1.MaxRows = vaSpread1.MaxRows + 1
             vaSpread1.Row = vaSpread1.MaxRows
             vaSpread1.Col = -1
             vaSpread1.BackColor = RGB(208, 207, 71)
             vaSpread1.Font.Bold = True
             
             vaSpread1.Col = 1
             vaSpread1.text = "Total Vales Días "
             vaSpread1.Col = 2
             vaSpread1.text = ContadorCodBarra
             ContadorCodBarra = 0
          End If
          CodBarra = RS!cli_codigo
          Fecha = RS!fechavale
       End If
       
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       vaSpread1.Row = vaSpread1.MaxRows
       
       vaSpread1.Col = 1
       vaSpread1.text = IIf(IsNull(RS!ate_codatencion), "", Trim(RS!ate_codatencion)) & " - " & IIf(IsNull(RS!ate_descripcion), "", Trim(RS!ate_descripcion))
       vaSpread1.Col = 2
       vaSpread1.text = IIf(IsNull(RS!reg_codigo), "", Trim(RS!reg_codigo)) & " - " & IIf(IsNull(RS!reg_nombre), "", Trim(RS!reg_nombre))
       vaSpread1.Col = 3
       vaSpread1.text = IIf(IsNull(RS!ser_codigo), "", Trim(RS!ser_codigo)) & " - " & IIf(IsNull(RS!ser_nombre), "", Trim(RS!ser_nombre))
       
       vaSpread1.Col = 4
       vaSpread1.text = IIf(IsNull(RS!codigobarra), "", Trim(RS!codigobarra))
       vaSpread1.Col = 5
       vaSpread1.text = IIf(IsNull(RS!per_rut), "", Trim(RS!per_rut))
       vaSpread1.Col = 6
       vaSpread1.text = IIf(IsNull(RS!per_nombre), "", Trim(RS!per_nombre))
       vaSpread1.Col = 7
       vaSpread1.text = IIf(IsNull(RS!FechaHoravale), "", Trim(RS!FechaHoravale))
       vaSpread1.Col = 8
       vaSpread1.text = IIf(IsNull(RS!cli_codigo), "", Trim(RS!cli_codigo))
       vaSpread1.Col = 9
       vaSpread1.text = IIf(IsNull(RS!cli_nombre), "", Trim(RS!cli_nombre))
        
       RS.MoveNext
       i = i + 1
       ContadorCodBarra = ContadorCodBarra + 1
    Loop
    If CodBarra <> "" Then
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       vaSpread1.Row = vaSpread1.MaxRows
       vaSpread1.Col = -1
       vaSpread1.BackColor = RGB(208, 207, 71)
       vaSpread1.Font.Bold = True
             
       vaSpread1.Col = 1
       vaSpread1.text = "Total Vales Días "
       
       vaSpread1.Col = 2
       vaSpread1.text = ContadorCodBarra
       ContadorCodBarra = 0
    End If
    
    RS.Close: Set RS = Nothing
    vaSpread1.Visible = True
    Label1(0).Visible = False
    Bar1(0).Visible = False
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
End Sub

Sub ExportarExcel()
Dim NashXl As Excel.Application
Dim iRow As Long, irow2 As Long
Dim NColumnas As Integer

On Error GoTo Ex_Error

fg_carga ""

Set NashXl = CreateObject("excel.application")
Set NashXl = New Excel.Application
NashXl.SheetsInNewWorkbook = 1
NashXl.Workbooks.Add
NashXl.Range("A1").Select
NashXl.ActiveCell.FormulaR1C1 = "Contrato          : " & MuestraCasino(1) & " - " & Trim(MuestraCasino(2))
NashXl.Range("A2").Select
NashXl.ActiveCell.FormulaR1C1 = "Punto de Atención : " & Val(fpLongInteger1(0).Value) & " - " & Trim(fpayuda(1).Caption)
NashXl.Range("A3").Select
NashXl.ActiveCell.FormulaR1C1 = "Regimen           : " & Val(fpLongInteger1(1).Value) & " - " & Trim(fpayuda(2).Caption)
NashXl.Range("A4").Select
NashXl.ActiveCell.FormulaR1C1 = "Servicio          : " & Val(fpLongInteger1(2).Value) & " - " & Trim(fpayuda(3).Caption)
NashXl.Range("A5").Select
NashXl.ActiveCell.FormulaR1C1 = "Fecha Desde       : " & Format(fpDateTime1(0).text, "dd/mm/yyyy") & " Fecha Hasta  " & Format(fpDateTime1(1).text, "dd/mm/yyyy")


MaxColumna = 10
NColumnas = 10
vaSpread1.AllowMultiBlocks = True
vaSpread1.SetSelection 1, -1, NColumnas, vaSpread1.MaxRows
vaSpread1.ClipboardCopy

iRow = vaSpread1.MaxRows
'------- Pegar vaspread1(0) - Planilla Excel
NashXl.Cells.Select
NashXl.Selection.NumberFormat = "@"

NashXl.Range("A6").Select
NashXl.ActiveSheet.Paste
'------- Asignar color
'------- Colorear titulo
NashXl.Range("A6:I6").Select
With NashXl.Selection.Interior
     .ColorIndex = 15
     .Pattern = xlSolid
End With
'------- Dibujar marco
NashXl.Range("A6:I" & iRow + 6).Select
NashXl.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
NashXl.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With NashXl.Selection.Borders(xlEdgeLeft)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeTop)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeBottom)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeRight)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideVertical)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideHorizontal)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
NashXl.Range("E1" & ":" & "E" & iRow + 6).Select
NashXl.Selection.NumberFormat = "@"

'------- Aplicar totales
NashXl.Selection.Font.Bold = True
'------- Ajustar columna
NashXl.Cells.Select
NashXl.Cells.EntireColumn.AutoFit
vaSpread1.AllowMultiBlocks = False: vaSpread1.SetSelection 1, 0, vaSpread1.MaxCols, vaSpread1.MaxRows

fg_descarga
NashXl.Visible = True

Ex_Error:
    Resume Next

End Sub
