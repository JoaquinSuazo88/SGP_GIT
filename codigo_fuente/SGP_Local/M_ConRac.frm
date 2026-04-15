VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_ConRac 
   Caption         =   "Control de Raciones"
   ClientHeight    =   8805
   ClientLeft      =   2835
   ClientTop       =   2670
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   14835
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Datos Raciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6435
      Left            =   30
      TabIndex        =   6
      Top             =   2310
      Width           =   14715
      Begin VB.Frame Frame5 
         Height          =   2895
         Left            =   5160
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   6975
         Begin VB.CommandButton Cmd1 
            Caption         =   "&Aceptar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   27
            Top             =   2280
            Width           =   1425
         End
         Begin VB.CommandButton Cmd2 
            Caption         =   "&Cancelar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   28
            Top             =   2280
            Width           =   1425
         End
         Begin EditLib.fpText Nombre 
            Height          =   315
            Index           =   1
            Left            =   2790
            TabIndex        =   25
            Top             =   1440
            Width           =   1695
            _Version        =   196608
            _ExtentX        =   2990
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
            MaxLength       =   255
            MultiLine       =   0   'False
            PasswordChar    =   "*"
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
         Begin EditLib.fpText Nombre 
            Height          =   315
            Index           =   0
            Left            =   2790
            TabIndex        =   24
            Top             =   1080
            Width           =   1695
            _Version        =   196608
            _ExtentX        =   2990
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
         Begin EditLib.fpLongInteger fpNRac 
            Height          =   315
            Left            =   2790
            TabIndex        =   26
            Top             =   1800
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha : "
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
            Left            =   240
            TabIndex        =   33
            Top             =   720
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nr. Ración"
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
            Left            =   1440
            TabIndex        =   32
            Top             =   1920
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Password"
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
            Left            =   1440
            TabIndex        =   31
            Top             =   1500
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Login"
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
            Left            =   1440
            TabIndex        =   30
            Top             =   1125
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Para actualizar comensales diarios minuta real, tiene comunicarse su monitor"
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
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   6540
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5760
         Left            =   150
         TabIndex        =   7
         Top             =   270
         Width           =   14415
         _Version        =   393216
         _ExtentX        =   25426
         _ExtentY        =   10160
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   88
         MaxRows         =   1
         SpreadDesigner  =   "M_ConRac.frx":0000
         ScrollBarTrack  =   3
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Clientes"
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
         Left            =   5190
         TabIndex        =   10
         Top             =   6150
         Width           =   690
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   4800
         Top             =   6180
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dias Bloqueados"
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
         Left            =   6645
         TabIndex        =   9
         Top             =   6135
         Width           =   1440
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008484FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   6285
         Top             =   6165
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dias Habilitados"
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
         Left            =   8850
         TabIndex        =   8
         Top             =   6135
         Width           =   1395
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   8490
         Top             =   6165
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   1620
      TabIndex        =   5
      Top             =   420
      Width           =   11685
      Begin VB.CommandButton Command2 
         Caption         =   "Exportar Vtas. Diarias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   22
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Importar Lectura Vales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   21
         Top             =   255
         Width           =   2175
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   3180
         TabIndex        =   1
         Top             =   615
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
         Left            =   3180
         TabIndex        =   2
         Top             =   975
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   3165
         TabIndex        =   0
         Top             =   255
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   3180
         TabIndex        =   3
         Top             =   1335
         Width           =   1050
         _Version        =   196608
         _ExtentX        =   1852
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
         ButtonStyle     =   2
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
         Text            =   "08/2025"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Minuta"
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
         Left            =   1770
         TabIndex        =   17
         Top             =   1410
         Width           =   1170
      End
      Begin VB.Label Label2 
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
         Index           =   3
         Left            =   1770
         TabIndex        =   16
         Top             =   1080
         Width           =   705
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
         Index           =   2
         Left            =   1770
         TabIndex        =   15
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   1770
         TabIndex        =   14
         Top             =   375
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   4410
         Picture         =   "M_ConRac.frx":2067
         Top             =   180
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   4410
         Picture         =   "M_ConRac.frx":2371
         Top             =   540
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   4410
         Picture         =   "M_ConRac.frx":267B
         Top             =   900
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   4845
         TabIndex        =   13
         Top             =   255
         Width           =   3855
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   4845
         TabIndex        =   12
         Top             =   615
         Width           =   3855
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   4845
         TabIndex        =   11
         Top             =   975
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   4890
         TabIndex        =   18
         Top             =   300
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   4890
         TabIndex        =   19
         Top             =   660
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   4890
         TabIndex        =   20
         Top             =   1020
         Width           =   3855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_ConRac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS         As New ADODB.Recordset
Dim RS1        As New ADODB.Recordset
Dim MsgTitulo  As String
Dim modo       As String
Dim i          As Long
Dim x          As Long
Dim v_columnas As Long
Dim est        As Boolean
Dim ciedia     As Long
Dim fecini     As Long
Dim fecfin     As Long
Dim EstCheck   As Boolean
Dim XRow       As Long
Dim xcol       As Long

Private Sub Cmd1_Click()

On Error GoTo Man_Error

Dim RS      As New ADODB.Recordset
Dim Fecha   As Long
Dim Nracion As Long

    '-------> Validar usuario
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT * FROM a_param WHERE par_valor = '" & LimpiaDato(Trim(Nombre(0).text)) & "' AND par_codigo = 'usulimbas' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    If RS.EOF Then
       
       MsgBox "Usuario no existe..."
       RS.Close
       Set RS = Nothing
       Nombre(0).text = ""
       Nombre(0).SetFocus
       Exit Sub
    
    End If
    
    RS.Close
    Set RS = Nothing
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT * FROM a_param WHERE par_codigo = 'parcomdia' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    
    If Not RS.EOF And UCase(Nombre(1).text) <> UCase(fg_Desencripta(TipoDato(RS!par_valor, ""))) Then
       
       MsgBox "La clave no corresponde al login..."
       RS.Close
       Set RS = Nothing
       Nombre(0).text = ""
       Nombre(0).SetFocus
       Exit Sub
    
    End If
    
    RS.Close
    Set RS = Nothing
    
    '-------> Validar que racion sea mayor que cero
    If Val(fpNRac.Value) <= 0 Then
    
       MsgBox "Ración debe ser mayor que cero..., proceso cancelado"
       Exit Sub
       
    End If
    
    Frame5.Visible = False
    Nombre(0).text = ""
    Nombre(1).text = ""
    
    vaSpread1.Row = 0
    vaSpread1.Col = xcol
    Fecha = Val(Format(Right(vaSpread1.text, 10), "yyyymmdd"))
        
    Nracion = fpNRac.Value
    
    '-------> Grabar de racion del dia b_minuta
    vg_db.Execute ("update b_minuta set min_racrea= " & Nracion & " where min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                   "AND min_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
                   "AND min_codser = " & Val(fpLongInteger1(2).Value) & " " & _
                   "AND min_fecmin = " & Fecha & "")
    
    
    '-------> Grabar de racion del dia b_minutaraciones
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("SELECT 1 FROM b_minutaraciones WHERE mir_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                           "AND mir_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
                           "AND mir_codser = " & Val(fpLongInteger1(2).Value) & " " & _
                           "AND mir_fecmin = " & Fecha & " " & _
                           "AND mir_rutcli = 'PRODUCIDAS'")
    
    If Not RS.EOF Then
    
       '-------> Grabar de racion del dia b_minuta
       vg_db.Execute ("update b_minutaraciones set  mir_nrorac = " & Nracion & " where mir_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                      "AND mir_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
                      "AND mir_codser = " & Val(fpLongInteger1(2).Value) & " " & _
                      "AND mir_fecmin = " & Fecha & " AND mir_rutcli = 'PRODUCIDAS'")
       
    Else
    
     vg_db.Execute ("INSERT  INTO dbo.b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli) " & _
                    "VALUES  ('" & LimpiaDato(Trim(fpText.text)) & "', " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & Fecha & ", 'PRODUCIDAS', " & Nracion & ", NULL, '')")
    
    End If
    
    RS.Close
    Set RS = Nothing
    
    vaSpread1.Row = XRow
    vaSpread1.Col = xcol
    vaSpread1.text = fpNRac.Value
    
    Toolbar1.Enabled = True
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame1.Enabled = True
    Toolbar1.Enabled = True
    vaSpread1.Enabled = True

    
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
    
End Sub

Private Sub Cmd2_Click()

On Error GoTo Man_Error

vaSpread1.text = "0"

Frame5.Visible = False

Frame1.Enabled = True
Toolbar1.Enabled = True
vaSpread1.Enabled = True

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim sql1 As String
Dim sql2 As String
Dim sql3 As String
Dim Fecha As Long
Dim numrac As Long
Dim codreg As Long
Dim codser As Long
Dim cbar_largo As Long
Dim cbar_posinicial As Long
Fecha = 0
numrac = 0
sql1 = IIf(vg_tipbase = "1", " format(a.fechahoravale, 'yyyymmdd') fecha  ", " CONVERT(VARCHAR(8), a.fechahoravale, 112) fecha ")
sql2 = IIf(vg_tipbase = "1", " format(a.fechahoravale, 'yyyymmdd') ", " CONVERT(VARCHAR(6), a.fechahoravale, 112) ")
sql3 = IIf(vg_tipbase = "1", " format(a.fechahoravale, 'yyyymmdd') ", " CONVERT(VARCHAR(8), a.fechahoravale, 112) ")
'-------> Importar lectura vales
'----->Validar si periodo esta abierto
If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), 0, 0) Then modo = "E": MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
'----->Traer contabilizado vales
fg_carga ""

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_DetalleLecturaxPeriodo '" & MuestraCasino(1) & "', " & Format(fpDateTime1.text, "yyyymm") & "")
Do While Not RS.EOF
    '-------> borrar minuta raciones
    vg_db.Execute ("sgp_Del_MinutaRaciones '" & MuestraCasino(1) & "', " & RS!reg_codigo & ", " & RS!ser_codigo & ", '" & RS!cli_codigo_rutcliente & "', " & RS!Fecha & "")
    '-------> Insertar datos
    vg_db.Execute ("sgp_Ins_MinutaRaciones '" & MuestraCasino(1) & "', " & RS!reg_codigo & ", " & RS!ser_codigo & ", " & RS!Fecha & ", '" & RS!cli_codigo_rutcliente & "', " & RS!numvale & "")
    numrac = numrac + RS!numvale
    RS.MoveNext
Loop
RS.Close: Set RS = Nothing

fg_descarga

If numrac > 0 Then
   
   MsgBox "Actualización raciones finalizada sin problema...", vbExclamation + vbOKOnly, MsgTitulo

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Command2_Click()

E_ImportarVentasDiarias.Show 1, Me

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

Me.Height = 9210
Me.Width = 14955

est = False
EstCheck = False

fg_centra Me
MsgTitulo = "Control de Raciones"
Me.HelpContextID = vg_OpcM
modo = "": vaSpread1.MaxRows = 0
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 3, 1), modo
OpGr = False: vaSpread1.MaxRows = 0
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
fpDateTime1.text = Format(Date, "mm/yyyy")
GenerarTitulo

End Sub

Private Sub fpDateTime1_Change()

If est Then Exit Sub
vaSpread1.MaxRows = 0: Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 3, 1), modo: est = True
MoverDatos
est = False

End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
MoverDatos

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

vaSpread1.MaxRows = 0: Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 3, 1), modo
fpayuda(Index).Caption = ""

Select Case Index

Case 1
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS.Open RutinaLectura.Regimen(2, Val(fpLongInteger1(1).Value), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    MoverDatos

Case 2
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open RutinaLectura.Servicio(2, Val(fpLongInteger1(2).Value), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "": Exit Sub
    fpayuda(2).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    MoverDatos

End Select

End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
If Trim(fpLongInteger1(Index).text) = "" Or Val(fpLongInteger1(Index).Value) < 1 Then fpLongInteger1(Index).text = ""
SendKeys "{Tab}"

End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Select Case KeyCode
  
  Case 120
    
    If Index = 1 Then Image1_Click 1
    If Index = 2 Then Image1_Click 2

End Select

End Sub

Private Sub fpNRac_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub fpText_Change()

If fpText.text = "" Then fpayuda(0).Caption = "": Exit Sub

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText.text)), ""), vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": fpLongInteger1(1).Value = "": fpayuda(1).Caption = "": fpLongInteger1(2).Value = "": fpayuda(2).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
fpLongInteger1(2).Value = "": fpayuda(2).Caption = ""
'MoverDatos

End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

Case 120
    
    Image1_Click 0

End Select

End Sub

Private Sub Image1_Click(Index As Integer)

Select Case Index

Case 0
    
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.text = vg_codigo
    fpayuda(0).Caption = vg_nombre
    fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(2).Value = "": fpayuda(2).Caption = ""
    fpLongInteger1(1).SetFocus

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
'    fpDateTime1.SetFocus

Case 3
    
    If fpText.text = "" Or Val(fpLongInteger1(1).Value) < 1 Or Val(fpLongInteger1(2).Value) < 0 Or fpDateTime1.text = "" Then Exit Sub
    B_HistPm.LlenarHistPlan "Histórico Estructura Fija", fpText.text, fpLongInteger1(1).text & "|" & fpLongInteger1(2).text & "|", 3
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    fpDateTime1.text = vg_codigo
'    accion = False: Combo1(0).ListIndex = vg_auxfecha - 1: accion = True
    MoverDatos

End Select

End Sub

Private Sub Nombre_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim Fecha      As Long
Dim RS         As New ADODB.Recordset
Dim RS1        As New ADODB.Recordset
Dim EstGrabado As Integer

With vaSpread1
    
    Select Case Button.Index
    
    Case 1 '-------> Agregar registro
        
        GenerarTitulo
        Dim auxrutcli As String
        auxrutcli = ""
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS.Open "SELECT DISTINCT a.cli_codigo, a.cli_nombre, Max(b.prv_fecvig) AS prv_fecvig " & _
                "FROM b_clientes a, b_preciovta b, b_minuta c " & _
                "WHERE b.prv_cencos = c.min_cencos " & _
                "AND   b.prv_codreg = c.min_codreg " & _
                "AND   b.prv_codser = c.min_codser " & _
                "AND   c.min_fecmin >= b.prv_fecvig " & _
                "AND   b.prv_rutcli = a.cli_codigo " & _
                "AND   c.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                "AND   c.min_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
                "AND   c.min_codser = " & Val(fpLongInteger1(2).Value) & " " & _
                "AND   c.min_fecmin >= " & fecini & " AND c.min_fecmin <= " & fecfin & " AND a.cli_tipo = 1 AND a.cli_activo = '1' " & _
                "GROUP BY a.cli_codigo, a.cli_nombre ORDER BY a.cli_codigo", vg_db, adOpenStatic
        
        If Not RS.EOF Then
           
           .MaxRows = .MaxRows + 1
           .Row = .MaxRows
           .Col = 1: .text = ""
           .Font.Bold = True
           .Col = 2
           .CellType = CellTypeStaticText
           .Font.Size = 9:
           .TypeHAlign = TypeHAlignCenter
           .text = " "
           .Font.Bold = True
           
           For i = 3 To .MaxCols
           
               .Col = i
               ' Set cell background color
               ' light gray, RGB(192, 192, 192)
               .BackColor = &HC0C0C0
               ' Define cell type as check box
               .CellType = CellTypeCheckBox
               ' Define the check box text
               .TypeCheckText = "Facturable"
               ' Align text left of graphic
               .TypeCheckTextAlign = TypeCheckTextAlignLeft
               .TypeHAlign = TypeHAlignCenter
               ' Set the column width
               .ColWidth(i) = 11
               
           Next i
        
           Do While Not RS.EOF
              
              If RS!cli_codigo <> auxrutcli Then
                 
                 .MaxRows = .MaxRows + 1
                 .Row = .MaxRows
                 .Col = 1
                 .text = fg_PintaRut(RS!cli_codigo)
                 .Col = 2
                 .text = RS!cli_nombre
                 auxrutcli = RS!cli_codigo
              
              End If
              
              RS.MoveNext
           Loop
        End If
        RS.Close: Set RS = Nothing
       
       '-------> Agregar fila totales
       .MaxRows = .MaxRows + 1: .Row = .MaxRows: .Col = 2: .Font.Bold = True: .Font.Size = 9: .text = "Total Cliente"
       .Col = -1: .BackColor = &HE0E0E0
       For i = 3 To .MaxCols
           
           .Col = i
           .Font.Bold = True
           .Font.Size = 9
           .CellType = CellTypeStaticText
           .TypeHAlign = 1
       
       Next i
       '-------> Fin agregar fila totales
       
       '-------> Agregar fila raciones personal
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1
       .text = "PERSONAL"
       .Col = 2
       .text = "PERSONAL"
       
       '-------> Agregar fila muestra referencia
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows: .Col = 1
       .text = "MUESTRA R"
       .Col = 2
       .text = "MUESTRA REFERENCIA"
       
       '-------> Agregar fila raciones producidas
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1
       .Lock = True
       .text = "PRODUCIDAS"
       .Col = 2
       .Lock = True
       .text = "PRODUCIDAS"
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS.Open "SELECT DISTINCT min_fecmin, min_racrea " & _
               "FROM b_minuta " & _
               "WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
               "AND   min_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
               "AND   min_codser = " & Val(fpLongInteger1(2).Value) & " " & _
               "AND   min_fecmin >= " & fecini & " AND min_fecmin <= " & fecfin & "", vg_db, adOpenStatic
       If Not RS.EOF Then
          
          Do While Not RS.EOF
             
             .Row = 0
             
             For i = 3 To .MaxCols
                 
                 .Col = i
                 If Val(Format(Right(.text, 10), "yyyymmdd")) = RS!min_fecmin Then j = i: Exit For
             
             Next i
             
             .Row = .MaxRows
             .Col = j
             .Lock = True
             
             .text = IIf(IsNull(RS!min_racrea) Or RS!min_racrea < 1, "", RS!min_racrea)
             
             RS.MoveNext
          
          Loop
       
       End If
       RS.Close: Set RS = Nothing
      
       '-------> Bloquea días de cierre en color rojo
       Dim diablq As Date
       If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Date, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
'       If Format(Date, "dd/mm/yyyy") > diablq Or Format(CDate(fpDateTime1.text), "mm/yyyy") < Format(Month(Date) - 1 & "/" & Year(Date), "mm/yyyy") Then
        If CierrePeriodo(Format(fpDateTime1.text, "yyyyMM"), 0, 34) Or CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), 0, 35) Then
          If Month(Date) = 1 Then
             v_columnas = ((dEoM(Format("01/" & "12/" & Year(Date) - 1, "dd/mm/yyyy")) - CDate(CDate("01/" + Mid(fpDateTime1.text, 1, 8))) + 1) * 2) - 1
          Else
'             v_columnas = ((dEoM(Format("01/" & (Month(Date) - 1) & "/" & Year(Date), "dd/mm/yyyy")) - CDate(CDate("01/" + Mid(fpDateTime1.text, 1, 8))) + 1) * 2) - 1
             v_columnas = (((CDate(vg_ciedia) - 1) - CDate(CDate("01/" + Mid(fpDateTime1.text, 1, 8))) + 1))
          End If
       Else
          v_columnas = 0
       End If
            
       If v_columnas > 0 Then
          
          .Row = -1
          
          For i = 3 To v_columnas + 2
              
              .Col = i
              .Lock = True
              .BackColor = Shape1(0).FillColor
          
          Next i
          
          .SetActiveCell i, 1
       
       End If
       '-------> Fin Bloqueo de celdas
       
       If ciedia > 0 Then
          
          .Row = 0
          .Col = 3
          
          If Format(Date, "dd/mm/yyyy") > diablq Or Format(CDate(Right(.text, 10)), "mm/yyyy") < Format(Month(Date) - 1 & "/" & Year(Date), "mm/yyyy") Then
             
             If Month(Date) = 1 Then
                
                v_columnas = ((dEoM(Format(ciedia + 1 & "/" & "12/" & Year(Date) - 1)) - CDate(CDate(Right(.text, 10))) + 1))
             
             Else
                
                v_columnas = ((dEoM(Format(ciedia + 1 & " /" & (Month(Date) - 1) & "/" & Year(Date))) - CDate(CDate(Right(.text, 10))) + 1))
             
             End If
          
          Else
             
             v_columnas = 0
          
          End If
          
          .Row = -1
          
          For i = 3 To v_columnas + 2
              
              .Col = i
              .Lock = True
              .BackColor = Shape1(0).FillColor
          
          Next i
       
       End If
       
       .Row = 1
       .Col = .MaxCols
       If .BackColor <> Shape1(0).FillColor Then
       
           modo = "A":
           Gl_Ac_Botones Me, 1, 0, modo
       
       End If
       
       .SetActiveCell 1, 1
       If v_columnas < .MaxCols Then
       
          fpLongInteger1(1).Enabled = False
          Image1(1).Enabled = False
          fpLongInteger1(2).Enabled = False
          fpDateTime1.Enabled = False
          Image1(2).Enabled = False
          
       End If
    
    Case 3 '-------> Activar modo modificación
        
        If .MaxRows < 1 Then Exit Sub
        .Row = 1: .Col = .MaxCols
        If .BackColor <> Shape1(0).FillColor Then modo = "M": Gl_Ac_Botones Me, 1, 0, modo Else Exit Sub
        fpLongInteger1(1).Enabled = False: Image1(1).Enabled = False: fpLongInteger1(2).Enabled = False: fpDateTime1.Enabled = False: Image1(2).Enabled = False
    
    Case 5 '-------> Borrar información
        
        If .ActiveRow < 1 Then MsgBox "No existe información a borrar...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS.Open "select DISTINCT isnull(mir_SPRS,'') as mir_SPRS FROM b_minutaraciones " & _
                "WHERE mir_cencos='" & fpText.text & "' " & _
                "AND mir_codreg=" & Val(fpLongInteger1(1).Value) & " " & _
                "AND mir_codser=" & Val(fpLongInteger1(2).Value) & " " & _
                "AND mir_fecmin>=" & fecini & " AND mir_fecmin<=" & fecfin & "", vg_db, adOpenForwardOnly
         
        If Not RS.EOF Then
        
           Do While Not RS.EOF
           
              If RS!mir_SPRS = "1" Then
              
                 MsgBox "Existen información integración SPRS, no puede ser borrado el periodo...", vbExclamation + vbOKOnly, MsgTitulo
                 
                 RS.Close: Set RS = Nothing
                 
                 Exit Sub
                 
              End If
              
              RS.MoveNext
        
           Loop
           RS.Close: Set RS = Nothing

        End If
        
        If MsgBox("Elimina Documento...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        
        EstGrabado = 1
        vg_db.BeginTrans
        vg_db.Execute "DELETE b_minutaraciones FROM b_minutaraciones WHERE mir_cencos='" & fpText.text & "' AND mir_codreg=" & Val(fpLongInteger1(1).Value) & " AND mir_codser=" & Val(fpLongInteger1(2).Value) & " AND mir_fecmin>=" & fecini & " AND mir_fecmin<=" & fecfin & ""
        vg_db.CommitTrans
        EstGrabado = 0
        
        If RS1.State = 1 Then
           RS1.Close
        End If
        
        Set RS1 = vg_db.Execute("sgp_Del_MinutaRacionFacturable '" & fpText.text & "', " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & fecini & ", " & fecfin & "")
        
        If Not RS1.EOF Then
       
           If RS1(0) > 0 Then
          
              MsgBox RS1(0) & " " & RS1(1), vbCritical + vbOKOnly, Me.Caption
       
           End If
    
        End If
        RS1.Close
        Set RS1 = Nothing
        
        .MaxRows = 0
        modo = "": Gl_Ac_Botones Me, 1, 3, modo
    
    Case 7 '-------> Actualizar lista
        
        MoverDatos
    
    Case 10 '-------> Cancelar
        
        If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        If modo = "A" Then
           .MaxRows = 0
        Else
           MoverDatos
        End If
        modo = "": Gl_Ac_Botones Me, 1, IIf(.MaxRows = 0, 2, 4), modo
        fpLongInteger1(1).Enabled = True: Image1(1).Enabled = True: fpLongInteger1(2).Enabled = True: fpDateTime1.Enabled = True: Image1(2).Enabled = True
    
    Case 12 '-------> Grabar información
        
        fg_carga ""
        
        Dim rutcli As String
        Dim nrorac As Long
        Dim Fac    As Integer
        
        rutcli = ""
        If fpText.text = "" Or Val(fpLongInteger1(1).Value) < 1 Or Val(fpLongInteger1(2).Value) < 1 Or fpDateTime1.text = "" Then Exit Sub
        If modo = "A" Then
           
           For i = 2 To .MaxRows
               
               .Row = i
               .Col = 1
               rutcli = Trim(fg_DespintaRut(.text))
               
               If rutcli <> "" Then
                  
                  For x = 3 To .MaxCols
                      .Row = i
                      .Col = x
                      nrorac = 0
                      
                      If Trim(.text) <> "" Then
                         
                         nrorac = Val(.text)
                         .Row = 0
                         .Col = x
                         EstGrabado = 1
                         
                         vg_db.BeginTrans
                         
                         vg_db.Execute "INSERT INTO b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac) SELECT DISTINCT min_cencos, min_codreg, min_codser, min_fecmin, '" & IIf(vg_tipbase = "1", Trim(rutcli), LTrim(rutcli)) & "', " & nrorac & " FROM b_minuta WHERE min_codigo IN (SELECT mid_codigo FROM b_minutadet WHERE mid_tipmin = '2') AND min_cencos = '" & fpText.text & "' AND min_codreg = " & Val(fpLongInteger1(1).Value) & " AND min_codser = " & Val(fpLongInteger1(2).Value) & " AND min_fecmin = " & Val(Format(Right(.text, 10), "yyyymmdd")) & ""
                         
                         '-------> Actualizar raciones planificación real
                         vg_db.Execute "UPDATE b_minuta SET min_racrea = " & nrorac & " WHERE b_minuta.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                                       "AND b_minuta.min_codreg = " & Val(fpLongInteger1(1).Value) & " AND b_minuta.min_codser = " & Val(fpLongInteger1(2).Value) & " AND b_minuta.min_fecmin = " & Val(Format(Right(.text, 10), "yyyymmdd")) & " AND '" & Trim(rutcli) & "' = 'PRODUCIDAS' AND " & nrorac & " > 0"
                         
                         vg_db.CommitTrans
                      
                         EstGrabado = 0
                         
                      End If
                                      
                       '-------> Grabar minuta raciones facturadas
                      .Row = 1
                      .Col = x
                      Fac = IIf(.text = "1", 1, 0)
                      .Row = 0
                      Fecha = Val(Format(Right(.text, 10), "yyyymmdd"))
                      
                      If RS1.State = 1 Then
                         RS1.Close
                      End If
                      
                      Set RS1 = vg_db.Execute("sgp_Ins_MinutaRacionFacturable '" & LimpiaDato(Trim(fpText.text)) & "', " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & Fecha & ", '" & Fac & "'")
                  
                      If Not RS1.EOF Then
       
                         If RS1(0) > 0 Then
          
                            MsgBox RS1(0) & " " & RS1(1), vbCritical + vbOKOnly, Me.Caption
       
                         End If
    
                      End If
                      RS1.Close
                      Set RS1 = Nothing
                  
                  Next x
               
               End If
           
           Next i
           
           modo = "M"
        Else
           For i = 2 To .MaxRows
               .Row = i
               .Col = 1
               rutcli = Trim(fg_DespintaRut(.text))
               
               If rutcli <> "" Then
                  
                  For x = 3 To .MaxCols
                      
                      .Row = i
                      .Col = x
                      nrorac = 0
                      nrorac = Val(.text)
                      .Row = 0
                      .Col = x
                      
                      If RS.State = 1 Then RS.Close
                      RS.CursorLocation = adUseClient
                      vg_db.CursorLocation = adUseClient
                      
                      RS.Open "SELECT b.mir_nrorac " & _
                              "FROM   b_clientes a, b_minutaraciones b " & _
                              "WHERE  b.mir_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                              "AND    b.mir_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
                              "AND    b.mir_codser = " & Val(fpLongInteger1(2).Value) & " " & _
                              "AND    b.mir_rutcli = '" & rutcli & "' " & _
                              "AND    b.mir_fecmin = " & Val(Format(Right(.text, 10), "yyyymmdd")) & "", vg_db, adOpenStatic
                      
                      If RS.EOF And nrorac > 0 Then
                         
                         EstGrabado = 1
                         vg_db.BeginTrans
                         vg_db.Execute "INSERT INTO b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac) SELECT DISTINCT min_cencos, min_codreg, min_codser, min_fecmin, '" & IIf(vg_tipbase = "1", Trim(rutcli), LTrim(rutcli)) & "', " & nrorac & " FROM b_minuta WHERE min_codigo IN (SELECT mid_codigo FROM b_minutadet WHERE mid_tipmin = '2') AND min_cencos = '" & Trim(fpText.text) & "' AND min_codreg = " & Val(fpLongInteger1(1).Value) & " AND min_codser = " & Val(fpLongInteger1(2).Value) & " AND min_fecmin = " & Val(Format(Right(.text, 10), "yyyymmdd")) & ""
                         vg_db.CommitTrans
                         EstGrabado = 0
                      
                      ElseIf Not RS.EOF Then
                         If RS!mir_nrorac <> nrorac Then
                            
                            EstGrabado = 1
                            vg_db.BeginTrans
                            vg_db.Execute "UPDATE b_minutaraciones SET mir_nrorac = " & nrorac & " WHERE mir_cencos = '" & Trim(fpText.text) & "' AND mir_codreg = " & Val(fpLongInteger1(1).Value) & " AND mir_codser = " & Val(fpLongInteger1(2).Value) & " AND mir_fecmin = " & Val(Format(Right(.text, 10), "yyyymmdd")) & " AND mir_rutcli = '" & rutcli & "'"
                            '-------> Actualizar raciones planificación real
                            vg_db.Execute "UPDATE b_minuta SET min_racrea = " & nrorac & " WHERE b_minuta.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                                          "AND b_minuta.min_codreg = " & Val(fpLongInteger1(1).Value) & " AND b_minuta.min_codser = " & Val(fpLongInteger1(2).Value) & " AND b_minuta.min_fecmin = " & Val(Format(Right(.text, 10), "yyyymmdd")) & " AND '" & Trim(rutcli) & "' = 'PRODUCIDAS'"
                            vg_db.CommitTrans
                            EstGrabado = 0
                         
                         End If
                      End If
                      RS.Close: Set RS = Nothing
                      
                       '-------> Grabar minuta raciones facturadas
                      .Row = 1
                      .Col = x
                      Fac = IIf(.text = "1", 1, 0)
                      .Row = 0
                      Fecha = Val(Format(Right(.text, 10), "yyyymmdd"))
                      
                      If RS1.State = 1 Then
                         
                         RS1.Close
                      
                      End If
                      
                      Set RS1 = vg_db.Execute("sgp_Ins_MinutaRacionFacturable '" & LimpiaDato(Trim(fpText.text)) & "', " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & Fecha & ", '" & Fac & "'")
                      
                      If Not RS1.EOF Then
       
                         If RS1(0) > 0 Then
          
                            MsgBox RS1(0) & " " & RS1(1), vbCritical + vbOKOnly, Me.Caption
       
                         End If
    
                      End If
                      RS1.Close
                      Set RS1 = Nothing
                  
                  Next x
               
               End If
           
           Next i
           modo = "M"
        
        End If
        modo = "": Gl_Ac_Botones Me, 1, IIf(.MaxRows = 0, 2, 4), modo
        fpLongInteger1(1).Enabled = True
        Image1(1).Enabled = True
        fpLongInteger1(2).Enabled = True
        fpDateTime1.Enabled = True
        Image1(2).Enabled = True
        
        fg_descarga
    
    Case 15 '-------> Imprimir
        
        If .MaxRows < 1 Then Exit Sub
        I_ConRac fpText.text, Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), fpDateTime1.text
    
    Case 18 '-------> Salir
        
        Me.Hide
        Unload Me
    
    End Select

End With

Exit Sub
Man_Error:
If Err = -2147467259 Then
   
   If EstGrabado = 1 Then
      
      vg_db.RollbackTrans
   
   End If
   
   MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
   Exit Sub

End If

If Err = 3034 Then
   
   If EstGrabado = 1 Then
      
      vg_db.RollbackTrans
   
   End If
   
   Exit Sub

End If
If EstGrabado = 1 Then
   
   vg_db.RollbackTrans

End If
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Sub GenerarTitulo()

Dim auxdia As Long, diafin As Long, auxma As String

With vaSpread1
    '-------> Traer día cierre
    ciedia = 0
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    RS.Open "SELECT MIN(cli_ciedia) AS cli_ciedia FROM b_clientes WHERE cli_tipo=1 AND cli_cievta='2' AND cli_activo='1'", vg_db, adOpenStatic
    If Not RS.EOF Then ciedia = IIf(IsNull(RS!cli_ciedia) Or Trim(RS!cli_ciedia) = "", 0, RS!cli_ciedia)
    RS.Close: Set RS = Nothing
    .MaxRows = 0
    If ciedia <> 0 Then
       
       auxdia = Left(BoM("01/" & fpDateTime1.text), 2)
       diafin = auxdia
       auxdia = (auxdia - (ciedia + 1))
       auxma = Mid(BoM("01/" & fpDateTime1.text), 4, 10)
       .MaxCols = 2 + Left(dEoM("01/" & fpDateTime1.text), 2) + auxdia + 1
       fecini = Format(auxma, "yyyymm") & fg_pone_cero(ciedia + 1, 2)
    
    Else
       
       .MaxCols = 2 + Left(dEoM("01/" & fpDateTime1.text), 2)
       auxma = fpDateTime1.text
       fecini = Format(fpDateTime1.text, "yyyymm") & fg_pone_cero(1, 2)
    
    End If
    
    fecfin = Format(dEoM("01/" & fpDateTime1.text), "yyyymmdd")
    x = IIf(ciedia = 0, 1, IIf(Format(ciedia & "/" & auxma, "dd/mm/yyyy") < Format(dEoM("01/" & auxma), "dd/mm/yyyy"), ciedia + 1, Format(dEoM("01/" & auxma), "dd")))
    For i = 3 To .MaxCols
        
        .Row = 0
        .Col = i
        .text = Trim(fg_Fecha_Dia1(Format(Mid(auxma, 4, 4) & "/" & Mid(auxma, 1, 2) & "/" & fg_pone_cero(Str(x), 2), "yyyymmdd"), 2) & "/" & auxma)
    '       .text = fg_Fecha_Dia(Format(Mid(fpDateTime1.text, 4, 4) & "/" & Mid(fpDateTime1.text, 1, 2) & "/" & fg_pone_cero(Str(X), 2), "yyyymmdd"), 2) & "/" & fpDateTime1.text
        x = x + 1
        
        If x > diafin And ciedia <> 0 Then
           
           auxma = fpDateTime1.text
           diafin = Left(dEoM("01/" & fpDateTime1.text), 2): x = 1
        
        End If
    
    Next i
    .Col = -1: .Row = -1
    .BackColor = Shape1(1).FillColor
    .Lock = False
    .Row = -1
    .Col = 1: .BackColor = Shape1(2).FillColor: .Col = 2: .BackColor = Shape1(2).FillColor

End With

End Sub

Sub MoverDatos()

Dim RS     As New ADODB.Recordset
Dim sql1   As String
Dim sql2   As String
Dim sql3   As String
Dim indper As Boolean
Dim indpro As Boolean
Dim aAp    As String
Dim rutcli As String
Dim ret    As Variant
Dim i      As Long
Dim j      As Long
Dim x      As Long
Dim Sql    As String

sql1 = IIf(vg_tipbase = "1", " format(a.fechahoravale, 'yyyymmdd') fecha  ", " CONVERT(VARCHAR(8), a.fechahoravale, 112) fecha ")
sql2 = IIf(vg_tipbase = "1", " format(a.fechahoravale, 'yyyymmdd') ", " CONVERT(VARCHAR(6), a.fechahoravale, 112) ")
sql3 = IIf(vg_tipbase = "1", " format(a.fechahoravale, 'yyyymmdd') ", " CONVERT(VARCHAR(8), a.fechahoravale, 112) ")

If fpText.text = "" Or Val(fpLongInteger1(1).Value) < 1 Or Val(fpLongInteger1(2).Value) < 1 Or fpDateTime1.text = "" Then GenerarTitulo: Exit Sub
With vaSpread1
    .MaxRows = 0: indper = False: indpro = False
    .Visible = False
    '-------> Rutina generar titulo
    GenerarTitulo
    Dim auxrutcli As String
    '-----------> Validar si existe planificación real
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT COUNT(a.mid_codigo) AS nreg FROM b_minutadet a, b_minuta b " & _
            "WHERE b.min_codigo = a.mid_codigo " & _
            "AND   b.min_cencos = '" & fpText.text & "' " & _
            "AND   b.min_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
            "AND   b.min_codser = " & Val(fpLongInteger1(2).Value) & " " & _
            "AND   b.min_fecmin >= " & fecini & " AND b.min_fecmin <= " & fecfin & " AND a.mid_tipmin = '2'", vg_db, adOpenForwardOnly
    If RS!nreg = 0 Then RS.Close: Set RS = Nothing: .Visible = True: Exit Sub
    RS.Close: Set RS = Nothing
    '-----------> Fin validar si existe planificación real
    
    auxrutcli = ""
   
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT DISTINCT a.mir_fecmin, a.mir_rutcli, a.mir_nrorac " & _
            "FROM   b_minutaraciones a, b_clientes b " & _
            "WHERE (a.mir_rutcli = b.cli_codigo OR a.mir_rutcli IN ('PERSONAL','PRODUCIDAS','MUESTRA R')) AND (b.cli_activo = '1' OR a.mir_nrorac > 0) " & _
            "AND    a.mir_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
            "AND    a.mir_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
            "AND    a.mir_codser = " & Val(fpLongInteger1(2).Value) & " " & _
            "AND    a.mir_fecmin >= " & fecini & " AND mir_fecmin <= " & fecfin & " ORDER BY mir_rutcli", vg_db, adOpenForwardOnly
    If Not RS.EOF Then
       
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       '-------> Insert tabla Proveedores
        RS1.Open "SELECT DISTINCT a.cli_codigo, a.cli_nombre, a.cli_activo, Max(b.prv_fecvig) AS prv_fecvig " & _
                 "FROM b_clientes a, b_preciovta b, b_minuta c " & _
                 "WHERE b.prv_cencos = c.min_cencos " & _
                 "AND   b.prv_codreg = c.min_codreg " & _
                 "AND   b.prv_codser = c.min_codser " & _
                 "AND   c.min_fecmin >= b.prv_fecvig " & _
                 "AND   (b.prv_rutcli = a.cli_codigo OR b.prv_rutcli = 'PERSONAL' OR b.prv_rutcli = 'MUESTRA R')" & _
                 "AND   c.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                 "AND   c.min_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
                 "AND   c.min_codser = " & Val(fpLongInteger1(2).Value) & " " & _
                 "AND   c.min_fecmin >= " & fecini & " AND c.min_fecmin <= " & fecfin & " AND a.cli_tipo = 1 " & _
                 "GROUP BY a.cli_codigo, a.cli_nombre, a.cli_activo ORDER BY a.cli_codigo", vg_db, adOpenForwardOnly
        If Not RS1.EOF Then
           
           .MaxRows = .MaxRows + 1
           .Row = .MaxRows
           .Col = 1: .text = ""
           .Font.Bold = True
           .Col = 2
           .CellType = CellTypeStaticText
           .Font.Size = 9:
           .TypeHAlign = TypeHAlignCenter
           .text = " "
           .Font.Bold = True
           
           For i = 3 To .MaxCols
           
               .Col = i
               ' Set cell background color
               ' light gray, RGB(192, 192, 192)
               .BackColor = &HC0C0C0
               ' Define cell type as check box
               .CellType = CellTypeCheckBox
               ' Define the check box text
               .TypeCheckText = "Facturable"
               ' Align text left of graphic
               .TypeCheckTextAlign = TypeCheckTextAlignLeft
               .TypeHAlign = TypeHAlignCenter
               ' Set the column width
               .ColWidth(i) = 11
               
           Next i
        
        End If
        
        Do While Not RS1.EOF
           
           If RS1!cli_codigo <> auxrutcli Then
              
              .MaxRows = .MaxRows + 1
              .Row = .MaxRows
              .Col = 1: .text = fg_PintaRut(RS1!cli_codigo)
              .Col = 2: .text = Trim(RS1!cli_nombre)
              auxrutcli = Trim(RS1!cli_codigo)
           
           End If
           
           RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
       
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS1.Open "SELECT DISTINCT a.mir_fecmin, a.mir_rutcli, a.mir_nrorac, isnull(a.mir_SPRS,'') as mir_SPRS " & _
                 "FROM   b_minutaraciones a, b_clientes b " & _
                 "WHERE (a.mir_rutcli = b.cli_codigo) AND (b.cli_activo = '1' OR a.mir_nrorac > 0) " & _
                 "AND    a.mir_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                 "AND    a.mir_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
                 "AND    a.mir_codser = " & Val(fpLongInteger1(2).Value) & " " & _
                 "AND    a.mir_fecmin >= " & fecini & " AND mir_fecmin <= " & fecfin & " ORDER BY mir_rutcli", vg_db, adOpenForwardOnly
        
        Do While Not RS1.EOF
           
           .Row = 0
          
           For i = 3 To .MaxCols
 
               .Col = i
 
               If Val(Format(Right(.text, 10), "yyyymmdd")) = RS1!mir_fecmin Then
 
                  j = i
                  Exit For
 
               End If
 
           Next i
           
           For i = 2 To .MaxRows

               .Row = i
               .Col = 1
               rutcli = fg_PintaRut(RS1!mir_rutcli)

               If .text = rutcli Then

                  .Row = i
                  .Col = j
                  If RS1!mir_SPRS = "1" Then
                  
                        For x = 3 To .MaxCols
                            
                            .Col = x
                            .Lock = True
                        
                        Next x
                    
                  End If
                  
                  .Col = j
                  .text = IIf(IsNull(RS1!mir_nrorac) Or RS1!mir_nrorac < 1, "", RS1!mir_nrorac)
                  Exit For

               End If

           Next i
           
           RS1.MoveNext
        
        Loop
        RS1.Close: Set RS1 = Nothing
       
       Dim EstEnc As Boolean
       
       
       '-------> Traer raciones facturables
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Sql = ""
       Sql = Sql & " '" & LimpiaDato(Trim(fpText.text)) & "' , "
       Sql = Sql & " " & Val(fpLongInteger1(1).Value) & ", "
       Sql = Sql & " " & Val(fpLongInteger1(2).Value) & " "
       Set RS1 = vg_db.Execute("sgp_Sel_MinutaRacionesFacturable " & Sql & " ")
       Do While Not RS1.EOF
       
           .Row = 0
          
           EstEnc = False
           For i = 3 To .MaxCols
 
               .Col = i
 
               If Val(Format(Right(.text, 10), "yyyymmdd")) = RS1!mrf_fecmin Then
                
                  EstEnc = True
                  j = i
                  Exit For
 
               End If
 
           Next i
          
          If EstEnc Then
             
             vaSpread1.Row = 1
             vaSpread1.Col = j
          
             If vaSpread1.MaxRows > 0 Then
                    
                EstCheck = True
                vaSpread1.text = IIf(RS1!mrf_facturado, "1", "0")
                vaSpread1.TypeCheckText = IIf(RS1!mrf_facturado, "No Facturable", "Facturable")
                EstCheck = False
          
             End If
                    
             For i = 2 To vaSpread1.MaxRows
          
                 vaSpread1.Row = i
                 vaSpread1.Col = j
                 If vaSpread1.Lock = False Then
                    vaSpread1.Lock = IIf(RS1!mrf_facturado, True, False)
                 End If
             Next i
          
          End If
          
          RS1.MoveNext
          
       Loop
       RS1.Close: Set RS1 = Nothing
       
       '-------> Agregar fila totales
       .MaxRows = .MaxRows + 1: .Row = .MaxRows: .Col = 2: .Font.Bold = True: .Font.Size = 9: .text = "Total Cliente"
       .Col = -1: .BackColor = &HE0E0E0
       For i = 3 To .MaxCols
           .Col = i: .Font.Bold = True: .Font.Size = 9: .CellType = CellTypeStaticText: .TypeHAlign = 1
       Next i
       '-------> Fin agregar fila totales
       
       '-------> Sumar totales
       Dim totnrorac As Long
       For i = 3 To .MaxCols
           .Col = i: totnrorac = 0
           For x = 2 To .MaxRows - 1
               .Row = x
               If Trim(.text) <> "" Then totnrorac = CCur(totnrorac + .text)
           Next x
           If totnrorac > 0 Then .Row = .MaxRows: .Col = i: .text = totnrorac
       Next i
    
       '-------> Mover datos personal
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1
       .text = "PERSONAL"
       .Col = 2
       .text = "PERSONAL"
       
       '-------> Mover datos muestra referencia
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1
       .text = "MUESTRA R"
       .Col = 2
       .text = "MUESTRA REFERENCIA"
       
       '------> Bloquear raciones producidas
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = -1
       .Lock = True
       .Col = 1
       .text = "PRODUCIDAS"
       .Lock = True
       .Col = 2
       .text = "PRODUCIDAS"
       
       auxrutcli = ""
       Do While Not RS.EOF
          
          If Trim(RS!mir_rutcli) = "PERSONAL" Or Trim(RS!mir_rutcli) = "PRODUCIDAS" Or Trim(RS!mir_rutcli) = "MUESTRA R" Then
             
             .Row = IIf(Trim(RS!mir_rutcli) = "PRODUCIDAS", (.MaxRows), IIf(Trim(RS!mir_rutcli) = "MUESTRA R", .MaxRows - 1, .MaxRows - 2))
             .Row = 0
             For i = 3 To .MaxCols
             
                 .Col = i
                 If Val(Format(Right(.text, 10), "yyyymmdd")) = RS!mir_fecmin Then j = i: Exit For
             
             Next i
             
             .Row = IIf(Trim(RS!mir_rutcli) = "PRODUCIDAS", (.MaxRows), IIf(Trim(RS!mir_rutcli) = "MUESTRA R", .MaxRows - 1, .MaxRows - 2))
             .Col = j
             .text = IIf(IsNull(RS!mir_nrorac) Or RS!mir_nrorac = 0, "0", RS!mir_nrorac)
          
          End If
          
          RS.MoveNext
       
       Loop
       
       '-------> Mover raciones reales si no existe
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient

       RS1.Open "SELECT DISTINCT min_fecmin, min_racrea " & _
                "FROM b_minuta " & _
                "WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                "AND   min_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
                "AND   min_codser = " & Val(fpLongInteger1(2).Value) & " " & _
                "AND   min_fecmin >= " & fecini & " AND min_fecmin<=" & fecfin & "", vg_db, adOpenForwardOnly
       If Not RS1.EOF Then
          Do While Not RS1.EOF
             .Row = 0
             For i = 3 To .MaxCols
                 .Col = i
                 If Val(Format(Right(.text, 10), "yyyymmdd")) = RS1!min_fecmin Then j = i: Exit For
             Next i
             .Row = .MaxRows
             .Col = j
             'If Val(.text) < 1 Then
             .text = IIf(IsNull(RS1!min_racrea) Or RS1!min_racrea < 1, "0", RS1!min_racrea)
             RS1.MoveNext
          Loop
       End If
       RS1.Close: Set RS1 = Nothing
       
       '-------> Bloquea días de cierre en color rojo
       Dim diablq As Date
       If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Date, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
'       If Format(Date, "dd/mm/yyyy") > diablq Or Format(CDate(fpDateTime1.text), "mm/yyyy") <= Format(Month(Date) - 1 & "/" & Year(Date), "mm/yyyy") Then
        If CierrePeriodo(Format(fpDateTime1.text, "yyyyMM"), 0, 34) Or CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), 0, 35) Then
'       If Format(Date, "dd/mm/yyyy") > diablq Or Format(CDate(fpDateTime1.text), "mm/yyyy") <= Format(Month(Date) - 1 & "/" & Year(Date), "mm/yyyy") Then
          If Month(Date) = 1 Then
             v_columnas = ((dEoM(Format("01/" & "12/" & Year(Date) - 1)) - CDate(CDate("01/" + Mid(fpDateTime1.text, 1, 8))) + 1) * 2) - 1
          Else
'             v_columnas = ((dEoM(Format("01/" & (Month(Date) - 1) & "/" & Year(Date))) - CDate(CDate("01/" + Mid(fpDateTime1.text, 1, 8))) + 1) * 2) - 1
             v_columnas = (((CDate(vg_ciedia) - 1) - CDate(CDate("01/" + Mid(fpDateTime1.text, 1, 8))) + 1))
          End If
       Else
          v_columnas = 0
       End If
            
       If v_columnas > 0 Then
          .Row = -1
          For i = 3 To v_columnas + 2
              .Col = i
              .Lock = True
              .BackColor = Shape1(0).FillColor
          Next i
          .SetActiveCell i, 1
       End If
       If ciedia > 0 Then
          .Row = 0
          .Col = 3
          If Format(Date, "dd/mm/yyyy") > diablq Or Format(CDate(Right(.text, 10)), "mm/yyyy") <= Format(Month(Date) - 1 & "/" & Year(Date), "mm/yyyy") Then
             If Month(Date) = 1 Then
                v_columnas = ((dEoM(Format(ciedia + 1 & "/" & "12/" & Year(Date) - 1)) - CDate(CDate(Right(.text, 10))) + 1))
             Else
                v_columnas = ((dEoM(Format(IIf(Month(Date) - 1 = 2 And ciedia = 28 Or ciedia = 29, ciedia, ciedia + 1) & " /" & (Month(Date) - 1) & "/" & Year(Date))) - CDate(CDate(Right(.text, 10))) + 1))
             End If
          Else
             v_columnas = 0
          End If
          .Row = -1
          For i = 3 To v_columnas + 2
              .Col = i
              .Lock = True
              .BackColor = Shape1(0).FillColor
          Next i
       End If
       '-------> Fin Bloqueo de celdas
       Gl_Ac_Botones Me, 1, 4, modo
       .SetActiveCell 1, 1
    Else
       If fpText.text = "" Or Val(fpLongInteger1(1).Value) < 1 Or Val(fpLongInteger1(2).Value) < 0 Or fpDateTime1.text = "" Then
          Gl_Ac_Botones Me, 1, 3, modo
       Else
          Gl_Ac_Botones Me, 1, 2, modo
       End If
    End If
    RS.Close: Set RS = Nothing
    .Visible = True
End With
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If Row <> 1 Then Exit Sub

Dim NombreDia As String
Dim Check     As String
Dim mensaje   As String
Dim i         As Long

If EstCheck = True Then Exit Sub

    If Row = 1 Then
       
        vaSpread1.Row = Row
        vaSpread1.Col = Col
        Check = vaSpread1.text
        
        If Check = 1 Then
           
           mensaje = "Esta seguro no facturar día "
        
        Else
           
           mensaje = "Esta seguro Activa ingreso facturación del día "
        
        End If
        
        vaSpread1.Row = 0
        vaSpread1.Col = Col
        NombreDia = vaSpread1.text
        
        If MsgBox(mensaje & NombreDia & " ... ??", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
           
           vaSpread1.Row = Row
           vaSpread1.Col = Col
           EstCheck = True
           vaSpread1.text = IIf(Check = 1, "0", "1")
           EstCheck = False
           Exit Sub
           
        End If
       
       vaSpread1.Row = 1
       vaSpread1.Col = Col
       If Check = 1 Then
       
          EstCheck = True
          vaSpread1.TypeCheckText = "No Facturable"
          vaSpread1.text = "1"
          EstCheck = False
          For i = 2 To vaSpread1.MaxRows - 3
           
              vaSpread1.Row = i
              vaSpread1.Col = Col
              vaSpread1.text = ""
              vaSpread1.Lock = True
           
          Next i
          
       Else
       
          EstCheck = True
          vaSpread1.TypeCheckText = "Facturable"
          vaSpread1.text = "0"
          EstCheck = False
          For i = 2 To vaSpread1.MaxRows - 3
           
              vaSpread1.Row = i
              vaSpread1.Col = Col
              vaSpread1.text = ""
              vaSpread1.Lock = False
           
          Next i
       
       
       End If
       
       modo = "M"
       Gl_Ac_Botones Me, 1, 0, modo
       fpLongInteger1(1).Enabled = False
       Image1(1).Enabled = False
       fpLongInteger1(2).Enabled = False
       fpDateTime1.Enabled = False
       Image1(2).Enabled = False
       
       Exit Sub
    
    End If

End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

Dim RS        As New ADODB.Recordset
Dim Sql       As String
Dim Fecha     As Long
Dim EstMinuta As Boolean
Dim EstFactu  As Boolean

'-------> Grilla en rojo
vaSpread1.Row = Row
vaSpread1.Col = Col

If Row < vaSpread1.MaxRows Then

   Exit Sub

End If

If vaSpread1.BackColor = Shape1(0).FillColor Then

   Exit Sub
   
End If

If Toolbar1.Buttons(12).Visible = True Then

   
   MsgBox "Estan activados los botones de cancelar y confirmar, completar esa operación...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub
   
End If
'-------> valida si raciones es mayor que cero entonces sale
vaSpread1.Row = Row
vaSpread1.Col = Col

If Val(vaSpread1.text) > 0 Then

   Exit Sub
   
End If
'-------> Sacar fecha de la grilla
vaSpread1.Row = 0
vaSpread1.Col = Col
Fecha = Val(Format(Right(vaSpread1.text, 10), "yyyymmdd"))

'-------> Sacar dato facturable de la grilla
EstFactu = True
vaSpread1.Row = 1
vaSpread1.Col = Col
EstFactu = Val(vaSpread1.text)

If EstFactu Then

   MsgBox "Para ese día esta seleccionado ticket no facturable...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub
   
End If

vaSpread1.Row = Row
vaSpread1.Col = Col
       

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
       
Sql = ""
Sql = Sql & " '" & LimpiaDato(Trim(fpText.text)) & "' , "
Sql = Sql & " " & Val(fpLongInteger1(1).Value) & ", "
Sql = Sql & " " & Val(fpLongInteger1(2).Value) & ", "
Sql = Sql & " " & Fecha & " "

Set RS = vg_db.Execute("sgp_Sel_MinutaconcomensalesCeroConRac " & Sql & " ")

EstMinuta = False

If Not RS.EOF Then

    EstMinuta = True

Else

   MsgBox "Para ese día no existe detalle de la minuta o bien no tiene asignado raciones por receta...", vbExclamation + vbOKOnly, MsgTitulo
   RS.Close
   Set RS = Nothing

   Exit Sub

End If
RS.Close
Set RS = Nothing


If EstMinuta And Not EstFactu And Row = vaSpread1.MaxRows And Col > 2 And (vaSpread1.text = "" Or vaSpread1.text = "0") And Toolbar1.Buttons(12).Visible = False Then

   XRow = Row
   xcol = Col
   
   Label1(5).Caption = "Fecha : " & Mid(Fecha, 7, 2) & "/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)
   
   Frame1.Enabled = False
   Toolbar1.Enabled = False
   vaSpread1.Enabled = False
   
   Nombre(1).text = ""
   Nombre(0).text = ""
   Frame5.Visible = True
   Nombre(1).text = ""
   Nombre(0).text = ""
   fpNRac.text = 0

Else

   Exit Sub

End If

End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

With vaSpread1
    
    If .MaxRows < 1 Or ChangeMade = False Then
       
       Exit Sub
    
    End If
    
    If modo = "" Then
       
       modo = "M"
    
    End If
    
    If ChangeMade = True And modo = "M" Then
       
       Gl_Ac_Botones Me, 1, 0, modo
       fpLongInteger1(1).Enabled = False
       Image1(1).Enabled = False
       fpLongInteger1(2).Enabled = False
       fpDateTime1.Enabled = False
       Image1(2).Enabled = False
    
    End If
    
    If .MaxRows = Row Then
       
       Exit Sub
    
    End If
    
    .Row = Row
    .Col = Col
    
    If Col > 2 Then
       
       Dim totnrorac As Long
       totnrorac = 0
       
       For i = 1 To .MaxRows
           
           .Row = i: .Col = 1
           
           If Trim(.text) = "" Then
              
              Exit For
           
           End If
           
           .Col = Col
           
           If Trim(.text) <> "" Then
              
              totnrorac = CCur(totnrorac + .text)
           
           End If
       
       Next i
       
       If totnrorac > 0 Then
          
          EstCheck = True
          .Col = Col
          .text = totnrorac
          EstCheck = False
       
       Else
          
          EstCheck = True
          .Col = Col
          .text = ""
          EstCheck = False
       
       End If
    
    End If

End With

End Sub
