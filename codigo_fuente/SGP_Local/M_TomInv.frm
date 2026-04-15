VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_TomInv 
   Caption         =   "Toma de Inventario"
   ClientHeight    =   8130
   ClientLeft      =   1470
   ClientTop       =   2100
   ClientWidth     =   11685
   Icon            =   "M_TomInv.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.ComboBox Fplist1 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   3810
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   20
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1470
      Left            =   45
      TabIndex        =   3
      Top             =   405
      Width           =   11520
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   960
         Width           =   2550
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Mostrar Familia Producto"
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
         Left            =   7440
         TabIndex        =   19
         Top             =   630
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cierre de Mes"
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
         Left            =   9390
         TabIndex        =   14
         Top             =   825
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   550
         Width           =   2550
      End
      Begin EditLib.fpDateTime Date1 
         Height          =   345
         Index           =   0
         Left            =   1785
         TabIndex        =   0
         Top             =   555
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
         _ExtentY        =   609
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
         DateCalcMethod  =   3
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1800
         TabIndex        =   21
         Top             =   210
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
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   1845
         TabIndex        =   29
         Top             =   1020
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Inv."
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
         Left            =   510
         TabIndex        =   27
         Top             =   1080
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3030
         Picture         =   "M_TomInv.frx":0442
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
         Left            =   3540
         TabIndex        =   22
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
         Index           =   2
         Left            =   480
         TabIndex        =   20
         Top             =   320
         Width           =   735
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   10725
         TabIndex        =   16
         Top             =   870
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label2 
         Caption         =   "Comentario Fecha Cierre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8415
         TabIndex        =   15
         Top             =   255
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   4365
         TabIndex        =   6
         Top             =   610
         Width           =   2550
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
         Index           =   0
         Left            =   510
         TabIndex        =   5
         Top             =   700
         Width           =   540
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   3540
         TabIndex        =   4
         Top             =   630
         Width           =   660
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3585
         TabIndex        =   23
         Top             =   255
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6030
      Left            =   45
      TabIndex        =   7
      Top             =   1995
      Width           =   11520
      Begin VB.Frame Frame7 
         Height          =   2895
         Left            =   2520
         TabIndex        =   31
         Top             =   720
         Visible         =   0   'False
         Width           =   6975
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
            TabIndex        =   37
            Top             =   2280
            Width           =   1425
         End
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
            TabIndex        =   35
            Top             =   2280
            Width           =   1425
         End
         Begin EditLib.fpText Nombre 
            Height          =   315
            Index           =   1
            Left            =   2790
            TabIndex        =   33
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
            TabIndex        =   32
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Para anular la caratula inventario, tiene comunicarse su monitor"
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
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Width           =   5445
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
            Index           =   7
            Left            =   1440
            TabIndex        =   38
            Top             =   1125
            Width           =   585
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
            Index           =   6
            Left            =   1440
            TabIndex        =   36
            Top             =   1500
            Width           =   930
         End
         Begin VB.Label LbIn1 
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
            Left            =   240
            TabIndex        =   34
            Top             =   720
            Width           =   720
         End
      End
      Begin VB.Frame Frame6 
         Height          =   5175
         Left            =   1080
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   9735
         Begin VB.TextBox Text1 
            Height          =   4575
            Index           =   0
            Left            =   840
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   26
            Top             =   360
            Width           =   7935
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   4695
            Left            =   960
            TabIndex        =   25
            Top             =   360
            Width           =   7935
         End
      End
      Begin VB.Frame Frame5 
         Height          =   675
         Left            =   3060
         TabIndex        =   17
         Top             =   1860
         Visible         =   0   'False
         Width           =   4965
         Begin VB.Label Label3 
            Caption         =   "Un momento, recalculando Precio"
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
            Left            =   1110
            TabIndex        =   18
            Top             =   300
            Width           =   2880
         End
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   1905
         TabIndex        =   11
         Top             =   5385
         Width           =   4125
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   12
            Top             =   135
            Width           =   4020
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   90
         TabIndex        =   9
         Top             =   5385
         Width           =   1785
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   10
            Top             =   135
            Width           =   1680
         End
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   420
         Left            =   7305
         TabIndex        =   13
         Top             =   5460
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   741
         ButtonWidth     =   3307
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar Producto"
               Description     =   "Agregar Productos"
               Object.ToolTipText     =   "Agregar Producto"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar Producto "
               Description     =   "Eliminar Producto "
               Object.ToolTipText     =   "Eliminar Producto "
               ImageIndex      =   2
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10470
         Top             =   5235
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
               Picture         =   "M_TomInv.frx":074C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_TomInv.frx":0A66
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5105
         Left            =   90
         TabIndex        =   8
         Top             =   180
         Width           =   11340
         _Version        =   393216
         _ExtentX        =   20002
         _ExtentY        =   8996
         _StockProps     =   64
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
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
         MaxCols         =   8
         MaxRows         =   50
         SpreadDesigner  =   "M_TomInv.frx":0D80
         ScrollBarTrack  =   3
      End
   End
End
Attribute VB_Name = "M_TomInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String, auxmodo As String, est As Boolean
Dim MsgTitulo As String, diablq As Date
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset
Dim vg_codigo_Inv As Long
Dim Segundos As Byte
Const INTERVALO_EN_MINUTOS As Integer = 5

Private Sub Check2_Click()
If est Then Exit Sub
'-------> Actualizar parametro visualizar familia productos
vg_db.Execute "UPDATE a_param SET par_valor = '" & IIf(Check2.Value = 1, 1, 0) & "' WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'opfampro'"
MuestraDatosGrilla
On Error Resume Next: vaSpread1.SetFocus
End Sub

Private Sub Cmd1_Click()

On Error GoTo Man_Error

Dim RS       As New ADODB.Recordset
Dim v_codbod As Long
Dim v_fecinv As Variant


If Trim(Date1(0).text) = "" Then Exit Sub
v_fecinv = Format(Date1(0).text, "yyyymmdd")
v_codbod = fg_codigocbo(Combo1, 0, 10, 0)

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
    
    RS.Open "SELECT * FROM a_param WHERE par_codigo = 'parconaein' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    
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
        
    Frame7.Visible = False
    Nombre(0).text = ""
    Nombre(1).text = ""
    
    If ValidarOpEnvio(MuestraCasino(1), 2) Then
        
        If Not isInternetConnected(False, False, False) Then
        
           MsgBox "No hay conexión a internet, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        Frame1.Enabled = False
        Frame2.Enabled = False
    
    End If
    
    If Not GenerarArcInvSap(v_codbod, Date1(0).text, False) Then
       
       If ValidarOpEnvio(MuestraCasino(1), 2) Then
          
          Text1(0).text = VgLinea & Text1(0).text & FechaHora & "Generación envió anulado, Finalizado Con Problema" & VgLinea
          I_EnvioSap "3"
          Frame1.Enabled = True
          Frame2.Enabled = True
          Frame6.Visible = False
          Exit Sub
       
       End If
    
       'Se envía inventario a OPTIMUM aunque falle el envío a SAP, debe pasar por aquí cuando se desconecte SAP
       If ValidarOpEnvio(MuestraCasino(1), 5) Then
          
          If Not GeneraInvAX(vg_codigo_Inv, Format(Date1(0).text, "yyyymm"), Format(Date1(0).text, "yyyymmdd")) Then
             
             Call MsgBox("No genero correctamente archivos OPTIMUM, trate de generar por envio Inventario OPTIMUM", vbInformation)
          
          End If
       
       End If
       Exit Sub
    
    Else
       
       If ValidarOpEnvio(MuestraCasino(1), 5) Then
          
          If Not GeneraInvAX(vg_codigo_Inv, Format(Date1(0).text, "yyyymm"), Format(Date1(0).text, "yyyymmdd")) Then
             
             Call MsgBox("No genero correctamente archivos OPTIMUM, trate de generar por envio Inventario OPTIMUM", vbInformation)
          
          End If
       
       End If
       
       If ValidarOpEnvio(MuestraCasino(1), 2) Then
          
          Text1(0).text = VgLinea & Text1(0).text & FechaHora & "Generación envió anulado, Finalizado Sin Problema" & VgLinea
          I_EnvioSap "3"
       
       End If
    
    End If
    
    If Not DiferenciaInventario(Format(Date1(0).text, "yyyymmdd")) Then
       
       vg_db.Execute "UPDATE b_tomainv SET tin_autaju = '1' WHERE tin_fectom = " & Format(Date1(0).text, "yyyymmdd") & " AND tin_codbod = " & v_codbod & ""
    
    Else
       
       '-------> Activar autorización de envio
       vg_db.Execute "UPDATE b_tomainv SET tin_autaju = '0' WHERE tin_fectom = " & Format(Date1(0).text, "yyyymmdd") & " AND tin_codbod = " & v_codbod & ""
    
    End If

    '-------> INI: Mover estado a la tabla parametro toma inventario
    vg_db.Execute "update a_param set par_valor = '1' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "' and par_valor = '0'"
    '-------> FIN: Mover estado a la tabla parametro toma inventario
           
    '-------> INI : Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
        
    Set RS = vg_db.Execute("sgp_Upd_ValidarInventarioCalendarizado '" & MuestraCasino(1) & "', " & Format(Date1(0).text, "yyyymmdd") & ", '0'")
    If Not RS.EOF Then
    
       If RS(0) > 0 And Trim(RS(1)) <> "" Then
       
          RS.Close
          Set RS = Nothing
          
          MsgBox "Existe error grabar inventario calendarizado..", vbExclamation + vbOKOnly, MsgTitulo
          Exit Sub
    
       End If
    
    End If
    RS.Close
    Set RS = Nothing
    '-------> FIN : Validar inventario calendarizado Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado


    Frame1.Enabled = True
    
    vaSpread1.Enabled = True
    
    Toolbar1.Enabled = True
    Toolbar2.Enabled = True
    Text1(1).Enabled = True
    Text1(2).Enabled = True
    Gl_Ac_Botones Me, 6, 1, modo
           
Exit Sub
Man_Error:

fg_descarga

    Frame1.Enabled = True
    vaSpread1.Enabled = True
    Toolbar1.Enabled = True
    Toolbar2.Enabled = True
    Text1(1).Enabled = True
    Text1(2).Enabled = True

MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub Cmd2_Click()

On Error GoTo Man_Error

Frame7.Visible = False

Frame1.Enabled = True
vaSpread1.Enabled = True
Toolbar1.Enabled = True
Toolbar2.Enabled = True
Text1(1).Enabled = True
Text1(2).Enabled = True

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub Combo1_Click(Index As Integer)

If est Then Exit Sub
Select Case Index

Case 0
    
    est = True: Date1(0).text = "": est = False
    MuestraDatosGrilla
    On Error Resume Next: vaSpread1.SetFocus

End Select

End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub Date1_Change(Index As Integer)

Dim v_fecinv As Variant, v_codbod As Long
If est Then Exit Sub
Label2.Caption = ""
Check1.Enabled = False
est = True: Check1.Value = 0: est = False
v_codbod = fg_codigocbo(Combo1, 0, 10, 0)
v_fecinv = Format(Date1(0).text, "yyyymmdd")
If modo = "A" And Val(v_fecinv) < Val(Date1(0).DateMin) Or Val(v_fecinv) > Val(Date1(0).DateMax) Then Date1(0).text = Format(CDate(Mid(Date1(0).DateMax, 7, 2) & "/" & Mid(Date1(0).DateMax, 5, 2) & "/" & Mid(Date1(0).DateMax, 1, 4)), "dd/mm/yyyy"): Exit Sub
RS1.Open "SELECT DISTINCT tin_fectom FROM b_tomainv WHERE tin_codbod = " & v_codbod & " ORDER BY tin_fectom desc", vg_db, adOpenStatic
If Not RS1.EOF Then If Val(v_fecinv) <= RS1!tin_fectom Then est = True: Date1(0).text = Str(CDate(fg_Ctod1(RS1!tin_fectom)) + 1): est = False
RS1.Close: Set RS1 = Nothing

End Sub

Private Sub Date1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub Date1_LostFocus(Index As Integer)

If IsDate(Date1(0).text) = False Then On Error Resume Next: Date1(0).SetFocus

End Sub

Private Sub Form_Activate()

fg_descarga
'If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
'If vaSpread1.MaxRows < 1 Or modo = "E" Then
'If modo = "E" Then modo = ""
'Exit Sub
'End If
'auxmodo = modo: modo = "X"
'MuestraDatosGrilla
'modo = auxmodo

End Sub

Private Sub Form_Load()

On Local Error GoTo Error_Partida

Me.Width = 11805
Me.Height = 8640
EspFecha Date1(0)
Me.HelpContextID = vg_OpcM
MsgTitulo = "Toma de Inventario"
fg_centra Me
modo = ""

If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then
   diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy")
Else
   diablq = 0
End If
Gl_Mo_Botones Me, 6

'-------> Formato de Celdas
vaSpread1.Row = -1
vaSpread1.Col = 4: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 6: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
vaSpread1.Col = 7: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr

'-------> Trae todos los registros de las Bodegas Disponibles
est = True
Combo1(0).Clear

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("SELECT a.* FROM a_bodega a, b_clientes b WHERE a.bod_codigo = b.cli_codbod AND b.cli_codigo = '" & vg_contra & "' ORDER BY bod_nombre")
Do While Not RS1.EOF
    Combo1(0).AddItem RS1!bod_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!bod_codigo), 10) & ")"
    RS1.MoveNext
Loop
RS1.Close
Set RS1 = Nothing

If Combo1(0).listcount > 0 Then Combo1(0).ListIndex = 0
Date1(0).text = ""
'-------> cargar tipo inventario
Combo1(1).Clear
Combo1(1).AddItem "Inventario Rotativo" & Space(150) & "(" & "1" & ")"
Combo1(1).AddItem "Inventario Full" & Space(150) & "(" & "2" & ")"
Combo1(1).ListIndex = -1
Combo1(1).Enabled = ValidarInventarioRotativo(MuestraCasino(1))
Check2.Value = IIf(0 = (fg_CambiaChar(GetParametro("opfampro"), ";", "','")), 0, 1)

'-------> Formatear grilla
If vg_pais = "CO" Then
   fg_OcultarGrilla vaSpread1, -1, 4, True
   fg_OcultarGrilla vaSpread1, -1, 6, True
   fg_OcultarGrilla vaSpread1, -1, 7, True
   vaSpread1.Row = 0
   vaSpread1.Col = 2
   vaSpread1.ColWidth(2) = 62.7
   Check2.Visible = False
   Check2.Value = 0
End If

'-------> Mover datos contrato
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
Fplist1.Visible = False
If vg_invrot = "1" Then
   
   If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 4) Then
      
      Combo1(1).ListIndex = 0
      FuncionAgregado
   
   Else
      
      MuestraDatosGrilla
      est = False
      modo = "E"
   
   End If
   Toolbar1.Buttons(19).Enabled = False

Else
   
   MuestraDatosGrilla
   est = False
   modo = "E"

End If

Exit Sub
Error_Partida:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, MsgTitulo
End Sub

Private Sub Form_Resize()

If Me.WindowState = 2 Then
    
    Frame1.Left = (Me.Width \ 2) - (Frame1.Width \ 2)
    Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)

ElseIf Me.WindowState = 0 Then
    
    Frame1.Left = 45
    Frame2.Left = 45

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
        
On Local Error GoTo Error_Salida
        
Dim RS1 As New ADODB.Recordset

If Trim(Date1(0).text) = "" Or Not IsDate(Date1(0).text) Or Combo1(0).ListIndex < 0 Then Exit Sub

If Format(Date1(0).text, "yyyymmdd") = Format(CDate(vg_ciedia) - 1, "yyyymmdd") Then

   If Not CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), fg_codigocbo(Combo1, 0, 10, 0), 45) Then
        
      
      If MsgBox("Esta seguro cerrar inventario, ya que no hay diferencia de ajuste... ?? ", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
      
         Exit Sub
      
      End If
      
      '-------> INI: Mover estado a la tabla parametro toma inventario
      vg_db.Execute "update a_param set par_valor = '0' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
      '-------> FIN: Mover estado a la tabla parametro toma inventario

      '-------> INI : Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado
      If RS1.State = 1 Then RS1.Close
      RS1.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
        
      Set RS1 = vg_db.Execute("sgp_Upd_ValidarInventarioCalendarizado '" & MuestraCasino(1) & "', " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ", '1'")
      If Not RS1.EOF Then
    
         If RS1(0) > 0 And Trim(RS1(1)) <> "" Then
       
            RS1.Close
            Set RS1 = Nothing
          
            MsgBox "Existe error grabar inventario calendarizado..", vbExclamation + vbOKOnly, MsgTitulo
            Exit Sub
    
         End If
    
      End If
      RS1.Close
      Set RS1 = Nothing
      '-------> FIN : Validar inventario calendarizado Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado
           
   End If

End If

Exit Sub
Error_Salida:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub fpList1_Click()

If est Then Exit Sub
Fplist1.Visible = False
Me.Refresh
est = True: Date1(0).text = Fplist1.List(Fplist1.ListIndex): est = False
MuestraDatosGrilla
On Error Resume Next: vaSpread1.SetFocus

End Sub

Private Sub fpList1_LostFocus()

Fplist1.Visible = False

End Sub

Private Sub fpList1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

est = True
Fplist1.SelLength = True
'.Selected(Fplist1.MouseIcon) = True
est = False

End Sub

Private Sub fpText_Change()

If est Then Exit Sub
Set RS1 = vg_db.Execute("SELECT * FROM b_clientes WHERE cli_codigo = '" & fpText.text & "' AND cli_tipo = 0")
If RS1.EOF Then
   RS1.Close
   Set RS = Nothing
   fpayuda(0).Caption = ""
   Exit Sub
End If
fpayuda(0).Caption = Trim(RS1!cli_nombre)
RS1.Close: Set RS1 = Nothing

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
End Select

End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 1, 2
    vaSpread1.Visible = False
    If Trim(Text1(Index).text) <> "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread1.Col = 1
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
              'Activar familia productos
              If Check2.Value = 1 Then
                 For j = i To 1 Step -1
                     vaSpread1.Row = j: vaSpread1.Col = 1
                     If Trim(vaSpread1.text) = "" Then vaSpread1.RowHidden = False: Exit For
                 Next j
              End If
           Else
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           End If
        Next i
        vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1_Click Index, 0
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    'vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).Text) = "", Index, 8)) = ColUserSortIndicatorAscending
    'vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).Text) = "", Index, 8): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    
    vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
    If Trim(Text1(Index).text) = "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Next
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1.Visible = True
End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)

Select Case Index
Case 1, 2
    Text1(1) = "": Text1(2) = ""
End Select

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

If Index = 0 Then Exit Sub
If KeyAscii <> 13 Then Exit Sub
vaSpread1.SetActiveCell 5, vaSpread1.ActiveRow
On Error Resume Next: vaSpread1.SetFocus

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim i           As Long
Dim v_codbod    As Long

Dim codpro      As String
Dim stofis      As Double
Dim stosis      As Double
Dim v_fecinv    As Variant

Dim sqlTMP      As String
Dim Fecha       As String
Dim ciemes      As Long
Dim Casino      As String
Dim RS1         As New ADODB.Recordset
Dim RS2         As New ADODB.Recordset
Dim RS3         As New ADODB.Recordset

On Error GoTo Man_Error

If Trim(Date1(0).text) = "" Then Exit Sub
v_fecinv = Format(Date1(0).text, "yyyymmdd")
v_codbod = fg_codigocbo(Combo1, 0, 10, 0)
Frame6.Visible = False
TraerFechaCierre

Select Case Button.Index

Case 1 '-------> Agregar toma
   
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 9) Then
    
       modo = "E"
       MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 46) Then
    
       modo = "E"
       MsgBox "Existen documentos pendientes, en la ventas servicios especiales. Debe cerrar las ventas servicios especiales", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 8) Then
    
       modo = "E"
       MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
'    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 48) Then
'
'       modo = "E"
'       MsgBox "Existe ajuste de inventario. proceso cancelado.", vbExclamation + vbOKOnly, MsgTitulo
'       Exit Sub
'
'    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 9) Then
    
       modo = "E"
       MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 49) Then
    
       MsgBox "Existen generación de caratula inventario, debe anular la generación de caratula inventario. Proceso cancelado.", vbExclamation + vbOKOnly, MsgTitulo
       Toolbar1.Enabled = True
       Exit Sub
    
    End If
    
    If CierreAjuste Then Exit Sub
    
    FuncionAgregado

Case 3 '-------> Modifica
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 0) Then
    
       modo = "E"
       MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 2) Then
    
       modo = "E"
       MsgBox "Existen documentos posteriores a esta toma inventario...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 46) Then
    
       modo = "E"
       MsgBox "Existen documentos pendientes, en la ventas servicios especiales. Debe cerrar las ventas servicios especiales", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 9) Then
       
       modo = "E"
       MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CDate(Date1(0).text) <> (CDate(vg_ciedia) - 1) Then
       
       modo = "E"
       MsgBox "Día esta bloqueado", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS1 = vg_db.Execute("SELECT MAX(tin_fectom) AS fecha FROM b_tomainv WHERE tin_codbod = " & Val(fg_codigocbo(M_TomInv.Combo1, 0, 10, 0)) & "")
    If Not RS1.EOF Then
        
        If fg_Ctod1(RS1!Fecha) <> Date1(0).text Then
            
            MsgBox "Solo puede modificar el último inventario" & vbCrLf & _
                   "si no se ha generado el ajuste...", vbExclamation + vbOKOnly, MsgTitulo
            RS1.Close: Set RS1 = Nothing: Exit Sub
        
        End If
    
    End If
    RS1.Close: Set RS1 = Nothing
    modo = "M"
    vaSpread1.Row = -1: vaSpread1.Col = 5
    vaSpread1.EditMode = True
    vaSpread1.Lock = False
    Gl_Ac_Botones Me, 6, 0, modo

Case 5 '-------> Borra_Datos
    
    '-------> Validar si el contrato tiene asignado inventario rotativo
    If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
       Format(Date1(0).Value, "dd/mm/yyyy") = (CDate(vg_ciedia) - 1) And Not CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 29) Then MsgBox "No es posible borrar documento, debe reabrir día...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 0) Then
    
       modo = "E"
       MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo
       modo = "E"
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 7) Then
    
       modo = "E"
       MsgBox "Existen documentos posteriores a la fecha de esta toma de inventario...", vbExclamation + vbOKOnly, MsgTitulo
       modo = "E"
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 9) Then
       
       modo = "E"
       MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 46) Then
    
       modo = "E"
       MsgBox "Existen documentos pendientes, en la ventas servicios especiales. Debe cerrar las ventas servicios especiales", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CDate(Date1(0).text) <> (CDate(vg_ciedia) - 1) And Not CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 27) Then
    
       modo = "E"
       MsgBox "Día esta bloqueado", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If vaSpread1.ActiveRow < 1 Then
    
       MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo
       modo = "E"
       Exit Sub
    
    End If
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("SELECT DISTINCT tin_fectom FROM b_tomainv WHERE tin_codbod = " & v_codbod & " ORDER BY tin_fectom DESC")
    If Not RS1.EOF Then
       
       If Str(v_fecinv) <> Str(RS1!tin_fectom) Then
          
          MsgBox "Solo puede eliminar la ultima toma de inventario...", vbExclamation + vbOKOnly, MsgTitulo
          RS1.Close
          Set RS1 = Nothing
          modo = "E"
          Exit Sub
          
        End If
    
    End If
    
    RS1.Close
    Set RS1 = Nothing
    
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
    
       modo = "E": Exit Sub
    
    End If
    
    Toolbar1.Enabled = False
    
    'Detalle - Devuelve stock
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS1 = vg_db.Execute("SELECT dev.dev_codmer, dev.dev_canmer, aju.aju_tipo FROM b_totventas tov, b_detventas dev, a_tipoajuste aju " & _
              "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc AND tov.tov_numdoc = dev.dev_numdoc " & _
              "AND   tov.tov_codser = aju.aju_codigo AND tov.tov_fecemi = '" & Format(Date1(0).text, "yyyymmdd") & "' AND tov_codbod = " & v_codbod & " " & _
              "AND   tov.tov_tipdoc = 'AI' AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' ORDER BY dev.dev_numlin")

    If Not RS1.EOF Then

        Do While Not RS1.EOF

            vg_db.Execute "UPDATE b_bodegas SET bod_canmer = bod_canmer" & IIf(RS1!aju_tipo = "A", "-", "+") & RS1!dev_canmer & " " & _
                          "WHERE bod_codpro = '" & RS1!dev_codmer & "' AND bod_codbod = " & v_codbod & ""
            RS1.MoveNext

        Loop

    End If
    RS1.Close: Set RS1 = Nothing

    '-------> Borrar toma inventario
    vg_db.Execute "DELETE b_tomainv FROM b_tomainv WHERE tin_fectom = " & Val(v_fecinv) & " AND tin_codbod = " & v_codbod & ""
    
'    If RS1.State = 1 Then RS1.Close
'    RS1.CursorLocation = adUseClient
'    vg_db.CursorLocation = adUseClient
'
'    Set RS1 = vg_db.Execute("sgp_UpdDel_Actualizar_Stock_Bodega_Anulacion_Borrado '" & MuestraCasino(1) & "', " & v_codbod & ", " & Val(v_fecinv) & "")
'
'    If Not RS1.EOF Then
'
'       If RS1(0) > 0 And Trim(RS1(1)) <> "" Then
'
'          RS1.Close
'          Set RS1 = Nothing
'
'          Toolbar1.Enabled = True
'          MsgBox "Existe error en la actualización de la bodega y eliminación bodega. Proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo
'          Exit Sub
'
'       End If
'
'    End If
'    RS1.Close
'    Set RS1 = Nothing
    
    
'    '-------> INI: Mover estado a la tabla parametro toma inventario
'    vg_db.Execute "update a_param set par_valor = '0' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
'    '-------> FIN: Mover estado a la tabla parametro toma inventario
        
    '-------> INI : Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
        
    Set RS1 = vg_db.Execute("sgp_Upd_ValidarInventarioCalendarizado '" & MuestraCasino(1) & "', " & Format(Date1(0).text, "yyyymmdd") & ", '0'")
    If Not RS1.EOF Then
    
       If RS1(0) > 0 And Trim(RS1(1)) <> "" Then
       
          RS1.Close
          Set RS1 = Nothing
          
          MsgBox "Existe error grabar inventario calendarizado..", vbExclamation + vbOKOnly, MsgTitulo
          Exit Sub
    
       End If
    
    End If
    RS1.Close
    Set RS1 = Nothing
    '-------> FIN : Validar inventario calendarizado Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado
        
    '-------> Encabezado

    vg_db.Execute "UPDATE b_totventas SET tov_estdoc = 'A' WHERE tov_fecemi = '" & Format(Date1(0).text, "yyyymmdd") & "' " & _
                  "AND tov_codbod = " & v_codbod & " AND tov_tipdoc = 'AI' AND tov_estdoc <> 'A' " & _
                  "AND tov_estdoc <> 'P'"
        
    If vg_invrot = "1" Then
       
       Toolbar1.Enabled = True
       Combo1(1).ListIndex = 0
       FuncionAgregado
       Toolbar1.Buttons(19).Enabled = False
       Combo1(1).Enabled = True
    
    Else
       
       Frame5.Visible = True
       Label3.Enabled = True
       Label3.Caption = "Un momento, Recalculando día"
       
       If vg_tipbase = "1" Then
          
          CalcularPMPDiaAccess Me, False, True
       
       Else
          
          CalcularPMPDiaSql Me, False, True
       
       End If
       Frame5.Visible = False
       Label3.Enabled = False
       est = True
       
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS1 = vg_db.Execute("SELECT DISTINCT tin_fectom FROM b_tomainv WHERE tin_codbod = " & v_codbod & " ORDER BY tin_fectom DESC")
       If Not RS1.EOF Then Date1(0).text = fg_Ctod1(RS1!tin_fectom) Else Date1(0).text = ""
       RS1.Close: Set RS1 = Nothing
       
       est = False
       modo = ""
       MuestraDatosGrilla
       Toolbar1.Enabled = True
       Combo1(1).Enabled = True
    
    End If
    
    '-------> Actualizar productospmpdia saldo
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("SELECT ppd_propon, ppd_saldo, ppd_codpro FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(Date1(0).Value, "yyyymmdd") & "")
    If Not RS1.EOF Then
       
       vg_db.Execute "UPDATE b_productospmpdia SET ppd_saldo = 0 WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & ""
    
       Do While Not RS1.EOF
       
          vg_db.Execute "UPDATE b_productospmpdia SET ppd_propon = " & RS1!ppd_propon & ", ppd_saldo = " & RS1!ppd_saldo & " WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & " AND ppd_codpro = '" & RS1!ppd_codpro & "'"
          RS1.MoveNext
    
       Loop
    End If
    RS1.Close
    Set RS1 = Nothing

Case 7 '-------> Actualizar
    
    modo = ""
    Toolbar1.Enabled = False
    MuestraDatosGrilla
    Toolbar1.Enabled = True

Case 10 '-------> Cancelar
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("SELECT COUNT(tov_fecemi) AS suma FROM b_totventas WHERE " & _
         "tov_codbod = " & Val(fg_codigocbo(Combo1, 0, 10, 0)) & " AND tov_tipdoc = 'AI' AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'")
    
    If vaSpread1.MaxRows = 0 And modo = "A" And RS1!Suma = 0 Then
       
       RS1.Close
       Set RS1 = Nothing
       Exit Sub
    
    End If
    
    RS1.Close
    Set RS1 = Nothing
    
    If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    '-------> Muestra el ultimo inventario
    est = True
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("SELECT DISTINCT tin_fectom FROM b_tomainv WHERE tin_codbod = " & v_codbod & " ORDER BY tin_fectom DESC")
    If Not RS1.EOF Then
       
       Date1(0).text = fg_Ctod1(RS1!tin_fectom)
    
    End If
    
    RS1.Close
    Set RS1 = Nothing
    
    est = False
    modo = ""
    MuestraDatosGrilla

Case 12 '-------> Confirmar
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 48) Then

       modo = "E"
       MsgBox "Existe ajuste de inventario. proceso cancelado.", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub

    End If

    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 2) Then
    
       modo = "E"
       MsgBox "Existen documentos posteriores a esta toma inventario...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 0) Then
    
       modo = "E"
       MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 9) Then
    
       modo = "E"
       MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 46) Then
    
       modo = "E"
       MsgBox "Existen documentos pendientes, en la ventas servicios especiales. Debe cerrar las ventas servicios especiales", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If Trim(Date1(0).text) = "" Or Trim(fpayuda(0).Caption) = "" Then
    
       MsgBox "Falta dato importante...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If Val(v_fecinv) = Val(Date1(0).DateMax) Then
    
       Check1.Value = 1
       Label2.Caption = "Cierre de Mes "
    
    Else
       
       Check1.Value = 0
       Label2.Caption = "Precierre de mes "
    
    End If
    
    Toolbar1.Enabled = False
    If modo = "M" Then
        
        For i = 1 To vaSpread1.MaxRows
            
            DoEvents
            vaSpread1.Row = i
            vaSpread1.Col = 1
            codpro = vaSpread1.text
            
            If Trim(codpro) <> "" Then
               
               vaSpread1.Col = 4
               stosis = vaSpread1.text
               
               vaSpread1.Col = 5
               stofis = vaSpread1.text
               
               vg_db.Execute "UPDATE b_tomainv SET tin_stofis = " & stofis & ", tin_stosis = " & stosis & " " & _
                             "WHERE tin_fectom = " & v_fecinv & " AND tin_codbod = " & v_codbod & " AND tin_codpro = '" & codpro & "'"
            
            End If
        
        Next i
        modo = ""
        MuestraDatosGrilla
        
        If vg_pais = "CO" And vg_invrot = "1" Then
           
           M_AjuInv.Show 1
           Me.Hide
           Unload Me
        
        End If
    
    ElseIf modo = "A" Then
        
        modo = "A"
        
        If vg_invrot = "1" And Combo1(1).ListIndex = -1 Then
           
           Toolbar1.Enabled = True
           MsgBox "Debe seleccionar tipo inventario...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If vg_invrot = "1" And Not CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 29) Then
           
           If vg_tipbase = "1" Then
              
              CalcularPMPDiaAccess Me, False, False
           
           Else
              
              CalcularPMPDiaSql Me, False, False
           
           End If
        
        End If
        
        If Combo1(1).ListIndex = 0 Then
           
           vg_codigo = ""
           If TraerParametroStock(MuestraCasino(1)) = "1" And TraerTipoInventarioRotativo(MuestraCasino(1)) = "1" Then
              
              If Not ValidarDatoCurvaABC Then MsgBox "No existen datos en tabla curva ABC. Proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
              CalcularInvRotCuarvaABC vg_codbod, "1"
           
           ElseIf TraerParametroStock(MuestraCasino(1)) = "1" And TraerTipoInventarioRotativo(MuestraCasino(1)) = "2" Then
              
              If TraerPorcentajeInventario(MuestraCasino(1)) = 0 Then MsgBox "El valor del porcentaje inventario esta con valor cero. Proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
              CalcularInvRotPorInventario vg_codbod, "1"
           
           ElseIf TraerParametroStock(MuestraCasino(1)) = "2" And TraerTipoInventarioRotativo(MuestraCasino(1)) = "1" Then
              
              If Not ValidarDatoCurvaABC Then MsgBox "No existen datos en tabla curva ABC. Proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
              CalcularInvRotCuarvaABC vg_codbod, "2"
           
           ElseIf TraerParametroStock(MuestraCasino(1)) = "2" And TraerTipoInventarioRotativo(MuestraCasino(1)) = "2" Then
              
              If TraerPorcentajeInventario(MuestraCasino(1)) = 0 Then MsgBox "El valor del porcentaje inventario esta con valor cero. Proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
              CalcularInvRotPorInventario vg_codbod, "2"
           
           End If
           If vg_codigo <> "|Ok|" Then Exit Sub
           MuestraDatosGrilla
        
        Else
           
           MuestraDatosGrilla
        
        End If
    
    End If
    
    Date1(0).Enabled = False
    
    If Not DiferenciaInventario(Format(Date1(0).text, "yyyymmdd")) Then
       
       '-------> Activar autorización de envio
       vg_db.Execute "UPDATE b_tomainv SET tin_autaju = '1' WHERE tin_fectom = " & Format(Date1(0).text, "yyyymmdd") & " AND tin_codbod = " & v_codbod & ""
    
    Else
       
       vg_db.Execute "UPDATE b_tomainv SET tin_autaju = '0' WHERE tin_fectom = " & Format(Date1(0).text, "yyyymmdd") & " AND tin_codbod = " & v_codbod & ""
    
    End If
    
    modo = ""
    Gl_Ac_Botones Me, 6, 1, modo
    Toolbar1.Enabled = True
    On Error Resume Next: vaSpread1.SetFocus

Case 15 '-------> Imprimir
    
    I_TomInv.Show 1

Case 18 '-------> Historico
    
    If Fplist1.Visible = True Then Fplist1.Visible = False: On Error Resume Next: vaSpread1.SetFocus: Exit Sub
    v_codbod = fg_codigocbo(Combo1, 0, 10, 0)
    Fplist1.Clear
    Fplist1.Visible = True
    Fplist1.ZOrder
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS1 = vg_db.Execute("SELECT DISTINCT tin_fectom FROM b_tomainv WHERE tin_codbod = " & v_codbod & " ORDER BY tin_fectom DESC")
    i = 75
    
    Do While Not RS1.EOF
       
       Fplist1.AddItem fg_Ctod1(RS1!tin_fectom)
       RS1.MoveNext: i = i + 195
    
    Loop
    RS1.Close
    Set RS1 = Nothing
    
 '   Fplist1.Height = IIf(i > 2025, 2025, i)
    est = True:
    'Fplist1.SelLength = True'
    est = False
    On Error Resume Next: Fplist1.SetFocus

Case 19 '-------> Filtrar
    
    If Date1(0).text = "" Then
       
       MsgBox "Debe ingresar fecha...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
       
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 48) Then

       modo = "E"
       MsgBox "Existe ajuste de inventario. proceso cancelado.", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub

    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 2) Then
       
       modo = "E"
       MsgBox "Existen documentos posteriores, a esta toma inventario...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 9) Then
    
       modo = "E"
       MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
       
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 46) Then
    
       modo = "E"
       MsgBox "Existen documentos pendientes, en la ventas servicios especiales. Debe cerrar las ventas servicios especiales", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    vg_codigo = ""
    vg_bodega = 0
    vg_bodega = Val(fg_codigocbo(Combo1, 0, 10, ""))
    
    B_Produc.Show 1
    If vg_codigo <> "|Ok|" Then Exit Sub
    MuestraDatosGrilla
    modo = "M"

Case 21 '-------> Autorizaciòn ajuste
    
    modo = "E"
    If MsgBox("Autoriza ajuste...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then modo = "E": Exit Sub
    vg_db.Execute "UPDATE b_tomainv SET tin_autaju = '1' WHERE tin_fectom = " & Format(Date1(0).text, "yyyymmdd") & " AND tin_codbod = " & v_codbod
    Gl_Ac_Botones Me, 6, 1, modo

Case 23 '-------> Generar envio Inventario
    
    auxmodo = modo
    modo = "E"
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 9) Then
       
       modo = "E"
       MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 46) Then
    
       modo = "E"
       MsgBox "Existen documentos pendientes, en la ventas servicios especiales. Debe cerrar las ventas servicios especiales", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If

    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 8) Then
    
       modo = "E"
       MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierreAjuste Then Exit Sub
'    If Not isNetwork(NETWORK_ALIVE_LAN) Then MsgBox "No hay conexión a internet, proceso cancelado", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    '--> Validar envio inventario sap
    If ValidarOpEnvio(MuestraCasino(1), 2) Then
       
       If Not isInternetConnected(False, False, False) Then
          
          MsgBox "No hay conexión a internet, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo
          Exit Sub
       
       End If
    
    End If
    
    Frame1.Enabled = False: Frame2.Enabled = False
    If Not GenerarArcInvSap(v_codbod, Date1(0).text, True) Then
       
       If ValidarOpEnvio(MuestraCasino(1), 2) Then
          
          Text1(0).text = Text1(0).text & FechaHora & "Generación envió Finalizado Con Problema" & VgLinea
          I_EnvioSap "3"
          Frame1.Enabled = True: Frame2.Enabled = True: Frame6.Visible = False
       
       End If

       '***********************************************************************************************************************
       'Se envía inventario a OPTIMUM aunque falle el envío a SAP, debe pasar por aquí cuando se desconecte SAP
       If ValidarOpEnvio(MuestraCasino(1), 5) Then
          
          If Not GeneraInvAX(vg_codigo_Inv, Format(Date1(0).text, "yyyymm"), Format(Date1(0).text, "yyyymmdd")) Then
             
             Call MsgBox("No genero correctamente archivos OPTIMUM, trate de generar por envio Inventario OPTIMUM", vbInformation)
          
          End If
       
       End If
'       Text1(0).text = Text1(0).text & FechaHora & "Generación envió Finalizado Sin Problema" & VgLinea
'       I_EnvioSap "3"
       '***********************************************************************************************************************
       Exit Sub
    
    Else
       
       If ValidarOpEnvio(MuestraCasino(1), 5) Then
          
          If Not GeneraInvAX(vg_codigo_Inv, Format(Date1(0).text, "yyyymm"), Format(Date1(0).text, "yyyymmdd")) Then
             
             Call MsgBox("No genero correctamente archivos OPTIMUM, trate de generar por envio Inventario OPTIMUM", vbInformation)
          
          End If
       
       End If
       
       If ValidarOpEnvio(MuestraCasino(1), 2) Then
          
          Text1(0).text = Text1(0).text & FechaHora & "Generación envió Finalizado Sin Problema" & VgLinea
          I_EnvioSap "3"
       
       End If
    
    End If
    
    If ValidarOpEnvio(MuestraCasino(1), 2) Then
       
       Frame1.Enabled = True
       Frame2.Enabled = True
       Frame6.Visible = False
    
    End If
    
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame6.Visible = False
    
    '-------> INI: Mover estado a la tabla parametro toma inventario
    vg_db.Execute "update a_param set par_valor = '0' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "' and par_valor = '1'"
    '-------> FIN: Mover estado a la tabla parametro toma inventario

    '-------> INI : Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
        
    Set RS1 = vg_db.Execute("sgp_Upd_ValidarInventarioCalendarizado '" & MuestraCasino(1) & "', " & Format(Date1(0).text, "yyyymmdd") & ", '1'")
    If Not RS1.EOF Then
    
       If RS1(0) > 0 And Trim(RS1(1)) <> "" Then
       
          RS1.Close
          Set RS1 = Nothing
          
          MsgBox "Existe error grabar inventario calendarizado..", vbExclamation + vbOKOnly, MsgTitulo
          Exit Sub
    
       End If
    
    End If
    RS1.Close
    Set RS1 = Nothing
    '-------> FIN : Validar inventario calendarizado Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado

    Dim SumaBloqueo
    SumaBloqueo = 0
    
    If SumaBloqueo = 0 And CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 49) Then

       SumaBloqueo = 1

    End If
    
    vaSpread1.Lock = IIf(SumaBloqueo = 0, False, True)
    
    Gl_Ac_Botones Me, 6, 1, modo
    Date1(0).Enabled = False
    
Case 24 '-------> Generar anular envio Inventario
    
    auxmodo = modo
    modo = "E"
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 0) Then
    
       modo = "E"
       MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 9) Then
       
       modo = "E"
       MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 46) Then
    
       modo = "E"
       MsgBox "Existen documentos pendientes, en la ventas servicios especiales. Debe cerrar las ventas servicios especiales", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If

    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 8) Then
    
       modo = "E"
       MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierreAjuste Then Exit Sub
    
'Mod Ini 20240801    Frame1.Enabled = False
'Mod Ini 20240801    vaSpread1.Enabled = False
'Mod Ini 20240801    Toolbar1.Enabled = False
'Mod Ini 20240801    Toolbar2.Enabled = False
'Mod Ini 20240801    Text1(1).Enabled = False
'Mod Ini 20240801    Text1(2).Enabled = False
    
'Mod Ini 20240801    Frame7.Visible = True
'Mod Ini 20240801    Frame7.Enabled = True
'Mod Ini 20240801    LbIn1.Caption = "Fecha : " & Format(Date1(0).text, "dd/mm/yyyy")
    
'    If Not isNetwork(NETWORK_ALIVE_LAN) Then MsgBox "No hay conexión a internet, proceso cancelado", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    
    If ValidarOpEnvio(MuestraCasino(1), 2) Then

        If Not isInternetConnected(False, False, False) Then MsgBox "No hay conexión a internet, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        Frame1.Enabled = False
        Frame2.Enabled = False

    End If

    If Not GenerarArcInvSap(v_codbod, Date1(0).text, False) Then

       If ValidarOpEnvio(MuestraCasino(1), 2) Then

          Text1(0).text = VgLinea & Text1(0).text & FechaHora & "Generación envió anulado, Finalizado Con Problema" & VgLinea
          I_EnvioSap "3"
          Frame1.Enabled = True: Frame2.Enabled = True: Frame6.Visible = False
          Exit Sub

       End If

       'Se envía inventario a OPTIMUM aunque falle el envío a SAP, debe pasar por aquí cuando se desconecte SAP
       If ValidarOpEnvio(MuestraCasino(1), 5) Then

          If Not GeneraInvAX(vg_codigo_Inv, Format(Date1(0).text, "yyyymm"), Format(Date1(0).text, "yyyymmdd")) Then

             Call MsgBox("No genero correctamente archivos OPTIMUM, trate de generar por envio Inventario OPTIMUM", vbInformation)

          End If

       End If
       Exit Sub

    Else

       If ValidarOpEnvio(MuestraCasino(1), 5) Then

          If Not GeneraInvAX(vg_codigo_Inv, Format(Date1(0).text, "yyyymm"), Format(Date1(0).text, "yyyymmdd")) Then

             Call MsgBox("No genero correctamente archivos OPTIMUM, trate de generar por envio Inventario OPTIMUM", vbInformation)

          End If

       End If

       If ValidarOpEnvio(MuestraCasino(1), 2) Then

          Text1(0).text = VgLinea & Text1(0).text & FechaHora & "Generación envió anulado, Finalizado Sin Problema" & VgLinea
          I_EnvioSap "3"

       End If

    End If

    If Not DiferenciaInventario(Format(Date1(0).text, "yyyymmdd")) Then

       vg_db.Execute "UPDATE b_tomainv SET tin_autaju = '1' WHERE tin_fectom = " & Format(Date1(0).text, "yyyymmdd") & " AND tin_codbod = " & v_codbod & ""

    Else

       '-------> Activar autorización de envio
       vg_db.Execute "UPDATE b_tomainv SET tin_autaju = '0' WHERE tin_fectom = " & Format(Date1(0).text, "yyyymmdd") & " AND tin_codbod = " & v_codbod & ""

    End If

    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame6.Visible = False
    Gl_Ac_Botones Me, 6, 1, modo

Case 26 '-------> Ajustar Inventario
    
    Toolbar1.Enabled = False
    '-------> Valida que exista casino en operación
    Casino = "": Casino = vg_contra
    If Trim(Casino) = "" Then MsgBox "No existe casino en operación...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    If RS2.State = 1 Then RS2.Close
    RS2.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    If vg_tipbase = "1" Then
       
       RS2.Open "SELECT COUNT(tov_fecemi) AS suma FROM b_totventas WHERE tov_fecemi = Cdate('" & Date1(0).text & "') " & _
                "AND tov_codbod = " & v_codbod & " AND tov_tipdoc = 'AI' AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'", vg_db, adOpenStatic
    
    Else
       
       RS2.Open "SELECT COUNT(tov_fecemi) AS suma FROM b_totventas WHERE tov_fecemi = '" & Format(Date1(0).text, "yyyymmdd") & "' " & _
                "AND tov_codbod = " & v_codbod & " AND tov_tipdoc = 'AI' AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'", vg_db, adOpenStatic
    
    End If
       '-------> Validar si existe toma inventario
       
    If RS3.State = 1 Then RS3.Close
    RS3.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
       
    If vg_tipbase = "1" Then
          
       RS3.Open "SELECT tov.tov_fecemi " & _
                "FROM b_totventas tov, b_detventas dev " & _
                "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc AND tov.tov_numdoc = dev.dev_numdoc " & _
                "AND tov.tov_fecemi = Cdate('" & Date1(0).text & "') AND tov_codbod = " & v_codbod & " AND tov.tov_tipdoc = 'AI' AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P'", vg_db, adOpenStatic
       
    Else
         
       RS3.Open "SELECT tov.tov_fecemi " & _
                "FROM b_totventas tov, b_detventas dev " & _
                "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc AND tov.tov_numdoc = dev.dev_numdoc " & _
                "AND tov.tov_fecemi = '" & Format(Date1(0).text, "yyyymmdd") & "' AND tov_codbod = " & v_codbod & " AND tov.tov_tipdoc = 'AI' AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P'", vg_db, adOpenStatic
       
    End If
       
    If RS3.EOF And CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 2) Then
          
       RS3.Close: Set RS3 = Nothing: MsgBox "Existen documentos posteriores, a esta toma inventario...", vbExclamation + vbOKOnly, MsgTitulo
       
       If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 9) Then
       
          MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo
       
       End If
    
       If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 46) Then
    
          MsgBox "Existen documentos pendientes, en la ventas servicios especiales. Debe cerrar las ventas servicios especiales", vbExclamation + vbOKOnly, MsgTitulo
    
       End If
    
    Else
          
       RS3.Close
       Set RS3 = Nothing
       M_AjuInv.Show 1
       
    End If
    RS2.Close
    Set RS2 = Nothing
    modo = ""
    MuestraDatosGrilla
    Toolbar1.Enabled = True

Case 27 '-------> Anular Ajuste Inventario
    
    Casino = ""
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'casino'", vg_db, adOpenStatic
    If Not RS1.EOF Then Casino = Trim(TipoDato(RS1!par_valor, ""))
    RS1.Close: Set RS1 = Nothing
    If Trim(Casino) = "" Then MsgBox "No existe casino en operación...", vbExclamation + vbOKOnly, MsgTitulo: modo = "E": Toolbar1.Enabled = True: Exit Sub
    
    '-------> Validar si el contrato tiene asignado inventario rotativo
    If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
       Format(Date1(0).Value, "dd/mm/yyyy") = (CDate(vg_ciedia) - 1) And Not CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 29) Then MsgBox "No es posible anular ajuste, debe reabrir día...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 0) Then
    
       MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo
       modo = "E"
       Toolbar1.Enabled = True: Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 2) Then
    
       MsgBox "No puede anular ajuste inventario, existen documentos posteriores...", vbExclamation + vbOKOnly, MsgTitulo
       modo = "E"
       Toolbar1.Enabled = True
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 9) Then
    
       MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo
       Toolbar1.Enabled = True
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 46) Then
    
       MsgBox "Existen documentos pendientes, en la ventas servicios especiales. Debe cerrar las ventas servicios especiales", vbExclamation + vbOKOnly, MsgTitulo
       Toolbar1.Enabled = True
       Exit Sub
    
    End If

    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 49) Then
    
       MsgBox "Existen un envio documento, debe anular envio documento. Proceso cancelado.", vbExclamation + vbOKOnly, MsgTitulo
       Toolbar1.Enabled = True
       Exit Sub
    
    End If

    If MsgBox("Anula ajuste de inventario...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
    
       modo = "E"
       Toolbar1.Enabled = True
       Exit Sub
    
    End If
    
    Toolbar1.Enabled = False
    
    
    '-------> Detalle - Devuelve stock
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS1.Open "SELECT dev.dev_codmer, dev.dev_canmer, aju.aju_tipo FROM b_totventas tov, b_detventas dev, a_tipoajuste aju " & _
             "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc AND tov.tov_numdoc = dev.dev_numdoc " & _
             "AND tov.tov_codser = aju.aju_codigo AND tov.tov_fecemi = '" & Format(Date1(0).text, "yyyymmdd") & "' AND tov_codbod = " & v_codbod & " " & _
             "AND tov.tov_tipdoc = 'AI' AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' ORDER BY dev.dev_numlin", vg_db, adOpenStatic

    If Not RS1.EOF Then

        Do While Not RS1.EOF

            vg_db.Execute "UPDATE b_bodegas SET bod_canmer = bod_canmer" & IIf(RS1!aju_tipo = "A", "-", "+") & RS1!dev_canmer & " " & _
                          "WHERE bod_codpro = '" & RS1!dev_codmer & "' AND bod_codbod = " & v_codbod & ""
            RS1.MoveNext

        Loop

    End If
    RS1.Close
    Set RS1 = Nothing

    '-------> Encabezado
    vg_db.Execute "UPDATE b_totventas SET tov_estdoc = 'A' WHERE tov_fecemi = '" & Format(Date1(0).text, "yyyymmdd") & "' " & _
                  "AND tov_codbod = " & Val(fg_codigocbo(Combo1, 0, 10, 0)) & " AND tov_tipdoc = 'AI' AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'"



'    Dim CodBodAju As Long
'    Dim FecInvAju As Long
'
'    CodBodAju = Val(fg_codigocbo(Combo1, 0, 10, 0))
'    FecInvAju = Format(Date1(0).text, "yyyymmdd")
'
'    If RS1.State = 1 Then RS1.Close
'    RS1.CursorLocation = adUseClient
'    vg_db.CursorLocation = adUseClient
'
'    Set RS1 = vg_db.Execute("sgp_Upd_actualizar_Stock_Bodega_Anulacion '" & MuestraCasino(1) & "', " & CodBodAju & ", " & FecInvAju & "")
'
'    If Not RS1.EOF Then
'
'       If RS1(0) > 0 And Trim(RS1(1)) <> "" Then
'
'          RS1.Close
'          Set RS1 = Nothing
'
'          Toolbar1.Enabled = True
'          MsgBox "Existe error en la actualización de la bodega. Proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo
'          Exit Sub
'
'       End If
'
'    End If
'    RS1.Close
'    Set RS1 = Nothing
    
    MsgBox "Ajuste anulado...", vbInformation, MsgTitulo
    
    If Not DiferenciaInventario(Format(Date1(0).text, "yyyymmdd")) Then
       
       '-------> Activar autorización de envio
       vg_db.Execute "UPDATE b_tomainv SET tin_autaju = '1', tin_envioADMSGP = '0' WHERE tin_fectom = " & Format(Date1(0).text, "yyyymmdd") & " AND tin_codbod = " & v_codbod
    
    Else
       
       vg_db.Execute "UPDATE b_tomainv SET tin_autaju = '0', tin_envioADMSGP = '0' WHERE tin_fectom = " & Format(Date1(0).text, "yyyymmdd") & " AND tin_codbod = " & v_codbod
    
    End If
    
    '-------> INI: Mover estado a la tabla parametro toma inventario
    vg_db.Execute "update a_param set par_valor = '1' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
    '-------> FIN: Mover estado a la tabla parametro toma inventario
           
    '-------> INI : Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
        
    Set RS1 = vg_db.Execute("sgp_Upd_ValidarInventarioCalendarizado '" & MuestraCasino(1) & "', " & Format(Date1(0).text, "yyyymmdd") & ", '0'")
    If Not RS1.EOF Then
    
       If RS1(0) > 0 And Trim(RS1(1)) <> "" Then
       
          RS1.Close
          Set RS1 = Nothing
          
          MsgBox "Existe error grabar inventario calendarizado..", vbExclamation + vbOKOnly, MsgTitulo
          Exit Sub
    
       End If
    
    End If
    RS1.Close
    Set RS1 = Nothing
    '-------> FIN : Validar inventario calendarizado Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado
    
    modo = ""
    MuestraDatosGrilla
    Toolbar1.Enabled = True

Case 29 '-------> Exportar Inventario
    
    Toolbar1.Enabled = False
    P_EIInve.Inicio "Exportar Inventario", "E", Format(Date1(0).text, "yyyymmdd")
    P_EIInve.Show 1
    Toolbar1.Enabled = True

Case 30
    
    vg_codigo = ""
       
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 0) Then
    
       MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo
       modo = "E"
       Toolbar1.Enabled = True: Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 49) Then
    
       MsgBox "Existen generación de caratula inventario, debe anular la generación de caratula inventario. Proceso cancelado.", vbExclamation + vbOKOnly, MsgTitulo
       Toolbar1.Enabled = True
       Exit Sub
    
    End If
    
    Toolbar1.Enabled = False
    P_EIInve.Inicio "Importar Inventario", "I", Format(Date1(0).text, "yyyymmdd")
    P_EIInve.Show 1
    If vg_codigo <> "" Then
       
       modo = ""
       MuestraDatosGrilla
    End If
    
    Toolbar1.Enabled = True

Case 32 '-------> Generar inventario OPTIMUM
    
    P_GenInvAx.Show 1, Partida

Case 34 '-------> Explorar carpeta envio inventario OPTIMUM
    
    ExplorarCarpeta dir_trabajo & "InformesAXInventario"

Case 36 '-------> Salir
    modo = "E"
    Me.Hide
    Unload Me
    
End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then
   
   MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
   Exit Sub

End If

If Err = 3034 Then
   
   Exit Sub

End If
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Sub FuncionAgregado()

Dim RS1 As New ADODB.Recordset
'-------> Traer fecha cierre de mes
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT cie_fecini, cie_fecter FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_estado = 1", vg_db, adOpenStatic
If Not RS1.EOF Then
   
   v_fecinv = Format(CDate(Mid(RS1!cie_fecter, 7, 2) & "/" & Mid(RS1!cie_fecter, 5, 2) & "/" & Mid(RS1!cie_fecter, 1, 4)), "dd/mm/yyyy")
   Date1(0).DateMin = CStr(RS1!cie_fecini)
   Date1(0).DateMax = CStr(RS1!cie_fecter)
   Date1(0).text = ""
   Date1(0).text = v_fecinv

End If
RS1.Close
Set RS1 = Nothing
'------->  Fin fecha cierre de mes

'-------> Mover fecha cierre de día
If vg_ciedia <> "" Then
   
   est = True
   v_fecinv = IIf(ValidarInventarioRotativo(MuestraCasino(1)) And Not CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 29), Format(CDate(vg_ciedia), "dd/mm/yyyy"), Format(CDate(vg_ciedia) - 1, "dd/mm/yyyy"))
   Date1(0).DateMin = IIf(ValidarInventarioRotativo(MuestraCasino(1)) And Not CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 29), Format(CDate(vg_ciedia), "yyyymmdd"), Format(CDate(vg_ciedia) - 1, "yyyymmdd")) 'Format(CDate(vg_ciedia) - 1, "yyyymmdd")
   Date1(0).DateMax = IIf(ValidarInventarioRotativo(MuestraCasino(1)) And Not CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 29), Format(CDate(vg_ciedia), "yyyymmdd"), Format(CDate(vg_ciedia) - 1, "yyyymmdd")) 'Format(CDate(vg_ciedia) - 1, "yyyymmdd")
   Date1(0).text = IIf(ValidarInventarioRotativo(MuestraCasino(1)) And Not CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 29), Format(CDate(vg_ciedia), "dd/mm/yyyy"), Format(CDate(vg_ciedia) - 1, "dd/mm/yyyy")) 'Format(CDate(vg_ciedia) - 1, "dd/mm/yyyy")
   est = False

End If
Date1(0).Enabled = True
vaSpread1.MaxRows = 0
Check1.Value = 0
modo = "A": Gl_Ac_Botones Me, 6, 0, modo
modo = "A"
'-------> determinar tipo inventario
If CierrePeriodo(Format(CDate(Date1(0).text), "yyyymmdd"), vg_codbod, 29) Then
   
   Combo1(1).ListIndex = 1

End If
On Error Resume Next: vaSpread1.SetFocus

End Sub

Sub MuestraDatosGrilla()

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset

On Local Error GoTo Error_Mover

fg_carga ""
Dim v_fecinv As Long, v_codbod As Long, aAp As String, i As Long, FecCie As Long, fecper As Long, aAp1 As String
Dim sqlTMP As String, Fecha As String, sqlPROPON As String, AñoMes As Long
v_codbod = fg_codigocbo(Combo1, 0, 10, 0)

If Trim(Date1(0).text) = "" Then
    
    '-------> Muestra el ultimo inventario
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    est = True
    RS1.Open "SELECT DISTINCT tin_fectom FROM b_tomainv WHERE tin_codbod = " & v_codbod & " ORDER BY tin_fectom DESC", vg_db, adOpenStatic
    
    If Not RS1.EOF Then
        
        Date1(0).text = fg_Ctod1(RS1!tin_fectom)
        modo = ""
    
    Else
        
        RS1.Close: Set RS1 = Nothing
        Date1(0).text = IIf(Trim(vg_ciedia) <> "", vg_ciedia, "")
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        est = False
        fg_descarga
        Exit Sub
    
    End If
    RS1.Close
    Set RS1 = Nothing
    est = False

End If

v_fecinv = Format(Date1(0).text, "yyyymmdd")

'-------> mover fecha final de mes para validar y no ocultar la columna stock
Dim fecterper As Long
fecterper = 0

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT cie_fecini, cie_fecter FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_estado = 1", vg_db, adOpenStatic
If Not RS1.EOF Then fecterper = RS1!cie_fecter
RS1.Close
Set RS1 = Nothing

'-------> modifica la columna 13-12-2023 If vg_pais = "CL" And (modo = "A" Or modo = "M") And v_fecinv <> fecterper Then
If vg_pais = "CL" And (modo = "A" Or modo = "M") Then
   
   fg_OcultarGrilla vaSpread1, -1, 4, True
   vaSpread1.ColWidth(2) = 34.5 + 9

'-------> modifica la columna 13-12-2023 ElseIf vg_pais = "CL" And v_fecinv <> fecterper Then
ElseIf vg_pais = "CL" Then

   fg_OcultarGrilla vaSpread1, -1, 4, IIf(v_fecinv <> Format(CDate(vg_ciedia) - 1, "yyyymmdd"), False, True)
   vaSpread1.ColWidth(2) = IIf(v_fecinv <> Format(CDate(vg_ciedia) - 1, "yyyymmdd"), 34.5, 34.5 + 9)

End If

'--------- Revisa Cierre de Mes ----------
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT DISTINCT tin_fectom, tin_ciemes FROM b_tomainv WHERE left(tin_fectom,6) = " & Val(Format(Date1(0).text, "yyyymm")) & " AND tin_codbod = " & v_codbod & "", vg_db, adOpenStatic
If Not RS1.EOF Then
   
   Do While Not RS1.EOF
      
      est = True:  est = False
      If RS1!tin_fectom = Val(Format(Date1(0).text, "yyyymmdd")) Then Label2.Caption = IIf(RS1!tin_ciemes > 0, "Cierre de Mes ", "Precierre de Mes ")  ' & fg_Ctod1(RS1!tin_fectom)
      FecCie = RS1!tin_fectom
      RS1.MoveNext
   
   Loop

End If
RS1.Close
Set RS1 = Nothing

'-----------------------------------------
'--------- Reviso si hay ajuste ----------
'-------> revisa fecha periodo
fecper = 0
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT cie_fecini, cie_fecter FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_estado = 1", vg_db, adOpenStatic
If Not RS1.EOF Then fecper = RS1!cie_fecter
RS1.Close
Set RS1 = Nothing

Toolbar2.Enabled = False
sqlPROPON = "tin.tin_propon"
sqlTMP = IIf(vg_codigo = "|Ok|", " AND pro.pro_codigo IN (SELECT * FROM " & Trim(vg_NUsr) & "_tmp_filtomainv) ", "")

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT COUNT(tov_fecemi) AS suma FROM b_totventas WHERE tov_fecemi = '" & Format(Date1(0).text, "yyyymmdd") & "' " & _
         "AND tov_codbod = " & Val(fg_codigocbo(Combo1, 0, 10, 0)) & " AND tov_tipdoc = 'AI' AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'", vg_db, adOpenStatic

If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS2.Open "SELECT MAX(tin_fectom) AS fecha FROM b_tomainv WHERE tin_codbod = " & Val(fg_codigocbo(Combo1, 0, 10, 0)), vg_db, adOpenStatic
If Not RS2.EOF And Not IsNull(RS2!Fecha) Then Fecha = fg_Ctod1(RS2!Fecha) Else Fecha = Date1(0).text
RS2.Close
Set RS2 = Nothing

If RS1!Suma = 0 And CDate(Date1(0).text) >= CDate(Fecha) And Not CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), Val(fg_codigocbo(Combo1, 0, 10, 0)), 2) Then

   If modo = "A" Then
        
        '------- Recalcular promedio precio ponderado
        '------- Fin recalcular promedio precio ponderado
        '-------> Agrega productos si no existen en la bodega
        vg_db.Execute "INSERT INTO b_tomainv (tin_fectom, tin_codbod, tin_codpro, tin_stofis, tin_stosis, tin_propon) " & _
                      "SELECT DISTINCT " & v_fecinv & ", " & v_codbod & ", pro.pro_codigo, 0, 0, 0 FROM b_productos pro, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR pro.pro_maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = pro.pro_maepro OR pro.pro_maepro < 1) AND pro.pro_codigo NOT IN " & _
                      "(SELECT tin_codpro FROM b_tomainv WHERE tin_fectom = " & v_fecinv & " AND tin_codbod = " & v_codbod & ") AND pro.pro_ctrsto = 1 AND (pro.pro_fecven > " & Format(Date, "YYYYMMDD") & " OR pro.pro_fecven <= 0)" & sqlTMP
        DoEvents
        '-------> insertar productos que tienen fecha de vencimiento, pero tienen stock.
        vg_db.Execute "INSERT INTO b_tomainv (tin_fectom, tin_codbod, tin_codpro, tin_stofis, tin_stosis, tin_propon) " & _
                      "SELECT DISTINCT " & v_fecinv & ", " & v_codbod & ", pro.pro_codigo, 0, 0, 0 FROM b_productos pro, b_bodegas bod, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR pro.pro_maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = pro.pro_maepro OR pro.pro_maepro < 1) AND pro.pro_codigo = bod.bod_codpro AND pro.pro_codigo NOT IN (SELECT tin_codpro FROM b_tomainv WHERE tin_fectom = " & v_fecinv & " AND tin_codbod = " & v_codbod & ") AND pro.pro_ctrsto = 1 AND bod.bod_codbod = " & v_codbod & " AND bod.bod_canmer > 0" & sqlTMP
        DoEvents
        
        '-------> INI: Mover estado a la tabla parametro toma inventario
        vg_db.Execute "update a_param set par_valor = '1' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
        '-------> FIN: Mover estado a la tabla parametro toma inventario
        
        '-------> Actualizar precio pmp
        vg_db.Execute "UPDATE b_tomainv SET b_tomainv.tin_propon = b_productospmpdia.ppd_propon FROM b_tomainv, b_productospmpdia WHERE b_tomainv.tin_codpro = b_productospmpdia.ppd_codpro " & _
                      "AND b_tomainv.tin_codbod = " & vg_codbod & " AND b_productospmpdia.ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & " AND b_productospmpdia.ppd_cencos = '" & MuestraCasino(1) & "' AND b_tomainv.tin_fectom = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & ""
        
        If Val(v_fecinv) = fecper Then Check1.Value = 1: Label2.Caption = "Cierre de Mes " Else Check1.Value = 0: Label2.Caption = "Precierre de mes "
        vg_db.Execute "UPDATE b_tomainv SET tin_ciemes = " & Val(IIf(Check1.Value = 1, Mid(v_fecinv, 1, 6), 0)) & " WHERE tin_fectom = " & v_fecinv & " AND tin_codbod = " & v_codbod & ""
        DoEvents
        vg_db.Execute "UPDATE b_tomainv SET tin_tipinv = '" & fg_codigocbo(Combo1, 1, 1, 0) & "' WHERE tin_codbod = " & v_codbod & " AND tin_fectom = " & v_fecinv & ""
   
   End If
   
       DoEvents
       
       If modo = "A" Then

          If RS3.State = 1 Then RS3.Close
          RS3.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient

          Set RS3 = vg_db.Execute("sgp_Upd_actualizar_Stock_Bodega '" & MuestraCasino(1) & "', " & vg_codbod & ", " & v_fecinv & "")

          If Not RS3.EOF Then

             If RS3(0) > 0 And Trim(RS3(1)) <> "" Then

                RS3.Close
                Set RS3 = Nothing

                MsgBox "Existe error en la actualización de la bodega. Proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo
                Exit Sub

              End If

          End If
          RS3.Close
          Set RS3 = Nothing

       End If
       
       vg_db.Execute ("sgp_s_tomainventario '" & MuestraCasino(1) & "', " & v_codbod & ", " & v_fecinv & ", '" & Format(Date1(0).text, "dd/mm/yyyy") & "', " & vg_DCa & ", '" & modo & "', " & fecterper & "")
       vg_db.Execute "Update b_productospmpdia " & _
                     "Set    b_productospmpdia.ppd_saldo = b_tomainv.tin_stofis " & _
                     "From   b_productospmpdia, b_tomainv " & _
                     "Where  b_productospmpdia.ppd_codpro = b_tomainv.tin_codpro " & _
                     "AND    b_productospmpdia.ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & " " & _
                     "AND    b_productospmpdia.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                     "AND    b_tomainv.tin_codbod         = " & vg_codbod & " " & _
                     "AND    b_tomainv.tin_fectom         = " & v_fecinv & ""
    
    
    If Format(Date1(0).text, "yyyymmdd") = Format(CDate(vg_ciedia) - 1, "yyyymmdd") Then
    
       If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), fg_codigocbo(Combo1, 0, 10, 0), 8) Then
            
          '-------> INI: Mover estado a la tabla parametro toma inventario
          vg_db.Execute "update a_param set par_valor = '1' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
          '-------> FIN: Mover estado a la tabla parametro toma inventario
                  
          '-------> INI : Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado
          If RS3.State = 1 Then RS3.Close
          RS3.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient
                    
          Set RS3 = vg_db.Execute("sgp_Upd_ValidarInventarioCalendarizado '" & MuestraCasino(1) & "', " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ", '0'")
          If Not RS3.EOF Then
                
             If RS3(0) > 0 And Trim(RS3(1)) <> "" Then
                   
                RS3.Close
                Set RS3 = Nothing
                      
                MsgBox "Existe error grabar inventario calendarizado..", vbExclamation + vbOKOnly, MsgTitulo
                Exit Sub
                
              End If
                
          End If
          RS3.Close
          Set RS3 = Nothing
          '-------> FIN : Validar inventario calendarizado Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado
       
       End If
    
    End If
    DoEvents
    Toolbar2.Enabled = True
'--------------------------------------------
End If
RS1.Close: Set RS1 = Nothing
Dim opord As String, codtip As Long
opord = IIf(Check2.Value = 0, " ORDER BY pro.pro_nombre", " ORDER BY pro.pro_ctacon, pro.pro_codtip, pro.pro_nombre")

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT tin.tin_stofis, tin.tin_stosis, pro.pro_codigo, pro.pro_nombre, pro.pro_codtip, pro.pro_ctacon, uni.uni_nombre, " & sqlPROPON & ", tin.tin_tipinv " & _
         "FROM b_tomainv tin, b_productos pro, a_unidad uni " & _
         "WHERE uni.uni_codigo = pro.pro_coduni AND pro.pro_codigo = tin.tin_codpro " & _
         "AND tin_fectom = " & v_fecinv & " AND tin_codbod = " & v_codbod & sqlTMP & "  " & opord & "", vg_db, adOpenForwardOnly
codtip = 0
If modo = "X" Then
   
   If Not RS1.EOF Then
      
      Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 1, (RS1!tin_tipinv))
      i = 1
      vaSpread1.Visible = False
      
      Do While Not RS1.EOF
         
         If vaSpread1.SearchCol(1, 0, vaSpread1.MaxRows, Trim(RS1!pro_codigo), SearchFlagsNone) <> -1 Then
            
            vaSpread1.Row = vaSpread1.SearchCol(1, 0, vaSpread1.MaxRows, Trim(RS1!pro_codigo), SearchFlagsNone)
            vaSpread1.Col = 4
            vaSpread1.text = IIf(modo = "A", Format(0, fg_Pict(9, vg_DCa)), Format(RS1!tin_stosis, fg_Pict(9, vg_DCa)))
         
         End If
         
         RS1.MoveNext
      
      Loop
      
      vaSpread1.Visible = True
      vaSpread1.SetActiveCell 5, 1
   
   End If

Else
   
   vaSpread1.Visible = False
   vaSpread1.MaxRows = 0
   
   If Not RS1.EOF Then
      
      i = 1
      Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 1, (RS1!tin_tipinv))
      
      Do While Not RS1.EOF
         
         vaSpread1.MaxRows = vaSpread1.MaxRows + 1
         vaSpread1.Row = vaSpread1.MaxRows
         
         If RS1!pro_codtip <> codtip And Check2.Value = 1 Then
            
            vaSpread1.Col = 1
            vaSpread1.Lock = True
            vaSpread1.text = ""
            
            vaSpread1.Col = 2
            vaSpread1.Lock = True
            vaSpread1.Font.Bold = True
            vaSpread1.text = fg_BuscaenArbol(RS1!pro_codtip, "a_tipopro", "tip_codigo")
            
            vaSpread1.Col = 3
            vaSpread1.Lock = True
            vaSpread1.text = ""
            
            vaSpread1.Col = 4
            vaSpread1.Lock = True
            vaSpread1.text = ""
            
            vaSpread1.Col = 5
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.Lock = True
            vaSpread1.text = ""
            
            vaSpread1.Col = 6
            vaSpread1.Lock = True
            vaSpread1.text = ""
            
            vaSpread1.Col = 7
            vaSpread1.Lock = True
            vaSpread1.text = ""
            
            codtip = RS1!pro_codtip
            vaSpread1.MaxRows = vaSpread1.MaxRows + 1
            vaSpread1.Row = vaSpread1.MaxRows
         
         End If
         
         vaSpread1.Col = 1
         vaSpread1.text = Trim(RS1!pro_codigo)
         
         vaSpread1.Col = 2
         vaSpread1.text = Trim(RS1!pro_nombre)
         
         vaSpread1.Col = 3
         vaSpread1.text = Trim(RS1!uni_nombre)
         
         vaSpread1.Col = 4
         'modifica 13/12/2023 vaSpread1.text = IIf(modo = "A" And v_fecinv <> fecterper, Format(0, fg_Pict(9, vg_DCa)), Format(RS1!tin_stosis, fg_Pict(9, vg_DCa)))
         vaSpread1.Col = 4
         vaSpread1.text = IIf(modo = "A", Format(0, fg_Pict(9, vg_DCa)), Format(RS1!tin_stosis, fg_Pict(9, vg_DCa)))

         vaSpread1.Col = 5
         vaSpread1.text = Format(RS1!tin_stofis, fg_Pict(9, vg_DCa))
         
         vaSpread1.Col = 6
         vaSpread1.text = Format(RS1(7), fg_Pict(9, vg_DPr))
         
         vaSpread1.Col = 7
         vaSpread1.text = Format(Format(RS1!tin_stofis, fg_Pict(9, vg_DCa)) * IIf(IsNull(RS1(7)), 0, Format(RS1(7), fg_Pict(9, vg_DPr))), fg_Pict(9, vg_DPr))
         
         RS1.MoveNext
      
      Loop
      vaSpread1.SetActiveCell 5, 1
   
   End If
   vaSpread1.Visible = True

End If
RS1.Close: Set RS1 = Nothing

If vg_codigo = "|Ok|" Then
   
   vg_db.Execute "DROP TABLE " & Trim(vg_NUsr) & "_tmp_filtomainv"

End If

'-------> Reviso si hay ajuste y bloqueo --------
Combo1(1).Enabled = False
vaSpread1.Row = -1
vaSpread1.Col = 5

Dim Suma As Long
Suma = 0

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
  
RS1.Open "SELECT COUNT(tov_fecemi) AS suma FROM b_totventas WHERE tov_fecemi = '" & Format(Date1(0).text, "yyyymmdd") & "' " & _
         "AND tov_codbod = " & Val(fg_codigocbo(Combo1, 0, 10, 0)) & " AND tov_tipdoc = 'AI' AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'", vg_db, adOpenStatic

If Not RS1.EOF Then

   Suma = RS1!Suma
   
End If
RS1.Close
Set RS1 = Nothing

If Suma = 0 And CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), v_codbod, 0) Then
    
   Suma = 1
    
End If

If Suma = 0 And CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 49) Then

   Suma = 1

End If

If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS2.Open "SELECT MAX(tin_fectom) AS fecha FROM b_tomainv WHERE tin_codbod = " & Val(fg_codigocbo(Combo1, 0, 10, 0)), vg_db, adOpenStatic
If Not RS2.EOF Then
   
   Fecha = fg_Ctod1(RS2!Fecha)

Else
   
   Fecha = Form.Date1(0).text

End If
RS2.Close
Set RS2 = Nothing

vaSpread1.Lock = IIf(Suma = 0 And Fecha = Date1(0).text, False, True)

'--------------------------------------------
vg_codigo = ""
Gl_Ac_Botones Me, 6, 1, modo
Date1(0).Enabled = False
fg_descarga

Exit Sub
Error_Mover:
    fg_descarga
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, MsgTitulo
    Resume Next

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim i As Long, v_fecinv As Variant, v_codbod As Long

On Local Error GoTo Error_Mover1
modo = "E"
If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then MsgBox "Debe grabar información, antes de " & IIf(Button.Index = 1, "Agregar", "Eliminar") & " producto...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub

Select Case Button.Index

Case 1
    
    vg_nombre = "": vg_codigo = "": vg_bodega = 0: vg_bodega = Val(fg_codigocbo(Combo1, 0, 10, ""))
    v_codbod = fg_codigocbo(Combo1, 0, 10, 0)
    v_fecinv = Format(Date1(0).text, "yyyymmdd")
    vg_left = Toolbar2.Width
    B_TabEst.LlenaDatos Trim(CStr(v_fecinv)), Trim(Str(v_codbod)), "Productos", "ProInv"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    Me.Refresh
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Col = 1: vaSpread1.Row = i
        If Trim(vaSpread1.text) = Trim(vg_codigo) Then MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    Next i
    
    '-------------------- Reviso si hay ajuste ----------------
    '-------> Agrega producto si no existen en la toma
    vg_db.Execute "INSERT INTO b_tomainv (tin_fectom, tin_codbod, tin_codpro, tin_stofis, tin_stosis, tin_propon) " & _
                  "SELECT " & v_fecinv & ", " & v_codbod & ", pro.pro_codigo, 0, 0, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = pro.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") FROM b_productos pro " & _
                  "WHERE pro.pro_codigo NOT IN (SELECT tin_codpro FROM b_tomainv WHERE tin_fectom = " & v_fecinv & " " & _
                  "AND tin_codbod = " & v_codbod & ") AND pro.pro_codigo = '" & Trim(vg_codigo) & "'"
    
    '-------> INI: Mover estado a la tabla parametro toma inventario
    vg_db.Execute "update a_param set par_valor = '1' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
    '-------> FIN: Mover estado a la tabla parametro toma inventario
        
    '-------> Trae los Stock de la fecha de sistema
    If vg_tipbase = "1" Then
       
       vg_db.Execute "UPDATE b_tomainv a INNER JOIN b_bodegas b ON a.tin_codbod=b.bod_codbod AND a.tin_codpro=b.bod_codpro " & _
                     "SET a.tin_stosis = b.bod_canmer WHERE a.tin_fectom = " & v_fecinv & " AND a.tin_codbod = " & v_codbod & " " & _
                     "AND a.tin_codpro = '" & Trim(vg_codigo) & "'"
    
    Else
       vg_db.Execute "UPDATE b_tomainv SET b_tomainv.tin_stosis = b.bod_canmer FROM b_tomainv a, b_bodegas b WHERE a.tin_codbod = b.bod_codbod AND a.tin_codpro = b.bod_codpro " & _
                     "AND a.tin_fectom = " & v_fecinv & " AND a.tin_codbod = " & v_codbod & " " & _
                     "AND a.tin_codpro = '" & Trim(vg_codigo) & "'"
    
    End If
    
    vg_db.Execute "UPDATE b_tomainv SET tin_ciemes = " & Val(IIf(Check1.Value = 1, Mid(v_fecinv, 1, 6), 0)) & " WHERE tin_fectom = " & v_fecinv & " AND tin_codbod = " & v_codbod
    
    '-------> Actualizar b_productospmpdia
    If vg_tipbase = "1" Then
       
       vg_db.Execute "UPDATE b_productospmpdia INNER JOIN b_tomainv ON b_productospmpdia.ppd_codpro = b_tomainv.tin_codpro SET b_productospmpdia.ppd_saldo = b_tomainv.tin_stofis " & _
                     "WHERE b_productospmpdia.ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & " AND b_productospmpdia.ppd_cencos = '" & MuestraCasino(1) & "' AND b_tomainv.tin_codbod = " & vg_codbod & " AND b_tomainv.tin_fectom = " & v_fecinv & ""
    
    Else
       
       vg_db.Execute "UPDATE b_productospmpdia SET b_productospmpdia.ppd_saldo = b_tomainv.tin_stofis FROM b_productospmpdia, b_tomainv WHERE b_productospmpdia.ppd_codpro = b_tomainv.tin_codpro " & _
                     "AND b_productospmpdia.ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & " AND b_productospmpdia.ppd_cencos = '" & MuestraCasino(1) & "' AND b_tomainv.tin_codbod = " & vg_codbod & " AND b_tomainv.tin_fectom = " & v_fecinv & ""
    
    
    End If
    vg_db.Execute "UPDATE b_tomainv SET tin_tipinv = '" & fg_codigocbo(Combo1, 1, 1, 0) & "' WHERE tin_codbod = " & v_codbod & " AND tin_fectom = " & v_fecinv & ""
    
    '----------------------------------------------------------
    If Check2.Value = 1 Then
       
       Dim opord As String, codtip As Long
       opord = IIf(Check2.Value = 0, " ORDER BY pro.pro_nombre", " ORDER BY pro.pro_ctacon, pro.pro_codtip, pro.pro_nombre")
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS1.Open "SELECT tin.tin_stofis, tin.tin_stosis, pro.pro_codigo, pro.pro_nombre, pro.pro_codtip, pro.pro_ctacon, uni.uni_nombre, tin.tin_propon " & _
                "FROM b_tomainv tin, b_productos pro, a_unidad uni " & _
                "WHERE uni.uni_codigo = pro.pro_coduni AND pro.pro_codigo = tin.tin_codpro " & _
                "AND tin_fectom = " & v_fecinv & " AND tin_codbod = " & v_codbod & "  " & opord & "", vg_db, adOpenStatic
       codtip = 0
       vaSpread1.Visible = False
       vaSpread1.MaxRows = 0
       If Not RS1.EOF Then
          
          i = 1
          
          Do While Not RS1.EOF
             
             vaSpread1.MaxRows = i: vaSpread1.Row = i
             
             If RS1!pro_codtip <> codtip And Check2.Value = 1 Then
                
                vaSpread1.Col = 1: vaSpread1.Lock = True: vaSpread1.text = ""
                vaSpread1.Col = 2: vaSpread1.Lock = True: vaSpread1.Font.Bold = True: vaSpread1.text = fg_BuscaenArbol(RS1!pro_codtip, "a_tipopro", "tip_codigo")
                vaSpread1.Col = 3: vaSpread1.Lock = True: vaSpread1.text = ""
                vaSpread1.Col = 4: vaSpread1.Lock = True: vaSpread1.text = ""
                vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.Lock = True: vaSpread1.text = ""
                vaSpread1.Col = 6: vaSpread1.Lock = True: vaSpread1.text = ""
                vaSpread1.Col = 7: vaSpread1.Lock = True: vaSpread1.text = ""
                codtip = RS1!pro_codtip
                i = i + 1: vaSpread1.MaxRows = i: vaSpread1.Row = i
             
             End If
             
             vaSpread1.Col = 1: vaSpread1.text = RS1!pro_codigo
             vaSpread1.Col = 2: vaSpread1.text = RS1!pro_nombre
             vaSpread1.Col = 3: vaSpread1.text = RS1!uni_nombre
             vaSpread1.Col = 4: vaSpread1.text = IIf(modo = "A", Format(0, fg_Pict(9, vg_DCa)), Format(RS1!tin_stosis, fg_Pict(9, vg_DCa)))
             vaSpread1.Col = 5: vaSpread1.text = Format(RS1!tin_stofis, fg_Pict(9, vg_DCa))
             vaSpread1.Col = 6: vaSpread1.text = Format(RS1!tin_propon, fg_Pict(9, vg_DPr))
             vaSpread1.Col = 7: vaSpread1.text = Format(Format(RS1!tin_stofis, fg_Pict(9, vg_DCa)) * Format(IIf(IsNull(RS1!tin_propon), 0, RS1!tin_propon), fg_Pict(9, vg_DPr)), fg_Pict(9, vg_DPr))
             RS1.MoveNext: i = i + 1
          
          Loop
          
          vaSpread1.SetActiveCell 5, 1
       
       End If
       vaSpread1.Visible = True
       RS1.Close: Set RS1 = Nothing
       If vaSpread1.SearchCol(1, 0, vaSpread1.MaxRows, Trim(vg_codigo), SearchFlagsNone) <> -1 Then
          
          vaSpread1.Row = vaSpread1.SearchCol(1, 0, vaSpread1.MaxRows, Trim(vg_codigo), SearchFlagsNone)
          vaSpread1.SetActiveCell 5, vaSpread1.Row
       
       End If
    
    Else
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
   
       vaSpread1.Row = vaSpread1.ActiveRow
       RS1.Open "SELECT pro.pro_codigo, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = pro.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS ppd_propon, pro.pro_nombre, pro.pro_codtip, pro.pro_ctacon, uni.uni_nombre " & _
                "FROM b_productos AS pro, a_unidad AS uni " & _
                "WHERE pro.pro_coduni = uni.uni_codigo AND pro.pro_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic
       If Not RS1.EOF Then
          
          i = vaSpread1.MaxRows + 1
          
          Do While Not RS1.EOF
             
             vaSpread1.MaxRows = i: vaSpread1.Row = i
             vaSpread1.Col = 1: vaSpread1.text = RS1!pro_codigo
             vaSpread1.Col = 2: vaSpread1.text = RS1!pro_nombre
             vaSpread1.Col = 3: vaSpread1.text = RS1!uni_nombre
             vaSpread1.Col = 4: vaSpread1.text = Format(0, fg_Pict(9, vg_DCa))
             vaSpread1.Col = 5: vaSpread1.text = Format(0, fg_Pict(9, vg_DCa))
             vaSpread1.Col = 6: vaSpread1.text = Format(RS1!ppd_propon, fg_Pict(9, vg_DPr))
             vaSpread1.Col = 7: vaSpread1.text = Format(0, fg_Pict(9, vg_DPr))
             RS1.MoveNext: i = i + 1
          
          Loop
       
       End If
       RS1.Close: Set RS1 = Nothing
       vaSpread1.Col = 5: vaSpread1.Row = vaSpread1.MaxRows
       vaSpread1.SetActiveCell 5, vaSpread1.MaxRows
       If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
    
    End If

Case 2
    
    If vaSpread1.MaxRows = 0 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    If Trim(vaSpread1.text) = "" Then MsgBox "No puede eliminar familia producto...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    modo = "E"
    vg_db.Execute "DELETE b_tomainv FROM b_tomainv WHERE tin_fectom = " & Val(Format(Date1(0).text, "yyyymmdd")) & " AND tin_codbod = " & Val(fg_codigocbo(Combo1, 0, 10, "")) & " " & _
                  "AND tin_codpro = '" & Trim(vaSpread1.text) & "'"
    i = vaSpread1.Row
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    
    If (vaSpread1.ActiveRow - 1) >= 0 Then
        
        vaSpread1.Row = i: vaSpread1.Col = 1
        If Trim(vaSpread1.text) <> "" Then Exit Sub
        vaSpread1.Row = IIf(vaSpread1.ActiveRow - 1 = 0, 1, (i - 1))
        vaSpread1.Col = 1
        If Trim(vaSpread1.text) = "" Then vaSpread1.DeleteRows (vaSpread1.Row), 1: vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    
    End If
    
    If vaSpread1.MaxRows = 0 Then
    
'       '-------> INI: Mover estado a la tabla parametro toma inventario
'       vg_db.Execute "update a_param set par_valor = '0' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
'       '-------> FIN: Mover estado a la tabla parametro toma inventario
'
    End If
    
    If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus

End Select

Exit Sub
Error_Mover1:
    fg_descarga
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, MsgTitulo
    Resume Next
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

If Row <> 0 Then Exit Sub
vaSpread1.Col = -1: vaSpread1.Row = -1: vaSpread1.ForeColor = RGB(0, 0, 0)
vaSpread1.Col = Col: vaSpread1.Row = -1: vaSpread1.ForeColor = RGB(255, 0, 0)

End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

vaSpread1.Col = Col: vaSpread1.Row = Row
If Round(vaSpread1.text, vg_DCa) < 0 Then vaSpread1.text = 0: Exit Sub
If modo = "" Or modo = "E" Then modo = "M"
Gl_Ac_Botones Me, 6, 0, modo

End Sub

Function GenerarArcInvSap(codbod As Long, FecInv As Variant, op As Boolean) As Boolean

Dim nomarch As String, nomarchzip As String, mdir As String, parametro1 As String, parametro2 As String, parametro3 As String, sql1 As String
Dim Fecha As String, vRet As Variant, fecenv As String
Dim socsap As String, cladoc As String, n_a As String, bkpf_bukrs As String, bkpf_blart As String, bkpf_budat As String, bkpf_bldat As String, bkpf_xblnr As String
Dim bkpf_bktxt As String, bkpf_waers As String, bseg_newbs As String, bseg_newko As String, bseg_wrbtr As String, bseg_zuonr As String, bseg_sgtxt As String, bseg_kostl As String
Dim n_acodimpto As String, n_actaimpto As String, n_amonimp As String, n_aimprecu As String, n_aotrimp As String
Dim invali As Long, invdes As Long, codigo As Long, numlin As Long, numero As Long, Sql As String
Dim RS1 As New ADODB.Recordset
Const INTERVALO_EN_MINUTOS As Integer = 5

On Error GoTo Man_EnvioInventario

Fecha = FecInv
modo = "E"
GenerarArcInvSap = False
socsap = "": cladoc = "": n_a = "": bkpf_bukrs = "": bkpf_blart = "": bkpf_budat = "": bkpf_bldat = "": bkpf_xblnr = "": bkpf_bktxt = "": bkpf_waers = "": bseg_newbs = "": bseg_newko = "": bseg_wrbtr = "": bseg_zuonr = "": bseg_sgtxt = "": bseg_kostl = ""
n_acodimpto = "": n_actaimpto = "": n_amonimp = "": n_aimprecu = "": n_aotrimp = ""
DoEvents

'------- Abrir mensaje de text
If ValidarOpEnvio(MuestraCasino(1), 2) Then
   
   Frame6.Visible = True
   Me.Refresh
   Text1(0).text = FechaHora & "CENCO : " & MuestraCasino(1) & " - " & MuestraCasino(2) & VgLinea
   Text1(0).text = Text1(0).text & FechaHora & "USUARIO : " & Environ("USERNAME") & VgLinea
   Text1(0).text = Text1(0).text & FechaHora & "Inicio del Proceso. " & IIf(op, "Inventario : " & FecInv, "Anulación Inventario : " & FecInv) & VgLinea

   '-------> Validar si existe usuario sap
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS1 = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'sapusu'")
   If RS1.EOF Then
      
      RS1.Close: Set RS1 = Nothing
      Text1(0).text = Text1(0).text & FechaHora & "No tiene creado usuario, para Web Service" & VgLinea
      GenerarArcInvSap = False: Exit Function
   
   ElseIf IsNull(RS1!par_valor) Or Trim(RS1!par_valor) = "" Then
      
      RS1.Close: Set RS1 = Nothing
      Text1(0).text = Text1(0).text & FechaHora & "Usuario fue borrado, para Web Service" & VgLinea
      GenerarArcInvSap = False: Exit Function
   
   End If
   RS1.Close: Set RS1 = Nothing

   '-------> Validar si existe password sap
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS1 = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'sappas'")
   If RS1.EOF Then
      
      RS1.Close: Set RS1 = Nothing
      Text1(0).text = Text1(0).text & FechaHora & "No tiene creado password, para Web Service" & VgLinea
      GenerarArcInvSap = False: Exit Function
   
   ElseIf IsNull(RS1!par_valor) Or Trim(RS1!par_valor) = "" Then
      
      RS1.Close: Set RS1 = Nothing
      Text1(0).text = Text1(0).text & FechaHora & "Password fue borrada, para Web Service" & VgLinea
      GenerarArcInvSap = False: Exit Function
   
   End If
   RS1.Close: Set RS1 = Nothing

End If

'------- Traer sociedad del contrato
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("SELECT * FROM b_clientes WHERE cli_codigo = '" & MuestraCasino(1) & "' AND cli_tipo = 0")
If RS1.EOF Or IsNull(RS1!cli_socsap) Or Trim(RS1!cli_socsap) = "" Then
   
   RS1.Close: Set RS1 = Nothing
   Text1(0).text = Text1(0).text & FechaHora & "No tiene asignado la sociedad de SAP." & VgLinea
   GenerarArcInvSap = False
   Exit Function

End If

bkpf_bukrs = Trim(RS1!cli_socsap)
bkpf_bktxt = "Inv. Ceco:" & Trim(RS1!cli_codigo) & " " & Mid(Meses(Fecha), 1, 3) & " " & Mid(fg_pone_cero(Fecha, 8), 7, 4) 'Trim(RS1!cli_nombre)
RS1.Close: Set RS1 = Nothing

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = IIf(op, " (b.tin_envsap='0' OR (b.tin_envsap) IS NULL) ", " (b.tin_envsap='1' OR (b.tin_envsap) IS NULL) ")

Set RS1 = vg_db.Execute("SELECT a.pro_ctacon, ROUND(SUM(b.tin_stofis*b.tin_propon),0) AS cosinv FROM b_productos a, b_tomainv b " & _
         "WHERE  b.tin_codpro = a.pro_codigo AND (a.pro_ctacon NOT IN ('" & fg_CambiaChar(GetParametro("ctagastos"), ";", "','") & "') OR a.pro_ctacon NOT IN ('" & fg_CambiaChar(GetParametro("ctagastos2"), ";", "','") & "')) " & _
         "AND    b.tin_codbod = " & codbod & " AND b.tin_fectom = " & Format(FecInv, "yyyymmdd") & " " & _
         "AND    b.tin_stofis <> 0 AND b.tin_propon <> 0 AND " & Sql & " GROUP BY a.pro_ctacon")
If Not RS1.EOF Then
   
   Do While Not RS1.EOF
      
      If RS1!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
         
         invali = Round(invali + RS1!cosinv, vg_DPr)
      
      ElseIf RS1!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
         
         invdes = Round(invdes + RS1!cosinv, vg_DPr)
      
      End If
      RS1.MoveNext
   
   Loop

Else
   fg_descarga
   RS1.Close: Set RS1 = Nothing
   Text1(0).text = Text1(0).text & FechaHora & "No existe datos registrado, en inventario..." & VgLinea
   GenerarArcSap = False
   Exit Function
End If
RS1.Close: Set RS1 = Nothing

'------- generar ultimo codigo tabla envio inventario
codigo = 0: numlin = 1
Set RS1 = vg_db.Execute("SELECT inv_codigo FROM sap_inv ORDER BY inv_codigo DESC")
If Not RS1.EOF Then RS1.MoveFirst: codigo = RS1!inv_codigo + 1 Else codigo = 1
RS1.Close: Set RS1 = Nothing

'------- Mover parametro Web Service
parametro1 = IIf(op, "2", "3")
parametro2 = codigo
parametro3 = MuestraCasino(1)

bkpf_budat = Format(FecInv, "DDMMYYYY")
bkpf_bldat = Format(FecInv, "DDMMYYYY")
bkpf_xblnr = MuestraCasino(1)
bkpf_waers = vg_tipmonsap
bkpf_blart = "TP"

If invali > 0 Then
   
   '------- Mover Alimentación cuenta 124010
   n_a = "X"
   bseg_newbs = IIf(op, "40", "50")
   bseg_newko = "124010"
   bseg_wrbtr = invali
   bseg_zuonr = MuestraCasino(1)
   bseg_sgtxt = "Inventario Final " & Trim(Mid(MuestraCasino(2), 1, 24)) & " (" & Trim(MuestraCasino(1)) & ")"
   bseg_kostl = MuestraCasino(1)
   bseg_kostl = MuestraCasino(1)
   n_acodimpto = "NA"
   n_actaimpto = ""
   n_amonimp = ""
   n_aimprecu = ""
   n_aotrimp = ""
   
   '------- Grabar encabezado inventario
   vg_db.Execute "INSERT INTO sap_inv VALUES (" & codigo & ", " & numlin & ", '" & n_a & "', '" & bkpf_bukrs & "', '" & bkpf_blart & "', '" & bkpf_budat & "', " & _
                 "'" & bkpf_bldat & "', '" & bkpf_xblnr & "', '" & bkpf_bktxt & "', '" & bkpf_waers & "', '" & bseg_newbs & "', " & _
                 "'" & bseg_newko & "', '" & bseg_wrbtr & "', '" & bseg_zuonr & "', '" & bseg_sgtxt & "', '" & bseg_kostl & "', " & _
                 "'" & n_acodimpto & "', '" & n_actaimpto & "', '" & n_amonimp & "', '" & n_aimprecu & "', '" & n_aotrimp & "')"
   '------- Mover Alimentación cuenta 410001
   n_a = ""
   bseg_newbs = IIf(op, "50", "40")
   bseg_newko = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','"))
   bseg_zuonr = MuestraCasino(1)
   bseg_sgtxt = "Inventario Final " & Trim(Mid(MuestraCasino(2), 1, 24)) & " (" & Trim(MuestraCasino(1)) & ")" 'MuestraCasino(2)
   bseg_kostl = MuestraCasino(1)
   
   '------- Grabar detalle inventario
   numlin = numlin + 1
   vg_db.Execute "INSERT INTO sap_inv VALUES (" & codigo & ", " & numlin & ", '" & n_a & "', '" & bkpf_bukrs & "', '" & bkpf_blart & "', '" & bkpf_budat & "', " & _
                 "'" & bkpf_bldat & "', '" & bkpf_xblnr & "', '" & bkpf_bktxt & "', '" & bkpf_waers & "', '" & bseg_newbs & "', " & _
                 "'" & bseg_newko & "', '" & bseg_wrbtr & "', '" & bseg_zuonr & "', '" & bseg_sgtxt & "', '" & bseg_kostl & "', " & _
                 "'" & n_acodimpto & "', '" & n_actaimpto & "', '" & n_amonimp & "', '" & n_aimprecu & "', '" & n_aotrimp & "')"
End If
If invdes > 0 Then
   
   '------- Mover Alimentación cuenta 124010
   n_a = IIf(invali > 0, "", "X")
   bseg_newbs = IIf(op, "40", "50")
   bseg_newko = "124020"
   bseg_wrbtr = invdes
   bseg_zuonr = MuestraCasino(1)
   bseg_sgtxt = "Inventario Final " & Trim(Mid(MuestraCasino(2), 1, 24)) & " (" & Trim(MuestraCasino(1)) & ")" 'MuestraCasino(2)
   bseg_kostl = MuestraCasino(1)
   n_acodimpto = "NA"
   n_actaimpto = ""
   n_amonimp = ""
   n_aimprecu = ""
   n_aotrimp = ""
   numlin = IIf(invali > 0, (numlin + 1), 1)
   '------- Grabar encabezado inventario
   vg_db.Execute "INSERT INTO sap_inv VALUES (" & codigo & ", " & numlin & ", '" & n_a & "', '" & bkpf_bukrs & "', '" & bkpf_blart & "', '" & bkpf_budat & "', " & _
                 "'" & bkpf_bldat & "', '" & bkpf_xblnr & "', '" & bkpf_bktxt & "', '" & bkpf_waers & "', '" & bseg_newbs & "', " & _
                 "'" & bseg_newko & "', '" & bseg_wrbtr & "', '" & bseg_zuonr & "', '" & bseg_sgtxt & "', '" & bseg_kostl & "', " & _
                 "'" & n_acodimpto & "', '" & n_actaimpto & "', '" & n_amonimp & "', '" & n_aimprecu & "', '" & n_aotrimp & "')"
   
   '------- Mover Alimentación cuenta 410001
   n_a = ""
   bseg_newbs = IIf(op, "50", "40")
   bseg_newko = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))
   bseg_zuonr = MuestraCasino(1)
   bseg_sgtxt = "Inventario Final " & Trim(Mid(MuestraCasino(2), 1, 24)) & " (" & Trim(MuestraCasino(1)) & ")" 'MuestraCasino(2)
   bseg_kostl = MuestraCasino(1)
   n_acodimpto = "NA" '"C1" '""
   n_actaimpto = ""
   n_amonimp = ""
   n_aimprecu = ""
   n_aotrimp = ""
   '------- Grabar detalle inventario
   numlin = numlin + 1
   vg_db.Execute "INSERT INTO sap_inv VALUES (" & codigo & ", " & numlin & ", '" & n_a & "', '" & bkpf_bukrs & "', '" & bkpf_blart & "', '" & bkpf_budat & "', " & _
                 "'" & bkpf_bldat & "', '" & bkpf_xblnr & "', '" & bkpf_bktxt & "', '" & bkpf_waers & "', '" & bseg_newbs & "', " & _
                 "'" & bseg_newko & "', '" & bseg_wrbtr & "', '" & bseg_zuonr & "', '" & bseg_sgtxt & "', '" & bseg_kostl & "', " & _
                 "'" & n_acodimpto & "', '" & n_actaimpto & "', '" & n_amonimp & "', '" & n_aimprecu & "', '" & n_aotrimp & "')"

End If
'------ Grabar log proceso inventario
fecenv = IIf(vg_tipbase = "1", Format(Date, "dd-mm-yyyy") & " " & Format(Time, "h:m:s"), Format(Date, "yyyymmdd") & " " & Format(Time, "h:m:s"))
numero = 0
Set RS1 = vg_db.Execute("SELECT numero FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' ORDER BY numero DESC")
If Not RS1.EOF Then RS1.MoveFirst: numero = RS1!numero + 1 Else numero = 1
RS1.Close: Set RS1 = Nothing
vg_db.Execute "INSERT INTO log_procesos (cencos, numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mensaje, envio) " & _
              "VALUES ('" & MuestraCasino(1) & "', " & numero & ", '" & fecenv & "', '" & IIf(op, "2", "3") & "', '', '', '" & Format(FecInv, "YYYYMMDD") & "',  0, '0', '', " & codigo & ")"
vg_codigo_Inv = codigo

'------- Proceso envio Web Service
If ValidarOpEnvio(MuestraCasino(1), 2) Then
   
   DoEvents
   If vg_tipbase = "1" Then
      
      vRet = Shell(Trim(dir_trabajo) & "WsSapPortal.exe " & Trim(parametro1) & "|" & Trim(parametro2) & "|" & Trim(parametro3) & "|" & LCase(App.Path) & "\" & "|" & "" & "|" & "" & "|" & "" & "|" & "" & "|")
   
   Else
      
      vRet = Shell(Trim(dir_trabajo) & "WsSapPortal.exe " & Trim(parametro1) & "|" & Trim(parametro2) & "|" & Trim(parametro3) & "|" & LCase(App.Path) & "\" & "|" & vg_SqlNSvr & "|" & vg_SqlBase & "|" & vg_SqlNUsr & "|" & vg_SqlPass & "|")
   
   End If
   If vRet = 0 Then Text1(0).text = Text1(0).text & "Proceso cancelado, no hay comunicación con Web Service": GenerarArcInvSap = False: Exit Function

   DoEvents
   Set RS1 = vg_db.Execute("SELECT * FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND numero = " & numero & " AND tipo_proceso = '" & IIf(op, "2", "3") & "' AND estado = '0'")
   Do While Not RS1.EOF
      
      DoEvents
      RS1.Close: Set RS1 = Nothing
      Set RS1 = vg_db.Execute("SELECT * FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND numero = " & numero & " AND tipo_proceso = '" & IIf(op, "2", "3") & "' AND estado = '0'")
   
   Loop
   RS1.Close: Set RS1 = Nothing

   '------- Proceso de estado de envio
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS1 = vg_db.Execute("SELECT * FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND numero = " & numero & " AND tipo_proceso = '" & IIf(op, "2", "3") & "'")
   If Not RS1.EOF Then
      
      StrMensaje = Trim(RS1!mensaje)
      
      If Len(StrMensaje) <> 0 Then
         
         Text1(0).text = Text1(0).text & VgLinea
         Text1(0).text = Text1(0).text & FechaHora & IIf(RS1!estado = "3", "Mensaje Error : ", "Mensaje SAP : ") & VgLinea
         Text1(0).text = Text1(0).text & FechaHora & "---------------------------------------------------------" & VgLinea
         
         Do While InStr(StrMensaje, ";") <> 0 And InStr(StrMensaje, ";") <> 1
            
            If StrMensaje <> "" Then
               
               nommen = Mid(StrMensaje, 1, InStr(StrMensaje, "|") - 1)
               StrMensaje = Mid(StrMensaje, InStr(StrMensaje, "|") + 1)
               Text1(0).text = Text1(0).text & FechaHora & Trim(nommen) & VgLinea
               If InStr(nommen, "timed out") <> 0 Or InStr(nommen, "No esta conectado a la internet") <> 0 Then RS1.Close: Set RS1 = Nothing: GenerarArcInvSap = False: Exit Function
            
            End If
         
         Loop
         
         Text1(0).text = Text1(0).text & FechaHora & "---------------------------------------------------------" & VgLinea
         Text1(0).text = Text1(0).text & VgLinea
         If RS1!estado = "2" Or RS1!estado = "0" Or RS1!estado = "3" Then RS1.Close: Set RS1 = Nothing: GenerarArcInvSap = False: Exit Function
      
      End If
   
   End If
   RS1.Close: Set RS1 = Nothing

End If

If ValidarOpEnvio(MuestraCasino(1), 5) Then
   
   vg_db.Execute "update log_procesos set estado = '1' where cencos = '" & MuestraCasino(1) & "' AND numero = " & numero & " AND tipo_proceso = '" & IIf(op, "2", "3") & "' AND estado = '0'"

End If

'------ Actualizar log proceso anulado registro
vg_db.Execute "UPDATE log_procesos SET anulado = '1' WHERE cencos = '" & MuestraCasino(1) & "' AND tipo_proceso = '" & IIf(op, "3", "2") & "' AND num_documento = '" & Format(FecInv, "YYYYMMDD") & "'"
'------- Grabar tabla b_totcompras si se genero sin problema
vg_db.Execute "UPDATE b_tomainv SET tin_envsap = '" & IIf(op, "1", "0") & "' WHERE tin_fectom = " & Format(FecInv, "yyyymmdd") & " AND tin_codbod = " & vg_codbod & ""

GenerarArcInvSap = True

Exit Function
Man_EnvioInventario:
If Err = 53 Then
   Text1(0).text = Text1(0).text & FechaHora & "No existe Ejecutable de envio..." & VgLinea
   vg_db.Execute "UPDATE log_procesos SET estado = '4', mensaje='No existe ejecutable, para procesar Web Service' WHERE cencos = '" & MuestraCasino(1) & "' AND tipo_proceso='" & IIf(op, "2", "3") & "' AND numero = " & numero & ""
   GenerarArcInvSap = False
   Exit Function
End If

End Function
