VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form M_WebToMDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WebToMDB"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12795
   Icon            =   "M_WebToMDB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraParamExp 
      Caption         =   "Parámetros de Exportación"
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
      Height          =   1545
      Left            =   150
      TabIndex        =   5
      Top             =   1110
      Width           =   12500
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "M_WebToMDB.frx":08CA
         Left            =   210
         List            =   "M_WebToMDB.frx":08F2
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   555
         Width           =   1680
      End
      Begin VB.ComboBox cboCentralCompra 
         Height          =   315
         Left            =   3450
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   570
         Width           =   3075
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6645
         ScaleHeight     =   255
         ScaleWidth      =   3180
         TabIndex        =   6
         Top             =   570
         Width           =   3240
         Begin EditLib.fpBoolean chkTipoSolicitud 
            Height          =   300
            Index           =   0
            Left            =   45
            TabIndex        =   7
            Tag             =   "1"
            Top             =   0
            Width           =   810
            _Version        =   196608
            _ExtentX        =   1429
            _ExtentY        =   529
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ThreeDInsideStyle=   0
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   0
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            AutoToggle      =   -1  'True
            BooleanStyle    =   0
            ToggleFalse     =   ""
            TextFalse       =   "Normal"
            BooleanPicture  =   2
            AlignPictureH   =   3
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   0   'False
            AlignTextH      =   0
            AlignTextV      =   1
            ToggleTrue      =   ""
            TextTrue        =   "Normal"
            Value           =   0
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   -1  'True
            BackColor       =   16777215
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   "Normal"
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin EditLib.fpBoolean chkTipoSolicitud 
            Height          =   300
            Index           =   1
            Left            =   1065
            TabIndex        =   8
            Tag             =   "3"
            Top             =   0
            Width           =   780
            _Version        =   196608
            _ExtentX        =   1376
            _ExtentY        =   529
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ThreeDInsideStyle=   0
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   0
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            AutoToggle      =   -1  'True
            BooleanStyle    =   0
            ToggleFalse     =   ""
            TextFalse       =   "Extra"
            BooleanPicture  =   2
            AlignPictureH   =   3
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   0   'False
            AlignTextH      =   0
            AlignTextV      =   1
            ToggleTrue      =   ""
            TextTrue        =   "Extra"
            Value           =   0
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   -1  'True
            BackColor       =   16777215
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   "Extra"
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin EditLib.fpBoolean chkTipoSolicitud 
            Height          =   300
            Index           =   2
            Left            =   2025
            TabIndex        =   9
            Tag             =   "4"
            Top             =   0
            Width           =   1065
            _Version        =   196608
            _ExtentX        =   1879
            _ExtentY        =   529
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ThreeDInsideStyle=   0
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   0
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            AutoToggle      =   -1  'True
            BooleanStyle    =   0
            ToggleFalse     =   ""
            TextFalse       =   "Anulación"
            BooleanPicture  =   2
            AlignPictureH   =   3
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   0   'False
            AlignTextH      =   0
            AlignTextV      =   1
            ToggleTrue      =   ""
            TextTrue        =   "Anulación"
            Value           =   0
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   -1  'True
            BackColor       =   16777215
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   "Anulación"
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
      End
      Begin EditLib.fpDoubleSingle txtYear 
         Height          =   315
         Left            =   1980
         TabIndex        =   12
         Top             =   555
         Width           =   600
         _Version        =   196608
         _ExtentX        =   1058
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
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
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
         Text            =   "0"
         DecimalPlaces   =   0
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDoubleSingle txtSemana 
         Height          =   315
         Left            =   2700
         TabIndex        =   13
         Top             =   555
         Width           =   615
         _Version        =   196608
         _ExtentX        =   1085
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
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
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
         Text            =   "0"
         DecimalPlaces   =   0
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
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
         Left            =   2820
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5940
         _Version        =   196608
         _ExtentX        =   10477
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
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   2
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
         ControlType     =   3
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
         Appearance      =   0
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
         Caption         =   "Seleccione Directorio *.MDB"
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
         TabIndex        =   23
         Top             =   1155
         Width           =   2445
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
         Left            =   2700
         TabIndex        =   18
         Top             =   345
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Left            =   2010
         TabIndex        =   17
         Top             =   345
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         Left            =   225
         TabIndex        =   16
         Top             =   345
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Central de Compra"
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
         Index           =   10
         Left            =   3450
         TabIndex        =   15
         Top             =   345
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Solicitud"
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
         Left            =   6645
         TabIndex        =   14
         Top             =   345
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   12735
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   12795
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "WebToMDB - Exportación de datos Web Pedidos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   765
         TabIndex        =   3
         Top             =   195
         Width           =   4035
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   105
         Picture         =   "M_WebToMDB.frx":095B
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   12
      Left            =   0
      TabIndex        =   1
      Top             =   1035
      Width           =   13755
   End
   Begin EditLib.fpBoolean chkAll 
      Height          =   165
      Left            =   675
      TabIndex        =   0
      Tag             =   "0"
      Top             =   2985
      Width           =   210
      _Version        =   196608
      _ExtentX        =   370
      _ExtentY        =   291
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      AutoToggle      =   -1  'True
      BooleanStyle    =   0
      ToggleFalse     =   ""
      TextFalse       =   ""
      BooleanPicture  =   2
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   0
      AlignTextV      =   1
      ToggleTrue      =   ""
      TextTrue        =   ""
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   ""
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin MSComctlLib.ImageList Icons 
      Index           =   0
      Left            =   7395
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_WebToMDB.frx":1225
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_WebToMDB.frx":1AFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_WebToMDB.frx":23D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_WebToMDB.frx":2773
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_WebToMDB.frx":4335
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpread.vaSpread spdData 
      Height          =   4680
      Left            =   150
      TabIndex        =   4
      Top             =   2880
      Width           =   12495
      _Version        =   393216
      _ExtentX        =   22040
      _ExtentY        =   8255
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
      MaxCols         =   9
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "M_WebToMDB.frx":464F
   End
   Begin MSComctlLib.Toolbar tlbAction 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   660
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "Icons(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar Pedidos"
            Object.Tag             =   "SearchData"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exporta a Excel antes de procesar"
            Object.Tag             =   "ExportExcel"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Procesar"
            Object.Tag             =   "ExportData"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   "Cancel"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   "Close"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prbStatus 
      Height          =   405
      Left            =   375
      TabIndex        =   20
      Top             =   5835
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   600
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "lblStatus"
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
      Left            =   390
      TabIndex        =   21
      Top             =   5595
      Width           =   750
   End
End
Attribute VB_Name = "M_WebToMDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Msgtitulo As String
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim CNN As New ADODB.Connection
Dim CMD As New ADODB.Command
Dim db7 As Database
Dim fila_selec As Long
Dim flagBaja As Boolean


Private Function blnExportPedidos() As Boolean

On Error GoTo blnExportPedidos_Err

Dim lngRow As Long
Dim strFileName As String
Dim strLocalFileName As String
Dim strFileNameMDB As String
Dim strLineReg As String
Dim i As Long
Dim intPorcentResp  As Long
Dim strCurrentDateServerSQL As String
Dim strCurrentTimeServerSQL As String
Dim intCasino As String
Dim lngSequence As Long
Dim lngCurrentReg As Long
Dim ccompra As String
Dim tipo As Long
Dim anomes As String
Dim Semana As Long
Dim strSQL As String
Dim estarc As Integer
Dim StrFileNameLdb As String

    blnExportPedidos = False

    If Trim(fpText1.text) = "" Then MsgBox "Debe seleccionar Directorio", vbCritical, Msgtitulo: Exit Function
    If Dir(fpText1.text, vbDirectory) = "" Then MsgBox "No existe directorio seleccionado", vbCritical, Msgtitulo: Exit Function
    
    For lngRow = 1 To spdData.MaxRows
        spdData.Row = lngRow: spdData.Col = 1: If spdData.Value = 1 Then Exit For
    Next
    
    If lngRow > spdData.MaxRows Then MsgBox "Debe seleccionar al menos un pedido.", vbCritical, Msgtitulo: Exit Function
    If MsgBox("¿ Desea Exportar Pedidos seleccionados ?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Function
    
    If (tlbAction.Buttons(2).Value = tbrPressed) Then
        If (Not blnRptExportToExcel()) Then
            If MsgBox("No fue posible exportar datos a Excel." & Chr(13) & _
                      "¿ Desea continuar con el proceso ?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Function
        End If
    End If
    
'    MesD = fg_pone_cero(Str(cboMes.ListIndex + 1), 2)
'    AnoD = fg_pone_cero(Str(txtYear.Value), 4)
'    PerD = AnoD & MesD
'    SemD = Val(txtSemana.Value)

    Screen.MousePointer = vbHourglass
    DoEvents

    fraParamExp.Enabled = False

    DoEvents
    
    vg_dbpedweb.Execute "pedweb_d_bajapedido"
    For lngRow = 1 To spdData.MaxRows
        spdData.Row = lngRow: spdData.Col = 1
        If spdData.Value = 1 Then
           spdData.Col = 2
           vg_dbpedweb.Execute "pedweb_i_bajapedido " & Trim(spdData.Value) & ""
        End If
    Next
    
    '-------> Crear directorio DatosTxt
    If Dir(dir_trabajo & "DatosTxt", vbDirectory) = "" Then MkDir dir_trabajo & "DatosTxt"
    '-------> Fin crear directorio DatosTxt

    strFileName = dir_trabajo & ("DatosTxt\BajaData.txt")
    If Dir(strFileName) <> "" Then Kill (strFileName)
    
    intPorcentResp = 0
    
    Set RS1 = vg_dbpedweb.Execute("pedweb_s_traerfechahora")
    If Not RS1.EOF Then
       strCurrentDateServerSQL = RS1(0)
       strCurrentTimeServerSQL = RS1(1)
    End If
    RS1.Close: Set RS1 = Nothing
    
    Set RS1 = vg_dbpedweb.Execute("pedweb_s_traersolicitudcompras")
'    lngCount = rsData.RecordCount
    Open strFileName For Output As #1
    intCasino = 0
    lngSequence = 0
    lngCurrentReg = 0
    
    While Not RS1.EOF
        lngCurrentReg = lngCurrentReg + 1
    
        If (intCasino = CInt(RS1("Casino"))) Then
            lngSequence = lngSequence + 1
        Else
            lngSequence = 1
        End If
        
        intCasino = CInt(RS1!Casino)
        
        Set RS2 = vg_dbpedweb.Execute("pedweb_s_traerdetallesolicitudcompras " & RS1!codigo & "")
        If Not RS2.EOF Then
'            If IsNull(RS2!Fecha2) Or Trim(RS2!Fecha2) = "" Then
'               MsgBox "Casino no fue procesado : " & Trim(RS1!centrocosto) & " Por tener fecha web en blanco o nula", vbCritical, Msgtitulo
'            Else
               Print #1, "CC" & Trim(RS1!centrocosto) & Trim(RS1!CentralDeCompra) & "_" & lngSequence
            
               Print #1, "create table CADFIL (CADFIL_CDFIL char(10), CADFIL_NMFIL char(50))"
               Print #1, "create table SOLFIL (SOLFIL_IDSOL Integer, CADFIL_IDFIL Integer, TIPSOL_IDSOL Integer, SOLFIL_DTSOL Datetime, SOLFIL_DTREF Char(6), SOLFIL_NRSEM Integer, TIME_STAMP Datetime)"
               Print #1, "create table SOLITE (SOLFIL_IDSOL Integer, CPOPRO_CDPRO Char(20), SOLITE_DTENT Datetime, SOLITE_QTSOL Double, SOLITE_FLPRO Char(1), SOLITE_FLATU Integer, SOLITE_FLCPA Integer, SOLFIL_DTADI Datetime)"
               Print #1, "create table TABCEN (TABCEN_CDCEN char(4), TABCEN_DSCEN char(50))"
               Print #1, "create table TABPAR (TABPAR_NRVFL char(5))"

               Print #1, "INSERT INTO CadFil VALUES( '" & Trim(RS1!centrocosto) & "', '" & Trim(RS1!Nombre) & "' )"
               Print #1, "INSERT INTO TabCen VALUES( '" & Trim(RS1!CentralDeCompra) & "', 'XX' )"
               Print #1, "INSERT INTO TabPar VALUES( '1.9' )"
            
               ccompra = Trim(RS1!CentralDeCompra)
               tipo = RS2!tipo
               anomes = RS2!anomes
               Semana = RS2!Semana
            
                        
               strSQL = "INSERT INTO SolFil VALUES( "
               strSQL = strSQL & RS2!codigo & ", "
               strSQL = strSQL & Trim(RS2!Casino) & ", "
               strSQL = strSQL & RS2!tipo & ", "
               strSQL = strSQL & "'" & strCurrentDateServerSQL & "', "
               strSQL = strSQL & "'" & RS2!anomes & "', "
               If ccompra = "SA" And tipo = 1 And anomes = "200807" And Semana = 31 Then
                  strSQL = strSQL & "27, "
               Else
                  strSQL = strSQL & RS2!Semana & ", "
               End If
               strSQL = strSQL & "'" & strCurrentDateServerSQL & " " & strCurrentTimeServerSQL & "' )"
               Print #1, strSQL
            
               While Not RS2.EOF
'                     strSQL = "INSERT INTO SolIte VALUES( "
'                     strSQL = strSQL & RS2!codigo & ", "
'                     strSQL = strSQL & "'" & Trim(RS2!CodigoProducto) & "', "
'                     strSQL = strSQL & "'" & RS2!FchEntrega & "', "
'                     strSQL = strSQL & RS2!cantidad & ", "
'                     strSQL = strSQL & "'" & RS2!TipoProducto & "', "
'                     strSQL = strSQL & "-1, "
'                     strSQL = strSQL & CInt(RS2!flagExtra) & ", "
'                     strSQL = strSQL & "'" & RS2!Fecha2 & "')"
                     strSQL = "INSERT INTO SolIte VALUES( "
                     strSQL = strSQL & RS2!codigo & ", "
                     strSQL = strSQL & "'" & Trim(RS2!CodigoProducto) & "', "
                     strSQL = strSQL & "'" & RS2!FchEntrega & "', "
                     strSQL = strSQL & RS2!cantidad & ", "
                     strSQL = strSQL & "'" & RS2!TipoProducto & "', "
                     strSQL = strSQL & "-1, "
                     'Se corrige tipos de datos en ambiente QAS 05/05/2011 AVL
                     'strSQL = strSQL & CInt(RS2!flagExtra) & ", "
                     'strSQL = strSQL & "'" & RS2!Fecha2 & "')"
                     strSQL = strSQL & "0, "
                     strSQL = strSQL & "'01/01/3000')"
                     Print #1, strSQL
                     RS2.MoveNext
               Wend

           End If
           RS2.Close
           vg_dbpedweb.Execute "pedweb_u_detallesolicitudcompras " & RS1!codigo & ""
           spdData.SetActiveCell 2, lngCurrentReg
           RS1.MoveNext
 '       End If
    Wend  ' While ( Not rsData.EOF )
    Close #1
    RS1.Close
    Set RS1 = Nothing
    Set RS2 = Nothing
    
    fg_carga ""
    strFileNameMDB = ""
    prbStatus.Max = 1
    strLocalFileName = strFileName
    Open strLocalFileName For Input As #1
    Do While Not EOF(1)
       Line Input #1, strLineReg: prbStatus.Max = prbStatus.Max + 1
    Loop
    Close #1
    
'   '-------> Crear directorio DatosTxt
'    If Dir(dir_trabajo & "Datos", vbDirectory) = "" Then MkDir dir_trabajo & "Datos"
'    '-------> Fin crear directorio DatosTxt
    
    lblStatus.Visible = True: prbStatus.Visible = True: prbStatus.Min = 0: lngRow = 0
    estarc = 0
    Open strLocalFileName For Input As #1
    If Not EOF(1) Then
        Line Input #1, strLineReg
        estarc = 1
        Do While Not EOF(1)
            lblStatus.Caption = "Procesando registros, " & Trim(Str(lngRow)) & "/" & Trim(Str(prbStatus.Max))
            DoEvents
            If Mid(strLineReg, 1, 2) = "CC" Then
                If Trim(strFileNameMDB) <> "" Then
                   db7.Close: Set db7 = Nothing
                   If Dir(StrFileNameLdb) <> "" Then
                      Kill (StrFileNameLdb)
                   End If
                End If
'                strFileNameMDB = dir_trabajo & "datos\" & Trim(strLineReg) & ".mdb"
'                strFileNameMDB = fpText1.text & "\" & Trim(strLineReg) & ".mdb"
'                StrFileNameLdb = fpText1.text & "\" & Trim(strLineReg) & ".ldb"
                
'                strFileNameMDB = dir_trabajo & "datos\" & Trim(strLineReg) & ".mdb"
' 05/05/2011 AVL
                'strFileNameMDB = fpText1.text & "\" & Trim(strLineReg) & ".mdb"
                'StrFileNameLdb = fpText1.text & "\" & Trim(strLineReg) & ".ldb"
                strFileNameMDB = fpText1.text & Trim(strLineReg) & ".mdb"
                StrFileNameLdb = fpText1.text & Trim(strLineReg) & ".ldb"
                If Dir(strFileNameMDB) <> "" Then
                   Kill (strFileNameMDB)
                End If
'                Set db7 = DBEngine(0).CreateDatabase(strFileNameMDB, dbLangGeneral, dbLangGeneral + dbVersion20)
'                Set db7 = DBEngine(0).CreateDatabase(strFileNameMDB, dbLangGeneral, dbVersion20)
                Set db7 = DBEngine(0).CreateDatabase(strFileNameMDB, dbLangGeneral, IIf(vg_VAccess = "dbVersion10", 1, IIf(vg_VAccess = "dbVersion20", 16, IIf(vg_VAccess = "dbVersion30", 32, 64))))
            Else
                db7.Execute Trim(strLineReg)
            End If
            strLineReg = ""
            Line Input #1, strLineReg
            lngRow = lngRow + 1
            prbStatus.Value = lngRow
        Loop
        If Trim(strLineReg) <> "" Then db7.Execute Trim(strLineReg)
        If Trim(strFileNameMDB) <> "" Then
           db7.Close: Set db7 = Nothing
           If Dir(StrFileNameLdb) <> "" Then
              Kill (StrFileNameLdb)
           End If
        End If
    End If
    lblStatus.Visible = False: prbStatus.Visible = False
    Close #1
    
    Screen.MousePointer = vbDefault
    DoEvents
    
    'spdData.MaxRows = 0: chkAll.Value = ValueFalse
    'spdData.Visible = True: Check1(3).Visible = True
    Call InicializaForm
    MsgBox "Proceso finalizado OK.", vbInformation, Msgtitulo
    blnExportPedidos = True
    Exit Function
    
blnExportPedidos_Err:
    If estarc = 1 Then Close #1
    Screen.MousePointer = vbDefault
    MsgBox "Se ha detectado el siguiente error: " & Chr(13) & Err.Number & " - " & Err.Description, vbCritical, "blnExportPedidos"
End Function

Private Function blnRptExportToExcel() As Boolean

On Error GoTo blnRptExportToExcel_Err

Dim xlsRpt As Object

    blnRptExportToExcel = False

    Set xlsRpt = CreateObject("Excel.Application")
    
    xlsRpt.Workbooks.Add
    xlsRpt.Workbooks(1).Worksheets(1).Select
    xlsRpt.Workbooks(1).Worksheets(1).Name = "Pedidos"
    spdData.Redraw = False
    spdData.OperationMode = OperationModeNormal
    spdData.SetSelection 2, 0, spdData.MaxCols, spdData.MaxRows
    spdData.ClipboardCopy
    spdData.ClearSelection
    spdData.OperationMode = OperationModeSingle
    spdData.Redraw = True
    
    xlsRpt.Cells.Select
    xlsRpt.Selection.NumberFormat = "@"
    xlsRpt.Range("A1").Select

    xlsRpt.ActiveSheet.Paste
    xlsRpt.Cells.Select
    xlsRpt.Cells.EntireColumn.AutoFit
    xlsRpt.Range("A1").Select
    xlsRpt.Visible = True
    xlsRpt.WindowState = -4140 ' xlMinimized
    
    Clipboard.Clear
    Set xlsRpt = Nothing
    blnRptExportToExcel = True
    Exit Function

blnRptExportToExcel_Err:
    Set xlsRpt = Nothing
    MsgBox "Se ha detectado el siguiente error: " & Chr(13) & Err.Number & " - " & Err.Description, vbCritical, "RptExportToExcel"
End Function

Private Sub GetCentralDeCompras()
Dim strSQL As String
Dim rsData As New ADODB.Recordset
Dim lencen As Integer

    Set rsData = vg_dbpedweb.Execute("sac_s_centralcompras")
    cboCentralCompra.Clear
    
    While (Not rsData.EOF)
        lencen = Len(Trim(rsData!TABCEN_CDCEN))
        cboCentralCompra.AddItem Trim(rsData!TABCEN_CDCEN) & String(3 - lencen, " ") & " - " & rsData!TABCEN_DSCEN
        rsData.MoveNext
    Wend
    
    rsData.Close
    Set rsData = Nothing

End Sub

Private Function GetRegional(idReg As Long) As String
Dim strSQL As String
Dim rsData As New ADODB.Recordset

    Set rsData = vg_dbpedweb.Execute("sac_s_regional 2, '" & idReg & "'")
    If Not rsData.EOF Then
        GetRegional = rsData("TABRGI_DSRGI")
    Else
        GetRegional = "-"
    End If
    
    rsData.Close
    Set rsData = Nothing

End Function

Private Function blnSearchPedidos()

Dim strURL As String
Dim strTipoPed1 As Long
Dim strTipoPed3 As Long
Dim strTipoPed4 As Long
Dim strYearMes As String
Dim strFileName As String
Dim strLocalFileName As String
Dim strLineReg As String
Dim intPos As String

    strTipoPed1 = IIf(chkTipoSolicitud(0).Value = ValueTrue, 1, 0)
'    strTipoPed = strTipoPed & IIf(chkTipoSolicitud(1).Value = ValueTrue And Trim(strTipoPed) <> "", ";", "") & IIf(chkTipoSolicitud(1).Value = ValueTrue, "3", "")
    strTipoPed3 = IIf(chkTipoSolicitud(1).Value = ValueTrue, 3, 0)
'    strTipoPed = strTipoPed & IIf(chkTipoSolicitud(2).Value = ValueTrue And Trim(strTipoPed) <> "", ";", "") & IIf(chkTipoSolicitud(2).Value = ValueTrue, "4", "")
    strTipoPed4 = IIf(chkTipoSolicitud(2).Value = ValueTrue, 4, 0)
    If cboMes.ListIndex < 0 Then MsgBox "Debe indicar mes de los pedidos a bajar.", vbCritical, Msgtitulo: Exit Function
    If Val(txtYear.Value) = 0 Then MsgBox "Debe indicar año de los pedidos a bajar.", vbCritical, Msgtitulo: Exit Function
    If Val(txtSemana.Value) = 0 Then MsgBox "Debe indicar semana de los pedidos a bajar.", vbCritical, Msgtitulo: Exit Function
    If cboCentralCompra.ListIndex < 0 Then MsgBox "Debe seleccionar Central de Compra.", vbCritical, Msgtitulo: Exit Function
    If (strTipoPed1 = 0 And strTipoPed3 = 0 And strTipoPed4 = 0) Then MsgBox "Debe seleccionar al menos un Tipo de Pedido.", vbCritical, Msgtitulo: Exit Function
    If Trim(fpText1.text) = "" Then MsgBox "Debe seleccionar Directorio", vbCritical, Msgtitulo: Exit Function
    If Dir(fpText1.text, vbDirectory) = "" Then MsgBox "No existe directorio seleccionado", vbCritical, Msgtitulo: Exit Function
    Screen.MousePointer = vbHourglass
    DoEvents
    
    strYearMes = fg_pone_cero(Str(txtYear.Value), 4) & fg_pone_cero(Str(cboMes.ListIndex + 1), 2)
    
    flagBaja = False
    
    fg_carga ""
    spdData.MaxRows = 0
    Set RS1 = vg_dbpedweb.Execute("pedweb_s_traerpedidos '" & Trim(Mid(cboCentralCompra.text, 1, 3)) & "', '" & strYearMes & "', " & txtSemana.Value & ", " & strTipoPed1 & ", " & strTipoPed3 & ", " & strTipoPed4 & "")
'jpaz    Open strLocalFileName For Input As #1
    Do While Not RS1.EOF
        DoEvents
        spdData.MaxRows = spdData.MaxRows + 1: spdData.Row = spdData.MaxRows
        spdData.Col = 2: spdData.Value = RS1!NPedido
        spdData.Col = 3: spdData.Value = RS1!tipo
        spdData.Col = 4: spdData.Value = RS1!CentralDeCompra
        spdData.Col = 5: spdData.Value = RS1!centrocosto
        spdData.Col = 6: spdData.Value = RS1!Nombre
        spdData.Col = 7: spdData.Value = GetRegional(RS1!idregional)
        spdData.Col = 8: spdData.Value = RS1!fecahelaboracion
        spdData.Col = 8: spdData.Value = spdData.Value & " " & RS1!horaelaboracion
        spdData.Col = 9: spdData.Value = RS1!descripcion
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    fg_descarga
    
    Screen.MousePointer = vbDefault
    DoEvents
    
    tlbAction.Buttons(4).Enabled = (spdData.MaxRows > 0)
'jpaz    WebBrowser.Visible = False
'jpaz    WebBrowser.Navigate "about:blank"
    DoEvents
    
    If spdData.MaxRows < 1 Then
        MsgBox "No hay pedidos por bajar.", vbInformation, Msgtitulo
    Else
        MsgBox spdData.MaxRows & " pedidos encontrados para exportar.", vbInformation, Msgtitulo
    End If

End Function

Private Sub InicializaForm()

    cboMes.ListIndex = Month(Date) - 1
    txtYear.Value = Year(Date)
    txtSemana.Value = 0
    cboCentralCompra.ListIndex = -1
    chkTipoSolicitud(0).Value = ValueFalse
    chkTipoSolicitud(1).Value = ValueFalse
    chkTipoSolicitud(2).Value = ValueFalse
    spdData.Visible = True: chkAll.Visible = True
    chkAll.Value = ValueFalse
    spdData.MaxRows = 0
'jpaz    WebBrowser.Navigate "about:blank"
'jpaz    WebBrowser.Visible = False
    lblStatus.Visible = False
    prbStatus.Visible = False
    tlbAction.Buttons(1).Enabled = True
    tlbAction.Buttons(2).Value = tbrPressed
    tlbAction.Buttons(4).Enabled = False
    fraParamExp.Enabled = True
    '-------> Traer parametro de ruta
    fpText1.text = ""
    Set RS2 = vg_dbpedweb.Execute("pedweb_s_parametroruta '" & vg_NUsr & "'")
    If Not RS2.EOF Then
       fpText1.text = RS2!par_ruta
    End If
    RS2.Close: Set RS2 = Nothing
    DoEvents

End Sub

Private Sub Double1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub chkAll_Change()
    spdData.Row = -1: spdData.Col = 1: spdData.Value = IIf(chkAll.Value = ValueTrue, 1, 0)
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
    
    'Me.Height = 6630
    'Me.Width = 11100
    Msgtitulo = "WebToMDB"
    fg_centra Me
    '-------> Abrir base sac
    'AbrirBaseSac
    Call InicializaForm
    'V_Clave.Show 1
    Call GetCentralDeCompras
    
End Sub

Private Sub fpText1_ButtonHit(Button As Integer, NewIndex As Integer)
M_ExpDir.Show 1, Me
If vg_dir <> "" Then fpText1.text = vg_dir Else Exit Sub
'-------> Grabar parametro ruta
Set RS2 = vg_dbpedweb.Execute("pedweb_s_parametroruta '" & vg_NUsr & "'")
If Not RS2.EOF Then
   vg_dbpedweb.Execute ("pedweb_iu_parametroruta 'M', '" & vg_NUsr & "', '" & fpText1.text & "'")
Else
   vg_dbpedweb.Execute ("pedweb_iu_parametroruta 'A', '" & vg_NUsr & "', '" & fpText1.text & "'")
End If
RS2.Close: Set RS2 = Nothing
End Sub

Private Sub tlbAction_ButtonClick(ByVal Button As MSComctlLib.Button)

If Trim(vg_VAccess) = "" Then
   MsgBox "No esta definida la versión Access, en archivo gestión.ini", vbCritical, Msgtitulo
   Exit Sub
End If
    Select Case Button.Tag
        Case "SearchData": Call blnSearchPedidos
            
        Case "ExportData":
            Call blnExportPedidos
            Call InicializaForm
        
        Case "Cancel": Call InicializaForm
        
        Case "Close"
'                Set db7 = DBEngine(0).CreateDatabase(strFileNameMDB, dbLangGeneral, IIf(vg_VAccess = "dbVersion10", 1, IIf(vg_VAccess = "dbVersion20", 16, IIf(vg_VAccess = "dbVersion30", 32, 64))))
            Unload Me
'            End
    
    End Select

End Sub
