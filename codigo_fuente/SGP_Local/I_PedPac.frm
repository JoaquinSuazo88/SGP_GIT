VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form I_PedPac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe"
   ClientHeight    =   5415
   ClientLeft      =   195
   ClientTop       =   1230
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   840
      TabIndex        =   20
      Top             =   480
      Width           =   7695
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   1515
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
         Index           =   7
         Left            =   240
         TabIndex        =   23
         Top             =   285
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2760
         Picture         =   "I_PedPac.frx":0000
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
         Left            =   3195
         TabIndex        =   22
         Top             =   210
         Width           =   4215
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3240
         TabIndex        =   24
         Top             =   255
         Width           =   4215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabMaxWidth     =   4
      TabCaption(0)   =   "Parámetros"
      TabPicture(0)   =   "I_PedPac.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraControls(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Parámetros"
      TabPicture(1)   =   "I_PedPac.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraControls(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Parámetros"
      TabPicture(2)   =   "I_PedPac.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraControls(2)"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraControls 
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
         Height          =   3375
         Index           =   2
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   7440
      End
      Begin VB.Frame fraControls 
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
         Height          =   3375
         Index           =   1
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   9600
         Begin VB.ListBox lstAporte 
            Height          =   1410
            Index           =   0
            Left            =   7275
            Style           =   1  'Checkbox
            TabIndex        =   38
            Top             =   1290
            Width           =   2205
         End
         Begin VB.ComboBox cboMiscRpt 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            ItemData        =   "I_PedPac.frx":035E
            Left            =   240
            List            =   "I_PedPac.frx":0360
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   510
            Width           =   2745
         End
         Begin VB.ListBox lstDepto 
            Height          =   1410
            Index           =   1
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   10
            Top             =   1290
            Width           =   2265
         End
         Begin VB.ListBox lstServicio 
            Height          =   1410
            Index           =   1
            Left            =   4870
            Style           =   1  'Checkbox
            TabIndex        =   9
            Top             =   1290
            Width           =   2205
         End
         Begin VB.ListBox lstRegimen 
            Height          =   1410
            Index           =   1
            Left            =   2530
            Style           =   1  'Checkbox
            TabIndex        =   8
            Top             =   1290
            Width           =   2190
         End
         Begin EditLib.fpDateTime Date1 
            Height          =   315
            Index           =   2
            Left            =   3360
            TabIndex        =   25
            Top             =   510
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
         Begin EditLib.fpDateTime Date1 
            Height          =   315
            Index           =   3
            Left            =   4890
            TabIndex        =   26
            Top             =   510
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   1
            Left            =   1875
            TabIndex        =   32
            Top             =   2970
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
            MaxLength       =   20
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
            Caption         =   "Aportes"
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
            Index           =   12
            Left            =   7275
            TabIndex        =   39
            Top             =   960
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Formato Informe"
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
            Index           =   11
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   3555
            TabIndex        =   34
            Top             =   2970
            Width           =   4335
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   3120
            Picture         =   "I_PedPac.frx":0362
            Top             =   2880
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Paciente"
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
            Left            =   960
            TabIndex        =   33
            Top             =   3045
            Width           =   765
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
            Index           =   10
            Left            =   4870
            TabIndex        =   31
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Left            =   3360
            TabIndex        =   30
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Grupo Usuario"
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
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Régimen"
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
            Left            =   2530
            TabIndex        =   28
            Top             =   960
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Left            =   4890
            TabIndex        =   27
            Top             =   240
            Width           =   510
         End
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   3600
            TabIndex        =   35
            Top             =   3015
            Width           =   4335
         End
      End
      Begin VB.Frame fraControls 
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
         Height          =   3375
         Index           =   0
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   8640
         Begin VB.ComboBox cboMiscRpt 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            ItemData        =   "I_PedPac.frx":066C
            Left            =   120
            List            =   "I_PedPac.frx":066E
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   555
            Width           =   2745
         End
         Begin VB.ListBox lstDepto 
            Height          =   1860
            Index           =   0
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   1290
            Width           =   2745
         End
         Begin VB.ListBox lstServicio 
            Height          =   1860
            Index           =   0
            Left            =   5895
            Style           =   1  'Checkbox
            TabIndex        =   4
            Top             =   1290
            Width           =   2565
         End
         Begin VB.ListBox lstRegimen 
            Height          =   1860
            Index           =   0
            Left            =   3045
            Style           =   1  'Checkbox
            TabIndex        =   3
            Top             =   1290
            Width           =   2670
         End
         Begin EditLib.fpDateTime Date1 
            Height          =   315
            Index           =   0
            Left            =   3045
            TabIndex        =   13
            Top             =   555
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
         Begin EditLib.fpDateTime Date1 
            Height          =   315
            Index           =   1
            Left            =   4575
            TabIndex        =   14
            Top             =   555
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
            Index           =   5
            Left            =   5895
            TabIndex        =   19
            Top             =   1005
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Régimen"
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
            Left            =   3045
            TabIndex        =   18
            Top             =   1005
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Grupo Usuario"
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
            TabIndex        =   17
            Top             =   1005
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Left            =   4575
            TabIndex        =   16
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Left            =   3045
            TabIndex        =   15
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Formato Informe"
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
            TabIndex        =   12
            Top             =   285
            Width           =   1380
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_PedPac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim MsgTitulo As String, est As Boolean
Public lc_Aux As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 5895
Me.Width = 10185
est = True
fg_centra Me
If lc_Aux = "Produc" Then
   Me.Caption = "Informe de Producción"
   MsgTitulo = "Informe de Producción"
   SSTab1.Tab = 0
   SSTab1.TabVisible(1) = False
   SSTab1.TabVisible(2) = False
ElseIf lc_Aux = "DetCon" Then
   Me.Caption = "Informe Detalle de Consumos"
   MsgTitulo = "Informe Detalle de Consumos"
   SSTab1.Tab = 1
   SSTab1.TabVisible(0) = False
   SSTab1.TabVisible(2) = False
   cboMiscRpt(1).Enabled = False
   lstAporte(0).Enabled = False
ElseIf lc_Aux = "ANutPa" Then
   Me.Caption = "Informe Aporte Nutricional"
   MsgTitulo = "Informe Aporte Nutriconal"
   SSTab1.Tab = 1
   SSTab1.TabVisible(0) = False
   SSTab1.TabVisible(2) = False
End If
EspFecha Date1(0)
EspFecha Date1(1)
'Dim btnX As Button
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
Date1(0).text = Format(Date, "dd/mm/yyyy")
Date1(1).text = Format(Date, "dd/mm/yyyy")
Date1(2).text = Format(Date, "dd/mm/yyyy")
Date1(3).text = Format(Date, "dd/mm/yyyy")
With cboMiscRpt(0)
    .Clear
    .AddItem "Detalle" & Space(150) & "(0)"
    .AddItem "Resumen" & Space(150) & "(1)"
    .ListIndex = 0
End With
With cboMiscRpt(1)
    .Clear
    .AddItem "Detalle" & Space(150) & "(0)"
    .AddItem "Resumen" & Space(150) & "(1)"
    .ListIndex = 0
End With
fpText(0).Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText(0).text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)

'------- Llenar tabla grupo paciente
With lstDepto(i)
    RS.Open RutinaLectura.GrupoPaciente(1), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No existe grupo usuario", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    For i = 0 To 1
        .AddItem "[Todos]" & Space(150), 0
        .ListIndex = 0
        .Selected(0) = True
        RS.MoveFirst
        Do While Not RS.EOF
           .AddItem Trim(RS!grp_nombre) & Space(150) & ";" & RS!grp_codigo
           RS.MoveNext
        Loop
    Next i
    RS.Close: Set RS = Nothing
End With
'------- Llenar tabla regimen
With lstRegimen(i)
    RS.Open RutinaLectura.Regimen(1, 0, ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No existe régimen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    For i = 0 To 1
        .AddItem "[Todos]" & Space(150), 0
        .ListIndex = 0
        .Selected(0) = True
        RS.MoveFirst
        Do While Not RS.EOF
           .AddItem Trim(RS!reg_nombre) & Space(150) & ";" & RS!reg_codigo
           RS.MoveNext
        Loop
    Next i
    RS.Close: Set RS = Nothing
End With
'------- Llenar tabla servicio
With lstServicio(i)
    RS.Open RutinaLectura.Servicio(1, 0, ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No existe servicio", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    For i = 0 To 1
        .AddItem "[Todos]" & Space(150), 0
        .ListIndex = 0
        .Selected(0) = True
        RS.MoveFirst
        Do While Not RS.EOF
           .AddItem Trim(RS!ser_nombre) & Space(150) & ";" & RS!ser_codigo
           RS.MoveNext
        Loop
    Next i
    RS.Close: Set RS = Nothing
End With
'------- Llenar tabla nutrientes
With lstAporte(i)
    RS.Open RutinaLectura.Nutriente(1, 0, ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No existe nutriente", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    For i = 0 To 0
        .AddItem "[Todos]" & Space(150), 0
        .ListIndex = 0
    '    .Selected(0) = True
        RS.MoveFirst
        j = 1
        Do While Not RS.EOF
           .AddItem Trim(RS!nut_nombre) & Space(150) & ";" & RS!nut_codigo
           If RS!nut_indpri = 1 Then .Selected(j) = True
           RS.MoveNext: j = j + 1
        Loop
        .ListIndex = 0
    Next i
    RS.Close: Set RS = Nothing
End With
est = False
End Sub

Private Sub fpText_Change(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText(0).text)), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!cli_nombre)
    RS.Close: Set RS = Nothing
Case 1
    fpayuda(1).Caption = ""
End Select
End Sub

Private Sub fpText_GotFocus(Index As Integer)
Select Case Index
Case 1
    If Trim(fpText(Index).text) = "" Or vg_Dig = "N" Then Exit Sub
    fpText(Index).text = fg_DespintaRut(fpText(Index).text)
    fpText(Index).text = Mid(fpText(Index).text, 1, Len(Trim(fpText(Index).text)) - 1)
End Select
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_LostFocus(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 1
    If fpText(Index).text = "" Or est Then Exit Sub
    fpText(Index).text = fg_RutDig(Trim(fpText(Index).text))
    RS.Open RutinaLectura.Paciente(1, Trim(fpText(Index).text)), vg_db, adOpenStatic
    codreg = 0
    If Not RS.EOF Then
       fpText(Index).text = fg_PintaRut(fpText(Index).text)
       fpayuda(1).Caption = Trim(RS!pac_nombre) & " " & Trim(RS!pac_appaterno) & " " & Trim(RS!pac_apmaterno)
    Else
        RS.Close: Set RS = Nothing: MsgBox "Pacientes no existe...", vbCritical, MsgTitulo
        fpText(Index).text = "": fpayuda(Index).Caption = ""
        Exit Sub
    End If
    RS.Close: Set RS = Nothing
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
    fpText(0).text = vg_codigo
    fpayuda(0).Caption = vg_nombre
Case 1
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_pacientes", "pac_", "Pacientes", "Pac"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(Index).text = fg_PintaRut(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
End Select
End Sub

Private Sub lstDepto_Click(Index As Integer)
    Call SetSelectionParamList(lstDepto(Index))
End Sub

Private Sub lstRegimen_Click(Index As Integer)
    Call SetSelectionParamList(lstRegimen(Index))
End Sub

Private Sub lstServicio_Click(Index As Integer)
    Call SetSelectionParamList(lstServicio(Index))
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codreg As String, codser As String, codgrp As String, codapo As String
Select Case Button.Index
Case 1
    If Trim(fpayuda(0).Caption) = "" Then MsgBox "No existe contrato", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If (Date1(SSTab1.Tab).Value > Date1(SSTab1.Tab).Value) Then
        MsgBox "Fecha desde debe ser menor o igual a Fecha hasta.", vbCritical, App.Title
        Exit Sub
    End If
    If (Not blnValidParam(codgrp, lstDepto(SSTab1.Tab), "Debe seleccionar al menos un Departamento.")) Then Exit Sub
    If (Not blnValidParam(codser, lstServicio(SSTab1.Tab), "Debe seleccionar al menos un Servicio.")) Then Exit Sub
    If (Not blnValidParam(codreg, lstRegimen(SSTab1.Tab), "Debe seleccionar al menos un Régimen.")) Then Exit Sub
    If (Not blnValidParam(codapo, lstAporte(0), "Debe seleccionar al menos un Aporte Nutricional.")) Then Exit Sub
    If lc_Aux = "Produc" Then
        I_ProduccionPaciente fpText(0).text, codreg, codser, codgrp, Date1(0).text, Date1(1).text, IIf(Val(fg_codigocbo(cboMiscRpt, 0, 1, "")) = 0, True, False)
    ElseIf lc_Aux = "DetCon" Then
        I_DetalleConsumoPaciente fpText(0).text, fg_DespintaRut(fpText(1).text), codreg, codser, codgrp, Date1(2).text, Date1(3).text
    ElseIf lc_Aux = "ANutPa" Then
        I_AporteNutPaciente fpText(0).text, Trim(fpayuda(0).Caption), fg_DespintaRut(fpText(1).text), Trim(fpayuda(1).Caption), codreg, codser, codgrp, codapo, Date1(2).text, Date1(3).text
    End If
Case 3
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub SetSelectionParamList(ByRef lstControl As ListBox)
Dim intIndex As Integer
Dim strItems As String

    If (lstControl.ListIndex < 0) Or (blnSetDataCtrl) Then Exit Sub
    
    DoEvents
    blnSetDataCtrl = True
    'lstControl.Selected(lstControl.ListIndex) = (Not lstControl.Selected(lstControl.ListIndex))
    
    If (lstControl.ListIndex = 0) Then ' TODOS
        For intIndex = 1 To lstControl.listcount - 1
            lstControl.Selected(intIndex) = False
        Next
    Else
        lstControl.Selected(0) = False
        ' Check ALL selected
        If (blnSelectionList(strItems, lstControl)) Then
            If ((CountItems(strItems) / 2) = (lstControl.listcount - 1)) Then
                blnSetDataCtrl = False
                lstControl.Selected(0) = True
            End If
        End If
    End If
    
    blnSetDataCtrl = False
    DoEvents

End Sub

Private Function blnSelectionList(ByRef strItem As String, lstControl As ListBox, Optional intIndexAll = 0) As Boolean
Dim intIndex As Integer
    blnSelectionList = False
    strItem = ""
    For intIndex = 0 To lstControl.listcount - 1
        If (lstControl.Selected(intIndex)) Then
            strItem = strItem & GetItem(lstControl.List(intIndex), 2) & ";" & GetItem(lstControl.List(intIndex), 1) & ";"
        End If
    Next
    blnSelectionList = (strItem <> "")
End Function

Function blnValidParam(ByRef strItem As String, lstControl As ListBox, strMsg As String) As Boolean

    blnValidParam = True

    If (Not blnSelectionList(strItem, lstControl)) Then
        MsgBox strMsg, vbCritical, MsgTitulo
        blnValidParam = False
    End If

End Function
