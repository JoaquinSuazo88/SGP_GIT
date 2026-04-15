VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_Copia_Minuta_Lideres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualiza - Copia Minuta Lideres a Seguidores"
   ClientHeight    =   10020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19785
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   19785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Datos Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   105
      TabIndex        =   20
      Top             =   3195
      Width           =   14550
      Begin VB.Frame Frame5 
         Height          =   435
         Index           =   1
         Left            =   5280
         TabIndex        =   37
         Top             =   5400
         Width           =   900
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   38
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.CheckBox ChComensales 
         Caption         =   "Comensales Totales"
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
         Left            =   7440
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox ChPonderaciones 
         Caption         =   "Ponderaciones"
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
         Left            =   5280
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox ChReceta 
         Caption         =   "Receta"
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
         Left            =   3600
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Left            =   2640
         TabIndex        =   25
         Top             =   5400
         Width           =   2550
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   26
            Top             =   135
            Width           =   4845
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Index           =   0
         Left            =   1560
         TabIndex        =   23
         Top             =   5400
         Width           =   900
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   24
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Actualiza - Copia Minuta Lideres"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   10695
         TabIndex        =   11
         Top             =   5970
         Width           =   1635
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   12525
         TabIndex        =   12
         Top             =   5970
         Width           =   1275
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   4320
         Left            =   105
         TabIndex        =   10
         Top             =   870
         Width           =   14295
         _Version        =   393216
         _ExtentX        =   25215
         _ExtentY        =   7620
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
         SpreadDesigner  =   "M_Copia_Minuta_Lideres.frx":0000
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   210
         Left            =   60
         TabIndex        =   21
         Top             =   6180
         Visible         =   0   'False
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   1
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Actualiza Minuta"
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
         Left            =   10200
         TabIndex        =   36
         Top             =   5520
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   9600
         Top             =   5550
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Copia Minuta"
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
         Left            =   12720
         TabIndex        =   35
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   12120
         Top             =   5550
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Org. Compras"
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
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label lbl_proceso 
         Alignment       =   2  'Center
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   11340
         TabIndex        =   22
         Top             =   3780
         Visible         =   0   'False
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Origenes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   1935
      TabIndex        =   13
      Top             =   105
      Width           =   10905
      Begin VB.Frame Frame2 
         Height          =   2490
         Left            =   150
         TabIndex        =   14
         Top             =   210
         Width           =   10575
         Begin EditLib.fpDateTime FpFecDesde 
            Height          =   315
            Left            =   1395
            TabIndex        =   3
            Top             =   1425
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
            Text            =   "01/09/2013"
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
         Begin EditLib.fpDateTime FpFecHasta 
            Height          =   315
            Left            =   8580
            TabIndex        =   4
            Top             =   1425
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
            Text            =   "28/09/2013"
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
            Left            =   1410
            TabIndex        =   0
            Top             =   285
            Width           =   1275
            _Version        =   196608
            _ExtentX        =   2249
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
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   360
            Left            =   10080
            TabIndex        =   5
            Top             =   1350
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   0
            Left            =   1395
            TabIndex        =   1
            Top             =   675
            Width           =   945
            _Version        =   196608
            _ExtentX        =   1676
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
            Index           =   1
            Left            =   1395
            TabIndex        =   2
            Top             =   1035
            Width           =   945
            _Version        =   196608
            _ExtentX        =   1676
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
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   2685
            Picture         =   "M_Copia_Minuta_Lideres.frx":1A52
            Top             =   960
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   3135
            TabIndex        =   30
            Top             =   1035
            Width           =   6735
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
            Left            =   120
            TabIndex        =   29
            Top             =   1140
            Width           =   705
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   2685
            Picture         =   "M_Copia_Minuta_Lideres.frx":1D5C
            Top             =   600
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   3135
            TabIndex        =   28
            Top             =   675
            Width           =   6735
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
            Left            =   120
            TabIndex        =   27
            Top             =   780
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta"
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
            Left            =   7275
            TabIndex        =   18
            Top             =   1515
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde"
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
            TabIndex        =   17
            Top             =   1515
            Width           =   1140
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   2700
            Picture         =   "M_Copia_Minuta_Lideres.frx":2066
            Top             =   210
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3150
            TabIndex        =   16
            Top             =   285
            Width           =   6735
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
            Left            =   135
            TabIndex        =   15
            Top             =   390
            Width           =   735
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
            Index           =   0
            Left            =   3195
            TabIndex        =   19
            Top             =   330
            Width           =   6735
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
            Left            =   3195
            TabIndex        =   33
            Top             =   720
            Width           =   6735
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
            Left            =   3195
            TabIndex        =   34
            Top             =   1080
            Width           =   6735
         End
      End
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
            Picture         =   "M_Copia_Minuta_Lideres.frx":2370
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   9615
      Left            =   14880
      TabIndex        =   32
      Top             =   240
      Width           =   4695
      _Version        =   393216
      _ExtentX        =   8281
      _ExtentY        =   16960
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
      MaxCols         =   4
      MaxRows         =   1
      SpreadDesigner  =   "M_Copia_Minuta_Lideres.frx":270A
   End
End
Attribute VB_Name = "M_Copia_Minuta_Lideres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lc_Aux As String
Dim MsgTitulo As String

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim RS           As New ADODB.Recordset
Dim seleccion    As String
Dim seleccionx   As String
Dim Fecha        As Date
Dim i            As Long
Dim j            As Long
Dim X            As Long
Dim CecoOrigen   As String
Dim RegOrigen    As Long
Dim SerOrigen    As Long
Dim CecoDestino  As String
Dim RegDestino   As Long
Dim SerDestino   As Long
Dim FecIni       As String
Dim FecFin       As String
Dim Bloque       As String
Dim Conta        As Long
Dim Sql          As String
Dim EstCopiado   As Boolean
Dim LargoDia     As Long
Dim FechaDesFin  As Date
Dim Id_Bloque    As Long
Dim MyBuffer     As String
Dim codest       As Long
Dim NumLin       As Long
Dim OpReceta     As String
Dim OpPorcentaje As String
Dim OpComensales As String

  If Not ValidarDatosEntrada Then
  
     Exit Sub
     
  End If
  
  '-------> Validar que exista un dato seleccionado encabezado
  seleccion = 0
  For i = 1 To vaSpread1.MaxRows
       
       vaSpread1.Row = i
       vaSpread1.Col = 1 'Seleccion
       seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
       If seleccion = 1 And vaSpread1.RowHidden = False Then
          
          Exit For
       
       End If
  
  Next i
  
  If seleccion = 0 Then
     
     MsgBox "Debe seleccionar a lo menos una estructura de servicio...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If
  
  If seleccion = 0 And (ChReceta.Value = 1 Or ChPonderaciones.Value = 1) Then
     
     MsgBox "Debe seleccionar a lo menos una estructura de servicio...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If

  Let MyBuffer = ""
  Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
  Let MyBuffer = MyBuffer & "<GrabaEst>"
            
  For i = 1 To vaSpread1.MaxRows
       
       vaSpread1.Row = i
       vaSpread1.Col = 1 'Seleccion
       seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
       If seleccion = 1 And vaSpread1.RowHidden = False Then
          
          MyBuffer = MyBuffer & " <Est"
          
          vaSpread1.Col = 3
          NumLin = vaSpread1.text
          
          vaSpread1.Col = 4
          codest = vaSpread1.text
          
          MyBuffer = MyBuffer & " Est = " & Chr(34) & codest & Chr(34)
          MyBuffer = MyBuffer & " Lin = " & Chr(34) & NumLin & Chr(34)
          MyBuffer = MyBuffer & "/>"
          
       End If
  
  Next i
  
  MyBuffer = MyBuffer & "</GrabaEst>"
  
  OpReceta = "0"
  OpPorcentaje = "0"
  OpComensales = "0"
  
  OpReceta = IIf(ChReceta.Value = 1, "1", "0")
  OpPorcentaje = IIf(ChPonderaciones.Value = 1, "1", "0")
  OpComensales = IIf(ChComensales.Value = 1, "1", "0")

  seleccion = 0
  Conta = 0
  For i = 1 To vaSpread2.MaxRows
       
       vaSpread2.Row = i
       vaSpread2.Col = 1 'Seleccion
       seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
       If seleccion = 1 And vaSpread2.RowHidden = False Then
          
          If vaSpread2.BackColor <> &HC0FFFF And OpReceta = "0" And OpPorcentaje = "0" And OpComensales = "0" Then
                    
             MsgBox "Debe seleccionar a lo menos unas de las opciones: Recetas - Porcentaje - Comensales Totales...", vbExclamation + vbOKOnly, MsgTitulo
             Exit Sub
                    
          End If
          
          Conta = Conta + 1
       
       End If
       
       vaSpread2.Col = 9
       vaSpread2.text = ""
  
  Next i

  If Conta = 0 Then
     
     MsgBox "Debe seleccionar a lo menos un ceco destino...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If
  
  ProgressBar1.Scrolling = ccScrollingSmooth
  ProgressBar1.Max = 100
  ProgressBar1.Visible = True
  ProgressBar1.Value = 0
  lbl_proceso.Caption = "0 %"
  lbl_proceso.Visible = True
  
  Toolbar2.Enabled = False
  FpFecDesde.Enabled = False
  FpFecHasta.Enabled = False
  fpText.Enabled = False
  fpText1.Enabled = False

  fg_carga ""
  EstCopiado = True
  
  CecoOrigen = LimpiaDato(fpText.text)
  RegOrigen = fpLongInteger1(0).text
  SerOrigen = fpLongInteger1(1).text
  FecIni = Format(FpFecDesde.text, "yyyymmdd")
  FecFin = Format(FpFecHasta.text, "yyyymmdd")
   
  j = 1
  For i = 1 To vaSpread2.MaxRows
  
      vaSpread2.Row = i
      vaSpread2.Col = 1 'Seleccion
      seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
      If seleccion = 1 And vaSpread2.RowHidden = False Then
          
         CecoDestino = ""
         vaSpread2.Col = 2
         CecoDestino = vaSpread2.text
          
         RegDestino = 0
         vaSpread2.Col = 4
         RegDestino = vaSpread2.text
          
         SerDestino = 0
         vaSpread2.Col = 8
         SerDestino = vaSpread2.text

         If RS.State = 1 Then RS.Close
         RS.CursorLocation = adUseClient
         vg_db.CursorLocation = adUseClient

         vaSpread2.Col = 1
         If vaSpread2.BackColor = &HC0FFFF Then
         
            Set RS = vg_db.Execute("sgpadm_Ins_XmlCopiaMinutaBloqueLiderSeguidor_V01 '" & MyBuffer & "', '" & CecoOrigen & "', " & RegOrigen & ", " & SerOrigen & ", " & FecIni & ", " & FecFin & ", '" & CecoDestino & "', " & RegDestino & ", " & SerDestino & "")
         
         Else
        
            'Actualizar minuta % Ponderaciones o raciones
            Set RS = vg_db.Execute("sgpadm_Upd_XmlMinutaBloqueLiderSeguidor_V01 '" & MyBuffer & "', '" & CecoOrigen & "', " & RegOrigen & ", " & SerOrigen & ", " & FecIni & ", " & FecFin & ", '" & CecoDestino & "', " & RegDestino & ", " & SerDestino & ", '" & OpReceta & "', '" & OpPorcentaje & "', '" & OpComensales & "', '" & IIf(ChPonderaciones.Caption = "Ponderaciones", "1", "2") & "'")
         
            
         End If
         
         If Not RS.EOF Then
                    
            If RS(0) > 0 Then
                       
               vaSpread2.Col = 7
               vaSpread2.text = RS(0) & " " & RS(1)
              
            Else
                       
               vaSpread2.Col = 7
               vaSpread2.text = "Proceso finalizado sin problema"
              
               vaSpread2.Row = i
               vaSpread2.Col = -1
               vaSpread2.BackColor = &HFFFFFF
              
            End If
                 
         End If
         RS.Close
         Set RS = Nothing

         
         ProgressBar1.Value = ((j / Conta) * 100)
         lbl_proceso.Caption = CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %"
         j = j + 1
         
      End If
  
      DoEvents
       
  Next i

  fg_descarga
  MsgBox "Proceso Finalizado", vbInformation, Me.Caption
                
  ProgressBar1.Visible = False
  lbl_proceso.Visible = False
  
  Toolbar2.Enabled = True
  FpFecDesde.Enabled = True
  FpFecHasta.Enabled = True
  fpText.Enabled = True
  fpText1.Enabled = True
                
Exit Sub
Man_Error:
    
    fg_descarga
    
    ProgressBar1.Visible = False
    lbl_proceso.Visible = False
    
    Toolbar2.Enabled = True
    FpFecDesde.Enabled = True
    FpFecHasta.Enabled = True
    fpText.Enabled = True
    fpText1.Enabled = True
    
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub Command2_Click()

On Error GoTo Man_Error

'-------> Salir de la opción
Unload Me

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
fg_centra Me
MsgTitulo = "Actualizar - Copiar Minuta Lideres a Seguidores"

FpFecHasta.text = Format(Date, "dd/mm/yyyy")
FpFecDesde.text = Format(Date, "dd/mm/yyyy")
vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0

vaSpread1.MaxRows = 0

vaSpread2.MaxRows = 0

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0

If IsDate(FpFecDesde.text) = False Then Exit Sub

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FPFecDestino_Change()

On Error GoTo Man_Error

'vaSpread2.MaxRows = 0
If IsDate(FPFecDestino.text) = False Then Exit Sub

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FPFecDestino_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0

If IsDate(FpFecHasta.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

Dim sdql_mvi As String
Dim RS       As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Select Case Index
    
    Case 0
        
        Set RS = vg_db.Execute("sgpadm_Sel_RegimenBloque " & IIf(Val(fpLongInteger1(0).Value) = 0, -1, Val(fpLongInteger1(0).Value)) & "")
        If RS.EOF Then
            
            RS.Close
            Set RS = Nothing
            fpayuda(1).Caption = ""
            Exit Sub
        
        End If
        fpayuda(1).Caption = Trim(RS!reg_nombre)
        RS.Close: Set RS = Nothing
        vaSpread1.MaxRows = 0
        vaSpread2.MaxRows = 0

    Case 1
        
        Set RS = vg_db.Execute("sgpadm_Sel_ServicioBloque " & IIf(Val(fpLongInteger1(1).Value) = 0, -1, Val(fpLongInteger1(1).Value)) & "")
        
        If RS.EOF Then
            
            RS.Close
            Set RS = Nothing
            fpayuda(2).Caption = ""
            
            Exit Sub
        
        End If
        fpayuda(2).Caption = Trim(RS!ser_nombre)
        RS.Close
        Set RS = Nothing
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
   
        Set RS = vg_db.Execute("sgpadm_Sel_ServicioPonderacionRaciones " & Val(fpLongInteger1(1).Value) & "")
        If Not RS.EOF Then
       
           ChPonderaciones.Caption = "Raciones"
            
        Else
            
           ChPonderaciones.Caption = "Ponderaciones"
    
        End If
        RS.Close
        Set RS = Nothing

        vaSpread1.MaxRows = 0
        vaSpread2.MaxRows = 0

End Select
    
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpText_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS = vg_db.Execute("sgpadm_s_cliente_V02 45, '" & LimpiaDato(fpText.text) & "', ''")
    If RS.EOF Then
        
        RS.Close
        Set RS = Nothing
        fpayuda(0).Caption = ""
        vaSpread1.MaxRows = 0
        vaSpread2.MaxRows = 0
        Exit Sub
    
    End If
    fpayuda(0).Caption = Trim(IIf(IsNull(RS!cli_nombre), "", RS!cli_nombre))
    fpText.text = RS!cli_codigo
    RS.Close
    Set RS = Nothing
    vaSpread1.MaxRows = 0
    vaSpread2.MaxRows = 0
 
Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpText1_Change()

On Error GoTo Man_Error

vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpText1_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

    Select Case Index
        
        Case 0
            
            vg_left = fpayuda(0).Left + 2300
            vg_nombre = "": vg_codigo = ""
            Call B_TabEst.LlenaDatos("b_clientes", "cli_", "Clientes", "Cliente_SitioRemoto")
            Call B_TabEst.Show(1)
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpText.text = vg_codigo
            fpayuda(0).Caption = vg_nombre
            fpLongInteger1(0).SetFocus
        
        Case 1
            
            vg_left = fpayuda(1).Left + 2300
            vg_nombre = "": vg_codigo = ""
            Call B_TabEst.LlenaDatos("a_regimen", "", "Regimen", "RegBlo")
            Call B_TabEst.Show(1)
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(0).Value = Val(vg_codigo)
            fpLongInteger1(0).SetFocus
            fpayuda(1).Caption = vg_nombre
            fpLongInteger1(1).SetFocus
        
        Case 2
            
            Let vg_left = fpayuda(2).Left + 2300
            Let vg_nombre = ""
            Let vg_codigo = ""
            Call B_TabEst.LlenaDatos("a_servicio", "", "Servicio", "SerBlo")
            Call B_TabEst.Show(1)
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(1).Value = Val(vg_codigo)
            fpLongInteger1(1).SetFocus
            fpayuda(2).Caption = vg_nombre
            FpFecDesde.Enabled = True
            FpFecHasta.Enabled = True
            
            Call FpFecDesde.SetFocus
    
    End Select

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub LlenarGrillaCeco()

On Error GoTo Man_Error

Dim RS        As New ADODB.Recordset
Dim OrgCompra As String
Dim Sql       As String

vaSpread2.Row = -1: vaSpread2.Col = -1
vaSpread2.BackColor = &HC0FFFF
vaSpread2.MaxRows = 0

' Control displays text tips aligned to pointer with focus
vaSpread2.TextTip = 2
vaSpread2.TextTipDelay = 250
X = vaSpread2.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)

OrgCompra = LimpiaDato(Trim(fpText1.text))

Sql = ""
Sql = " sgpadm_Sel_OrgComprasxCeco "
Sql = Sql & " '" & OrgCompra & "'"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

Do While Not RS.EOF
   
   vaSpread2.MaxRows = vaSpread2.MaxRows + 1
   vaSpread2.Row = vaSpread2.MaxRows
   
   vaSpread2.Col = 1
   vaSpread2.text = "0"
   
   vaSpread2.Col = 2
   vaSpread2.text = RS!cli_codigo
   
   vaSpread2.Col = 3
   vaSpread2.text = Trim(IIf(IsNull(RS!cli_nombre), "", RS!cli_nombre))
   
   vaSpread2.Col = 5
   vaSpread2.text = 0
   
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub TextDet2_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(TextDet2(Index).text, ",")

If Index = 2 Then
   
   TextDet2(3).text = ""

ElseIf Index = 3 Then
   
   TextDet2(2).text = ""

ElseIf Index = 4 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""

End If

For i = 1 To vaSpread2.MaxRows
           
    vaSpread2.Row = i
    vaSpread2.Col = 9
    vaSpread2.text = 0
    
Next

Select Case Index

Case 2, 3, 4
    
    vaSpread2.Visible = False
    
    If Trim(TextDet2(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread2.MaxRows
           
           vaSpread2.Row = i
           vaSpread2.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like IIf(Index = 2 Or Index = 4, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
           'UCase(Trim(vaSpread2.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread2.Col = 1
           
           If indactivo = -1 And Trim(vaSpread2.text) <> "" Then
              
              vaSpread2.Col = 9
              
              If Val(vaSpread2.Value) <> 1 Then
                              
                 vaSpread2.Col = 1
              
                 If vaSpread2.RowHidden = True Then
                 
                    vaSpread2.RowHidden = False
                    vaSpread2.Col = 9
                    vaSpread2.text = 1
                 
                 Else
                 
                    vaSpread2.Col = 9
                    vaSpread2.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread2.Col = 9
              EstBuq = vaSpread2.Value
              vaSpread2.Col = 1
              
              If vaSpread2.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread2.RowHidden = True
                 
                 vaSpread2.Col = 9
                 vaSpread2.text = 0
                 
              End If
           
           End If
        
        Next i
        
        vaSpread2.SetActiveCell Index + 1, 1
        vaSpread2.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread2.ColUserSortIndicator(IIf(Trim(wvarArr8020(X)) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread2.SortKey(1) = IIf(Trim(wvarArr8020(X)) = "", 0, 0): vaSpread2.SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread2.Sort -1, -1, vaSpread2.maxcols, vaSpread2.MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(TextDet2(Index).text) = "" Then
       
       For i = 1 To vaSpread2.MaxRows
           
           vaSpread2.Row = i
           If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
           
           vaSpread2.Col = 9
           vaSpread2.text = 0
       
       Next
       
       vaSpread2.SetActiveCell Index, vaSpread2.SearchCol(Index, 0, vaSpread2.MaxRows, Trim(TextDet2(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread2.SetActiveCell Index, 1
    
    End If
    
    vaSpread2.Visible = True

End Select

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS        As New ADODB.Recordset
Dim Sql       As String
Dim i         As Long
Dim xmlceco   As String
Dim seleccion As String
Dim codCeco   As String

Select Case Button.Index

Case 1

  If Not ValidarDatosEntrada Then
  
     Exit Sub
     
  End If
   
  vaSpread1.MaxRows = 0
  vaSpread2.MaxRows = 0
  ' vaSpread1.Row = -1: vaSpread1.Col = -1
  ' vaSpread1.BackColor = &HC0FFFF
   
  Sql = ""
  Sql = Sql & "'" & Trim(fpText.text) & "', "
  Sql = Sql & "" & fpLongInteger1(0).text & ", "
  Sql = Sql & "" & fpLongInteger1(1).text & ", "
  Sql = Sql & "" & Format(FpFecDesde, ("YYYYMMDD")) & ", "
  Sql = Sql & "" & Format(FpFecHasta, ("YYYYMMDD")) & " "
   
  fg_carga ""
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
  Set RS = vg_db.Execute("sgpadm_Sel_TraerEstructuraOrigenLideres_V01 " & Sql & "")
  If Not RS.EOF Then
   
        Do While Not RS.EOF
           
           vaSpread1.MaxRows = vaSpread1.MaxRows + 1
           vaSpread1.Row = vaSpread1.MaxRows
           
           vaSpread1.Col = 1
           vaSpread1.text = "0"
           
           vaSpread1.Col = 2
           vaSpread1.text = RS!ess_nombre
           
           vaSpread1.Col = 3
           vaSpread1.text = RS!mid_numlin
                    
           vaSpread1.Col = 4
           vaSpread1.text = RS!ess_codigo
                    
           RS.MoveNext
        
        Loop
   
  Else
      
      fg_descarga
      vaSpread1.MaxRows = 0
      MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo
   
  End If
  RS.Close
  Set RS = Nothing

  Sql = ""
  Sql = Sql & "'" & Trim(fpText.text) & "', "
  Sql = Sql & "'" & Trim(fpText1.text) & "', "
  Sql = Sql & "" & fpLongInteger1(1).text & ", "
  Sql = Sql & "" & Format(FpFecDesde, ("YYYYMMDD")) & ", "
  Sql = Sql & "" & Format(FpFecHasta, ("YYYYMMDD")) & " "
   
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
  Set RS = vg_db.Execute("sgpadm_Sel_TraerMinutaSeguidores_V01 " & Sql & "")
  If Not RS.EOF Then
   
        Do While Not RS.EOF
           
           vaSpread2.MaxRows = vaSpread2.MaxRows + 1
           vaSpread2.Row = vaSpread2.MaxRows
           
           If RS!EstadoMinuta = "2" Then
           
              vaSpread2.Row = vaSpread2.MaxRows: vaSpread2.Col = -1
              vaSpread2.BackColor = &HC0FFFF
           
           End If
           
           vaSpread2.Col = 1
           vaSpread2.text = "0"
           
           vaSpread2.Col = 2
           vaSpread2.text = RS!cli_codigo
           
           vaSpread2.Col = 3
           vaSpread2.text = RS!cli_nombre
           
           vaSpread2.Col = 4
           vaSpread2.text = RS!min_codreg
           
           vaSpread2.Col = 5
           vaSpread2.text = Trim(RS!reg_nombre)
           
           vaSpread2.Col = 6
           vaSpread2.text = RS!min_codser & " - " & Trim(RS!ser_nombre)
           
           vaSpread2.Col = 7
           vaSpread2.text = ""
                     
           vaSpread2.Col = 8
           vaSpread2.text = RS!min_codser
           
           vaSpread2.Col = 9
           vaSpread2.text = 0
                    
           RS.MoveNext
        
        Loop
   
  Else
      
      fg_descarga
      vaSpread2.MaxRows = 0
      MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo
   
  End If
  RS.Close
  Set RS = Nothing

  fg_descarga
   
End Select

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Function ValidarDatosEntrada() As Boolean

On Error GoTo Man_Error
  
Dim RS        As New ADODB.Recordset
Dim Sql       As String
  
ValidarDatosEntrada = True

  '-------> Validar fechas
  If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
     
     MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
     ValidarDatosEntrada = False
     
     Exit Function
  
  End If
    
  If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
     
     MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
     ValidarDatosEntrada = False
     Exit Function
  
  End If
  
  '-------> Validar que exista Ceco
  If Trim(fpayuda(0).Caption) = "" Then
      
      MsgBox "No existe ceco seleccionado...", vbExclamation + vbOKOnly, MsgTitulo
      ValidarDatosEntrada = False
      Exit Function
   
  End If

  '-------> Validar que exista Regimen
  If Trim(fpayuda(1).Caption) = "" Then
      
      MsgBox "No existe regimen seleccionado...", vbExclamation + vbOKOnly, MsgTitulo
      ValidarDatosEntrada = False
      Exit Function
   
  End If

  '-------> Validar que exista Servicio
  If Trim(fpayuda(2).Caption) = "" Then
      
      MsgBox "No existe servicio seleccionado...", vbExclamation + vbOKOnly, MsgTitulo
      ValidarDatosEntrada = False
      Exit Function
   
  End If

  '-------> Validar que exista Org. Compras
  If fpText1.text = "" Then

      MsgBox "No existe org. compras seleccionado...", vbExclamation + vbOKOnly, MsgTitulo
      ValidarDatosEntrada = False
      Exit Function

  End If
  
  Sql = ""
  Sql = Sql & "'" & Trim(fpText.text) & "', "
  Sql = Sql & "" & fpLongInteger1(0).text & ", "
  Sql = Sql & "" & fpLongInteger1(1).text & ", "
  Sql = Sql & "" & Format(FpFecDesde, ("YYYYMMDD")) & ", "
  Sql = Sql & "" & Format(FpFecHasta, ("YYYYMMDD")) & " "
  
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
  Set RS = vg_db.Execute("  sgpadm_Sel_ValidarsiExistenMasBloqueLideres_V01 " & Sql & "")
  If Not RS.EOF Then
   
     RS.Close
     Set RS = Nothing
   
     MsgBox "Para la fecha desde y hasta existen dos bloques distincto, debe seleccionar una fecha donde exista un solo bloque...", vbExclamation + vbOKOnly, MsgTitulo
     ValidarDatosEntrada = False
     Exit Function
   
  End If
  RS.Close
  Set RS = Nothing


Exit Function
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Function

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1
    
    Dim i As Long
    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
    For i = 1 To vaSpread1.MaxRows 'BlockRow To BlockRow2
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    
        End If
        
    Next
    
End Select

'vaSpread2.MaxRows = 0
'Command1.Enabled = False

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1
    
    Dim i As Long
    vaSpread2.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread2.Row = i
        
        If vaSpread2.RowHidden = False Then
            
           vaSpread2.Value = IIf(vaSpread2.Value = "1", "0", "1")
        
        End If
    
    Next
    
    For i = 1 To vaSpread2.MaxRows 'BlockRow To BlockRow2
        
        vaSpread2.Row = i
        
        If vaSpread2.RowHidden = False Then
            
           vaSpread2.Value = IIf(vaSpread2.Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    vaSpread2.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread2.Row = i
        If vaSpread2.RowHidden = False Then
            
           vaSpread2.Value = IIf(vaSpread2.Value = "1", "0", "1")
    
        End If
        
    Next
    
End Select

Exit Sub
Man_Error:

    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub
