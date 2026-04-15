VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form E_PlanMinuta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Detalle Minuta"
   ClientHeight    =   10305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13140
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   13140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Caption         =   "Modifica columna excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   10560
      TabIndex        =   32
      Top             =   5760
      Width           =   2415
      Begin VB.CheckBox Check1 
         Caption         =   "Realiza cambio Q Total Día"
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
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Racion"
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
         Left            =   600
         TabIndex        =   34
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "% Ponderación"
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
         Left            =   600
         TabIndex        =   33
         Top             =   1080
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin FPSpread.vaSpread vaSpread3 
      Height          =   855
      Left            =   4680
      TabIndex        =   24
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
      _Version        =   393216
      _ExtentX        =   2143
      _ExtentY        =   1508
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
      SpreadDesigner  =   "E_PlanMinuta.frx":0000
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
      Left            =   11640
      TabIndex        =   12
      Top             =   9360
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar"
      Enabled         =   0   'False
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
      Left            =   11640
      TabIndex        =   11
      Top             =   8640
      Width           =   1275
   End
   Begin VB.Frame Frame2 
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
      Height          =   4335
      Left            =   120
      TabIndex        =   18
      Top             =   5760
      Width           =   10335
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   2595
         TabIndex        =   20
         Top             =   3720
         Width           =   7110
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   10
            Top             =   135
            Width           =   7005
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   1
         Left            =   1680
         TabIndex        =   19
         Top             =   3720
         Width           =   900
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   9
            Top             =   135
            Width           =   795
         End
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   3255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   9975
         _Version        =   393216
         _ExtentX        =   17595
         _ExtentY        =   5741
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
         SpreadDesigner  =   "E_PlanMinuta.frx":01F7
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ceco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   720
      TabIndex        =   13
      Top             =   120
      Width           =   11655
      Begin VB.CheckBox Check2 
         Caption         =   "No Mostrar Casinos Propuesta"
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
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   8160
         TabIndex        =   23
         Top             =   4200
         Width           =   2895
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
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
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
            Left            =   1560
            TabIndex        =   5
            Top             =   300
            Width           =   735
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   2
            Left            =   2280
            Picture         =   "E_PlanMinuta.frx":1AE3
            Top             =   160
            Width           =   480
         End
      End
      Begin VB.Frame Frame7 
         Height          =   435
         Left            =   9840
         TabIndex        =   22
         Top             =   3750
         Width           =   900
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   3
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Left            =   2715
         TabIndex        =   15
         Top             =   3750
         Width           =   7110
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   2
            Top             =   135
            Width           =   7005
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   1800
         TabIndex        =   14
         Top             =   3750
         Width           =   900
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            MaxLength       =   246
            TabIndex        =   1
            Top             =   135
            Width           =   795
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3135
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   10845
         _Version        =   393216
         _ExtentX        =   19129
         _ExtentY        =   5530
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
         MaxCols         =   5
         SpreadDesigner  =   "E_PlanMinuta.frx":1DED
         VisibleCols     =   4
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10200
         Top             =   240
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
               Picture         =   "E_PlanMinuta.frx":3730
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   2895
         TabIndex        =   6
         Top             =   5040
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
         Left            =   6660
         TabIndex        =   7
         Top             =   5040
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   8280
         TabIndex        =   16
         Top             =   4980
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Plato"
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
         TabIndex        =   29
         Top             =   4725
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   1355
         Picture         =   "E_PlanMinuta.frx":3ACA
         Top             =   4560
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   1860
         TabIndex        =   28
         Top             =   4650
         Width           =   6075
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   1860
         TabIndex        =   26
         Top             =   4305
         Width           =   6075
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C. Dietetica"
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
         Left            =   240
         TabIndex        =   25
         Top             =   4365
         Width           =   1020
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   1335
         Picture         =   "E_PlanMinuta.frx":3DD4
         Top             =   4200
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha desde"
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
         Left            =   1560
         TabIndex        =   21
         Top             =   5115
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha hasta"
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
         Left            =   5355
         TabIndex        =   17
         Top             =   5115
         Width           =   1065
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   1905
         TabIndex        =   27
         Top             =   4350
         Width           =   6075
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   1905
         TabIndex        =   30
         Top             =   4695
         Width           =   6075
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "E_PlanMinuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public lc_Aux As String
Dim MsgTitulo As String
Dim FilCatDie As Long
Dim FilTipPla As Long

Private Sub Check1_Click()

On Error GoTo Man_Error

If Check1.Value = 1 Then

    Option2(0).Enabled = False
    Option2(1).Enabled = False

Else

    Option2(0).Enabled = True
    Option2(1).Enabled = True

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Check2_Click()

Form_Load

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim AuxFilCatDie As Long
Dim AuxFilTipPla As Long

Me.HelpContextID = vg_OpcM
fg_centra Me
MsgTitulo = "Exportar Excel Detalle Minuta II"

'-------> leer datos parametros categoria ditetica
FilCatDie = 0
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
Set RS = vg_db.Execute("SELECT par_valor FROM b_paramtiporeceta WITH ( HOLDLOCK ) WHERE par_codigo = 'CatDiePlan-" & Trim(vg_NUsr) & "'")
If Not RS.EOF Then FilCatDie = RS!par_valor: AuxFilCatDie = RS!par_valor
RS.Close
Set RS = Nothing

If FilCatDie = 0 Then
   
   fpayuda(0).Caption = "Todos"

Else
   
   fpayuda(0).Caption = fg_BuscaenArbol(AuxFilCatDie, "a_recetacatdie", "car_codigo")

End If

'-------> leer datos parametros tipo plato
FilTipPla = 0
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
Set RS = vg_db.Execute("SELECT par_valor FROM b_paramtiporeceta WITH ( HOLDLOCK ) WHERE par_codigo = 'TipPlaPlan-" & Trim(vg_NUsr) & "'")
If Not RS.EOF Then FilTipPla = RS!par_valor: AuxFilTipPla = RS!par_valor
RS.Close
Set RS = Nothing

If FilTipPla = 0 Then
   
   fpayuda(1).Caption = "Todos"

Else
   
   fpayuda(1).Caption = fg_BuscaenArbol(AuxFilTipPla, "a_recetatippla", "tip_codigo")

End If

FpFecHasta.text = Format(Date, "dd/mm/yyyy")
FpFecDesde.text = Format(Date, "dd/mm/yyyy")
vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0

TextDet1(2).text = ""
TextDet1(3).text = ""

TextDet2(2).text = ""
TextDet2(3).text = ""

CargarCeco
CargarServicio

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub CargarCeco()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i As Long

vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &HC0FFFF

vaSpread1.MaxRows = 0
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
Set RS = vg_db.Execute("sgpadm_Sel_FiltrarCecoxPropoProd_V02 '" & IIf(Check2.Value = 1, "C", "") & "'")

If Not RS.EOF Then
  Do While Not RS.EOF
      
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      
      vaSpread1.Col = 1
      vaSpread1.text = "0"
      
      vaSpread1.Col = 2
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS(0)
      
      vaSpread1.Col = 3
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = Trim(RS(1))
      
      vaSpread1.Col = 4
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = Trim(RS(2))
         
      vaSpread1.Col = 5
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = 0
      
      RS.MoveNext
  Loop
Else
   vaSpread1.MaxRows = 0
   MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo
End If
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub CargarServicio()

On Error GoTo Man_Error

Dim RS        As New ADODB.Recordset
Dim i         As Long
Dim codCeco   As String
Dim Sql       As String
Dim seleccion As Integer

'--> Concatenar codigo ceco
codCeco = ""

For i = 1 To vaSpread1.MaxRows
       
    vaSpread1.Row = i
    vaSpread1.Col = 1 'Seleccion
    seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
    If seleccion = 1 And vaSpread1.RowHidden = False Then

       vaSpread1.Col = 2
       codCeco = codCeco & "'" & vaSpread1.text & "', "

    End If
  
Next i

Sql = ""
If Trim(codCeco) <> "" Then
   
   Sql = Sql & Replace(Mid(codCeco, 1, Len(codCeco) - 2), "'", """")

End If

vaSpread3.MaxRows = 0
'Set RS = vg_db.Execute("sgpadm_s_servicio 10, '', 0, 0")
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
Set RS = vg_db.Execute("sgpadm_Sel_ServiciosClientes '" & Sql & "', '" & Format(FpFecDesde.Value, "yyyymmdd") & "', '" & Format(FpFecHasta.Value, "yyyymmdd") & "'")

If Not RS.EOF Then
  Do While Not RS.EOF
      
      vaSpread3.MaxRows = vaSpread3.MaxRows + 1
      vaSpread3.Row = vaSpread3.MaxRows
      
      vaSpread3.Col = 1
      vaSpread3.text = "0"
      
      vaSpread3.Col = 2
      vaSpread3.text = RS(0)
      
      vaSpread3.Col = 3
      vaSpread3.text = Trim(RS(1))
      
      vaSpread3.Col = 4
      vaSpread3.text = Trim(RS(2))
         
      RS.MoveNext
  
  Loop

Else
   
   vaSpread3.MaxRows = 0
'   MsgBox "No existe información requerida servicio", vbExclamation + vbOKOnly, Msgtitulo

End If

RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim RS1             As New ADODB.Recordset
Dim Sql             As String
Dim Sql1            As String
Dim Sql2            As String
Dim codCeco         As String
Dim CodReceta       As String
Dim codser          As String
Dim NomArchivoExcel As String
Dim Extension       As String
Dim seleccion       As Integer
Dim i               As Long

Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object

Dim ClaveExcel      As String

If Not ValidarDatos Then Exit Sub

'--> Concatenar codigo ceco
codCeco = ""
For i = 1 To vaSpread1.MaxRows
       
    vaSpread1.Row = i
    vaSpread1.Col = 1 'Seleccion
    seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
    If seleccion = 1 And vaSpread1.RowHidden = False Then

       vaSpread1.Col = 2
       codCeco = codCeco & "'" & vaSpread1.text & "', "

    End If
  
Next i

If Trim(codCeco) = "" Then
  
   MsgBox "No existen datos seleccionados grilla cecos...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub

End If

'--> Concatenar codigo receta
CodReceta = ""
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then

       vaSpread2.Col = 2
       CodReceta = CodReceta & "" & vaSpread2.text & ", "

    End If
  
Next i

codser = ""
For i = 1 To vaSpread3.MaxRows
       
    vaSpread3.Row = i
    vaSpread3.Col = 1 'Seleccion
    seleccion = IIf(vaSpread3.text = "", 0, vaSpread3.text)
    
    If seleccion = 1 And vaSpread3.RowHidden = False Then

       vaSpread3.Col = 2
       codser = codser & "" & vaSpread3.text & ", "

    End If
  
Next i

'-------> Validar cantidad registro se sobre pase hoja excel
Sql = ""
Sql = Sql & Replace(Mid(codCeco, 1, Len(codCeco) - 2), "'", """")
Sql1 = ""
If CodReceta <> "" Then
   
   Sql1 = Sql1 & Mid(CodReceta, 1, Len(CodReceta) - 2)

End If

Sql2 = ""
If codser <> "" Then

   Sql2 = Sql2 & Mid(codser, 1, Len(codser) - 2)
   
End If

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
If Check1.Value = 1 Then

   Set RS = vg_db.Execute("sgpadm_Sel_ComensalesMinutaBloque '" & Sql & "', '" & Sql2 & "', '" & Format(FpFecDesde, ("YYYYmmdd")) & "', '" & Format(FpFecHasta, ("YYYYmmdd")) & "'")

Else

   Set RS = vg_db.Execute("sgpadm_Sel_ValidarNRegDetalleMinuta_2 '" & Sql & "', '" & Sql1 & "', '" & Sql2 & "', '" & Format(FpFecDesde, ("YYYYmmdd")) & "', '" & Format(FpFecHasta, ("YYYYmmdd")) & "'")

End If

If Not RS.EOF Then
  
   If Check1.Value = 1 Then
   
      If RS.RecordCount > 1020000 Then
   
         MsgBox "El resultado sobrepasa maximo de fila en excel, Debera seleccionar menos Ceco", vbCritical
         Exit Sub
   
      End If
      
   Else
   
      If RS!nReg > 1020000 Then
      
         MsgBox "El resultado sobrepasa maximo de fila en excel, Debera seleccionar menos Ceco", vbCritical
         Exit Sub
   
      End If
   
   End If
   
End If
  
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
          
Toolbar2.Enabled = False
FpFecDesde.Enabled = False
FpFecHasta.Enabled = False
Frame1.Enabled = False
Frame2.Enabled = False

fg_carga ""
  
Sql = ""
Sql = Sql & Replace(Mid(codCeco, 1, Len(codCeco) - 2), "'", """")

Sql1 = ""

If Trim(CodReceta) <> "" Then
   
   Sql1 = Sql1 & Mid(CodReceta, 1, Len(CodReceta) - 2)
   
End If

Sql2 = ""
If codser <> "" Then

   Sql2 = Sql2 & Mid(codser, 1, Len(codser) - 2)
   
End If

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
If Check1.Value = 1 Then

   Set RS = vg_db.Execute("sgpadm_Sel_ComensalesMinutaBloque '" & Sql & "', '" & Sql2 & "', '" & Format(FpFecDesde, ("YYYYmmdd")) & "', '" & Format(FpFecHasta, ("YYYYmmdd")) & "'")

Else

   Set RS = vg_db.Execute("sgpadm_Sel_DetalleMinuta_V04 '" & Sql & "', '" & Sql1 & "' , '" & Sql2 & "', '" & Format(FpFecDesde, ("YYYYmmdd")) & "', '" & Format(FpFecHasta, ("YYYYmmdd")) & "', '" & IIf(Option2(0).Value = True, "0", "1") & "'")
   
End If
  
'-------> Create an instance of Excel and add a workbook
Set xlApp = CreateObject("Excel.Application")
Set xlWb = xlApp.Workbooks.Add
Set xlWs = xlWb.Worksheets("Hoja1")
  
'-------> Display Excel and give user control of Excel's lifetime
xlApp.UserControl = True
    
'-------> Check version of Excel
Call encabezado(RS, xlWs)
          
xlWs.Cells(2, 1).CopyFromRecordset RS

Dim Filas As Long

Filas = RS.RecordCount
'-------> desprotecger columnas
If Check1.Value = 1 Then

   xlApp.Columns("H:H").Select
   xlApp.Selection.Locked = False
   xlApp.Selection.FormulaHidden = False

Else

    xlApp.Columns("t:t").Select
    xlApp.Selection.NumberFormat = "0"
        
    xlApp.Range("t:t").Select
    xlApp.Range("t:t").Activate
    xlApp.Selection.Replace What:="0", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        
    xlApp.Columns("T:T").Select
    xlApp.Selection.Locked = False
    xlApp.Selection.FormulaHidden = False
    
    If Option2(0).Value = True Then
    
       xlApp.Columns("P:P").Select
       xlApp.Selection.NumberFormat = "0"
    
       xlApp.Columns("P:P").Select
       xlApp.Selection.Locked = False
       xlApp.Selection.FormulaHidden = False
       
       xlApp.Range("Q2").Select
       xlApp.ActiveCell.FormulaR1C1 = "=(+RC[-1]*RC[1])/100"
'       xlApp.Range("Q2").Select
       
'       xlApp.Selection.AutoFill Destination:=xlApp.Range("Q2:Q6265")
       xlApp.Selection.AutoFill Destination:=xlApp.Range("Q2:Q" & Filas + 1)
       xlApp.Range("Q2:Q" & Filas + 1).Select
       
       xlApp.Columns("Q:Q").Select
       xlApp.Selection.NumberFormat = "0"

       xlApp.Range("P2:P" & Filas + 1).Select
'       xlApp.Range(Selection, Selection.End(xlDown)).Select
       With xlApp.Selection.Interior
            
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        
        End With
    
    Else
    
       xlApp.Columns("Q:Q").Select
       xlApp.Selection.NumberFormat = "0"
       
       xlApp.Columns("Q:Q").Select
       xlApp.Selection.Locked = False
       xlApp.Selection.FormulaHidden = False
    
       xlApp.Range("P2").Select
       xlApp.ActiveCell.FormulaR1C1 = "=(+RC[1]/RC[2])*100"
'       xlApp.Range("P2").Select
       xlApp.Selection.AutoFill Destination:=xlApp.Range("P2:P" & Filas + 1)
       xlApp.Range("P2:P" & Filas + 1).Select
    
       xlApp.Columns("P:P").Select
       xlApp.Selection.NumberFormat = "0"
    
       xlApp.Range("Q2:Q" & Filas + 1).Select
'       xlApp.Range(Selection, Selection.End(xlDown)).Select
       With xlApp.Selection.Interior
            
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        
        End With
    
    End If

End If

'-------> Auto-fit the column widths and row heights
xlApp.Selection.CurrentRegion.Columns.AutoFit
xlApp.Selection.CurrentRegion.Rows.AutoFit
    
xlApp.Columns("A:A").Select
xlApp.Selection.Delete Shift:=xlToLeft
  
ClaveExcel = "Jp123456"
             
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
Set RS1 = vg_db.Execute("sgpadm_s_parametro 1, 'parhojaexc', ''")
If Not RS1.EOF Then
                
   ClaveExcel = RS1(0)
             
End If
RS1.Close
Set RS1 = Nothing

'comenta bloque de plantilla 20191017
'xlApp.xlWs.Select
'xlApp.ActiveSheet.Protect password:=ClaveExcel, DrawingObjects:=True, _
'Contents:=True, Scenarios:=True, AllowFormattingCells:=True

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
                
'ProgressBar1.Visible = False
'lbl_proceso.Visible = False
  
Toolbar2.Enabled = True
FpFecDesde.Enabled = True
FpFecHasta.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True


Exit Sub
Man_Error:
    Frame1.Enabled = True
    Frame2.Enabled = True
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Command2_Click()

On Error GoTo Man_Error

Unload Me

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub FpFecDesde_Change()
On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub

Command1.Enabled = False
vaSpread2.MaxRows = 0
CargarServicio

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub

Command1.Enabled = False
vaSpread2.MaxRows = 0
CargarServicio

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
End Sub

Private Sub Image1_Click(Index As Integer)
On Error GoTo Man_Error

Select Case Index

Case 0
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(0).Left + 2400
    B_ArbEst.MoverDatosTvwDir "a_recetacatdie", "car_", "Categoria Dietetica", "1"
    B_ArbEst.Show 1
    If vg_codigo = "" Then Exit Sub
    FilCatDie = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre: vg_nombre = ""

Case 1
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(1).Left + 2400
    B_ArbEst.MoverDatosTvwDir "a_recetatippla", "tip_", "Tipo Plato", "1"
    B_ArbEst.Show 1
    If Trim(vg_codigo) = "" Then Exit Sub
    FilTipPla = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre: vg_nombre = ""

Case 2

    vg_left = Image1(2).Left + 2300
    vg_nombre = "": vg_codigo = ""
    'OpcionLectura = "1"
'    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread3, 0, 0, 0, 0, 0, "4", 0
    B_MTaEst.LlenaDatos "Servicio", Me, 0, 0, 0, 0, 0, "7", 0
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub

End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0

    Image1(2).Enabled = False
    
Case 1

    Image1(2).Enabled = True
    
End Select

Command1.Enabled = False
vaSpread2.MaxRows = 0

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

'Private Sub TextDet1_Change(Index As Integer)
'
'On Error GoTo Man_Error
'
'Dim i          As Long
'Dim indactivo  As Integer
'
'If Index = 2 Then
'   TextDet1(3).text = ""
'ElseIf Index = 3 Then
'   TextDet1(2).text = ""
'ElseIf Index = 4 Then
'   TextDet1(2).text = ""
'   TextDet1(3).text = ""
'End If
'Select Case Index
'Case 2, 3, 4
'    vaSpread1.Visible = False
'    If Trim(TextDet1(Index).text) <> "" Then
'       For i = 1 To vaSpread1.MaxRows
'           vaSpread1.Row = i
'           vaSpread1.Col = Index
'           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(TextDet1(Index).text) & "*"
'           vaSpread1.Col = 1
'           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
'              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
'           Else
'              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
'           End If
'        Next i
'        vaSpread1.SetActiveCell Index + 1, 1
'    End If
'    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
'    vaSpread1.ColUserSortIndicator(IIf(Trim(TextDet1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
'    vaSpread1.SortKey(1) = IIf(Trim(TextDet1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
'    vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
'    If Trim(TextDet1(Index).text) = "" Then
'       For i = 1 To vaSpread1.MaxRows
'           vaSpread1.Row = i
'           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
'       Next
'       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextDet1(Index).text), SearchFlagsGreaterOrEqual)
'       vaSpread1.SetActiveCell Index, 1
'    End If
'    vaSpread1.Visible = True
'End Select
'
'Exit Sub
'Man_Error:
'    fg_descarga
'    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
'    ins_log_error Date & Time & Err & ":  " & Error$(Err)
'End Sub

Private Sub TextDet1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(TextDet1(Index).text, ",")

If Index = 2 Then
   TextDet1(3).text = ""
ElseIf Index = 3 Then
   TextDet1(2).text = ""
ElseIf Index = 4 Then
   TextDet1(2).text = ""
   TextDet1(3).text = ""
End If

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 5
    vaSpread1.text = 0
    
Next

Select Case Index

Case 2, 3, 4
    
    vaSpread1.Visible = False
    
    If Trim(TextDet1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread1.Col = 1
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 5
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 5
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 5
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 5
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 1
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 5
                 vaSpread1.text = 0
                 
              End If
           
           End If
        
        Next i
        
        vaSpread1.SetActiveCell Index + 1, 1
        vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread1.ColUserSortIndicator(IIf(Trim(wvarArr8020(X)) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread1.SortKey(1) = IIf(Trim(wvarArr8020(X)) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(TextDet1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 5
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextDet1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

'Private Sub TextDet2_Change(Index As Integer)
'
'On Error GoTo Man_Error
'
'Dim i          As Long
'Dim indactivo  As Integer
'
'If Index = 2 Then
'   TextDet2(3).text = ""
'ElseIf Index = 3 Then
'   TextDet2(2).text = ""
'End If
'Select Case Index
'Case 2, 3
'    vaSpread2.Visible = False
'    If Trim(TextDet2(Index).text) <> "" Then
'       For i = 1 To vaSpread2.MaxRows
'           vaSpread2.Row = i
'           vaSpread2.Col = Index
'           indactivo = UCase(Trim(vaSpread2.Value)) Like "*" & UCase(TextDet2(Index).text) & "*"
'           vaSpread2.Col = 1
'           If indactivo = -1 And Trim(vaSpread2.text) <> "" Then
'              If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
'           Else
'              If vaSpread2.RowHidden = False Then vaSpread2.RowHidden = True
'           End If
'        Next i
'        vaSpread2.SetActiveCell Index + 1, 1
'    End If
'    vaSpread2.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
'    vaSpread2.ColUserSortIndicator(IIf(Trim(TextDet2(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
'    vaSpread2.SortKey(1) = IIf(Trim(TextDet2(Index).text) = "", 0, 0): vaSpread2.SortKeyOrder(1) = SortKeyOrderAscending
'    vaSpread2.Sort -1, -1, vaSpread2.maxcols, vaSpread2.MaxRows, SortByRow
'    If Trim(TextDet2(Index).text) = "" Then
'       For i = 1 To vaSpread2.MaxRows
'           vaSpread2.Row = i
'           If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
'       Next
'       vaSpread2.SetActiveCell Index, vaSpread2.SearchCol(Index, 0, vaSpread2.MaxRows, Trim(TextDet2(Index).text), SearchFlagsGreaterOrEqual)
'       vaSpread2.SetActiveCell Index, 1
'    End If
'    vaSpread2.Visible = True
'End Select
'
'Exit Sub
'Man_Error:
'    fg_descarga
'    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
'    ins_log_error Date & Time & Err & ":  " & Error$(Err)
'
'End Sub

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
End If

For i = 1 To vaSpread2.MaxRows
           
    vaSpread2.Row = i
    vaSpread2.Col = 4
    vaSpread2.text = 0
    
Next

Select Case Index

Case 2, 3
    
    vaSpread2.Visible = False
    
    If Trim(TextDet2(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread2.MaxRows
           
           vaSpread2.Row = i
           vaSpread2.Col = Index
           indactivo = UCase(Trim(vaSpread2.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread2.Col = 1
           
           If indactivo = -1 And Trim(vaSpread2.text) <> "" Then
              
              vaSpread2.Col = 4
              
              If Val(vaSpread2.Value) <> 1 Then
                              
                 vaSpread2.Col = 1
              
                 If vaSpread2.RowHidden = True Then
                 
                    vaSpread2.RowHidden = False
                    vaSpread2.Col = 4
                    vaSpread2.text = 1
                 
                 Else
                 
                    vaSpread2.Col = 4
                    vaSpread2.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread2.Col = 4
              EstBuq = vaSpread2.Value
              vaSpread2.Col = 1
              
              If vaSpread2.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread2.RowHidden = True
                 
                 vaSpread2.Col = 4
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
           
           vaSpread2.Col = 4
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
Dim Sql1      As String
Dim i         As Long
Dim xmlceco   As String
Dim seleccion As String
Dim codCeco   As String
Dim codser    As String

Select Case Button.Index
Case 1

   If Not ValidarDatos Then Exit Sub
  
  vaSpread2.MaxRows = 0
  vaSpread2.Row = -1: vaSpread2.Col = -1
  vaSpread2.BackColor = &HC0FFFF
   
  TextDet2(2).text = ""
  TextDet2(3).text = ""
   
  codCeco = ""
  For i = 1 To vaSpread1.MaxRows
       
    vaSpread1.Row = i
    vaSpread1.Col = 1 'Seleccion
    seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
    If seleccion = 1 And vaSpread1.RowHidden = False Then

       vaSpread1.Col = 2
       codCeco = codCeco & "'" & vaSpread1.text & "', "

    End If
  
  Next i
  

  If Option1(0).Value = True Then
  
     vaSpread3.Col = 1
     vaSpread3.Row = -1
     vaSpread3.Value = "0"
  
  End If
  
  codser = ""
  For i = 1 To vaSpread3.MaxRows
       
    vaSpread3.Row = i
    vaSpread3.Col = 1 'Seleccion
    seleccion = IIf(vaSpread3.text = "", 0, vaSpread3.text)
    
    If seleccion = 1 And vaSpread3.RowHidden = False Then

       vaSpread3.Col = 2
       codser = codser & "" & vaSpread3.text & ", "

    End If
  
  Next i
  
  If Trim(codCeco) = "" Then
  
   MsgBox "No existen datos seleccionados grilla cecos...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub
  
  End If
  
  Sql = ""
  Sql = Sql & Replace(Mid(codCeco, 1, Len(codCeco) - 2), "'", """")
  
  Sql1 = ""
  If Trim(codser) <> "" Then
  
     Sql1 = Sql1 & Sql1 & Mid(codser, 1, Len(codser) - 2)
     
  End If
  
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
    
  Set RS = vg_db.Execute("sgpadm_Sel_DetalleMinuta_3 '" & Sql & "', '" & Sql1 & "', '" & Format(FpFecDesde, ("YYYYmmdd")) & "', '" & Format(FpFecHasta, ("YYYYmmdd")) & "', '" & IIf(FilCatDie = 0, "", FilCatDie) & "', '" & IIf(FilTipPla = 0, "", FilTipPla) & "'")
  If Not RS.EOF Then
     
     Do While Not RS.EOF
      
        vaSpread2.MaxRows = vaSpread2.MaxRows + 1
        vaSpread2.Row = vaSpread2.MaxRows
      
        vaSpread2.Col = 1
        vaSpread2.text = "0"
      
        vaSpread2.Col = 2
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.text = RS(0)
      
        vaSpread2.Col = 3
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.text = Trim(RS(1))
      
        vaSpread2.Col = 4
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.text = 0
      
        RS.MoveNext
     Loop
     
     Command1.Enabled = True
  
  Else
     
     Command1.Enabled = False
     vaSpread2.MaxRows = 0
     MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo
  
  End If
  RS.Close
  Set RS = Nothing

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

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

CargarServicio

vaSpread2.MaxRows = 0
Command1.Enabled = False

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Function ValidarDatos() As Boolean

Dim seleccion As Integer
Dim i As Long

ValidarDatos = True

''-------> Validar org. compras
'If Trim(fpText.text) = "" Then
'
'     ValidarDatos = False
'     MsgBox "Debe ingresar Org. Compras...", vbExclamation + vbOKOnly, Msgtitulo
'     Exit Function
'
'End If
  
'-------> Validar fechas
If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
   
   MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If
    
If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
   
   MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

'-------> Validar que exista un dato seleccionado
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
     
   MsgBox " Se debe seleccionar un Bloque por lo menos", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function
  
End If

'-------> Validar que exista un dato seleccionado servicio
If Option1(1).Value = True Then
   seleccion = 0
   For i = 1 To vaSpread3.MaxRows
       
       vaSpread3.Row = i
       vaSpread3.Col = 1 'Seleccion
       seleccion = IIf(vaSpread3.text = "", 0, vaSpread3.text)
    
       If seleccion = 1 And vaSpread3.RowHidden = False Then
          
          Exit For
          
       End If
  
    Next i
  
    If seleccion = 0 Then
     
       MsgBox " Se debe seleccionar un Servicio por lo menos", vbExclamation + vbOKOnly, MsgTitulo
       ValidarDatos = False
       Exit Function
  
    End If

End If

End Function

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
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub encabezado(ByRef RS As ADODB.Recordset, ByRef xlWs As Object)
On Error GoTo Man_Error

Dim fldCount As Long
Dim icol As Long

'-------> Copy field names to the first row of the worksheet
fldCount = RS.Fields.count
For icol = 1 To fldCount
    xlWs.Cells(1, icol).Value = RS.Fields(icol - 1).Name
Next

Exit Sub
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo
End Sub

