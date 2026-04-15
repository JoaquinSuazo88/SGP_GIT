VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{1DF3AFED-99E0-4474-9900-954B8FD24E86}#1.0#0"; "ChilkatMail2.dll"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_GenPed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación Pedidos"
   ClientHeight    =   10605
   ClientLeft      =   1365
   ClientTop       =   1635
   ClientWidth     =   17775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10605
   ScaleWidth      =   17775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1425
      Index           =   1
      Left            =   4650
      TabIndex        =   1
      Top             =   120
      Width           =   8175
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1380
         TabIndex        =   2
         Top             =   285
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
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   360
         Left            =   6945
         TabIndex        =   3
         Top             =   690
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
         Left            =   1380
         TabIndex        =   4
         Top             =   780
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
         Index           =   1
         Left            =   5055
         TabIndex        =   5
         Top             =   780
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
         Index           =   1
         Left            =   3840
         TabIndex        =   10
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3045
         TabIndex        =   8
         Top             =   285
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2610
         Picture         =   "M_GenPed.frx":0000
         Top             =   195
         Width           =   480
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
         Left            =   90
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
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
         Left            =   90
         TabIndex        =   6
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3090
         TabIndex        =   9
         Top             =   330
         Width           =   4935
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   10605
      Left            =   17265
      TabIndex        =   0
      Top             =   0
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   18706
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   300
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
            Picture         =   "M_GenPed.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8385
      Left            =   135
      TabIndex        =   11
      Top             =   1785
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   14790
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ingreso Pedido"
      TabPicture(0)   =   "M_GenPed.frx":06A4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vaSpread1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Resumen Pedido"
      TabPicture(1)   =   "M_GenPed.frx":06C0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vaSpread4"
      Tab(1).Control(1)=   "vaSpread2"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   5190
         TabIndex        =   19
         Top             =   2685
         Width           =   5055
         Begin VB.Label Label3 
            Caption         =   "Un Momento Generando Necesidad"
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
            Left            =   1110
            TabIndex        =   20
            Top             =   420
            Width           =   3735
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4455
         Left            =   3390
         TabIndex        =   13
         Top             =   1650
         Width           =   9015
         Begin VB.CommandButton cmd_Aceptar 
            Caption         =   "Aceptar"
            Height          =   360
            Left            =   5955
            TabIndex        =   16
            Top             =   3825
            Width           =   1200
         End
         Begin VB.CommandButton cmd_cancelar 
            Caption         =   "Cancelar"
            Height          =   405
            Left            =   4110
            TabIndex        =   15
            Top             =   3855
            Width           =   1635
         End
         Begin VB.CommandButton cmd_generar_pedido 
            Caption         =   "Generar Pedido"
            Height          =   465
            Left            =   2340
            TabIndex        =   14
            Top             =   3840
            Width           =   1380
         End
         Begin FPSpread.vaSpread vaSpread3 
            Height          =   2400
            Left            =   1095
            TabIndex        =   17
            Top             =   1065
            Width           =   6900
            _Version        =   393216
            _ExtentX        =   12171
            _ExtentY        =   4233
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
            SpreadDesigner  =   "M_GenPed.frx":06DC
         End
         Begin VB.Label lbl_mensaje 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1005
            TabIndex        =   18
            Top             =   390
            Width           =   7080
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   7260
         Left            =   30
         TabIndex        =   12
         Top             =   555
         Width           =   17010
         _Version        =   393216
         _ExtentX        =   30004
         _ExtentY        =   12806
         _StockProps     =   64
         AutoClipboard   =   0   'False
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   14
         MaxRows         =   1
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_GenPed.frx":1FA2
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   2535
         Left            =   -72570
         TabIndex        =   21
         Top             =   1695
         Width           =   9075
         _Version        =   393216
         _ExtentX        =   16007
         _ExtentY        =   4471
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
         MaxCols         =   5
         SpreadDesigner  =   "M_GenPed.frx":2715
      End
      Begin FPSpread.vaSpread vaSpread4 
         Height          =   2535
         Left            =   -72585
         TabIndex        =   22
         Top             =   4875
         Width           =   11130
         _Version        =   393216
         _ExtentX        =   19632
         _ExtentY        =   4471
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
         MaxCols         =   12
         SpreadDesigner  =   "M_GenPed.frx":402C
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   180
      Top             =   1000
      Visible         =   0   'False
      Width           =   300
   End
   Begin CHILKATMAILLib2Ctl.ChilkatMailMan2 oMail 
      Left            =   2640
      OleObjectBlob   =   "M_GenPed.frx":59E4
      Top             =   600
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP oFTP 
      Left            =   1920
      OleObjectBlob   =   "M_GenPed.frx":5A08
      Top             =   870
   End
End
Attribute VB_Name = "M_GenPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim MsgTitulo As String

Private Sub cmd_Aceptar_Click()
vaSpread1.Enabled = True
Frame1(1).Enabled = True
Toolbar1.Enabled = True
Frame3.Visible = False
vaSpread1.MaxRows = 0
Toolbar1.Buttons(1).Visible = False
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub cmd_cancelar_Click()
vaSpread1.Enabled = True
Frame1(1).Enabled = True
Toolbar1.Enabled = True
Frame3.Visible = False
vaSpread1.MaxRows = 0
Toolbar1.Buttons(1).Visible = False
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub cmd_generar_pedido_Click()
vaSpread1.Enabled = True
Frame1(1).Enabled = True
Toolbar1.Enabled = True
Frame3.Visible = False
Call generar_pedido
vaSpread1.MaxRows = 0
Toolbar1.Buttons(1).Visible = False
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub LimpiarControles()
    vaSpread1.MaxRows = 0
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
fg_carga ""
Me.Height = 11025
Me.Width = 17895
fg_centra Me
MsgTitulo = "Generación Pedidos"
 Toolbar1.ImageList = Partida.IL1
 Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
 Set BtnX = Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
 Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
 Set BtnX = Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = "Eliminar Pedido "
 Set BtnX = Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = ""
 Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
 Set BtnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): BtnX.Visible = False: BtnX.ToolTipText = "Enviar a PEL": BtnX.Enabled = True
 Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
 Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = False: BtnX.ToolTipText = "Exportar a Excel "
 Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
 Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"


vaSpread4.Visible = False

fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
fpText.text = ""
Label3.Visible = False
Frame2.Visible = False
Frame3.Visible = False

fg_descarga
Call LimpiarControles
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then Exit Sub
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then Exit Sub
End Sub

Private Sub fpText_Change()
Dim RS As New ADODB.Recordset
Dim Sql As String
If fpText.text = "" Then fpayuda.Caption = "": Exit Sub
Sql = Trim(LimpiaDato(fpText.text))
Set RS = vg_db.Execute("sgpadm_s_cliente_V02 29, '" & Sql & "', ''")
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda.Caption = "": Exit Sub
fpayuda.Caption = Trim(RS!Cli_nombre)
Call busca_encabezado
RS.Close: Set RS = Nothing
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    Image1_Click
End Select
End Sub

Private Sub Image1_Click()
vg_left = fpayuda.Left + 2300
vg_nombre = "": vg_codigo = ""
B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "Clientesimap"
B_TabEst.Show 1
Me.Refresh
If vg_codigo = "" Then Exit Sub
fpText.text = vg_codigo: fpayuda.Caption = vg_nombre
If Me.Visible Then fpDateTime1(0).SetFocus
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
    
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(7).Visible = True
    Toolbar1.Buttons(9).Visible = True


Else
    Toolbar1.Buttons(4).Visible = False
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(9).Visible = False
    
End If
Call busca_encabezado
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 11 '-------> Salir
    Me.Hide
    Unload Me
 
Case 1 'Generar el Pedido
    
    vaSpread1.Enabled = False
    Frame1(1).Enabled = False
    Toolbar1.Enabled = False
    
         valida_Pedidos M_GenPed.vaSpread1, fpText, Format(fpDateTime1(0).text, "yyyymmdd"), Format(fpDateTime1(1).text, "yyyymmdd"), 2
         
        If pedidos = 1 Then
            lbl_mensaje.Caption = "No de puede generar Pedido ya fue enviado a PEL"
            Frame3.Visible = True
            vaSpread3.Visible = True
            cmd_generar_pedido.Visible = False
            cmd_cancelar.Visible = False
            cmd_Aceptar.Visible = True
        Else
            valida_Pedidos M_GenPed.vaSpread1, fpText, Format(fpDateTime1(0).text, "yyyymmdd"), Format(fpDateTime1(1).text, "yyyymmdd"), 1
           If pedidos = 1 Then
             lbl_mensaje.Caption = "Existen los siguentes Pedidos, y si presiona Generar Pedido se Borrraran y se generara un Nuevo Pedido"
             vaSpread3.Visible = True
           Else
            lbl_mensaje.Caption = "Presione el Botón Generar Pedido, para Confirmar la Creación"
            vaSpread3.Visible = False
            cmd_generar_pedido.Visible = True
            cmd_cancelar.Visible = True
            cmd_Aceptar.Visible = False
           End If
           Frame3.Visible = True
           cmd_generar_pedido.Visible = True
           cmd_cancelar.Visible = True
           cmd_Aceptar.Visible = False
        End If
 
 Case 4 'eliminar pedido
    Dim estado As String
    Dim pedido As Integer
    If vaSpread2.MaxRows < 1 Then Exit Sub
    vaSpread2.Row = vaSpread2.ActiveRow
    vaSpread2.Col = 1:  pedido = vaSpread2.text
    vaSpread2.Col = 2: estado = vaSpread2.text
    
    
    If (estado = "Enviado" Or estado = "Descarga Minuta") Then
         MsgBox "No se puede eliminar el pedido, fue enviado a PEL"
         Exit Sub
    Else
    msg = "¿Esta seguro que desea eliminar " + "el pedido: " + CStr(pedido) + "?"
                Response = MsgBox(msg, 4 + 32, "Sistema Gestión")
                If Response = 6 Then
                      Sql = " sgpadm_del_EliminarPedidosEncabezado "
                      Sql = Sql & pedido
                      Set RS = vg_db.Execute(Sql)
                      MsgBox "Eliminacion Termino Correctamente", vbInformation + vbOKOnly, MsgTitulo
                      Call busca_encabezado
                      Exit Sub
                End If
                
    
    End If

 Case 7 'enviar pedido a PEL
    If vaSpread2.MaxRows < 1 Then Exit Sub
    vaSpread2.Row = vaSpread2.ActiveRow
    vaSpread2.Col = 1:  pedido = vaSpread2.text
    vaSpread2.Col = 2: estado = vaSpread2.text
    
    Set RS = vg_db.Execute("sgpadm_sel_seleccionaDetallePedido " & pedido & ", 'A'")
    If Not RS.EOF Then
            MsgBox "El Pedido no puede ser enviado a PEL, debido a que tiene items sin rutas o convenios ", 16
            Exit Sub
       RS.Close: Set RS = Nothing

    End If
    
    If estado = "Enviado" Or estado = "Descarga Minuta" Then
         MsgBox "El Pedido ya que se Encuentra Enviado a PEL"
         Exit Sub
    Else
    msg = "¿Esta seguro que desea enviar a PEL el pedido: " + CStr(pedido) + "?"
                Response = MsgBox(msg, 4 + 32, "Sistema Gestión")
                If Response = 6 Then
                      Sql = " sgpadm_iu_actualizaestado "
                      Sql = Sql & pedido & "," & 2
                      Set RS = vg_db.Execute(Sql)
                      MsgBox "Envio a PEL Termino Correctamente", vbInformation + vbOKOnly, MsgTitulo
                      Call busca_encabezado
                      Exit Sub
                End If
                
    
    End If

 
    Call busca_encabezado
    
 Case 9 'Llevar a Excel
  If vaSpread2.MaxRows < 1 Then Exit Sub
  Call carga_excel
    
    
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Private Sub carga_excel()

    Dim estado As String
    Dim pedido As Integer
    
    vaSpread2.Row = vaSpread2.ActiveRow
    vaSpread2.Col = 1:  pedido = vaSpread2.text
    vaSpread2.Col = 3: estado = vaSpread2.text
    
    
   Sql = " sgpadm_sel_seleccionaDetallePedido "
   Sql = Sql & pedido & ",'S'"
   Set RS = vg_db.Execute(Sql)
    
   '-------> Inicio LLenar grilla
   vaSpread4.MaxRows = 0
 
    Do While Not RS.EOF
    
        vaSpread4.MaxRows = vaSpread4.MaxRows + 1
        vaSpread4.Row = vaSpread4.MaxRows
        
        vaSpread4.Col = 1 ' Ruta
        vaSpread4.text = Val(RS(0))
        vaSpread4.TypeHAlign = TypeHAlignCenter
        
        vaSpread4.Col = 2 ' Cod. Ingrediente
        vaSpread4.text = RS(1)
        vaSpread4.TypeHAlign = TypeHAlignCenter
        
        vaSpread4.Col = 3 ' Des. Ingrediente
        vaSpread4.text = RS(2)
        vaSpread4.TypeHAlign = TypeHAlignCenter
        
        vaSpread4.Col = 4 ' Probveedor
        vaSpread4.text = RS(3)
        vaSpread4.TypeHAlign = TypeHAlignCenter
        
        vaSpread4.Col = 5 ' familia
        vaSpread4.text = Val(RS(4))
        vaSpread4.TypeHAlign = TypeHAlignCenter
        
        vaSpread4.Col = 6 ' Cencos
        vaSpread4.text = Val(RS(5))
        vaSpread4.TypeHAlign = TypeHAlignCenter
        
        
        
        vaSpread4.Col = 7 ' Formato sap
        vaSpread4.text = RS(6)
        vaSpread4.TypeHAlign = TypeHAlignCenter
        
        vaSpread4.Col = 8 ' Material Sap
        vaSpread4.text = RS(7)
        vaSpread4.TypeHAlign = TypeHAlignCenter
        
        vaSpread4.Col = 9 ' Unidad
        vaSpread4.text = RS(8)
        vaSpread4.TypeHAlign = TypeHAlignCenter
        
        vaSpread4.Col = 10 ' total
        vaSpread4.text = Val(RS(9))
        vaSpread4.TypeHAlign = TypeHAlignCenter
        
        
        vaSpread4.Col = 11 ' Fecha despacho
        vaSpread4.text = Format(RS(10), "DD/MM/YYYY")
        vaSpread4.TypeHAlign = TypeHAlignCenter
        
        
        
        vaSpread4.Col = 12 ' Des. Ingrediente
        
        If RS(11) = False Then
            vaSpread4.text = "No Activo"
        Else
            vaSpread4.text = "Activo"
        
        End If
        
        vaSpread4.TypeHAlign = TypeHAlignCenter
        
        
        
        
        
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    
    
    
    If vaSpread4.MaxRows < 1 Then Exit Sub
    
    Dim X As Boolean
    vaSpread4.Row = -1
    vaSpread4.Col = -1
    vaSpread4.RowHidden = False
    
    
    ' Export Excel file and set result to x
    
    If Dir(dir_trabajo & "DetallePedido.XLS") <> "" Then Kill dir_trabajo & "DetallePedido.XLS"
    X = vaSpread4.ExportToExcel(dir_trabajo & "DetallePedido.XLS", "Test Sheet 1", dir_trabajo & "LOGFILE.TXT")
    ' Display result to user based on T/F value of x
    
    If X = True Then
        Dim XL As excel.Application
        Set XL = CreateObject("Excel.application")
        
        XL.Workbooks.Open FileName:=dir_trabajo & "DetallePedido.XLS"
        XL.Cells.Select ''-------> Desactivar proteción
        XL.ActiveSheet.Unprotect
        
 
        
        
        XL.Rows("1:1").Select '------> Insert Fila
        XL.Selection.Insert 'Shift:=xlDown
        XL.Range("A1").Select
        XL.ActiveCell.FormulaR1C1 = "Ruta"
        XL.Range("B1").Select
        XL.ActiveCell.FormulaR1C1 = "Código Ingrediente"
        XL.Range("C1").Select
        XL.ActiveCell.FormulaR1C1 = "Descripción"
        XL.Range("D1").Select
        XL.ActiveCell.FormulaR1C1 = "Proveedor"
        XL.Range("E1").Select
        XL.ActiveCell.FormulaR1C1 = "Familia"
        XL.Range("F1").Select
        XL.ActiveCell.FormulaR1C1 = "Cecos"
        XL.Range("G1").Select
        XL.ActiveCell.FormulaR1C1 = "Material Sap"
        XL.Range("H1").Select
        XL.ActiveCell.FormulaR1C1 = "Descripcion Material"
        XL.Range("I1").Select
        XL.ActiveCell.FormulaR1C1 = "Unidad Medida"
        XL.Range("J1").Select
        XL.ActiveCell.FormulaR1C1 = "Total"
        XL.Range("K1").Select
        XL.ActiveCell.FormulaR1C1 = "Fecha Despacho"
        XL.Range("L1").Select
        XL.ActiveCell.FormulaR1C1 = "Estado"
        XL.ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        XL.Visible = True '------->Visualizar
    Else
        MsgBox "Archivo esta abierto, grabe con otro nombre y luego cierre libro", , "Result"
    End If


End Sub

Private Sub generar_pedido()
    Dim MyBuffer As Variant
    Dim IdRuta As Long
    Dim CodIngrediente As String
    Dim CodProveedor As String
    Dim FamProducto As String
    Dim CenCosto As String
    Dim codproducto As String
    Dim FechaDespacho As String
    Dim total As Double
    Dim CodProductoSGP As String
    Dim CantidadIngrediente As Double
    Dim CantidadProducto As Double
    Dim pedido   As Integer
    Dim Linea    As Integer
    Dim Activo   As Integer
    Dim observacion   As String
    Dim proveedor   As String
    Dim familia   As String
    Dim descmaterial  As String
    Dim unidad        As String
    
    
    
    
    '-------> General Pedido & Minuta Real
    
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<GrabaDetallePedido>"

    For i = 1 To M_GenPed.vaSpread1.MaxRows
        Let MyBuffer = MyBuffer & " <DetallePedido"
        MyBuffer = MyBuffer & " Op = " & Chr(34) & 0 & Chr(34)
        desc = Replace(Trim(desc), Chr(34), "&quot;")
        desc = Replace(Trim(desc), Chr(38), "&amp;")
        desc = Replace(Trim(desc), Chr(39), "&apos;")
        desc = Replace(Trim(desc), Chr(60), "&lt;")
        desc = Replace(Trim(desc), Chr(62), "&gt;")

        M_GenPed.vaSpread1.Row = i
        
        M_GenPed.vaSpread1.Col = 1 'Id Ruta de Compras
        IdRuta = IIf(M_GenPed.vaSpread1.text = "", 0, M_GenPed.vaSpread1.text)

        M_GenPed.vaSpread1.Col = 2 'Código Ingrediente
        CodIngrediente = M_GenPed.vaSpread1.text

        M_GenPed.vaSpread1.Col = 3 'Descripción Ingrediente

        M_GenPed.vaSpread1.Col = 4 'Código Proveedor SAP
        CodProveedor = M_GenPed.vaSpread1.text

        M_GenPed.vaSpread1.Col = 5 'Código Familia Producto
        familia = M_GenPed.vaSpread1.text

        M_GenPed.vaSpread1.Col = 6 'Centro costo

        M_GenPed.vaSpread1.Col = 7 'Código Producto SAP
        codproducto = M_GenPed.vaSpread1.text
    

        M_GenPed.vaSpread1.Col = 8 'Descripción Producto
        descmaterial = M_GenPed.vaSpread1.text
        
        M_GenPed.vaSpread1.Col = 9 'Unidad
        unidad = M_GenPed.vaSpread1.text
        
        M_GenPed.vaSpread1.Col = 10 'Fecha Despacho
        FechaDespacho = IIf(M_GenPed.vaSpread1.text = "", "", Format(M_GenPed.vaSpread1.text, "yyyymmdd"))

        M_GenPed.vaSpread1.Col = 11 'Total
        total = M_GenPed.vaSpread1.text

        M_GenPed.vaSpread1.Col = 12 'Còdigo Producto SGP
        CodProductoSGP = M_GenPed.vaSpread1.text
        
        
        M_GenPed.vaSpread1.Col = 13 'Cantidad Ingrediente SGP
        CantidadIngrediente = M_GenPed.vaSpread1.text

      
        MyBuffer = MyBuffer & " pedido  = " & Chr(34) & 99999 & Chr(34)
        MyBuffer = MyBuffer & " linea  = " & Chr(34) & i & Chr(34)
        MyBuffer = MyBuffer & " CodIngrediente = " & Chr(34) & CodIngrediente & Chr(34)
        MyBuffer = MyBuffer & " CodProductoSGP  = " & Chr(34) & CodProductoSGP & Chr(34)
        MyBuffer = MyBuffer & " CodProducto  = " & Chr(34) & codproducto & Chr(34)
        MyBuffer = MyBuffer & " FechaDespacho  = " & Chr(34) & FechaDespacho & Chr(34)
       
       
       ' MyBuffer = MyBuffer & " FechaDespacho  = " & Chr(34) & FechaDespacho & Chr(34)
       
        MyBuffer = MyBuffer & " CantidadIngrediente  = " & Chr(34) & CantidadIngrediente & Chr(34)
        MyBuffer = MyBuffer & " CantidadProducto  = " & Chr(34) & CantidadProducto & Chr(34)
        MyBuffer = MyBuffer & " Total  = " & Chr(34) & total & Chr(34)
        MyBuffer = MyBuffer & " IdRuta  = " & Chr(34) & IdRuta & Chr(34)
     
       If IdRuta = 0 Or CodProductoSGP = "" Then
          MyBuffer = MyBuffer & " activo  = " & Chr(34) & 0 & Chr(34)
       Else
        MyBuffer = MyBuffer & " activo  = " & Chr(34) & 1 & Chr(34)
      End If
        MyBuffer = MyBuffer & " observacion  = " & Chr(34) & Null & Chr(34)
        MyBuffer = MyBuffer & " CodProveedor  = " & Chr(34) & CodProveedor & Chr(34)
        MyBuffer = MyBuffer & " familia  = " & Chr(34) & familia & Chr(34)
        MyBuffer = MyBuffer & " descmaterial  = " & Chr(34) & descmaterial & Chr(34)
        MyBuffer = MyBuffer & " unidad  = " & Chr(34) & unidad & Chr(34)
        
  
        
        Let MyBuffer = MyBuffer & "/>"

    Next i

    Let MyBuffer = MyBuffer & "</GrabaDetallePedido>"
    vg_db.Execute ("sgpadm_Ins_GrabaDetallePedidoNuevo '" & MyBuffer & "', '" & LimpiaDato(fpText.text) & "', " & Format(fpDateTime1(0).text, "yyyymmdd") & ", " & Format(fpDateTime1(1).text, "yyyymmdd") & "")

    MsgBox "Generación pedido finalizado sin problema", vbInformation + vbOKOnly, MsgTitulo
    Toolbar1.Buttons(3).Enabled = False
    Toolbar1.Enabled = True
    fg_descarga


Exit Sub

Man_Error:
Toolbar1.Enabled = True
Label3.Visible = False
Frame2.Visible = False
Label3.Caption = ""
vaSpread3.MaxRows = 0
Frame3.Visible = False

fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
End Sub
Private Sub busca_encabezado()
   
   Sql = " sgpadm_s_seleccionaPedidosEncabezado"
   Sql = Sql & " '" & fpText & "'"
   Set RS = vg_db.Execute(Sql)
    
   '-------> Inicio LLenar grilla
   Dim AuxCodIngrediente As String
   AuxIngrediente = ""
   vaSpread2.MaxRows = 0
 
    Do While Not RS.EOF
    
        vaSpread2.MaxRows = vaSpread2.MaxRows + 1
        vaSpread2.Row = vaSpread2.MaxRows
        
        vaSpread2.Col = 1 ' IdCompra
        vaSpread2.text = Val(RS(0))
        vaSpread2.TypeHAlign = TypeHAlignCenter
        
        vaSpread2.Col = 2 ' Cod. Ingrediente
        vaSpread2.text = RS(1)
        vaSpread2.TypeHAlign = TypeHAlignCenter
        
        vaSpread2.Col = 3 ' Des. Ingrediente
        vaSpread2.text = Format(RS(2), "DD/MM/YYYY")
        vaSpread2.TypeHAlign = TypeHAlignCenter
        
        vaSpread2.Col = 4 ' Des. Ingrediente
        vaSpread2.text = Format(RS(3), "DD/MM/YYYY")
        vaSpread2.TypeHAlign = TypeHAlignCenter
        
        vaSpread2.Col = 5 ' Des. Ingrediente
        vaSpread2.text = Format(RS(4), "DD/MM/YYYY")
        vaSpread2.TypeHAlign = TypeHAlignCenter
        
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing

End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Dim RS As New ADODB.Recordset
Dim Sql As String
Dim NomExcelZip As String
Dim i As Long
Dim NameTemp As String

Select Case Button.Index
Case 1
    '-------> Validar centro de costo
    Sql = Trim(LimpiaDato(fpText.text))
    Set RS = vg_db.Execute("sgpadm_s_cliente_V02 29, '" & Sql & "', ''")
    If Not RS.EOF Then
        If RS("cli_tipominuta") <> 3 Then
            MsgBox "Tipo de Minuta Contrato no Corresponde, Contrato Inactivo, pedido cancelado", vbExclamation + vbOKOnly, MsgTitulo
            fpayuda.Caption = ""
            Exit Sub
        End If
       RS.Close: Set RS = Nothing

    End If
    
    '-------> Validar fecha nulas
    
    If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then
       MsgBox "Fecha no corresponde, pedido cancelado", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    End If
    
    '-------> Validar si fecha final es menor inicial
    If fpDateTime1(1).text < fpDateTime1(0).text Then
       MsgBox "Fecha Inicial es mayor Final, pedido cancelado", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    End If
    '-------> Validar pedido bloqueado
    
    '-------> Validar si existe datos ruta carga
    'sql = Trim(LimpiaDato(fpText.text))
    'Set RS = vg_db.Execute("sgpadm_Sel_ValidarRutaDespacho '" & sql & "'")
    'If RS.EOF Then
    '    RS.Close: Set RS = Nothing
    '    MsgBox "No existe datos cargados rutas compras, pedido cancelado", vbExclamation + vbOKOnly, Msgtitulo
    '    Exit Sub
    'End If
    'RS.Close: Set RS = Nothing
    
    '-------> Validar si existe datos convenios
    'sql = Trim(LimpiaDato(fpText.text))
    'Set RS = vg_db.Execute("sgpadm_Sel_ValidarConvenios '" & sql & "'")
    'If RS.EOF Then
    '   RS.Close: Set RS = Nothing
    '   MsgBox "No existe datos cargados convenios, pedido cancelado", vbExclamation + vbOKOnly, Msgtitulo
    '   Exit Sub
    'End If
    'RS.Close: Set RS = Nothing

    '-------> Validar si existe minuta bloque
    
  ' FIN ARI
    
    Sql = Trim(LimpiaDato(fpText.text))
    Set RS = vg_db.Execute("sgpadm_Sel_ValidarMinutaBloqueACT '" & Sql & "', " & Format(fpDateTime1(0).text, "yyyymmdd") & ", " & Format(fpDateTime1(1).text, "yyyymmdd") & "")
   
    If RS.EOF Then
       RS.Close: Set RS = Nothing
       MsgBox "No existe minuta bloque, pedido cancelado", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    End If
    RS.Close: Set RS = Nothing
   
   'FIN ARI1
    
    
    Label3.Visible = True
    Frame2.Visible = True
    Label3.Caption = "Un momento generando la necesidad ..."
    
    vaSpread1.MaxRows = 0
    vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = &HC0FFFF
    estexi = True
   
   'dispara sp ppal.
   fg_carga ""
   Toolbar1.Enabled = False
   Sql = " sgpadm_Sel_GeneracionPedido_FDespacho_V09 "
   Sql = Sql & " '" & fpText & "'"
   Sql = Sql & " , " & Format(fpDateTime1(0).text, "yyyymmdd") & " "
   Sql = Sql & " , " & Format(fpDateTime1(1).text, "yyyymmdd") & ""

   Set RS = vg_db.Execute(Sql)
    
   '-------> Inicio LLenar grilla
   Dim AuxCodIngrediente As String
   AuxIngrediente = ""
   vaSpread1.MaxRows = 0
    If Not RS.EOF Then
      Toolbar1.Buttons(1).Visible = True
  
    Else
      Toolbar1.Buttons(1).Visible = False
    End If
    Do While Not RS.EOF
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
      '  If AuxCodIngrediente <> RS(1) Then
      '     For i = 1 To 12
      '         vaSpread1.Col = i
      '         vaSpread1.BackColor = Shape1(2).FillColor
      '     Next i
      '     AuxCodIngrediente = RS(1)
      '  End If
        vaSpread1.Col = 1 ' IdCompra
        vaSpread1.text = IIf(RS(0) = 0, "", RS(0))
        
        vaSpread1.Col = 2 ' Cod. Ingrediente
        vaSpread1.text = RS(1)
        vaSpread1.Col = 3 ' Des. Ingrediente
        vaSpread1.text = RS(2)
        vaSpread1.Col = 4 ' Proveedor
        vaSpread1.text = RS(3)
        vaSpread1.Col = 5 ' Familia Producto
        vaSpread1.text = RS(4)
        vaSpread1.Col = 6 ' Centro Costo
        vaSpread1.text = RS(5)
        vaSpread1.Col = 7 ' Codigo Producto SAP
        vaSpread1.text = RS(6)
        vaSpread1.Col = 8 ' Des. Producto SAp
        vaSpread1.text = RS(7)
        vaSpread1.Col = 9 ' Unidad
        vaSpread1.text = RS(8)
        vaSpread1.Col = 10 ' Fecha Despacho
        vaSpread1.text = RS(9)
        vaSpread1.Col = 11 ' Cantidad Solicitar
        vaSpread1.text = RS(10)
        vaSpread1.Col = 12 'Código Productos
        vaSpread1.text = RS(11)
        vaSpread1.Col = 13 ' Cantidad Ingrediente
        vaSpread1.text = RS(12)
        vaSpread1.Col = 14 ' Cantidad Producto
        vaSpread1.text = RS(13)
        
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    
   
   
    Label3.Visible = False
    Frame2.Visible = False
    Label3.Caption = ""
    
    '-------> validar si existe pedido
    
    If vaSpread1.MaxRows < 1 Then
       fg_descarga
       Label3.Visible = False
       Frame2.Visible = False
       Label3.Caption = ""
'       DropTeblaTmp (NameTemp)
       MsgBox "Por favor verificar si existen " & VgLinea & VgLinea & "- Rutas para la fecha consultada " & VgLinea & "- Convenios vigentes para la fecha consultada " & VgLinea, vbInformation + vbOKOnly, MsgTitulo
       Toolbar1.Enabled = True
       Exit Sub
    End If
       Toolbar1.Enabled = True
       
       fg_descarga
End Select
Exit Sub
End Sub

