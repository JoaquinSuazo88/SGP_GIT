VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_Pedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación Pedido Mensual"
   ClientHeight    =   8700
   ClientLeft      =   3015
   ClientTop       =   2130
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   7800
      TabIndex        =   11
      Top             =   7920
      Width           =   4170
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   150
         TabIndex        =   12
         Top             =   180
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   688
         ButtonWidth     =   1138
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1425
      Index           =   1
      Left            =   2490
      TabIndex        =   3
      Top             =   120
      Width           =   7095
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1380
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
         Left            =   1380
         TabIndex        =   1
         Top             =   630
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1658
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
         Text            =   "10/2016"
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
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   360
         Left            =   2520
         TabIndex        =   2
         Top             =   600
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
      Begin EditLib.fpDateTime fpDateTime2 
         Height          =   315
         Left            =   1380
         TabIndex        =   14
         Top             =   980
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1658
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
         Text            =   "10/2016"
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   3840
         TabIndex        =   16
         Top             =   980
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
         Caption         =   "Nro. Semana"
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
         Left            =   2520
         TabIndex        =   15
         Top             =   1080
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo SAC"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha(mm/aa)"
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
         Top             =   690
         Width           =   1230
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
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2610
         Picture         =   "M_Pedido.frx":0000
         Top             =   195
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3045
         TabIndex        =   4
         Top             =   285
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3090
         TabIndex        =   7
         Top             =   330
         Width           =   3855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8700
      Left            =   12285
      TabIndex        =   8
      Top             =   0
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   15346
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   6075
      Left            =   30
      TabIndex        =   9
      Top             =   1710
      Width           =   12195
      _Version        =   393216
      _ExtentX        =   21511
      _ExtentY        =   10716
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
      MaxCols         =   17
      MaxRows         =   1
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "M_Pedido.frx":030A
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _Version        =   393216
      _ExtentX        =   2143
      _ExtentY        =   1085
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
      MaxRows         =   1
      SpreadDesigner  =   "M_Pedido.frx":0BF6
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
            Picture         =   "M_Pedido.frx":0F04
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim Fecha As Long
Dim Msgtitulo As String
Dim est As Boolean, etapa5 As Boolean, aAp1 As String, aAp2 As String
Dim estexi As Boolean
Dim vecdes() As Variant

Private Sub Form_Activate()
fg_descarga
'-------> Traer fecha cierre día
 TraerFechaCierre
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
fg_carga ""
Me.Height = 9165
Me.Width = 12885
fg_centra Me
est = True
Msgtitulo = "Generación pedidos Mensual"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.Enabled = False: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): BtnX.Visible = True: BtnX.Enabled = False: BtnX.ToolTipText = "Enviar"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = "Borrar "
Set BtnX = Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = ""
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Enabled = False: BtnX.ToolTipText = "Imprimir Pedidos mensual"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Imprimir1 ", , tbrDefault, "A_Imprimir "): BtnX.Visible = IIf(vg_pais = "CO", True, False): BtnX.Enabled = True: BtnX.ToolTipText = "Imprimir Productos no Asociado SAC"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
Toolbar2.ImageList = Partida.IL1
Set BtnX = Toolbar2.Buttons.Add(, "", , tbrDefault, 0): BtnX.Visible = True: BtnX.Enabled = True: BtnX.Caption = "Agregar Producto ": BtnX.ToolTipText = "Agregar Producto "
If vg_pais = "CO" Then Set BtnX = Toolbar2.Buttons.Add(, "", , tbrDefault, 0): BtnX.Visible = True: BtnX.Enabled = True: BtnX.Caption = "Cambiar Producto ": BtnX.ToolTipText = "Cambiar Producto"
'Set btnX = Toolbar2.Buttons.Add(, "A_EliminarF", , tbrDefault, "A_EliminarF"): btnX.Visible = True: btnX.Enabled = True: btnX.Caption = "Eliminar Producto ": btnX.ToolTipText = "Eliminar Producto"
Toolbar2.Enabled = False
vaSpread1.MaxRows = 0
fpDateTime1.text = Format(Date, "mm/yyyy")
fpDateTime2.text = Format(Date, "mm/yyyy")
fpLongInteger1(0).Value = ""
fpText.Enabled = ModCasino
Image1.Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda.Caption = MuestraCasino(2)
If vg_pais = "CL" Then
   Label1(1).Visible = False
   Label1(2).Visible = False
   fpDateTime2.Visible = False
   fpLongInteger1(0).Visible = False
End If
Dim X As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
'X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFC0, &H800000)
est = False
If vg_pais = "CO" Then
   MarcaPredeterminadoFormatoCompras
End If
fg_descarga
End Sub

Sub CalcularStockSACRECMINDIA()
Dim RS As New ADODB.Recordset
Dim sql1 As String, sql2 As String, sql3 As String, sql4 As String
Dim fechoy As Date, fecaux As Date, fecini As Date, fecfin As Date
Dim fecval As Long, fecpin As Long, fecpfi As Long, fecxin As Long, fecxfi As Long
Dim aAp As String
Dim estfij As Boolean

   '-------> Validar productos vigentes toma de pedido
   ValidarProductoVigente
   '-------> Traer stock actual
   vaSpread2.MaxRows = 0
   RS.Open "SELECT DISTINCT bod_codpro, SUM(bod_canmer) AS bod_canmer FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " GROUP BY bod_codpro", vg_db, adOpenStatic
   If Not RS.EOF Then
      Do While Not RS.EOF
         vaSpread2.MaxRows = vaSpread2.MaxRows + 1
         vaSpread2.Row = vaSpread2.MaxRows
         vaSpread2.Col = 1: vaSpread2.text = RS!bod_codpro
         vaSpread2.Col = 2: vaSpread2.text = RS!bod_canmer
         vaSpread2.Col = 3: vaSpread2.text = 0
         vaSpread2.Col = 4: vaSpread2.text = 0
         vaSpread2.Col = 5: vaSpread2.text = 0
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
   '-------> Fin traer stock actual
   '-------> Traer ordernes de compras x recibir
   sql1 = IIf(vg_tipbase = "1", " SUM(IIF(a.tipsol_idsol = 4,(-1 * a.pedite_qtcpa), a.pedite_qtcpa)) AS CanEnt ", " SUM(CASE WHEN a.tipsol_idsol = 4 THEN (-1 * a.pedite_qtcpa) ELSE a.pedite_qtcpa END) AS CanEnt ")
   sql2 = IIf(vg_tipbase = "1", " '" & Fecha & "' ", " '" & Fecha & "' ")
   sql3 = IIf(vg_tipbase = "1", " ,(SELECT DISTINCT bod_canmer FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " AND bod_codpro = b.pro_codigo) as bod_canmer ", " ,(SELECT DISTINCT bod_canmer FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " AND bod_codpro = b.pro_codigo) as bod_canmer ")
   sql4 = IIf(vg_tipbase = "1", " Format(a.solite_dtent, 'yyyymm') ", " convert(varchar(6),solite_dtent,112) ")
   RS.Open "SELECT DISTINCT c.foc_codcat, a.cpopro_cdpro, a.cadfor_nrcgc, b.pro_codigo, b.pro_nombre, b.pro_ctrsto, b.pro_ctacon, b.pro_propon, " & _
           "f.uni_nomcor, c.foc_nomsac, a.solite_dtent, a.pedite_vlpco, " & sql1 & " " & _
           "" & sql3 & " " & _
           "FROM b_ocsac a, b_productos b, b_formatocompras c, b_formatocomprassgp d, a_unidad f " & _
           "Where c.foc_codsac   = d.fcs_codsac " & _
           "AND   a.cpopro_cdpro = c.foc_codsac " & _
           "AND   b.pro_codigo   = d.fcs_codsgp " & _
           "AND   b.pro_coduni   = f.uni_codigo " & _
           "AND   a.cadfil_cdfil = '" & MuestraCasino(1) & "' " & _
           "AND   " & sql4 & "   = " & sql2 & " " & _
           "AND   a.pedite_flafo IN (0,1) " & _
           "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven = 0) " & _
           "GROUP BY c.foc_codcat, a.cpopro_cdpro, a.cadfor_nrcgc, b.pro_codigo, b.pro_nombre, b.pro_ctrsto, b.pro_ctacon, b.pro_propon, f.uni_nomcor, c.foc_nomsac, a.cpopro_cdpro, a.solite_dtent, a.pedite_vlpco ORDER BY c.foc_codcat, a.solite_dtent, a.cadfor_nrcgc, c.foc_nomsac", vg_db, adOpenStatic
   If Not RS.EOF Then
      sql1 = IIf(vg_tipbase = "1", " cdate('" & RS!solite_dtent & "') ", " '" & Format(RS!solite_dtent, "yyyymmdd") & "' ")
   End If
   Do While Not RS.EOF
      RS1.Open "SELECT SUM(ocr_cancom) AS difer " & _
               "FROM   b_ocsacrecibido " & _
               "WHERE  ocr_rutpro = '" & RutPro & "' " & _
               "AND    ocr_fecoc  = " & sql1 & " " & _
               "AND    ocr_codprodsac = '" & Trim(RS!cpopro_cdpro) & "' AND ocr_codprodsgp = '" & Trim(RS!pro_codigo) & "'", vg_db, adOpenStatic
      If Not RS1.EOF And ((RS1!difer - RS!canent) <> 0 Or IsNull(RS1!difer - RS!canent)) And RS!canent > 0 Then
         If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone) <> -1 Then
            vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone)
            vaSpread2.Col = 5
            If Not IsNull(RS1!difer - RS!canent) Then
               vaSpread2.text = (vaSpread2.text + (RS!canent - RS1!difer))
            Else
               vaSpread2.text = (vaSpread2.text + RS!canent)
            End If
            If Val(vaSpread2.text) < 0 Then vaSpread2.text = 0
         End If
      
      End If
      RS1.Close: Set RS1 = Nothing
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
   '-------> Fin traer ordernes de compras x recibir
   fechoy = CDate(Date) 'amorgadoIIf(Fecha = Format(CDate(Date), "yyyymm"), fg_pone_cero(Str(Day(Now)), 2) & "/" & Format(BoM(CDate(Date)), "mm/yyyy"), CDate(Date))
   '-------> Traer consumo a la fecha
   aAp = Trim(vg_NUsr) & "_tmp_PedMensualFecha"
   '-------> Creo tabla temporal y chequeo si existe antes
   fg_CheckTmp aAp
   RS.Open "SELECT e.min_fecmin, a.pro_codtip, a.pro_codigo, SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))) AS cantidad1, " & _
           "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))/a.pro_facing) AS cantidad2, 0 AS fecped INTO " & aAp & " " & _
           "FROM   b_productos a, b_contlistpreing b, b_receta c, b_recetadet d, b_minuta e, b_minutadet f, a_servicio h " & _
           "WHERE  e.min_codigo = f.mid_codigo " & _
           "AND    f.mid_codrec = d.red_codigo " & _
           "AND    f.mid_tiprec = d.red_tiprec AND ((d.red_tiprec<>0 AND d.red_cencos = '" & MuestraCasino(1) & "') OR (d.red_tiprec = 0 AND d.red_cencos = '0')) " & _
           "AND    d.red_codigo = c.rec_codigo " & _
           "AND    d.red_codpro = b.cpi_coding AND b.cpi_cencos = '" & MuestraCasino(1) & "' " & _
           "AND    b.cpi_codped = a.pro_codigo " & _
           "AND    e.min_cencos = '" & Trim(fpText.text) & "' " & _
           "AND    e.min_fecmin >= " & Format(fechoy, "yyyymmdd") & " " & _
           "AND    e.min_fecmin <= " & Val(Format(dEoM("27/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)), "yyyymmdd")) & " " & _
           "AND    f.mid_tipmin = '2' and e.min_codser = h.ser_codigo and h.ser_activo = '1' " & _
           "AND    a.pro_facing > 0 AND (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) GROUP BY e.min_fecmin, a.pro_codtip, a.pro_codigo", vg_db, adOpenStatic
   Set RS = Nothing
   '-------> Rutina buscar estructura fija
   RS.Open "SELECT DISTINCT a.mif_codreg, a.mif_codser FROM b_minutafija a, a_servicio b WHERE a.mif_codser = b.ser_codigo and b.ser_activo = '1' and a.mif_cencos = '" & fpText.text & "'", vg_db, adOpenStatic
   If Not RS.EOF Then
      Do While Not RS.EOF
         estfij = False
         '-------> Buscar datos estructura fija día
         RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
                  "WHERE mfd_cencos = '" & Trim(fpText.text) & "' " & _
                  "AND   mfd_codreg = " & RS!mif_codreg & " " & _
                  "AND   mfd_codser = " & RS!mif_codser & " " & _
                  "AND   mfd_fecha >= " & Format(fechoy, "yyyymmdd") & " " & _
                  "AND   mfd_fecha <= " & Val(Format(dEoM("27/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)), "yyyymmdd")) & " AND mfd_tipmin = '2'", vg_db, adOpenForwardOnly
         If Not RS1.EOF Then estfij = True
         RS1.Close: Set RS1 = Nothing
         fecval = 0
         If Not estfij Then
            '-------> Buscar fecha mayor de estructura fija
            RS1.Open "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija WHERE mif_cencos = '" & Trim(fpText.text) & "' AND mif_codreg = " & RS!mif_codreg & " AND mif_codser = " & RS!mif_codser & "", vg_db, adOpenForwardOnly
            If Not RS1.EOF Then fecval = IIf(IsNull(RS1!fecval), 0, RS1!fecval)
            RS1.Close: Set RS1 = Nothing
            If fecval > 0 Then
               '-------> Calcular datos desde tabla estructura fija
                fecaux = fechoy
                Do While fecaux <= dEoM(dEoM("27/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + 1)
                   If CDate(fg_Ctod1(fecval)) <= fecaux Then
                      vg_db.Execute "INSERT INTO " & aAp & " SELECT " & Format(fecaux, "yyyymmdd") & " AS min_fecmin, b.pro_codtip, b.pro_codigo AS pro_codigo, 0 AS cantidad1, " & _
                                    "a.mif_canpro AS cantidad2, 0 AS fecped FROM b_minutafija a, b_productos b " & _
                                    "WHERE a.mif_codpro = b.pro_codigo " & _
                                    "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                    "AND   a.mif_cencos = '" & Trim(fpText.text) & "' " & _
                                    "AND   a.mif_codreg = " & RS!mif_codreg & " " & _
                                    "AND   a.mif_codser = " & RS!mif_codser & " " & _
                                    "AND   a.mif_fecval = " & fecval & " " & _
                                    "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(Year(fecaux) & fg_pone_cero(Str(Month(fecaux)), 2) & fg_pone_cero(Str(Day(fecaux)), 2), 2), Len(fg_Fecha_Dia(Year(fecaux) & fg_pone_cero(Str(Month(fecaux)), 2) & fg_pone_cero(Str(Day(fecaux)), 2), 2)) - 2))) & " " & _
                                    "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0)"
                   End If
                   Set RS1 = Nothing
                   fecaux = fecaux + 1
               Loop
            End If
         ElseIf estfij Then
            '-------> Calcular datos desde tabla estructura fija día
            vg_db.Execute "INSERT INTO " & aAp & " SELECT a.mfd_fecha AS min_fecmin, b.pro_codtip, b.pro_codigo AS pro_codigo, 0 AS cantidad1, " & _
                          "a.mfd_canpro AS cantidad2, 0 AS fecped FROM b_minutafijadia a, b_productos b " & _
                          "WHERE a.mfd_codpro = b.pro_codigo " & _
                          "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                          "AND   a.mfd_codreg = " & RS!mif_codreg & " " & _
                          "AND   a.mfd_codser = " & RS!mif_codser & " " & _
                          "AND   a.mfd_fecha >= " & Format(fechoy, "yyyymmdd") & " " & _
                          "AND   a.mfd_fecha <= " & Val(Format(dEoM(dEoM("27/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + 1), "yyyymmdd")) & " " & _
                          "AND   a.mfd_tipmin = '2' " & _
                          "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0)"
            Set RS1 = Nothing
         End If
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
   '-------> Calcular días despachos
   '-------> actualizar fecha pedido mensual
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'M')", vg_db, adOpenForwardOnly
   If Not RS.EOF Then
      fecxin = Format(fechoy + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("27/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.fecped = 1 WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'M' AND c.pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
         vg_db.Execute "UPDATE " & aAp & "  " & _
                       "SET " & aAp & ".fecped = substring(convert(varchar(8),a.min_fecmin),1,6) + '01' " & _
                       "FROM " & aAp & "  a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'M' AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
   End If
   RS.Close: Set RS = Nothing
   '-------> actualizar fecha pedido quincenal 1-15
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'Q1')", vg_db, adOpenForwardOnly
   If Not RS.EOF Then
      fecxin = Format(fechoy + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate("15/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.fecped = 1 WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q1' AND c.pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".fecped = 1 " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q1' AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
      fecxin = Format(CDate("16/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("27/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.fecped = 1 WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q1' AND c.pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".fecped = 1 " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q1' AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
   End If
   RS.Close: Set RS = Nothing
   '-------> actualizar fecha pedido quincenal 2-16
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'Q2')", vg_db, adOpenForwardOnly
   If Not RS.EOF Then
      fecxin = Format(CDate("02/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate("16/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.fecped = 1 WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q2' AND c.pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".fecped = 1 " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q2' AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
      fecxin = Format(CDate("17/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("27/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.fecped = 1 WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q2' AND c.pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".fecped = 1 " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q2' AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
   End If
   RS.Close: Set RS = Nothing
   '-------> actualizar fecha pedido quincenal 3-17
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'Q3')", vg_db, adOpenForwardOnly
   If Not RS.EOF Then
      fecxin = Format(CDate("03/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate("27/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.fecped = 1 WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q3' AND c.pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".fecped = 1 " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q3' AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
      fecxin = Format(CDate("18/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("17/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.fecped = 1 WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q3' AND c.pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".fecped = 1 " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q3' AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
   End If
   RS.Close: Set RS = Nothing
   '-------> actualizar fecha pedido quincenal 4-18
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'Q4')", vg_db, adOpenForwardOnly
   If Not RS.EOF Then
      fecxin = Format(CDate("04/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate("18/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.fecped = 1 WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q4' AND c.pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".fecped = 1 " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q4' AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
      fecxin = Format(CDate("19/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("18/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.fecped = 1 WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q4' AND c.pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecpin & " AND a.min_fecmin <= " & fecpfi & ""
      Else
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".fecped = 1 " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q4' AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
   End If
   RS.Close: Set RS = Nothing
   '-------> actualizar fecha pedido cada 10 días.
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'D')", vg_db, adOpenForwardOnly
   If Not RS.EOF Then
      fecxin = Format(CDate("01/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate("10/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.ped_fecped = 1 WHERE c.cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'D' AND c.pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".fecped = 1 " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'D' AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
      fecxin = Format(CDate("11/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("20/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.fecped = 1 WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'D' AND c.pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecpin & " AND a.min_fecmin <= " & fecpfi & ""
      Else
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".fecped = 1 " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'D' AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
      fecxin = Format(CDate("21/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("27/" & Mid(fechoy, 4, 2) & "/" & Mid(fechoy, 7, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.fecped = 1 WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'D' AND c.pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecpin & " AND a.min_fecmin <= " & fecpfi & ""
      Else
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".fecped = 1 " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'D' AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
   End If
   RS.Close: Set RS = Nothing
   '-------> actualizar fecha diario y semanal
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'E' OR pad_tipo = 'S')", vg_db, adOpenForwardOnly
   Do While Not RS.EOF
      fecini = fechoy
      fecfin = "07/" & IIf(Mid(fechoy, 4, 2) = 12, "01/" & Mid(fechoy, 7, 4) + 1, Mid(fechoy, 4, 2) + 1 & "/" & Mid(fechoy, 7, 4))
      fecpin = 0: fecpfi = 0
      Do While fecini <= fecfin
         '-------> Buscar fecha inicial y fecha final
         For j = 1 To 7
             If (DatePart("w", fecini, 2)) = Val(Mid(RS!pad_diario, j, 1)) Then
                If fecpin = 0 Then
                   fecpin = Format(fecini, "yyyymmdd")
                ElseIf fecpfi = 0 Then
                   fecpfi = Format(fecini, "yyyymmdd")
                End If
             End If
             If fecpin > 0 And fecpfi > 0 Then
                fecxin = Format(CDate(fg_Ctod1(fecpin)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
                fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
                fecxfi = Format(CDate(fg_Ctod1(fecpfi)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
                fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
                If vg_tipbase = "1" Then
                   vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                                 "SET a.fecped = 1 WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (c.pad_tipo = 'E' OR c.pad_tipo = 'S') AND c.pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
                Else
                   vg_db.Execute "UPDATE " & aAp & " SET " & aAp & ".fecped = 1 FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d WHERE a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo " & _
                                 "AND c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'E' OR pad_tipo = 'S')AND pad_codigo = " & RS!pad_codigo & " AND a.fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
                End If
                fecpin = fecpfi: fecpfi = 0
                Exit For
             End If
         Next j
         fecini = fecini + 1
      Loop
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
   '-------> Leer archivo temporales
   vg_db.Execute ("DELETE " & aAp & " FROM  " & aAp & " WHERE fecped = 0")
   '-------> Fin calcular días despachos
   RS.Open "SELECT b.pro_codigo, b.pro_facsto, SUM(a.cantidad1) AS cantidad1, SUM(a.cantidad2) AS cantidad2 " & _
           "FROM " & aAp & " a, b_productos b " & _
           "WHERE a.pro_codigo = b.pro_codigo " & _
           "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
           "GROUP BY b.pro_codigo, b.pro_facsto", vg_db, adOpenForwardOnly
   If Not RS.EOF Then
      Do While Not RS.EOF
         If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone) <> -1 Then
            vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone)
            vaSpread2.Col = 4
            If Not IsNull(RS!cantidad2) Then
               vaSpread2.text = (vaSpread2.text + IIf(Int(RS!cantidad2 / RS!pro_facsto) <> (RS!cantidad2 / RS!pro_facsto), Int(RS!cantidad2 / RS!pro_facsto) + 1, Round(RS!cantidad2 / RS!pro_facsto, 0)) * RS!pro_facsto)
            Else
               vaSpread2.text = (vaSpread2.text - 0)
            End If
            If Val(vaSpread2.text) < 0 Then vaSpread2.text = 0
         End If
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
   '-------> Fin traer consumo a la fecha
End Sub

Private Sub MoverDatos()
Dim codtip As Long, fecenv As Long, fecval As Long, diaini As Long, diafin As Long, i As Long, j As Long, X As Long
Dim aAp As String, proc1 As String, proc2 As String, proc3 As String
Dim canped As Double, proped As Double, estfij As Boolean
Dim sql1 As String, sql2 As String, sql3 As String, sql4 As String, sql5 As String, sql6 As String
Dim fecxin As Long, fecxfi As Long
Dim fechoy As Date, fecaux As Date
Dim fecini As Date, fecfin As Date, fecped As Date, EstPed As Integer, fecpin As Long, fecpfi As Long
fg_carga ""
vaSpread1.MaxRows = 0
Toolbar2.Enabled = False
vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = &HC0FFFF
Fecha = 0: codtip = 0: fecenv = 0: auxing = 0
Fecha = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)
estexi = True
'-------> Cargar despacho vector
RS.Open "SELECT DISTINCT pad_tipo FROM b_paramdesp WHERE pad_cencos='" & LimpiaDato(Trim(fpText.text)) & "'", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe parametros de despacho, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
ReDim vecdes(RS.RecordCount, 5)
i = 1
Do While Not RS.EOF
   vecdes(i, 1) = Trim(RS!pad_tipo)
   vecdes(i, 2) = "01/" & fpDateTime1.text
   vecdes(i, 3) = IIf(Trim(RS!pad_tipo) = "S", "08/" & fpDateTime1.text, IIf(Trim(RS!pad_tipo) = "D", "11/" & fpDateTime1.text, ""))
   vecdes(i, 4) = IIf(Trim(RS!pad_tipo) = "S" Or Trim(RS!pad_tipo) = "Q", "15/" & fpDateTime1.text, IIf(Trim(RS!pad_tipo) = "D", "21/" & fpDateTime1.text, ""))
   vecdes(i, 5) = IIf(Trim(RS!pad_tipo) = "S", "22/" & fpDateTime1.text, "")
   If RS!pad_tipo = "E" Then
      RS1.Open "SELECT DISTINCT pad_tipo FROM b_paramdesp WHERE pad_cencos='" & LimpiaDato(Trim(fpText.text)) & "' AND pad_tipo='E' AND (pad_diario='' or pad_diario='0000000' or (pad_diario) IS NULL)", vg_db, adOpenStatic
      If Not RS1.EOF Then fg_descarga: RS.Close: Set RS = Nothing: RS1.Close: Set RS1 = Nothing: MsgBox "Falta definir los parametros despachos díarios, en una familia productos, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
      RS1.Close: Set RS1 = Nothing:
   End If
   RS.MoveNext: i = i + 1
Loop
RS.Close: Set RS = Nothing
aAp1 = Trim(vg_NUsr) & "_tmp_paramdesp"
'-------> Creo tabla temporal y chequeo si existe antes
fg_CheckTmp aAp1
'------->
vg_db.Execute "SELECT DISTINCT pro_codtip, 0 AS pro_previo INTO " & aAp1 & " FROM b_productos"
If vg_tipbase = "1" Then
   aAp2 = Trim(vg_NUsr) & "_tmp_productospmpdiaGenPed"
   '-------> Creo tabla temporal y b_productospmpdia
   fg_CheckTmp aAp2
   vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                 "INTO " & aAp2 & " " & _
                 "FROM b_productospmpdia " & _
                 "WHERE ppd_cencos='" & MuestraCasino(1) & "' " & _
                 "AND   ppd_fecdia>=" & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia<=" & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                 "AND   ppd_propon>0 " & _
                 "GROUP BY ppd_cencos, ppd_codpro"
   vg_db.Execute "ALTER TABLE " & aAp2 & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
   vg_db.Execute "UPDATE " & aAp2 & " INNER JOIN b_productospmpdia ON (" & aAp2 & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp2 & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp2 & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp2 & ".ppd_propon=b_productospmpdia.ppd_propon"
   vg_db.Execute "INSERT INTO " & aAp2 & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp2 & ")"
End If
RS.Open "SELECT * FROM " & aAp1 & "", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      vg_db.Execute "UPDATE " & aAp1 & " SET pro_previo = " & IIf(fg_BuscaenArbolNivel2(RS!pro_codtip, "a_tipopro", "tip_codigo") = 0, RS!pro_codtip, fg_BuscaenArbolNivel2(RS!pro_codtip, "a_tipopro", "tip_codigo")) & "  WHERE pro_codtip = " & RS!pro_codtip & ""
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
'-------> Validar si existe pedidos mensual
sql1 = IIf(vg_tipbase = "1", " val(mid(a.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) ")
'-------> Validar si existe pedido generado
RS.Open "SELECT DISTINCT ped_anomes, ped_persac, ped_semsac FROM b_minutapedidos WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " AND ped_tipped = 1", vg_db, adOpenStatic
If Not RS.EOF Then
   fpDateTime2.text = IIf(IsNull(RS!ped_persac), Format(Date, "mm/yyyy"), Mid(RS!ped_persac, 5, 2) & "/" & Mid(RS!ped_persac, 1, 4))
   fpLongInteger1(0).Value = IIf(IsNull(RS!ped_semsac), 0, RS!ped_semsac)
   RS.Close: Set RS = Nothing
   '-------> Validar si pedido fue enviado
   RS.Open "SELECT DISTINCT b.mid_fecval FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND " & sql1 & " = " & Val(Fecha) & " AND b.mid_fecval > 0", vg_db, adOpenStatic
   If Not RS.EOF Then fecenv = RS!mid_fecval Else fecenv = 0
   RS.Close: Set RS = Nothing
   If fecenv > 0 Then
      Toolbar1.Buttons(1).Enabled = False
      Toolbar1.Buttons(3).Enabled = False
      If Not CierrePeriodo(Fecha, vg_codbod, 10) Then
         Toolbar1.Buttons(5).Visible = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
         Toolbar1.Buttons(6).Visible = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
      Else
         Toolbar1.Buttons(5).Visible = False
         Toolbar1.Buttons(6).Visible = True
      End If
      Toolbar1.Buttons(8).Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "0", False, True)
   Else
      CalcularStockSACRECMINDIA
      Toolbar1.Buttons(1).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
      Toolbar1.Buttons(3).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
      Toolbar1.Buttons(5).Visible = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
      Toolbar1.Buttons(6).Visible = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
      Toolbar1.Buttons(8).Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "0", False, True) 'False
   End If
   proc1 = "": proc2 = "": proc3 = ""
   MoverDatosGrilla proc1, proc2, proc3
   estexi = IIf(fecenv = 0, False, True)
Else
   RS.Close: Set RS = Nothing
   estexi = False
   CalcularStockSACRECMINDIA
   '-------> Traer consumo minuta actual
   fecenv = 1
   aAp = Trim(vg_NUsr) & "_tmp_PedMensual"
   '-------> Creo tabla temporal y chequeo si existe antes
   fg_CheckTmp aAp
   If vg_tipbase = "1" Then
      RS.Open "SELECT e.min_fecmin, b.cpi_coding AS ing_codigo, a.pro_codigo, a.pro_codtip, " & _
              "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))) AS cantidad1, " & _
              "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))/a.pro_facing) AS cantidad2, 0 AS ped_fecped INTO " & aAp & " " & _
              "FROM b_productos a, b_contlistpreing b, b_receta c, b_recetadet d, b_minuta e, b_minutadet f, a_servicio h " & _
              "WHERE e.min_codigo = f.mid_codigo AND f.mid_codrec = d.red_codigo AND f.mid_tiprec = d.red_tiprec AND ((d.red_tiprec <> 0 AND d.red_cencos = '" & MuestraCasino(1) & "') OR (d.red_tiprec = 0 AND d.red_cencos = '0')) " & _
              "AND   d.red_codigo = c.rec_codigo AND d.red_codpro = b.cpi_coding AND b.cpi_cencos = '" & MuestraCasino(1) & "' AND b.cpi_codped = a.pro_codigo " & _
              "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) " & _
              "AND   e.min_cencos = '" & Trim(fpText.text) & "' AND val(mid(e.min_fecmin,1,6)) = " & Fecha & " " & _
              "AND   f.mid_tipmin = '1' AND a.pro_facing > 0 and e.min_codser = h.ser_codigo and h.ser_activo = '1' GROUP BY e.min_fecmin, b.cpi_coding, a.pro_codigo, a.pro_codtip", vg_db, adOpenForwardOnly
   Else
      RS.Open "SELECT e.min_fecmin, b.cpi_coding AS ing_codigo, a.pro_codigo, a.pro_codtip, " & _
              "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))) AS cantidad1, " & _
              "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))/a.pro_facing) AS cantidad2, 0 AS ped_fecped INTO " & aAp & " " & _
              "FROM b_productos a, b_contlistpreing b, b_receta c, b_recetadet d, b_minuta e, b_minutadet f, a_servicio h " & _
              "WHERE e.min_codigo = f.mid_codigo AND f.mid_codrec = d.red_codigo AND f.mid_tiprec = d.red_tiprec AND ((d.red_tiprec <> 0 AND d.red_cencos = '" & MuestraCasino(1) & "') OR (d.red_tiprec = 0 AND d.red_cencos = '0')) " & _
              "AND   d.red_codigo = c.rec_codigo AND d.red_codpro = b.cpi_coding AND b.cpi_cencos = '" & MuestraCasino(1) & "' AND b.cpi_codped = a.pro_codigo " & _
              "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) " & _
              "AND   e.min_cencos = '" & Trim(fpText.text) & "' AND convert(int,substring(convert(varchar(8),e.min_fecmin),1,6)) = " & Fecha & " " & _
              "AND   f.mid_tipmin = '1' AND a.pro_facing > 0 and e.min_codser = h.ser_codigo and h.ser_activo = '1' GROUP BY e.min_fecmin, b.cpi_coding, a.pro_codigo, a.pro_codtip", vg_db, adOpenForwardOnly
   End If
   Set RS = Nothing
   diaini = 0: diafin = 0
   diaini = 1: diafin = Mid(dEoM("01/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)), 1, 2)
   '-------> Rutina buscar estructura fija
   RS.Open "SELECT DISTINCT a.mif_codreg, a.mif_codser FROM  b_minutafija a, a_servicio b WHERE a.mif_cencos='" & Trim(fpText.text) & "' and a.mif_codser = b.ser_codigo and b.ser_activo ='1'", vg_db, adOpenStatic
   If Not RS.EOF Then
      Do While Not RS.EOF
         estfij = False
         '-------> Buscar datos estructura fija día
         If vg_tipbase = "1" Then
            RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
                     "WHERE mfd_cencos='" & Trim(fpText.text) & "' " & _
                     "AND   mfd_codreg=" & RS!mif_codreg & " " & _
                     "AND   mfd_codser=" & RS!mif_codser & " " & _
                     "AND mid(mfd_fecha,1,6)>=" & Fecha & " AND mid(mfd_fecha,1,6)<=" & Fecha & " AND mfd_tipmin='1'", vg_db, adOpenStatic
         Else
            RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
                     "WHERE mfd_cencos='" & Trim(fpText.text) & "' " & _
                     "AND   mfd_codreg=" & RS!mif_codreg & " " & _
                     "AND   mfd_codser=" & RS!mif_codser & " " & _
                     "AND   convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) >= " & Fecha & " AND convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) <= " & Fecha & " AND mfd_tipmin = '1'", vg_db, adOpenStatic
         End If
         If Not RS1.EOF Then estfij = True
         RS1.Close: Set RS1 = Nothing
         fecval = 0
         If Not estfij Then
            '-------> Buscar fecha mayor de estructura fija
            RS1.Open "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija " & _
                     "WHERE mif_cencos='" & Trim(fpText.text) & "' " & _
                     "AND   mif_codreg=" & RS!mif_codreg & " " & _
                     "AND   mif_codser=" & RS!mif_codser & "", vg_db, adOpenStatic
            If Not RS1.EOF Then fecval = IIf(IsNull(RS1!fecval), 0, RS1!fecval)
            RS1.Close: Set RS1 = Nothing
            If fecval > 0 Then
               '-------> Traer estructura fija
               For i = diaini To diafin
                   If fecval <= Fecha & Right("0" & i, 2) Then
                      vg_db.Execute "INSERT INTO " & aAp & " SELECT " & Fecha & Right("0" & i, 2) & " AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                                    "0 AS cantidad1, a.mif_canpro AS cantidad2, 0 AS ped_fecped " & _
                                    "FROM  b_minutafija a, b_productos b, b_ingrediente c, b_productosing d " & _
                                    "WHERE a.mif_codpro=b.pro_codigo " & _
                                    "AND   b.pro_codigo=d.pri_codpro " & _
                                    "AND   c.ing_codigo=d.pri_coding " & _
                                    "AND  (b.pro_fecven>" & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                    "AND   a.mif_cencos='" & Trim(fpText.text) & "' " & _
                                    "AND   a.mif_codreg=" & RS!mif_codreg & " " & _
                                    "AND   a.mif_codser=" & RS!mif_codser & " " & _
                                    "AND   a.mif_fecval=" & fecval & " " & _
                                    "AND   a.mif_dianro=" & fg_NumDia(Trim(Left(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                    "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
                    
                      vg_db.Execute "INSERT INTO " & aAp & " SELECT " & Fecha & Right("0" & i, 2) & " AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                                    "0 AS cantidad1, a.mif_canpro AS cantidad2, 0 AS ped_fecped " & _
                                    "FROM  b_minutafija a, b_productos b, b_ingrediente c, b_productosing d " & _
                                    "WHERE a.mif_codpro=b.pro_codigo " & _
                                    "AND   b.pro_codigo=d.pri_codpro " & _
                                    "AND   c.ing_codigo=d.pri_coding " & _
                                    "AND  (b.pro_fecven>" & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                    "AND   a.mif_cencos='" & Trim(fpText.text) & "' " & _
                                    "AND   a.mif_codreg=" & RS!mif_codreg & " " & _
                                    "AND   a.mif_codser=" & RS!mif_codser & " " & _
                                    "AND   a.mif_fecval=" & fecval & " " & _
                                    "AND   a.mif_dianro=" & fg_NumDia(Trim(Left(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                    "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
                   End If
                   Set RS1 = Nothing
               Next i
            End If
         ElseIf estfij Then
             '-------> Calcular datos desde tabla estructura fija día
             If vg_tipbase = "1" Then
                vg_db.Execute "INSERT INTO " & aAp & " SELECT a.mfd_fecha AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                              "0 AS cantidad1, a.mfd_canpro AS cantidad2, 0 AS ped_fecped " & _
                              "FROM b_minutafijadia a, b_productos b, b_ingrediente c, b_productosing d " & _
                              "WHERE a.mfd_codpro=b.pro_codigo " & _
                              "AND   b.pro_codigo=d.pri_codpro " & _
                              "AND   c.ing_codigo=d.pri_coding " & _
                              "AND  (b.pro_fecven>" & Format(Date, "yyyymmdd") & " OR b.pro_fecven<=0) " & _
                              "AND   a.mfd_cencos='" & Trim(fpText.text) & "' " & _
                              "AND   a.mfd_codreg=" & RS!mif_codreg & " " & _
                              "AND   a.mfd_codser=" & RS!mif_codser & " " & _
                              "AND mid(a.mfd_fecha,1,6)=" & Fecha & " AND a.mfd_tipmin = '1' " & _
                              "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
             
                vg_db.Execute "INSERT INTO " & aAp & " SELECT a.mfd_fecha AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                              "0 AS cantidad1, a.mfd_canpro AS cantidad2, 0 AS ped_fecped " & _
                              "FROM b_minutafijadia a, b_productos b, b_ingrediente c, b_productosing d " & _
                              "WHERE a.mfd_codpro = b.pro_codigo " & _
                              "AND   b.pro_codigo = d.pri_codpro " & _
                              "AND   c.ing_codigo = d.pri_coding " & _
                              "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                              "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                              "AND   a.mfd_codreg = " & RS!mif_codreg & " " & _
                              "AND   a.mfd_codser = " & RS!mif_codser & " " & _
                              "AND mid(a.mfd_fecha,1,6) = " & Fecha & " AND a.mfd_tipmin='1' " & _
                              "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
             Else
                vg_db.Execute "INSERT INTO " & aAp & " SELECT a.mfd_fecha AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                              "0 AS cantidad1, a.mfd_canpro AS cantidad2, 0 AS ped_fecped " & _
                              "FROM b_minutafijadia a, b_productos b, b_ingrediente c, b_productosing d " & _
                              "WHERE a.mfd_codpro = b.pro_codigo " & _
                              "AND   b.pro_codigo = d.pri_codpro " & _
                              "AND   c.ing_codigo = d.pri_coding " & _
                              "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                              "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                              "AND   a.mfd_codreg = " & RS!mif_codreg & " " & _
                              "AND   a.mfd_codser = " & RS!mif_codser & " " & _
                              "AND convert(int,substring(convert(varchar(8),a.mfd_fecha),1,6)) = " & Fecha & " AND a.mfd_tipmin = '1' " & _
                              "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
             
                vg_db.Execute "INSERT INTO " & aAp & " SELECT a.mfd_fecha AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                              "0 AS cantidad1, a.mfd_canpro AS cantidad2, 0 AS ped_fecped " & _
                              "FROM b_minutafijadia a, b_productos b, b_ingrediente c, b_productosing d " & _
                              "WHERE a.mfd_codpro = b.pro_codigo " & _
                              "AND   b.pro_codigo = d.pri_codpro " & _
                              "AND   c.ing_codigo = d.pri_coding " & _
                              "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                              "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                              "AND   a.mfd_codreg = " & RS!mif_codreg & " " & _
                              "AND   a.mfd_codser = " & RS!mif_codser & " " & _
                              "AND convert(int,substring(convert(varchar(8),a.mfd_fecha),1,6)) = " & Fecha & " AND a.mfd_tipmin = '1' " & _
                              "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
             End If
         End If
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
   '-------> Mover fecha pedido
   '-------> actualizar fecha pedido mensual
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'M')", vg_db, adOpenStatic
   If Not RS.EOF Then
      fecxin = Format(CDate("01/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("15/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.ped_fecped = " & fecxin & " WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'M' AND c.pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
'         vg_db.Execute "UPDATE " & aAp & "  " & _
'                       "SET " & aAp & ".ped_fecped = substring(convert(varchar(8),a.min_fecmin),1,6) + '01' " & _
'                       "FROM " & aAp & "  a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
'                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'M' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      
         vg_db.Execute "UPDATE " & aAp & "  " & _
                       "SET " & aAp & ".ped_fecped = " & fecxin & " " & _
                       "FROM " & aAp & "  a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'M' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
   End If
   RS.Close: Set RS = Nothing
      
   '-------> actualizar fecha pedido quincenal 1-15
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'Q1')", vg_db, adOpenStatic
   If Not RS.EOF Then
      fecxin = Format(CDate("01/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate("15/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.ped_fecped = " & fecxin & " WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND pad_tipo = 'Q1' AND pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
'         vg_db.Execute "UPDATE " & aAp & " " & _
'                       "SET " & aAp & ".ped_fecped = substring(convert(varchar(8),a.min_fecmin),1,6) + '01' " & _
'                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
'                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q1' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".ped_fecped = '" & fecxin & "' " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q1' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
      fecxin = Format(CDate("15/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("15/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.ped_fecped = " & fecxin & "  WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND pad_tipo = 'Q1' AND pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
'         vg_db.Execute "UPDATE " & aAp & " " & _
'                       "SET " & aAp & ".ped_fecped = substring(convert(varchar(8),a.min_fecmin),1,6) + '15' " & _
'                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
'                      "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q1' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".ped_fecped = " & fecxin & " " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                      "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q1' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
   End If
   RS.Close: Set RS = Nothing
   '-------> actualizar fecha pedido quincenal 2-16
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'Q2')", vg_db, adOpenStatic
   If Not RS.EOF Then
      fecxin = Format(CDate("02/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate("16/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.ped_fecped = " & fecxin & "  WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND pad_tipo = 'Q2' AND pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
'          vg_db.Execute "UPDATE " & aAp & " " & _
'                        "SET " & aAp & ".ped_fecped = substring(convert(varchar(8),a.min_fecmin),1,6) + '02' " & _
'                        "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
'                        "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q2' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
          vg_db.Execute "UPDATE " & aAp & " " & _
                        "SET " & aAp & ".ped_fecped = " & fecxin & " " & _
                        "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                        "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q2' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
      fecxin = Format(CDate("16/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("16/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.ped_fecped = " & fecxin & " WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q2' AND c.pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
'         vg_db.Execute "UPDATE " & aAp & " " & _
'                       "SET " & aAp & ".ped_fecped = substring(convert(varchar(8),a.min_fecmin),1,6) + '16' " & _
'                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
'                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q2' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".ped_fecped = " & fecxin & " " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q2' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
   End If
   RS.Close: Set RS = Nothing
   '-------> actualizar fecha pedido quincenal 3-17
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'Q3')", vg_db, adOpenStatic
   If Not RS.EOF Then
      fecxin = Format(CDate("03/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate("17/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.ped_fecped = " & fecxin & " WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q3' AND c.pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
'         vg_db.Execute "UPDATE " & aAp & " " & _
'                       "SET " & aAp & ".ped_fecped = substring(convert(varchar(8),a.min_fecmin),1,6) + '03' " & _
'                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
'                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q3' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".ped_fecped = " & fecxin & " " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q3' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
      fecxin = Format(CDate("17/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("17/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.ped_fecped = " & fecxin & "  WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q3' AND c.pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
'         vg_db.Execute "UPDATE " & aAp & " " & _
'                       "SET " & aAp & ".ped_fecped = substring(convert(varchar(8),a.min_fecmin),1,6) + '17' " & _
'                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
'                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q3' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".ped_fecped = " & fecxin & " " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q3' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
   End If
   RS.Close: Set RS = Nothing
   '-------> actualizar fecha pedido quincenal 4-18
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'Q4')", vg_db, adOpenStatic
   If Not RS.EOF Then
      fecxin = Format(CDate("04/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate("18/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.ped_fecped = " & fecxin & " WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q4' AND c.pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
'         vg_db.Execute "UPDATE " & aAp & " " & _
'                       "SET " & aAp & ".ped_fecped = substring(convert(varchar(8),a.min_fecmin),1,6) + '04' " & _
'                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
'                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q4' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".ped_fecped = " & fecxin & " " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q4' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
      fecxin = Format(CDate("18/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("18/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.ped_fecped =  " & fecxin & "  WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q4' AND c.pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
'         vg_db.Execute "UPDATE " & aAp & " " & _
'                       "SET " & aAp & ".ped_fecped = substring(convert(varchar(8),a.min_fecmin),1,6) + '18' " & _
'                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
'                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q4' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".ped_fecped = " & fecxin & " " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q4' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
   End If
   RS.Close: Set RS = Nothing
   '-------> actualizar fecha pedido cada 10 días.
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'Q4')", vg_db, adOpenStatic
   If Not RS.EOF Then
      fecxin = Format(CDate("01/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate("10/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.ped_fecped = " & fecxin & " WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q4' AND c.pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
'         vg_db.Execute "UPDATE " & aAp & " " & _
'                       "SET " & aAp & ".ped_fecped = substring(convert(varchar(8),a.min_fecmin),1,6) + '01' " & _
'                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
'                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q4' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".ped_fecped = " & fecxin & " " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q4' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
      fecxin = Format(CDate("11/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("20/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.ped_fecped = " & fecxin & "  WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q4' AND c.pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
'         vg_db.Execute "UPDATE " & aAp & " " & _
'                       "SET " & aAp & ".ped_fecped = substring(convert(varchar(8),a.min_fecmin),1,6) + '11' " & _
'                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
'                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q4' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".ped_fecped = " & fecxin & " " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q4' AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      End If
      fecxin = Format(CDate("21/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
      fecxfi = Format(CDate(dEoM("27/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
      fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
      If vg_tipbase = "1" Then
         vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip=b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo=d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                       "SET a.ped_fecped = " & fecxin & "  WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND c.pad_tipo = 'Q4' AND c.pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
      Else
'         vg_db.Execute "UPDATE " & aAp & " " & _
'                       "SET " & aAp & ".ped_fecped=substring(convert(varchar(8),a.min_fecmin),1,6) + '21' " & _
'                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
'                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip=b.tip_codigo AND b.tip_codigo=d.pro_codtip AND d.pro_previo=c.pad_codigo AND c.pad_tipo='Q4' AND a.ped_fecped=0 AND a.min_fecmin>=" & fecxin & " AND a.min_fecmin<=" & fecxfi & ""
         vg_db.Execute "UPDATE " & aAp & " " & _
                       "SET " & aAp & ".ped_fecped= " & fecxin & " " & _
                       "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                       "WHERE c.pad_cencos = '" & MuestraCasino(1) & "' AND a.pro_codtip=b.tip_codigo AND b.tip_codigo=d.pro_codtip AND d.pro_previo=c.pad_codigo AND c.pad_tipo='Q4' AND a.ped_fecped=0 AND a.min_fecmin>=" & fecxin & " AND a.min_fecmin<=" & fecxfi & ""
      End If
   End If
   RS.Close: Set RS = Nothing
   '-------> actualizar fecha diario y semanal
   RS.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos='" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo='E' OR pad_tipo='S')", vg_db, adOpenStatic
   Do While Not RS.EOF
      fecini = "01/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)
      fecfin = "07/" & IIf(Mid(Fecha, 5, 2) = 12, "01/" & Mid(Fecha, 1, 4) + 1, Mid(Fecha, 5, 2) + 1 & "/" & Mid(Fecha, 1, 4)) 'dEoM("27/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))
      fecpin = 0: fecpfi = 0
      Do While fecini <= fecfin
         '-------> Buscar fecha inicial y fecha final
         For j = 1 To 7
             If (DatePart("w", fecini, 2)) = Val(Mid(RS!pad_diario, j, 1)) Then
                If fecpin = 0 Then
                   fecpin = Format(fecini, "yyyymmdd")
                ElseIf fecpfi = 0 Then
                   fecpfi = Format(fecini, "yyyymmdd")
                End If
             End If
             If fecpin > 0 And fecpfi > 0 Then
                fecxin = Format(CDate(fg_Ctod1(fecpin)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
                fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
                fecxfi = Format(CDate(fg_Ctod1(fecpfi)) + IIf(IsNull(RS!pad_diaseg), 0, RS!pad_diaseg), "yyyymmdd")
                fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
                If vg_tipbase = "1" Then
                   vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo=c.pad_codigo " & _
                                 "SET a.ped_fecped = " & fecxin & " WHERE c.pad_cencos='" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo='E' OR pad_tipo='S') AND pad_codigo=" & RS!pad_codigo & " AND a.ped_fecped=0 AND a.min_fecmin>=" & fecxin & " AND a.min_fecmin <= " & fecxfi & ""
                Else
                   vg_db.Execute "UPDATE " & aAp & " SET " & aAp & ".ped_fecped=" & fecxin & " FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d WHERE a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo " & _
                                 "AND c.pad_cencos='" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo='E' OR pad_tipo='S')AND pad_codigo=" & RS!pad_codigo & " AND a.ped_fecped=0 AND a.min_fecmin>=" & fecxin & " AND a.min_fecmin<=" & fecxfi & ""
                End If
                fecpin = fecpfi: fecpfi = 0
                Exit For
             End If
         Next j
         fecini = fecini + 1
      Loop
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
   '-------> Leer archivo temporales
   vg_db.Execute ("DELETE " & aAp & " FROM  " & aAp & " WHERE ped_fecped=0 or ((cantidad1=0 or cantidad1 is null)And (cantidad2=0 or cantidad2 is null))")
   sql1 = " ": sql2 = " ": sql3 = " ": sql4 = " ": sql5 = " ": sql6 = " "
   If vg_pais = "CO" Then
      sql1 = IIf(vg_tipbase = "1", " AND cdate(i.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), i.foc_vigfin,101) >  '" & Date & "'")
      sql6 = IIf(vg_tipbase = "1", " AND cdate(b_formatocompras.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), b_formatocompras.foc_vigfin,101) >  '" & Date & "'")
      sql2 = " , b_formatocompras i, b_formatocomprassgp j "
      sql3 = " i.foc_codsac, i.foc_nomsac, i.foc_unisac, i.foc_faccon, "
'      sql4 = "AND   i.foc_codsac = j.fcs_codsac AND   i.foc_codsac = (SELECT TOP 1 b_formatocomprassgp.fcs_codsac FROM b_formatocomprassgp WHERE b_formatocomprassgp.fcs_codsgp = j.fcs_codsgp) AND   j.fcs_codsgp = c.pro_codigo AND (j.fcs_cenpre = 1 or j.fcs_sgppre = 1) AND (i.foc_flexec = 0 OR (i.foc_flexec = -1 " & sql1 & ")) "
      sql4 = "AND   i.foc_codsac = j.fcs_codsac AND   i.foc_faccon > 0 AND i.foc_codsac = (SELECT TOP 1 b_formatocomprassgp.fcs_codsac FROM b_formatocomprassgp, b_formatocompras WHERE b_formatocomprassgp.fcs_codsac = b_formatocompras.foc_codsac and b_formatocomprassgp.fcs_codsgp = j.fcs_codsgp AND b_formatocomprassgp.fcs_cenpre = 1 AND (b_formatocompras.foc_flexec = 0 OR (b_formatocompras.foc_flexec = -1 " & sql6 & "))) AND j.fcs_codsgp = c.pro_codigo AND (i.foc_flexec = 0 OR (i.foc_flexec = -1 " & sql1 & ")) "
      sql5 = " , i.foc_codsac, i.foc_nomsac, i.foc_unisac, i.foc_faccon "
   Else
      sql3 = " c.pro_codigo as foc_codsac, c.pro_nombre as foc_nomsac, e.unm_nomcor as foc_unisac, 1 as foc_faccon, "
'      sql5 = " , c.pro_codigo, c.pro_nombre, e.unm_nomcor, i.foc_faccon "
      sql5 = " , c.pro_codigo, c.pro_nombre, e.unm_nomcor "
   End If
   proc1 = "SELECT a.ped_fecped, b.ing_codigo, b.ing_nombre, " & _
           "e.unm_nomcor, c.pro_codigo, c.pro_nombre, " & _
           "c.pro_coduni, c.pro_facsto, c.pro_ctacon, h.pro_previo AS tip_previo, d.uni_nomcor, " & sql3 & " " & _
           "SUM(a.cantidad1) AS cantidad1, SUM(a.cantidad2) AS cantidad2, (SELECT DISTINCT cpi_precos FROM b_contlistpreing WHERE cpi_coding = b.ing_codigo AND cpi_cencos = '" & MuestraCasino(1) & "') AS cpi_precos " & _
           "FROM  " & aAp & " a, b_ingrediente b, b_productos c, a_unidad d, a_unidadmed e, a_tipopro f, b_paramdesp g, " & aAp1 & " h " & sql2 & " " & _
           "WHERE a.ing_codigo=b.ing_codigo " & _
           "AND   a.pro_codigo=c.pro_codigo " & _
           "AND   c.pro_codtip=f.tip_codigo " & _
           "AND   f.tip_codigo=h.pro_codtip AND h.pro_previo = g.pad_codigo AND g.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
           "AND   b.ing_unimed=e.unm_codigo " & _
           "AND   c.pro_coduni=d.uni_codigo " & _
           "AND  (c.pro_fecven>" & Format(Date, "yyyymmdd") & " OR c.pro_fecven <= 0) " & _
           "AND   c.pro_facsto>0 " & _
           "     " & sql4 & "  " & _
           "GROUP BY a.ped_fecped, b.ing_codigo, b.ing_nombre, " & _
           "e.unm_nomcor, c.pro_codigo, c.pro_nombre, c.pro_coduni, " & _
           "c.pro_facsto, d.uni_nomcor, c.pro_ctacon, h.pro_previo, d.uni_nomcor " & sql5 & " ORDER BY c.pro_ctacon, h.pro_previo, b.ing_codigo, b.ing_nombre, c.pro_nombre, a.ped_fecped"
   MoverDatosGrilla proc1, proc2, proc3
End If
Toolbar2.Enabled = IIf(estexi, False, True)
End Sub

Sub MoverDatosGrilla(proc1 As String, proc2 As String, proc3 As String)
Dim RS As New ADODB.Recordset
Dim proc11 As String, proc22 As String, proc33 As String, sql1 As String, sql2 As String, sql3 As String
Dim codpre As Long, despa As Long, cansol As Double, canppr As Double, pcanrea As Double, pcanmin As Double
Dim necmin As Double, stoact As Double, ordrec As Double, conmir As Double, canres As Double, faccon As Double, stodia As Double, canrea As Double
Dim auxing As String, auxpro As String
Dim i As Long
auxpro = "": auxing = "": codpre = 0: despa = 0: i = 0: cansol = 0: canrea = 0
With vaSpread1
    .Visible = False
    
    If Not estexi Then
       RS.Open proc1 & proc2 & proc3, vg_db, adOpenStatic
       If Not RS.EOF Then Toolbar1.Buttons(1).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(8).Enabled = False Else Frame2.Enabled = False: Toolbar1.Buttons(1).Enabled = False
    ElseIf estexi Then
       sql1 = "": sql2 = " ": sql3 = " ": sql4 = " "
       proc11 = "": proc22 = "": proc33 = ""
       If vg_pais = "CO" Then
          sql1 = IIf(vg_tipbase = "1", " AND cdate(i.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), i.foc_vigfin,101) >  '" & Date & "'")
          sql2 = ", b_formatocompras i, b_formatocomprassgp j "
          sql3 = " AND   i.foc_codsac = j.fcs_codsac AND   j.fcs_codsgp = b.pro_codigo AND i.foc_codsac = a.ped_codsac "
          sql4 = " i.foc_codsac, i.foc_nomsac, i.foc_unisac, i.foc_faccon, "
       Else
          sql4 = " b.pro_codigo AS foc_codsac, b.pro_nombre AS foc_nomsac, d.uni_nomcor AS foc_unisac, 1 AS foc_faccon, "
       End If
       If vg_tipbase = "1" Then
            proc11 = "SELECT c.ing_codigo, c.ing_nombre, (SELECT DISTINCT cpi_precos FROM b_contlistpreing WHERE cpi_coding=c.ing_codigo AND cpi_cencos='" & MuestraCasino(1) & "'), e.unm_nomcor, " & _
                    "b.pro_codigo, b.pro_nombre, b.pro_coduni, b.pro_facsto, " & _
                    "d.uni_nomcor, " & sql4 & " a.ped_canped AS cantidad2, a.ped_canmin, (SELECT DISTINCT ppd_propon FROM " & aAp2 & " WHERE ppd_codpro=b.pro_codigo AND ppd_cencos='" & MuestraCasino(1) & "' AND ppd_fecdia=" & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS cpi_precos, " & _
                    "a.ped_stoact, a.ped_conrea, a.ped_ordrec, a.ped_proped, a.ped_fecped, b.pro_ctacon, h.pro_previo AS tip_previo, a.ped_persac, a.ped_semsac " & _
                    "FROM  b_minutapedidos a, b_productos b, b_ingrediente c, a_unidad d, a_unidadmed e, a_tipopro f, b_paramdesp g, " & aAp1 & " h " & sql2 & " " & _
                    "WHERE a.ped_codpro=b.pro_codigo " & _
                    "AND   b.pro_codtip=f.tip_codigo " & _
                    "AND   f.tip_codigo=h.pro_codtip AND h.pro_previo=g.pad_codigo AND g.pad_cencos='" & LimpiaDato(Trim(fpText.text)) & "' " & _
                    "AND   a.ped_coding=c.ing_codigo " & _
                    "AND   c.ing_unimed=e.unm_codigo " & _
                    "AND   b.pro_coduni=d.uni_codigo " & _
                    "AND   a.ped_codcas='" & LimpiaDato(Trim(fpText.text)) & "' " & _
                    "AND   a.ped_anomes=" & Fecha & " " & _
                    "AND   a.ped_tipped=1 " & _
                    "      " & sql3 & " " & _
                    "ORDER BY b.pro_ctacon, tip_previo, c.ing_codigo, c.ing_nombre, b.pro_nombre, a.ped_fecped "
            proc22 = "UNION "
            proc33 = "SELECT 'zzzfija', 'Estructura Fija', 0, '', " & _
                    "b.pro_codigo, b.pro_nombre, b.pro_coduni, b.pro_facsto, " & _
                    "d.uni_nomcor, " & sql4 & " a.ped_canped AS cantidad2, a.ped_canmin, (SELECT DISTINCT ppd_propon FROM " & aAp2 & " WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS cpi_precos, " & _
                    "a.ped_stoact, a.ped_conrea, a.ped_ordrec, a.ped_proped, a.ped_fecped, b.pro_ctacon, h.pro_previo AS tip_previo, a.ped_persac, a.ped_semsac " & _
                    "FROM  b_minutapedidos a, b_productos b, b_ingrediente c, a_unidad d, a_unidadmed e, a_tipopro f, b_paramdesp g, " & aAp1 & " h " & sql2 & " " & _
                    "WHERE a.ped_codpro = b.pro_codigo " & _
                    "AND   b.pro_codtip = f.tip_codigo " & _
                    "AND   f.tip_codigo = h.pro_codtip AND h.pro_previo = g.pad_codigo AND g.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                    "AND   a.ped_coding = '' " & _
                    "AND   c.ing_unimed = e.unm_codigo " & _
                    "AND   b.pro_coduni = d.uni_codigo " & _
                    "AND   a.ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                    "AND   a.ped_anomes = " & Fecha & " " & _
                    "AND   a.ped_tipped = 1 " & _
                    "      " & sql3 & " " & _
                    "ORDER BY b.pro_ctacon, tip_previo, c.ing_codigo, c.ing_nombre, b.pro_nombre, a.ped_fecped"
       Else
          proc11 = "SELECT c.ing_codigo, c.ing_nombre, (SELECT DISTINCT cpi_precos FROM b_contlistpreing WHERE cpi_coding = c.ing_codigo AND cpi_cencos = '" & MuestraCasino(1) & "'), e.unm_nomcor, " & _
                  "b.pro_codigo, b.pro_nombre, b.pro_coduni, b.pro_facsto, " & _
                  "d.uni_nomcor, " & sql4 & " a.ped_canped AS cantidad2, a.ped_canmin, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS cpi_precos, " & _
                  "a.ped_stoact, a.ped_conrea, a.ped_ordrec, a.ped_proped, a.ped_fecped, b.pro_ctacon, h.pro_previo AS tip_previo, a.ped_persac, a.ped_semsac " & _
                  "FROM  b_minutapedidos a, b_productos b, b_ingrediente c, a_unidad d, a_unidadmed e, a_tipopro f, b_paramdesp g, " & aAp1 & " h " & sql2 & " " & _
                  "WHERE a.ped_codpro = b.pro_codigo " & _
                  "AND   b.pro_codtip = f.tip_codigo " & _
                  "AND   f.tip_codigo = h.pro_codtip AND h.pro_previo = g.pad_codigo AND g.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                  "AND   a.ped_coding = c.ing_codigo " & _
                  "AND   c.ing_unimed = e.unm_codigo " & _
                  "AND   b.pro_coduni = d.uni_codigo " & _
                  "AND   a.ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                  "AND   a.ped_anomes = " & Fecha & " " & _
                  "AND   a.ped_tipped = 1 " & _
                  "      " & sql3 & " " & _
                  " "
          proc22 = "UNION "
          proc33 = "SELECT 'zzzfija', 'Estructura Fija', 0, '', " & _
                  "b.pro_codigo, b.pro_nombre, b.pro_coduni, b.pro_facsto, " & _
                  "d.uni_nomcor, " & sql4 & " a.ped_canped AS cantidad2, a.ped_canmin, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS cpi_precos, " & _
                  "a.ped_stoact, a.ped_conrea, a.ped_ordrec, a.ped_proped, a.ped_fecped, b.pro_ctacon, h.pro_previo AS tip_previo, a.ped_persac, a.ped_semsac " & _
                  "FROM  b_minutapedidos a, b_productos b, b_ingrediente c, a_unidad d, a_unidadmed e, a_tipopro f, b_paramdesp g, " & aAp1 & " h " & sql2 & " " & _
                  "WHERE a.ped_codpro = b.pro_codigo " & _
                  "AND   b.pro_codtip = f.tip_codigo " & _
                  "AND   f.tip_codigo = h.pro_codtip AND h.pro_previo = g.pad_codigo AND g.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                  "AND   a.ped_coding = '' " & _
                  "AND   c.ing_unimed = e.unm_codigo " & _
                  "AND   b.pro_coduni = d.uni_codigo " & _
                  "AND   a.ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                  "AND   a.ped_anomes = " & Fecha & " " & _
                  "AND   a.ped_tipped = 1 " & _
                  "      " & sql3 & " " & _
                  "ORDER BY b.pro_ctacon, tip_previo, c.ing_codigo, c.ing_nombre, b.pro_nombre, a.ped_fecped"
       End If
       RS.Open proc11 & proc22 & proc33, vg_db, adOpenStatic
    End If
    If Not RS.EOF Then
       If fecenv > 0 Then
          fpDateTime2.text = IIf(IsNull(RS!ped_persac), fpDateTime1.text, RS!ped_persac)
          fpLongInteger1(0).Value = IIf(IsNull(RS!ped_semsac), "", RS!ped_semsac)
       End If
       Do While Not RS.EOF
          If RS!tip_previo <> codpre Then
             .MaxRows = .MaxRows + 1
             .Row = .MaxRows
             .Col = -1: .BackColor = &HFFFFC0
             .Col = 2: .Font.Bold = True: .CellType = CellTypeStaticText: .text = fg_BuscaenArbol(RS!tip_previo, "a_tipopro", "tip_codigo")
             .Col = 3: .CellType = CellTypeStaticText: .text = ""
             .Col = 4: .CellType = CellTypeStaticText: .text = ""
             .Col = 5: .CellType = CellTypeStaticText: .text = ""
             .Col = 6: .CellType = CellTypeStaticText: .text = ""
             .Col = 7: .CellType = CellTypeStaticText: .text = ""
             .Col = 8: .CellType = CellTypeStaticText: .text = ""
             .Col = 16: .CellType = CellTypeStaticText: .text = RS!tip_previo
             codpre = RS!tip_previo
          End If
          If RS!pro_codigo <> auxpro Or RS!ing_codigo <> auxing Then
             If i > 0 Then
                .Row = i
                If Not estexi Then
                   .Col = 4: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = Format(cansol, fg_Pict(6, vg_DCa))
                Else
                   .Col = 4: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = Format(pcanmin, fg_Pict(6, vg_DCa))
                End If
                If Not estexi Then
                   necmin = IIf(vg_pais = "CO", canrea, (cansol - stoact - ordrec + conmir))
                Else
                   necmin = IIf(vg_pais = "CO", canrea, (pcanmin - stoact - ordrec + pcanrea))
                End If
                necmin = IIf(necmin < 0, 0, necmin)
                canrea = IIf(canrea < 0, 0, canrea)
                .Col = 5: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight
                .text = canrea
                .Col = 8: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight
                .text = canrea
                cansol = 0: stoact = 0: ordrec = 0: conmir = 0: pcanmin = 0: pcanrea = 0: stodia = 0
             End If
             .MaxRows = .MaxRows + 1
             .Row = .MaxRows
             i = .Row
             .Col = 1: .CellType = CellTypeStaticText: .text = Trim(RS!foc_codsac)
             .Col = 2: .CellType = CellTypeStaticText: .text = Trim(RS!foc_nomsac)
             .Col = 3: .CellType = CellTypeStaticText: .text = Trim(RS!foc_unisac)
             .Col = 4: .CellType = CellTypeStaticText: .text = ""
             .Col = 5: .CellType = CellTypeStaticText: .text = ""
             .Col = 6: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = RS!foc_faccon '"1"
             .Col = 7: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = ""
             .Col = 8: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = ""
             .Col = 9: .CellType = CellTypeStaticText: .text = Trim(RS!ing_codigo)
             .Col = 10: .CellType = CellTypeStaticText: .text = Trim(RS!ing_nombre)
             .Col = 11: .CellType = CellTypeStaticText: .text = Trim(RS!pro_codigo)
             .Col = 12: .CellType = CellTypeStaticText: .text = Trim(RS!pro_nombre)
             conmir = 0: stoact = 0: ordrec = 0: stodia = 0
             If Not estexi Then
                If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone) <> -1 Then
                   vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone)
                   vaSpread2.Col = 2: .Col = 13: .CellType = CellTypeStaticText: .text = vaSpread2.text: stoact = Val(vaSpread2.text)
                   vaSpread2.Col = 4: .Col = 14: .CellType = CellTypeStaticText: .text = vaSpread2.text: conmir = Val(vaSpread2.text)
                   vaSpread2.Col = 5: .Col = 15: .CellType = CellTypeStaticText: .text = vaSpread2.text: ordrec = Val(vaSpread2.text)
                   stodia = stoact
'                   stodia = (stodia + ordrec)
                   If (stodia - conmir + ordrec) > 0 Then
                       stodia = (stodia - conmir + ordrec)
                   ElseIf (stodia - conmir + ordrec) <= 0 Then
                       stodia = 0
                   End If
                End If
             Else
                .Col = 13: .CellType = CellTypeStaticText: .text = RS!ped_stoact: stoact = RS!ped_stoact: stodia = RS!ped_stoact
                .Col = 14: .CellType = CellTypeStaticText: .text = RS!ped_conrea: conmir = RS!ped_conrea: stodia = (stodia + RS!ped_ordrec)
                .Col = 15: .CellType = CellTypeStaticText: .text = RS!ped_ordrec: ordrec = RS!ped_ordrec: stodia = (stodia - RS!ped_conrea)
                pcanmin = RS!ped_canmin: pcanrea = RS!ped_conrea
             End If
             .Col = 16: .CellType = CellTypeStaticText: .text = Format(RS!cpi_precos, fg_Pict(6, 2))
             necmin = IIf((conmir - stoact - ordrec) < 0, (conmir - stoact - ordrec) * -1, (conmir - stoact - ordrec))
             auxpro = RS!pro_codigo
             auxing = RS!ing_codigo
             faccon = RS!foc_faccon
             canrea = 0
             .MaxRows = .MaxRows + 1
             .Row = .MaxRows
             .Col = 1: .CellType = CellTypeStaticText: .text = ""
             .Col = 2: .CellType = CellTypeStaticText: .text = ""
             .Col = 3: .CellType = CellTypeStaticText: .text = ""
             .Col = 4: .CellType = CellTypeStaticText: .text = ""
             .Col = 5: .CellType = CellTypeStaticText: .text = ""
             .Col = 6: .CellType = CellTypeStaticText: .text = ""
             .Col = 7: .CellType = CellTypeStaticText: .text = CDate(fg_Ctod1(RS!ped_fecped))
             If Not estexi Then
'                necmin = IIf(Int(((RS!cantidad2 + conmir) - stodia) / RS!foc_faccon) <> (((RS!cantidad2 + conmir) - stodia) / RS!foc_faccon), Int(((RS!cantidad2 + conmir) - stodia) / RS!foc_faccon) + 1, Round(((RS!cantidad2 + conmir) - stodia) / RS!foc_faccon, 0))
                necmin = IIf(Int(((RS!cantidad2) - stodia) / RS!foc_faccon) <> (((RS!cantidad2) - stodia) / RS!foc_faccon), Int(((RS!cantidad2) - stodia) / RS!foc_faccon) + 1, Round(((RS!cantidad2) - stodia) / RS!foc_faccon, 0))
                If Not IsNull(RS!cantidad2) Then
                   If Not estexi Then
                      cansol = cansol + IIf(Int((RS!cantidad2) / RS!pro_facsto) <> ((RS!cantidad2) / RS!pro_facsto), ((RS!cantidad2) / RS!pro_facsto), Round((RS!cantidad2) / RS!pro_facsto, 0)) * RS!pro_facsto
                   End If
                End If
                canres = 0
'                stodia = IIf((stodia + (necmin * RS!foc_faccon) - (RS!cantidad2 + conmir)) > 0, (stodia + (IIf(necmin < 0, 0, necmin) * RS!foc_faccon) - (RS!cantidad2 + conmir)), 0)
                stodia = IIf((stodia + -(RS!cantidad2)) > 0, (stodia + (IIf(necmin < 0, 0, necmin) * RS!foc_faccon) - (RS!cantidad2)), 0)
'                canres = Round(IIf(stodia > 0, 0, (necmin * RS!foc_faccon)) - IIf((stodia - (RS!cantidad2 + conmir)) > 0, 0, (stodia - (RS!cantidad2 + conmir)) * -1), vg_DCa)
                canres = Round(IIf(stodia > 0, 0, (necmin * RS!foc_faccon)) - IIf((stodia - (RS!cantidad2)) > 0, 0, (stodia - (RS!cantidad2)) * -1), vg_DCa)
                If canres < 0 Then canres = 0
             Else
                necmin = RS!ped_proped
             End If
             canrea = canrea + IIf(necmin < 0, 0, necmin)
             .Col = 8
             If Not estexi Then
                .CellType = CellTypeStaticText
             Else
                .ForeColor = IIf(RS!ped_conrea = 0 And RS!ped_canmin = 0, &H8000000D, &H80000012)
                .CellType = IIf(RS!ped_conrea = 0 And RS!ped_canmin = 0, CellTypeNumber, CellTypeStaticText)
             End If
             .TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(IIf(necmin < 0, 0, necmin), fg_Pict(6, 2)) 'Format(IIf(Int(RS!cantidad2 / RS!pro_facsto) <> (RS!cantidad2 / RS!pro_facsto), Int(RS!cantidad2 / RS!pro_facsto) + 1, Round(RS!cantidad2 / RS!pro_facsto, 0)) * RS!pro_facsto, fg_Pict(6, 2))
             .Col = 9: .CellType = CellTypeStaticText: .text = Trim(RS!ing_codigo)
             .Col = 10: .CellType = CellTypeStaticText: .text = Trim(RS!ing_nombre)
             .Col = 11: .CellType = CellTypeStaticText: .text = Trim(RS!pro_codigo)
             .Col = 12: .CellType = CellTypeStaticText: .text = Trim(RS!pro_nombre)
             necmin = 0
          Else
             .MaxRows = vaSpread1.MaxRows + 1
             .Row = .MaxRows
             .Col = 1: .CellType = CellTypeStaticText: .text = ""
             .Col = 2: .CellType = CellTypeStaticText: .text = ""
             .Col = 3: .CellType = CellTypeStaticText: .text = ""
             .Col = 4: .CellType = CellTypeStaticText: .text = ""
             .Col = 5: .CellType = CellTypeStaticText: .text = ""
             .Col = 6: .CellType = CellTypeStaticText: .text = ""
             .Col = 7: .CellType = CellTypeStaticText: .text = CDate(fg_Ctod1(RS!ped_fecped))
             If necmin < 0 Then
                If Not estexi Then
                   cantid = IIf((RS!cantidad2 - stodia) > 0, (RS!cantidad2 - stodia), (RS!cantidad2 - stodia))
                   necmin = IIf(Int((cantid - canres) / RS!foc_faccon) <> ((cantid - canres) / RS!foc_faccon), Int((cantid - canres) / RS!foc_faccon) + 1, Round((cantid - canres) / RS!foc_faccon, 0))
                   If Not IsNull(RS!cantidad2) Then
                      If Not estexi Then
                         cansol = cansol + IIf(Int(RS!cantidad2 / RS!pro_facsto) <> (RS!cantidad2 / RS!pro_facsto), (RS!cantidad2 / RS!pro_facsto), Round((RS!cantidad2) / RS!pro_facsto, 0)) * RS!pro_facsto
                      End If
                   End If
                   canres = Round((necmin * RS!foc_faccon) - (RS!cantidad2 - canres), vg_DCa)
                   stodia = IIf((stodia - RS!cantidad2) > 0, (stodia - RS!cantidad2), 0)
                Else
                   pcanmin = pcanmin + RS!ped_canmin
                   necmin = RS!ped_proped
                End If
                canrea = canrea + necmin
             Else
                If Not estexi Then
                   cantid = IIf((RS!cantidad2 - stodia) > 0, (RS!cantidad2 - stodia), (RS!cantidad2 - stodia) - 1)
                   necmin = IIf(Int((cantid - canres) / RS!foc_faccon) <> ((cantid - canres) / RS!foc_faccon), Int((cantid - canres) / RS!foc_faccon) + 1, Round((cantid - canres) / RS!foc_faccon, 0))
                   If Not IsNull(RS!cantidad2) Then
                      If Not estexi Then
                         cansol = cansol + IIf(Int(RS!cantidad2 / RS!pro_facsto) <> (RS!cantidad2 / RS!pro_facsto), (RS!cantidad2 / RS!pro_facsto), Round((RS!cantidad2) / RS!pro_facsto, 0)) * RS!pro_facsto
                      End If
                   End If
                   stodia = IIf((stodia - RS!cantidad2) > 0, (stodia - RS!cantidad2), 0)
                   canres = Round(IIf(stodia > 0, 0, (necmin * RS!foc_faccon)) - IIf((stodia - (cantid - canres)) > 0, 0, (stodia - (cantid - canres)) * -1), vg_DCa)
                   If canres < 0 Then canres = 0
                Else
                   pcanmin = pcanmin + RS!ped_canmin
                   necmin = RS!ped_proped
                End If
                canrea = canrea + IIf(necmin < 0, 0, necmin)
             End If
             .Col = 8
             If Not estexi Then
                .CellType = CellTypeStaticText
             Else
                .ForeColor = IIf(RS!ped_conrea = 0 And RS!ped_canmin = 0, &H8000000D, &H80000012)
                .CellType = IIf(RS!ped_conrea = 0 And RS!ped_canmin = 0, CellTypeNumber, CellTypeStaticText)
             End If
             .TypeHAlign = TypeHAlignRight: .text = Format(IIf(necmin < 0, 0, necmin), fg_Pict(6, 2))
             necmin = 0
             .Col = 9: .CellType = CellTypeStaticText: .text = Trim(RS!ing_codigo)
             .Col = 10: .CellType = CellTypeStaticText: .text = Trim(RS!ing_nombre)
             .Col = 11: .CellType = CellTypeStaticText: .text = Trim(RS!pro_codigo)
             .Col = 12: .CellType = CellTypeStaticText: .text = Trim(RS!pro_nombre)
          End If
          .Col = 17
          .text = 0
          If Not IsNull(RS!cantidad2) Then
             If Not estexi Then
                .text = Format(RS!cantidad2, fg_Pict(6, vg_DCa))
             Else
                .text = Format(RS!ped_canmin, fg_Pict(6, vg_DCa))
             End If
          End If
          RS.MoveNext
       Loop
       .Row = i
       If Not estexi Then
          .Col = 4: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = Format(cansol, fg_Pict(6, vg_DCa))
       Else
          .Col = 4: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = Format(pcanmin, fg_Pict(6, vg_DCa))
       End If
       necmin = (cansol - stoact - ordrec + conmir)
       .Col = 5: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = canrea 'Format(IIf(necmin < 0, (necmin * -1), necmin), fg_Pict(6, 2))
       .Col = 8: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = canrea 'Format(IIf(necmin < 0, 0, necmin), fg_Pict(6, 2))
       cansol = 0: stoact = 0: ordrec = 0: conmir = 0
    End If
    RS.Close: Set RS = Nothing: fg_descarga
    '-------> Borrar tablas temporales
    'If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
    'If Trim(aAp1) <> "" Then vg_db.Execute "DROP TABLE " & aAp1 & ""
    'If Trim(aAp2) <> "" Then vg_db.Execute "DROP TABLE " & aAp2 & ""
    .Visible = True
    If Me.Visible Then .SetFocus
End With
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If Not IsDate(fpDateTime1.text) Then Exit Sub
MoverDatos
End Sub

Private Sub fpDateTime2_Change()
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If Not IsDate(fpDateTime2.text) Then Exit Sub
End Sub

Private Sub fpText_Change()
If fpText.text = "" Then fpayuda.Caption = "": Exit Sub
Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(8).Enabled = False: vaSpread1.MaxRows = 0
RS.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & fpText.text & "' AND cli_tipo = 0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda.Caption = "": Exit Sub
fpayuda.Caption = Trim(RS!cli_nombre)
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
B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
B_TabEst.Show 1
Me.Refresh
If vg_codigo = "" Then Exit Sub
fpText.text = vg_codigo: fpayuda.Caption = vg_nombre
Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(8).Enabled = False: vaSpread1.MaxRows = 0
If Me.Visible Then fpDateTime1.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codpro As String, coding As String, codsac As String, i As Long
Dim fechasis As Long, fecdes As Long, nrosem As Long
Dim canmin As Double, cospro As Double, cosali As Double, CosDes As Double
Dim canped As Double, stoact As Double, proped As Double, pedpro As Double
Dim aAp As String, persac As String, sql1 As String, sql2 As String
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

On Error GoTo Man_Error

Select Case Button.Index
Case 1 '-------> Grabar pedido
    If vg_pais <> "CL" And Trim(fpLongInteger1(0).text) = "" Then MsgBox "Debe ingresar semana SAC...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If vg_paid <> "CL" And Trim(fpDateTime2.text) = "" Then MsgBox "Debe seleccionar periodo SAC...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText.text)) & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda.Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '-------> Validar si existen datos en planificación teórica
    persac = Format(fpDateTime2.text, "yyyymm")
    nrosem = IIf(fpLongInteger1(0).Visible = True, fpLongInteger1(0).Value, 0)
    If vg_tipbase = "1" Then
       RS.Open "SELECT DISTINCT a.min_cencos FROM b_minuta a, a_servicio b WHERE a.min_codser = b.ser_codigo and b.ser_activo = '1' and a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(a.min_fecmin,1,6)) = " & Fecha & "", vg_db, adOpenStatic
    Else
       RS.Open "SELECT DISTINCT a.min_cencos FROM b_minuta a, a_servicio b WHERE a.min_codser = b.ser_codigo and b.ser_activo = '1' and a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & Fecha & "", vg_db, adOpenStatic
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    fg_carga ""
    fechasis = 0: fechasis = Val(Mid(Date, 7, 4)) & Mid(Date, 4, 2) & Mid(Date, 1, 2)
    '-------> Grabar tabla b_minutapedidos
    Toolbar1.Enabled = False
    Toolbar2.Enabled = False
    Frame1(1).Enabled = False
    
    With vaSpread1
    
        .Enabled = False
        vg_db.BeginTrans
        canmin = 0: codsac = "": codpro = "": coding = "": canped = 0: stoact = 0: proped = 0: fecdes = 0: pedpro = 0
        '-------> Eliminar pedido
        vg_db.Execute "DELETE b_minutapedidos FROM b_minutapedidos WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " AND ped_tipped = 1"
        For i = 1 To .MaxRows
            DoEvents
            .Row = i: .Col = 1
            canped = 0: stoact = 0: proped = 0: pedpro = 0
            If .BackColor <> &HFFFFC0 Then
               .Col = 1
               If Trim(.text) <> "" Then
                  codsac = .text
                  'se saco el 4
    
                  .Col = 5: pedpro = IIf(Trim(.text) = "", 0, .text)
                  .Col = 9: coding = .text
                  .Col = 11: codpro = .text
                  .Col = 13: stoact = IIf(Trim(.text) = "", 0, .text)
                  .Col = 14: mincon = IIf(Trim(.text) = "", 0, .text)
                  .Col = 15: orcrec = IIf(Trim(.text) = "", 0, .text)
                  If vg_pais = "CO" Then
                     vg_db.Execute "UPDATE b_formatocomprassgp SET fcs_cenpre = 0 WHERE fcs_codsgp = '" & codpro & "'"
                     vg_db.Execute "UPDATE b_formatocomprassgp SET fcs_cenpre = 1 WHERE fcs_codsac = '" & codsac & "' AND fcs_codsgp = '" & codpro & "'"
                  End If
                  i = i + 1
               End If
               .Row = i: .Col = 7
               If .BackColor <> &HFFFFC0 And Trim(.text) <> "" Then
                  .Col = 17: canmin = IIf(Trim(.text) = "", 0, .text)
                  .Col = 7: fecdes = Format(.text, "yyyymmdd")
                  .Col = 8: cansol = IIf(Trim(.text) = "", 0, .text)
                  vg_db.Execute "INSERT INTO b_minutapedidos (ped_codcas, ped_fecped, ped_anomes, ped_tipped, ped_coding, ped_codpro, ped_codsac, ped_canmin, ped_canped, ped_fecenv, ped_stoact, ped_proped, ped_ordrec, ped_conrea, ped_persac, ped_semsac) " & _
                  "VALUES ('" & fpText.text & "', " & fecdes & ", " & Fecha & ", 1, '" & coding & "', '" & codpro & "', '" & codsac & "', " & Round(canmin, 2) & ", 0, 0, " & Round(stoact, 2) & ", " & Round(cansol, 2) & ", " & Round(orcrec, 2) & ", " & Round(mincon, 2) & ", '" & persac & "', " & nrosem & ")"
                  '-------> Actualizar codigo pedido en ingrediente
                  vg_db.Execute "UPDATE b_contlistpreing SET cpi_codped = '" & codpro & "' WHERE cpi_cencos = '" & MuestraCasino(1) & "' AND (cpi_codped = '' OR (cpi_codped) IS NULL)"
               End If
            End If
        Next i
        If vg_tipbase = "1" Then
           '-------> Insert tabla productospmpdia
           aAp = Trim(vg_NUsr) & "_tmp_ProductoPMPGenPed2"
           fg_CheckTmp aAp
           vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                         "INTO " & aAp & " " & _
                         "FROM b_productospmpdia " & _
                         "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                         "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                         "AND   ppd_propon>0 " & _
                         "GROUP BY ppd_cencos, ppd_codpro"
           vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
           vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon"
           vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
        End If
        '-------> Traer estructura fija
        Dim fecval As Long
        RS.Open "SELECT DISTINCT a.mif_cencos, a.mif_codreg, a.mif_codser FROM b_minutafija a, a_servicio b " & _
                "WHERE a.mif_codser = b.ser_codigo and b.ser_activo = '1' and a.mif_cencos = '" & LimpiaDato(Trim(fpText.text)) & "'", vg_db, adOpenStatic
        If Not RS.EOF Then
           Do While Not RS.EOF
              DoEvents
              '-------> Validar si existe estructura fija día
              If vg_tipbase = "1" Then
                 RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
                          "WHERE  mfd_cencos = '" & RS!mif_cencos & "' " & _
                          "AND    mfd_codreg = " & RS!mif_codreg & " " & _
                          "AND    mfd_codser = " & RS!mif_codser & " " & _
                          "AND mid(mfd_fecha,1,6) = " & Fecha & " " & _
                          "AND    mfd_tipmin = '1'", vg_db, adOpenStatic
              Else
                 RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
                          "WHERE  mfd_cencos = '" & RS!mif_cencos & "' " & _
                          "AND    mfd_codreg = " & RS!mif_codreg & " " & _
                          "AND    mfd_codser = " & RS!mif_codser & " " & _
                          "AND    convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) = " & Fecha & " " & _
                          "AND    mfd_tipmin = '1'", vg_db, adOpenStatic
              End If
              If RS1.EOF Then
                 RS1.Close: Set RS1 = Nothing
                 '-------> Buscar fecha mayor de estructura fija
                 fecval = 0
                 RS1.Open "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija WHERE mif_cencos = '" & Trim(fpText.text) & "' AND mif_codreg = " & RS!mif_codreg & " AND mif_codser = " & RS!mif_codser & "", vg_db, adOpenStatic
                 If Not RS1.EOF Then fecval = IIf(IsNull(RS1!fecval), 0, RS1!fecval)
                 RS1.Close: Set RS1 = Nothing
                 If fecval > 0 Then
                    '-------> Traer estructura fija
                    For i = 1 To Val(Mid(dEoM("26/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)), 1, 2))
                        If fecval <= Fecha & fg_pone_cero(Str(i), 2) Then
                           '-------> Grabar estructura fija día teorica
                           If vg_tipbase = "1" Then
                              vg_db.Execute "INSERT INTO b_minutafijadia (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro) SELECT a.mif_cencos, a.mif_codreg, a.mif_codser, " & Fecha & fg_pone_cero(i, 2) & ", b.pro_codigo, '1', a.mif_canpro, (SELECT DISTINCT ppd_propon FROM " & aAp & " WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") " & _
                                            "FROM b_minutafija a, b_productos b " & _
                                            "WHERE a.mif_codpro = b.pro_codigo " & _
                                            "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                            "AND   a.mif_cencos = '" & Trim(fpText.text) & "' " & _
                                            "AND   a.mif_codreg = " & RS!mif_codreg & " " & _
                                            "AND   a.mif_codser = " & RS!mif_codser & " " & _
                                            "AND   a.mif_fecval = " & fecval & " " & _
                                            "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                            "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0)"
                           Else
                              vg_db.Execute "INSERT INTO b_minutafijadia (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro) SELECT a.mif_cencos, a.mif_codreg, a.mif_codser, " & Fecha & fg_pone_cero(i, 2) & ", b.pro_codigo, '1', a.mif_canpro, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") " & _
                                            "FROM b_minutafija a, b_productos b " & _
                                            "WHERE a.mif_codpro = b.pro_codigo " & _
                                            "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                            "AND   a.mif_cencos = '" & Trim(fpText.text) & "' " & _
                                            "AND   a.mif_codreg = " & RS!mif_codreg & " " & _
                                            "AND   a.mif_codser = " & RS!mif_codser & " " & _
                                            "AND   a.mif_fecval = " & fecval & " " & _
                                            "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                            "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0)"
                           End If
                        End If
                        
                        Set RS1 = Nothing
                    Next i
                 End If
              Else
                 RS1.Close: Set RS1 = Nothing
                 '-------> Actualizar precio propon tabla estructura fija x día
                 If vg_tipbase = "1" Then
                    vg_db.Execute "UPDATE b_minutafijadia INNER JOIN " & aAp & " ON b_minutafijadia.mfd_codpro = " & aAp & ".ppd_codpro SET b_minutafijadia.mfd_cospro = " & aAp & ".ppd_propon " & _
                                  "WHERE b_minutafijadia.mfd_cencos = '" & Trim(fpText.text) & "' AND b_minutafijadia.mfd_codreg = " & RS!mif_codreg & " AND b_minutafijadia.mfd_codser = " & RS!mif_codser & " AND mid(b_minutafijadia.mfd_fecha,1,6) = " & Fecha & " AND b_minutafijadia.mfd_tipmin = '1' AND " & aAp & ".ppd_cencos = '" & MuestraCasino(1) & "'"
                 Else
                    vg_db.Execute "UPDATE b_minutafijadia SET b_minutafijadia.mfd_cospro = b_productospmpdia.ppd_propon FROM  b_productospmpdia WHERE b_minutafijadia.mfd_codpro = b_productospmpdia.ppd_codpro " & _
                                  "AND b_minutafijadia.mfd_cencos = '" & Trim(fpText.text) & "' AND b_minutafijadia.mfd_codreg = " & RS!mif_codreg & " AND b_minutafijadia.mfd_codser = " & RS!mif_codser & " AND convert(int,substring(convert(varchar(8),b_minutafijadia.mfd_fecha),1,6)) = " & Fecha & " AND b_minutafijadia.mfd_tipmin = '1' AND b_productospmpdia.ppd_cencos = '" & MuestraCasino(1) & "'"
                 End If
              End If
              RS.MoveNext
           Loop
        End If
        RS.Close: Set RS = Nothing
        '-------> Borrar tablas temporales
        If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
        vg_db.CommitTrans
        Toolbar1.Enabled = True
        Toolbar2.Enabled = True
        Frame1(1).Enabled = True
        .Enabled = True
    
    End With
    
    Toolbar1.Buttons(3).Enabled = True: Toolbar1.Buttons(5).Visible = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(6).Visible = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(8).Enabled = True
    fg_descarga

Case 3 '-------> Generar pedido
    
    If vg_pais <> "CL" And Trim(fpLongInteger1(0).text) = "" Then MsgBox "Debe ingresar semana SAC...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If vg_pais <> "CL" And Trim(fpDateTime2.text) = "" Then MsgBox "Debe seleccionar periodo SAC...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If vg_pais <> "CL" And Not ValidarCodInternoSac(LimpiaDato(Trim(fpText.text))) Then MsgBox "Codigo Interno SAC, no esta definido en casino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If vg_pais <> "CL" And Not ValidarCentralCompraSac(LimpiaDato(Trim(fpText.text))) Then MsgBox "Codigo central de compras, no esta definido casino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText.text)) & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda.Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '-------> Validar si existen datos en planificación teórica
    persac = Format(fpDateTime2.text, "yyyymm")
    nrosem = IIf(fpLongInteger1(0).Visible = True, fpLongInteger1(0).Value, 0)
    If vg_tipbase = "1" Then
       RS.Open "SELECT DISTINCT a.min_cencos FROM b_minuta a, a_servicio b WHERE a.min_codser = b.ser_codigo and b.ser_activo = '1' and a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND VAL(MID(a.min_fecmin,1,6)) = " & Fecha & "", vg_db, adOpenStatic
    Else
       RS.Open "SELECT DISTINCT a.min_cencos FROM b_minuta a, a_servicio b WHERE a.min_codser = b.ser_codigo and b.ser_activo = '1' and a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & Fecha & "", vg_db, adOpenStatic
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    fecdes = 0: fechasis = 0: fechasis = Val(Mid(Date, 7, 4)) & Mid(Date, 4, 2) & Mid(Date, 1, 2)
    If MsgBox("¿ Esta seguro generar pedido ?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    '-------> Definir vector costo recetas
    Dim vecrec As Variant
    fg_carga ""
    Toolbar1.Enabled = False
    Toolbar2.Enabled = False
    Frame1(1).Enabled = False
    
    With vaSpread1
        
        .Enabled = False
        vg_db.BeginTrans
        If vg_tipbase = "1" Then
           '-------> Insert tabla productospmpdia
           aAp = Trim(vg_NUsr) & "_tmp_ProductoPMPGenPed3"
           fg_CheckTmp aAp
           vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                         "INTO " & aAp & " " & _
                         "FROM b_productospmpdia " & _
                         "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                         "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                         "AND   ppd_propon > 0 " & _
                         "GROUP BY ppd_cencos, ppd_codpro"
           vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
           vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon"
           vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
        End If
        '-------> Actualizar fecha envio minuta pedido
        vg_db.Execute "UPDATE b_minutapedidos SET ped_fecenv = " & fechasis & " WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " AND ped_tipped = 1"
        '-------> Grabar minuta costo teórico & real
        sql1 = IIf(vg_tipbase = "1", " AND val(mid(b.min_fecmin,1,6)) = " & Fecha & " ", " AND convert(int,substring(convert(varchar(8),b.min_fecmin),1,6)) = " & Fecha & " ")
        vg_db.Execute "DELETE b_minutacosto FROM b_minutacosto WHERE mic_cencos = '" & MuestraCasino(1) & "' AND mic_fecval IN (SELECT a.mid_fecval FROM b_minutadet a, b_minuta b WHERE a.mid_codigo = b.min_codigo AND b.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & sql1 & ") AND mic_codpro = (SELECT top 1 d.red_codpro FROM b_minutadet a, b_minuta b, b_receta c, b_recetadet d WHERE a.mid_codigo = b.min_codigo AND a.mid_codrec = c.rec_codigo AND c.rec_codigo = d.red_codigo AND a.mid_codrec = d.red_codigo AND a.mid_tiprec = d.red_tiprec AND ((d.red_tiprec<>0 AND d.red_cencos = '" & MuestraCasino(1) & "') OR (d.red_tiprec = 0 AND d.red_cencos = '0')) AND b.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & sql1 & ")"
        Dim tmin As String
        i = 1
        tmin = "1"
        sql1 = IIf(vg_tipbase = "1", " AND val(mid(a.min_fecmin,1,6)) = " & Fecha & " ", " AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & Fecha & " ")
        vg_db.Execute "INSERT INTO b_minutacosto (mic_cencos, mic_fecval, mic_tipmin, mic_codpro, mic_cospro) SELECT DISTINCT a.min_cencos, " & fechasis & ", '" & tmin & "', d.red_codpro, (SELECT top 1 cpi_precos FROM b_contlistpreing WHERE cpi_cencos = '" & MuestraCasino(1) & "' AND cpi_coding = d.red_codpro) AS cpi_precos " & _
                      "FROM b_minuta a, b_minutadet b, b_receta c, b_recetadet d, a_servicio e " & _
                      "WHERE a.min_codigo = b.mid_codigo " & _
                      "AND   b.mid_codrec = c.rec_codigo " & _
                      "AND   c.rec_codigo = d.red_codigo " & _
                      "AND   a.min_cencos = '" & MuestraCasino(1) & "' AND b.mid_tipmin = '1' and a.min_codser = e.ser_codigo and e.ser_activo = '1' " & _
                      " " & sql1 & " AND d.red_codpro not in (SELECT top 1 mic_codpro from b_minutacosto where mic_fecval = " & fechasis & " and mic_cencos = a.min_cencos AND mic_tipmin = '" & tmin & "' AND mic_codpro = d.red_codpro)"
        tmin = "2"
        vg_db.Execute "INSERT INTO b_minutacosto (mic_cencos, mic_fecval, mic_tipmin, mic_codpro, mic_cospro) SELECT DISTINCT a.min_cencos, " & fechasis & ", '" & tmin & "', d.red_codpro, (SELECT top 1 cpi_precos FROM b_contlistpreing WHERE cpi_cencos = '" & MuestraCasino(1) & "' AND cpi_coding = d.red_codpro) AS cpi_precos " & _
                      "FROM b_minuta a, b_minutadet b, b_receta c, b_recetadet d , a_servicio e " & _
                      "WHERE a.min_codigo = b.mid_codigo " & _
                      "AND   b.mid_codrec = c.rec_codigo " & _
                      "AND   c.rec_codigo = d.red_codigo " & _
                      "AND   a.min_cencos = '" & MuestraCasino(1) & "' AND b.mid_tipmin = '1' and a.min_codser = e.ser_codigo and e.ser_activo = '1' " & _
                      " " & sql1 & " AND d.red_codpro not in (SELECT top 1 mic_codpro from b_minutacosto where mic_fecval = " & fechasis & " and mic_cencos = a.min_cencos AND mic_tipmin = '" & tmin & "' AND mic_codpro = d.red_codpro)"
        If vg_tipbase = "1" Then
           vg_db.Execute "UPDATE b_minutacosto INNER JOIN b_contlistpreing ON (b_contlistpreing.cpi_coding = b_minutacosto.mic_codpro) AND (b_minutacosto.mic_cencos = b_contlistpreing.cpi_cencos) SET b_minutacosto.mic_cospro = b_contlistpreing.cpi_precos " & _
                         "Where b_minutacosto.mic_cencos = '" & MuestraCasino(1) & "' And b_minutacosto.mic_fecval = " & fechasis & " And b_minutacosto.mic_tipmin IN ('1','2')"
        Else
           vg_db.Execute "UPDATE b_minutacosto SET b_minutacosto.mic_cospro = b_contlistpreing.cpi_precos FROM b_minutacosto, b_contlistpreing " & _
                         "Where b_contlistpreing.cpi_coding = b_minutacosto.mic_codpro AND b_minutacosto.mic_cencos = b_contlistpreing.cpi_cencos AND  b_minutacosto.mic_cencos = '" & MuestraCasino(1) & "' And b_minutacosto.mic_fecval = " & fechasis & " And b_minutacosto.mic_tipmin IN ('1','2')"
        End If
        vg_db.Execute "UPDATE b_minutacosto SET mic_cospro = 0 WHERE mic_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND mic_fecval = " & fechasis & " AND mic_cospro IS NULL"
        '-------> Generar minuta costo estructura fija
        '-------> Eliminar estructura fija día real si existen datos
        If vg_tipbase = "1" Then
           vg_db.Execute "DELETE b_minutafijadia FROM b_minutafijadia WHERE mfd_cencos = '" & Trim(fpText.text) & "' AND mid(mfd_fecha,1,6) = " & Fecha & " AND mfd_tipmin = '2'"
           '-------> Grabar estructura fija día real
           vg_db.Execute "INSERT INTO b_minutafijadia (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro) SELECT a.mfd_cencos, a.mfd_codreg, a.mfd_codser, a.mfd_fecha, a.mfd_codpro, '2', a.mfd_canpro, (SELECT ppd_propon FROM " & aAp & " WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") " & _
                         "FROM b_minutafijadia a, b_productos b, a_servicio c " & _
                         "WHERE   a.mfd_codpro = b.pro_codigo " & _
                         "AND    (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                         "AND     a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                         "AND mid(a.mfd_fecha,1,6) = " & Fecha & " " & _
                         "AND     a.mfd_tipmin = '1' " & _
                         "AND    (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) and a.mfd_codser = c.ser_codigo and c.ser_activo = '1'"
        Else
           vg_db.Execute "DELETE b_minutafijadia FROM b_minutafijadia WHERE mfd_cencos = '" & Trim(fpText.text) & "' AND convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) = " & Fecha & " AND mfd_tipmin = '2'"
           '-------> Grabar estructura fija día real
           vg_db.Execute "INSERT INTO b_minutafijadia (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro) SELECT a.mfd_cencos, a.mfd_codreg, a.mfd_codser, a.mfd_fecha, a.mfd_codpro, '2', a.mfd_canpro, (SELECT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") " & _
                         "FROM b_minutafijadia a, b_productos b, a_servicio c  " & _
                         "WHERE a.mfd_codpro = b.pro_codigo " & _
                         "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                         "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                         "AND   convert(int,substring(convert(varchar(8),a.mfd_fecha),1,6)) = " & Fecha & " " & _
                         "AND   a.mfd_tipmin = '1' " & _
                         "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) and a.mfd_codser = c.ser_codigo and c.ser_activo = '1'"
        End If
        '-------> Traer total de receta desde planificación de minutas y luego calcular costo
        If vg_tipbase = "1" Then
           RS.Open "SELECT COUNT(b.mid_codrec) AS nreg FROM b_minuta a, b_minutadet b, a_servicio c WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                   "AND val(mid(a.min_fecmin,1,6)) = " & Fecha & " AND b.mid_tipmin = '1' and a.min_codser = c.ser_codigo and c.ser_activo = '1'", vg_db, adOpenStatic
        Else
           RS.Open "SELECT COUNT(b.mid_codrec) AS nreg FROM b_minuta a, b_minutadet b, a_servicio c WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                   "AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & Fecha & " AND b.mid_tipmin = '1' and a.min_codser = c.ser_codigo and c.ser_activo = '1' ", vg_db, adOpenStatic
        End If
        If RS.EOF Or RS!nreg < 1 Then RS.Close: Set RS = Nothing: vg_db.RollbackTrans: MsgBox "No existe información", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
        ReDim vecrec(RS!nreg, 4)
        RS.Close: Set RS = Nothing
        For i = 1 To UBound(vecrec)
            DoEvents
            vecrec(i, 1) = 0 '-------> codigo receta
            vecrec(i, 2) = 0 '-------> tipo receta
            vecrec(i, 3) = 0 '-------> costo receta alimentación
            vecrec(i, 4) = 0 '-------> costo receta desechable
        Next i
        i = 1
        If vg_tipbase = "1" Then
           RS.Open "SELECT DISTINCT b.mid_codrec, b.mid_tiprec FROM b_minuta a, b_minutadet b, a_servicio c WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                   "AND val(mid(a.min_fecmin,1,6)) = " & Fecha & " AND b.mid_tipmin = '1' and a.min_codser = c.ser_codigo and c.ser_activo = '1'  ", vg_db, adOpenStatic
        Else
           RS.Open "SELECT DISTINCT b.mid_codrec, b.mid_tiprec FROM b_minuta a, b_minutadet b, a_servicio c WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                   "AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & Fecha & " AND b.mid_tipmin = '1' and a.min_codser = c.ser_codigo and c.ser_activo = '1'  ", vg_db, adOpenStatic
        End If
        If RS.EOF Then RS.Close: Set RS = Nothing: vg_db.RollbackTrans: MsgBox "No existe información", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
        Do While Not RS.EOF
           DoEvents
           vecrec(i, 1) = RS!mid_codrec
           vecrec(i, 2) = RS!mid_tiprec
           vecrec(i, 3) = Format(fg_CalCtoRecPlan(fechasis, 1, RS!mid_codrec, IIf(IsNull(RS!mid_tiprec), 0, RS!mid_tiprec), (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','"))), fg_Pict(6, 2))
           vecrec(i, 4) = Format(fg_CalCtoRecPlan(fechasis, 1, RS!mid_codrec, IIf(IsNull(RS!mid_tiprec), 0, RS!mid_tiprec), (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))), fg_Pict(6, 2))
           RS.MoveNext: i = i + 1
        Loop
        RS.Close: Set RS = Nothing
        
        '-------> Generar planificación real & actualizar costo teórica
        If vg_tipbase = "1" Then
           RS.Open "SELECT b.* FROM b_minuta a, b_minutadet b, a_servicio c WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                   "AND val(mid(a.min_fecmin,1,6)) = " & Fecha & " AND b.mid_tipmin = '1'and a.min_codser = c.ser_codigo and c.ser_activo = '1'  ", vg_db, adOpenStatic
        Else
           RS.Open "SELECT b.* FROM b_minuta a, b_minutadet b, a_servicio c WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                   "AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & Fecha & " AND b.mid_tipmin = '1' and a.min_codser = c.ser_codigo and c.ser_activo = '1' ", vg_db, adOpenStatic
        End If
        If RS.EOF Then RS.Close: Set RS = Nothing: vg_db.RollbackTrans: MsgBox "No existe información", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
        
        Set RS1 = vg_db.Execute("select 1 from b_minutadet as bmd WITH (NOLOCK) inner join b_minuta as bmd2 WITH (NOLOCK) on bmd.mid_codigo = bmd2.min_codigo where bmd2.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),bmd2.min_fecmin),1,6)) = " & Fecha & " AND bmd.mid_tipmin = '2'")
        If Not RS1.EOF Then
           vg_db.Execute ("DELETE b_minutadet FROM b_minuta a INNER JOIN b_minutadet b ON a.min_codigo = b.mid_codigo WHERE a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & Fecha & " AND b.mid_tipmin = '2'")
        End If
        RS1.Close: Set RS1 = Nothing
        
        Do While Not RS.EOF
           DoEvents
           For i = 1 To UBound(vecrec)
               If RS!mid_codrec = vecrec(i, 1) And RS!mid_tiprec = vecrec(i, 2) Then
                  cosali = CCur(vecrec(i, 3))
                  CosDes = CCur(vecrec(i, 4))
                  Exit For
               End If
           Next
           vg_db.Execute "INSERT INTO b_minutadet (mid_codigo, mid_tipmin, mid_numlin, mid_estser, mid_codrec, mid_numrac, mid_descri, mid_cosrec, mid_fecval, mid_tiprec, mid_nummer, mid_rec5eta, mid_cosdes) " & _
                         "VALUES (" & RS!mid_codigo & ", '2', " & RS!mid_numlin & ", " & IIf(IsNull(RS!mid_estser), 0, RS!mid_estser) & ", " & RS!mid_codrec & ", " & IIf(IsNull(RS!mid_numrac), "NULL", RS!mid_numrac) & ", '" & RS!mid_descri & "', " & cosali & ", " & fechasis & ", " & RS!mid_tiprec & ", 0, " & IIf(IsNull(RS!mid_rec5eta) Or Trim(RS!mid_rec5eta) = "", "Null", RS!mid_rec5eta) & ", " & CosDes & ")"
           
           vg_db.Execute "UPDATE b_minutadet SET mid_fecval = " & fechasis & ", mid_cosrec = " & cosali & ", mid_cosdes = " & CosDes & " WHERE mid_codigo = " & RS!mid_codigo & " AND mid_tipmin = '1' AND mid_codrec = " & RS!mid_codrec & " AND mid_tiprec = " & RS!mid_tiprec & ""
           
           RS.MoveNext
        Loop
        RS.Close: Set RS = Nothing
        '-------> Bloquear planificación teórica
        If vg_tipbase = "1" Then
'           vg_db.Execute "UPDATE b_minuta SET min_indblo = 1, min_racrea = min_racteo WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(min_fecmin,1,6)) = " & Fecha & " AND (min_indblo = 0 OR (min_indblo) IS NULL)"
           vg_db.Execute "UPDATE b_minuta SET min_indblo = 1, min_racrea = 0 WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(min_fecmin,1,6)) = " & Fecha & " AND (min_indblo = 0 OR (min_indblo) IS NULL)"
           RS.Open "SELECT DISTINCT a.min_codreg, a.min_codser, a.min_fecmin, a.min_racrea FROM b_minuta a, a_servicio b WHERE a.min_codser = b.ser_codigo and b.ser_activo = '1' and a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(a.min_fecmin,1,6)) = " & Fecha & " AND a.min_racrea > 0", vg_db, adOpenStatic
        Else
'           vg_db.Execute "UPDATE b_minuta SET min_indblo = 1, min_racrea = min_racteo WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),min_fecmin),1,6)) = " & Fecha & " AND (min_indblo = 0 OR (min_indblo) IS NULL)"
           vg_db.Execute "UPDATE b_minuta SET min_indblo = 1, min_racrea = 0 WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),min_fecmin),1,6)) = " & Fecha & " AND (min_indblo = 0 OR (min_indblo) IS NULL or min_indblo = 1)"
           RS.Open "SELECT DISTINCT a.min_codreg, a.min_codser, a.min_fecmin, a.min_racrea FROM b_minuta a, a_servicio b WHERE a.min_codser = b.ser_codigo and b.ser_activo = '1' and a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & Fecha & "", vg_db, adOpenStatic ' AND a.min_racrea > 0", vg_db, adOpenStatic
        End If
        '-------> Grabar raciones en minutas raciones
        Do While Not RS.EOF
           DoEvents
           RS1.Open "SELECT * FROM b_minutaraciones WHERE mir_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND mir_codreg = " & RS!min_codreg & " AND mir_codser = " & RS!min_codser & " AND mir_fecmin = " & RS!min_fecmin & " AND mir_rutcli = 'PRODUCIDAS'", vg_db, adOpenStatic
           If RS1.EOF Then
'              vg_db.Execute "INSERT INTO b_minutaraciones VALUES ('" & LimpiaDato(Trim(fpText.text)) & "', " & RS!min_codreg & ", " & RS!min_codser & ", " & RS!min_fecmin & ", 'PRODUCIDAS', " & RS!min_racrea & ", NULL, '')"
              vg_db.Execute "INSERT INTO b_minutaraciones VALUES ('" & LimpiaDato(Trim(fpText.text)) & "', " & RS!min_codreg & ", " & RS!min_codser & ", " & RS!min_fecmin & ", 'PRODUCIDAS', 0 , NULL, '')"
           Else
'              vg_db.Execute "UPDATE b_minutaraciones SET mir_nrorac = " & RS!min_racrea & " WHERE mir_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND mir_codreg = " & RS!min_codreg & " AND mir_codser = " & RS!min_codser & " AND mir_fecmin = " & RS!min_fecmin & " AND mir_rutcli = 'PRODUCIDAS' AND mir_nrorac < 1"
              vg_db.Execute "UPDATE b_minutaraciones SET mir_nrorac =  0 WHERE mir_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND mir_codreg = " & RS!min_codreg & " AND mir_codser = " & RS!min_codser & " AND mir_fecmin = " & RS!min_fecmin & " AND mir_rutcli = 'PRODUCIDAS' " 'AND mir_nrorac < 1"
           End If
           RS1.Close: Set RS1 = Nothing
           RS.MoveNext
        Loop
        RS.Close: Set RS = Nothing
        vg_db.CommitTrans
        If Trim(aAp) <> "" And vg_tipbase = "1" Then vg_db.Execute "DROP TABLE " & aAp & ""
        If vg_pais = "CO" Then GenerarArchivoMdb
        fg_descarga
        Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(8).Enabled = True: Frame2.Enabled = False
        MsgBox "Generación pedido Finalizado Sin Problema", vbInformation + vbOKOnly, Msgtitulo
        I_PedidosNew LimpiaDato(Trim(fpText.text)), Fecha, 0
        Toolbar1.Enabled = True
        Frame1(1).Enabled = True
        .Enabled = True
    End With
Case 5 '-------> Borrar pedido
    If CierrePeriodo(Fecha, vg_codbod, 10) Then MsgBox "Existen documentos realizados, en la salida producción. Proceso cancelado", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina pedido...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    Toolbar1.Enabled = False
    vg_db.BeginTrans
    '-------> Eliminar minutapedido
    vg_db.Execute "DELETE b_minutapedidos FROM b_minutapedidos WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " AND ped_tipped = 1"
    If vg_tipbase = "1" Then
       '-------> Eliminar minutacosto
       vg_db.Execute "DELETE b_minutacosto FROM b_minutacosto WHERE mic_cencos = '" & MuestraCasino(1) & "' AND mic_fecval IN (SELECT a.mid_fecval FROM b_minutadet a, b_minuta b WHERE a.mid_codigo = b.min_codigo AND b.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(b.min_fecmin,1,6)) = " & Fecha & ") AND mic_codpro = (SELECT top 1 d.red_codpro FROM b_minutadet a, b_minuta b, b_receta c, b_recetadet d WHERE a.mid_codigo = b.min_codigo AND a.mid_codrec = c.rec_codigo AND c.rec_codigo = d.red_codigo AND a.mid_codrec = d.red_codigo AND a.mid_tiprec = d.red_tiprec AND ((d.red_tiprec<>0 AND d.red_cencos = '" & MuestraCasino(1) & "') OR (d.red_tiprec = 0 AND d.red_cencos = '0')) AND b.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(b.min_fecmin,1,6)) = " & Fecha & ")"
       '-------> Eliminar minutas real
       vg_db.Execute "DELETE b_minutadet FROM b_minutadet WHERE mid_codigo IN (SELECT min_codigo FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(min_fecmin,1,6)) = " & Fecha & ") AND mid_tipmin = '2'"
       '-------> Desbloquear planificación teórica
       vg_db.Execute "UPDATE b_minuta SET min_indblo = 0, min_racrea = 0 WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(min_fecmin,1,6)) = " & Fecha & " AND min_indblo = 1"
       '-------> Actualizar detalle planificación teórica al campo fecval
       vg_db.Execute "UPDATE b_minuta INNER JOIN b_minutadet ON b_minuta.min_codigo = b_minutadet.mid_codigo SET b_minutadet.mid_fecval = 0 " & _
                     "WHERE b_minutadet.mid_tipmin = '1' AND b_minuta.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(b_minuta.min_fecmin,1,6)) = " & Fecha & ""
       '-------> Eliminar estructura fija día real
       vg_db.Execute "DELETE b_minutafijadia FROM b_minutafijadia WHERE mfd_cencos = '" & Trim(fpText.text) & "' AND mid(mfd_fecha,1,6) = " & Fecha & " AND mfd_tipmin = '2'"
    Else
       '-------> Eliminar minutacosto
       vg_db.Execute "DELETE b_minutacosto FROM b_minutacosto WHERE mic_cencos = '" & MuestraCasino(1) & "' AND mic_fecval IN (SELECT a.mid_fecval FROM b_minutadet a, b_minuta b WHERE a.mid_codigo = b.min_codigo AND b.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),b.min_fecmin),1,6)) = " & Fecha & ")"
       '-------> Eliminar minutas real
       vg_db.Execute "DELETE b_minutadet FROM b_minutadet WHERE mid_codigo IN (SELECT min_codigo FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),min_fecmin),1,6)) = " & Fecha & ") AND mid_tipmin = '2'"
       '-------> Desbloquear planificación teórica
       vg_db.Execute "UPDATE b_minuta SET min_indblo = 0, min_racrea = 0 WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),min_fecmin),1,6)) = " & Fecha & " AND min_indblo = 1"
       '-------> Actualizar detalle planificación teórica al campo fecval
       vg_db.Execute "UPDATE b_minutadet SET b_minutadet.mid_fecval = 0 FROM b_minutadet, b_minuta WHERE b_minuta.min_codigo = b_minutadet.mid_codigo " & _
                     "AND b_minutadet.mid_tipmin = '1' AND b_minuta.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),b_minuta.min_fecmin),1,6)) = " & Fecha & ""
       '-------> Eliminar estructura fija día real
       vg_db.Execute "DELETE b_minutafijadia FROM b_minutafijadia WHERE mfd_cencos = '" & Trim(fpText.text) & "' AND convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) = " & Fecha & " AND mfd_tipmin = '2'"
    End If
    vaSpread1.MaxRows = 0
    Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = True
    vg_db.CommitTrans
    Toolbar1.Enabled = True
Case 8 '-------> Imprimir
    If vaSpread1.MaxRows < 1 Then Exit Sub
    RS.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText.text)) & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda.Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '-------> Validar si existen datos en planificación teórica
    RS.Open "SELECT DISTINCT ped_codcas FROM b_minutapedidos WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " AND ped_tipped = 1", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    I_PedidosNew LimpiaDato(Trim(fpText.text)), Fecha, 0 'IIf(Check1(0).Value = 1, IIf(etapa5, 0, 1), 0)
Case 10 '-------> Imprimir Productos no Asociado a Sac
    I_ProductosNoAsociadoSac
Case 12 '-------> Cerrar
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
Toolbar1.Enabled = True
Frame1(1).Enabled = True
Toolbar2.Enabled = True
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = -2147217900 Then vg_db.RollbackTrans: MsgBox "Error datos duplicados, Comuniquese con deparatemnto de informatica de la región...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Dim codpro As String, aux_codigo As String
Dim necmin As Double, cansol As Double, canres As Double, canmin As Double, faccon As Double, facsto As Double, canrea As Double
Dim stoact As Double, ordrec As Double, conmiras As Double, stodia As Double, cantid  As Double, conmir As Double

Dim i As Long, X As Long, j As Long
Select Case Button.Index
Case 1
    vg_nombre = "": vg_codigo = ""
    vg_left = fpayuda.Left + 2300
    If vg_pais = "CO" Then
       B_TabEst.LlenaDatos "b_formatocompras", "foc_", "Productos SAC", "PSAC"
    Else
       B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Pst"
    End If
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    Dim codtip As String, tipdes As String, coding As String, nomuni As String
    Dim proc1 As String, proc2 As String, sql1 As String
    Dim diaseg As Long
    Dim fecini As Date, fecfin As Date
    sql1 = " "
    If vg_pais = "CO" Then
       RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, c.pad_codigo, c.pad_tipo, f.foc_unisac, f.foc_faccon " & _
               "FROM b_productos a, a_tipopro b, b_paramdesp c, " & aAp1 & " d, b_formatocomprassgp e, b_formatocompras f " & _
               "WHERE a.pro_codtip = b.tip_codigo " & _
               "AND   b.tip_codigo = d.pro_codtip " & _
               "AND   d.pro_previo = c.pad_codigo " & _
               "AND   c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
               "AND   e.fcs_codsac = '" & vg_codigo & "' " & _
               "AND   a.pro_facing > 0 AND a.pro_facsto > 0 AND a.pro_codigo = e.fcs_codsgp AND e.fcs_codsac = f.foc_codsac", vg_db, adOpenStatic
    Else
       RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, c.pad_codigo, c.pad_tipo, 1 AS foc_faccon " & _
               "FROM b_productos a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
               "WHERE a.pro_codtip = b.tip_codigo " & _
               "AND   b.tip_codigo = d.pro_codtip " & _
               "AND   d.pro_previo = c.pad_codigo " & _
               "AND   c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
               "AND   a.pro_codigo = '" & vg_codigo & "' " & _
               "AND   a.pro_facing > 0 AND a.pro_facsto > 0", vg_db, adOpenStatic
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Producto no tiene asignado los factores", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    '-------> Validar si existe producto en grilla
    If vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone) <> -1 Then
       RS.Close: Set RS = Nothing: MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    End If
    Toolbar1.Buttons(3).Enabled = False
    
    With vaSpread1
    
        .Row = .ActiveRow
        codpro = RS!pro_codigo
        codtip = RS!pad_codigo
        tipdes = RS!pad_tipo
        faccon = RS!foc_faccon
        If vg_pais = "CO" Then nomuni = RS!foc_unisac Else nomuni = ""
        RS.Close: Set RS = Nothing
        '-------> Validar si existe familia productos
        est = True
        If .SearchCol(16, 0, .MaxRows, codtip, SearchFlagsNone) <> -1 Then
           For i = .SearchCol(16, 0, .MaxRows, codtip, SearchFlagsNone) + 1 To .MaxRows
               .Row = i
               .Col = 16
               If Trim(.text) <> "" Then
                  Exit For
               End If
           Next i
           X = i
        Else
           .MaxRows = vaSpread1.MaxRows + 1
           .Row = vaSpread1.MaxRows
           .Col = -1: .BackColor = &HFFFFC0
           .Col = 2: .Font.Bold = True: .CellType = CellTypeStaticText: .text = fg_BuscaenArbol(Val(codtip), "a_tipopro", "tip_codigo")
           .Col = 3: .CellType = CellTypeStaticText: .text = ""
           .Col = 4: .CellType = CellTypeStaticText: .text = ""
           .Col = 5: .CellType = CellTypeStaticText: .text = ""
           .Col = 6: .CellType = CellTypeStaticText: .text = ""
           .Col = 7: .CellType = CellTypeStaticText: .text = ""
           .Col = 8: .CellType = CellTypeStaticText: .text = ""
           .Col = 9: .CellType = CellTypeStaticText: .text = ""
           .Col = 10: .CellType = CellTypeStaticText: .text = ""
           .Col = 11: .CellType = CellTypeStaticText: .text = ""
           .Col = 12: .CellType = CellTypeStaticText: .text = ""
           .Col = 13: .CellType = CellTypeStaticText: .text = ""
           .Col = 14: .CellType = CellTypeStaticText: .text = ""
           .Col = 15: .CellType = CellTypeStaticText: .text = ""
           .Col = 16: .CellType = CellTypeStaticText: .text = Val(codtip)
           X = .MaxRows + 1
        End If
        '-------> validar si existe mas de un ingrediente
        RS.Open "SELECT COUNT(pri_coding) AS nreg FROM b_productosing WHERE pri_codpro = '" & codpro & "'", vg_db, adOpenStatic
        If RS.EOF Or IsNull(RS!nreg) Or RS!nreg = 0 Then RS.Close: Set RS = Nothing: Exit Sub
        If RS!nreg > 1 Then
           aux_codigo = vg_codigo
           vg_nombre = ""
           vg_left = fpayuda.Left + 2300
           B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "Proing"
           SendKeys "+{Tab}"
           B_TabEst.Show 1
           If vg_codigo = "" Then RS.Close: Set RS = Nothing: Exit Sub
           coding = vg_codigo
           vg_codigo = aux_codigo
           proc2 = " AND  (c.ing_codigo = '" & coding & "')"
        End If
        RS.Close: Set RS = Nothing
        '-------> Agregar Producto
        proc1 = "SELECT a.pro_codigo, a.pro_nombre, a.pro_facsto, " & _
                "c.ing_codigo, (SELECT DISTINCT cpi_precos FROM b_contlistpreing WHERE cpi_coding = c.ing_codigo AND cpi_cencos = '" & MuestraCasino(1) & "') AS cpi_precos, c.ing_nombre, d.uni_nomcor " & _
                "FROM  b_productos a, b_productosing b, b_ingrediente c, a_unidad d " & _
                "WHERE a.pro_codigo = b.pri_codpro " & _
                "AND   c.ing_codigo = b.pri_coding " & _
                "AND   a.pro_coduni = d.uni_codigo " & _
                "AND   a.pro_codigo = '" & codpro & "'"
        RS.Open proc1 & proc2, vg_db, adOpenStatic
        If Not RS.EOF Then
           .MaxRows = vaSpread1.MaxRows + 1
           .InsertRows X, 1
           .Row = X
           .Col = 1: .CellType = CellTypeStaticText: .text = Trim(vg_codigo)
           .Col = 2: .CellType = CellTypeStaticText: .text = Trim(vg_nombre)
           .Col = 3: .CellType = CellTypeStaticText: .text = IIf(vg_pais = "CO", Trim(nomuni), Trim(RS!uni_nomcor))
           .Col = 4: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = 0
           .Col = 5: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = 0
           .Col = 6: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = faccon '"1"
           .Col = 7: .CellType = CellTypeStaticText: .text = ""
           .Col = 8: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = 0
           .Col = 9: .CellType = CellTypeStaticText: .text = Trim(RS!ing_codigo)
           .Col = 10: .CellType = CellTypeStaticText: .text = Trim(RS!ing_nombre)
           .Col = 11: .CellType = CellTypeStaticText: .text = Trim(RS!pro_codigo)
           .Col = 12: .CellType = CellTypeStaticText: .text = Trim(RS!pro_nombre)
           .Col = 13: .CellType = CellTypeStaticText: .text = ""
           .Col = 14: .CellType = CellTypeStaticText: .text = ""
           .Col = 15: .CellType = CellTypeStaticText: .text = ""
           .Col = 16: .CellType = CellTypeStaticText: .text = ""
           conmir = 0: stoact = 0: ordrec = 0
           If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone) <> -1 Then
              vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone)
              vaSpread2.Col = 2: .Col = 13: .CellType = CellTypeStaticText: .text = vaSpread2.text: stoact = Val(vaSpread2.text)
              vaSpread2.Col = 4: .Col = 14: .CellType = CellTypeStaticText: .text = vaSpread2.text: conmir = Val(vaSpread2.text)
              vaSpread2.Col = 5: .Col = 15: .CellType = CellTypeStaticText: .text = vaSpread2.text: ordrec = Val(vaSpread2.text)
           End If
           necmin = IIf((conmir - stoact - ordrec) < 0, (conmir - stoact - ordrec) * -1, (conmir - stoact - ordrec))
           Select Case Trim(tipdes)
           Case "M"
                RS1.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'M')", vg_db, adOpenStatic
                If Not RS1.EOF Then
                   fecxin = Format(CDate("01/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
                   fecxfi = Format(CDate(dEoM("15/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
                   X = X + 1
                   .MaxRows = .MaxRows + 1
                   .InsertRows X, 1
                   .Row = X
                   .Col = 7: .CellType = CellTypeStaticText: .text = fg_Ctod1(fecxin)
                   fg_FormatearCeldaGrillaNuemrica vaSpread1, 8
                   .Col = 11: .CellType = CellTypeStaticText: .text = codpro
                   .SetActiveCell 8, X
                End If
                RS1.Close: Set RS1 = Nothing
           Case "Q1"
                '-------> actualizar fecha pedido quincenal 1-15
                RS1.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'Q1')", vg_db, adOpenStatic
                If Not RS1.EOF Then
                   fecxin = Format(CDate("01/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
                   fecxfi = Format(CDate("15/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
                   X = X + 1
                   .MaxRows = .MaxRows + 1
                   .InsertRows X, 1
                   .Row = X
                   .Col = 7: .CellType = CellTypeStaticText: .text = fg_Ctod1(fecxin)
                   fg_FormatearCeldaGrillaNuemrica vaSpread1, 8
                   .Col = 11: .CellType = CellTypeStaticText: .text = codpro
                   .SetActiveCell 8, X
                   
                   fecxin = Format(CDate("15/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
                   fecxfi = Format(CDate(dEoM("15/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
                   X = X + 1
                   .MaxRows = .MaxRows + 1
                   .InsertRows X, 1
                   .Row = X
                   .Col = 7: .CellType = CellTypeStaticText: .text = fg_Ctod1(fecxin)
                   .Col = 11: .CellType = CellTypeStaticText: .text = codpro
                   fg_FormatearCeldaGrillaNuemrica vaSpread1, 8
                   .SetActiveCell 8, X
                End If
                RS1.Close: Set RS1 = Nothing
           Case "Q2"
                '-------> actualizar fecha pedido quincenal 2-16
                RS1.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'Q2')", vg_db, adOpenStatic
                If Not RS1.EOF Then
                   fecxin = Format(CDate("02/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
                   fecxfi = Format(CDate("16/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
                   X = X + 1
                   .MaxRows = .MaxRows + 1
                   .InsertRows X, 1
                   .Row = X
                   .Col = 7: .CellType = CellTypeStaticText: .text = fg_Ctod1(fecxin)
                   fg_FormatearCeldaGrillaNuemrica vaSpread1, 8
                   .Col = 11: .CellType = CellTypeStaticText: .text = codpro
                   .SetActiveCell 8, X
                   
                   fecxin = Format(CDate("16/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
                   fecxfi = Format(CDate(dEoM("16/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
                   X = X + 1
                   .MaxRows = .MaxRows + 1
                   .InsertRows X, 1
                   .Row = X
                   .Col = 7: .CellType = CellTypeStaticText: .text = fg_Ctod1(fecxin)
                   fg_FormatearCeldaGrillaNuemrica vaSpread1, 8
                   .Col = 11: .CellType = CellTypeStaticText: .text = codpro
                   .SetActiveCell 8, X
                End If
                RS1.Close: Set RS1 = Nothing
           Case "Q3"
                '-------> actualizar fecha pedido quincenal 3-17
                RS1.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'Q3')", vg_db, adOpenStatic
                If Not RS1.EOF Then
                   fecxin = Format(CDate("03/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
                   fecxfi = Format(CDate("17/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
                   X = X + 1
                   .MaxRows = .MaxRows + 1
                   .InsertRows X, 1
                   .Row = X
                   .Col = 7: .CellType = CellTypeStaticText: .text = fg_Ctod1(fecxin)
                   fg_FormatearCeldaGrillaNuemrica vaSpread1, 8
                   .Col = 11: .CellType = CellTypeStaticText: .text = codpro
                   .SetActiveCell 8, X
                   
                   fecxin = Format(CDate("17/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
                   fecxfi = Format(CDate(dEoM("17/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
                   X = X + 1
                   .MaxRows = .MaxRows + 1
                   .InsertRows X, 1
                   .Row = X
                   .Col = 7: .CellType = CellTypeStaticText: .text = fg_Ctod1(fecxin)
                   fg_FormatearCeldaGrillaNuemrica vaSpread1, 8
                   .Col = 11: .CellType = CellTypeStaticText: .text = codpro
                   .SetActiveCell 8, X
                End If
                RS1.Close: Set RS1 = Nothing
           Case "Q4"
                '-------> actualizar fecha pedido quincenal 4-18
                RS1.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'Q4')", vg_db, adOpenStatic
                If Not RS1.EOF Then
                   fecxin = Format(CDate("04/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
                   fecxfi = Format(CDate("18/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
                   X = X + 1
                   .MaxRows = .MaxRows + 1
                   .InsertRows X, 1
                   .Row = X
                   .Col = 7: .CellType = CellTypeStaticText: .text = fg_Ctod1(fecxin)
                   fg_FormatearCeldaGrillaNuemrica vaSpread1, 8
                   .Col = 11: .CellType = CellTypeStaticText: .text = codpro
                   .SetActiveCell 8, X
                   
                   fecxin = Format(CDate("18/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
                   fecxfi = Format(CDate(dEoM("18/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                   fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
                   X = X + 1
                   .MaxRows = .MaxRows + 1
                   .InsertRows X, 1
                   .Row = X
                   .Col = 7: .CellType = CellTypeStaticText: .text = fg_Ctod1(fecxin)
                   fg_FormatearCeldaGrillaNuemrica vaSpread1, 8
                   .Col = 11: .CellType = CellTypeStaticText: .text = codpro
                   .SetActiveCell 8, X
                End If
                RS1.Close: Set RS1 = Nothing
           Case "E", "S" 'Semanal y diario
                RS1.Open "SELECT pad_codigo, pad_diario, pad_diaseg FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND (pad_tipo = 'S' or pad_tipo = 'E')", vg_db, adOpenStatic
                If Not RS1.EOF Then
                    fecini = "01/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)
                    fecfin = "07/" & IIf(Mid(Fecha, 5, 2) = 12, "01/" & Mid(Fecha, 1, 4) + 1, Mid(Fecha, 5, 2) + 1 & "/" & Mid(Fecha, 1, 4)) 'dEoM("27/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))
                    fecpin = 0: fecpfi = 0
                    Do While fecini <= fecfin
                       '-------> Buscar fecha inicial y fecha final
                       For j = 1 To 7
                           If (DatePart("w", fecini, 2)) = Val(Mid(RS1!pad_diario, j, 1)) Then
                               If fecpin = 0 Then
                                  fecpin = Format(fecini, "yyyymmdd")
                               ElseIf fecpfi = 0 Then
                                  fecpfi = Format(fecini, "yyyymmdd")
                               End If
                            End If
                            If fecpin > 0 And fecpfi > 0 Then
                               fecxin = Format(CDate(fg_Ctod1(fecpin)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                               fecxin = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxin))), "yyyymmdd")
                               fecxfi = Format(CDate(fg_Ctod1(fecpfi)) + IIf(IsNull(RS1!pad_diaseg), 0, RS1!pad_diaseg), "yyyymmdd")
                               fecxfi = Format(CalcularDiasFeriados(LimpiaDato(Trim(fpText.text)), CDate(fg_Ctod1(fecxfi))), "yyyymmdd")
                               X = X + 1
                               .MaxRows = .MaxRows + 1
                               .InsertRows X, 1
                               .Row = X
                               .Col = 7: .CellType = CellTypeStaticText: .text = fg_Ctod1(fecpin)
                               fg_FormatearCeldaGrillaNuemrica vaSpread1, 8
                               .Col = 11: .CellType = CellTypeStaticText: .text = codpro
                               .SetActiveCell 8, X
                               fecpin = fecpfi: fecpfi = 0
                               Exit For
                            End If
                       Next j
                       fecini = fecini + 1
                    Loop
                End If
                RS1.Close: Set RS1 = Nothing
           End Select
        End If
        RS.Close: Set RS = Nothing
        est = False
    End With
Case 2
    Dim codsac As String
    With vaSpread1
        
        .Row = .ActiveRow
        .Col = 11
        codpro = ""
        codpro = Trim(.text)
        .Col = 1
        codsac = ""
        codsac = Trim(.text)
        If codpro = "" Or codsac = "" Then Exit Sub
        vg_codigo = ""
        If vg_pais = "CO" Then
           B_TabEst.LlenaDatos codsac, codpro, "Productos SAC", "CamPSAC"
        Else
           B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Pst"
        End If
        B_TabEst.Show 1
        If vg_codigo = "" Then Exit Sub
        '-------> Validar si existe producto en grilla
        If .SearchCol(1, 0, .MaxRows, Trim(vg_codigo), SearchFlagsNone) <> -1 Then
           MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        End If
        If vg_pais = "CO" Then
           RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, c.pad_codigo, a.pro_facsto, c.pad_tipo, f.foc_unisac, f.foc_faccon, f.foc_nomsac " & _
                   "FROM b_productos a, a_tipopro b, b_paramdesp c, " & aAp1 & " d, b_formatocomprassgp e, b_formatocompras f " & _
                   "WHERE a.pro_codtip = b.tip_codigo " & _
                   "AND   b.tip_codigo = d.pro_codtip " & _
                   "AND   d.pro_previo = c.pad_codigo " & _
                   "AND   c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                   "AND   e.fcs_codsac = '" & vg_codigo & "' " & _
                   "AND   a.pro_facing > 0 AND a.pro_facsto > 0 AND a.pro_codigo = e.fcs_codsgp AND e.fcs_codsac = f.foc_codsac ", vg_db, adOpenStatic
        Else
           RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, a.pro_facsto, c.pad_codigo, c.pad_tipo, e.uni_nomcor AS foc_unisac, 1 AS foc_faccon " & _
                   "FROM b_productos a, a_tipopro b, b_paramdesp c, " & aAp1 & " d, a_unidad e " & _
                   "WHERE a.pro_codtip = b.tip_codigo " & _
                   "AND   b.tip_codigo = d.pro_codtip " & _
                   "AND   d.pro_previo = c.pad_codigo " & _
                   "AND   d.pro_coduni = e.uni_codigo " & _
                   "AND   c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                   "AND   a.pro_codigo = '" & vg_codigo & "' " & _
                   "AND   a.pro_facing > 0 AND a.pro_facsto > 0", vg_db, adOpenStatic
        End If
        If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Producto no tiene asignado los factores", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        Toolbar1.Buttons(3).Enabled = False
        .Row = .ActiveRow
        .Col = 1: .text = vg_codigo
        .Col = 2: .text = Trim(RS!foc_nomsac)
        .Col = 3: .text = Trim(RS!foc_unisac)
        .Col = 6: .text = IIf(IsNull(RS!foc_faccon), 0, RS!foc_faccon)
        faccon = IIf(IsNull(RS!foc_faccon), 0, RS!foc_faccon)
        facsto = RS!pro_facsto
        RS.Close: Set RS = Nothing
        .Row = .ActiveRow + 1
        '-------> Traer STOCK DIA
        .Col = 11: codpro = .text
        stodia = 0
        If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(codpro), SearchFlagsNone) <> -1 Then
           vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(codpro), SearchFlagsNone)
           vaSpread2.Col = 2: stodia = Val(vaSpread2.text)
           vaSpread2.Col = 5: stodia = stodia + Val(vaSpread2.text)
           vaSpread2.Col = 4: conmir = Val(vaSpread2.text)
        End If
        If (stodia - conmir + ordrec) > 0 Then
           stodia = (stodia - conmir + ordrec)
        ElseIf (stodia - conmir + ordrec) <= 0 Then
           stodia = 0
        End If


        .Col = 17: canmin = IIf(Trim(.text) = "", 0, .text)
        If faccon > 0 Then
'           necmin = IIf(Int(((canmin + conmir) - stodia) / faccon) <> (((canmin + conmir) - stodia) / faccon), Int(((canmin + conmir) - stodia) / faccon) + 1, Round(((canmin + conmir) - stodia) / faccon, 0))
           necmin = IIf(Int(((canmin) - stodia) / faccon) <> (((canmin) - stodia) / faccon), Int(((canmin) - stodia) / faccon) + 1, Round(((canmin) - stodia) / faccon, 0))
           If Not IsNull(canmin) Then
              cansol = cansol + IIf(Int((canmin) / facsto) <> ((canmin) / facsto), ((canmin) / facsto), Round((canmin) / facsto, 0)) * facsto
           End If
        Else
           necmin = 0
           cansol = 0
        End If
        canres = 0
'        stodia = IIf((stodia - (canmin + conmir)) > 0, (stodia - (canmin + conmir)), 0)
        stodia = IIf((stodia - (canmin)) > 0, (stodia - (canmin)), 0)
'        canres = Round(IIf(stodia > 0, 0, (necmin * faccon)) - IIf((stodia - (canmin + conmir)) > 0, 0, (stodia - (canmin + conmir)) * -1), vg_DCa)
        canres = Round(IIf(stodia > 0, 0, (necmin * faccon)) - IIf((stodia - (canmin)) > 0, 0, (stodia - (canmin)) * -1), vg_DCa)
        If canres < 0 Then canres = 0
        canrea = canrea + IIf(necmin < 0, 0, necmin)
        .Col = 8: .CellType = IIf(canmin = 0, CellTypeNumber, CellTypeStaticText): .TypeHAlign = TypeHAlignRight: .text = Format(IIf(necmin < 0, 0, necmin), fg_Pict(6, 2))
        For i = .Row + 1 To .MaxRows
            .Row = i
            .Col = 11
            If Trim(.text) = Trim(codpro) Then
               .Col = 17
               If faccon > 0 Then
                  canmin = IIf(Trim(.text) = "", 0, vaSpread1.text)
                  cantid = IIf((canmin - stodia) > 0, (canmin - stodia), (cantid - stodia) - 1)
                  necmin = IIf(Int((cantid - canres) / faccon) <> ((cantid - canres) / faccon), Int((cantid - canres) / faccon) + 1, Round((cantid - canres) / faccon, 0))
                  If Not IsNull(canmin) Then
                     cansol = cansol + IIf(Int(canmin / facsto) <> (canmin / facsto), (canmin / facsto), Round((canmin) / facsto, 0)) * facsto
                  End If
                  stodia = IIf((stodia - canmin) > 0, (stodia - canmin), 0)
                  canres = Round(IIf(stodia > 0, 0, (necmin * faccon)) - IIf((stodia - (cantid - canres)) > 0, 0, (stodia - (cantid - canres)) * -1), vg_DCa)
                  .Col = 8: .CellType = IIf(canmin = 0, CellTypeNumber, CellTypeStaticText): .TypeHAlign = TypeHAlignRight: .text = Format(IIf(necmin < 0, 0, necmin), fg_Pict(6, 2))
               Else
                  .Col = 8: .CellType = IIf(canmin = 0, CellTypeNumber, CellTypeStaticText): .TypeHAlign = TypeHAlignRight: .text = Format(IIf(necmin < 0, 0, necmin), fg_Pict(6, 2))
               End If
               canrea = canrea + IIf(necmin < 0, 0, necmin)
            Else
                Exit For
            End If
        Next i
        .Row = .ActiveRow
        necmin = (cansol - stoact - ordrec + conmir)
        necmin = IIf(necmin < 0, 0, necmin)
        If faccon > 0 Then
           .Col = 5: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = canrea
           .Col = 8: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = canrea
        Else
           .Col = 5: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = 0
           .Col = 8: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .text = 0
        End If
    
    End With
End Select
Exit Sub
Man_Error:
    Resume Next
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RS As New ADODB.Recordset
Select Case Button.Index
Case 1
    If Not IsDate(fpDateTime1.text) Or Not IsDate(fpDateTime2.text) Then Exit Sub
    '-------> Validar si la minuta es teorica normal
    sql1 = IIf(vg_tipbase = "1", " val(mid(a.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) ")
    RS.Open "SELECT DISTINCT a.min_codigo FROM b_minuta a, a_servicio b WHERE a.min_codser = b.ser_codigo and b.ser_activo = '1' and a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND " & sql1 & " = " & Format(fpDateTime1.text, "yyyymm") & " AND min_indblo IN (2,11,99)", vg_db, adOpenStatic
    If Not RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Existe Bloque Minuta, pedido se cancela", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    MoverDatos
End Select
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
With vaSpread1

If .MaxRows < 1 Then Exit Sub
Dim codpro As String
Dim faccon As Double, canmin As Double, cansol As Double
Dim indrow As Long, i As Long
.Row = Row
.Col = 11: codpro = .text
.Col = 8: canmin = .text
If .SearchCol(11, 0, .MaxRows, Trim(codpro), SearchFlagsNone) <> -1 Then
   .Row = .SearchCol(11, 0, .MaxRows, Trim(codpro), SearchFlagsNone)
   indrow = .SearchCol(11, 0, .MaxRows, Trim(codpro), SearchFlagsNone)
   .Col = 6: faccon = .text
'   If Int(canmin / faccon) <> (canmin / faccon) Then
   If faccon <= 0 Then
      MsgBox "No esta definodo factor conversión", vbExclamation + vbOKOnly, Msgtitulo
      .Row = Row
      .Col = 8
      .text = 0
      Exit Sub
   End If
   If Int(canmin) <> (canmin) Then
'      MsgBox "Cantidad Digitada no corresponde unidad despacho", vbExclamation + vbOKOnly, Msgtitulo
      MsgBox "Cantidad despacho debe ser entero", vbExclamation + vbOKOnly, Msgtitulo
      .Row = Row
      .Col = 8
      .text = 0
      Exit Sub
   End If
   cansol = 0
   For i = indrow + 1 To .MaxRows
       .Row = i
       .Col = 11
       If Trim(.text) = Trim(codpro) Then
          .Col = 8
          cansol = cansol + .text
          If i = .MaxRows Then
             .Row = indrow
             .Col = 8
             .TypeHAlign = TypeHAlignRight
             .text = Format(cansol, fg_Pict(6, 0))
            Exit For
          End If
       Else
          .Row = indrow
          .Col = 8
          .TypeHAlign = TypeHAlignRight
          .text = Format(cansol, fg_Pict(6, 0))
          Exit For
       End If
   Next i
   Toolbar1.Buttons(3).Enabled = False
End If

End With
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
'If vaSpread1.MaxRows < 1 Then Exit Sub
'If ChangeMade = True Then
'   Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(8).Enabled = False
'   '-------> Rebajar o aumentar proposición pedido
'   Dim canped As Double, codpro As String, vStock As Double
'   vaSpread1.Row = Row
'   vaSpread1.Col = 11
'   codpro = vaSpread1.text
'   '-------> Fin rebajar o aumentar proposición pedido
'   '-------> recalcular stock
'   If vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) <> -1 Then
'      vaSpread1.Row = vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone)
'      vaSpread1.Col = 6: vStock = vaSpread1.text
'      For i = vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) To vaSpread1.MaxRows
'          vaSpread1.Row = i: vaSpread1.Col = 11
'          If vaSpread1.text = codpro Then
'             vaSpread1.Col = 5: canped = vaSpread1.text
'             vaSpread1.Col = 8: vaSpread1.text = Format(vStock, fg_Pict(6, 2))
'             canped = IIf(canped > vStock, (canped - vStock), IIf(Val(vStock) >= canped, 0, canped))
'             vaSpread1.Col = 9: vaSpread1.text = Format(canped, fg_Pict(6, 2))
'             vaSpread1.Col = 5: canped = vaSpread1.text
'             vStock = IIf((vStock - canped) <= 0, 0, (vStock - canped))
'          Else
'              Exit For
'          End If
'      Next i
'   End If
'   vaSpread1.EditEnterAction = EditEnterActionDown
'End If
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If Row = 0 Then Exit Sub
Dim coding As String, noming As String, codpro As String, NomPro As String, color As String
Dim necmin As Double, stoact As Double, ordrec As Double, conmir As Double, stodes As Double

With vaSpread1

.Row = Row
.Col = 1 ': color = Right(vaSpread1.text, 1)
If Trim(.text) = "" Then Exit Sub
TipWidth = 8000
ShowTip = True
MultiLine = 2
.Col = 4: necmin = IIf(Trim(.text) = "", 0, .text)
.Col = 9: coding = .text
.Col = 10: noming = .text
.Col = 11: codpro = .text
.Col = 12: NomPro = .text
.Col = 13: stoact = IIf(Trim(.text) = "", 0, .text)
.Col = 14: conmir = IIf(Trim(.text) = "", 0, .text)
.Col = 15: ordrec = IIf(Trim(.text) = "", 0, .text)
stodes = 0
If (stoact - conmir + ordrec) > 0 Then
   stodes = necmin - (stoact - conmir + ordrec)
ElseIf (stoact - conmir + ordrec) <= 0 Then
   stodes = necmin
End If
'TipText = "Ingrediente   : " & Trim(coding) & " - " & Trim(noming) & vbCrLf & _
'          "Producto       : " & Trim(codpro) & " - " & Trim(NomPro) & vbCrLf & _
'          "(+) Necesidad según minuta Teorica : ! " & Space(20) & fg_pone_espacio(Format(necmin, fg_Pict(6, 2)), 20) & vbCrLf & _
'          "(-) Stock actual                                         : ! " & Space(20) & fg_pone_espacio(Format(stoact, fg_Pict(6, 2)), 20) & vbCrLf & _
'          "(-) Ordenes de compra por recibir       : ! " & Space(20) & fg_pone_espacio(Format(ordrec, fg_Pict(6, 2)), 20) & vbCrLf & _
'          "(+) Por consumir según minuta Real : ! " & Space(20) & fg_pone_espacio(Format(conmir, fg_Pict(6, 2)), 20) & vbCrLf & _
'          "(=) Pedido propuesto                             : ! " & Space(20) & Format((necmin - stoact - ordrec + conmir), fg_Pict(6, 2)) & vbCrLf & _
'          "(=) Cantidad a solicitar                          : ! " & Space(20) & IIf((necmin - stoact - ordrec + conmir) < 0, Format(0, fg_Pict(6, 2)), Format((necmin - stoact - ordrec + conmir), fg_Pict(6, 2)))
TipText = "Ingrediente   : " & Trim(coding) & " - " & Trim(noming) & vbCrLf & _
          "Producto       : " & Trim(codpro) & " - " & Trim(NomPro) & vbCrLf & _
          "(+) Necesidad según minuta Teorica : ! " & Space(20) & fg_pone_espacio(Format(necmin, fg_Pict(6, 2)), 20) & vbCrLf & _
          "(-) Stock actual                                         : ! " & Space(20) & fg_pone_espacio(Format(stoact, fg_Pict(6, 2)), 20) & vbCrLf & _
          "(-) Ordenes de compra por recibir       : ! " & Space(20) & fg_pone_espacio(Format(ordrec, fg_Pict(6, 2)), 20) & vbCrLf & _
          "(+) Por consumir según minuta Real : ! " & Space(20) & fg_pone_espacio(Format(conmir, fg_Pict(6, 2)), 20) & vbCrLf & _
          "(=) Pedido propuesto                             : ! " & Space(20) & Format(stodes, fg_Pict(6, 2)) & vbCrLf & _
          "(=) Cantidad a solicitar                          : ! " & Space(20) & IIf((stoact - conmir + ordrec) <= 0, Format(necmin, fg_Pict(6, 2)), IIf((necmin - stoact - ordrec + conmir) < 0, Format(0, fg_Pict(6, 2)), Format((necmin - stoact - ordrec + conmir), fg_Pict(6, 2))))

End With
End Sub

Function CalcularDiasFeriados(cencos As String, Fecha As Variant) As String
Dim RS3 As New ADODB.Recordset
Dim diafer As Boolean
Dim sql1 As String
diafer = True
'-------> validar si existen dias feriado
sql1 = IIf(vg_tipbase = "1", " AND cdate(CFI_Fecha) = '" & Fecha & "' ", " AND Convert(VarChar(10), CFI_Fecha, 103) = '" & Fecha & "' ")
RS3.Open "SELECT CFI_Fecha FROM b_Fecha_Inhabiles WHERE CFI_CeCo = '" & cencos & "' " & sql1 & "", vg_db, adOpenStatic
If Not RS3.EOF Then
   RS3.Close: Set RS3 = Nothing
   diaferi = True
   Do While diaferi
      Fecha = CDate((Fecha)) + 1
      sql1 = IIf(vg_tipbase = "1", " AND cdate(CFI_Fecha) = '" & Fecha & "' ", " AND Convert(VarChar(10), CFI_Fecha, 103) = '" & Fecha & "' ")
      RS3.Open "SELECT CFI_Fecha FROM b_Fecha_Inhabiles WHERE CFI_CeCo = '" & cencos & "' " & sql1 & "", vg_db, adOpenStatic
      If RS3.EOF Then diaferi = False
      RS3.Close: Set RS3 = Nothing
   Loop
Else
   RS3.Close: Set RS3 = Nothing
End If
CalcularDiasFeriados = Fecha
End Function

Sub GenerarArchivoMdb()
Dim RS As New ADODB.Recordset
Dim strFileName As String
Dim strSQL  As String
Dim persac As String
Dim fecper As Long
Dim strFileNameMDB As String
Dim strLocalFileName As String
Dim StrFileNameLdb As String

persac = Mid(fpDateTime2.text, 4, 4) & Mid(fpDateTime2.text, 1, 2)
'-------> Crear directorio DatosTxt
If Dir(dir_trabajo & "DatosTxt", vbDirectory) = "" Then MkDir dir_trabajo & "DatosTxt"
'-------> Fin crear directorio DatosTxt

strFileName = dir_trabajo & ("DatosTxt\BajaData.txt")
If Dir(strFileName) <> "" Then Kill (strFileName)

fecper = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)
RS.Open "SELECT a.*, b.cli_nombre, b.cli_ccisac, b.cli_cecsac FROM b_minutapedidos a, b_clientes b WHERE a.ped_codcas = b.cli_codigo AND a.ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND a.ped_anomes = " & fecper & " AND a.ped_tipped = 1 AND a.ped_proped > 0", vg_db, adOpenStatic
Open strFileName For Output As #1

If Not RS.EOF Then
   Print #1, "CC" & Trim(RS!ped_codcas) & Trim(RS!cli_cecsac) & "_" & 1
   Print #1, "create table CADFIL (CADFIL_CDFIL char(10), CADFIL_NMFIL char(50))"
   Print #1, "create table SOLFIL (SOLFIL_IDSOL Integer, CADFIL_IDFIL Integer, TIPSOL_IDSOL Integer, SOLFIL_DTSOL Datetime, SOLFIL_DTREF Char(6), SOLFIL_NRSEM Integer, TIME_STAMP Datetime)"
   Print #1, "create table SOLITE (SOLFIL_IDSOL Integer, CPOPRO_CDPRO Char(20), SOLITE_DTENT Datetime, SOLITE_QTSOL Double, SOLITE_FLPRO Char(1), SOLITE_FLATU Integer, SOLITE_FLCPA Integer)"
   Print #1, "create table TABCEN (TABCEN_CDCEN char(4), TABCEN_DSCEN char(50))"
   Print #1, "create table TABPAR (TABPAR_NRVFL char(5))"
   Print #1, "INSERT INTO CadFil VALUES( '" & Trim(RS!ped_codcas) & "', '" & Trim(RS!cli_nombre) & "' )"
   Print #1, "INSERT INTO TabCen VALUES( '" & Trim(RS!cli_cecsac) & "', 'XX' )"
   Print #1, "INSERT INTO TabPar VALUES( '1.9' )"
   strSQL = "INSERT INTO SolFil VALUES( "
   strSQL = strSQL & 10000000 & ", "
   strSQL = strSQL & Trim(RS!cli_ccisac) & ", "
   strSQL = strSQL & RS!ped_tipped & ", "
   strSQL = strSQL & "'" & Format(Date, "dd/mm/yyyy") & "', "
   strSQL = strSQL & "'" & persac & "', "
   strSQL = strSQL & Val(fpLongInteger1(0).Value) & ", "
   strSQL = strSQL & " '" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "h:m:s") & "' " & ")"
   Print #1, strSQL
   While Not RS.EOF
         strSQL = "INSERT INTO SolIte VALUES( "
         strSQL = strSQL & 10000000 & ", "
         strSQL = strSQL & "'" & Trim(RS!ped_codsac) & "', "
         strSQL = strSQL & "'" & fg_Ctod1(RS!ped_fecped) & "', "
         strSQL = strSQL & RS!ped_proped & ", "
         strSQL = strSQL & "'I', "
         strSQL = strSQL & "-1, "
         strSQL = strSQL & 0 & ")"
         Print #1, strSQL
         RS.MoveNext
   Wend
End If
Close #1
RS.Close
fg_carga ""
strFileNameMDB = ""
strLocalFileName = strFileName
Open strLocalFileName For Input As #1
Do While Not EOF(1)
   Line Input #1, strLineReg
Loop
Close #1

'-------> Crear directorio DatosTxt
If Dir(dir_trabajo & "Datos", vbDirectory) = "" Then MkDir dir_trabajo & "Datos"
'-------> Fin crear directorio DatosTxt

'    lblStatus.Visible = True: prbStatus.Visible = True: prbStatus.Min = 0: lngRow = 0
Open strLocalFileName For Input As #1
If Not EOF(1) Then
   Line Input #1, strLineReg
   Do While Not EOF(1)
      DoEvents
      If Mid(strLineReg, 1, 2) = "CC" Then
         If Trim(strFileNameMDB) <> "" Then
            db7.Close: Set db7 = Nothing
            If Dir(StrFileNameLdb) <> "" Then
                Kill (StrFileNameLdb)
             End If
         End If
         strFileNameMDB = dir_trabajo & "datos\" & Trim(strLineReg) & ".mdb"
         StrFileNameLdb = dir_trabajo & "datos\" & Trim(strLineReg) & ".ldb"
         If Dir(strFileNameMDB) <> "" Then
            Kill (strFileNameMDB)
         End If
         Set db7 = DBEngine(0).CreateDatabase(strFileNameMDB, dbLangGeneral, dbVersion20)
      Else
         db7.Execute Trim(strLineReg)
      End If
      strLineReg = ""
      Line Input #1, strLineReg
      lngRow = lngRow + 1
   Loop
   If Trim(strLineReg) <> "" Then db7.Execute Trim(strLineReg)
   If Trim(strFileNameMDB) <> "" Then
      db7.Close: Set db7 = Nothing
      If Dir(StrFileNameLdb) <> "" Then
         Kill (StrFileNameLdb)
      End If
   End If
    Close #1
End If
    Screen.MousePointer = vbDefault
    DoEvents
End Sub
