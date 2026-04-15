VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form M_Minu03 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Plato Menú"
   ClientHeight    =   6300
   ClientLeft      =   2055
   ClientTop       =   1395
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6300
   ScaleWidth      =   7980
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7335
      Begin VB.TextBox FptNombre 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1815
         LinkTimeout     =   0
         MaxLength       =   30
         TabIndex        =   7
         Top             =   885
         Width           =   3195
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1815
         TabIndex        =   8
         Top             =   255
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
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
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   1
         Left            =   3120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   570
         Width           =   3675
         _Version        =   196608
         _ExtentX        =   6482
         _ExtentY        =   556
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483638
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
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   16777215
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   2
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1815
         TabIndex        =   10
         Top             =   570
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
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
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   0
         Left            =   3120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   255
         Width           =   3675
         _Version        =   196608
         _ExtentX        =   6482
         _ExtentY        =   556
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483638
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
         AlignTextV      =   2
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   2
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
         Caption         =   "Registro 0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   5040
         TabIndex        =   15
         Top             =   1000
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Texto"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   555
         TabIndex        =   14
         Top             =   1000
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Recetario"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   555
         TabIndex        =   13
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Subrecetario"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   555
         TabIndex        =   12
         Top             =   675
         Width           =   1305
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2680
         Picture         =   "M_Minu03.frx":0000
         Top             =   160
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2680
         Picture         =   "M_Minu03.frx":030A
         Top             =   480
         Width           =   480
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4140
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   7335
      _Version        =   393216
      _ExtentX        =   12938
      _ExtentY        =   7303
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
      MaxRows         =   20
      OperationMode   =   2
      RestrictRows    =   -1  'True
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "M_Minu03.frx":0614
      VisibleCols     =   3
      VisibleRows     =   20
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu03.frx":0A64
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu03.frx":0D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu03.frx":1098
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu03.frx":13B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6300
      Left            =   7350
      TabIndex        =   1
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   11113
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Receta"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2070
      TabIndex        =   5
      Top             =   1485
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2070
      TabIndex        =   4
      Top             =   1845
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Plato"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   315
      TabIndex        =   3
      Top             =   1845
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "Categoria Dietetica"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   315
      TabIndex        =   2
      Top             =   1485
      Width           =   1800
   End
End
Attribute VB_Name = "M_Minu03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConSql As ADODB.Recordset, Consql1 As ADODB.Recordset
Dim i As Long, irow As Long
Dim findstring As String, sourcestring As String
Dim swactiva As Integer, iayuda As Integer
Dim SubCRecetas As Long, codsegmento As Long, codsubsegmento As Long
Private Sub Form_Activate()
fg_descarga
End Sub
Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me
fg_carga (ss)
vaSpread1.MaxRows = 0
vg_plancategoria1 = 0: vg_plancategoria2 = 0
vg_plancategoria3 = 0: vg_plancategoria4 = 0
vg_codregimen = 0: iayuda = 0

Set ConSql = vg_db.Execute("select * " & _
             "From PB00062", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_paramreceta 1, 0, ''", , adCmdStoredProc)
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
      vg_auxcodsegmento = ConSql!Deflt_Rcpe_Wstg_Rt
      fpLongInteger1(0).Value = ConSql!Deflt_Rcpe_Wstg_Rt
      fpLongInteger1(1).Value = ConSql!Deflt_Rcpe_Mrgn_Rt
      ConSql.MoveNext
   Loop
End If
ConSql.Close: Set ConSql = Nothing

If Val(fpLongInteger1(0).Value) < 1 Then fpLongInteger1(0).Value = "": Exit Sub
codsegmento = Val(fpLongInteger1(0).Value)
vg_auxcodsegmento = Val(fpLongInteger1(0).Value)

Set ConSql = vg_db.Execute("select * " & _
             "From Sdx_PB00074 " & _
             "where Unit_Dfnd_No=" & codsegmento & "", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_segmento 4, " & codsegmento & ", ''", , adCmdStoredProc)
If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: iayuda = 0: fpayuda(0).Text = "": vg_codigo = 0: vg_auxcodsegmento = 0: fpayuda(1).Text = "": fpLongInteger1(1).Value = "": fpLongInteger1(0).SetFocus: Exit Sub
fpayuda(0).Text = ConSql!Unit_Dfnd_Desc
fpayuda(1).Text = ""
vaSpread1.MaxRows = 0
FptNombre.Text = ""
ConSql.Close: Set ConSql = Nothing

If Val(fpLongInteger1(1).Value) < 1 Then fpLongInteger1(1).Value = "": Exit Sub
codsubsegmento = Val(fpLongInteger1(1).Value)
Set ConSql = vg_db.Execute("select * " & _
             "From PB00074 " & _
             "where Unit_Dfnd_No=" & codsubsegmento & " " & _
             "and   Prev_Unit_Dfnd_No=" & codsegmento & "", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_subsegmento 5, " & codsubsegmento & ", " & codsegmento & ", ''", , adCmdStoredProc)
If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: fpayuda(1).Text = "": vg_codigo = 0: fpLongInteger1(1).SetFocus: Exit Sub
fpayuda(1).Text = ConSql!Unit_Dfnd_Desc
ConSql.Close: Set ConSql = Nothing
If codsegmento > 0 And codsubsegmento > 0 Then
   MoverRecetas
End If

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cargando Tabla Anexa Plantilla Menú"
End Sub
Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
Select Case Index
  Case 0
    If Val(fpLongInteger1(0).Value) < 1 Then fpLongInteger1(0).Value = "": Exit Sub
    codsegmento = Val(fpLongInteger1(0).Value)
    vg_auxcodsegmento = Val(fpLongInteger1(0).Value)
    Set ConSql = vg_db.Execute("select * " & _
                 "From Sdx_PB00074 " & _
                 "where Unit_Dfnd_No=" & codsegmento & "", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_segmento 4, " & codsegmento & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: iayuda = 0: fpayuda(0).Text = "": vg_codigo = 0: vg_auxcodsegmento = 0: fpayuda(1).Text = "": fpLongInteger1(1).Value = 0: fpLongInteger1(0).SetFocus: Exit Sub
    fpayuda(0).Text = ConSql!Unit_Dfnd_Desc
    fpayuda(1).Text = "": fpLongInteger1(1).Value = ""
    vaSpread1.MaxRows = 0
    FptNombre.Text = ""
    ConSql.Close: Set ConSql = Nothing
   Case 1
    If Val(fpLongInteger1(1).Value) < 1 Then fpLongInteger1(1).Value = "": Exit Sub
    codsubsegmento = Val(fpLongInteger1(1).Value)
    Set ConSql = vg_db.Execute("select * " & _
                 "From PB00074 " & _
                 "where Unit_Dfnd_No=" & codsubsegmento & " " & _
                 "and   Prev_Unit_Dfnd_No=" & codsegmento & "", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_subsegmento 5, " & codsubsegmento & ", " & codsegmento & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: iayuda = 0: fpayuda(1).Text = "": vg_codigo = 0: fpLongInteger1(1).SetFocus: Exit Sub
    fpayuda(1).Text = ConSql!Unit_Dfnd_Desc
    ConSql.Close: Set ConSql = Nothing
    If codsegmento > 0 And codsubsegmento > 0 Then
       MoverRecetas
       Label1(1).Caption = "Todos"
       Label1(0).Caption = "Todos"
    End If
End Select
End Sub
Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    If Index = 0 Then image1_Click 0
    If Index = 1 Then image1_Click 1
End Select
End Sub
'Private Sub fpLongInteger1_LostFocus(Index As Integer)
'Select Case Index
'  Case 0
'    If Val(fpLongInteger1(0).Value) < 1 Then fpLongInteger1(0).Value = "": Exit Sub
'    If iayuda = 1 Then Exit Sub
'    vg_auxcodsegmento = Val(fpLongInteger1(0).Value)
'    codsegmento = Val(fpLongInteger1(0).Value)
'    Set ConSql = vg_db.Execute("sod_s_segmento 4, " & codsegmento & ", ''", , adCmdStoredProc)
'    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: fpayuda(0).Text = "": vg_codigo = 0: vg_auxcodsegmento = 0: fpayuda(1).Text = "": fpLongInteger1(1).Value = 0: fpLongInteger1(0).SetFocus: Exit Sub
'    fpayuda(0).Text = ConSql!Unit_Dfnd_Desc
'    fpayuda(1).Text = "": fpLongInteger1(1).Value = ""
'    vaSpread1.MaxRows = 0
'    FptNombre.Text = ""
'    ConSql.Close: Set ConSql = Nothing
'   Case 1
'    If Val(fpLongInteger1(0).Value) < 1 Then fpLongInteger1(1).Value = "": Exit Sub
'    If iayuda = 1 Then Exit Sub
'    codsubsegmento = Val(fpLongInteger1(1).Value)
'    Set ConSql = vg_db.Execute("sod_s_subsegmento 5, " & codsubsegmento & ", " & Val(fpLongInteger1(0).Value) & ", ''", , adCmdStoredProc)
'    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: fpayuda(1).Text = "": vg_codigo = 0: fpLongInteger1(1).SetFocus: Exit Sub
'    fpayuda(1).Text = ConSql!Unit_Dfnd_Desc
'    ConSql.Close: Set ConSql = Nothing
'    If codsegmento > 0 And codsubsegmento > 0 Then
'       MoverRecetas
'    End If
'End Select
'End Sub
Private Sub fpTnombre_Change()
If vaSpread1.MaxRows < 1 Then Exit Sub
findstring = Trim(FptNombre.Text)
If FptNombre.Text = "" Then
   vaSpread1.Visible = False
   swactiva = 0
   For i = 1 To vaSpread1.MaxRows
       vaSpread1.Row = i
       vaSpread1.RowHidden = False
       If swactiva = 0 Then vaSpread1.OperationMode = 2: vaSpread1.Action = 0: swactiva = 1
   Next i
   Label1(0).Caption = "Registro " & Format(vaSpread1.MaxRows, fg_Pict(6, 0))
   vaSpread1.Visible = True
Else
   swactiva = 0
   vaSpread1.Visible = False
   irow = 0
   For i = 1 To vaSpread1.MaxRows
       vaSpread1.Row = i
       vaSpread1.Col = 2
       sourcestring = Trim(vaSpread1.Value)
       indactivo = UCase(Trim(sourcestring)) Like "*" + UCase(findstring) + "*"
       If indactivo = -1 Then
          If swactiva = 0 Then vaSpread1.OperationMode = 2: vaSpread1.Action = 0: swactiva = 1
          If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
          irow = irow + 1
       Else
          If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
       End If
   Next i
   Label1(0).Caption = "Reg. Enc. " & Format(irow, fg_Pict(6, 0))
   vaSpread1.Visible = True
End If
End Sub
Private Sub image1_Click(Index As Integer)
Select Case Index
  Case 0
    iayuda = 1
    vg_codigo = 0
    vg_left = fpayuda(0).Left + 2000
    B_Segmen.Show 1
    M_Minu03.Refresh
    iayuda = 0
    If vg_codigo = 0 Then Exit Sub
    fpLongInteger1(0).Value = vg_codigo
    codsegmento = vg_codigo
    vg_auxcodsegmento = vg_codigo
    Set ConSql = vg_db.Execute("select * " & _
                 "From Sdx_PB00074 " & _
                 "where Unit_Dfnd_No=" & codsegmento & "", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_segmento 4, " & codsegmento & ", ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(0).Text = ConSql!Unit_Dfnd_Desc
       fpayuda(1).Text = ""
       fpLongInteger1(1).Value = ""
       vaSpread1.MaxRows = 0
       FptNombre.Text = ""
       fpLongInteger1(1).SetFocus
    Else
       fpayuda(0).Text = "": fpayuda(1).Text = ""
       vg_codigo = 0: fpLongInteger1(0).Value = "": fpLongInteger1(1).Value = ""
       MsgBox "Segmento No Existe", vbExclamation + vbOKOnly, "Cambiar Plato Menú"
    End If
    ConSql.Close: Set ConSql = Nothing
  Case 1
    iayuda = 1
    vg_codigo = 0
    vg_opcion = 0
    vg_left = fpayuda(1).Left + 2000
    B_SubSeg.Show 1
    M_Minu03.Refresh
    If vg_codigo = 0 Then iayuda = 0: Exit Sub
    fpLongInteger1(1).Value = vg_codigo
    codsubsegmento = vg_codigo
    Set ConSql = vg_db.Execute("select * " & _
                 "From PB00074 " & _
                 "where Unit_Dfnd_No=" & codsubsegmento & " " & _
                 "and   Prev_Unit_Dfnd_No=" & codsegmento & "", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_subsegmento 5, " & codsubsegmento & ", " & codsegmento & ", ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(1).Text = ConSql!Unit_Dfnd_Desc
       fpLongInteger1(1).SetFocus
    Else
       fpayuda(1).Text = ""
       vg_codigo = 0
       MsgBox "Subsegmento No Existe", vbExclamation + vbOKOnly, "Cambiar Plato Menú"
    End If
    ConSql.Close: Set ConSql = Nothing
    iayuda = 0
    If codsegmento > 0 And codsubsegmento Then
       MoverRecetas
       FptNombre.SetFocus
    End If
End Select
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 1
    If vaSpread1.MaxRows < 1 Then Exit Sub
    ICGrilla = 1
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1_DblClick vaSpread1.Col, vaSpread1.Row
  Case 3
    B_DieReP.Show 1
    If vg_opcion <> 2 Then
       MoverRecetas
       Label2(0).Caption = "Todos"
       Label2(1).Caption = "Todos"
       If vg_desctplatoplan <> "" Then Label2(0).Caption = vg_desctplatoplan
       If vg_descdieteticoplan <> "" Then Label2(1).Caption = vg_descdieteticoplan
    End If
  Case 5
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    vg_vercodreceta = Val(vaSpread1.Text)
    V_Recetas.Show 1
''    If vaSpread1.MaxRows < 1 Then Exit Sub
''    vaSpread1.Row = vaSpread1.ActiveRow
''    vaSpread1.Col = 1
''    SubCReceta = Val(vaSpread1.Value)
''    vRet = Shell(dir_trabajo & "\Subrecet.exe " & vg_NUsr & "," & vg_Pass & "," & CStr(SubCReceta) & "," & CStr(WsCodPVenta) & ",", 1)
''    vRet = Shell(dir_trabajo & "\Subrecet.exe " & vg_NUsr & "," & vg_Pass & "," & CStr(SubCReceta) & ",", 1)
  Case 7
    ICGrilla = 0
    Me.Hide
End Select
End Sub
Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If Row > 0 Then vaSpread1.Row = vaSpread1.ActiveRow
End Sub
Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim NewRow10 As Long

If Row > 0 Then
   ICGrilla = 1
   vaSpread1.Row = vaSpread1.ActiveRow
   NewRow10 = Program.Row
   For i = 1 To 100
       Program.Row = i
       Program.Col = IndColumna + 6
       vaSpread1.Col = 1
       If Val(Program.Value) = Val(vaSpread1.Value) Then
          TITLE = "Ingreso de Recetas"
          msg = "Opción Ya Existe ¿ Esta Seguro De Ingresarla ?"
          Style = vbYesNo + vbQuestion + vbDefaultButton2
          Help = "DEMO.HLP"
          Ctxt = 1000
          ws_respuesta = MsgBox(msg, Style, TITLE, Help, Ctxt)
          Select Case ws_respuesta
            Case Is = vbNo
              Program.Row = NewRow10
              Exit Sub
            Case Is = vbYes
              Exit For
          End Select
       End If
   Next i
   Program.Row = NewRow10
   Select Case IndPlantilla
     Case 1
       Program.Col = IndColumna
       Program.CellType = 5
       Program.TypeHAlign = 2
       Program.Value = "R"
       Program.ForeColor = &HFF&
       Program.BackColor = &H80FF80
       Program.Col = IndColumna + 4
       Program.Value = 1
   
       vaSpread1.Col = 2
       Program.Col = IndColumna + 1
  ' Limpiar Datos y Formato Celda
       Program.Action = 3
  ' Retorna Modo de la columna
       Program.BlockMode = False
       Program.Font.Bold = False
       Program.Font.Size = 8
       Program.Value = vaSpread1.Value
              
       Program.Col = IndColumna + 2
       Program.CellType = 3
       Program.TypeIntegerMin = 1
       Program.TypeIntegerMax = 9999999
       Program.TypeHAlign = 1
       Program.TypeSpin = False
       Program.TypeIntegerSpinInc = 1
       Program.TypeIntegerSpinWrap = False
       Program.Value = 0
       Program.ForeColor = &HFF0000 '&HFF&
              
       vaSpread1.Col = 3
       Program.Col = IndColumna + 6
       Program.TypeHAlign = 1
       Program.Value = Format(Val(vaSpread1.Value), fg_Pict(6, 2))
       Program.ForeColor = &HFF0000
              
       vaSpread1.Col = 1
       Program.Col = IndColumna + 5
       Program.Value = vaSpread1.Value
   End Select
   M_Minu02.Plantilla(0).Enabled = True
   IndGrabadoDetalle = 1
   Me.Hide
'   Unload Me
End If
End Sub
Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If vaSpread1.MaxRows > 0 Then
   ICGrilla = 1
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1_DblClick vaSpread1.Col, vaSpread1.Row
End If
End Sub
Private Sub MoverRecetas()
fg_carga ""
vaSpread1.MaxRows = 0
Set ConSql = vg_db.Execute("select PB00078.Rcpe_No, PB00078.Rcpe_Desc, PB00078.Rcpe_Menu_Narr, " & _
             "PB00077.Rcpe_Cost_Val " & _
             "From PB00077, PB00078, PB00083, PB00357 " & _
             "Where PB00077.Rcpe_No = PB00078.Rcpe_No " & _
             "and   PB00077.Rcpe_No = PB00083.Rcpe_No " & _
             "and  (PB00077.Rcpe_No=PB00357.Rcpe_No) " & _
             "and   PB00083.Unit_Dfnd_No=" & codsubsegmento & " " & _
             "and  (PB00357.Diet_Cat_No=" & vg_codregimen & " or " & vg_codregimen & "=0) " & _
             "and  (PB00077.Rcpe_Cat_1_No=" & vg_plancategoria1 & " or " & vg_plancategoria1 & "=0) " & _
             "and  (PB00077.Rcpe_Cat_2_No=" & vg_plancategoria2 & " or " & vg_plancategoria2 & "=0) " & _
             "and  (PB00077.Rcpe_Cat_3_No=" & vg_plancategoria3 & " or " & vg_plancategoria3 & "=0) " & _
             "and  (PB00077.Rcpe_Cat_4_No=" & vg_plancategoria4 & " or " & vg_plancategoria4 & "=0) " & _
             "and   PB00077.Del_Ind=0 " & _
             "and   PB00078.Del_Ind=0 " & _
             "and   PB00083.Del_Ind=0 " & _
             "order by PB00078.Rcpe_Desc", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_receta 6, 0, " & codsubsegmento & ", " & vg_codregimen & ", " & vg_plancategoria1 & ", " & vg_plancategoria2 & ", " & vg_plancategoria3 & ", " & vg_plancategoria4 & ", ''", , adCmdStoredProc)
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
     
      vaSpread1.Col = 1
      vaSpread1.CellType = 5
      vaSpread1.TypeHAlign = 1
      vaSpread1.Value = ConSql!Rcpe_No
      
      vaSpread1.Col = 2
      vaSpread1.CellType = 5
      vaSpread1.TypeHAlign = 0
      vaSpread1.Value = Trim(ConSql!Rcpe_Desc)
      
      vaSpread1.Col = 3
      vaSpread1.CellType = 5
      vaSpread1.TypeHAlign = 1
      vaSpread1.Value = Format(ConSql!Rcpe_Cost_Val, fg_Pict(6, 2))
              
      ConSql.MoveNext
   Loop
'   Command1(0).Enabled = True
'   Command1(1).Enabled = True
'   Command1(3).Enabled = True
   Label1(0).Visible = True
   Label1(0).Caption = "Registro " & Format(vaSpread1.MaxRows, fg_Pict(6, 0))
   ConSql.Close: Set ConSql = Nothing: fg_descarga
Else
   fg_descarga
'   Command1(0).Enabled = False
'   Command1(1).Enabled = False
'   Command1(3).Enabled = False
   Label1(0).Visible = True
   Label1(0).Caption = "Registro " & Format(vaSpread1.MaxRows, fg_Pict(6, 0))
   If indopccion = 11 Then MsgBox "No Existe Receta Con Esos Parametros", vbExclamation + vbOKOnly, "Detalle Plantilla Menú": ConSql.Close: Set ConSql = Nothing: fg_descarga
   If IndOpcion <> 11 Then MsgBox "No Existen Recetas Ha Buscar", vbExclamation + vbOKOnly, "Detalle Plantilla Menú": ConSql.Close: Set ConSql = Nothing: fg_descarga ': Command1_Click (3) 'M_MINU20.Show 1
'   Command1(0).Enabled = False
'   Command1(1).Enabled = False
'   Command1(3).Enabled = False
End If
vaSpread1.Row = 1
End Sub
