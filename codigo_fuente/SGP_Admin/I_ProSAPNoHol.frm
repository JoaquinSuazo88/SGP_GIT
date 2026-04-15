VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_ProSAPNoHol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos no Homologados SAP"
   ClientHeight    =   2160
   ClientLeft      =   4320
   ClientTop       =   2955
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2160
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin EditLib.fpText fpText 
      Height          =   315
      Index           =   1
      Left            =   1800
      TabIndex        =   0
      Top             =   600
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
   Begin EditLib.fpDateTime fpDateTime1 
      Height          =   315
      Index           =   0
      Left            =   1815
      TabIndex        =   2
      Top             =   1440
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
      Left            =   7350
      TabIndex        =   3
      Top             =   1440
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
      Index           =   2
      Left            =   1800
      TabIndex        =   1
      Top             =   960
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
      ButtonAlign     =   1
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   3120
      Picture         =   "I_ProSAPNoHol.frx":0000
      Top             =   880
      Width           =   480
   End
   Begin VB.Label fpayuda 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   3600
      TabIndex        =   11
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Producto SAP"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label2 
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
      Index           =   2
      Left            =   6225
      TabIndex        =   9
      Top             =   1515
      Width           =   1005
   End
   Begin VB.Label Label2 
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
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1515
      Width           =   1110
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   3120
      Picture         =   "I_ProSAPNoHol.frx":030A
      Top             =   510
      Width           =   480
   End
   Begin VB.Label fpayuda 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   3600
      TabIndex        =   6
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Centro De Costo"
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
      Left            =   240
      TabIndex        =   5
      Top             =   660
      Width           =   1410
   End
   Begin VB.Label sombra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   3630
      TabIndex        =   7
      Top             =   645
      Width           =   5055
   End
   Begin VB.Label sombra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   3630
      TabIndex        =   12
      Top             =   1005
      Width           =   5055
   End
End
Attribute VB_Name = "I_ProSAPNoHol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
fg_centra Me
'Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = True
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpDateTime1(0).text = Date
fpDateTime1(1).text = Date
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then Exit Sub
End Sub

Private Sub fpText_Change(Index As Integer)
Dim RS As New ADODB.Recordset
Select Case Index
Case 1
    RS.Open "sgpadm_s_cliente_V02 19, '" & Trim(LimpiaDato(fpText(1).text)) & "', ''", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!cli_nombre)
    RS.Close: Set RS = Nothing
Case 2
    RS.Open "sgpadm_s_productossap 1, '" & Trim(LimpiaDato(fpText(2).text)) & "', ''", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "": Exit Sub
    fpayuda(2).Caption = Trim(RS!fcs_CodMaterial)
    RS.Close: Set RS = Nothing
End Select
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Centro Costo", "CecoPortalElec"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(1).text = vg_codigo
    fpayuda(1).Caption = vg_nombre
    fpText(2).SetFocus
Case 2
    vg_left = fpayuda(2).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_formatocompras_sap", "", "Formato Compras Sap", "ForComSap"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(2).text = vg_codigo
    fpayuda(2).Caption = vg_nombre
    fpDateTime1(0).SetFocus
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If fpDateTime1(0).Value > fpDateTime1(1).Value Then MsgBox "Fecha inicial debe menor o igual fecha final", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then MsgBox "Formato fecha no corresponde", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Trim(fpayuda(1).Caption) = "" And Trim(fpText(1).text) <> "" Then MsgBox "Centro costo no corresponde", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Trim(fpayuda(2).Caption) = "" And Trim(fpText(2).text) <> "" Then MsgBox "Código Mercaderia SAP no corresponde", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    I_ProductosnoHomologados fpText(1).text, fpText(2).text, Format(fpDateTime1(0).text, "yyyymmdd"), Format(fpDateTime1(1).text, "yyyymmdd")
Case 3
    Me.Hide
    Unload Me
End Select
End Sub
