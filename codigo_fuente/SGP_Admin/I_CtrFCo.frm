VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form I_CtrFCo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Facturas Compras"
   ClientHeight    =   1635
   ClientLeft      =   2145
   ClientTop       =   1455
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1215
      Index           =   0
      Left            =   30
      TabIndex        =   3
      Top             =   390
      Width           =   7395
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   0
         Left            =   3135
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   390
         Width           =   3765
         _Version        =   196608
         _ExtentX        =   6641
         _ExtentY        =   556
         Enabled         =   0   'False
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   390
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Top             =   720
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
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
         NegFormat       =   1
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
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Folio"
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
         Left            =   360
         TabIndex        =   6
         Top             =   750
         Width           =   690
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2685
         Picture         =   "I_CtrFCo.frx":0000
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Casino"
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
         Left            =   360
         TabIndex        =   4
         Top             =   465
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_CtrFCo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim i As Integer, isel As Integer
Dim MsgTitulo As String, tipinf As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 2115
Me.Width = 7530
fg_centra Me
MsgTitulo = "Control Facturas Compras"
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): btnX.Visible = True: btnX.ToolTipText = "Vista Previa": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): btnX.Visible = True: btnX.Enabled = False: btnX.ToolTipText = "Enviar": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): btnX.Visible = True: btnX.ToolTipText = "Historico Planificacón Teórica"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.Text = MuestraCasino(1)
fpayuda(0).Text = MuestraCasino(2)
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
'    Image1_Click 0
End Select
End Sub

Private Sub fpText_LostFocus()
If fpText.Text = "" Then fpayuda(0).Text = "": Exit Sub
RS.Open "select * from b_clientes where cli_codigo='" & fpText.Text & "' and cli_tipo=0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Text = "": Exit Sub
fpayuda(0).Text = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    Dim fecemi As Long
    RS.Open "select * from b_clientes where cli_codigo='" & fpText.Text & "' and cli_tipo=0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Text = "": Exit Sub
    fpayuda(0).Text = Trim(RS!cli_nombre)
    RS.Close: Set RS = Nothing
    If tipinf = "C" Then
       RS.Open "select max(toc_fecemi) as fecemi from b_totcompras where (toc_tipdoc='FA' or toc_tipdoc='NC' or toc_tipdoc='ND') and toc_numinf=" & Val(fpLongInteger1(0).Value) & " and toc_tipinf='C'", vg_db, adOpenStatic
    ElseIf tipinf = "T" Then
       RS.Open "select max(tov_fecemi) as fecemi from b_totventas where  tov_tipdoc='TR' and tov_numinf=" & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    ElseIf tipinf = "F" Then
       RS.Open "select max(toc_fecemi) as fecemi from b_totcompras where toc_numinf=" & Val(fpLongInteger1(0).Value) & " and toc_tipinf='F'", vg_db, adOpenStatic
    End If
    If RS.EOF Or IsNull(RS!fecemi) Then RS.Close: Set RS = Nothing: MsgBox "No existe información", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fecemi = Format(RS!fecemi, "ddmmyyyy")
    RS.Close: Set RS = Nothing
    If tipinf = "C" Then I_CFC fpText.Text, fecemi, Val(fpLongInteger1(0).Value)
    If tipinf = "T" Then I_CTC fpText.Text, fecemi, Val(fpLongInteger1(0).Value)
    If tipinf = "F" Then I_FoFi fpText.Text, fecemi, Val(fpLongInteger1(0).Value)
Case 3
    Dim numero As Long
    RS.Open "select inf_feccie from a_infcfcfofi where inf_tipo='" & tipinf & "' and inf_numero=" & (fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    If Not RS.EOF Then If RS!inf_feccie > 0 Then RS.Close: Set RS = Nothing: MsgBox "Nº Documento fue generado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    RS.Open "select max(inf_numero) as numero from a_infcfcfofi where inf_tipo='" & tipinf & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información, para enviar", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    numero = RS!numero
    RS.Close: Set RS = Nothing
    If tipinf = "C" Then
       RS.Open "select distinct toc_tipinf from b_totcompras where (toc_tipdoc='FA' or toc_tipdoc='NC' or toc_tipdoc='ND') and toc_numinf=" & numero & " and toc_tipinf='C'", vg_db, adOpenStatic
    ElseIf tipinf = "T" Then
       RS.Open "select distinct tov_tipdoc from b_totventas where  tov_tipdoc='TR' and tov_numinf=" & numero & "", vg_db, adOpenStatic
    ElseIf tipinf = "F" Then
       RS.Open "select distinct toc_tipinf from b_totcompras where toc_numinf=" & numero & " and toc_tipinf='F'", vg_db, adOpenStatic
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información, para enviar", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    vg_db.Execute "update a_infcfcfofi set inf_feccie=" & Format(Date, "yyyymmdd") & ", inf_usuario='" & vg_NUsr & "' where inf_tipo='" & tipinf & "' and inf_numero=" & numero & ""
    vg_db.Execute "insert into a_infcfcfofi (inf_tipo, inf_numero) values ('" & tipinf & "', " & (numero + 1) & ")"
    MsgBox "Generación envio Finalizado Sin Problema", vbExclamation + vbOKOnly, MsgTitulo
Case 5
    vg_codigo = ""
    Dim titform As String
    titform = ""
    If tipinf = "C" Then
       titform = "Histórico Control Facturas Compras"
    ElseIf tipinf = "T" Then
       titform = "Histórico Control Traspasos Casinos"
    ElseIf tipinf = "F" Then
       titform = "Histórico Control Fondo Fijo (Fofi)"
    End If
    B_HistPm.LlenarHistPlan titform, fpText.Text, tipinf, 4
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
Case 7
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
RS.Close: Set RS = Nothing
If Err.Number = -2147467259 Then
    MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"
Else
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If
End Sub

Sub Inicio(tfor As String, tf As String)
Me.Caption = tfor
MsgTitulo = tfor
tipinf = tf
'--- Buscar numero folio ----'
RS.Open "select max(inf_numero) as numero from a_infcfcfofi " & _
"where inf_tipo='" & tipinf & "'", vg_db, adOpenStatic
If Not RS.EOF Then fpLongInteger1(0).Value = RS!numero
RS.Close: Set RS = Nothing
'---------------------------'
End Sub

