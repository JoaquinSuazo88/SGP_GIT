VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form M_Usua 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantención Usuario "
   ClientHeight    =   5955
   ClientLeft      =   270
   ClientTop       =   1515
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   14760
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   971
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "M_UsuCtr.frx":0000
         Left            =   1680
         List            =   "M_UsuCtr.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   555
         Width           =   2505
         _Version        =   196608
         _ExtentX        =   4410
         _ExtentY        =   870
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
         ThreeDOutsideHighlightColor=   -2147483628
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
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
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
         Left            =   4260
         TabIndex        =   6
         Top             =   640
         Width           =   525
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
         Left            =   255
         TabIndex        =   5
         Top             =   645
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Columna"
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
         Left            =   255
         TabIndex        =   4
         Top             =   345
         Width           =   1320
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":001E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":0338
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":0652
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":096C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":0C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":0FA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":15D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":18EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":1C08
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":1F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":223C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":2556
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":2870
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_UsuCtr.frx":2B8A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incluir"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Alterar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Borrar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Actualizar Lista"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   15
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_Usua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset, RS2 As New ADODB.Recordset
Dim modo As String, loginusuario As String, hijo As String, hijo2 As String
Dim vecdatos(5) As String
Dim ivalidar As Integer, itexto As Integer, itab As Integer, fin As Integer
Dim i As Long, j As Long, ibusca As Long, nivel As Long, codhijo As Long, codhijo2 As Long, indindex As Long
Dim dest As Node, sourcenode As Node, nd As Node, rootnode As Node

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 5670
Me.Width = 7425
fg_centra Me
SSTab1.Tab = 0
itab = 0
modo = "M"
MoverDatosGrillas
End Sub

Private Sub fpText_Change()
If LimpiaDato(Trim(fpText1.Text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    RS1.Open "select count(usu_codigo) as nreg From a_usuarios where usu_codigo like '%" + UCase(LimpiaDato(fpText1.Text)) & "%'", vg_db, adOpenStatic
    ibusca = RS1!NReg: vaSpread1.MaxRows = RS1!NReg: RS1.Close: Set RS = Nothing
    RS1.Open "select * From a_usuarios where imp_codigo like '%" + UCase(LimpiaDato(fpText1.Text)) & "%'", vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    RS1.Open "select count(usu_codigo) as nreg From a_usuarios where ucase(usu_nombre) like '%" + UCase(LimpiaDato(fpText1.Text)) & "%'", vg_db, adOpenStatic
    ibusca = RS1!NReg: vaSpread1.MaxRows = RS1!NReg: RS1.Close: Set RS = Nothing
    RS1.Open "select * From a_usuarios where ucase(usu_nombre) like '%" + UCase(LimpiaDato(fpText1.Text)) & "%'", vg_db, adOpenStatic
End If
i = 1
If Not RS1.EOF Then
   Do While Not RS1.EOF
      vaSpread1.Row = i
        i = i + 1
        vaSpread1.Col = 1
        vaSpread1.TypeHAlign = 1
        vaSpread1.Value = RS1!usu_codigo
        vaSpread1.Col = 2
        vaSpread1.Value = Trim(RS1!usu_nombre)
        RS1.MoveNext
        Loop
        Ac_Botones
End If
RS1.Close: Set RS = Nothing
If fpText1.Text = "" Then
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Else
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Encontrado"
End If
End Sub


Private Sub fpText1_Change(Index As Integer)
If fpText1(Index).Text <> vecdatos(Index) And itexto = 0 And modo = "M" Then
   If Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True Then Exit Sub
   Ac_HabDes 4
   Ac_Boton 1
End If
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
'  Case 13 And Toolbar1.Buttons(9).Visible = True And Toolbar1.Buttons(11).Visible = True
'    Actualiza_Datos
  Case 27 And Toolbar1.Buttons(11).Visible = True And Toolbar1.Buttons(13).Visible = True
'    Cancela_Fila
  Case 113 And Toolbar1.Buttons(1).Visible = True
    modo = "A"
    Agrega_Dato
  Case 114 And Toolbar1.Buttons(3).Visible = True
    modo = "M"
    Agrega_Dato
  Case 115 And Toolbar1.Buttons(5).Visible = True
'    Borra_Fila
End Select
End Sub

Private Sub Label3_Click()

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
  Case 0
    If itab = 2 Then itab = 0: Ac_Boton 2
  Case 1
    If vaSpread1.MaxRows > 0 And modo = "M" Then
       If itab = 2 Then itab = 1: Ac_Boton 2
       modo = "M"
       MoverDetalleDatos
'       M_Receta.Refresh
    ElseIf vaSpread1.MaxRows < 1 And modo = "M" Then
       SSTab1.Tab = 0
       Exit Sub
    End If
  Case 2
    If vaSpread1.MaxRows < 1 Then SSTab1.Tab = 0: Exit Sub
    Ac_Boton 4
    itab = 2
Case 3
    If vaSpread1.MaxRows < 1 Then SSTab1.Tab = 0: Exit Sub
    Ac_Boton 4
    itab = 2
    
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 1
    Agrega_Dato
  Case 3
    If vaSpread1.MaxRows < 1 Then Exit Sub
    Altera_Dato
  Case 5
    Borra_Fila
  Case 7
'    fpText1.Text = ""
'    MoverDatosGrillas
  Case 10
    Cancela_Dato
  Case 12
    Actualiza_Dato
  Case 15
'    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, "Color": Exit Sub
'    vg_opimp = 14
'    Preview.Show 1
  Case 18
    Me.Hide
    Unload Me
End Select
End Sub

Sub Agrega_Dato()
itexto = 1
modo = "A"
Ac_Boton 1
Ac_HabDes 2
LimpiarVariable
itexto = 0
End Sub

Sub Altera_Dato()
Ac_HabDes 2
Ac_Boton 1
If SSTab1.Tab = 1 Or SSTab1.Tab = 0 Then
   SSTab1.TabEnabled(2) = False
   SSTab1.Tab = 1
'   MoverDetalleDatos
End If
End Sub

Sub Borra_Fila()

On Error GoTo Man_Error
Resp_Delete ("Mantención")
If respuesta = vbYes Then
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = 1: loginusuario = vaSpread1.Value
   vg_db.BeginTrans
     vg_db.Execute "delete Sdx_UsuCtrlAcceso from Sdx_UsuCtrlAcceso where login='" & loginusuario & "'"
     vg_db.Execute "delete Sdx_Usuario from Sdx_Usuario where loginusuario='" & loginusuario & "'"
   vg_db.CommitTrans
   vaSpread1.Action = 5
   vaSpread1.MaxRows = vaSpread1.MaxRows - 1
   Set RS1 = vg_db.Execute("select count(loginusuario) as nreg " & _
                "From Sdx_Usuario " & _
                "where ucase(nombre) like '%" + UCase(("")) + "%'", , adCmdText)
'   Set RS1 = vg_db.Execute("sod_s_usuario 3, '', '%" + UCase(("")) + "%'", , adCmdStoredProc)
   If RS1.EOF Or RS1!NReg = 0 Then
      RS1.Close: Set RS1 = Nothing
      Ac_Boton 3
   ElseIf RS1!NReg > 0 Then
      RS1.Close: Set RS1 = Nothing
      Ac_Boton 2
   End If
End If

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub Cancela_Dato()
TITLE = "Usuario"
msg = "Cancelar Operación"
Style = vbYesNo + vbQuestion + vbDefaultButton2
Help = "DEMO.HLP"
Ctxt = 1000
ws_respuesta = MsgBox(msg, Style, TITLE, Help, Ctxt)
Select Case ws_respuesta
  Case Is = vbYes
    If SSTab1.Tab = 1 Then
       Set RS1 = vg_db.Execute("select count(loginusuario) as nreg " & _
                    "From Sdx_Usuario " & _
                    "where ucase(nombre) like '%" + UCase(LimpiaDato("")) + "%'", , adCmdText)
'       Set RS1 = vg_db.Execute("sod_s_usuario 3, '', '%" + UCase(LimpiaDato("")) + "%'", , adCmdStoredProc)
       If RS1.EOF Or RS1!NReg = 0 Then
          RS1.Close: Set RS1 = Nothing
          Ac_HabDes 1
          Ac_Boton 3
          SSTab1.Tab = 0
       ElseIf RS1!NReg > 0 Then
          modo = "M"
          RS1.Close: Set RS1 = Nothing
          If modo = "A" Then
             SSTab1.Tab = 0
          ElseIf modo = "M" Then
             MoverDetalleDatos
'             SSTab1.TabEnabled(2) = True
          End If
          Ac_HabDes 3
          Ac_Boton 2
       End If
    End If
  Case Is = vbCancel
    Exit Sub
End Select
End Sub

Sub Actualiza_Dato()

On Error GoTo Man_Error
ivalidar = 0
ValidarCampos
If ivalidar = 1 Then Exit Sub
If modo = "A" Then
   vg_db.BeginTrans
     vg_db.Execute "insert into Sdx_Usuario (loginusuario, passwordusuario, " & _
                   "nombre, telefono, oficina, departamento) " & _
                   "values ('" & LimpiaDato(Trim(fpText1(0).Text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText1(2).Text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText1(1).Text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText1(3).Text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText1(4).Text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText1(5).Text)) & "')"
     Set RS1 = vg_db.Execute("select * " & _
                  "From Sdx_Usuario " & _
                  "where loginusuario='" & LimpiaDato(Trim(fpText1(0).Text)) & "'", , adCmdText)
'     Set RS1 = vg_db.Execute("sod_i_usuario '" & xxLimpiaDato(Trim(fpText1(0).Text)) & "', '" & xxLimpiaDato(Trim(fpText1(2).Text)) & "', '" & xxLimpiaDato(Trim(fpText1(1).Text)) & "', '" & xxLimpiaDato(Trim(fpText1(3).Text)) & "', '" & xxLimpiaDato(Trim(fpText1(4).Text)) & "', '" & LimpiaDato(Trim(fpText1(5).Text)) & "'", , adCmdText)  'adCmdStoredProc)
     If Not RS1.EOF Then
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 1
        vaSpread1.TypeHAlign = 1
        vaSpread1.Text = Trim(RS1!loginusuario)
        vaSpread1.Col = 2
        vaSpread1.Text = Trim(RS1!Nombre)
     End If
     RS1.Close: Set RS1 = Nothing
     modo = "M"
   vg_db.CommitTrans
ElseIf modo = "M" Then
   vg_db.BeginTrans
     vg_db.Execute "update Sdx_Usuario set passwordusuario='" & LimpiaDato(Trim(fpText1(2).Text)) & "', nombre='" & LimpiaDato(Trim(fpText1(1).Text)) & "', telefono='" & LimpiaDato(Trim(fpText1(3).Text)) & "', oficina='" & LimpiaDato(Trim(fpText1(4).Text)) & "', departamento='" & LimpiaDato(Trim(fpText1(5).Text)) & "' WHERE loginusuario='" & LimpiaDato(Trim(fpText1(0).Text)) & "'"
     vaSpread1.Row = vaSpread1.ActiveRow
     vaSpread1.Col = 2: vaSpread1.Text = LimpiaDato(Trim(fpText1(1).Text))
   vg_db.CommitTrans
End If
modo = "M"
Ac_HabDes 3
Ac_Boton 2

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub ValidarCampos()
If ivalidar = 0 And fpText1(0).Text = "" Then ivalidar = 1: MsgBox "Debe Ingresar Login", vbExclamation + vbOKOnly, "Usuario": fpText1(0).SetFocus
If ivalidar = 0 And fpText1(1).Text = "" Then ivalidar = 1: MsgBox "Debe Ingresar Nombre", vbExclamation + vbOKOnly, "Usuario": fpText1(1).SetFocus
If ivalidar = 0 And fpText1(2).Text = "" Then ivalidar = 1: MsgBox "Debe Ingresar Pasword", vbExclamation + vbOKOnly, "Usuario": fpText1(2).SetFocus
If modo = "A" Then
   Set RS1 = vg_db.Execute("select * " & _
                "From Sdx_Usuario " & _
                "where loginusuario='" & LimpiaDato(Trim(fpText1(0).Text)) & "'", , adCmdText)
'   Set RS1 = vg_db.Execute("sod_s_usuario 2, '" & LimpiaDato(Trim(fpText1(0).Text)) & "', ''", , adCmdStoredProc)
   If Not RS1.EOF Then ivalidar = 1: MsgBox "Usuario ya existe...", vbExclamation + vbOKOnly, "Mantención de usuarios": RS1.Close: Set RS1 = Nothing: Exit Sub
End If
End Sub

Sub LimpiarVariable()
For i = 0 To 5
    vecdatos(i) = ""
    fpText1(i).Text = ""
    If modo = "M" And i = 0 Then
       fpText1(i).Enabled = False
    Else
       fpText1(i).Enabled = True
    End If
Next i
End Sub

Sub MoverDetalleDatos()
itexto = 1
LimpiarVariable
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: loginusuario = vaSpread1.Text
Set RS1 = vg_db.Execute("select * " & _
             "From Sdx_Usuario " & _
             "where loginusuario='" & loginusuario & "'", , adCmdText)
'Set RS1 = vg_db.Execute("sod_s_usuario 2, '" & loginusuario & "', ''", , adCmdStoredProc)
If Not RS1.EOF Then
   Do While Not RS1.EOF
      fpText1(0).Text = Trim(RS1!loginusuario)
      fpText1(1).Text = Trim(RS1!Nombre)
      fpText1(2).Text = Trim(RS1!passwordusuario)
      fpText1(3).Text = Trim(RS1!telefono)
      fpText1(4).Text = Trim(RS1!oficina)
      fpText1(5).Text = Trim(RS1!departamento)
      RS1.MoveNext
   Loop
Else
   Ac_HabDes 1
   Ac_Boton 3
End If
RS1.Close: Set RS1 = Nothing
itexto = 0
End Sub

Sub MoverDatosGrillas()
vaSpread1.MaxRows = 0
RS1.Open "select *  From a_usuarios order by usu_nombre", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = 1
      vaSpread1.Text = RS1!usu_codigo
      vaSpread1.Col = 2
      vaSpread1.Text = Trim(RS1!usu_nombre)
      RS1.MoveNext
   Loop
   vaSpread1.Row = 1: vaSpread1.Col = 1: vaSpread1.EditMode = True
Else
   Ac_HabDes 1
   Ac_Boton 3
End If
RS1.Close: Set RS1 = Nothing
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
End Sub

Function Ac_Boton(Boton As Integer)
Select Case Boton
  Case 1
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(8).Visible = True

    Toolbar1.Buttons(10).Visible = True
    Toolbar1.Buttons(11).Visible = False
    Toolbar1.Buttons(12).Visible = True
    Toolbar1.Buttons(13).Visible = False
    fpText.Enabled = False
  Case 2
    Toolbar1.Buttons(1).Visible = True: Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(3).Visible = True: Toolbar1.Buttons(4).Visible = False
    Toolbar1.Buttons(5).Visible = True: Toolbar1.Buttons(6).Visible = False
    
    Toolbar1.Buttons(7).Visible = True
    Toolbar1.Buttons(8).Visible = False

    Toolbar1.Buttons(10).Visible = False
    Toolbar1.Buttons(11).Visible = True
    Toolbar1.Buttons(12).Visible = False
    Toolbar1.Buttons(13).Visible = True
  Case 3
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(8).Visible = True

    Toolbar1.Buttons(10).Visible = False
    Toolbar1.Buttons(11).Visible = True
    Toolbar1.Buttons(12).Visible = False
    Toolbar1.Buttons(13).Visible = True

  Case 4
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(8).Visible = True

    Toolbar1.Buttons(10).Visible = False
    Toolbar1.Buttons(11).Visible = True
    Toolbar1.Buttons(12).Visible = False
    Toolbar1.Buttons(13).Visible = True

End Select
End Function

Function Ac_HabDes(Opcion As Integer)
Select Case Opcion
    Case 1
        fpText.Enabled = False
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
    Case 2
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.Tab = 1
    Case 3
        fpText1(0).Enabled = False
        fpText.Enabled = True
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
    Case 4
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
End Select
End Function



