VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form T_TipPla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Plato"
   ClientHeight    =   4815
   ClientLeft      =   2835
   ClientTop       =   2505
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView TvwDir 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7858
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "T_TipPla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim codigo As Long, codpadre As Long
Dim ivalidar As Integer, CodPat As Integer
Dim modo As String, Nombre As String, Msgtitulo As String
Dim dest As Node, sourcenode As Node, nd As Node, rootnode As Node

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 5190
Me.Width = 6120
fg_centra Me
modo = "M"
Msgtitulo = "Tipo de Plato"
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, IIf(vg_modrec = True, 1, 6), modo
MoverDatosGrillas
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then TvwDir.Move 0, Toolbar1.Height, ScaleWidth, ScaleHeight - Toolbar1.Height
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    Agrega_Fila
Case 3
    Altera_Fila
Case 5
    Borra_Fila
Case 7
    MoverDatosGrillas
Case 10
    Cancela_Fila
Case 12
    SendKeys "{enter}"
    Actualiza_Dato
    TvwDir.SetFocus
Case 15
    Set dest = TvwDir.SelectedItem
    If dest Is Nothing Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_TipPla
Case 18
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub MoverDatosGrillas()
fg_carga ""
TvwDir.Nodes.Clear
Set nd = TvwDir.Nodes.Add(, , "R", "Tipo de Plato ")
Set dest = nd
RS1.Open RutinaLectura.RecetaTipoPlato(1, 0, ""), vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
        Set rootnode = TvwDir.Nodes.Add(nd, tvwChild, "H" & RS1!tip_codigo, Trim(RS1!tip_nombre))
            ' agregar un nodo hijo postizo, si fuera necesario
        If rootnode.Children = 0 Then
           RS2.Open RutinaLectura.RecetaTipoPlato(2, RS1!tip_codigo, ""), vg_db, adOpenStatic
            If Not RS2.EOF Then
               ' la propiedad Texto de los nodos postizos es "***"
               TvwDir.Nodes.Add rootnode.Index, tvwChild, , "*"
            End If
            RS2.Close: Set RS2 = Nothing
        End If
        RS1.MoveNext
    Loop
Else
    Gl_Ac_Botones Me, 1, IIf(vg_modrec = True, 2, 6), modo
End If
RS1.Close: Set RS1 = Nothing
TvwDir.Nodes.Item(dest.Key).Selected = True
fg_descarga
End Sub

Private Sub Actualiza_Dato()
Dim indice As Long
On Error GoTo Man_Error
If modo = "A" Or modo = "M" Then
   ivalidar = 0
   ValidarCampos
   If ivalidar = 1 Then Exit Sub
   If modo = "A" Then
      vg_db.BeginTrans
      RS1.Open RutinaLectura.RecetaTipoPlato(3, 0, ""), vg_db, adOpenStatic
      If Not RS1.EOF Then RS1.MoveFirst: indice = RS1!tip_codigo + 1 Else indice = 1
      RS1.Close: Set RS1 = Nothing
      vg_db.Execute "INSERT INTO a_recetatippla VALUES (" & indice & ", '" & Mid(LimpiaDato(Trim(TvwDir.Nodes(dest.Index).text)), 1, 50) & "'," & codpadre & ")"
      vg_db.CommitTrans
      TvwDir.Nodes(dest.Index).Key = "H" & indice
      dest.EnsureVisible
  ElseIf modo = "M" Then
      vg_db.BeginTrans
      vg_db.Execute "UPDATE a_recetatippla SET tip_nombre = '" & Mid(LimpiaDato(Trim(TvwDir.Nodes(dest.Index).text)), 1, 50) & "' WHERE tip_codigo = " & Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20)) & ""
      vg_db.CommitTrans
  End If
  modo = "M"
  Gl_Ac_Botones Me, 1, 1, modo
End If
Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub ValidarCampos()
Dim sql1 As String
If ivalidar = 0 And LimpiaDato(Trim(TvwDir.Nodes(dest.Index).text)) = "" Then
    ivalidar = 1: MsgBox "Descripción Debe ser Informada", vbExclamation + vbOKOnly, Msgtitulo
    Set TvwDir.SelectedItem = dest
    TvwDir.StartLabelEdit
    Exit Sub
ElseIf ivalidar = 0 Then
    If modo = "M" Then
        RS1.Open RutinaLectura.RecetaTipoPlato(4, codpadre, ""), vg_db, adOpenStatic
        If Not RS1.EOF Then codigo = RS1!tip_previo
        RS1.Close: Set RS1 = Nothing
    Else
        codigo = codpadre
    End If
    RS1.Open RutinaLectura.RecetaTipoPlato(5, codigo, UCase(LimpiaDato(Trim(TvwDir.Nodes(dest.Index).text)))), vg_db, adOpenStatic
    If Not RS1.EOF Then
        If RS1!nreg = 1 And Nombre <> TvwDir.Nodes(dest.Index).text Then
            RS1.Close: Set RS1 = Nothing
            ivalidar = 1: MsgBox "Ya existe descripción", vbExclamation + vbOKOnly, Msgtitulo
            Set TvwDir.SelectedItem = dest
            TvwDir.StartLabelEdit
            Exit Sub
        Else
            RS1.Close: Set RS1 = Nothing
            Exit Sub
        End If
    Else
        RS1.Close: Set RS1 = Nothing
    End If
End If
End Sub

Private Sub Agrega_Fila()
Set dest = TvwDir.SelectedItem
codpadre = Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20))
tvwDir_Expand dest
Set sourcenode = TvwDir.Nodes.Add(dest.Index, tvwChild, "H" & 999999999, "")
Set TvwDir.SelectedItem = sourcenode
Set dest = sourcenode
TvwDir.StartLabelEdit
modo = "A"
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub Altera_Fila()
If TvwDir.SelectedItem.Index = 1 Then Exit Sub
modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
Set dest = TvwDir.SelectedItem
codpadre = Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20))
Nombre = TvwDir.Nodes(dest.Index).text
TvwDir.StartLabelEdit
End Sub

Private Sub Borra_Fila()
On Error GoTo Man_Error
If TvwDir.SelectedItem.Index = 1 Then Exit Sub
Set dest = TvwDir.SelectedItem
RS1.Open RutinaLectura.RecetaTipoPlato(4, Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20)), ""), vg_db, adOpenStatic
If Not RS1.EOF Then
    RS2.Open RutinaLectura.RecetaTipoPlato(2, RS1!tip_codigo, ""), vg_db, adOpenStatic
    If Not RS2.EOF Then
        MsgBox "No se puede Eliminar Dato, esta asociado a un Tipo Plato", vbCritical + vbOKOnly, Msgtitulo
        RS1.Close: Set RS1 = Nothing: RS2.Close: Set RS2 = Nothing
        TvwDir.SetFocus
        Exit Sub
    End If
    RS2.Close: Set RS2 = Nothing
End If
RS1.Close: Set RS1 = Nothing
If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
RS1.Open RutinaLectura.RecetaTipoPlato(4, Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20)), ""), vg_db, adOpenStatic
If Not RS1.EOF Then CodPat = RS1!tip_previo
RS1.Close: Set RS1 = Nothing
vg_db.BeginTrans
vg_db.Execute "DELETE a_recetatippla FROM a_recetatippla WHERE tip_codigo = " & Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20))
TvwDir.Nodes.Remove dest.Index
vg_db.CommitTrans
TvwDir.SetFocus
If dest Is Nothing Then Gl_Ac_Botones Me, 1, 2, modo Else Gl_Ac_Botones Me, 1, 1, modo
Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Cancela_Fila()
If MsgBox("Cancela registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
If modo = "A" Then
   TvwDir.Nodes.Remove dest.Index
   TvwDir.SetFocus
   modo = "M"
Else
   RS1.Open RutinaLectura.RecetaTipoPlato(4, Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20)), ""), vg_db, adOpenStatic
   If Not RS1.EOF Then TvwDir.Nodes(dest.Index).text = Trim(RS1!tip_nombre)
   RS1.Close: Set RS1 = Nothing
   TvwDir.SetFocus
End If
If dest Is Nothing Then Gl_Ac_Botones Me, 1, 2, modo Else Gl_Ac_Botones Me, 1, 1, modo
If modo = "A" Then
   Set TvwDir.SelectedItem = dest
   codpadre = Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20))
   TvwDir.StartLabelEdit
ElseIf modo = "M" Then
   Set dest = TvwDir.SelectedItem
   codpadre = Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20))
   TvwDir.StartLabelEdit
End If
End Sub

Private Sub tvwDir_AfterLabelEdit(Cancel As Integer, NewString As String)
TvwDir.Nodes(dest.Index).text = LimpiaDato(Trim(NewString))
Actualiza_Dato
End Sub

Private Sub tvwDir_BeforeLabelEdit(Cancel As Integer)
If TvwDir.SelectedItem.Index = 1 Then Cancel = 1: Exit Sub
End Sub

Private Sub TvwDir_Collapse(ByVal Node As MSComctlLib.Node)
If Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True Then ValidarCampos
End Sub

Private Sub TvwDir_Click()
If Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True Then Actualiza_Dato
End Sub

Private Sub tvwDir_Expand(ByVal Node As MSComctlLib.Node)
Dim dest1 As Node
Set dest1 = Node
If Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True Then ValidarCampos
If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
If Node.Child.text <> "*" Then Exit Sub
'------- Eliminar el elemento hijo positivo
TvwDir.Nodes.Remove Node.Child.Index
RS1.Open RutinaLectura.RecetaTipoPlato(6, Val(Mid(TvwDir.Nodes(dest1.Index).Key, 2, 20)), ""), vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      Set nd = TvwDir.Nodes.Add(dest1.Index, tvwChild, "H" & RS1!tip_codigo, RS1!tip_nombre)
      If nd.Children = 0 Then
         Set RS2 = vg_db.Execute(RutinaLectura.RecetaTipoPlato(2, RS1!tip_codigo, ""), , adCmdText)
         If Not RS2.EOF Then
            '------- la propiedad Texto de los nodos positivos es
            TvwDir.Nodes.Add nd.Index, tvwChild, , "*"
         End If
         RS2.Close: Set RS2 = Nothing
      End If
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
End Sub

Private Sub TvwDir_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13 And Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True
    Actualiza_Dato
Case 27 And Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True
    Cancela_Fila
Case 113 And Toolbar1.Buttons(1).Visible = True
    Agrega_Fila
Case 114 And Toolbar1.Buttons(3).Visible = True
    Altera_Fila
Case 115 And Toolbar1.Buttons(5).Visible = True
    Borra_Fila
End Select
End Sub
