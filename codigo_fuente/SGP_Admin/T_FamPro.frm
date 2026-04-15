VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form T_FamPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Familias de Producto"
   ClientHeight    =   5235
   ClientLeft      =   2220
   ClientTop       =   5100
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   6405
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
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "T_FamPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codigo As Long, codpadre As Long
Dim ivalidar As Integer
Dim modo As String, Nombre As String
Dim dest As Node, sourcenode As Node, nd As Node, rootNode As Node

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
Me.Height = 5190
Me.Width = 6120
MsgTitulo = "Categoría de Productos"
fg_centra Me
modo = "M"
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
MoverDatosGrillas

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Resize()

On Error GoTo Man_Error

If Me.WindowState <> 1 Then TvwDir.Move 0, Toolbar1.Height, ScaleWidth, ScaleHeight - Toolbar1.Height

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

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
    
        On Error Resume Next
        SendKeys "{enter}"
        On Error Resume Next
        
        TvwDir.SetFocus
        
        Actualiza_Dato
        TvwDir.SetFocus
    
    Case 15
        
        Set dest = TvwDir.SelectedItem
        If dest Is Nothing Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        i_Fampro
    
    Case 18
        
        Me.Hide
        Unload Me
    
    End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub MoverDatosGrillas()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

fg_carga ""
TvwDir.Nodes.Clear
Set nd = TvwDir.Nodes.Add(, , "R", "Familia de Producto ")
Set dest = nd

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_familiaproducto 4, 0,''")

If Not RS.EOF Then
    
    Do While Not RS.EOF
        
        Set rootNode = TvwDir.Nodes.Add(nd, tvwChild, "H" & RS!tip_codigo, Trim(RS!tip_nombre))
        ' agregar un nodo hijo postizo, si fuera necesario
        If rootNode.Children = 0 Then
           
           If RS!hijo = 1 Then
                
                ' la propiedad Texto de los nodos postizos es "***"
                TvwDir.Nodes.Add rootNode.Index, tvwChild, , "*"
            
            End If
        
        End If
        
        RS.MoveNext
    
    Loop

Else
   
   Gl_Ac_Botones Me, 1, 2, modo

End If
RS.Close
Set RS = Nothing
TvwDir.Nodes.item(dest.Key).Selected = True
fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Actualiza_Dato()

On Error GoTo Man_Error

Dim RS        As New ADODB.Recordset
Dim indice    As Long
Dim indtec    As String
Dim nivpro    As String
Dim arvpro    As String
Dim inipro    As Long
Dim finpro    As Long
Dim cdproduto As String

If modo = "A" Or modo = "M" Then
    
    ivalidar = 0
    
    ValidarCampos
    
    If ivalidar = 1 Then Exit Sub
    
    If modo = "A" Then
        
        indice = 0
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("sgpadm_iu_fampro 'A', 0, '" & Mid(LimpiaDato(Trim(TvwDir.Nodes(dest.Index).text)), 1, 35) & "', " & codpadre & "")
        
        If Not RS.EOF Then
           
           indice = RS!indice
        
        End If
        RS.Close
        Set RS = Nothing
        
        '------- Crear arbol de producto tecfood
        If indice > 0 Then
           
           nivpro = "1"
           arvpro = ""
           inipro = 1
           finpro = 1
           
           If codpadre <> 0 Then
              
              If RS.State = 1 Then RS.Close
              RS.CursorLocation = adUseClient
              vg_db.CursorLocation = adUseClient
              
              Set RS = vg_db.Execute("SELECT tip_codigo, tip_previo FROM a_tipopro WHERE tip_codigo = " & codpadre & "")
              
              If Not RS.EOF Then arvpro = CStr(RS!tip_codigo): nivpro = IIf(RS!tip_previo = 0, "2", "3"): inipro = IIf(RS!tip_previo = 0, 2, 4): finpro = 2
              
              RS.Close
              Set RS = Nothing
           
           End If
           TvwDir.Nodes(dest.Index).Key = "H" & indice
           dest.EnsureVisible
        
        End If
        
        '------- Fin crear arbol de producto tecfood
    ElseIf modo = "M" Then
        
        vg_db.Execute "sgpadm_iu_fampro  'M', " & Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20)) & ", '" & Mid(LimpiaDato(Trim(TvwDir.Nodes(dest.Index).text)), 1, 35) & "', 0"
    
    End If
    modo = "M"
    Gl_Ac_Botones Me, 1, 1, modo

End If

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub ValidarCampos()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If ivalidar = 0 And LimpiaDato(Trim(TvwDir.Nodes(dest.Index).text)) = "" Then
    
    ivalidar = 1
    MsgBox "Descripción debe ser informada...", vbExclamation + vbOKOnly, MsgTitulo
    Set TvwDir.SelectedItem = dest
    TvwDir.StartLabelEdit
    Exit Sub

ElseIf ivalidar = 0 Then
    
    If modo = "M" Then
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        Set RS = vg_db.Execute("SELECT * FROM a_tipopro WHERE tip_codigo=" & codpadre & "")
        If Not RS.EOF Then codigo = RS!tip_codigo
        RS.Close
        Set RS = Nothing
    
    Else
        
        codigo = codpadre
    
    End If

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Agrega_Fila()

On Error GoTo Man_Error

Set dest = TvwDir.SelectedItem
a = TvwDir.SelectedItem.FullPath
a = TvwDir.PathSeparator
codpadre = Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20))
tvwDir_Expand dest
Set sourcenode = TvwDir.Nodes.Add(dest.Index, tvwChild, "H" & 999999999, "")
Set TvwDir.SelectedItem = sourcenode
Set dest = sourcenode
TvwDir.StartLabelEdit
modo = "A"
Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Altera_Fila()

On Error GoTo Man_Error

If TvwDir.SelectedItem.Index = 1 Then Exit Sub
modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
Set dest = TvwDir.SelectedItem
codpadre = Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20))
Nombre = TvwDir.Nodes(dest.Index).text
TvwDir.StartLabelEdit

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Borra_Fila()

On Error GoTo Man_Error

Dim RS     As New ADODB.Recordset
Dim RS1    As New ADODB.Recordset
Dim sw     As Integer
Dim nivpro As String

If TvwDir.SelectedItem.Index = 1 Then Exit Sub
Set dest = TvwDir.SelectedItem

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_s_familiaproducto 2," & Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20)) & ",''")
sw = RS("tip_codigo")
If Not RS.EOF Then
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS1 = vg_db.Execute("sgpadm_s_familiaproducto 5," & sw & ",''")
   
   If Not RS1.EOF Then
      
      MsgBox "No se puede Eliminar Dato, contiene subcategorias...", vbCritical + vbOKOnly, MsgTitulo
      RS.Close
      Set RS = Nothing
      
      RS1.Close
      Set RS1 = Nothing
      TvwDir.SetFocus
      Exit Sub
   
   End If
   
   RS1.Close
   Set RS1 = Nothing

End If
RS.Close
Set RS = Nothing

If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub

    vg_db.Execute "DELETE FROM a_tipopro WHERE tip_codigo=" & Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20)) & ""
    TvwDir.Nodes.Remove dest.Index

    TvwDir.SetFocus

If dest Is Nothing Then
   
   Gl_Ac_Botones Me, 1, 2, modo

Else
   
   Gl_Ac_Botones Me, 1, 1, modo

End If

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Cancela_Fila()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Resp_Cancel ("Mantención")

If respuesta = vbYes Then
   
   If modo = "A" Then
      
      TvwDir.Nodes.Remove dest.Index
      TvwDir.SetFocus
      modo = "M"
   
   Else
      
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      Set RS = vg_db.Execute("sgpadm_s_familiaproducto 2, " & Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20)) & ",''")
      If Not RS.EOF Then TvwDir.Nodes(dest.Index).text = Trim(RS!tip_nombre)
      RS.Close
      Set RS = Nothing
      TvwDir.SetFocus
   
   End If
   
   If dest Is Nothing Then Gl_Ac_Botones Me, 1, 2, modo Else Gl_Ac_Botones Me, 1, 1, modo

Else
   
   If modo = "A" Then
      
      Set TvwDir.SelectedItem = dest
      codpadre = Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20))
      TvwDir.StartLabelEdit
   
   ElseIf modo = "M" Then
      
      Set dest = TvwDir.SelectedItem
      codpadre = Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20))
      TvwDir.StartLabelEdit
   
   End If

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub tvwDir_AfterLabelEdit(Cancel As Integer, NewString As String)

On Error GoTo Man_Error

TvwDir.Nodes(dest.Index).text = LimpiaDato(Trim(NewString))
Actualiza_Dato

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub tvwDir_BeforeLabelEdit(Cancel As Integer)

On Error GoTo Man_Error

If TvwDir.SelectedItem.Index = 1 Then Cancel = 1: Exit Sub

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub TvwDir_Collapse(ByVal Node As MSComctlLib.Node)

On Error GoTo Man_Error

If Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True Then ValidarCampos

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub TvwDir_Click()

On Error GoTo Man_Error

If Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True Then Actualiza_Dato

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub tvwDir_Expand(ByVal Node As MSComctlLib.Node)

On Error GoTo Man_Error

Dim RS    As New ADODB.Recordset
Dim dest1 As Node

Set dest1 = Node
If Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True Then
   
   ValidarCampos

End If
If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
If Node.Child.text <> "*" Then Exit Sub
' eliminar el elemento hijo positivo
TvwDir.Nodes.Remove Node.Child.Index

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_s_familiaproducto 3, " & Val(Mid(TvwDir.Nodes(dest1.Index).Key, 2, 20)) & ",''")
If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      Set nd = TvwDir.Nodes.Add(dest1.Index, tvwChild, "H" & RS!tip_codigo, RS!tip_nombre)
      
      If nd.Children = 0 Then
         
         If RS!hijo = 1 Then
            
            ' la propiedad Texto de los nodos positivos es "***"
            TvwDir.Nodes.Add nd.Index, tvwChild, , "*"
         
         End If
      
      End If
      
      RS.MoveNext
   
   Loop

End If
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub TvwDir_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

    Case 13 And Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True
    
        Actualiza_Dato
    
    Case 27 And Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True
        
        Cancela_Fila
    
    Case 113 And Toolbar1.Buttons(1).Visible = True
        
        Agrega_Fila
    
    Case 119 And Toolbar1.Buttons(3).Visible = True
        
        Altera_Fila
    
    Case 115 And Toolbar1.Buttons(5).Visible = True
        
        Borra_Fila

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub
