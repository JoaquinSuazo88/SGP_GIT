VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form B_TipPla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo Plato"
   ClientHeight    =   4695
   ClientLeft      =   2100
   ClientTop       =   1290
   ClientWidth     =   5370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView TvwDir 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8281
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   4695
      Left            =   4830
      TabIndex        =   1
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   8281
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "B_TipPla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset, RS2 As New ADODB.Recordset
Dim i As Long
Dim dest As Node, sourcenode As Node, nd As Node, rootNode As Node
Dim nivel As String, Nombre As String
Dim codigo As Long

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()

fg_centra Me
Me.Left = vg_left
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
codigo = 0
MoverDatosTvwDir

End Sub

Sub MoverDatosTvwDir()

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset
fg_carga (ss)
TvwDir.Nodes.Clear

Set RS1 = vg_db.Execute("SELECT * FROM a_recetatippla WHERE tip_previo=0 ORDER BY tip_codigo")
If Not RS1.EOF Then
   
   Do While Not RS1.EOF
      
      Set rootNode = TvwDir.Nodes.Add(, , "H" & RS1!tip_codigo, Trim(RS1!tip_nombre))
      ' agregar un nodo hijo postizo, si fuera necesario
      
      If rootNode.Children = 0 Then
         
         Set RS2 = vg_db.Execute("SELECT DISTINCT tip_codigo FROM a_recetatippla WHERE tip_previo=" & RS1!tip_codigo & "")
         
         If Not RS2.EOF Then
            
            ' la propiedad Texto de los nodos postizos es "*"
            TvwDir.Nodes.Add rootNode.Index, tvwChild, , "*"
         
         End If
         
         RS2.Close: Set RS2 = Nothing
      
      End If
      
      RS1.MoveNext
   Loop

End If
RS1.Close: Set RS1 = Nothing

fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Tipo Plato"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

Case 1
    
    If TvwDir.Nodes.count < 1 Then Exit Sub
    vg_nombre = TvwDir.SelectedItem.FullPath
    vg_codigo = codigo

End Select

Cerrar

End Sub

Private Sub tvwDir_Expand(ByVal Node As MSComctlLib.Node)

Dim dest1 As Node
Set dest1 = Node
If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
If Node.Child.text <> "*" Then Exit Sub
' eliminar el elemento hijo positivo
TvwDir.Nodes.Remove Node.Child.Index

Set RS1 = vg_db.Execute("SELECT * FROM a_recetatippla WHERE tip_previo=" & Val(Mid(TvwDir.Nodes(dest1.Index).Key, 2, 20)) & " ORDER BY tip_codigo")
If Not RS1.EOF Then
   
   Do While Not RS1.EOF
      
      Set nd = TvwDir.Nodes.Add(dest1.Index, tvwChild, "H" & RS1!tip_codigo, RS1!tip_nombre)
      
      If nd.Children = 0 Then
         
         Set RS2 = vg_db.Execute("SELECT DISTINCT tip_codigo FROM a_recetatippla WHERE tip_previo=" & RS1!tip_codigo & "")
         
         If Not RS2.EOF Then
            
            ' la propiedad Texto de los nodos postizos es "*"
            TvwDir.Nodes.Add rootNode.Index, tvwChild, , "*"
         
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

Case 27
    
    Cerrar

End Select

End Sub

Private Sub TvwDir_NodeClick(ByVal Node As MSComctlLib.Node)

Set dest = Node
nivel = Mid(TvwDir.Nodes(dest.Index).Key, 1, 1)
codigo = Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20))
If nivel = "H" Then vg_nombre = Nombre & "\" & Trim((TvwDir.Nodes(dest.Index).text))

End Sub

Sub Cerrar()

Me.Hide
Unload Me

End Sub
