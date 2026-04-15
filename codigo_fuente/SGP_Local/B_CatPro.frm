VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form B_CatPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categoria Producto"
   ClientHeight    =   4695
   ClientLeft      =   2100
   ClientTop       =   1290
   ClientWidth     =   5370
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
Attribute VB_Name = "B_CatPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConSql As ADODB.Recordset
Dim i As Long, ibusca As Long
Dim dest As Node, sourcenode As Node, nd As Node, rootnode As Node
Private Sub Form_Activate()
fg_descarga
End Sub
Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me
Me.Left = vg_left
fg_carga (ss)
Mover_Botones
TvwDir.Nodes.Clear
Set ConSql = vg_db.Execute("select * " & _
             "From PB00353 " & _
             "Where Own_Anal_Code = 0 " & _
             "order by Anal_Code", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_categoriaprod 1, " & 0 & ", ''", , adCmdStoredProc)
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
'      Set rootnode = TvwDir.Nodes.Add(nd, tvwChild, "H" & ConSql!Anal_Code, Trim(ConSql!Anal_Code_Desc))
      Set rootnode = TvwDir.Nodes.Add(, , "H" & ConSql!Anal_Code, ConSql!Anal_Code & "  " & Trim(ConSql!Anal_Code_Desc))
          ' agregar un nodo hijo postizo, si fuera necesario
      If rootnode.Children = 0 And ConSql!Anal_Code_Type = 1 Then
         ' la propiedad Texto de los nodos postizos es "*"
         TvwDir.Nodes.Add rootnode.Index, tvwChild, , "*"
      End If
      ConSql.MoveNext
   Loop
End If
ConSql.Close: Set ConSql = Nothing
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Categoria Producto"
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 1
    MoverDatos
  Case 3
    Cerrar
End Select
End Sub
Private Sub MoverDatos()
If TvwDir.Nodes.count > 0 Then
   Set dest = TvwDir.SelectedItem
   vg_codigo = Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20))
   Cerrar
End If
End Sub
Private Sub tvwDir_Expand(ByVal Node As MSComctlLib.Node)
Dim dest1 As Node
Set dest1 = Node
If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
If Node.Child.Text <> "*" Then Exit Sub
' eliminar el elemento hijo positivo
TvwDir.Nodes.Remove Node.Child.Index
Set ConSql = vg_db.Execute("select * " & _
             "From PB00353 " & _
             "where Own_Anal_Code=" & Val(Mid(TvwDir.Nodes(dest1.Index).Key, 2, 20)) & " " & _
             "order by Anal_Code", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_categoriaprod 9, " & Val(Mid(TvwDir.Nodes(dest1.Index).Key, 2, 20)) & ", ''", , adCmdStoredProc)
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
      Set nd = TvwDir.Nodes.Add(dest1.Index, tvwChild, "H" & ConSql!Anal_Code, ConSql!Anal_Code & "  " & ConSql!Anal_Code_Desc)
      If nd.Children = 0 And ConSql!Anal_Code_Type = 1 Then
         ' la propiedad Texto de los nodos positivos es "***"
         TvwDir.Nodes.Add nd.Index, tvwChild, , "*"
      End If
      ConSql.MoveNext
   Loop
End If
ConSql.Close: Set ConSql = Nothing
End Sub
Private Sub TvwDir_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 27
    Cerrar
End Select
End Sub
Sub Cerrar()
Me.Hide
Unload Me
End Sub
Sub Mover_Botones()

   Toolbar1.ImageList = Partida.IL1
   Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar "
   Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
   Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"

End Sub

