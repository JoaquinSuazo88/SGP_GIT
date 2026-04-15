VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form B_ArbEst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   4695
   ClientLeft      =   3750
   ClientTop       =   1980
   ClientWidth     =   5085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView TvwDir 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4540
      _ExtentX        =   8017
      _ExtentY        =   8281
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4500
      Top             =   4200
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
            Picture         =   "B_ArbEst.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_ArbEst.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_ArbEst.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_ArbEst.frx":078E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   4695
      Left            =   4545
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
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "B_ArbEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dest As Node, sourcenode As Node, nd As Node, nd2 As Node, nd3 As Node, rootNode As Node
Dim nivel As String, Nombre As String
Dim codigo As Long
Dim Tabla As String, Suf As String

Private Sub Agrega_Fila(RS2 As ADODB.Recordset)

Set nd3 = TvwDir.SelectedItem
codpadre = Val(Mid(TvwDir.Nodes(nd3.Index).Key, 2, 20))
'tvwDir_Expand nd3
Set sourcenode = TvwDir.Nodes.Add(nd3.Index, tvwChild, "J" & RS2(0), RS2(1))
Set TvwDir.SelectedItem = sourcenode
Set nd3 = sourcenode
nd3.ExpandedImage = 4

End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()

fg_centra Me
Me.Left = vg_left
nivel = "R"
codigo = 0

End Sub

Sub MoverDatosTvwDir(TablaGen As String, SufGen As String, titgen As String, Optional op As String)

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset

Tabla = TablaGen
Suf = SufGen
Me.Caption = titgen
fg_carga ""
TvwDir.Nodes.Clear

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_Sel_ArbolCatDieteticaTipPlato_V01 " & IIf(Tabla = "a_recetacatdie", 1, IIf(Tabla = "a_tipopro", 3, 5)) & ", 0")

If Not RS1.EOF Then
   
   If op = "1" Then
      
      Set rootNode = TvwDir.Nodes.Add(, , "R" & "Todos", "Todos", 3)
      op = ""
      vg_nombre = "Todos"
   
   End If
    
   Do While Not RS1.EOF
      Set rootNode = TvwDir.Nodes.Add(, , "R" & RS1(0), Trim(RS1(1)), 3)
          ' agregar un nodo hijo postizo, si fuera necesario
      rootNode.ExpandedImage = 4
     If rootNode.Children = 0 And RS1(2) = 1 Then
                
        ' la propiedad Texto de los nodos postizos es "***"
        TvwDir.Nodes.Add rootNode.Index, tvwChild, , "*"
        
     End If
     RS1.MoveNext
    
    Loop

    Nombre = Trim((TvwDir.Nodes(1).text))
    codigo = Val(Mid(TvwDir.Nodes(1).Key, 2, 20)):

End If

RS1.Close
Set RS1 = Nothing
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Categoria Dietetica"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

Case 1
    
    Screen.MousePointer = 11
    If TvwDir.Nodes.count < 1 Then Exit Sub
    If codigo > 0 Then vg_nombre = TvwDir.SelectedItem.FullPath    'Nombre
    vg_codigo = codigo
    
End Select

Cerrar

End Sub

Private Sub tvwDir_Expand(ByVal Node As MSComctlLib.Node)

Dim RS1 As New ADODB.Recordset

Set dest = Node
If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
If Node.Child.text <> "*" And Node.Child.text <> "**" Then Exit Sub
' eliminar el elemento hijo positivo
TvwDir.Nodes.Remove Node.Child.Index

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_Sel_ArbolCatDieteticaTipPlato_V01 " & IIf(Tabla = "a_recetacatdie", 2, IIf(Tabla = "a_tipopro", 4, 6)) & ", " & Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20)) & "")
If Not RS1.EOF Then
   
   Do While Not RS1.EOF
      
      Set nd = TvwDir.Nodes.Add(dest.Index, tvwChild, "H" & RS1(0), RS1(1), 4)
      
      dest.ExpandedImage = 4
      
      If nd.Children = 0 And RS1(2) = 1 Then
            
            ' la propiedad Texto de los nodos positivos es "***"
            TvwDir.Nodes.Add nd.Index, tvwChild, , "**"
      
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

If nivel = "R" Then
   
   Nombre = Trim((TvwDir.Nodes(dest.Index).text))

ElseIf nivel = "H" Then
   
   vg_nombre = Nombre & "\" & Trim((TvwDir.Nodes(dest.Index).text)) ' & "\" & Trim((TvwDir.Nodes(nd.Index).Text))
   Nombre = vg_nombre

End If

End Sub

Sub Cerrar()

Me.Hide
Unload Me

End Sub
