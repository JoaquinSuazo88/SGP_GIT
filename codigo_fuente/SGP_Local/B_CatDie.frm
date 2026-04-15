VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form B_CatDie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categoria Dietetica"
   ClientHeight    =   4695
   ClientLeft      =   4620
   ClientTop       =   1650
   ClientWidth     =   5085
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
            Picture         =   "B_CatDie.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_CatDie.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_CatDie.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_CatDie.frx":078E
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
Attribute VB_Name = "B_CatDie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset, RS2 As New ADODB.Recordset
Dim dest As Node, sourcenode As Node, nd As Node, rootnode As Node
Dim nivel As String, Nombre As String
Dim codigo As Long
Dim Tabla As String, Suf As String
Private Sub Form_Activate()
fg_descarga
End Sub
Private Sub Form_Load()
fg_centra Me
Me.Left = vg_left
nivel = "R"
codigo = 0
'MoverDatosTvwDir
End Sub
Sub MoverDatosTvwDir(TablaGen As String, SufGen As String, TitGen As String)

On Error GoTo Man_Error
Tabla = TablaGen
Suf = SufGen
Me.Caption = TitGen
fg_carga "ss"
TvwDir.Nodes.Clear
RS1.Open "select distinct " & Suf & "codigo, " & Suf & "nombre " & _
         "from  " & Tabla & " " & _
         "where " & Suf & "previo=0", vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
        Set rootnode = TvwDir.Nodes.Add(, , "R" & RS1(0), Trim(RS1(1)), 3)
            ' agregar un nodo hijo postizo, si fuera necesario
        If rootnode.Children = 0 Then
            RS2.Open "select distinct " & Suf & "codigo " & _
                     "from  " & Tabla & " " & _
                     "where " & Suf & "previo=" & RS1(0) & "", vg_db, adOpenStatic
           If Not RS2.EOF Then
                ' la propiedad Texto de los nodos postizos es "***"
                TvwDir.Nodes.Add rootnode.Index, tvwChild, , "**"
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
MsgBox Err & ":  " & Error$(Err), vbCritical, "Categoria Dietetica"
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 1
    If TvwDir.Nodes.Count < 1 Then Exit Sub
    vg_nombre = TvwDir.SelectedItem.FullPath
    vg_codigo = codigo
End Select
Cerrar
End Sub
Private Sub tvwDir_Expand(ByVal Node As MSComctlLib.Node)
Set dest = Node
TvwDir_NodeClick dest
If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
If Node.Child.Text <> "**" Then Exit Sub
If Node.Child.Text = "**" Then
   ' eliminar el elemento hijo positivo
   TvwDir.Nodes.Remove Node.Child.Index
   RS1.Open "select * from  " & Tabla & " where " & Suf & "previo=" & Val(Mid(TvwDir.Nodes(dest.Index).Key, 2, 20)) & " " & _
            "order by " & Suf & "nombre", vg_db, adLockReadOnly
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         Set nd = TvwDir.Nodes.Add(dest.Index, tvwChild, "H" & RS1(0), RS1(1), 3)
         dest.ExpandedImage = 4
         RS1.MoveNext
      Loop
   End If
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
   Nombre = Trim((TvwDir.Nodes(dest.Index).Text))
ElseIf nivel = "H" Then
   vg_nombre = Trim((TvwDir.Nodes(dest.Index).Text))
   'vg_nombre = Nombre & "\" & Trim((TvwDir.Nodes(dest.Index).Text))
End If
End Sub
Sub Cerrar()
Me.Hide
Unload Me
End Sub
