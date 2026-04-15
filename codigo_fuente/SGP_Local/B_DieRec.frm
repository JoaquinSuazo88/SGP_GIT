VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form B_DieRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtrar Recetas"
   ClientHeight    =   5460
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9420
   Begin VB.Frame Frame1 
      Caption         =   "Tipo Plato"
      Height          =   5415
      Index           =   1
      Left            =   4440
      TabIndex        =   1
      Top             =   0
      Width           =   4335
      Begin MSComctlLib.TreeView TvwDir 
         Height          =   4935
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   8705
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Categoria Dietetica"
      Height          =   5415
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin MSComctlLib.TreeView TvwDir 
         Height          =   4935
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   8705
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_DieRec.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_DieRec.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_DieRec.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_DieRec.frx":094E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_DieRec.frx":0AA8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5460
      Left            =   8790
      TabIndex        =   4
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   9631
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Deshacer"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "B_DieRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Consql1 As ADODB.Recordset, Consql2 As ADODB.Recordset
Dim dest As Node, sourcenode As Node, nd As Node, rootnode As Node
Dim nivel As String, nomtplato1 As String, nomtplato2 As String, nomtplato3 As String, nomtplato4 As String
Dim nomdietetico As String
Dim auxdiet As Long, auxcat1 As Long, auxcat2 As Long, auxcat3 As Long, auxcat4 As Long
Dim Opcion As Integer
Private Sub Form_Activate()
fg_descarga
auxdiet = 0: auxcat1 = 0: auxcat2 = 0: auxcat3 = 0: auxcat4 = 0: Opcion = 0
auxdiet = vg_codregimen: auxcat1 = vg_auxcategoria1: auxcat2 = vg_auxcategoria2: auxcat3 = vg_auxcategoria3: auxcat3 = vg_auxcategoria4
End Sub
Private Sub Form_Load()
fg_centra Me
Opcion = 0
nomtplato1 = "": nomtplato2 = "": nomtplato3 = "": nomtplato4 = "": nomdietetico = ""
MoverDatosTvwDir
End Sub
Sub MoverDatosTvwDir()
fg_carga "ss"
TvwDir(0).Nodes.Clear
Set Consql1 = vg_db.Execute("select distinct Diet_Grp_No, Diet_Grp_Desc " & _
              "From PB00384 " & _
              "order by Diet_Grp_No", , adCmdText)
'Set Consql1 = vg_db.Execute("sod_s_categoriadietetica 1, " & 0 & ", ''", , adCmdStoredProc)
If Not Consql1.EOF Then
   Do While Not Consql1.EOF
      Set rootnode = TvwDir(0).Nodes.Add(, , "R" & Consql1!Diet_Grp_No, Trim(Consql1!Diet_Grp_Desc), 4)
          ' agregar un nodo hijo postizo, si fuera necesario
      If rootnode.Children = 0 Then
         Set Consql2 = vg_db.Execute("select distinct Diet_Cat_Group_No " & _
                       "From  PB00358 " & _
                       "where PB00358.Diet_Cat_Group_No=" & Consql1!Diet_Grp_No & "", , adCmdText)
         If Not Consql2.EOF Then
            ' la propiedad Texto de los nodos postizos es "***"
            TvwDir(0).Nodes.Add rootnode.Index, tvwChild, , "*"
         End If
         Consql2.Close: Set Consql2 = Nothing
      End If
      Consql1.MoveNext
   Loop
End If
Consql1.Close: Set Consql1 = Nothing

TvwDir(1).Nodes.Clear
Set Consql1 = vg_db.Execute("select distinct Rcpe_Cat_No, Rcpe_Cat_Desc " & _
              "From PB00085 " & _
              "order by Rcpe_Cat_No", , adCmdText)
'Set Consql1 = vg_db.Execute("sod_s_categoriareceta 1, " & 0 & ", ''", , adCmdStoredProc)
If Not Consql1.EOF Then
   Set rootnode = TvwDir(1).Nodes.Add(, , "RT" & 0, "Todos", 4)
   Do While Not Consql1.EOF
      Set rootnode = TvwDir(1).Nodes.Add(, , "R" & Consql1!Rcpe_Cat_No, Trim(Consql1!Rcpe_Cat_Desc), 4)
          ' agregar un nodo hijo postizo, si fuera necesario
      If rootnode.Children = 0 Then
         Set Consql2 = vg_db.Execute("select distinct Prev_Rcpe_Cat_No " & _
                       "From  PB00086 " & _
                       "where PB00086.Prev_Rcpe_Cat_No=" & Consql1!Rcpe_Cat_No & "", , adCmdText)
         If Not Consql2.EOF Then
            ' la propiedad Texto de los nodos postizos es "***"
            TvwDir(1).Nodes.Add rootnode.Index, tvwChild, , "*"
         End If
         Consql2.Close: Set Consql2 = Nothing
      End If
      Consql1.MoveNext
   Loop
End If
Consql1.Close: Set Consql1 = Nothing
fg_descarga
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 1
    If TvwDir(0).Nodes.Count < 1 And TvwDir(1).Nodes.Count < 1 Then Exit Sub
    If nivel <> "" And nivel = "R" Then
      MsgBox " Categoria Dietetica Seleccionada Debe Ser De Un Nivel Mas Bajo "
      Exit Sub
    End If
    vg_codregimen = auxdiet: vg_auxcategoria1 = auxcat1: vg_auxcategoria2 = auxcat2: vg_auxcategoria3 = auxcat3: vg_auxcategoria4 = auxcat4
    If nomdietetico <> "" Then vg_descdieteticorecet = nomdietetico
    If nomtplato1 <> "" Then vg_desctplatorecet = nomtplato1
    If nomtplato2 <> "" Then vg_desctplatorecet = vg_desctplatorecet & "\" & nomtplato2
    If nomtplato3 <> "" Then vg_desctplatorecet = vg_desctplatorecet & "\" & nomtplato3
    If nomtplato4 <> "" Then vg_desctplatorecet = vg_desctplatorecet & "\" & nomtplato4
    If Opcion = 0 Then vg_desctplatorecet = TvwDir(1).SelectedItem.FullPath
    vg_opcion = 0
    Me.Hide
  Case 3
    If TvwDir(0).Nodes.Count < 1 And TvwDir(1).Nodes.Count < 1 Then Exit Sub
    auxdiet = 0: auxcat1 = 0: auxcat2 = 0: auxcat3 = 0: auxcat4 = 0
    vg_codregimen = 0: vg_auxcategoria1 = 0: vg_auxcategoria2 = 0: vg_auxcategoria3 = 0: vg_auxcategoria4 = 0
    nomtplato1 = "Todos": nomtplato2 = "": nomtplato3 = "": nomtplato4 = "": nomdietetico = "Todos"
    MoverDatosTvwDir
    Opcion = 1
  Case 5
    vg_opcion = 2
    Cerrar
End Select
End Sub
Private Sub tvwDir_Expand(Index As Integer, ByVal Node As MSComctlLib.Node)
Set dest = Node

Select Case Index
  Case 0
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.Text <> "*" Then Exit Sub
    If Node.Child.Text = "*" Then
 ' eliminar el elemento hijo positivo
       TvwDir(0).Nodes.Remove Node.Child.Index
       Set Consql1 = vg_db.Execute("select PB00358.* " & _
                     "From PB00358, PB00384 " & _
                     "Where PB00358.Diet_Cat_Group_No = PB00384.Diet_Grp_No " & _
                     "and   PB00358.Diet_Cat_Group_No=" & Val(Mid(TvwDir(0).Nodes(dest.Index).Key, 2, 20)) & " " & _
                     "order by PB00358.Diet_Cat_No", , adCmdText)
'       Set Consql1 = vg_db.Execute("sod_s_categoriadietetica 6, " & Val(Mid(TvwDir(0).Nodes(dest.Index).Key, 2, 20)) & ", ''", , adCmdStoredProc)
       If Not Consql1.EOF Then
          Do While Not Consql1.EOF
             Set nd = TvwDir(0).Nodes.Add(dest.Index, tvwChild, "H" & Consql1!Diet_Cat_No, Consql1!Diet_Cat_Desc, 4)
             dest.ExpandedImage = 5
             Consql1.MoveNext
          Loop
       End If
       Consql1.Close: Set Consql1 = Nothing
    End If
  Case 1
    TvwDir_NodeClick Index, dest
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.Text <> "*" And Node.Child.Text <> "**" And Node.Child.Text <> "***" Then Exit Sub
    If Node.Child.Text = "*" Then
       ' eliminar el elemento hijo positivo
       TvwDir(1).Nodes.Remove Node.Child.Index
       Set Consql1 = vg_db.Execute("select distinct Rcpe_Cat_No, Rcpe_Cat_Desc " & _
                     "From PB00086 " & _
                     "where  Prev_Rcpe_Cat_No=" & Val(Mid(TvwDir(1).Nodes(dest.Index).Key, 2, 20)) & " " & _
                     "order by Rcpe_Cat_No", , adCmdText)
'       Set Consql1 = vg_db.Execute("sod_s_categoriareceta 14, " & Val(Mid(TvwDir(1).Nodes(dest.Index).Key, 2, 20)) & ", ''", , adCmdStoredProc)
       If Not Consql1.EOF Then
          Set rootnode = TvwDir(1).Nodes.Add(dest.Index, tvwChild, "HT" & Consql1!Rcpe_Cat_No, "Todos", 4)
          Do While Not Consql1.EOF
             Set nd = TvwDir(1).Nodes.Add(dest.Index, tvwChild, "H" & Consql1!Rcpe_Cat_No, Trim(Consql1!Rcpe_Cat_Desc), 4)
             dest.ExpandedImage = 5
             If nd.Children = 0 Then
                Set Consql2 = vg_db.Execute("select distinct Prev_Rcpe_Cat_No " & _
                              "From  PB00087 " & _
                              "where PB00087.Prev_Rcpe_Cat_No=" & Consql1!Rcpe_Cat_No & "", , adCmdText)
                If Not Consql2.EOF Then
                   ' la propiedad Texto de los nodos positivos es "***"
                   TvwDir(1).Nodes.Add nd.Index, tvwChild, , "**"
                End If
                Consql2.Close: Set Consql2 = Nothing
             End If
             Consql1.MoveNext
          Loop
       End If
       Consql1.Close: Set Consql1 = Nothing
    ElseIf Node.Child.Text = "**" Then
       ' eliminar el elemento hijo positivos ' ***
       TvwDir(1).Nodes.Remove Node.Child.Index
       Set Consql1 = vg_db.Execute("select distinct Rcpe_Cat_No, Rcpe_Cat_Desc " & _
                     "From PB00087 " & _
                     "where  Prev_Rcpe_Cat_No=" & Val(Mid(TvwDir(1).Nodes(dest.Index).Key, 2, 20)) & " " & _
                     "order by Rcpe_Cat_No", , adCmdText)
'       Set Consql1 = vg_db.Execute("sod_s_categoriareceta 15, " & Val(Mid(TvwDir(1).Nodes(dest.Index).Key, 2, 20)) & ", ''", , adCmdStoredProc)
       If Not Consql1.EOF Then
          Set rootnode = TvwDir(1).Nodes.Add(dest.Index, tvwChild, "JT" & Consql1!Rcpe_Cat_No, "Todos", 4)
          Do While Not Consql1.EOF
             Set nd = TvwDir(1).Nodes.Add(dest.Index, tvwChild, "J" & Consql1!Rcpe_Cat_No, Trim(Consql1!Rcpe_Cat_Desc), 4)
             dest.ExpandedImage = 5
             If nd.Children = 0 Then
                Set Consql2 = vg_db.Execute("select distinct Prev_Rcpe_Cat_No " & _
                             "From  PB00088 " & _
                             "where PB00088.Prev_Rcpe_Cat_No=" & Consql1!Rcpe_Cat_No & "", , adCmdText)
                If Not Consql2.EOF Then
                   ' la propiedad Texto de los nodos posotivos es "***"
                   TvwDir(1).Nodes.Add nd.Index, tvwChild, , "***"
                End If
                Consql2.Close: Set Consql2 = Nothing
             End If
             Consql1.MoveNext
          Loop
       End If
       Consql1.Close: Set Consql1 = Nothing
    ElseIf Node.Child.Text = "***" Then
       ' eliminar el elemento hijo positivos
       TvwDir(1).Nodes.Remove Node.Child.Index
       Set Consql1 = vg_db.Execute("select distinct Rcpe_Cat_No, Rcpe_Cat_Desc " & _
                     "From PB00088 " & _
                     "where Prev_Rcpe_Cat_No=" & Val(Mid(TvwDir(1).Nodes(dest.Index).Key, 2, 20)) & " " & _
                     "order by Rcpe_Cat_No", , adCmdText)
'       Set Consql1 = vg_db.Execute("sod_s_categoriareceta 16, " & Val(Mid(TvwDir(1).Nodes(dest.Index).Key, 2, 20)) & ", ''", , adCmdStoredProc)
       If Not Consql1.EOF Then
          Set rootnode = TvwDir(1).Nodes.Add(dest.Index, tvwChild, "KT" & Consql1!Rcpe_Cat_No, "Todos", 4)
          Do While Not Consql1.EOF
             Set nd = TvwDir(1).Nodes.Add(dest.Index, tvwChild, "K" & Consql1!Rcpe_Cat_No, Trim(Consql1!Rcpe_Cat_Desc), 4)
              dest.ExpandedImage = 5
'             If nd.Children = 0 And Consql1!hijo = "1" Then
'                 ' la propiedad Texto de los nodos positivos es "***"
'                TvwDir(1).Nodes.Add nd.Index, tvwChild, , "*****"
'             End If
             Consql1.MoveNext
          Loop
       End If
       Consql1.Close: Set Consql1 = Nothing
    End If
End Select
End Sub
Private Sub TvwDir_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 27
    vg_opcion = 2
    Cerrar
End Select
End Sub
Private Sub TvwDir_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
Set dest = Node
Select Case Index
   Case 0
     nivel = Mid(TvwDir(0).Nodes(dest.Index).Key, 1, 1)
     auxdiet = Val(Mid(TvwDir(0).Nodes(dest.Index).Key, 2, 20))
     nomdietetico = Trim((TvwDir(0).Nodes(dest.Index).Text))
   Case 1
     If Mid(TvwDir(1).Nodes(dest.Index).Key, 1, 1) = "R" Then
        If Mid(TvwDir(1).Nodes(dest.Index).Key, 1, 2) = "RT" Then
           auxcat1 = 0
        Else
           auxcat1 = Val(Mid(TvwDir(1).Nodes(dest.Index).Key, 2, 20))
        End If
        nomtplato1 = Trim((TvwDir(1).Nodes(dest.Index).Text))
        nomtplato2 = ""
        nomtplato3 = ""
        nomtplato4 = ""
        auxcat2 = 0
        auxcat3 = 0
        auxcat4 = 0
     ElseIf Mid(TvwDir(1).Nodes(dest.Index).Key, 1, 1) = "H" Then
        If Mid(TvwDir(1).Nodes(dest.Index).Key, 1, 1) = "HT" Then
           auxcat2 = 0
        Else
           auxcat2 = Val(Mid(TvwDir(1).Nodes(dest.Index).Key, 2, 20))
        End If
        nomtplato2 = Trim((TvwDir(1).Nodes(dest.Index).Text))
        nomtplato3 = ""
        nomtplato4 = ""
        auxcat3 = 0
        auxcat4 = 0
     ElseIf Mid(TvwDir(1).Nodes(dest.Index).Key, 1, 1) = "J" Then
        If Mid(TvwDir(1).Nodes(dest.Index).Key, 1, 1) = "JT" Then
           auxcat3 = 0
        Else
           auxcat3 = Val(Mid(TvwDir(1).Nodes(dest.Index).Key, 2, 20))
        End If
        nomtplato3 = Trim((TvwDir(1).Nodes(dest.Index).Text))
        nomtplato4 = ""
        auxcat4 = 0
     ElseIf Mid(TvwDir(1).Nodes(dest.Index).Key, 1, 1) = "K" Then
        If Mid(TvwDir(1).Nodes(dest.Index).Key, 1, 1) = "KT" Then
           auxcat4 = 0
        Else
           auxcat4 = Val(Mid(TvwDir(1).Nodes(dest.Index).Key, 2, 20))
        End If
        nomtplato4 = Trim((TvwDir(1).Nodes(dest.Index).Text))
     End If
End Select
End Sub
Sub Cerrar()
Me.Hide
End Sub
