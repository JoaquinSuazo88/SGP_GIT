VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form B_DieTip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtrar Recetas"
   ClientHeight    =   5460
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   9420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
            Picture         =   "B_DieTip.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_DieTip.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_DieTip.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_DieTip.frx":094E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_DieTip.frx":0AA8
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
Attribute VB_Name = "B_DieTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset, RS2 As New ADODB.Recordset
Dim dest As Node, nd As Node, rootNode As Node
Dim nomTippla As String, nomcatdie As String
Dim codcatdie As Long, codTippla As Long

Private Sub Form_Activate()

fg_descarga
codcatdie = 0: codTippla = 0
codcatdie = vg_filcatdie: codTippla = vg_filtippla

End Sub

Private Sub Form_Load()

fg_centra Me
nomTippla = "": nomcatdie = ""
MoverDatosTvwDir

End Sub

Sub MoverDatosTvwDir()

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset

fg_carga "ss"
codcatdie = 0: codTippla = 0
' *** Llenar Categoria dietetica ***'
TvwDir(0).Nodes.Clear
'RS1.Open "SELECT * FROM a_recetacatdie WHERE car_previo=0 ORDER BY car_codigo", vg_db, adOpenStatic
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS1 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaPrimerNivel_V02")
If Not RS1.EOF Then
   
   Do While Not RS1.EOF
      
      Set rootNode = TvwDir(0).Nodes.Add(, , "R" & RS1!car_codigo, Trim(RS1!car_nombre), 4)
          ' agregar un nodo hijo postizo, si fuera necesario
      
      If rootNode.Children = 0 Then
         
'         RS2.Open "SELECT DISTINCT car_previo FROM  a_recetacatdie WHERE car_previo=" & RS1!car_codigo & "", vg_db, adOpenForwardOnly ', adOpenStatic
         RS2.CursorLocation = adUseClient
         vg_db.CursorLocation = adUseClient
         Set RS2 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02 " & RS1!car_codigo & ", '1'")
         If Not RS2.EOF Then
            
            ' la propiedad Texto de los nodos postizos es "***"
            TvwDir(0).Nodes.Add rootNode.Index, tvwChild, , "*"
         
         End If
         RS2.Close: Set RS2 = Nothing
      
      End If
      
      RS1.MoveNext
   
   Loop

End If
RS1.Close: Set RS1 = Nothing

'*** Llenar Tipo Plato *** '
TvwDir(1).Nodes.Clear
'RS1.Open "SELECT * FROM a_recetatippla WHERE tip_previo=0 ORDER BY tip_nombre", vg_db, adOpenStatic
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS1 = vg_db.Execute("sgpadm_Sel_TipoPlatoPrimerNivel_V02")
If Not RS1.EOF Then
   
   Do While Not RS1.EOF
      
      Set rootNode = TvwDir(1).Nodes.Add(, , "R" & RS1!tip_codigo, Trim(RS1!tip_nombre), 4)
          ' agregar un nodo hijo postizo, si fuera necesario
      
      If rootNode.Children = 0 Then
         
'         RS2.Open "SELECT DISTINCT tip_previo FROM a_recetatippla WHERE tip_previo=" & RS1!tip_codigo & "", vg_db, adOpenForwardOnly ', adOpenStatic
         RS2.CursorLocation = adUseClient
         vg_db.CursorLocation = adUseClient
         Set RS2 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & RS1!tip_codigo & ", '1'")
         
         If Not RS2.EOF Then
            
            ' la propiedad Texto de los nodos postizos es "***"
            TvwDir(1).Nodes.Add rootNode.Index, tvwChild, , "*"
         
         End If
         
         RS2.Close: Set RS2 = Nothing
      
      End If
      
      RS1.MoveNext
   
   Loop

End If
RS1.Close: Set RS1 = Nothing

fg_descarga

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

Case 1
    
    vg_filcatdie = codcatdie: vg_filtippla = codTippla: vg_filnomcatdie = "Todos": vg_filnomtippla = "Todos"
    If codcatdie > 0 Then vg_filnomcatdie = TvwDir(0).SelectedItem.FullPath
    If codTippla > 0 Then vg_filnomtippla = TvwDir(1).SelectedItem.FullPath
    vg_opcion = 0
    Me.Hide

Case 3
    
    If TvwDir(0).Nodes.count < 1 And TvwDir(1).Nodes.count < 1 Then Exit Sub
    codcatdie = 0: codTippla = 0
    nomTippla = "Todos": nomcatdie = "Todos"
    MoverDatosTvwDir

Case 5
    
    vg_opcion = 2
    Me.Hide

End Select

End Sub

Private Sub tvwDir_Expand(Index As Integer, ByVal Node As MSComctlLib.Node)

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset

Set dest = Node
Select Case Index

Case 0
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" Then Exit Sub
    If Node.Child.text = "*" Then
 ' eliminar el elemento hijo positivo
       TvwDir(0).Nodes.Remove Node.Child.Index
'       RS1.Open "SELECT * FROM a_recetacatdie WHERE car_previo=" & Val(Mid(TvwDir(0).Nodes(dest.Index).Key, 2, 20)) & " ORDER BY car_codigo", vg_db, adOpenStatic
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       Set RS1 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02 " & Val(Mid(TvwDir(0).Nodes(dest.Index).Key, 2, 20)) & ", '2'")
       If Not RS1.EOF Then
          
          Do While Not RS1.EOF
             
             Set nd = TvwDir(0).Nodes.Add(dest.Index, tvwChild, "H" & RS1!car_codigo, Trim(RS1!car_nombre), 4)
             dest.ExpandedImage = 5
             
             If nd.Children = 0 Then
                
'                RS2.Open "SELECT DISTINCT car_previo FROM a_recetacatdie WHERE car_previo=" & RS1!car_codigo & "", vg_db, adOpenForwardOnly ', adOpenStatic
                RS2.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                Set RS2 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02 " & RS1!car_codigo & ", '1'")
                If Not RS2.EOF Then
                   
                   ' la propiedad Texto de los nodos positivos es "***"
                   TvwDir(0).Nodes.Add nd.Index, tvwChild, , "*"
                
                End If
                
                RS2.Close: Set RS2 = Nothing
             
             End If
             
             RS1.MoveNext
          Loop
       
       End If
       
       RS1.Close: Set RS1 = Nothing
    
    End If

Case 1
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" Then Exit Sub
    ' eliminar el elemento hijo positivo
    TvwDir(1).Nodes.Remove Node.Child.Index
'    RS1.Open "SELECT * FROM a_recetatippla WHERE tip_previo=" & Val(Mid(TvwDir(1).Nodes(dest.Index).Key, 2, 20)) & " ORDER BY tip_codigo", vg_db, adOpenStatic
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS1 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & Val(Mid(TvwDir(1).Nodes(dest.Index).Key, 2, 20)) & ", '2'")
    If Not RS1.EOF Then
       
       Do While Not RS1.EOF
          
          Set nd = TvwDir(1).Nodes.Add(dest.Index, tvwChild, "H" & RS1!tip_codigo, Trim(RS1!tip_nombre), 4)
          dest.ExpandedImage = 5
          
          If nd.Children = 0 Then
             
'             RS2.Open "SELECT DISTINCT tip_previo FROM a_recetatippla WHERE tip_previo=" & RS1!tip_codigo & "", vg_db, adOpenForwardOnly ', adOpenStatic
             RS2.CursorLocation = adUseClient
             vg_db.CursorLocation = adUseClient
             Set RS2 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & RS1!tip_codigo & ", '1'")
             If Not RS2.EOF Then
                
                ' la propiedad Texto de los nodos positivos es "***"
                TvwDir(1).Nodes.Add nd.Index, tvwChild, , "**"
             
             End If
             
             RS2.Close: Set RS2 = Nothing
          
          End If
          
          RS1.MoveNext
       
       Loop
    
    End If
    RS1.Close: Set RS1 = Nothing

End Select

End Sub

Private Sub TvwDir_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Select Case KeyCode

Case 27
    
    vg_opcion = 2
    Me.Hide

End Select

End Sub

Private Sub TvwDir_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)

Set dest = Node

Select Case Index

Case 0
     
     codcatdie = Val(Mid(TvwDir(0).Nodes(dest.Index).Key, 2, 20))
     nomcatdie = Trim((TvwDir(0).Nodes(dest.Index).text))

Case 1
     
     codTippla = Val(Mid(TvwDir(1).Nodes(dest.Index).Key, 2, 20))
     nomTippla = Trim((TvwDir(1).Nodes(dest.Index).text))

End Select

End Sub
