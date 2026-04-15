VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form B_DieTipExcel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtrar Recetas Excel"
   ClientHeight    =   6945
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   10890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Tipo Plato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Index           =   1
      Left            =   5400
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      Begin MSComctlLib.TreeView TvwDir 
         Height          =   6135
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   10821
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
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
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Categoria Dietetica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin MSComctlLib.TreeView TvwDir 
         Height          =   6135
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   10821
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
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
            Picture         =   "B_DieTipExcel.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_DieTipExcel.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_DieTipExcel.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_DieTipExcel.frx":094E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_DieTipExcel.frx":0AA8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6945
      Left            =   10260
      TabIndex        =   4
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   12250
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
Attribute VB_Name = "B_DieTipExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dest As Node, nd As Node, nd2 As Node, rootNode As Node, nd1 As Node
Dim MsgTitulo As String

Dim Nivel2 As Long
Dim Nivel3 As Long
Dim Nivel4 As Long
Dim Nivel5 As Long
Dim Nivel6 As Long

Dim Nodx   As Node
Dim Nod2   As Node
Dim Nod3   As Node
Dim Nod4   As Node
Dim Nod5   As Node
Dim Nod6   As Node

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me
MsgTitulo = "Filtro Cátegoria Diétetica & Tipo de Plato"

'MoverDatosTvwDir

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub MoverDatosTvwDir()

On Error GoTo Man_Error

'tvwFirst : Añade el nodo al principio
'tvwLast : Añade el nodo al final
'tvwNext : Lo añade al siguiente nodo indicado
'tvwPrevious : Lo añade al lugar anterior al nodo indicado
'tvwChild : Nuevo nodo Hijo o secundario del nodo indicado

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset

fg_carga ""
CodCatDie = 0
codTippla = 0
Nivel2 = 0

' *** Llenar Categoria dietetica ***'
TvwDir(0).Nodes.Clear

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS1 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaPrimerNivel_V02")
If Not RS1.EOF Then
   
   Do While Not RS1.EOF
           
      Set Nodx = TvwDir(0).Nodes.Add(, , "R" & RS1!car_codigo, RS1!car_codigo & " - " & Trim(RS1!car_nombre))
      ' agregar un nodo hijo postizo, si fuera necesario

      If Nodx.Children = 0 Then
         
         If RS2.State = 1 Then RS2.Close
         RS2.CursorLocation = adUseClient
         vg_db.CursorLocation = adUseClient
         
         Set RS2 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02 " & RS1!car_codigo & ", '1'")
         
         If Not RS2.EOF Then
            
            ' la propiedad Texto de los nodos postizos es "***"
            TvwDir(0).Nodes.item(TvwDir(0).Nodes.count).Selected = True
            TvwDir(0).Nodes.Add Nodx.Index, tvwChild, , "*"
            
            Set nd1 = TvwDir(0).SelectedItem
            tvwDir_Expand_2 0, nd1
         
         End If
         RS2.Close
         Set RS2 = Nothing
      
      End If
      
      RS1.MoveNext
   
   Loop

End If
RS1.Close
Set RS1 = Nothing

'*** Llenar Tipo Plato *** '
Nivel2 = 0
TvwDir(1).Nodes.Clear

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_Sel_TipoPlatoPrimerNivel_V02")

If Not RS1.EOF Then

   Do While Not RS1.EOF

      Set Nodx = TvwDir(1).Nodes.Add(, , "R" & RS1!tip_codigo, RS1!tip_codigo & " - " & Trim(RS1!tip_nombre))

      ' agregar un nodo hijo postizo, si fuera necesario
      If Nodx.Children = 0 Then

         If RS2.State = 1 Then RS2.Close
         RS2.CursorLocation = adUseClient
         vg_db.CursorLocation = adUseClient
         
         Set RS2 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & RS1!tip_codigo & ", '1'")

         If Not RS2.EOF Then

            ' la propiedad Texto de los nodos postizos es "***"
            TvwDir(1).Nodes.item(TvwDir(1).Nodes.count).Selected = True
            TvwDir(1).Nodes.Add Nodx.Index, tvwChild, , "*"

            Set nd1 = TvwDir(1).SelectedItem
            tvwDir_Expand_2 1, nd1

         End If

         RS2.Close
         Set RS2 = Nothing

      End If

      RS1.MoveNext

   Loop

End If
RS1.Close
Set RS1 = Nothing

For Iini = 1 To TvwDir(0).Nodes.count

    TvwDir(0).Nodes.item(Iini).Checked = True

Next

For Iini = 1 To TvwDir(1).Nodes.count

    TvwDir(1).Nodes.item(Iini).Checked = True

Next

fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index

Case 5
    
    Me.Hide

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub tvwDir_Expand_2(Index As Integer, ByVal Node As MSComctlLib.Node)

On Error GoTo Man_Error

Dim RS1       As New ADODB.Recordset
Dim RS2       As New ADODB.Recordset
Dim estnivel2 As Boolean

estnivel2 = True
Set dest = Node
Nivel3 = 0

Select Case Index

Case 0
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" Then Exit Sub
    
    ' eliminar el elemento hijo positivo
       
       TvwDir(0).Nodes.Remove Node.Child.Index
       
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Nivel3 = 0
       Nivel2 = Val(Mid(TvwDir(0).Nodes(dest.Index).key, 2, 10))
       
       Set RS1 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02 " & Nivel2 & ", '2'")
       
       If Not RS1.EOF Then
          
          Do While Not RS1.EOF
                                      
             Set Nod2 = TvwDir(0).Nodes.Add(Nodx, tvwChild, "H" & RS1!car_codigo & fg_pone_espacio(Nivel2, 10), RS1!car_codigo & " - " & Trim(RS1!car_nombre))
             
             If Nod2.Children = 0 Then
                
                If RS2.State = 1 Then RS2.Close
                RS2.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS2 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02 " & RS1!car_codigo & ", '1'")
                
                If Not RS2.EOF Then
                   
                Nivel3 = RS1!car_codigo
   
                ' la propiedad Texto de los nodos positivos es "***"
                TvwDir(0).Nodes.item(TvwDir(0).Nodes.count).Selected = True
                TvwDir(0).Nodes.Add Nod2.Index, tvwChild, , "**"
                
                Set nd = TvwDir(0).SelectedItem
                   
                Set nd1 = TvwDir(0).SelectedItem
                tvwDir_Expand_3 0, nd1
                estnivel2 = False
                
                End If
                
                RS2.Close
                Set RS2 = Nothing
                
             End If
                       
             RS1.MoveNext
          
          Loop
       
       End If
       
       RS1.Close
       Set RS1 = Nothing

Case 1
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" Then Exit Sub
    
    ' eliminar el elemento hijo positivo
    TvwDir(1).Nodes.Remove Node.Child.Index
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Nivel3 = 0
    Nivel2 = Val(Mid(TvwDir(1).Nodes(dest.Index).key, 2, 10))
    
    Set RS1 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & Nivel2 & ", '2'")
    
    If Not RS1.EOF Then
       
       Do While Not RS1.EOF
          
          Set Nod2 = TvwDir(1).Nodes.Add(Nodx, tvwChild, "H" & RS1!tip_codigo & fg_pone_espacio(Nivel2, 10), RS1!tip_codigo & " - " & Trim(RS1!tip_nombre))
          
          If Nod2.Children = 0 Then
             
             If RS2.State = 1 Then RS2.Close
             RS2.CursorLocation = adUseClient
             vg_db.CursorLocation = adUseClient
             
             Set RS2 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & RS1!tip_codigo & ", '1'")
             
             If Not RS2.EOF Then
                
                Nivel3 = RS1!tip_codigo
            
                ' la propiedad Texto de los nodos positivos es "***"
                TvwDir(1).Nodes.item(TvwDir(1).Nodes.count).Selected = True
                TvwDir(1).Nodes.Add Nod2.Index, tvwChild, , "**"

                Set nd1 = TvwDir(1).SelectedItem
                tvwDir_Expand_3 1, nd1
                estnivel2 = False

             End If
             
             RS2.Close
             Set RS2 = Nothing
          
            
          End If
          
          RS1.MoveNext
       
       Loop
    
    End If
    RS1.Close
    Set RS1 = Nothing

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub tvwDir_Expand_3(Index As Integer, ByVal Node As MSComctlLib.Node)

On Error GoTo Man_Error

Dim RS5       As New ADODB.Recordset
Dim RS6       As New ADODB.Recordset
Dim estnivel3 As Boolean

estnivel3 = True
Nivel4 = 0
Set dest = Node

Select Case Index

Case 0
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" Then Exit Sub

    ' eliminar el elemento hijo positivo
       TvwDir(0).Nodes.Remove Node.Child.Index
       
       If RS5.State = 1 Then RS5.Close
       RS5.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS5 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02  " & Nivel3 & ", '1'")
       
       If Not RS5.EOF Then
          
          Do While Not RS5.EOF
             
             Set Nod3 = TvwDir(0).Nodes.Add(Nod2, tvwChild, "H" & RS5!car_codigo & fg_pone_espacio(Nivel3, 10), RS5!car_codigo & " - " & Trim(RS5!car_nombre))
             
             If Nod3.Children = 0 Then
                
                If RS6.State = 1 Then RS6.Close
                RS6.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS6 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02 " & RS5!car_codigo & ", '1'")
                
                If Not RS6.EOF Then
                   
                   Nivel4 = RS5!car_codigo
                
                   ' la propiedad Texto de los nodos positivos es "***"
                   TvwDir(0).Nodes.item(TvwDir(0).Nodes.count).Selected = True
                   TvwDir(0).Nodes.Add Nod3.Index, tvwChild, , "****"
                
                   Set nd = TvwDir(0).SelectedItem
                
                   Set ndl = TvwDir(0).SelectedItem
                   tvwDir_Expand_4 0, ndl 'dest
                   estnivel3 = False
                
                End If
                
                RS6.Close
                Set RS6 = Nothing
                
             End If
             
             RS5.MoveNext
          
          Loop
       
       RS5.Close
       Set RS5 = Nothing
    
    End If

Case 1
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" Then Exit Sub
    
    ' eliminar el elemento hijo positivo
    TvwDir(1).Nodes.Remove Node.Child.Index
    
    If RS5.State = 1 Then RS5.Close
    RS5.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS5 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & Nivel3 & ", '2'")
    
    If Not RS5.EOF Then
       
       Do While Not RS5.EOF
          
          Set Nod3 = TvwDir(1).Nodes.Add(Nod2, tvwChild, "H" & RS5!tip_codigo & fg_pone_espacio(Val(Nivel3), 10), RS5!tip_codigo & " - " & Trim(RS5!tip_nombre))
          
          If Nod3.Children = 0 Then
             
             If RS6.State = 1 Then RS6.Close
             RS6.CursorLocation = adUseClient
             vg_db.CursorLocation = adUseClient
             
             Set RS6 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & RS5!tip_codigo & ", '1'")
             
             If Not RS6.EOF Then
                
                Nivel4 = RS5!tip_codigo
                
                ' la propiedad Texto de los nodos positivos es "***"
                TvwDir(1).Nodes.item(TvwDir(1).Nodes.count).Selected = True
                TvwDir(1).Nodes.Add Nod3.Index, tvwChild, , "****"

                Set ndl = TvwDir(1).SelectedItem
                tvwDir_Expand_4 1, ndl 'dest
                estnivel3 = False
             
             End If
             
             RS6.Close
             Set RS6 = Nothing
          
          End If
          
          RS5.MoveNext
       
       Loop
    
    End If
    RS5.Close
    Set RS5 = Nothing

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub tvwDir_Expand_4(Index As Integer, ByVal Node As MSComctlLib.Node)

On Error GoTo Man_Error

Dim RS7       As New ADODB.Recordset
Dim RS8       As New ADODB.Recordset
Dim estnivel4 As Boolean

Set dest = Node
Select Case Index

Case 0
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" And Node.Child.text <> "****" Then Exit Sub

    ' eliminar el elemento hijo positivo
       TvwDir(0).Nodes.Remove Node.Child.Index
       
       If RS7.State = 1 Then RS7.Close
       RS7.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS7 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02 " & Nivel4 & ", '1'")
       
       If Not RS7.EOF Then
          
          Do While Not RS7.EOF
             
              Set Nod4 = TvwDir(0).Nodes.Add(Nod3, tvwChild, "H" & RS7!car_codigo & fg_pone_espacio(Val(Nivel4), 10), RS7!car_codigo & " - " & Trim(RS7!car_nombre))
             
             If Nod4.Children = 0 Then
                
                If RS8.State = 1 Then RS8.Close
                RS8.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS8 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02 " & RS7!car_codigo & ", '1'")
                
                If Not RS8.EOF Then
                   
                   Nivel5 = RS7!car_codigo
                   
                   ' la propiedad Texto de los nodos positivos es "*****"
                   TvwDir(0).Nodes.item(TvwDir(0).Nodes.count).Selected = True
                   TvwDir(0).Nodes.Add Nod4.Index, tvwChild, , "*****"
                
                   Set ndl = TvwDir(0).SelectedItem
                   tvwDir_Expand_5 0, ndl 'dest
                   estnivel4 = False
                
                End If
                
                RS8.Close
                Set RS8 = Nothing
                
             End If
             
             RS7.MoveNext
          
          Loop
       
       End If
       
       RS7.Close
       Set RS7 = Nothing
    
Case 1
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" And Node.Child.text <> "****" Then Exit Sub
    
    ' eliminar el elemento hijo positivo
    TvwDir(1).Nodes.Remove Node.Child.Index
    
    If RS7.State = 1 Then RS7.Close
    RS7.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS7 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & Nivel4 & ", '2'")
    
    If Not RS7.EOF Then
       
       Do While Not RS7.EOF
          
          Set Nod4 = TvwDir(1).Nodes.Add(Nod3, tvwChild, "H" & RS7!tip_codigo & fg_pone_espacio(Val(Nivel4), 10), RS7!tip_codigo & " - " & Trim(RS7!tip_nombre))
          
          If Nod4.Children = 0 Then
             
             If RS8.State = 1 Then RS8.Close
             RS8.CursorLocation = adUseClient
             vg_db.CursorLocation = adUseClient
             
             Set RS8 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & RS7!tip_codigo & ", '1'")
             
             If Not RS8.EOF Then
                
                Nivel5 = RS7!tip_codigo
                
                ' la propiedad Texto de los nodos positivos es "*****"
                TvwDir(1).Nodes.item(TvwDir(1).Nodes.count).Selected = True
                TvwDir(1).Nodes.Add Nod4.Index, tvwChild, , "*****"

                Set ndl = TvwDir(1).SelectedItem
                tvwDir_Expand_5 1, ndl 'dest
                estnivel4 = False
             
             End If
             
             RS8.Close
             Set RS8 = Nothing
          
          End If
          
          RS7.MoveNext
       
       Loop
    
    End If
    RS7.Close
    Set RS7 = Nothing

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub tvwDir_Expand_5(Index As Integer, ByVal Node As MSComctlLib.Node)

On Error GoTo Man_Error

Dim RS9        As New ADODB.Recordset
Dim RS10       As New ADODB.Recordset
Dim estnivel5  As Boolean

Set dest = Node
Select Case Index

Case 0
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" And Node.Child.text <> "****" And Node.Child.text <> "*****" Then Exit Sub

    ' eliminar el elemento hijo positivo
       TvwDir(0).Nodes.Remove Node.Child.Index
       
       If RS9.State = 1 Then RS9.Close
       RS9.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS9 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02 " & Nivel5 & ", '1'")
       
       If Not RS9.EOF Then
          
          Do While Not RS9.EOF
             
              Set Nod5 = TvwDir(0).Nodes.Add(Nod4, tvwChild, "H" & RS9!car_codigo & fg_pone_espacio(Val(Nivel5), 10), RS9!car_codigo & " - " & Trim(RS9!car_nombre))
             
             If Nod5.Children = 0 Then
                
                If RS10.State = 1 Then RS10.Close
                RS10.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS10 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02 " & RS9!car_codigo & ", '1'")
                
                If Not RS10.EOF Then
                   
                   Nivel6 = RS9!car_codigo
                   
                   ' la propiedad Texto de los nodos positivos es "*****"
                   TvwDir(0).Nodes.item(TvwDir(0).Nodes.count).Selected = True
                   TvwDir(0).Nodes.Add Nod5.Index, tvwChild, , "******"
                
                   Set ndl = TvwDir(0).SelectedItem
                   tvwDir_Expand_6 0, ndl 'dest
                   estnivel5 = False
                
                End If
                
                RS10.Close
                Set RS10 = Nothing
                
             End If
             
             RS9.MoveNext
          
          Loop
       
       End If
       
       RS9.Close
       Set RS9 = Nothing
    
Case 1
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" And Node.Child.text <> "****" And Node.Child.text <> "*****" Then Exit Sub
    
    ' eliminar el elemento hijo positivo
    TvwDir(1).Nodes.Remove Node.Child.Index
    
    If RS9.State = 1 Then RS9.Close
    RS9.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS9 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & Nivel5 & ", '2'")
    
    If Not RS9.EOF Then
       
       Do While Not RS9.EOF
          
          Set Nod5 = TvwDir(1).Nodes.Add(Nod4, tvwChild, "H" & RS9!tip_codigo & fg_pone_espacio(Val(Nivel5), 10), RS9!tip_codigo & " - " & Trim(RS9!tip_nombre))
          
          If Nod5.Children = 0 Then
             
             If RS10.State = 1 Then RS10.Close
             RS10.CursorLocation = adUseClient
             vg_db.CursorLocation = adUseClient
             
             Set RS10 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & RS9!tip_codigo & ", '1'")
             
             If Not RS10.EOF Then
                
                Nivel6 = RS9!tip_codigo
                
                ' la propiedad Texto de los nodos positivos es "*****"
                TvwDir(1).Nodes.item(TvwDir(1).Nodes.count).Selected = True
                TvwDir(1).Nodes.Add Nod5.Index, tvwChild, , "******"

                Set ndl = TvwDir(1).SelectedItem
                tvwDir_Expand_6 1, ndl 'dest
                estnivel5 = False
             
             End If
             
             RS10.Close
             Set RS10 = Nothing
          
          End If
          
          RS9.MoveNext
       
       Loop
    
    End If
    RS9.Close
    Set RS9 = Nothing

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub tvwDir_Expand_6(Index As Integer, ByVal Node As MSComctlLib.Node)

On Error GoTo Man_Error

Dim RS11       As New ADODB.Recordset
Dim RS12       As New ADODB.Recordset
Dim estnivel6 As Boolean

Set dest = Node
Select Case Index

Case 0
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" And Node.Child.text <> "****" And Node.Child.text <> "*****" And Node.Child.text <> "******" Then Exit Sub

    ' eliminar el elemento hijo positivo
       TvwDir(0).Nodes.Remove Node.Child.Index
       
       If RS11.State = 1 Then RS11.Close
       RS11.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS11 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02 " & Nivel6 & ", '1'")
       
       If Not RS11.EOF Then
          
          Do While Not RS11.EOF
             
              Set Nod6 = TvwDir(0).Nodes.Add(Nod5, tvwChild, "H" & RS11!car_codigo & fg_pone_espacio(Val(Nivel6), 10), RS11!car_codigo & " - " & Trim(RS11!car_nombre))
             
             If Nod6.Children = 0 Then
                
                If RS12.State = 1 Then RS12.Close
                RS12.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS12 = vg_db.Execute("sgpadm_Sel_CategoriaDieteticaOtrosNiveles_V02 " & RS11!car_codigo & ", '1'")
                
                If Not RS12.EOF Then
                   
                   ' la propiedad Texto de los nodos positivos es "******"
                   TvwDir(0).Nodes.item(TvwDir(0).Nodes.count).Selected = True
                   TvwDir(0).Nodes.Add Nod6.Index, tvwChild, , "******"
                
                   Set nd = TvwDir(0).SelectedItem
                   tvwDir_Expand_6 0, dest
                   estnivel6 = False
                
                End If
                
                RS12.Close
                Set RS12 = Nothing
                
             End If
             
             RS11.MoveNext
          
          Loop
       
       End If
       
       RS11.Close
       Set RS11 = Nothing
    
Case 1
    
    If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
    If Node.Child.text <> "*" And Node.Child.text <> "**" And Node.Child.text <> "***" And Node.Child.text <> "****" And Node.Child.text <> "*****" And Node.Child.text <> "******" Then Exit Sub
    
    ' eliminar el elemento hijo positivo
    TvwDir(1).Nodes.Remove Node.Child.Index
    
    If RS11.State = 1 Then RS11.Close
    RS11.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS11 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & Nivel6 & ", '2'")
    
    If Not RS11.EOF Then
       
       Do While Not RS11.EOF
          
          Set Nod6 = TvwDir(1).Nodes.Add(Nod5, tvwChild, "H" & RS11!tip_codigo & fg_pone_espacio(Val(Nivel6), 10), RS11!tip_codigo & " - " & Trim(RS11!tip_nombre))
          
          If Nod6.Children = 0 Then
             
             If RS12.State = 1 Then RS12.Close
             RS12.CursorLocation = adUseClient
             vg_db.CursorLocation = adUseClient
             
             Set RS12 = vg_db.Execute("sgpadm_Sel_TipoPlatoOtrosNiveles_V02 " & RS11!tip_codigo & ", '1'")
             
             If Not RS12.EOF Then
                
                ' la propiedad Texto de los nodos positivos es "*****"
                TvwDir(1).Nodes.item(TvwDir(1).Nodes.count).Selected = True
                TvwDir(1).Nodes.Add Nod6.Index, tvwChild, , "******"

                Set nd = TvwDir(1).SelectedItem
                tvwDir_Expand_6 1, dest
                estnivel6 = False
             
             End If
             
             RS12.Close
             Set RS12 = Nothing
          
          End If
          
          RS11.MoveNext
       
       Loop
    
    End If
    RS11.Close
    Set RS11 = Nothing

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub TvwDir_NodeCheck(Index As Integer, ByVal Node As MSComctlLib.Node)

'Dim lCheck  As Boolean
'Dim lCheck1 As Boolean
'Dim lCheck2 As Boolean
'Dim itesel  As Node
'Dim i       As Long
'Dim j       As Long
'Dim p       As Long
'Dim cKey    As String
'Dim cKey2   As String
'
'fg_carga ""
'
'TvwDir(Index).Nodes.item(Node.Key).Selected = True
'Set itesel = TvwDir(Index).SelectedItem
''tvwDir_Expand itesel
'TvwDir(Index).Nodes.item(Node.Key).Selected = True
'lCheck = TvwDir(Index).Nodes.item(TvwDir(Index).SelectedItem.Index).Checked
'lCheck1 = TvwDir(Index).Nodes.item(TvwDir(Index).SelectedItem.Index).Checked
'cKey = Trim(TvwDir(Index).Nodes.item(TvwDir(Index).SelectedItem.Index).Key)
''MsgBox Val(Mid(cKey, 2, 10))
'If TvwDir(Index).SelectedItem.Children > 0 Then
'
'   For i = TvwDir(Index).SelectedItem.Index + 1 To TvwDir(Index).Nodes.count
'
'       If (Val(Mid(cKey, 2, 11)) = Val(Mid(TvwDir(Index).Nodes.item(i).Key, 12, 21))) Or (Val(Mid(cKey2, 2, 11)) = Val(Mid(TvwDir(Index).Nodes.item(i).Key, 12, 21)) And Val(Mid(cKey2, 2, 11)) > 0) Then
'
'          TvwDir(Index).Nodes.item(i).Checked = lCheck1
'
'          If Index = 1 Then
'
'             If TvwDir(Index).Nodes.item(i).Children > 0 Then
'
'                cKey2 = Trim(TvwDir(Index).Nodes.item(i).Key)
'
'             End If
'
'          End If
'
'       ElseIf Val(Mid(cKey, 2, 10)) = Val(Mid(TvwDir(Index).Nodes.item(i).Key, 11, 22)) And TvwDir(Index).Nodes.item(i).Children = 0 Then
'
'          TvwDir(Index).Nodes.item(i).Checked = lCheck1
'
'          'TvwDir(Index).Nodes.item(i).Children
'
'       End If
'
'   Next i
'
'Else
'
'   For i = 1 To TvwDir(Index).Nodes.count
'
'       If TvwDir(Index).Nodes.item(i).Children = 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDir(Index).Nodes.item(i).Key, 12, 21)) Then
'
'          j = i
'          Exit For
'
'       ElseIf Val(Mid(cKey, 2, 10)) = Val(Mid(TvwDir(Index).Nodes.item(i).Key, 11, 22)) And TvwDir(Index).Nodes.item(i).Children = 0 Then
'
'          j = i
'          Exit For
'
'       End If
'
'   Next i
'
'   For i = j To TvwDir(Index).Nodes.count
'
'       If TvwDir(Index).Nodes.item(i).Children > 0 Then
'
'            Exit For
'
'       End If
'
'       If TvwDir(Index).Nodes.item(i).Checked = True Then
'
'           For p = i - 1 To TvwDir(Index).Nodes.count
'
'               If Val(Mid(cKey, 11, 21)) = Val(Mid(TvwDir(Index).Nodes.item(p).Key, 12, 21)) Then
'
'                  lCheck1 = TvwDir(Index).Nodes.item(p).Checked
'
'               End If
'
'            Next p
'
''          lCheck1 = True 'TvwDir.Nodes.Item(i).Checked
'
'          Exit For
'
'       End If
'
'   Next i
'
'   For i = (TvwDir(Index).SelectedItem.Index - 1) To 1 Step -1
'
'       If TvwDir(Index).Nodes.item(i).Children > 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDir(Index).Nodes.item(i).Key, 2, 11)) Then
'
'          TvwDir(Index).Nodes.item(i).Checked = lCheck1
'          If Index = 0 Then
'
'             Exit For
'
'          End If
'
'       ElseIf Val(Mid(cKey, 11, 22)) = Val(Mid(TvwDir(Index).Nodes.item(i).Key, 2, 10)) And TvwDir(Index).Nodes.item(i).Children > 0 Then
'
'          TvwDir(Index).Nodes.item(i).Checked = lCheck
'          Exit For
'
'       ElseIf CStr(Mid(TvwDir(Index).Nodes.item(i).Key, 1, 1)) = "R" Then
'
''          lCheck2 = TvwDir(Index).Nodes.item(i).Checked
''
''          For p = i + 1 To TvwDir(Index).Nodes.count
''
''              If TvwDir(Index).Nodes.item(i).Checked = lCheck2 Then
''
''                 lCheck = TvwDir(Index).Nodes.item(i).Checked
''                 Exit For
''
''              End If
''
''          Next p
'
'          TvwDir(Index).Nodes.item(i).Checked = lCheck
'          Exit For
'
'       ElseIf TvwDir(Index).Nodes.item(i).Checked = True And TvwDir(Index).Nodes.item(i).Children > 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDir(Index).Nodes.item(i).Key, 2, 11)) Then
'
'          Exit For
'
'       ElseIf (TvwDir(Index).Nodes.item(i).Checked = True Or TvwDir(Index).Nodes.item(i).Children > 0) Then
'
''          Exit For
'
'       End If
'
'   Next i
'
'End If
'fg_descarga
'
'Exit Sub
'Man_Error:
'MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"


On Error GoTo Man_Error

Dim lCheck        As Boolean
Dim lCheck1       As Boolean
Dim itesel        As Node
Dim i             As Long
Dim j             As Long
Dim p             As Long
Dim cKey          As String
Dim cKey2         As String
Dim cKey3         As String
Dim cKeyFullPath  As String

fg_carga ""

cKey2 = ""
cKey3 = ""
cKey4 = ""
cKey5 = ""
cKey = ""
cKeyFullPath = ""

TvwDir(Index).Nodes.item(Node.key).Selected = True
Set itesel = TvwDir(Index).SelectedItem
TvwDir(Index).Nodes.item(Node.key).Selected = True
lCheck = TvwDir(Index).Nodes.item(TvwDir(Index).SelectedItem.Index).Checked
lCheck1 = TvwDir(Index).Nodes.item(TvwDir(Index).SelectedItem.Index).Checked
cKey = Trim(TvwDir(Index).Nodes.item(TvwDir(Index).SelectedItem.Index).key)
cKeyFullPath = TvwDir(Index).Nodes.item(TvwDir(Index).SelectedItem.Index).text '.Parent '.Root '.LastSibling '.Children '.FirstSibling '.FullPath

If TvwDir(Index).SelectedItem.Children > 0 Then
   
   For i = TvwDir(Index).SelectedItem.Index + 1 To TvwDir(Index).Nodes.count
      
       If Mid(TvwDir(Index).Nodes.item(i).key, 1, 1) = "R" Then
       
          Exit For
       
       End If
       
       If (Val(Mid(cKey, 2, 11)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 12, 21))) Or _
          (ValidarKey(cKey3, Mid(TvwDir(Index).Nodes.item(i).key, 12, 21), ",") _
          And Val(Mid(cKey2, 2, 11)) > 0) Then


'       (Val(Mid(cKey2, 2, 10)) = Val(Mid(TvwDir(Index).Nodes.item(i).Key, 12, 21)) Or _
'       And Val(Mid(cKey2, 2, 11)) > 0) Then
       
          TvwDir(Index).Nodes.item(i).Checked = lCheck1
                   
             If TvwDir(Index).Nodes.item(i).Children > 0 Then
               
                cKey2 = Trim(TvwDir(Index).Nodes.item(i).key)
                cKey3 = Trim(Mid(Trim(TvwDir(Index).Nodes.item(i).key), 2, 10)) & "," & cKey3

                
             End If
       
       ElseIf Val(Mid(cKey, 2, 10)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 11, 22)) And TvwDir(Index).Nodes.item(i).Children = 0 Then
       
          TvwDir(Index).Nodes.item(i).Checked = lCheck1
          
       ElseIf ValidarKey(TvwDir(Index).Nodes.item(i).FullPath, cKeyFullPath, "\") Then
       
          TvwDir(Index).Nodes.item(i).Checked = lCheck1
       
       End If
       
   Next i


   For i = 1 To TvwDir(Index).Nodes.count
       
       If TvwDir(Index).Nodes.item(i).Children = 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 12, 21)) Then
          
          j = i
          Exit For
       
       ElseIf Val(Mid(cKey, 2, 10)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 11, 22)) And TvwDir(Index).Nodes.item(i).Children = 0 Then
          
          j = i
          Exit For
       
       End If
   
   Next i

   If j > 0 Then

      For i = j To TvwDir(Index).Nodes.count
       
          If TvwDir(Index).Nodes.item(i).Checked = True And Val(Mid(cKey, 11, 10)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 11, 10)) Then
          
             lCheck1 = True
             Exit For
   
          End If
     
          If TvwDir(Index).Nodes.item(i).Children > 0 Then
       
             Exit For
        
          End If
       
      Next i
      
    End If
   Dim lCheck2 As Boolean
   lCheck2 = False
   
   For i = (TvwDir(Index).SelectedItem.Index - 1) To 1 Step -1

       cKey2 = Trim(TvwDir(Index).Nodes.item(i).key)
       If TvwDir(Index).Nodes.item(i).Children > 0 And Val(Mid(cKey, 22, 30)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 2, 11)) Then
          
          TvwDir(Index).Nodes.item(i).Checked = lCheck1
''          Exit For
       
       ElseIf Val(Mid(cKey, 11, 22)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 2, 10)) And TvwDir(Index).Nodes.item(i).Children > 0 Then
        
          TvwDir(Index).Nodes.item(i).Checked = lCheck1
'          Exit For
          
'       ElseIf Index = 1 And (Val(Mid(cKey2, 2, 11)) = Val(Mid(TvwDir(Index).Nodes.item(i).Key, 2, 10))) And CStr(Mid(TvwDir(Index).Nodes.item(i).Key, 1, 1)) = "R" Then
       ElseIf (Val(Mid(cKey2, 2, 11)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 2, 10))) And CStr(Mid(TvwDir(Index).Nodes.item(i).key, 1, 1)) = "R" Then
          
          For p = i + 1 To TvwDir(Index).Nodes.count

              If (Val(Mid(cKey, 2, 11)) = Val(Mid(TvwDir(Index).Nodes.item(p).key, 12, 21))) Or (Val(Mid(cKey2, 2, 11)) = Val(Mid(TvwDir(Index).Nodes.item(p).key, 12, 21)) And Val(Mid(cKey2, 2, 11)) > 0) Then
       
                 If TvwDir(Index).Nodes.item(p).Checked <> lCheck1 Then
                     
                        lCheck2 = TvwDir(Index).Nodes.item(p).Checked
                        
                 End If
                     
                 If TvwDir(Index).Nodes.item(p).Children > 0 Then

                    cKey2 = Trim(TvwDir(Index).Nodes.item(p).key)

                 End If

              End If
             
          Next p
             
          If (Val(Mid(cKey, 2, 11)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 2, 11))) Then
                    
             TvwDir(Index).Nodes.item(i).Checked = IIf(Not lCheck2, lCheck1, lCheck2)
          
          End If
          
          Exit For
          
       ElseIf TvwDir(Index).Nodes.item(i).Checked = True And TvwDir(Index).Nodes.item(i).Children > 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 2, 11)) Then
          
          Exit For
       
       ElseIf (TvwDir(Index).Nodes.item(i).Checked = True Or TvwDir(Index).Nodes.item(i).Children > 0) Then
          
       
       End If
   
   Next i

Else
   
   For i = 1 To TvwDir(Index).Nodes.count
       
       If TvwDir(Index).Nodes.item(i).Children = 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 12, 21)) Then
          
          j = i
          Exit For
       
       ElseIf Val(Mid(cKey, 2, 10)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 11, 22)) And TvwDir(Index).Nodes.item(i).Children = 0 Then
          
          j = i
          Exit For
       
       End If
   
   Next i

   If j > 0 Then

      For i = j To TvwDir(Index).Nodes.count
       
          If TvwDir(Index).Nodes.item(i).Checked = True And Val(Mid(cKey, 11, 10)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 11, 10)) Then
          
             lCheck1 = True
             Exit For
   
          End If
     
          If TvwDir(Index).Nodes.item(i).Children > 0 Then
       
             Exit For
        
          End If
       
      Next i
   
   End If
   
'   Dim lCheck2 As Boolean
   lCheck2 = False
   
   For i = (TvwDir(Index).SelectedItem.Index - 1) To 1 Step -1

       cKey2 = Trim(TvwDir(Index).Nodes.item(i).key)
       If TvwDir(Index).Nodes.item(i).Children > 0 And Val(Mid(cKey, 22, 30)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 2, 11)) Then
          
          TvwDir(Index).Nodes.item(i).Checked = lCheck1
''          Exit For
       
       ElseIf Val(Mid(cKey, 11, 22)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 2, 10)) And TvwDir(Index).Nodes.item(i).Children > 0 Then
        
          TvwDir(Index).Nodes.item(i).Checked = lCheck1
'          Exit For
          
'       ElseIf Index = 1 And (Val(Mid(cKey2, 2, 11)) = Val(Mid(TvwDir(Index).Nodes.item(i).Key, 2, 10))) And CStr(Mid(TvwDir(Index).Nodes.item(i).Key, 1, 1)) = "R" Then
       ElseIf (Val(Mid(cKey2, 2, 11)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 2, 10))) And CStr(Mid(TvwDir(Index).Nodes.item(i).key, 1, 1)) = "R" Then
          
          For p = i + 1 To TvwDir(Index).Nodes.count

              If (Val(Mid(cKey, 2, 11)) = Val(Mid(TvwDir(Index).Nodes.item(p).key, 12, 21))) Or (Val(Mid(cKey2, 2, 11)) = Val(Mid(TvwDir(Index).Nodes.item(p).key, 12, 21)) And Val(Mid(cKey2, 2, 11)) > 0) Then
       
                 If TvwDir(Index).Nodes.item(p).Checked <> lCheck1 Then
                     
                        lCheck2 = TvwDir(Index).Nodes.item(p).Checked
                        
                 End If
                     
                 If TvwDir(Index).Nodes.item(p).Children > 0 Then

                    cKey2 = Trim(TvwDir(Index).Nodes.item(p).key)

                 End If

              End If
             
          Next p
             
          If (Val(Mid(cKey, 2, 11)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 2, 11))) Then
          
              TvwDir(Index).Nodes.item(i).Checked = IIf(Not lCheck2, lCheck1, lCheck2)
          
          End If
          
          Exit For
          
       ElseIf TvwDir(Index).Nodes.item(i).Checked = True And TvwDir(Index).Nodes.item(i).Children > 0 And Val(Mid(cKey, 12, 21)) = Val(Mid(TvwDir(Index).Nodes.item(i).key, 2, 11)) Then
          
          Exit For
       
       ElseIf (TvwDir(Index).Nodes.item(i).Checked = True Or TvwDir(Index).Nodes.item(i).Children > 0) Then
          
       
       End If
   
   Next i

End If
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Function ValidarKey(ArregloKey As String, key As String, Caracter As String) As Boolean

Dim Ind           As Long
Dim cKeyArreglo() As String

ValidarKey = False
If Trim(ArregloKey) <> "" Then
cKeyArreglo = Split(ArregloKey, Caracter)

For Ind = 0 To UBound(cKeyArreglo)

    If cKeyArreglo(Ind) = Trim(key) Then
    
       ValidarKey = True
       Exit For
       
    End If

Next

End If
End Function


Private Sub TvwDir_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 27
    
    vg_opcion = 2
    Me.Hide

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub TvwDir_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)

On Error GoTo Man_Error

Set dest = Node

Select Case Index

Case 0
     
     CodCatDie = Val(Mid(TvwDir(0).Nodes(dest.Index).key, 2, 20))
     nomcatdie = Trim((TvwDir(0).Nodes(dest.Index).text))

Case 1
     
     codTippla = Val(Mid(TvwDir(1).Nodes(dest.Index).key, 2, 20))
     nomTippla = Trim((TvwDir(1).Nodes(dest.Index).text))

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub


