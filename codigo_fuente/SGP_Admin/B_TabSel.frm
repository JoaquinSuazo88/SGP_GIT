VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form B_TabSel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Servicio con Estructura"
   ClientHeight    =   7575
   ClientLeft      =   6030
   ClientTop       =   2130
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   7770
   Begin VB.Frame Frame1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin MSComctlLib.TreeView TvwDir 
         Height          =   6885
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   12144
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7575
      Left            =   7230
      TabIndex        =   2
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   13361
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "B_TabSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rootNode     As Node
Dim ArNivTree(2) As Variant
Dim fso
Private BtnX     As Variant
Dim Arbol1       As Object
Dim MsgTitulo    As String

Private Sub Form_Activate()
    
    Call fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error
    
    Call fg_centra(Me)
    Me.Width = 7890
    Me.Left = vg_left
    fg_carga ""
    MsgTitulo = "Estructura Servicio"
    Toolbar1.ImageList = Partida.IL1
    Toolbar1.Buttons.Clear
    Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
    Call fg_descarga
    Exit Sub

Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, MsgTitulo)

End Sub

Sub LlenaDatos(Arbol As Object, CodigoSubSeg As Long, CodigoRegimen As Long, FechaInicial As Long, FechaFinal As Long, IndProReal As String, OpcionLectura As String)

On Error GoTo Man_Error

Dim RS                   As New ADODB.Recordset
Dim AuxCodigoServicio    As Long
Dim AuxCodigoEstServicio As Long
Dim pcodser              As String
Dim i                    As Long

MsgTitulo = "Estructura Servicio"
Set Arbol1 = Arbol

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
Set RS = vg_db.Execute("sgpadm_Sel_ServicioEstructura " & CodigoSubSeg & ", " & CodigoRegimen & ", " & FechaInicial & ", " & FechaFinal & "")
ArNivTree(0) = "Servicio"   'Código Servicio
ArNivTree(1) = 6  'Largo de Subsegmento
i = 1
TvwDir.Nodes.Clear

Do While Not RS.EOF
   
   If RS(0) <> AuxCodigoServicio Then
      
      padre = Chr(nivel)
      Set rootNode = TvwDir.Nodes.Add(, , "N" & fg_pone_espacio(RS(0), 5), RS(0) & " - " & Trim(RS(1)))
      pcodser = "": pcodser = "N" & fg_pone_espacio(RS(0), 5): AuxCodigoServicio = RS(0)
      xcencos = "": xcodreg = 0: xcodser = 0
   
   End If
   
   If RS(2) <> AuxCodigoEstServicio Then
      
      padre = Chr(nivel)
      Set rootNode = TvwDir.Nodes.Add(pcodser, tvwChild, pcodser & "Servicio" & fg_pone_espacio(RS(2), 10), Trim(RS(2)) & " - " & Trim(RS(3)))
      AuxCodigoEstServicio = RS(2)
   
   End If
   
   RS.MoveNext
   i = i + 1

Loop
RS.Close
Set RS = Nothing

For i = 1 To Arbol.Nodes.count
    
    DoEvents
    If i <= TvwDir.Nodes.count Then
       TvwDir.Nodes.item(i).Checked = Arbol.Nodes.item(i).Checked
    End If
Next i

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Sub LlenaDatosBloque(Arbol As Object, Ceco As String, CodigoRegimen As Long, FechaInicial As Long, FechaFinal As Long, IndProReal As String, OpcionLectura As String)

On Error GoTo Man_Error

Dim RS                   As New ADODB.Recordset
Dim AuxCodigoServicio    As Long
Dim AuxCodigoEstServicio As Long
Dim pcodser              As String
Dim i                    As Long

MsgTitulo = "Estructura Servicio"
Set Arbol1 = Arbol

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
Set RS = vg_db.Execute("sgpadm_Sel_ServicioEstructuraMinutaBloqueI_V02 '" & Ceco & "', " & CodigoRegimen & ", " & FechaInicial & ", " & FechaFinal & "")
ArNivTree(0) = "Servicio"   'Código Servicio
ArNivTree(1) = 6  'Largo de Subsegmento
i = 1
TvwDir.Nodes.Clear
Do While Not RS.EOF
   
   If RS(0) <> AuxCodigoServicio Then
      
      padre = Chr(nivel)
      Set rootNode = TvwDir.Nodes.Add(, , "N" & fg_pone_espacio(RS(0), 5), RS(0) & " - " & Trim(RS(1)))
      pcodser = ""
      pcodser = "N" & fg_pone_espacio(RS(0), 5)
      AuxCodigoServicio = RS(0)
      xcencos = ""
      xcodreg = 0
      xcodser = 0
   
   End If
   
   If RS(2) <> AuxCodigoEstServicio Then
      
      padre = Chr(nivel)
      Set rootNode = TvwDir.Nodes.Add(pcodser, tvwChild, pcodser & "Servicio" & fg_pone_espacio(RS(2), 10), Trim(RS(2)) & " - " & Trim(RS(3)))
      AuxCodigoEstServicio = RS(2)
   
   End If
   
   RS.MoveNext
   i = i + 1

Loop
RS.Close
Set RS = Nothing

For i = 1 To Arbol.Nodes.count
    
    DoEvents
    If i <= TvwDir.Nodes.count Then
       TvwDir.Nodes.item(i).Checked = Arbol.Nodes.item(i).Checked
    End If
Next i

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Sub LlenaDatosCecoFechas(Arbol As Object, Ceco As String, FechaInicial As Long, FechaFinal As Long)

On Error GoTo Man_Error

Dim RS                As New ADODB.Recordset
Dim AuxCodigoRegimen  As Long
Dim AuxCodigoServicio As Long
Dim pcodreg           As String
Dim i                 As Long

MsgTitulo = "Regimen / Servicios"

Set Arbol1 = Arbol
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_RegimenServicioMinutaBloque '" & Ceco & "', " & FechaInicial & ", " & FechaFinal & "")
ArNivTree(0) = "Regimen"   'Código Servicio
ArNivTree(1) = 6  'Largo de Servicio
i = 1
TvwDir.Nodes.Clear
Do While Not RS.EOF
   
   If RS(0) <> AuxCodigoRegimen Then
      
      padre = Chr(nivel)
      Set rootNode = TvwDir.Nodes.Add(, , "N" & fg_pone_espacio(RS(0), 5), RS(0) & " - " & Trim(RS(1)))
      pcodreg = ""
      pcodreg = "N" & fg_pone_espacio(RS(0), 5)
      AuxCodigoRegimen = RS(0)
   
   End If
   
   If RS(2) <> AuxCodigoServicio Then
      
      padre = Chr(nivel)
      Set rootNode = TvwDir.Nodes.Add(pcodreg, tvwChild, pcodreg & "Servicio" & fg_pone_espacio(RS(2), 10), Trim(RS(2)) & " - " & Trim(RS(3)))
      AuxCodigoServicio = RS(2)
   
   End If
   
   RS.MoveNext
   i = i + 1

Loop
RS.Close
Set RS = Nothing

For i = 1 To Arbol.Nodes.count
    
    DoEvents
    If i <= TvwDir.Nodes.count Then
       TvwDir.Nodes.item(i).Checked = Arbol.Nodes.item(i).Checked
    End If
Next i

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim i As Long

Select Case Button.Index

Case 1
    
    For i = 1 To Arbol1.Nodes.count
        
        DoEvents
        If i <= TvwDir.Nodes.count Then
           Arbol1.Nodes.item(i).Checked = TvwDir.Nodes.item(i).Checked
        End If
    Next i
    vg_codigo = "1"
    Me.Hide
    Unload Me

Case 3
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub TvwDir_NodeCheck(ByVal Node As MSComctlLib.Node)

On Error GoTo Man_Error

Dim cKey As String, lKey As Integer, i As Long, lCheck As Boolean, MarcarAsc As Boolean, lGraba As Boolean
TvwDir.Nodes.item(Node.key).Selected = True
lCheck = TvwDir.Nodes.item(TvwDir.SelectedItem.Index).Checked
cKey = Trim(TvwDir.Nodes.item(TvwDir.SelectedItem.Index).key)
lKey = Len(cKey)

Dim MarcarDesc As Boolean, INiv As Integer, RecNivel As String
MarcarAsc = False
If lCheck Then
    
    MarcarDesc = True: INiv = 1
    RecNivel = Mid(TvwDir.Nodes.item(TvwDir.SelectedItem.Index).key, 1, ArNivTree(INiv))

End If
'------->
For i = 1 To TvwDir.Nodes.count
    
    If cKey = Mid(TvwDir.Nodes.item(i).key, 1, lKey) Then
        
        TvwDir.Nodes.item(i).Checked = lCheck
        lGraba = True
    
    End If
    
    '-------> Comando marcas descendentes
    If MarcarDesc And Trim(TvwDir.Nodes.item(i).key) = RecNivel Then
        
        INiv = INiv + 1
        RecNivel = Mid(TvwDir.Nodes.item(TvwDir.SelectedItem.Index).key, 1, ArNivTree(INiv))
        TvwDir.Nodes.item(i).Checked = True
    
    End If
    '------->

Next i
fg_descarga

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub
