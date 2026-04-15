VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_PlantillaMenuII 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plantilla Menu"
   ClientHeight    =   9300
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   18360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   18360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   9735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   9735
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   7935
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   18015
      _Version        =   393216
      _ExtentX        =   31776
      _ExtentY        =   13996
      _StockProps     =   64
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   16
      SpreadDesigner  =   "M_PlantillaMenu.frx":0000
      VisibleCols     =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   18360
      _ExtentX        =   32385
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Servicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Menu Main 
      Caption         =   "Menu"
      Begin VB.Menu Principal 
         Caption         =   "Guardar"
         Index           =   1
      End
      Begin VB.Menu Estructura1 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu MenuDetalle 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu opgrilla 
         Caption         =   "Deshacer"
         Index           =   0
      End
      Begin VB.Menu opgrilla 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu opgrilla 
         Caption         =   "Cambiar Tipo de Plato"
         Index           =   3
      End
      Begin VB.Menu opgrilla 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu opgrilla 
         Caption         =   "Insertar Línea"
         Index           =   5
      End
      Begin VB.Menu opgrilla 
         Caption         =   "Eliminar Línea"
         Index           =   6
      End
      Begin VB.Menu opgrilla 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu opgrilla 
         Caption         =   "Subir Línea"
         Index           =   8
      End
      Begin VB.Menu opgrilla 
         Caption         =   "Bajar Línea"
         Index           =   9
      End
      Begin VB.Menu opgrilla 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu opgrilla 
         Caption         =   "Cortar"
         Index           =   11
      End
      Begin VB.Menu opgrilla 
         Caption         =   "Copiar"
         Index           =   12
      End
      Begin VB.Menu opgrilla 
         Caption         =   "Pegar"
         Index           =   13
      End
      Begin VB.Menu opgrilla 
         Caption         =   "Agregar Estructura"
         Index           =   14
         Begin VB.Menu Estructura2 
            Caption         =   ""
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "M_PlantillaMenuII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vectorcol() As Long
Dim modo            As String
Dim IndGrabado      As Integer
Private iblockrow   As Integer
Private iblockrow2  As Integer
Private iblockcol   As Integer
Private iblockcol2  As Integer
Private aiblockrow  As Integer
Private aiblockrow2 As Integer
Private aiblockcol  As Integer
Private aiblockcol2 As Integer

Private Sub Form_Activate()
    
On Error GoTo Man_Error

    Call fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    
End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Call fg_carga("")
Me.HelpContextID = vg_OpcM
Call fg_centra(Me)
modo = ""

    Toolbar1.ImageList = Partida.IL1
    Set BtnX = Toolbar1.Buttons.Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): BtnX.Visible = True: BtnX.ToolTipText = " "
    Set BtnX = Toolbar1.Buttons.Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): BtnX.Visible = False: BtnX.ToolTipText = "Grabar Datos": BtnX.Enabled = IIf(Mid(ValidarUsuario(M_MinSR1), 2, 2) = "0", False, True)
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Cortar", , tbrDefault, "A_Cortar"): BtnX.Visible = True: BtnX.ToolTipText = "Cortar"
    Set BtnX = Toolbar1.Buttons.Add(, "A_Copiar", , tbrDefault, "A_Copiar"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar"
    Set BtnX = Toolbar1.Buttons.Add(, "I_Pegar", , tbrDefault, "I_Pegar"): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Toolbar1.Buttons.Add(, "A_Pegar", , tbrDefault, "A_Pegar"): BtnX.Visible = False: BtnX.ToolTipText = "Pegar"
    Set BtnX = Toolbar1.Buttons.Add(, "A_Buscar", , tbrDefault, "A_Buscar"): BtnX.Visible = True: BtnX.ToolTipText = "Buscar Receta"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_InsertarF", , tbrDefault, "A_InsertarF"): BtnX.Visible = True: BtnX.ToolTipText = "Insertar"
    Set BtnX = Toolbar1.Buttons.Add(, "A_EliminarF", , tbrDefault, "A_EliminarF"): BtnX.Visible = True: BtnX.ToolTipText = "Eliminar"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_SubirF", , tbrDefault, "A_SubirF"): BtnX.Visible = True: BtnX.ToolTipText = "Subir"
    Set BtnX = Toolbar1.Buttons.Add(, "A_BajarF", , tbrDefault, "A_BajarF"): BtnX.Visible = True: BtnX.ToolTipText = "Bajar"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Planificación Teórica"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False

    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False

    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False
    Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDropdown, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Minuta Bloque a Excel ": BtnX.ButtonMenus.Add text:="Formato I": BtnX.ButtonMenus.Add text:="Formato II Resumido"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Deshacer", , tbrDefault, "A_Deshacer"): BtnX.Visible = True: BtnX.Enabled = False: BtnX.ToolTipText = "Deshacer"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Retrocede", , tbrDefault, "A_Retrocede"): BtnX.Visible = True: BtnX.Enabled = True: BtnX.ToolTipText = "Retrocede Minuta Bloque"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Avanza", , tbrDefault, "A_Avanza"): BtnX.Visible = True: BtnX.Enabled = True: BtnX.ToolTipText = "Avanza Minuta Bloque"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
    Me.HelpContextID = vg_OpcM

Text2.text = ""
Text2.text = vg_codservicio & " - " & vg_nombre
Text1.Enabled = False

If Vg_PlaSer = "1" Then

   vaSpread1.MaxRows = 10
   Text1.Enabled = True
   Text1.text = ""
   
End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index
        
    Case 2
            
        If Toolbar1.Buttons(2).Enabled = False Then IndGrabado = 0: Exit Sub
        If MsgBox(" Actualiza plantilla Menú...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then
               
           IndGrabado = 0
'           Plantilla(0).Enabled = False
           Toolbar1.Buttons(1).Visible = True
           Toolbar1.Buttons(2).Visible = False
           Toolbar1.Buttons(31).Enabled = False
           Cancel = -1
           Exit Sub
            
        End If
            
        If IndGrabado = 1 Then GrabarPlantillaMenu
        IndGrabado = 0
'        Plantilla(0).Enabled = False
        Toolbar1.Buttons(1).Visible = True
        Toolbar1.Buttons(2).Visible = False
'        Toolbar1.Buttons(31).Enabled = False
    
    Case 5 'copiar
            
        If Index = 11 Then
            
           'Validar recetas 5 etapas
           j = 0
           For i = 1 To MaxColumna
               
               If (vectorcol(i) - 2) = iblockcol Or vectorcol(i) = iblockcol Then j = (vectorcol(i) - 2): Exit For
           
           Next i
           If j = 0 Then Exit Sub
        
        End If
        
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = vaSpread1.ActiveCol
        
        aiblockrow = iblockrow: aiblockrow2 = iblockrow2
        aiblockcol = iblockcol: aiblockcol2 = iblockcol2
        If vaSpread1.Col = 1 Or vaSpread1.Col = 2 Then Exit Sub
        If vaSpread1.MaxRows > 1000 Then Del_Row = vaSpread1.MaxRows - 1000: vaSpread1.MaxRows = vaSpread1.MaxRows - Del_Row
'        Plato(13).Enabled = True: opgrilla(13).Enabled = True
'        Plato(14).Enabled = True: opgrilla(14).Enabled = True
        Toolbar1.Buttons(6).Visible = False: Toolbar1.Buttons(7).Visible = True
        If iblockcol < 1 Then aiblockcol = 1: aiblockcol2 = vaSpread1.maxcols
        indcortarpegar = 1
        If Index = 11 Then
           
'           indcortarpegar = 0
'           Toolbar1.Buttons(8).Visible = True
'           Toolbar1.Buttons(9).Visible = False
'           Plato(14).Enabled = False
'           opgrilla(14).Enabled = False
           
        Else
           
'           Toolbar1.Buttons(8).Visible = False
'           Toolbar1.Buttons(9).Visible = True
'           Plato(14).Enabled = True
'           opgrilla(14).Enabled = True
        
        End If

    Case 6 'Pegar
    
    
End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

   If Col < 1 Then Exit Sub
                                                                   
'    Let OpGrilla(15).Enabled = IIf(Col = 1 And (vg_codregimen > 9999 And AddReceta = 0), False, True)
'    Let Plato(15).Enabled = IIf(Col = 1 And (vg_codregimen > 9999 And AddReceta = 0), False, True)
    Let indactivo = 1
    Let iblockrow = vaSpread1.ActiveRow
    Let iblockrow2 = vaSpread1.ActiveRow
    Let iblockcol = vaSpread1.ActiveCol
    Let iblockcol2 = vaSpread1.ActiveCol
    Let vaSpread1.Row = vaSpread1.ActiveRow
    Let vaSpread1.Col = vaSpread1.ActiveCol
     
'    If Col = 1 Then Plato_Click (16): Exit Sub
     
Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim cod As Variant
Dim estructuraservicio As Long
Dim nombreEstructura As String
If Row = 1 Then
    
    vaSpread1.Col = 1
    nombreEstructura = vaSpread1.text
    
    If nombreEstructura = "" Then
       
       MsgBox "No puede ingresar tipo plato sin no tener Estructura de Servicio", 16
       Exit Sub
    
    End If

End If

'If Row < 1 Or Col = 1 Then Exit Sub
If Col = 1 Then

   Let vg_nombre = ""
   Let vg_codigo = ""
   Call B_TabEst.LlenaDatos("a_estservicio", "" & vg_codservicio & "", "Estructura Servicio", "EstSer")
   Call B_TabEst.Show(1) '
   Me.Refresh
   If vg_codigo = "" Then Exit Sub
   vaSpread1.Row = Row
   vaSpread1.Col = Col
   vaSpread1.text = vg_nombre
   
   vaSpread1.Col = vaSpread1.maxcols
   vaSpread1.text = Val(vg_codigo)

   vaSpread1.Row = vaSpread1.ActiveRow
   IndGrabado = 1
'   Plato(0).Enabled = True: opgrilla(0).Enabled = True
'   Plato(13).Enabled = False: opgrilla(13).Enabled = False
'   Plato(14).Enabled = False: opgrilla(14).Enabled = False
'   Plantilla(0).Enabled = True
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = True
   Toolbar1.Buttons(6).Visible = True
   Toolbar1.Buttons(7).Visible = False

Else
    
    Call Plato_Click(0)

End If

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

If vaSpread1.MaxRows < 1 Then Exit Sub
iblockrow = NewRow
iblockrow2 = NewRow
iblockcol = NewCol
iblockcol2 = NewCol
If NewRow < 0 Then iblockrow = 1
If NewRow < 0 Then iblockrow2 = (vaSpread1.MaxRows - 1)
If NewRow >= vaSpread1.MaxRows Then iblockrow2 = (vaSpread1.MaxRows - 1)
If indcos = False Or NewCol < 1 Then Exit Sub

End Sub

Private Sub vaspread1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

On Error GoTo Man_Error

Select Case Button

Case 2
    
    If vaSpread1.Visible <> True Then Exit Sub
    '-------> Validar si minuta esta bloqueada
    If ValidarBloqueoMinuta Then Exit Sub
    Indvaspread1 = 0
    PopupMenu MenuDetalle

End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub

Private Sub Estructura2_Click(Index As Integer)

On Error GoTo Man_Error

    LlenaSubMenu Estructura2, Index

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub

Sub LlenaSubMenu(SubMenu As Object, Index As Integer)

On Error GoTo Man_Error
    
    Dim i           As Long
    Dim j           As Long
    Dim colgrupo    As Long
    Dim CodigoGrupo As Long
    Dim RowGrupoMin As Long
    Dim RowGrupoMax As Long
    
    If vaSpread1.ActiveRow = vaSpread1.MaxRows Then
       
       Call MsgBox("No puede insertar estructura fija ultima fila", vbCritical + vbOKOnly, Msgtitulo)
       Exit Sub
    
    End If
    
    '-------> Rescata El Codigo de Agrupacion
    Dim RS2  As New ADODB.Recordset
    Set RS2 = vg_db.Execute("sgpadm_Sel_CodigodeAgrupacionServicio " & vg_codservicio & ", " & SubMenu(Index).HelpContextID & "")
    CodigoGrupo = RS2!ess_agrupacionestructura
    RS2.Close
    Set RS2 = Nothing
    
    GrabarCambios 1, 1, "Estructura Servicio"
    'columna de grupo de estructura y Encabezado
    colgrupo = vaSpread1.GetColFromID("Grupo") + 1
    '-------> Buscar grupo estructura servicio
    RowGrupoMin = vaSpread1.SearchCol(colgrupo, 0, -1, CodigoGrupo, SearchFlagsValue)
    If RowGrupoMin > 0 Then
       
       For i = RowGrupoMin To vaSpread1.MaxRows - 1
           
           vaSpread1.Row = i
           vaSpread1.Col = vaSpread1.maxcols
           
           If Val(vaSpread1.text) <> CodigoGrupo And Trim(vaSpread1.text) <> "" Then
              
              Exit For
           
           End If
           
           RowGrupoMax = i
       
       Next i
       
       If vaSpread1.ActiveRow >= RowGrupoMin And vaSpread1.ActiveRow <= RowGrupoMax Then
          
          vaSpread1.Row = vaSpread1.ActiveRow
          
          If Trim(vaSpread1.text) = "" Then
             
             RowGrupoMax = vaSpread1.ActiveRow
          
          Else
             
             RowGrupoMax = RowGrupoMax + 1
             vaSpread1.MaxRows = vaSpread1.MaxRows + 1
             vaSpread1.InsertRows RowGrupoMax, 1
          
          End If
       
       Else
          
          RowGrupoMax = RowGrupoMax + 1
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
          vaSpread1.InsertRows RowGrupoMax, 1
       
       End If
    
    Else
       
       vaSpread1.Row = vaSpread1.ActiveRow
       
       If Val(vaSpread1.text) >= 0 Then
          
          RowGrupoMax = vaSpread1.Row
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
          vaSpread1.InsertRows RowGrupoMax, 1
       
       Else
          
          'columna de grupo de estructura y Encabezado
          colgrupo = vaSpread1.GetColFromID("Grupo") + 1
          '-------> Buscar grupo estructura servicio
          vaSpread1.Col = colgrupo
          RowGrupoMin = vaSpread1.SearchCol(colgrupo, 0, -1, Trim(vaSpread1.text), SearchFlagsValue)
          RowGrupoMax = 0
          
          For i = RowGrupoMin To vaSpread1.MaxRows - 1
              
              vaSpread1.Row = i
              vaSpread1.Col = vaSpread1.maxcols
              
              If Val(vaSpread1.text) <> CodigoGrupo And Trim(vaSpread1.text) <> "" Then
                 
                 Exit For
              
              End If
              
              RowGrupoMax = i
          
          Next i
          
          RowGrupoMax = RowGrupoMax + 1
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
          vaSpread1.InsertRows RowGrupoMax, 1
       
       End If

    End If

    vaSpread1.Row = RowGrupoMax
    '-------> Mover codigo estructura
    vaSpread1.Col = vaSpread1.maxcols - 1
    vaSpread1.text = SubMenu(Index).HelpContextID
    '-------> Mover grupo estructura
    vaSpread1.Col = vaSpread1.maxcols
    vaSpread1.text = CodigoGrupo
    '-------> Mover descripción estructura servico
    vaSpread1.Col = 1
    vaSpread1.text = SubMenu(Index).Caption
    vaSpread1.Col = vaSpread1.maxcols: vaSpread1.text = CodigoGrupo
    vaSpread1.Col = vaSpread1.maxcols - 1: vaSpread1.text = SubMenu(Index).HelpContextID

    '-------> Mover color a las lineas nuevas
    vaSpread1.Row = RowGrupoMax
    vaSpread1.Col = -1
'    vaSpread1.BackColor = Shape1(0).FillColor
    '-------> Mover color a la columna estructura servicio
    vaSpread1.Col = 1
'    vaSpread1.BackColor = Shape1(2).FillColor

    Estructura1(Index).Enabled = False: Estructura2(Index).Enabled = False
'    IndGrabado = 1
'    Plantilla(0).Enabled = True
'    Toolbar1.Buttons(1).Visible = False
'    Toolbar1.Buttons(2).Visible = True
    
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Opgrilla_Click(Index As Integer)
    
    Select Case Index
        
        Case 0
            
            Call Plato_Click(2)
        
        Case 2
            
            Call Plato_Click(2)
        
        Case 3
            
            Call Plato_Click(3)
        
        Case 5
            
            Call Plato_Click(5)
        
        Case 6
            
            Call Plato_Click(6)
        
        Case 8
            
            Call Plato_Click(8)
        
        Case 9
            
            Call Plato_Click(9)
        
        Case 11
            Call Plato_Click(11)
        Case 12
            
            Call Plato_Click(12)
        
        Case 13
            
            Call Plato_Click(13)
        
        Case 14
            
            Call Plato_Click(14)
        
        Case 15
            
            Call Plato_Click(15)
    
    End Select

End Sub

Private Sub Plato_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS                  As New ADODB.Recordset
Dim Del_Row             As Integer
Dim c                   As Long
Dim IndCol              As Integer
Dim indrow              As Integer
Dim indcol2             As Integer
Dim indrow2             As Integer
Dim indrow3             As Integer
Dim FilaAct             As Long
Dim FilaAnt             As Long
Dim FilaPos             As Long
Dim AuxIblockrow        As Integer
Dim addrec              As Long
Dim codest              As Long
Dim cosali              As Double
Dim CosDes              As Double
Dim NroMes              As String
Dim FinGrilla           As String
Dim MesInicio           As String
Dim FechaBusqueda       As String
Dim SumaMes             As String
Dim MesInicio3          As String
Dim xx                  As Long
Dim xp                  As Long
Dim FechaDia            As Long
Dim SeleccionOpt        As Long
Dim CodGrupoEstBaj      As Long
Dim CantTotalPorcentaje As Double
            
Dim VecSelGrid          As Variant
Dim VecRacPegar         As Variant
Dim contador            As Long
Dim contador_b          As Long
Dim cantCol             As Long
Dim LargoVec            As Long
Dim accion              As String
Dim ColumnaActiva       As Long
Dim FilaActiva          As Long
Dim ColumnaAntActiva    As Long
Dim n                   As Long
Dim n1                  As Long
Dim NFilas              As Long
Dim CantCol1            As Long
Dim d                   As Variant
Dim Max                 As Long
Dim max1                As Long
Dim ff                  As Long
Dim f                   As Long
Dim desc                As String
Dim g                   As Long
Dim j                   As Long
Dim tope                As Long
Dim jjj                 As Long

contador = 0
contador_b = 0
cantCol = 0
LargoVec = 0
accion = ""
n1 = 0
n = 0
NFilas = 0

'Fila Maestra de Grupo
    
    
    Select Case Index
        
        Case 0 '-------> Ingresar recetas
            
            ws_respuesta = ""
            Let ColumnaReceta = Col
            vg_codigo = "": vg_nombre = ""
            'vg_left = fpayuda(1).Left + 2400
            B_ArbEst.MoverDatosTvwDir "a_recetatippla", "tip_", "Tipo Plato", "0"
            B_ArbEst.Show 1
            If Trim(vg_codigo) = "" Then Exit Sub
            FilTipPla = Val(vg_codigo)
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.Col = vaSpread1.ActiveCol
            vaSpread1.text = vg_nombre
            vaSpread1.Col = vaSpread1.ActiveCol + 1
            vaSpread1.text = Val(vg_codigo)
            
            vaSpread1.Row = vaSpread1.ActiveRow
            IndGrabado = 1
'            Plato(0).Enabled = True: opgrilla(0).Enabled = True
'            Plato(13).Enabled = False: opgrilla(13).Enabled = False
'            Plato(14).Enabled = False: opgrilla(14).Enabled = False
'            Plantilla(0).Enabled = True
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = True
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False
            
            fg_descarga

        
    End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub GrabarCambios(ifil As Long, icol As Long, estado As String)

On Error GoTo Man_Error

Dim ret
ContadorDeshacer = ContadorDeshacer + 1
ret = vaSpread1.SaveToFile(LCase(App.Path) & "\" & "spreadPMenu" & vg_NUsr & ContadorDeshacer & ".ss6", False)
'Toolbar1.Buttons(34).Visible = True
'Toolbar1.Buttons(34).Enabled = True

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub

Private Sub GrabarPlantillaMenu()

On Error GoTo Man_Error
    
Dim RS               As New ADODB.Recordset
Dim FecMinuta        As Long
Dim IndDia           As Long
Dim MyBuffer         As String
Dim TipoPlato        As Long
Dim PorcenDiario     As Long
Dim PorcenTotal      As Long
Dim CodEstructura    As Long
Dim NumLin           As Long
Dim Fecha            As Long

Dim dia              As Variant
Dim Pos              As Variant
Dim Cabecera         As Long

Dim Sql              As String
    
IndDia = 1
'gauge1.Value = 0
'gauge.Value = 0
'Picture1.Visible = True
'Label3.Visible = True
'gauge.Visible = True
'Picture1.Refresh
'Label3.Refresh
'gauge.Refresh
'gauge1.Refresh
    
fg_carga ""
    
vaSpread1.Enabled = False
Toolbar1.Enabled = False
'Main(0).Enabled = False
'Main(1).Enabled = False
Let IndDia = 1
Let EstGrpEst = True
        
        For i = 2 To (vaSpread1.maxcols - 1) Step 1
          
            DoEvents
'            gauge1.Value = Val((IndDia / MaxColumna) * 100)
            
            vaSpread1.Row = 0
            vaSpread1.Col = i
            dia = Right(vaSpread1.text, 1)
'            Label3.Caption = "": Label3.Caption = "Día : " & dia
                    
            Let MyBuffer = ""
            Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
            Let MyBuffer = MyBuffer & "<GrabaPlantilla>"
            Let NumLin = 1

            For j = 1 To (vaSpread1.MaxRows - 1)
                       
'                gauge.Value = Val((j / (vaSpread1.MaxRows - 1)) * 100)
                TipoPlato = 0
                        
                vaSpread1.Row = j
                
                '-------> Sacar codigo estructura servicio
                vaSpread1.Col = vaSpread1.maxcols
                If Trim(vaSpread1.text) <> "" And Val(vaSpread1.text) > 0 Then CodEstructura = vaSpread1.text
                
                vaSpread1.Col = i + 1
                TipoPlato = Val(vaSpread1.text)
                
                If TipoPlato > 0 And CodEstructura > 0 Then
                   
                            
                   MyBuffer = MyBuffer & " <Menu"
                   MyBuffer = MyBuffer & " Op = " & Chr(34) & 0 & Chr(34)

                   MyBuffer = MyBuffer & " CodEstructura = " & Chr(34) & CodEstructura & Chr(34)
                   MyBuffer = MyBuffer & " TipoPlato = " & Chr(34) & TipoPlato & Chr(34)
                   MyBuffer = MyBuffer & " NumLin = " & Chr(34) & NumLin & Chr(34)
                   MyBuffer = MyBuffer & " Dia = " & Chr(34) & dia & Chr(34)
                   MyBuffer = MyBuffer & "/>"
                   
                End If
                NumLin = NumLin + 1
            
            Next j
            
            IndDia = IndDia + 1
            MyBuffer = MyBuffer & "</GrabaPlantilla>"
            
            Set RS = vg_db.Execute("sgpadm_Ins_XmlPlantillaMenu '" & MyBuffer & "', '" & vg_codservicio & "', " & dia & ", '" & LimpiaDato(Trim(Text1.text)) & "', " & vg_IDBloque & "")
            If Not RS.EOF Then
            
               If RS(0) > 0 Then
                  
                  MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, Msgtitulo
               
               End If
            
            End If
            RS.Close: Set RS = Nothing
            Let EstGrpEst = False
        
        Next i
        
        Picture1.Visible = False: gauge.Visible = False
        vaSpread1.Enabled = True
        'Main(0).Enabled = True
        'Main(1).Enabled = True
        vaSpread1.Refresh
        Toolbar1.Enabled = True
        '-------> Mover Datos grilla principal del formulario m_minsr1
'        Sql = ""
'        Sql = LimpiaDato(Trim(M_MinSR1.fpText.text))
'        Sql = Sql & ", " & M_MinSR1.fpLongInteger1(0).Value & ", " & M_MinSR1.fpLongInteger1(1).Value & ""
'        Set RS = vg_db.Execute("sgpadm_Sel_ListarMinutaBloquexCeco " & Sql & "")
'        M_MinSR1.vaSpread1.MaxRows = 0
'        M_MinSR1.vaSpread1.Row = -1
'        M_MinSR1.vaSpread1.Col = -1
'        M_MinSR1.vaSpread1.BackColor = Shape1(0).FillColor  'Amarillo
'        Do While Not RS.EOF = True
'
'            M_MinSR1.vaSpread1.MaxRows = M_MinSR1.vaSpread1.MaxRows + 1
'            M_MinSR1.vaSpread1.Row = M_MinSR1.vaSpread1.MaxRows
'
'            If RS!IdEstadoMinuta <> 11 Then
'
'                M_MinSR1.vaSpread1.Col = -1
'                M_MinSR1.vaSpread1.BackColor = M_MinSR1.Shape1(1).FillColor ' Rojo
'
'            End If
'
'            M_MinSR1.vaSpread1.Col = 2
'            M_MinSR1.vaSpread1.text = CStr(RS!Id_Bloque)
'            M_MinSR1.vaSpread1.Col = 3
'            M_MinSR1.vaSpread1.text = RS!reg_codigo & " - " & Trim(RS!reg_nombre)
'            M_MinSR1.vaSpread1.Col = 4
'            M_MinSR1.vaSpread1.text = RS!ser_codigo & " - " & Trim(RS!ser_nombre)
'            M_MinSR1.vaSpread1.Col = 5
'            M_MinSR1.vaSpread1.text = Format(RS!fechadesde, "dd/mm/yyyy")
'            M_MinSR1.vaSpread1.Col = 6
'            M_MinSR1.vaSpread1.text = Format(RS!fechahasta, "dd/mm/yyyy")
'            M_MinSR1.vaSpread1.Col = 7
'            M_MinSR1.vaSpread1.text = RS!reg_codigo
'            M_MinSR1.vaSpread1.Col = 8
'            M_MinSR1.vaSpread1.text = RS!ser_codigo
'            M_MinSR1.vaSpread1.Col = 9
'            M_MinSR1.vaSpread1.text = Trim(RS!reg_nombre)
'            M_MinSR1.vaSpread1.Col = 10
'            M_MinSR1.vaSpread1.text = Trim(RS!ser_nombre)
'            RS.MoveNext
'
'        Loop
'        RS.Close: Set RS = Nothing
'        M_MinSR1.FpFecDesde.Enabled = True
'        M_MinSR1.FpFecHasta.Enabled = True
        fg_descarga

Exit Sub
Man_Error:
    Picture1.Visible = False: gauge.Visible = False
    vaSpread1.Enabled = True
'    Main(0).Enabled = True
'    Main(1).Enabled = True
'    Toolbar1.Enabled = True
    Call fg_descarga
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Msgtitulo)
    Call ins_log_error(Date & Time & Err & ":  " & Error$(Err))

End Sub

