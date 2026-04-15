VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_Plami2 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planificación Teórica"
   ClientHeight    =   8070
   ClientLeft      =   2370
   ClientTop       =   2640
   ClientWidth     =   11970
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11400
      Top             =   4680
   End
   Begin VB.Frame Frame2 
      Height          =   2625
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   5190
      Visible         =   0   'False
      Width           =   15195
      Begin VB.Frame Frame2 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   56
         Top             =   120
         Visible         =   0   'False
         Width           =   2775
         Begin MSComctlLib.ProgressBar Bar1 
      End
   End
   Begin VB.Menu Main 
      Caption         =   "Menú"
      Index           =   0
      NegotiatePosition=   2  'Middle
      Begin VB.Menu Plantilla 
         Caption         =   "&Grabar Semana"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Ver &Receta"
         Index           =   5
      End
      Begin VB.Menu Plantilla 
         Caption         =   "C&opiar Minutas"
         Index           =   8
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Aporte &Nutricionales x Días"
         Index           =   10
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Costo Receta"
         Index           =   11
      End
      Begin VB.Menu Plantilla 
         Caption         =   "&Frecuencia Recetas"
         Index           =   12
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Ac&tualizar Costo Planificación"
         Index           =   13
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Exportar Recetas"
         Index           =   14
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Parámetro de Grabado"
         Index           =   16
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu Plantilla 
         Caption         =   "&Cerrar"
         Index           =   22
      End
   End
   Begin VB.Menu Main 
      Caption         =   "&Plato Menú"
      Index           =   1
      NegotiatePosition=   2  'Middle
      Begin VB.Menu Plato 
         Caption         =   "&Deshacer"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu Plato 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu Plato 
         Caption         =   "Cambiar Plato &Menú"
         Index           =   2
      End
      Begin VB.Menu Plato 
         Caption         =   "Come&ntario"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu Plato 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu Plato 
         Caption         =   "&Insertar"
         Index           =   5
      End
      Begin VB.Menu Plato 
         Caption         =   "&Eliminar"
         Index           =   6
      End
      Begin VB.Menu Plato 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu Plato 
         Caption         =   "&Subir"
         Index           =   8
      End
      Begin VB.Menu Plato 
         Caption         =   "&Bajar"
         Index           =   9
      End
      Begin VB.Menu Plato 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu Plato 
         Caption         =   "Cor&tar"
         Index           =   11
         Shortcut        =   ^X
      End
      Begin VB.Menu Plato 
         Caption         =   "C&opiar"
         Index           =   12
         Shortcut        =   ^C
      End
      Begin VB.Menu Plato 
         Caption         =   "&Pegar"
         Enabled         =   0   'False
         Index           =   13
         Shortcut        =   ^V
      End
      Begin VB.Menu Plato 
         Caption         =   "Pegado &Especial"
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu Plato 
         Caption         =   "&Buscar Recetas o Ingredientes"
         Index           =   15
         Shortcut        =   ^B
      End
      Begin VB.Menu Plato 
         Caption         =   "Crear Estr&uctura"
         Index           =   17
      End
      Begin VB.Menu Plato 
         Caption         =   "&Agrega Estructura"
         Index           =   18
         Begin VB.Menu Estructura1 
            Caption         =   ""
            Index           =   0
         End
      End
   End
   Begin VB.Menu Main 
      Caption         =   "&Ver"
      Index           =   2
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu Ver 
         Caption         =   "Días &Pantalla"
         Index           =   0
         Visible         =   0   'False
         Begin VB.Menu Dias 
            Caption         =   "&1"
            Index           =   0
         End
         Begin VB.Menu Dias 
            Caption         =   "&2"
            Index           =   1
         End
         Begin VB.Menu Dias 
            Caption         =   "&3"
            Index           =   2
         End
         Begin VB.Menu Dias 
            Caption         =   "&4"
            Index           =   3
         End
         Begin VB.Menu Dias 
            Caption         =   "&5"
            Index           =   4
         End
         Begin VB.Menu Dias 
            Caption         =   "&6"
            Index           =   5
         End
         Begin VB.Menu Dias 
            Caption         =   "&7"
            Index           =   6
         End
      End
      Begin VB.Menu Ver 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu Ver 
         Caption         =   "&Semana Siguiente"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu Ver 
         Caption         =   "Semana &Anterior"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu Ver 
         Caption         =   "-"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu Ver 
         Caption         =   "Costo Minutas"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu Ver 
         Caption         =   "Aporte &Nutricional x Día"
         Index           =   6
      End
      Begin VB.Menu Ver 
         Caption         =   "&Gramos Productos Mensual"
         Index           =   7
      End
      Begin VB.Menu Ver 
         Caption         =   "&Frecuencia De Recetas"
         Index           =   8
      End
      Begin VB.Menu Ver 
         Caption         =   "&Costo Minuta Resumido"
         Index           =   9
      End
   End
   Begin VB.Menu MenuDetalle 
      Caption         =   ""
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu OpGrilla 
         Caption         =   "Deshacer"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Cambiar Plato &Menú"
         Index           =   2
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Come&ntario"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Insertar"
         Index           =   5
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Eliminar"
         Index           =   6
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Subir"
         Index           =   8
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Bajar"
         Index           =   9
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Cor&tar"
         Index           =   11
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "C&opiar"
         Index           =   12
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Pegar"
         Enabled         =   0   'False
         Index           =   13
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Pegado Especial"
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Buscar Receta"
         Index           =   15
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Ag&rega Estructura Personalizada"
         Index           =   16
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Crear Estr&uctura"
         Index           =   17
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Agrega Estructura"
         Index           =   18
         Begin VB.Menu Estructura2 
            Caption         =   ""
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "M_Plami2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset, RSTem As New ADODB.Recordset
Dim i As Long, j As Long, indcortarpegar As Long, fecha As Long, maxColumna As Long, maxfila As Long
Dim iblockrow As Integer, iblockrow2 As Integer, iblockcol As Integer, iblockcol2 As Integer, SwSalir As Integer
Dim aiblockrow As Integer, aiblockrow2 As Integer, aiblockcol As Integer, aiblockcol2 As Integer, indactivo As Integer
Dim indcos As Boolean, estgra As Boolean, estapo As Boolean
Dim veccos() As Variant
Dim vectorcol() As Long
Dim Msgtitulo As String
Dim TipoCopia As String, NameTemp As String
Dim SpresdText As String
Dim CellTex As String
Dim SpreadClon As New M_Plami2
Dim xColIni As Variant, xRowIni As Variant, xColFin As Variant, xRowFin As Variant
Dim CorDes As Long
'**** Samuel melendez ------------------------------------
Private Declare Function sendmessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wmsg As Long, _
    ByVal wparam As Long, lparam As Any) As Long
    Private Const EM_CANUNDO = &HC6
    Private Const EM_UNDO = &HC7
'***-------------------------------------------------------


'******* Id de Proceso SQL ***********************************************
'** Las siguientes variables RSSpid y Spid, sirven a
'** algunos procesos los cuelaes necesitan identificar el
'** turno de usuario que nos asigna SQL Server de manera unica.
'** estas variables solo deben ocuparse
'** para consultar el numero de proceso, el cual sera siempre el mismo
'** mientras no se cierre este formulario, asi mismo estas se destruyen
'** cuando es cerrado el formulario.
Dim RSSpid As New ADODB.Recordset
Dim spid As Long
'**----------------------------------------------------------------------
'************************************************************************
Enum DeshacerType
    AddFile = 1
    DelFile = 2
End Enum

Private Sub Check1_Click(Index As Integer)
HabilitaCol Index
End Sub

Private Sub Estructura1_Click(Index As Integer)
    LlenaSubMenu Estructura1, Index
End Sub

Sub LlenaSubMenu(SubMenu As Object, Index As Integer)
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1

DesqloqSubMenu vaSpread1.text
vaSpread1.text = SubMenu(Index).Caption

'DeshacerInto vaSpread1.Col, vaSpread1.Row, vaSpread1, DeshacerUltimoCammbio + 1

ActualizaEstructuraInferior vaSpread1, SubMenu(Index).Caption

vaSpread1.Col = vaSpread1.MaxCols: vaSpread1.text = SubMenu(Index).HelpContextID
'DeshacerInto vaSpread1.Col, vaSpread1.Row, vaSpread1, DeshacerUltimoCammbio


Estructura1.Item(Index).Enabled = False: Estructura2.Item(Index).Enabled = False

Plantilla(0).Enabled = True
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = True
End Sub

Private Sub Estructura2_Click(Index As Integer)
LlenaSubMenu Estructura2, Index
End Sub

Private Sub Form_Load()
Dim RS As New ADODB.Recordset
' **La variable "SpreadClon" contiene una copia de la grilla "vaSpread1"
' **tal como se inicio cuando se cargó el formulario
'Set SpreadClon = New vaSpread1

'*********-----> Identificacion que asigna el servidor SQL
        '** se mantiene mientras este abierto formulario
Set RSSpid = vg_db.Execute("Select @@Spid")
If Not (RSSpid.EOF And RSSpid.BOF) Then spid = RSSpid.Fields(0)
RSSpid.Close
Set RSSpid = Nothing

'********---->Validar minuta en uso <-

Me.HelpContextID = vg_OpcM
Me.Height = 6765
Me.Width = 11055
fg_centra Me
Msgtitulo = "Planificación Teórica"
fg_carga ""

' Ejecuta el timer cada 1 segundo
Timer1.Interval = 1000
vg_TemSeg = 0
CorDes = 0
Label4.Caption = M_Plami1.fpayuda(0).Caption & "(" & M_Plami1.fpLongInteger1(0).Value & ")" & " - " & M_Plami1.fpayuda(1).Caption & " - " & M_Plami1.fpayuda(2).Caption & " - " & " Tipo: " & IIf(vg_IndpprSelec = "1", "Real", "Propuesta") & " - Zona : " & Trim(Mid(M_Plami1.Combo2(0).text, 1, 150))
Label1(1).Caption = "Nota : Las raciones debe incluir las raciones del personal "
indcos = False
estapo = False
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): btnX.Visible = True: btnX.ToolTipText = " "
Set btnX = Toolbar1.Buttons.Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): btnX.Visible = False: btnX.ToolTipText = "Grabar Datos": btnX.Enabled = IIf(Mid(ValidarUsuario(M_Plami1), 2, 2) = "0", False, True)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Cortar", , tbrDefault, "A_Cortar"): btnX.Visible = True: btnX.ToolTipText = "Cortar"
Set btnX = Toolbar1.Buttons.Add(, "A_Copiar", , tbrDefault, "A_Copiar"): btnX.Visible = True: btnX.ToolTipText = "Copiar"
Set btnX = Toolbar1.Buttons.Add(, "I_Pegar", , tbrDefault, "I_Pegar"): btnX.Visible = True: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, "A_Pegar", , tbrDefault, "A_Pegar"): btnX.Visible = False: btnX.ToolTipText = "Pegar"
'Set btnX = Toolbar1.Buttons.Add(, "I_PegadoEspecial", , tbrDefault, "I_PegadoEspecial"): btnX.Visible = True: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, "I_PegadoEspecial", , tbrDefault, "I_PegadoEspecial"): btnX.Visible = True: btnX.ToolTipText = ""  ' Activé visiblemente esta opcion (True)  02/09/09 Samuel Melendez
Set btnX = Toolbar1.Buttons.Add(, "A_PegadoEspecial", , tbrDefault, "A_PegadoEspecial"): btnX.Visible = False: btnX.ToolTipText = "Pegado Especial"
Set btnX = Toolbar1.Buttons.Add(, "A_Buscar", , tbrDefault, "A_Buscar"): btnX.Visible = True: btnX.ToolTipText = "Buscar Recetas o Ingredientes"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_InsertarF", , tbrDefault, "A_InsertarF"): btnX.Visible = True: btnX.ToolTipText = "Insertar"
Set btnX = Toolbar1.Buttons.Add(, "A_EliminarF", , tbrDefault, "A_EliminarF"): btnX.Visible = True: btnX.ToolTipText = "Eliminar"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_SubirF", , tbrDefault, "A_SubirF"): btnX.Visible = True: btnX.ToolTipText = "Subir"
Set btnX = Toolbar1.Buttons.Add(, "A_BajarF", , tbrDefault, "A_BajarF"): btnX.Visible = True: btnX.ToolTipText = "Bajar"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_VerReceta", , tbrDefault, "A_VerReceta"): btnX.Visible = True: btnX.ToolTipText = "Ver Recetas"
Set btnX = Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): btnX.Visible = True: btnX.ToolTipText = "Copiar Planificación Teórica"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Aportes", , tbrDefault, "A_Aportes"): btnX.Visible = True: btnX.ToolTipText = "Aportes Nutricionales x Días"
Set btnX = Toolbar1.Buttons.Add(, "A_Costo", , tbrDefault, "A_Costo"): btnX.Visible = True: btnX.ToolTipText = "Visualizar Costo"
Set btnX = Toolbar1.Buttons.Add(, "A_Frecuencia", , tbrDefault, "A_Frecuencia"): btnX.Visible = True: btnX.ToolTipText = "Frecuencia Recetas"
Set btnX = Toolbar1.Buttons.Add(, "A_ExporReceta", , tbrDefault, "A_ExporReceta"): btnX.Visible = True: btnX.ToolTipText = "Exportar Recetas"
Set btnX = Toolbar1.Buttons.Add(, "A_ActCostoReceta", , tbrDefault, "A_ActCostoReceta"): btnX.Visible = True: btnX.ToolTipText = "Actualizar Planificación"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): btnX.Visible = False: btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): btnX.Visible = True: btnX.ToolTipText = "Exporta Planificación Minuta a Excel "
Set btnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): btnX.Visible = False: btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "Calorias", , tbrDefault, "Calorias"): btnX.Visible = True: btnX.ToolTipText = "Minuta con Calorias"
Set btnX = Toolbar1.Buttons.Add(, "Ingrediente", , tbrDefault, "Ingrediente"): btnX.Visible = True: btnX.ToolTipText = "Frecuencia de Ingrediente"
Set btnX = Toolbar1.Buttons.Add(, "A_Deshacer", , tbrDefault, "A_Deshacer"): btnX.Visible = True: btnX.ToolTipText = "Deshacer"
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
vg_ActCalorias = False
DetallePlantillaMinuta


'ActualizaAportesMinuta
'-------> Llena Sub Menu estructura

Toolbar1.Buttons(31).Enabled = False

'-------> Llena Sub Menu estructura

Dim x As Long

Set RS = vg_db.Execute("sgpadm_s_estservicio 1, " & vg_codservicio & ",''")

If Not RS.EOF Then

    x = 1

    Do While Not RS.EOF

        Load Estructura1(x): Load Estructura2(x)

        Estructura1(x).Caption = Trim(RS!ess_nombre): Estructura2(x).Caption = Trim(RS!ess_nombre)

        Estructura1(x).HelpContextID = RS!ess_codigo: Estructura2(x).HelpContextID = RS!ess_codigo

        Estructura1(x).Enabled = True: Estructura2(x).Enabled = True

        For i = 1 To vaSpread1.MaxRows

            vaSpread1.Col = vaSpread1.MaxCols: vaSpread1.Row = i

            If Trim(vaSpread1.text) <> "" Then

                If Val(vaSpread1.text) = RS!ess_codigo Then Estructura1(x).Enabled = False: Estructura2(x).Enabled = False

            End If

        Next

        x = x + 1

        RS.MoveNext

    Loop

End If

RS.Close: Set RS = Nothing

Estructura1(0).Visible = False: Estructura2(0).Visible = False

If Mid(ValidaPerfil(M_Plami2), 1, 4) = "1000" = True Then BlocSoloAcceso

'vg_db.Execute "sgpadm_iu_deshacer 3,  '" & vg_NUsr & "', " & spid & " "


End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then Label4.Move 0, 360, ScaleWidth, 435 'ScaleHeight - Toolbar1.Height
If Me.WindowState <> 1 Then Frame1.Move 0, 840, ScaleWidth, 675
If Me.WindowState <> 1 Then vaSpread1.Move 0, 1560, ScaleWidth, ScaleHeight - 1560
End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Man_Error
M_Plami1.DropTebleTmp (NameTemp)
If SwSalir <> 0 Then Exit Sub
If Toolbar1.Buttons(2).Visible = False Then Me.Hide: Unload Me: M_Plami1.WindowState = 0: Exit Sub
If MsgBox(" Actualiza planificación ...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Cancel = -1: Me.Hide: Unload Me: M_Plami1.WindowState = 0
If Toolbar1.Buttons(2).Visible = True And Cancel <> -1 Then GrabarPlantillaMinuta
Toolbar1.Buttons(1).Visible = True
Toolbar1.Buttons(2).Visible = False
SwSalir = 1
vg_PartePlani = False
Me.Hide
Unload Me
Set SpreadClon = Nothing
vg_db.Execute "sgpadm_iu_deshacer 3,  '" & vg_NUsr & "', " & spid & " "
M_Plami1.WindowState = 0
Man_Error:
End Sub

Private Sub Plantilla_Click(Index As Integer)
Dim StrRec As String, StrRecb As String
Dim j As Long, i As Long, codrec As Long, tiprec As Long
Dim cosali As Double, cosdes As Double
Dim desc As String
vg_RecetaReal = 0
estgra = False
Select Case Index
Case 0 '-------> Actualizar planificación
    If Toolbar1.Buttons(2).Enabled = False Then estgra = False: Exit Sub
    If MsgBox(" Actualiza planificación ...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Cancel = -1: estgra = False: Exit Sub
    If ValidaEstructuras = False Then MsgBox "No puede grabar, si exiten recetas sin ser asignadas a una estructura": Exit Sub
    If Toolbar1.Buttons(2).Visible = True Then
       Toolbar1.Enabled = False
       Toolbar1.Buttons(31).Enabled = False
       CorDes = 0
       GrabarPlantillaMinuta
       CorDes = 0
       If Dir(LCase(App.Path) & "\" & "Spread" & vg_NUsr & "*" & ".ss6") <> "" Then Kill LCase(App.Path) & "\" & "Spread" & vg_NUsr & "*" & ".ss6"
       Toolbar1.Enabled = True
    End If
    vg_TemSeg = 0
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = False
Case 3 '-------> Visualizar costo
    If Frame2(0).Visible = True Then Frame2(0).Visible = False: vaSpread1.Move 0, 1380, ScaleWidth, ScaleHeight - 1380: estgra = False: Exit Sub
    vaSpread1.Move 0, 1380, ScaleWidth, ScaleHeight - 4000
    Frame2(0).Move 0, ScaleHeight - 2600, ScaleWidth, ScaleHeight - 1200
    Frame2(0).Visible = True
    CargarCosto
Case 5 '-------> Visualizar receta
    Dim xcol As Integer, auxtiprec  As Long
    vaSpread1.Row = vaSpread1.ActiveRow ': cand = vaSpread1.text
    vaSpread1.Col = vaSpread1.ActiveCol: desc = vaSpread1.text
    If vaSpread1.Col = 1 Then estgra = False: Exit Sub
'    If vaSpread1.BackColor = Shape1(1).FillColor Then vg_newestrec = True Else vg_newestrec = False
    vg_newestrec = True
    vg_modreceta = True
'    If vaSpread1.BackColor = Shape1(1).FillColor Then vg_modreceta = True Else vg_modreceta = False
    xcol = 0
    For i = 1 To maxColumna
        If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or vectorcol(i) = (vaSpread1.Col - 2) Or vectorcol(i) = (vaSpread1.Col - 4)) And Trim(vaSpread1.text) <> "" Then xcol = vectorcol(i): Exit For
    Next i
    If xcol = 0 Then MsgBox "No existe receta ha vizualizar", vbCritical + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
    If vg_newestrec = True Then
       vg_fecval = 0: vg_fecval = Val(vg_fecha) & Right("0" & (Int(xcol / 6) + 1), 2)
       Set RS = vg_db.Execute("sgpadm_s_planifminuta 3, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & "," & vg_Zona & ", " & vg_fecval & ", 0, 0,'" & vg_IndpprSelec & "'")
       If Not RS.EOF Then vg_fecval = RS!mid_fecval: vg_opcion = 2
       RS.Close: Set RS = Nothing
    End If
    vaSpread1.Col = xcol '+ 3
    vaSpread1.Row = 0
    If vaSpread1.text = "R" Then
      vaSpread1.Col = vaSpread1.ActiveCol + 1: desc = vaSpread1.text
      vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = xcol + 4
      StrRec = vaSpread1.text
    ElseIf vaSpread1.text = "N.Rac." Then
      vaSpread1.Col = vaSpread1.ActiveCol - 1: desc = vaSpread1.text
      vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = xcol + 2
      StrRec = vaSpread1.text
    ElseIf vaSpread1.text = "Costo" Then
      vaSpread1.Col = vaSpread1.ActiveCol - 2: desc = vaSpread1.text
      vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = xcol + 1
      StrRec = vaSpread1.text
    ElseIf vaSpread1.text = "Calorias" Then
      vaSpread1.Col = vaSpread1.ActiveCol - 1: desc = vaSpread1.text
      vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = xcol - 1
      StrRec = vaSpread1.text
    Else
      vaSpread1.Row = vaSpread1.ActiveRow: desc = vaSpread1.text
      vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = xcol + 3
      StrRec = vaSpread1.text
    End If

    If Len(StrRec) <> 0 Then
       Do While InStr(StrRec, ";") <> 0
          StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
          StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
          vg_newcodrec = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)): StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
          vg_newcodrec = IIf(vg_newcodrec = 0, BuscarCodReceta(desc), vg_newcodrec)
          vg_tiprec = Val(Mid(StrRecb, 1))
          vg_PartePlani = True
       Loop
    End If
'    vg_newnomrec = ""
    auxtiprec = vg_tiprec
    Dim Receta As New M_Receta
    vg_RecetaReal = 1
    Receta.Show 1, Me
    Set Receta = Nothing

    vg_newestrec = False
    If vg_newcodrec <> 0 And Trim(vg_newnomrec) <> "" And vaSpread1.BackColor <> Shape1(1).FillColor And auxtiprec = vg_tiprec Then
        vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = vaSpread1.ActiveCol
        vaSpread1.Col = xcol + 3
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = xcol
        '-------> Limpiar Datos y Formato Celda
        vaSpread1.Action = 3
        '-------> Retorna Modo de la columna
        vaSpread1.BlockMode = False
        vaSpread1.Font.Bold = False
        vaSpread1.Font.Size = 8
        vaSpread1.text = vg_newnomrec
        
        vaSpread1.Col = xcol + 2
        vaSpread1.CellType = 5
        vaSpread1.TypeHAlign = 1
        '-------> Calcular costo alimentación y deshechable
        cosali = Format(fg_CalCtoRecListaPrecio(Val(vg_newcodrec), vg_tiprec, vg_codlpr, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")), Val(vg_fecha)), fg_Pict(6, 2))
        cosdes = Format(fg_CalCtoRecListaPrecio(Val(vg_newcodrec), vg_tiprec, vg_codlpr, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")), Val(vg_fecha)), fg_Pict(6, 2))
        vaSpread1.text = Format((cosali + cosdes), fg_Pict(6, 2))
        
        vaSpread1.Col = xcol + 3
        vaSpread1.text = vg_newcodrec & "&" & vg_tiprec & "&;"
        
        '-------> Revizar si existe receta iguales en el mes y actualizar
        For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 5
            vaSpread1.Col = i + 4
            For j = 1 To (vaSpread1.MaxRows - 1)
                vaSpread1.Row = j: codrec = 0
                vaSpread1.Col = i + 1
                If vaSpread1.BackColor = Shape1(1).FillColor Then Exit For
                vaSpread1.Col = i + 4
                If Trim(vaSpread1.text) <> "" Then
                   StrRec = vaSpread1.text
                   If Len(StrRec) <> 0 Then
                      Do While InStr(StrRec, ";") <> 0
                         StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
                         StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
                         codrec = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)): StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
                         tiprec = Val(Mid(StrRecb, 1))
                      Loop
                   End If
                   If codrec = vg_newcodrec Then
                      vaSpread1.Col = i + 4
                      vaSpread1.text = vg_newcodrec & "&" & vg_tiprec & "&;"
                      
                      vaSpread1.Col = i + 3
                      vaSpread1.CellType = 5
                      vaSpread1.TypeHAlign = 1
                      vaSpread1.text = Format((cosali + cosdes), fg_Pict(6, 2))
                   
'                      vaSpread1.Col = (maxColumna * 6 + 1) + ((i + 3) / 6)
'                      If Trim(vaSpread1.text) = "" And vaSpread1.Col < vaSpread1.MaxCols Then vaSpread1.text = 1

                   End If
                End If
            Next j
        Next i
        If indcos = True Then
           For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
               Calctodia 1, i + 1
           Next i
           MostrarCosto vaSpread1.ActiveCol
        End If
        '-------> Actualizar lista receta
        If B_Receta.vaSpread1.MaxRows > 0 Then
            B_Receta.vaSpread1.Row = B_Receta.vaSpread1.SearchCol(1, -1, B_Receta.vaSpread1.MaxRows, Val(vg_newcodrec), SearchFlagsEqual)
            B_Receta.vaSpread1.Col = 3: B_Receta.vaSpread1.text = Format((cosali + cosdes), fg_Pict(6, 2))
        End If
        vg_newcodrec = 0: vg_newnomrec = "": vg_tiprec = -1
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
    End If
    If indcos = True Then Me.Refresh: Toolbar1.Refresh: Frame2(0).Refresh: Frame2(1).Refresh: Frame2(2).Refresh: Frame2(3).Refresh: Frame2(4).Refresh
    vg_newcodrec = 0
Case 8 '-------> Copiar planificación
    M_CPlaTe.Show 1, Me
Case 10 '-------> Visualizar aportes x día
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col = 1 Then estgra = False: Exit Sub
    j = 0
    For i = 1 To maxColumna
        If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or vectorcol(i) = (vaSpread1.Col - 2)) Then j = vectorcol(i): Exit For
    Next i
    vaSpread1.Col = j: vaSpread1.Row = 0
    C_ApoPla.LlenarApoPlan Me, "Aporte Planificación Real " & vaSpread1.text, vg_codsubseg, vg_codregimen, vg_codservicio, Val(vg_fecha), 2, j
    C_ApoPla.Show 1, Me
Case 11 '-------> Visualizar frecuencia de recetas
    If Toolbar1.Buttons(2).Visible = True Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
    C_FrePla.LlenarFrecPlan "Frecuencia Planificación " & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4), vg_codsubseg, vg_codregimen, vg_codservicio, Val(vg_fecha), 1
    C_FrePla.Show 1, Me
Case 13 '-------> Actualizar costo recetas y planificación
    If indgrabado = 1 Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
    fg_carga ""
    '-------> Rutina actualizar precio planificación
'    vg_db.Execute "sgpadm_p_actualizarplanif " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & vg_codlpr & ", " & Val(vg_fecha) & ", 0, '" & vg_NUsr & "'"
    vg_db.Execute "sgpadm_p_actuaplanif " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & vg_codlpr & ", " & Val(vg_fecha) & ""
    Dim vecactrec As Variant
    '-------> Traer total de receta desde planificación de minutas y luego calcular costo
    Set RS = vg_db.Execute("sgpadm_s_planifminuta 11, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & vg_Zona & "," & Val(vg_fecha) & ", 0,0,'" & vg_IndpprSelec & "'")
    If RS.EOF Or RS!nReg < 1 Then RS.Close: Set RS = Nothing: MsgBox "No existe información", vbInformation + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
    ReDim vecactrec(RS!nReg, 4)
    RS.Close: Set RS = Nothing
    For i = 1 To UBound(vecactrec)
        DoEvents
        vecactrec(i, 1) = 0 '-------> codigo receta
        vecactrec(i, 2) = 0 '-------> tipo receta
        vecactrec(i, 3) = 0 '-------> costo receta alimentación
        vecactrec(i, 4) = 0 '-------> costo receta desechable
    Next i
    i = 1
    Dim inddia As Long
    gauge1.Value = 0: gauge.Value = 0: fecha = 0: inddia = 1: fecha = 0: cosali = 0: cosdes = 0
    Picture1.Visible = True: Label2.Visible = False: Label3.Visible = True: Label3.Caption = "Recopilando información, un momento....": gauge.Visible = True: gauge.Visible = False
    Set RS = vg_db.Execute("sgpadm_s_planifminuta 12, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & vg_Zona & " , " & Val(vg_fecha) & ", 0,0,'" & vg_IndpprSelec & "' ")
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información", vbInformation + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
    Do While Not RS.EOF
       DoEvents
       vecactrec(i, 1) = RS!mid_codrec
       vecactrec(i, 2) = RS!mid_tiprec
       vecactrec(i, 3) = Format(fg_CalCtoRecListaPrecio(Val(RS!mid_codrec), RS!mid_tiprec, vg_codlpr, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")), Val(vg_fecha)), fg_Pict(6, 2))       'Format(IIf(IsNull(RS!mid_cosrec), 0, RS!mid_cosrec), fg_Pict(6, 2))
       vecactrec(i, 4) = Format(IIf(IsNull(RS!mid_cosdes), 0, RS!mid_cosdes), fg_Pict(6, 2))
       
       RS.MoveNext: i = i + 1
    Loop
    RS.Close: Set RS = Nothing
    
    gauge1.Value = 0: gauge.Value = 0: fecha = 0: inddia = 1: fecha = 0: cosali = 0: cosdes = 0
    Picture1.Visible = True: Label2.Visible = False: Label3.Visible = True: Label3.Caption = "Actualizando costo receta, en planificación": gauge.Visible = True: gauge.Visible = False
    For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
        DoEvents
        gauge1.Value = Val((i / vaSpread1.MaxCols) * 100)
        existedat = 0
        vaSpread1.Row = 1: vaSpread1.Col = i
        fecha = Val(vg_fecha) & fg_pone_cero(inddia, 2)
        If vaSpread1.BackColor <> Shape1(1).FillColor Then
           For j = 1 To (vaSpread1.MaxRows - 1)
               vaSpread1.Row = j
               vaSpread1.Col = i + 1
               If Trim(vaSpread1.text) <> "" Then existedat = 1: Exit For
           Next j
           If existedat > 0 Then
              For j = 1 To (vaSpread1.MaxRows - 1)
                  vaSpread1.Row = j: vaSpread1.Col = i + 1: codrec = 0
                  If Trim(vaSpread1.text) <> "" Then
                    vaSpread1.Col = i + 4: StrRec = vaSpread1.text
                    If Len(StrRec) <> 0 Then
                       Do While InStr(StrRec, ";") <> 0
                          StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
                          StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
                          codrec = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)): StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
                          tiprec = Val(Mid(StrRecb, 1))
                       Loop
                    End If
                    vaSpread1.Col = i + 3
                    '-------> Traer costo alimentación y desechables
                    For x = 1 To UBound(vecactrec)
                        If codrec = vecactrec(x, 1) And tiprec = vecactrec(x, 2) Then
                           cosali = vecactrec(x, 3)
                           cosdes = vecactrec(x, 4)
                           Exit For
                        End If
                    Next
                    vaSpread1.text = Format((cosali + cosdes), fg_Pict(6, 2))
'                    vaSpread1.Col = (maxColumna * 6 + 1) + ((i + 3) / 6)
'                    If Trim(vaSpread1.text) = "" And vaSpread1.Col < vaSpread1.MaxCols Then vaSpread1.text = 2
                    indgrabado = 1
                  End If
              Next j
           End If
        End If
        inddia = inddia + 1
    Next i
    Label2.Visible = True: Picture1.Visible = False: gauge.Visible = False
    vaSpread1.Refresh
    If indgrabado = 1 Then fg_descarga: MsgBox "Actualización costo receta finalizado sin problema, luego grabe información", vbInformation + vbOKOnly, Msgtitulo: Plantilla(0).Enabled = True: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = True: estgra = False: Exit Sub
    fg_descarga
Case 14 '-------> Visualizar detalle de recetas
    If Toolbar1.Buttons(2).Visible = True Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
    C_ExpRec.LlenarExporReceta "Exportar Recetas " & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4), vg_codsubseg, vg_codregimen, vg_codservicio, Val(vg_fecha)
    C_ExpRec.Show 1, Me
Case 16 '-------> Parámetro de grabado
    M_ParGra.Show 1, Me
'    Timer1.Interval = 60000 ' corresponde a un minuto
Case 20
    If Toolbar1.Buttons(2).Visible = True Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
    C_IngPla.LlenarFrecIng "Frecuencia Planificación Ingrediente " & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4), vg_codsubseg, vg_codregimen, vg_codservicio, Val(vg_fecha), 1
    C_IngPla.Show 1, Me
Case 21
'    Combo1.ListIndex = Combo1.ListCount - 1
    Deshacer "Spread" & vg_NUsr & CorDes & ".ss6" 'Combo1.Text
    If CorDes < 1 Then: Toolbar1.Buttons(31).Visible = True: Toolbar1.Buttons(31).Enabled = False
'    If (Combo1.ListCount - 1) > -1 Then Combo1.RemoveItem Combo1.ListCount - 1
'    If (Combo1.ListCount - 1) < 0 Then: Toolbar1.Buttons(31).Visible = True: Toolbar1.Buttons(31).Enabled = False
'    Toolbar1.Buttons(31).Visible = True
'    DeshacerSelect
Case 22 '-------> Salir
    vg_PartePlani = False
    SwSalir = 0
    If Toolbar1.Buttons(2).Visible = False Then SwSalir = 1: Me.Hide: Unload Me: M_Plami1.WindowState = 0: estgra = False: Exit Sub
    If MsgBox(" Actualiza planificación ...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then SwSalir = 1: Me.Hide: Unload Me: M_Plami1.WindowState = 0
    If Toolbar1.Buttons(2).Visible = True Then GrabarPlantillaMinuta
    SwSalir = 1: Me.Hide: Unload Me: M_Plami1.WindowState = 0
End Select
estgra = False
End Sub

Private Sub Plato_Click(Index As Integer)
On Error GoTo Man_Error
If Toolbar1.Buttons(2).Enabled = False Then estgra = False: Exit Sub
Dim Del_Row As Integer, indcol As Integer, indrow As Integer, indcol2 As Integer, indrow2 As Integer, indrow3 As Long, xx As Long
Dim Col As Long, fil As Long, codest As Long, cosali As Double, cosdes As Double
Dim VecSelGrid As Variant: Dim VecRacPegar As Variant
Dim contador, contador_b, cantCol As Integer, LargoVec As Integer
Dim accion As String
Dim ColumnaActiva, FilaActiva, ColumnaAntActiva, n, n1, NFilas As Integer
contador = 0: contador_b = 0: cantCol = 0: LargoVec = 0:  accion = "": n1 = 0: n = 0: NFilas = 0
estgra = True
Select Case Index
Case 2 '-------> Ingresa recetas
    If vaSpread1.ActiveRow = vaSpread1.MaxRows Then Exit Sub
    iblockcol = vaSpread1.ActiveCol: aiblockcol = vaSpread1.ActiveCol
    iblockcol2 = vaSpread1.ActiveCol: aiblockcol2 = vaSpread1.ActiveCol
    iblockrow = vaSpread1.ActiveRow: aiblockrow = vaSpread1.ActiveRow
    iblockrow2 = vaSpread1.ActiveRow: aiblockrow2 = vaSpread1.ActiveRow
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    vaSpread1.Col = vaSpread1.ActiveCol + 4: vaSpread1.Row = 0:
    If vaSpread1.text = "Calorias" Then
       If vaSpread1.ColHidden = False Then vg_ActCalorias = True Else vg_ActCalorias = False
    End If
    
    vg_RecetaReal = 1
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
    
'    If vg_IndpprSelec = 1 And vaSpread1.Col = 1 Then
'        vaSpread1.Row = vaSpread1.ActiveRow
'        vaSpread1.Col = vaSpread1.ActiveCol
'        vaSpread1.CellType = CellTypeEdit
'    End If -- samuel
    
    
    j = 0
    For i = 1 To maxColumna
        If vaSpread1.Col = vectorcol(i) Then j = vectorcol(i): Exit For
    Next i
    If j = 0 Then estgra = False: Exit Sub
    vg_codigo = "": vg_nombre = "": vg_tiprec = -1
    vaSpread1.Row = vaSpread1.ActiveRow
    B_Receta.vaSpread1.Col = 6
    If vg_ActCalorias = True Then
       B_Receta.vaSpread1.ColHidden = False
    Else
       B_Receta.vaSpread1.ColHidden = True
    End If
    B_Receta.Show 1, Me
    
    vg_RecetaReal = 0
    B_Receta.vaSpread1.Col = 6
    If vg_ActCalorias = True Then
       B_Receta.vaSpread1.ColHidden = False
    Else
       B_Receta.vaSpread1.ColHidden = True
    End If

    If Trim(vg_codigo) = "" Or Trim(vg_nombre) = "" Then estgra = False: Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = j - 1
    vaSpread1.CellType = 5
    vaSpread1.TypeHAlign = 2
    vaSpread1.Value = "R"
    vaSpread1.ForeColor = &HFF&
    vaSpread1.BackColor = &H80FF80
    
    'DeshacerInto vaSpread1.Col, vaSpread1.Row, vaSpread1, DeshacerUltimoCammbio + 1
    GrabarCambios 1, 1, ""
    
    vaSpread1.Col = j
    '-------> Limpiar Datos y Formato Celda
    vaSpread1.Action = 3
    '-------> Retorna Modo de la columna
    vaSpread1.BlockMode = False
    vaSpread1.Font.Bold = False
    vaSpread1.Font.Size = 8

    vaSpread1.text = vg_nombre
    
    'DeshacerInto vaSpread1.Col, vaSpread1.Row, vaSpread1, DeshacerUltimoCammbio

    
    vaSpread1.Col = j + 1
    If Trim(vaSpread1.text) = "" Then
       '-------> Asignar raciones estimadas
'       vaSpread1.Row = 0: vaSpread1.col = j
'       RS.Open "select sra_serdia, sum(sra_raciones) as sra_raciones from a_serviciorac where sra_codser=" & vg_codservicio & " and sra_serdia=" & IIf(fg_Dia(Format(CDate(Mid(Trim(vaSpread1.Text), 5, Len(Trim(vaSpread1.Text)))), "yyyymmdd")) = 1, 7, (fg_Dia(Format(CDate(Mid(Trim(vaSpread1.Text), 5, Len(Trim(vaSpread1.Text)))), "yyyymmdd")) - 1)) & " group by sra_serdia", vg_db, adOpenStatic
       codest = 0
       vaSpread1.Row = vaSpread1.ActiveRow
       For i = (IIf(vaSpread1.Row = 1, 1, vaSpread1.Row + 1 - 1)) To 1 Step -1
           vaSpread1.Row = i
           vaSpread1.Col = 1
           If Trim(vaSpread1.text) <> "" Then vaSpread1.Col = vaSpread1.MaxCols: codest = Val(vaSpread1.text): Exit For
       Next i
       Set RS = vg_db.Execute("SELECT * FROM a_estservicio WHERE ess_codser=" & vg_codservicio & " AND ess_codigo=" & codest & "")
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = j + 1
       vaSpread1.CellType = 3
       vaSpread1.TypeIntegerMin = 1
       vaSpread1.TypeIntegerMax = 9999999
       vaSpread1.TypeHAlign = 1
       vaSpread1.TypeSpin = False
       vaSpread1.TypeIntegerSpinInc = 1
       vaSpread1.TypeIntegerSpinWrap = False
       vaSpread1.text = IIf(RS.EOF, 0, RS!ess_racmin)
       vaSpread1.ForeColor = &HFF0000
       RS.Close: Set RS = Nothing
       'DeshacerInto vaSpread1.Col, vaSpread1.Row, vaSpread1, DeshacerUltimoCammbio

    End If
    
    vaSpread1.Col = j + 2
    vaSpread1.CellType = 5
    vaSpread1.TypeHAlign = 1
    '------> Calcular costo planificación alimento y desechable
    cosali = 0: cosdes = 0
    cosali = Format(fg_CalCtoRecListaPrecio(Val(vg_codigo), vg_tiprec, vg_codlpr, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")), Val(vg_fecha)), fg_Pict(6, 2))
    cosdes = Format(fg_CalCtoRecListaPrecio(Val(vg_codigo), vg_tiprec, vg_codlpr, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")), Val(vg_fecha)), fg_Pict(6, 2))
    vaSpread1.text = Format((cosali + cosdes), fg_Pict(6, 2))
    
    'DeshacerInto vaSpread1.Col, vaSpread1.Row, vaSpread1, DeshacerUltimoCammbio
    
    vaSpread1.Col = j + 3
'    If Trim(vaSpread1.text) <> Val(vg_codigo) And ((maxColumna * 6 + 1) + ((j + 2) / 6)) < vaSpread1.MaxCols Then
'       vaSpread1.Col = (maxColumna * 6 + 1) + ((j + 2) / 6)
'       vaSpread1.text = 1
'       ': If indcos = True Then vaSpread1.col = J + 2: veccos((Int(J / 5) + 1), 1) = Round(veccos((Int(J / 5) + 1), 1) - vaSpread1.Text, vg_DCa)
'    End If
    
    
    vaSpread1.Col = vaSpread1.MaxCols - 1
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.text = EstructuraSuperior(vaSpread1, vaSpread1.Row)

    vaSpread1.Col = j + 3
    vaSpread1.text = Val(vg_codigo) & "&" & vg_tiprec & "&;"
    'DeshacerInto vaSpread1.Col, vaSpread1.Row, vaSpread1, DeshacerUltimoCammbio

    If indcos = True Then Calctodia vaSpread1.Row, j
    
    RS1.Open "sgpadm_s_AporteNutricionales 2," & IIf(codrec = 0, BuscarCodReceta(vg_nombre), codrec) & "," & vg_codsubseg & "," & vg_codregimen & ", " & vg_Zona & "", vg_db, adOpenForwardOnly ', adOpenStatic
    If Not RS1.EOF Then
      vg_Calorias = RS1!candiet
    End If
    RS1.Close: Set RS1 = Nothing
    
    vaSpread1.Col = j + 4
    vaSpread1.Row = 0
    If vaSpread1.text = "Calorias" Then
      vaSpread1.Col = j + 4
      vaSpread1.Row = vaSpread1.ActiveRow
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.text = Format(vg_Calorias, fg_Pict(9, 2))
    'DeshacerInto vaSpread1.Col, vaSpread1.Row, vaSpread1, DeshacerUltimoCammbio

    End If
    
'    vaSpread1.Row = vaSpread1.ActiveRow
    
    If Mid(ValidaPerfil(M_Plami2), 1, 4) = "1000" = True Then
    
        BlocSoloAcceso
    Else
        vaSpread1.Row = iblockrow
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
    End If
Case 5 '-------> Insertar linea
    ' se agregó esta asignacion a estas variables, para indicarle la seleccion de las celdas
    vaSpread1.GetSelection 1, xColIni, xRowIni, xColFin, xRowFin
    GrabarCambios Val(xRowIni), Val(xRowFin), "Insertar"
    vaSpread1.Enabled = False
    indcol = iblockcol
    iblockcol = 1: iblockcol2 = vaSpread1.MaxCols
    vaSpread1.MaxRows = vaSpread1.MaxRows + ((xRowFin - xRowIni) + 1) '1
    vaSpread1.InsertRows xRowIni, ((xRowFin - xRowIni) + 1)
    'DeshacerModFile xRowIni, ((xRowFin - xRowIni) + 1), DeshacerUltimoCammbio + 1, AddFile
    
    
    If vg_IndpprSelec <> "2" Then
       For i = 3 To (vaSpread1.MaxCols - maxColumna) Step 6
           vaSpread1.Row = 0: vaSpread1.Col = i
           If CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
              Dim f As Long, c As Long
              For c = i - 1 To i + 2
                  vaSpread1.Row = xRowIni: vaSpread1.Col = c
                  vaSpread1.BackColor = Shape1(1).FillColor
              Next c
           End If
       Next i
    End If
    '-------> Validar días modificados
'    For j = iblockrow To (vaSpread1.MaxRows - 1) '((vaSpread1.MaxRows - 1) - ((iblockrow2 - iblockrow) + 1))
'        For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
'            vaSpread1.Row = j
''            vaSpread1.Col = i + 1
''            If Trim(vaSpread1.Text) <> "" Then
'               vaSpread1.Col = (maxColumna * 6 + 1) + ((i + 3) / 6)
'               If Trim(vaSpread1.text) = "" And vaSpread1.Col < vaSpread1.MaxCols Then vaSpread1.text = 1
''            End If
'        Next i
'    Next j
    '-------> Fin validar días modificados
    'For i = 3 To vaSpread1.MaxCols Step 5
    '    vaSpread1.Row = 0: vaSpread1.Col = i
    '    If InStr(1, Trim(vaSpread1.Text), "Día " & Format(Date, "d/mm/yyyy")) = 1 Then
    '        For Col = 0 To i - 4
    '            vaSpread1.Row = iblockrow: vaSpread1.Col = Col + 2
    '            vaSpread1.BackColor = Shape1(1).FillColor
    '        Next Col
    '    End If
    'Next i
    iblockcol = indcol
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    fg_descarga
    vaSpread1.Enabled = True
Case 6 '-------> Eliminar línea
    vaSpread1.Enabled = False
    Dim x_iblockrow As Variant, x_iblockrow2 As Variant, x_iblockcol As Variant, x_iblockcol2 As Variant
    
    vaSpread1.GetSelection 1, xColIni, xRowIni, xColFin, xRowFin
    
    '-- INICIO -- ULTIMA FILA
        'Saca de la seleccion la ultima fila cuando se encuantra seleccionada por el usuario para borrar
      If xRowFin = vaSpread1.MaxRows Then xRowFin = xRowFin - 1
    '-- FIN --  ULTIMA FILA
  
    
    ' se agregó esta asignacion a estas variablea, las cuales corresponden
    ' al rango de celda seleccionado, ya que como se estaban asignando
    ' a veces se producia inconsistencias en la asignacion devolviendo rangos malos
    iblockrow = xRowIni
    iblockrow2 = xRowFin
    iblockcol = xColIni
    iblockcol2 = xColFin
    '*********************------------------
    
    fg_carga ""
    indcol = iblockcol
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    'If indactivo = 0 Then iblockcol = vaSpread1.ActiveCol: iblockcol2 = vaSpread1.ActiveCol: iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    If iblockcol < 0 Then iblockcol = 2: iblockcol2 = vaSpread1.MaxCols
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    

    
    NFilas = (aiblockrow2 - aiblockrow) + 1
    

    
    If NFilas > 1 Then
        'debido a que las siguientes variable cambian de valor, al salir el mensaje
        ' al usuario, debido a que se ejecuta un evento que las cambia al
        ' perder el foco, aqui se intenta rescatar su valor en variables de paso
        ' para una vez mostrado el mensaje, se les devuelva su valor anterior

    
        If MsgBox("Cuando se intenta eliminar mas de una fila, no es posible recuperar la informacion contenida en ella mediante la opcion deshacer  ¿Desea Continuar? ", vbInformation + vbYesNo) = vbNo Then
                fg_descarga
                vaSpread1.Enabled = True
                Exit Sub
        End If
        
        For i = xRowIni To xRowFin
            vaSpread1.Col = 1
            vaSpread1.Row = i
            DesqloqSubMenu (vaSpread1.text)
        Next i
        'aqui se recupera su valor anterior
        iblockrow = xRowIni
        iblockrow2 = xRowFin
        iblockcol = xColIni
        iblockcol2 = xColFin
    End If
    
    If vaSpread1.BackColor = Shape1(1).FillColor And Trim(vaSpread1.text) <> "" Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, Msgtitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
    If vaSpread1.BackColor = Shape1(1).FillColor Or Trim(vaSpread1.text) = "" Then GoTo Paso
'    j = 0
'    For i = 1 To maxColumna
'        If (vectorcol(i) - 1) = vaSpread1.Col Or vectorcol(i) = vaSpread1.Col Then j = (vectorcol(i) - 1): Exit For
'    Next i
'    If j = 0 Then estgra = False:Exit Sub
    
    'Validación Dias Bloqueados
    If vg_IndpprSelec <> "2" Then
       If iblockcol < 1 Or (iblockcol = 1 And Trim(vaSpread1.text) <> "") Then
          For i = 1 To maxColumna
              vaSpread1.Col = vectorcol(i)
              vaSpread1.Row = 1
              If vaSpread1.BackColor = Shape1(1).FillColor And Shape1(1).FillColor = &H8080FF Then MsgBox "Existen Días Bloqueado, utilizar opción Suprimir para eliminar recetas.", vbCritical + vbOKOnly, Msgtitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
          Next i
       Else
          For i = iblockcol To iblockcol2
              vaSpread1.Col = i
              For j = iblockrow To iblockrow2
                 vaSpread1.Row = j
                 If vaSpread1.BackColor = Shape1(1).FillColor And Shape1(1).FillColor = &H8080FF Then MsgBox "Bloque seleccionado existen días bloqueado, utilizar opción Suprimir para eliminar recetas.", vbCritical + vbOKOnly, Msgtitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
              Next j
          Next i
       End If
    End If
    For i = 1 To maxColumna
        If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
    Next i
    For i = 1 To maxColumna
        If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
        If vectorcol(i) = iblockcol2 Then iblockcol2 = (vectorcol(i) + 3): Exit For
    Next i
    indcol = aiblockcol: indcol2 = iblockcol2
    indrow = aiblockrow: indrow2 = aiblockrow2
    If indcos = True Then
       For i = iblockcol To iblockcol2 Step 6
           Calctodia 1, i + 1
       Next i
       MostrarCosto vaSpread1.ActiveCol
      
    End If
    '-------> Validar días modificados
'    For j = iblockrow To ((vaSpread1.MaxRows - 1) - ((iblockrow2 - iblockrow) + 1) + 1)
'        For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
'            vaSpread1.Row = j
'            vaSpread1.Col = i + 1
'            If Trim(vaSpread1.text) <> "" Then
'               vaSpread1.Col = (maxColumna * 6 + 1) + ((i + 3) / 6)
'               If Trim(vaSpread1.text) = "" And vaSpread1.Col < vaSpread1.MaxCols Then vaSpread1.text = 1
'            End If
'        Next i
'    Next j
    '-------> Fin validar días modificados
    iblockcol = auxcol
    vaSpread1.BlockMode = False
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    indactivo = 0
Paso:
    vaSpread1.Row = vaSpread1.ActiveRow
    
    'Validación Dias Bloqueados
    If vg_IndpprSelec <> "2" Then
       If iblockcol < 1 Or (iblockcol = 1 And Trim(vaSpread1.text) <> "") Then
          For i = 1 To maxColumna
              vaSpread1.Col = vectorcol(i)
              vaSpread1.Row = 1
              If vaSpread1.BackColor = Shape1(1).FillColor And Shape1(1).FillColor = &H8080FF Then MsgBox "Existen Días Bloqueado, utilizar opción Suprimir para eliminar recetas.", vbCritical + vbOKOnly, Msgtitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
          Next i
       Else
          For i = iblockcol To iblockcol2
              vaSpread1.Col = i
              For j = iblockrow To iblockrow2
                  vaSpread1.Row = j
                  If vaSpread1.BackColor = Shape1(1).FillColor And Shape1(1).FillColor = &H8080FF Then MsgBox "Bloque seleccionado existen días bloqueado,utilizar opción Suprimir para eliminar recetas.", vbCritical + vbOKOnly, Msgtitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
              Next j
          Next i
       End If
    End If
    'For i = 1 To vaSpread1.MaxCols
    '    vaSpread1.Col = i
    '    If Trim(vaSpread1.text) <> "" Then MsgBox "Existe mas información en la linea, no puede eliminarla completamente", vbCritical + vbOKOnly, Msgtitulo: estgra = False:Exit Sub
    'Next i
    
    vaSpread1.Row = iblockrow2
    SpreadClon.vaSpread1.Row = iblockrow2
    vaSpread1.Col = iblockcol
    SpreadClon.vaSpread1.Col = iblockcol
    'vaSpread1.Visible = False
    
    
    GrabarCambios Val(iblockrow), Val(NFilas), "Eliminar"
   
'Esta función pega datos de una grila a otra
'    'Single Block Selected
'    Dim array1, array2 As Long
'    'Get the size of the block
'    array1 = vaSpread1.SelBlockRow2 - vaSpread1.SelBlockRow
'    array2 = vaSpread1.SelBlockCol2 - vaSpread1.SelBlockCol
'    'Init array size
'    ReDim fparray(array1, array2) As Variant
'    'Get data: ColLeft, RowTop
'    vaSpread1.GetArray vaSpread1.SelBlockCol, vaSpread1.SelBlockRow, fparray
'    'Display the selected data
'    vaSpread2.SetArray 1, 1, fparray
    
    
    'If NFilas = 1 Then DeshacerDelFile iblockrow, NFilas
    vaSpread1.DeleteRows iblockrow, NFilas
    SpreadClon.vaSpread1.DeleteRows iblockrow, NFilas
    vaSpread1.MaxRows = vaSpread1.MaxRows - NFilas
    SpreadClon.vaSpread1.MaxRows = vaSpread1.MaxRows - NFilas
    'DeshacerModFile iblockrow, NFilas, DeshacerUltimoCammbio, DelFile
    vaSpread1.Col = vaSpread1.Row
    vaSpread1.Visible = True
    SpreadClon.vaSpread1.Visible = True
    
    '-------> Validar días modificados
'    For j = iblockrow To vaSpread1.MaxRows - 1
'        For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
'            vaSpread1.Row = j
'            vaSpread1.Col = i + 1
'            If Trim(vaSpread1.text) <> "" Then
'               vaSpread1.Col = (maxColumna * 6 + 1) + ((i + 3) / 6)
'               If Trim(vaSpread1.text) = "" And vaSpread1.Col < vaSpread1.MaxCols Then
'                  vaSpread1.text = 1
'               End If
'            End If
'        Next i
'    Next j
    '-------> Fin validar días modificados
    
    iblockcol = indcol
    If vg_IndpprSelec <> "2" Then
       For i = 3 To (vaSpread1.MaxCols - maxColumna) Step 6
           vaSpread1.Row = 0: vaSpread1.Col = i
           If CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
               For Col = 0 To i - 4
                   vaSpread1.Row = (vaSpread1.MaxRows - 1): vaSpread1.Col = Col + 2
                   vaSpread1.BackColor = Shape1(1).FillColor
               Next Col
           End If
       Next i
    End If
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    fg_descarga
    vaSpread1.Enabled = True
Case 8 '-------> Subir línea
    vaSpread1.Enabled = False
    
    vaSpread1.GetSelection 1, xColIni, xRowIni, xColFin, xRowFin
    iblockrow = xRowIni
    iblockrow2 = xRowFin
    iblockcol = xColIni
    iblockcol2 = xColFin
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Row = 1 Or vaSpread1.Row = vaSpread1.MaxRows Then estgra = False: vaSpread1.Enabled = True: Exit Sub
    If iblockcol < 1 Or (iblockcol = 1 And Trim(vaSpread1.text) <> "") Then
       For i = 1 To maxColumna
           vaSpread1.Col = vectorcol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, Msgtitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
       Next i
    Else
       For i = iblockcol To iblockcol2
           vaSpread1.Col = i
           For j = iblockrow To iblockrow2
              vaSpread1.Row = j
              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, Msgtitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
           Next j
       Next i
    End If
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col > 1 Then
        indcol = iblockcol
        vaSpread1.Col = 1
        If Trim(vaSpread1.text) <> "" Then MsgBox "No puede salirse del rango...", vbCritical + vbOKOnly, Msgtitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
        If (iblockrow - ((iblockrow2 - iblockrow) + 1)) < 1 Then
           MsgBox "Imposible subir la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, Msgtitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
        End If
        If iblockcol < 0 Then iblockcol = 1: iblockcol2 = vaSpread1.MaxCols
        For i = 1 To maxColumna
            If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
        Next i
        For i = 1 To maxColumna
            If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
            If vectorcol(i) = iblockcol2 Or vectorcol(i) + 1 = iblockcol2 Or vectorcol(i) + 2 = iblockcol2 Then iblockcol2 = (vectorcol(i) + 4): Exit For
        Next i
        '-------> Validar días modificados
'        For j = (iblockrow - 1) To (vaSpread1.MaxRows - 1)
'            For i = iblockcol To iblockcol2 Step 6
'                vaSpread1.Row = j
'                vaSpread1.Col = i + 1
'                If Trim(vaSpread1.text) <> "" Then
'                   vaSpread1.Col = (maxColumna * 6 + 1) + ((i + 3) / 6)
'                   If Trim(vaSpread1.text) = "" And vaSpread1.Col < vaSpread1.MaxCols Then vaSpread1.text = 1
'                End If
'            Next i
'        Next j
        '-------> Fin validar días modificados
        GrabarCambios vaSpread1.Row, 1, "Subir Linea"
        '-------> Copiar datos ultima fila
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.RowHidden = True
        vaSpread1.MoveRange iblockcol, (iblockrow - 1), iblockcol2, (iblockrow - 1), iblockcol, vaSpread1.MaxRows
        
        
        '-------> Copiar datos fila seleccionada
        vaSpread1.ClearRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow - 1), False
        vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, iblockrow, iblockcol, (iblockrow - 1)
        
        '---> SpreadClon
        SpreadClon.vaSpread1.ClearRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow - 1), False
        SpreadClon.vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, iblockrow, iblockcol, (iblockrow - 1)
        '---
        
        '-------> Devolver datos fila y restar ultima fila
        vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
        vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow
        
        '-------> Devolver datos fila y restar ultima fila
        SpreadClon.vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
        SpreadClon.vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow
       
        vaSpread1.Row = iblockrow - 1: vaSpread1.Col = iblockcol
        vaSpread1.DeleteRows vaSpread1.MaxRows, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        vaSpread1.SetActiveCell vaSpread1.Col, vaSpread1.Row
    ElseIf vaSpread1.Col = 1 Then
        If Trim(vaSpread1.text) = "" Then estgra = False: vaSpread1.Enabled = True: Exit Sub
        For i = iblockrow - 1 To 1 Step -1 'Recorre el espacio que hay entre la estructura seleccioneda y la anterior
            vaSpread1.Row = i
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next i
        For z = iblockrow + 1 To (vaSpread1.MaxRows - 1) 'Recorre el espacio que hay entre la estructura seleccioneda y la posterior
            vaSpread1.Row = z
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next z
        If z > (vaSpread1.MaxRows - 1) Then
            For fil = (vaSpread1.MaxRows - 1) To 1 Step -1
                For colu = 1 To vaSpread1.MaxCols
                    vaSpread1.Col = colu: vaSpread1.Row = fil
                    If Trim(vaSpread1.text) <> "" Then
                        z = fil + 1: Exit For
                    End If
                Next colu
'                If z <= (vaSpread1.MaxRows - 1) Then Exit For
                If z <= (vaSpread1.MaxRows) Then Exit For
            Next fil
        End If
        filaAct = iblockrow         'Fila actual
        filaAnt = IIf(i < 1, 1, i)  'Fila anterior
        filaPos = z                 'Fila posterior
        
        '-------> Validar días modificados
'        For j = filaAnt To (vaSpread1.MaxRows - 1)
'            For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
'                vaSpread1.Row = j
'                vaSpread1.Col = i + 1
'                If Trim(vaSpread1.text) <> "" Then
'                   vaSpread1.Col = (maxColumna * 6 + 1) + ((i + 3) / 6)
'                   If Trim(vaSpread1.text) = "" And vaSpread1.Col < vaSpread1.MaxCols Then vaSpread1.text = 1
'                End If
'            Next i
'        Next j
        '-------> Fin validar días modificados
        '------- Agregar filas temporales y respaldar
        vaSpread1.MaxRows = vaSpread1.MaxRows + (filaAct - filaAnt)
        For i = vaSpread1.MaxRows - (filaAct - filaAnt) + 1 To vaSpread1.MaxRows + (filaAct - filaAnt)
            vaSpread1.Row = i
            vaSpread1.RowHidden = True
        Next i
        vaSpread1.MoveRange 1, filaAnt, vaSpread1.MaxCols, (filaAct - 1), 1, vaSpread1.MaxRows - (filaAct - filaAnt) + 1
        
        '-------> Mover estructura
        vaSpread1.MoveRange 1, filaAct, vaSpread1.MaxCols, (filaPos - 1), 1, filaAnt
        
        '-------> Devolver respaldo
        vaSpread1.MoveRange 1, vaSpread1.MaxRows - (filaAct - filaAnt) + 1, vaSpread1.MaxCols, vaSpread1.MaxRows - (filaAct - filaAnt) + 1 + (filaAct - filaAnt - 1), 1, filaAnt + (filaPos - filaAct)
        

        
        For i = vaSpread1.MaxRows - (filaAct - filaAnt) + 1 To vaSpread1.MaxRows
            vaSpread1.DeleteRows i, 1
            vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        Next i
        vaSpread1.SetActiveCell 1, filaAnt
    End If
    vaSpread1.Row = iblockrow - 1: vaSpread1.Col = iblockcol
    iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow: iblockcol = vaSpread1.ActiveCol
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
    Plato(14).Enabled = False
    OpGrilla(14).Enabled = False
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    vaSpread1.Col = 1
    For i = 1 To (vaSpread1.MaxRows - 1)
        vaSpread1.Row = i
        vaSpread1.BackColor = Shape1(2).FillColor
    Next i
    vaSpread1.Enabled = True
Case 9 '-------> Bajar línea
    vaSpread1.Enabled = False
    vaSpread1.GetSelection 1, xColIni, xRowIni, xColFin, xRowFin
    
    iblockrow = xRowIni
    iblockrow2 = xRowFin
    iblockcol = xColIni
    iblockcol2 = xColFin
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Row = vaSpread1.MaxRows Then estgra = False: vaSpread1.Enabled = True: Exit Sub
    If iblockcol < 1 Or (iblockcol = 1 And Trim(vaSpread1.text) <> "") Then
       For i = 1 To maxColumna
           vaSpread1.Col = vectorcol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, Msgtitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
       Next i
    Else
       For i = iblockcol To iblockcol2
           vaSpread1.Col = i
           For j = iblockrow To iblockrow2
              vaSpread1.Row = j
              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, Msgtitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
           Next j
       Next i
    End If
    vaSpread1.Col = vaSpread1.ActiveCol
    '-------> Grabar Evento
    GrabarCambios vaSpread1.Row, j, "Bajar Linea"
    If vaSpread1.Col > 1 Then
        vaSpread1.Col = 1
        vaSpread1.Row = vaSpread1.ActiveRow + 1
        If Trim(vaSpread1.text) <> "" Then MsgBox "No puede salirse del rango...", vbCritical + vbOKOnly, Msgtitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow - 1
        If (iblockrow2 + ((iblockrow2 - iblockrow) + 1)) > (vaSpread1.MaxRows - 1) Then
           MsgBox "Imposible bajar la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, Msgtitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
        End If
        indcol = iblockcol
        If iblockcol < 0 Then iblockcol = 1: iblockcol2 = vaSpread1.MaxCols
        For i = 1 To maxColumna
            If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
        Next i
        For i = 1 To maxColumna
            If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
            If vectorcol(i) = iblockcol2 Or vectorcol(i) + 1 = iblockcol2 Or vectorcol(i) + 2 = iblockcol2 Then iblockcol2 = (vectorcol(i) + 4): Exit For
        Next i
        '-------> Validar días modificados
'        For j = iblockrow To (vaSpread1.MaxRows - 1)
'            For i = iblockcol To iblockcol2 Step 6
'                vaSpread1.Row = j
'                vaSpread1.Col = i + 1
'                If Trim(vaSpread1.text) <> "" Then
'                   vaSpread1.Col = (maxColumna * 6 + 1) + ((i + 3) / 6)
'                   If Trim(vaSpread1.text) = "" And vaSpread1.Col < vaSpread1.MaxCols Then vaSpread1.text = 1
'                End If
'            Next i
'        Next j
        '-------> Fin validar días modificados
        '-------> Copiar datos ultima fila
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.RowHidden = True
        vaSpread1.MoveRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow + 1), iblockcol, vaSpread1.MaxRows
    
    
        '-------> Copiar datos fila Seleccionada
        vaSpread1.ClearRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow + 1), False
        vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, iblockrow, iblockcol, (iblockrow + 1)
                
        '------->  SpreadClon
        SpreadClon.vaSpread1.ClearRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow + 1), False
        SpreadClon.vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, iblockrow, iblockcol, (iblockrow + 1)
       
        
    
        '-------> Devolver datos fila y restar ultima fila
        vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
        vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow
        
        '------->  SpreadClon
        SpreadClon.vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
        SpreadClon.vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow

        
        vaSpread1.DeleteRows vaSpread1.MaxRows, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        vaSpread1.Row = iblockrow + 1: vaSpread1.Col = iblockcol
        vaSpread1.SetActiveCell vaSpread1.Col, vaSpread1.Row
    ElseIf vaSpread1.Col = 1 Then
        If Trim(vaSpread1.text) = "" Then estgra = False: vaSpread1.Enabled = True: Exit Sub
        For z = iblockrow + 1 To (vaSpread1.MaxRows - 1) 'Recorre el espacio que hay entre la estructura seleccioneda y la posterior
            vaSpread1.Row = z
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next z
        If z > (vaSpread1.MaxRows - 1) Then estgra = False: vaSpread1.Enabled = True: Exit Sub
        vaSpread1.Col = vaSpread1.ActiveCol
        auxIblockrow = z
        For i = auxIblockrow - 1 To 1 Step -1 'Recorre el espacio que hay entre la estructura seleccioneda y la anterior
            vaSpread1.Row = i
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next i
        For z = auxIblockrow + 1 To (vaSpread1.MaxRows - 1) 'Recorre el espacio que hay entre la estructura seleccioneda y la posterior
            vaSpread1.Row = z
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next z
        If z > (vaSpread1.MaxRows - 1) Then
            For fil = (vaSpread1.MaxRows - 1) To 1 Step -1
                For colu = 1 To vaSpread1.MaxCols
                    vaSpread1.Col = colu: vaSpread1.Row = fil
                    If Trim(vaSpread1.text) <> "" Then
                        z = fil + 1: Exit For
                    End If
                Next colu
                If z <= (vaSpread1.MaxRows - 1) Then Exit For
            Next fil
        End If
        filaAct = auxIblockrow         'Fila actual
        filaAnt = IIf(i < 1, 1, i)  'Fila anterior
        filaPos = z                 'Fila posterior
        '-------> Validar días modificados
'        For j = filaAnt To (vaSpread1.MaxRows - 1)
'            For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
'                vaSpread1.Row = j
'                vaSpread1.Col = i + 1
'                If Trim(vaSpread1.text) <> "" Then
'                   vaSpread1.Col = (maxColumna * 6 + 1) + ((i + 3) / 6)
'                   If Trim(vaSpread1.text) = "" And vaSpread1.Col < vaSpread1.MaxCols Then vaSpread1.text = 1
'                End If
'            Next i
'        Next j
        '-------> Fin validar días modificados
        
        '-------> Agregar filas temporales y respaldar
        vaSpread1.MaxRows = vaSpread1.MaxRows + (filaAct - filaAnt)
        For i = vaSpread1.MaxRows - (filaAct - filaAnt) + 1 To vaSpread1.MaxRows + (filaAct - filaAnt)
            vaSpread1.Row = i
            vaSpread1.RowHidden = True
        Next i
        vaSpread1.MoveRange 1, filaAnt, vaSpread1.MaxCols, (filaAct - 1), 1, vaSpread1.MaxRows - (filaAct - filaAnt) + 1
        
        '-------> Mover estructura
        vaSpread1.MoveRange 1, filaAct, vaSpread1.MaxCols, (filaPos - 1), 1, filaAnt
        
        
        '-------> Devolver respaldo
        vaSpread1.MoveRange 1, vaSpread1.MaxRows - (filaAct - filaAnt) + 1, vaSpread1.MaxCols, vaSpread1.MaxRows - (filaAct - filaAnt) + 1 + (filaAct - filaAnt - 1), 1, filaAnt + (filaPos - filaAct)
        

        
        For i = vaSpread1.MaxRows - (filaAct - filaAnt) + 1 To vaSpread1.MaxRows
            vaSpread1.DeleteRows i, 1
            vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        Next i
        vaSpread1.SetActiveCell 1, filaAnt + (filaPos - filaAct)
    End If
    iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow: iblockcol = vaSpread1.ActiveCol
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
    Plato(14).Enabled = False
    OpGrilla(14).Enabled = False
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    vaSpread1.Col = 1
    For i = 1 To (vaSpread1.MaxRows - 1)
        vaSpread1.Row = i
        vaSpread1.BackColor = Shape1(2).FillColor
    Next i
    vaSpread1.Enabled = True
Case 11, 12 '-------> Copiar y pegar linea
    If vaSpread1.ActiveRow = vaSpread1.MaxRows Then estgra = False: Exit Sub
    If Index = 11 Then
       If iblockcol < 1 Then
          For i = 1 To maxColumna
              vaSpread1.Col = vectorcol(i)
              vaSpread1.Row = 1
              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede usar cortar", vbCritical + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
          Next i
       Else
          For i = iblockcol To iblockcol2
              vaSpread1.Col = i
              For j = iblockrow To iblockrow2
                 vaSpread1.Row = j
                 If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede usar cortar", vbCritical + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
              Next j
          Next i
       End If
    End If
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    '------> Verificar si copiar receta o raciones solamente
    vaSpread1.Row = 0
    If vaSpread1.text = "N.Rac." Then
      TipoCopia = "Copiar Raciones"
    Else
      TipoCopia = "Copiar Receta"
    End If
       
    aiblockrow = iblockrow: aiblockrow2 = iblockrow2
    aiblockcol = iblockcol: aiblockcol2 = iblockcol2
    
    If vaSpread1.Col = 1 Then estgra = False: Exit Sub
    Toolbar1.Buttons(6).Visible = False: Toolbar1.Buttons(7).Visible = True
       
    Plato(13).Enabled = True: OpGrilla(13).Enabled = True
    Plato(14).Enabled = True: OpGrilla(14).Enabled = True
'    Plato(14).Enabled = False: OpGrilla(14).Enabled = False
    If iblockcol < 1 Then aiblockcol = 2: aiblockcol2 = vaSpread1.MaxCols
    indcortarpegar = 1
    If Index = 11 Then
       indcortarpegar = 0
       Toolbar1.Buttons(8).Visible = True
       Toolbar1.Buttons(9).Visible = False
       Plato(14).Enabled = False
       OpGrilla(14).Enabled = False
    Else
       Toolbar1.Buttons(8).Visible = False
'       Toolbar1.Buttons(9).Visible = True
       Toolbar1.Buttons(9).Visible = True ' Cambié opcion a "True"  02/09/09 Samuel Melendez
       Plato(14).Enabled = True
       OpGrilla(14).Enabled = True
 '      Plato(14).Enabled = False
 '      OpGrilla(14).Enabled = False
    End If
Case 13, 14 '-------> Copiar y pegar
    
    
    If indcortarpegar = 0 Then
       If (iblockcol2 - iblockcol) > (aiblockcol2 - aiblockcol) Or (iblockrow2 - iblockrow) > (aiblockrow2 - aiblockrow) Then MsgBox "Imposible Pegar la infomación ya que el área de Cortar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
'      If iblockcol2 > aiblockcol2 Then
'         MsgBox "Imposible Cortar la infomación ya que el área de Cortar y el área de Pegado tienen formas distintas", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
'         estgra = False:Exit Sub
'      End If
       indcortarpegar = 0
    Else
       If (iblockrow2 - iblockrow) > (aiblockrow2 - aiblockrow) Then MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
       If iblockcol2 + (aiblockcol2 - aiblockcol) > (vaSpread1.MaxCols - maxColumna) Then MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única misma celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, Msgtitulo: estgra = False: Exit Sub   'aiblockcol <> iblockcol2 Or aiblockcol = 1 Then MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única misma celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
'       If aiblockcol <> iblockcol2 And aiblockcol = 1 Then MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única misma celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
    End If
    If iblockcol < 1 Then
       For i = 1 To maxColumna
           vaSpread1.Col = vectorcol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
       Next i
    Else
       For i = iblockcol To iblockcol2
           vaSpread1.Col = i
           For j = iblockrow To iblockrow2
              vaSpread1.Row = j
              If vaSpread1.BackColor = Shape1(1).FillColor And Index <> 14 Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, Msgtitulo: estgra = False: Exit Sub
           Next j
       Next i
    End If
    
    vaSpread1.Col = 1
    If vaSpread1.text = "Comensales" Then estgra = False: Exit Sub ' Valida que no se peguen recetas en la Línea de Comensales.
    
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col = 1 Then estgra = False: Exit Sub
    If indcortarpegar = 0 Then Toolbar1.Buttons(6).Visible = True: Toolbar1.Buttons(7).Visible = False
    '-------> Destinacion de copiar y pegar datos
    If iblockcol < 1 Then
       iblockcol = 2: iblockcol2 = vaSpread1.MaxCols
    End If
    
    If aiblockcol2 = vaSpread1.MaxCols Then aiblockcol2 = vaSpread1.MaxCols - 1
    If aiblockcol2 = (vaSpread1.MaxCols - maxColumna - 1) Then aiblockcol2 = (vaSpread1.MaxCols - maxColumna - 1)
    vaSpread1.Row = 0: vaSpread1.Col = iblockcol

    vaSpread1.Row = 0
    If vaSpread1.text = "N.Rac." And TipoCopia = "Copiar Raciones" Then
        cantCol = aiblockcol2 - aiblockcol
        CantCol1 = iblockcol2 - iblockcol
    Else
        For i = 1 To maxColumna
            If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For  'Inicio pegar
        Next i
        For i = 1 To maxColumna
             If (vectorcol(i) - 1) = iblockcol2 Or vectorcol(i) = iblockcol2 Or (vectorcol(i) + 1) = iblockcol2 Or (vectorcol(i) + 2) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 3)): Exit For ' Fin pegar
        Next i
    
        For i = 1 To maxColumna
            If (vectorcol(i) - 1) = aiblockcol Or vectorcol(i) = aiblockcol Or (vectorcol(i) + 1) = aiblockcol Or (vectorcol(i) + 2) = aiblockcol Then aiblockcol = (vectorcol(i) - 1): Exit For ' Columna de inicio copia
        Next i
        For i = 1 To maxColumna
            If (vectorcol(i) - 1) = aiblockcol2 Or vectorcol(i) = aiblockcol2 Or (vectorcol(i) + 1) = aiblockcol2 Or (vectorcol(i) + 2) = aiblockcol2 = (vectorcol(i) + 1) Then aiblockcol2 = (vectorcol(i) + 3): Exit For 'Fin copia   aiblockcol2 = (vectorcol(i) + 3) Copiaba hasta el cod receta ahora hasta Calorias
        Next i
        
        cantCol = aiblockcol2 - aiblockcol
        CantCol1 = iblockcol2 - iblockcol
    End If
    '-----> Llena vectores con las raciones
    LargoVec = aiblockrow2 - aiblockrow + 1
    If aiblockcol > 1 And aiblockrow > 0 Then
       ReDim VecSelGrid(0)
       ReDim VecSelGrid(20000)
       For i = aiblockcol To aiblockcol2
           vaSpread1.Col = i
           vaSpread1.Row = 0
           d = vaSpread1.text
           If vaSpread1.text = "N.Rac." Then
              For j = aiblockrow To aiblockrow + LargoVec - 1
                  vaSpread1.Col = i: vaSpread1.Row = j: d = vaSpread1.text
                  contador = contador + 1
                  If Trim(vaSpread1.text) <> "" Then VecSelGrid(contador) = vaSpread1.text    ' Almacena las raciones a copiar
              Next j
           End If
       Next i
    End If
    
    If vaSpread1.ActiveCol > 1 And vaSpread1.ActiveRow > 0 Then
       ReDim VecRacPegar(0)
       ReDim VecRacPegar(20000)
       For i = iblockcol To iblockcol2
           vaSpread1.Col = i
           vaSpread1.Row = 0
           If vaSpread1.text = "N.Rac." Then
              For j = vaSpread1.ActiveRow To vaSpread1.ActiveRow + contador - 1 'vaSpread1.MaxRows - 1
                  vaSpread1.Col = i: vaSpread1.Row = j
                  contador_b = contador_b + 1
                  If Trim(vaSpread1.text) <> "" Then VecRacPegar(contador_b) = vaSpread1.text ' Almacena las raciones a reemplazar
              Next j
           End If
       Next i
    End If
    
    indcol = aiblockcol: indcol2 = iblockcol2
    indrow = aiblockrow: indrow2 = aiblockrow2
    If Index = 14 And indcortarpegar = 1 Then
       If (aiblockrow2 - aiblockrow) <> 0 Or (aiblockcol2 - aiblockcol) <> 4 Then MsgBox "Por esta opción solamente puede copiar una receta", vbInformation + vbOKOnly, Msgtitulo: iblockcol = vaSpread1.ActiveCol: iblockcol2 = indcol2: estgra = False: Exit Sub
       '-------> Rutina pegado especial
       Dim nrodia As String
       vaSpread1.Row = 0: nrodia = ""
       For i = aiblockcol To aiblockcol2 Step 6
           vaSpread1.Col = i + 1
           nrodia = nrodia & Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 7, 2)) & ";"
       Next i
       For i = 1 To maxColumna
           vaSpread1.Col = vectorcol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then vaSpread1.Row = 0: nrodia = nrodia & Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 7, 2)) & ";"
       Next i
       
       vg_codigo = ""
       M_CpRPla.Inicio "Copia Especial Recetas en Planificación Real", "PLAREA", vg_fecha, nrodia
       M_CpRPla.Show 1
       If Trim(vg_codigo) = "" Then
          iblockcol = vaSpread1.ActiveCol: iblockcol2 = indcol2
          estgra = False: Exit Sub
       End If
       '-------> Grabar Evento Pegado Especial
       GrabarCambios 1, 1, "Pegado Especial"
       Dim vecdia() As String
       Dim xser As Long, iser As Long
       '-------> Mover días no permitidos
       ReDim Preserve vecdia(0)
       ValLcntH = "": i = 0
       For j = 1 To Len(vg_codigo)
           If Asc(Mid(vg_codigo, j, 1)) <> 59 Then
              ValLcntH = ValLcntH + Mid(vg_codigo, j, 1)
           Else
              ReDim Preserve vecdia(i): vecdia(i) = ValLcntH: ValLcntH = "": i = i + 1
           End If
       Next j
       If Trim(ValLcntH) <> "" Then ReDim Preserve vecdia(i): vecdia(i) = ValLcntH
'       Dim DeshacerLastModifi As Long
'       DeshacerLastModifi = DeshacerUltimoCammbio + 1
       vaSpread1.Enabled = False
       fg_carga ""
       For i = 3 To (vaSpread1.MaxCols - maxColumna) Step 6
           vaSpread1.Row = aiblockrow
           vaSpread1.Col = vaSpread1.MaxCols
           iser = Val(vaSpread1.text)
           vaSpread1.Row = 0
           vaSpread1.Col = i
           l = 0
           nrodia = Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 7, 2))
           For j = 0 To UBound(vecdia)
               If nrodia = vecdia(j) Then
                  vaSpread1.Row = aiblockrow: vaSpread1.Col = i - 1
                  If Trim(vaSpread1.text) <> "" Then
                     For x = aiblockrow + 1 To vaSpread1.MaxRows
                         vaSpread1.Row = x: vaSpread1.Col = vaSpread1.MaxCols: xser = Val(vaSpread1.text)
                         vaSpread1.Col = i + 1
                         If vaSpread1.Row = vaSpread1.MaxRows Then vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.InsertRows x, 1: l = x: Exit For
                         If xser <> iser And xser > 0 Then
                            vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.InsertRows x, 1: l = x: Exit For
                         ElseIf Trim(vaSpread1.text) <> "" And xser > 0 Then
                            vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.InsertRows x + 1, 1: x = x + 1: l = x: Exit For
                         ElseIf Trim(vaSpread1.text) = "" Then
                            Exit For
                         End If
                     Next x
                     vaSpread1.CopyRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i - 1, x
                     'DeshacerIntoMatriz aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i - 1, x, DeshacerLastModifi
                     
                     vaSpread1.Row = x: accion = "Copiar"
                  Else
                  '-----> Copia los elemenos seleccionados
                     vaSpread1.CopyRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i - 1, aiblockrow
                       
                     '--***--> este procedimiento guarda los cambios de manera que despues se puedan deshacer
                     'DeshacerIntoMatriz aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i - 1, aiblockrow, DeshacerLastModifi
                     vaSpread1.Row = aiblockrow: accion = "Copiar"

                  End If
                  '-------> Asignar colores
                  For x = (i - 1) To (i - 1) + 4
                      vaSpread1.Col = x
                      vaSpread1.BackColor = Shape1(0).FillColor
                      For xx = 1 To maxColumna
                          If (vectorcol(xx) - 1) = vaSpread1.Col Then
                              vaSpread1.Col = x + 2
                              vaSpread1.CellType = CellTypeNumber
                              vaSpread1.TypeNumberDecPlaces = 0
                              vaSpread1.TypeIntegerMin = 1
                              vaSpread1.TypeIntegerMax = 9999999
                              vaSpread1.TypeHAlign = TypeHAlignRight
                              vaSpread1.TypeSpin = False
                              vaSpread1.TypeIntegerSpinInc = 1
                              vaSpread1.TypeIntegerSpinWrap = False
                              Exit For
                          End If
                      Next xx
                      vaSpread1.Col = x
                      If x = (i - 1) Then vaSpread1.ForeColor = &HFF&: vaSpread1.BackColor = &H80FF80
                  Next x
                  If l > 0 Then
                     z = l
                     For l = 3 To (vaSpread1.MaxCols - maxColumna) Step 6
                         vaSpread1.Row = 1: vaSpread1.Col = l
                         If vaSpread1.BackColor = Shape1(1).FillColor Then
                            vaSpread1.Row = z
                            For x = (l - 1) To (l - 1) + 4
                                vaSpread1.Col = x
                                vaSpread1.BackColor = Shape1(1).FillColor
                            Next x
                         End If
                     Next l
                  End If
                  '-------> Fin asignar colores
                  '-------> Validar días modificados
'                  For z = aiblockrow To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow)
'                      vaSpread1.Row = z
'                      vaSpread1.Col = i ' + 1
'                      If Trim(vaSpread1.text) <> "" And ((maxColumna * 6 + 1) + ((i + 2) / 6)) < vaSpread1.MaxCols Then
'                         vaSpread1.Col = (maxColumna * 6 + 1) + ((i + 2) / 6)
'                         vaSpread1.text = 1
'                      End If
'                  Next z
                  '-------> Fin validar días modificados
                  Exit For
               End If
           Next j
       Next i
    Else
       '-------> Grabar Evento Copiado y Pegado
       GrabarCambios vaSpread1.Row, j, "Copiado y Pegado"
       indrow3 = vaSpread1.MaxRows
       For i = iblockcol To iblockcol2 Step 6
           If indcortarpegar = 1 Then
              vaSpread1.Row = aiblockrow: vaSpread1.Col = aiblockcol
              If vaSpread1.BackColor = Shape1(1).FillColor Then
                 vaSpread1.MaxRows = vaSpread1.MaxRows + (aiblockrow2 - aiblockrow) + 1
                vaSpread1.CopyRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow)
                'DeshacerIntoMatriz aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow)

                 accion = "Copiar"
                 '-------> Asignar colores
                 For j = vaSpread1.MaxRows - (aiblockrow2 - aiblockrow) To vaSpread1.MaxRows
                     vaSpread1.Row = j
                     For x = (i) To (i) + 4
                         vaSpread1.Col = x
                         vaSpread1.BackColor = Shape1(0).FillColor
                         For xx = 1 To maxColumna
                             If (vectorcol(xx) - 1) = vaSpread1.Col Then
                                vaSpread1.Col = x + 2
                                vaSpread1.CellType = CellTypeNumber
                                vaSpread1.TypeNumberDecPlaces = 0
                                vaSpread1.TypeIntegerMin = 1
                                vaSpread1.TypeIntegerMax = 9999999
                                vaSpread1.TypeHAlign = TypeHAlignRight
                                vaSpread1.TypeSpin = False
                                vaSpread1.TypeIntegerSpinInc = 1
                                vaSpread1.TypeIntegerSpinWrap = False
                                Exit For
                             End If
                         Next xx
                         vaSpread1.Col = x
                         If x = (i) And Trim(vaSpread1.text) <> "" Then vaSpread1.ForeColor = &HFF&: vaSpread1.BackColor = &H80FF80
                     Next x
                 Next j
                 '-------> Fin asignar colores
                 vaSpread1.CopyRange iblockcol, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow), iblockcol2, vaSpread1.MaxRows, i, vaSpread1.ActiveRow
                 'DeshacerIntoMatriz iblockcol, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow), iblockcol2, vaSpread1.MaxRows, i, vaSpread1.ActiveRow
                 vaSpread1.MaxRows = indrow3: accion = "Copiar"
              Else
                 vaSpread1.CopyRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.ActiveRow
                 accion = "Copiar"
                 'DeshacerIntoMatriz aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.ActiveRow

            End If
           ElseIf indcortarpegar = 0 Then
              vaSpread1.Row = aiblockrow: vaSpread1.Col = aiblockcol
              If vaSpread1.BackColor = Shape1(1).FillColor Then
                 vaSpread1.MaxRows = vaSpread1.MaxRows + (aiblockrow2 - aiblockrow) + 1
                 vaSpread1.MoveRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow)
                 
                  '--***--> este procedimiento guarda los cambios de manera que despues se puedan deshacer
                 'DeshacerIntoMatriz aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow)
                
                 
                 '-------> Asignar colores
                 For j = vaSpread1.MaxRows - (aiblockrow2 - aiblockrow) To vaSpread1.MaxRows
                     vaSpread1.Row = j
                     For x = (i) To (i) + 4
                         vaSpread1.Col = x
                         vaSpread1.BackColor = Shape1(0).FillColor
                         For xx = 1 To maxColumna
                             If (vectorcol(xx) - 1) = vaSpread1.Col Then
                                vaSpread1.Col = x + 2
                                vaSpread1.CellType = CellTypeNumber
                                vaSpread1.TypeNumberDecPlaces = 0
                                vaSpread1.TypeIntegerMin = 1
                                vaSpread1.TypeIntegerMax = 9999999
                                vaSpread1.TypeHAlign = TypeHAlignRight
                                vaSpread1.TypeSpin = False
                                vaSpread1.TypeIntegerSpinInc = 1
                                vaSpread1.TypeIntegerSpinWrap = False
                                Exit For
                             End If
                         Next xx
                         vaSpread1.Col = x
                         If x = (i) And Trim(vaSpread1.text) <> "" Then vaSpread1.ForeColor = &HFF&: vaSpread1.BackColor = &H80FF80
                     Next x
                 Next j
                 '-------> Fin asignar colores
                 vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow), iblockcol2, vaSpread1.MaxRows, i, vaSpread1.ActiveRow
                 
                 
                 '--***--> este procedimiento guarda los cambios de manera que despues se puedan deshacer
                 'DeshacerIntoMatriz iblockcol, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow), iblockcol2, vaSpread1.MaxRows, i, vaSpread1.ActiveRow

                 vaSpread1.MaxRows = indrow3: accion = "Cortar"
              Else
                 '------- Funcion CORTAR Y PEGAR
                 vaSpread1.MoveRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.ActiveRow
                 accion = "Cortar"
                 
                 '--***--> este procedimiento guarda los cambios de manera que despues se puedan deshacer
                 'DeshacerIntoMatriz aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.ActiveRow

                 ' aiblockcol = columna inicial del origen
                 ' aiblockrow = Fila inicial origen
                 ' aiblockcol2 = columna final origen
                 ' aiblockrow2 = fila final origen
                 ' i = columna inicial destino
                 ' vaSpread1.ActiveRow = fila inicial destino
                 

                 
                 
                 OpGrilla(13).Enabled = False
                 Toolbar1.Buttons(6).Visible = True
                 Toolbar1.Buttons(7).Visible = False
              End If
           End If
           For j = vaSpread1.ActiveRow To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow)
               vaSpread1.Row = j
               If indcortarpegar = 0 Then
'                  For x = aiblockcol To aiblockcol2 Step 6
'                      vaSpread1.Col = x + 1
'                      If Trim(vaSpread1.text) <> "" And ((maxColumna * 6 + 1) + ((x + 3) / 6)) < vaSpread1.MaxCols Then
'                         vaSpread1.Col = (maxColumna * 6 + 1) + ((x + 3) / 6)
'                         vaSpread1.text = 1
'                      End If
'                  Next x
               Else
                  If vaSpread1.ActiveRow = vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow) And ((iblockcol + (aiblockcol2 - aiblockcol)) - 6) > 6 Then
'                     For x = iblockcol To (iblockcol + (aiblockcol2 - aiblockcol)) Step 6
'                         vaSpread1.Col = x + 1
'                         If Trim(vaSpread1.text) <> "" And ((maxColumna * 6 + 1) + ((x + 3) / 6)) < vaSpread1.MaxCols Then
'                            vaSpread1.Col = (maxColumna * 6 + 1) + ((x + 3) / 6)
'                            If vaSpread1.Col < vaSpread1.MaxCols Then vaSpread1.text = 1
'                         End If
'                     Next x
                  Else
''                     For x = vaSpread1.ActiveRow To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow)
''                         vaSpread1.Row = j
'                         For xx = IIf((i - 1) = 1, 2, (i - 1)) To (iblockcol2 + 1) Step 6
'                            vaSpread1.Col = xx + 2
'                            If Trim(vaSpread1.text) <> "" And ((maxColumna * 6 + 1) + ((xx + 3) / 6)) < vaSpread1.MaxCols Then
'                               vaSpread1.Col = (maxColumna * 6 + 1) + ((xx + 3) / 6)
'                               If vaSpread1.Col < vaSpread1.MaxCols Then vaSpread1.text = 1
'                            End If
'                         Next xx
''                     Next x
                  End If
               End If
           Next j
           '-------> Fin validar días modificados
       Next i
    End If
    If indcos = True Then
       For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
           Calctodia 1, i + 1
       Next i
       MostrarCosto vaSpread1.ActiveCol
       
    End If
    '------> Se trabaja como excel las raciones
    ColumnaActiva = vaSpread1.ActiveCol: FilaActiva = vaSpread1.ActiveRow: ColumnaAntActiva = ColumnaActiva - 1
    vaSpread1.Col = ColumnaActiva: vaSpread1.Row = 0
    If ColumnaActiva > 1 And accion = "Copiar" Then
      vaSpread1.Row = 0
      '-------->  Copia en posición Ración
      If vaSpread1.text = "N.Rac." Then
            If contador = 1 Then
              n = 1: n1 = 1: Max = contador: Max1 = contador_b
              For ff = FilaActiva To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow) 'vaSpread1.MaxRows
                vaSpread1.Row = FilaActiva
                If Trim(VecRacPegar(n1)) = "" Then
                  vaSpread1.Col = f - 1: desc = vaSpread1.text
                  vaSpread1.Col = f: vaSpread1.Row = FilaActiva
                  vaSpread1.text = IIf(Trim(desc) = "", "", 0)
                ElseIf Trim(VecRacPegar(n1)) = "0" Then
                  vaSpread1.text = IIf(Trim(VecRacPegar(n1)) = "0", VecSelGrid(n), VecRacPegar(n1))
                Else
                  If TipoCopia = "Copiar Raciones" Then
                    vaSpread1.text = Trim(VecSelGrid(n))
                  'Else
                  '  vaSpread1.text = Trim(VecRacPegar(n1))
                  End If
                End If
                If n <= Max Then n = n + 1
                If n1 <= Max1 Then n1 = n1 + 1
                If n > Max Then n = 1
                If n1 > Max1 Then Exit For
              Next ff
            ElseIf contador > 1 Then
              n = 1: n1 = 1: Max = contador: Max1 = contador_b
              For g = FilaActiva To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow) 'vaSpread1.MaxRows
                vaSpread1.Col = ColumnaActiva: vaSpread1.Row = g
                    If Trim(VecRacPegar(n1)) = "" Then
                      vaSpread1.Col = ColumnaAntActiva: vaSpread1.Row = g: desc = vaSpread1.text
                      vaSpread1.Col = ColumnaActiva: vaSpread1.Row = g
                      vaSpread1.text = IIf(Trim(desc) = "", "", 0)
                    ElseIf Trim(VecRacPegar(n1)) = "0" Then
                      vaSpread1.text = IIf(Trim(VecRacPegar(n1)) = "0", VecSelGrid(n), VecRacPegar(n1))
                    Else
                      If TipoCopia = "Copiar Raciones" Then
                        vaSpread1.text = Trim(VecSelGrid(n))
                      Else
                        vaSpread1.text = Trim(VecRacPegar(n1))
                      End If
                    End If
                    If n <= Max Then n = n + 1
                    If n1 <= Max1 Then n1 = n1 + 1
                    If n > Max Then n = 1
                    If n1 > Max1 Then Exit For
              Next g
            End If
      '-------->  Copia en posición Costo
      ElseIf vaSpread1.text = "Costo" Then
        Tope = ColumnaActiva - CantCol1
        For f = ColumnaActiva To Tope Step -1
          vaSpread1.Col = f: vaSpread1.Row = 0
          If vaSpread1.text = "N.Rac." Then
              n = 1: n1 = 1: Max = contador: Max1 = contador_b
              For g = FilaActiva To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow) 'vaSpread1.MaxRows
                vaSpread1.Col = f: vaSpread1.Row = g
                  If Trim(VecRacPegar(n1)) = "" Then
                    vaSpread1.Col = f - 1: desc = vaSpread1.text
                    vaSpread1.Col = f
                    vaSpread1.text = IIf(Trim(desc) = "", "", 0)
                  ElseIf Trim(VecRacPegar(n1)) > 0 Then
                    vaSpread1.text = Trim(VecRacPegar(n1))
                  Else
                    vaSpread1.text = VecSelGrid(n)
                  End If
                  If n <= Max Then n = n + 1
                  If n1 <= Max1 Then n1 = n1 + 1
                  If n > Max Then n = 1
                  If n1 > Max1 Then Exit For
              Next g
          End If
        Next f
      Else
        '-------->  Distinta posición a la anterior
        For f = ColumnaActiva To vaSpread1.ActiveCol + CantCol1 ' vaSpread1.MaxCols
          vaSpread1.Col = f: vaSpread1.Row = 0
          If vaSpread1.text = "N.Rac." Then
            If contador = 1 Then
              n = 1: n1 = 1: Max = contador: Max1 = contador_b
              For ff = FilaActiva To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow) 'vaSpread1.MaxRows
                vaSpread1.Row = FilaActiva
                If Trim(VecRacPegar(n1)) = "" Then
                  vaSpread1.Col = f - 1: desc = vaSpread1.text
                  vaSpread1.Col = f: vaSpread1.Row = FilaActiva
                  vaSpread1.text = IIf(Trim(desc) = "", "", 0)
                Else
                  vaSpread1.text = IIf(Trim(VecRacPegar(n1)) = 0, VecSelGrid(n), VecRacPegar(n1))
                End If
                If n <= Max Then n = n + 1
                If n1 <= Max1 Then n1 = n1 + 1
                If n > Max Then n = 1
                If n1 > Max1 Then Exit For
              Next ff
            Else
              n = 1: n1 = 1: Max = contador: Max1 = contador_b
              For g = FilaActiva To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow) 'vaSpread1.MaxRows
                vaSpread1.Col = f: vaSpread1.Row = g
                  If Trim(VecRacPegar(n1)) = "" Then
                    vaSpread1.Col = f - 1: desc = vaSpread1.text
                    vaSpread1.Col = f
                    vaSpread1.text = IIf(Trim(desc) = "", "", 0)
                  ElseIf Trim(VecRacPegar(n1)) > 0 Then
                    vaSpread1.text = Trim(VecRacPegar(n1))
                  Else
                    vaSpread1.text = VecSelGrid(n)
                  End If
                  If n <= Max Then n = n + 1
                  If n1 <= Max1 Then n1 = n1 + 1
                  If n > Max Then n = 1
                  If n1 > Max1 Then Exit For
              Next g
            End If
          End If
        Next f
      End If
    End If
    '------>
    aiblockcol = indcol: iblockcol2 = indcol2
    aiblockrow = indrow: aiblockrow2 = indrow2
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    vaSpread1.Enabled = True
    fg_descarga
Case 15
    B_BusVas.Partidas Me
    B_BusVas.Show 1
Case 16
    If vaSpread1.ActiveCol = 1 And vaSpread1.ActiveRow <> vaSpread1.MaxRows And Trim(vaSpread1.text) <> "" And vg_IndpprSelec = 2 Then
        G_Proc.CellEdite B_CelEdi, "Editar Estructura", "Nombre Estructura", vaSpread1
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
    End If
Case 17
    
  
    DoEvents
    T_Servic.SSTab1.TabEnabled(0) = False
    DoEvents
    T_Servic.SSTab1.TabEnabled(2) = False
    DoEvents
    T_Servic.CallForm = Me.Name
    DoEvents
    T_Servic.SSTab1.Tab = 1
    DoEvents
    T_Servic.MoverDatosGrillas2 vg_codservicio
    DoEvents
    T_Servic.Show 1
    DoEvents
End Select
CargarAporteCalorico
estgra = False

Exit Sub
Man_Error:
vaSpread1.Enabled = True
fg_descarga
End Sub

Private Sub Timer1_Timer()
On Error GoTo Man_Error
If Toolbar1.Buttons(2).Visible = False Then Exit Sub
' variable estática para acumular la cantidad de segundos
'Static Temp_Seg As Long
' incrementa
vg_TemSeg = vg_TemSeg + 1
' comprueba que los segundos no sea igual a la cantidad de minutos _
  que queremos , en este caso 5 minutos
If (vg_TemSeg * 60) >= (vg_IntMin * 60) * 60 And Not estgra Then
   ' reestablece
   estgra = True
   If MsgBox(" Actualiza planificación ...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Cancel = -1: vg_TemSeg = 0: estgra = False: Exit Sub
   If Toolbar1.Buttons(2).Visible = True Then
      Toolbar1.Enabled = False
      Toolbar1.Buttons(31).Enabled = False
      CorDes = 0
      GrabarPlantillaMinuta
      CorDes = 0
      If Dir(LCase(App.Path) & "\" & "Spread" & vg_NUsr & "*" & ".ss6") <> "" Then Kill LCase(App.Path) & "\" & "Spread" & vg_NUsr & "*" & ".ss6"
      Toolbar1.Enabled = True
   End If
   Toolbar1.Buttons(1).Visible = True
   Toolbar1.Buttons(2).Visible = False
   estgra = False
   vg_TemSeg = 0
End If
Exit Sub
Man_Error:
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 2
    Plantilla_Click (0)
Case 4
    Plato_Click (11)
Case 5
    Plato_Click (12)
Case 7
    Plato_Click (13)
Case 9
    Plato_Click (14)
Case 10
   Plato_Click (15)
Case 12
    Plato_Click (5)
Case 13
    Plato_Click (6)
Case 15
    Plato_Click (8)
Case 16
    Plato_Click (9)
Case 18
    Plantilla_Click (5)
Case 19
    Plantilla_Click (8)
Case 21
    Plantilla_Click (10)
Case 22
    Plantilla_Click (3)
Case 23
    Plantilla_Click (11)
Case 24
    Plantilla_Click (14)
Case 25
    Plantilla_Click (13)
Case 27
    ExportarExcel
Case 29
    HabilitaCeldaCalorias
Case 30
    Plantilla_Click (20)
Case 31
    Plantilla_Click (21)
Case 32
    Plantilla_Click (22)
End Select
End Sub

Sub ExportarExcel()
CargarAporteCalorico
Dim NashXl As Excel.Application
Dim irow As Long, irow2 As Long
Dim NColumnas As Integer
fg_carga ""
Set NashXl = CreateObject("excel.application")
Set NashXl = New Excel.Application
NashXl.SheetsInNewWorkbook = 1
NashXl.Workbooks.Add
NashXl.Range("A1").Select
NashXl.ActiveCell.FormulaR1C1 = "Sub-Segmento : " & vg_codsubseg & "-" & vg_nomsubseg
NashXl.Range("A2").Select
NashXl.ActiveCell.FormulaR1C1 = "Regimen      : " & vg_codregimen & "-" & vg_nomreg
NashXl.Range("A3").Select
NashXl.ActiveCell.FormulaR1C1 = "Servicio     : " & vg_codservicio & "-" & vg_nomser
NashXl.Range("A4").Select
NashXl.ActiveCell.FormulaR1C1 = "Fecha        : " & vg_fecha

maxColumna = Val(fg_mes(Mid(vg_fecha, 5, 2) & Mid(vg_fecha, 1, 4)))
NColumnas = (maxColumna * 6) + 1
vaSpread1.AllowMultiBlocks = True
'vaSpread1.SetSelection 1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows + 3
vaSpread1.SetSelection 1, -1, NColumnas, vaSpread1.MaxRows + 3
vaSpread1.ClipboardCopy

irow = vaSpread1.MaxRows + 5
'------- Pegar vaspread1(0) - Planilla Excel
NashXl.Range("A5").Select
NashXl.ActiveSheet.Paste
'------- Asignar color
'------- Colorear titulo
NashXl.Range("A5:GE5").Select
With NashXl.Selection.Interior
     .ColorIndex = 15
     .Pattern = xlSolid
End With
'------- Dibujar marco
NashXl.Range("A5:GE" & irow).Select
NashXl.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
NashXl.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With NashXl.Selection.Borders(xlEdgeLeft)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeTop)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeBottom)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeRight)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideVertical)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideHorizontal)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
NashXl.Range("A2" & ":" & "A" & irow).Select
NashXl.Selection.NumberFormat = "#,##0.00"

'------- Asigna Colores a Estructura de Servicio
NashXl.Range("A6:" & "A" & irow).Select
With NashXl.Selection.Interior
     .ColorIndex = 10
     .Pattern = xlSolid
End With
'------- Aplicar totales

NashXl.Selection.Font.Bold = True

NashXl.Range("B" & irow & ":" & "B" & 2).Select
NashXl.Selection.NumberFormat = "#,##0.00"
'------- Ajustar columna
NashXl.Cells.Select
NashXl.Cells.EntireColumn.AutoFit
vaSpread1.AllowMultiBlocks = False: vaSpread1.SetSelection 1, 0, vaSpread1.MaxCols, vaSpread1.MaxRows

NashXl.Cells.Replace What:="&0&;", Replacement:="", LookAt:=xlPart, SearchOrder _
      :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
NashXl.Cells.Replace What:="&-1&;", Replacement:="", LookAt:=xlPart, SearchOrder _
      :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
NashXl.Cells.Replace What:="&1&;", Replacement:="", LookAt:=xlPart, SearchOrder _
      :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
fg_descarga
NashXl.Visible = True
End Sub



Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
indactivo = 1
iblockrow = BlockRow
iblockrow2 = BlockRow2
iblockcol = BlockCol
iblockcol2 = IIf(estapo = False, BlockCol2 + 1, BlockCol2)
If BlockRow < 0 Then iblockrow = 1
'jpaz If BlockRow2 < 0 Then iblockrow2 = 100
'jpaz If BlockRow2 > 100 Then iblockrow2 = 100
If BlockRow2 < 0 Then iblockrow2 = (vaSpread1.MaxRows - 1)
If BlockRow2 >= vaSpread1.MaxRows Then iblockrow2 = (vaSpread1.MaxRows - 1)
'CargarAporteCalorico
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If Row < 1 Then Exit Sub
OpGrilla(15).Enabled = IIf(Col = 1, True, False)
Plato(15).Enabled = IIf(Col = 1, True, False)
indactivo = 1
iblockrow = vaSpread1.ActiveRow
iblockrow2 = vaSpread1.ActiveRow
iblockcol = vaSpread1.ActiveCol
iblockcol2 = vaSpread1.ActiveCol
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.ActiveCol
'CargarAporteCalorico
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
If Col = 1 Then Plato_Click (16): Exit Sub
If Row < 1 Or Col = 1 Then Exit Sub
Plato_Click (2)
'CargarAporteCalorico
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
        
        
    'If ChangeMade = True Then DeshacerInto Col, Row, vaSpread1, DeshacerUltimoCammbio + 1

    
    If Toolbar1.Buttons(2).Enabled = False Then Exit Sub
    vaSpread1.Row = Row: vaSpread1.Col = Col
    If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    If vaSpread1.ChangeMade = False Or Col = 1 Or Mode = 1 Then i = IIf(vaSpread1.text = "", "0", vaSpread1.text): Exit Sub
    'If vaSpread1.ChangeMade = False Or Col = 1 Or Mode = 1 Then Exit Sub
    '-------> Grabar Evento Modificar
    GrabarCambios 1, j, "Modificar Estructura"
    If vaSpread1.ChangeMade = True Then vaSpread1.Col = (maxColumna * 6 + 1) + (vaSpread1.Col / 6): vaSpread1.text = 1: If indcos = True Then vaSpread1.Col = Col: j = Col - 1:  Calctodia vaSpread1.Row, j 'veccos((Int(J / 5) + 1), 4) = Round(veccos((Int(J / 5) + 1), 4) - (i), vg_DPr): veccos((Int(J / 5) + 1), 4) = Round(veccos((Int(J / 5) + 1), 4) + (vaSpread1.Text), vg_DPr)
    vaSpread1.Row = Row
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    CargarAporteCalorico
End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
If Toolbar1.Buttons(2).Enabled = False Then Exit Sub
Dim delrow As Integer, indcol As Integer, indrow As Integer, indcol2 As Integer, indrow2 As Integer
Select Case KeyCode
Case 65 To 90
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    ws_respuesta = ""
    ws_respuesta = Chr(KeyCode)
    Plato_Click (2)
Case 86
    Exit Sub
Case 46
    Select Case vaSpread1.ActiveCol
    Case 1
        Dim xRow As Integer
        Dim codest As Integer
        Dim AvisoEst As Boolean
        vaSpread1.GetSelection 1, xColIni, xRowIni, xColFin, xRowFin
        
        iblockcol = xColIni
        iblockrow = xRowIni
        iblockcol2 = xColIni
        iblockrow2 = xRowFin
        
        If vaSpread1.MaxRows = vaSpread1.ActiveRow Or vaSpread1.MaxRows = iblockrow Or vaSpread1.MaxRows = iblockrow2 Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = vaSpread1.ActiveCol
        If vaSpread1.Col = 1 And vaSpread1.Row <> vaSpread1.MaxRows Then
           '-------> Grabar Evento Modificar Estructura
           GrabarCambios 1, 1, "Modificar Estructura"
            
            DesqloqSubMenu (vaSpread1.text)
            vaSpread1.text = ""
             'DeshacerInto vaSpread1.Col, vaSpread1.Row, vaSpread1, DeshacerUltimoCammbio + 1
            xRow = vaSpread1.ActiveRow
            AvisoEst = False
            For i = vaSpread1.ActiveRow To 1 Step -1
                vaSpread1.Row = i
                If Trim(vaSpread1.text) <> "" Then
                    vaSpread1.Col = vaSpread1.MaxCols
                    codest = vaSpread1.text
                    vaSpread1.Row = xRow
                    vaSpread1.text = codest
                    vaSpread1.Col = 1
                    AvisoEst = True
                    Exit For
                End If
                
            Next i
            
            If AvisoEst = False And Trim(StrRec) <> "" Then
               If Dir(LCase(App.Path) & "\" & StrRec) <> "" Then Kill LCase(App.Path) & "\" & StrRec
               CorDes = CorDes - 1
               MsgBox "Al borrar esta estructura, dejará recetas sin asignar", vbCritical
            End If
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = True
        End If
        If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
        
        j = 0
        For i = 1 To maxColumna
            If (vectorcol(i) - 1) = vaSpread1.Col Or vectorcol(i) = vaSpread1.Col Then j = (vectorcol(i) - 1): Exit For
        Next i
        If j = 0 Then Exit Sub
        Plato(0).Enabled = True
        OpGrilla(0).Enabled = True
        Plato(13).Enabled = False
        OpGrilla(13).Enabled = False
        If indactivo = 0 Or iblockcol < 1 Or iblockrow < 1 Then iblockcol = vaSpread1.ActiveCol: iblockcol2 = vaSpread1.ActiveCol: iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow
        
       
        aiblockcol = iblockcol
        aiblockrow = iblockrow
        aiblockcol2 = iblockcol2
        aiblockrow2 = iblockrow2
        If iblockcol < 0 Then iblockcol = 2: iblockcol2 = vaSpread1.MaxCols
        aiblockcol = iblockcol
        aiblockrow = iblockrow
        aiblockcol2 = iblockcol2
        aiblockrow2 = iblockrow2
        For i = 1 To maxColumna
            If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
        Next i
        For i = 1 To maxColumna
            If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
            If vectorcol(i) = iblockcol2 Then iblockcol2 = (vectorcol(i) + 4): Exit For
        Next i
        indcol = aiblockcol: indcol2 = iblockcol2
        indrow = aiblockrow: indrow2 = IIf(aiblockrow2 = vaSpread1.MaxRows, (aiblockrow2 - 1), aiblockrow2)
        vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow2, False
        
        'DeshacerMatriz iblockcol, iblockrow, iblockcol2, iblockrow2, DeshacerUltimoCammbio + 1
        
        If indcos = True Then
           For i = iblockcol To iblockcol2 Step 6
               Calctodia 1, i + 1
           Next i
           MostrarCosto vaSpread1.ActiveCol
        End If
        iblockcol = auxcol
        vaSpread1.BlockMode = False
        Plantilla(0).Enabled = True
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
        indactivo = 0
    Case Is > 1
        vaSpread1.GetSelection 1, xColIni, xRowIni, xColFin, xRowFin
        
        iblockcol = xColIni
        iblockrow = xRowIni
        iblockcol2 = xColFin
        iblockrow2 = xRowFin
        
        If vaSpread1.MaxRows = vaSpread1.ActiveRow Or vaSpread1.MaxRows = iblockrow Or vaSpread1.MaxRows = iblockrow2 Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = vaSpread1.ActiveCol
        If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
        j = 0
        For i = 1 To maxColumna
            If (vectorcol(i) - 1) = vaSpread1.Col Or vectorcol(i) = vaSpread1.Col Then j = (vectorcol(i) - 1): Exit For
        Next i
        If j = 0 Then Exit Sub
        Plato(0).Enabled = True
        OpGrilla(0).Enabled = True
        Plato(13).Enabled = False
        OpGrilla(13).Enabled = False
        If indactivo = 0 Or iblockcol < 1 Or iblockrow < 1 Then iblockcol = vaSpread1.ActiveCol: iblockcol2 = vaSpread1.ActiveCol: iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow
         
        '-------> Grabar Evento Eliminación recetas
        GrabarCambios 1, 1, "Eliminación Recetas"
         
        aiblockcol = iblockcol
        aiblockrow = iblockrow
        aiblockcol2 = iblockcol2
        aiblockrow2 = iblockrow2
        If iblockcol < 0 Then iblockcol = 2: iblockcol2 = vaSpread1.MaxCols
        aiblockcol = iblockcol
        aiblockrow = iblockrow
        aiblockcol2 = iblockcol2
        aiblockrow2 = iblockrow2
        For i = 1 To maxColumna
            If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
        Next i
        For i = 1 To maxColumna
            If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
            If vectorcol(i) = iblockcol2 Then iblockcol2 = (vectorcol(i) + 4): Exit For
        Next i
        indcol = aiblockcol: indcol2 = iblockcol2
        indrow = aiblockrow: indrow2 = IIf(aiblockrow2 = vaSpread1.MaxRows, (aiblockrow2 - 1), aiblockrow2)
        vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow2, False
        If indcos = True Then
           For i = iblockcol To iblockcol2 Step 6
               Calctodia 1, i + 1
           Next i
           MostrarCosto vaSpread1.ActiveCol
        End If
        iblockcol = auxcol
        vaSpread1.BlockMode = False
        Plantilla(0).Enabled = True
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
        indactivo = 0
   
    End Select
End Select
CargarAporteCalorico
End Sub



Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
'iblockrow = BlockRow
'iblockrow2 = BlockRow2
'iblockcol = BlockCol
'iblockcol2 = BlockCol2
'If BlockRow < 0 Then iblockrow = 1
'If BlockRow2 < 0 Then iblockrow2 = (vaSpread1.MaxRows - 1)
'If BlockRow2 >= vaSpread1.MaxRows Then iblockrow2 = (vaSpread1.MaxRows - 1)
iblockrow = NewRow
iblockrow2 = NewRow
iblockcol = NewCol
iblockcol2 = NewCol
If NewRow < 0 Then iblockrow = 1
If NewRow < 0 Then iblockrow2 = (vaSpread1.MaxRows - 1)
If NewRow >= vaSpread1.MaxRows Then iblockrow2 = (vaSpread1.MaxRows - 1)
If indcos = False Or NewCol < 1 Then Exit Sub
MostrarCosto NewCol

'CargarAporteCalorico
End Sub

Private Sub vaspread1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Button
Case 2
    If vaSpread1.Visible <> True Then Exit Sub
    Indvaspread1 = 0
    If Mid(ValidaPerfil(M_Plami2), 1, 4) <> "1000" = True Then
        PopupMenu MenuDetalle
        
    End If
End Select
'CargarAporteCalorico
End Sub

Private Sub Opgrilla_Click(Index As Integer)
Select Case Index
Case 0
    Plato_Click (0)
Case 2
    Plato_Click (2)
Case 3
    Plato_Click (3)
Case 5
    Plato_Click (5)
Case 6
    Plato_Click (6)
Case 8
    Plato_Click (8)
Case 9
    Plato_Click (9)
Case 11
    Plato_Click (11)
Case 12
    Plato_Click (12)
Case 13
    Plato_Click (13)
Case 14
    Plato_Click (14)
Case 15
    Plato_Click (15)
Case 16
    Plato_Click (16)
Case 17
    Plato_Click (17)
End Select
End Sub

Private Sub GrabarPlantillaMinuta()
Dim desc As String, StrRec As String, StrRecb As String, NameEstManual As String, NameEst As String
Dim codrec As Long, numrac As Long, estser As Long, fecha As Long, conregdet As Long, indice As Long, existedat As Long, inddia As Long, tiprec As Long
Dim fechasis As Long, fecini As Long, fecfin As Long, totrac As Long
Dim cosali As Double, cospro As Double, cosdes As Double
On Error GoTo Man_Error
NameEstManual = ""
inddia = 1: conregdet = 0: gauge1.Value = 0: gauge.Value = 0: fecha = 0: fecini = 0: fecfin = 0
Picture1.Visible = True: Label3.Visible = True: gauge.Visible = True
Picture1.Refresh: Label3.Refresh: gauge.Refresh: gauge1.Refresh
fechasis = 0: fechasis = Val(Mid(Date, 7, 4)) & Mid(Date, 4, 2) & Mid(Date, 1, 2)
fg_carga ""
'-------> Grabar planificación minutas
For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
    DoEvents
    If inddia > maxColumna Then Exit Sub
    gauge1.Value = Val((inddia / maxColumna) * 100)
    Label3.Caption = "": Label3.Caption = "Día : " & inddia
    existedat = 0: vaSpread1.Row = 1: vaSpread1.Col = i
    If (vaSpread1.MaxRows - 1) = 0 Then
        existedat = 0
        fecha = Val(vg_fecha) & fg_pone_cero(inddia, 2)
    Else
       For j = 1 To (vaSpread1.MaxRows - 1)
           vaSpread1.Row = j
           fecha = Val(vg_fecha) & fg_pone_cero(inddia, 2)
           vaSpread1.Col = i + 1
           If Trim(vaSpread1.text) <> "" Then existedat = 1: Exit For
       Next j
    End If
    indice = 0
    vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = i + 2: totrac = Val(vaSpread1.text)
    Set RS = vg_db.Execute("sgpadm_s_planifminuta 4, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & "," & vg_Zona & ", " & Val(fecha) & ", 0, 0,'" & vg_IndpprSelec & "'")
    If Not RS.EOF Then
       indice = RS!min_codigo: RS.Close: Set RS = Nothing
       If indice > 0 And existedat = 0 Then
          vg_db.Execute "sgpadm_d_minutadet 'E2', " & indice & ", '1', 0, 0, 0, 0, '', 0, 0, 0, 0"
          vg_db.Execute "sgpadm_d_minuta 'E', " & indice & ", 0, 0, 0, 0, 0, 0, 0, 0, '', '" & vg_IndpprSelec & "'"
       Else
          vaSpread1.Row = 1
          If vaSpread1.BackColor <> Shape1(1).FillColor Then
             vg_db.Execute "sgpadm_iu_minuta 'M1', " & indice & ", " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & Val(fecha) & ", 0, " & totrac & ", " & totrac & ", 0, '', '" & vg_IndpprSelec & "'"
          End If
       End If
    Else
       RS.Close: Set RS = Nothing
       If existedat > 0 Then
          Set RS = vg_db.Execute("sgpadm_iu_minuta 'A', 0, '" & vg_codsubseg & "', " & vg_codregimen & ", " & vg_codservicio & ", " & Val(fecha) & ", 0, " & totrac & ", " & totrac & ", 0, '', '" & vg_IndpprSelec & "'")
          If Not RS.EOF Then
             indice = RS!indice
          End If
          RS.Close: Set RS = Nothing
       End If
    End If
    gauge.Value = 0: conregdet = 0: estser = 0
    If existedat > 0 Then
       If maxfila > vaSpread1.MaxRows Then
          '-------> Si maximo de fila es mayor que grilla borra detalle
          For j = vaSpread1.MaxRows To maxfila
              vg_db.Execute "sgpadm_d_minutadet 'E1', " & indice & ", '1', " & j & ", 0, 0, 0, '', 0, 0, 0, 0"
          Next j
       End If
       '-------> Actualizar detalle minutas
       For j = 1 To (vaSpread1.MaxRows - 1)
           conregdet = conregdet + 1
           gauge.Value = Val((conregdet / (vaSpread1.MaxRows - 1)) * 100)
           desc = "": codrec = 0: numrec = 0: cosali = 0: cosdes = 0
           vaSpread1.Row = j
          
           '---------- Samuel Melendez 28/09/09
           '** Si el nombre de la estructura fue ingresado manualmente
           '** por el usuario se llena la variable "NameEstManual", sino queda vacia
           'NameEstManual = ValidaNombreEstructura(j, vaSpread1)
'jpaz           NameEstManual = EstructuraSuperior(vaSpread1, j)
           '-----------------------------------
                      
           
           vaSpread1.Col = vaSpread1.MaxCols
           If Trim(vaSpread1.text) <> "" Then
              estser = vaSpread1.text
              vaSpread1.Col = 1
              If vg_IndpprSelec = "2" And Trim(vaSpread1.text) <> "" Then
                 NameEst = vaSpread1.text
                 vaSpread1.Col = vaSpread1.MaxCols - 1
                 NameEstManual = vaSpread1.text
                 
                 If NameEstManual <> NameEst Then
                    NameEstManual = NameEst
                 Else
                    NameEstManual = IIf(Trim(NameEstManual) = "", "", NameEstManual)
                 End If
              ElseIf vg_IndpprSelec = "1" Then
                 NameEstManual = ""
              End If
           End If
           vaSpread1.Col = i + 1: desc = Trim(vaSpread1.text)
           
           If desc <> "" And estser > 0 Then
              vaSpread1.Col = i + 2: numrac = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
              vaSpread1.Col = i + 3: cosali = IIf(Trim(vaSpread1.text) = "", 0, Val(vaSpread1.text))
              vaSpread1.Col = i + 4: d = vaSpread1.text
              
              StrRec = vaSpread1.text
              If Len(StrRec) <> 0 Then
                 Do While InStr(StrRec, ";") <> 0
                    StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
                    StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
                    codrec = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)):
                    StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
                    tiprec = Val(Mid(StrRecb, 1))
                 Loop
              End If
              
              '----> Traer costo receta alimentación y desechable
              cosali = 0
              cosdes = 0
              If codrec = 0 Then
                 codrec = BuscarCodReceta(desc)
              End If
              Set RS = vg_db.Execute("sgpadm_s_minutadet 1, " & indice & ", '1', " & j & ", 0, 0, 0, '', 0, 0, 0, 0")
              If Not RS.EOF Then
                 RS.Close: Set RS = Nothing
                 vaSpread1.Col = (maxColumna * 6 + 1) + ((i + 3) / 6) ' este estaba
                 vg_db.Execute "sgpadm_iu_minutadet 'M', " & indice & ", '1', " & j & ", " & estser & ", " & codrec & ", " & numrac & ", '" & Mid(desc, 1, 50) & "', " & cosali & ", 0, " & tiprec & ", " & cosdes & ", '" & NameEstManual & "' "
              Else
                 RS.Close: Set RS = Nothing
                 vg_db.Execute "sgpadm_iu_minutadet 'A', " & indice & ", '1', " & j & ", " & estser & ", " & codrec & ", " & numrac & ", '" & Trim(Mid(desc, 1, 50)) & "', " & cosali & ", 0, " & 1 & ", " & cosdes & ", '" & NameEstManual & "' "
              End If
           Else
               vg_db.Execute "sgpadm_d_minutadet 'E1', " & indice & ", '1', " & j & ", 0, 0, 0, '', 0, 0, 0, 0"
           End If
       Next j
    End If
    inddia = inddia + 1
Next i
fecfin = fecha
Picture1.Visible = False: gauge.Visible = False
vaSpread1.Refresh
fg_descarga
CargarAporteCalorico
'vg_db.Execute "sgpadm_iu_deshacer 3,  '" & vg_NUsr & "', " & spid & " "
Toolbar1.Buttons(31).Enabled = False
Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Resume
End Sub

Sub HabilitaCeldaCalorias()
CargarAporteCalorico
vaSpread1.Visible = False
For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
  vaSpread1.Row = 0
  vaSpread1.Col = i + 5
  If vaSpread1.ColHidden = True Then
     vaSpread1.ColHidden = False
     estapo = False
  Else
     vaSpread1.ColHidden = True
     estapo = True
  End If
Next i
vaSpread1.Visible = True
End Sub

Sub DetallePlantillaMinuta()
fg_carga ""
Dim indrow3 As Long, inddia As Long, fecha As String, spid As Long
Dim sw As Boolean: sw = False

SwSalir = 0: maxColumna = 0: indactivo = 0
iblockrow = 0: iblockrow2 = 0: iblockcol = 0: iblockcol2 = 0: SwSalir = 0
aiblockrow = 0: aiblockrow2 = 0: aiblockcol = 0: aiblockcol2 = 0

vg_db.Execute "DELETE paso_servicio WHERE ser_spid = @@spid and ser_usr = '" & vg_NUsr & "'"
'--isel = 0
'-------> Buscar spid
Set RS = vg_db.Execute("SELECT @@spid spid")
If Not RS.EOF Then spid = RS!spid: vg_db.Execute "INSERT INTO paso_servicio VALUES (" & spid & ", '" & vg_NUsr & "', " & Val(vg_codservicio) & ")"
RS.Close: Set RS = Nothing

'-------> Formatear columna
maxColumna = Val(fg_mes(Mid(vg_fecha, 5, 2) & Mid(vg_fecha, 1, 4)))
vaSpread1.MaxRows = 100
vaSpread1.MaxCols = 0: vaSpread1.MaxCols = 6 * maxColumna + 1: vaSpread1.Row = 0
vaSpread1.Col = 1
vaSpread1.ColsFrozen = 1
vaSpread1.VisibleCols = 1
vaSpread1.ColWidth(1) = 15
vaSpread1.text = "Estructura Servicio"
ReDim Preserve vectorcol(0)
For i = 2 To vaSpread1.MaxCols Step 6
   
    
    vaSpread1.Col = i
    vaSpread1.ColWidth(i) = 1.5
    vaSpread1.text = " "
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = i + 1
    vaSpread1.ColWidth(i + 1) = 21
    If i = 2 Then
       ReDim Preserve vectorcol(1)
       vectorcol(1) = 3
       vaSpread1.text = " " & Mid(fg_Fecha_Dia(vg_fecha & Right("0" & (i - 1), 2), 1), 1, 3) & " " & (i - 1) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)
    Else
       vaSpread1.text = " " & Mid(fg_Fecha_Dia(vg_fecha & Right("0" & CLng((i / 6) + 1), 2), 1), 1, 3) & " " & CLng((i / 6) + 1) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)
       ReDim Preserve vectorcol(CLng((i / 6) + 1))
       vectorcol(CLng((i / 6) + 1)) = i + 1
    End If
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = i + 2
    vaSpread1.ColWidth(i + 2) = 6
    vaSpread1.text = "N.Rac."
    vaSpread1.ColHidden = False
   
    vaSpread1.Col = i + 3
    vaSpread1.ColWidth(i + 3) = 9
    vaSpread1.text = "Costo"
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = i + 4
    vaSpread1.text = "Cod. Receta"
    vaSpread1.ColHidden = True
    
    vaSpread1.Col = i + 5
    vaSpread1.ColWidth(i + 3) = 9
    vaSpread1.text = "Calorias"
    vaSpread1.ColHidden = True
    
    For j = 1 To vaSpread1.MaxRows
        vaSpread1.Row = j

        vaSpread1.Col = i
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.text = ""

        vaSpread1.Col = i + 1
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.text = " "

        vaSpread1.Col = i + 2
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.text = " "

        vaSpread1.Col = i + 3
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.text = " "

        vaSpread1.Col = i + 4
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.text = " "

        vaSpread1.Col = i + 5
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignRight
        vaSpread1.text = " " 'aca debe venir el codigo receta

    Next j
    vaSpread1.Row = 0
Next i

vaSpread1.Row = 0
For i = 1 To maxColumna
   vaSpread1.MaxCols = vaSpread1.MaxCols + 1
   vaSpread1.Col = vaSpread1.MaxCols
   vaSpread1.text = "Estado"
   vaSpread1.ColHidden = True
Next i
vaSpread1.MaxCols = vaSpread1.MaxCols + 1
vaSpread1.Col = vaSpread1.MaxCols
vaSpread1.ColWidth(vaSpread1.MaxCols) = 5
vaSpread1.text = "Còd. Est."
vaSpread1.ColHidden = True

vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(0).FillColor  'Amarillo
vaSpread1.Row = -1: vaSpread1.Col = 1
vaSpread1.Font.Bold = True
vaSpread1.Font.Size = 9
vaSpread1.BackColor = Shape1(2).FillColor 'Verde
If vg_Zona = "" Then vg_Zona = 0
j = 0: i = 0: indrow3 = 0 'sgpadm_s_PlanMinutaDetreal 50, 10013,10001,1, 200811,2, 'adm', 66
Set RS = vg_db.Execute("sgpadm_s_PlanMinutaDetreal " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & "," & vg_Zona & ", " & Val(vg_fecha) & ", " & vg_codlpr & ",'" & vg_NUsr & "'," & spid & ",'" & vg_IndpprSelec & "'")
DoEvents
If Not RS.EOF Then
  sw = True   '-------> Calcula el costo plato según su gramaje
   Do While Not RS.EOF
      DoEvents
      j = (((Val(Mid(RS!min_fecmin, 7, 2)) * 6) - 6) + 1) + 1
      vaSpread1.Row = RS!mid_numlin
      If indrow3 < vaSpread1.Row Then indrow3 = vaSpread1.Row
      If RS!ess_codigo <> i Then
         vaSpread1.Col = 1
         If IIf(IsNull(RS!mid_desest), "", RS!mid_desest) <> "" And vg_IndpprSelec = 2 Then
            vaSpread1.text = RS!mid_desest
            
            vaSpread1.Col = vaSpread1.MaxCols - 1
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignCenter
            vaSpread1.text = IIf(IsNull(RS!mid_desest), "", RS!mid_desest)
         
         Else
            vaSpread1.text = Trim(RS!ess_nombre)
            
            vaSpread1.Col = vaSpread1.MaxCols - 1
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignCenter
            vaSpread1.text = Trim(RS!ess_nombre)
         
         End If
         
         vaSpread1.Col = vaSpread1.MaxCols
         vaSpread1.CellType = CellTypeStaticText
         vaSpread1.TypeHAlign = TypeHAlignCenter
         vaSpread1.text = RS!ess_codigo
         i = RS!ess_codigo
        
      End If
      
      vaSpread1.Col = j
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignCenter
      vaSpread1.Value = "R"
      vaSpread1.ForeColor = &HFF&
      vaSpread1.BackColor = &H80FF80
           
      vaSpread1.Col = j + 1
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = Trim(RS!pas_nombre)
                         
      vaSpread1.Col = j + 2
      vaSpread1.CellType = CellTypeNumber
      vaSpread1.TypeNumberDecPlaces = 0
      vaSpread1.TypeIntegerMin = 1
      vaSpread1.TypeIntegerMax = 9999999
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.TypeSpin = False
      vaSpread1.TypeIntegerSpinInc = 1
      vaSpread1.TypeIntegerSpinWrap = False
      vaSpread1.Value = RS!mid_numrac
      vaSpread1.ForeColor = &HFF0000
                       
      vaSpread1.Col = j + 3
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignRight
      precio = Format(IIf(IsNull(RS!pas_prerec) Or Trim(RS!pas_prerec) = 0, 0, RS!pas_prerec), fg_Pict(6, 2))
      vaSpread1.text = precio
      
      vaSpread1.Col = j + 4: vaSpread1.text = RS!pas_codrec & "&" & RS!mid_tiprec & "&;"
          
      vaSpread1.Col = j + 5
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.text = IIf(IsNull(RS!candiet) Or RS!candiet = 0, "", Format(Trim(RS!candiet), fg_Pict(6, 2)))
      
'      vaSpread1.Col = vaSpread1.MaxCols - 1
'      vaSpread1.CellType = CellTypeStaticText
'      vaSpread1.TypeHAlign = TypeHAlignCenter
'      vaSpread1.Text = IIf(IsNull(RS!mid_desest), "", RS!mid_desest)
'      vaSpread1.Visible = True
'      vaSpread1.ColHidden = False
     
      vaSpread1.Col = vaSpread1.MaxCols
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignCenter
      vaSpread1.text = RS!ess_codigo
      
      
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing: fg_descarga
Else
    '-------> Retorna minuta sin precio
    Set RS = vg_db.Execute("sgpadm_s_planifminuta 1, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & "," & vg_Zona & ", " & Val(vg_fecha) & ", 0,0,'" & vg_IndpprSelec & "'")
    DoEvents
    If Not RS.EOF Then '-------> Consulta trae productos sin costo
      sw = True
        Do While Not RS.EOF
              DoEvents
              j = (((Val(Mid(RS!min_fecmin, 7, 2)) * 6) - 6) + 1) + 1
              vaSpread1.Row = RS!mid_numlin
              If indrow3 < vaSpread1.Row Then indrow3 = vaSpread1.Row
              If RS!mid_estser <> i Then
                 vaSpread1.Col = 1
                 vaSpread1.text = RS!ess_nombre
                 vaSpread1.Col = vaSpread1.MaxCols
                 vaSpread1.CellType = CellTypeStaticText
                 vaSpread1.TypeHAlign = TypeHAlignCenter
                 vaSpread1.text = RS!mid_estser
                 i = RS!mid_estser
              End If
              vaSpread1.Col = j
              vaSpread1.CellType = CellTypeStaticText
              vaSpread1.TypeHAlign = TypeHAlignCenter
              vaSpread1.Value = "R"
              vaSpread1.ForeColor = &HFF&
              vaSpread1.BackColor = &H80FF80
                   
              vaSpread1.Col = j + 1
              vaSpread1.CellType = CellTypeStaticText
              vaSpread1.TypeHAlign = TypeHAlignLeft
              vaSpread1.text = Trim(RS!mid_descri)
                                 
              vaSpread1.Col = j + 2
              vaSpread1.CellType = CellTypeNumber
              vaSpread1.TypeNumberDecPlaces = 0
              vaSpread1.TypeIntegerMin = 1
              vaSpread1.TypeIntegerMax = 9999999
              vaSpread1.TypeHAlign = TypeHAlignRight
              vaSpread1.TypeSpin = False
              vaSpread1.TypeIntegerSpinInc = 1
              vaSpread1.TypeIntegerSpinWrap = False
              vaSpread1.Value = RS!mid_numrac
              vaSpread1.ForeColor = &HFF0000
                               
              vaSpread1.Col = j + 3
              vaSpread1.CellType = CellTypeStaticText
              vaSpread1.TypeHAlign = TypeHAlignRight
              vaSpread1.text = Format((IIf(IsNull(RS!mid_cosrec), 0, RS!mid_cosrec) + IIf(IsNull(RS!mid_cosrec), 0, RS!mid_cosrec)), fg_Pict(6, 2))
              
              vaSpread1.Col = j + 4: vaSpread1.text = Val(RS!mid_codrec) & "&" & vg_tiprec & "&;"
              
              vaSpread1.Col = j + 5
              vaSpread1.CellType = CellTypeStaticText
              vaSpread1.TypeHAlign = TypeHAlignRight
              vaSpread1.text = Format(Trim(RS!candiet), fg_Pict(6, 2))
                          
             RS.MoveNext
           Loop
        End If
   RS.Close: Set RS = Nothing: fg_descarga
End If

If Not sw And vg_IndpprSelec = 1 Then    '--->Trae estructura completa si no hay registros de minuta.
   Set RS = vg_db.Execute("sgpadm_s_estservicio 1, " & vg_codservicio & ",''")
   If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
   Do While Not RS.EOF
      vaSpread1.Row = RS!ess_orden
      If indrow3 < vaSpread1.Row Then indrow3 = vaSpread1.Row
      vaSpread1.Col = 1
      vaSpread1.text = RS!ess_nombre
      For i = 2 To vaSpread1.MaxCols Step 6
          vaSpread1.Col = vaSpread1.MaxCols
          vaSpread1.text = RS!ess_codigo
      Next i
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
ElseIf Not sw And vg_IndpprSelec = 2 Then
   indrow3 = 20
End If

If vg_IndpprSelec <> 2 Then
    For i = 3 To (vaSpread1.MaxCols - maxColumna) Step 6
        vaSpread1.Row = 0: vaSpread1.Col = i
        If CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
            Dim fil As Long, Col As Long
            For fil = 1 To (vaSpread1.MaxRows - 1)
                For Col = i - 1 To i + 2
                    vaSpread1.Row = fil: vaSpread1.Col = Col
                    If vaSpread1.CellType = CellTypeNumber Then
                       vaSpread1.CellType = CellTypeStaticText
                       vaSpread1.TypeHAlign = TypeHAlignRight
                    End If
                    vaSpread1.BackColor = Shape1(1).FillColor
                Next Col
            Next fil
        End If
    Next i

   For i = 3 To (vaSpread1.MaxCols - maxColumna) Step 6
       vaSpread1.Row = 0: vaSpread1.Col = i
       If CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
          For fil = 1 To (vaSpread1.MaxRows - 1)
              For Col = i - 1 To i + 4
                  vaSpread1.Row = fil: vaSpread1.Col = Col
                  If vaSpread1.CellType = CellTypeNumber Then
                     vaSpread1.CellType = CellTypeStaticText
                     vaSpread1.TypeHAlign = TypeHAlignRight
                  End If
                  vaSpread1.BackColor = Shape1(1).FillColor
              Next Col
          Next fil
       End If
   Next i
End If

vaSpread1.MaxRows = indrow3 + 1
vaSpread1.Row = vaSpread1.MaxRows
maxfila = vaSpread1.MaxRows
vaSpread1.Col = 1
vaSpread1.text = "Comensales"
vaSpread1.Col = -1: vaSpread1.BackColor = &HE0E0E0
'-------> formatear ultima columna
For i = 2 To (vaSpread1.MaxCols - maxColumna) Step 6
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = i + 2
    vaSpread1.CellType = CellTypeNumber
    vaSpread1.TypeNumberDecPlaces = 0
    vaSpread1.TypeIntegerMin = 1
    vaSpread1.TypeIntegerMax = 9999999
    vaSpread1.TypeHAlign = TypeHAlignRight
    vaSpread1.TypeSpin = False
    vaSpread1.TypeIntegerSpinInc = 1
    vaSpread1.TypeIntegerSpinWrap = False
    vaSpread1.Value = Format(0, fg_Pict(6, 0))
    vaSpread1.ForeColor = &HFF0000
Next i

Set RS = vg_db.Execute("sgpadm_s_planifminuta 2, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & vg_Zona & "," & Val(vg_fecha) & ", 0, 0,'" & vg_IndpprSelec & "'")
DoEvents
If Not RS.EOF Then
   Do While Not RS.EOF
      DoEvents
      j = (((Val(Mid(RS!min_fecmin, 7, 2)) * 6) - 6) + 1) + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = j + 2
      vaSpread1.CellType = CellTypeNumber
      vaSpread1.TypeNumberDecPlaces = 0
      vaSpread1.TypeIntegerMin = 1
      vaSpread1.TypeIntegerMax = 9999999
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.TypeSpin = False
      vaSpread1.TypeIntegerSpinInc = 1
      vaSpread1.TypeIntegerSpinWrap = False
      vaSpread1.Value = IIf(IsNull(RS!min_racteo), 0, RS!min_racteo)
      vaSpread1.ForeColor = &HFF0000
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
Else
   RS.Close: Set RS = Nothing
   Set RS = vg_db.Execute("sgpadm_s_servraciones " & vg_codservicio & "")
   DoEvents
   If Not RS.EOF Then
      Do While Not RS.EOF
         inddia = 1
         For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
             If RS!sra_serdia = IIf(fg_Dia(vg_fecha & fg_pone_cero(inddia, 2)) = 1, 7, Val(fg_Dia(vg_fecha & fg_pone_cero(inddia, 2)) - 1)) Then
                vaSpread1.Col = i + 2
                vaSpread1.CellType = CellTypeNumber
                vaSpread1.TypeNumberDecPlaces = 0
                vaSpread1.TypeIntegerMin = 1
                vaSpread1.TypeIntegerMax = 9999999
                vaSpread1.TypeHAlign = TypeHAlignRight
                vaSpread1.TypeSpin = False
                vaSpread1.TypeIntegerSpinInc = 1
                vaSpread1.TypeIntegerSpinWrap = False
                vaSpread1.Value = RS!Raciones
                vaSpread1.ForeColor = &HFF0000
             End If
             inddia = inddia + 1
         Next i
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
End If

If vg_IndpprSelec <> 2 Then
   For i = 3 To (vaSpread1.MaxCols - maxColumna) Step 6
       vaSpread1.Row = 0: vaSpread1.Col = i
       If CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
          For fil = 1 To (vaSpread1.MaxRows - 1)
              For Col = i - 1 To i + 2
                  vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = Col
                  If vaSpread1.CellType = CellTypeNumber Then
                     vaSpread1.CellType = CellTypeStaticText
                     vaSpread1.TypeHAlign = TypeHAlignRight
                  End If
              Next Col
          Next fil
       End If
   Next i
End If
vaSpread1.Row = 1: vaSpread1.Col = 1
iblockrow = vaSpread1.Row: aiblockrow = vaSpread1.Row
iblockrow2 = vaSpread1.Row: aiblockrow2 = vaSpread1.Row
iblockcol = vaSpread1.Col: aiblockcol = vaSpread1.Col
iblockcol2 = vaSpread1.Col: aiblockcol2 = vaSpread1.Col

End Sub

Sub Calctodia(Row As Long, Col As Long)
Dim x As Long, numrac As Long
Dim cosdia As Double
veccos((Int(Col / 6) + 1), 1) = 0: veccos((Int(Col / 6) + 1), 4) = 0
For x = 1 To (vaSpread1.MaxRows - 1)
    vaSpread1.Row = x
    vaSpread1.Col = Col + 1: numrac = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
    vaSpread1.Col = Col + 2: cosdia = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
    vaSpread1.Col = Col + 3
    If Trim(vaSpread1.text) <> "" And numrac > 0 Then
       vaSpread1.Col = Col + 2: veccos((Int(Col / 6) + 1), 1) = Round(veccos((Int(Col / 6) + 1), 1) + (cosdia * numrac), vg_DCa)
'       vaSpread1.Col = Col + 1: veccos((Int(Col / 5) + 1), 4) = Round(veccos((Int(Col / 5) + 1), 4) + numrac, vg_DCa)
    End If
Next x
vaSpread1.Row = vaSpread1.MaxRows
vaSpread1.Col = Col + 1: veccos((Int(Col / 6) + 1), 4) = Round(veccos((Int(Col / 6) + 1), 4) + IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text), vg_DCa)
End Sub

Sub MostrarCosto(Col As Long)
Dim xcol As Long
Dim toapla As Double, toaesf As Double, toafoo As Double, totdia As Double, totesf As Double, nracre As Double, nracfo As Double, totrac As Double
vaSpread1.Col = Col
xcol = 0
For i = 1 To maxColumna
    If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or vectorcol(i) = (vaSpread1.Col - 2)) Then xcol = vectorcol(i): Exit For
Next i
vaSpread1.Row = 0: vaSpread1.Col = xcol: Frame2(2).Caption = vaSpread1.text: Frame2(3).Caption = "Acumulado hasta " & vaSpread1.text
toapla = 0: toaesf = 0: toafoo = 0: totdia = 0: totesf = 0: nracre = 0: nracfo = 0: totrac = 0
For i = 1 To UBound(veccos)
    If i <= (Int(xcol / 5) + 1) Then
       toapla = CCur(toapla + veccos(i, 1))
       toaesf = CCur(toaesf + veccos(i, 2))
       toafoo = CCur(toafoo + veccos(i, 3))
       nracre = CCur(nracre + veccos(i, 4))
       nracfo = CCur(nracfo + veccos(i, 5))
    End If
    totrac = CCur(totrac + veccos(i, 4))
    totdia = CCur(totdia + veccos(i, 1))
    totesf = CCur(totesf + veccos(i, 2))
Next i
Label1(7).Caption = Format(totdia, fg_Pict(6, 2))
Label1(11).Caption = Format(totesf, fg_Pict(6, 2))
Label1(12).Caption = Format(CCur(totdia + totesf), fg_Pict(6, 2))
If totrac > 0 Then Label1(40).Caption = Format(CCur(totdia / totrac), fg_Pict(6, 2)) Else Label1(40).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(41).Caption = Format(CCur(totesf / totrac), fg_Pict(6, 2)) Else Label1(41).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(8).Caption = Format(CCur((totdia + totesf) / totrac), fg_Pict(6, 2)) Else Label1(8).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(48).Caption = Format(totrac, fg_Pict(6, 2)) Else Label1(48).Caption = Format(0, fg_Pict(6, 2))
Label1(20).Caption = Format(veccos((Int(xcol / 6) + 1), 1), fg_Pict(6, 2))
Label1(21).Caption = Format(veccos((Int(xcol / 6) + 1), 2), fg_Pict(6, 2))
Label1(22).Caption = Format(CCur(veccos((Int(xcol / 6) + 1), 1) + (veccos((Int(xcol / 6) + 1), 2))), fg_Pict(6, 2))
Label1(23).Caption = Format(veccos((Int(xcol / 6) + 1), 3), fg_Pict(6, 2))
Label1(44).Caption = Format(veccos((Int(xcol / 6) + 1), 4), fg_Pict(6, 2))
If veccos((Int(xcol / 6) + 1), 4) > 0 Then Label1(45).Caption = Format(CCur((veccos((Int(xcol / 6) + 1), 1) + (veccos((Int(xcol / 6) + 1), 2))) / veccos((Int(xcol / 6) + 1), 4)), fg_Pict(6, 2)) Else Label1(45).Caption = Format(0, fg_Pict(6, 2))
Label1(46).Caption = Format(veccos((Int(xcol / 6) + 1), 5), fg_Pict(6, 2))
If veccos((Int(xcol / 6) + 1), 5) > 0 Then Label1(47).Caption = Format(CCur(veccos((Int(xcol / 6) + 1), 3) / veccos((Int(xcol / 6) + 1), 5)), fg_Pict(6, 2)) Else Label1(47).Caption = Format(0, fg_Pict(6, 2))
Label1(31).Caption = Format(toapla, fg_Pict(6, 2))
Label1(32).Caption = Format((toaesf), fg_Pict(6, 2))
Label1(33).Caption = Format(CCur(toapla + (toaesf)), fg_Pict(6, 2))
Label1(34).Caption = Format(nracre, fg_Pict(6, 2))
If nracre > 0 Then Label1(35).Caption = Format(CCur((toapla + toaesf) / nracre), fg_Pict(6, 2)) Else Label1(35).Caption = Format(0, fg_Pict(6, 2))
Label1(36).Caption = Format(toafoo, fg_Pict(6, 2))
Label1(37).Caption = Format(nracfo, fg_Pict(6, 2))
If nracfo > 0 Then Label1(38).Caption = Format(CCur(toafoo / nracfo), fg_Pict(6, 2)) Else Label1(38).Caption = Format(0, fg_Pict(6, 2))
End Sub

Sub CargarAporteCalorico()
Dim i, j, racion As Long
Dim caloria, RacxCal, SumCal, TotalAporte As Double
racion = 0: caloria = 0: RacxCal = 0: SumCal = 0: TotalAporte = 0
For i = 2 To vaSpread1.MaxCols Step 6
  SumCal = 0: TotalAporte = 0
  For j = 1 To vaSpread1.MaxRows
    vaSpread1.Col = i: vaSpread1.Row = j
    If vaSpread1.text = "R" Then
      RacxCal = 0: caloria = 0
      vaSpread1.Col = i + 2 'Racion
      If Trim(vaSpread1.text) <> "" Then
         racion = Val(vaSpread1.text)
      Else
         racion = 0
      End If
      
      vaSpread1.Col = i + 5 'Caloria
      
      caloria = Trim(vaSpread1.text)
      If caloria <> "" Then
      RacxCal = (racion * caloria) ' Ración por Caloria
      End If
      SumCal = SumCal + RacxCal
    End If
  Next j
  vaSpread1.Col = i + 2 ' Posición Ración
  vaSpread1.Row = vaSpread1.MaxRows
  If IsNull(vaSpread1.text) = False And vaSpread1.text > "0" Then
     TotalAporte = SumCal / vaSpread1.text
  
     vaSpread1.Col = i + 5: vaSpread1.Row = vaSpread1.MaxRows
     vaSpread1.text = Format(TotalAporte, fg_Pict(6, 2))
     vaSpread1.ForeColor = &HFF0000
'     vaSpread1.Lock = True
  Else
     vaSpread1.Col = i + 5: vaSpread1.Row = vaSpread1.MaxRows
     vaSpread1.text = ""
     vaSpread1.ForeColor = &HFF0000
'     vaSpread1.Lock = True
  End If
Next i
End Sub

Sub CargarCosto()
fg_carga ""
vaSpread1.Col = vaSpread1.ActiveCol
If vaSpread1.Col = 1 Then vaSpread1.Col = 3
Dim cosdia As Double, totdia As Double, totesf As Double, totrac As Double
Dim fecha As Long, xcol As Long, inddia As Long, fecesf As Double, nracre As Long, nracfo As Long
Dim aAp As String
j = 0: fecval = 0: cosdia = 0: totdia = 0: totesf = 0: fecesf = 0: inddia = 1: numrac = 0: totrac = 0
For i = 1 To maxColumna
    If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or vectorcol(i) = (vaSpread1.Col - 2)) Then xcol = vectorcol(i): Exit For
Next i
vaSpread1.Row = 0: vaSpread1.Col = xcol: Frame2(2).Caption = vaSpread1.text: Frame2(3).Caption = "Acumulado hasta " & vaSpread1.text
ReDim veccos(maxColumna, 5)
'------------ Buscar fecha estructura fija
'RS.Open "select max(mif_fecval) as fecval from b_minutafija " & _
'        "where mif_cencos='" & vg_codcasino & "' " & _
'        "and   mif_codreg=" & vg_codregimen & " " & _
'        "and   mif_codser=" & vg_codservicio & "", vg_db, adOpenStatic
'If Not RS.EOF And IsNull(RS!fecval) = False Then fecesf = RS!fecval
'RS.Close: Set RS = Nothing
'------------
'------------ Traer Producto estructura fija, luego actualizar los precio a la ultima toma inventario o biene desde maestro producto
'aAp = Trim(vg_NUsr) & "_tmp_PrecioEstFijaReal"
'fg_CheckTmp aAp
'RS.Open "select distinct mif_codpro as codpro, Round(0,2) as propon into " & aAp & " from b_minutafija where mif_cencos='" & vg_codcasino & "' and mif_codreg=" & vg_codregimen & " and mif_codser=" & vg_codservicio & "", vg_db, adOpenStatic
'Set RS = Nothing
'vg_db.Execute "update " & aAp & " inner join b_tomainv on " & aAp & ".codpro = b_tomainv.tin_codpro set " & aAp & ".propon=b_tomainv.tin_propon " & _
'              "where val(mid(b_tomainv.tin_fectom,1,6))=" & Val(Format(BoM("01/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)), "yyyymm")) & " and b_tomainv.tin_ciemes<>0"
'vg_db.Execute "update " & aAp & " inner join b_productos on " & aAp & ".codpro=b_productos.pro_codigo set " & aAp & ".propon=b_productos.pro_propon where " & aAp & ".propon=0"
'------------
'------------ Calcular costo día planificado & estructura fija & salida
Bar1(0).Min = 0: Bar1(0).Value = 0: Bar1(0).Max = maxColumna: Frame2(4).Visible = True: Bar1(0).Visible = True
For j = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
    Bar1(0).Value = Bar1(0).Value + 1
    If inddia > maxColumna Then Exit Sub
    fecha = Val(vg_fecha) & Right("0" & inddia, 2)
    veccos(inddia, 1) = 0: veccos(inddia, 2) = 0: veccos(inddia, 3) = 0: veccos(inddia, 4) = 0: veccos(inddia, 5) = 0
    For i = 1 To (vaSpread1.MaxRows - 1)
        vaSpread1.Row = i
        vaSpread1.Col = j + 2: numrac = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
        vaSpread1.Col = j + 3: cosdia = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
        vaSpread1.Col = j + 4
        If Trim(vaSpread1.text) <> "" And numrac > 0 Then
           totdia = Round(totdia + (cosdia * numrac), vg_DCa)
           veccos(inddia, 1) = Round(veccos(inddia, 1) + (cosdia * numrac), vg_DCa)
'           veccos(inddia, 4) = Round(veccos(inddia, 4) + numrac, vg_DPr)
        End If
    Next i
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = j + 2
    veccos(inddia, 4) = Round(veccos(inddia, 4) + IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text), vg_DPr)
    totrac = Round(totrac + IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text), vg_DPr)
    If fecesf > 0 Then
''       RS.Open "select b_minutafija.mif_dianro, sum(b_productos.pro_propon*b_minutafija.mif_canpro) as cosesf " & _
''               "from   b_productos, b_minutafija " & _
''               "where  b_minutafija.mif_codpro=b_productos.pro_codigo " & _
''               "and    b_productos.pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') " & _
''               "and    b_minutafija.mif_cencos='" & vg_codcasino & "' " & _
''               "and    b_minutafija.mif_codreg=" & vg_codregimen & " " & _
''               "and    b_minutafija.mif_codser=" & vg_codservicio & " " & _
''               "and    b_minutafija.mif_fecval=" & fecesf & " " & _
''               "and    b_minutafija.mif_dianro=" & fg_NumDia(Trim(Left(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2), Len(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2)) - 2))) & " " & _
''               "group by b_minutafija.mif_dianro", vg_db, adOpenStatic
'       RS.Open "select b_minutafija.mif_dianro, sum(" & aAp & ".propon*b_minutafija.mif_canpro) as cosesf " & _
'               "from   b_productos, b_minutafija, " & aAp & " " & _
'               "where  b_minutafija.mif_codpro=" & aAp & ".codpro " & _
'               "and    " & aAp & ".codpro=b_productos.pro_codigo " & _
'               "and    b_productos.pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') " & _
'               "and    b_minutafija.mif_cencos='" & vg_codcasino & "' " & _
'               "and    b_minutafija.mif_codreg=" & vg_codregimen & " " & _
'               "and    b_minutafija.mif_codser=" & vg_codservicio & " " & _
'               "and    b_minutafija.mif_fecval=" & fecesf & " " & _
'               "and    b_minutafija.mif_dianro=" & fg_NumDia(Trim(Left(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2), Len(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2)) - 2))) & " " & _
'               "group by b_minutafija.mif_dianro", vg_db, adOpenStatic
'       If Not RS.EOF Then totesf = Round(totesf + RS!cosesf, vg_DCa): veccos(inddia, 2) = Round(veccos(inddia, 2) + RS!cosesf, vg_DCa)
'       RS.Close: Set RS = Nothing
    End If
'    RS.Open "select b_totventas.tov_codreg, b_totventas.tov_codser, " & _
'            "sum(IIf(b_totventas.tov_tipdoc='SP',b_detventas.dev_ptotal,'-' & b_detventas.dev_ptotal)) as totdoc " & _
'            "from  b_totventas, b_detventas, b_productos " & _
'            "where b_totventas.tov_rutcli=b_detventas.dev_rutcli " & _
'            "and   b_totventas.tov_tipdoc=b_detventas.dev_tipdoc " & _
'            "and   b_totventas.tov_numdoc=b_detventas.dev_numdoc " & _
'            "and   b_detventas.dev_codmer=b_productos.pro_codigo " & _
'            "and   b_productos.pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') " & _
'            "and   b_totventas.tov_codreg=" & vg_codregimen & " " & _
'            "and   b_totventas.tov_codser=" & vg_codservicio & " " & _
'            "and  (b_totventas.tov_tipdoc='SP' or b_totventas.tov_tipdoc='DP') " & _
'            "and   b_detventas.dev_canmer<>0 " & _
'            "and   b_totventas.tov_estdoc<>'A' " & _
'            "and   b_totventas.tov_fecpro=cdate('" & fg_Ctod1(Val(vg_fecha) & Right("0" & inddia, 2)) & "') " & _
'            "group by b_totventas.tov_codreg, b_totventas.tov_codser", vg_db, adOpenStatic
'    If Not RS.EOF Then veccos(inddia, 3) = Round(veccos(inddia, 3) + RS!totdoc, vg_DCa)
'    RS.Close: Set RS = Nothing
    
'    RS.Open "select sum(mir_nrorac) as mir_nrorac from b_minutaraciones " & _
'            "where  mir_cencos='" & vg_codcasino & "' " & _
'            "and    mir_codreg=" & vg_codregimen & " " & _
'            "and    mir_codser=" & vg_codservicio & " " & _
'            "and    mir_fecmin=" & Val(vg_fecha) & Right("0" & inddia, 2) & "", vg_db, adOpenStatic
'    If Not RS.EOF And Not IsNull(RS!mir_nrorac) Then veccos(inddia, 5) = Round(veccos(inddia, 5) + RS!mir_nrorac, vg_DPr) Else veccos(inddia, 5) = 0
'    RS.Close: Set RS = Nothing
    inddia = inddia + 1
Next j
Frame2(4).Visible = False
Bar1(0).Visible = False
'------------ Fin Calcular costo día
toapla = 0: toaesf = 0: toafoo = 0: numrac = 0: nracfo = 0
For i = 1 To (Int(xcol / 6) + 1)
    toapla = Round(toapla + veccos(i, 1), vg_DCa)
    toaesf = Round(toaesf + veccos(i, 2), vg_DCa)
    toafoo = Round(toafoo + veccos(i, 3), vg_DCa)
    nracre = Round(nracre + veccos(i, 4), vg_DPr)
    nracfo = Round(nracfo + veccos(i, 5), vg_DPr)
Next i
Label1(7).Caption = Format(totdia, fg_Pict(6, 2))
Label1(11).Caption = Format(totesf, fg_Pict(6, 2))
Label1(12).Caption = Format(CCur(totdia + totesf), fg_Pict(6, 2))
If totrac > 0 Then Label1(40).Caption = Format(CCur(totdia / totrac), fg_Pict(6, 2)) Else Label1(40).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(41).Caption = Format(CCur(totesf / totrac), fg_Pict(6, 2)) Else Label1(41).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(8).Caption = Format(CCur((totdia + totesf) / totrac), fg_Pict(6, 2)) Else Label1(8).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(48).Caption = Format(totrac, fg_Pict(6, 2)) Else Label1(48).Caption = Format(0, fg_Pict(6, 2))
Label1(20).Caption = Format(veccos((Int(xcol / 6) + 1), 1), fg_Pict(6, 2))
Label1(21).Caption = Format((veccos((Int(xcol / 6) + 1), 2)), fg_Pict(6, 2))
Label1(22).Caption = Format(CCur(veccos((Int(xcol / 6) + 1), 1) + (veccos((Int(xcol / 6) + 1), 2))), fg_Pict(6, 2))
Label1(23).Caption = Format(veccos((Int(xcol / 6) + 1), 3), fg_Pict(6, 2))
Label1(44).Caption = Format(nracre, fg_Pict(6, 2))
If nracre > 0 Then Label1(45).Caption = Format(CCur((veccos((Int(xcol / 6) + 1), 1) + (veccos((Int(xcol / 6) + 1), 2))) / nracre), fg_Pict(6, 2)) Else Label1(45).Caption = Format(0, fg_Pict(6, 2))
Label1(46).Caption = Format(nracfo, fg_Pict(6, 2))
If nracfo > 0 Then Label1(47).Caption = Format(CCur(veccos((Int(xcol / 6) + 1), 3) / nracfo), fg_Pict(6, 2)) Else Label1(47).Caption = Format(0, fg_Pict(6, 2))
Label1(31).Caption = Format(toapla, fg_Pict(6, 2))
Label1(32).Caption = Format((toaesf), fg_Pict(6, 2))
Label1(33).Caption = Format(CCur(toapla + (toaesf)), fg_Pict(6, 2))
Label1(34).Caption = Format(nracre, fg_Pict(6, 2))
If nracre > 0 Then Label1(35).Caption = Format(CCur((toapla + toaesf) / nracre), fg_Pict(6, 2)) Else Label1(35).Caption = Format(0, fg_Pict(6, 2))
Label1(36).Caption = Format(toafoo, fg_Pict(6, 2))
Label1(37).Caption = Format(nracfo, fg_Pict(6, 2))
If nracfo > 0 Then Label1(38).Caption = Format(CCur(toafoo / nracfo), fg_Pict(6, 2)) Else Label1(38).Caption = Format(0, fg_Pict(6, 2))
indcos = True
fg_descarga
End Sub

Function ValidaMinuta(SubSeg As String, reg As String, Serv As String, TipPlan As String, Zona As String, Fec As String) As Boolean
'*****************---->Validar minuta en uso <---------------------------
'------ Esta funcion crea una tabla temporal concatenando los parametros ingresaods
'------ para la minuta, de esta manera permanece una tabla temporal identificando
'------ que alguien se encuentra conectado a esa minuta, si alguien
'------ mas quiere acceder, se dara un aviso que esta en uso
'------ esta tabla temporal se destruye cuando se cierra este formulario (evento Unload)
'------ y tambien si el usuario cierra la sesion SQL Server la destruye automaticamente.
'----------------------------------------------------------------------
    
    Dim RSTempCheck As New ADODB.Recordset
    
    NameTemp = SubSeg & reg & Serv & TipPlan & Zona & Fec

    Set RSTempCheck = vg_db.Execute("select * from tempdb.dbo.sysobjects where xtype = 'U' and name = '##ValidaMinuta_" & NameTemp & "'")
    
    If RSTempCheck.EOF And RSTempCheck.BOF Then
        Set RSTem = vg_db.Execute("CREATE TABLE ##ValidaMinuta_" & NameTemp & " (usu_codigo VarChar(20))")
        Set RS = vg_db.Execute("INSERT INTO ##ValidaMinuta_" & NameTemp & " (usu_codigo) values ('" & vg_NUsr & "')")
        ValidaMinuta = True
    Else
        ValidaMinuta = False
        Set RS = vg_db.Execute("SELECT usu_codigo from ##ValidaMinuta_" & NameTemp & " ")
        If Not (RS.EOF = True And RS.BOF = True) Then
            RS.MoveFirst
            MsgBox "La minuta con los parametros ingresados, actualmente esta siendo usada por el usuario: '" & UCase(RS!usu_codigo) & "', podra ingresar cuando el usuario termine de trabajar en ella"
        End If
    End If

RSTempCheck.Close
Set RSTempCheck = Nothing
End Function


Sub BlocSoloAcceso()
        ' en caso si tiene solo autorizacion para ver sin modificar ni grabar
        Toolbar1.Buttons(25).Enabled = False ' Actualizar
        Toolbar1.Buttons(4).Enabled = False ' Cortar
        Toolbar1.Buttons(5).Enabled = False ' copiar
        Toolbar1.Buttons(2).Enabled = False
        Toolbar1.Buttons(6).Enabled = False
        'Toolbar1.Buttons(10).Enabled = False ' buscar
        Toolbar1.Buttons(12).Enabled = False ' insertar fila
        Toolbar1.Buttons(13).Enabled = False ' eliminar fila
        Toolbar1.Buttons(15).Enabled = False ' subir fila
        Toolbar1.Buttons(16).Enabled = False ' bajar fila
        Toolbar1.Buttons(19).Enabled = False ' Copiar Minuta
End Sub

Sub CellEditEstruct()
' este procedimiento no se esta usando
' deja editables solo las celdas de la columna uno cuando su contenido
' es distinto de vacio
Dim i As Integer
    If ExraeCodCombo(M_Plami1.Combo2(1)) = 2 Then
        vaSpread1.Col = 1
       
        For i = 1 To vaSpread1.MaxRows
        DoEvents
            vaSpread1.Row = i
            If vaSpread1.text <> "" Then
                vaSpread1.TypeEditCharSet = TypeEditCharSetAlphanumeric
                vaSpread1.CellType = CellTypeEdit
            End If
            DoEvents
        Next i
    End If
End Sub

Sub HabilitaCol(Index As Integer)
'CargarAporteCalorico
'vaSpread1.Visible = False
Select Case Index
Case 0
    For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
    DoEvents
      vaSpread1.Row = 0
      vaSpread1.Col = i + 2
      If vaSpread1.ColHidden = True Then
         vaSpread1.ColHidden = False
         estapo = False
         DoEvents
      Else
         vaSpread1.ColHidden = True
         estapo = True
         DoEvents
      End If
    Next i
Case 1
    For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
    DoEvents
      vaSpread1.Row = 0
      vaSpread1.Col = i + 3
      DoEvents
      If vaSpread1.ColHidden = True Then
         vaSpread1.ColHidden = False
         estapo = False
      Else
         vaSpread1.ColHidden = True
         estapo = True
      End If
      DoEvents
    Next i
Case 2

    'Samuel Quie me quede no vuelve a mostrar columna
      vaSpread1.Row = -1
      vaSpread1.Col = 2
'      vaSpread1.ColHidden = True
'      vaSpread1.ColHidden = False
      
       If vaSpread1.ColHidden = True Then

         vaSpread1.ColHidden = False
'            vaSpread1.ColWidth(2) = 1.5
'            vaSpread1.text = " "
'            vaSpread1.ColHidden = False
'            vaSpread1.CellType = CellTypeStaticText
'            vaSpread1.TypeHAlign = TypeHAlignLeft
'            vaSpread1.text = ""
            
            estapo = False
      Else
         vaSpread1.ColHidden = True
         estapo = True
      End If
    
    For i = 2 To (vaSpread1.MaxCols - maxColumna - 1) Step 6
    DoEvents
      vaSpread1.Row = 0
      vaSpread1.Col = i + 6
      If vaSpread1.ColHidden = True Then
         vaSpread1.ColHidden = False
         estapo = False
      Else
         vaSpread1.ColHidden = True
         estapo = True
      End If
      DoEvents
    Next i
End Select
vaSpread1.Visible = True
End Sub

Private Function ValidaEstructuras() As Boolean
'----- el objetivo de esta funcion es encontrar recetas que no esten asignadas
'----- a una estructura, para lo cual comienza recorriendo desde la columna uno
'----- hacia abajo, preguntando hasta que encuentre recetas sin estructura
'----- por ejemplo si la celda (columna 1, fila 1) esta en blanco y la celda (columna 1, fila 2)
'----- es distinta de vacia, devolvera FALSE
Dim i As Integer, j As Integer


        ValidaEstructuras = True
        xRow = vaSpread1.ActiveRow
        vaSpread1.Row = 1
        vaSpread1.Col = 1
        
        If Trim(vaSpread1.text) <> "" Then ValidaEstructuras = True: Exit Function
        
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Row = i
            
            If Trim(vaSpread1.text) = "" Then
            
                For j = 2 To vaSpread1.MaxCols
                DoEvents
                    vaSpread1.Col = j
                    If Trim(vaSpread1.text) <> "" Then
                        ValidaEstructuras = False
                        DoEvents
                        Exit Function
                    End If
                Next j

            Else
                ValidaEstructuras = True: Exit Function
            End If
            
        Next i
End Function

Function ValidaNombreEstructura(ByVal xRow As Integer, ByVal xSpread As vaSpread) As String
Dim Estruc1 As String, Estruc2 As String
Dim RSEst As New ADODB.Recordset
ValidaNombreEstructura = ""
Dim xResCol As Long, xResRow2 As Long


    xResCol = xSpread.Col
    xResRow2 = xSpread.Row
    
    xSpread.Row = xRow
    xSpread.Col = 1
    Estruc2 = xSpread.text
    
    If Estruc2 = "" Then ValidaNombreEstructura = "": Exit Function
    
    xSpread.Col = xSpread.MaxCols
    Estruc1 = G_Proc.fg_ExtraeServicio(vg_codservicio, xSpread.text)
    
    If Trim(Estruc1) = Trim(Estruc2) Then
        ValidaNombreEstructura = ""
    Else
        ValidaNombreEstructura = Estruc2
    End If
    
    xSpread.Col = xResCol
    xSpread.Row = xResRow2
    
End Function

Sub DeshacerInto(ByVal xcol As Long, ByVal xRow As Long, ByVal Spread As vaSpread, ByVal UltimoCambio As Long)
    'DeshacerFormatoCeldas
    DoEvents
    'fg_carga ""
    SpreadClon.vaSpread1.Col = xcol
    SpreadClon.vaSpread1.Row = xRow
    vg_db.Execute "sgpadm_iu_deshacer 1,  '" & vg_NUsr & "', " & spid & ", " & xcol & ", " & xRow & ", '" & Trim(SpreadClon.vaSpread1.text) & "', " & UltimoCambio & ", 0, 0, 0, " _
    & SpreadClon.vaSpread1.CellType _
    & ", " & SpreadClon.vaSpread1.TypeHAlign _
    & ", " & SpreadClon.vaSpread1.ForeColor _
    & " , " & SpreadClon.vaSpread1.BackColor _
    & ", " & SpreadClon.vaSpread1.TypeNumberDecPlaces _
    & ", " & SpreadClon.vaSpread1.TypeIntegerMin _
    & ", " & SpreadClon.vaSpread1.TypeIntegerMax _
    & ", " & IIf(SpreadClon.vaSpread1.TypeSpin = False, 0, 1) _
    & ", " & SpreadClon.vaSpread1.TypeIntegerSpinInc _
    & ", " & IIf(SpreadClon.vaSpread1.TypeIntegerSpinWrap = False, 0, 1) _
    & ", " & IIf(SpreadClon.vaSpread1.AutoSize = False, 0, 1) _
    & ", " & SpreadClon.vaSpread1.BackColorStyle _
    & ", " & SpreadClon.vaSpread1.BorderStyle _
    & ", " & IIf(SpreadClon.vaSpread1.Enabled = False, 0, 1) _
    & ", '" & SpreadClon.vaSpread1.Font _
    & "', " & IIf(SpreadClon.vaSpread1.FontBold = False, 0, 1) _
    & ", " & IIf(SpreadClon.vaSpread1.FontItalic = False, 0, 1) _
    & ", " & SpreadClon.vaSpread1.FontSize _
    & ", " & IIf(SpreadClon.vaSpread1.FontStrikethru = False, 0, 1) & " "

    

    Spread.Col = xcol
    Spread.Row = xRow
    'vg_db.Execute "sgpadm_iu_deshacer 1,  '" & vg_NUsr & "', " & spid & ", " & xCol & ", " & xRow & ", '" & Spread.Text & "' "
    SpreadClon.vaSpread1.text = Trim(Spread.text)
    SpreadClon.vaSpread1.CellType = Spread.CellType
    SpreadClon.vaSpread1.TypeHAlign = Spread.TypeHAlign
    SpreadClon.vaSpread1.ForeColor = Spread.ForeColor
    SpreadClon.vaSpread1.BackColor = Spread.BackColor
    SpreadClon.vaSpread1.TypeNumberDecPlaces = Spread.TypeNumberDecPlaces
    SpreadClon.vaSpread1.TypeIntegerMin = Spread.TypeIntegerMin
    SpreadClon.vaSpread1.TypeIntegerMax = Spread.TypeIntegerMax
    SpreadClon.vaSpread1.TypeSpin = Spread.TypeSpin
    SpreadClon.vaSpread1.TypeIntegerSpinInc = Spread.TypeIntegerSpinInc
    SpreadClon.vaSpread1.TypeIntegerSpinWrap = Spread.TypeIntegerSpinWrap
    SpreadClon.vaSpread1.AutoSize = Spread.AutoSize
    SpreadClon.vaSpread1.BackColorStyle = Spread.AutoSize
    SpreadClon.vaSpread1.BorderStyle = Spread.BorderStyle
    SpreadClon.vaSpread1.Enabled = Spread.Enabled
    SpreadClon.vaSpread1.Font = Spread.Font
    SpreadClon.vaSpread1.FontBold = Spread.FontBold
    SpreadClon.vaSpread1.FontItalic = Spread.FontItalic
    SpreadClon.vaSpread1.FontSize = Spread.FontSize
    SpreadClon.vaSpread1.FontStrikethru = Spread.FontStrikethru
    
    If Toolbar1.Buttons(31).Enabled = False Then Toolbar1.Buttons(31).Enabled = True

    DoEvents
    
    'fg_descarga
End Sub

Sub DeshacerModFile(ByVal xRowActive As Long, ByVal xCantRow As Long, ByVal UltimoCambio As Long, ByVal DesType As DeshacerType)
    'DeshacerFormatoCeldas
    DoEvents

    vg_db.Execute "sgpadm_iu_deshacer 1,  '" & vg_NUsr & "', " & spid & ", 0, 0, '', " & UltimoCambio & ", " & DesType & ", " & xRowActive & ", " & xCantRow & " "

  
    
    SpreadClon.vaSpread1.MaxRows = vaSpread1.MaxRows
    SpreadClon.vaSpread1.InsertRows xRowActive, xCantRow
    
    If Toolbar1.Buttons(31).Enabled = False Then Toolbar1.Buttons(31).Enabled = True

    DoEvents
End Sub

Private Sub DeshacerSelect()
On Error GoTo Man_Error
fg_carga ""
    Set RS = vg_db.Execute("sgpadm_iu_deshacer 2, '" & vg_NUsr & "', " & spid & ", 0, 0, '', " & DeshacerUltimoCammbio & " ")
    
    If Not (RS.EOF And RS.BOF) Then
    RS.MoveFirst
            
           
            Do While Not RS.EOF = True
            Select Case RS!des_tipo
                Case 0
                    
                    DoEvents
                        vaSpread1.Col = RS!des_colum
                        vaSpread1.Row = RS!des_fila
                        vaSpread1.text = Trim(RS!des_dato)
                    vaSpread1.CellType = RS!CellType
                    vaSpread1.ForeColor = RS!ForeColor
                    vaSpread1.BackColor = RS!BackColor
                    vaSpread1.TypeNumberDecPlaces = RS!TypeNumberDecPlaces
                    vaSpread1.TypeIntegerMin = RS!TypeIntegerMin
                    vaSpread1.TypeIntegerMax = RS!TypeIntegerMax
                    vaSpread1.TypeSpin = RS!TypeSpin
                    vaSpread1.TypeIntegerSpinInc = RS!TypeIntegerSpinInc
                    vaSpread1.TypeIntegerSpinWrap = RS!TypeIntegerSpinWrap
                    vaSpread1.AutoSize = RS!AutoSize
                    vaSpread1.BackColorStyle = RS!BackColorStyle
                    vaSpread1.BorderStyle = RS!BorderStyle
                    vaSpread1.Enabled = RS!Enabled
                    vaSpread1.Font = RS!Font
                    vaSpread1.FontBold = RS!FontBold
                    vaSpread1.FontItalic = RS!FontItalic
                    vaSpread1.FontSize = RS!FontSize
                    vaSpread1.FontStrikethru = RS!FontStrikethru
                    vaSpread1.TypeHAlign = RS!TypeHAlign
                        
                        
                        SpreadClon.vaSpread1.Col = RS!des_colum
                        SpreadClon.vaSpread1.Row = RS!des_fila
                        SpreadClon.vaSpread1.text = Trim(RS!des_dato)
                        vaSpread1.CellType = RS!CellType
                    SpreadClon.vaSpread1.TypeHAlign = RS!TypeHAlign
                    SpreadClon.vaSpread1.ForeColor = RS!ForeColor
                    SpreadClon.vaSpread1.BackColor = RS!BackColor
                    SpreadClon.vaSpread1.TypeNumberDecPlaces = RS!TypeNumberDecPlaces
                    SpreadClon.vaSpread1.TypeIntegerMin = RS!TypeIntegerMin
                    SpreadClon.vaSpread1.TypeIntegerMax = RS!TypeIntegerMax
                    SpreadClon.vaSpread1.TypeSpin = RS!TypeSpin
                    SpreadClon.vaSpread1.TypeIntegerSpinInc = RS!TypeIntegerSpinInc
                    SpreadClon.vaSpread1.TypeIntegerSpinWrap = RS!TypeIntegerSpinWrap
                    SpreadClon.vaSpread1.AutoSize = RS!AutoSize
                    SpreadClon.vaSpread1.BackColorStyle = RS!BackColorStyle
                    SpreadClon.vaSpread1.BorderStyle = RS!BorderStyle
                    SpreadClon.vaSpread1.Enabled = RS!Enabled
                    SpreadClon.vaSpread1.Font = RS!Font
                    SpreadClon.vaSpread1.FontBold = RS!FontBold
                    SpreadClon.vaSpread1.FontItalic = RS!FontItalic
                    SpreadClon.vaSpread1.FontSize = RS!FontSize
                    SpreadClon.vaSpread1.FontStrikethru = RS!FontStrikethru
                    DoEvents
                Case 1 '---> Se Agregó
                    vaSpread1.DeleteRows RS!des_activfil, RS!des_cantfil
                    SpreadClon.vaSpread1.DeleteRows RS!des_activfil, RS!des_cantfil
                    vaSpread1.MaxRows = vaSpread1.MaxRows - RS!des_cantfil
                    SpreadClon.vaSpread1.MaxRows = vaSpread1.MaxRows
                Case 2 '---> Se Eliminó
                    vaSpread1.MaxRows = vaSpread1.MaxRows + RS!des_cantfil
                    SpreadClon.vaSpread1.MaxRows = vaSpread1.MaxRows
                    vaSpread1.InsertRows RS!des_activfil, RS!des_cantfil
                    SpreadClon.vaSpread1.InsertRows RS!des_activfil, RS!des_cantfil
                End Select
            
            RS.MoveNext
            Loop
                'RS.MoveFirst
                'DeshacerFormatoCeldas RS!des_colum, RS!des_fila
    Else
        Toolbar1.Buttons(31).Enabled = False
    End If
    vg_db.Execute ("sgpadm_iu_deshacer 5, '" & vg_NUsr & "', " & spid & ", 0, 0, '', " & DeshacerUltimoCammbio & " ")
    RS.Close
    Set RS = Nothing
fg_descarga
Man_Error:
If 3704 = Err.Number Or 0 = Err.Number Or 3265 = Err.Number Then
    On Error Resume Next
Else
    MsgBox Err.Number & " - " & Err.Description, vbCritical
End If

fg_descarga

End Sub


Function DeshacerUltimoCammbio() As Long
    Set RS = vg_db.Execute("sgpadm_iu_deshacer 4,  '" & vg_NUsr & "', " & spid & " ")
    If Not (RS.EOF And RS.BOF) Then
        DeshacerUltimoCammbio = RS.Fields(0)
    Else
        DeshacerUltimoCammbio = 0
    End If
End Function




Sub DeshacerIntoMatriz(ByVal ColIniOri As Long, ByVal FilIniOri As Long, ByVal ColFinOri As Long, ByVal FilFinOri As Long, ByVal ColIniDes As Long, ByVal FilIniDes As Long, Optional LastModifi As Long)
Dim f As Long, c As Long, LastChange As Long, FilFinDes As Long, ColFinDes  As Long
'vaSpread1.Visible = False
'Id que identifica todo este cambio como un solo movimiento para deshacer
If LastModifi = 0 Then
    LastChange = DeshacerUltimoCammbio + 1
Else
    LastChange = LastModifi
End If

'Ciclo que guarda celdas de origen antes de ser cortadas
    For c = ColIniOri To ColFinOri
        For f = FilIniOri To FilFinOri
            DeshacerInto c, f, vaSpread1, LastChange
        Next f
    Next c

ColFinDes = ColIniDes + ((ColFinOri - ColIniOri))
FilFinDes = FilIniDes + ((FilFinOri - FilIniOri))

'Ciclo que guarda las celdas de destino antes de ser pegada informacion en ellas
    For c = ColIniDes To ColFinDes
        For f = FilIniDes To FilFinDes
            
            DeshacerInto c, f, vaSpread1, LastChange
        Next f
    Next c
'vaSpread1.Visible = True
End Sub


Sub DeshacerMatriz(ByVal ColIniOri As Long, ByVal FilIniOri As Long, ByVal ColFinOri As Long, ByVal FilFinOri As Long, ByVal LastModifi As Long)
Dim f As Long, c As Long ', LastChange As Long
'vaSpread1.Visible = False
'Id que identifica todo este cambio como un solo movimiento para deshacer
'LastChange = DeshacerUltimoCammbio + 1

'Ciclo que guarda celdas de origen antes de ser cortadas
    For c = ColIniOri To ColFinOri
    DoEvents
        For f = FilIniOri To FilFinOri
            DoEvents
            DeshacerInto c, f, vaSpread1, LastModifi
        Next f
    Next c
'vaSpread1.Visible = True
End Sub

Private Sub DeshacerDelFile(ByVal ActiveFile As Long, ByVal CantFile As Long)
Dim ret As Integer, ret2 As Integer, Colini As Integer, IdChangeMod As Long
Dim srow As Long

'vaSpread1.Visible = False
IdChangeMod = DeshacerUltimoCammbio + 1
ret = 0
ret2 = 0
Colini = -4
 vaSpread1.Col = 1
 vaSpread1.Row = 1
 For srow = ActiveFile To (ActiveFile + CantFile) - 1
     DoEvents
                        
            Do While ret <> -1
                ret = vaSpread1.SearchRow(srow, ret, vaSpread1.MaxCols - 1, " ", SearchFlagsValue)
                
                ret2 = vaSpread1.SearchRow(srow, ret2, vaSpread1.MaxCols - 1, "", SearchFlagsValue)
                
                If ret = -1 And ret2 <> -1 Then ret = ret2
                If ret2 = -1 And ret <> -1 Then ret2 = ret
                
                If ret > ret2 Then
                    ret = ret2
                Else
                    ret2 = ret
                End If
                
                'vaSpread1.SetActiveCell ret, srow
                If (ret - Colini) > 1 Then
                    If Colini = -4 Then Colini = 1
                    'If Colini > 0 Then DeshacerMatriz Colini, srow, ret, srow, IdChangeMod
                End If
                Colini = ret
    
            DoEvents
            Loop
        
        ret = 0
        ret2 = 0
        Colini = -4
     Next srow
'vaSpread1.Visible = True
End Sub

Sub AddEstructuraMenu(ByVal CodigoEst As Long, ByVal NombreEst As String)
On Error GoTo ErrSub
    Dim i As Long
            Load Estructura1.Item(Estructura1.count)
            Load Estructura2.Item(Estructura2.count)

            Estructura1.Item(Estructura1.count - 1).Caption = NombreEst
            Estructura1.Item(Estructura1.count - 1).HelpContextID = CodigoEst
            Estructura1.Item(Estructura1.count - 1).Enabled = True
            Estructura1.Item(Estructura1.count - 1).Visible = True

            Estructura2.Item(Estructura2.count - 1).Caption = NombreEst
            Estructura2.Item(Estructura2.count - 1).HelpContextID = CodigoEst
            Estructura2.Item(Estructura2.count - 1).Enabled = True
            Estructura2.Item(Estructura2.count - 1).Visible = True
            'Estructura2.Item(Estructura2.count - 1) = True

Exit Sub
ErrSub:
    
    On Local Error Resume Next
    MsgBox Err.Description, vbCritical

End Sub

Function EstructuraSuperior(ByVal Spread As vaSpread, ByVal Fila As Long) As String
Dim xRespRow As Long, xRespCol As Long
Dim x1 As Long

EstructuraSuperior = ""
xRespRow = Spread.Row
xRespCol = Spread.Col

        'AvisoEst = False
        Spread.Row = Fila
        Spread.Col = Spread.MaxCols - 1
        For x1 = Spread.Row To 1 Step -1
            Spread.Row = x1
            If Trim(Spread.text) <> "" Then
                EstructuraSuperior = Spread.text
                Exit For
            End If
            
        Next x1

        
Spread.Row = xRespRow
Spread.Col = xRespCol
End Function


Sub ActualizaEstructuraInferior(ByVal Spread As vaSpread, ByVal NameEstruct As String, Optional UltimoCambio As Long)
Dim xRespRow As Long, xRespCol As Long, EstructuraAnterior As String
Dim x1 As Long


xRespRow = Spread.Row
xRespCol = Spread.Col



     

        Spread.Col = Spread.MaxCols - 1
        EstructuraAnterior = Spread.text
        For x1 = Spread.Row To Spread.MaxRows - 1
            Spread.Row = x1
            If Trim(Spread.text) = "" Or Spread.Row = Spread.MaxRows Or Trim(EstructuraAnterior) <> Trim(Spread.text) Then Exit For
            Spread.text = NameEstruct
            'DeshacerInto vaSpread1.Col, vaSpread1.Row, vaSpread1, DeshacerUltimoCammbio
        Next x1

        
Spread.Row = xRespRow
Spread.Col = xRespCol
End Sub




Sub DesqloqSubMenu(OpcioneMenu As String)
Dim iA As Integer
    For iA = 1 To Estructura2.count - 1
            If Trim(Estructura2.Item(iA).Caption) = Trim(OpcioneMenu) Then
                Estructura1.Item(iA).Enabled = True
                Estructura2.Item(iA).Enabled = True
            End If
    Next iA
End Sub


Sub Deshacer(StrRec As Variant)
'load in file
Dim ret As Integer
Screen.MousePointer = 11
ret = vaSpread1.LoadFromFile(LCase(App.Path) & "\" & StrRec)
If Dir(LCase(App.Path) & "\" & StrRec) <> "" Then Kill LCase(App.Path) & "\" & StrRec
CorDes = CorDes - 1
Screen.MousePointer = 0
'vaSpread1.ProcessTab = True

'Dim i As Long
'Dim nomrec As String, racion As Long, cosrec As Double, codrec As Long, caloria As Double, ifil As Long, icol As Long
'If Len(StrRec) <> 0 Then
'   Do While InStr(StrRec, ";") <> 0
'      nomrec = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
'      StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
'      ifil = Val(Mid(StrRec, 1, InStr(StrRec, ";") - 1)) ': StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
'      StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
'      icol = Val(Mid(StrRec, 1, InStr(StrRec, ";") - 1))
'      StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
'      racion = Val(Mid(StrRec, 1, InStr(StrRec, ";") - 1))
'      StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
'      cosrec = Val(Mid(StrRec, 1, InStr(StrRec, ";") - 1))
'      StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
'      codrec = Val(Mid(StrRec, 1, InStr(StrRec, ";") - 1))
'      StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
'      caloria = Val(Mid(StrRec, 1))
'   Loop
'   vaSpread1.Row = ifil
'   vaSpread1.Col = icol
'   If Trim(nomrec) = "Insertar" Then
'      vaSpread1.DeleteRows ifil, (icol - ifil) + 1
'      vaSpread1.MaxRows = vaSpread1.MaxRows - ((icol - ifil) + 1)
'   ElseIf Trim(nomrec) = "Eliminar" Then
'     'load in file
'     Dim ret As Integer
'     Screen.MousePointer = 11
'     ret = vaSpread1.LoadFromFile(LCase(App.Path) & "\spreadss.ss6")
'     Screen.MousePointer = 0
'     vaSpread1.ProcessTab = True
'   ElseIf Trim(nomrec) = "" Then
'      vaSpread1.Col = icol - 1
'      vaSpread1.Action = 3
'      For i = icol To icol + 4
'      vaSpread1.Col = i
'      vaSpread1.Action = 3
'      Next i
'   Else
'      vaSpread1.Col = icol
'      vaSpread1.Text = nomrec
'      vaSpread1.Col = icol + 1
'      vaSpread1.Text = racion
'      vaSpread1.Col = icol + 2
'      vaSpread1.Text = cosrec
'      vaSpread1.Col = icol + 3
'      vaSpread1.Text = codrec
'      vaSpread1.Col = icol + 4
'      vaSpread1.Text = caloria
'   End If
'   vaSpread1.SetActiveCell icol, ifil ': vaSpread1.SetFocus
'End If
End Sub

Sub GrabarCambios(ifil As Long, icol As Long, estado As String)
Dim ret
CorDes = CorDes + 1
ret = vaSpread1.SaveToFile(LCase(App.Path) & "\" & "spread" & vg_NUsr & CorDes & ".ss6", False)

'Dim nomrec As String, racion As Long, cosrec As Double, codrec As Long, caloria As Double
'vaSpread1.Row = ifil
'vaSpread1.Col = icol
'nomrec = IIf(estado = "Insertar", "Insertar", vaSpread1.Text)
'vaSpread1.Col = j + 1
'racion = Val(vaSpread1.Text)
'vaSpread1.Col = j + 2
'cosrec = Val(vaSpread1.Text)
'vaSpread1.Col = j + 3
'codrec = Val(vaSpread1.Text)
'vaSpread1.Col = j + 4
'caloria = Val(vaSpread1.Text)
'Combo1.AddItem nomrec & Space(150) & ";" & ifil & ";" & icol & ";" & racion & ";" & cosrec & ";" & codrec & ";" & caloria
''    Combo1.AddItem nomrec & ";" & vaSpread1.Row & ";" & j & ";" & racion & ";" & cosrec & ";" & codrec & ";" & caloria
Toolbar1.Buttons(31).Visible = True
Toolbar1.Buttons(31).Enabled = True
End Sub

