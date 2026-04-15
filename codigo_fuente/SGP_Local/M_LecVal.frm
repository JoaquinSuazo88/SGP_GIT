VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_LecVal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lectura Vales"
   ClientHeight    =   7710
   ClientLeft      =   5370
   ClientTop       =   2055
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8775
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   5445
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   5445
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   5445
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5055
         Left            =   600
         TabIndex        =   4
         Top             =   1800
         Width           =   7725
         _Version        =   393216
         _ExtentX        =   13626
         _ExtentY        =   8916
         _StockProps     =   64
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   0
         SpreadDesigner  =   "M_LecVal.frx":0000
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   2580
         TabIndex        =   10
         Top             =   1245
         Width           =   5430
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   2580
         TabIndex        =   9
         Top             =   765
         Width           =   5430
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   2580
         TabIndex        =   8
         Top             =   285
         Width           =   5430
      End
      Begin VB.Label Label3 
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
         Height          =   225
         Index           =   1
         Left            =   840
         TabIndex        =   7
         Top             =   1320
         Width           =   1560
      End
      Begin VB.Label Label3 
         Caption         =   "Regimen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   840
         TabIndex        =   6
         Top             =   840
         Width           =   1560
      End
      Begin VB.Label Label3 
         Caption         =   "Punto Atención"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   1560
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_LecVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim Msgtitulo  As String
Dim est As Boolean
Dim strString As String

Private Sub Combo1_Click(Index As Integer)
Select Case Index
Case 0
    '-------> Cargar Combo Regimen
    vg_ptoate = fg_codigocbo(Combo1, 0, 10, 0)
    CargarDatoCombo Combo1, 1, "a_pto_lectura_vales_pto_servicio", strString, "LecReg", "N"
    If Combo1(1).listcount = 1 Then
       Combo1(1).ListIndex = 0
    End If
    vg_ptoate = ""
Case 1
    '-------> Cargar Combo Servicio
    vg_ptoate = fg_codigocbo(Combo1, 0, 10, 0)
    vg_codreg = fg_codigocbo(Combo1, 1, 10, 0)
    CargarDatoCombo Combo1, 2, "a_pto_lectura_vales_pto_servicio", strString, "LecSer", "N"
    If Combo1(2).listcount = 1 Then
       Combo1(2).ListIndex = 0
    End If
    vg_ptoate = ""
    vg_codreg = ""
End Select
If Combo1(0).ListIndex > -1 And Combo1(1).ListIndex > -1 And Combo1(2).ListIndex > -1 Then
   MoverDatosGrillas
End If
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Msgtitulo = "Lectura de Vales"
fg_centra Me
modo = ""
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 3, modo
Dim dwLen As Long

dwLen = MAX_COMPUTERNAME_LENGTH + 1
strString = String(dwLen, "X")
GetComputerName strString, dwLen
strString = Left(strString, dwLen)
'-------> Cargar Combo Punto de Atención
CargarDatoCombo Combo1, 0, "a_pto_lectura_vales_pto_atencion", strString, "PunVen", "N"
If Combo1(0).listcount = 1 Then
   Combo1(0).ListIndex = 0
End If
'vaSpread1.MaxRows = 0
End Sub

Sub MoverDatosGrillas()
Dim RS As New ADODB.Recordset
Dim codreg As Long
Dim codser As Long
Dim codate As Long
est = True
codate = fg_codigocbo(Combo1, 0, 10, 0)
codreg = fg_codigocbo(Combo1, 1, 10, 0)
codser = fg_codigocbo(Combo1, 2, 10, 0)
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
'RS.Open RutinaLectura.DetalleLectura(1, MuestraCasino(1), codreg, codser, codate, "", 0), vg_db, adOpenStatic
Set RS = vg_db.Execute("sgp_Sel_DetalleLectura '" & MuestraCasino(1) & "', " & codreg & ", " & codser & ", " & codate & "")
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   vaSpread1.Lock = True
   vaSpread1.Value = IIf(IsNull(RS!codigobarra), "", RS!codigobarra)
       
   vaSpread1.Col = 2
   vaSpread1.CellType = CellTypeStaticText
   vaSpread1.Value = RS!FechaHoravale
       
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread1.MaxRows = vaSpread1.MaxRows + 1
vaSpread1.Row = vaSpread1.MaxRows
vaSpread1.Visible = True
vaSpread1.Col = 1
vaSpread1.SetActiveCell 1, vaSpread1.MaxRows
If vaSpread1.MaxRows > 0 And vaSpread1.Visible = True Then vaSpread1.SetFocus
est = False
Exit Sub
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 1 'Incluir
Case 3 'Alterar
Case 5 'Borrar
Case 7 'Actualizar Lista
Case 10 'Cancelar
Case 12
Case 15
'    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
'    I_PtoLecturaVales Text1.text
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
est = False
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
If NewCol <> 2 And Row = vaSpread1.MaxRows Then
    GrabaRegistro Row
Else
    vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus
End If
End Sub

Private Sub GrabaRegistro(Fila As Long)
On Error GoTo Man_Error
If est Then Exit Sub
Dim RS As New ADODB.Recordset
Dim i As Long
Dim codreg As Long
Dim cadena As String
Dim codser As Long
Dim codate As Long
Dim CodBar As String
Dim CodigoCliente As String
Dim contcli As Long
Dim PosInicialPin As Long
Dim PosInicialFecha As Long
Dim PosTipo As Long
Dim LargoPin As Long
Dim LargoFecha As Long
Dim LargoTipo As Long
Dim sql1 As String
Dim fecgra As Date
Dim horgra As Timer
codate = fg_codigocbo(Combo1, 0, 10, 0)
codreg = fg_codigocbo(Combo1, 1, 10, 0)
codser = fg_codigocbo(Combo1, 2, 10, 0)

    vaSpread1.Row = Fila
    vaSpread1.Col = 1: CodBar = LimpiaDato(vaSpread1.text)
    If Trim(CodBar) = "" Then Exit Sub
    '-------> Validar a_par_codigo_barra_cas
'    Set RS = vg_db.Execute(RutinaLectura.ParametroCodigoBarra(2, MuestraCasino(1), ""))
    Set RS = vg_db.Execute("sgp_Sel_ParametroCodigoBarra '" & MuestraCasino(1) & "'")
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS!cbar_posinicial) Or RS!cbar_posinicial < 1 Then RS.Close: Set RS = Nothing: MsgBox "Posicion incial del atributo codigo barra, esta con valor cero o nulo. Proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Row = Fila: vaSpread1.Col = Col: vaSpread1.text = "": vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
          If IsNull(RS!cbar_largo) Or RS!cbar_largo < 1 Then RS.Close: Set RS = Nothing: MsgBox "Largo del atributo codigo barra, esta en valor cero o nulo. Proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Row = Fila: vaSpread1.Col = Col: vaSpread1.text = "": vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
          If RS!atr_codigo_barra = 3 Then PosInicialPin = RS!cbar_posinicial: LargoPin = RS!cbar_largo
          If RS!atr_codigo_barra = 2 Then PosInicialFecha = RS!cbar_posinicial: LargoFecha = RS!cbar_largo
          If RS!atr_codigo_barra = 1 Then PosInicialtipo = RS!cbar_posinicial: LargoTipo = RS!cbar_largo
          RS.MoveNext
       Loop
     End If
     RS.Close: Set RS = Nothing
    CodigoCliente = ""
    If MuestraCasino(1) = "23260" Then
       Set RS = vg_db.Execute("sgp_Sel_LargoMaximoParCodigoBarra '" & MuestraCasino(1) & "'")
       If RS.EOF Then
          RS.Close: Set RS = Nothing
          vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.text = "": MsgBox "No existe datos, para validar el largo del código barra...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
       Else
          If Len(CodBar) > RS!LargoMaximo Then RS.Close: Set RS = Nothing: vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.text = "": MsgBox "Código barra no se encuentra configurado ...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
       End If
       RS.Close: Set RS = Nothing
       
       Set RS = vg_db.Execute("sgp_Sel_ClienteTipoVale")
       If Not RS.EOF Then
          Do While Not RS.EOF
             If PosInicialPin = 0 Then RS.Close: Set RS = Nothing: vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.text = "": MsgBox "Posicion incial del atributo codigo barra, esta con valor cero o nulo. Proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
             If LargoPin = 0 Then RS.Close: Set RS = Nothing: vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.text = "": MsgBox "Largo del atributo codigo barra, esta en valor cero o nulo. Proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
             If RS!cli_tipo_vale = Mid(CodBar, PosInicialPin, LargoPin) Then
                CodigoCliente = RS!cli_codigo
                Exit Do
             End If
             RS.MoveNext
          Loop
          If Trim(Mid(CodBar, PosInicialPin, LargoPin)) = "" Then
             RS.Close: Set RS = Nothing
             MsgBox "Posicion incial y final del atributo del tipo codigo barra, esta con valor cero o nulo. Proceso cancelado......", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
          End If
          If Trim(CodigoCliente) = "" Then
             RS.Close: Set RS = Nothing
             MsgBox "Cliente no tiene asignado tipo vale...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
          End If
       Else
          RS.Close: Set RS = Nothing
          If PosInicialPin = 0 Then
             MsgBox "No existe parametro código barra...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
          ElseIf CodigoCliente = "" Then
             MsgBox "Cliente no existe o bien no esta definido los parametro codigo barra...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
          End If
       End If
       RS.Close: Set RS = Nothing
    
       '-------> validar Codigo barra si existe en la base
       RS.Open ("sgp_Sel_ValidarLecturaCodigoBarra '" & MuestraCasino(1) & "', '" & CodBar & "'"), vg_db, adOpenStatic
       If Not RS Is Nothing Then
       If Not RS.EOF Then
          If Trim(RS(0)) <> "" Then
              RS.Close: Set RS = Nothing
              MsgBox "Código de Barra ya fue ingresado. Proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo
              vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.text = ""
              vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus
              Exit Sub
          End If
       End If
       RS.Close: Set RS = Nothing
       End If
    Else
       RS.Open RutinaLectura.Personal(1, "", ""), vg_db, adOpenStatic
       If RS.EOF Then
          RS.Close: Set RS = Nothing
          vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.text = ""
          MsgBox "Existen más de un cliente a facturar y no esta definido en mantenedor personal. Proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo:  vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
       End If
       RS.Close: Set RS = Nothing
       '-------> Traer codigo cliente
       If PosInicialPin = 0 Then vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.text = "": MsgBox "Posicion incial del atributo codigo barra, esta con valor cero o nulo. Proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
       If LargoPin = 0 Then vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.text = "": MsgBox "Largo del atributo codigo barra, esta en valor cero o nulo. Proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
       Set RS = vg_db.Execute(RutinaLectura.Personal(6, Mid(CodBar, PosInicialPin, LargoPin), ""))
       If Not RS.EOF Then
          CodigoCliente = RS!cli_codigo
       Else
          RS.Close: Set RS = Nothing
          vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.text = ""
          MsgBox "Código barra no tiene asignado su cliente. Proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: Exit Sub
       End If
       RS.Close: Set RS = Nothing
    End If
    vg_db.Execute ("sgp_Ins_DetalleLectura '" & MuestraCasino(1) & "', '" & CodigoCliente & "', " & codreg & ", " & codser & ", " & codate & ", '" & CodBar & "'")
    MoverDatosGrillas
Exit Sub
Man_Error:
fg_descarga
If Err = 13 Then MsgBox "No corresponde formato fecha, indicado codigo barra", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub
