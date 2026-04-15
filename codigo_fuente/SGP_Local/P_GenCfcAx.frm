VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form P_GenCfcAx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación Cfc OPTIMUM"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ver Carpeta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   3720
      TabIndex        =   10
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar AX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   5280
      TabIndex        =   3
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   6720
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.Frame Frame3 
         Height          =   555
         Index           =   2
         Left            =   4080
         TabIndex        =   11
         Top             =   5520
         Width           =   3225
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   12
            Top             =   170
            Width           =   3000
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   300
         Width           =   4095
      End
      Begin VB.Frame Frame3 
         Height          =   555
         Index           =   1
         Left            =   2880
         TabIndex        =   6
         Top             =   5520
         Width           =   1185
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   7
            Top             =   170
            Width           =   960
         End
      End
      Begin VB.Frame Frame3 
         Height          =   555
         Index           =   0
         Left            =   1680
         TabIndex        =   4
         Top             =   5520
         Width           =   1185
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   170
            Width           =   960
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4575
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   7335
         _Version        =   393216
         _ExtentX        =   12938
         _ExtentY        =   8070
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         SpreadDesigner  =   "P_GenCfcAx.frx":0000
      End
      Begin VB.Label Label1 
         Caption         =   "Lugar Fisico"
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
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "P_GenCfcAx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgTitulo As String
Dim est As Boolean
Dim OpFormatoAx As Boolean
Dim OpFormato As Integer

Private Sub Combo1_Click(Index As Integer)

On Error GoTo error

If est Then Exit Sub

Dim RS As New ADODB.Recordset
Dim LugFis As String
LugFis = fg_codigocbo(Combo1, 1, 4, 1)

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', 'ParLugFis'")
If Not RS.EOF Then
   
   est = True
   vg_db.Execute ("sgp_Upd_Param 1, '" & MuestraCasino(1) & "', 'ParLugFis', '', '', '" & LugFis & "'")
   est = False

Else
   
   vg_db.Execute ("sgp_Ins_Param 'ParLugFis','Parametro Lugar Fisico','C', '" & LugFis & "', '" & MuestraCasino(1) & "'")

End If
RS.Close
Set RS = Nothing

Exit Sub
error:
fg_descarga
MsgBox Err.Description, vbCritical

End Sub

Private Sub Command1_Click(Index As Integer)

On Error GoTo error

Dim RS          As New ADODB.Recordset
Dim RSAx        As New ADODB.Recordset
Dim isel        As Boolean
Dim i           As Long
Dim j           As Long
Dim seleccion   As String
Dim xmlperiodo  As String
Dim periodo     As String
Dim periodommaa As String
Dim Sql         As String
Dim AnJes       As Object
Dim hora        As String
Dim Ceco        As String
Dim Folio       As Long
Dim numdoc      As Long
Dim rutpro      As String
Dim LugarFisico As String
Dim Inf_Tipo    As String

Select Case Index

Case 0
    
    '--> Lugar Fisico
    If Combo1(1).ListIndex = -1 And OpFormatoAx = True Then
       
       MsgBox "Debe selecionar lugar fisico", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    LugarFisico = fg_codigocbo(Combo1, 1, 4, 1)
    
    '--> Validar que haya seleccionado un item de la lista
    isel = False
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        
        If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
           
           isel = True
           Exit For
        
        End If
    
    Next i
    
    If Not isel Then
       
       MsgBox "Debe seleccionar a lo menos un item de la lista o bien no hay datos grilla...", vbCritical + vbOKOnly, MsgTitulo
    
    End If
    
    '--> Generar archivos AX
    seleccion = 0
    Ceco = MuestraCasino(1)
    
    For i = 1 To vaSpread1.MaxRows
       
        vaSpread1.Row = i
        vaSpread1.Col = 1 'Seleccion
        seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
        If seleccion = 1 And vaSpread1.RowHidden = False Then
           
           fg_carga ""
           vaSpread1.Row = i
           
           vaSpread1.Col = 2 'Folio
           Folio = IIf(vaSpread1.text = "", 0, vaSpread1.text)
           
           vaSpread1.Col = 3 'Periodo
           periodo = IIf(vaSpread1.text = "", 0, Mid(vaSpread1.text, 4, 4) & Mid(vaSpread1.text, 1, 2))

           vaSpread1.Col = 5 'Inf_tipo
           Inf_Tipo = IIf(vaSpread1.text = "", 0, vaSpread1.text)
           
           If OpFormatoAx Then
              
              If Not GeneraCfcAX(Folio, periodo, LugarFisico) Then
              
                 Call MsgBox("No genero correctamente archivos OPTIMUM, trate de generar por envio Facturación OPTIMUM", vbInformation)
           
              End If
           
           ElseIf Not OpFormatoAx And OpFormato = 1 Then
        
              If Not GeneraCfcDigitado(Folio, periodo, Inf_Tipo) Then
              
                 Call MsgBox("No genero correctamente archivos CFC MANUAL, trate de generar por envio Facturación CFC MANUAL", vbInformation)
           
              End If
        
           ElseIf Not OpFormatoAx And OpFormato = 3 Then
           
             If Not GenerarTraspasoSalidaAX(Folio, periodo) Then
             
                 Call MsgBox("No genero correctamente archivos Traspaso salida manual, trate de generar por envio Facturación Traspaso salida manual", vbInformation)
             
             End If
           
           End If
           
        End If
    
    Next i
    
    Text1(2).text = ""
    Text1(3).text = ""
    Text1(4).text = ""
    
    CargarInventarioGrilla

Case 1
    
    Me.Hide
    Unload Me

Case 2 'explorar carpeta
    
    If OpFormatoAx And OpFormato = 2 Then
       
       ExplorarCarpeta dir_trabajo_Inf & "InformesAXFacturacion"
    
    ElseIf Not OpFormatoAx And (OpFormato = 1 Or OpFormato = 3) Then
    
       ExplorarCarpeta dir_trabajo_Inf & "InformesAXFacturacionManual"
    
    End If

End Select

Exit Sub
error:
    fg_descarga
    If Err.Number = 70 Then
       
       Call MsgBox("Archivo se envuentra en uso, cerrar archivo csv para continuar", vbInformation, Me.Caption)
       Screen.MousePointer = DEFAULT
    
    Else
       
       MsgBox Err.Description, vbCritical
    
    End If

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo error

Dim RS As New ADODB.Recordset

OpFormatoAx = True
'-------> Validar si el contrato tiene opción de envio sap
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT * FROM b_casinointerfaz WHERE cai_cencos = '" & MuestraCasino(1) & "'and cai_codtii = 6")

If Not RS.EOF Then
   
   Combo1(1).Visible = False
   Label1.Visible = False
   OpFormatoAx = False
    
End If
RS.Close: Set RS = Nothing

Me.HelpContextID = vg_OpcM
fg_centra Me
modo = ""
vaSpread1.MaxRows = 0
Frame1.Caption = MuestraCasino(1)
Frame1.Caption = Trim(Frame1.Caption) & " - " & Trim(MuestraCasino(2))

est = True
'-------> Cargar Lugar Fisico
CargarDatoCombo Combo1, 1, "LugarFisico_AX", "cli_", "LugFis", "A"

'-------> buscar a_param lugar fisico
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', 'ParLugFis'")
If Not RS.EOF Then
   
   Combo1(1).ListIndex = fg_buscacbostring(Combo1, 1, 4, (RS!par_valor))

End If
RS.Close
Set RS = Nothing

CargarInventarioGrilla

est = False
Exit Sub
error:
    MsgBox Err.Description, vbCritical

End Sub

Sub Inicio(op As Integer)

On Error GoTo error

OpFormato = op
MsgTitulo = IIf(op = 2, "Generar Facturación AX", IIf(op = 1, "Generar Facturación Manual", "Generar Traspaso Salida"))
Me.Caption = IIf(op = 2, "Generar Facturación AX", IIf(op = 1, "Generar Facturación Manual", "Generar Traspaso Salida"))
Command1(0).Caption = IIf(op = 2, "Generar AX", IIf(op = 1, "Genera Manual", "Genera Traspaso Salida"))


Exit Sub
error:
    MsgBox Err.Description, vbCritical

End Sub

Sub CargarInventarioGrilla()

On Error GoTo error

Dim RS   As New ADODB.Recordset
Dim Ceco As String
Dim Sql  As String

'--> Validar homologación Ceco y cuentas contable AX
Ceco = MuestraCasino(1)

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT 1 FROM Cecos_Sap_AX csa with (nolock) inner join b_clientes bc on bc.cli_codigo = csa.Cecos_Sap and bc.cli_socsap = csa.Sociedad_Sap and bc.cli_activo = '1' and bc.cli_tipo = 0 and bc.cli_codbod > 0 WHERE bc.cli_codigo = '" & Ceco & "' ")
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe homologación Ceco con OPTIMUM..., Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Sub

End If
RS.Close
Set RS = Nothing

Dim Mglosa As String
Mglosa = ""
     
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_ValidarCuentasFacAx '" & Ceco & "'")
If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      Mglosa = Mglosa & RS!pro_ctacon & VgLinea
      RS.MoveNext
   
   Loop

End If
RS.Close
Set RS = Nothing

If Trim(Mglosa) <> "" Then
   
   fg_descarga
   MsgBox "Existe cuentas SAP que no estan hologados OPTIMUM, estas son las siguintes : " & VgLinea & Mglosa & "Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Sub

End If

'--> Cargar datos grilla
vaSpread1.MaxRows = 0

Sql = ""
If OpFormato = 1 Then

    Sql = Sql & "sgp_Sel_EstadoEnvioFacturacionManual "
    
ElseIf OpFormato = 2 Then

    Sql = Sql & "sgp_Sel_EstadoEnvioFacturacionAX "

ElseIf OpFormato = 3 Then
    
    Sql = Sql & "sgp_Sel_EstadoEnvioTraspasoSalidaManual"
    
End If

Sql = Sql & "'" & Ceco & "'"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute(" " & Sql & "")
Do While Not RS.EOF
   
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   
   vaSpread1.Col = 1
   vaSpread1.text = ""
   
   vaSpread1.Col = 2
   vaSpread1.text = RS!inf_numero
   
   vaSpread1.Col = 3
   vaSpread1.text = RS!toc_fecper
   
   vaSpread1.Col = 4
   vaSpread1.text = RS!Glosa
   
   If OpFormato = 1 Then
      
      vaSpread1.Col = 5
      vaSpread1.text = RS!Inf_Tipo
   
   End If
   
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing

Exit Sub
error:
    MsgBox Err.Description, vbCritical

End Sub

Private Sub Text1_Change(Index As Integer)

Dim i As Long
Dim IndActivo As Integer

On Error GoTo error

Select Case Index

    Case 2
    
        Text1(3).text = ""
        Text1(4).text = ""
    
    Case 3
    
        Text1(2).text = ""
        Text1(4).text = ""
    
    Case 4
    
        Text1(2).text = ""
        Text1(3).text = ""

End Select

    vaSpread1.Visible = False
    If Trim(Text1(Index).text) <> "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           IndActivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread1.Col = Index
           
           If IndActivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           Else
              
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           
           End If
        
        Next i
        
        vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1_Click Index, 0
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    
    vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
    
    If Trim(Text1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

Exit Sub
error:
    MsgBox Err.Description, vbCritical

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo error

If Col = 1 And Row = 0 Then
   
   vaSpread1.Row = -1
   vaSpread1.Col = 1
   vaSpread1.Value = IIf(vaSpread1.Value = "1", "", "1")

End If

Exit Sub
error:
    MsgBox Err.Description, vbCritical

End Sub

