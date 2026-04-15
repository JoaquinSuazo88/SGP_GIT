VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form P_GenInvAx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación Inventario OPTIMUM"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9105
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
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   6
      Top             =   6480
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
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
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
      Height          =   375
      Index           =   0
      Left            =   6240
      TabIndex        =   4
      Top             =   6480
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
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8775
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   1560
         TabIndex        =   2
         Top             =   5520
         Width           =   1785
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   1560
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5055
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   8295
         _Version        =   393216
         _ExtentX        =   14631
         _ExtentY        =   8916
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
         SpreadDesigner  =   "P_GenInvAx.frx":0000
      End
   End
End
Attribute VB_Name = "P_GenInvAx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msgtitulo As String

Private Sub Command1_Click(Index As Integer)

On Error GoTo error

Dim RS          As New ADODB.Recordset
Dim RSAx        As New ADODB.Recordset
Dim isel        As Boolean
Dim i           As Long
Dim seleccion   As String
Dim xmlperiodo  As String
Dim periodo     As String
Dim periodommaa As String
Dim sql         As String
Dim AnJes       As Object
Dim hora        As String
Dim Ceco        As String
Dim cuenta      As String
Dim CodigoInv   As Long
Dim Fecha       As Long

Select Case Index

Case 0
    
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
       
       MsgBox "Debe seleccionar a lo menos un item de la lista o bien no hay datos grilla...", vbCritical + vbOKOnly, Msgtitulo
    
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

           vaSpread1.Col = 2 'codigo
           CodigoInv = IIf(vaSpread1.text = "", 0, vaSpread1.text)

           vaSpread1.Col = 3 'Periodo
           periodo = IIf(vaSpread1.text = "", 0, Mid(vaSpread1.text, 4, 4) & Mid(vaSpread1.text, 1, 2))
           
           vaSpread1.Col = 5 'Fecha
           Fecha = IIf(vaSpread1.text = "", 0, vaSpread1.text)
           
           If GeneraInvAX(CodigoInv, periodo, Fecha) Then
           
           End If
        End If
    Next i
    
    CargarInventarioGrilla

Case 1
    
    Me.Hide
    Unload Me

Case 2 '-------> Explorar carpeta envio inventario OPTIMUM
    
    ExplorarCarpeta dir_trabajo_Inf & "InformesAXInventario"

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

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo error

Me.HelpContextID = vg_OpcM
Msgtitulo = "Generar Inventario OPTIMUM"
fg_centra Me
modo = ""
vaSpread1.MaxRows = 0
Frame1.Caption = MuestraCasino(1)
Frame1.Caption = Trim(Frame1.Caption) & " - " & Trim(MuestraCasino(2))
CargarInventarioGrilla

Exit Sub
error:
    MsgBox Err.Description, vbCritical

End Sub

Sub CargarInventarioGrilla()

On Error GoTo error

Dim RS As New ADODB.Recordset
Dim Ceco As String

'--> Validar homologación Ceco y cuentas contable AX
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Ceco = MuestraCasino(1)
Set RS = vg_db.Execute("SELECT 1 FROM Cecos_Sap_AX csa with (nolock) inner join b_clientes bc on bc.cli_codigo = csa.Cecos_Sap and bc.cli_socsap = csa.Sociedad_Sap and bc.cli_activo = '1' and bc.cli_tipo = 0 and bc.cli_codbod > 0 WHERE bc.cli_codigo = '" & Ceco & "'")
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe homologación Ceco con AX..., Proceso Cancelado", vbCritical + vbOKOnly, Msgtitulo
   Exit Sub

End If
RS.Close
Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_ValidarCuentasAx '" & Ceco & "'")
If Not RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe homologación cuentas contables AX..., Proceso Cancelado", vbCritical + vbOKOnly, Msgtitulo
   Exit Sub

End If
RS.Close
Set RS = Nothing

'--> Cargar datos grilla
vaSpread1.MaxRows = 0
Set RS = vg_db.Execute("sgp_Sel_EstadoEnvioAX '" & Ceco & "'")
Do While Not RS.EOF
   
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   vaSpread1.text = ""
   
   vaSpread1.Col = 2
   vaSpread1.text = RS!codigoenv
   
   vaSpread1.Col = 3
   vaSpread1.text = RS!tin_ciemes
   
   vaSpread1.Col = 4
   vaSpread1.text = RS!glosa
   
   vaSpread1.Col = 5
   vaSpread1.text = RS!tin_fectom
   
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

If Col = 1 And Row = 0 Then
   
   vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.Value = IIf(vaSpread1.Value = "1", "", "1")

End If

End Sub
