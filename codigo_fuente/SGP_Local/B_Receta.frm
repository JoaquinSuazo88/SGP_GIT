VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form B_Receta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Plato Menú"
   ClientHeight    =   6345
   ClientLeft      =   2280
   ClientTop       =   2475
   ClientWidth     =   12180
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6345
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   1560
      TabIndex        =   3
      Top             =   0
      Width           =   8325
      Begin VB.TextBox FptNombre 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2610
         LinkTimeout     =   0
         MaxLength       =   80
         TabIndex        =   0
         Top             =   915
         Width           =   3195
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2610
         TabIndex        =   10
         Top             =   570
         Width           =   4005
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2145
         Picture         =   "B_Receta.frx":0000
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2145
         Picture         =   "B_Receta.frx":030A
         Top             =   135
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Plato"
         Height          =   195
         Index           =   5
         Left            =   1155
         TabIndex        =   7
         Top             =   645
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C. Dietetica"
         Height          =   195
         Index           =   3
         Left            =   1170
         TabIndex        =   6
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Texto"
         Height          =   195
         Index           =   1
         Left            =   1155
         TabIndex        =   5
         Top             =   1035
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Registro 0"
         Height          =   195
         Index           =   0
         Left            =   5820
         TabIndex        =   4
         Top             =   1035
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2610
         TabIndex        =   8
         Top             =   240
         Width           =   4005
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   2655
         TabIndex        =   9
         Top             =   285
         Width           =   4005
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2655
         TabIndex        =   11
         Top             =   615
         Width           =   4005
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4740
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   11325
      _Version        =   393216
      _ExtentX        =   19976
      _ExtentY        =   8361
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   20
      OperationMode   =   2
      RestrictRows    =   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "B_Receta.frx":0614
      VisibleCols     =   3
      VisibleRows     =   20
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6345
      Left            =   11550
      TabIndex        =   2
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   11192
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "B_Receta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim i As Long, iRow As Long, filcatdie As Long, filtippla As Long
Dim findstring As String, sourcestring As String
Dim swactiva As Integer, iayuda As Integer

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga
If Trim(ws_respuesta) <> "" Then fpTnombre.text = ws_respuesta: fpTnombre.SelStart = Len(ws_respuesta): ws_respuesta = ""

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
fg_centra Me
fg_carga ""
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirma"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Deshacer", , tbrDefault, "A_Deshacer"): BtnX.Visible = True: BtnX.ToolTipText = "Deshacer"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_VerReceta", , tbrDefault, "A_VerReceta"): BtnX.Visible = True: BtnX.ToolTipText = "Ver Recetas"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
vaSpread1.MaxRows = 0
fpayuda(0).Caption = "Todos"
fpayuda(1).Caption = "Todos"
iayuda = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'catdefecto'")
If Not RS.EOF Then filcatdie = RS!par_valor: fpayuda(0).Caption = fg_BuscaenArbol(RS!par_valor, "a_recetacatdie", "car_codigo")
RS.Close
Set RS = Nothing

filtippla = 0
MoverRecetasGrilla

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpTnombre_Change()

On Error GoTo Man_Error

With vaSpread1
    
    If .MaxRows < 1 Then Exit Sub
    findstring = fpTnombre.text 'Trim(FptNombre.Text)
    If fpTnombre.text = "" Then
       
       .Visible = False
       swactiva = 0
       
       For i = 1 To .MaxRows
           
           .Row = i
           .RowHidden = False
           If swactiva = 0 Then .OperationMode = 2: .Action = 0: swactiva = 1
       
       Next i
       Label1(0).Caption = "Registro " & Format(.MaxRows, fg_Pict(6, 0))
       .Visible = True
    
    Else
       
       swactiva = 0
       .Visible = False
       iRow = 0
       
       For i = 1 To .MaxRows
           
           .Row = i
           .Col = 2
           sourcestring = .Value 'Trim(.Value)
           IndActivo = UCase(Trim(sourcestring)) Like "*" & UCase(findstring) & "*"
           If IndActivo = -1 Then
              
              If swactiva = 0 Then .OperationMode = 2: .Action = 0: swactiva = 1
              If .RowHidden = True Then .RowHidden = False
              iRow = iRow + 1
           
           Else
              
              If .RowHidden = False Then .RowHidden = True
           
           End If
       
       Next i
       Label1(0).Caption = "Reg. Enc. " & Format(iRow, fg_Pict(6, 0))
       .Visible = True
    
    End If

End With

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FptNombre_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If KeyCode = 27 Then ICGrilla = 0: Me.Hide: Exit Sub
If KeyCode = 40 Or KeyCode = 34 And iRow > 0 Then vaSpread1.SetFocus

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

    Case 0
    
        vg_codigo = "": vg_nombre = ""
        vg_left = fpayuda(0).Left + 2400
        B_ArbEst.MoverDatosTvwDir "a_recetacatdie", "car_", "Categoria Dietetica"
        B_ArbEst.Show 1
        If vg_codigo = "" Then Exit Sub
        filcatdie = Val(vg_codigo)
        fpayuda(0).Caption = vg_nombre
        vg_nombre = ""
        fpTnombre.text = ""
        MoverRecetasGrilla
    
    Case 1
        
        vg_codigo = "": vg_nombre = ""
        vg_left = fpayuda(1).Left + 2400
        B_ArbEst.MoverDatosTvwDir "a_recetatippla", "tip_", "Tipo Plato"
        B_ArbEst.Show 1
        If Trim(vg_codigo) = "" Then Exit Sub
        tippla = Val(vg_codigo)
        filtippla = Val(vg_codigo)
        fpayuda(1).Caption = vg_nombre
        vg_nombre = ""
        fpTnombre.text = ""
        MoverRecetasGrilla

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index
    
    Case 1
        
        If vaSpread1.MaxRows < 1 Then Exit Sub
        ICGrilla = 1
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1_DblClick vaSpread1.Col, vaSpread1.Row
    
    Case 3
        
        fpTnombre.text = ""
        
        filcatdie = 0
        filtippla = 0
        
        fpayuda(0).Caption = "Todos"
        fpayuda(1).Caption = "Todos"
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        RS.Open "SELECT par_valor FROM a_param WHERE par_cencos='" & MuestraCasino(1) & "' AND par_codigo='catdefecto'", vg_db, adOpenStatic
        If Not RS.EOF Then
        
           filcatdie = RS!par_valor
           fpayuda(0).Caption = fg_BuscaenArbol(RS!par_valor, "a_recetacatdie", "car_codigo")
           
        End If
        
        RS.Close
        Set RS = Nothing
        
        MoverRecetasGrilla
        
      Case 5
        
        vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1
        vg_newcodrec = Val(vaSpread1.text)
        vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 6
        vg_tiprec = Val(vaSpread1.text)
        vg_auxtiprec = vg_tiprec
        vg_5etapas = IIf(vg_codregimen < 10000, False, True)
    
    '    M_Receta.Show 0
        Dim Receta As New M_Receta
        vg_RecetaReal = 1
        Receta.Show 1, Me
        Set Receta = Nothing
    
    Case 7
        
        ICGrilla = 0
        Me.Hide
        
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

fpTnombre.SetFocus

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If Row > 0 Then vaSpread1.Row = vaSpread1.ActiveRow

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If Row < 1 Then Exit Sub
With vaSpread1
    .Row = Row
    .Col = 1: vg_codigo = .text
    .Col = 2: vg_nombre = .text
    .Col = 6: vg_tiprec = .text
End With
Me.Hide

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
ICGrilla = 1
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1_DblClick vaSpread1.Col, vaSpread1.Row

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub MoverRecetasGrilla()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim arr
Dim CodRec As Long
Dim i As Long
Dim X As Long

fg_carga ""
CodRec = 0
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
'Modificación Jpaz 20130208 RS.Open RutinaLectura.Receta(3, 0, filcatdie, filtippla, "", 0), vg_db, adOpenStatic

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_s_traerlistareceta '" & MuestraCasino(1) & "', " & filcatdie & ", " & filtippla & ", " & vg_codregimen & "")

If Not RS.EOF Then
   
   i = 1
   X = 1
   vaSpread1.MaxRows = RS!nreg
   arr = RS.GetRows
   RS.Close: Set RS = Nothing
   For i = 0 To UBound(arr, 2)
      
      If arr(0, i) <> CodRec Then
'         vaSpread1.MaxRows = vaSpread1.MaxRows + 1
'                  vaSpread1.Row = vaSpread1.MaxRows
         vaSpread1.Row = X

         grdCellTypeStatic vaSpread1, 1, vaSpread1.Row, 1
         grdSetText vaSpread1, 1, vaSpread1.Row, arr(0, i)

         grdCellTypeStatic vaSpread1, 2, vaSpread1.Row, 0
         grdSetText vaSpread1, 2, vaSpread1.Row, arr(1, i)
         
         grdCellTypeStatic vaSpread1, 3, vaSpread1.Row, 0
         grdSetText vaSpread1, 3, vaSpread1.Row, arr(3, i)
         
         CodRec = arr(0, i)
         X = X + 1
         
      End If
      
      vaSpread1.Col = 4
      If arr(2, i) = -1 Or (arr(2, i) > 0 And vg_codregimen = arr(2, i)) Then
         
         grdCellTypeStatic vaSpread1, 4, vaSpread1.Row, 1
         grdSetText vaSpread1, 4, vaSpread1.Row, arr(4, i)
         
         grdCellTypeStatic vaSpread1, 5, vaSpread1.Row, 2
         grdSetText vaSpread1, 5, vaSpread1.Row, IIf(arr(2, i) > 0, "x Regimen", "Local")
         
         grdCellTypeStatic vaSpread1, 6, vaSpread1.Row, 2
         grdSetText vaSpread1, 6, vaSpread1.Row, arr(2, i)
      
      ElseIf arr(2, i) = 0 And Trim(vaSpread1.text) = "" Then
         
         grdCellTypeStatic vaSpread1, 4, vaSpread1.Row, 1
         grdSetText vaSpread1, 4, vaSpread1.Row, Format(arr(4, i), fg_Pict(6, 2))

         grdCellTypeStatic vaSpread1, 5, vaSpread1.Row, 2
         grdSetText vaSpread1, 5, vaSpread1.Row, "Patrón"
         
         grdCellTypeStatic vaSpread1, 6, vaSpread1.Row, 2
         grdSetText vaSpread1, 6, vaSpread1.Row, arr(2, i)
         
      End If
      
   Next i
   vaSpread1.MaxRows = X - 1
   Label1(0).Visible = True
   Label1(0).Caption = "Registro " & Format(vaSpread1.MaxRows, fg_Pict(6, 0))
   fg_descarga
   
End If
   
   If RS.State = 1 Then
      
      fg_descarga
      vaSpread1.Visible = True
      Label1(0).Visible = True
      Label1(0).Caption = "Registro " & Format(vaSpread1.MaxRows, fg_Pict(6, 0))
      RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No Existen Recetas", vbExclamation + vbOKOnly, "Busqueda Recetas"
   
   End If
   
If vaSpread1.MaxRows > 0 Then vaSpread1.Row = 1
vaSpread1.Visible = True

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If KeyCode = 27 Then ICGrilla = 0: Me.Hide: Exit Sub
If TeclasNoPermitidas(KeyCode) = True Then fpTnombre.text = IIf(KeyCode = 8, fpTnombre.text, fpTnombre.text & Chr(KeyCode)): fpTnombre.SetFocus: fpTnombre.SelStart = Len(fpTnombre.text)

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub
