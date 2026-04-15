VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form B_RecMinutaBloqueAMD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Receta Minuta Bloque"
   ClientHeight    =   7005
   ClientLeft      =   4215
   ClientTop       =   2310
   ClientWidth     =   15885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   15885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   3315
      TabIndex        =   0
      Top             =   0
      Width           =   10125
      Begin VB.CheckBox IncluirColCalorias 
         Caption         =   "Incluir Columna Calorias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8370
         TabIndex        =   13
         Top             =   930
         Width           =   1695
      End
      Begin VB.CheckBox IncluirColCosto 
         Caption         =   "Incluir Columna Costo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8400
         TabIndex        =   12
         Top             =   270
         Width           =   1695
      End
      Begin VB.TextBox FptNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   1830
         LinkTimeout     =   0
         MaxLength       =   80
         TabIndex        =   1
         Top             =   915
         Width           =   3375
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   1830
         TabIndex        =   7
         Top             =   570
         Width           =   6075
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   1305
         Picture         =   "B_RecMinutaBloqueAMD.frx":0000
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   1305
         Picture         =   "B_RecMinutaBloqueAMD.frx":030A
         Top             =   135
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   6
         Top             =   645
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C. Dietetica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   5
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Texto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   4
         Top             =   1035
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Registro 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   5310
         TabIndex        =   3
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
         Left            =   1830
         TabIndex        =   2
         Top             =   240
         Width           =   6075
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   1875
         TabIndex        =   8
         Top             =   285
         Width           =   6075
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   1875
         TabIndex        =   9
         Top             =   615
         Width           =   6075
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   5220
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   15075
      _Version        =   393216
      _ExtentX        =   26591
      _ExtentY        =   9208
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
      MaxCols         =   10
      MaxRows         =   20
      OperationMode   =   2
      RestrictRows    =   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "B_RecMinutaBloqueAMD.frx":0614
      VisibleCols     =   3
      VisibleRows     =   20
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7005
      Left            =   15255
      TabIndex        =   11
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   12356
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "B_RecMinutaBloqueAMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FilCatDie As Long, FilTipPla As Long
Dim SwActiva As Integer, iayuda As Integer
Dim FormActivado As Boolean ' esta variable es para saber si ya esta cargado el form ya que daba un error=
Dim Est As Boolean

Private Sub Form_Activate()

If Not Est Then
   
   LoadForm

End If
Est = False
fg_descarga
If Trim(ws_respuesta) <> "" Then FptNombre.text = ws_respuesta: FptNombre.SelStart = Len(ws_respuesta): ws_respuesta = ""

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Dim RS           As New ADODB.Recordset
Dim AuxFilCatDie As Long
Dim AuxFilTipPla As Long
'-------> Leer datos parametros para el calculo costo
Est = True
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_ParamTipoReceta 'Mcr-" & Trim(vg_NUsr) & "'")

If Not RS.EOF Then IncluirColCosto.Value = RS!par_valor
RS.Close: Set RS = Nothing

'-------> Leer datos parametros para el calculo costo
Est = True

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_ParamTipoReceta 'Mcc-" & Trim(vg_NUsr) & "'")

If Not RS.EOF Then IncluirColCalorias.Value = RS!par_valor
RS.Close: Set RS = Nothing

'-------> leer datos parametros categoria ditetica
FilCatDie = 0

If vg_filcatdieMin = 0 Then

    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_ParamTipoReceta 'CatDiePlan-" & Trim(vg_NUsr) & "'")
    
    If Not RS.EOF Then FilCatDie = RS!par_valor: AuxFilCatDie = RS!par_valor
    RS.Close: Set RS = Nothing

Else

   FilCatDie = vg_filcatdieMin
   AuxFilCatDie = vg_filcatdieMin
   
End If

If FilCatDie = 0 Then
   
   fpayuda(0).Caption = "Todos"

Else
   
   fpayuda(0).Caption = fg_BuscaenArbol(AuxFilCatDie, "a_recetacatdie", "car_codigo")

End If

'-------> leer datos parametros tipo plato
FilTipPla = 0

If vg_filtipplaMin = 0 Then

    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS = vg_db.Execute("sgpadm_Sel_ParamTipoReceta 'TipPlaPlan-" & Trim(vg_NUsr) & "'")
    If Not RS.EOF Then FilTipPla = RS!par_valor: AuxFilTipPla = RS!par_valor
    RS.Close: Set RS = Nothing

Else

    FilTipPla = vg_filtipplaMin
    AuxFilTipPla = vg_filtipplaMin

End If

If FilTipPla = 0 Then
   
   fpayuda(1).Caption = "Todos"

Else
   
   fpayuda(1).Caption = fg_BuscaenArbol(AuxFilTipPla, "a_recetatippla", "tip_codigo")

End If

Est = False
FormActivado = True

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Public Sub fpTnombre_Change()

On Error GoTo Man_Error

Dim FindString     As String
Dim SourceString   As String
Dim FinsStringTipo As String
Dim inactivo1      As Long
Dim inactivo2      As Long
Dim i              As Long
Dim IRow           As Long

If vaSpread1.MaxRows < 1 Then Exit Sub

FindString = Trim(FptNombre.text)
If FptNombre.text = "" Then
   
   vaSpread1.Visible = False
   SwActiva = 0
   
   For i = 1 To vaSpread1.MaxRows
       
       vaSpread1.Row = i
       vaSpread1.Col = 5
       SourceString = Trim(vaSpread1.Value)
       indactivo2 = UCase(Trim(SourceString)) Like "*" & UCase(FinsStringTipo) & "*"
       
       If indactivo2 = -1 Then
          
          If SwActiva = 0 Then vaSpread1.OperationMode = 2: vaSpread1.Action = 0: SwActiva = 1
          If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
          IRow = IRow + 1
       
       Else
          
          If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
       
       End If
   
   Next i
   
   Label1(0).Caption = "Registro " & Format(vaSpread1.MaxRows, fg_Pict(6, 0))
   vaSpread1.Visible = True
Else
   SwActiva = 0
   vaSpread1.Visible = False
   IRow = 0
   
   For i = 1 To vaSpread1.MaxRows
       
       vaSpread1.Row = i
       vaSpread1.Col = 2
       SourceString = Trim(vaSpread1.Value)
       indactivo1 = UCase(Trim(SourceString)) Like "*" & UCase(FindString) & "*"
       vaSpread1.Col = 5
       SourceString = Trim(vaSpread1.Value)
       indactivo2 = UCase(Trim(SourceString)) Like "*" & UCase(FinsStringTipo) & "*"
       
       If indactivo1 = -1 And indactivo2 = -1 Then
          
          If SwActiva = 0 Then vaSpread1.OperationMode = 2: vaSpread1.Action = 0: SwActiva = 1
          If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
          IRow = IRow + 1
       
       Else
          
          If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
       
       End If
   
   Next i
   
   Label1(0).Caption = "Reg. Enc. " & Format(IRow, fg_Pict(6, 0))
   vaSpread1.Visible = True

End If

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fptnombre_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If KeyCode = 27 Then ICGrilla = 0: Me.Hide: Exit Sub
If KeyCode = 40 Or KeyCode = 34 And IRow > 0 Then vaSpread1.SetFocus

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0
    
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(0).Left + 2400
    B_ArbEst.MoverDatosTvwDir "a_recetacatdie", "car_", "Categoria Dietetica", "1"
    B_ArbEst.Show 1
    If vg_codigo = "" Then Exit Sub
'    DoEvents
'    Screen.MousePointer = vbHourglass
    FilCatDie = Val(vg_codigo)
    vg_filcatdieMin = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
    vg_nombre = ""
    FptNombre.text = ""
    MoverRecetasGrilla

Case 1
    
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(1).Left + 2400
    B_ArbEst.MoverDatosTvwDir "a_recetatippla", "tip_", "Tipo Plato", "1"
    B_ArbEst.Show 1
    If Trim(vg_codigo) = "" Then Exit Sub
    FilTipPla = Val(vg_codigo)
    vg_filtipplaMin = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    vg_nombre = ""
    FptNombre.text = ""
    MoverRecetasGrilla

End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Option1_Click(Index As Integer)

fpTnombre_Change
FptNombre.SetFocus

End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)

FptNombre.SetFocus

End Sub

Private Sub Option1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

FptNombre.SetFocus

End Sub

Private Sub IncluirColCalorias_Click()

If Est Then Exit Sub
FptNombre.text = ""
MoverRecetasGrilla

End Sub

Private Sub IncluirColCosto_Click()

If Est Then Exit Sub
FptNombre.text = ""
MoverRecetasGrilla

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Select Case Button.Index

Case 1
    
    If vaSpread1.MaxRows < 1 Then Exit Sub
    ICGrilla = 1
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1_DblClick vaSpread1.Col, vaSpread1.Row

Case 3
    
    FptNombre.text = "": FilCatDie = 0: FilTipPla = 0
    fpayuda(0).Caption = "Todos": fpayuda(1).Caption = "Todos"
    MoverRecetasGrilla

Case 5
  
  vaSpread1.Col = 1: vaSpread1.Row = vaSpread1.ActiveRow
  
  If vaSpread1.text <> "" And vaSpread1.text <> "Código" Then
    vg_newestrec = True
    vg_modreceta = True 'False
    
    If vg_newestrec = True Then
       
       vg_fecval = 0: vg_fecval = Val(vg_fecha) & Right("0" & (Int(xcol / 7) + 2), 2)
       
       If Len(vg_Zona) = 0 Then
            
            Let vg_Zona = 0
        
       End If
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
'       Set RS = vg_db.Execute("sgpadm_s_planifminuta 3, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & "," & vg_Zona & ", " & vg_fecval & ", 0, 0," & vg_IndpprSelec & "")
       Set RS = vg_db.Execute("sgpadm_s_planifminuta 3, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & "," & vg_Zona & ", " & vg_fecval & ", 0, 0," & vg_Indppr & "")
       If Not RS.EOF Then vg_fecval = RS!mid_fecval: vg_opcion = 2
       RS.Close: Set RS = Nothing
    
    End If
    
    StrRec = vaSpread1.text
    vg_newcodrec = StrRec
    
    auxtiprec = vg_tiprec
    Dim Receta As New M_Receta
    Dim auxrecetareal As Integer
    auxrecetareal = vg_RecetaReal
    vg_auxtiprec = vg_tiprec
    Let VarSitioRemoto = True
    Receta.Show 1, Me
    Set Receta = Nothing
    vg_RecetaReal = auxrecetareal
    Est = True
  
  End If

Case 7
    
    ICGrilla = 0
    Me.Hide

End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

FptNombre.SetFocus

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If Row > 0 Then vaSpread1.Row = vaSpread1.ActiveRow

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim NomUnidadReceta As String

If Row < 1 Then Exit Sub
vaSpread1.Row = Row
'vaSpread1.Col = 8: NomUnidadReceta = vaSpread1.text
vaSpread1.Col = 9: NomUnidadReceta = vaSpread1.text

vaSpread1.Col = 1: vg_codigo = vaSpread1.text
vaSpread1.Col = 2: vg_nombre = vaSpread1.text & " [" & NomUnidadReceta & "]"
'vaSpread1.Col = 4: vg_Valor = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
vaSpread1.Col = 5: vg_Valor = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
'vaSpread1.Col = 6: vg_Calorias = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
vaSpread1.Col = 7: vg_Calorias = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
Me.Hide

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

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
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub MoverRecetasGrilla()

On Error GoTo Man_Error

fg_carga ""
Dim RS     As New ADODB.Recordset
Dim X      As Boolean
Dim Fecha  As Long
Dim dato   As Variant
Dim IndRec As Long
'------- Grabar Parametro mostrar costo receta min bloque
vg_db.Execute ("sgpadm_iu_param 'UI', 'Mcr-" & Trim(vg_NUsr) & "', 'Parametro Mostrar Costo Receta min.bloque', 'N', '" & IIf(IncluirColCosto.Value = 1, 1, 0) & "'")

'-------> Leer datos
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_Sel_ParamTipoReceta 'Mcr-" & Trim(vg_NUsr) & "'")
If Not RS.EOF Then IncluirColCosto.Value = RS!par_valor
RS.Close: Set RS = Nothing

vg_db.Execute ("sgpadm_iu_param 'UI', 'Mcc-" & Trim(vg_NUsr) & "', 'Parametro Mostrar Calorias Receta min.bloque', 'N', '" & IIf(IncluirColCalorias.Value = 1, 1, 0) & "'")

'-------> Leer datos
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_Sel_ParamTipoReceta 'Mcc-" & Trim(vg_NUsr) & "'")
If Not RS.EOF Then IncluirColCalorias.Value = RS!par_valor
RS.Close: Set RS = Nothing

'------- Grabar Parametro categoria dietetica

vg_db.Execute ("sgpadm_iu_param 'UI', 'CatDiePlan-" & Trim(vg_NUsr) & "', 'Parametro Categoria Dietetica Plan', 'N', '" & FilCatDie & "'")
'-------> Leer datos
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_Sel_ParamTipoReceta 'CatDiePlan-" & Trim(vg_NUsr) & "'")
If Not RS.EOF Then FilCatDie = RS!par_valor
RS.Close: Set RS = Nothing

'------- Grabar Parametro tipo plato
vg_db.Execute ("sgpadm_iu_param 'UI', 'TipPlaPlan-" & Trim(vg_NUsr) & "', 'Parametro Tipo Plato Plan', 'N', '" & FilTipPla & "'")

'-------> Leer datos
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_Sel_ParamTipoReceta 'TipPlaPlan-" & Trim(vg_NUsr) & "'")
If Not RS.EOF Then FilTipPla = RS!par_valor
RS.Close: Set RS = Nothing

Est = False
' Control displays text tips aligned to pointer with focus
   
    vaSpread1.TextTip = 2
    vaSpread1.TextTipDelay = 250
    X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
    vaSpread1.Visible = False
    vaSpread1.MaxRows = 0
    Toolbar1.Enabled = False
    Frame1.Enabled = False
    Label1(0).Caption = ""
    DoEvents
    Screen.MousePointer = 11
    Dim SeleccionOpt As Integer
    
    '-------> Ocultar columna costo
    vaSpread1.Col = 5 '4
    vaSpread1.ColHidden = IIf(IncluirColCosto.Value = 1, False, True)
    vaSpread1.Col = 7 '6
    vaSpread1.ColHidden = IIf(IncluirColCalorias.Value = 1, False, True)


    M_MinBloqueADM2.vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
    M_MinBloqueADM2.vaSpread1.Col = xColIni
    M_MinBloqueADM2.vaSpread1.Row = SpreadHeader + 3
    Fecha = CLng(Format(Mid(M_MinBloqueADM2.vaSpread1.text, 5, Len(M_MinBloqueADM2.vaSpread1.text)), "yyyymmdd"))
    SeleccionOpt = IIf(M_MinSR1.optPrecioConvenio.Value = True, 1, IIf(M_MinSR1.optPrecioGenerico = True, 2, 3))
    Me.Refresh
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS = vg_db.Execute("sgpadm_Sel_ResumenCostoReceta_V03 '" & vg_codcasino & "', " & vg_codregimen & ", " & vg_codservicio & ", " & FilCatDie & ", " & FilTipPla & ", " & Fecha & ", " & SeleccionOpt & ", '" & IIf(IncluirColCosto.Value = 1, 1, 0) & "'")
    vaSpread1.MaxRows = 0
    vaSpread1.MaxRows = RS.RecordCount
    IndRec = 1
    
    If Not RS.EOF Then
        
        DoEvents
        Screen.MousePointer = 11
        
        Do While Not RS.EOF
            
'            vaSpread1.MaxRows = vaSpread1.MaxRows + 1
            vaSpread1.Row = IndRec 'vaSpread1.MaxRows
    
            vaSpread1.Col = 1
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignRight
            vaSpread1.text = RS!rec_codigo
          
            vaSpread1.Col = 2
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignLeft
            vaSpread1.text = Trim(RS!rec_nombre)
          
            vaSpread1.Col = 3
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignLeft
            vaSpread1.text = Trim(RS!nom_catdie)
            
            vaSpread1.Col = 4 '3
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignLeft
            vaSpread1.text = Trim(RS!nom_tippla)
          
            vaSpread1.Col = 5 '4
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignRight
            vaSpread1.text = Format(RS!promedioreceta, fg_Pict(6, 2))
                  
            vaSpread1.Col = 6 '5
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignRight
            If IsNull(RS!rec_indppr) Or RS!rec_indppr = "" Then
               
               vaSpread1.text = ""
            
            Else
               
               vaSpread1.text = IIf(Trim(RS!rec_indppr) = "1", "Real", "Propuesta")
            
            End If
            
            vaSpread1.Col = 7
'            If vg_ActCalorias = True Then vaSpread1.ColHidden = False Else vaSpread1.ColHidden = True
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignRight
            vaSpread1.text = Format(IIf(IsNull(RS!AporteNut), 0, RS!AporteNut), fg_Pict(6, 2))
            
            vaSpread1.Col = 9
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignLeft
            vaSpread1.text = RS!descripcioncorta
            
            vaSpread1.Col = 10
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignLeft
            vaSpread1.text = RS!Estacionalidad_asoc

            IndRec = IndRec + 1
            
            RS.MoveNext
          
       Loop
       Label1(0).Visible = True
       Label1(0).Caption = "Registro " & Format(vaSpread1.MaxRows, fg_Pict(6, 0))
       RS.Close: Set RS = Nothing: fg_descarga
       fpTnombre_Change
    
    Else
       
       fg_descarga
       Label1(0).Visible = True
       Label1(0).Caption = "Registro " & Format(vaSpread1.MaxRows, fg_Pict(6, 0))
       RS.Close
       Set RS = Nothing
       fg_descarga
       Call MsgBox("No Existen Recetas", vbExclamation + vbOKOnly, "Busqueda Recetas")
    
    End If
    vaSpread1.Row = 1
    vaSpread1.Visible = True
    Toolbar1.Enabled = True
    Frame1.Enabled = True
    If Me.Visible Then
       FptNombre.SetFocus
    End If

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If KeyCode = 27 Then ICGrilla = 0: Me.Hide: Exit Sub
If TeclasNoPermitidas(KeyCode) = True Then FptNombre.text = IIf(KeyCode = 8, FptNombre.text, FptNombre.text & Chr(KeyCode)): FptNombre.SetFocus: FptNombre.SelStart = Len(FptNombre.text)

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub LoadForm()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim AuxFilCatDie  As Long
Dim AuxFilTipPla As Long

If FormActivado = True Then
    
    FormActivado = False
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
End If

vaSpread1.MaxRows = 0

iayuda = 0

If vg_filcatdieMin > 0 Then

   FilCatDie = vg_filcatdieMin
   AuxFilCatDie = vg_filcatdieMin
   
End If

If FilCatDie = 0 Then
   
   fpayuda(0).Caption = "Todos"

Else
   
   AuxFilCatDie = FilCatDie
   fpayuda(0).Caption = fg_BuscaenArbol(AuxFilCatDie, "a_recetacatdie", "car_codigo")

End If

'-------> leer datos parametros tipo plato
If vg_filtipplaMin > 0 Then
    
    FilTipPla = vg_filtipplaMin
    AuxFilTipPla = vg_filtipplaMin

End If

If FilTipPla = 0 Then
   
   fpayuda(1).Caption = "Todos"

Else
   
   AuxFilTipPla = FilTipPla
   fpayuda(1).Caption = fg_BuscaenArbol(AuxFilTipPla, "a_recetatippla", "tip_codigo")

End If

MoverRecetasGrilla

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread1.Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
Select Case Col
Case 1
    vaSpread1.Col = Col
    TipText = "Código : " & vaSpread1.text
Case 2
    vaSpread1.Col = Col
    TipText = "Nombre Plato : " & Trim(vaSpread1.text)
Case 3
    vaSpread1.Col = Col
    TipText = "Cat. Dietetica : " & Trim(vaSpread1.text)
Case 4 '3
    vaSpread1.Col = Col
    TipText = "Tipo Plato : " & Trim(vaSpread1.text)
Case 5 '4
    vaSpread1.Col = Col
    TipText = "Precio unitario : " & Trim(vaSpread1.text)
Case 6 '5
    vaSpread1.Col = Col
    TipText = "Tipo Receta : " & Trim(vaSpread1.text)
Case 7 '5
    vaSpread1.Col = Col
    TipText = "Calorias : " & Trim(vaSpread1.text)
End Select

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub
