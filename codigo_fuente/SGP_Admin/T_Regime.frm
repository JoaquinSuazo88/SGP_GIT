VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form T_Regime 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Régimen"
   ClientHeight    =   4965
   ClientLeft      =   2835
   ClientTop       =   2505
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   971
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "T_Regime.frx":0000
         Left            =   2010
         List            =   "T_Regime.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2500
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   2010
         TabIndex        =   2
         Top             =   555
         Width           =   2505
         _Version        =   196608
         _ExtentX        =   4410
         _ExtentY        =   870
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483637
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
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
         Left            =   4590
         TabIndex        =   5
         Top             =   645
         Width           =   585
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
         Left            =   525
         TabIndex        =   4
         Top             =   645
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Columna"
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
         Left            =   525
         TabIndex        =   3
         Top             =   345
         Width           =   1380
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3405
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   7680
      _Version        =   393216
      _ExtentX        =   13547
      _ExtentY        =   6006
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      AutoCalc        =   0   'False
      ButtonDrawMode  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
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
      FormulaSync     =   0   'False
      MaxCols         =   5
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "T_Regime.frx":001E
      ScrollBarTrack  =   3
      ClipboardOptions=   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "T_Regime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim OpGr As Boolean
Dim ibusca As Long
Dim vTipoReg() As Variant 'en  general

Private Sub GrabaRegistro(Fila As Long)

On Error GoTo Man_Error

Dim RS     As New ADODB.Recordset
Dim codigo As Long, Nombre As String, Activo As String, Indicador As String, ind As String
Dim codaux As Long, z As Long

OpGr = True
vaSpread1.Row = Fila
vaSpread1.Col = 1
codigo = Val(vaSpread1.Value)

vaSpread1.Col = 2
Nombre = Trim(LimpiaDato(Mid(vaSpread1.Value, 1, 50)))

vaSpread1.Col = 3
Activo = Trim(LimpiaDato(vaSpread1.Value))

If vg_Indppr = "1" Or vg_Indppr = "2" Then
    
    vaSpread1.Col = 4
    If Trim(Nombre) = "" Or Trim(vaSpread1.text) = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub
    Indicador = IIf(vg_Indppr = "1", "1", "2")

Else
    
    vaSpread1.Col = 4
    If Trim(Nombre) = "" Or Trim(vaSpread1.text) = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub
    Indicador = IIf(vaSpread1.TypeComboBoxCurSel = 0, 1, 2)

End If

If Trim(Nombre) = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub

If modo = "A" Then
   
   codigo = 0
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_iu_regimen 'A', 0, '" & Trim(Nombre) & "', '" & Activo & "','" & Indicador & "'")
   If Not RS.EOF Then
   
      codigo = RS!indice
      vaSpread1.Col = 1
      vaSpread1.Value = codigo
   
   End If
   RS.Close
   Set RS = Nothing

Else
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("SELECT DISTINCT reg_indppr FROM a_regimen WHERE reg_codigo = " & codigo & "")
    
    If Not RS.EOF Then
       
       ind = RS!reg_indppr
       RS.Close
       Set RS = Nothing
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       Set RS = vg_db.Execute("SELECT DISTINCT min_codreg, min_indppr FROM b_minuta WHERE min_codreg = " & codigo & " AND min_indppr = '" & ind & "'")
       If Not RS.EOF Then
          
          If RS!min_indppr <> Indicador Then
             
             RS.Close
             Set RS = Nothing
             vaSpread1.Col = 5
             codaux = -1
             
             For z = 0 To vaSpread1.TypeComboBoxCount
                 
                 vaSpread1.TypeComboBoxCurSel = z
                 If vaSpread1.text = ind Then codaux = z: Exit For
                 codaux = -1
             
             Next z
             
             vaSpread1.Col = 4
             vaSpread1.TypeComboBoxCurSel = codaux
             Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
             Combo1.Enabled = True
             fpText1.Enabled = True
             modo = ""
             Gl_Ac_Botones Me, 1, 1, modo
             MsgBox "No se puede actualizar regimen, ya que existe minuta asociada...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False
             Exit Sub
          
          End If
       
       End If
    
    End If
    RS.Close
    Set RS = Nothing
    vg_db.Execute "sgpadm_iu_regimen 'M', " & codigo & ", '" & Trim(Nombre) & "', '" & Activo & "', '" & Indicador & "'"
    vaSpread1.Col = 5
    vaSpread1.Value = Indicador
    
End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Combo1.Enabled = True
fpText1.Enabled = True
modo = ""
Gl_Ac_Botones Me, 1, 1, modo
OpGr = False

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
Me.Height = 5445
'Me.Width = 8055
MsgTitulo = "Regimen"
fg_centra Me
modo = ""
ibusca = 0
'--->Carga en la Form Load
ReDim vTipoReg(2, 2)
If vg_Indppr = "1" Or vg_Indppr = "2" Then
  
  vTipoReg(1, 1) = IIf(vg_Indppr = "1", "1", "2")
  vTipoReg(1, 2) = IIf(vg_Indppr = "1", "Real", "Propuesta")

Else
  
  vTipoReg(1, 1) = 1
  vTipoReg(1, 2) = "Real"
  vTipoReg(2, 1) = 2
  vTipoReg(2, 2) = "Propuesta"

End If

Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1
MoverDatosGrillas
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Resize()

On Error GoTo Man_Error

If Me.WindowState = 0 Then
   
   Frame1.Move 240, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440

ElseIf Me.WindowState = 2 Then
   
   Frame1.Move 5800, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440

End If
Toolbar1.Refresh

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpText1_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
vaSpread1.Visible = False

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
If vg_Indppr = "1" Or vg_Indppr = "2" Then
    
   If Combo1.ItemData(Combo1.ListIndex) = 0 Then
        
      Set RS = vg_db.Execute("sgpadm_s_regimen 3, " & vg_Indppr & ", '%" & UCase(LimpiaDato(fpText1.text)) & "%'")
    
   ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
        
      Set RS = vg_db.Execute("sgpadm_s_regimen 4, " & vg_Indppr & ", '%" & UCase(LimpiaDato(fpText1.text)) & "%'")
    
   End If

Else
    
   If Combo1.ItemData(Combo1.ListIndex) = 0 Then
        
      Set RS = vg_db.Execute("sgpadm_s_regimen 5, 0, '%" & UCase(LimpiaDato(fpText1.text)) & "%'")
    
   ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
        
      Set RS = vg_db.Execute("sgpadm_s_regimen 6, 0, '%" & UCase(LimpiaDato(fpText1.text)) & "%'")
    
   End If

End If
If RS.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS!nReg
i = 1
If Not RS.EOF Then
   OpGr = True
   
   Do While Not RS.EOF
      
      vaSpread1.Row = i
      i = i + 1
      
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = 1
      vaSpread1.Value = RS!reg_codigo
      
      vaSpread1.Col = 2
      vaSpread1.Value = IIf(IsNull(RS!reg_nombre), "", Trim(RS!reg_nombre))
      
      vaSpread1.Col = 3
      vaSpread1.Value = IIf(IsNull(RS!reg_activo), "0", RS!reg_activo)
      
      For z = 1 To UBound(vTipoReg)
      
          vaSpread1.Col = 4
          lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vTipoReg(z, 2))
          
          vaSpread1.Col = 5
          liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vTipoReg(z, 1)
          
          vaSpread1.Col = 4
          vaSpread1.TypeComboBoxList = lisnom
          
          vaSpread1.Col = 5
          vaSpread1.TypeComboBoxList = liscod
      
      Next z
      
      vaSpread1.Col = 5
      codaux = -1
      For z = 0 To vaSpread1.TypeComboBoxCount
          
          vaSpread1.TypeComboBoxCurSel = z
          If vaSpread1.text = IIf(IsNull(RS!reg_indppr), 0, RS!reg_indppr) Then codaux = z: Exit For
          codaux = -1
      
      Next z
      vaSpread1.Col = 4
      vaSpread1.TypeComboBoxCurSel = codaux
    
      RS.MoveNext
   Loop
   OpGr = False
   Gl_Ac_Botones Me, 1, 1, modo

End If
RS.Close
Set RS = Nothing
vaSpread1.Visible = True
If fpText1.text = "" Then Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro" Else Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim codigo As Long, Nombre As String, orden As String, j As Integer, z As Integer

Select Case Button.Index
    
    Case 1
        
        modo = "A"
        Gl_Ac_Botones Me, 1, 0, modo
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        lisnom = ""
        liscod = ""
        
        For j = 1 To UBound(vTipoReg)
            
            If vTipoReg(j, 1) <> "" Then
               
               lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vTipoReg(j, 2))
               liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vTipoReg(j, 1)
            
            End If
        
        Next j
        
       If vg_Indppr = 1 Or vg_Indppr = 2 Then
         
          lisnom = IIf(vg_Indppr = "1", "Real", "Propuesta"): liscod = IIf(vg_Indppr = "1", "1", "2")
          vaSpread1.Col = 4
          vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
          
          vaSpread1.Col = 5
          vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")
       
       Else
          
          vaSpread1.Col = 4
          vaSpread1.TypeComboBoxList = lisnom
          
          vaSpread1.Col = 5
          vaSpread1.TypeComboBoxList = liscod
       
       End If
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 4
        vaSpread1.Lock = False: vaSpread1.TypeComboBoxList = lisnom
        
        vaSpread1.Col = 5
        vaSpread1.TypeComboBoxList = liscod
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 2
        vaSpread1.SetActiveCell 2, vaSpread1.MaxRows
        vaSpread1.SetFocus
        'vaSpread1.Col = 4: vaSpread1.TypeComboBoxCurSel = 0
        
    Case 3
    
        modo = "M"
        If vaSpread1.MaxRows < 1 Then Exit Sub
        Gl_Ac_Botones Me, 1, 0, modo
        
    Case 5
        
        If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = Val(vaSpread1.Value)
        
        vg_db.Execute "DELETE a_regimen FROM a_regimen WHERE reg_codigo=" & codigo
        
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        modo = ""
        Gl_Ac_Botones Me, 1, 1, modo
        
    Case 7
        
        fpText1.text = ""
        MoverDatosGrillas
        
    Case 10
        
        If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        If modo = "A" Then
            
           vaSpread1.Row = vaSpread1.ActiveRow
           vaSpread1.DeleteRows vaSpread1.Row, 1
           vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        
        Else
            
           Cancela
        
        End If
        modo = ""
        Gl_Ac_Botones Me, 1, 1, modo
        Combo1.Enabled = True
        fpText1.Enabled = True
    
    Case 12
        
        GrabaRegistro vaSpread1.ActiveRow
    
    Case 15
    
        If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        I_Regime
        
    Case 18
        
        Me.Hide
        Unload Me
        
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Or 2147217900 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If modo = "" Then modo = "M"
If vg_Indppr = 1 Or vg_Indppr = 2 Then
 
 vaSpread1.Col = 4
 vaSpread1.TypeComboBoxClear Col, Row
 vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
 
 vaSpread1.Col = 5
 vaSpread1.TypeComboBoxClear Col, Row
 vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")

End If
Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

If vg_Indppr = 1 Or vg_Indppr = 2 Then
 
 vaSpread1.Col = 4
 vaSpread1.TypeComboBoxClear Col, Row
 vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
 
 vaSpread1.Col = 5
 vaSpread1.TypeComboBoxClear Col, Row
 vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")

End If
If Col <> 3 Or Row = 0 Or OpGr Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo Man_Error

If vg_Indppr = 1 Or vg_Indppr = 2 Then
 
 vaSpread1.Col = 4
 vaSpread1.TypeComboBoxClear Col, Row
 vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
 
 vaSpread1.Col = 5
 vaSpread1.TypeComboBoxClear Col, Row
 vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")

End If

If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    
   GrabaRegistro Row

ElseIf Toolbar1.Buttons(12).Visible = False Then
    
   Cancela

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Select Case Col
    
    Case 4
      
      If vg_Indppr = 1 Or vg_Indppr = 2 Then
       
         vaSpread1.Col = 4
         vaSpread1.TypeComboBoxClear Col, Row
         vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
         
         vaSpread1.Col = 5
         vaSpread1.TypeComboBoxClear Col, Row
         vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")
      
      End If
        
        Dim indice As Long
        vaSpread1.Row = Row
        
        vaSpread1.Col = 4
        indice = vaSpread1.TypeComboBoxCurSel
        
        vaSpread1.Col = 5
        vaSpread1.TypeComboBoxCurSel = indice
        
        If modo = "" Then modo = "M"
        Gl_Ac_Botones Me, 1, 0, modo
        vaSpread1.EditEnterAction = EditEnterActionNone

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub MoverDatosGrillas()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

vaSpread1.Visible = False
vaSpread1.MaxRows = 0
OpGr = True
lisnom = ""

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
If vg_Indppr = "1" Or vg_Indppr = "2" Then
   
   Set RS = vg_db.Execute("SELECT reg_codigo, reg_nombre, reg_activo, reg_indppr FROM a_regimen  where reg_indppr=" & vg_Indppr & " ORDER BY reg_codigo")

Else
   
   Set RS = vg_db.Execute("SELECT reg_codigo, reg_nombre, reg_activo, reg_indppr FROM a_regimen ORDER BY reg_codigo")

End If

Do While Not RS.EOF
    
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    
    vaSpread1.Col = 1
    vaSpread1.Value = RS!reg_codigo
    
    vaSpread1.Col = 2
    vaSpread1.Value = IIf(IsNull(RS!reg_nombre), "", Trim(RS!reg_nombre))
    
    vaSpread1.Col = 3
    vaSpread1.Value = IIf(IsNull(RS!reg_activo), "0", RS!reg_activo)
    
    lisnom = ""
    liscod = ""
    cParam = ""
    encuentra = False
    
    For j = 1 To UBound(vTipoReg)
        
        If vTipoReg(j, 1) <> "" Then
           
           'vaSpread1.Col = 4:
           lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vTipoReg(j, 2))
           'vaSpread1.Col = 5:
           liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vTipoReg(j, 1)
        
        End If
    
    Next j
    
    vaSpread1.Col = 4
    vaSpread1.TypeComboBoxList = lisnom
    
    vaSpread1.Col = 5
    vaSpread1.TypeComboBoxList = liscod
    
    vaSpread1.Col = 5
    codaux = -1
    For z = 0 To vaSpread1.TypeComboBoxCount
        
        vaSpread1.TypeComboBoxCurSel = z
        If vaSpread1.text = IIf(RS!reg_indppr = "1", "1", "2") Then codaux = z: Exit For
        codaux = -1
    
    Next z
    vaSpread1.Col = 4
    vaSpread1.TypeComboBoxCurSel = codaux
    
    RS.MoveNext

Loop
RS.Close
Set RS = Nothing

OpGr = False
Gl_Ac_Botones Me, 1, 1, modo
vaSpread1.Visible = True
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Cancela()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Dim codigo As Long
OpGr = True
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = Val(vaSpread1.Value)

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("SELECT reg_codigo, reg_nombre, reg_activo FROM a_regimen WHERE reg_codigo=" & codigo & "")
If Not RS.EOF Then

   vaSpread1.Col = 2
   vaSpread1.Value = IIf(IsNull(RS!reg_nombre), "", Trim(RS!reg_nombre))
   
   vaSpread1.Col = 3
   vaSpread1.Value = IIf(IsNull(RS!reg_activo), "0", RS!reg_activo)

End If
RS.Close
Set RS = Nothing
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub
