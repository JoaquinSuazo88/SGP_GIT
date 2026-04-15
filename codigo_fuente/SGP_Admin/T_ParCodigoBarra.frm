VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form T_ParCodigoBarra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros Lectura Vales"
   ClientHeight    =   4980
   ClientLeft      =   5370
   ClientTop       =   2280
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   971
      Left            =   930
      TabIndex        =   0
      Top             =   360
      Width           =   5865
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "T_ParCodigoBarra.frx":0000
         Left            =   1680
         List            =   "T_ParCodigoBarra.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2500
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1680
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
         Left            =   255
         TabIndex        =   5
         Top             =   345
         Width           =   1380
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
         Left            =   255
         TabIndex        =   4
         Top             =   645
         Width           =   1140
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
         Left            =   4260
         TabIndex        =   3
         Top             =   645
         Width           =   585
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3435
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   7695
      _Version        =   393216
      _ExtentX        =   13573
      _ExtentY        =   6059
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
      MaxCols         =   2
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "T_ParCodigoBarra.frx":001E
      ScrollBarTrack  =   3
      ClipboardOptions=   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "T_ParCodigoBarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim OpGr As Boolean
Public lc_Aux As String

Private Sub GrabaRegistro(Fila As Long)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim descripcion As String
Dim codigo As Long
Dim i As Long
Dim OpcionUsuario As Boolean

OpGr = True
vaSpread1.Row = Fila
vaSpread1.Col = 1
codigo = IIf(Trim(vaSpread1.text) <> "", vaSpread1.Value, 0)

vaSpread1.Col = 2
descripcion = Trim(LimpiaDato(Mid(vaSpread1.Value, 1, 100)))
For i = 1 To vaSpread1.MaxRows
    
    vaSpread1.Col = 2: vaSpread1.Row = i
    If UCase(Trim(vaSpread1.text)) = UCase(Trim(descripcion)) And Fila <> i Then MsgBox "Descripción ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.Row: vaSpread1.SetFocus: OpGr = False: Exit Sub

Next i

If Trim(descripcion) = "" Then MsgBox "Favor ingresar descripción, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.Row: vaSpread1.SetFocus: OpGr = False: Exit Sub

If modo = "A" Then
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Ins_ParametroCodigoBarra '" & Trim(descripcion) & "'")
   
   If Not RS.EOF Then
      
      codigo = RS!atr_codigo_barra
      vaSpread1.Col = 1
      vaSpread1.Value = codigo
   
   End If
   RS.Close
   Set RS = Nothing

Else
   
   vg_db.Execute ("sgpadm_Upd_AtributoCodigoBarra " & codigo & ", '" & Trim(descripcion) & "'")

End If
OpcionUsuario = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", falso, True)
vaSpread1.Col = 2
vaSpread1.Lock = OpcionUsuario

If modo = "A" Then
   
   MsgBox "Registro guardo exitosamente", vbInformation + vbOKOnly, MsgTitulo

Else
   
   MsgBox "Registro modificado exitosamente", vbInformation + vbOKOnly, MsgTitulo

End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Combo1.Enabled = True
fpText1.Enabled = True
modo = ""
Gl_Ac_Botones Me, 1, 1, modo
OpGr = False

Exit Sub
Man_Error:
If Err.Number = -2147467259 Then
    
    MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"

Else
    
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"

End If

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
Me.Height = 5380
Me.Width = 8850
MsgTitulo = "Parametro Código Barra"
fg_centra Me
modo = ""
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
   
   Frame1.Move 1200, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440

ElseIf Me.WindowState = 2 Then
   
   Frame1.Move 4200, 360, 6015, 971
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
Dim i As Long
Dim OpcionUsuario As Boolean
If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
OpcionUsuario = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", falso, True)
vaSpread1.Visible = False
vaSpread1.MaxRows = 0

If Combo1.ItemData(Combo1.ListIndex) = 0 Then
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   Set RS = vg_db.Execute("sgpadm_Sel_ParametroCodigoBarraxCodigo '%" & UCase(LimpiaDato(fpText1.text)) & "%'")

ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
   
   Set RS = vg_db.Execute("sgpadm_Sel_ParametroCodigoBarraxNombre '%" & UCase(LimpiaDato(fpText1.text)) & "%'")

End If

i = 1
If Not RS.EOF Then
   
   vaSpread1.MaxRows = RS!nReg
   
   Do While Not RS.EOF
      
      vaSpread1.Row = i
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = 1
      vaSpread1.Value = RS!atr_codigo_barra
      
      vaSpread1.Col = 2
      vaSpread1.Lock = OpcionUsuario
      vaSpread1.Value = Trim(RS!atr_nombre)
      
      RS.MoveNext
      i = i + 1
   
   Loop
   Gl_Ac_Botones Me, 1, 1, modo

End If
RS.Close
Set RS = Nothing

vaSpread1.Visible = True
If fpText1.text = "" Then
   
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"

Else
   
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim codigo As Long, descripcion As String

Select Case Button.Index

    Case 1
        
        modo = "A"
        Gl_Ac_Botones Me, 1, 0, modo
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 2
        vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus
    
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
        
        vg_db.Execute ("sgpadm_Del_ParametroCodigoBarra " & codigo & "")
        
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
        modo = ""
        Gl_Ac_Botones Me, 1, 1, modo
        MsgBox "Registro eliminado exitosamente", vbInformation + vbOKOnly, MsgTitulo
    
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
        I_ParametroCodigoBarra
    
    Case 18
        
        Me.Hide
        Unload Me

End Select

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
If Err.Number = -2147467259 Or Err.Number = -2147217900 Then
   
   MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"

Else
    
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"

End If

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If Col <> 4 Or OpGr Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub MoverDatosGrillas()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim OpcionUsuario As Boolean

OpGr = True
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
OpcionUsuario = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", falso, True)

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_Sel_ParametroCodigoBarra 0")
Do While Not RS.EOF
    
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    
    vaSpread1.Col = 1
    vaSpread1.Value = RS!atr_codigo_barra
    
    vaSpread1.Col = 2
    vaSpread1.Lock = OpcionUsuario
    vaSpread1.Value = Trim(RS!atr_nombre)
    
    RS.MoveNext

Loop
RS.Close
Set RS = Nothing
Gl_Ac_Botones Me, 1, 1, modo
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
vaSpread1.Visible = True
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo Man_Error

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
Set RS = vg_db.Execute("sgpadm_Sel_ParametroCodigoBarra " & codigo & "")
If Not RS.EOF Then
    
    vaSpread1.Col = 2
    vaSpread1.Value = Trim(RS!atr_nombre)

End If
RS.Close
Set RS = Nothing
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub
