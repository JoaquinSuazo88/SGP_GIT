VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form T_Mermas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Mermas"
   ClientHeight    =   4950
   ClientLeft      =   2940
   ClientTop       =   1110
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   971
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "T_Mermas.frx":0000
         Left            =   2010
         List            =   "T_Mermas.frx":000A
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
         Left            =   525
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
         Left            =   4590
         TabIndex        =   3
         Top             =   645
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3405
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   11550
      _Version        =   393216
      _ExtentX        =   20373
      _ExtentY        =   6006
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      AutoCalc        =   0   'False
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
      MaxCols         =   6
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "T_Mermas.frx":001E
      ScrollBarTrack  =   3
      ClipboardOptions=   0
   End
End
Attribute VB_Name = "T_Mermas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim OpGr As Boolean
Dim ibusca As Long
Dim MsgTitulo As String
Public lc_Aux  As String

Private Sub GrabaRegistro(Fila As Long)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim codigo As Long
Dim Nombre As String
Dim Activo As String
Dim MensajeNuevo As String
Dim MensajeAnterior As String

OpGr = True
vaSpread1.Row = Fila
vaSpread1.Col = 1
codigo = Val(vaSpread1.Value)

vaSpread1.Col = 2
Nombre = Trim(LimpiaDato(Mid(vaSpread1.Value, 1, 100)))

vaSpread1.Col = 6
Activo = Trim(LimpiaDato(Mid(vaSpread1.Value, 1, 1)))

If Trim(Nombre) = "" Then

   MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo
   vaSpread1.Row = Fila
   vaSpread1.Col = 2
   vaSpread1.SetActiveCell 2, vaSpread1.MaxRows
   vaSpread1.SetFocus
   OpGr = False
   Exit Sub

End If

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If modo = "A" Then
   
   codigo = 0
   Set RS = vg_db.Execute("sgpadm_Ins_TipoMermas '" & Trim(Nombre) & "', '" & vg_NUsr & "'")
   If Not RS.EOF Then
      
      If RS(0) = 0 Then
         
         codigo = RS(3)
         vaSpread1.Col = 1
         vaSpread1.Value = codigo
         
         vaSpread1.Col = 6
         vaSpread1.Value = "1"
         vaSpread1.Lock = False
         
         MensajeNuevo = codigo & ";" & Trim(Nombre) & ";" & 1 & ";" & vg_NUsr
         
         'registrar Log sistema Grabar
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Agregar"), Me.HelpContextID, MensajeNuevo, "", "")
         
         MostrarRegistro (Fila)
         MsgBox "Registro grabo exitosamente", vbInformation + vbOKOnly, MsgTitulo
      
      Else
         
         MensajeNuevo = Trim(Nombre) & ";" & 1 & ";" & vg_NUsr
         
         'registrar Log sistema error Grabar
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Agregado"), Me.HelpContextID, MensajeNuevo, "", "")
         
         MsgBox "Registro finalizo con error " & RS(0) & ":" & RS(1), vbInformation + vbOKOnly, MsgTitulo
                        
         vaSpread1.DeleteRows Fila, 1
         vaSpread1.MaxRows = vaSpread1.MaxRows - 1

      End If
      
   End If
   RS.Close
   Set RS = Nothing
   
Else
   
   MensajeAnterior = ""
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   '-------> INI : traer dato anterior
   Set RS = vg_db.Execute("sgpadm_Sel_CodigoTipoMermas " & codigo & "")
   If Not RS.EOF Then
       
      MensajeAnterior = codigo & ";" & IIf(IsNull(RS!aju_nombre), "", Trim(RS!aju_nombre)) & ";" & IIf(IsNull(RS!aju_Activo), "", Trim(RS!aju_Activo)) & ";" & vg_NUsr
        
   End If
   RS.Close
   Set RS = Nothing

   '-------> FIN : traer dato anterior
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Upd_TipoMermas " & codigo & ", '" & Trim(Nombre) & "', '" & Activo & "', '" & vg_NUsr & "'")

   If Not RS.EOF Then
      
      If RS(0) = 0 Then
           
         MensajeNuevo = codigo & ";" & Trim(Nombre) & ";" & Activo & ";" & vg_NUsr
         
         'registrar Log sistema Modificación
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificar"), Me.HelpContextID, MensajeNuevo, MensajeAnterior, "")
         
         MsgBox "Registro modifico exitosamente", vbInformation + vbOKOnly, MsgTitulo
           
      Else
         
         'registrar Log sistema error Modificación
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), Me.HelpContextID, MensajeNuevo, MensajeAnterior, "")
         
         MsgBox "Registro finalizo con error " & RS(0) & ":" & RS(1), vbInformation + vbOKOnly, MsgTitulo
                        
      End If
      
   End If
   RS.Close
   Set RS = Nothing

   MostrarRegistro (Fila)

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
Me.Height = 5385
Me.Width = 11820
MsgTitulo = "Tipo de Mermas"

fg_centra Me
modo = ""
ibusca = 0

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
   
   Frame1.Move 0, 360, 6015, 971
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

If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

vaSpread1.Visible = False
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    
   Set RS = vg_db.Execute("sgpadm_Sel_BuscarCodigoTipoMermas '%" & UCase(LimpiaDato(fpText1.text)) & "%'")

ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    
   Set RS = vg_db.Execute("sgpadm_Sel_BuscarNombreTipoMermas '%" & UCase(LimpiaDato(fpText1.text)) & "%'")

End If

If RS.EOF Then
   
   vaSpread1.MaxRows = 0
Else

   vaSpread1.MaxRows = RS.RecordCount

End If

i = 1

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      vaSpread1.Row = i
      i = i + 1
      
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = 1
      vaSpread1.Value = RS!aju_codigo
      
      vaSpread1.Col = 2
      vaSpread1.Value = IIf(IsNull(RS!Nombre), "", Trim(RS!Nombre))
      
      vaSpread1.Col = 3
      vaSpread1.Lock = True
      vaSpread1.Value = IIf(IsNull(RS!Fecha_Creacion), "", Trim(RS!Fecha_Creacion))
    
      vaSpread1.Col = 4
      vaSpread1.Lock = True
      vaSpread1.Value = IIf(IsNull(RS!Fecha_Modificacion), "", Trim(RS!Fecha_Modificacion))
    
      vaSpread1.Col = 5
      vaSpread1.Lock = True
      vaSpread1.Value = IIf(IsNull(RS!usuario), "", Trim(RS!usuario))
      
      vaSpread1.Col = 6
      vaSpread1.Lock = False
      vaSpread1.Value = IIf(IsNull(RS!Activo), "", Trim(RS!Activo))
     
      RS.MoveNext
   
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

Dim codigo As Long
Dim Nombre As String
Dim Activo As String
Dim Mensaje As String
Dim orden As String
Dim RS As New ADODB.Recordset

On Error GoTo Man_Error

Select Case Button.Index

Case 1
    
    'registrar Log sistema preparando Agregar
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Agregar"), Me.HelpContextID, "", "", "")
    
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    
    vaSpread1.Col = 6
    vaSpread1.Lock = True
    
    vaSpread1.Col = 2
    vaSpread1.SetActiveCell 2, vaSpread1.MaxRows
    vaSpread1.SetFocus

Case 3
    
    'registrar Log sistema preparando Modificación
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), Me.HelpContextID, "", "", "")
    
    modo = "M"
    If vaSpread1.MaxRows < 1 Then Exit Sub
    Gl_Ac_Botones Me, 1, 0, modo

Case 5
    
    If vaSpread1.ActiveRow < 1 Then
       
       MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    codigo = Val(vaSpread1.Value)
    
    vaSpread1.Col = 2
    Nombre = Trim(vaSpread1.text)
        
    vaSpread1.Col = 6
    Activo = vaSpread1.text
    
    Mensaje = ""
    Mensaje = codigo & ";" & Nombre & ";" & Activo & vg_NUsr
    
    'registrar Log sistema preparando Desactivar
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Eliminar"), Me.HelpContextID, Mensaje, "", "")
    
    If MsgBox("Desactiva registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
    
       'registrar Log sistema Anula cambio
       Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Anular"), Me.HelpContextID, Mensaje, "", "")
       
       Exit Sub
    
    End If
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    codigo = Val(vaSpread1.Value)

    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Del_TipoMermas " & codigo & ", '0', '" & vg_NUsr & "'")

    If Not RS.EOF Then
      
       If Trim(RS(0)) <> "" Then
           
          'registrar Log sistema desactivar el item
          Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminar"), Me.HelpContextID, Mensaje, "", "")
          
          MsgBox "Registro desactivado exitosamente", vbInformation + vbOKOnly, MsgTitulo
           
       Else
         
          'registrar Log sistema error desactivar el item
          Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, Mensaje, "", "")
          
          MsgBox "Registro finalizo con error " & RS(0), vbInformation + vbOKOnly, MsgTitulo
                        
       End If
      
    End If
    RS.Close
    Set RS = Nothing

    Cancela

'    vaSpread1.DeleteRows vaSpread1.Row, 1
'    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    
    modo = ""
    Gl_Ac_Botones Me, 1, 1, modo

Case 7
    
    fpText1.text = ""
    MoverDatosGrillas

Case 10
    
    If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
    
       Exit Sub
    
    End If
    
   'registrar Log sistema cancela proceso agregado
   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Cancelar"), Me.HelpContextID, "", "", "")
    
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
    
    'registrar Log sistema imprimir
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Informe"), Me.HelpContextID, "", "", "")
    
    I_TipoMermas

Case 18
    
    'registrar Log sistema salir de la opción
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "", "")
    
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

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

If (Col <> 6) Or Row = 0 Or OpGr Then Exit Sub

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

vaSpread1.Visible = False
vaSpread1.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_TipoMermas")

Do While Not RS.EOF
    
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    
    vaSpread1.Col = 1
    vaSpread1.Value = RS!aju_codigo
    
    vaSpread1.Col = 2
    vaSpread1.Value = IIf(IsNull(RS!Nombre), "", Trim(RS!Nombre))
    
    vaSpread1.Col = 3
    vaSpread1.Lock = True
    vaSpread1.Value = IIf(IsNull(RS!Fecha_Creacion), "", Trim(RS!Fecha_Creacion))
    
    vaSpread1.Col = 4
    vaSpread1.Lock = True
    vaSpread1.Value = IIf(IsNull(RS!Fecha_Modificacion), "", Trim(RS!Fecha_Modificacion))
    
    vaSpread1.Col = 5
    vaSpread1.Lock = True
    vaSpread1.Value = IIf(IsNull(RS!usuario), "", Trim(RS!usuario))
    
    vaSpread1.Col = 6
    vaSpread1.Lock = False
    vaSpread1.Value = IIf(IsNull(RS!Activo), "", Trim(RS!Activo))
    
    RS.MoveNext

Loop

RS.Close
Set RS = Nothing

Gl_Ac_Botones Me, 1, 1, modo
vaSpread1.Visible = True
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"

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

Set RS = vg_db.Execute("sgpadm_Sel_CodigoTipoMermas " & codigo & "")
If Not RS.EOF Then
   
   vaSpread1.Col = 2
   vaSpread1.Value = IIf(IsNull(RS!aju_nombre), "", Trim(RS!aju_nombre))

    vaSpread1.Col = 3
    vaSpread1.Lock = True
    vaSpread1.Value = IIf(IsNull(RS!Fecha_Creacion), "", Trim(RS!Fecha_Creacion))
    
    vaSpread1.Col = 4
    vaSpread1.Lock = True
    vaSpread1.Value = IIf(IsNull(RS!aju_fecmod), "", Trim(RS!aju_fecmod))
    
    vaSpread1.Col = 5
    vaSpread1.Lock = True
    vaSpread1.Value = IIf(IsNull(RS!usuario), "", Trim(RS!usuario))
    
    vaSpread1.Col = 6
    vaSpread1.Lock = False
    vaSpread1.Value = IIf(IsNull(RS!aju_Activo), "", Trim(RS!aju_Activo))
    
End If
RS.Close
Set RS = Nothing
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub MostrarRegistro(IRow As Long)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim codigo As Long

OpGr = True
vaSpread1.Row = IRow
vaSpread1.Col = 1
codigo = Val(vaSpread1.Value)

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_CodigoTipoMermas " & codigo & "")
If Not RS.EOF Then
   
   vaSpread1.Col = 2
   vaSpread1.Value = IIf(IsNull(RS!aju_nombre), "", Trim(RS!aju_nombre))

   vaSpread1.Col = 3
   vaSpread1.Lock = True
   vaSpread1.Value = IIf(IsNull(RS!Fecha_Creacion), "", Trim(RS!Fecha_Creacion))
    
   vaSpread1.Col = 4
   vaSpread1.Lock = True
   vaSpread1.Value = IIf(IsNull(RS!aju_fecmod), "", Trim(RS!aju_fecmod))
    
   vaSpread1.Col = 5
   vaSpread1.Lock = True
   vaSpread1.Value = IIf(IsNull(RS!usuario), "", Trim(RS!usuario))
    
   vaSpread1.Col = 6
   vaSpread1.Lock = False
   vaSpread1.Value = IIf(IsNull(RS!aju_Activo), "", Trim(RS!aju_Activo))
    
End If
RS.Close
Set RS = Nothing
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub


