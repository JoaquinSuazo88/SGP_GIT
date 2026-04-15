VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_AsoFamSGPGrupoDespacho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asociar Famlilia SGP & Grupo Despacho"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9180
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
   ScaleHeight     =   8955
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8895
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   7815
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   8535
         _Version        =   393216
         _ExtentX        =   15055
         _ExtentY        =   13785
         _StockProps     =   64
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   20
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "M_AsoFamSGPGrupoDespacho.frx":0000
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_AsoFamSGPGrupoDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo      As String
Dim MsgTitulo As String
Dim Est       As Boolean
Public lc_Aux As String

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
MsgTitulo = "Asociar Familia SGP & Grupo Despacho"
fg_centra Me
modo = ""

Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 13, modo

Est = True

Moverdetalle

Est = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Moverdetalle()

On Error GoTo Man_Error

Dim RS                  As New ADODB.Recordset
Dim codaux              As Long
Dim lisnom              As String
Dim liscod              As String
Dim i                   As Long
Dim vLisGrupoDespacho() As Variant

fg_carga ""

'-------> Mover grupo despacho
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_GrupoDespacho")

i = 1

If Not RS.EOF Then
   
   ReDim vLisGrupoDespacho(RS.RecordCount, 2)
   
   Do While Not RS.EOF
      
      vLisGrupoDespacho(i, 1) = RS!idgrupodespacho
      vLisGrupoDespacho(i, 2) = RS!Nombre
      i = i + 1
      RS.MoveNext
   
   Loop

End If

RS.Close
Set RS = Nothing

vaSpread1.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_FamiliaSGPGrupoDespacho")

Do While Not RS.EOF
             
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   
   vaSpread1.Col = 1 'codigo familia sgp
   
   If RS("tip_nivel") = 0 Then
      
      vaSpread1.Font.Bold = True
   
   End If
   
   vaSpread1.CellType = CellTypeStaticText
   vaSpread1.text = RS("tip_codigo")
                 
   vaSpread1.Col = 2 ' Nombre familia sgp
   
   If RS("tip_nivel") = 0 Then
      
      vaSpread1.Font.Bold = True
   
   End If
   
   vaSpread1.CellType = CellTypeStaticText
   vaSpread1.text = Space(RS("tip_nivel") + IIf(RS("tip_nivel") = 0, 0, 5)) & RS("tip_nombre")
        
   lisnom = ""
   liscod = ""
   
   '-------> Mover grupo despacho
   For i = 1 To UBound(vLisGrupoDespacho)
       
       vaSpread1.Col = 3
       lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vLisGrupoDespacho(i, 2))
       
       vaSpread1.Col = 4
       liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vLisGrupoDespacho(i, 1)
       
       vaSpread1.Col = 3
       vaSpread1.TypeComboBoxList = lisnom
       
       vaSpread1.Col = 4
       vaSpread1.TypeComboBoxList = liscod
   
   Next i
   
   vaSpread1.Col = 4
   codaux = -1
   
   For i = 0 To vaSpread1.TypeComboBoxCount
       
       vaSpread1.TypeComboBoxCurSel = i
       If vaSpread1.text = RS!grupodespacho Then
       
          codaux = i
          Exit For
          
       End If
       codaux = -1
   
   Next i
   vaSpread1.Col = 3
   vaSpread1.TypeComboBoxCurSel = codaux

   RS.MoveNext
     
Loop
    
RS.Close
Set RS = Nothing

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim i        As Long
Dim codfam   As Long
Dim codgru   As Long
Dim MyBuffer As String

Select Case Button.Index
    
    Case 3 '-------> Modificar
        
        'registrar Log sistema modificar
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), Me.HelpContextID, "", "", "")
        
        modo = "M"
        Gl_Ac_Botones Me, 1, 0, modo
    
    Case 7, 10 '-------> Actualizar lista y cancelar
        
        'registrar Log sistema actualizar lista
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Actualizar"), Me.HelpContextID, "", "", "")
        
        Moverdetalle
        Gl_Ac_Botones Me, 1, 13, modo
    
    Case 12 '------> Confirmar
        
    fg_carga ""
    
    'registrar Log sistema modificar
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), CStr(Me.HelpContextID), "", "", "")
    
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<GrabaFamDes>"
    
    For i = 1 To vaSpread1.MaxRows

        vaSpread1.Row = i
        vaSpread1.Col = 3
        
        If Trim(vaSpread1.text) <> "" Then
           
           vaSpread1.Col = 1: codfam = vaSpread1.text
           vaSpread1.Col = 4: codgru = vaSpread1.text
        
           MyBuffer = MyBuffer & " <FamDes"
           MyBuffer = MyBuffer & " codfam = " & Chr(34) & codfam & Chr(34)
           MyBuffer = MyBuffer & " codgru = " & Chr(34) & codgru & Chr(34)
           MyBuffer = MyBuffer & "/>"
        
        End If
    
    Next i
    
    MyBuffer = MyBuffer & "</GrabaFamDes>"

    Set RS = vg_db.Execute("sgpadm_Ins_XmlParametroGrupoDespacho '" & MyBuffer & "'")
    If Not RS.EOF Then
            
       If Trim(RS(1)) <> "" Then
                  
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
        
       Else
       
          MsgBox "Generación grabado finalizado sin problema", vbInformation + vbOKOnly, MsgTitulo
               
       End If
            
    End If
    RS.Close: Set RS = Nothing
    
    'registrar Log sistema actualizar lista
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")
    
    modo = ""
    Gl_Ac_Botones Me, 1, 13, modo

    fg_descarga
    
    Case 15 '-------> Imprimir
        
        'registrar Log sistema actualizar lista
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Imprimir"), Me.HelpContextID, "", "", "")
        
        I_ParametroGrupoDespacho
    
    Case 18 '-------> Salir
       
       Me.Hide
       Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim indice As Long

Select Case Col

    Case 3
    
        vaSpread1.Row = Row
        vaSpread1.Col = 3
        indice = vaSpread1.TypeComboBoxCurSel
    
        vaSpread1.Col = 4
        vaSpread1.TypeComboBoxCurSel = indice
    
        If modo = "" Then modo = "M"
    
        If Toolbar1.Buttons(12).Visible = False Then
       
           Gl_Ac_Botones Me, 1, 0, modo
    
        End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If modo = "" Then modo = "M"

Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub

Select Case KeyCode
    Case 46
        
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = vaSpread1.ActiveCol
        
        If vaSpread1.Col <> 3 Then Exit Sub
        
        vaSpread1.text = ""
        vaSpread1.TypeComboBoxCurSel = -1
        If modo = "" Then modo = "M"
        If Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub
