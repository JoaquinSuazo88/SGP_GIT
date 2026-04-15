VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form M_Generar_Archivo_Rutas 
   Caption         =   "Generar Archivo Rutas"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   400
      Left            =   3120
      TabIndex        =   5
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar"
      Height          =   400
      Left            =   840
      TabIndex        =   4
      Top             =   3360
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   106692609
      CurrentDate     =   41731
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   3240
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   106692609
      CurrentDate     =   41731
   End
   Begin EditLib.fpText fpText1 
      Height          =   315
      Index           =   0
      Left            =   3240
      TabIndex        =   8
      Top             =   240
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2990
      _ExtentY        =   556
      Enabled         =   0   'False
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
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
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
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
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
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fpText1 
      Height          =   315
      Index           =   1
      Left            =   3240
      TabIndex        =   9
      Top             =   840
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2990
      _ExtentY        =   556
      Enabled         =   0   'False
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
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
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
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
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
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Ultima Fecha de Pedidos"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Ultima Fecha Generada"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Hasta"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Desde"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "M_Generar_Archivo_Rutas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fecha_parametros As String
Dim opsistema As String
Dim opaccesoa As String

Private Sub Command1_Click()
 
On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim Sql             As String
Dim fecha_parametro As Variant

If opsistema = "1" Then

    Sql = " sgpadm_sel_BuscaExisteFechaParametros "

ElseIf opsistema = "2" Then

    Sql = "sgpadm_Sel_Parametros 'fecrutagde'"

End If

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute(Sql)

If Not RS.EOF Then
     
   fecha_parametro = RS(0)

End If
RS.Close

Dim FechaInicial As String
Dim FechaFinal   As String
Dim fechamaxima  As String

If Format(DTPicker1, "YYYYMMDD") < fecha_parametro Then
   
   MsgBox "La fecha inicial no puede ser menor que la ultima fecha generada ", vbExclamation
   Exit Sub

End If

If DTPicker2 < DTPicker1 Then
   
   MsgBox "La fecha hasta no puede ser menor que la fecha desde", vbExclamation
   Exit Sub

End If

' Lee maxima fecha de Pedido

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = "sgpadm_sel_maximadelospedido "
Set RS = vg_db.Execute(Sql)
If Not RS.EOF Then
    
   fechamaxima = Format(RS(0), "YYYYMMDD")
  
End If
RS.Close
  
FechaInicial = Format(DTPicker1, "YYYYMMDD")
FechaFinal = Format(DTPicker2, "YYYYMMDD")
  
If FechaInicial < fechamaxima Then
   
   MsgBox "La fecha desde no puede ser menor que la Ultima Fecha de los Pedidos", vbExclamation
   Exit Sub
  
End If
  
'--> Validar rutas grupo despacho
  
  
'registrar Log sistema
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Generación_Rutas_Masivas"), CStr(Me.HelpContextID), "", "", "")
  
' Actualiza la fecha en el tabla A_Param
Sql = "sgpadm_Upd_ParametroSistema"

If opsistema = "1" Then
   
   Sql = Sql & " 'fecharuta', "
   
ElseIf opsistema = "2" Then

   Sql = Sql & " 'fecrutagde', "

End If

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = Sql & FechaFinal
Set RS = vg_db.Execute(Sql)
'RS.Close
  
' Genera Archivo de Rutas

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If opsistema = "1" Then

    Sql = "sgpadm_Ins_GenerarRutaCecoProveedor "

ElseIf opsistema = "2" Then

    Sql = "sgpadm_Ins_GenerarRutaGrupoDespachoCecoProveedor "

End If

Sql = Sql & "'" & FechaInicial & "',"
Sql = Sql & "'" & FechaFinal & "',  '" & vg_NUsr & "'"
Set RS = vg_db.Execute(Sql)
  
If Not RS.EOF Then
    
   If RS(0) > 0 Then
       
       MsgBox RS(1)
    
      'registrar Log sistema
      Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Agregado"), CStr(Me.HelpContextID), "", "", "")
    
   End If
  
End If
RS.Close
Set RS = Nothing

'registrar Log sistema
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Agregado_Generación_Rutas_Masivas"), CStr(Me.HelpContextID), "", "", "")

'
MsgBox "Se Genero las Rutas para los Cecos", vbExclamation
Unload Me

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Command2_Click()

On Error GoTo Man_Error

Unload Me

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Load()
  
On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim Fecha           As String
Dim Sql             As String
Dim fechamaxina     As Variant
Dim fecha_parametro As String

Fecha = ""

Command1.Enabled = IIf(Mid(ValidarUsuarioAcceso(CStr(opaccesoa), vg_NUsr), 2, 1) = 1, True, False)

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = "sgpadm_sel_maximadelospedido "
Set RS = vg_db.Execute(Sql)
If Not RS.EOF Then
    
   fechamaxima = Format(RS(0), "YYYYMMDD")
  
End If
RS.Close
  
If fechamaxima <> "" Then
     
   fpText1(1) = Mid(fechamaxima, 7, 2) + "/" + Mid(fechamaxima, 5, 2) + "/" + Mid(fechamaxima, 1, 4)
  
End If
  
Dim fecha_parametro_aux As Date
    
If opsistema = "1" Then
   
   Sql = " sgpadm_sel_BuscaExisteFechaParametros "

ElseIf opsistema = "2" Then

   Sql = "sgpadm_Sel_Parametros 'fecrutagde'"

End If

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute(Sql)
If Not RS.EOF Then
     
   fecha_parametro = IIf(IsNull(RS(0)), "", RS(0))
  
End If
RS.Close
    
If fecha_parametro <> "" Then
     
   fpText1(0) = Mid(fecha_parametro, 7, 2) + "/" + Mid(fecha_parametro, 5, 2) + "/" + Mid(fecha_parametro, 1, 4)
  
End If
    
If fecha_parametro <> "" Then
   
   fecha_parametro_aux = Mid(fecha_parametro, 7, 2) + "/" + Mid(fecha_parametro, 5, 2) + "/" + Mid(fecha_parametro, 1, 4)
   fecha_parametro_aux = fecha_parametro_aux + 1
   fecha_parametro = Format(fecha_parametro_aux, "YYYYMMDD")
  
End If
  
If fecha_parametro = "" Then
        
   Sql = " sgpadm_Sel_RecuperaFechaServidor "
   Set RS = vg_db.Execute(Sql)
        
   If Not RS.EOF Then
        
      DTPicker1 = RS(0)
      DTPicker2 = RS(0)
     
   End If
  
Else
       
   Fecha = Mid(fecha_parametro, 7, 2) + "/" + Mid(fecha_parametro, 5, 2) + "/" + Mid(fecha_parametro, 1, 4)
   DTPicker1 = Fecha
   DTPicker2 = Fecha
 
End If
 
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub Inicio(Glosa As String, Op As String, opacceso As String)

On Error GoTo Man_Error

opsistema = Op
opaccesoa = opacceso
Me.HelpContextID = opacceso
Me.Caption = Glosa

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Man_Error

'registrar Log sistema salir
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "", "")

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub
