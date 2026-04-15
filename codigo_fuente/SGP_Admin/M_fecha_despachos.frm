VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_fecha_despachos 
   Caption         =   "Mantención días de Despacho Proveedor"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16005
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   16005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mostrar en Grilla"
      Height          =   1815
      Left            =   9840
      TabIndex        =   11
      Top             =   5880
      Width           =   5775
      Begin VB.CheckBox Check1 
         Caption         =   "Sitios Simap"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sitios No Simap"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Sitios FM"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Value           =   1  'Checked
         Width           =   3375
      End
   End
   Begin VB.TextBox Txt_proveedor 
      Height          =   285
      Index           =   0
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   0
      Text            =   " "
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   2520
      TabIndex        =   4
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   6120
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15900
      _ExtentX        =   28046
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin EditLib.fpText fpText1 
      Height          =   315
      Index           =   1
      Left            =   10680
      TabIndex        =   1
      Top             =   880
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
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4215
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   15735
      _Version        =   393216
      _ExtentX        =   27755
      _ExtentY        =   7435
      _StockProps     =   64
      AutoClipboard   =   0   'False
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
      MaxCols         =   19
      MaxRows         =   1
      ProcessTab      =   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "M_fecha_despachos.frx":0000
   End
   Begin VB.Label Label1 
      Caption         =   "Ulitima Fecha Generada"
      Height          =   195
      Index           =   1
      Left            =   8280
      TabIndex        =   12
      Top             =   960
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   2400
      Picture         =   "M_fecha_despachos.frx":0B3E
      Top             =   840
      Width           =   480
   End
   Begin VB.Label fpayuda 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   3000
      TabIndex        =   10
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedor"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "M_fecha_despachos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BtnX As Variant
Dim estdes As Boolean
Dim parametro As Integer
Public contador As Integer
Public codigo_anterior As Integer
Dim fecha_parametro As String
Public modo As String
Public codigoceco As String
Public buff As String
Public collec As String
Public SW_VALIDACION

Private Sub Check1_Click()

On Error GoTo Man_Error

Call lee_fechas_cecos

'registrar Log sistema
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Ordenar"), Me.HelpContextID, "", "", "")

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Check2_Click()

On Error GoTo Man_Error

Call lee_fechas_cecos

'registrar Log sistema
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Ordenar"), Me.HelpContextID, "", "", "")

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Check3_Click()

On Error GoTo Man_Error

Call lee_fechas_cecos

'registrar Log sistema
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Ordenar"), Me.HelpContextID, "", "", "")

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error
    
  fg_centra Me
  Me.HelpContextID = vg_OpcM
  Toolbar1.ImageList = Partida.IL1
  
  Set BtnX = Toolbar1.Buttons.Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): BtnX.Visible = True: BtnX.ToolTipText = "Grabar ": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, False)
  Set BtnX = Toolbar1.Buttons.Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): BtnX.Visible = False: BtnX.ToolTipText = ""
  Set BtnX = Toolbar1.Buttons.Add(, "A_BajarF", , tbrDefault, "A_BajarF"): BtnX.Visible = True: BtnX.ToolTipText = "Ordenar"
 ' Set BtnX = Toolbar1.Buttons.Add(, "Proceso", , tbrDefault, "Proceso"): BtnX.Visible = True: BtnX.ToolTipText = "Generar Archivo Rutas "
  Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
  Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
 
  Toolbar1.Buttons(3).Enabled = False
  
  vaSpread1.MaxRows = 0
  
  modo = ""
  Sql = " sgpadm_sel_BuscaExisteFechaParametros "
  Set RS = vg_db.Execute(Sql)
  If Not RS.EOF Then
     
     fecha_parametro = IIf(IsNull(RS(0)), "", RS(0))
  
  End If
  RS.Close
  
  If fecha_parametro <> "" Then
     
     fpText1(1) = Mid(fecha_parametro, 7, 2) + "/" + Mid(fecha_parametro, 5, 2) + "/" + Mid(fecha_parametro, 1, 4)
  
  End If
  
Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)
  
End Sub

Private Sub lee_fechas_cecos()

 On Error GoTo Man_Error
 
 Dim Sql    As String
 Dim RS     As New ADODB.Recordset
 Dim codigo As Integer
 Dim AccMod As Boolean
 Dim i      As Long
 
 '-------> Permiso de acceso
 AccMod = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, False)
 
 Text1(2) = ""
 Text1(3) = ""
 
 Toolbar1.Buttons(3).Enabled = True
 
 parametro = parametro + 1
 
 If parametro = 4 Then
    
    parametro = 1
 
 End If
    
 vaSpread1.MaxRows = 0
    
    Dim filtro As String
    filtro = ""
    
    If Check1.Value = 1 Then
       
       filtro = filtro + "SI"
    
    Else
      
      filtro = filtro + "XX"
    
    End If
    
    If Check2.Value = 1 Then
       
       filtro = filtro + "NS"
    
    Else
      
      filtro = filtro + "XX"
    
    End If
    
    If Check3.Value = 1 Then
       
       filtro = filtro + "FM"
    
    Else
      
      filtro = filtro + "XX"
    
    End If
    
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Sql = " sgpadm_sel_parametros_despacho_proveedor_casino_V01  '" & Trim(Txt_proveedor(0)) & "'," & parametro & ",'" & filtro & "'"
    '," & parametro
    Set RS = vg_db.Execute(Sql)
    
   '-------> Inicio LLenar grilla
   
   vaSpread1.Visible = False
   vaSpread1.MaxRows = 0
   vaSpread1.MaxRows = RS.RecordCount
   i = 1
 
    Do While Not RS.EOF
              
'        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = i 'vaSpread1.MaxRows
        estdes = True
        
        vaSpread1.Col = 2 ' Org. Compras
        vaSpread1.text = RS("Id_OrgCompras")
        vaSpread1.TypeHAlign = TypeHAlignCenter
        vaSpread1.Lock = IIf(AccMod, False, True)
        
        vaSpread1.Col = 3 ' Ceco
        vaSpread1.text = RS("cli_codigo")
        vaSpread1.TypeHAlign = TypeHAlignCenter
        vaSpread1.Lock = IIf(AccMod, False, True)
        
        vaSpread1.Col = 4 ' Nombre
        vaSpread1.text = RS("cli_nombre")
        vaSpread1.Lock = IIf(AccMod, False, True)
        
        vaSpread1.Col = 5 ' Codigo Despacho
        vaSpread1.text = IIf(RS("cod_despacho") = 0, "", RS("cod_despacho"))
        vaSpread1.Lock = IIf(AccMod, False, True)
        
        vaSpread1.Col = 6 ' Descripcion Despacho
        vaSpread1.text = RS("desc_despacho")
        vaSpread1.Lock = IIf(AccMod, False, True)
        
        vaSpread1.Col = 7 ' Lunes
        vaSpread1.text = RS("lu")
        vaSpread1.Lock = IIf(AccMod, False, True)
        
        vaSpread1.Col = 8 ' Martes
        vaSpread1.text = RS("ma")
        vaSpread1.Lock = IIf(AccMod, False, True)
                
        vaSpread1.Col = 9 ' Miercoles
        vaSpread1.text = RS("mi")
        vaSpread1.Lock = IIf(AccMod, False, True)
                
        vaSpread1.Col = 10 ' Jueves
        vaSpread1.text = RS("ju")
        vaSpread1.Lock = IIf(AccMod, False, True)
                
        vaSpread1.Col = 11 ' Viernes
        vaSpread1.text = RS("vi")
        vaSpread1.Lock = IIf(AccMod, False, True)
                
        vaSpread1.Col = 12 ' Sabado
        vaSpread1.text = RS("sa")
        vaSpread1.Lock = IIf(AccMod, False, True)
                
        vaSpread1.Col = 13 ' Domingo
        vaSpread1.text = RS("do")
        vaSpread1.Lock = IIf(AccMod, False, True)
                 
        vaSpread1.Col = 14 ' Cross
        vaSpread1.text = RS("crossdocking")
        vaSpread1.Lock = IIf(AccMod, False, True)
               
        vaSpread1.Col = 15 ' Cross
        vaSpread1.text = RS("cli_blockmintrabajafinsemana")
        vaSpread1.Lock = IIf(AccMod, False, True)
           
        If RS("cli_blockmintrabajafinsemana") = 0 Then
          
          vaSpread1.Col = 12
          vaSpread1.Lock = True
          vaSpread1.Col = 13
          vaSpread1.Lock = True
        
        End If
        
        vaSpread1.Col = 16 ' Fecha desde
        vaSpread1.text = ""
        vaSpread1.Lock = True
        vaSpread1.Lock = IIf(AccMod, False, True)
        
        vaSpread1.Col = 17 ' Actualizado
        vaSpread1.text = "0"
        vaSpread1.Lock = IIf(AccMod, False, True)
           
        vaSpread1.Col = 18 ' MarcaVista
        vaSpread1.text = "0"
           
        vaSpread1.Col = 19 ' Excluye calculo CD
        vaSpread1.text = RS("ExcluyecalculoCDMinuta")
        vaSpread1.Lock = IIf(AccMod, False, True)
        
        RS.MoveNext
        i = i + 1
    
    Loop

   vaSpread1.Visible = True
   
 Toolbar1.Buttons(1).Visible = True
 Toolbar1.Buttons(2).Visible = False
    
  ret = vaSpread1.ExportToXMLBuffer("ParamCeco", collec, buff, ExportToXMLFormattedData, "")

estdes = True

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)
  
End Sub

Private Sub graba_fechas_despachos()

 On Error GoTo Man_Error
 
 Dim i                      As Integer
 Dim p                      As Integer
 
 Dim Ceco                   As String
 Dim codigo                 As String
 Dim Nombre                 As String
 Dim lunes                  As Integer
 Dim martes                 As Integer
 Dim miercoles              As Integer
 Dim jueves                 As Integer
 Dim viernes                As Integer
 Dim sabado                 As Integer
 Dim domingo                As Integer
 Dim crossking              As Integer
 Dim ExcluyecalculoCDMinuta As Integer
 Dim proveedor              As String
 proveedor = Txt_proveedor(0)
 Dim estext                 As String
 
 Dim buffact                As String
 Dim collecact              As String
 Dim modulo                 As String
 Dim buffantes              As String
 Dim buffactual             As String
 Dim existeActualizado      As Integer
 Dim fechaavalidar          As String
 
' Dim xml_ant                As New DOMDocument40
' Dim xml_act                As New DOMDocument40
 Dim xml_ant                As New DOMDocument60
 Dim xml_act                As New DOMDocument60
 
 
 Screen.MousePointer = 11
 
'--> Validar que no exista fecha rutas normal
If Not ValidarRutasxGrupo Then Exit Sub
    
      
      'solo reprcesa el ceco Modificado
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 16 'Marca Actualiza
        fechaavalidar = IIf(vaSpread1.text = "", 0, vaSpread1.text)
       
        vaSpread1.Col = 17 'Marca Actualiza
        actualiza = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
        
        If fechaavalidar <> "0" Then
          
          vaSpread1.Col = 16 'fecha
          fechamodif = Format(vaSpread1.text, "YYYYMMDD")
          SW_VALIDACION = 0
          Call validar
          vaSpread1.SetActiveCell 15, vaSpread1.ActiveRow
             
          If SW_VALIDACION = 1 Then
                
             Exit For
          
          End If
        
        End If
       
    Next i
 
If SW_VALIDACION = 0 Then

 modulo = "Mantendedor de Fechas de Despacho"
 ret = vaSpread1.ExportToXMLBuffer("ParamCeco", collecact, buffact, ExportToXMLFormattedData, "")

  buffantes = buff
  buffactual = buffact

  xml_ant.LoadXml (buffantes)
  xml_act.LoadXml (buffact)
  
existeActualizado = vaSpread1.SearchCol(16, 0, -1, "1", SearchFlagsValue)
If existeActualizado > 0 Then

'20180323comenta ya que es muy lento
'    Sql = " sgpadm_Ins_Log_Fecha_Despacho  "
'    Sql = Sql & "'" & UCase(vg_NUsr) & "',"
'    Sql = Sql & "'" & modulo & "',"
'    Sql = Sql & "'" & xml_ant.DocumentElement.XML & "',"
'    Sql = Sql & "'" & xml_act.DocumentElement.XML & "'"
'
'    'Debug.Print sql
'    Set RS = vg_db.Execute(Sql)
    
End If

 For i = 1 To vaSpread1.MaxRows
        estext = False
        
        vaSpread1.Row = i
        vaSpread1.Col = 3 'ceco
        Ceco = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        vaSpread1.Col = 5 'codigo
        codigo = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        vaSpread1.Col = 6 'Descripcion
        Nombre = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        vaSpread1.Col = 7 'Lunes
        lunes = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        vaSpread1.Col = 8 'Martes
        martes = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        vaSpread1.Col = 9 'Miercoles
        miercoles = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        vaSpread1.Col = 10 'Jueves
        jueves = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        vaSpread1.Col = 11 'Viernes
        viernes = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        vaSpread1.Col = 12 'Sabado
        sabado = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        vaSpread1.Col = 13 'Domingo
        domingo = IIf(vaSpread1.text = "", 0, vaSpread1.text)
          
        vaSpread1.Col = 14 'Domingo
        crossking = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        vaSpread1.Col = 19 'ExcluyecalculoCDMinuta
        ExcluyecalculoCDMinuta = IIf(vaSpread1.text = "", 0, vaSpread1.text)
          
        ' vaSpread1.ExportToXML(
        
        vaSpread1.Col = 17 'Marca Actualiza
        actualiza = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
                
        If codigo <> 0 And actualiza = "1" Then
           
           If lunes <> 0 Or martes <> 0 Or miercoles <> 0 Or jueves <> 0 Or viernes <> 0 Or sabado <> 0 Or domingo <> 0 Or crossking <> 0 Or ExcluyecalculoCDMinuta <> 0 Then
              
              Sql = ""
              Sql = " sgpadm_iu_Fechasdespachoproveedorcecos_V01 "
              Sql = Sql & " '" & Ceco & "',"
              Sql = Sql & " '" & Trim(proveedor) & "',"
              Sql = Sql & codigo & ","
              Sql = Sql & " '" & Nombre & "',"
              Sql = Sql & lunes & ","
              Sql = Sql & martes & ","
              Sql = Sql & miercoles & ","
              Sql = Sql & jueves & ","
              Sql = Sql & viernes & ","
              Sql = Sql & sabado & ","
              Sql = Sql & domingo & ","
              Sql = Sql & crossking & ","
              Sql = Sql & ExcluyecalculoCDMinuta
              Set RS = vg_db.Execute(Sql)
           
           Else
              
              Sql = ""
              Sql = " sgpadm_Del_ParamProveedorCecos "
              Sql = Sql & " '" & Ceco & "',"
              Sql = Sql & " '" & Trim(proveedor) & "'"
              Set RS = vg_db.Execute(Sql)
           
           End If
        
        End If
       
    Next i
        'solo reprcesa el ceco Modificado
    
    For i = 1 To vaSpread1.MaxRows
        
        estext = False
        vaSpread1.Row = i
        
        vaSpread1.Col = 16 'Marca Actualiza
        fechaavalidar = IIf(vaSpread1.text = "", 0, vaSpread1.text)
       
        vaSpread1.Col = 17 'Marca Actualiza
        actualiza = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        If fechaavalidar <> "0" Then
          
           vaSpread1.Col = 16 'fecha
           fechamodif = Format(vaSpread1.text, "YYYYMMDD")
        
           vaSpread1.Col = 3 'ceco
           Ceco = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
           XML = ""
           XML = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
           XML = XML & "<GrabaRutaCecoPro>"
           vaSpread1.Row = i
           vaSpread1.Col = 3 'Id Ruta de Compras
           codigo = IIf(vaSpread1.text = "", 0, vaSpread1.text)
           XML = XML & "<RutaCecoPro Ceco = " & Chr(34) & codigo & Chr(34)
           XML = XML & " Prov = " & Chr(34) & proveedor & Chr(34)
           XML = XML & "/>"
           XML = XML & "</GrabaRutaCecoPro>"
          
           Sql = "sgpadm_Ins_XmlGenerarRutaProveedor  "
           Sql = Sql & " '" & XML & "',"
           Sql = Sql & fechamodif & ","
           Sql = Sql & fecha_parametro & ","
           Sql = Sql & UCase(vg_NUsr)
           Set RS1 = vg_db.Execute(Sql)
           If Not RS1.EOF Then
              
              If RS1(0) > 0 Then
                 
                 MsgBox RS1(1)
              
              End If
           
           End If
           RS1.Close
           Set RS1 = Nothing
        
        End If
       
    Next i
    
    Screen.MousePointer = 0
    MsgBox "Se actualizaron las fechas de despachos de los cecos y se Generaron las Rutas Correspondiente", vbExclamation
    parametro = 0
    Call lee_fechas_cecos
 
 End If
     
 fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Unload(Cancel As Integer)

'registrar Log sistema salir
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "", "")

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error
    
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_proveedor", "prv_", "Casino", "ProveedorSimap"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    Txt_proveedor(0).text = vg_codigo
    fpayuda(0).Caption = vg_nombre
    estdes = True

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(Text1(Index).text, ",")

If Index = 2 Then
   
   Text1(3).text = ""
   Text1(4).text = ""

ElseIf Index = 3 Then
   
   Text1(2).text = ""
   Text1(4).text = ""

ElseIf Index = 4 Then
   
   Text1(2).text = ""
   Text1(3).text = ""

End If

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 18
    vaSpread1.text = 0

Next

Select Case Index

Case 2, 3, 4
    
    vaSpread1.Visible = False
    
    If Trim(Text1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread1.Col = 2
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 18
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 18
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 18
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 18
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 18
                 vaSpread1.text = 0
                 
              End If
           
           End If
        
        Next i
        
        vaSpread1.SetActiveCell Index + 1, 1
        vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread1.ColUserSortIndicator(IIf(Trim(wvarArr8020(X)) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread1.SortKey(1) = IIf(Trim(wvarArr8020(X)) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(Text1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 18
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

'Private Sub Text1_Change(Index As Integer)
'
'Select Case Index
'
'Case 2, 3
'    vaSpread1.Visible = False
'    If Trim(Text1(Index).text) <> "" Then
'       For i = 1 To vaSpread1.MaxRows
'           vaSpread1.Row = i
'           vaSpread1.Col = Index
'           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(Text1(Index).text) & "*"
'           vaSpread1.Col = 2
'           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
'              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
'           Else
'              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
'           End If
'        Next i
'        vaSpread1.SetActiveCell Index, 1
'    End If
''    vaSpread1_Click Index, 0
'    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
'    vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
'    vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
'    vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
'    If Trim(Text1(Index).text) = "" Then
'       For i = 1 To vaSpread1.MaxRows
'           vaSpread1.Row = i
'           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
'       Next
'       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
'       vaSpread1.SetActiveCell Index, 1
'    End If
'    vaSpread1.Visible = True
'
'End Select
'
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim seleccion As String
Dim i As Integer
    
If (vaSpread1.MaxRows < 1 Or Trim(Txt_proveedor(0).text) = "") And Button.Index <> 5 Then

  MsgBox "no existen datos seleccionado en la lista", vbCritical, MsgTitulo
  Exit Sub

End If
    
    Select Case Button.Index
        
    Case 1 ' Grabo Fecha de Despacho
        
     
      'registrar Log sistema modificar
       Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), CStr(Me.HelpContextID), "", "", "")
       
       Call graba_fechas_despachos
        
       'registrar Log sistema actualizar lista
       Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")
        
    Case 3 ' Grabo Fecha de Despacho
        
       Call lee_fechas_cecos
       
       'registrar Log sistema
       Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Ordenar"), Me.HelpContextID, "", "", "")
    
    Case 4 ' Generar ARchivo Rutas
             
        M_Generar_Archivo_Rutas.Show 0, Partida
      
    Case 5 ' Salir del Programa
        
        Me.Hide
        Unload Me
    
    End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Txt_proveedor_Change(Index As Integer)

On Error GoTo Man_Error
    
    vaSpread1.MaxRows = 0
      
    Dim RS1 As New ADODB.Recordset
    Dim Sql As String
    If Txt_proveedor(0).text = "" Then fpayuda(0).Caption = "": Exit Sub
    Sql = Trim(LimpiaDato(Txt_proveedor(0).text))
    Set RS1 = vg_db.Execute("sgpadm_s_proveedor 4, '" & Sql & "', ''")
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS1!prv_nombre)
    parametro = 0
    
    Call lee_fechas_cecos

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1
    
    Dim i As Long
    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    
    Next

End Select
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

If (Col <> 7 And Col <> 8 And Col <> 9 And Col <> 10 And Col <> 11 And Col <> 12 And Col <> 13 And Col <> 19) Or Row = 0 Or estdes Then
   
   Exit Sub

End If

vaSpread1.Row = Row
vaSpread1.Col = 5

If vaSpread1.text <> 1 Then
   
   For i = 7 To 13
       
       If i <> Col Then
          
          estdes = True
          vaSpread1.Col = i
          vaSpread1.text = "0"
          estdes = False
       
       End If
       
   Next i

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error
  
Dim codigo As Long

Dim findessemana As String

If vaSpread1.MaxRows < 1 Then Exit Sub

If (Col) = 5 Then
            
            modo = "M"
            vaSpread1.SetFocus
            vaSpread1.Row = vaSpread1.ActiveRow
            
            vaSpread1.Col = 5
            codigo = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
              
            vaSpread1.Col = 14
            findessemana = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
               
            vaSpread1.Col = 3
            codigoceco = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
            
            ' Rescata el Parametro correspondiente
            
            If codigo = 1 Then
              
              vaSpread1.Col = 6 ' Periodo
              vaSpread1.text = "SEMANAL"
            
            End If
            
            If codigo = 2 Then
              
              vaSpread1.Col = 6 ' Periodo
              vaSpread1.text = "QUINCENAL"
            
            End If
            
            If codigo > 2 Then
              
              vaSpread1.Col = 6 ' Distintos Periodo
              vaSpread1.text = "CADA " + CStr(codigo) + " SEMANAS"
            
            End If
            
            If codigo = 0 Then
                
                If codigo_anterior = 1 Then
                    
                    vaSpread1.Col = 6 ' Periodo
                    vaSpread1.text = "SEMANAL"
                
                End If
                
                If codigo_anterior = 2 Then
                    
                    vaSpread1.Col = 6 ' Periodo
                    vaSpread1.text = "QUINCENAL"
                
                End If
                
                If codigo_anterior > 2 Then
                    
                    vaSpread1.Col = 6 ' Periodo
                    vaSpread1.text = "CADA " + CStr(codigo_anterior) + " SEMANAS"
                
                End If
                
                vaSpread1.Col = 5 ' Periodo
                vaSpread1.text = codigo_anterior
                 
            End If
       
            estdes = True
       
     If codigo > 0 Then
      
      If codigo <> codigo_anterior Then
            
            With vaSpread1
             
              .Col = 7:  .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter:  .Lock = False
              .Col = 8:  .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter:  .Lock = False
              .Col = 9:  .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter:  .Lock = False
              .Col = 10:  .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter:  .Lock = False
              .Col = 11: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter:  .Lock = False
              .Col = 12: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter:  .Lock = False
              .Col = 13: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter:  .Lock = False
              
              ' Rescata la Ultima Fecha del Cecos
               
               Dim valor As String
           
               If fecha_parametro <> "" Then
                  
                  Sql = " sgpadm_sel_maximafechadesdedelcecos " & codigoceco
                  Set RS = vg_db.Execute(Sql)
                  
                  If Not RS.EOF Then
                    
                    If IsNull(RS(0)) Then
                           
                           vaSpread1.Col = 16 ' Fecha Hastas
                           vaSpread1.Lock = False
                           vaSpread1.CellType = CellTypeDate
                           vaSpread1.text = Mid(fecha_parametro, 7, 2) + "/" + Mid(fecha_parametro, 5, 2) + "/" + Mid(fecha_parametro, 1, 4)
                           vaSpread1.SetFocus
                           vaSpread1.Col = 17
                           vaSpread1.text = "1"
                        
                        Else
                        
                           vaSpread1.Col = 16 ' Fecha Hastas
                           vaSpread1.Lock = False
                           vaSpread1.CellType = CellTypeDate
                           vaSpread1.text = Format(RS(0), "DD/MM/YyyY")
                           vaSpread1.SetFocus
                           vaSpread1.Col = 17
                           vaSpread1.text = "1"
                         
                         End If
                  
                  End If
                  
                  RS.Close
               
               End If
             
             End With
          
        End If
       
       End If
      
       estdes = False
     
     If findessemana = "0" Then
     
         With vaSpread1
              
              .Row = vaSpread1.ActiveRow
              .Col = 12: .Lock = True
              .Col = 13: .Lock = True
           
         End With
     
     End If
     
End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.ActiveCol

If Row = 0 Or vaSpread1.MaxRows < 1 Then Exit Sub

If vaSpread1.Lock = False Then

   If Col = 16 Then
      
      modo = "M"
      
   End If
      
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = 5
   codigo_anterior = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)

   If Col = 14 Or Col = 19 Then
      
      Toolbar1.Buttons(1).Visible = True
      Toolbar1.Buttons(2).Visible = False
  
   End If
  
   If (Col = 7 Or Col = 8 Or Col = 9 Or Col = 10 Or Col = 11 Or Col = 12 Or Col = 13) Then
      
      modo = "M"
      
      Toolbar1.Buttons(1).Visible = True
      Toolbar1.Buttons(2).Visible = False
          
      vaSpread1.Col = 14
      findessemana = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
         
      vaSpread1.Col = 3
      codigoceco = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
            
      If fecha_parametro <> "" Then
       
       Sql = " sgpadm_sel_maximafechadesdedelcecos " & codigoceco
       
       Set RS = vg_db.Execute(Sql)
       If Not RS.EOF Then
            
            If IsNull(RS(0)) Then
               
               vaSpread1.Col = 16 ' Fecha Hastas
               vaSpread1.Lock = False
               vaSpread1.CellType = CellTypeDate
               vaSpread1.text = Mid(fecha_parametro, 7, 2) + "/" + Mid(fecha_parametro, 5, 2) + "/" + Mid(fecha_parametro, 1, 4)
               vaSpread1.SetFocus
               vaSpread1.Col = 17
               vaSpread1.text = "1"
               
            Else
            
               vaSpread1.Col = 16 ' Fecha Hastas
               vaSpread1.Lock = False
               vaSpread1.CellType = CellTypeDate
               vaSpread1.text = Format(RS(0), "DD/MM/YyyY")
               vaSpread1.SetFocus
               vaSpread1.Col = 17
               vaSpread1.text = "1"
             
             End If
            
          End If
          RS.Close
      
      Else
          
          modo = ""
      
      End If
      
      If codigo_anterior <> 1 Then
          
          estdes = False
      
      End If
  
  ElseIf Col = 14 Or Col = 19 Then
      
      vaSpread1.Col = 17
      vaSpread1.text = "1"
  
  End If

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(1).Visible = True Then
   
   Call validar

End If

End Sub

Private Sub validar()

Dim Fecha As Date
Dim fechaproceso As String
Dim fechahasta As String

vaSpread1.Col = 16 'Lunes
Fecha = IIf(vaSpread1.text = "", 0, vaSpread1.text)
fechaproceso = Format(Fecha, "YYYYMMDD")
fechahasta = fecha_parametro

If fechahasta <> "" Then
 
' If fechaproceso > fechahasta Then
'      vaSpread1.Col = 16
'      vaSpread1.text = "0"
'      SW_VALIDACION = 1
'      MsgBox "La fecha desde no puede ser mayor a la ultima fecha generada " + fpText1(1).text, vbExclamation
'      vaSpread1.TypeHAlign = TypeHAlignCenter
'      vaSpread1.SetActiveCell 14, vaSpread1.ActiveRow
'    '  Toolbar1.Buttons(1).Enabled = False
'
'      Exit Sub
'
' End If
'
' sql = " sgpadm_sel_ExistenPedidoentrerangodeFecha " & "'" & codigoceco & "'," & "'" & fechaproceso & "'"
' Set RS = vg_db.Execute(sql)
' If Not RS.EOF Then
'    If RS(0) > 0 Then
'      vaSpread1.Col = 16
'      vaSpread1.text = "0"
'      SW_VALIDACION = 1
'      MsgBox "Existen pedidos confirmados en esta fecha", vbExclamation
'      vaSpread1.SetActiveCell 14, vaSpread1.ActiveRow
'     ' Toolbar1.Buttons(1).Enabled = False
'      Exit Sub
'
'  End If
' End If
'
' sql = " sgpadm_sel_maximafechadesdedelcecospedido " & "'" & codigoceco & "'"
' Set RS = vg_db.Execute(sql)
' If Not RS.EOF Then
'  If Not IsNull(RS(0)) Then
'   fechahasta = IIf(IsNull(RS(0)), Format(Now, "YYYYMMDD"), Format(RS(0), "YYYYMMDD"))
'
'  If fechaproceso < fechahasta Then
'      vaSpread1.Col = 16
'      vaSpread1.text = "0"
'      SW_VALIDACION = 1
'      MsgBox "La Fecha desde no puede ser menor que la fecha del ultimo pedido confirmado ", vbExclamation
'      vaSpread1.SetActiveCell 14, vaSpread1.ActiveRow
'     ' Toolbar1.Buttons(1).Enabled = False
'   Else
'      modo = ""
'      vaSpread1.Col = 16 ' Periodo
'      vaSpread1.text = "1"
'      SW_VALIDACION = 0
'      vaSpread1.TypeHAlign = TypeHAlignCenter
'     ' Toolbar1.Buttons(1).Enabled = True
'    End If
'    Else
'     modo = ""
'      vaSpread1.Col = 16 ' Periodo
'      vaSpread1.text = "1"
'      SW_VALIDACION = 0
'      vaSpread1.TypeHAlign = TypeHAlignCenter
'     ' Toolbar1.Buttons(1).Enabled = True
'  End If
' End If

End If

End Sub

Function ValidarRutasxGrupo() As Boolean

Dim RS            As New ADODB.Recordset
Dim Sql           As String
Dim EstDia        As Boolean
Dim fechaavalidar As String
Dim fechamodif    As Long
Dim actualiza     As Long
Dim xmlfamilia    As String
Dim i             As Long
Dim codigoceco    As String

Dim xlApp As Object
Dim xlWb As Object
Dim xlWs As Object

ValidarRutasxGrupo = True

EstDia = False
fechaavalidar = ""
actualiza = 0
xmlfamilia = ""
xmlfamilia = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlfamilia = xmlfamilia & "<GrabaRutaCeco>"
    
For i = 1 To vaSpread1.MaxRows
    
    vaSpread1.Row = i
    vaSpread1.Col = 16 'Marca Actualiza
    fechaavalidar = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
    vaSpread1.Col = 17 'Marca Actualiza
    actualiza = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
    If fechaavalidar <> "0" And actualiza = 1 Then
        
       vaSpread1.Col = 3  'ceco
       codigoceco = vaSpread1.text
        
       vaSpread1.Col = 16 'fecha
       fechamodif = Format(vaSpread1.text, "YYYYMMDD")
          
       xmlfamilia = xmlfamilia & " <RutaCeco"
       xmlfamilia = xmlfamilia & " Ceco = " & Chr(34) & codigoceco & Chr(34)
       xmlfamilia = xmlfamilia & " Fec = " & Chr(34) & fechamodif & Chr(34)
       xmlfamilia = xmlfamilia & "/>"
    
    End If
        
Next i
    
xmlfamilia = xmlfamilia & "</GrabaRutaCeco>"
          
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
          
Sql = ""
Sql = "sgpadm_Sel_XmlValidasiExisteRutaGrupoPAP "
Sql = Sql & " '" & xmlfamilia & "'"
         
Set RS = vg_db.Execute(Sql)
If Not RS.EOF Then
            
   If RS(0) > 0 And RS("estado") <> "0" Then
               
      fg_descarga
      ValidarRutasxGrupo = False
      ' Create an instance of Excel and add a workbook
      Set xlApp = CreateObject("Excel.Application")
      Set xlWb = xlApp.Workbooks.Add
      Set xlWs = xlWb.Worksheets("Hoja1")
          
      If RS.RecordCount > xlWs.Range("A1", xlWs.Range("A1").End(xlDown)).Rows.count Then
    
         ' Close ADO objects
         RS.Close
         Set RS = Nothing
       
         ' Release Excel references
         Set xlWs = Nothing
         Set xlWb = Nothing

         Set xlApp = Nothing
       
         MsgBox "Excede numero filas, debera bajar la fecha despacho", vbCritical
       
         Exit Function
    
      End If
          
      MsgBox "Existen rutas normales en la fechas solicitadas por pantalla. se generara una planilla excel con las rutas que impiden generar las nuevas rutas grupo despacho. Pasos seguir es borrar las rutas normales", vbCritical + vbOKOnly, MsgTitulo
          
      ' Display Excel and give user control of Excel's lifetime
      xlApp.Visible = True
      xlApp.UserControl = True
    
      ' Check version of Excel
      Call encabezado(RS, xlWs)
          
      xlWs.Cells(2, 1).CopyFromRecordset RS
    
      ' Auto-fit the column widths and row heights
      xlApp.Selection.CurrentRegion.Columns.AutoFit
      xlApp.Selection.CurrentRegion.Rows.AutoFit

      ' Release Excel references
      Set xlWs = Nothing
      Set xlWb = Nothing

      Set xlApp = Nothing
          
      Exit Function
            
   End If
         
End If
          
' Close ADO objects
RS.Close
Set RS = Nothing
    
Exit Function
Man_Error:

ValidarRutasxGrupo = False
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Function


