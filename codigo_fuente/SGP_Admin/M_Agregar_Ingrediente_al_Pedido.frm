VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_Agregar_Ingrediente_al_Pedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Ingrediente al Pedido"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   15720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin EditLib.fpText txt_Codigo 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1095
      _Version        =   196608
      _ExtentX        =   1931
      _ExtentY        =   661
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
      Height          =   2355
      Left            =   390
      TabIndex        =   1
      Top             =   1530
      Width           =   15045
      _Version        =   393216
      _ExtentX        =   26538
      _ExtentY        =   4154
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
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
      MaxCols         =   12
      MaxRows         =   1
      SpreadDesigner  =   "M_Agregar_Ingrediente_al_Pedido.frx":0000
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Label lblval1 
      AutoSize        =   -1  'True
      Caption         =   "Solo se pueden agregar ingredientes al pedido que tengan convenio"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   375
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   4845
   End
   Begin VB.Label fpayuda 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   1965
      TabIndex        =   3
      Top             =   1020
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   1485
      Picture         =   "M_Agregar_Ingrediente_al_Pedido.frx":06C3
      Top             =   930
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrediente"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "M_Agregar_Ingrediente_al_Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql   As String
Dim BtnX  As Variant
Dim RS    As New ADODB.Recordset

Private Sub Form_Load()

Toolbar1.ImageList = Partida.IL1
  
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Agregar Ingrediente "
Set BtnX = Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error
    
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""

    vg_Indppr = "1"
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "IngReal"
    B_TabEst.Show 1
    Me.Refresh

    If vg_codigo = "" Then Exit Sub
    txt_Codigo.text = vg_codigo: fpayuda(0).Caption = vg_nombre
    Call txt_Codigo_KeyPress(13)


Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub busca_detalle()

On Error GoTo Man_Error

  Dim RS_ing As New ADODB.Recordset
  Dim tipo As Integer
  Dim fechini As String
  Dim fechfin As String
  Dim msg As String

  If M_Lista_Pedido.fpText3 = "Pedido PAP" Then tipo = 1
  If M_Lista_Pedido.fpText3 = "Pedido CD" Then tipo = 3
  If M_Lista_Pedido.fpText3 = "Proyectado" Then tipo = 2

  fechini = Format(M_Lista_Pedido.fpDateTime1(2), "YYYYMMDD")
  fechfin = Format(M_Lista_Pedido.fpDateTime1(3), "YYYYMMDD")
                
  Sql = "sgpadm_sel_Ingredientes_Especiales "
  Sql = Sql & "'" & vg_cencos & "',"
  Sql = Sql & fechini & ","
  Sql = Sql & fechfin & ","
  Sql = Sql & "'" & txt_Codigo.text & "',"
  
  Sql = Sql & tipo
  
  Set RS_ing = vg_db.Execute(Sql)
                
  '-------> Inicio LLenar grilla
                
  vaSpread1.MaxRows = 0
  If Not RS_ing.EOF Then
    msg = ""
    
    If RS_ing("CodigoSap") = "" Then msg = msg & "- No se encontro un formato de compra SAP para este ingrediente." & vbCrLf
    If RS_ing("Proveedor") = "" Then msg = msg & "- No se encontro un convenio para este ingrediente." & vbCrLf
    If IsNull(RS_ing("fecha_despacho")) Then msg = msg & "- No se encontro una ruta para este ingrediente." & vbCrLf
    If IsNull(RS_ing("ProductoSgp")) Then msg = msg & "- No se encontro un Producto SGP para este ingrediente." & vbCrLf
                 
    If msg <> "" Then
      
      MsgBox msg, vbCritical + vbExclamation
      GoTo fin
    
    End If
              
    Do While Not RS_ing.EOF
             
      'If Not IsNull(RS_ing(0)) Then
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 1 ' Proveedor
        vaSpread1.text = RS_ing("Proveedor")
                 
        
        vaSpread1.Col = 2 ' Nombre Proveedor
        vaSpread1.text = RS_ing("NombreProveedor")
        
        vaSpread1.Col = 3 ' Sap
        vaSpread1.text = RS_ing("CodigoSap")
         
        vaSpread1.Col = 4 ' Descripcion
        vaSpread1.text = RS_ing("DesMaterial")
        
        vaSpread1.Col = 5 ' Undidsd
        vaSpread1.text = RS_ing("Unidad")
        
        vaSpread1.Col = 6 ' fecha despacho
        vaSpread1.text = IIf(IsNull(RS_ing("fecha_despacho")), "", RS_ing("fecha_despacho"))
        vaSpread1.TypeHAlign = TypeHAlignCenter
       
        
        vaSpread1.Col = 7 ' Cantidad Despachada
        If RS_ing("Proveedor") = "" Then
          
          vaSpread1.Lock = True
          vaSpread1.ForeColor = vbBlue
          Me.lblval1.Visible = True
        
        End If
        vaSpread1.text = 0
        
        vaSpread1.Col = 8 ' Perfil Reondeo
        vaSpread1.text = RS_ing("Perfil")
        vaSpread1.TypeHAlign = TypeHAlignCenter
        
        vaSpread1.Col = 9 ' Familia
        vaSpread1.text = IIf(IsNull(RS_ing("Familia")), "", RS_ing("Familia"))
        
        vaSpread1.Col = 10 ' IdRuta
         vaSpread1.text = IIf(IsNull(RS_ing("Ruta")), "", RS_ing("Ruta"))
        
        vaSpread1.Col = 11 ' SGP
        vaSpread1.text = RS_ing("ProductoSgp") + " " + RS_ing("Pro_nombre")
        
        vaSpread1.Col = 12 ' Perfil
        vaSpread1.text = RS_ing("Ingrediente")
       'End If
             
        Toolbar1.Buttons(1).Visible = True
        Toolbar1.Buttons(2).Visible = False
             
        RS_ing.MoveNext
     
     Loop
   
   End If

fin:
                
    RS_ing.Close
    Set RS_ing = Nothing
                
Exit Sub
Man_Error:
    If RS_ing.State = 1 Then
      RS_ing.Close: Set RS_ing = Nothing
    End If
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error
    
    Select Case Button.Index
    
    Case 1 ' Agrega Ingrediente
        
        Call Agrega_Ingrediente
        
    Case 4 ' Salir del Programa
        
        Me.Hide
        Unload Me
    
    End Select
    
Exit Sub
Man_Error:
    fg_descarga
    If Err = 438 Or Err = 70 Then
       MsgBox "Hay un Excel Abierto debe cerrarlo para poder bajar el nuevo", vbExclamation
       Exit Sub
    End If
    
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Agrega_Ingrediente()

 On Error GoTo Man_Error
 
 Dim proveedor As String
 Dim codMateria As String
 Dim unidad As String
 Dim Fecha As String
 Dim cantdad As Double
 Dim perfil As Double
 Dim familia As String
 Dim desmateria As String
 Dim IdRuta As Long
 Dim cod_prod As String
 'Dim perfil As Long
 Dim contador As Integer
 contador = 0
 
 Dim pedido As Integer
 pedido = M_Lista_Pedido.fpText2
    
  
    
For i = 1 To vaSpread1.MaxRows
    
    vaSpread1.Row = i
    vaSpread1.Col = 7
    cantidad = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
    vaSpread1.Col = 1
    proveedor = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
 
 If cantidad > 0 Then
    
    If proveedor <> "-1" Then
     
     contador = contador + 1
     vaSpread1.Col = 3
     codMateria = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
     vaSpread1.Col = 4
     desmateria = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
     vaSpread1.Col = 5
     unidad = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
     vaSpread1.Col = 6
     Fecha = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
     vaSpread1.Col = 7
     cantdad = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
     vaSpread1.Col = 8
     perfil = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
     vaSpread1.Col = 9
     familia = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
     vaSpread1.Col = 10
     IdRuta = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
     vaSpread1.Col = 11
     cod_prod = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
   ' vaSpread1.Col = 12
   ' perfil = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
 
 
     Sql = "sgpadm_iu_Ingredientes_Especiales  "
     Sql = Sql & pedido & ","
     Sql = Sql & "'" & vg_cencos & "',"
     Sql = Sql & "'" & txt_Codigo & "',"
     Sql = Sql & "'" & cod_prod & "',"
     Sql = Sql & "'" & codMateria & "',"
     Sql = Sql & "'" & Format(Fecha, "YYYYMMDD") & "',"
     Sql = Sql & cantdad & ","
     Sql = Sql & IdRuta & ","
     Sql = Sql & "'" & proveedor & "',"
     Sql = Sql & "'" & familia & "',"
     Sql = Sql & "'" & desmateria & "',"
     Sql = Sql & "'" & unidad & "',"
     Sql = Sql & "'" & perfil & "'"
     
     Set RS = vg_db.Execute(Sql)
    
    End If
 'Else
  
  
 End If

Next i
    
    
    If contador > 0 Then
     
     MsgBox "Se Actualizo el Detalle del Pedido OK", vbExclamation
     
    Else
     
     MsgBox "No existen Registros para agregar al Pedido", vbExclamation
    
    End If

Exit Sub
Man_Error:
    fg_descarga
    
    If Err = 438 Or Err = 70 Then
       
       MsgBox "Hay un Excel Abierto debe cerrarlo para poder bajar el nuevo", vbExclamation
       Exit Sub
    End If
    
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub txt_Codigo_Change()

fpayuda(0) = ""
vaSpread1.MaxRows = 0
Me.lblval1.Visible = False

End Sub

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
  
  On Error GoTo Man_Error
  
  Dim Sql As String

  If KeyAscii = 13 Then
  Me.lblval1.Visible = False
    fpayuda(0).Caption = ""
  
  
    'txt_Codigo.text = Trim(LimpiaDato(txt_Codigo.text))
  
    Sql = "SELECT DISTINCT a.ing_codigo, a.ing_nombre, a.ing_indppr FROM b_ingrediente a where a.ing_codigo = '" & Trim(LimpiaDato(txt_Codigo.text)) & "'"
  
      vaSpread1.MaxRows = 0
      RS.Open Sql, vg_db, adOpenStatic
      If RS.EOF Then
        
        MsgBox "Ingrediente no encontrado", vbExclamation
        RS.Close: Set RS = Nothing
        Exit Sub
      
      Else
      If RS(2) = 2 Then
        
        MsgBox "Este ingrediente corresponde a propuesta", vbExclamation
        RS.Close: Set RS = Nothing
        Exit Sub
      
      End If
        
        fpayuda(0).Caption = Trim(RS!ing_nombre)
      
      End If
      RS.Close
      Set RS = Nothing
          
     'validacion: el ingrediente no puede estar en el pedido
     
      Dim pedido As Integer
      pedido = M_Lista_Pedido.fpText2
      Sql = "sgpadm_sel_Ingredientes_Pedido " & pedido & "," & "'" & txt_Codigo & "'"
      Set RS = vg_db.Execute(Sql)
      If Not RS.EOF Then
        
        If RS(0) > 0 Then
           
           MsgBox "El Ingrediente ya esta incorporado en el Pedido", vbExclamation
        
        Else
          
          Call busca_detalle
        
        End If
      
      End If
      RS.Close
      Set RS = Nothing
  End If
    
Exit Sub
Man_Error:
      If RS.State = 1 Then
        RS.Close: Set RS = Nothing
      End If
      fg_descarga
      MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
      ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub

If Mode = 0 And ChangeMade = True Then
    Dim proveedor As String
    Dim perfil_redondeo As Double
    Dim cantidad As Double
    
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    proveedor = Trim(vaSpread1.text)
    
   vaSpread1.Col = 8
   perfil_redondeo = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
   vaSpread1.Col = 7
   cantidad = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
 
 If cantidad - (perfil_redondeo * Fix(cantidad / perfil_redondeo)) <> 0 Then
     
     MsgBox "La no es multiplo de perfil de redondeo", vbCritical + vbExclamation
     vaSpread1.Col = 7 ' Familia
     vaSpread1.text = 0
     
     Exit Sub
     
 End If
 
End If

Man_Error:
    fg_descarga
    If Err.Number = 11 Then
      MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
     ins_log_error Date & Time & Err & ":  " & Error$(Err)
    End If

End Sub
