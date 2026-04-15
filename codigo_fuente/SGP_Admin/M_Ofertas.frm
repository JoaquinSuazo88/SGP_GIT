VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_Ofertas 
   Caption         =   "Mantencion de Ofertas"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Ofertas"
      Height          =   4335
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   7095
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3615
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   6735
         _Version        =   393216
         _ExtentX        =   11880
         _ExtentY        =   6376
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   4
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
         MaxCols         =   4
         MaxRows         =   1
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "M_Ofertas.frx":0000
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_Ofertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public modo As String
Private Sub Form_Load()
On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
fg_centra Me
modo = "": ibusca = 0
Toolbar1.ImageList = Partida.IL1

  Set BtnX = Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
  Set BtnX = Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
  Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
  
  Set BtnX = Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = "Cancelar "
  Set BtnX = Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = ""
  
 
  Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
  Set BtnX = Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = ""
  Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
  Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
     
  Toolbar1.Buttons(4).Enabled = False
  Toolbar1.Buttons(6).Enabled = False
  
  Call lee_ofertas

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)
  
End Sub
Private Sub lee_ofertas()

 On Error GoTo Man_Error
    
    Sql = " sgpadm_Sel_Ofertas "
    Set RS = vg_db.Execute(Sql)
    
   '-------> Inicio LLenar grilla
   
    vaSpread1.MaxRows = 0
 
    Do While Not RS.EOF
        
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 1 ' Codigo
        vaSpread1.text = Val(RS(0))
        vaSpread1.TypeHAlign = TypeHAlignCenter
        
        vaSpread1.Col = 2 ' Descripcion
        vaSpread1.text = IIf(IsNull(RS(1)), " ", RS(1))
        
        vaSpread1.Col = 3 ' Descripcion Corta
        vaSpread1.text = IIf(IsNull(RS(3)), " ", RS(2))
        
        vaSpread1.Col = 4 ' Activo
        vaSpread1.text = IIf(IsNull(RS(2)), " ", RS(3))
        
        
        
   
        RS.MoveNext
    Loop
    
 RS.Close: Set RS = Nothing

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
    Select Case Button.Index
    Case 1 ' Genera Pedcido Nuevo por Pedido o Proyectado
         modo = "A"
         Toolbar1.Buttons(1).Enabled = False
         Toolbar1.Buttons(4).Enabled = True
         Toolbar1.Buttons(6).Enabled = True
         vaSpread1.MaxRows = vaSpread1.MaxRows + 1
         vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus
        vaSpread1.Col = 4 ' Activo
        vaSpread1.text = 1
    Case 4 'Confirmar
       If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
         Toolbar1.Buttons(1).Enabled = True
         Toolbar1.Buttons(4).Enabled = False
         Toolbar1.Buttons(6).Enabled = False
         Call lee_ofertas
         
        
        
    
         
    Case 6 'Confirmar
         Call grabar
    
    Case 9 'Salir
        Me.Hide
        Unload Me
    End Select
        
 
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Private Sub validar()
On Error GoTo Man_Error

Dim codigo As Integer
Dim descripcion As String
Dim activo As Integer
Dim descripcioncorta As String
Dim RS1 As New ADODB.Recordset

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
vaSpread1.Col = 2: descripcion = Trim(LimpiaDato(vaSpread1.Value))
vaSpread1.Col = 4: activo = Val(vaSpread1.Value)
vaSpread1.Col = 3: descripcioncorta = vaSpread1.text
 
    If Trim(descripcion) = "" Then
        MsgBox "Favor ingresar nombre oferta, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, Msgtitulo
        vaSpread1.Col = 2: vaSpread1.SetActiveCell vaSpread1.ActiveCol, vaSpread1.ActiveRow: vaSpread1.SetFocus
        OpGr = False
        Exit Sub
    End If
    
      
    If Trim(descripcioncorta) = "" Then
        MsgBox "Favor ingresar nombre corto, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, Msgtitulo
        vaSpread1.Col = 3: vaSpread1.SetActiveCell vaSpread1.ActiveCol, vaSpread1.ActiveRow: vaSpread1.SetFocus
        OpGr = False
        Exit Sub
    End If
 Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
   
End Sub

Private Sub grabar()

On Error GoTo Man_Error

Dim codigo As Integer
Dim descripcion As String
Dim activo As Integer
Dim descripcioncorta As String
Dim RS1 As New ADODB.Recordset


If modo = "" Then
  modo = "M"

End If
vaSpread1.Row = vaSpread1.ActiveRow

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
    vaSpread1.Col = 2: descripcion = Trim(LimpiaDato(vaSpread1.Value))
    vaSpread1.Col = 4: activo = Val(vaSpread1.Value)
    vaSpread1.Col = 3: descripcioncorta = vaSpread1.text
 
    If Trim(descripcion) = "" Then
        MsgBox "Favor ingresar nombre ofreta, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, Msgtitulo
        vaSpread1.Col = 2: vaSpread1.SetActiveCell vaSpread1.ActiveCol, vaSpread1.ActiveRow: vaSpread1.SetFocus
        OpGr = False
        Exit Sub
    End If
    
      
    If Trim(descripcioncorta) = "" Then
        MsgBox "Favor ingresar nombre corto, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, Msgtitulo
        vaSpread1.Col = 3: vaSpread1.SetActiveCell vaSpread1.ActiveCol, vaSpread1.ActiveRow: vaSpread1.SetFocus
        OpGr = False
        Exit Sub
    End If
    
    If codigo <> 0 Then
        modo = "M"
    Else
        modo = "A"
    
    End If
    
          
          Sql = " sgpadm_iu_ofertas "
          Sql = Sql & "'" & modo & "',"
          Sql = Sql & codigo & ","
          Sql = Sql & "'" & Trim(descripcion) & "',"
          Sql = Sql & activo & ","
          Sql = Sql & "'" & descripcioncorta & "',"
          Sql = Sql & "'" & UCase(vg_NUsr) & "'"
          Set RS1 = vg_db.Execute(Sql)
          
 Next i
          
         
         Toolbar1.Buttons(1).Enabled = True
         Toolbar1.Buttons(4).Enabled = False
         Toolbar1.Buttons(6).Enabled = False

Call lee_ofertas
modo = ""
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub
Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
  Toolbar1.Buttons(4).Enabled = True
  Toolbar1.Buttons(6).Enabled = True
  Toolbar1.Buttons(1).Enabled = False
  vaSpread1.Row = vaSpread1.ActiveRow
  vaSpread1.Col = 4 ' Activo
  vaSpread1.Lock = False
       
End Sub
Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If modo = "" Then
   modo = "M"
   
End If

If NewCol = 4 Then
   Toolbar1.Buttons(4).Enabled = True
   Toolbar1.Buttons(6).Enabled = True
   Toolbar1.Buttons(1).Enabled = False


End If



If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(3).Visible = True Then
    Call validar
'Else
   
   

End If

End Sub



