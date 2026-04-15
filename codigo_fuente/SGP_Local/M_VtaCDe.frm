VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_VtaCDe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Venta Contado x Centro Costo"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9015
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4830
         Left            =   180
         TabIndex        =   1
         Top             =   600
         Width           =   8670
         _Version        =   393216
         _ExtentX        =   15293
         _ExtentY        =   8520
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "M_VtaCDe.frx":0000
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "fff"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   5280
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_VtaCDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String
Dim OpGr As Boolean
Dim ibusca As Long, codcli As String

Private Sub GrabaRegistro(Fila As Long)
Dim codigo As Long, descripcion As String, detmto As Double
On Error GoTo Man_Error
OpGr = True
vaSpread1.Row = Fila
vaSpread1.Col = 1: codigo = vaSpread1.text
vaSpread1.Col = 3: descripcion = Trim(LimpiaDato(vaSpread1.text))
vaSpread1.Col = 4: detmto = vaSpread1.text
If detmto < 1 Then MsgBox OpGr = False: Exit Sub
vg_db.BeginTrans

vg_db.CommitTrans
modo = "": Gl_Ac_Botones Me, 1, 1, modo
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = True
OpGr = False

Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 6630
Me.Width = 9360
Msgtitulo = "Detalle Venta Contado x Centro Costo"
fg_centra Me
modo = "": ibusca = 0
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = True
MoverDatosGrillas
OpGr = False
End Sub

Private Sub fpText1_Change()
If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    RS2.Open "SELECT reg_codigo,reg_nombre FROM a_regimen WHERE reg_codigo LIKE '%" & UCase(LimpiaDato(fpText1.text)) & "%'", vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    RS2.Open "SELECT reg_codigo, reg_nombre FROM a_regimen WHERE UCASE(reg_nombre) LIKE '%" & UCase(LimpiaDato(fpText1.text)) & "%'", vg_db, adOpenStatic
End If
'If ibusca <> RS2.RecordCount Then ibusca = RS2.RecordCount:
vaSpread1.MaxRows = RS2.RecordCount
i = 1
If Not RS2.EOF Then
   Do While Not RS2.EOF
      vaSpread1.Row = i
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.Value = RS2!reg_codigo
      vaSpread1.Col = 2
      vaSpread1.CellType = IIf(RS2!reg_codigo > 9999, CellTypeStaticText, CellTypeEdit)
      vaSpread1.Value = IIf(IsNull(RS2!reg_nombre), "", Trim(RS2!reg_nombre))
      RS2.MoveNext: i = i + 1
   Loop
   Gl_Ac_Botones Me, 1, 1, modo
End If
RS2.Close: Set RS2 = Nothing
If fpText1.text = "" Then
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Else
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codigo As Long, Nombre As String, orden As String
On Error GoTo Man_Error
Select Case Button.INDEX
Case 1
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 2: vaSpread1.text = "": vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus
Case 3
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
Case 5
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
    vg_db.BeginTrans
    vg_db.Execute "DELETE a_regimen FROM a_regimen WHERE reg_codigo=" & codigo
    vg_db.CommitTrans
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
Case 7
    fpText1.text = ""
    MoverDatosGrillas
Case 10
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    If modo = "A" Then
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.DeleteRows vaSpread1.Row, 1
       vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    Else
       Cancela
    End If
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Combo1.Enabled = True: fpText1.Enabled = True
Case 12
    GrabaRegistro vaSpread1.ActiveRow
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_Regime
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub MoverDatosGrillas()
vaSpread1.MaxRows = 0
RS2.Open "SELECT DISTINCT * FROM b_clientecencos WHERE clc_codcli='" & codcli & "'", vg_db, adOpenStatic
Do While Not RS2.EOF
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = 1: vaSpread1.Value = RS2!clc_codigo
    vaSpread1.Col = 2
    vaSpread1.CellType = CellTypeStaticText
    vaSpread1.Value = IIf(IsNull(RS2!clc_nombre), "", Trim(RS2!clc_nombre))
    RS2.MoveNext
Loop
RS2.Close: Set RS2 = Nothing
Gl_Ac_Botones Me, 1, 1, modo
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = True
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    GrabaRegistro Row
ElseIf Toolbar1.Buttons(12).Visible = False Then
    Cancela
End If
End Sub

Private Sub Cancela()
OpGr = True
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
RS1.Open "SELECT * FROM a_regimen WHERE reg_codigo=" & codigo, vg_db, adOpenStatic
If Not RS1.EOF Then
   vaSpread1.Col = 2: vaSpread1.Value = IIf(IsNull(RS1!reg_nombre), "", Trim(RS1!reg_nombre))
End If
RS1.Close: Set RS1 = Nothing
OpGr = False
End Sub

Sub LlenarDatos(Fecha As String, rutcli As String)
Label1(0).Caption = Fecha
codcli = rutcli
End Sub
