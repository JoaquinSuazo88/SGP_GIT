VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form B_PedPac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos Paciente : JPAZ"
   ClientHeight    =   3045
   ClientLeft      =   3270
   ClientTop       =   2970
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _Version        =   393216
      _ExtentX        =   14631
      _ExtentY        =   4895
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      SpreadDesigner  =   "B_PedPac.frx":0000
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   3045
      Left            =   8565
      TabIndex        =   1
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   5371
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "B_PedPac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim Msgtitulo As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_centra Me
fg_carga ""
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fg_descarga
End Sub

Sub LlenarPedidoPaciente(rutpac As String)
Me.Caption = "Pedidos Paciente : "
RS.Open "SELECT pac_nombre, pac_appaterno, pac_apmaterno FROM b_pacientes WHERE pac_codigo = '" & rutpac & "'", vg_db, adOpenStatic
If Not RS.EOF Then Me.Caption = "Pedidos Paciente : " & fg_PintaRut(rutpac) & " - " & Trim(UCase(RS!pac_nombre)) & " " & Trim(UCase(RS!pac_appaterno)) & " " & Trim(UCase(RS!pac_apmaterno))
RS.Close: Set RS = Nothing
RS.Open "SELECT a.top_codigo, a.top_fecped, a.top_codreg, c.reg_nombre, b.usu_nombre " & _
        "FROM b_tomapedido a, a_usuarios b, a_regimen c " & _
        "WHERE a.top_codusu = b.usu_codigo " & _
        "AND   a.top_codreg = c.reg_codigo " & _
        "AND   a.top_codpac = '" & rutpac & "'", vg_db, adOpenStatic
With vaSpread1
    .MaxRows = 0
    If Not RS.EOF Then
       Do While Not RS.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1: .text = RS!top_codigo
          .Col = 2: .text = RS!top_fecped
          .Col = 3: .text = "Normal"
          .Col = 4: .text = Trim(RS!reg_nombre)
          .Col = 6: .text = Trim(RS!usu_nombre)
          RS.MoveNext
       Loop
    End If
End With
RS.Close: Set RS = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If vaSpread1.MaxRows < 1 Then Exit Sub
    MoverDatos
Case 3
    vg_codigo = ""
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
MoverDatos
End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
MoverDatos
End Sub

Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 27
    Cerrar
End Select
End Sub

Private Sub MoverDatos()
With vaSpread1
    If .MaxRows < 1 Then Exit Sub
    .Row = .ActiveRow
    .Col = 1: vg_codigo = Val(.text)
    .Col = 2: vg_nombre = .text
End With
Cerrar
End Sub

Sub Cerrar()
Me.Hide
Unload Me
End Sub
