VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_PNRece 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametro Nº Recetas a Incluir"
   ClientHeight    =   6540
   ClientLeft      =   3780
   ClientTop       =   1830
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   6135
      Left            =   105
      TabIndex        =   0
      Top             =   360
      Width           =   6975
      _Version        =   393216
      _ExtentX        =   12303
      _ExtentY        =   10821
      _StockProps     =   64
      ButtonDrawMode  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   2
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
      MaxCols         =   3
      SpreadDesigner  =   "M_PNRece.frx":0000
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_PNRece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim est As Boolean
Dim Msgtitulo As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 7050
Me.Width = 7245
Msgtitulo = "Parametro Nº Recetas a Incluir"
fg_centra Me
modo = "": est = False
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = False
Toolbar1.Buttons(5).Visible = False
Toolbar1.Buttons(6).Visible = False
MoverDatosGrilla
End Sub

Sub MoverDatosGrilla()
On Error GoTo Man_Error
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
RS.Open "sgpadm_s_paramnreceta", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 1: vaSpread1.text = RS!sub_codigo
      vaSpread1.Col = 2: vaSpread1.text = Trim(RS!sub_nombre)
      vaSpread1.Col = 3: vaSpread1.text = IIf(RS!pnr_nreceta = 0, "", Format(RS!pnr_nreceta, fg_Pict(6, 0)))
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
vaSpread1.Col = -1: vaSpread1.Row = -1
vaSpread1.Lock = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1" Or Mid(ValidarUsuario(Me), 2, 1) = "1", False, True)
vaSpread1.SetActiveCell 3, 1
vaSpread1.Visible = True
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 3
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
Case 7
    MoverDatosGrilla
Case 10
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    MoverDatosGrilla
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
Case 12
    Dim vNumRec As Long, CodSeg As Long
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1: CodSeg = vaSpread1.text
        vaSpread1.Col = 3: vNumRec = IIf(Trim(vaSpread1.text) <> "", vaSpread1.text, 0)
        RS.Open "SELECT pnr_codseg FROM b_paramnreceta WHERE pnr_codseg=" & CodSeg & "", vg_db, adOpenStatic
        If RS.EOF Then
           vg_db.Execute "INSERT INTO b_paramnreceta (pnr_codseg, pnr_nreceta) VALUES (" & CodSeg & ", " & vNumRec & ")"
        Else
           vg_db.Execute "UPDATE b_paramnreceta SET pnr_nreceta=" & vNumRec & " WHERE pnr_codseg=" & CodSeg & ""
        End If
        RS.Close: Set RS = Nothing
    Next i
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
Case 15
    RS.Open "SELECT DISTINCT pnr_codseg FROM b_paramnreceta", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    I_ParametroNReceta
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

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
If modo = "" Then modo = "M"
If modo = "M" And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = False
vaSpread1.Row = Row
vaSpread1.Col = Col: If vaSpread1.text = "0" Then vaSpread1.text = ""
End Sub
