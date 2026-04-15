VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_AsoLPr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asociar Lista de Precio"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11655
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "M_AsoLPr.frx":0000
         Left            =   1560
         List            =   "M_AsoLPr.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   390
         Width           =   4575
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4875
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   11415
         _Version        =   393216
         _ExtentX        =   20135
         _ExtentY        =   8599
         _StockProps     =   64
         ButtonDrawMode  =   1
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
         MaxCols         =   7
         MaxRows         =   1000000
         SpreadDesigner  =   "M_AsoLPr.frx":0004
         VirtualMode     =   -1  'True
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   5760
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Sub-Segmento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   1605
         TabIndex        =   5
         Top             =   510
         Width           =   4575
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000018&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_AsoLPr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim modo As String, Codigo As Long
Dim Msgtitulo As String
Dim vLisPre() As Variant
Dim Est As Boolean

Private Sub Combo1_Click(Index As Integer)
If Est Then Exit Sub
Codigo = Trim(fg_codigocbo(Combo1, 0, 10, ""))
MoverDatosGrillas
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Dim RS As New ADODB.Recordset
Me.HelpContextID = vg_OpcM
Me.Height = 7215
Me.Width = 12015
Msgtitulo = "Asociar Lista de Precio"
fg_centra Me
modo = "": Codigo = 0: Est = True
Gl_Mo_Botones Me, 1
Combo1(0).Clear
Set RS = vg_db.Execute("sgpadm_s_subsegmento 1, 0, '', ''")
'Combo1(0).AddItem "Todos" & Space(150) & "(" & fg_pone_cero(Str(0), 10) & ")"
Do While Not RS.EOF
    Combo1(0).AddItem IIf(IsNull(RS!sub_nombre), "", RS!sub_nombre) & Space(150) & "(" & fg_pone_cero(Str(RS!sub_codigo), 10) & ")"
    RS.MoveNext
Loop
RS.Close: Set RS = Nothing
Combo1(0).ListIndex = -1
vaSpread1.MaxRows = 0
Gl_Ac_Botones Me, 1, 3, modo
'MoverDatosGrillas
Est = False
End Sub

Private Sub MoverDatosGrillas()
Dim RS As New ADODB.Recordset
Dim v_inicio As Long, v_final As Long, i As Long, ii As Long, j As Long, z As Long, auxsub As Long
fg_carga ""
'-------> Mover lista precio vector
Set RS = vg_db.Execute("sgpadm_s_listaprecio 4, 0, 0, '" & vg_NUsr & "'")
i = 1
If Not RS.EOF Then
   ReDim vLisPre(RS!nReg, 2)
   Do While Not RS.EOF
      vLisPre(i, 1) = RS!lpr_codigo
      vLisPre(i, 2) = RS!lpr_nombre
      i = i + 1
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing

'-------> Mover datos grilla
With vaSpread1
    .Visible = False
    .MaxRows = 0
    .Row = -1: .Col = -1
    .BackColor = Shape1(0).FillColor
    Bar1(0).Visible = True: Bar1(0).Value = 0
    auxsub = 0
    Set RS = vg_db.Execute("sgpadm_s_subsegmento 2, " & Codigo & ", 0, ''")
    If Not RS.EOF Then
        .MaxRows = RS!nReg
        ii = 1
        Do While Not RS.EOF
        '    .MaxRows = .MaxRows + 1
        '    .Row = .MaxRows
            Bar1(0).Value = Val((ii / .MaxRows) * 100)
            .Row = ii
            If RS!sub_codigo <> auxsub Then
               .Col = 2: .Value = RS!sub_codigo & " - " & Trim(RS!sub_nombre)
               auxsub = RS!sub_codigo
            End If
            .Col = 1: .Value = RS!sub_codigo
            .Col = 3: .Value = RS!reg_codigo & " - " & Trim(RS!reg_nombre)
            .Col = 4: .Value = RS!reg_codigo
            If i > 1 Then
               lisnom = "": liscod = "": encuentra = False
               '-------> Mover lista precio
               For j = 1 To UBound(vLisPre)
                   .Col = 5: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vLisPre(j, 2))
                   .Col = 6: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vLisPre(j, 1)
                   .Col = 5: .TypeComboBoxList = lisnom
                   .Col = 6: .TypeComboBoxList = liscod
               Next j
               .Col = 6
               codaux = -1
               For z = 0 To .TypeComboBoxCount
                   .TypeComboBoxCurSel = z
                   If .text = RS!lpr_codigo Then codaux = z: Exit For
                   codaux = -1
               Next z
               .Col = 5: .TypeComboBoxCurSel = codaux
            End If
            .Col = 7: .Value = RS!sub_codigo & " - " & Trim(RS!sub_nombre)
            RS.MoveNext: ii = ii + 1
        Loop
    End If
    RS.Close: Set RS = Nothing
    .Visible = True
    Gl_Ac_Botones Me, 1, IIf(.MaxRows > 0, 4, 2), modo
End With
Bar1(0).Visible = False
fg_descarga
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codsse As Long, codreg As Long, codlpr As Long, i As Long
On Error GoTo Man_Error
Select Case Button.Index
Case 3 '-------> Modificar
    modo = "M"
    Combo1(0).Enabled = False
    Gl_Ac_Botones Me, 1, 0, modo
Case 5 '-------> Eliminar
    If vaSpread1.MaxRows < 1 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 6
    If Trim(vaSpread1.text) = "" Then Exit Sub
    If MsgBox("Eliminar Dato", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codsse = vaSpread1.text
    vaSpread1.Col = 4: codreg = vaSpread1.text
    vaSpread1.Col = 6: codlpr = vaSpread1.text
    vg_db.Execute "DELETE FROM b_asolistaprecio WHERE alp_codsse=" & codsse & " AND alp_codreg=" & codreg & " AND alp_codlpr=" & codlpr & ""
    vaSpread1.Col = 6: vaSpread1.TypeComboBoxCurSel = -1: vaSpread1.text = ""
    vaSpread1.Col = 5: vaSpread1.TypeComboBoxCurSel = -1: vaSpread1.text = ""
Case 7, 10 '-------> Actualizar lista y cancelar
    MoverDatosGrillas
    Combo1(0).Enabled = True
Case 12 '------> Confirmar
    fg_carga ""
    Bar1(0).Visible = True: Bar1(0).Value = 0
    For i = 1 To vaSpread1.MaxRows
        Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
        vaSpread1.Row = i
        vaSpread1.Col = 6
        If Trim(vaSpread1.text) <> "" Then
           vaSpread1.Col = 1: codsse = vaSpread1.text
           vaSpread1.Col = 4: codreg = vaSpread1.text
           vaSpread1.Col = 6: codlpr = vaSpread1.text
           '-------> Borrar registro
           vg_db.Execute "DELETE FROM b_asolistaprecio WHERE alp_codsse=" & codsse & " AND alp_codreg=" & codreg & ""
           '------>  Agregar registro
           vg_db.Execute "INSERT INTO b_asolistaprecio (alp_codsse, alp_codreg, alp_codlpr) VALUES (" & codsse & ", " & codreg & ", " & codlpr & ")"
        End If
    Next i
    Bar1(0).Visible = False: Bar1(0).Value = 0
    modo = ""
    Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows > 0, 4, 2), modo
    Combo1(0).Enabled = True
    fg_descarga
    MsgBox "Generación grabado finalizado sin problema", vbInformation + vbOKOnly, Msgtitulo
Case 15 '-------> Imprimir
    I_AsociarListaPrecio
Case 18 '-------> Salir
   Me.Hide
   Unload Me
End Select
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = NewRow
vaSpread1.Col = 7
Frame1.Caption = vaSpread1.text
End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Select Case Col
Case 5
    Dim indice As Long
    vaSpread1.Row = Row
    vaSpread1.Col = 5: indice = vaSpread1.TypeComboBoxCurSel
    vaSpread1.Col = 6: vaSpread1.TypeComboBoxCurSel = indice
    If modo = "" Then modo = "M"
    If Toolbar1.Buttons(12).Visible = False Then Combo1(0).Enabled = False: Gl_Ac_Botones Me, 1, 0, modo
'    vaSpread1.EditEnterAction = EditEnterActionNone
End Select
End Sub
