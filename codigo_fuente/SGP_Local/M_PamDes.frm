VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_PamDes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametro Despacho"
   ClientHeight    =   5550
   ClientLeft      =   3345
   ClientTop       =   3120
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   700
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   6855
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   915
         TabIndex        =   3
         Top             =   210
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
         _ExtentY        =   556
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
         ThreeDOutsideHighlightColor=   16777215
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
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   16777215
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   10
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
         ButtonAlign     =   1
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2145
         Picture         =   "M_PamDes.frx":0000
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contrato"
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
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   735
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2655
         TabIndex        =   4
         Top             =   210
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2700
         TabIndex        =   6
         Top             =   255
         Width           =   3975
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   11880
      _Version        =   393216
      _ExtentX        =   20955
      _ExtentY        =   7223
      _StockProps     =   64
      ButtonDrawMode  =   1
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
      MaxCols         =   14
      MaxRows         =   20
      ProcessTab      =   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "M_PamDes.frx":030A
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_PamDes"
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
Me.Height = 6030
Me.Width = 12225
Msgtitulo = "Parametro Despacho"
fg_centra Me
modo = "": est = False
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = False
'Toolbar1.Buttons(5).Visible = False
'Toolbar1.Buttons(6).Visible = False
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
RS.Open "SELECT b.tip_codigo, b.tip_nombre FROM a_tipopro a INNER JOIN a_tipopro AS b ON a.tip_codigo = b.tip_previo WHERE a.tip_previo = 0  AND a.tip_activo in ('1','S')", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      RS1.Open "SELECT DISTINCT pad_codigo FROM b_paramdesp WHERE pad_cencos = '" & MuestraCasino(1) & "' AND pad_codigo = " & RS!tip_codigo & "", vg_db, adOpenStatic
      If RS1.EOF Then
         vg_db.Execute "INSERT INTO b_paramdesp VALUES (" & RS!tip_codigo & ", '', '" & MuestraCasino(1) & "', '', 0)"
      End If
      RS1.Close: Set RS1 = Nothing
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing

MoverDatosGrilla
End Sub

Sub MoverDatosGrilla()
Dim i As Long, codaux As Long
On Error GoTo Man_Error
With vaSpread1
    .Visible = False
    .MaxRows = 0
    RS.Open "SELECT * FROM a_tipopro WHERE tip_previo = 0 AND tip_activo in ('1','S')", vg_db, adOpenStatic
    If Not RS.EOF Then
       Do While Not RS.EOF
          i = 1
          RS1.Open "SELECT a.pad_tipo, a.pad_diario, a.pad_diaseg, b.tip_codigo, b.tip_previo, b.tip_nombre FROM b_paramdesp a, a_tipopro b WHERE a.pad_codigo = b.tip_codigo AND b.tip_previo = " & RS!tip_codigo & " AND a.pad_cencos = '" & fpText.text & "' AND b.tip_activo in ('1','S') ORDER BY b.tip_nombre", vg_db, adOpenStatic
          If Not RS1.EOF Then
             Do While Not RS1.EOF
                If i = 1 Then
                   .MaxRows = .MaxRows + 1
                   .Row = .MaxRows
                   .Col = 1: .text = ""
                   .Col = 2: .FontBold = True: .text = Trim(RS!tip_nombre)
                   .Col = 3: .CellType = CellTypeStaticText
                   .Col = 4: .CellType = CellTypeStaticText
                   .Col = 5: .CellType = CellTypeStaticText
                   .Col = 6: .CellType = CellTypeStaticText
                   .Col = 7: .CellType = CellTypeStaticText
                   .Col = 8: .CellType = CellTypeStaticText
                   .Col = 9: .CellType = CellTypeStaticText
                   .Col = 10: .CellType = CellTypeStaticText
                   .Col = 11: .CellType = CellTypeStaticText
                   .Col = 12: .CellType = CellTypeStaticText
                   .Col = 13: .CellType = CellTypeStaticText
                   .Col = 14: .CellType = CellTypeStaticText
                End If
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 1: .text = RS1!tip_codigo
                .Col = 2: .text = IIf(IsNull(RS1!tip_nombre), "", Trim(RS1!tip_nombre))
                .Col = 3: .text = IIf(IsNull(RS1!pad_diaseg) Or RS1!pad_diaseg = 0, "", RS1!pad_diaseg)
                .Col = 4: .TypeComboBoxList = "MENSUAL" & Chr$(9) & "QUINCENAL" & Chr$(9) & "SEMANAL" & Chr$(9) & "10 DIAS" & Chr$(9) & "DIARIO"
                .Col = 5: .TypeComboBoxList = "M" & Chr$(9) & "Q" & Chr$(9) & "S" & Chr$(9) & "D" & Chr$(9) & "E"
                For i = 0 To .TypeComboBoxCount
                    .TypeComboBoxCurSel = i
                    If .text = Mid(RS1!pad_tipo, 1, 1) Then j = i: Exit For
                    j = -1
                Next i
                .Col = 4: .TypeComboBoxCurSel = j
                If j = 1 Then
                   .Col = 6: .CellType = CellTypeComboBox: .TypeComboBoxList = "QUINCENAL 1-15" & Chr$(9) & "QUINCENAL 2-16" & Chr$(9) & "QUINCENAL 3-17" & Chr$(9) & "QUINCENAL 4-18"
                   .Col = 7: .CellType = CellTypeComboBox: .TypeComboBoxList = "Q1" & Chr$(9) & "Q2" & Chr$(9) & "Q3" & Chr$(9) & "Q4"
                   For i = 0 To .TypeComboBoxCount
                       .TypeComboBoxCurSel = i
                       If .text = Trim(RS1!pad_tipo) Then j = i: Exit For
                       j = -1
                  Next i
                  .Col = 6: .TypeComboBoxCurSel = j
                Else
                   .Col = 6: .CellType = CellTypeStaticText
                   .Col = 7: .CellType = CellTypeStaticText
                End If
                est = True
                If Trim(RS1!pad_tipo) <> "E" And Trim(RS1!pad_tipo) <> "S" Then
                   .Col = 8: .CellType = CellTypeStaticText
                   .Col = 9: .CellType = CellTypeStaticText
                   .Col = 10: .CellType = CellTypeStaticText
                   .Col = 11: .CellType = CellTypeStaticText
                   .Col = 12: .CellType = CellTypeStaticText
                   .Col = 13: .CellType = CellTypeStaticText
                   .Col = 14: .CellType = CellTypeStaticText
                Else
                   .Col = 8: .CellType = CellTypeCheckBox: .TypeHAlign = TypeHAlignCenter: .text = IIf(IsNull(RS1!pad_diario), 0, Mid(RS1!pad_diario, 1, 1))
                   .Col = 9: .CellType = CellTypeCheckBox: .TypeHAlign = TypeHAlignCenter: .text = IIf(IsNull(RS1!pad_diario), 0, Mid(RS1!pad_diario, 2, 1))
                   .Col = 10: .CellType = CellTypeCheckBox: .TypeHAlign = TypeHAlignCenter: .text = IIf(IsNull(RS1!pad_diario), 0, Mid(RS1!pad_diario, 3, 1))
                   .Col = 11: .CellType = CellTypeCheckBox: .TypeHAlign = TypeHAlignCenter: .text = IIf(IsNull(RS1!pad_diario), 0, Mid(RS1!pad_diario, 4, 1))
                   .Col = 12: .CellType = CellTypeCheckBox: .TypeHAlign = TypeHAlignCenter: .text = IIf(IsNull(RS1!pad_diario), 0, Mid(RS1!pad_diario, 5, 1))
                   .Col = 13: .CellType = CellTypeCheckBox: .TypeHAlign = TypeHAlignCenter: .text = IIf(IsNull(RS1!pad_diario), 0, Mid(RS1!pad_diario, 6, 1))
                   .Col = 14: .CellType = CellTypeCheckBox: .TypeHAlign = TypeHAlignCenter: .text = IIf(IsNull(RS1!pad_diario), 0, Mid(RS1!pad_diario, 7, 1))
                End If
                est = False
                i = 2
                RS1.MoveNext
             Loop
          End If
          RS1.Close: Set RS1 = Nothing
          RS.MoveNext
       Loop
    End If
    RS.Close: Set RS = Nothing
    .SetActiveCell 3, 1
    .Visible = True
End With
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub fpText_Change()
RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & fpText.text & "' AND cli_tipo = 0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.text = vg_codigo
    fpayuda(0).Caption = vg_nombre
    MoverDatosGrilla
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 3 '-------> Modificar
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
'    Toolbar1.Buttons(5).Visible = False
'    Toolbar1.Buttons(6).Visible = False
Case 7 '-------> Actualizar lista
    MoverDatosGrilla
Case 5 '-------> eliminar datos
    With vaSpread1
        .Row = .ActiveRow
        .Col = 3: .text = ""
        .Col = 4: .CellType = CellTypeStaticText: .text = ""
        .Col = 5: .CellType = CellTypeStaticText: .text = ""
        .Col = 6: .CellType = CellTypeStaticText: .text = ""
        .Col = 7: .CellType = CellTypeStaticText: .text = ""
        .Col = 8: .CellType = CellTypeStaticText: .text = ""
        .Col = 9: .CellType = CellTypeStaticText: .text = ""
        .Col = 10: .CellType = CellTypeStaticText: .text = ""
        .Col = 11: .CellType = CellTypeStaticText: .text = ""
        .Col = 12: .CellType = CellTypeStaticText: .text = ""
        .Col = 13: .CellType = CellTypeStaticText: .text = ""
        .Col = 14: .CellType = CellTypeStaticText: .text = ""
    '    If modo = "" Then modo = "M"
    '    Gl_Ac_Botones Me, 1, 0, modo
    
        .Row = .ActiveRow
        .Col = 1
        If .ActiveRow < 1 Or Trim(.text) = "" Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    '    vg_db.BeginTrans
    '    vg_db.Execute "DELETE b_paramdesp from b_paramdesp WHERE pad_codigo = " & codpad & " AND pad_cencos = '" & Trim(fpText.text) & "'"
    '    vg_db.CommitTrans
    '    .DeleteRows .Row, 1
    '    .MaxRows = .MaxRows - 1
        Dim codpad As Long
        .Row = .ActiveRow
        .Col = 1: codpad = Val(.Value)
        vg_db.Execute ("UPDATE b_paramdesp SET pad_tipo = '', pad_diario = '', pad_diaseg = 0 WHERE pad_codigo = " & codpad & " AND pad_cencos = '" & Trim(fpText.text) & "'")
        modo = "": Gl_Ac_Botones Me, 1, IIf(.MaxRows = 0, 2, 1), modo
    End With
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
Case 10 '-------> Cancelar
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    MoverDatosGrilla
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
'    Toolbar1.Buttons(5).Visible = False
'    Toolbar1.Buttons(6).Visible = False
Case 12 '-------> Gabar datos
    Dim vCodFam As Long, tipdes As String, tipdes1 As String, i As Long, j As Long, X As Long, estext As Boolean, desdia As String
    With vaSpread1
        '-------> Validar datos
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If Trim(.text) <> "" Then
               .Col = 5: tipdes = .text
               If tipdes = "E" Or tipdes = "S" Then
                  estext = False
                  For j = 8 To .MaxCols
                      .Col = j
                      If .text <> "0" And Trim(.text) <> "" Then estext = True
                  Next j
                  If Not estext Then MsgBox "No a especificado los días de despachos", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
               End If
               .Col = 7: tipdes1 = .text
               If Trim(tipdes) = "Q" And Trim(tipdes1) = "" Then
                  MsgBox "No a especificado la segunda quincena de despachos", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
               End If
            End If
        Next i
        
        vg_db.BeginTrans
        vg_db.Execute "DELETE b_paramdesp FROM b_paramdesp WHERE pad_cencos = '" & Trim(fpText.text) & "'"
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            vCodFam = Val(.text)
            .Col = 5
            tipdes = "": tipdes = .text
            If Trim(vCodFam) <> "" And tipdes <> "" Then
               .Col = 1
               vCodFam = Val(.text)
               .Col = 3: candiaseg = Val(.text)
               .Col = 5: tipdes = .text
               If tipdes = "Q" Then
                  .Col = 7: tipdes = .text
               End If
               desdia = ""
               If tipdes = "E" Or tipdes = "S" Then
                  X = 1
                  For j = 8 To .MaxCols
                      .Col = j
                      desdia = desdia & IIf(Trim(.text) = "" Or Trim(.text) = "0", "0", X) 'vaSpread4.text)
                      X = X + 1
                  Next j
               End If
               vg_db.Execute "INSERT INTO b_paramdesp VALUES (" & vCodFam & ", '" & tipdes & "', '" & Trim(fpText.text) & "', '" & desdia & "', " & candiaseg & ")"
            End If
        Next i
        vg_db.CommitTrans
        modo = "": Gl_Ac_Botones Me, 1, 1, modo
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = False
    '    Toolbar1.Buttons(5).Visible = False
    '    Toolbar1.Buttons(6).Visible = False
    End With
Case 15 '-------> Imprimir
    RS.Open "SELECT DISTINCT pad_codigo FROM b_paramdesp WHERE pad_cencos = '" & Trim(fpText.text) & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    I_ParametroDespacho Trim(fpText.text)
Case 18 '-------> Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Or 2147217900 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
If modo = "" Then modo = "M"
'If modo = "M" And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = False
If modo = "M" And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False
End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
With vaSpread1
    If .MaxRows < 1 Then Exit Sub
    If modo = "" Then modo = "M"
    'If modo = "M" And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = False
    If modo = "M" And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False:
    Dim tipdes As String
    .Row = Row
    Select Case Col
    Case 4
        .Col = 4: tipdes = .TypeComboBoxCurSel
        .Col = 5: .TypeComboBoxCurSel = tipdes
        .EditEnterAction = EditEnterActionNone
        est = True
        If .text <> "S" And .text <> "E" Then
           .Col = 6: .CellType = CellTypeStaticText: .text = ""
           .Col = 7: .CellType = CellTypeStaticText: .text = ""
           .Col = 8: .CellType = CellTypeStaticText: .text = ""
           .Col = 9: .CellType = CellTypeStaticText: .text = ""
           .Col = 10: .CellType = CellTypeStaticText: .text = ""
           .Col = 11: .CellType = CellTypeStaticText: .text = ""
           .Col = 12: .CellType = CellTypeStaticText: .text = ""
           .Col = 13: .CellType = CellTypeStaticText: .text = ""
           .Col = 14: .CellType = CellTypeStaticText: .text = ""
           '-------> Mover datos segunda quincena
           .Col = 6
           If tipdes = "1" Then
              .Col = 6: .CellType = CellTypeComboBox: .TypeComboBoxList = "QUINCENAL 1-15" & Chr$(9) & "QUINCENAL 2-16" & Chr$(9) & "QUINCENAL 3-17" & Chr$(9) & "QUINCENAL 4-18"
              .Col = 7: .CellType = CellTypeComboBox: .TypeComboBoxList = "Q1" & Chr$(9) & "Q2" & Chr$(9) & "Q3" & Chr$(9) & "Q4"
           End If
        Else
           .Col = 6: .CellType = CellTypeStaticText: .text = ""
           .Col = 7: .CellType = CellTypeStaticText: .text = ""
           .Col = 8: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
           .Col = 9: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
           .Col = 10: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
           .Col = 11: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
           .Col = 12: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
           .Col = 13: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
           .Col = 14: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter
        End If
        est = False
    Case 6
        .Col = 6: tipdes = .TypeComboBoxCurSel
        .Col = 7: .TypeComboBoxCurSel = tipdes
        .EditEnterAction = EditEnterActionNone
    End Select
    est = False
End With
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim i As Long
If (Col <> 8 And Col <> 9 And Col <> 10 And Col <> 11 And Col <> 12 And Col <> 13 And Col <> 14) Or Row = 0 Or est Then Exit Sub
vaSpread1.Row = Row
vaSpread1.Col = 4
If vaSpread1.TypeComboBoxCurSel = 2 Then
   For i = 8 To 14
       If i <> Col Then est = True: vaSpread1.Col = i: vaSpread1.text = "0": est = False
   Next i
End If
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
With vaSpread1
    .Row = .ActiveRow
    .Col = .ActiveCol
    If .MaxRows < 1 Or .Col <> 4 Then Exit Sub
    Select Case KeyCode
    Case 46
        .Row = .ActiveRow
        .Col = 4: .CellType = CellTypeStaticText: .text = ""
        .Col = 5: .CellType = CellTypeStaticText: .text = ""
        .Col = 6: .CellType = CellTypeStaticText: .text = ""
        .Col = 7: .CellType = CellTypeStaticText: .text = ""
        .Col = 8: .CellType = CellTypeStaticText: .text = ""
        .Col = 9: .CellType = CellTypeStaticText: .text = ""
        .Col = 10: .CellType = CellTypeStaticText: .text = ""
        .Col = 11: .CellType = CellTypeStaticText: .text = ""
        .Col = 12: .CellType = CellTypeStaticText: .text = ""
        .Col = 13: .CellType = CellTypeStaticText: .text = ""
        .Col = 14: .CellType = CellTypeStaticText: .text = ""
        If modo = "" Then modo = "M"
        Gl_Ac_Botones Me, 1, 0, modo
    End Select
End With
End Sub
