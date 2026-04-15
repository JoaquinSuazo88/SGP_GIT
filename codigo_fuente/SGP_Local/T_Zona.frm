VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form T_Zona 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zona"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   971
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "T_Zona.frx":0000
         Left            =   2025
         List            =   "T_Zona.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2500
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   2025
         TabIndex        =   2
         Top             =   555
         Width           =   2505
         _Version        =   196608
         _ExtentX        =   4410
         _ExtentY        =   870
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
         ThreeDOutsideHighlightColor=   -2147483628
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
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
         ThreeDFrameColor=   -2147483637
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Columna"
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
         Left            =   525
         TabIndex        =   5
         Top             =   345
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Texto"
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
         Index           =   1
         Left            =   525
         TabIndex        =   4
         Top             =   645
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
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
         Left            =   4605
         TabIndex        =   3
         Top             =   645
         Width           =   585
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3375
      Left            =   15
      TabIndex        =   6
      Top             =   1425
      Width           =   9795
      _Version        =   393216
      _ExtentX        =   17277
      _ExtentY        =   5953
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      AutoCalc        =   0   'False
      ButtonDrawMode  =   1
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
      FormulaSync     =   0   'False
      MaxCols         =   4
      ProcessTab      =   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "T_Zona.frx":001E
      ScrollBarTrack  =   1
      ClipboardOptions=   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "T_Zona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String
Dim OpGr As Boolean
Dim vLisPre() As Variant
Dim EstVect As Boolean

Private Sub GrabaRegistro(Fila As Long)
Dim Codigo As Long, Nombre As String, codlpr As Long
On Error GoTo Man_Error
OpGr = True
vaSpread1.Row = Fila
vaSpread1.Col = 1: Codigo = Val(vaSpread1.Value)
vaSpread1.Col = 2: Nombre = Trim(LimpiaDato(vaSpread1.Value))
vaSpread1.Col = 3:
If Trim(Nombre) = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, Fila: vaSpread1.SetFocus: OpGr = False: Exit Sub
If vaSpread1.TypeComboBoxCurSel = -1 Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 3, Fila: vaSpread1.SetFocus: OpGr = False: Exit Sub
vaSpread1.Col = 4: codlpr = Val(vaSpread1.text)
If modo = "A" Then
    Set RS1 = vg_db.Execute("sgpadm_s_zona 2, 0,''")
    If Not RS1.EOF Then RS1.MoveFirst: Codigo = RS1!zon_codigo + 1 Else Codigo = 1
    RS1.Close: Set RS1 = Nothing
    vg_db.Execute "INSERT INTO a_zona (zon_codigo, zon_nombre, zon_codlpr) " & _
                  "VALUES (" & Codigo & ", '" & Trim(Nombre) & "', " & codlpr & ")"
    vaSpread1.Col = 1: vaSpread1.text = Codigo
Else
    vg_db.Execute "UPDATE a_zona SET zon_nombre='" & Trim(Nombre) & "', zon_codlpr=" & codlpr & " WHERE zon_codigo=" & Codigo & ""
End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Combo1.Enabled = True: fpText1.Enabled = True
modo = "": Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
OpGr = False
Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 5355
Me.Width = 9960
Msgtitulo = "Zona"
fg_centra Me
EstVect = False
modo = ""
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
Combo1.ListIndex = 1
MoverDatosGrillas
OpGr = False
End Sub

Private Sub Form_Resize()
Frame1.Move IIf(Me.WindowState = 2, 4200, 2000), 360, 6015, 971
If Me.WindowState = 0 Then
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
End If
Toolbar1.Refresh
End Sub

Private Sub fpText1_Change()
If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
Dim i As Long, z As Long, lisnom As String, liscod As String, codaux As Long
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
   Set RS2 = vg_db.Execute("sgpadm_s_zona 3, 0,'%" & UCase(LimpiaDato(fpText1.text)) & "%'")
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
   Set RS2 = vg_db.Execute("sgpadm_s_zona 4, 0,'%" & UCase(LimpiaDato(fpText1.text)) & "%'")
End If
If RS2.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS2!nReg
i = 1
If Not RS2.EOF Then
   Do While Not RS2.EOF
      vaSpread1.Row = i
      i = i + 1
      vaSpread1.Col = 1: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = RS2!zon_codigo
      vaSpread1.Col = 2: vaSpread1.text = Trim(RS2!Zon_nombre)
      lisnom = "": liscod = ""
      '-------> Mover lista precio
      If EstVect Then
         For z = 1 To UBound(vLisPre)
             vaSpread1.Col = 3: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vLisPre(z, 2))
             vaSpread1.Col = 4: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vLisPre(z, 1)
             vaSpread1.Col = 3: vaSpread1.TypeComboBoxList = lisnom
             vaSpread1.Col = 4: vaSpread1.TypeComboBoxList = liscod
         Next z
      End If
      vaSpread1.Col = 4
      codaux = -1
      For z = 0 To vaSpread1.TypeComboBoxCount
          vaSpread1.TypeComboBoxCurSel = z
          If vaSpread1.text = IIf(IsNull(RS2!zon_codlpr), 0, RS2!zon_codlpr) Then codaux = z: Exit For
          codaux = -1
      Next z
      vaSpread1.Col = 3: vaSpread1.TypeComboBoxCurSel = codaux
      RS2.MoveNext
   Loop
   Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
End If
RS2.Close: Set RS2 = Nothing
If fpText1.text = "" Then
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Else
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Codigo As Long, Nombre As String, orden As String, z As Long, lisnom As String, liscod As String
'On Error GoTo Man_Error
Select Case Button.Index
Case 1
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    lisnom = "": liscod = ""
    '-------> Mover lista precio
    If EstVect Then
       For z = 1 To UBound(vLisPre)
           vaSpread1.Col = 3: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vLisPre(z, 2))
           vaSpread1.Col = 4: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vLisPre(z, 1)
           vaSpread1.Col = 3: vaSpread1.TypeComboBoxList = lisnom
           vaSpread1.Col = 4: vaSpread1.TypeComboBoxList = liscod
       Next z
    End If
    vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus
Case 3
    If vaSpread1.MaxRows < 1 Then Exit Sub
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
Case 5
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: Codigo = Val(vaSpread1.Value)
    vg_db.Execute "DELETE a_zona FROM a_zona WHERE zon_codigo=" & Codigo & ""
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    modo = "": Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
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
    modo = "": Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
    Combo1.Enabled = True: fpText1.Enabled = True
Case 12
    GrabaRegistro vaSpread1.ActiveRow
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_Zona
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Select Case Col
Case 3
    Dim indice As Long
    vaSpread1.Row = Row
    vaSpread1.Col = 3: indice = vaSpread1.TypeComboBoxCurSel
    vaSpread1.Col = 4: vaSpread1.TypeComboBoxCurSel = indice
    If modo = "" Then modo = "M"
    If Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo
End Select
End Sub

Private Sub MoverDatosGrillas()
Dim i As Long, z As Long, lisnom As String, liscod As String, codaux As Long
'-------> Mover lista precio vector
EstVect = False
Set RS = vg_db.Execute("sgpadm_s_listaprecio 4, 0, 0, '" & vg_NUsr & "'")
i = 1
If Not RS.EOF Then
   ReDim vLisPre(RS!nReg, 2)
   Do While Not RS.EOF
      EstVect = True
      vLisPre(i, 1) = RS!lpr_codigo
      vLisPre(i, 2) = RS!lpr_nombre
      i = i + 1
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
'------> Mover zona grilla
vaSpread1.MaxRows = 0
Set RS2 = vg_db.Execute("sgpadm_s_zona 5, 0,''")
Do While Not RS2.EOF
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = 1: vaSpread1.TypeHAlign = 1: vaSpread1.text = RS2!zon_codigo
    vaSpread1.Col = 2: vaSpread1.text = Trim(RS2!Zon_nombre)
    lisnom = "": liscod = ""
    '-------> Mover lista precio
    If EstVect Then
        For z = 1 To UBound(vLisPre)
            vaSpread1.Col = 3: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vLisPre(z, 2))
            vaSpread1.Col = 4: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vLisPre(z, 1)
            vaSpread1.Col = 3: vaSpread1.TypeComboBoxList = lisnom
            vaSpread1.Col = 4: vaSpread1.TypeComboBoxList = liscod
        Next z
    End If
    vaSpread1.Col = 4
    codaux = -1
    For z = 0 To vaSpread1.TypeComboBoxCount
        vaSpread1.TypeComboBoxCurSel = z
        If vaSpread1.text = IIf(IsNull(RS2!zon_codlpr), 0, RS2!zon_codlpr) Then codaux = z: Exit For
        codaux = -1
    Next z
    vaSpread1.Col = 3: vaSpread1.TypeComboBoxCurSel = codaux
    RS2.MoveNext
Loop
RS2.Close: Set RS2 = Nothing
Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
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
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    GrabaRegistro Row
ElseIf Toolbar1.Buttons(12).Visible = False Then
    Cancela
End If
End Sub

Private Sub Cancela()
Dim z As Long, codaux As Long
OpGr = True
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: Codigo = Val(vaSpread1.Value)
Set RS1 = vg_db.Execute("sgpadm_s_zona 1, " & Codigo & ",''")
If Not RS1.EOF Then
   vaSpread1.Col = 2: vaSpread1.text = Trim(RS1!Zon_nombre)
    vaSpread1.Col = 4
    codaux = -1
    For z = 0 To vaSpread1.TypeComboBoxCount
        vaSpread1.TypeComboBoxCurSel = z
        If vaSpread1.text = IIf(IsNull(RS1!zon_codlpr), 0, RS1!zon_codlpr) Then codaux = z: Exit For
        codaux = -1
    Next z
    vaSpread1.Col = 3: vaSpread1.TypeComboBoxCurSel = codaux
End If
RS1.Close: Set RS1 = Nothing
OpGr = False
End Sub

