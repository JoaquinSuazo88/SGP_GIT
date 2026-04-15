VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form T_Depart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departamento"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   971
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "T_Depart.frx":0000
         Left            =   1950
         List            =   "T_Depart.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2500
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1950
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
         Left            =   450
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
         Left            =   450
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
         Left            =   4530
         TabIndex        =   3
         Top             =   645
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3345
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   8325
      _Version        =   393216
      _ExtentX        =   14684
      _ExtentY        =   5900
      _StockProps     =   64
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
      MaxCols         =   4
      SpreadDesigner  =   "T_Depart.frx":001E
      ScrollBarTrack  =   3
   End
   Begin VB.Label Label3 
      Caption         =   "Ejemplo: MAT-A, MAT-B"
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
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   5160
      Width           =   7455
   End
   Begin VB.Label Label3 
      Caption         =   "(*) Utilice comas (,) para identificar más de un código de homologación."
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
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   4920
      Width           =   7455
   End
End
Attribute VB_Name = "T_Depart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim modo As String
Dim OpGr As Boolean

Private Sub GrabaRegistro(Fila As Long)
Dim Nombre As String, othervalue As String, codigo As Long, estado As String, indlec As Integer
On Error GoTo Man_Error
indlec = 0
OpGr = True
vaSpread1.Row = Fila
vaSpread1.Col = 1: codigo = Val(vaSpread1.text)
vaSpread1.Col = 2: Nombre = Trim(LimpiaDato(vaSpread1.text))
vaSpread1.Col = 3: othervalue = Trim(LimpiaDato(vaSpread1.text))
vaSpread1.Col = 4: estado = IIf(vaSpread1.text = "1", "1", "0")
If Trim(Nombre) = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub
If modo = "A" Then
    vg_db.BeginTrans
    RS.Open "SELECT dep_codigo FROM a_departamento ORDER BY dep_codigo DESC", vg_db, adOpenStatic
    indlec = 1
    If Not RS.EOF Then RS.MoveFirst: codigo = RS!dep_codigo + 1 Else codigo = 1
    RS.Close: Set RS = Nothing
    indlec = 0
    vg_db.Execute "INSERT INTO a_departamento (dep_codigo, dep_nombre, dep_othervalue, dep_estado) VALUES (" & codigo & ", '" & Nombre & "', '" & othervalue & "', '" & estado & "')"
    vg_db.CommitTrans
    vaSpread1.Col = 1: vaSpread1.text = codigo
Else
    vg_db.BeginTrans
    vg_db.Execute "UPDATE a_departamento SET dep_nombre='" & Nombre & "', dep_othervalue='" & othervalue & "', dep_estado='" & estado & "' WHERE dep_codigo=" & codigo
    vg_db.CommitTrans
End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Combo1.Enabled = True: fpText1.Enabled = True
modo = "": Gl_Ac_Botones Me, 1, 1, modo
OpGr = False
Exit Sub
Man_Error:
If indlec = 1 Then RS.Close: Set RS = Nothing
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
Me.Height = 6105
Me.Width = 8625
Msgtitulo = "Departamento"
fg_centra Me
modo = "": OpGr = True
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1
MoverDatosGrillas
OpGr = False
End Sub

Private Sub Form_Resize()
If Me.WindowState = 0 Then
   Frame1.Move 810, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 2200 '1440
ElseIf Me.WindowState = 2 Then
   Frame1.Move 4200, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
End If
Toolbar1.Refresh
End Sub

Private Sub fpText1_Change()
If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    RS.Open "SELECT dep_codigo, dep_nombre, dep_othervalue, dep_estado FROM a_departamento WHERE dep_codigo LIKE '%" & UCase(LimpiaDato(fpText1.text)) & "%'", vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    RS.Open "SELECT dep_codigo, dep_nombre, dep_othervalue, dep_estado FROM a_departamento WHERE UCASE(dep_nombre) LIKE '%" & UCase(LimpiaDato(fpText1.text)) & "%'", vg_db, adOpenStatic
End If
vaSpread1.MaxRows = RS.RecordCount
i = 1
If Not RS.EOF Then
    Do While Not RS.EOF
       vaSpread1.Row = i
       vaSpread1.Col = 1
       vaSpread1.TypeHAlign = TypeHAlignRight
       vaSpread1.text = RS!dep_codigo
       vaSpread1.Col = 2: vaSpread1.text = IIf(IsNull(RS!dep_nombre), "", Trim(RS!dep_nombre))
       vaSpread1.Col = 3: vaSpread1.text = IIf(IsNull(RS!dep_othervalue), "", Trim(RS!dep_othervalue))
       vaSpread1.Col = 4: vaSpread1.text = IIf(IsNull(RS!dep_estado), "0", IIf(RS!dep_estado = "1", "1", "0"))
       RS.MoveNext: i = i + 1
    Loop
    Gl_Ac_Botones Me, 1, 1, modo
End If
RS.Close: Set RS = Nothing
If fpText1.text = "" Then
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Else
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codigo As Long, Nombre As String, NomCor As String
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus
Case 3
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
Case 5
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
    vg_db.BeginTrans
    vg_db.Execute "DELETE a_departamento FROM a_departamento WHERE dep_codigo=" & codigo
    vg_db.CommitTrans
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
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
    Combo1.Enabled = True: fpText1.Enabled = True
Case 12
    GrabaRegistro vaSpread1.ActiveRow
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_Departamento
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

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If OpGr Or Toolbar1.Buttons(12).Visible = True Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub MoverDatosGrillas()
vaSpread1.MaxRows = 0
RS.Open "SELECT * FROM a_departamento ORDER BY dep_codigo", vg_db, adOpenStatic
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1: vaSpread1.text = RS!dep_codigo
   vaSpread1.Col = 2: vaSpread1.text = IIf(IsNull(RS!dep_nombre), "", Trim(RS!dep_nombre))
   vaSpread1.Col = 3: vaSpread1.text = IIf(IsNull(RS!dep_othervalue), "", Trim(RS!dep_othervalue))
   vaSpread1.Col = 4: vaSpread1.text = IIf(IsNull(RS!dep_estado), "0", IIf(RS!dep_estado = "1", "1", "0"))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
Gl_Ac_Botones Me, 1, 1, modo
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
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
RS.Open "SELECT * FROM a_departamento WHERE dep_codigo=" & codigo, vg_db, adOpenStatic
If Not RS.EOF Then
   vaSpread1.Col = 2: vaSpread1.text = IIf(IsNull(RS!dep_nombre), "", Trim(RS!dep_nombre))
   vaSpread1.Col = 3: vaSpread1.text = IIf(IsNull(RS!dep_othervalue), "", Trim(RS!dep_othervalue))
   vaSpread1.Col = 4: vaSpread1.text = IIf(IsNull(RS!dep_estado), "0", IIf(RS!dep_estado = "1", "1", "0"))
End If
RS.Close: Set RS = Nothing
OpGr = False
End Sub

