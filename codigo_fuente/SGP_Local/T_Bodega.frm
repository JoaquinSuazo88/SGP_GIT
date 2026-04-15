VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form T_Bodega 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bodegas"
   ClientHeight    =   4845
   ClientLeft      =   1890
   ClientTop       =   2475
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
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
         ItemData        =   "T_Bodega.frx":0000
         Left            =   1950
         List            =   "T_Bodega.frx":000A
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
         TabIndex        =   5
         Top             =   645
         Width           =   585
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
         TabIndex        =   3
         Top             =   345
         Width           =   1380
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8325
      _ExtentX        =   14684
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
      Width           =   8115
      _Version        =   393216
      _ExtentX        =   14314
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
      MaxCols         =   3
      SpreadDesigner  =   "T_Bodega.frx":001E
      ScrollBarTrack  =   3
   End
End
Attribute VB_Name = "T_Bodega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String
Dim OpGr As Boolean
Dim indlec As Integer
Dim ibusca As Long

Private Sub GrabaRegistro(Fila As Long)
On Error GoTo Man_Error
Dim RS1 As New ADODB.Recordset
Dim codigo As Long
indlec = 0
OpGr = True
vaSpread1.Row = Fila
vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
vaSpread1.Col = 2: Nombre = Trim(LimpiaDato(vaSpread1.Value))
vaSpread1.Col = 3: NomCor = Trim(LimpiaDato(vaSpread1.Value))
If Trim(Nombre) = "" Or Trim(NomCor) = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub
If modo = "A" Then
    vg_db.BeginTrans
    RS1.Open RutinaLectura.Bodega(2, 0, ""), vg_db, adOpenStatic
    indlec = 1
    If Not RS1.EOF Then RS1.MoveFirst: codigo = RS1!bod_codigo + 1 Else codigo = 1
    RS1.Close: Set RS1 = Nothing
    indlec = 0
    vg_db.Execute "INSERT INTO a_bodega (bod_codigo, bod_nombre, bod_ubicac) VALUES (" & codigo & ", '" & Trim(Nombre) & "', '" & Trim(NomCor) & "')"
    vg_db.CommitTrans
    vaSpread1.Col = 1: vaSpread1.Value = codigo
Else
    vg_db.BeginTrans
    vg_db.Execute "UPDATE a_bodega SET bod_nombre='" & Trim(Nombre) & "', bod_ubicac='" & Trim(NomCor) & "' WHERE bod_codigo=" & codigo
    vg_db.CommitTrans
End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Combo1.Enabled = True: fpText1.Enabled = True
modo = "": Gl_Ac_Botones Me, 1, 1, modo
OpGr = False
Exit Sub
Man_Error:
If indlec = 1 Then RS1.Close: Set RS1 = Nothing
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 5355
Me.Width = 8445
Msgtitulo = "Bodegas"
fg_centra Me
modo = "": ibusca = 0: indlec = 0
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1
MoverDatosGrillas
OpGr = False
End Sub

Private Sub Form_Resize()
If Me.WindowState = 0 Then
   Frame1.Move 810, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
ElseIf Me.WindowState = 2 Then
   Frame1.Move 4200, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
End If
Toolbar1.Refresh
End Sub

Private Sub fpText1_Change()
Dim RS2 As New ADODB.Recordset
If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    RS2.Open RutinaLectura.Bodega(4, 0, UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    RS2.Open RutinaLectura.Bodega(5, 0, UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
End If
With vaSpread1
    If ibusca <> RS2.RecordCount Then ibusca = RS2.RecordCount: .MaxRows = RS2.RecordCount
    i = 1
    If Not RS2.EOF Then
        Do While Not RS2.EOF
           .Row = i
           MoverGrilla RS2!bod_codigo, RS2!bod_nombre, RS2!bod_ubicac, .Row
           RS2.MoveNext: i = i + 1
        Loop
        Gl_Ac_Botones Me, 1, 1, modo
    End If
    RS2.Close: Set RS2 = Nothing
    If fpText1.text = "" Then
       Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registro"
    Else
       Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
    End If
End With
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
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
    RS1.Open "SELECT DISTINCT a.* FROM a_bodega a, b_clientes b WHERE a.bod_codigo=b.cli_codbod AND a.bod_codigo=" & codigo & "", vg_db, adOpenStatic
    If Not RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
    RS1.Close: Set RS1 = Nothing
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
    vg_db.BeginTrans
    vg_db.Execute "DELETE a_bodega FROM a_bodega WHERE bod_codigo=" & codigo
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
    I_Bodega
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
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub MoverDatosGrillas()
Dim RS2 As New ADODB.Recordset
With vaSpread1
    .MaxRows = 0
    RS2.Open RutinaLectura.Bodega(3, 0, ""), vg_db, adOpenStatic
    Do While Not RS2.EOF
       grdAddRow vaSpread1
       .Row = .MaxRows
       
       MoverGrilla RS2!bod_codigo, RS2!bod_nombre, RS2!bod_ubicac, .Row

       RS2.MoveNext
    Loop
    RS2.Close: Set RS2 = Nothing
    Gl_Ac_Botones Me, 1, 1, modo
    Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registro"
End With
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
   GrabaRegistro Row
ElseIf Toolbar1.Buttons(12).Visible = False Then
   Cancela
End If
End Sub

Private Sub Cancela()
Dim RS1 As New ADODB.Recordset
Dim codigo As Long
OpGr = True
With vaSpread1
    .Row = .ActiveRow
    .Col = 1: codigo = Val(.Value)
    RS1.Open RutinaLectura.Bodega(3, codigo, ""), vg_db, adOpenStatic
    If Not RS1.EOF Then
       MoverGrilla Val(codigo), RS1!bod_nombre, RS1!bod_ubicac, .Row
    End If
    RS1.Close: Set RS1 = Nothing
    OpGr = False
End With
End Sub

Private Sub MoverGrilla(codbod As Long, nombod As String, ubibod As String, Row As Long)

grdCellTypeStatic vaSpread1, 1, Row, 1
grdSetText vaSpread1, 1, Row, IIf(IsNull(codbod), 1, codbod)

grdCellTypeEdit vaSpread1, 2, Row, 0, 25
grdSetText vaSpread1, 2, Row, IIf(IsNull(nombod), "", Trim(nombod))
       
grdCellTypeEdit vaSpread1, 3, Row, 0, 35
grdSetText vaSpread1, 3, Row, IIf(IsNull(ubibod), "", Trim(ubibod))

End Sub
