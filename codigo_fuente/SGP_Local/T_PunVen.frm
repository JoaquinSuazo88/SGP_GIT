VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form T_PunVen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punto Atención"
   ClientHeight    =   5010
   ClientLeft      =   5700
   ClientTop       =   2190
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   971
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "T_PunVen.frx":0000
         Left            =   2010
         List            =   "T_PunVen.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2500
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   2010
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
         Left            =   4590
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
         Left            =   525
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
         Left            =   525
         TabIndex        =   3
         Top             =   345
         Width           =   1380
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3405
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   6030
      _Version        =   393216
      _ExtentX        =   10636
      _ExtentY        =   6006
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      AutoCalc        =   0   'False
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
      FormulaSync     =   0   'False
      MaxCols         =   2
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "T_PunVen.frx":001E
      ScrollBarTrack  =   3
      ClipboardOptions=   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "T_PunVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim OpGr As Boolean
Dim ibusca As Long

Private Sub GrabaRegistro(Fila As Long)
On Error GoTo Man_Error
Dim RS As New ADODB.Recordset
Dim opusu As Boolean
Dim codigo As Long
Dim Nombre As String
Dim i As Long
OpGr = True
With vaSpread1
    .Row = Fila
    .Col = 1: codigo = Val(.Value)
    .Col = 2: Nombre = Trim(LimpiaDato(.Value))
    For i = 1 To .MaxRows
        .Col = 2: .Row = i
        If UCase(Trim(.text)) = UCase(Trim(Nombre)) And Fila <> i Then MsgBox "Descripción ya existe en la grilla...", vbExclamation + vbOKOnly, Msgtitulo: .Row = Fila: .Col = 2: .SetActiveCell 2, .Row: .SetFocus: OpGr = False: Exit Sub
    Next i
    If Trim(Nombre) = "" Then MsgBox "Favor ingresar descripción, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, Msgtitulo: .Row = Fila: .Col = 2: .SetActiveCell 2, .Row: .SetFocus: OpGr = False: Exit Sub
    If modo = "A" Then
        RS.Open RutinaLectura.PuntoAtencion(1, 0, ""), vg_db, adOpenStatic
        If Not RS.EOF Then RS.MoveFirst: codigo = RS!ate_codatencion + 1 Else codigo = 1
        RS.Close: Set RS = Nothing
        vg_db.Execute "INSERT INTO a_pto_atencion (ate_codatencion, ate_descripcion) VALUES (" & codigo & ", '" & Trim(Nombre) & "')"
        .Col = 1: .Value = codigo

    Else
        vg_db.Execute "UPDATE a_pto_atencion SET ate_descripcion='" & Trim(Nombre) & "' WHERE ate_codatencion=" & codigo

    End If
    Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registro"
    opusu = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", falso, True)
    .Col = 2
    .Lock = opusu
End With

Frame1.Enabled = True
modo = "": Gl_Ac_Botones Me, 1, 1, modo
OpGr = False
If modo = "A" Then
        MsgBox "Registro guardo exitosamente", vbInformation + vbOKOnly, Msgtitulo
Else
        MsgBox "Registro modificado exitosamente", vbInformation + vbOKOnly, Msgtitulo
End If
Exit Sub
Man_Error:
If Err.Number = -2147467259 Then
    MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"
Else
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 5400
Me.Width = 6210
Msgtitulo = "Punto Atención"
fg_centra Me
modo = "": ibusca = 0
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1
MoverDatosGrillas
OpGr = False
End Sub

Private Sub Form_Resize()
If Me.WindowState = 0 Then
   Frame1.Move 0, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
ElseIf Me.WindowState = 2 Then
   Frame1.Move 4200, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
End If
Toolbar1.Refresh
End Sub

Private Sub fpText1_Change()
Dim RS As New ADODB.Recordset
Dim opusu As Boolean
If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
opusu = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", falso, True)
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    RS.Open RutinaLectura.PuntoAtencion(3, 0, UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    RS.Open RutinaLectura.PuntoAtencion(4, 0, UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
End If
With vaSpread1
    ibusca = RS.RecordCount: .MaxRows = RS.RecordCount
    i = 1
    If Not RS.EOF Then
       Do While Not RS.EOF
          .Row = i
          i = i + 1
          .Col = 1: .Value = RS!ate_codatencion
          .Col = 2: .Lock = opusu: .Value = IIf(IsNull(RS!ate_descripcion), "", Trim(RS!ate_descripcion))
          RS.MoveNext
       Loop
       Gl_Ac_Botones Me, 1, 1, modo
    End If
    RS.Close: Set RS = Nothing
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
    Frame1.Enabled = False
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus
Case 3
    If vaSpread1.MaxRows < 1 Then Exit Sub
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    Frame1.Enabled = False
Case 5
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
    vg_db.Execute "DELETE a_pto_atencion FROM a_pto_atencion WHERE ate_codatencion=" & codigo
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    MsgBox "Registro eliminado exitosamente", vbInformation + vbOKOnly, Msgtitulo
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
    Frame1.Enabled = True
Case 12
    GrabaRegistro vaSpread1.ActiveRow
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_PuntoAtencion
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err.Number = -2147467259 Or Err.Number = -2147217900 Then
    MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"
Else
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
Frame1.Enabled = False
End Sub

Private Sub MoverDatosGrillas()
Dim RS As New ADODB.Recordset
Dim opusu As Boolean
With vaSpread1
    .MaxRows = 0
    opusu = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", falso, True)
    RS.Open RutinaLectura.PuntoAtencion(2, 0, ""), vg_db, adOpenStatic
    Do While Not RS.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1: .Value = RS!ate_codatencion
        .Col = 2: .Lock = opusu: .Value = IIf(IsNull(RS!ate_descripcion), "", Trim(RS!ate_descripcion))
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
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
Dim RS As New ADODB.Recordset
Dim codigo As Long
OpGr = True
With vaSpread1
    If .ActiveRow < 1 Then Exit Sub
    .Row = .ActiveRow
    .Col = 1: codigo = Val(.Value)
    RS.Open RutinaLectura.PuntoAtencion(2, codigo, ""), vg_db, adOpenStatic
    If Not RS.EOF Then
       .Row = .ActiveRow
       .Col = 2: .Value = IIf(IsNull(RS!ate_descripcion), "", Trim(RS!ate_descripcion))
    End If
    RS.Close: Set RS = Nothing
End With
OpGr = False
End Sub

