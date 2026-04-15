VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_Personas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personas"
   ClientHeight    =   5445
   ClientLeft      =   2415
   ClientTop       =   2310
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   14085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6645
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2335
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   971
      Left            =   3600
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "M_Personas.frx":0000
         Left            =   2010
         List            =   "M_Personas.frx":000A
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
         Left            =   4590
         TabIndex        =   3
         Top             =   645
         Width           =   585
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3885
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   13845
      _Version        =   393216
      _ExtentX        =   24421
      _ExtentY        =   6853
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      AutoCalc        =   0   'False
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
      FormulaSync     =   0   'False
      MaxCols         =   5
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "M_Personas.frx":001E
      ScrollBarTrack  =   3
      ClipboardOptions=   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_Personas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim OpGr As Boolean
Dim ibusca As Long, itop As Long, iRow As Long
Dim rutexi As Boolean

Private Sub GrabaRegistro(Fila As Long)
On Error GoTo Man_Error
Dim RS As New ADODB.Recordset
Dim codigo As String
Dim Nombre As String
Dim rut As String
Dim CodBar As String
OpGr = True
If Command1.Visible = True Then Command1.Visible = False
With vaSpread1
    .Row = Fila
    .Col = 1: rut = fg_DespintaRut(.Value)
    .Col = 2: Nombre = Trim(LimpiaDato(.Value))
    .Col = 3: codigo = fg_DespintaRut(.Value)
    .Col = 5: CodBar = Trim(LimpiaDato(.Value))
    If Trim(rut) = "" Then MsgBox "Favor ingresar rut, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, Msgtitulo: .Row = Fila: .Col = 1: .SetActiveCell 1, .Row: .SetFocus: OpGr = False: Exit Sub
    If Trim(Nombre) = "" Then MsgBox "Favor ingresar nombre, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, Msgtitulo: .Row = Fila: .Col = 2: .SetActiveCell 2, .Row: .SetFocus: OpGr = False: Exit Sub
    If Trim(codigo) = "" Then MsgBox "Favor ingresar Rut cliente, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, Msgtitulo: .Row = Fila: .Col = 3: .SetActiveCell 3, .Row: .SetFocus: OpGr = False: Exit Sub
    If Trim(CodBar) <> "" Then
       RS.Open RutinaLectura.Personal(7, Trim(LimpiaDato(CodBar)), ""), vg_db, adOpenStatic
       If Not RS.EOF Then
          vaSpread1.Col = 1
            If rut <> RS!per_rut Then RS.Close: Set RS = Nothing: MsgBox "Ya existe Código Barra...", vbExclamation + vbOKOnly, Msgtitulo:  .Row = Fila: .Col = 5: .SetActiveCell 5, .Row: .SetFocus: OpGr = False: Exit Sub
        End If
        RS.Close: Set RS = Nothing
    End If
    If modo = "A" Then
        'Validar que no exista rut
        RS.Open RutinaLectura.Personal(2, rut, ""), vg_db, adOpenStatic
        If Not RS.EOF Then
           RS.Close: Set RS = Nothing
           MsgBox "Rut ya esta informado...", vbExclamation + vbOKOnly, Msgtitulo: .Row = Fila: .Col = 2: .SetActiveCell 2, .Row: .SetFocus: OpGr = False: Exit Sub
        End If
        RS.Close: Set RS = Nothing
        
        vg_db.BeginTrans
        vg_db.Execute "INSERT INTO b_persona (per_rut, per_nombre, cli_codigo, per_codbarra) VALUES ('" & rut & "',  '" & Trim(Nombre) & "', '" & codigo & "', '" & CodBar & "')"
        vg_db.CommitTrans

'        .Col = 1: .Value = codigo
    Else
        vg_db.BeginTrans
        vg_db.Execute "UPDATE b_persona SET per_nombre = '" & Trim(Nombre) & "', cli_codigo = '" & codigo & "', per_codbarra = '" & CodBar & "' WHERE per_rut='" & rut & "'"
        vg_db.CommitTrans

    End If
    Frame1.Enabled = True
    Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registro"
End With
rutexi = False
'Command1.Visible = True
modo = "": Gl_Ac_Botones Me, 1, 1, modo
OpGr = False
If modo = "A" Then
        MsgBox "Registro guardo exitosamente", vbInformation + vbOKOnly, Msgtitulo
Else
        MsgBox "Registro modificado exitosamente", vbInformation + vbOKOnly, Msgtitulo
End If
Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
If Err.Number = -2147467259 Then
    MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"
Else
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If
End Sub

Private Sub Command1_Click()
vg_left = Command1.Left + 3801
vg_nombre = "": vg_codigo = ""
B_TabEst.LlenaDatos "b_clientes", "cli_", "Cliente", "Cliente"
B_TabEst.Show 1
Me.Refresh
With vaSpread1
    If vg_codigo = "" Then .Col = 3: .Row = iRow: .SetActiveCell 4, iRow: .EditMode = True: .EditModeReplace = True: .SetFocus: Exit Sub
    .Row = iRow
    .Col = 3
    .Value = vg_codigo
    .Col = 4
    .Value = vg_nombre
    .Col = 5
    .EditMode = True
    .EditModeReplace = True
    .SetActiveCell 5, iRow: .SetFocus
End With
If modo <> "A" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 5955
Me.Width = 14205
Msgtitulo = "Personal"
fg_centra Me
modo = "": ibusca = 0: itop = 1
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1
MoverDatosGrillas
OpGr = False
rutexi = False
End Sub

Private Sub Form_Resize()
'If Me.WindowState = 0 Then
'   Frame1.Move 3840, 360, 6015, 971
'   vaSpread1.Move 120, 1440, ScaleWidth, ScaleHeight - 1440
'ElseIf Me.WindowState = 2 Then
'   Frame1.Move 5200, 360, 6015, 971
'   vaSpread1.Move 120, 1440, ScaleWidth, ScaleHeight - 1440
'End If
'Toolbar1.Refresh
End Sub

Private Sub fpText1_Change()
Dim RS As New ADODB.Recordset
Dim opusu  As Boolean
If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
opusu = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", falso, True)
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    RS.Open RutinaLectura.Personal(3, 0, UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    RS.Open RutinaLectura.Personal(4, 0, UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
End If
With vaSpread1
    ibusca = RS.RecordCount: .MaxRows = RS.RecordCount
    i = 1
    If Not RS.EOF Then
       Do While Not RS.EOF
          .Row = i
          i = i + 1
          .Col = 1: .Lock = True: .text = fg_PintaRut(RS!per_rut)
          .Col = 2: .Lock = opusu: .text = IIf(IsNull(RS!per_nombre), "", Trim(RS!per_nombre))
          .Col = 3: .Lock = opusu: .text = IIf(IsNull(RS!cli_codigo), "", fg_PintaRut(RS!cli_codigo))
          .Col = 4: .Lock = opusu: .text = IIf(IsNull(RS!cli_nombre), "", Trim(RS!cli_nombre))
          .Col = 5: .Lock = opusu: .text = IIf(IsNull(RS!per_codbarra), "", Trim(RS!per_codbarra))
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
Dim codigo As String, Nombre As String, rut As String, CodBar As String
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 1: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus
    Command1.Visible = False
    Frame1.Enabled = False
    rutexi = False
Case 3
    If vaSpread1.MaxRows < 1 Then Exit Sub
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    Frame1.Enabled = False
    rutexi = False
'    Command1.Visible = False
Case 5
    rutexi = False
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: rut = fg_DespintaRut(vaSpread1.Value)
    vg_db.Execute "DELETE b_persona FROM b_persona WHERE per_rut = '" & rut & "'"
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    MsgBox "Registro eliminado exitosamente", vbInformation + vbOKOnly, Msgtitulo
Case 7
    rutexi = False
    fpText1.text = ""
    MoverDatosGrillas
Case 10
    rutexi = False
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    Frame1.Enabled = True
    If modo = "A" Then
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    Else
        Cancela
    End If
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
'    Command1.Visible = True
    Frame1.Enabled = True
Case 12
    GrabaRegistro vaSpread1.ActiveRow
    rutexi = False
Case 15
    rutexi = False
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_Personas
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err.Number = -2147467259 Then
    MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"
Else
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows > 1 Then Exit Sub
vaSpread1.Row = Row
vaSpread1.Col = Col
If Col <> 3 Then Command1.Visible = False
Select Case Col
Case 1, 3 'ingreso rut
    If InStr(vaSpread1.text, "-") = 0 Or Trim(vaSpread1.text) = "" Or (vaSpread1.ActiveCol = 2 Or vaSpread1.ActiveCol = 4 Or vaSpread1.ActiveCol = 5) Or vaSpread1.Lock = True Then Exit Sub
    If Trim(vaSpread1.text) = "" Or vg_Dig = "N" Then Exit Sub
    vaSpread1.text = fg_DespintaRut(vaSpread1.text)
    vaSpread1.text = Mid(vaSpread1.text, 1, Len(Trim(vaSpread1.text)) - 1)
    If Col = 3 Then
       Command1.Top = IIf(Row = 1, 2335, 2335 + (240 * (Row - itop)))
       Command1.Visible = True
       vaSpread1.EditMode = True
       vaSpread1.EditModeReplace = True
       vaSpread1.Row = Row
       iRow = Row
       vaSpread1.Col = 4
       vaSpread1.TypeHAlign = TypeHAlignLeft
    End If
End Select
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub MoverDatosGrillas()
Dim RS As New ADODB.Recordset
Dim opusu As Boolean
Command1.Visible = False
With vaSpread1
    opusu = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", falso, True)
    .MaxRows = 0
    RS.Open RutinaLectura.Personal(5, "", ""), vg_db, adOpenStatic
    Do While Not RS.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1: .Lock = True: .text = fg_PintaRut(RS!per_rut)
        .Col = 2: .Lock = opusu: .Value = IIf(IsNull(RS!per_nombre), "", Trim(RS!per_nombre))
        .Col = 3: .Lock = opusu: .Value = IIf(IsNull(RS!cli_codigo), "", fg_PintaRut(RS!cli_codigo))
        .Col = 4: .Lock = opusu: .Value = IIf(IsNull(RS!cli_nombre), "", Trim(RS!cli_nombre))
        .Col = 5: .Lock = opusu: .Value = IIf(IsNull(RS!per_codbarra), "", Trim(RS!per_codbarra))
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    Gl_Ac_Botones Me, 1, 1, modo
    Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registro"
End With
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
Dim RS As New ADODB.Recordset
Dim rut As String
Dim posinicial As Long
Dim largo As Long
Dim atributo As String
Dim cadena As String
iRow = Row
posinicial = 0
largo = 0
Command1.Top = IIf(Row = 1, 2335, 2335 + (240 * (Row - itop)))
Command1.Visible = True
If ChangeMade = False And Col <> 4 Then
   If Col <> 3 Then Command1.Visible = False
   Exit Sub
End If
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
Frame1.Enabled = False
Select Case Col
Case Is <> 3
    Command1.Visible = False
    If Col = 5 Then '-------> Validar código barra
'        '-------> Traer parametro código barra
'        RS.Open RutinaLectura.ParametroCodBarra(1, MuestraCasino(1), "'Rut comensal'"), vg_db, adOpenStatic
'        If RS.EOF Then
'           RS.Close: Set RS = Nothing
''           MsgBox "No existe parametrización código barra. Proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Col = Col: vaSpread1.text = "": vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus: Exit Sub
'        Else
'           If IsNull(RS!cbar_posinicial) Or RS!cbar_posinicial < 1 Then RS.Close: Set RS = Nothing: MsgBox "Posicion Incial del atributo código barra, esta con valor cero o nulo. Proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Col = Col: vaSpread1.text = "": vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus: Exit Sub
'           If IsNull(RS!cbar_largo) Or RS!cbar_largo < 1 Then RS.Close: Set RS = Nothing: MsgBox "Largo del atributo código barra, esta en valor cero o nulo. Proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Col = Col: vaSpread1.text = "": vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus: Exit Sub
'           posinicial = IIf(IsNull(RS!cbar_posinicial), 0, RS!cbar_posinicial)
'           largo = IIf(IsNull(RS!cbar_largo), 0, RS!cbar_largo)
'           RS.Close: Set RS = Nothing
'           '-------> Validar que valor contenido tenga valor
'           vaSpread1.Row = Row
'           vaSpread1.Col = Col
'           If Trim(Mid(vaSpread1.text, posinicial, largo)) = "" Then MsgBox "La posición inicial y largo del atributo código barra, , da como resultado valor cero o nulo. Proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Col = Col: vaSpread1.text = "": vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus: Exit Sub
'           cadena = Mid(vaSpread1.text, posinicial, largo)
'           vaSpread1.text = cadena 'Mid(vaSpread1.text, posinicial, largo)
'        End If
           vaSpread1.Row = Row
           vaSpread1.Col = Col
        If Trim(vaSpread1.text) <> "" Then
        RS.Open RutinaLectura.Personal(7, Trim(LimpiaDato(vaSpread1.text)), ""), vg_db, adOpenStatic
        If Not RS.EOF Then
           vaSpread1.Col = 1
           rut = fg_DespintaRut(Trim(vaSpread1.text))
'           If rut <> RS!per_rut Then MsgBox "Ya existe Código Barra...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Col = Col: vaSpread1.text = "": vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus: Exit Sub
           If rut <> RS!per_rut Then RS.Close: Set RS = Nothing: MsgBox "Ya existe Código Barra...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Col = Col: vaSpread1.text = "": vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus: Exit Sub
        End If
        RS.Close: Set RS = Nothing
        End If
'        vaSpread1.text = IIf(KeyCode = 8, vaSpread1.text, vaSpread1.text & Chr(KeyCode)): vaSpread1.SetFocus: vaSpread1.SelStart = Len(vaSpread1.text)
'        vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus: Exit Sub
    End If
Case 3
    Command1.Top = IIf(Row = 1, 2335, 2335 + (240 * (Row - itop)))
    Command1.Visible = True
    vaSpread1.Row = Row
    vaSpread1.Col = Col
    vaSpread1.text = fg_RutDig(Trim(vaSpread1.text))
    vaSpread1.text = fg_PintaRut(vaSpread1.text)
    'Validar que no exista rut
    rut = fg_DespintaRut(Trim(vaSpread1.text))
    RS.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(rut)) & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1.text = "": vaSpread1.Col = 4: vaSpread1.text = "": Exit Sub
    vaSpread1.Col = 4: vaSpread1.text = Trim(RS!cli_nombre)
    RS.Close: Set RS = Nothing
    Command1.Visible = False
    vaSpread1.Col = 5
    vaSpread1.EditMode = True
    vaSpread1.EditModeReplace = True
    vaSpread1.SetActiveCell 5, Row: vaSpread1.SetFocus
End Select
End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RS As New ADODB.Recordset
Dim rut As String
If vaSpread1.MaxRows < 1 Then Exit Sub
Select Case KeyCode
Case 39, 40
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If InStr(vaSpread1.text, "-") = 0 Or Trim(vaSpread1.text) = "" Or (vaSpread1.ActiveCol = 2 Or vaSpread1.ActiveCol = 4 Or vaSpread1.ActiveCol = 5) Or vaSpread1.Lock = True Then Exit Sub
    vaSpread1.text = fg_DespintaRut(vaSpread1.text)
    vaSpread1.text = Mid(vaSpread1.text, 1, Len(Trim(vaSpread1.text)) - 1)
'Case 46
'    a = a
Case 86
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
End Select
End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If vaSpread1.ActiveCol = 5 Then
    SendKeys "{Tab}"
'    SendKeys "+{Tab}"
End If
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim rut As String
Dim RS As New ADODB.Recordset
If rutexi Then Exit Sub
If (Row <> NewRow And NewRow > 0) Or (Col <> NewCol And NewCol > 0) Or Col <> 3 Then Command1.Visible = False
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
   GrabaRegistro Row
ElseIf Toolbar1.Buttons(12).Visible = False Then
   Cancela
ElseIf vaSpread1.MaxRows > 0 Then
    vaSpread1.Row = Row
    vaSpread1.Col = Col
    If Trim(vaSpread1.text) = "" Or (vaSpread1.ActiveCol = 2 Or vaSpread1.ActiveCol = 4 Or vaSpread1.ActiveCol = 5) Or vaSpread1.Lock = True Then Exit Sub
'    MsgBox "1"
    vaSpread1.text = fg_RutDig(Trim(vaSpread1.text))
    vaSpread1.text = fg_PintaRut(vaSpread1.text)
    rut = fg_DespintaRut(Trim(vaSpread1.text))
    '-------> Validar largo del campo rut
    If Len(rut) < 6 Or Len(rut) > 10 Then
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = 1
       vaSpread1.text = ""
       rut = ""
       MsgBox "Largo del rut debe estar entre mínimo 6 caracteres y Max 10...", vbExclamation + vbOKOnly, Msgtitulo
       vaSpread1.SetActiveCell 1, vaSpread1.Row
'       vaSpread1.SetFocus
       OpGr = False
'       Exit Sub
    Else
       'Validar que no exista rut
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = Col
       vaSpread1.text = fg_PintaRut(rut)
       
       RS.Open RutinaLectura.Personal(5, LimpiaDato(Trim(rut)), ""), vg_db, adOpenStatic
       If Not RS.EOF Then
          rutexi = True
           vaSpread1.Row = vaSpread1.ActiveRow
          vaSpread1.Col = 1
          vaSpread1.text = fg_PintaRut(rut)
          MsgBox "Rut personal ya esta informado...", vbExclamation + vbOKOnly, Msgtitulo
          rutexi = False
          vaSpread1.Row = vaSpread1.ActiveRow
          vaSpread1.Col = 1
          vaSpread1.text = ""
          vaSpread1.SetActiveCell 1, vaSpread1.Row
 '         vaSpread1.SetFocus
          OpGr = False
'          Exit Sub
         RS.Close: Set RS = Nothing
       Else
          RS.Close: Set RS = Nothing
       End If
       vaSpread1.SetActiveCell 2, vaSpread1.Row
       rutexi = False
'       vaSpread1.SetFocus
    End If
End If
End Sub

Private Sub vaSpread1_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
itop = NewTop
Command1.Visible = False
End Sub

Private Sub Cancela()
Dim RS As New ADODB.Recordset
Dim rut As String
OpGr = True
With vaSpread1
    .Row = .ActiveRow
    .Col = 1: rut = fg_DespintaRut(Trim(.text))
    RS.Open RutinaLectura.Personal(5, rut, ""), vg_db, adOpenStatic
    If Not RS.EOF Then
       .Col = 2: .Value = IIf(IsNull(RS!per_nombre), "", Trim(RS!per_nombre))
       .Col = 3: .Value = IIf(IsNull(RS!cli_codigo), "", fg_PintaRut(RS!cli_codigo))
       .Col = 4: .Value = IIf(IsNull(RS!cli_nombre), "", Trim(RS!cli_nombre))
       .Col = 5: .Value = IIf(IsNull(RS!per_codbarra), "", Trim(RS!per_codbarra))
    End If
    RS.Close: Set RS = Nothing
End With
OpGr = False
End Sub

