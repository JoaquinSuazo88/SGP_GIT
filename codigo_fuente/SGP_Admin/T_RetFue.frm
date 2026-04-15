VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form T_RetFue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retención en la Fuente"
   ClientHeight    =   5145
   ClientLeft      =   780
   ClientTop       =   1590
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5145
   ScaleWidth      =   11925
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
      Left            =   6455
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   971
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "T_RetFue.frx":0000
         Left            =   1875
         List            =   "T_RetFue.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1875
         TabIndex        =   3
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
         Left            =   345
         TabIndex        =   6
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
         Left            =   345
         TabIndex        =   5
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
         Left            =   4455
         TabIndex        =   4
         Top             =   645
         Width           =   585
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3585
      Left            =   120
      TabIndex        =   0
      Top             =   1350
      Width           =   11715
      _Version        =   393216
      _ExtentX        =   20664
      _ExtentY        =   6324
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
      MaxCols         =   7
      SpreadDesigner  =   "T_RetFue.frx":001E
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "T_RetFue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim OpGr As Boolean
Dim ibusca As Long, IRow As Long, itop As Long

Private Sub GrabaRegistro(Fila As Long)

On Error GoTo Man_Error

Dim RS     As New ADODB.Recordset
Dim codigo As Long
Dim Nombre As String
Dim codcta As String
Dim tipret As String
Dim indret As String
Dim portar As Double

If Command1.Visible = True Then Command1.Visible = False
OpGr = True
vaSpread1.Row = Fila

vaSpread1.Col = 1
codigo = Val(vaSpread1.text)

vaSpread1.Col = 2
Nombre = Trim(LimpiaDato(Mid(vaSpread1.text, 1, 100)))

vaSpread1.Col = 3
portar = Val(vaSpread1.text)

vaSpread1.Col = 4
codcta = Trim(LimpiaDato(Mid(vaSpread1.text, 1, 10)))

vaSpread1.Col = 6
tipret = Trim(LimpiaDato(Mid(vaSpread1.text, 1, 10)))

vaSpread1.Col = 7
indret = Trim(LimpiaDato(Mid(vaSpread1.text, 1, 10)))

If Trim(Nombre) = "" Or portar < 0 Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub

If modo = "A" Then
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient

   codigo = 0
   Set RS = vg_db.Execute("sgpadm_iu_retencionfuente 'A', 0, '" & Trim(Nombre) & "', " & portar & ", '" & codcta & "', '" & tipret & "', '" & indret & "'")
   If Not RS.EOF Then
      
      codigo = RS!indice
      vaSpread1.Col = 1
      vaSpread1.Lock = True
      vaSpread1.text = codigo
   
   End If
   RS.Close
   Set RS = Nothing

Else
    
    vg_db.Execute "sgpadm_iu_retencionfuente 'M', " & codigo & ", '" & Trim(Nombre) & "', " & portar & ", '" & codcta & "', '" & tipret & "', '" & indret & "'"

End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Combo1.Enabled = True
fpText1.Enabled = True
modo = ""
Gl_Ac_Botones Me, 1, 1, modo
OpGr = False
Command1.Visible = False

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
Me.Height = 5625
Me.Width = 12015
MsgTitulo = "Retención en la Fuente"
fg_centra Me
modo = "": ibusca = 0: itop = 1
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1
MoverDatosGrillas
Command1.Visible = False: Command1.Top = 2010
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Resize()

On Error GoTo Man_Error

If Me.WindowState = 0 Then
   
   Frame1.Move 2400, 360, 6015, 971
   vaSpread1.Move 180, 1440, ScaleWidth - 200, ScaleHeight - 1440
   Command1.Left = 6510

ElseIf Me.WindowState = 2 Then
   
   Frame1.Move 4200, 360, 6015, 971
   vaSpread1.Move 120, 1440, ScaleWidth - 200, ScaleHeight - 1440
   Command1.Left = 6455

End If
Toolbar1.Refresh

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpText1_Change()

On Error GoTo Man_Error

If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub

Dim RS As New ADODB.Recordset

vaSpread1.Visible = False
Command1.Visible = False

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Combo1.ItemData(Combo1.ListIndex) = 0 Then
   
   Set RS = vg_db.Execute("sgpadm_s_retencionfuente 3, 0, '%" & UCase(LimpiaDato(fpText1.text)) & "%'")

ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
   
   Set RS = vg_db.Execute("sgpadm_s_retencionfuente 4, 0, '%" & UCase(LimpiaDato(fpText1.text)) & "%'")

End If

If RS.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS!nReg
i = 1

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      vaSpread1.Row = i
      i = i + 1
      
      vaSpread1.Col = 1
      vaSpread1.Lock = True
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = RS!ref_codigo
      
      vaSpread1.Col = 2
      vaSpread1.Lock = False
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = IIf(IsNull(RS!ref_nombre), "", Trim(RS!ref_nombre))
      
      vaSpread1.Col = 3
      vaSpread1.Lock = False
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.text = IIf(IsNull(RS!ref_portar), "", Trim(RS!ref_portar))
      
      vaSpread1.Col = 4
      vaSpread1.Lock = False
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = IIf(IsNull(RS!ref_codcta), "", Trim(RS!ref_codcta))
      
      vaSpread1.Col = 5
      vaSpread1.Lock = True
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = IIf(IsNull(RS!cta_nombre), "", Trim(RS!cta_nombre))
      
      vaSpread1.Col = 6
      vaSpread1.Lock = False
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = IIf(IsNull(RS!ref_tipret), "", Trim(RS!ref_tipret))
      
      vaSpread1.Col = 7
      vaSpread1.Lock = False
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = IIf(IsNull(RS!ref_indret), "", Trim(RS!ref_indret))
      
      RS.MoveNext
   
   Loop
   Gl_Ac_Botones Me, 1, 1, modo

End If
RS.Close
Set RS = Nothing
vaSpread1.Visible = True
If fpText1.text = "" Then Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro" Else Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim codigo As Long, Nombre As String, orden As String

On Error GoTo Man_Error

Select Case Button.Index

Case 1
    
    Command1.Visible = False
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = 2
    vaSpread1.SetActiveCell 2, vaSpread1.MaxRows
    vaSpread1.SetFocus
    
    Command1.Visible = False

Case 3
    
    modo = "M"
    If vaSpread1.MaxRows < 1 Then Exit Sub
    Gl_Ac_Botones Me, 1, 0, modo

Case 5
    
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    codigo = Val(vaSpread1.text)
    vg_db.Execute "DELETE b_retencionfuente FROM b_retencionfuente WHERE ref_codigo = " & codigo & ""
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    modo = ""
    Gl_Ac_Botones Me, 1, 1, modo

Case 7
    
    fpText1.text = ""
    MoverDatosGrillas

Case 10
    
    If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If modo = "A" Then
       
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.DeleteRows vaSpread1.Row, 1
       vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    
    Else
       
       Cancela
    
    End If
    modo = ""
    Gl_Ac_Botones Me, 1, 1, modo
    Combo1.Enabled = True
    fpText1.Enabled = True
    Command1.Visible = False

Case 12
    
    GrabaRegistro vaSpread1.ActiveRow

Case 15
    
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    I_RetencionFuente

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

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Select Case Col

Case Is <> 4
    
    Command1.Visible = False

Case 4
    
    Command1.Top = IIf(Row = 1, 2010, 2010 + (240 * (Row - itop)))
    Command1.Visible = True
    vaSpread1.EditMode = True
    vaSpread1.EditModeReplace = True
    vaSpread1.Row = Row
    IRow = Row
    vaSpread1.Col = 4
    vaSpread1.TypeHAlign = TypeHAlignLeft

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub MoverDatosGrillas()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

itop = 1
Command1.Visible = False
vaSpread1.Visible = False
vaSpread1.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_retencionfuente 2, 0, ''")

Do While Not RS.EOF
    
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    
    vaSpread1.Col = 1
    vaSpread1.Lock = True
    vaSpread1.TypeHAlign = TypeHAlignLeft
    vaSpread1.text = RS!ref_codigo
    
    vaSpread1.Col = 2
    vaSpread1.Lock = False
    vaSpread1.TypeHAlign = TypeHAlignLeft
    vaSpread1.text = IIf(IsNull(RS!ref_nombre), "", Trim(RS!ref_nombre))
    
    vaSpread1.Col = 3
    vaSpread1.Lock = False
    vaSpread1.TypeHAlign = TypeHAlignRight
    vaSpread1.text = IIf(IsNull(RS!ref_portar), "", Trim(RS!ref_portar))
    
    vaSpread1.Col = 4
    vaSpread1.Lock = False
    vaSpread1.TypeHAlign = TypeHAlignLeft
    vaSpread1.text = IIf(IsNull(RS!ref_codcta), "", Trim(RS!ref_codcta))
    
    vaSpread1.Col = 5
    vaSpread1.Lock = True
    vaSpread1.TypeHAlign = TypeHAlignLeft
    vaSpread1.text = IIf(IsNull(RS!cta_nombre), "", Trim(RS!cta_nombre))
    
    vaSpread1.Col = 6
    vaSpread1.Lock = False
    vaSpread1.TypeHAlign = TypeHAlignLeft
    vaSpread1.text = IIf(IsNull(RS!ref_tipret), "", Trim(RS!ref_tipret))
    
    vaSpread1.Col = 7
    vaSpread1.Lock = False
    vaSpread1.TypeHAlign = TypeHAlignLeft
    vaSpread1.text = IIf(IsNull(RS!ref_indret), "", Trim(RS!ref_indret))
    
    RS.MoveNext

Loop
RS.Close
Set RS = Nothing
Gl_Ac_Botones Me, 1, 1, modo
vaSpread1.Visible = True
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

IRow = Row
Command1.Top = IIf(Row = 1, 2010, 2010 + (240 * (Row - itop)))
Command1.Visible = True
If ChangeMade = False And Col <> 6 Then
   
   If Col <> 4 Then Command1.Visible = False
   Exit Sub

End If
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
Select Case Col

Case Is <> 4
    
    Command1.Visible = False

Case 4
    
    Command1.Top = IIf(Row = 1, 2010, 2010 + (240 * (Row - itop)))
    Command1.Visible = True
    vaSpread1.Row = Row
    vaSpread1.Col = Col
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS = vg_db.Execute("SELECT cta_nombre FROM a_ctacontable WHERE cta_codigo = '" & vaSpread1.Value & "'")
    If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1.text = "": vaSpread1.Col = 5: vaSpread1.text = "": Exit Sub
    vaSpread1.Col = 5
    vaSpread1.text = Trim(RS!cta_nombre)
    
    RS.Close
    Set RS = Nothing
    
    Command1.Visible = False

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo Man_Error

If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then

'    Command1.Visible = False
    GrabaRegistro Row

ElseIf Toolbar1.Buttons(12).Visible = False Then

'    Command1.Visible = False
    Cancela

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)

On Error GoTo Man_Error

itop = NewTop
Command1.Visible = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Cancela()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim codigo As Long

OpGr = True
Command1.Visible = False
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = Val(vaSpread1.text)

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_retencionfuente 1, " & codigo & ", ''")
If Not RS.EOF Then
   
   vaSpread1.Col = 2
   vaSpread1.Lock = False
   vaSpread1.TypeHAlign = TypeHAlignLeft
   vaSpread1.text = IIf(IsNull(RS!ref_nombre), "", Trim(RS!ref_nombre))
   
   vaSpread1.Col = 3
   vaSpread1.Lock = False
   vaSpread1.TypeHAlign = TypeHAlignRight
   vaSpread1.text = IIf(IsNull(RS!ref_portar), "", Trim(RS!ref_portar))
   
   vaSpread1.Col = 4
   vaSpread1.Lock = False
   vaSpread1.TypeHAlign = TypeHAlignLeft
   vaSpread1.text = IIf(IsNull(RS!ref_codcta), "", Trim(RS!ref_codcta))
   
   vaSpread1.Col = 5
   vaSpread1.Lock = True
   vaSpread1.TypeHAlign = TypeHAlignLeft
   vaSpread1.text = IIf(IsNull(RS!cta_nombre), "", Trim(RS!cta_nombre))
   
   vaSpread1.Col = 6
   vaSpread1.Lock = False
   vaSpread1.TypeHAlign = TypeHAlignLeft
   vaSpread1.text = IIf(IsNull(RS!ref_tipret), "", Trim(RS!ref_tipret))
   
   vaSpread1.Col = 7
   vaSpread1.Lock = False
   vaSpread1.TypeHAlign = TypeHAlignLeft
   vaSpread1.text = IIf(IsNull(RS!ref_indret), "", Trim(RS!ref_indret))

End If
RS.Close
Set RS = Nothing
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Command1_Click()

On Error GoTo Man_Error

vg_left = Command1.Left + 3801
vg_nombre = ""
vg_codigo = ""
B_TabEst.LlenaDatos "a_ctacontable", "cta_", "Cta. Contable", "Gen"
B_TabEst.Show 1
Me.Refresh

If vg_codigo = "" Then

   vaSpread1.Col = 4
   vaSpread1.Row = IRow
   vaSpread1.SetActiveCell 4, IRow
   vaSpread1.EditMode = True
   vaSpread1.EditModeReplace = True
   vaSpread1.SetFocus
   Exit Sub

End If

vaSpread1.Row = IRow

vaSpread1.Col = 4
vaSpread1.text = vg_codigo

vaSpread1.Col = 5
vaSpread1.text = vg_nombre

If modo <> "A" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub
