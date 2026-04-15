VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form T_Impues 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impuestos"
   ClientHeight    =   5550
   ClientLeft      =   2130
   ClientTop       =   2415
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   13215
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   971
      Left            =   3375
      TabIndex        =   2
      Top             =   480
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "T_Impues.frx":0000
         Left            =   1875
         List            =   "T_Impues.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1875
         TabIndex        =   4
         Top             =   555
         Width           =   2385
         _Version        =   196608
         _ExtentX        =   4207
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
         Left            =   375
         TabIndex        =   7
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
         Left            =   375
         TabIndex        =   6
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
         Left            =   4335
         TabIndex        =   5
         Top             =   645
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3945
      Left            =   120
      TabIndex        =   1
      Top             =   1500
      Width           =   13005
      _Version        =   393216
      _ExtentX        =   22939
      _ExtentY        =   6959
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
      MaxCols         =   11
      SpreadDesigner  =   "T_Impues.frx":001E
      ScrollBarTrack  =   3
   End
End
Attribute VB_Name = "T_Impues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim OpGr As Boolean
Dim ibusca As Long
Dim indlec As Integer

Private Sub GrabaRegistro(Fila As Long)
On Error GoTo Man_Error
Dim RS As New ADODB.Recordset
Dim codigo As Long, incos As Long
Dim impues As Double
Dim Nombre As String, indmod As String, impadi As String, codsap As String, cimpsap1 As String, cimpsap2 As String, cimpsap3 As String, cimpsap4 As String
OpGr = True
indlec = 0
With vaSpread1
    .Row = Fila
    .Col = 1: codigo = Val(.Value)
    .Col = 2: Nombre = Trim(LimpiaDato(.Value))
    .Col = 3: impues = .Value
    .Col = 4: incos = IIf(.text = "1", 1, 0)
    .Col = 5: indmod = IIf(.text = "1", "S", "N")
    .Col = 6: codsap = Trim(LimpiaDato(.Value))
    .Col = 7: cimpsap1 = Trim(LimpiaDato(.Value))
    .Col = 8: cimpsap2 = Trim(LimpiaDato(.Value))
    .Col = 9: cimpsap3 = Trim(LimpiaDato(.Value))
    .Col = 10: cimpsap4 = Trim(LimpiaDato(.Value))
    .Col = 11: impadi = IIf(.text = "1", "1", "0")
    If Trim(Nombre) = "" Or impues = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub
    If modo = "A" Then
        vg_db.BeginTrans
        RS.Open RutinaLectura.Impuesto(2, 0, ""), vg_db, adOpenStatic
        indlec = 1
        If Not RS.EOF Then RS.MoveFirst: codigo = RS!imp_codigo + 1 Else codigo = 1
        RS.Close: Set RS = Nothing
        indlec = 0
        vg_db.Execute "INSERT INTO a_impuesto (imp_codigo, imp_nombre, imp_pctimp, imp_inccos, imp_codsap, imp_indmod, imp_adicional, imp_cimsap1, imp_cimsap2, imp_cimsap3, imp_cimsap4) VALUES (" & codigo & ", '" & Trim(Nombre) & "', " & impues & ", " & incos & ", '" & codsap & "', '" & indmod & "', '" & impadi & "', '" & cimpsap1 & "', '" & cimpsap2 & "', '" & cimpsap3 & "', '" & cimpsap4 & "')"
        vg_db.CommitTrans
        .Col = 1: .Value = codigo
    Else
        vg_db.BeginTrans
        vg_db.Execute "UPDATE a_impuesto SET imp_nombre = '" & Trim(Nombre) & "', imp_pctimp = " & impues & ", imp_inccos = " & incos & ", imp_codsap = '" & codsap & "', imp_indmod = '" & indmod & "', imp_adicional = '" & impadi & "', imp_cimsap1 = '" & cimpsap1 & "', imp_cimsap2 = '" & cimpsap2 & "', imp_cimsap3 = '" & cimpsap3 & "', imp_cimsap4 = '" & cimpsap4 & "' WHERE imp_codigo = " & codigo
        vg_db.CommitTrans
    End If
    Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registro"
End With
Combo1.Enabled = True: fpText1.Enabled = True
modo = "": Gl_Ac_Botones Me, 1, 1, modo
OpGr = False

Exit Sub
Man_Error:
If indlec = 1 Then RS1.Close: Set RS1 = Nothing
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
If Err.Number = -2147467259 Then MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error" Else MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 5970
Me.Width = 13335
MsgTitulo = "Impuesto"
fg_centra Me
modo = ""
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1
If vg_pais = "CL" Then
   vaSpread1.Row = 0
   vaSpread1.Col = 7
   vaSpread1.text = "Código Impuesto SAP"
   vaSpread1.Row = -1
   vaSpread1.Col = 8
   vaSpread1.ColHidden = True
   vaSpread1.Col = 9
   vaSpread1.ColHidden = True
   vaSpread1.Col = 10
   vaSpread1.ColHidden = True
End If
MoverDatosGrillas
OpGr = False
End Sub

Private Sub Form_Resize()
If Me.WindowState = 0 Then
   Frame1.Move 1575, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
ElseIf Me.WindowState = 2 Then
   Frame1.Move 4200, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
End If
Toolbar1.Refresh
End Sub

Private Sub fpText1_Change()
Dim RS As New ADODB.Recordset
If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
   RS.Open RutinaLectura.Impuesto(4, 0, UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
   RS.Open RutinaLectura.Impuesto(5, 0, UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
End If
With vaSpread1
    ibusca = RS.RecordCount: .MaxRows = RS.RecordCount
    i = 1
    If Not RS.EOF Then
       Do While Not RS.EOF
          .Row = i
          .Col = 1
          .Value = RS!imp_codigo
          
          .Col = 2
          .Value = IIf(IsNull(RS!imp_nombre), "", Trim(RS!imp_nombre))
          
          .Col = 3
          .Value = IIf(IsNull(RS!imp_pctimp), "", Trim(RS!imp_pctimp))
          
          .Col = 4
          .CellType = 10
          .TypeCheckText = ""
          .TypeCheckCenter = True
          .text = IIf(RS!imp_inccos = 1, "1", "")
          
          .Col = 5
          .Value = IIf(IsNull(RS!imp_indmod), "", IIf(Trim(RS!imp_codsap) = "S", "1", ""))
          
          .Col = 6
          .Value = IIf(IsNull(RS!imp_codsap), "", Trim(RS!imp_codsap))
          
          .Col = 7
          .Value = IIf(IsNull(RS!imp_cimsap1), "", Trim(RS!imp_cimsap1))
          
          .Col = 8
          .Value = IIf(IsNull(RS!imp_cimsap2), "", Trim(RS!imp_cimsap2))
          
          .Col = 9
          .Value = IIf(IsNull(RS!imp_cimsap3), "", Trim(RS!imp_cimsap3))
          
          .Col = 10
          .Value = IIf(IsNull(RS!imp_cimsap4), "", Trim(RS!imp_cimsap4))
          
          .Col = 11
          .CellType = 10
          .TypeCheckText = ""
          .TypeCheckCenter = True
          .text = IIf(RS!imp_adicional = 1, "1", "")
          RS.MoveNext: i = i + 1
       Loop
       Gl_Ac_Botones Me, 1, IIf(vg_modprod = True, 0, 6), modo
    End If
    RS.Close: Set RS = Nothing
    If vg_modprod = False Then
       .Col = 1: .Col2 = .MaxCols: .Row = 1: .Row2 = .MaxRows
       .BlockMode = True: .Lock = True: .Protect = True: .BlockMode = False
    End If
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
    vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.Row: vaSpread1.SetFocus
Case 3
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
Case 5
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
    If codigo = 1 Then MsgBox "No puede eliminar el IVA...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    vg_db.Execute "DELETE a_impuesto FROM a_impuesto WHERE imp_codigo = " & codigo
    vg_db.CommitTrans
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
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
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Combo1.Enabled = True: fpText1.Enabled = True
Case 12
    GrabaRegistro vaSpread1.ActiveRow
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    I_impues
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
vg_db.RollbackTrans
If Err = 3034 Then Exit Sub
If Err.Number = -2147467259 Then MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error" Else MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If (Col <> 4 And Col <> 5 Or Col <> 11) Or Row = 0 Or OpGr Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, IIf(vg_modprod = True, 0, 6), modo
End Sub
    
Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, IIf(vg_modprod = True, 0, 6), modo
End Sub

Private Sub MoverDatosGrillas()
Dim RS As New ADODB.Recordset
OpGr = True
With vaSpread1
    .MaxRows = 0
    RS.Open RutinaLectura.Impuesto(3, 0, ""), vg_db, adOpenStatic
    Do While Not RS.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1
       .Value = RS!imp_codigo
       
       .Col = 2
       If RS!imp_codigo = 1 Then .CellType = CellTypeStaticText
       .Value = IIf(IsNull(RS!imp_nombre), "", Trim(RS!imp_nombre))
    
       
       .Col = 3
       .Value = IIf(IsNull(RS!imp_pctimp), 0, RS!imp_pctimp)
       
       .Col = 4
       .CellType = 10
       .TypeCheckText = ""
       .TypeCheckCenter = True
       .text = IIf(RS!imp_inccos = 1, "1", "")
       
       .Col = 5
       .Value = IIf(IsNull(RS!imp_indmod), "", IIf(Trim(RS!imp_indmod) = "S", "1", ""))
       
       .Col = 6
       .Value = IIf(IsNull(RS!imp_codsap), "", Trim(RS!imp_codsap))
       
       .Col = 7
       .Value = IIf(IsNull(RS!imp_cimsap1), "", Trim(RS!imp_cimsap1))
       
       .Col = 8
       .Value = IIf(IsNull(RS!imp_cimsap2), "", Trim(RS!imp_cimsap2))
       
       .Col = 9
       .Value = IIf(IsNull(RS!imp_cimsap3), "", Trim(RS!imp_cimsap3))
       
       .Col = 10
       .Value = IIf(IsNull(RS!imp_cimsap4), "", Trim(RS!imp_cimsap4))
       
       .Col = 11
       .CellType = 10
       .TypeCheckText = ""
       .TypeCheckCenter = True
       .text = IIf(RS!imp_adicional = 1, "1", "")
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    Gl_Ac_Botones Me, 1, IIf(vg_modprod = True, IIf(.MaxRows > 0, 1, 2), 6), modo
    If vg_modprod = False Then
       .Col = 1: .Col2 = .MaxCols: .Row = 1: .Row2 = .MaxRows
       .BlockMode = True: .Lock = True: .Protect = True: .BlockMode = False
    End If
    Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registro"
End With
OpGr = False
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
    .Row = .ActiveRow
    .Col = 1: codigo = Val(.Value)
    RS.Open RutinaLectura.Impuesto(3, codigo, ""), vg_db, adOpenStatic
    If Not RS.EOF Then
       .Col = 2
       .Value = IIf(IsNull(RS!imp_nombre), "", Trim(RS!imp_nombre))
       
       .Col = 3
       .Value = IIf(IsNull(RS!imp_pctimp), 0, RS!imp_pctimp)
       
       .Col = 4
       .CellType = 10
       .TypeCheckText = ""
       .TypeCheckCenter = True
       .text = IIf(RS!imp_inccos = 1, "1", "")
       
       .Col = 5
       .Value = IIf(IsNull(RS!imp_indmod), "", IIf(Trim(RS!imp_indmod) = "S", "1", ""))
       
       .Col = 6
       .Value = IIf(IsNull(RS!imp_codsap), "", Trim(RS!imp_codsap))
       
       .Col = 7
       .Value = IIf(IsNull(RS!imp_cimsap1), "", Trim(RS!imp_cimsap1))
       
       .Col = 8
       .Value = IIf(IsNull(RS!imp_cimsap2), "", Trim(RS!imp_cimsap2))
       
       .Col = 9
       .Value = IIf(IsNull(RS!imp_cimsap3), "", Trim(RS!imp_cimsap3))
       
       .Col = 10
       .Value = IIf(IsNull(RS!imp_cimsap4), "", Trim(RS!imp_cimsap4))
       
       .Col = 11
       .CellType = 10
       .TypeCheckText = ""
       .TypeCheckCenter = True
       .text = IIf(RS!imp_adicional = 1, "1", "")
    End If
    RS.Close: Set RS = Nothing
End With
OpGr = False
End Sub
