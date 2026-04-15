VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form T_Unienb 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unidad embalaje"
   ClientHeight    =   4725
   ClientLeft      =   3420
   ClientTop       =   2385
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   7425
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3345
      Left            =   0
      TabIndex        =   7
      Top             =   1350
      Width           =   7395
      _Version        =   393216
      _ExtentX        =   13044
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
      SpreadDesigner  =   "T_Unienb.frx":0000
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   971
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "T_Unienb.frx":188F
         Left            =   1680
         List            =   "T_Unienb.frx":1899
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
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
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4260
         TabIndex        =   4
         Top             =   640
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Texto"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   255
         TabIndex        =   3
         Top             =   645
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Columna"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   255
         TabIndex        =   2
         Top             =   345
         Width           =   1320
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "T_Unienb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String
Dim MsgTitulo As String
Dim OpGr As Boolean

Private Sub GrabaRegistro(Fila As Long)
OpGr = True
vaSpread1.Row = Fila
vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
vaSpread1.Col = 2: Nombre = Trim(LimpiaDato(vaSpread1.Value))
vaSpread1.Col = 3: NomCor = Trim(LimpiaDato(vaSpread1.Value))
If Trim(Nombre) = "" Or Trim(NomCor) = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub
If modo = "A" Then
    vg_db.BeginTrans
    RS1.Open "select uni_codigo from a_unidad order by uni_codigo desc", vg_db, adOpenStatic
    If Not RS1.EOF Then
        RS1.MoveFirst
        codigo = RS1!uni_codigo + 1
    Else
        codigo = 1
    End If
    RS1.Close: Set RS1 = Nothing
    vg_db.Execute "insert into a_unidad (uni_codigo, uni_nombre, uni_nomcor) values (" & codigo & ", '" & Trim(Nombre) & "', '" & Trim(NomCor) & "')"
    vg_db.CommitTrans
    vaSpread1.Col = 1: vaSpread1.Value = codigo
Else
    vg_db.BeginTrans
    vg_db.Execute "UPDATE a_unidad SET uni_nombre='" & Trim(Nombre) & "', uni_nomcor='" & Trim(NomCor) & "' WHERE uni_codigo=" & codigo
    vg_db.CommitTrans
End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Combo1.Enabled = True: fpText1.Enabled = True
modo = "": Ac_Botones
OpGr = False
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 5235
Me.Width = 7540
MsgTitulo = "Unidades de Envase"
fg_centra Me
modo = ""
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): btnX.Visible = True: btnX.ToolTipText = "Incluir"
Set btnX = Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): btnX.Visible = False: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, "A_Alterar", , tbrDefault, "A_Alterar"): btnX.Visible = True: btnX.ToolTipText = "Alterar"
Set btnX = Toolbar1.Buttons.Add(, "I_Alterar", , tbrDefault, "I_Alterar"): btnX.Visible = False: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): btnX.Visible = True: btnX.ToolTipText = "Borrar "
Set btnX = Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): btnX.Visible = False: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): btnX.Visible = True: btnX.ToolTipText = "Actualizar Lista   "
Set btnX = Toolbar1.Buttons.Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): btnX.Visible = False: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): btnX.Visible = False: btnX.ToolTipText = "Cancelar "
Set btnX = Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): btnX.Visible = True: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = False: btnX.ToolTipText = "Confirmar "
Set btnX = Toolbar1.Buttons.Add(, "I_Conformar ", , tbrDefault, "I_Confirmar "): btnX.Visible = True: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): btnX.Visible = True: btnX.ToolTipText = "Imprimir "
Set btnX = Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): btnX.Visible = False: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
Combo1.ListIndex = 1
Ac_Botones
MoverDatosGrillas
OpGr = False
End Sub

Private Sub Form_Resize()
If T_Unienv.WindowState = 0 Then
   Frame1.Move 0, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
ElseIf T_Unienv.WindowState = 2 Then
   Frame1.Move 4200, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
End If
Toolbar1.Refresh
End Sub

Private Sub fpText1_Change()
If LimpiaDato(Trim(fpText1.Text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    RS2.Open "select count(uni_codigo) as nreg From a_unidad where uni_codigo like '%" + UCase(LimpiaDato(fpText1.Text)) & "%'", vg_db, adOpenStatic
    ibusca = RS2!NReg: vaSpread1.MaxRows = RS2!NReg: RS2.Close: Set rs = Nothing
    RS2.Open "select uni_codigo,uni_nombre,uni_nomcor From a_unidad where uni_codigo like '%" + UCase(LimpiaDato(fpText1.Text)) & "%'", vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    RS2.Open "select count(uni_codigo) as nreg From a_unidad where ucase(uni_nombre) like '%" + UCase(LimpiaDato(fpText1.Text)) & "%'", vg_db, adOpenStatic
    ibusca = RS2!NReg: vaSpread1.MaxRows = RS2!NReg: RS2.Close: Set rs = Nothing
    RS2.Open "select uni_codigo, uni_nombre, uni_nomcor From a_unidad where ucase(uni_nombre) like '%" + UCase(LimpiaDato(fpText1.Text)) & "%'", vg_db, adOpenStatic
End If
i = 1
If Not RS2.EOF Then
   Do While Not RS2.EOF
      vaSpread1.Row = i
      i = i + 1
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = 1
      vaSpread1.Value = RS2!uni_codigo
      vaSpread1.Col = 2
      vaSpread1.Value = Trim(RS2!uni_nombre)
      vaSpread1.Col = 3
      vaSpread1.Value = Trim(RS2!uni_nomcor)
      RS2.MoveNext
   Loop
   Ac_Botones
End If
RS2.Close: Set rs = Nothing
If fpText1.Text = "" Then
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Else
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Encontrado"
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codigo As Long, Nombre As String, NomCor As String
Select Case Button.Index
Case 1
    modo = "A"
    Ac_Botones
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus
Case 3
    modo = "M"
    Ac_Botones
Case 5
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
    vg_db.BeginTrans
    vg_db.Execute "delete a_unidad from a_unidad where uni_codigo=" & codigo
    vg_db.CommitTrans
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    modo = "": Ac_Botones
Case 7
    fpText1.Text = ""
    MoverDatosGrillas
Case 10
    If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If modo = "A" Then
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    Else
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
        RS1.Open "select * from a_unidad where uni_codigo=" & codigo, vg_db, adOpenStatic
        If Not RS1.EOF Then
           vaSpread1.Col = 2: vaSpread1.Value = Trim(RS1!uni_nombre)
           vaSpread1.Col = 3: vaSpread1.Value = Trim(RS1!uni_nomcor)
        End If
        RS1.Close: Set RS1 = Nothing
    End If
    modo = "": Ac_Botones
    Combo1.Enabled = True: fpText1.Enabled = True
Case 12
    GrabaRegistro vaSpread1.ActiveRow
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    I_UniEnv
Case 18
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Ac_Botones
End Sub

Function Ac_Botones()
If modo = "A" Or modo = "M" Then
    Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(3).Visible = False: Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False: Toolbar1.Buttons(8).Visible = True
    Toolbar1.Buttons(10).Visible = True: Toolbar1.Buttons(11).Visible = False
    Toolbar1.Buttons(12).Visible = True: Toolbar1.Buttons(13).Visible = False
    Toolbar1.Buttons(15).Visible = False: Toolbar1.Buttons(16).Visible = True
    Combo1.Enabled = False: fpText1.Enabled = False
Else
    Toolbar1.Buttons(1).Visible = True: Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(3).Visible = True: Toolbar1.Buttons(4).Visible = False
    Toolbar1.Buttons(5).Visible = True: Toolbar1.Buttons(6).Visible = False
    Toolbar1.Buttons(7).Visible = True: Toolbar1.Buttons(8).Visible = False
    Toolbar1.Buttons(10).Visible = False: Toolbar1.Buttons(11).Visible = True
    Toolbar1.Buttons(12).Visible = False: Toolbar1.Buttons(13).Visible = True
    Toolbar1.Buttons(15).Visible = True: Toolbar1.Buttons(16).Visible = False
    Combo1.Enabled = True: fpText1.Enabled = True
End If
End Function

Private Sub MoverDatosGrillas()
vaSpread1.MaxRows = 0
RS2.Open "select * From a_unidad order by uni_codigo", vg_db, adOpenStatic
Do While Not RS2.EOF
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = 1: vaSpread1.Value = RS2!uni_codigo
    vaSpread1.Col = 2: vaSpread1.Value = Trim(RS2!uni_nombre)
    vaSpread1.Col = 3: vaSpread1.Value = Trim(RS2!uni_nomcor)
    RS2.MoveNext
Loop
RS2.Close: Set RS2 = Nothing
Ac_Botones
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") Then GrabaRegistro Row
End Sub

