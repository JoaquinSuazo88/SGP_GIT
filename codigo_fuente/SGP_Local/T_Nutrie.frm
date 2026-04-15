VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form T_Nutrie 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nutrientes"
   ClientHeight    =   4800
   ClientLeft      =   1860
   ClientTop       =   1995
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3435
      Left            =   0
      TabIndex        =   7
      Top             =   1350
      Width           =   8715
      _Version        =   393216
      _ExtentX        =   15372
      _ExtentY        =   6059
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
      MaxCols         =   5
      SpreadDesigner  =   "T_Nutrie.frx":0000
      ScrollBarTrack  =   3
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   971
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "T_Nutrie.frx":19E1
         Left            =   1965
         List            =   "T_Nutrie.frx":19EB
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1965
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
         Left            =   4545
         TabIndex        =   4
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
         Left            =   480
         TabIndex        =   3
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
         Left            =   480
         TabIndex        =   2
         Top             =   345
         Width           =   1380
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "T_Nutrie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String
Dim OpGr As Boolean
Dim ibusca As Long

Private Sub GrabaRegistro(Fila As Long)
On Error GoTo Man_Error
Dim codigo As Long
Dim Nombre As String
Dim NomCor As String
OpGr = True
With vaSpread1
    .Row = Fila
    .Col = 1: codigo = Val(.Value)
    .Col = 2: Nombre = Trim(LimpiaDato(.Value))
    .Col = 3: NomCor = Trim(LimpiaDato(.Value))
    .Col = 4
    If .text = "1" Then
        .Col = 4: indpri = 1
    Else
        .Col = 4: indpri = 0
    End If
    .Col = 5: secnro = Trim(LimpiaDato(.Value))
    If Trim(Nombre) = "" Or Trim(NomCor) = "" Or secnro = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: .Row = Fila: .Col = 2: .SetActiveCell 2, .Row: .SetFocus: OpGr = False: Exit Sub
    If modo = "A" Then
        vg_db.BeginTrans
        RS1.Open RutinaLectura.Nutriente(2, 0, ""), vg_db, adOpenStatic
        If Not RS1.EOF Then
            RS1.MoveFirst
            codigo = RS1!nut_codigo + 1
        Else
            codigo = 1
        End If
        RS1.Close: Set RS1 = Nothing
        vg_db.Execute "INSERT INTO a_nutriente (nut_codigo, nut_nombre, nut_nomuni,nut_indpri,nut_secnro) VALUES (" & codigo & ", '" & Trim(Nombre) & "', '" & Trim(NomCor) & "'," & indpri & "," & secnro & ")"
        vg_db.CommitTrans
        .Col = 1: .Value = codigo
    Else
        vg_db.BeginTrans
        vg_db.Execute "UPDATE a_nutriente SET nut_nombre='" & Trim(Nombre) & "', nut_nomuni='" & Trim(NomCor) & "',nut_indpri=" & indpri & ",nut_secnro=" & secnro & " WHERE nut_codigo=" & codigo
        vg_db.CommitTrans
    End If
    Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registro"
    Combo1.Enabled = True: fpText1.Enabled = True
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    OpGr = False
End With
Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
If Err.Number = -2147467259 Then
    vg_db.RollbackTrans
    MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"
Else
    vg_db.RollbackTrans
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 5235
Me.Width = 8840
Msgtitulo = "Nutriente"
fg_centra Me
modo = "": ibusca = 0
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, IIf(vg_modprod = True, 1, 6), modo
Combo1.ListIndex = 1
MoverDatosGrillas
OpGr = False
End Sub

Private Sub Form_Resize()
If Me.WindowState = 0 Then
   Frame1.Move 1200, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
ElseIf Me.WindowState = 2 Then
   Frame1.Move 4200, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
End If
Toolbar1.Refresh
End Sub

Private Sub fpText1_Change()
If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    RS2.Open RutinaLectura.Nutriente(4, 0, UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    RS2.Open RutinaLectura.Nutriente(5, 0, UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
End If
With vaSpread1
    ibusca = RS2.RecordCount: .MaxRows = RS2.RecordCount
    i = 1
    If Not RS2.EOF Then
        Do While Not RS2.EOF
           .Row = i
           i = i + 1
           .Col = 1: .Value = RS2!nut_codigo
           .Col = 2: .Value = IIf(IsNull(RS2!nut_nombre), "", Trim(RS2!nut_nombre))
           .Col = 3: .Value = IIf(IsNull(RS2!nut_nomuni), "", Trim(RS2!nut_nomuni))
           If RS2!nut_indpri = 1 Then
              .Col = 4
              .CellType = 10
              .TypeCheckText = ""
              .TypeCheckCenter = True
              .text = "1"
           End If
           .Col = 5
           .Value = IIf(IsNull(RS2!nut_secnro), "", RS2!nut_secnro)
           RS2.MoveNext
       Loop
       Gl_Ac_Botones Me, 1, IIf(vg_modprod = True, 1, 6), modo
    End If
    RS2.Close: Set RS2 = Nothing
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
    vg_db.Execute "delete a_nutriente from a_nutriente where nut_codigo=" & codigo
    vg_db.CommitTrans
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
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
    I_Nutri
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
If Err.Number = -2147467259 Then
    vg_db.RollbackTrans
    MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"
Else
    vg_db.RollbackTrans
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If Col <> 4 Or OpGr Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, IIf(vg_modprod = True, 0, 6), modo
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, IIf(vg_modprod = True, 0, 6), modo
End Sub

Private Sub MoverDatosGrillas()
OpGr = True
With vaSpread1
    .MaxRows = 0
    RS2.Open RutinaLectura.Nutriente(3, 0, ""), vg_db, adOpenStatic
    Do While Not RS2.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1: .Value = RS2!nut_codigo
        .Col = 2: .Value = IIf(IsNull(RS2!nut_nombre), "", Trim(RS2!nut_nombre))
        .Col = 3: .Value = IIf(IsNull(RS2!nut_nomuni), "", Trim(RS2!nut_nomuni))
        If RS2!nut_indpri = 1 Then
            .Col = 4
            .CellType = 10
            .TypeCheckText = ""
            .TypeCheckCenter = True
            .text = "1"
        End If
        .Col = 5: .Value = IIf(IsNull(RS2!nut_secnro), "", Trim(RS2!nut_secnro))
        RS2.MoveNext
    Loop
    RS2.Close: Set RS2 = Nothing
    Gl_Ac_Botones Me, 1, IIf(vg_modprod = True, 1, 6), modo
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
Dim codigo As Long
OpGr = True
With vaSpread1
    .Row = .ActiveRow
    .Col = 1: codigo = Val(.Value)
    RS1.Open RutinaLectura.Nutriente(3, codigo, ""), vg_db, adOpenStatic
    If Not RS1.EOF Then
       .Col = 2: .Value = IIf(IsNull(RS1!nut_nombre), "", Trim(RS1!nut_nombre))
       .Col = 3: .Value = IIf(IsNull(RS1!nut_nomuni), "", Trim(RS1!nut_nomuni))
       .Col = 4
       .CellType = 10
       .TypeCheckText = ""
       .TypeCheckCenter = True
       .text = IIf(IsNull(RS1!nut_indpri), "", Trim(Str(RS1!nut_indpri)))
        
       .Col = 5: .Value = IIf(IsNull(RS1!nut_secnro), "", Trim(RS1!nut_secnro))
    End If
    RS1.Close: Set RS1 = Nothing
End With
OpGr = False
End Sub
