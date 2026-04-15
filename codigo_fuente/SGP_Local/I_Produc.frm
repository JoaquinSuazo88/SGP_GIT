VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form I_Produc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Productos"
   ClientHeight    =   6165
   ClientLeft      =   1860
   ClientTop       =   2220
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6165
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5775
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   8895
      Begin VB.OptionButton Option1 
         Caption         =   "Nombre Fantasia"
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   9
         Top             =   5280
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nombre Producto"
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   8
         Top             =   5280
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Selección Productos"
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   135
         TabIndex        =   7
         Top             =   960
         Width           =   4455
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3735
            Left            =   135
            TabIndex        =   12
            Top             =   330
            Width           =   4215
            _Version        =   393216
            _ExtentX        =   7435
            _ExtentY        =   6588
            _StockProps     =   64
            AllowMultiBlocks=   -1  'True
            DisplayRowHeaders=   0   'False
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
            MaxRows         =   20
            SpreadDesigner  =   "I_Produc.frx":0000
            ScrollBarTrack  =   3
         End
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "I_Produc.frx":047A
         Left            =   1080
         List            =   "I_Produc.frx":047C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   3795
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Seleccción Nutrientes"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   4215
         Index           =   1
         Left            =   4695
         TabIndex        =   2
         Top             =   960
         Width           =   4050
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   " P%| G%| CHO%"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   3720
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   3345
            Left            =   240
            MultiSelect     =   1  'Simple
            TabIndex        =   4
            Top             =   360
            Width           =   3570
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "Código Prod."
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   2280
            TabIndex        =   3
            Top             =   3720
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Metodo Preparación"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5580
         TabIndex        =   11
         Top             =   375
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblREGSEL 
         Caption         =   "lblREGSEL"
         Height          =   270
         Left            =   5190
         TabIndex        =   13
         Top             =   450
         Width           =   2565
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Informes"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   435
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_Produc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim i As Integer, iselecc As Integer
Dim Opx As String, cuenta As Long
Dim aAp As String

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
    Check2(1).Enabled = False
    Check2(1).Value = 0
    Check2(2).Enabled = False
    Check2(2).Value = 0
    List1.Clear
    List1.Enabled = False
    Frame1(1).Enabled = False
    Frame1(1).Caption = ""
    Check1.Enabled = False
    Check1.Value = 0
    Option1(0).Visible = IIf(Opx = "P", False, True)
    Option1(1).Visible = IIf(Opx = "P", False, True)
ElseIf Combo1.ListIndex = 1 And Opx = "P" Then
    Dim contimp As Long
    '------- Llenar tabla contimp
    RS.Open RutinaLectura.Impuesto(1, 0, ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No existe maestro impuestos", vbExclamation + vbOKOnly, Msgtitulo: Me.Hide: Unload Me
    List1.Clear: contimp = 0
    Do While Not RS.EOF
        List1.AddItem Trim(RS!imp_nombre)
        List1.ItemData(List1.NewIndex) = RS!imp_codigo
        List1.Selected(contimp) = True
        contimp = contimp + 1
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    '-------
    Check2(1).Enabled = False
    Check2(1).Value = 0
    Check2(2).Enabled = False
    Check2(2).Value = 0
    List1.Enabled = True
    Frame1(1).Caption = "Selección Impuesto"
    Frame1(1).Enabled = True
    Check1.Enabled = True
    Check1.Value = 0
ElseIf Combo1.ListIndex = 1 And Opx = "I" Then
    Dim iaporte As Long
    '------- Llenar tabla nutrientes
    RS.Open RutinaLectura.Nutriente(1, 0, ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No existe maestro nutrientes", vbExclamation + vbOKOnly, Msgtitulo: Me.Hide: Unload Me
    List1.Clear: iaporte = 0
    Do While Not RS.EOF
        List1.AddItem Trim(RS!nut_nombre)
        List1.ItemData(List1.NewIndex) = RS!nut_codigo
        If RS!nut_indpri = 1 Then List1.Selected(iaporte) = True
        iaporte = iaporte + 1
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    '-------
    Check1.Enabled = False
    Check1.Value = 0
    Check2(1).Enabled = True
    Check2(2).Enabled = True
    List1.Enabled = True
    Frame1(1).Enabled = True
    Frame1(1).Caption = "Selección Nutrientes"
End If
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
On Error GoTo Man_Error
fg_centra Me
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
cuenta = 0: aAp = ""
lblREGSEL.Caption = cuenta & " registros seleccionados"
Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, "Informe Productos"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Dim i As Long, codpro As String, NomPro As String, NomFan As String
aAp = ""
Select Case Button.Index
Case 1
    fg_carga ""
    '-----Crea tabla temporal-----
    If Combo1.ListIndex = 1 Then
       aAp = Trim(vg_NUsr) & "_tmp_imppro"
       fg_CheckTmp aAp
       vg_db.BeginTrans
       vg_db.Execute "CREATE TABLE " & aAp & " (tem_codigo varchar(20), tem_nombre varchar(50), tem_nomfan varchar(100), tem_codpat varchar(20))"
       vg_db.CommitTrans
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Col = 1: vaSpread1.Row = i
           If vaSpread1.text = "1" Then
              vaSpread1.Col = 2: codpro = Trim(vaSpread1.text)
              vaSpread1.Col = 3: NomPro = Trim(vaSpread1.text)
              vaSpread1.Col = 4: NomFan = Trim(vaSpread1.text)
              vg_db.BeginTrans
              vg_db.Execute "INSERT INTO " & aAp & " (tem_codigo, tem_nombre, tem_nomfan, tem_codpat) " & _
                            "VALUES ('" & codpro & "', '" & NomPro & "', '" & NomFan & "', '0')"
              If Combo1.ListIndex = 1 And Opx = "P" Then
                  vg_db.Execute "INSERT INTO " & aAp & " (tem_codigo, tem_nombre, tem_nomfan, tem_codpat) " & _
                                "SELECT ipr_codimp, 'x', '', '" & codpro & "' " & _
                                "FROM b_productosimp WHERE ipr_codpro = '" & codpro & "'"
              ElseIf Combo1.ListIndex = 1 And Opx = "I" Then
                  vg_db.Execute "INSERT INTO " & aAp & " (tem_codigo, tem_nombre, tem_nomfan, tem_codpat) " & _
                                "SELECT pnu_codapo, pnu_canapo, '', '" & codpro & "' " & _
                                "FROM b_productonut WHERE pnu_codpro = '" & codpro & "'"
              End If
              vg_db.CommitTrans
          End If
       Next i
    End If
    '----------------------------------------
    iselecc = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then iselecc = 1: Exit For
    Next i
    If iselecc = 0 Then MsgBox "Debe seleccionar a lo menor un producto", vbExclamation + vbOKOnly, Msgtitulo: fg_descarga: Exit Sub
    If Combo1.ListIndex = 0 Then
       If Opx = "P" Then I_Productos Else I_Ingrediente
    ElseIf Combo1.ListIndex = 1 And Opx = "P" Then
       iselecc = 0
       For i = 0 To List1.listcount - 1
           If List1.Selected(i) = True Then iselecc = 1: Exit For
       Next i
       If iselecc = 0 Then MsgBox "Debe seleccionar a lo menos un Impuesto", vbExclamation + vbOKOnly, Msgtitulo: fg_descarga: Exit Sub
       I_ImpuestoProductos
    ElseIf Combo1.ListIndex = 2 And Opx = "P" Then
       I_IngredientesProductos
    ElseIf Combo1.ListIndex = 1 And Opx = "I" Then
       iselecc = 0
       For i = 0 To List1.listcount - 1
           If List1.Selected(i) = True Then iselecc = 1: Exit For
       Next i
       If iselecc = 0 Then MsgBox "Debe seleccionar a lo menos un Aporte Nutricional", vbExclamation + vbOKOnly, Msgtitulo: fg_descarga: Exit Sub
       I_AporteProductos
    End If
    fg_descarga
Case 3
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
If Err.Number = -2147467259 Then
    MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"
Else
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If
vg_db.RollbackTrans
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
Dim i As Long
vaSpread1.Col = 1
For i = BlockRow To BlockRow2
    vaSpread1.Row = i
    vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
Next
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
vaSpread1.Row = Row: vaSpread1.Col = Col
If Row = -1 And vaSpread1.text = 0 Then
    cuenta = 0
ElseIf Row = -1 And vaSpread1.text = 1 Then
    cuenta = vaSpread1.MaxRows
Else
    If vaSpread1.text = 1 Then cuenta = cuenta + 1 Else cuenta = cuenta - 1
End If
lblREGSEL.Caption = cuenta & " registros seleccionados"
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
'If Col = 1 And Row = 0 Then vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.Value = IIf(vaSpread1.Value = "1", "", "1")
End Sub

Sub TraspasoGrilla(vaSrepadX As vaSpread, op As String)
Dim i As Long, j As Long
Dim vecfamprod() As Variant
fg_carga ""
Opx = op
If Opx = "P" Then
    With Combo1
        .Clear
        .AddItem "Productos"
        .AddItem "Impuesto Productos"
        .AddItem "Ingredientes & Productos"
    End With
    Msgtitulo = "Impresión de Productos"
    Me.Caption = "Imprimir Productos"
    RS1.Open RutinaLectura.Producto(1, "", ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
    ReDim vecfamprod(RS1.RecordCount, 2)
    i = 1
    Do While Not RS1.EOF
       vecfamprod(i, 1) = RS1!pro_codtip
       vecfamprod(i, 2) = fg_BuscaenArbol(RS1!pro_codtip, "a_tipopro", "tip_codigo")
       RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing
    With vaSpread1
        .MaxRows = vaSrepadX.MaxRows
        For i = 1 To vaSrepadX.MaxRows
            .Row = i: vaSrepadX.Row = i
            .Col = 2: vaSrepadX.Col = 1
            .TypeHAlign = 1: .CellType = 5
            .Value = vaSrepadX.Value
            .Col = 3: vaSrepadX.Col = 2
            .TypeHAlign = 0: .CellType = 5
            .Value = vaSrepadX.Value
            .Col = 4: vaSrepadX.Col = 5
            For j = 1 To UBound(vecfamprod)
                If vecfamprod(j, 1) = vaSrepadX.text Then .text = vecfamprod(j, 2): Exit For
            Next j
    '        .Value = vaSrepadX.Value
        Next i
    End With
ElseIf Opx = "I" Then
    With Combo1
        .Clear
        .AddItem "Ingredientes"
        .AddItem "Aportes Nutricionales Ingredientes"
    End With
    Msgtitulo = "Impresión de Ingredientes"
    Me.Caption = "Imprimir Ingredientes"
    With vaSpread1
        RS1.Open RutinaLectura.Ingrediente(1, "", ""), vg_db, adOpenStatic
        i = 1
        Do While Not RS1.EOF
            .MaxRows = i: .Row = i
            .Col = 2: .TypeHAlign = 1: .CellType = 5
            .Value = RS1!ing_codigo
            .Col = 3: .TypeHAlign = 0: .CellType = 5
            .Value = RS1!ing_nombre
            .Col = 4: .Value = RS1!ing_nomfan
            RS1.MoveNext: i = i + 1
        Loop
        RS1.Close: Set RS1 = Nothing
    End With
End If
Combo1.ListIndex = 0
fg_descarga
End Sub
