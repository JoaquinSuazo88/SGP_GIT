VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_SacSgp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asociar Productos SAC vs SGP"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   14070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   13875
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   1080
         TabIndex        =   13
         Top             =   6000
         Width           =   1425
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   2
            Top             =   135
            Width           =   1320
         End
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   2520
         TabIndex        =   12
         Top             =   6000
         Width           =   3645
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   3
            Top             =   135
            Width           =   3540
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "M_SacSgp.frx":0000
         Left            =   2160
         List            =   "M_SacSgp.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4575
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5055
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   13650
         _Version        =   393216
         _ExtentX        =   24077
         _ExtentY        =   8916
         _StockProps     =   64
         ButtonDrawMode  =   2
         EditEnterAction =   2
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
         MaxCols         =   10
         MaxRows         =   20
         SpreadDesigner  =   "M_SacSgp.frx":0004
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   615
         _Version        =   393216
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   1
         MaxRows         =   1
         SpreadDesigner  =   "M_SacSgp.frx":0646
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2205
         TabIndex        =   11
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Familia productos SAC"
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   165
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   7080
      Visible         =   0   'False
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H80000003&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   10920
      Top             =   7080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   8865
      Top             =   7110
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Productos Vigentes"
      Height          =   195
      Index           =   0
      Left            =   9225
      TabIndex        =   7
      Top             =   7080
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00D9D9FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   6720
      Top             =   7110
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Productos No Vigentes"
      Height          =   195
      Index           =   1
      Left            =   7080
      TabIndex        =   6
      Top             =   7080
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "M_SacSgp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim Est As Boolean, modo As String, Codigo As String
Dim Msgtitulo As String

Private Sub Combo1_Click(Index As Integer)
Codigo = Trim(fg_codigocbo(Combo1, 0, 10, ""))
MoverDatosGrilla
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 7830
Me.Width = 14160
Me.HelpContextID = vg_OpcM
fg_centra Me
Msgtitulo = "Asociar Productos SAC vs SGP"
modo = "": Est = True: Codigo = ""
Gl_Mo_Botones Me, 12
Combo1(0).Clear
Set RS = vg_db.Execute("sgpadm_s_catproductosac 1, '', ''")
Combo1(0).AddItem "Todos" & Space(150) & "(" & fg_pone_cero(Str(0), 10) & ")"
Do While Not RS.EOF
    Combo1(0).AddItem IIf(IsNull(RS!cap_nombre), "", RS!cap_nombre) & Space(150) & "(" & fg_pone_cero(Str(RS!cap_codigo), 10) & ")"
    RS.MoveNext
Loop
RS.Close: Set RS = Nothing
Combo1(0).ListIndex = 0
MoverDatosGrilla
modo = "": Est = False
End Sub

Sub MoverDatosGrilla()
fg_carga ""
Dim x As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 250
x = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
Set RS = vg_db.Execute("sgpadm_s_formatocompras 1, '" & Codigo & "', ''")
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = Shape1(0).FillColor
If Not RS.EOF Then
    Do While Not RS.EOF
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       vaSpread1.Row = vaSpread1.MaxRows
       
       vaSpread1.Col = 1
       vaSpread1.TypeHAlign = TypeHAlignLeft
       vaSpread1.text = IIf(IsNull(RS!cap_nombre), "", Trim(RS!cap_nombre))
       
       vaSpread1.Col = 2
       vaSpread1.TypeHAlign = TypeHAlignLeft
       vaSpread1.text = ""
       
       vaSpread1.Col = 3
       vaSpread1.TypeHAlign = TypeHAlignLeft
       vaSpread1.text = IIf(IsNull(RS!foc_codsac), "", Trim(RS!foc_codsac))
                
       vaSpread1.Col = 4
       vaSpread1.TypeHAlign = TypeHAlignLeft
       vaSpread1.text = IIf(IsNull(RS!foc_nomsac), "", Trim(RS!foc_nomsac))
                
       vaSpread1.Col = 5
       vaSpread1.TypeHAlign = TypeHAlignLeft
       vaSpread1.text = IIf(IsNull(RS!foc_unisac), "", Trim(RS!foc_unisac))
       
       vaSpread1.Col = 6
       vaSpread1.TypeHAlign = TypeHAlignLeft
       vaSpread1.text = IIf(IsNull(RS!foc_vigfin), "", Format(RS!foc_vigfin, "dd/mm/yyyy"))
                
       RS.MoveNext
    Loop
End If
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
Text1(1).text = ""
Text1(2).text = ""
Combo1(0).Enabled = True
Gl_Ac_Botones Me, 12, 0, modo
Toolbar1.Buttons(1).ToolTipText = "Incluir Formato SGP"
fg_descarga
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 1, 2
    vaSpread1.Visible = False
    If Trim(Text1(Index).text) <> "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = Index + 2
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread1.Col = 1 + 2
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           Else
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           End If
        Next i
        vaSpread1.SetActiveCell Index + 2, 1
    End If
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
    If Trim(Text1(Index).text) = "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Next
       vaSpread1.SetActiveCell Index + 2, vaSpread1.SearchCol(Index + 2, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index + 2, 1
    End If
    vaSpread1.Visible = True
End Select
End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Row As Long, i As Long, codsac As String, codsgp As String, auxsgp As String, sgppre As Integer
Select Case Button.Index
Case 1 '------> Incluir formato sgp
    If vaSpread1.MaxRows < 1 Then Exit Sub
    vg_nombre = "": vg_codigo = ""
    vg_left = Me.Left + 7550
    vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = -1
    If vaSpread1.RowHidden = True Then Exit Sub
    vaSpread1.BackColor = Shape1(2).FillColor
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "ProdActi"
    B_TabEst.Show 1
    vaSpread1.Row = -1: vaSpread1.Col = -1
    vaSpread1.BackColor = Shape1(0).FillColor
    If vg_codigo = "" Then Exit Sub
    Row = vaSpread1.ActiveRow
'    '-------> Validar productos en grilla
'    For i = 1 To vaSpread1.MaxRows
'        vaSpread1.row = i
'        vaSpread1.Col = 4
'        If Trim(vaSpread1.text) = Trim(vg_codigo) Then MsgBox "Existe producto en la grilla...", vbExclamation, Msgtitulo: Exit Sub
'    Next i
    vaSpread1.Row = Row
    vaSpread1.Col = 7: vaSpread1.text = Trim(vg_codigo)
    vaSpread1.Col = 8: vaSpread1.text = Trim(vg_nombre)
    vaSpread1.Col = 9: vaSpread1.text = Trim(vg_ames)
    Set RS = vg_db.Execute("sgpadm_s_productos 26, '" & Trim(LimpiaDato(vg_codigo)) & "', '', '" & vg_NUsr & "'")
    If Not RS.EOF Then
       vaSpread1.text = IIf(IsNull(RS!pro_codigo), "", Trim(RS!pro_codigo))
       vaSpread1.Col = 8
       vaSpread1.text = IIf(IsNull(RS!pro_nombre), "", Trim(RS!pro_nombre))
       vaSpread1.Col = 9
       vaSpread1.text = IIf(IsNull(RS!uni_nomcor), "", Trim(RS!uni_nomcor))
       vaSpread1.Col = 10
       vaSpread1.text = IIf(IsNull(RS!pro_indppr), "", IIf(Trim(RS!pro_indppr) = "1", "Real", "Propuesta"))
    Else
       vaSpread1.text = ""
       vaSpread1.Col = 8: vaSpread1.text = ""
       vaSpread1.Col = 9: vaSpread1.text = ""
       vaSpread1.Col = 10: vaSpread1.text = ""
'       RS.Close: Set RS = Nothing
'       MsgBox "No existe producto...", vbExclamation, Msgtitulo: Exit Sub
    End If
    RS.Close: Set RS = Nothing
    modo = ""
    If Toolbar1.Buttons(12).Visible = False Then Combo1(0).Enabled = False: Gl_Ac_Botones Me, 12, 1, modo
Case 5 '-------> borrar vinculo sac
    If vaSpread1.MaxRows < 1 Then Exit Sub
    Row = vaSpread1.ActiveRow
    vaSpread1.Row = Row: vaSpread1.Col = 2
    If vaSpread1.CellType <> CellTypePicture Or vaSpread1.RowHidden = True Then Exit Sub
    If MsgBox("Elimina vinculo sac...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = Row
    vaSpread1.Col = 2: vaSpread1.text = ""
    vaSpread1.CellType = CellTypeStaticText
'    vaSpread1.Col = 6: vaSpread1.text = ""
    modo = ""
    If Toolbar1.Buttons(12).Visible = False Then Combo1(0).Enabled = False: Gl_Ac_Botones Me, 12, 1, modo
Case 7 '-------> Exportar excel
    Dim NashXl As Excel.Application
    fg_carga ""
    Set NashXl = CreateObject("excel.application")
    Set NashXl = New Excel.Application
    NashXl.SheetsInNewWorkbook = 1
    NashXl.Workbooks.Add
    NashXl.Range("A1").Select
'    NashXl.ActiveCell.FormulaR1C1 = Label1.Caption & ": " & Trim(Mid(Combo1(0).text, 1, 150))
    vaSpread1.AllowMultiBlocks = True
    vaSpread1.SetSelection 1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows
    vaSpread1.ClipboardCopy
    '-------> formatear columna texto
    NashXl.Range("C:C").Select
    NashXl.Selection.NumberFormat = "@"
    
    NashXl.Range("A3").Select
    NashXl.ActiveSheet.Paste

    NashXl.Cells.Select
    NashXl.Cells.EntireColumn.AutoFit
    vaSpread1.AllowMultiBlocks = False: vaSpread1.SetSelection 1, 0, vaSpread1.MaxCols, vaSpread1.MaxRows
    fg_descarga
    NashXl.Visible = True
Case 10 '-------> Cancelar
    If vaSpread1.MaxRows < 1 Then Exit Sub
    MoverDatosGrilla
Case 12 '-------> Grabar datos
    If vaSpread1.MaxRows < 1 Then Exit Sub
    fg_carga ""
    Dim estvin As Boolean
    If MsgBox("Esta Seguro ?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    estvin = False
    Bar1(0).Visible = True: Bar1(0).Value = 0
    For i = 1 To vaSpread1.MaxRows
        Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
        vaSpread1.Row = i
        vaSpread1.Col = 2: sgppre = 0:
        If vaSpread1.CellType = CellTypePicture Then sgppre = 1
        vaSpread1.Col = 3: codsca = "": codsac = Trim(vaSpread1.text)
        vaSpread1.Col = 7: codsgp = "": codsgp = Trim(vaSpread1.text)
        If codsgp <> "" Then
           estvin = True
           Set RS = vg_db.Execute("sgpadm_s_productos 27, '" & codsgp & "', '" & codsac & "', '" & vg_NUsr & "'")
'           RS.Open "SELECT DISTINCT b.fcs_codsac FROM b_formatocompras a, b_formatocomprassgp b " & _
'                    "WHERE a.foc_codsac = b.fcs_codsac " & _
'                    "AND   b.fcs_codsgp = '" & codsgp & "' " & _
'                    "AND   b.fcs_codsac = '" & codsac & "' " & _
'                    "AND   a.foc_flexec = 0", vg_db, adOpenStatic
           If RS.EOF Then
              vg_db.Execute "INSERT INTO b_formatocomprassgp (fcs_codsac, fcs_codsgp, fcs_sgppre) VALUES ('" & codsac & "', '" & codsgp & "', " & sgppre & ")"
           End If
           RS.Close: Set RS = Nothing
        End If
    Next i
    '-------> Sp para activar los formato de compras sac que no han sido vinculado
    If estvin Then vg_db.Execute "sgpadm_p_actuvinculofcompras"
    estvin = False
    Bar1(0).Visible = False: fg_descarga
    MsgBox "Generación grabado finalizado sin problema", vbInformation + vbOKOnly, Msgtitulo
    MoverDatosGrilla
Case 15 '-------> Fijar Vinculo
    If vaSpread1.MaxRows < 1 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow: Row = vaSpread1.ActiveRow
    If vaSpread1.RowHidden = True Then Exit Sub
    vaSpread1.Col = 7
    If Trim(vaSpread1.text) = "" Then MsgBox "Debe tener ingresado un producto SGP", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
    '-------> validar que exista vinculo
    Set RS = vg_db.Execute("sgpadm_s_formatocompras 2, '" & Trim(LimpiaDato(vaSpread1.text)) & "', ''")
    If Not RS.EOF Then MsgBox "Formato de compras SAC, ya esta vinculado" & VgLinea & VgLinea & Trim(RS!foc_codsac) & VgLinea & Trim(RS!foc_nomsac), vbInformation + vbOKOnly, Msgtitulo: RS.Close: Set RS = Nothing: Exit Sub
    RS.Close: Set RS = Nothing
    '-------> Validar en la grilla si ya existe vinculo
    vaSpread1.Col = 7: codsgp = "": codsgp = Trim(vaSpread1.text)
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 7: auxsgp = Trim(vaSpread1.text)
        vaSpread1.Col = 1
        If auxsgp <> "" And codsgp = auxsgp And vaSpread1.CellType = CellTypePicture And i <> Row Then MsgBox "Formato de compra SAC, esta viculado en la grilla...", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
    Next i
    vaSpread2.Row = 1: vaSpread2.Col = 1
    vaSpread1.Row = Row
    vaSpread1.Col = 2
    vaSpread1.CellType = CellTypePicture
    vaSpread1.TypePictCenter = True
    vaSpread1.TypePictMaintainScale = True
    vaSpread1.TypePictStretch = True
    vaSpread1.TypePictPicture = vaSpread2.TypePictPicture
    vaSpread1.text = vaSpread2.text
Case 17 '-------> Cerrar opción
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
Select Case Col
Case 7
    vaSpread1.Row = Row
    vaSpread1.Col = Col
    Set RS = vg_db.Execute("sgpadm_s_productos 26, '" & Trim(LimpiaDato(vaSpread1.text)) & "', '', '" & vg_NUsr & "'")
    If Not RS.EOF Then
       vaSpread1.text = IIf(IsNull(RS!pro_codigo), "", Trim(RS!pro_codigo))
       vaSpread1.Col = 8
       vaSpread1.text = IIf(IsNull(RS!pro_nombre), "", Trim(RS!pro_nombre))
       vaSpread1.Col = 9
       vaSpread1.text = IIf(IsNull(RS!uni_nomcor), "", Trim(RS!uni_nomcor))
       vaSpread1.Col = 10
       vaSpread1.text = IIf(IsNull(RS!pro_indppr), "", IIf(Trim(RS!pro_indppr) = "1", "Real", "Propuesta"))
    Else
       vaSpread1.text = ""
       vaSpread1.Col = 8: vaSpread1.text = ""
       vaSpread1.Col = 9: vaSpread1.text = ""
       vaSpread1.Col = 10: vaSpread1.text = ""
'       RS.Close: Set RS = Nothing
'       MsgBox "No existe producto...", vbExclamation, Msgtitulo: Exit Sub
    End If
    RS.Close: Set RS = Nothing
    If Toolbar1.Buttons(12).Visible = False Then Combo1(0).Enabled = False: Gl_Ac_Botones Me, 12, 1, modo
End Select

End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
Select Case Col
Case 7
'Preview.Toolbar1_ButtonClick Preview.Toolbar1.Buttons("word")
    Toolbar1_ButtonClick Toolbar1.Buttons(1)
End Select
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
If ChangeMade = False Then Exit Sub
Select Case Col
Case 7
    vaSpread1.Row = Row
    vaSpread1.Col = Col
    Set RS = vg_db.Execute("sgpadm_s_productos 26, '" & Trim(LimpiaDato(vaSpread1.text)) & "', '', '" & vg_NUsr & "'")
    If Not RS.EOF Then
       vaSpread1.text = IIf(IsNull(RS!pro_codigo), "", Trim(RS!pro_codigo))
       vaSpread1.Col = 8
       vaSpread1.text = IIf(IsNull(RS!pro_nombre), "", Trim(RS!pro_nombre))
       vaSpread1.Col = 9
       vaSpread1.text = IIf(IsNull(RS!uni_nomcor), "", Trim(RS!uni_nomcor))
    Else
       vaSpread1.text = ""
       vaSpread1.Col = 8: vaSpread1.text = ""
       vaSpread1.Col = 9: vaSpread1.text = ""
'       RS.Close: Set RS = Nothing
'       MsgBox "No existe producto...", vbExclamation, Msgtitulo: Exit Sub
    End If
    RS.Close: Set RS = Nothing
    If Toolbar1.Buttons(12).Visible = False Then Combo1(0).Enabled = False: Gl_Ac_Botones Me, 12, 1, modo
End Select
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread1.Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
Select Case Col
Case 3
    vaSpread1.Col = Col
    TipText = "Código Sac: " & vaSpread1.text
Case 4
    vaSpread1.Col = Col
    TipText = "Descripción " & Trim(vaSpread1.text)
Case 5
    vaSpread1.Col = Col
    TipText = "Unidad Medida SAC : " & Trim(vaSpread1.text)
Case 6
    vaSpread1.Col = Col
    TipText = "Fecha Vigencia : " & Trim(vaSpread1.text)
Case 7
    vaSpread1.Col = Col
    TipText = "Código SGP : " & Trim(vaSpread1.text)
Case 8
    vaSpread1.Col = Col
    TipText = "Descripción SGP : " & Trim(vaSpread1.text)
Case 9
    vaSpread1.Col = Col
    TipText = "Unidad Medida SGP : " & Trim(vaSpread1.text)
End Select
End Sub
