VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form B_RutPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscador Ruta Productos"
   ClientHeight    =   8820
   ClientLeft      =   1200
   ClientTop       =   1740
   ClientWidth     =   9690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   7095
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   8775
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   2
         Left            =   600
         TabIndex        =   13
         Top             =   6480
         Width           =   1395
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   14
            Top             =   135
            Width           =   1290
         End
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   2085
         TabIndex        =   11
         Top             =   6480
         Width           =   1125
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   12
            Top             =   135
            Width           =   1020
         End
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Left            =   3240
         TabIndex        =   9
         Top             =   6480
         Width           =   4125
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   10
            Top             =   135
            Width           =   4020
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6165
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   8520
         _Version        =   393216
         _ExtentX        =   15028
         _ExtentY        =   10874
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
         MaxCols         =   5
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "B_RutPro.frx":0000
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         ItemData        =   "B_RutPro.frx":1AC2
         Left            =   2640
         List            =   "B_RutPro.frx":1AC4
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   660
         Width           =   4095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "B_RutPro.frx":1AC6
         Left            =   2640
         List            =   "B_RutPro.frx":1AC8
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2715
         TabIndex        =   8
         Top             =   765
         Width           =   4065
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2715
         TabIndex        =   7
         Top             =   345
         Width           =   4065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Familia Productos"
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
         Left            =   840
         TabIndex        =   3
         Top             =   780
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Central de Compras"
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
         Index           =   4
         Left            =   840
         TabIndex        =   2
         Top             =   345
         Width           =   1665
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8820
      Left            =   9060
      TabIndex        =   15
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   15558
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "B_RutPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim codrut As Long
Dim codcen As String
Dim codfpr As String
Dim opcion As String
Dim est As Boolean
Dim spid As Long

Private Sub Combo1_Click(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    codcen = Trim(fg_codigocbo(Combo1, 0, 10, ""))
    If Mid(codcen, 1, 1) = 0 Then
       codcen = ""
    End If
Case 1
    codfpr = Val(fg_codigocbo(Combo1, 1, 10, ""))
End Select
CargarProductosnoIncludosRuta
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
On Error GoTo Man_Error
fg_centra Me
fg_carga ""
est = True
Toolbar1.ImageList = Partida.IL1
Toolbar1.Buttons.Clear
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fg_descarga
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Titulo
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 2, 3, 4
    vaSpread1.Visible = False
    If Trim(Text1(Index).text) <> "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = Index: nom = UCase(Trim(vaSpread1.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread1.Col = 1
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           Else
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           End If
        Next i
        vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
    If Trim(Text1(Index).text) = "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Next
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1.Visible = True
End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim EstSel As Boolean
Dim i As Long
Select Case Button.Index
Case 1
   Dim codpro As String, nompro As String, codcco As String, fecvig As String
   EstSel = False
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then EstSel = True: Exit For
    Next i
    If Not EstSel Then MsgBox "Debe Seleccionar A lo menor un producto", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    EstSel = False
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then
           vaSpread1.Col = 2
           codpro = vaSpread1.text
           vaSpread1.Col = 3
           codcco = vaSpread1.text
           vaSpread1.Col = 4
           nompro = vaSpread1.text
           vaSpread1.Col = 5
           fecvig = vaSpread1.text
           If opcion = "rutpro" Then
              '-------> Asignar datos a la grilla ruta productos
              M_Ruta.vaSpread2.MaxRows = M_Ruta.vaSpread2.MaxRows + 1
              M_Ruta.vaSpread2.Row = M_Ruta.vaSpread2.MaxRows
              If EstSel = False Then vg_codigo = M_Ruta.vaSpread2.MaxRows: EstSel = True
              M_Ruta.vaSpread2.Col = -1
              M_Ruta.vaSpread2.BackColor = &H80000013
              M_Ruta.vaSpread2.Col = 2
              M_Ruta.vaSpread2.text = codpro
              M_Ruta.vaSpread2.Col = 3
              M_Ruta.vaSpread2.text = codcco
              M_Ruta.vaSpread2.Col = 4
              M_Ruta.vaSpread2.text = Trim(Mid(Combo1(1).text, 1, 150)) 'codfpr
              M_Ruta.vaSpread2.Col = 5
              M_Ruta.vaSpread2.text = nompro
              M_Ruta.vaSpread2.Col = 6
              M_Ruta.vaSpread2.text = fecvig
              M_Ruta.vaSpread2.SetActiveCell 2, M_Ruta.vaSpread2.MaxRows
           ElseIf opcion = "regpro" Then
              '-------> Asignar datos a la grilla ruta productos
              M_RegNeg.vaSpread3.MaxRows = M_RegNeg.vaSpread3.MaxRows + 1
              M_RegNeg.vaSpread3.Row = M_RegNeg.vaSpread3.MaxRows
              If EstSel = False Then vg_codigo = M_RegNeg.vaSpread3.MaxRows: EstSel = True
              M_RegNeg.vaSpread3.Col = -1
              M_RegNeg.vaSpread3.BackColor = &H80000013
              M_RegNeg.vaSpread3.Col = 1
              M_RegNeg.vaSpread3.text = "1"
              M_RegNeg.vaSpread3.Col = 2
              M_RegNeg.vaSpread3.text = Trim(Mid(Combo1(1).text, 1, 150))
              M_RegNeg.vaSpread3.Col = 3
              M_RegNeg.vaSpread3.text = codpro
              M_RegNeg.vaSpread3.Col = 4
              M_RegNeg.vaSpread3.text = nompro
              M_RegNeg.vaSpread3.Col = 5
              M_RegNeg.vaSpread3.text = "1"
              M_RegNeg.vaSpread3.Col = 6
              M_RegNeg.vaSpread3.text = "1"
              M_RegNeg.vaSpread3.Col = 7
              M_RegNeg.vaSpread3.text = "1"
              M_RegNeg.vaSpread3.SetActiveCell 2, M_RegNeg.vaSpread3.MaxRows
           ElseIf opcion = "regunpro" Then
              vg_codigo = codpro: vg_nombre = nompro
           End If
        End If
    Next i
    Me.Hide
    Unload Me
Case 3
    vg_codigo = ""
    Me.Hide
    Unload Me
End Select
End Sub

Sub LlenaDatos(codigo As String, titgen As String, op As String, spi As Long)
'-------> cargar familia productos
Me.Caption = IIf(op = "rutpro", "Agregar Productos Ruta : ", "Agregar Productos Reglas de Negocios : ") & titgen
opcion = op
codrut = codigo
spid = spi
Combo1(0).AddItem "Todos" & Space(150) & "(" & fg_pone_cero(Str(0), 10) & ")"
Set RS = vg_dbpedweb.Execute("pedweb_s_centralcompras 1, '', ''")
If Not RS.EOF Then
   Do While Not RS.EOF
      Combo1(0).AddItem IIf(IsNull(RS!CentralDeCompra), "", RS!CentralDeCompra) & Space(150) & "(" & fg_pone_espacio((RS!CentralDeCompra), 10) & ")"
      RS.MoveNext
   Loop
   Combo1(0).ListIndex = 0
End If
RS.Close: Set RS = Nothing
codcen = ""

Set RS = vg_dbpedweb.Execute("pedweb_s_familiaproducto 1, '', ''")
If Not RS.EOF Then
   Do While Not RS.EOF
      Combo1(1).AddItem IIf(IsNull(RS!descripcioncompleta), "", RS!descripcioncompleta) & Space(150) & "(" & fg_pone_cero(Str(RS!codigo), 10) & ")"
      RS.MoveNext
   Loop
   Combo1(1).ListIndex = 0
End If
RS.Close: Set RS = Nothing
codfpr = Val(fg_codigocbo(Combo1, 1, 10, ""))
If op = "rutpro" Then
   CargarProductosnoIncludosRuta
ElseIf op = "regpro" Or op = "regunpro" Then
   CargarReglasdeNegociosProductos
End If
est = False
End Sub

Sub CargarProductosnoIncludosRuta()
fg_carga ""
'-------> cargar datos vector
vaSpread1.MaxRows = 0
vaSpread1.Visible = False
Set RS = vg_dbpedweb.Execute("pedweb_s_prodoctosnoincluidoruta '" & codrut & "', '" & codcen & "', " & codfpr & "")
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   vaSpread1.text = "0"
   vaSpread1.Col = 2
   vaSpread1.text = RS!pce_codpro
   vaSpread1.Col = 3
   vaSpread1.text = RS!pce_codcen
   vaSpread1.Col = 4
   vaSpread1.text = RS!descripcion
   vaSpread1.Col = 5
   vaSpread1.text = RS!vigencia
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
fg_descarga
End Sub

Sub CargarReglasdeNegociosProductos()
'-------> cargar datos vector
fg_carga ""
vaSpread1.MaxRows = 0
vaSpread1.Visible = False
Set RS = vg_dbpedweb.Execute("pedweb_s_prodoctosnoreglasdenegocios '" & codrut & "', '" & codfpr & "', '" & codcen & "', '" & IIf(opcion = "regpro", vg_NUsr, "") & "' , " & spid & "")
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   vaSpread1.text = "0"
   vaSpread1.Col = 2
   vaSpread1.text = RS!codigo
   vaSpread1.Col = 3
   vaSpread1.text = RS!CentralDeCompra
   vaSpread1.Col = 4
   vaSpread1.text = RS!descripcion
   vaSpread1.Col = 5
   vaSpread1.text = RS!vigencia
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
fg_descarga
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
If opcion <> "regunpro" Then
    Dim i As Long
    vaSpread1.Col = 1
    For i = BlockRow To BlockRow2
        vaSpread1.Row = i
        vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    Next i
End If
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If opcion = "regunpro" And ButtonDown = 1 Then
   Dim i As Long
   For i = 1 To vaSpread1.MaxRows
       vaSpread1.Row = i
       vaSpread1.Col = 1
       If i <> Row Then vaSpread1.text = "0"
   Next i
End If
End Sub
