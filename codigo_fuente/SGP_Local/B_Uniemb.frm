VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form B_Uniemb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingrediente"
   ClientHeight    =   4815
   ClientLeft      =   2520
   ClientTop       =   1530
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4540
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "B_Uniemb.frx":0000
         Left            =   1680
         List            =   "B_Uniemb.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   555
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Buscar Columna"
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
         Left            =   225
         TabIndex        =   4
         Top             =   345
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Buscar Texto"
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
         Left            =   225
         TabIndex        =   3
         Top             =   660
         Width           =   1080
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3615
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   4545
      _Version        =   393216
      _ExtentX        =   8017
      _ExtentY        =   6376
      _StockProps     =   64
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
      MaxCols         =   2
      MaxRows         =   30
      OperationMode   =   3
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "B_Uniemb.frx":001E
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   4815
      Left            =   4545
      TabIndex        =   6
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   8493
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "B_Uniemb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Consql1 As ADODB.Recordset
Dim i As Long, ibusca As Long
Dim icombo As Integer
Private Sub Combo1_Click()
If icombo = 0 Then
   Text1.Text = ""
End If
End Sub
Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 27
    Cerrar
End Select
End Sub
Private Sub Form_Activate()
fg_descarga
End Sub
Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me
If vg_opcioningrediente = 1 Then B_CamIng.Left = vg_left
fg_carga (ss)
Mover_Botones
icombo = 1
Combo1.ListIndex = 1
vaSpread1.MaxRows = 0
If vg_opcioning = 1 Then
   Set Consql1 = vg_db.Execute("select PB00080.Ing_No, PB00080.Ltst_Price, " & _
                 "PB00080.Uom_Code_No, PB00080.Grnsh_Ind, PB00354.Uom_Name, " & _
                 "PB00080.Del_Ind , PB00081.Ing_Desc " & _
                 "From PB00080, PB00081, PB00354 " & _
                 "Where PB00080.Ing_No = PB00081.Ing_No " & _
                 "and   PB00080.Uom_Code_No = PB00354.Uom_Code " & _
                 "and   PB00080.Del_Ind =0 " & _
                 "and   PB00081.Del_Ind =0 " & _
                 "order by PB00081.Ing_Desc", , adCmdText)
'   Set Consql1 = vg_db.Execute("sod_s_ingrediente 1, 0, ''", , adCmdStoredProc)
ElseIf vg_opcioning = 2 Then
   Set Consql1 = vg_db.Execute("sod_s_tablagringrediente 1, '" & vg_auxgcodcasino & "', " & vg_auxgcodregimen & ", " & vg_auxgcategoria1 & ", " & vg_auxgcategoria2 & ", " & vg_auxgcategoria3 & ", " & vg_auxgcategoria4 & ", '" & vg_auxgmes & "', '" & vg_auxgano & "', '', 0", , adCmdStoredProc)
ElseIf vg_opcioning = 3 Then
   Set Consql1 = vg_db.Execute("sod_s_tablagrseringrediente 1, '" & vg_auxgcodcasino & "', " & vg_auxgcodpventa & ", " & vg_auxgcodservicio & ", " & vg_auxgcodregimen & ", " & vg_auxgcategoria1 & ", " & vg_auxgcategoria2 & ", " & vg_auxgcategoria3 & ", " & vg_auxgcategoria4 & ", '" & vg_auxgmes & "', '" & vg_auxgano & "', '', 0", , adCmdStoredProc)
End If
If Not Consql1.EOF Then
   Do While Not Consql1.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
     
      vaSpread1.Col = 1
      vaSpread1.CellType = 5
      vaSpread1.TypeHAlign = 1
      vaSpread1.Text = Consql1!Ing_No
      
      vaSpread1.Col = 2
      vaSpread1.CellType = 5
      vaSpread1.TypeHAlign = 0
      vaSpread1.Text = Trim(Consql1!Ing_Desc)
      Consql1.MoveNext
   Loop
End If
Consql1.Close: Set Consql1 = Nothing
icombo = 0
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Ingrediente"
End Sub
Private Sub Text1_Change()
If LimpiaDato(Trim(Text1.Text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
   If vg_opcioning = 1 Then
      Set Consql1 = vg_db.Execute("sod_s_ingrediente 3, 0, '%" + UCase(LimpiaDato(Text1.Text)) + "%'", , adCmdStoredProc)
      If Consql1.EOF Or Consql1!nreg = 0 Then Consql1.Close: Set Consql1 = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: Exit Sub
      If ibusca <> Consql1!nreg Then ibusca = Consql1!nreg: vaSpread1.MaxRows = Consql1!nreg: Consql1.Close: Set Consql1 = Nothing
      Set Consql1 = vg_db.Execute("sod_s_ingrediente 2, 0, '%" + UCase(LimpiaDato(Text1.Text)) + "%'", , adCmdStoredProc)
   ElseIf vg_opcioning = 2 Then
      Set Consql1 = vg_db.Execute("sod_s_tablagringrediente 3, '" & vg_auxgcodcasino & "', " & vg_auxgcodregimen & ", " & vg_auxgcategoria1 & ", " & vg_auxgcategoria2 & ", " & vg_auxgcategoria3 & ", " & vg_auxgcategoria4 & ", '" & vg_auxgmes & "', '" & vg_auxgano & "', '%" + UCase(LimpiaDato(Text1.Text)) + "%', 0", , adCmdStoredProc)
      If Consql1.EOF Or Consql1!nreg = 0 Then Consql1.Close: Set Consql1 = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: Exit Sub
      If ibusca <> Consql1!nreg Then ibusca = Consql1!nreg: vaSpread1.MaxRows = Consql1!nreg: Consql1.Close: Set Consql1 = Nothing
      Set Consql1 = vg_db.Execute("sod_s_tablagringrediente 2, '" & vg_auxgcodcasino & "', " & vg_auxgcodregimen & ", " & vg_auxgcategoria1 & ", " & vg_auxgcategoria2 & ", " & vg_auxgcategoria3 & ", " & vg_auxgcategoria4 & ", '" & vg_auxgmes & "', '" & vg_auxgano & "', '%" + UCase(LimpiaDato(Text1.Text)) + "%', 0", , adCmdStoredProc)
   ElseIf vg_opcioning = 3 Then
      Set Consql1 = vg_db.Execute("sod_s_tablagrseringrediente 3, '" & vg_auxgcodcasino & "', " & vg_auxgcodpventa & ", " & vg_auxgcodservicio & ", " & vg_auxgcodregimen & ", " & vg_auxgcategoria1 & ", " & vg_auxgcategoria2 & ", " & vg_auxgcategoria3 & ", " & vg_auxgcategoria4 & ", '" & vg_auxgmes & "', '" & vg_auxgano & "', '%" + UCase(LimpiaDato(Text1.Text)) + "%', 0", , adCmdStoredProc)
      If Consql1.EOF Or Consql1!nreg = 0 Then Consql1.Close: Set Consql1 = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: Exit Sub
      If ibusca <> Consql1!nreg Then ibusca = Consql1!nreg: vaSpread1.MaxRows = Consql1!nreg: Consql1.Close: Set Consql1 = Nothing
      Set Consql1 = vg_db.Execute("sod_s_tablagrseringrediente 2, '" & vg_auxgcodcasino & "', " & vg_auxgcodpventa & ", " & vg_auxgcodservicio & ", " & vg_auxgcodregimen & ", " & vg_auxgcategoria1 & ", " & vg_auxgcategoria2 & ", " & vg_auxgcategoria3 & ", " & vg_auxgcategoria4 & ", '" & vg_auxgmes & "', '" & vg_auxgano & "', '%" + UCase(LimpiaDato(Text1.Text)) + "%', 0", , adCmdStoredProc)
   End If
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
   If vg_opcioning = 1 Then
      Set Consql1 = vg_db.Execute("sod_s_ingrediente 5, 0, '%" + UCase(LimpiaDato(Text1.Text)) + "%'", , adCmdStoredProc)
      If Consql1.EOF Or Consql1!nreg = 0 Then Consql1.Close: Set Consql1 = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: Exit Sub
      If ibusca <> Consql1!nreg Then ibusca = Consql1!nreg: vaSpread1.MaxRows = Consql1!nreg: Consql1.Close: Set Consql1 = Nothing
      Set Consql1 = vg_db.Execute("sod_s_ingrediente 4, 0, '%" + UCase(LimpiaDato(Text1.Text)) + "%'", , adCmdStoredProc)
   ElseIf vg_opcioning = 2 Then
      Set Consql1 = vg_db.Execute("sod_s_tablagringrediente 5, '" & vg_auxgcodcasino & "', " & vg_auxgcodregimen & ", " & vg_auxgcategoria1 & ", " & vg_auxgcategoria2 & ", " & vg_auxgcategoria3 & ", " & vg_auxgcategoria4 & ", '" & vg_auxgmes & "', '" & vg_auxgano & "', '%" + UCase(LimpiaDato(Text1.Text)) + "%', 0", , adCmdStoredProc)
      If Consql1.EOF Or Consql1!nreg = 0 Then Consql1.Close: Set Consql1 = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: Exit Sub
      If ibusca <> Consql1!nreg Then ibusca = Consql1!nreg: vaSpread1.MaxRows = Consql1!nreg: Consql1.Close: Set Consql1 = Nothing
      Set Consql1 = vg_db.Execute("sod_s_tablagringrediente 4, '" & vg_auxgcodcasino & "', " & vg_auxgcodregimen & ", " & vg_auxgcategoria1 & ", " & vg_auxgcategoria2 & ", " & vg_auxgcategoria3 & ", " & vg_auxgcategoria4 & ", '" & vg_auxgmes & "', '" & vg_auxgano & "', '%" + UCase(LimpiaDato(Text1.Text)) + "%', 0", , adCmdStoredProc)
   ElseIf vg_opcioning = 3 Then
      Set Consql1 = vg_db.Execute("sod_s_tablagrseringrediente 5, '" & vg_auxgcodcasino & "', " & vg_auxgcodpventa & ", " & vg_auxgcodservicio & ", " & vg_auxgcodregimen & ", " & vg_auxgcategoria1 & ", " & vg_auxgcategoria2 & ", " & vg_auxgcategoria3 & ", " & vg_auxgcategoria4 & ", '" & vg_auxgmes & "', '" & vg_auxgano & "', '%" + UCase(LimpiaDato(Text1.Text)) + "%', 0", , adCmdStoredProc)
      If Consql1.EOF Or Consql1!nreg = 0 Then Consql1.Close: Set Consql1 = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: Exit Sub
      If ibusca <> Consql1!nreg Then ibusca = Consql1!nreg: vaSpread1.MaxRows = Consql1!nreg: Consql1.Close: Set Consql1 = Nothing
      Set Consql1 = vg_db.Execute("sod_s_tablagrseringrediente 4, '" & vg_auxgcodcasino & "', " & vg_auxgcodpventa & ", " & vg_auxgcodservicio & ", " & vg_auxgcodregimen & ", " & vg_auxgcategoria1 & ", " & vg_auxgcategoria2 & ", " & vg_auxgcategoria3 & ", " & vg_auxgcategoria4 & ", '" & vg_auxgmes & "', '" & vg_auxgano & "', '%" + UCase(LimpiaDato(Text1.Text)) + "%', 0", , adCmdStoredProc)
   End If
End If
i = 1
If Not Consql1.EOF Then
   Do While Not Consql1.EOF
      vaSpread1.Row = i
     
      vaSpread1.Col = 1
      vaSpread1.CellType = 5
      vaSpread1.TypeHAlign = 1
      vaSpread1.Text = Consql1!Ing_No
      
      vaSpread1.Col = 2
      vaSpread1.CellType = 5
      vaSpread1.TypeHAlign = 0
      vaSpread1.Text = Trim(Consql1!Ing_Desc)
      
      i = i + 1
      Consql1.MoveNext
   Loop
End If
Consql1.Close: Set Consql1 = Nothing
vaSpread1.SetActiveCell 1, 1
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 27
    Cerrar
End Select
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 1
    MoverDatos
  Case 3
    vg_codigo = 0
    vg_nombre = ""
    Cerrar
End Select
End Sub
Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
MoverDatos
End Sub
Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
MoverDatos
End Sub
Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 27
    Cerrar
End Select
End Sub
Private Sub MoverDatos()
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
vg_codigo = Val(vaSpread1.Text)
vaSpread1.Col = 2
vg_nombre = vaSpread1.Text
Cerrar
End Sub
Sub Cerrar()
Me.Hide
'Unload B_Ingred
End Sub
Sub Mover_Botones()

   Toolbar1.ImageList = Partida.IL1
   Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar "
   Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
   Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"

End Sub

