VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form B_MTaEst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   1815
   ClientTop       =   1830
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5985
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   0
      ItemData        =   "B_MTaEst.frx":0000
      Left            =   120
      List            =   "B_MTaEst.frx":0002
      TabIndex        =   8
      Top             =   3600
      Width           =   5115
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1710
         TabIndex        =   3
         Top             =   555
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "B_MTaEst.frx":0004
         Left            =   1710
         List            =   "B_MTaEst.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Buscar Texto"
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
         Left            =   165
         TabIndex        =   5
         Top             =   660
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Buscar Columna"
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
         Left            =   165
         TabIndex        =   4
         Top             =   345
         Width           =   1440
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Index           =   1
      Left            =   720
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2220
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   5415
      _Version        =   393216
      _ExtentX        =   9551
      _ExtentY        =   3916
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
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
      MaxRows         =   13
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "B_MTaEst.frx":0022
      StartingColNumber=   6
      ScrollBarTrack  =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5685
      Left            =   5445
      TabIndex        =   7
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   10028
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   0
      TabIndex        =   9
      Top             =   3360
      Width           =   5415
   End
End
Attribute VB_Name = "B_MTaEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim i As Long, j As Long, ibusca As Long, codigo As String, opc As String, ind As Long
Dim icombo As Integer, est As Boolean, est1 As Boolean
Dim findstring  As String, sourcestring As String
Dim prog As Object

Private Sub Combo1_Click()
If icombo = 0 Then Text1.text = ""
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
fg_centra Me
Me.Left = vg_left
est1 = True
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
icombo = 1
Combo1.ListIndex = 1
End Sub

Sub LlenaDatos(TitGen As String, programa As Object, cencos As String, codreg As String, fecini As Long, fecfin As Long, opcion As String, lc_Aux As String, indgri As Long, tipmin As String)

Dim aAp  As String
Dim sql1 As String
Dim sql2 As String

fg_carga ""
opc = opcion
ind = indgri
Me.Caption = TitGen
Set prog = programa
vaSpread1.MaxRows = 0
List1(0).Clear: List1(1).Clear
'List1 (0) ' 0: List1(1).ListIndex = 0
sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecini) & "') ", " '" & Format(fg_Ctod1(fecini), "yyyymmdd") & "' ")
sql2 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecfin) & "') ", " '" & Format(fg_Ctod1(fecfin), "yyyymmdd") & "' ")

If opcion = "0" Then
   
   If TitGen = "Regimen" Then
      
      RS1.Open "SELECT reg_codigo, reg_nombre FROM a_regimen ORDER BY reg_codigo", vg_db, adOpenStatic
    
    Else
      
      RS1.Open "SELECT ser_codigo, ser_nombre, ser_orden FROM a_servicio WHERE ser_activo = '1' ORDER BY ser_orden, ser_codigo", vg_db, adOpenStatic
    
    End If
    
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fg_descarga: Exit Sub

ElseIf opcion = "1" Then
   
   If TitGen = "Regimen" Then
      
      RS1.Open "SELECT DISTINCT b.min_codreg, a.reg_nombre " & _
               "FROM a_regimen a, b_minuta b, b_minutadet c " & _
               "WHERE b.min_codigo = c.mid_codigo " & _
               "AND   b.min_codreg = a.reg_codigo " & _
               "AND   b.min_cencos = '" & cencos & "' " & _
               "AND   b.min_fecmin >= " & fecini & " " & _
               "AND   b.min_fecmin <= " & fecfin & " " & _
               "AND   c.mid_tipmin IN (" & tipmin & ") " & _
               "ORDER BY a.reg_nombre", vg_db, adOpenStatic
    
    Else
      
      RS1.Open "SELECT DISTINCT b.min_codser, a.ser_nombre, a.ser_orden " & _
               "FROM a_servicio a, b_minuta b, b_minutadet c " & _
               "WHERE b.min_codigo = c.mid_codigo " & _
               "AND   b.min_codser = a.ser_codigo " & _
               "AND   b.min_cencos = '" & cencos & "' " & _
               "AND   b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
               "AND   b.min_fecmin >= " & fecini & " " & _
               "AND   b.min_fecmin <= " & fecfin & " " & _
               "AND   c.mid_tipmin IN (" & tipmin & ") AND a.ser_activo = '1' " & _
               "ORDER BY a.ser_orden, a.ser_nombre", vg_db, adOpenStatic
    
    End If
    
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fg_descarga: Exit Sub

ElseIf opcion = "2" Then
   
   RS1.Open "SELECT nut_codigo, nut_nombre, nut_indpri, nut_secnro FROM a_nutriente ORDER BY nut_secnro", vg_db, adOpenStatic

ElseIf opcion = "3" Then
   
   RS1.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_tipo = 1 AND cli_activo = '1' ORDER BY cli_codigo", vg_db, adOpenStatic

ElseIf opcion = "4" Then
   
   If TitGen = "Regimen" Then
      
      RS1.Open "SELECT DISTINCT a.reg_codigo, a.reg_nombre " & _
               "FROM a_regimen a, b_minuta b, b_minutadet c " & _
               "WHERE b.min_codigo = c.mid_codigo " & _
               "AND   b.min_codreg = a.reg_codigo " & _
               "AND   b.min_cencos = '" & cencos & "' " & _
               "AND   b.min_fecmin >= " & fecini & " " & _
               "AND   b.min_fecmin <= " & fecfin & " " & _
               "AND   c.mid_tipmin = '2' " & _
               "ORDER BY a.reg_nombre", vg_db, adOpenStatic
   
   Else
      
      RS1.Open "SELECT DISTINCT a.ser_codigo, a.ser_nombre, a.ser_orden " & _
               "FROM a_servicio a, b_minuta b, b_minutadet c " & _
               "WHERE b.min_codigo = c.mid_codigo " & _
               "AND   b.min_codser = a.ser_codigo " & _
               "AND   b.min_cencos = '" & cencos & "' " & _
               "AND   b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
               "AND   b.min_fecmin >= " & fecini & " " & _
               "AND   b.min_fecmin <= " & fecfin & " " & _
               "AND   c.mid_tipmin = '2' AND a.ser_activo='1' " & _
               "ORDER BY a.ser_orden, a.ser_nombre", vg_db, adOpenStatic
   
   End If

ElseIf opcion = "5" Then
   
   'Crea tabla temporal
   aAp = Trim(vg_NUsr) & "_tmp_" & lc_Aux
   fg_CheckTmp aAp
   If TitGen = "Regimen" Then
      
      RS1.Open "SELECT DISTINCT a.tov_codreg AS min_codreg, b.reg_nombre INTO " & aAp & " " & _
               "FROM b_totventas a, a_regimen b " & _
               "WHERE a.tov_codreg = b.reg_codigo " & _
               "AND   a.tov_rutcli = '" & cencos & "' " & _
               "AND   a.tov_codbod = " & vg_codbod & " AND a.tov_tipdoc= 'SP' " & _
               "AND   a.tov_fecpro >= " & sql1 & " " & _
               "AND   a.tov_fecpro <= " & sql2 & " " & _
               "ORDER BY b.reg_nombre", vg_db, adOpenStatic

      vg_db.Execute "INSERT INTO " & aAp & " SELECT DISTINCT b.min_codreg AS min_codreg, a.reg_nombre AS reg_nombre " & _
                    "FROM a_regimen a, b_minuta b, b_minutadet c " & _
                    "WHERE b.min_codigo = c.mid_codigo " & _
                    "AND   b.min_codreg = a.reg_codigo " & _
                    "AND   b.min_cencos = '" & cencos & "' " & _
                    "AND   b.min_fecmin >= " & fecini & " " & _
                    "AND   b.min_fecmin <= " & fecfin & " " & _
                    "AND   c.mid_tipmin IN ('1','2') " & _
                    "ORDER BY a.reg_nombre"
   Else
      
      RS1.Open "SELECT DISTINCT a.tov_codser AS min_codser, b.ser_nombre, b.ser_orden INTO " & aAp & " " & _
               "FROM b_totventas a, a_servicio b " & _
               "WHERE a.tov_codser = b.ser_codigo " & _
               "AND   a.tov_rutcli = '" & cencos & "' " & _
               "AND   a.tov_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
               "AND   a.tov_codbod =" & vg_codbod & " AND a.tov_tipdoc= 'SP' " & _
               "AND   a.tov_fecpro >= " & sql1 & " " & _
               "AND   a.tov_fecpro <= " & sql2 & " AND b.ser_activo = '1' " & _
               "ORDER BY b.ser_orden, b.ser_nombre", vg_db, adOpenStatic

      vg_db.Execute "INSERT INTO " & aAp & " SELECT DISTINCT b.min_codser AS min_codser, a.ser_nombre AS ser_nombre, a.ser_orden AS ser_orden " & _
                    "FROM a_servicio a, b_minuta b, b_minutadet c " & _
                    "WHERE b.min_codigo = c.mid_codigo " & _
                    "AND   b.min_codser = a.ser_codigo " & _
                    "AND   b.min_cencos = '" & cencos & "' " & _
                    "AND   b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                    "AND   b.min_fecmin >= " & fecini & " " & _
                    "AND   b.min_fecmin <= " & fecfin & " " & _
                    "AND   c.mid_tipmin IN ('1','2') " & _
                    "ORDER BY a.ser_orden, a.ser_nombre"
   End If
   Set RS1 = Nothing
   RS1.Open "SELECT DISTINCT * FROM " & aAp & "", vg_db, adOpenStatic

ElseIf opcion = "6" Then
   
   If TitGen = "Regimen" Then
      
      RS1.Open "SELECT DISTINCT a.tov_codreg AS min_codreg, b.reg_nombre " & _
               "FROM b_totventas a, a_regimen b " & _
               "WHERE a.tov_codreg = b.reg_codigo " & _
               "AND   a.tov_rutcli = '" & cencos & "' " & _
               "AND   a.tov_codbod = " & vg_codbod & " AND a.tov_tipdoc = 'SP' " & _
               "AND   a.tov_fecpro >= " & sql1 & " " & _
               "AND   a.tov_fecpro <= " & sql2 & " " & _
               "ORDER BY b.reg_nombre", vg_db, adOpenStatic
   
   Else
      
      RS1.Open "SELECT DISTINCT a.tov_codser AS min_codser, b.ser_nombre, b.ser_orden " & _
               "FROM b_totventas a, a_servicio b " & _
               "WHERE a.tov_codser = b.ser_codigo " & _
               "AND   a.tov_rutcli = '" & cencos & "' " & _
               "AND   a.tov_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
               "AND   a.tov_codbod = " & vg_codbod & " AND a.tov_tipdoc = 'SP' " & _
               "AND   a.tov_fecpro >= " & sql1 & " " & _
               "AND   a.tov_fecpro <= " & sql2 & " AND b.ser_activo = '1' " & _
               "ORDER BY b.ser_orden, b.ser_nombre", vg_db, adOpenStatic
   
   End If

End If

Do While Not RS1.EOF
   
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   
   programa(ind).Row = vaSpread1.MaxRows
   programa(ind).Col = 1
   
   If programa(ind).text = "1" Then
      
      vaSpread1.Col = 1
      vaSpread1.CellType = 10
      vaSpread1.TypeCheckText = ""
      vaSpread1.TypeCheckCenter = True
      vaSpread1.text = "1" ' checked
      List1(0).AddItem RS1(0) & Space(5) & RS1(1)
      List1(1).AddItem RS1(0) & RS1(1) & Space(150) & "(" & fg_pone_espacio(CStr(RS1(0)), 10) & ")" '& RS1(0)
   
   ElseIf programa(ind).text = "" Then
         
      vaSpread1.Col = 1
      vaSpread1.CellType = 10
      vaSpread1.TypeCheckText = ""
      vaSpread1.TypeCheckCenter = True
      vaSpread1.text = "" ' checked
   
   End If
   
   vaSpread1.Col = 2
   vaSpread1.text = IIf(opcion = "3", IIf(Len(RS1(0)) < 10, RS1(0) & Space(10 - Len(CStr(RS1(0)))), RS1(0)), RS1(0))
   vaSpread1.Col = 3
   vaSpread1.text = RS1(1)
   
   RS1.MoveNext

Loop

RS1.Close
Set RS1 = Nothing
fg_descarga
icombo = 0

End Sub

Private Sub List1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 27
    Cerrar
End Select
End Sub

Private Sub Text1_Change()
If vaSpread1.MaxRows < 1 Then Exit Sub
If LimpiaDato(Trim(Text1.text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
   findstring = Text1.text
   For i = 1 To vaSpread1.MaxRows
       vaSpread1.Row = i
       vaSpread1.Col = 2
       sourcestring = Trim(vaSpread1.text)
       IndActivo = UCase(Trim(sourcestring)) Like "*" & UCase(findstring) & "*"
       If IndActivo = -1 Then
          If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Else
          If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
       End If
   Next i
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
   findstring = Text1.text
   For i = 1 To vaSpread1.MaxRows
       vaSpread1.Row = i
       vaSpread1.Col = 3
       sourcestring = Trim(vaSpread1.text)
       IndActivo = UCase(Trim(sourcestring)) Like "*" & UCase(findstring) & "*"
       If IndActivo = -1 Then
          If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Else
          If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
       End If
   Next i
End If
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
    Cerrar
End Select
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If icombo = 1 Or Row = 0 Then Exit Sub
Select Case Col
Case 1
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    
    If vaSpread1.Value = "1" Then ' checked
       
       vaSpread1.Col = 2: codigo = vaSpread1.text
       vaSpread1.Col = 3
       List1(0).AddItem codigo & Space(5) & Trim(vaSpread1.text)
       List1(1).AddItem codigo & vaSpread1.text & Space(150) & "(" & fg_pone_espacio(CStr(codigo), 10) & ")" '& codigo
    
    ElseIf vaSpread1.Value = "0" Then
       
       vaSpread1.Col = 1
       vaSpread1.CellType = 10
       vaSpread1.TypeCheckText = ""
       vaSpread1.TypeCheckCenter = True
       vaSpread1.Value = "0" ' checked
       
       For i = 0 To List1(1).listcount - 1
           
           vaSpread1.Col = 2
           codigo = vaSpread1.text
           
           If Trim(vaSpread1.text) = Trim(Mid((List1(1).List(i)), Len((List1(1).List(i))) - 10, 10)) Then
              
              List1(1).RemoveItem i
              List1(0).RemoveItem i
              Exit Sub
           
           
           End If
       Next i
    
    End If
End Select
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
If Row = 0 Then Exit Sub
MoverDatos
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If Col = 1 And Row = 0 Then
   est = True
   icombo = 1
   vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.Value = IIf(vaSpread1.Value = "1", "", "1")
   est = IIf(vaSpread1.Value = "1", True, False)
   If Not est Then
      est1 = False
      List1(0).Clear: List1(1).Clear
   Else
      If est1 Then icombo = 0: Exit Sub
      For i = 1 To vaSpread1.MaxRows
          vaSpread1.Row = i
          vaSpread1.Col = 2
          codigo = vaSpread1.text
          vaSpread1.Col = 3
          List1(0).AddItem codigo & Space(5) & Trim(vaSpread1.text)
          List1(1).AddItem codigo & vaSpread1.text & Space(150) & "(" & fg_pone_espacio(CStr(codigo), 10) & ")" '& codigo
      Next i
   End If
   icombo = 0
End If
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

Sub MoverDatos()
If vaSpread1.MaxRows < 1 Then Exit Sub
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1
    If vaSpread1.text = "1" Then
       prog(ind).Row = i
       prog(ind).Col = 1
       prog(ind).CellType = 10
       prog(ind).TypeCheckText = ""
       prog(ind).TypeCheckCenter = True
       prog(ind).text = "1" ' checked
    ElseIf vaSpread1.text = "" Or vaSpread1.text = "0" Then
       prog(ind).Row = i
       prog(ind).Col = 1
       prog(ind).CellType = 10
       prog(ind).TypeCheckText = ""
       prog(ind).TypeCheckCenter = True
       prog(ind).text = "" ' checked
    End If
Next i
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
vg_codigo = vaSpread1.text
Cerrar
End Sub

Sub Cerrar()
Me.Hide
Unload Me
End Sub
