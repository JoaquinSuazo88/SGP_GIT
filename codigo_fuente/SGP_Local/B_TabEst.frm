VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form B_TabEst 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4815
   ClientLeft      =   3645
   ClientTop       =   3270
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5730
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5160
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   555
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "B_TabEst.frx":0000
         Left            =   1680
         List            =   "B_TabEst.frx":0002
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
         Left            =   105
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
         Left            =   90
         TabIndex        =   4
         Top             =   345
         Width           =   1440
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3615
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   5160
      _Version        =   393216
      _ExtentX        =   9102
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
      SelectBlockOptions=   0
      SpreadDesigner  =   "B_TabEst.frx":0004
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   4815
      Left            =   5190
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
Attribute VB_Name = "B_TabEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim Tabla As String, Suf As String, Opx As String, Titulo As String
Dim icombo As Integer, est As Boolean

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
If Trim(ws_respuesta) <> "" Then Text1.text = ws_respuesta: Text1.SelStart = Len(ws_respuesta): ws_respuesta = ""
End Sub

Private Sub Form_Load()
On Error GoTo Man_Error
fg_centra Me
fg_carga ""
icombo = 1
est = True
With Combo1
    .Clear
    .AddItem "Codigo"
    .AddItem "Nombre"
    .ListIndex = 1
End With
'-------> LlenaDatos
Toolbar1.ImageList = Partida.IL1
Toolbar1.Buttons.Clear
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
icombo = 0
fg_descarga
Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, Titulo
End Sub

Private Sub Text1_Change()
Dim i As Long, indactivo As Variant, indvec As Long, var As Long, indact As Boolean
indact = False
indvec = IIf(Combo1.ListIndex = 0, 1, 2)
With vaSpread1
    .Visible = False
    If Trim(Text1.text) <> "" Then
       For i = 1 To .MaxRows
           .Row = i
           .Col = indvec
           indactivo = UCase(Trim(.Value)) Like "*" & UCase(Text1.text) & "*"
           .Col = 1
           If indactivo = -1 And Trim(.text) <> "" Then
              If Not indact Then .OperationMode = 2: .Action = 0: indact = True
              If .RowHidden = True Then .RowHidden = False
           Else
              If .RowHidden = False Then .RowHidden = True
           End If
       Next i
    '   .SetActiveCell indvec, 1
    End If
    .ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    .ColUserSortIndicator(IIf(Trim(Text1.text) = "", 0, 0)) = ColUserSortIndicatorAscending
    .SortKey(1) = IIf(Trim(Text1.text) = "", 0, 0): .SortKeyOrder(1) = SortKeyOrderAscending
    .Sort -1, -1, .MaxCols, .MaxRows, SortByRow
    If Trim(Text1.text) = "" Then
       For i = 1 To .MaxRows
           .Row = i
           If .RowHidden = True Then .RowHidden = False
       Next
       .SetActiveCell Index, .SearchCol(indvec, 0, .MaxRows, Trim(Text1.text), SearchFlagsGreaterOrEqual)
       .SetActiveCell indvec, 1
    End If
    .Visible = True
End With

'Dim z As Long, i As Long, var As Long
'On Error GoTo Man_Error
'If LimpiaDato(Trim(Text1.text)) & Chr(KeyAscii) = "" Then Exit Sub
'If Combo1.ListIndex = 0 Then
'    If Opx = "Gen" Then
'        RS1.Open "SELECT " & Suf & "codigo,  " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "Ser" Then
'        RS1.Open "SELECT " & Suf & "codigo,  " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "activo='1' AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "Gpr" Then
'        RS1.Open "SELECT DISTINCT " & Suf & "codigo,  " & Suf & "nombre FROM " & Tabla & " a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo=c.cli_codtis OR " & Suf & "maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=" & Suf & "maepro OR " & Suf & "maepro<1) AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "Tpc" Then '-------> Tabla lista precio
'        RS1.Open "SELECT " & Suf & "codigo,  " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "cencos='" & MuestraCasino(1) & "' AND " & Suf & "activo='1' AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "Bod" Then '-------> Filtrar codigo bodega
'        RS1.Open "SELECT a.* FROM a_bodega a, b_clientes b WHERE a.bod_codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND (a.bod_codigo NOT IN (SELECT DISTINCT cli_codbod FROM b_clientes WHERE (cli_codbod) Is Not Null  AND cli_codbod>0)  OR a.bod_codigo=b.cli_codbod) AND b.cli_codigo='" & Trim(Suf) & "'", vg_db, adOpenStatic
'    ElseIf Opx = "Pac" Then '-------> Tabla paciente
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre, " & Suf & "appaterno, " & Suf & "apmaterno FROM " & Tabla & " WHERE " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND " & Suf & "estado='0' ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf Opx = "Reg" Then '-------> Tabla regimen con filtro
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND " & Suf & "codigo=" & vg_codregimen & " ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf Opx = "Usu" Then  'Tabla usuario grupo paciente
'        RS1.Open "SELECT DISTINCT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " a, b_usuariogrupopac b, b_pacientes c WHERE " & Suf & "codigo=b.ugp_codusu AND b.ugp_codgrp=c.pac_codgrp AND c.pac_codigo='" & vg_Aux & "' AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf Opx = "ProVig" Then
'        RS1.Open "SELECT DISTINCT " & Suf & "codigo,  " & Suf & "nombre FROM " & Tabla & " a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo=c.cli_codtis OR " & Suf & "maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=" & Suf & "maepro OR " & Suf & "maepro<1) AND (" & Suf & "fecven>" & Format(Date, "yyyymmdd") & " OR " & Suf & "fecven <= 0) AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "ProInv" Or Opx = "ProInv1" Then
'         RS1.Open "SELECT DISTINCT pro.pro_codigo, pro.pro_nombre FROM b_productos pro, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo=c.cli_codtis OR pro.pro_maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=pro.pro_maepro OR pro.pro_maepro<1) AND pro.pro_codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND pro.pro_ctrsto=1 AND (pro.pro_fecven>" & Format(Date, "yyyymmdd") & " OR pro.pro_fecven <= 0) ORDER BY pro.pro_nombre UNION SELECT DISTINCT pro_codigo, pro_nombre FROM b_productos pro, b_bodegas bod WHERE pro.pro_codigo=bod.bod_codpro AND pro_codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND bod.bod_codbod=" & vg_codbod & " AND bod.bod_canmer>0 ORDER BY pro.pro_nombre", vg_db, adOpenStatic
'    ElseIf Opx = "ProInvNoStock" Then  'Tabla generica
'         RS1.Open "SELECT DISTINCT pro.pro_codigo, pro.pro_nombre FROM b_productos pro, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo=c.cli_codtis OR pro.pro_maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=pro.pro_maepro OR pro.pro_maepro<1) AND pro.pro_codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND (pro.pro_fecven>" & Format(Date, "yyyymmdd") & " OR pro.pro_fecven <= 0) ORDER BY pro.pro_nombre UNION SELECT DISTINCT pro_codigo, pro_nombre FROM b_productos pro, b_bodegas bod WHERE pro.pro_codigo=bod.bod_codpro AND UCASE(pro_nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND bod.bod_codbod=" & vg_codbod & " AND bod.bod_canmer>0 ORDER BY pro.pro_nombre", vg_db, adOpenStatic
'    ElseIf Opx = "ProGrl" Then
'        RS1.Open "SELECT DISTINCT " & Suf & "codigo,  " & Suf & "nombre FROM " & Tabla & " a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo=c.cli_codtis OR " & Suf & "maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=" & Suf & "maepro OR " & Suf & "maepro<1) AND (" & Suf & "fecven>" & Format(Date, "yyyymmdd") & " OR " & Suf & "fecven <= 0) AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND (" & Suf & "ctacon IN ('" & fg_CambiaChar(GetParametro("ctagastos"), ";", "','") & "') OR " & Suf & "ctacon IN ('" & fg_CambiaChar(GetParametro("ctagastos2"), ";", "','") & "'))", vg_db, adOpenStatic
'    ElseIf Opx = "0" Or Opx = "1" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE (" & Suf & "tiprec='" & Opx & "' or '" & Opx & "'='') AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf Opx = "Contrato" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=0 AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
''    ElseIf Opx = "Traspaso" Then
'    ElseIf Mid(Opx, 1, 8) = "Traspaso" Then
''        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=2 AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.Text)) & "%' ", vg_db, adOpenStatic
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "codigo<>'" & Trim(Mid(Opx, 9, 10)) & "' AND (" & Suf & "tipo=2 OR " & Suf & "tipo=0) AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "Sucursal" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "direccion FROM " & Tabla & " WHERE " & Suf & "codcli='" & vg_codigo & "' AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "Cliente" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=1 AND " & Suf & "activo='1' AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "CliSap" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=1 AND " & Suf & "clisap='1' AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "CliAlum" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=3 AND ucase(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "Proing" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & ", b_productosing WHERE " & Suf & "codigo=b_productosing.pri_coding AND b_productosing.pri_codpro='" & vg_codigo & "' AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "Pst" Then 'Tabla productos controla stock
'        RS1.Open "SELECT DISTINCT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo=c.cli_codtis OR " & Suf & "maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=" & Suf & "maepro OR " & Suf & "maepro<1) AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND (" & Suf & "fecven>" & Format(Date, "yyyymmdd") & " OR " & Suf & "fecven <= 0) AND " & Suf & "ctrsto=1 ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf Opx = "Pbo" Then 'Tabla productos Bodega
'        RS1.Open "SELECT " & Tabla & "." & Suf & "codigo, " & Tabla & "." & Suf & "nombre FROM " & Tabla & ", b_bodegas WHERE " & Tabla & "." & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND " & Tabla & "." & Suf & "ctrsto=1 AND " & Tabla & "." & Suf & "codigo=b_bodegas.bod_codpro AND ROUND(b_bodegas.bod_canmer,2)>0 AND b_bodegas.bod_codbod=" & vg_codbod & " ORDER BY " & Tabla & "." & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf Opx = "No5etapas" Then
'        RS1.Open "SELECT " & Suf & "codigo,  " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "codigo<10000 AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "5etapas" Then
'        RS1.Open "SELECT " & Suf & "codigo,  " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "codigo = " & vg_tiprec & " AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf vg_modrec = False Then
'        RS1.Open "SELECT DISTINCT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & ", b_productos,  b_productosing, a_tipopro " & _
'                 "WHERE  " & Suf & "codigo=b_productosing.pri_coding " & _
'                 "AND    b_productosing.pri_codpro=b_productos.pro_codigo " & _
'                 "AND    b_productos.pro_codtip=a_tipopro.tip_codigo " & _
'                 "AND    trim(str(a_tipopro.tip_codigo)) IN ('" & Opx & "') " & _
'                 "AND    UCASE(" & Suf & "codigo) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' " & _
'                 "ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    End If
'ElseIf Combo1.ListIndex = 1 Then
'    If Opx = "Gen" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf Opx = "Ser" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "activo='1' AND UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf Opx = "Gpr" Then
'        RS1.Open "SELECT DISTINCT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo=c.cli_codtis OR " & Suf & "maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=" & Suf & "maepro OR " & Suf & "maepro<1) AND UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf Opx = "Tpc" Then '-------> Tabla lista precio
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "cencos='" & MuestraCasino(1) & "' AND " & Suf & "activo='1' AND UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf Opx = "Bod" Then '-------> Filtrar codigo bodega
'        RS1.Open "SELECT a.* FROM a_bodega a, b_clientes b WHERE UCASE(a.bod_nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND (a.bod_codigo NOT IN (SELECT DISTINCT cli_codbod FROM b_clientes WHERE (cli_codbod) Is Not Null  AND cli_codbod>0)  OR a.bod_codigo=b.cli_codbod) AND b.cli_codigo='" & Trim(Suf) & "'", vg_db, adOpenStatic
'    ElseIf Opx = "Pac" Then '-------> Tabla paciente
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre, " & Suf & "appaterno, " & Suf & "apmaterno FROM " & Tabla & " WHERE (UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' OR UCASE(" & Suf & "appaterno) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' OR UCASE(" & Suf & "apmaterno) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%') AND " & Suf & "estado='0' ORDER BY " & Suf & "nombre", vg_db, adOpenStatic
'    ElseIf Opx = "Reg" Then '-------> Tabla regimen con filtro
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "codigo=" & vg_codregimen & "  AND UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "nombre", vg_db, adOpenStatic
'    ElseIf Opx = "Usu" Then '-------> Tabla usuario grupo paciente
'        RS1.Open "SELECT DISTINCT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " a, b_usuariogrupopac b, b_pacientes c WHERE " & Suf & "codigo=b.ugp_codusu AND b.ugp_codgrp=c.pac_codgrp AND c.pac_codigo='" & vg_Aux & "' AND UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "nombre", vg_db, adOpenStatic
'    ElseIf Opx = "ProVig" Then
'        RS1.Open "SELECT DISTINCT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo=c.cli_codtis OR " & Suf & "maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=" & Suf & "maepro OR " & Suf & "maepro<1) AND (" & Suf & "fecven>" & Format(Date, "yyyymmdd") & " OR " & Suf & "fecven <= 0) AND UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf Opx = "ProInv" Or Opx = "ProInv1" Then
'         RS1.Open "SELECT DISTINCT pro.pro_codigo, pro.pro_nombre FROM b_productos pro, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo=c.cli_codtis OR pro.pro_maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=pro.pro_maepro OR pro.pro_maepro<1) AND UCASE(pro.pro_nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND pro.pro_ctrsto=1 AND (pro.pro_fecven>" & Format(Date, "yyyymmdd") & " OR pro.pro_fecven <= 0) ORDER BY pro.pro_nombre UNION SELECT DISTINCT pro_codigo, pro_nombre FROM b_productos pro, b_bodegas bod WHERE pro.pro_codigo=bod.bod_codpro AND UCASE(pro_nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND bod.bod_codbod=" & vg_codbod & " AND bod.bod_canmer>0 ORDER BY pro.pro_nombre", vg_db, adOpenStatic
'    ElseIf Opx = "ProInvNoStock" Then  'Tabla generica
'         RS1.Open "SELECT DISTINCT pro.pro_codigo, pro.pro_nombre FROM b_productos pro, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo=c.cli_codtis OR pro.pro_maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=pro.pro_maepro OR pro.pro_maepro<1) AND UCASE(pro.pro_nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND (pro.pro_fecven>" & Format(Date, "yyyymmdd") & " OR pro.pro_fecven <= 0) ORDER BY pro.pro_nombre UNION SELECT DISTINCT pro_codigo, pro_nombre FROM b_productos pro, b_bodegas bod WHERE pro.pro_codigo=bod.bod_codpro AND UCASE(pro_nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND bod.bod_codbod=" & vg_codbod & " AND bod.bod_canmer>0 ORDER BY pro.pro_nombre", vg_db, adOpenStatic
'    ElseIf Opx = "ProGrl" Then
'        RS1.Open "SELECT DISTINCT " & Suf & "codigo,  " & Suf & "nombre FROM " & Tabla & " a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo=c.cli_codtis OR " & Suf & "maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=" & Suf & "maepro OR " & Suf & "maepro<1) AND (" & Suf & "fecven>" & Format(Date, "yyyymmdd") & " OR " & Suf & "fecven <= 0) AND ucase(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND (" & Suf & "ctacon IN ('" & fg_CambiaChar(GetParametro("ctagastos"), ";", "','") & "') OR " & Suf & "ctacon IN ('" & fg_CambiaChar(GetParametro("ctagastos2"), ";", "','") & "'))", vg_db, adOpenStatic
'    ElseIf Opx = "0" Or Opx = "1" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE (" & Suf & "tiprec='" & Opx & "' or '" & Opx & "'='') AND ucase(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf Opx = "Contrato" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=0 AND ucase(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
''    ElseIf Opx = "Traspaso" Then
''        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=2 AND ucase(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.Text)) & "%' ", vg_db, adOpenStatic
'     ElseIf Mid(Opx, 1, 8) = "Traspaso" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "codigo<>'" & Trim(Mid(Opx, 9, 10)) & "' AND (" & Suf & "tipo=2 OR " & Suf & "tipo=0) AND ucase(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "Sucursal" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "direccion FROM " & Tabla & " WHERE " & Suf & "codcli='" & vg_codigo & "' AND UCASE(" & Suf & "direccion) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "Cliente" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=1 AND " & Suf & "activo='1' AND UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "CliSap" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=1 AND " & Suf & "clisap='1' AND UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "CliAlum" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=3 AND ucase(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "Proing" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & ", b_productosing WHERE " & Suf & "codigo=b_productosing.pri_coding AND b_productosing.pri_codpro='" & vg_codigo & "' AND ucase(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ", vg_db, adOpenStatic
'    ElseIf Opx = "Pst" Then 'Tabla productos controla stock
'        RS1.Open "SELECT DISTINCT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo=c.cli_codtis OR " & Suf & "maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=" & Suf & "maepro OR " & Suf & "maepro<1) AND UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND (" & Suf & "fecven>" & Format(Date, "yyyymmdd") & " OR " & Suf & "fecven <= 0)  AND " & Suf & "ctrsto=1 ORDER BY " & Suf & "nombre", vg_db, adOpenStatic
'    ElseIf Opx = "Pbo" Then 'Tabla productos Bodega
'        RS1.Open "SELECT " & Tabla & "." & Suf & "codigo, " & Tabla & "." & Suf & "nombre FROM " & Tabla & ", b_bodegas WHERE UCASE(" & Tabla & "." & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND " & Tabla & "." & Suf & "ctrsto=1 AND " & Tabla & "." & Suf & "codigo=b_bodegas.bod_codpro AND ROUND(b_bodegas.bod_canmer,2)>0 AND b_bodegas.bod_codbod=" & vg_codbod & " ORDER BY " & Tabla & "." & Suf & "nombre", vg_db, adOpenStatic
'    ElseIf Opx = "No5etapas" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "codigo<10000 AND UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf Opx = "5etapas" Then
'        RS1.Open "SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "codigo = " & vg_tiprec & " AND UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    ElseIf vg_modrec = False Then
'        RS1.Open "SELECT DISTINCT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & ", b_productos,  b_productosing, a_tipopro " & _
'                 "WHERE  " & Suf & "codigo=b_productosing.pri_coding " & _
'                 "AND    b_productosing.pri_codpro=b_productos.pro_codigo " & _
'                 "AND    b_productos.pro_codtip=a_tipopro.tip_codigo " & _
'                 "AND    trim(str(a_tipopro.tip_codigo)) IN ('" & Opx & "') " & _
'                 "AND    UCASE(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' " & _
'                 "ORDER BY " & Suf & "codigo", vg_db, adOpenStatic
'    End If
'End If
'i = 1
'vaSpread1.Visible = False
'vaSpread1.MaxRows = 0
'If Not RS1.EOF Then
'    Do While Not RS1.EOF
'        var = -1
'        If Opx = "ProInv" Then
'           var = M_TomInv.vaSpread1.SearchCol(1, 0, M_TomInv.vaSpread1.MaxRows, Trim(RS1(0)), SearchFlagsNone)
'        ElseIf Opx = "ProInv1" Then
'           var = M_AjuInv.vaSpread1.SearchCol(1, 0, M_AjuInv.vaSpread1.MaxRows, Trim(RS1(0)), SearchFlagsNone)
'        End If
'        If (Mid(Opx, 1, 6) = "ProInv" And var = -1) Or (Mid(Opx, 1, 6) <> "ProInv") Then
'            vaSpread1.MaxRows = RS1.RecordCount
'            vaSpread1.MaxRows = i: vaSpread1.Row = i: i = i + 1
'            vaSpread1.Col = 1
'            If Opx = "Cliente" Or Opx = "CliAlum" Then
'               vaSpread1.text = fg_PintaRut(RS1(0))
'            Else
'               vaSpread1.text = RS1(0)
'            End If
''            vaSpread1.Col = 1: vaSpread1.text = IIf(Opx = "Cliente" Or Opx = "CliAlum", fg_PintaRut(RS1(0)), RS1(0))
'            vaSpread1.Col = 2: vaSpread1.text = Trim(RS1(1))
'            If Opx = "Pac" Then vaSpread1.text = vaSpread1.text & " " & Trim(RS1(2)) & " " & Trim(RS1(3))
'        End If
'        RS1.MoveNext
'    Loop
'End If
'RS1.Close: Set RS1 = Nothing
'vaSpread1.Visible = True
'vaSpread1.SetActiveCell 1, 1
'Exit Sub
'Man_Error:
'    Resume Next
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If est Then est = False: Exit Sub
If KeyCode = 27 Then Cerrar: Exit Sub
If KeyCode = 40 Or KeyCode = 34 And iRow > 0 Then vaSpread1.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    MoverDatos
Case 3
    vg_codigo = ""
    Cerrar
End Select
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
MoverDatos
End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
On Error Resume Next
SendKeys "{Tab}"
MoverDatos
SendKeys "+{Tab}"

Exit Sub
Man_Error:
    Resume Next

End Sub

Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)
If est Then est = False: Exit Sub
If KeyCode = 27 Then Cerrar: Exit Sub
If TeclasNoPermitidas(KeyCode) = True Then Text1.text = IIf(KeyCode = 8, Text1.text, Text1.text & Chr(KeyCode)): Text1.SetFocus: Text1.SelStart = Len(Text1.text)
End Sub

Private Sub MoverDatos()
With vaSpread1
    If .MaxRows < 1 Then Exit Sub
    .Row = .ActiveRow
    .Col = 1: vg_codigo = Trim(.text)
    .Col = 2: vg_nombre = Trim(.text)
End With
Cerrar
End Sub

Sub Cerrar()
Me.Hide
Unload Me
End Sub

Sub LlenaDatos(TablaGen As String, SufGen As String, TitGen As String, Op As String)
Dim est As Boolean, z As Long, var As Long, sql1 As String

On Error GoTo Man_Error

Me.Caption = TitGen
fg_carga ""
Opx = Op
Tabla = TablaGen
Suf = SufGen
Titulo = TitGen
vaSpread1.MaxRows = 0

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Opx = "Gen" Then '-------> Tabla generica
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " ORDER BY " & SufGen & "nombre", vg_db, adOpenForwardOnly ', adOpenStatic

ElseIf Opx = "Proveedor" Then '-------> Tabla Proveedor
   
   Set RS1 = vg_db.Execute("sgp_Sel_ListaProveedor")

ElseIf Opx = "PtoAte" Then
    
    RS1.Open "select distinct a.ate_codatencion, a.ate_descripcion from a_pto_atencion a inner join b_detallelectura b on a.ate_codatencion = b.ate_codatencion where b.cli_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic

ElseIf Opx = "RegVal" Then
    
    RS1.Open "select distinct a.reg_codigo, a.reg_nombre from a_regimen a inner join b_detallelectura b on a.reg_codigo = b.reg_codigo where b.cli_codigo = '" & vg_codigo & "' and b.ate_codatencion = " & Val(vg_ptoate) & "", vg_db, adOpenStatic

ElseIf Opx = "SerVal" Then
    
    RS1.Open "select distinct a.ser_codigo, a.ser_nombre from a_servicio a inner join b_detallelectura b on a.ser_codigo = b.ser_codigo where b.cli_codigo = '" & vg_codigo & "' and b.ate_codatencion = " & Val(vg_ptoate) & "", vg_db, adOpenStatic

ElseIf Opx = "Ser" Then '-------> Tabla generica
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "activo = '1' ORDER BY " & SufGen & "nombre", vg_db, adOpenForwardOnly ', adOpenStatic

ElseIf Opx = "SerBlo" Then  '-------> Tabla minuta bloque
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "codigo > 9999 and " & SufGen & "activo = '1' ORDER BY " & SufGen & "nombre", vg_db, adOpenForwardOnly ', adOpenStatic

ElseIf Opx = "RegBlo" Then  '-------> Tabla minuta bloque
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "codigo > 9999 ORDER BY " & SufGen & "nombre", vg_db, adOpenForwardOnly ', adOpenStatic

ElseIf Opx = "Gpr" Then '-------> Tabla producto
    
    SufGen = "a." & SufGen
    Suf = SufGen
    RS1.Open "SELECT DISTINCT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR " & SufGen & "maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = " & SufGen & "maepro OR " & SufGen & "maepro < 1) ORDER BY " & SufGen & "nombre", vg_db, adOpenStatic

ElseIf Opx = "Tpc" Then '-------> Tabla lista precio
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "cencos = '" & MuestraCasino(1) & "' AND " & SufGen & "activo = '1' ORDER BY " & SufGen & "nombre", vg_db, adOpenStatic

ElseIf Opx = "Bod" Then '-------> Filtrar codigo bodega
    
    RS1.Open "SELECT a.* FROM a_bodega a, b_clientes b WHERE (a.bod_codigo NOT IN (SELECT DISTINCT cli_codbod FROM b_clientes WHERE (cli_codbod) Is Not Null  AND cli_codbod > 0)  OR a.bod_codigo = b.cli_codbod) AND b.cli_codigo = '" & Trim(SufGen) & "'", vg_db, adOpenStatic

ElseIf Opx = "Pac" Then '-------> Tabla paciente
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre, " & SufGen & "appaterno, " & SufGen & "apmaterno FROM " & TablaGen & " WHERE " & SufGen & "estado = '0' ORDER BY " & SufGen & "nombre", vg_db, adOpenStatic

ElseIf Opx = "Reg" Then '-------> Tabla generica
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "codigo = " & vg_codregimen & " ORDER BY " & SufGen & "nombre", vg_db, adOpenStatic

ElseIf Opx = "Usu" Then '-------> Tabla usuario grupo paciente
    
    SufGen = "a." & SufGen: Suf = SufGen
    RS1.Open "SELECT DISTINCT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " a, b_usuariogrupopac b, b_pacientes c WHERE " & SufGen & "codigo = b.ugp_codusu AND b.ugp_codgrp = c.pac_codgrp AND c.pac_codigo = '" & vg_Aux & "' ORDER BY " & SufGen & "nombre", vg_db, adOpenStatic

ElseIf Opx = "ProVig" Then
    
    SufGen = "a." & SufGen
    Suf = SufGen
'    RS1.Open "SELECT DISTINCT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR " & SufGen & "maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = " & SufGen & "maepro OR " & SufGen & "maepro < 1) AND (" & SufGen & "fecven>" & Format(Date, "yyyymmdd") & " OR " & SufGen & "fecven <= 0) ORDER BY " & SufGen & "nombre", vg_db, adOpenForwardOnly ', adOpenStatic
    Set RS1 = vg_db.Execute("sgp_Sel_ListaProductosVigentes '" & MuestraCasino(1) & "'")

ElseIf Opx = "ProVigSac" Then
    
    SufGen = "a." & SufGen
    Suf = SufGen
    RS1.Open "SELECT DISTINCT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " a, a_tiposervicio b, b_clientes c, b_formatocompras d, b_formatocomprassgp e WHERE (b.tis_codigo = c.cli_codtis OR " & SufGen & "maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = " & SufGen & "maepro OR " & SufGen & "maepro < 1) AND (" & SufGen & "fecven>" & Format(Date, "yyyymmdd") & " OR " & SufGen & "fecven <= 0) AND d.foc_codsac = e.fcs_codsac AND " & SufGen & "codigo = e.fcs_codsgp AND d.foc_codsac = '" & vg_codigo & "' ORDER BY " & SufGen & "nombre", vg_db, adOpenStatic

ElseIf Opx = "ProInv" Or Opx = "ProInv1" Then '-------> Tabla generica
    
    If vg_tipbase = "1" Then
       
       RS1.Open "SELECT DISTINCT pro.pro_codigo, pro.pro_nombre FROM b_productos pro, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR pro.pro_maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = pro.pro_maepro OR pro.pro_maepro < 1) AND pro.pro_ctrsto = 1 AND (pro.pro_fecven > " & Format(Date, "yyyymmdd") & " OR pro.pro_fecven <= 0) ORDER BY pro.pro_nombre union SELECT DISTINCT pro_codigo, pro_nombre FROM b_productos pro, b_bodegas bod WHERE pro.pro_codigo = bod.bod_codpro AND bod.bod_codbod = " & vg_codbod & " AND bod.bod_canmer > 0 ORDER BY pro.pro_nombre", vg_db, adOpenStatic
    
    Else
       
       RS1.Open "SELECT DISTINCT pro.pro_codigo, pro.pro_nombre FROM b_productos pro, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR pro.pro_maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = pro.pro_maepro OR pro.pro_maepro < 1) AND pro.pro_ctrsto = 1 AND (pro.pro_fecven > " & Format(Date, "yyyymmdd") & " OR pro.pro_fecven <= 0) UNION SELECT DISTINCT pro_codigo, pro_nombre FROM b_productos pro, b_bodegas bod WHERE pro.pro_codigo = bod.bod_codpro AND bod.bod_codbod = " & vg_codbod & " AND bod.bod_canmer > 0 ORDER BY pro.pro_nombre", vg_db, adOpenStatic
    
    End If

ElseIf Opx = "ProInvNoStock" Then '-------> Tabla generica
    
    If vg_tipbase = "1" Then
       
       RS1.Open "SELECT DISTINCT pro.pro_codigo, pro.pro_nombre FROM b_productos pro, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR pro.pro_maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=pro.pro_maepro OR pro.pro_maepro<1) AND (pro.pro_fecven>" & Format(Date, "yyyymmdd") & " OR pro.pro_fecven <= 0) ORDER BY pro.pro_nombre union SELECT DISTINCT pro_codigo, pro_nombre FROM b_productos pro, b_bodegas bod WHERE pro.pro_codigo=bod.bod_codpro AND bod.bod_codbod=" & vg_codbod & " AND bod.bod_canmer>0 ORDER BY pro.pro_nombre", vg_db, adOpenStatic
    
    Else
       
       RS1.Open "SELECT DISTINCT pro.pro_codigo, pro.pro_nombre FROM b_productos pro, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR pro.pro_maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = pro.pro_maepro OR pro.pro_maepro < 1) AND (pro.pro_fecven > " & Format(Date, "yyyymmdd") & " OR pro.pro_fecven <= 0) union SELECT DISTINCT pro_codigo, pro_nombre FROM b_productos pro, b_bodegas bod WHERE pro.pro_codigo = bod.bod_codpro AND bod.bod_codbod = " & vg_codbod & " AND bod.bod_canmer > 0 ORDER BY pro.pro_nombre", vg_db, adOpenStatic
    
    End If

ElseIf Opx = "ProGrl" Then
    
    SufGen = "a." & SufGen
    Suf = SufGen
'20100209    RS1.Open "SELECT DISTINCT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR " & SufGen & "maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = " & SufGen & "maepro OR " & SufGen & "maepro < 1) AND (" & SufGen & "fecven > " & Format(Date, "yyyymmdd") & " OR " & SufGen & "fecven <= 0) AND (" & SufGen & "ctacon IN ('" & fg_CambiaChar(GetParametro("ctagastos"), ";", "','") & "') OR " & SufGen & "ctacon IN ('" & fg_CambiaChar(GetParametro("ctagastos2"), ";", "','") & "')) ORDER BY " & SufGen & "nombre", vg_db, adOpenStatic
    RS1.Open "SELECT DISTINCT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR " & SufGen & "maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = " & SufGen & "maepro OR " & SufGen & "maepro < 1) AND (" & SufGen & "fecven > " & Format(Date, "yyyymmdd") & " OR " & SufGen & "fecven <= 0) AND (" & SufGen & "ctacon NOT IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') AND " & SufGen & "ctacon NOT IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')) ORDER BY " & SufGen & "nombre", vg_db, adOpenStatic

ElseIf Opx = "0" Or Opx = "1" Then '-------> Tabla Recetas estado 0
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE (" & SufGen & "tiprec = '" & Op & "' OR '" & Op & "' = '') ORDER BY " & SufGen & "codigo", vg_db, adOpenStatic

ElseIf Opx = "Contrato" Then
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "tipo = 0 ORDER BY " & SufGen & "codigo", vg_db, adOpenStatic

ElseIf Mid(Opx, 1, 8) = "Traspaso" Then
'    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "tipo=2 ORDER BY " & SufGen & "codigo", vg_db, adOpenStatic
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "codigo <> '" & Mid(Opx, 9, 5) & "' AND (" & SufGen & "tipo = 2 OR " & SufGen & "tipo = 0) ORDER BY " & SufGen & "codigo", vg_db, adOpenStatic

ElseIf Opx = "Sucursal" Then
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "direccion FROM " & TablaGen & " WHERE " & SufGen & "codcli = '" & vg_codigo & "' ORDER BY " & SufGen & "codigo", vg_db, adOpenStatic

ElseIf Opx = "Cliente" Then
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "tipo = 1 AND " & SufGen & "activo = '1' ORDER BY " & SufGen & "codigo", vg_db, adOpenStatic

ElseIf Opx = "CliSap" Then
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "tipo = 1 AND " & SufGen & "clisap = '1'ORDER BY " & SufGen & "codigo", vg_db, adOpenStatic

ElseIf Opx = "CliAlum" Then
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "tipo = 3 ORDER BY " & SufGen & "codigo", vg_db, adOpenStatic

ElseIf Opx = "Proing" Then
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & ", b_productosing WHERE " & SufGen & "codigo = b_productosing.pri_coding AND b_productosing.pri_codpro = '" & vg_codigo & "' ORDER BY " & SufGen & "codigo", vg_db, adOpenStatic

ElseIf Opx = "Pst" Then 'Tabla productos controla stock
    
    SufGen = "a." & SufGen
    Suf = SufGen
    RS1.Open "SELECT DISTINCT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR " & SufGen & "maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = " & SufGen & "maepro OR " & SufGen & "maepro < 1) AND  (" & SufGen & "fecven > " & Format(Date, "yyyymmdd") & " OR " & SufGen & "fecven <= 0) AND " & SufGen & "ctrsto = 1 ORDER BY " & SufGen & "nombre", vg_db, adOpenStatic

ElseIf Opx = "Pbo" Then 'Tabla productos Bodega
    
    RS1.Open "SELECT " & TablaGen & "." & SufGen & "codigo, " & TablaGen & "." & SufGen & "nombre FROM " & TablaGen & ", b_bodegas WHERE " & TablaGen & "." & SufGen & "ctrsto=1 AND " & TablaGen & "." & SufGen & "codigo=b_bodegas.bod_codpro AND ROUND(b_bodegas.bod_canmer,2)>0 AND b_bodegas.bod_codbod=" & vg_codbod & " ORDER BY " & TablaGen & "." & SufGen & "nombre", vg_db, adOpenStatic

ElseIf Opx = "No5etapas" Then
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "codigo<10000 ORDER BY " & SufGen & "nombre", vg_db, adOpenStatic

ElseIf Opx = "5etapas" Then
    
    RS1.Open "SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "codigo= " & vg_tiprec & " ORDER BY " & SufGen & "nombre", vg_db, adOpenStatic

ElseIf Opx = "PSAC" Then
    
    vaSpread1.MaxCols = 4
    vaSpread1.Row = 0
    vaSpread1.Col = 3: vaSpread1.text = "Unidad"
    vaSpread1.Col = 4: vaSpread1.text = "F.Conver."
    Me.Width = 7820
    vaSpread1.Width = 7120
    Frame1.Left = 950
    sql1 = IIf(vg_tipbase = "1", " AND cdate(foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), foc_vigfin,101) >  '" & Date & "'")
    RS1.Open "SELECT foc_codsac, foc_nomsac, foc_unisac, foc_faccon FROM b_formatocompras WHERE (foc_flexec = 0 OR (foc_flexec = -1 " & sql1 & "))", vg_db, adOpenStatic

ElseIf Opx = "CamPSAC" Then
    
    sql1 = IIf(vg_tipbase = "1", " AND cdate(a.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), a.foc_vigfin,101) >  '" & Date & "'")
    RS1.Open "SELECT DISTINCT a.foc_codsac, a.foc_nomsac FROM b_formatocompras a, b_formatocomprassgp b WHERE a.foc_codsac = b.fcs_codsac and b.fcs_codsgp IN ('" & SufGen & "') and b.fcs_codsac <> '" & TablaGen & "' AND (a.foc_flexec = 0 OR (a.foc_flexec = -1 " & sql1 & "))", vg_db, adOpenStatic

ElseIf vg_modrec = False Then
    
    sql1 = IIf(vg_tipbase = "1", " trim(str(a_tipopro.tip_codigo)) ", " ltrim(convert(varchar(20),a_tipopro.tip_codigo)) ")
    RS1.Open "SELECT DISTINCT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & ", b_productos,  b_productosing, a_tipopro " & _
             "WHERE  " & SufGen & "codigo = b_productosing.pri_coding " & _
             "AND    b_productosing.pri_codpro = b_productos.pro_codigo " & _
             "AND    b_productos.pro_codtip = a_tipopro.tip_codigo " & _
             "AND    " & sql1 & " IN ('" & (Opx) & "') " & _
             "ORDER BY " & SufGen & "codigo", vg_db, adOpenStatic

ElseIf Opx = "VtaSerEsp" Then

    Set RS1 = vg_db.Execute("sgp_Sel_ListaDescripcionVentataServiciosEspeciales '" & MuestraCasino(1) & "', " & vg_codbod & "")


End If

With vaSpread1

If Not RS1.EOF Then
    
    Do While Not RS1.EOF
       
       var = -1
        If Opx = "ProInv" Then
           
           var = M_TomInv.vaSpread1.SearchCol(1, 0, M_TomInv.vaSpread1.MaxRows, Trim(RS1(0)), SearchFlagsNone)
        
        ElseIf Opx = "ProInv1" Then
           
           var = M_AjuInv.vaSpread1.SearchCol(1, 0, M_AjuInv.vaSpread1.MaxRows, Trim(RS1(0)), SearchFlagsNone)
        
        
        End If
        If (Mid(Opx, 1, 6) = "ProInv" And var = -1) Or (Mid(Opx, 1, 6) <> "ProInv") Then
'        If ((Opx = "ProInv" Or Opx = "ProInv1") And var = -1) Or (Opx <> "ProInv" And Opx <> "ProInv1") Then
            
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
            
            If Opx = "Cliente" Or Opx = "CliAlum" Then
               
               .text = fg_PintaRut(RS1(0))
            
            Else
               
               .TypeHAlign = IIf(Opx = "PSAC", 0, 1)
               .text = Trim(RS1(0))
               If .MaxCols > 2 Then
                  
                  .Col = 3
                  .text = IIf(IsNull(RS1(2)), "", Trim(RS1(2)))
                  .Col = 4
                  .text = IIf(IsNull(RS1(3)), 0, RS1(3))
               
               End If
            
            End If
'            .text = IIf(Opx = "Cliente" Or Opx = "CliAlum", fg_PintaRut(RS1(0)), RS1(0))
            
            .Col = 2
            .text = Trim(RS1(1))
            If Opx = "Pac" Then .text = .text & " " & Trim(RS1(2)) & " " & Trim(RS1(3))
        
        End If
        
        RS1.MoveNext
    
    Loop

'    If RS1(0).Type = adVarWChar Then
'        .Col = 1:
'        .Row = -1
'        .TypeHAlign = TypeHAlignLeft
'    Else
'        .Col = 1:
'        .Row = -1
'        .TypeHAlign = TypeHAlignRight
'    End If
End If
RS1.Close
Set RS1 = Nothing

End With
fg_descarga

Exit Sub
Man_Error:
    Resume Next
End Sub
