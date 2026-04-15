VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_ProPre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso Precio Producto"
   ClientHeight    =   7695
   ClientLeft      =   3090
   ClientTop       =   1830
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6555
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   6975
         _Version        =   393216
         _ExtentX        =   12303
         _ExtentY        =   11562
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   4
         MaxRows         =   30
         SpreadDesigner  =   "M_ProPre.frx":0000
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_ProPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim est As Boolean
Dim Msgtitulo As String, vcencos As String, vcodreg As Long, vcodser As Long, vtipmin As String, vanomes As Long
Private FeHasta As Long

Private Sub Form_Activate()
fg_descarga
TraerFechaCierre
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 8205
Me.Width = 7590
Msgtitulo = "Ingreso Precio Producto"
fg_centra Me
modo = "": est = False
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = False
Toolbar1.Buttons(5).Visible = False
Toolbar1.Buttons(6).Visible = False
Toolbar1.Buttons(15).Visible = False
Toolbar1.Buttons(16).Visible = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 3
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
    Toolbar1.Buttons(15).Visible = False
    Toolbar1.Buttons(16).Visible = False
Case 7
    Call LlenarListaPrecio(vcencos, vcodreg, vcodser, vanomes, vtipmin, FeHasta)
Case 10
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    Call LlenarListaPrecio(vcencos, vcodreg, vcodser, vanomes, vtipmin, FeHasta)
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
    Toolbar1.Buttons(15).Visible = False
    Toolbar1.Buttons(16).Visible = False
Case 12
    Dim vCodPro As Long, vPrePro As Double, fecini As Date, fecfin As Date, coding As String
    vg_db.BeginTrans
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 4
        If Trim(vaSpread1.text) <> "" And Val(vaSpread1.text) > 0 Then
           vaSpread1.Col = 1: vCodPro = vaSpread1.text
           vaSpread1.Col = 4: vPrePro = vaSpread1.text
           fecini = dBoM(vg_ciedia)
           fecfin = CDate(vg_ciedia) - 1
           vg_db.Execute "UPDATE b_productospmpdia SET ppd_propon = " & vPrePro & _
                          " WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                          " AND ppd_codpro = '" & vCodPro & "' " & _
                          " AND ppd_propon < 1 " & _
                          " AND ppd_fecdia >= " & Format(fecini, "yyyymmdd") & _
                          " AND ppd_fecdia <= " & Format(fecfin, "yyyymmdd")
           RS.Open "SELECT DISTINCT pri_coding FROM b_productosing WHERE pri_codpro = '" & vCodPro & "'", vg_db, adOpenStatic
           If Not RS.EOF Then
              RS1.Open "SELECT Round(AVG(a.ppd_propon/c.pro_facing), 2) AS cosing " & _
                      "FROM  b_productospmpdia a, b_productosing b, b_productos c " & _
                      "WHERE b.pri_codpro = c.pro_codigo " & _
                      "AND   c.pro_codigo = a.ppd_codpro " & _
                      "AND   a.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                      "AND   a.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
                      "AND   a.ppd_propon > 0 AND b.pri_coding = '" & RS!pri_coding & "'", vg_db, adOpenStatic
              If Not RS.EOF Then
                 vg_db.Execute "UPDATE b_contlistpreing " & _
                               " SET cpi_feccos = " & Format(Date, "yyyymmdd") & ", cpi_precos = " & IIf(IsNull(RS1!cosing), 0, RS1!cosing) & " " & _
                               " WHERE cpi_cencos = '" & MuestraCasino(1) & "' " & _
                               " AND cpi_coding = '" & RS!pri_coding & "'"
              End If
              RS1.Close: Set RS1 = Nothing
           End If
           RS.Close: Set RS = Nothing
         End If
    Next i
    vg_db.CommitTrans
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
    Toolbar1.Buttons(15).Visible = False
    Toolbar1.Buttons(16).Visible = False
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Or 2147217900 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
If modo = "" Then modo = "M"
If modo = "M" And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = False: Toolbar1.Buttons(15).Visible = False: Toolbar1.Buttons(16).Visible = False
End Sub

Sub LlenarListaPrecio(cencos As String, codreg As Long, codser As Long, anomes As Long, tipmin As String, FecHasta As Long)
On Error GoTo Man_Error
fg_carga ""
vaSpread1.Visible = False
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &HC0FFFF
vaSpread1.MaxRows = 0
vcencos = cencos
vcodreg = codreg
vcodser = codser
vanomes = anomes
Let FeHasta = FecHasta
vtipmin = tipmin
Dim aAp As String
If vg_tipbase = "1" Then
   '-------> Insert tabla productospmpdia
   aAp = Trim(vg_NUsr) & "_tmp_ProductoPMPLisPrecio"
   fg_CheckTmp aAp
   vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, 0 AS ppd_upreco, null AS ppd_fecuco, Max(ppd_fecdia) AS ppd_fecdia " & _
                 "INTO " & aAp & " " & _
                 "FROM b_productospmpdia " & _
                 "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                 "AND   ppd_propon > 0 " & _
                 "GROUP BY ppd_cencos, ppd_codpro"
                 
   vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
   
   vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon, " & aAp & ".ppd_upreco=b_productospmpdia.ppd_upreco, " & aAp & ".ppd_fecuco=b_productospmpdia.ppd_fecuco"
   
   vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_upreco, ppd_fecuco, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_upreco, ppd_fecuco, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
   
   RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, c.uni_nombre " & _
           "FROM b_productos a, a_unidad c, b_minuta d, b_minutadet e, b_recetadet f, b_ingrediente g, " & aAp & " h, b_contlistpreing i " & _
           "WHERE d.min_codigo = e.mid_codigo " & _
           "AND   e.mid_codrec = f.red_codigo " & _
           "AND   e.mid_tiprec = f.red_tiprec AND ((f.red_tiprec <> 0 AND f.red_cencos = '" & MuestraCasino(1) & "') OR (f.red_tiprec = 0 AND f.red_cencos = '0')) " & _
           "AND   f.red_codpro = g.ing_codigo " & _
           "AND   g.ing_codigo = i.cpi_coding " & _
           "AND   i.cpi_codcom = a.pro_codigo " & _
           "AND   a.pro_codigo = h.ppd_codpro " & _
           "AND   i.cpi_cencos = '" & cencos & "' " & _
           "AND   h.ppd_cencos = '" & cencos & "' " & _
           "AND  (a.pro_codigo IN (SELECT bod_codpro FROM b_bodegas WHERE bod_codbod=" & vg_codbod & " AND bod_canmer<1) OR a.pro_codigo NOT IN (SELECT bod_codpro FROM b_bodegas WHERE bod_codbod=" & vg_codbod & ")) " & _
           "AND   d.min_cencos = '" & cencos & "' " & _
           "AND   d.min_codreg = " & codreg & " " & _
           "AND   d.min_codser = " & codser & " " & _
           "AND   val(mid(d.min_fecmin,1,6)) >= " & anomes & " " & _
           "AND   val(mid(d.min_fecmin,1,6)) <= " & FecHasta & " " & _
           "AND   e.mid_tipmin = '" & tipmin & "'  " & _
           "AND   a.pro_coduni = c.uni_codigo " & _
           "AND   i.cpi_precos <= 0 AND h.ppd_propon = 0 " & _
           "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) ORDER BY a.pro_nombre", vg_db, adOpenStatic
 Else
    RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, c.uni_nombre " & _
            "FROM b_productos a, a_unidad c, b_minuta d, b_minutadet e, b_recetadet f, b_ingrediente g, b_productospmpdia h, b_contlistpreing i " & _
            "WHERE d.min_codigo = e.mid_codigo " & _
            "AND   e.mid_codrec = f.red_codigo " & _
            "AND   e.mid_tiprec = f.red_tiprec AND ((f.red_tiprec <> 0 AND f.red_cencos = '" & MuestraCasino(1) & "') OR (f.red_tiprec = 0 AND f.red_cencos = '0')) " & _
            "AND   f.red_codpro = g.ing_codigo " & _
            "AND   g.ing_codigo = i.cpi_coding " & _
            "AND   i.cpi_codcom = a.pro_codigo " & _
            "AND   a.pro_codigo = h.ppd_codpro " & _
            "AND   i.cpi_cencos = '" & cencos & "' " & _
            "AND   h.ppd_cencos = '" & cencos & "' " & _
            "AND  (a.pro_codigo IN (SELECT bod_codpro FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " AND bod_canmer < 1) OR a.pro_codigo NOT IN (SELECT bod_codpro FROM b_bodegas WHERE bod_codbod = " & vg_codbod & ")) " & _
            "AND   d.min_cencos = '" & cencos & "' " & _
            "AND   d.min_codreg = " & codreg & " " & _
            "AND   d.min_codser = " & codser & " " & _
            "AND   convert(int,substring(convert(varchar(8),d.min_fecmin),1,6)) >= " & anomes & " " & _
            "AND   convert(int,substring(convert(varchar(8),d.min_fecmin),1,6)) <= " & FecHasta & " " & _
            "AND   e.mid_tipmin = '" & tipmin & "'  " & _
            "AND   a.pro_coduni = c.uni_codigo " & _
            "AND   i.cpi_precos <= 0 AND h.ppd_propon = 0 AND h.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
            "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) ORDER BY a.pro_nombre", vg_db, adOpenStatic
End If
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!pro_codigo
      vaSpread1.Col = 2: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!pro_nombre
      vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!uni_nombre
      vaSpread1.Col = 4: vaSpread1.text = ""
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
   vaSpread1.SetActiveCell 4, 1
   vaSpread1.Visible = True
   fg_descarga
Else
   vg_codigo = ""
   RS.Close: Set RS = Nothing
   fg_descarga
   Gl_Ac_Botones Me, 1, 3, modo
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = False
   Toolbar1.Buttons(5).Visible = False
   Toolbar1.Buttons(6).Visible = False
   Toolbar1.Buttons(15).Visible = False
   Toolbar1.Buttons(16).Visible = False
   MsgBox "No existe información, con valores ceros", vbCritical + vbOKOnly, Msgtitulo
End If
'-------> Borrar tablas temporales
If Trim(aAp) <> "" And vg_tipbase = "1" Then vg_db.Execute "DROP TABLE " & aAp & ""
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub




















'Dim modo As String
'Dim RS As New ADODB.Recordset
'Dim RS1 As New ADODB.Recordset
'Dim est As Boolean
'Dim Msgtitulo As String, vcencos As String, vcodreg As Long, vcodser As Long, vtipmin As String, vanomes As Long
'
'Private Sub Form_Activate()
'fg_descarga
'TraerFechaCierre
'End Sub
'
'Private Sub Form_Load()
'Me.HelpContextID = vg_OpcM
'Me.Height = 8205
'Me.Width = 7590
'Msgtitulo = "Ingreso Precio Producto"
'fg_centra Me
'modo = "": est = False
'Gl_Mo_Botones Me, 1
'Gl_Ac_Botones Me, 1, 1, modo
'Toolbar1.Buttons(1).Visible = False
'Toolbar1.Buttons(2).Visible = False
'Toolbar1.Buttons(5).Visible = False
'Toolbar1.Buttons(6).Visible = False
'Toolbar1.Buttons(15).Visible = False
'Toolbar1.Buttons(16).Visible = False
'End Sub
'
'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Man_Error
'Select Case Button.Index
'Case 3
'    modo = "M"
'    Gl_Ac_Botones Me, 1, 0, modo
'    Toolbar1.Buttons(1).Visible = False
'    Toolbar1.Buttons(2).Visible = False
'    Toolbar1.Buttons(5).Visible = False
'    Toolbar1.Buttons(6).Visible = False
'    Toolbar1.Buttons(15).Visible = False
'    Toolbar1.Buttons(16).Visible = False
'Case 7
'    LlenarListaPrecio vcencos, vcodreg, vcodser, vanomes, vtipmin
'Case 10
'    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
'    LlenarListaPrecio vcencos, vcodreg, vcodser, vanomes, vtipmin
'    modo = "": Gl_Ac_Botones Me, 1, 1, modo
'    Toolbar1.Buttons(1).Visible = False
'    Toolbar1.Buttons(2).Visible = False
'    Toolbar1.Buttons(5).Visible = False
'    Toolbar1.Buttons(6).Visible = False
'    Toolbar1.Buttons(15).Visible = False
'    Toolbar1.Buttons(16).Visible = False
'Case 12
'    Dim vCodPro As Long, vPrePro As Double, fecini As Date, fecfin As Date, coding As String
'    vg_db.BeginTrans
'    For I = 1 To vaSpread1.MaxRows
'        vaSpread1.Row = I
'        vaSpread1.Col = 4
'        If Trim(vaSpread1.text) <> "" And Val(vaSpread1.text) > 0 Then
'           vaSpread1.Col = 1: vCodPro = vaSpread1.text
'           vaSpread1.Col = 4: vPrePro = vaSpread1.text
'           fecini = dBoM(vg_ciedia)
'           fecfin = CDate(vg_ciedia) - 1
'           vg_db.Execute "UPDATE b_productospmpdia SET ppd_propon = " & vPrePro & " WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_codpro = '" & vCodPro & "' AND ppd_propon < 1 AND ppd_fecdia >= " & Format(fecini, "yyyymmdd") & " AND ppd_fecdia <= " & Format(fecfin, "yyyymmdd") & ""
'           RS.Open "SELECT DISTINCT pri_coding FROM b_productosing WHERE pri_codpro = '" & vCodPro & "'", vg_db, adOpenStatic
'           If Not RS.EOF Then
'              RS1.Open "SELECT Round(AVG(a.ppd_propon/c.pro_facing), 2) AS cosing " & _
'                      "FROM  b_productospmpdia a, b_productosing b, b_productos c " & _
'                      "WHERE b.pri_codpro = c.pro_codigo " & _
'                      "AND   c.pro_codigo = a.ppd_codpro " & _
'                      "AND   a.ppd_cencos = '" & MuestraCasino(1) & "' " & _
'                      "AND   a.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
'                      "AND   a.ppd_propon > 0 AND b.pri_coding = '" & RS!pri_coding & "'", vg_db, adOpenStatic
'              If Not RS.EOF Then
'                 vg_db.Execute "UPDATE b_contlistpreing SET cpi_feccos = " & Format(Date, "yyyymmdd") & ", cpi_precos = " & IIf(IsNull(RS1!cosing), 0, RS1!cosing) & " " & _
'                               "WHERE cpi_cencos = '" & MuestraCasino(1) & "' AND cpi_coding = '" & RS!pri_coding & "'"
'              End If
'              RS1.Close: Set RS1 = Nothing
'           End If
'           RS.Close: Set RS = Nothing
'         End If
'    Next I
'    vg_db.CommitTrans
'    modo = "": Gl_Ac_Botones Me, 1, 1, modo
'    Toolbar1.Buttons(1).Visible = False
'    Toolbar1.Buttons(2).Visible = False
'    Toolbar1.Buttons(5).Visible = False
'    Toolbar1.Buttons(6).Visible = False
'    Toolbar1.Buttons(15).Visible = False
'    Toolbar1.Buttons(16).Visible = False
'Case 18
'    Me.Hide
'    Unload Me
'End Select
'Exit Sub
'Man_Error:
'If Err = -2147467259 Or 2147217900 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
'If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
'vg_db.RollbackTrans
'fg_descarga
'MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
'ins_log_error Date & Time & Err & ":  " & Error$(Err)
'End Sub
'
'Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
'If vaSpread1.MaxRows < 1 Then Exit Sub
'If modo = "" Then modo = "M"
'If modo = "M" And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = False: Toolbar1.Buttons(15).Visible = False: Toolbar1.Buttons(16).Visible = False
'End Sub
'
'Sub LlenarListaPrecio(cencos As String, codreg As Long, codser As Long, anomes As Long, tipmin As String)
'On Error GoTo Man_Error
'fg_carga ""
'vaSpread1.Visible = False
'vaSpread1.Row = -1: vaSpread1.Col = -1
'vaSpread1.BackColor = &HC0FFFF
'vaSpread1.MaxRows = 0
'vcencos = cencos
'vcodreg = codreg
'vcodser = codser
'vanomes = anomes
'vtipmin = tipmin
'Dim aAp As String
'If vg_tipbase = "1" Then
'   '-------> Insert tabla productospmpdia
'   aAp = Trim(vg_NUsr) & "_tmp_ProductoPMPLisPrecio"
'   fg_CheckTmp aAp
'   vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, 0 AS ppd_upreco, null AS ppd_fecuco, Max(ppd_fecdia) AS ppd_fecdia " & _
'                 "INTO " & aAp & " " & _
'                 "FROM b_productospmpdia " & _
'                 "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
'                 "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
'                 "AND   ppd_propon > 0 " & _
'                 "GROUP BY ppd_cencos, ppd_codpro"
'   vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
'   vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon, " & aAp & ".ppd_upreco=b_productospmpdia.ppd_upreco, " & aAp & ".ppd_fecuco=b_productospmpdia.ppd_fecuco"
'   vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_upreco, ppd_fecuco, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_upreco, ppd_fecuco, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
'   RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, c.uni_nombre " & _
'           "FROM b_productos a, a_unidad c, b_minuta d, b_minutadet e, b_recetadet f, b_ingrediente g, " & aAp & " h, b_contlistpreing i " & _
'           "WHERE d.min_codigo = e.mid_codigo " & _
'           "AND   e.mid_codrec = f.red_codigo " & _
'           "AND   e.mid_tiprec = f.red_tiprec AND ((f.red_tiprec <> 0 AND f.red_cencos = '" & MuestraCasino(1) & "') OR (f.red_tiprec = 0 AND f.red_cencos = '0')) " & _
'           "AND   f.red_codpro = g.ing_codigo " & _
'           "AND   g.ing_codigo = i.cpi_coding " & _
'           "AND   i.cpi_codcom = a.pro_codigo " & _
'           "AND   a.pro_codigo = h.ppd_codpro " & _
'           "AND   i.cpi_cencos = '" & cencos & "' " & _
'           "AND   h.ppd_cencos = '" & cencos & "' " & _
'           "AND  (a.pro_codigo IN (SELECT bod_codpro FROM b_bodegas WHERE bod_codbod=" & vg_codbod & " AND bod_canmer<1) OR a.pro_codigo NOT IN (SELECT bod_codpro FROM b_bodegas WHERE bod_codbod=" & vg_codbod & ")) " & _
'           "AND   d.min_cencos = '" & cencos & "' " & _
'           "AND   d.min_codreg = " & codreg & " " & _
'           "AND   d.min_codser = " & codser & " " & _
'           "AND   val(mid(d.min_fecmin,1,6)) = " & anomes & " " & _
'           "AND   e.mid_tipmin = '" & tipmin & "'  " & _
'           "AND   a.pro_coduni = c.uni_codigo " & _
'           "AND   i.cpi_precos <= 0 AND h.ppd_propon = 0 " & _
'           "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) ORDER BY a.pro_nombre", vg_db, adOpenStatic
' Else
'    RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, c.uni_nombre " & _
'            "FROM b_productos a, a_unidad c, b_minuta d, b_minutadet e, b_recetadet f, b_ingrediente g, b_productospmpdia h, b_contlistpreing i " & _
'            "WHERE d.min_codigo = e.mid_codigo " & _
'            "AND   e.mid_codrec = f.red_codigo " & _
'            "AND   e.mid_tiprec = f.red_tiprec AND ((f.red_tiprec <> 0 AND f.red_cencos = '" & MuestraCasino(1) & "') OR (f.red_tiprec = 0 AND f.red_cencos = '0')) " & _
'            "AND   f.red_codpro = g.ing_codigo " & _
'            "AND   g.ing_codigo = i.cpi_coding " & _
'            "AND   i.cpi_codcom = a.pro_codigo " & _
'            "AND   a.pro_codigo = h.ppd_codpro " & _
'            "AND   i.cpi_cencos = '" & cencos & "' " & _
'            "AND   h.ppd_cencos = '" & cencos & "' " & _
'            "AND  (a.pro_codigo IN (SELECT bod_codpro FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " AND bod_canmer < 1) OR a.pro_codigo NOT IN (SELECT bod_codpro FROM b_bodegas WHERE bod_codbod = " & vg_codbod & ")) " & _
'            "AND   d.min_cencos = '" & cencos & "' " & _
'            "AND   d.min_codreg = " & codreg & " " & _
'            "AND   d.min_codser = " & codser & " " & _
'            "AND   convert(int,substring(convert(varchar(8),d.min_fecmin),1,6)) = " & anomes & " " & _
'            "AND   e.mid_tipmin = '" & tipmin & "'  " & _
'            "AND   a.pro_coduni = c.uni_codigo " & _
'            "AND   i.cpi_precos <= 0 AND h.ppd_propon = 0 AND h.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
'            "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) ORDER BY a.pro_nombre", vg_db, adOpenStatic
'End If
'If Not RS.EOF Then
'   Do While Not RS.EOF
'      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
'      vaSpread1.Row = vaSpread1.MaxRows
'      vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!pro_codigo
'      vaSpread1.Col = 2: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!pro_nombre
'      vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!uni_nombre
'      vaSpread1.Col = 4: vaSpread1.text = ""
'      RS.MoveNext
'   Loop
'   RS.Close: Set RS = Nothing
'   vaSpread1.SetActiveCell 4, 1
'   vaSpread1.Visible = True
'   fg_descarga
'Else
'   vg_codigo = ""
'   RS.Close: Set RS = Nothing
'   fg_descarga
'   Gl_Ac_Botones Me, 1, 3, modo
'   Toolbar1.Buttons(1).Visible = False
'   Toolbar1.Buttons(2).Visible = False
'   Toolbar1.Buttons(5).Visible = False
'   Toolbar1.Buttons(6).Visible = False
'   Toolbar1.Buttons(15).Visible = False
'   Toolbar1.Buttons(16).Visible = False
'   MsgBox "No existe información, con valores ceros", vbCritical + vbOKOnly, Msgtitulo
'End If
''-------> Borrar tablas temporales
'If Trim(aAp) <> "" And vg_tipbase = "1" Then vg_db.Execute "DROP TABLE " & aAp & ""
'Exit Sub
'Man_Error:
'fg_descarga
'MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
'ins_log_error Date & Time & Err & ":  " & Error$(Err)
'End Sub
