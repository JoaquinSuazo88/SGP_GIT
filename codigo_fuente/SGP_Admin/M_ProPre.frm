VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_ProPre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso Precio Producto"
   ClientHeight    =   7725
   ClientLeft      =   1185
   ClientTop       =   1560
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   7485
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
         TabIndex        =   1
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
      TabIndex        =   2
      Top             =   0
      Width           =   7485
      _ExtentX        =   13203
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
Option Explicit
Option Compare Text

Private modo        As String
Private RS          As New ADODB.Recordset
Private RS1         As New ADODB.Recordset
Private est         As Boolean
Private Msgtitulo   As String
Private vCenCos     As String
Private vCodReg     As Long
Private vCodSer     As Long
Private vTipMin     As String
Private vAnoMes     As Long
Private FeHasta     As Long

Private Sub Form_Activate()
    Call fg_descarga
'TraerFechaCierre
End Sub

Private Sub Form_Load()
    Me.HelpContextID = vg_OpcM
    Me.Height = 8205
    Me.Width = 7590
    Msgtitulo = "Ingreso Precio Producto"
    fg_centra Me
    modo = "": est = False
    Call Gl_Mo_Botones(Me, 1)
    Call Gl_Ac_Botones(Me, 1, 1, modo)
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
    Toolbar1.Buttons(15).Visible = False
    Toolbar1.Buttons(16).Visible = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long

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
            Call LlenarListaPrecio(vCenCos, vCodReg, vCodSer, vAnoMes, vTipMin, FeHasta)
        Case 10
            If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
            Call LlenarListaPrecio(vCenCos, vCodReg, vCodSer, vAnoMes, vTipMin, FeHasta)
            modo = "": Gl_Ac_Botones Me, 1, 1, modo
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(15).Visible = False
            Toolbar1.Buttons(16).Visible = False
        Case 12
            Dim vCodPro As Long, vPrePro As Double, FecIni As Date, FecFin As Date, coding As String
            With vaSpread1
                
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = 4
                    If Trim(.text) <> "" And Val(.text) > 0 Then
                       .Col = 1: vCodPro = .text
                       .Col = 4: vPrePro = .text
                       FecIni = dBoM(vg_ciedia)
                       FecFin = CDate(vg_ciedia) - 1
                       vg_db.Execute "UPDATE cas_b_productospmpdia SET ppd_propon = " & vPrePro & " WHERE ppd_cecori = '" & vg_codcasino & "' AND ppd_codpro = '" & vCodPro & "' AND ppd_propon < 1 AND ppd_fecdia >= " & Format(FecIni, "yyyymmdd") & " AND ppd_fecdia <= " & Format(FecFin, "yyyymmdd") & ""
                       RS.Open "SELECT DISTINCT pri_coding FROM b_productosing WHERE pri_codpro = '" & vCodPro & "'", vg_db, adOpenStatic
                       If Not RS.EOF Then
                            RS1.Open "select cosing      = Round(AVG(" & vPrePro & " / b.pro_facing), 2)" & _
                                    " FROM  CAS_b_contlistpreing a, b_productos b " & _
                                    " WHERE a.cpi_cecori = '" & vg_codcasino & "' " & _
                                    " AND   a.cpi_codcom = b.pro_codigo " & _
                                    " AND   a.cpi_coding = '" & RS!pri_coding & "'", vg_db, adOpenStatic
                       
'                          RS1.Open "SELECT Round(AVG(a.ppd_propon/c.pro_facing), 2) AS cosing " & _
'                                  "FROM  b_productospmpdia a, b_productosing b, b_productos c " & _
'                                  "WHERE b.pri_codpro = c.pro_codigo " & _
'                                  "AND   c.pro_codigo = a.ppd_codpro " & _
'                                  "AND   a.ppd_cencos = '" & vg_codcasino & "' " & _
'                                  "AND   a.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
'                                  "AND   a.ppd_propon > 0 AND b.pri_coding = '" & RS!pri_coding & "'", vg_db, adOpenStatic
                          If Not RS.EOF Then
                             vg_db.Execute "UPDATE CAS_b_contlistpreing SET cpi_feccos = " & Format(Date, "yyyymmdd") & ", cpi_precos = " & IIf(IsNull(RS1!cosing), 0, RS1!cosing) & " " & _
                                           "WHERE cpi_cecori = '" & vg_codcasino & "' AND cpi_coding = '" & RS!pri_coding & "'"
                          End If
                          RS1.Close: Set RS1 = Nothing
                       End If
                       RS.Close: Set RS = Nothing
                     End If
                Next i
            
            End With
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
    If Err = -2147467259 Or 2147217900 Then: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
    If Err = 3034 Then Exit Sub
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
If modo = "" Then modo = "M"
If modo = "M" And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = False: Toolbar1.Buttons(15).Visible = False: Toolbar1.Buttons(16).Visible = False
End Sub

Sub LlenarListaPrecio(cencos As String, codReg As Long, codser As Long, anomes As Long, TipMin As String, FecHasta As Long)
On Error GoTo Man_Error
fg_carga ""
vaSpread1.Visible = False
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &HC0FFFF
vaSpread1.MaxRows = 0
vCenCos = cencos
vCodReg = codReg
vCodSer = codser
vAnoMes = anomes
Let FeHasta = FecHasta
vTipMin = TipMin
Dim aAp As String
    RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, g.uni_nombre FROM b_productos a With(NoLock), cas_b_receta b With(NoLock), cas_b_recetadet c With(NoLock)," & _
        " cas_b_minuta d With(NoLock), cas_b_minutadet e With(NoLock), " & _
        " cas_b_contlistpreing f With(NoLock), a_unidad g With(NoLock) " & _
        " Where d.min_cecori = e.mid_cecori  " & _
        " AND   d.min_codigo = e.mid_codigo  " & _
        " AND   d.min_cecori = b.rec_cecori  " & _
        " AND   b.rec_cecori = c.red_cecori  " & _
        " AND   e.mid_codrec = b.rec_codigo  " & _
        " AND   b.rec_codigo = c.red_codigo  " & _
        " AND   d.min_cecori = f.cpi_cecori  " & _
        " AND   c.red_codpro = f.cpi_coding  " & _
        " AND   f.cpi_codcom = a.pro_codigo AND   a.pro_coduni = g.uni_codigo " & _
        " AND   d.min_cecori    = '" & cencos & "' " & _
        " AND   d.min_codreg    = " & codReg & " " & _
        " AND   d.min_codser    = " & codser & " " & _
        " AND   convert(int,substring(convert(varchar(8),d.min_fecmin),1,6)) >= " & anomes & " " & _
        " AND   convert(int,substring(convert(varchar(8),d.min_fecmin),1,6)) <= " & FecHasta & " " & _
        " AND   f.cpi_precos    <=  0", vg_db, adOpenStatic
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
'        Me.Show
'        Me.Hide
'        Unload Me
    End If
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub


