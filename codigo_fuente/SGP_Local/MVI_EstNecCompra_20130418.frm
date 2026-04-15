VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Object = "{1DF3AFED-99E0-4474-9900-954B8FD24E86}#1.0#0"; "ChilkatMail2.dll"
Begin VB.Form MVI_EstNecCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descarga de pedidos de compra en Excel MVI"
   ClientHeight    =   8700
   ClientLeft      =   2925
   ClientTop       =   2145
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   13935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1425
      Index           =   1
      Left            =   3330
      TabIndex        =   3
      Top             =   120
      Width           =   7095
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1380
         TabIndex        =   0
         Top             =   285
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
         _ExtentY        =   556
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
         ThreeDOutsideHighlightColor=   16777215
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
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   16777215
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   10
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
         ButtonAlign     =   1
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1380
         TabIndex        =   1
         Top             =   630
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1658
         _ExtentY        =   556
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
         ThreeDOutsideHighlightColor=   16777215
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
         ButtonStyle     =   1
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
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   16777215
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   "04/2013"
         DateCalcMethod  =   4
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/yyyy"
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   0
         IncHour         =   0
         IncMinute       =   0
         IncSecond       =   0
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   360
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cargar Información"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha(mm/aa)"
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
         TabIndex        =   6
         Top             =   690
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contrato"
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
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2610
         Picture         =   "MVI_EstNecCompra.frx":0000
         Top             =   195
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3045
         TabIndex        =   4
         Top             =   285
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3090
         TabIndex        =   7
         Top             =   330
         Width           =   3855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8700
      Left            =   13425
      TabIndex        =   8
      Top             =   0
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   15346
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   6795
      Left            =   30
      TabIndex        =   9
      Top             =   1710
      Width           =   13275
      _Version        =   393216
      _ExtentX        =   23416
      _ExtentY        =   11986
      _StockProps     =   64
      AutoClipboard   =   0   'False
      ButtonDrawMode  =   1
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
      MaxCols         =   9
      MaxRows         =   1
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "MVI_EstNecCompra.frx":030A
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _Version        =   393216
      _ExtentX        =   2143
      _ExtentY        =   1085
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
      MaxCols         =   5
      MaxRows         =   1
      SpreadDesigner  =   "MVI_EstNecCompra.frx":0871
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MVI_EstNecCompra.frx":0B7F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ACTIVEZIPLib.ActiveZip AZ1 
      Left            =   1560
      Top             =   240
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin CHILKATMAILLib2Ctl.ChilkatMailMan2 oMail 
      Left            =   2640
      OleObjectBlob   =   "MVI_EstNecCompra.frx":0F19
      Top             =   600
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP oFTP 
      Left            =   1920
      OleObjectBlob   =   "MVI_EstNecCompra.frx":0F3D
      Top             =   870
   End
End
Attribute VB_Name = "MVI_EstNecCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim Fecha As Long
Dim Msgtitulo As String
Dim est As Boolean, etapa5 As Boolean, aAp1 As String, aAp2 As String
Dim estexi As Boolean
Dim vecdes() As Variant
Dim sql As String

Private Sub Command1_Click()
  If vaSpread1.MaxRows = 0 Then Exit Sub
  'exporta el recordset a excel
  Call Exportar_ADO_Excel(Me, sql, "C:\NecCompraExcel.xls")
End Sub

Private Sub Form_Activate()
fg_descarga
'-------> Traer fecha cierre día
 TraerFechaCierre
End Sub

Private Sub LimpiarControles()
    vaSpread1.MaxRows = 0
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
fg_carga ""
Me.Height = 9180
Me.Width = 14025
fg_centra Me
est = True
Msgtitulo = "Pedido Mensual Ruta"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_ImportarInventario", , tbrDefault, "A_ImportarInventario"): BtnX.Visible = True: BtnX.ToolTipText = "Importar Archivo "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_ExportarInventario", , tbrDefault, "A_ExportarInventario"): BtnX.Visible = False: BtnX.ToolTipText = "Exportar Excel "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
vaSpread1.MaxRows = 0
fpDateTime1.text = Format(Date, "mm/yyyy")
fpText.Enabled = ModCasino
Image1.Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda.Caption = MuestraCasino(2)
Dim X As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFC0, &H800000)
est = False
fg_descarga

'MVI_ImpArchivos.Show
Call LimpiarControles

End Sub

Private Sub LlenarGrilla(sql As String)

RS.Open sql, vg_db, adOpenStatic

vaSpread1.MaxRows = 0


If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      vaSpread2.Col = 1: vaSpread2.text = RS(0)
      vaSpread2.Col = 2: vaSpread2.text = RS(1)
      vaSpread2.Col = 3: vaSpread2.text = RS(2)
      vaSpread2.Col = 4: vaSpread2.text = RS(3)
      vaSpread2.Col = 5: vaSpread2.text = RS(4)
      vaSpread2.Col = 5: vaSpread2.text = RS(5)
      vaSpread2.Col = 5: vaSpread2.text = RS(6)
      vaSpread2.Col = 5: vaSpread2.text = RS(7)
      RS.MoveNext
   Loop
End If

Set RS = Nothing
End Sub

Private Sub MoverDatos()
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim codtip As Long, fecenv As Long, fecval As Long, diaini As Long, diafin As Long, i As Long, j As Long, X As Long
Dim aAp As String, proc1 As String, proc2 As String, proc3 As String
Dim canped As Double, proped As Double, estfij As Boolean
Dim sql1 As String, sql2 As String, sql3 As String, sql4 As String, sql5 As String, sql6 As String
Dim fecxin As Long, fecxfi As Long
Dim fechoy As Date, fecaux As Date
Dim fecini As Date, fecfin As Date, fecped As Date, EstPed As Integer, fecpin As Long, fecpfi As Long
fg_carga ""
vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = &HC0FFFF
Fecha = 0: codtip = 0: fecenv = 0: auxing = 0
Fecha = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)
estexi = True
sql1 = " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) "
   'de aqui comienza calculo necesidades
   '-------> Traer consumo minuta teorica actual
   fecenv = 1
   aAp = Trim(vg_NUsr) & "_tmp_PedMensual"
   '-------> Creo tabla temporal y chequeo si existe antes
   fg_CheckTmp aAp
'      RS.Open "SELECT e.min_fecmin, b.cpi_coding AS ing_codigo, a.pro_codigo, a.pro_codtip, " & _
'              "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))) AS cantidad1, " & _
'              "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))/a.pro_facing) AS cantidad2, 0 AS ped_fecped INTO " & aAp & " " & _
'              "FROM b_productos a, b_contlistpreing b, b_receta c, b_recetadet d, b_minuta e, b_minutadet f, a_servicio h " & _
'              "WHERE e.min_codigo = f.mid_codigo AND f.mid_codrec = d.red_codigo AND f.mid_tiprec = d.red_tiprec AND ((d.red_tiprec <> 0 AND d.red_cencos = '" & MuestraCasino(1) & "') OR (d.red_tiprec = 0 AND d.red_cencos = '0')) " & _
'              "AND   d.red_codigo = c.rec_codigo AND d.red_codpro = b.cpi_coding AND b.cpi_cencos = '" & MuestraCasino(1) & "' AND b.cpi_codped = a.pro_codigo " & _
'              "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) " & _
'              "AND   e.min_cencos = '" & Trim(fpText.text) & "' AND convert(int,substring(convert(varchar(8),e.min_fecmin),1,6)) = " & Fecha & " " & _
'              "AND   f.mid_tipmin = '1' AND a.pro_facing > 0 and e.min_codser = h.ser_codigo and h.ser_activo = '1' GROUP BY e.min_fecmin, b.cpi_coding, a.pro_codigo, a.pro_codtip", vg_db, adOpenForwardOnly
'
      Set RS = vg_db.Execute("SELECT e.min_fecmin, b.cpi_coding AS ing_codigo, a.pro_codigo, a.pro_codtip, " & _
              "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))) AS cantidad1, " & _
              "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))/a.pro_facing) AS cantidad2, 0 AS ped_fecped INTO " & aAp & " " & _
              "FROM b_productos a, b_contlistpreing b, b_receta c, b_recetadet d, b_minuta e, b_minutadet f, a_servicio h " & _
              "WHERE e.min_codigo = f.mid_codigo AND f.mid_codrec = d.red_codigo AND f.mid_tiprec = d.red_tiprec AND ((d.red_tiprec <> 0 AND d.red_cencos = '" & MuestraCasino(1) & "') OR (d.red_tiprec = 0 AND d.red_cencos = '0')) " & _
              "AND   d.red_codigo = c.rec_codigo AND d.red_codpro = b.cpi_coding AND b.cpi_cencos = '" & MuestraCasino(1) & "' AND b.cpi_codped = a.pro_codigo " & _
              "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) " & _
              "AND   e.min_cencos = '" & Trim(fpText.text) & "' AND convert(int,substring(convert(varchar(8),e.min_fecmin),1,6)) = " & Fecha & " " & _
              "AND   f.mid_tipmin = '1' AND a.pro_facing > 0 and e.min_codser = h.ser_codigo and h.ser_activo = '1' GROUP BY e.min_fecmin, b.cpi_coding, a.pro_codigo, a.pro_codtip")
   
   
   Set RS = Nothing
   diaini = 0: diafin = 0
   diaini = 1: diafin = Mid(dEoM("01/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)), 1, 2)
   '-------> Rutina buscar estructura fija
   Set RS = vg_db.Execute("SELECT DISTINCT a.mif_codreg, a.mif_codser FROM  b_minutafija a, a_servicio b WHERE a.mif_cencos='" & Trim(fpText.text) & "' and a.mif_codser = b.ser_codigo and b.ser_activo = '1'")
   If Not RS.EOF Then
      Do While Not RS.EOF
         estfij = False
         '-------> Buscar datos estructura fija día
            RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
                     "WHERE mfd_cencos='" & Trim(fpText.text) & "' " & _
                     "AND   mfd_codreg=" & RS!mif_codreg & " " & _
                     "AND   mfd_codser=" & RS!mif_codser & " " & _
                     "AND   convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) >= " & Fecha & " AND convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) <= " & Fecha & " AND mfd_tipmin = '1'", vg_db, adOpenStatic
         If Not RS1.EOF Then estfij = True
         RS1.Close: Set RS1 = Nothing
         fecval = 0
         If Not estfij Then
            '-------> Buscar fecha mayor de estructura fija
            RS1.Open "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija " & _
                     "WHERE mif_cencos='" & Trim(fpText.text) & "' " & _
                     "AND   mif_codreg=" & RS!mif_codreg & " " & _
                     "AND   mif_codser=" & RS!mif_codser & "", vg_db, adOpenStatic
            If Not RS1.EOF Then fecval = IIf(IsNull(RS1!fecval), 0, RS1!fecval)
            RS1.Close: Set RS1 = Nothing
            If fecval > 0 Then
               '-------> Traer estructura fija
               For i = diaini To diafin
                   If fecval <= Fecha & Right("0" & i, 2) Then
                      vg_db.Execute "INSERT INTO " & aAp & " SELECT " & Fecha & Right("0" & i, 2) & " AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                                    "0 AS cantidad1, a.mif_canpro AS cantidad2, 0 AS ped_fecped " & _
                                    "FROM  b_minutafija a, b_productos b, b_ingrediente c, b_productosing d " & _
                                    "WHERE a.mif_codpro=b.pro_codigo " & _
                                    "AND   b.pro_codigo=d.pri_codpro " & _
                                    "AND   c.ing_codigo=d.pri_coding " & _
                                    "AND  (b.pro_fecven>" & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                    "AND   a.mif_cencos='" & Trim(fpText.text) & "' " & _
                                    "AND   a.mif_codreg=" & RS!mif_codreg & " " & _
                                    "AND   a.mif_codser=" & RS!mif_codser & " " & _
                                    "AND   a.mif_fecval=" & fecval & " " & _
                                    "AND   a.mif_dianro=" & fg_NumDia(Trim(Left(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                    "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
                    
                      vg_db.Execute "INSERT INTO " & aAp & " SELECT " & Fecha & Right("0" & i, 2) & " AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                                    "0 AS cantidad1, a.mif_canpro AS cantidad2, 0 AS ped_fecped " & _
                                    "FROM  b_minutafija a, b_productos b, b_ingrediente c, b_productosing d " & _
                                    "WHERE a.mif_codpro=b.pro_codigo " & _
                                    "AND   b.pro_codigo=d.pri_codpro " & _
                                    "AND   c.ing_codigo=d.pri_coding " & _
                                    "AND  (b.pro_fecven>" & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                    "AND   a.mif_cencos='" & Trim(fpText.text) & "' " & _
                                    "AND   a.mif_codreg=" & RS!mif_codreg & " " & _
                                    "AND   a.mif_codser=" & RS!mif_codser & " " & _
                                    "AND   a.mif_fecval=" & fecval & " " & _
                                    "AND   a.mif_dianro=" & fg_NumDia(Trim(Left(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                    "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
                   End If
                   Set RS1 = Nothing
               Next i
            End If
         ElseIf estfij Then
             '-------> Calcular datos desde tabla estructura fija día
                vg_db.Execute "INSERT INTO " & aAp & " SELECT a.mfd_fecha AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                              "0 AS cantidad1, a.mfd_canpro AS cantidad2, 0 AS ped_fecped " & _
                              "FROM b_minutafijadia a, b_productos b, b_ingrediente c, b_productosing d " & _
                              "WHERE a.mfd_codpro = b.pro_codigo " & _
                              "AND   b.pro_codigo = d.pri_codpro " & _
                              "AND   c.ing_codigo = d.pri_coding " & _
                              "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                              "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                              "AND   a.mfd_codreg = " & RS!mif_codreg & " " & _
                              "AND   a.mfd_codser = " & RS!mif_codser & " " & _
                              "AND convert(int,substring(convert(varchar(8),a.mfd_fecha),1,6)) = " & Fecha & " AND a.mfd_tipmin = '1' " & _
                              "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
             
                vg_db.Execute "INSERT INTO " & aAp & " SELECT a.mfd_fecha AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                              "0 AS cantidad1, a.mfd_canpro AS cantidad2, 0 AS ped_fecped " & _
                              "FROM b_minutafijadia a, b_productos b, b_ingrediente c, b_productosing d " & _
                              "WHERE a.mfd_codpro = b.pro_codigo " & _
                              "AND   b.pro_codigo = d.pri_codpro " & _
                              "AND   c.ing_codigo = d.pri_coding " & _
                              "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                              "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                              "AND   a.mfd_codreg = " & RS!mif_codreg & " " & _
                              "AND   a.mfd_codser = " & RS!mif_codser & " " & _
                              "AND convert(int,substring(convert(varchar(8),a.mfd_fecha),1,6)) = " & Fecha & " AND a.mfd_tipmin = '1' " & _
                              "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
         End If
         RS.MoveNext
      Loop
   End If
   'hasta aca la tempo tiene datos
   RS.Close: Set RS = Nothing
   
   'fin calculo necesidades
   'Nueva cosecha Marcelo verdugo
   
   ' aca colocar la query que muestra los datos en la grilla y exporta los datos al excel.
   ' query que estará detallada en los docs. de CWalther.
   
   'pto. 3.2.5
   'Bloque 1:
   
   'si la tabla existe, la destruye
   vg_db.Execute "IF object_id('tmp_mes_actual') IS NOT NULL BEGIN DROP TABLE tmp_mes_actual END"
   
   sql = " select * into tmp_mes_actual from " & Trim(vg_NUsr) & "_tmp_PedMensual"
       
    Set RS = vg_db.Execute(sql)
   
   Set RS = Nothing
   
   Call Calcula_MesSgte
   
   Exit Sub
   
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If Not IsDate(fpDateTime1.text) Then Exit Sub
MoverDatos
End Sub

Private Sub fpDateTime2_Change()
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If Not IsDate(fpDateTime2.text) Then Exit Sub
End Sub

Private Sub fpText_Change()
If fpText.text = "" Then fpayuda.Caption = "": Exit Sub
Toolbar1.Buttons(1).Enabled = True: Toolbar1.Buttons(3).Enabled = False: vaSpread1.MaxRows = 0
RS.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & fpText.text & "' AND cli_tipo = 0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda.Caption = "": Exit Sub
fpayuda.Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    Image1_Click
End Select
End Sub

Private Sub Image1_Click()
vg_left = fpayuda.Left + 2300
vg_nombre = "": vg_codigo = ""
B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
B_TabEst.Show 1
Me.Refresh
If vg_codigo = "" Then Exit Sub
fpText.text = vg_codigo: fpayuda.Caption = vg_nombre
Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(8).Enabled = False: vaSpread1.MaxRows = 0
If Me.Visible Then fpDateTime1.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codpro As String, coding As String, codsac As String, i As Integer
Dim fechasis As Long, fecdes As Long, nrosem As Long
Dim canmin As Double, cospro As Double, cosali As Double, CosDes As Double
Dim canped As Double, stoact As Double, proped As Double, pedpro As Double
Dim aAp As String, persac As String, sql1 As String, sql2 As String
On Error GoTo Man_Error
Select Case Button.Index
Case 1 '-------> Importar
     MVI_ImpArchivos.Show 1
Case 3 '-------> Exportar
    Toolbar1.Enabled = False
    P_EIInve.Inicio "Exportar Pedido Mensual Ruta", "EP", 0
    P_EIInve.Show 1
    Toolbar1.Enabled = True
Case 5 '-------> Salir
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RS As New ADODB.Recordset
Select Case Button.Index
Case 1
    If Not IsDate(fpDateTime1.text) Then Exit Sub
    '-------> Validar si la minuta es teorica normal
    sql1 = " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) "
    RS.Open "SELECT DISTINCT a.min_codigo FROM b_minuta a, a_servicio b WHERE a.min_codser = b.ser_codigo and b.ser_activo = '1' and a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND " & sql1 & " = " & Format(fpDateTime1.text, "yyyymm") & " AND min_indblo IN (2,11)", vg_db, adOpenStatic
    If Not RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Existe Bloque Minuta, pedido se cancela", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    
    '-------> Validar si existe datos ruta carga
    RS.Open "select top 1 id_carga from ruta_compras", vg_db, adOpenStatic
    If RS.EOF Then
       RS.Close: Set RS = Nothing: MsgBox "No existe datos cargados rutas compras, pedido se cancela", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    End If
    RS.Close: Set RS = Nothing
    
    '-------> Validar si existe datos convenios
    RS.Open "select top 1 Reg_info from convenios_mvi", vg_db, adOpenStatic
    If RS.EOF Then
       RS.Close: Set RS = Nothing: MsgBox "No existe datos cargados convenios, pedido se cancela", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    End If
    RS.Close: Set RS = Nothing
    
    MoverDatos
End Select
End Sub

Private Sub Calcula_MesSgte()
Dim RS As New ADODB.Recordset
If Not IsDate(fpDateTime1.text) Then Exit Sub
'-------> Validar si la minuta es teorica normal
sql1 = " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) "
RS.Open "SELECT DISTINCT a.min_codigo FROM b_minuta a, a_servicio b WHERE a.min_codser = b.ser_codigo and b.ser_activo = '1' and a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND " & sql1 & " = " & Val(Fecha) & " AND min_indblo IN (2,11)", vg_db, adOpenStatic
If Not RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Existe Bloque Minuta, pedido se cancela", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
RS.Close: Set RS = Nothing
MoverDatosMesSgte
End Sub

Private Sub MoverDatosMesSgte()
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim codtip As Long, fecenv As Long, fecval As Long, diaini As Long, diafin As Long, i As Long, j As Long, X As Long
Dim aAp As String, proc1 As String, proc2 As String, proc3 As String
Dim canped As Double, proped As Double, estfij As Boolean
Dim sql1 As String, sql2 As String, sql3 As String, sql4 As String, sql5 As String, sql6 As String
Dim fecxin As Long, fecxfi As Long
Dim fechoy As Date, fecaux As Date
Dim fecini As Date, fecfin As Date, fecped As Date, EstPed As Integer, fecpin As Long, fecpfi As Long
Dim NomExcelZip As String
fg_carga ""
vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = &HC0FFFF
Fecha = 0: codtip = 0: fecenv = 0: auxing = 0
Fecha = Format(DateAdd("m", 1, "01/" & fpDateTime1.text), "YYYYMM")
estexi = True
sql1 = " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) "
   'de aqui comienza calculo necesidades
   '-------> Traer consumo minuta teorica actual
   fecenv = 1
   aAp = Trim(vg_NUsr) & "_tmp_PedMensual"
   '-------> Creo tabla temporal y chequeo si existe antes
   fg_CheckTmp aAp
   RS.Open "SELECT e.min_fecmin, b.cpi_coding AS ing_codigo, a.pro_codigo, a.pro_codtip, " & _
           "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))) AS cantidad1, " & _
           "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))/a.pro_facing) AS cantidad2, 0 AS ped_fecped INTO " & aAp & " " & _
           "FROM b_productos a, b_contlistpreing b, b_receta c, b_recetadet d, b_minuta e, b_minutadet f, a_servicio h " & _
           "WHERE e.min_codigo = f.mid_codigo AND f.mid_codrec = d.red_codigo AND f.mid_tiprec = d.red_tiprec AND ((d.red_tiprec <> 0 AND d.red_cencos = '" & MuestraCasino(1) & "') OR (d.red_tiprec = 0 AND d.red_cencos = '0')) " & _
           "AND   d.red_codigo = c.rec_codigo AND d.red_codpro = b.cpi_coding AND b.cpi_cencos = '" & MuestraCasino(1) & "' AND b.cpi_codped = a.pro_codigo " & _
           "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) " & _
           "AND   e.min_cencos = '" & Trim(fpText.text) & "' AND convert(int,substring(convert(varchar(8),e.min_fecmin),1,6)) = " & Fecha & " " & _
           "AND   f.mid_tipmin = '1' AND a.pro_facing > 0 and e.min_codser = h.ser_codigo and h.ser_activo = '1' GROUP BY e.min_fecmin, b.cpi_coding, a.pro_codigo, a.pro_codtip", vg_db, adOpenForwardOnly
   Set RS = Nothing
   diaini = 0: diafin = 0
   diaini = 1: diafin = Mid(dEoM("01/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)), 1, 2)
   '-------> Rutina buscar estructura fija
   RS.Open "SELECT DISTINCT a.mif_codreg, a.mif_codser FROM  b_minutafija a, a_servicio b WHERE a.mif_cencos='" & Trim(fpText.text) & "' and a.mif_codser = b.ser_codigo and b.ser_activo ='1'", vg_db, adOpenStatic
   If Not RS.EOF Then
      Do While Not RS.EOF
         estfij = False
         '-------> Buscar datos estructura fija día
            RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
                     "WHERE mfd_cencos='" & Trim(fpText.text) & "' " & _
                     "AND   mfd_codreg=" & RS!mif_codreg & " " & _
                     "AND   mfd_codser=" & RS!mif_codser & " " & _
                     "AND   convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) >= " & Fecha & " AND convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) <= " & Fecha & " AND mfd_tipmin = '1'", vg_db, adOpenStatic
         If Not RS1.EOF Then estfij = True
         RS1.Close: Set RS1 = Nothing
         fecval = 0
         If Not estfij Then
            '-------> Buscar fecha mayor de estructura fija
            RS1.Open "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija " & _
                     "WHERE mif_cencos='" & Trim(fpText.text) & "' " & _
                     "AND   mif_codreg=" & RS!mif_codreg & " " & _
                     "AND   mif_codser=" & RS!mif_codser & "", vg_db, adOpenStatic
            If Not RS1.EOF Then fecval = IIf(IsNull(RS1!fecval), 0, RS1!fecval)
            RS1.Close: Set RS1 = Nothing
            If fecval > 0 Then
               '-------> Traer estructura fija
               For i = diaini To diafin
                   If fecval <= Fecha & Right("0" & i, 2) Then
                      vg_db.Execute "INSERT INTO " & aAp & " SELECT " & Fecha & Right("0" & i, 2) & " AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                                    "0 AS cantidad1, a.mif_canpro AS cantidad2, 0 AS ped_fecped " & _
                                    "FROM  b_minutafija a, b_productos b, b_ingrediente c, b_productosing d " & _
                                    "WHERE a.mif_codpro=b.pro_codigo " & _
                                    "AND   b.pro_codigo=d.pri_codpro " & _
                                    "AND   c.ing_codigo=d.pri_coding " & _
                                    "AND  (b.pro_fecven>" & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                    "AND   a.mif_cencos='" & Trim(fpText.text) & "' " & _
                                    "AND   a.mif_codreg=" & RS!mif_codreg & " " & _
                                    "AND   a.mif_codser=" & RS!mif_codser & " " & _
                                    "AND   a.mif_fecval=" & fecval & " " & _
                                    "AND   a.mif_dianro=" & fg_NumDia(Trim(Left(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                    "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
                    
                      vg_db.Execute "INSERT INTO " & aAp & " SELECT " & Fecha & Right("0" & i, 2) & " AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                                    "0 AS cantidad1, a.mif_canpro AS cantidad2, 0 AS ped_fecped " & _
                                    "FROM  b_minutafija a, b_productos b, b_ingrediente c, b_productosing d " & _
                                    "WHERE a.mif_codpro=b.pro_codigo " & _
                                    "AND   b.pro_codigo=d.pri_codpro " & _
                                    "AND   c.ing_codigo=d.pri_coding " & _
                                    "AND  (b.pro_fecven>" & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                    "AND   a.mif_cencos='" & Trim(fpText.text) & "' " & _
                                    "AND   a.mif_codreg=" & RS!mif_codreg & " " & _
                                    "AND   a.mif_codser=" & RS!mif_codser & " " & _
                                    "AND   a.mif_fecval=" & fecval & " " & _
                                    "AND   a.mif_dianro=" & fg_NumDia(Trim(Left(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                    "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
                   End If
                   Set RS1 = Nothing
               Next i
            End If
         ElseIf estfij Then
             '-------> Calcular datos desde tabla estructura fija día
                vg_db.Execute "INSERT INTO " & aAp & " SELECT a.mfd_fecha AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                              "0 AS cantidad1, a.mfd_canpro AS cantidad2, 0 AS ped_fecped " & _
                              "FROM b_minutafijadia a, b_productos b, b_ingrediente c, b_productosing d " & _
                              "WHERE a.mfd_codpro = b.pro_codigo " & _
                              "AND   b.pro_codigo = d.pri_codpro " & _
                              "AND   c.ing_codigo = d.pri_coding " & _
                              "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                              "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                              "AND   a.mfd_codreg = " & RS!mif_codreg & " " & _
                              "AND   a.mfd_codser = " & RS!mif_codser & " " & _
                              "AND convert(int,substring(convert(varchar(8),a.mfd_fecha),1,6)) = " & Fecha & " AND a.mfd_tipmin = '1' " & _
                              "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
             
                vg_db.Execute "INSERT INTO " & aAp & " SELECT a.mfd_fecha AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                              "0 AS cantidad1, a.mfd_canpro AS cantidad2, 0 AS ped_fecped " & _
                              "FROM b_minutafijadia a, b_productos b, b_ingrediente c, b_productosing d " & _
                              "WHERE a.mfd_codpro = b.pro_codigo " & _
                              "AND   b.pro_codigo = d.pri_codpro " & _
                              "AND   c.ing_codigo = d.pri_coding " & _
                              "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                              "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                              "AND   a.mfd_codreg = " & RS!mif_codreg & " " & _
                              "AND   a.mfd_codser = " & RS!mif_codser & " " & _
                              "AND convert(int,substring(convert(varchar(8),a.mfd_fecha),1,6)) = " & Fecha & " AND a.mfd_tipmin = '1' " & _
                              "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
         End If
         RS.MoveNext
      Loop
   End If
   'hasta aca la tempo tiene datos
   RS.Close: Set RS = Nothing
   
   'fin calculo necesidades
   'Nueva cosecha Marcelo verdugo
   
   ' aca colocar la query que muestra los datos en la grilla y exporta los datos al excel.
   ' query que estará detallada en los docs. de CWalther.
   
   'pto. 3.2.5
   'Bloque 1:
   
      'si la tabla existe, la destruye
   vg_db.Execute "IF object_id('tmp_mes_sgte') IS NOT NULL BEGIN DROP TABLE tmp_mes_sgte END"

   
   sql = " select * into tmp_mes_sgte from " & Trim(vg_NUsr) & "_tmp_PedMensual"
   
      
   Set RS = vg_db.Execute(sql)
   Set RS = Nothing
   'dispara sp ppal.
   sql = " MVI_NEC_COMPRA"
   sql = sql & " '" & fpText & "'"
   sql = sql & " , " & Val(Left(fpDateTime1.text, 2))
   sql = sql & " ," & Val(Right(fpDateTime1.text, 4))
'    Sql = Sql & " "

   Set RS = vg_db.Execute(sql)
   
    
'-------> Inicio LLenar grilla
vaSpread1.MaxRows = 0
If Not RS.EOF Then
   Toolbar1.Buttons(3).Enabled = True
Else
   Toolbar1.Buttons(3).Enabled = False
End If
Do While Not RS.EOF
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = 1
    vaSpread1.text = RS(0)
    vaSpread1.Col = 2
    vaSpread1.text = RS(1)
    vaSpread1.Col = 3
    vaSpread1.text = RS(2)
    vaSpread1.Col = 4
    vaSpread1.text = RS(3)
    vaSpread1.Col = 5
    vaSpread1.text = RS(4)
    vaSpread1.Col = 6
    vaSpread1.text = RS(5)
    vaSpread1.Col = 7
    vaSpread1.text = RS(6)
    vaSpread1.Col = 8
    vaSpread1.text = RS(7)
    vaSpread1.Col = 9
    vaSpread1.text = RS(8)
   
    RS.MoveNext
Loop
RS.Close: Set RS = Nothing

'-------> validar si existe pedido

If vaSpread1.MaxRows < 1 Then
   fg_descarga
   MsgBox "Por favor verificar si existen " & VgLinea & VgLinea & "- Rutas para la fecha consultada " & VgLinea & "- Convenios vigentes para la fecha consultada " & VgLinea, vbInformation + vbOKOnly, Msgtitulo
   Exit Sub
End If

'-------> Grabar minuta
GenerarMinutaReal

'-------> generar excel
'Dim sql As String
Dim NomExcel As String
NomExcel = dir_trabajo & "Pedido" & Trim(fpText.text) & Format(Date, "yyyymmdd") & ".xls"

If Dir(NomExcel) <> "" Then Kill NomExcel   'borrar base datos si existe
If Dir(Mid((NomExcel), 1, Len((NomExcel)) - 3) & "txt") <> "" Then Kill Mid((NomExcel), 1, Len((NomExcel)) - 3) & "txt"
Open Mid((NomExcel), 1, Len((NomExcel)) - 3) & "txt" For Output As #1

MVI_EstNecCompra.vaSpread1.Row = 0
MVI_EstNecCompra.vaSpread1.Col = 1
sql = " " & MVI_EstNecCompra.vaSpread1.text & "|"
MVI_EstNecCompra.vaSpread1.Col = 2
sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
MVI_EstNecCompra.vaSpread1.Col = 3
sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
MVI_EstNecCompra.vaSpread1.Col = 4
sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
MVI_EstNecCompra.vaSpread1.Col = 5
sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
MVI_EstNecCompra.vaSpread1.Col = 6
sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
MVI_EstNecCompra.vaSpread1.Col = 7
sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
MVI_EstNecCompra.vaSpread1.Col = 8
sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
MVI_EstNecCompra.vaSpread1.Col = 9
sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
Print #1, sql
       
For i = 1 To MVI_EstNecCompra.vaSpread1.MaxRows
    sql = ""
'    PB.Value = Val((i / MVI_EstNecCompra.vaSpread1.MaxRows) * 100)
    MVI_EstNecCompra.vaSpread1.Row = i
    MVI_EstNecCompra.vaSpread1.Col = 1
    sql = MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 2
    sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 3
    sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 4
    sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 5
    sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 6
    sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 7
    sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 8
    sql = sql & Format(MVI_EstNecCompra.vaSpread1.text, "mm/dd/yyyy") & "|"
    MVI_EstNecCompra.vaSpread1.Col = 9
    sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
    Print #1, sql
'           Print #1, Trim(RS!pro_codigo) & ";" & Trim(RS!pro_nombre) & ";" & Trim(RS!uni_nomcor) & ";" & Round(RS!tin_stofis, vg_DCa)
Next i
Close #1
Set XL = CreateObject("Excel.application")
XL.Workbooks.OpenText Mid((NomExcel), 1, Len((NomExcel)) - 3) & "txt", , 1, 1, , , , , , , True, "|"
XL.ActiveWorkbook.SaveAs Filename:=NomExcel, _
                                  FileFormat:=xlNormal, password:="", WriteResPassword:="", _
                                  ReadOnlyRecommended:=False, CreateBackup:=False
XL.Quit
Set XL = Nothing

'-------> 1 Comprimir archivo excel
NomExcelZip = dir_trabajo & "Pedido" & Trim(fpText.text) & Format(Date, "yyyymmdd") & ".zip"
'-------> verificar si existe archivo zip destino si existe borrar
If Dir(NomExcelZip) <> "" Then Kill NomExcelZip
AZ1.CreateZip NomExcelZip, "": AZ1.AddFile NomExcel, "", True, "": AZ1.Close
'-------> verificar si existe archivo mdb destino si existe borrar
If Dir(NomExcel) <> "" Then Kill NomExcel

'-------> Traer dirección correo
Dim emailpedido As String
RS.Open "select isnull(par_valor,'') as par_valor from a_param where par_cencos = '" & Trim(fpText.text) & "' and par_codigo = 'emailenped'", vg_db, adOpenStatic
If Not RS.EOF Then
   emailpedido = RS!par_valor
End If
RS.Close: Set RS = Nothing

'-------> Enviar correo
Dim cBody As String
cBody = ""
cBody = "Generación Automatica Pedidos " & Format(fpDateTime1.text, "mm/yyyy") & VgLinea & VgLinea
cBody = cBody & "IMPORTANTE: " & VgLinea
cBody = cBody & "Este correo es informativo, favor no responder a esta dirección de correo, ya que no se encuentra habilitada para recibir mensajes." & VgLinea & VgLinea
cBody = cBody & "Atte." & VgLinea
cBody = cBody & "SGP Chile" & VgLinea

SendMail oMail, "SGP : Pedido Insumo (" & Trim(fpText.text) & ") " & Trim(fpayuda.Caption), cBody, NomExcelZip, "Email Pedido", emailpedido

fg_descarga
    'CW
    
    'si se descomenta la linea de abajo, llenará la grilla con los datos del proc. almac.
   'Call LlenarGrilla(Sql)
   
   'Shell ("""C:\Program Files\Microsoft Office\OFFICE11\EXCEL.EXE"" ""c:\EstNecBChile.xls"""), vbNormalFocus
   
   
   'si se borra este bloque, se omitira la carga del archivo excel
   
'   Dim xl As Excel.Application


 '   Set xl = CreateObject("excel.Application")
    
 '   xl.Workbooks.Open ("c:\EstNecBChile.xls") ' substitute your file here
 '     xl.Visible = False
    
 '    Print xl.Cells(1, 5).Value
    ' can do any operation here like running macro etc
    
 '   xl.Visible = True
    
 '     Set xl = Nothing
 
 'borrar hasta aca.
   
End Sub

Private Sub GenerarMinutaReal()
Dim RS As New ADODB.Recordset
Dim anomes As Long
Dim fecped As Long
Dim CodProdSap As String
Dim CantPedido As Double
Dim EstPed As Boolean
Dim fechasis As Long

fechasis = 0 'Val(Mid(Date, 7, 4)) & Mid(Date, 4, 2) & Mid(Date, 1, 2)"
anomes = Format(fpDateTime1.text, "yyyymm")
EstPed = True

    RS.Open "select convert(varchar(8),getdate(),112) as fecsis", vg_db, adOpenStatic
    If Not RS.EOF Then
       fechasis = RS!fecsis
    End If
    RS.Close: Set RS = Nothing
    With vaSpread1
        
        .Enabled = False
        '-------> ver minutapedidos
        RS.Open "select ped_fecenv from b_minutapedidos where ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & anomes & " AND ped_tipped = 1", vg_db, adOpenStatic
        If Not RS.EOF Then
            vg_db.Execute "delete b_minutapedidos where ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & anomes & " AND ped_tipped = 1"
            EstPed = False
        End If
        RS.Close: Set RS = Nothing
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Row = i
            vaSpread1.Col = 1
            '-------> Mover codigo producto sap
            vaSpread1.Col = 5
            CodProdSap = Trim(vaSpread1.text)
            '-------> Mover fecha pedido
            vaSpread1.Col = 8
            fecped = Format(vaSpread1.text, "yyyymmdd")
            '-------> Mover cantidad pedido
            vaSpread1.Col = 9
            CantPedido = vaSpread1.text
            vg_db.Execute "insert into b_minutapedidos (ped_codcas, ped_fecped, ped_anomes, ped_tipped, ped_coding, ped_codpro, ped_codsac, ped_canmin, ped_canped, ped_fecenv, ped_stoact, ped_proped, ped_ordrec, ped_conrea, ped_persac, ped_semsac ) " & _
                          " values ('" & LimpiaDato(Trim(fpText.text)) & "', " & fecped & ", " & anomes & ", 1, '" & CodProdSap & "', '" & CodProdSap & "', '', 0,  " & CantPedido & ", " & fecped & ", 0, 0, 0, 0, '', 0)"
        Next i
        If EstPed And vaSpread1.MaxRows > 0 Then
        '-------> Grabar minuta costo teórico & real
        sql1 = " AND convert(int,substring(convert(varchar(8),b.min_fecmin),1,6)) = " & anomes & " "
        vg_db.Execute "DELETE b_minutacosto FROM b_minutacosto WHERE mic_cencos = '" & MuestraCasino(1) & "' AND mic_fecval IN (SELECT a.mid_fecval FROM b_minutadet a, b_minuta b WHERE a.mid_codigo = b.min_codigo AND b.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & sql1 & ") AND mic_codpro = (SELECT top 1 d.red_codpro FROM b_minutadet a, b_minuta b, b_receta c, b_recetadet d WHERE a.mid_codigo = b.min_codigo AND a.mid_codrec = c.rec_codigo AND c.rec_codigo = d.red_codigo AND a.mid_codrec = d.red_codigo AND a.mid_tiprec = d.red_tiprec AND ((d.red_tiprec<>0 AND d.red_cencos = '" & MuestraCasino(1) & "') OR (d.red_tiprec = 0 AND d.red_cencos = '0')) AND b.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & sql1 & ")"
        Dim tmin As String
        i = 1
        tmin = "1"
        sql1 = " AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & anomes & " "
        vg_db.Execute "INSERT INTO b_minutacosto (mic_cencos, mic_fecval, mic_tipmin, mic_codpro, mic_cospro) SELECT DISTINCT a.min_cencos, " & fechasis & ", '" & tmin & "', d.red_codpro, (SELECT top 1 cpi_precos FROM b_contlistpreing WHERE cpi_cencos = '" & MuestraCasino(1) & "' AND cpi_coding = d.red_codpro) AS cpi_precos " & _
                      "FROM b_minuta a, b_minutadet b, b_receta c, b_recetadet d, a_servicio e " & _
                      "WHERE a.min_codigo = b.mid_codigo " & _
                      "AND   b.mid_codrec = c.rec_codigo " & _
                      "AND   c.rec_codigo = d.red_codigo " & _
                      "AND   a.min_cencos = '" & MuestraCasino(1) & "' AND b.mid_tipmin = '1' and a.min_codser = e.ser_codigo and e.ser_activo = '1' " & _
                      " " & sql1 & " AND d.red_codpro not in (SELECT top 1 mic_codpro from b_minutacosto where mic_fecval = " & fechasis & " and mic_cencos = a.min_cencos AND mic_tipmin = '" & tmin & "' AND mic_codpro = d.red_codpro)"
        tmin = "2"
        vg_db.Execute "INSERT INTO b_minutacosto (mic_cencos, mic_fecval, mic_tipmin, mic_codpro, mic_cospro) SELECT DISTINCT a.min_cencos, " & fechasis & ", '" & tmin & "', d.red_codpro, (SELECT top 1 cpi_precos FROM b_contlistpreing WHERE cpi_cencos = '" & MuestraCasino(1) & "' AND cpi_coding = d.red_codpro) AS cpi_precos " & _
                      "FROM b_minuta a, b_minutadet b, b_receta c, b_recetadet d, a_servicio e " & _
                      "WHERE a.min_codigo = b.mid_codigo " & _
                      "AND   b.mid_codrec = c.rec_codigo " & _
                      "AND   c.rec_codigo = d.red_codigo " & _
                      "AND   a.min_cencos = '" & MuestraCasino(1) & "' AND b.mid_tipmin = '1' and a.min_codser = e.ser_codigo and e.ser_activo = '1' " & _
                      " " & sql1 & " AND d.red_codpro not in (SELECT top 1 mic_codpro from b_minutacosto where mic_fecval = " & fechasis & " and mic_cencos = a.min_cencos AND mic_tipmin = '" & tmin & "' AND mic_codpro = d.red_codpro)"
        
        vg_db.Execute "UPDATE b_minutacosto SET b_minutacosto.mic_cospro = b_contlistpreing.cpi_precos FROM b_minutacosto, b_contlistpreing " & _
                      "Where b_contlistpreing.cpi_coding = b_minutacosto.mic_codpro AND b_minutacosto.mic_cencos = b_contlistpreing.cpi_cencos AND  b_minutacosto.mic_cencos = '" & MuestraCasino(1) & "' And b_minutacosto.mic_fecval = " & fechasis & " And b_minutacosto.mic_tipmin IN ('1','2')"
        
        vg_db.Execute "UPDATE b_minutacosto SET mic_cospro = 0 WHERE mic_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND mic_fecval = " & fechasis & " AND mic_cospro IS NULL"
        
        '-------> Traer estructura fija
        Dim fecval As Long
        RS.Open "SELECT DISTINCT a.mif_cencos, a.mif_codreg, a.mif_codser FROM b_minutafija a, a_servicio b " & _
                "WHERE a.mif_codser = b.ser_codigo and b.ser_activo = '1' and a.mif_cencos = '" & LimpiaDato(Trim(fpText.text)) & "'", vg_db, adOpenStatic
        If Not RS.EOF Then
           Do While Not RS.EOF
              DoEvents
              '-------> Validar si existe estructura fija día
              RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
                       "WHERE  mfd_cencos = '" & RS!mif_cencos & "' " & _
                       "AND    mfd_codreg = " & RS!mif_codreg & " " & _
                       "AND    mfd_codser = " & RS!mif_codser & " " & _
                       "AND    convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) = " & anomes & " " & _
                       "AND    mfd_tipmin = '1'", vg_db, adOpenStatic
              If RS1.EOF Then
                 RS1.Close: Set RS1 = Nothing
                 '-------> Buscar fecha mayor de estructura fija
                 fecval = 0
                 RS1.Open "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija WHERE mif_cencos = '" & Trim(fpText.text) & "' AND mif_codreg = " & RS!mif_codreg & " AND mif_codser = " & RS!mif_codser & "", vg_db, adOpenStatic
                 If Not RS1.EOF Then fecval = IIf(IsNull(RS1!fecval), 0, RS1!fecval)
                 RS1.Close: Set RS1 = Nothing
                 If fecval > 0 Then
                    '-------> Traer estructura fija
                    For i = 1 To Val(Mid(dEoM("26/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)), 1, 2))
                        If fecval <= Fecha & fg_pone_cero(Str(i), 2) Then
                           '-------> Grabar estructura fija día teorica
                           vg_db.Execute "INSERT INTO b_minutafijadia (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro) " & _
                                         "SELECT a.mif_cencos, a.mif_codreg, a.mif_codser, " & anomes & fg_pone_cero(i, 2) & ", b.pro_codigo, '1', a.mif_canpro, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") " & _
                                         "FROM b_minutafija a, b_productos b " & _
                                         "WHERE a.mif_codpro = b.pro_codigo " & _
                                         "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                         "AND   a.mif_cencos = '" & Trim(fpText.text) & "' " & _
                                         "AND   a.mif_codreg = " & RS!mif_codreg & " " & _
                                         "AND   a.mif_codser = " & RS!mif_codser & " " & _
                                         "AND   a.mif_fecval = " & fecval & " " & _
                                         "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(anomes & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(anomes & Right("0" & i, 2), 2)) - 2))) & " " & _
                                         "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0)"
                        End If
                        
                        Set RS1 = Nothing
                    Next i
                 End If
              Else
                 RS1.Close: Set RS1 = Nothing
                 '-------> Actualizar precio propon tabla estructura fija x día
                 vg_db.Execute "UPDATE b_minutafijadia SET b_minutafijadia.mfd_cospro = b_productospmpdia.ppd_propon FROM  b_productospmpdia WHERE b_minutafijadia.mfd_codpro = b_productospmpdia.ppd_codpro " & _
                               "AND b_minutafijadia.mfd_cencos = '" & Trim(fpText.text) & "' AND b_minutafijadia.mfd_codreg = " & RS!mif_codreg & " AND b_minutafijadia.mfd_codser = " & RS!mif_codser & " AND convert(int,substring(convert(varchar(8),b_minutafijadia.mfd_fecha),1,6)) = " & anomes & " AND b_minutafijadia.mfd_tipmin = '1' AND b_productospmpdia.ppd_cencos = '" & MuestraCasino(1) & "'"
              End If
              RS.MoveNext
           Loop
        End If
        RS.Close: Set RS = Nothing
        
        '-------> Generar minuta costo estructura fija
        '-------> Eliminar estructura fija día real si existen datos
           vg_db.Execute "DELETE b_minutafijadia FROM b_minutafijadia WHERE mfd_cencos = '" & Trim(fpText.text) & "' AND convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) = " & anomes & " AND mfd_tipmin = '2'"
           '-------> Grabar estructura fija día real
           vg_db.Execute "INSERT INTO b_minutafijadia (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro) SELECT a.mfd_cencos, a.mfd_codreg, a.mfd_codser, a.mfd_fecha, a.mfd_codpro, '2', a.mfd_canpro, (SELECT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") " & _
                         "FROM b_minutafijadia a, b_productos b, a_servicio c " & _
                         "WHERE a.mfd_codpro = b.pro_codigo " & _
                         "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                         "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                         "AND   convert(int,substring(convert(varchar(8),a.mfd_fecha),1,6)) = " & anomes & " " & _
                         "AND   a.mfd_tipmin = '1' " & _
                         "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) and a.mfd_codser = c.ser_codigo and c.ser_activo = '1'  "
        '-------> Traer total de receta desde planificación de minutas y luego calcular costo
        RS.Open "SELECT COUNT(b.mid_codrec) AS nreg FROM b_minuta a, b_minutadet b, a_servicio c WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                "AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & anomes & " AND b.mid_tipmin = '1' and a.min_codser = c.ser_codigo and c.ser_activo = '1'  ", vg_db, adOpenStatic
        If RS.EOF Or RS!nreg < 1 Then RS.Close: Set RS = Nothing:  MsgBox "No existe información", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
        ReDim vecrec(RS!nreg, 4)
        RS.Close: Set RS = Nothing
        For i = 1 To UBound(vecrec)
            DoEvents
            vecrec(i, 1) = 0 '-------> codigo receta
            vecrec(i, 2) = 0 '-------> tipo receta
            vecrec(i, 3) = 0 '-------> costo receta alimentación
            vecrec(i, 4) = 0 '-------> costo receta desechable
        Next i
        i = 1
        RS.Open "SELECT DISTINCT b.mid_codrec, b.mid_tiprec FROM b_minuta a, b_minutadet b, a_servicio c WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                "AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & anomes & " AND b.mid_tipmin = '1' and a.min_codser = c.ser_codigo and c.ser_activo = '1'  ", vg_db, adOpenStatic
        If RS.EOF Then RS.Close: Set RS = Nothing: vg_db.RollbackTrans: MsgBox "No existe información", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
        Do While Not RS.EOF
           DoEvents
           vecrec(i, 1) = RS!mid_codrec
           vecrec(i, 2) = RS!mid_tiprec
           vecrec(i, 3) = Format(fg_CalCtoRecPlan(fechasis, 1, RS!mid_codrec, IIf(IsNull(RS!mid_tiprec), 0, RS!mid_tiprec), (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','"))), fg_Pict(6, 2))
           vecrec(i, 4) = Format(fg_CalCtoRecPlan(fechasis, 1, RS!mid_codrec, IIf(IsNull(RS!mid_tiprec), 0, RS!mid_tiprec), (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))), fg_Pict(6, 2))
           RS.MoveNext: i = i + 1
        Loop
        RS.Close: Set RS = Nothing

        '-------> Generar planificación real & actualizar costo teórica
        RS.Open "SELECT b.* FROM b_minuta a, b_minutadet b, a_servicio c WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                "AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & anomes & " AND b.mid_tipmin = '1' and a.min_codser = c.ser_codigo and c.ser_activo = '1'", vg_db, adOpenStatic
        If RS.EOF Then RS.Close: Set RS = Nothing: vg_db.RollbackTrans: MsgBox "No existe información", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
        Do While Not RS.EOF
           DoEvents
           For i = 1 To UBound(vecrec)
               If RS!mid_codrec = vecrec(i, 1) And RS!mid_tiprec = vecrec(i, 2) Then
                  cosali = CCur(vecrec(i, 3))
                  CosDes = CCur(vecrec(i, 4))
                  Exit For
               End If
           Next
           vg_db.Execute "INSERT INTO b_minutadet (mid_codigo, mid_tipmin, mid_numlin, mid_estser, mid_codrec, mid_numrac, mid_descri, mid_cosrec, mid_fecval, mid_tiprec, mid_nummer, mid_rec5eta, mid_cosdes, mid_modmina, mid_modminb) " & _
                         "VALUES (" & RS!mid_codigo & ", '2', " & RS!mid_numlin & ", " & RS!mid_estser & ",  " & RS!mid_codrec & ", " & IIf(IsNull(RS!mid_numrac), "NULL", RS!mid_numrac) & ", '" & RS!mid_descri & "', " & cosali & ", " & fechasis & ", " & RS!mid_tiprec & ", 0, " & IIf(IsNull(RS!mid_rec5eta) Or Trim(RS!mid_rec5eta) = "", "Null", RS!mid_rec5eta) & ", " & CosDes & ", '0', '0')"
           vg_db.Execute "UPDATE b_minutadet SET mid_fecval = " & fechasis & ", mid_cosrec = " & cosali & ", mid_cosdes = " & CosDes & " WHERE mid_codigo = " & RS!mid_codigo & " AND mid_tipmin = '1' AND mid_codrec = " & RS!mid_codrec & " AND mid_tiprec = " & RS!mid_tiprec & ""
           RS.MoveNext
        Loop
        RS.Close: Set RS = Nothing
        '-------> Bloquear planificación teórica
        vg_db.Execute "UPDATE b_minuta SET min_indblo = 1, min_racrea = min_racteo WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),min_fecmin),1,6)) = " & anomes & " AND (min_indblo = 0 OR (min_indblo) IS NULL)"
        RS.Open "SELECT DISTINCT a.min_codreg, a.min_codser, a.min_fecmin, a.min_racrea FROM b_minuta a, a_servicio b WHERE a.min_codser = b.ser_codigo and b.ser_activo = '1' and a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & anomes & " AND a.min_racrea > 0", vg_db, adOpenStatic
        '-------> Grabar raciones en minutas raciones
        Do While Not RS.EOF
           DoEvents
           RS1.Open "SELECT * FROM b_minutaraciones WHERE mir_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND mir_codreg = " & RS!min_codreg & " AND mir_codser = " & RS!min_codser & " AND mir_fecmin = " & RS!min_fecmin & " AND mir_rutcli = 'PRODUCIDAS'", vg_db, adOpenStatic
           If RS1.EOF Then
              vg_db.Execute "INSERT INTO b_minutaraciones (mir_cencos,mir_codreg,mir_codser,mir_fecmin,mir_rutcli,mir_nrorac,mir_nroguia,mir_codcli) VALUES ('" & LimpiaDato(Trim(fpText.text)) & "', " & RS!min_codreg & ", " & RS!min_codser & ", " & RS!min_fecmin & ", 'PRODUCIDAS', " & RS!min_racrea & ", NULL, '')"
           Else
              vg_db.Execute "UPDATE b_minutaraciones SET mir_nrorac = " & RS!min_racrea & " WHERE mir_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND mir_codreg = " & RS!min_codreg & " AND mir_codser = " & RS!min_codser & " AND mir_fecmin = " & RS!min_fecmin & " AND mir_rutcli = 'PRODUCIDAS' AND mir_nrorac < 1"
           End If
           RS1.Close: Set RS1 = Nothing
           RS.MoveNext
        Loop
        RS.Close: Set RS = Nothing
        fg_descarga
'        Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(8).Enabled = True: Frame2.Enabled = False
        MsgBox "Generación pedido Finalizado Sin Problema", vbInformation + vbOKOnly, Msgtitulo
'        Toolbar1.Enabled = True
'        Frame1(1).Enabled = True
        End If
        .Enabled = True

    End With

End Sub

Function CalcularDiasFeriados(cencos As String, Fecha As Variant) As String
Dim RS3 As New ADODB.Recordset
Dim diafer As Boolean
Dim sql1 As String
diafer = True
'-------> validar si existen dias feriado
sql1 = IIf(vg_tipbase = "1", " AND cdate(CFI_Fecha) = '" & Fecha & "' ", " AND Convert(VarChar(10), CFI_Fecha, 103) = '" & Fecha & "' ")
RS3.Open "SELECT CFI_Fecha FROM b_Fecha_Inhabiles WHERE CFI_CeCo = '" & cencos & "' " & sql1 & "", vg_db, adOpenStatic
If Not RS3.EOF Then
   RS3.Close: Set RS3 = Nothing
   diaferi = True
   Do While diaferi
      Fecha = CDate((Fecha)) + 1
      sql1 = IIf(vg_tipbase = "1", " AND cdate(CFI_Fecha) = '" & Fecha & "' ", " AND Convert(VarChar(10), CFI_Fecha, 103) = '" & Fecha & "' ")
      RS3.Open "SELECT CFI_Fecha FROM b_Fecha_Inhabiles WHERE CFI_CeCo = '" & cencos & "' " & sql1 & "", vg_db, adOpenStatic
      If RS3.EOF Then diaferi = False
      RS3.Close: Set RS3 = Nothing
   Loop
Else
   RS3.Close: Set RS3 = Nothing
End If
CalcularDiasFeriados = Fecha
End Function

