VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form C_ExpRec 
   Caption         =   "Exportación Recetas"
   ClientHeight    =   7245
   ClientLeft      =   1950
   ClientTop       =   1950
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10095
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
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
         Index           =   3
         Left            =   1755
         TabIndex        =   15
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   2760
         TabIndex        =   14
         Top             =   1305
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   2760
         TabIndex        =   11
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
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
         Index           =   2
         Left            =   1755
         TabIndex        =   10
         Top             =   1005
         Width           =   660
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2760
         TabIndex        =   6
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label1 
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
         Left            =   1755
         TabIndex        =   4
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
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
         Left            =   1755
         TabIndex        =   3
         Top             =   645
         Width           =   750
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   2805
         TabIndex        =   7
         Top             =   285
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2805
         TabIndex        =   8
         Top             =   645
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   2805
         TabIndex        =   12
         Top             =   1005
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   7
         Left            =   2790
         TabIndex        =   16
         Top             =   1350
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   10095
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4815
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   9855
         _Version        =   393216
         _ExtentX        =   17383
         _ExtentY        =   8493
         _StockProps     =   64
         ColsFrozen      =   3
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
         MaxRows         =   18
         SpreadDesigner  =   "C_ExpRec.frx":0000
         VisibleCols     =   3
         VisibleRows     =   18
         ScrollBarTrack  =   3
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   5160
         Visible         =   0   'False
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   318
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7245
      Left            =   10125
      TabIndex        =   9
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   12779
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   1005
      Left            =   2655
      TabIndex        =   13
      Top             =   6660
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   1773
      _Version        =   393217
      TextRTF         =   $"C_ExpRec.frx":0430
   End
End
Attribute VB_Name = "C_ExpRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim acodcas As String, acodreg As Long, acodser As Long, aanomes As Long, numreg As Long, atipmin As String
Dim FechaIni As Long
Dim FechaFin As Long
Dim aAp As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_centra Me
aAp = ""
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Excel "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "word", , tbrDefault, "word"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Excel "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
End Sub

Sub LlenarExporReceta(tfor As String, codcas As String, codreg As Long, codser As Long, anomes As Long, tipmin As String)
Dim isalto As Integer
Dim codreceta As Long, iRow As Long, i As Long, condia As Long, auxfecha As Long
Dim cosreceta As Double, canreceta As Double, totgralreceta As Double, nomR As String, MetR As String
fg_carga ""
'------- Rutina frecuencia de recetas
acodcas = codcas
acodreg = codreg
acodser = codser
aanomes = anomes
atipmin = tipmin
Me.Caption = tfor
Msgtitulo = tfor
RS1.Open RutinaLectura.Cliente(1, codcas, ""), vg_db, adOpenStatic
If Not RS1.EOF Then fpayuda(0).Caption = Trim(RS1!cli_nombre) & IIf(tipmin = 1, " (Planificación Teórica)", " (Planificación Real)")
RS1.Close: Set RS1 = Nothing
RS1.Open RutinaLectura.Regimen(2, codreg, ""), vg_db, adOpenStatic
If Not RS1.EOF Then fpayuda(1).Caption = Trim(RS1!reg_nombre)
RS1.Close: Set RS1 = Nothing
fpayuda(3).Caption = Meses("01/" & Mid(fg_pone_cero(anomes, 6), 5, 2) & "/" & Mid(fg_pone_cero(anomes, 6), 1, 4)) & " " & Mid(fg_pone_cero(anomes, 6), 1, 4)
RS1.Open RutinaLectura.Servicio(8, codser, ""), vg_db, adOpenStatic
If Not RS1.EOF Then fpayuda(5).Caption = Trim(RS1!ser_nombre)
RS1.Close: Set RS1 = Nothing
With vaSpread1(0)
    .Visible = False
    .Row = -1: .Col = -1
    .MaxRows = 0
    aAp = Trim(vg_NUsr) & "_tmp_ExpRecetas"
    '------- Creo tabla temporal y chequeo si existe antes
    fg_CheckTmp aAp
    RS_Dato.Open "SELECT DISTINCT a.min_codser, d.ser_nombre, b.mid_codrec, c.red_codpro, c.red_nroite, c.red_canpro, c.red_pctapr, " & _
                 "c.red_pctcoc, c.red_pctnut INTO " & aAp & " " & _
                 "FROM   b_minuta a, b_minutadet b, b_recetadet c, a_servicio d " & _
                 "WHERE b.mid_codigo = a.min_codigo " & _
                 "AND   a.min_codser = d.ser_codigo " & _
                 "AND   b.mid_codrec = c.red_codigo " & _
                 "AND   b.mid_tiprec = c.red_tiprec AND ((c.red_tiprec <> 0 AND c.red_cencos = '" & MuestraCasino(1) & "') OR (c.red_tiprec = 0 AND c.red_cencos = '0')) " & _
                 "AND   a.min_cencos = '" & codcas & "' " & _
                 "AND   a.min_codreg = " & codreg & " " & _
                 "AND   a.min_codser = " & codser & " " & _
                 "AND   a.min_fecmin >= " & anomes & "01" & " " & _
                 "AND   b.mid_tipmin = '" & tipmin & "'  " & _
                 "AND   a.min_fecmin <= " & Format(dEoM("01/" & Mid(anomes, 5, 2) & "/" & Mid(anomes, 1, 4)), "yyyymmdd") & "", vg_db, adOpenForwardOnly
    Set RS_Dato = Nothing
    RS_Dato.Open "SELECT a.min_codser, a.ser_nombre, a.mid_codrec, b.rec_nombre, b.rec_nomfan, c.ing_codigo, (SELECT DISTINCT cpi_precos FROM b_contlistpreing WHERE cpi_cencos='" & MuestraCasino(1) & "' AND cpi_coding=c.ing_codigo) AS cpi_precos, a.red_nroite, c.ing_nombre, d.unm_nomcor, " & _
                 "a.red_canpro, a.red_pctapr, a.red_pctcoc, a.red_pctnut, b.rec_catdie, b.rec_tippla, b.rec_metpre " & _
                 "FROM  " & aAp & " a, b_receta b, b_ingrediente c, a_unidadmed d " & _
                 "WHERE a.mid_codrec = b.rec_codigo " & _
                 "AND   a.red_codpro = c.ing_codigo " & _
                 "AND   c.ing_unimed = d.unm_codigo " & _
                 "ORDER BY a.min_codser, b.rec_nombre, b.rec_nomfan, a.red_nroite", vg_db, adOpenForwardOnly
    If RS_Dato.EOF Then RS_Dato.Close: Set RS_Dato = Nothing: Exit Sub
    .Visible = False
    .MaxRows = 0
    nomR = "": RT1.text = ""
    numreg = 0
    Do While Not RS_Dato.EOF
        If nomR <> Trim(RS_Dato!rec_nomfan) Then
            If Trim(RT1.text) <> "" Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 1: .TypeHAlign = TypeHAlignLeft: .text = RT1.text
            ElseIf nomR <> "" Then
    '            .MaxRows = .MaxRows + 1
    '            .Row = .MaxRows
            End If
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
            .Font.Bold = True
            .Font.Size = 9
            
            .TypeHAlign = TypeHAlignLeft
            .text = Trim(RS_Dato!rec_nomfan)
            
            nomR = Trim(RS_Dato!rec_nomfan)
        End If
        If IsNull(RS_Dato!rec_metpre) Then
           RT1.TextRTF = ""
        Else
           Trim (RS_Dato!rec_metpre)
        End If
    '    RT1.TextRTF = IIf(IsNull(RS_Dato!rec_metpre) Or Trim(RS_Dato!rec_metpre) = "", "", Trim(RS_Dato!rec_metpre))
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1: .TypeHAlign = TypeHAlignLeft: .text = Trim(RS_Dato!ing_nombre)
        .Col = 2: .TypeHAlign = TypeHAlignLeft: .text = Trim(RS_Dato!unm_nomcor)
        .Col = 3: .TypeHAlign = TypeHAlignRight: .text = Format(RS_Dato!red_canpro, fg_Pict(6, vg_RDCa))
        RS_Dato.MoveNext: i = i + 1: numreg = numreg + 1
    Loop
    'RS1.Close: Set RS1 = Nothing
    If Trim(RT1.text) <> "" Then
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1: .TypeHAlign = TypeHAlignLeft: .text = RT1.text
    End If
    .Visible = True
End With
fg_descarga
End Sub

Sub LlenarExporRecetaBloque(tfor As String, codcas As String, codreg As Long, codser As Long, fecini As Long, fecfin As Long, tipmin As String)
Dim isalto As Integer
Dim codreceta As Long, iRow As Long, i As Long, condia As Long, auxfecha As Long
Dim cosreceta As Double, canreceta As Double, totgralreceta As Double, nomR As String, MetR As String
Dim RS1 As New ADODB.Recordset
fg_carga ""
'------- Rutina frecuencia de recetas
acodcas = codcas
acodreg = codreg
acodser = codser
FechaIni = fecini
FechaFin = fecfin
atipmin = tipmin
Me.Caption = tfor
Msgtitulo = tfor
RS1.Open RutinaLectura.Cliente(1, codcas, ""), vg_db, adOpenStatic
If Not RS1.EOF Then fpayuda(0).Caption = Trim(RS1!cli_nombre) & " (Planificación Bloque)"
RS1.Close: Set RS1 = Nothing
RS1.Open RutinaLectura.Regimen(2, codreg, ""), vg_db, adOpenStatic
If Not RS1.EOF Then fpayuda(1).Caption = Trim(RS1!reg_nombre)
RS1.Close: Set RS1 = Nothing
fpayuda(3).Caption = fg_Ctod1(fecini) & " - " & fg_Ctod1(fecfin)
RS1.Open RutinaLectura.Servicio(8, codser, ""), vg_db, adOpenStatic
If Not RS1.EOF Then fpayuda(5).Caption = Trim(RS1!ser_nombre)
RS1.Close: Set RS1 = Nothing
With vaSpread1(0)
    .Visible = False
    .Row = -1: .Col = -1
    .MaxRows = 0
    aAp = Trim(vg_NUsr) & "_tmp_ExpRecetas"
    '------- Creo tabla temporal y chequeo si existe antes
    fg_CheckTmp aAp
    RS_Dato.Open "SELECT DISTINCT a.min_codser, d.ser_nombre, b.mid_codrec, c.red_codpro, c.red_nroite, c.red_canpro, c.red_pctapr, " & _
                 "c.red_pctcoc, c.red_pctnut INTO " & aAp & " " & _
                 "FROM   b_minuta a, b_minutadet b, b_recetadet c, a_servicio d " & _
                 "WHERE b.mid_codigo = a.min_codigo " & _
                 "AND   a.min_codser = d.ser_codigo " & _
                 "AND   b.mid_codrec = c.red_codigo " & _
                 "AND   b.mid_tiprec = c.red_tiprec AND ((c.red_tiprec <> 0 AND c.red_cencos = '" & MuestraCasino(1) & "') OR (c.red_tiprec = 0 AND c.red_cencos = '0')) " & _
                 "AND   a.min_cencos = '" & codcas & "' " & _
                 "AND   a.min_codreg = " & codreg & " " & _
                 "AND   a.min_codser = " & codser & " " & _
                 "AND   a.min_fecmin >= " & fecini & " " & _
                 "AND   b.mid_tipmin = '" & tipmin & "' " & _
                 "AND   a.min_fecmin <= " & fecfin & "", vg_db, adOpenForwardOnly
    Set RS_Dato = Nothing
    RS_Dato.Open "SELECT a.min_codser, a.ser_nombre, a.mid_codrec, b.rec_nombre, b.rec_nomfan, c.ing_codigo, (SELECT DISTINCT cpi_precos FROM b_contlistpreing WHERE cpi_cencos='" & MuestraCasino(1) & "' AND cpi_coding=c.ing_codigo) AS cpi_precos, a.red_nroite, c.ing_nombre, d.unm_nomcor, " & _
                 "a.red_canpro, a.red_pctapr, a.red_pctcoc, a.red_pctnut, b.rec_catdie, b.rec_tippla, b.rec_metpre " & _
                 "FROM  " & aAp & " a, b_receta b, b_ingrediente c, a_unidadmed d " & _
                 "WHERE a.mid_codrec = b.rec_codigo " & _
                 "AND   a.red_codpro = c.ing_codigo " & _
                 "AND   c.ing_unimed = d.unm_codigo " & _
                 "ORDER BY a.min_codser, b.rec_nombre, b.rec_nomfan, a.red_nroite", vg_db, adOpenForwardOnly
    If RS_Dato.EOF Then RS_Dato.Close: Set RS_Dato = Nothing: Exit Sub
    .Visible = False
    .MaxRows = 0
    nomR = "": RT1.text = ""
    numreg = 0
    Do While Not RS_Dato.EOF
        If nomR <> Trim(RS_Dato!rec_nomfan) Then
            If Trim(RT1.text) <> "" Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 1: .TypeHAlign = TypeHAlignLeft: .text = RT1.text
            ElseIf nomR <> "" Then
    '            .MaxRows = .MaxRows + 1
    '            .Row = .MaxRows
            End If
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
            .Font.Bold = True
            .Font.Size = 9
            
            .TypeHAlign = TypeHAlignLeft
            .text = Trim(RS_Dato!rec_nomfan)
            
            nomR = Trim(RS_Dato!rec_nomfan)
        End If
        If IsNull(RS_Dato!rec_metpre) Then
           RT1.TextRTF = ""
        Else
           Trim (RS_Dato!rec_metpre)
        End If
    '    RT1.TextRTF = IIf(IsNull(RS_Dato!rec_metpre) Or Trim(RS_Dato!rec_metpre) = "", "", Trim(RS_Dato!rec_metpre))
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1: .TypeHAlign = TypeHAlignLeft: .text = Trim(RS_Dato!ing_nombre)
        .Col = 2: .TypeHAlign = TypeHAlignLeft: .text = Trim(RS_Dato!unm_nomcor)
        .Col = 3: .TypeHAlign = TypeHAlignRight: .text = Format(RS_Dato!red_canpro, fg_Pict(6, vg_RDCa))
        RS_Dato.MoveNext: i = i + 1: numreg = numreg + 1
    Loop
    'RS1.Close: Set RS1 = Nothing
    If Trim(RT1.text) <> "" Then
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1: .TypeHAlign = TypeHAlignLeft: .text = RT1.text
    End If
    .Visible = True
End With
fg_descarga
End Sub

Private Sub Form_Unload(Cancel As Integer)
If vaSpread1(0).MaxRows > O Then RS_Dato.Close: Set RS_Dato = Nothing
If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 2, 4
    If vaSpread1(0).MaxRows < 1 Then Exit Sub
    If Button.Index = 2 Then
       ExportarExcel
    Else
       If aanomes > 0 Then
          I_ListaRecetaPlanificacion acodcas, acodreg, acodser, aanomes & "01", Format(dEoM("01/" & Mid(aanomes, 5, 2) & "/" & Mid(aanomes, 1, 4)), "yyyymmdd") & "", numreg, atipmin
       Else
          I_ListaRecetaPlanificacion acodcas, acodreg, acodser, FechaIni, FechaFin, numreg, atipmin
       End If
    End If
Case 6
    Me.Hide
    Unload Me
End Select
End Sub

Sub ExportarExcel()
Dim NashXl As Excel.Application
Dim iRow As Long, irow2 As Long
fg_carga ""
Set NashXl = CreateObject("excel.application")
Set NashXl = New Excel.Application
NashXl.SheetsInNewWorkbook = 1
NashXl.Workbooks.Add
NashXl.Range("A1").Select
NashXl.ActiveCell.FormulaR1C1 = Label1(0).Caption & ": " & fpayuda(0).Caption
NashXl.Range("A2").Select
NashXl.ActiveCell.FormulaR1C1 = Label1(1).Caption & ": " & fpayuda(1).Caption
NashXl.Range("A3").Select
NashXl.ActiveCell.FormulaR1C1 = Label1(2).Caption & ": " & fpayuda(3).Caption
NashXl.Range("A4").Select
NashXl.ActiveCell.FormulaR1C1 = Label1(3).Caption & ": " & fpayuda(5).Caption
vaSpread1(0).AllowMultiBlocks = True
vaSpread1(0).SetSelection 1, -1, vaSpread1(0).MaxCols, vaSpread1(0).MaxRows + 3
vaSpread1(0).ClipboardCopy
iRow = vaSpread1(0).MaxRows + 5
'------- Pegar vaspread1(0) - Planilla Excel
NashXl.Range("A5").Select
NashXl.ActiveSheet.Paste
'------- Asignar color
'NashXl.Range("A1:D" & irow).Select
'With NashXl.Selection.Interior
'     .ColorIndex = 36
'     .Pattern = xlSolid
'End With
'------- Colorear titulo
NashXl.Range("A5:C5").Select
With NashXl.Selection.Interior
     .ColorIndex = 15
     .Pattern = xlSolid
End With
'------- Dibujar marco
NashXl.Range("A5:C" & iRow).Select
NashXl.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
NashXl.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With NashXl.Selection.Borders(xlEdgeLeft)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeTop)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeBottom)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeRight)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideVertical)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideHorizontal)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
NashXl.Range("C2" & ":" & "C" & iRow).Select
NashXl.Selection.NumberFormat = "#,##0." & fg_pone_cero(0, vg_RDCa)
'------- Aplicar totales

NashXl.Selection.Font.Bold = True
'With NashXl.Selection.Interior
'     .ColorIndex = 35
'     .Pattern = xlSolid
'End With
NashXl.Range("B" & iRow & ":" & "B" & 2).Select
NashXl.Selection.NumberFormat = "#,##0.00"
'------- Ajustar columna
NashXl.Cells.Select
NashXl.Cells.EntireColumn.AutoFit
vaSpread1(0).AllowMultiBlocks = False: vaSpread1(0).SetSelection 1, 0, vaSpread1(0).MaxCols, vaSpread1(0).MaxRows
fg_descarga
NashXl.Visible = True
End Sub
