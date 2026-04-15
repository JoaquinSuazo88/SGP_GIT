VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form C_FrePla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frecuencia Recetas"
   ClientHeight    =   5790
   ClientLeft      =   2100
   ClientTop       =   1620
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9885
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   9255
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9015
         _Version        =   393216
         _ExtentX        =   15901
         _ExtentY        =   5953
         _StockProps     =   64
         ColsFrozen      =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   46
         MaxRows         =   18
         SpreadDesigner  =   "C_FrePla.frx":0000
         VisibleCols     =   4
         VisibleRows     =   18
         ScrollBarTrack  =   3
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   11
         Left            =   7920
         TabIndex        =   9
         Top             =   3975
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   7920
         TabIndex        =   8
         Top             =   3675
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Costo Promedio Diario"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   5880
         TabIndex        =   7
         Top             =   3975
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Recetas Listadas"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   5880
         TabIndex        =   6
         Top             =   3675
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
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
         Index           =   2
         Left            =   1755
         TabIndex        =   3
         Top             =   1005
         Width           =   705
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
         TabIndex        =   2
         Top             =   645
         Width           =   750
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
         TabIndex        =   1
         Top             =   300
         Width           =   735
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2760
         TabIndex        =   12
         Top             =   600
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   2805
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   16
         Top             =   1005
         Width           =   4335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5790
      Left            =   9255
      TabIndex        =   10
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   10213
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "C_FrePla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private RS1     As New ADODB.Recordset
Private RS2     As New ADODB.Recordset
Private BtnX    As Variant
Private Dia     As Long

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Excel "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
End Sub

Sub LlenarFrecPlan(tfor As String, cencos As String, FechaHasta As String, codreg As Long, codser As Long, anomes As Long, tipmin As String)
'Sub LlenarFrecPlan(tfor As String, cencos As String, codreg As Long, codser As Long, anomes As Long, tipmin As String)
Dim isalto          As Integer
Dim sql1            As String
Dim codreceta       As Long
Dim iRow            As Long
Dim i               As Long
Dim condia          As Long
Dim auxfecha        As Long
Dim cosreceta       As Double
Dim canreceta       As Double
Dim totgralreceta   As Double
Dim ConLinea        As Long

fg_carga ""
'------- Rutina frecuencia de recetas
Me.Caption = tfor
Msgtitulo = tfor
RS1.Open RutinaLectura.Cliente(1, cencos, ""), vg_db, adOpenStatic
If Not RS1.EOF Then fpayuda(0).Caption = RS1!cli_nombre
RS1.Close: Set RS1 = Nothing
RS1.Open RutinaLectura.Regimen(2, codreg, ""), vg_db, adOpenStatic
If Not RS1.EOF Then fpayuda(1).Caption = RS1!reg_nombre
RS1.Close: Set RS1 = Nothing
RS1.Open RutinaLectura.Servicio(8, codser, ""), vg_db, adOpenStatic
If Not RS1.EOF Then fpayuda(3).Caption = RS1!ser_nombre
RS1.Close: Set RS1 = Nothing
With vaSpread1(0)
    .Row = -1: .Col = 1
    .BackColor = &HC0FFC0
    .Row = -1: .Col = 2
    .BackColor = &HC0FFC0
    .Row = -1: .Col = 3
    .BackColor = &HC0FFC0
    .Row = -1: .Col = 4
    .BackColor = &HC0FFC0
    
    .MaxRows = 0
    '------- Buscar Nº días
    sql1 = IIf(vg_tipbase = "1", " val(mid(a.min_fecmin,1 ,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1 ,6)) ")
    RS1.Open "SELECT DISTINCT a.min_fecmin FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo " & _
             "AND a.min_cencos = '" & cencos & "' AND a.min_codreg = " & codreg & " AND a.min_codser = " & codser & " " & _
             "AND " & sql1 & " = " & anomes & " AND b.mid_tipmin = '" & tipmin & "' ORDER BY a.min_fecmin", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
    Select Case fg_Dia(RS1!min_fecmin)
    Case 1
        Dia = 7
    Case 2
        Dia = 1
    Case 3
        Dia = 2
    Case 4
        Dia = 3
    Case 5
        Dia = 4
    Case 6
        Dia = 5
    Case 7
        Dia = 6
    End Select
    Dia = Dia - 1
    RS1.Close: Set RS1 = Nothing
    sql1 = IIf(vg_tipbase = "1", " val(mid(b.min_fecmin,1 ,6)) ", " convert(int,substring(convert(varchar(8),b.min_fecmin),1 ,6)) ")
    RS1.Open "SELECT c.mid_tipmin, c.mid_numlin, c.mid_codrec, c.mid_descri, isnull(c.mid_cosrec,0) as mid_cosrec, isnull(c.mid_cosdes,0) as mid_cosdes, b.min_fecmin, b.min_indblo, " & _
             "a.rec_codigo, a.rec_nombre, a.rec_nomfan, c.mid_numrac FROM b_receta a, b_minuta b, b_minutadet c WHERE b.min_codigo = c.mid_codigo " & _
             "AND c.mid_codrec = a.rec_codigo AND b.min_cencos = '" & cencos & "' AND b.min_codreg = " & codreg & " AND b.min_codser = " & codser & " " & _
             "AND " & sql1 & " = " & anomes & " AND c.mid_tipmin = '" & tipmin & "' ORDER BY a.rec_codigo, b.min_fecmin", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
    codreceta = 0: isalto = 0: auxfecha = 0: cosreceta = 0: canreceta = 0: totgralreceta = 0: condia = 0
    iRow = 0
    Do While Not RS1.EOF
       If RS1!rec_codigo <> codreceta Then
          If isalto = 1 Then
             iRow = iRow + 1
             sql1 = IIf(vg_tipbase = "1", " val(mid(min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),min_fecmin),1,6)) ")
             RS2.Open "SELECT COUNT(b_minutadet.mid_codrec) AS nreg FROM b_minutadet WHERE b_minutadet.mid_codigo IN (SELECT min_codigo FROM b_minuta WHERE min_cencos = '" & cencos & "' AND min_codreg = " & codreg & " AND min_codser = " & codser & " AND " & sql1 & " = " & anomes & ") " & _
                      "AND mid_tipmin = '" & tipmin & "' AND mid_codrec = " & codreceta & "", vg_db, adOpenStatic
             If Not RS2.EOF Then
                .Row = iRow - 1
                .Col = 3
                .CellType = 5
                .TypeHAlign = 1
                .text = Format(RS2!nreg, fg_Pict(6, 0))
                .ForeColor = &HFF0000
             
                .Col = 4
                .CellType = 5
                .TypeHAlign = 1
                .text = Format(CCur(cosreceta / RS2!nreg), fg_Pict(6, 2))
                .ForeColor = &HFF0000
             End If
             RS2.Close: Set RS2 = Nothing
          End If
          .MaxRows = .MaxRows + 1
          iRow = .MaxRows
          .Row = iRow
             
          .Col = 1
          .CellType = 5
          .TypeHAlign = 1
          .text = RS1!rec_codigo
             
          .Col = 2
          .CellType = 5
          .TypeHAlign = 0
          .text = Trim(RS1!rec_nombre)
             
          .Col = 4 + (Dia + Val(Mid(RS1!min_fecmin, 7, 2)))
          .CellType = 5
          .TypeHAlign = 2
          .text = 0
          .text = CCur(.text + IIf(IsNull(RS1!mid_numrac), 0, RS1!mid_numrac)) '"X"
          .ForeColor = &HFF0000
          codreceta = RS1!rec_codigo
          cosreceta = 0: canreceta = 0: auxfecha = 0
          isalto = 1
          ConLinea = ConLinea + 1
       Else
          .Col = 4 + (Dia + Val(Mid(RS1!min_fecmin, 7, 2)))
          .CellType = 5
          .TypeHAlign = 2
          .text = CCur(Val(.text) + IIf(IsNull(RS1!mid_numrac), 0, RS1!mid_numrac)) '"X"
          .ForeColor = &HFF0000
       End If
       cosreceta = (cosreceta + RS1!mid_cosrec + RS1!mid_cosdes)
       If RS1!min_fecmin <> auxfecha Then condia = condia + 1: canreceta = canreceta + 1: auxfecha = RS1!min_fecmin
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    sql1 = IIf(vg_tipbase = "1", " val(mid(min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),min_fecmin),1,6)) ")
    RS1.Open "SELECT COUNT(mid_codrec) AS nreg FROM b_minutadet WHERE mid_codigo IN (SELECT min_codigo FROM b_minuta WHERE min_cencos = '" & cencos & "' AND min_codreg = " & codreg & " AND min_codser = " & codser & " AND " & sql1 & " = " & anomes & ") " & _
             "AND mid_tipmin = '" & tipmin & "' AND mid_codrec = " & codreceta & "", vg_db, adOpenStatic
    If Not RS1.EOF Then
       .Row = iRow
       .Col = 3
       .CellType = 5
       .TypeHAlign = 1
       .text = Format(RS1!nreg, fg_Pict(6, 0))
       .ForeColor = &HFF0000
    
       .Col = 4
       .CellType = 5
       .TypeHAlign = 1
       .text = Format(CCur(cosreceta / RS1!nreg), fg_Pict(6, 2))
       .ForeColor = &HFF0000
    End If
    RS1.Close: Set RS1 = Nothing
    cosreceta = 0: canreceta = 0
    For i = 2 To .MaxRows
        .Row = i
        .Col = 3
        canreceta = Val(.text)
        .Col = 4
        cosreceta = Val(.text)
        totgralreceta = CCur(totgralreceta + (cosreceta * canreceta))
    Next i
    Label1(9).Caption = Format(.MaxRows, fg_Pict(6, 2))
End With

Label1(11).Caption = Format(totgralreceta, fg_Pict(6, 2))

sql1 = IIf(vg_tipbase = "1", " Val(Mid(min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),min_fecmin),1,6)) ")
RS1.Open "SELECT COUNT(min_codigo) AS nreg FROM b_minuta WHERE min_codigo IN (SELECT mid_codigo FROM b_minutadet WHERE mid_tipmin = '" & tipmin & "') AND min_cencos = '" & cencos & "' " & _
        "AND min_codreg = " & codreg & " AND min_codser = " & codser & " AND " & sql1 & " = " & anomes & "", vg_db, adOpenStatic
If Not RS1.EOF And RS1!nreg > 0 Then Label1(11).Caption = Format(CCur(totgralreceta / RS1!nreg), fg_Pict(6, 2))
RS1.Close: Set RS1 = Nothing
fg_descarga
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 2
    If vaSpread1(0).MaxRows < 1 Then Exit Sub
    ExportarExcel
Case 4
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
vaSpread1(0).AllowMultiBlocks = True
vaSpread1(0).SetSelection 1, -1, vaSpread1(0).MaxCols, vaSpread1(0).MaxRows
vaSpread1(0).ClipboardCopy
iRow = vaSpread1(0).MaxRows + 1
'------- Pegar vaspread1(1) - Planilla Excel
NashXl.Range("A1").Select
NashXl.ActiveSheet.Paste
'------- Asignar color
'NashXl.Range("A1:D" & irow).Select
'With NashXl.Selection.Interior
'     .ColorIndex = 36
'     .Pattern = xlSolid
'End With
'------- Colorear titulo
NashXl.Range("A1:AT1").Select
With NashXl.Selection.Interior
     .ColorIndex = 15
     .Pattern = xlSolid
End With
'------- Dibujar marco
NashXl.Range("A1:AT" & iRow).Select
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
NashXl.Range("D3" & ":" & "D" & iRow).Select
NashXl.Selection.NumberFormat = "#,##0.00"
'------- Aplicar totales

'------- Dibujar marco
iRow = iRow + 1
irow2 = iRow + 1
NashXl.Range("B" & iRow).Select
NashXl.ActiveCell.FormulaR1C1 = Label1(8).Caption
NashXl.Range("C" & iRow).Select
NashXl.ActiveCell.FormulaR1C1 = Label1(9).Caption
NashXl.Range("B" & irow2).Select
NashXl.ActiveCell.FormulaR1C1 = Label1(10).Caption
NashXl.Range("C" & irow2).Select
NashXl.ActiveCell.FormulaR1C1 = Label1(11).Caption
NashXl.Range("B" & iRow & ":" & "C" & irow2).Select
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
NashXl.Selection.Font.Bold = True
'With NashXl.Selection.Interior
'     .ColorIndex = 35
'     .Pattern = xlSolid
'End With
NashXl.Range("C" & iRow & ":" & "C" & irow2).Select
NashXl.Selection.NumberFormat = "#,##0.00"
'------- Ajustar columna
NashXl.Cells.Select
NashXl.Cells.EntireColumn.AutoFit
vaSpread1(0).AllowMultiBlocks = False: vaSpread1(0).SetSelection 1, 0, vaSpread1(0).MaxCols, vaSpread1(0).MaxRows
fg_descarga
NashXl.Visible = True
End Sub
