VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form C_ApoPla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aporte Planificación Día"
   ClientHeight    =   5190
   ClientLeft      =   525
   ClientTop       =   1230
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   12195
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   9255
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
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2760
         TabIndex        =   9
         Top             =   600
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
         Left            =   1875
         TabIndex        =   5
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
         Left            =   1875
         TabIndex        =   4
         Top             =   645
         Width           =   750
      End
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
         Left            =   1875
         TabIndex        =   3
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2760
         TabIndex        =   7
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   2805
         TabIndex        =   8
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
         TabIndex        =   10
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
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   11415
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11175
         _Version        =   393216
         _ExtentX        =   19711
         _ExtentY        =   5953
         _StockProps     =   64
         ColsFrozen      =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         MaxRows         =   18
         SpreadDesigner  =   "C_ApoPla.frx":0000
         VisibleCols     =   4
         VisibleRows     =   18
         ScrollBarTrack  =   3
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5190
      Left            =   11565
      TabIndex        =   6
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   9155
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "C_ApoPla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset

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

Sub LlenarApoPlan(ffor As Object, tfor As String, cencos As String, codreg As Long, codser As Long, Fecha As Long, tipmin As String, iCol As Long)
Dim CodRec As Long, indapo As Long, iRow As Long, i As Long, X As Long, tiprec As Long
Dim nomrec As String, StrRec As String, StrRecb As String
Dim totbru As Double, totser As Double, totnet As Double
Dim vecapo() As Long
Dim vectot() As Double
fg_carga ""
Me.Caption = tfor
Msgtitulo = tfor
'------- Llenar tabla nutrientes
indapo = 1
RS1.Open RutinaLectura.Nutriente(1, 0, ""), vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fg_descarga: MsgBox "No existe maestro nutrientes", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
With vaSpread1(0)
    .MaxRows = 0: .MaxCols = 5: .Row = 0: ReDim Preserve vecapo(0): ReDim Preserve vectot(0)
    totbru = 0: totser = 0: totnet = 0
    Do While Not RS1.EOF
       .MaxCols = .MaxCols + 1
       .Col = .MaxCols: .text = Trim(RS1!nut_nombre)
       ReDim Preserve vecapo(indapo)
       vecapo(indapo) = RS1!nut_codigo
       ReDim Preserve vectot(indapo)
       vectot(indapo) = 0
       indapo = indapo + 1
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    indapo = indapo - 1
    If indapo < 1 Then indapo = 1
    RS1.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(cencos)), ""), vg_db, adOpenStatic
    If Not RS1.EOF Then fpayuda(0).Caption = RS1!cli_nombre
    RS1.Close: Set RS1 = Nothing
    RS1.Open RutinaLectura.Regimen(2, codreg, ""), vg_db, adOpenStatic
    If Not RS1.EOF Then fpayuda(1).Caption = RS1!reg_nombre
    RS1.Close: Set RS1 = Nothing
    RS1.Open RutinaLectura.Servicio(8, codser, ""), vg_db, adOpenStatic
    If Not RS1.EOF Then fpayuda(3).Caption = RS1!ser_nombre
    RS1.Close: Set RS1 = Nothing
    
    '------- Formatear colores
    .Row = -1: .Col = -1: .BackColor = &HC0FFFF
    .Row = -1: .Col = 1: .BackColor = &HC0FFC0
    .Row = -1: .Col = 2: .BackColor = &HC0FFC0
    .Row = -1: .Col = 3: .BackColor = &HC0FFC0
    .Row = -1: .Col = 4: .BackColor = &HC0FFC0
    .Row = -1: .Col = 5: .BackColor = &HC0FFC0
    For i = 1 To (ffor.vaSpread1.MaxRows - 1)
        ffor.vaSpread1.Row = i
        ffor.vaSpread1.Col = iCol + 3
        If Trim(ffor.vaSpread1.text) <> "" Then
    '       codrec = ffor.vaSpread1.Text
           StrRec = ffor.vaSpread1.text
           If Len(StrRec) <> 0 Then
              Do While InStr(StrRec, ";") <> 0
                 StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
                 StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
                 CodRec = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)): StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
                 tiprec = Val(Mid(StrRecb, 1))
              Loop
           End If
           ffor.vaSpread1.Col = iCol: nomrec = ffor.vaSpread1.text
           
           '------- Calculo gramos brutos, cantidad servida y cantidad neta
           RS1.Open RutinaLectura.Receta(5, CodRec, 0, 0, "", tiprec), vg_db, adOpenStatic
           If Not RS1.EOF Then
              .MaxRows = .MaxRows + 1
              .Row = .MaxRows
              .Col = 1: .CellType = 5: .TypeHAlign = 1: .text = CodRec
              .Col = 2: .CellType = 5: .TypeHAlign = 0: .text = nomrec
              .Col = 3: .CellType = 5: .TypeHAlign = 1: .text = Format(RS1!canpro, fg_Pict(6, vg_RDCa))
              .Col = 4: .CellType = 5: .TypeHAlign = 1: .text = Format(RS1!canser, fg_Pict(6, vg_RDCa))
              .Col = 5: .CellType = 5: .TypeHAlign = 1: .text = Format(RS1!cannet, fg_Pict(6, vg_RDCa))
              totbru = CCur(totbru + RS1!canpro)
              totser = CCur(totser + RS1!canser)
              totnet = CCur(totnet + RS1!cannet)
              For X = 1 To indapo
                  .Col = X + 5
                  .CellType = 5
                  .TypeHAlign = 1
                  .text = Format(0, fg_Pict(6, 2))
              Next X
           End If
           RS1.Close: Set RS1 = Nothing
           
           '------- Calculo aporte nutricionales
           RS1.Open RutinaLectura.Receta(4, CodRec, 0, 0, "", tiprec), vg_db, adOpenStatic
           If Not RS1.EOF Then
              Do While Not RS1.EOF
                 For X = 1 To indapo
                    If RS1!nut_codigo = vecapo(X) Then
                       .Col = X + 5
                       .CellType = 5
                       .TypeHAlign = 1
                       .text = Format(RS1!candiet, fg_Pict(6, 2))
                       vectot(X) = CCur(vectot(X) + RS1!candiet)
                       Exit For
                    End If
                 Next X
                 RS1.MoveNext
              Loop
           End If
           RS1.Close: Set RS1 = Nothing
        End If
    Next i
    If .MaxRows < 1 Then Exit Sub
    .MaxRows = .MaxRows + 2
    .Row = .MaxRows
    .Col = 1: .CellType = 5: .TypeHAlign = 0: .Font.Bold = True: .Font.Size = 8: .BackColor = &HC0C0C0: .text = ""
    .Col = 2: .CellType = 5: .TypeHAlign = 0: .Font.Bold = True: .Font.Size = 8: .BackColor = &HC0C0C0: .text = "Total Gral. "
    .Col = 3: .CellType = 5: .TypeHAlign = 1:  .Font.Bold = True: .Font.Size = 8: .BackColor = &HC0C0C0: .text = Format(totbru, fg_Pict(6, vg_RDCa))
    .Col = 4: .CellType = 5: .TypeHAlign = 1:  .Font.Bold = True: .Font.Size = 8: .BackColor = &HC0C0C0: .text = Format(totser, fg_Pict(6, vg_RDCa))
    .Col = 5: .CellType = 5: .TypeHAlign = 1:  .Font.Bold = True: .Font.Size = 8: .BackColor = &HC0C0C0: .text = Format(totnet, fg_Pict(6, vg_RDCa))
    For X = 1 To indapo
        .Col = X + 5
        .CellType = 5
        .TypeHAlign = 1
        .Font.Bold = True: .Font.Size = 8
        .BackColor = &HC0C0C0:
        .text = Format(vectot(X), fg_Pict(6, 2))
    Next X
End With
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
NashXl.Range("C2" & ":" & "C" & iRow).Select
NashXl.Selection.NumberFormat = "#,##0." & fg_pone_cero(0, vg_RDCa)
NashXl.Range("D2" & ":" & "D" & iRow).Select
NashXl.Selection.NumberFormat = "#,##0." & fg_pone_cero(0, vg_RDCa)
NashXl.Range("E2" & ":" & "E" & iRow).Select
NashXl.Selection.NumberFormat = "#,##0." & fg_pone_cero(0, vg_RDCa)
'------- Dibujar marco
iRow = iRow + 1
irow2 = iRow + 1
'------- Ajustar columna
NashXl.Cells.Select
NashXl.Cells.EntireColumn.AutoFit
vaSpread1(0).AllowMultiBlocks = False: vaSpread1(0).SetSelection 1, 0, vaSpread1(0).MaxCols, vaSpread1(0).MaxRows
fg_descarga
NashXl.Visible = True
End Sub
