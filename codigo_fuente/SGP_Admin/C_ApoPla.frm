VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form C_ApoPla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aporte Planificación Día"
   ClientHeight    =   5340
   ClientLeft      =   525
   ClientTop       =   1230
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
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
         Caption         =   "Sub-Segmento"
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
         Left            =   1395
         TabIndex        =   5
         Top             =   300
         Width           =   1245
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
         Left            =   1395
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
         Left            =   1395
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
      Width           =   9255
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9015
         _Version        =   393216
         _ExtentX        =   15901
         _ExtentY        =   5953
         _StockProps     =   64
         ColsFrozen      =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         MaxRows         =   18
         SpreadDesigner  =   "C_ApoPla.frx":0000
         VisibleCols     =   5
         VisibleRows     =   18
         ScrollBarTrack  =   3
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5340
      Left            =   9405
      TabIndex        =   6
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   9419
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
Option Explicit
Option Compare Text

Private RS2     As New ADODB.Recordset
Private BtnX    As Variant

Private Sub Form_Activate()

On Error GoTo Man_Error

    fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Form_Load()
    
On Error GoTo Man_Error

    fg_centra Me
    Toolbar1.ImageList = Partida.IL1
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Excel "
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Sub LlenarApoPlan(ffor As Object, tfor As String, subseg As Variant, codReg As Long, codser As Long, Fecha As Long, TipMin As String, icol As Long)

On Error GoTo Man_Error

Dim RS1         As New ADODB.Recordset
Dim CodRec      As Long
Dim IndApo      As Long
Dim IRow        As Long
Dim i           As Long
Dim X           As Long
Dim tiprec      As Long
Dim nomrec      As String
Dim StrRec      As String
Dim StrRecb     As String
Dim TotBru      As Double
Dim TotSer      As Double
Dim TotNet      As Double
Dim TotNetApr   As Double
Dim vecapo()    As Long
Dim vectot()    As Double

    fg_carga ""
    Me.Caption = tfor
    MsgTitulo = tfor
    'Llenar tabla nutrientes
    IndApo = 1
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("sgpadm_s_nutriente 1, " & vg_codservicio & ",''")
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fg_descarga: MsgBox "No existe maestro nutrientes", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    vaSpread1(0).MaxRows = 0
    vaSpread1(0).maxcols = 6
    vaSpread1(0).Row = 0
    ReDim Preserve vecapo(0): ReDim Preserve vectot(0)
    TotBru = 0: TotSer = 0: TotNet = 0
    
    Do While Not RS1.EOF
       
       vaSpread1(0).maxcols = vaSpread1(0).maxcols + 1
       vaSpread1(0).Col = vaSpread1(0).maxcols: vaSpread1(0).text = Trim(RS1!nut_nombre)
       ReDim Preserve vecapo(IndApo)
       vecapo(IndApo) = RS1!nut_codigo
       ReDim Preserve vectot(IndApo)
       vectot(IndApo) = 0
       IndApo = IndApo + 1
       RS1.MoveNext
    
    Loop
    RS1.Close
    Set RS1 = Nothing
    
    If VarSitioRemoto = False Then
        
        Let Label1(0).Caption = "Sub-Segmento"
        IndApo = IndApo - 1
        If IndApo < 1 Then IndApo = 1
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS1 = vg_db.Execute("SELECT sub_codigo, sub_nombre FROM a_subsegmento WHERE sub_codigo = " & subseg & "")
        If Not RS1.EOF Then fpayuda(0).Caption = RS1!sub_nombre
        RS1.Close: Set RS1 = Nothing
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS1 = vg_db.Execute("SELECT reg_nombre FROM a_regimen WHERE reg_codigo = " & codReg & "")
        If Not RS1.EOF Then fpayuda(1).Caption = RS1!reg_nombre
        RS1.Close: Set RS1 = Nothing
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS1 = vg_db.Execute("SELECT ser_nombre FROM a_servicio WHERE ser_codigo = " & codser & "")
        If Not RS1.EOF Then fpayuda(3).Caption = RS1!ser_nombre
        RS1.Close: Set RS1 = Nothing
    
    Else
        
        Let Label1(0).Caption = "Cliente"
        IndApo = IndApo - 1
        If IndApo < 1 Then IndApo = 1
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS1 = vg_db.Execute("SELECT cli_codigo, cli_nombre FROM b_clientes wiht (nolock) WHERE cli_codigo = '" & subseg & "' and cli_tipo = 0")
        If Not RS1.EOF Then fpayuda(0).Caption = Trim(RS1!Cli_nombre)
        RS1.Close: Set RS1 = Nothing
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS1 = vg_db.Execute("SELECT reg_nombre FROM a_regimen with (nolock) WHERE reg_codigo = " & codReg & "")
        If Not RS1.EOF Then fpayuda(1).Caption = Trim(RS1!reg_nombre)
        RS1.Close: Set RS1 = Nothing
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS1 = vg_db.Execute("SELECT ser_nombre FROM a_servicio with (nolock) WHERE ser_codigo = " & codser & "")
        If Not RS1.EOF Then fpayuda(3).Caption = Trim(RS1!ser_nombre)
        RS1.Close: Set RS1 = Nothing
        
    End If

    '-------> Formatear colores
    vaSpread1(0).Row = -1
    vaSpread1(0).Col = -1
    vaSpread1(0).BackColor = &HE0FEFE
    vaSpread1(0).Row = -1
    vaSpread1(0).Col = 1
    vaSpread1(0).BackColor = &HDEFEDE
    vaSpread1(0).Row = -1
    vaSpread1(0).Col = 2
    vaSpread1(0).BackColor = &HDEFEDE
    vaSpread1(0).Row = -1
    vaSpread1(0).Col = 3
    vaSpread1(0).BackColor = &HDEFEDE
    vaSpread1(0).Row = -1
    vaSpread1(0).Col = 4
    vaSpread1(0).BackColor = &HDEFEDE
    vaSpread1(0).Row = -1
    vaSpread1(0).Col = 5
    vaSpread1(0).BackColor = &HDEFEDE
    vaSpread1(0).Row = -1
    vaSpread1(0).Col = 6
    vaSpread1(0).BackColor = &HDEFEDE
    
    
    For i = 1 To (ffor.vaSpread1.MaxRows - 1)
        ffor.vaSpread1.Row = i
        ffor.vaSpread1.Col = icol + IIf(VarSitioRemoto, 4, 3)
        If Trim(ffor.vaSpread1.text) <> "" Then
            StrRec = ffor.vaSpread1.text
            If Len(StrRec) <> 0 Then
                
                Do While InStr(StrRec, ";") <> 0
                    
                    StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
                    StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
                    CodRec = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)): StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
                    tiprec = Val(Mid(StrRecb, 1))
                
                Loop
            
            End If
            ffor.vaSpread1.Col = icol
            nomrec = ffor.vaSpread1.text
            If vg_Zona = "" Then
                
                Let vg_Zona = 0
            
            End If
            
            If RS1.State = 1 Then RS1.Close
            RS1.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            If VarSitioRemoto = False Then
                
                Set RS1 = vg_db.Execute("sgpadm_s_aporteNutriMinuta " & IIf(CodRec = 0, BuscarCodReceta(nomrec), CodRec) & "," & subseg & "," & codReg & ", " & vg_Zona & "")
            
            Else
                
                Set RS1 = vg_db.Execute("sgpadm_Sel_AporteNutriMinutaBloque_V02 1, " & IIf(CodRec = 0, BuscarCodRecetaSitRem(nomrec), CodRec) & ", " & tiprec & ", '" & subseg & "'")  '----------------" & subseg & ",   " & vg_Zona & "
            
            End If
            
            If Not RS1.EOF Then
                
                vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
                vaSpread1(0).Row = vaSpread1(0).MaxRows
                
                If VarSitioRemoto = False Then
                    
                    vaSpread1(0).Col = 1
                    vaSpread1(0).CellType = 5
                    vaSpread1(0).TypeHAlign = 1
                    vaSpread1(0).text = IIf(CodRec = 0, BuscarCodReceta(nomrec), CodRec)
                
                Else
                    
                    Let vaSpread1(0).Col = 1
                    Let vaSpread1(0).CellType = 5
                    Let vaSpread1(0).TypeHAlign = 1
                    Let vaSpread1(0).text = IIf(CodRec = 0, BuscarCodRecetaSitRem(nomrec), CodRec)
                
                End If
                
                vaSpread1(0).Col = 2
                vaSpread1(0).CellType = 5
                vaSpread1(0).TypeHAlign = 0
                vaSpread1(0).text = nomrec
                
                vaSpread1(0).Col = 3
                vaSpread1(0).CellType = 5
                vaSpread1(0).TypeHAlign = 1
                vaSpread1(0).text = Format(RS1!canpro, fg_Pict(6, vg_RDCa))
                
                vaSpread1(0).Col = 4
                vaSpread1(0).CellType = 5
                vaSpread1(0).TypeHAlign = 1
                vaSpread1(0).text = Format(RS1!cannetapro, fg_Pict(6, vg_RDCa))
                
                
                vaSpread1(0).Col = 5
                vaSpread1(0).CellType = 5
                vaSpread1(0).TypeHAlign = 1
                vaSpread1(0).text = Format(RS1!canser, fg_Pict(6, vg_RDCa))
                
                vaSpread1(0).Col = 6
                vaSpread1(0).CellType = 5
                vaSpread1(0).TypeHAlign = 1
                vaSpread1(0).text = Format(RS1!cannet, fg_Pict(6, vg_RDCa))
                
                TotBru = CCur(TotBru + RS1!canpro)
                TotSer = CCur(TotSer + RS1!canser)
                TotNet = CCur(TotNet + RS1!cannet)
                TotNetApr = CCur(TotNetApr + RS1!cannetapro)
                
                For X = 1 To IndApo
                    
                    vaSpread1(0).Col = X + 6
                    vaSpread1(0).CellType = 5
                    vaSpread1(0).TypeHAlign = 1
                    vaSpread1(0).text = Format(0, fg_Pict(6, 2))
                
                Next X
                
            End If
            RS1.Close
            Set RS1 = Nothing
           
           '-------> Calculo aporte nutricionales
           If RS1.State = 1 Then RS1.Close
           RS1.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           If VarSitioRemoto = False Then
                
                Set RS1 = vg_db.Execute("sgpadm_s_AporteNutricionales 1," & IIf(CodRec = 0, BuscarCodReceta(nomrec), CodRec) & "," & subseg & "," & codReg & ", " & vg_Zona & "")
            
            Else
                
                Set RS1 = vg_db.Execute("sgpadm_Sel_AporteNutriMinutaBloque_V02 2, " & IIf(CodRec = 0, BuscarCodRecetaSitRem(nomrec), CodRec) & "," & tiprec & ", '" & subseg & "'")  '------& "," & subseg   & ", " & vg_Zona & "")
            
            End If
            
            If Not RS1.EOF Then
                Do While Not RS1.EOF
                    
                    For X = 1 To IndApo
                        
                        If RS1!nut_codigo = vecapo(X) Then
                            
                            vaSpread1(0).Col = X + 6
                            vaSpread1(0).CellType = 5
                            vaSpread1(0).TypeHAlign = 1
                            vaSpread1(0).text = Format(RS1!candiet, fg_Pict(6, 2))
                            vectot(X) = CCur(vectot(X) + RS1!candiet)
                            Exit For
                        
                        End If
                    
                    Next X
                    
                    RS1.MoveNext
                
                Loop
            
            End If
            RS1.Close
            Set RS1 = Nothing
        
        End If
    
    Next i

    If vaSpread1(0).MaxRows < 1 Then Exit Sub
    vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 2
    vaSpread1(0).Row = vaSpread1(0).MaxRows
    vaSpread1(0).Col = 1
    vaSpread1(0).CellType = 5
    vaSpread1(0).TypeHAlign = 0
    vaSpread1(0).Font.Bold = True
    vaSpread1(0).Font.Size = 8
    vaSpread1(0).BackColor = &HC0C0C0
    vaSpread1(0).text = ""
    
    vaSpread1(0).Col = 2
    vaSpread1(0).CellType = 5
    vaSpread1(0).TypeHAlign = 0
    vaSpread1(0).Font.Bold = True
    vaSpread1(0).Font.Size = 8
    vaSpread1(0).BackColor = &HC0C0C0
    vaSpread1(0).text = "Total Gral. "
    
    vaSpread1(0).Col = 3
    vaSpread1(0).CellType = 5
    vaSpread1(0).TypeHAlign = 1
    vaSpread1(0).Font.Bold = True
    vaSpread1(0).Font.Size = 8
    vaSpread1(0).BackColor = &HC0C0C0
    vaSpread1(0).text = Format(TotBru, fg_Pict(6, vg_RDCa))
    
    vaSpread1(0).Col = 4
    vaSpread1(0).CellType = 5
    vaSpread1(0).TypeHAlign = 1
    vaSpread1(0).Font.Bold = True
    vaSpread1(0).Font.Size = 8
    vaSpread1(0).BackColor = &HC0C0C0
    vaSpread1(0).text = Format(TotNetApr, fg_Pict(6, vg_RDCa))
    
    vaSpread1(0).Col = 5
    vaSpread1(0).CellType = 5
    vaSpread1(0).TypeHAlign = 1
    vaSpread1(0).Font.Bold = True
    vaSpread1(0).Font.Size = 8
    vaSpread1(0).BackColor = &HC0C0C0
    vaSpread1(0).text = Format(TotSer, fg_Pict(6, vg_RDCa))
    
    vaSpread1(0).Col = 6
    vaSpread1(0).CellType = 5
    vaSpread1(0).TypeHAlign = 1
    vaSpread1(0).Font.Bold = True
    vaSpread1(0).Font.Size = 8
    vaSpread1(0).BackColor = &HC0C0C0
    vaSpread1(0).text = Format(TotNet, fg_Pict(6, vg_RDCa))
    
    For X = 1 To IndApo
        
        vaSpread1(0).Col = X + 6
        vaSpread1(0).CellType = 5
        vaSpread1(0).TypeHAlign = 1
        vaSpread1(0).Font.Bold = True: vaSpread1(0).Font.Size = 8
        vaSpread1(0).BackColor = &HC0C0C0:
        vaSpread1(0).text = Format(vectot(X), fg_Pict(6, 2))
    
    Next X
    fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
On Error GoTo Man_Error

    Select Case Button.Index
        
        Case 2
            
            If vaSpread1(0).MaxRows < 1 Then Exit Sub
            Call ExportarExcel
        
        Case 4
            
            Me.Hide
            Call Unload(Me)
    
    End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Sub ExportarExcel()

On Error GoTo Man_Error

Dim NashXl  As excel.Application
Dim IRow    As Long
Dim irow2   As Long

    fg_carga ""
    Set NashXl = CreateObject("excel.application")
    Set NashXl = New excel.Application
    NashXl.SheetsInNewWorkbook = 1
    NashXl.Workbooks.Add
    vaSpread1(0).AllowMultiBlocks = True
    vaSpread1(0).SetSelection 1, -1, vaSpread1(0).maxcols, vaSpread1(0).MaxRows
    vaSpread1(0).ClipboardCopy
    IRow = vaSpread1(0).MaxRows + 1
    '-------> Pegar vaspread1(1) - Planilla Excel
    NashXl.Range("A1").Select
    NashXl.ActiveSheet.Paste
    '-------> Colorear titulo
    NashXl.Range("A1:AT1").Select
    With NashXl.Selection.Interior
         .ColorIndex = 15
         .Pattern = xlSolid
    End With
    '-------> Dibujar marco
    NashXl.Range("A1:AT" & IRow).Select
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
    NashXl.Range("D3" & ":" & "D" & IRow).Select
    NashXl.Selection.NumberFormat = "#,##0." & fg_pone_cero(0, vg_RDCa) '"#,##0.00"
    'Dibujar marco
    IRow = IRow + 1
    irow2 = IRow + 1
    'Ajustar columna
    NashXl.Cells.Select
    NashXl.Cells.EntireColumn.AutoFit
    vaSpread1(0).AllowMultiBlocks = False: vaSpread1(0).SetSelection 1, 0, vaSpread1(0).maxcols, vaSpread1(0).MaxRows
    fg_descarga
    NashXl.Visible = True

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub
