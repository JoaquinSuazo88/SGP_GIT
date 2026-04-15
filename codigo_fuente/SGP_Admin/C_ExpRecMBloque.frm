VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form C_ExpRecMBloque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Width           =   10455
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6015
         Index           =   0
         Left            =   135
         TabIndex        =   14
         Top             =   240
         Width           =   10095
         _Version        =   393216
         _ExtentX        =   17806
         _ExtentY        =   10610
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
         SpreadDesigner  =   "C_ExpRecMBloque.frx":0000
         VisibleCols     =   3
         VisibleRows     =   18
         ScrollBarTrack  =   3
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   6480
         Visible         =   0   'False
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   318
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
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
         Left            =   795
         TabIndex        =   8
         Top             =   645
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ceco"
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
         Left            =   795
         TabIndex        =   7
         Top             =   300
         Width           =   450
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   7575
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Top             =   600
         Width           =   7575
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
         Left            =   795
         TabIndex        =   4
         Top             =   1005
         Width           =   660
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   1800
         TabIndex        =   3
         Top             =   960
         Width           =   7575
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   1800
         TabIndex        =   2
         Top             =   1305
         Width           =   7575
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
         Index           =   3
         Left            =   795
         TabIndex        =   1
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   1845
         TabIndex        =   9
         Top             =   285
         Width           =   7575
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   1845
         TabIndex        =   10
         Top             =   645
         Width           =   7575
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   1845
         TabIndex        =   11
         Top             =   1005
         Width           =   7575
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   7
         Left            =   1830
         TabIndex        =   12
         Top             =   1350
         Width           =   7575
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8685
      Left            =   10590
      TabIndex        =   16
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   15319
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
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   1773
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"C_ExpRecMBloque.frx":0413
   End
End
Attribute VB_Name = "C_ExpRecMBloque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim asubseg As Long, acodreg As Long, acodser As Long, aanomes As Long, numreg As Long
Dim aCeco As String
Dim FechaIni As Long
Dim FechaFin As Long

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Titulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Excel "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "word", , tbrDefault, "word"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Excel "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Titulo

End Sub

Sub LlenarExporReceta(tfor As String, subseg As Long, codReg As Long, codser As Long, anomes As Long)
On Error GoTo Man_Error

Dim isalto As Integer
Dim CodReceta As Long, IRow As Long, i As Long, condia As Long, auxfecha As Long
Dim cosreceta As Double, canreceta As Double, totgralreceta As Double, nomR As String, MetR As String
fg_carga ""
'-------> Rutina frecuencia de recetas
asubseg = subseg
acodreg = codReg
acodser = codser
aanomes = anomes
Me.Caption = tfor
MsgTitulo = tfor
Label1(0).Caption = "Subseg."

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("SELECT sub_codigo, sub_nombre FROM a_subsegmento WHERE sub_codigo=" & subseg & "")
If Not RS1.EOF Then fpayuda(0).Caption = RS1!sub_nombre
RS1.Close
Set RS1 = Nothing

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("SELECT reg_nombre FROM a_regimen WHERE reg_codigo=" & codReg & "")
If Not RS1.EOF Then fpayuda(1).Caption = RS1!reg_nombre
RS1.Close
Set RS1 = Nothing

fpayuda(3).Caption = Meses("01/" & Mid(fg_pone_cero(anomes, 6), 5, 2) & "/" & Mid(fg_pone_cero(anomes, 6), 1, 4)) & " " & Mid(fg_pone_cero(anomes, 6), 1, 4)

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("SELECT ser_codigo, ser_nombre FROM a_servicio WHERE ser_codigo=" & codser & "")
If Not RS1.EOF Then fpayuda(5).Caption = RS1!ser_nombre
RS1.Close
Set RS1 = Nothing

vaSpread1(0).Visible = False
vaSpread1(0).Row = -1: vaSpread1(0).Col = -1
vaSpread1(0).MaxRows = 0
'Clipboard.Clear
'Clipboard.SetText "sp_s_composicionminuta " & subseg & ", " & codreg & ", " & anomes & "01" & ", " & Format(dEoM("01/" & Mid(anomes, 5, 2) & "/" & Mid(anomes, 1, 4)), "yyyymmdd") & ""
'RS1.Open "sp_s_composicionminuta " & subseg & ", " & codreg & ", " & codser & ", " & anomes & "01" & ", " & Format(dEoM("01/" & Mid(anomes, 5, 2) & "/" & Mid(anomes, 1, 4)), "yyyymmdd") & "", vg_db, adOpenForwardOnly ', adOpenStatic
'RS_Dato.Open "sp_s_composicionminutaSS " & subseg & ", " & codreg & ", " & codser & ", " & anomes & "01" & ", " & Format(dEoM("01/" & Mid(anomes, 5, 2) & "/" & Mid(anomes, 1, 4)), "yyyymmdd") & "," & vg_Zona & "", vg_db, adOpenForwardOnly ', adOpenStatic
If RS_Dato.State = 1 Then RS_Dato.Close
RS_Dato.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS_Dato.Open "sgpadm_s_composicionminutas " & subseg & ", " & codReg & ", " & codser & ", " & anomes & "01" & ", " & Format(dEoM("01/" & Mid(anomes, 5, 2) & "/" & Mid(anomes, 1, 4)), "yyyymmdd") & "," & vg_Zona & "", vg_db, adOpenForwardOnly ', adOpenStatic
If RS_Dato.EOF Then RS_Dato.Close: Set RS_Dato = Nothing: Exit Sub
vaSpread1(0).Visible = False
vaSpread1(0).MaxRows = 0
nomR = "": RT1.text = ""
numreg = 0
Do While Not RS_Dato.EOF
    
    If nomR <> Trim(RS_Dato!rec_nomfan) Then
        
        If Trim(RT1.text) <> "" Then
            
            vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
            vaSpread1(0).Row = vaSpread1(0).MaxRows
            vaSpread1(0).Col = 1: vaSpread1(0).TypeHAlign = TypeHAlignLeft: vaSpread1(0).text = RT1.text
        
        ElseIf nomR <> "" Then

'            vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
'            vaSpread1(0).Row = vaSpread1(0).MaxRows
        
        End If
        vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
        vaSpread1(0).Row = vaSpread1(0).MaxRows
        vaSpread1(0).Col = 1
        vaSpread1(0).Font.Bold = True
        vaSpread1(0).Font.Size = 9
        
        vaSpread1(0).TypeHAlign = TypeHAlignLeft
        vaSpread1(0).text = RS_Dato!mid_codrec & " - " & Trim(RS_Dato!rec_nomfan)
        
        nomR = Trim(RS_Dato!rec_nomfan)
    
    End If
    
    RT1.TextRTF = IIf(IsNull(RS_Dato!rec_metpre), "", (RS_Dato!rec_metpre))
    vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
    vaSpread1(0).Row = vaSpread1(0).MaxRows
    vaSpread1(0).Col = 1: vaSpread1(0).TypeHAlign = TypeHAlignLeft: vaSpread1(0).text = Trim(RS_Dato!ing_nombre)
    vaSpread1(0).Col = 2: vaSpread1(0).TypeHAlign = TypeHAlignLeft: vaSpread1(0).text = Trim(RS_Dato!unm_nomcor)
    vaSpread1(0).Col = 3: vaSpread1(0).TypeHAlign = TypeHAlignRight: vaSpread1(0).text = Format(RS_Dato!red_canpro, fg_Pict(6, vg_RDCa))
    RS_Dato.MoveNext: i = i + 1: numreg = numreg + 1
Loop

'RS1.Close: Set RS1 = Nothing

If Trim(RT1.text) <> "" Then
    
    vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
    vaSpread1(0).Row = vaSpread1(0).MaxRows
    vaSpread1(0).Col = 1: vaSpread1(0).TypeHAlign = TypeHAlignLeft: vaSpread1(0).text = RT1.text

End If

vaSpread1(0).Visible = True
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Titulo

End Sub

Sub LlenarExporRecetaBloque(tfor As String, Ceco As String, codReg As Long, codser As Long, fechainib As Long, fechafinb As Long)

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset
Dim isalto As Integer
Dim CodReceta As Long, IRow As Long, i As Long, condia As Long, auxfecha As Long
Dim cosreceta As Double, canreceta As Double, totgralreceta As Double, nomR As String, MetR As String
fg_carga ""
'-------> Rutina frecuencia de recetas
asubseg = 0
aCeco = Ceco
acodreg = codReg
acodser = codser
aanomes = 0
FechaIni = fechainib
FechaFin = fechafinb
Me.Caption = tfor
MsgTitulo = tfor
Label1(0).Caption = "Ceco"

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("SELECT cli_codigo, cli_nombre FROM b_clientes with (nolock) WHERE cli_codigo='" & Ceco & "'")
If Not RS1.EOF Then fpayuda(0).Caption = RS1!Cli_nombre
RS1.Close
Set RS1 = Nothing

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("SELECT reg_nombre FROM a_regimen with (nolock) WHERE reg_codigo=" & codReg & "")
If Not RS1.EOF Then fpayuda(1).Caption = RS1!reg_nombre
RS1.Close
Set RS1 = Nothing
fpayuda(3).Caption = fg_Ctod1(FechaIni) & " - " & fg_Ctod1(FechaFin)

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("SELECT ser_codigo, ser_nombre FROM a_servicio with (nolock) WHERE ser_codigo=" & codser & "")
If Not RS1.EOF Then fpayuda(5).Caption = RS1!ser_nombre
RS1.Close
Set RS1 = Nothing

vaSpread1(0).Visible = False
vaSpread1(0).Row = -1: vaSpread1(0).Col = -1
vaSpread1(0).MaxRows = 0

RS_Dato.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS_Dato = vg_db.Execute("sgpadm_Sel_ExportarRecetaMinutaBloque_V03 '" & Ceco & "', " & codReg & ", " & codser & ", " & FechaIni & ", " & FechaFin & "")
'RS_Dato.Open "sgpadm_Sel_ExportarRecetaMinutaBloque_V02 '" & Ceco & "', " & codReg & ", " & codser & ", " & FechaIni & ", " & FechaFin & "", vg_db, , adOpenStatic 'adOpenForwardOnly
If RS_Dato.EOF Then RS_Dato.Close: Set RS_Dato = Nothing: Exit Sub
vaSpread1(0).Visible = False
vaSpread1(0).MaxRows = 0
'vaSpread1(0).MaxRows = RS_Dato.RecordCount
nomR = "": RT1.text = ""
numreg = 0

Do While Not RS_Dato.EOF
    
    If nomR <> Trim(RS_Dato!rec_nomfan) Then
        
        If Trim(RT1.text) <> "" Then
            
            vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
            vaSpread1(0).Row = vaSpread1(0).MaxRows
            vaSpread1(0).Col = 1
            vaSpread1(0).TypeHAlign = TypeHAlignLeft
            vaSpread1(0).text = RT1.text

'        ElseIf nomR <> "" Then
        End If
        vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
        vaSpread1(0).Row = vaSpread1(0).MaxRows
        
        vaSpread1(0).Col = 1
        vaSpread1(0).Font.Bold = True
        vaSpread1(0).Font.Size = 9
        vaSpread1(0).TypeHAlign = TypeHAlignLeft
        vaSpread1(0).text = RS_Dato!mid_codrec & " - " & Trim(RS_Dato!rec_nomfan)
        
        nomR = Trim(RS_Dato!rec_nomfan)

    End If
    RT1.TextRTF = Trim(RS_Dato!rec_metpre)
    vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
    vaSpread1(0).Row = vaSpread1(0).MaxRows
    
    vaSpread1(0).Col = 1
    vaSpread1(0).TypeHAlign = TypeHAlignLeft
    vaSpread1(0).text = Trim(RS_Dato!ing_nombre)
    
    vaSpread1(0).Col = 2
    vaSpread1(0).TypeHAlign = TypeHAlignLeft
    vaSpread1(0).text = Trim(RS_Dato!unm_nomcor)
    
    vaSpread1(0).Col = 3
    vaSpread1(0).TypeHAlign = TypeHAlignRight
    vaSpread1(0).text = Format(RS_Dato!red_canpro, fg_Pict(6, vg_RDCa))
    
    RS_Dato.MoveNext
    i = i + 1
    numreg = numreg + 1

Loop

If Trim(RT1.text) <> "" Then
    
    vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
    vaSpread1(0).Row = vaSpread1(0).MaxRows
    vaSpread1(0).Col = 1
    vaSpread1(0).TypeHAlign = TypeHAlignLeft
    vaSpread1(0).text = RT1.text

End If
vaSpread1(0).Visible = True
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Titulo
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Man_Error

If RS_Dato.State = 1 Then
   
   RS_Dato.Close
   Set RS_Dato = Nothing

End If

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Titulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error

Select Case Button.Index

Case 2, 4
    
    If vaSpread1(0).MaxRows < 1 Then Exit Sub
    
    If Button.Index = 2 Then
       
       ExportarExcel
    
    Else
       
       If aanomes > 0 Then
          
          I_ListaRecetaPlanificacion asubseg, acodreg, acodser, aanomes & "01", Format(dEoM("01/" & Mid(aanomes, 5, 2) & "/" & Mid(aanomes, 1, 4)), "yyyymmdd") & "", numreg
       
       Else
          
          I_ListaRecetaPlanificacionBloque aCeco, acodreg, acodser, FechaIni, FechaFin, numreg
       
       End If
    
    End If

Case 6
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Titulo
End Sub

Sub ExportarExcel()

On Error GoTo Man_Error

Dim NashXl As excel.Application
Dim IRow As Long, irow2 As Long
fg_carga ""
Set NashXl = CreateObject("excel.application")
Set NashXl = New excel.Application
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
vaSpread1(0).SetSelection 1, -1, vaSpread1(0).maxcols, vaSpread1(0).MaxRows + 3
vaSpread1(0).ClipboardCopy
IRow = vaSpread1(0).MaxRows + 5
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
NashXl.Range("A5:C" & IRow).Select
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
NashXl.Range("C2" & ":" & "C" & IRow).Select
NashXl.Selection.NumberFormat = "#,##0." & fg_pone_cero(0, vg_RDCa) '"#,##0.00"
'------- Aplicar totales

NashXl.Selection.Font.Bold = True
'With NashXl.Selection.Interior
'     .ColorIndex = 35
'     .Pattern = xlSolid
'End With
NashXl.Range("B" & IRow & ":" & "B" & 2).Select
NashXl.Selection.NumberFormat = "#,##0." & fg_pone_cero(0, vg_RDCa) '"#,##0.00"
'------- Ajustar columna
NashXl.Cells.Select
NashXl.Cells.EntireColumn.AutoFit
vaSpread1(0).AllowMultiBlocks = False: vaSpread1(0).SetSelection 1, 0, vaSpread1(0).maxcols, vaSpread1(0).MaxRows
fg_descarga
NashXl.Visible = True

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Titulo

End Sub
