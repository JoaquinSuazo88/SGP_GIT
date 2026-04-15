VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form E_TrabajosPorLotes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial Trabajo por Lotes"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   14025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   13695
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   5040
         TabIndex        =   6
         Top             =   5670
         Width           =   1620
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   7
            Top             =   135
            Width           =   1515
         End
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Left            =   11835
         TabIndex        =   4
         Top             =   5670
         Width           =   1350
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   5
            Top             =   135
            Width           =   1245
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4695
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   13005
         _Version        =   393216
         _ExtentX        =   22939
         _ExtentY        =   8281
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
         MaxCols         =   7
         SpreadDesigner  =   "E_TrabajosPorLotes.frx":0000
         VisibleCols     =   3
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   4815
         TabIndex        =   8
         Top             =   330
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ButtonStyle     =   2
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
         Text            =   "01/09/2013"
         DateCalcMethod  =   4
         DateTimeFormat  =   5
         UserDefinedFormat=   "dd/mm/yyyy"
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
      Begin EditLib.fpDateTime FpFecHasta 
         Height          =   315
         Left            =   8580
         TabIndex        =   9
         Top             =   330
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ButtonStyle     =   2
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
         Text            =   "28/09/2013"
         DateCalcMethod  =   4
         DateTimeFormat  =   5
         UserDefinedFormat=   "dd/mm/yyyy"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Un Momento Procesando Información ..."
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
         Left            =   360
         TabIndex        =   12
         Top             =   5880
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha hasta"
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
         Left            =   7275
         TabIndex        =   11
         Top             =   405
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha desde"
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
         Left            =   3480
         TabIndex        =   10
         Top             =   405
         Width           =   1110
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   11115
      TabIndex        =   1
      Top             =   6720
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   12525
      TabIndex        =   0
      Top             =   6720
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "E_TrabajosPorLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public lc_Aux As String
Dim MsgTitulo As String

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim Sql             As String
Dim Sql1            As String
Dim codCeco         As String
Dim NomArchivoExcel As String
Dim Extension       As String
Dim seleccion       As Integer
Dim i               As Long
Dim X               As Long

Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object

If Not ValidarDatos Then Exit Sub

'--> Concatenar codigo ceco
'codCeco = ""
'X = 0
'For i = 1 To vaSpread1.MaxRows
'
'    vaSpread1.Row = i
'    vaSpread1.Col = 1 'Seleccion
'    seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
'
'    If seleccion = 1 And vaSpread1.RowHidden = False Then
'
'       vaSpread1.Col = 2
'       codCeco = codCeco & "'" & vaSpread1.text & "', "
'
'       X = X + 1
'
'    End If
'
'Next i
'
'
'If Trim(codCeco) = "" Then
'
'   MsgBox "No existen datos seleccionados grilla cecos...", vbExclamation + vbOKOnly, MsgTitulo
'   Exit Sub
'
'End If


'procedimiento para los directores sgpadm_Sel_ResumidoQDirectores

Label1(0).Visible = True
Label1(0).Caption = "Un Momento Procesando Información ......"
Sql = ""
Sql = "SGP_S_CargaGrillaHisorial "
If X <> vaSpread1.MaxRows Then
   
   Sql = Sql & "'" & Format(FpFecDesde.text, "yyyymmdd") & "', '" & Format(FpFecHasta.text, "yyyymmdd") & "'"

End If

RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

If Not RS.EOF Then
  
   If RS.RecordCount > 1020000 Then
      
      RS.Close
      Set RS = Nothing
      MsgBox "El resultado sobrepasa el máximo de filas en excel, deberá seleccionar menos Cecos", vbCritical
      Label1(0).Visible = False
      Frame1.Enabled = True
      Exit Sub
   
   End If
  
End If
  
'-------> Guardar nombre archivo excel
NomArchivoExcel = ""
CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
CD.Filter = "Todos los archivos *.xlsx,*.xls"
On Error Resume Next
CD.ShowSave
           
'-------> JPAZ Permite controlar Boton Cancelar
If Err.Number = 32755 Then
   MsgBox "Proceso cancelado"
   Label1(0).Visible = False
   Frame1.Enabled = True
   Exit Sub
End If
            
If CD.FileName = "" Then
   
   MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
   Label1(0).Visible = False
   Frame1.Enabled = True
   Exit Sub

Else
   
   Extension = ""
   Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
   
   If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
      
      MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
      Label1(0).Visible = False
      Frame1.Enabled = True
      Exit Sub
   
   End If
   NomArchivoExcel = CD.FileName

End If
          
FpFecDesde.Enabled = False
FpFecHasta.Enabled = False
Frame1.Enabled = False

fg_carga ""
  
'Sql = ""
'Sql = Sql & Replace(Mid(codceco, 1, Len(codceco) - 2), "'", """")

'Set RS = vg_db.Execute("sgpadm_Sel_DetalleMinuta_2 '" & Sql & "', '" & Format(FpFecDesde, ("YYYYmmdd")) & "', '" & Format(FpFecHasta, ("YYYYmmdd")) & "'")
  
'-------> Create an instance of Excel and add a workbook
Set xlApp = CreateObject("Excel.Application")
Set xlWb = xlApp.Workbooks.Add
Set xlWs = xlWb.Worksheets("Hoja1")
  
'-------> Display Excel and give user control of Excel's lifetime
xlApp.UserControl = True
    
'-------> Check version of Excel
Call encabezado(RS, xlWs)
          
xlWs.Cells(2, 1).CopyFromRecordset RS

'-------> Auto-fit the column widths and row heights
xlApp.Selection.CurrentRegion.Columns.AutoFit
xlApp.Selection.CurrentRegion.Rows.AutoFit
    
'xlApp.Columns("A:A").Select
'xlApp.Selection.Delete Shift:=xlToLeft
  
xlWb.Close True, NomArchivoExcel

Dim XL As New Excel.Application 'Crea el objeto excel
XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
XL.Visible = True
XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
'-------> Close ADO objects
RS.Close
Set RS = Nothing
    
'-- Cerrar Excel
xlApp.Quit
'-------> Release Excel references
Set xlWs = Nothing
Set xlWb = Nothing
Set xlApp = Nothing
  
fg_descarga
Label1(0).Visible = False
MsgBox "Proceso Finalizado", vbInformation, Me.Caption
                
FpFecDesde.Enabled = True
FpFecHasta.Enabled = True
Frame1.Enabled = True


Exit Sub
Man_Error:
    Frame1.Enabled = True
    Label1(0).Visible = False
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    Resume
End Sub

Private Sub Command2_Click()

On Error GoTo Man_Error

Unload Me

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
fg_centra Me
MsgTitulo = "Historial Trabajos por Lotes"

FpFecHasta.text = Format(Date, "dd/mm/yyyy")
FpFecDesde.text = Format(Date, "dd/mm/yyyy")
vaSpread1.MaxRows = 0

TextDet1(2).text = ""
TextDet1(3).text = ""

CargaGrillaHistorial

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub



Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub

CargaGrillaHistorial

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

If IsDate(FpFecHasta.text) = False Then Exit Sub

CargaGrillaHistorial

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
End Sub

Private Sub TextDet1_Change(Index As Integer)
On Error GoTo Man_Error

Dim i          As Long
Dim indactivo  As Integer

If Index = 2 Then
   
   TextDet1(3).text = ""

ElseIf Index = 3 Then
   
   TextDet1(2).text = ""

End If

Select Case Index

Case 2, 3
    vaSpread1.Visible = False
    If Trim(TextDet1(Index).text) <> "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = IIf(Index = 2, Index + 1, Index + 4)
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(TextDet1(Index).text) & "*"
           vaSpread1.Col = 2
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           Else
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           End If
        Next i
        vaSpread1.SetActiveCell Index + 1, 1
    End If
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Trim(TextDet1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(TextDet1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
    If Trim(TextDet1(Index).text) = "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Next
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextDet1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1
    Dim i As Long
    vaSpread1.Col = 1

    For i = BlockRow To BlockRow2

        vaSpread1.Row = i

        If vaSpread1.RowHidden = False Then

           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")

        End If

    Next
    
Case Is <> 1

    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    
        End If
        
    Next
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Function ValidarDatos() As Boolean

Dim seleccion As Integer
Dim i As Long

ValidarDatos = True

'-------> Validar fechas
If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
   
   MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If
    
If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
   
   MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

'-------> Validar que exista un dato seleccionado
seleccion = 0
For i = 1 To vaSpread1.MaxRows
       
    vaSpread1.Row = i
    vaSpread1.Col = 1 'Seleccion
    seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
    If seleccion = 1 And vaSpread1.RowHidden = False Then
       Exit For
    End If
  
Next i
  
If seleccion = 0 Then
     
   MsgBox " Se debe seleccionar un ceco por lo menos", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function
  
End If

End Function

Sub encabezado(ByRef RS As ADODB.Recordset, ByRef xlWs As Object)
On Error GoTo Man_Error

Dim fldCount As Long
Dim icol As Long

'-------> Copy field names to the first row of the worksheet
fldCount = RS.Fields.count
For icol = 1 To fldCount
    xlWs.Cells(1, icol).Value = RS.Fields(icol - 1).Name
Next

Exit Sub
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo
End Sub






Private Sub CargaGrillaHistorial()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i As Long

vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &HC0FFFF

vaSpread1.MaxRows = 0
Set RS = vg_db.Execute("EXEC SGP_S_CargaGrillaHisorial '" & Format(FpFecDesde.text, "yyyymmdd") & "', '" & Format(FpFecHasta.text, "yyyymmdd") & "'")

If Not RS.EOF Then
  Do While Not RS.EOF
      
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      
      vaSpread1.Col = 1
      vaSpread1.CellType = 5
      vaSpread1.text = RS(0)
      
      vaSpread1.Col = 2
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS(1)
      
      vaSpread1.Col = 3
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = Trim(RS(2))
      
      vaSpread1.Col = 4
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = Trim(RS(3))
      
      vaSpread1.Col = 5
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = IIf(IsNull(RS(4)), "", (RS(4)))
      
      vaSpread1.Col = 6
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = IIf(IsNull(RS(5)), "", (RS(5)))

      vaSpread1.Col = 7
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = Trim(RS(6))


      
      RS.MoveNext
  Loop
Else
   vaSpread1.MaxRows = 0
   'MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo
End If
RS.Close
Set RS = Nothing
'vaSpread1.Col = 1
'vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
'vaSpread1.ColUserSortIndicator(IIf(Trim(vaSpread1.text) = "", 0, 0)) = ColUserSortIndicatorAscending
'vaSpread1.SortKey(1) = IIf(Trim(vaSpread1.text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
'vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    Resume
End Sub
