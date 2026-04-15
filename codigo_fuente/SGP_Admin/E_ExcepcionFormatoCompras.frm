VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form E_ExcepcionFormatoCompras 
   Caption         =   "Exportar Excel Excepción Formato Compras"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   435
      Index           =   0
      Left            =   2760
      TabIndex        =   7
      Top             =   5880
      Width           =   900
      Begin VB.TextBox TextDet2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   45
         TabIndex        =   8
         Top             =   135
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Height          =   435
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   5880
      Width           =   900
      Begin VB.TextBox TextDet2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   45
         TabIndex        =   6
         Top             =   135
         Width           =   795
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   7320
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   8640
      TabIndex        =   1
      Top             =   6360
      Width           =   1215
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   9855
      _Version        =   393216
      _ExtentX        =   17383
      _ExtentY        =   8493
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
      SpreadDesigner  =   "E_ExcepcionFormatoCompras.frx":0000
   End
   Begin EditLib.fpDateTime fpDateTime1 
      Height          =   315
      Left            =   2130
      TabIndex        =   3
      Top             =   360
      Width           =   1425
      _Version        =   196608
      _ExtentX        =   2514
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
      ButtonStyle     =   3
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   -1  'True
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
      Text            =   ""
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
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Un momento por favor procesando información ....."
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
      Left            =   240
      TabIndex        =   9
      Top             =   6600
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha (dd/mm/aaaa)"
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
      Left            =   240
      TabIndex        =   4
      Top             =   405
      Width           =   1800
   End
End
Attribute VB_Name = "E_ExcepcionFormatoCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgTitulo As String

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

    Case 0
    
        Me.Hide
        Unload Me

    Case 1
    
        ExportarPlanillaExcel
        
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.Height = 7710
Me.Width = 10560
MsgTitulo = "Exportar Excel Excepción Formato Compras"
fg_centra Me

LlenarCeco

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Sub LlenarCeco()

On Error GoTo Man_Error

Dim RS    As New ADODB.Recordset

vaSpread1.MaxRows = 0
vaSpread1.Visible = False

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
Set RS = vg_db.Execute("sgpadm_Sel_MostrarCecoExpecionFormatoExcel_V01")

If RS.EOF Then
           
   RS.Close
   Set RS = Nothing
   MsgBox "No existe información a exportar excel", vbInformation, "Ingredientes"
   vaSpread1.Visible = True
           
   Exit Sub
        
End If
        
Do While Not RS.EOF
                          
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   
   vaSpread1.Row = vaSpread1.MaxRows
            
   vaSpread1.Col = 1
   vaSpread1.text = "0"
            
   vaSpread1.Col = 2
   vaSpread1.text = RS!ID_Orgcompra
            
   vaSpread1.Col = 3
   vaSpread1.text = RS!Id_Ceco
   
   vaSpread1.Col = 4
   vaSpread1.text = RS!Cli_nombre
   
   vaSpread1.Col = 5
   vaSpread1.text = 0
            
   RS.MoveNext
        
Loop

RS.Close
Set RS = Nothing

vaSpread1.Visible = True

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Sub ExportarPlanillaExcel()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim i               As Long
Dim Selec           As Boolean
Dim OrgCompra       As String
Dim Ceco            As String
Dim MyBuffer        As String
Dim NomArchivoExcel As String
Dim Extension       As String

Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object

Selec = False

'Validar fecha

If Trim(fpDateTime1.text) = "" Then

   MsgBox "Debe seleccionar la fecha...", vbExclamation + vbOKOnly, MsgTitulo

   Exit Sub

End If

'Validar seleccion
For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    
    vaSpread1.Col = 1
    If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
    
       Selec = True
       
    End If
    
Next i

If Not Selec Then

   MsgBox "Debe haber por lo menos un ítem seleccionado de la lista...", vbExclamation + vbOKOnly, MsgTitulo

   Exit Sub
   
End If

'crear xml
Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<OrgCeco>"

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    
    vaSpread1.Col = 1
    If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
    
       vaSpread1.Col = 2
       OrgCompra = vaSpread1.text
       
       vaSpread1.Col = 3
       Ceco = vaSpread1.text
       
       MyBuffer = MyBuffer & " <OCE"
       MyBuffer = MyBuffer & " OC = " & Chr(34) & OrgCompra & Chr(34)
       MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
       MyBuffer = MyBuffer & "/>"
       
    End If
    
Next i

MyBuffer = MyBuffer & "</OrgCeco>"
      
Label1(0).Visible = True
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_XmlExportarExcelExcecionFormato_V01 '" & MyBuffer & "', " & LimpiaDato(fpDateTime1.DateValue) & "")

Label1(0).Visible = False

If Not RS.EOF Then
  
   If RS.RecordCount > 1020000 Then
      
      RS.Close
      Set RS = Nothing
      
      MsgBox "El resultado sobrepasa maximo de fila en excel, Debera seleccionar menos Ceco", vbCritical
      Exit Sub
   
   End If
  
End If

'-------> Guardar nombre archivo excel
NomArchivoExcel = ""
CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
CD.Filter = "Todos los archivos *.xls,*.xlsx"
On Error Resume Next
CD.ShowSave
           
'-------> JPAZ Permite controlar Boton Cancelar
If Err.Number = 32755 Then
   
   MsgBox "Proceso cancelado"
   Exit Sub

End If
            
If CD.FileName = "" Then
   
   MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
   Exit Sub

Else
   
   Extension = ""
   Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
   
   If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
      MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
      Exit Sub
   End If
   
   NomArchivoExcel = CD.FileName

End If
          
fg_carga ""
  
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

Dim XL As New excel.Application 'Crea el objeto excel
XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
XL.Visible = True
XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
'-------> Close ADO objects
RS.Close
Set RS = Nothing
    
' -- Cerrar Excel
xlApp.Quit
'-------> Release Excel references
Set xlWs = Nothing
Set xlWb = Nothing
Set xlApp = Nothing
  
fg_descarga
MsgBox "Proceso Finalizado", vbInformation, Me.Caption

Exit Sub
Man_Error:
fg_descarga
Label1(0).Visible = False
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub TextDet2_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(TextDet2(Index).text, ",")

If Index = 2 Then
   
   TextDet2(3).text = ""

ElseIf Index = 3 Then
   
   TextDet2(2).text = ""

End If

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 5
    vaSpread1.text = 0
    
Next

Select Case Index

Case 2, 3

    vaSpread1.Visible = False
    
    If Trim(TextDet2(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like IIf(Index = 2 And Index = 3, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
           vaSpread1.Col = Index
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 5
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 5
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 5
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 5
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 5
                 vaSpread1.text = 0
                 
              End If
           
           End If
        
        Next i
        
        vaSpread1.SetActiveCell Index + 1, 1
        vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread1.ColUserSortIndicator(IIf(Trim(wvarArr8020(X)) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread1.SortKey(1) = IIf(Trim(wvarArr8020(X)) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(TextDet2(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 5
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextDet2(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

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
    
    For i = 1 To vaSpread1.MaxRows 'BlockRow To BlockRow2
        
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

