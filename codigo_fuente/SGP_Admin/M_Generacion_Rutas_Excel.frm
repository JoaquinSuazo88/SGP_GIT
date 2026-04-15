VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form M_Generacion_Rutas_Excel 
   Caption         =   "Generacion de Rutas Despachos"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9600
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Tipo Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   15
      Top             =   1440
      Width           =   5655
      Begin VB.OptionButton Option1 
         Caption         =   "PAP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1320
      TabIndex        =   5
      Top             =   7320
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   4080
      TabIndex        =   7
      Top             =   7320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   2520
      TabIndex        =   6
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mostrar en Grilla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2280
      TabIndex        =   11
      Top             =   7800
      Width           =   3135
      Begin VB.CheckBox Check1 
         Caption         =   "Sitios Simap"
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
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sitios No Simap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Sitios FM"
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
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1575
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4215
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   8415
      _Version        =   393216
      _ExtentX        =   14843
      _ExtentY        =   7435
      _StockProps     =   64
      AutoClipboard   =   0   'False
      ButtonDrawMode  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   1
      ProcessTab      =   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "M_Generacion_Rutas_Excel.frx":0000
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   4920
      TabIndex        =   1
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   106692609
      CurrentDate     =   41744
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   106692609
      CurrentDate     =   41744
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Desde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   13
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Hasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "M_Generacion_Rutas_Excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public filtro As String
Public GenPed As Boolean

Private Sub Check1_Click()

Call lee_fechas_cecos

End Sub

Private Sub Check2_Click()

Call lee_fechas_cecos

End Sub

Private Sub Check3_Click()

Call lee_fechas_cecos

End Sub

Private Sub DTPicker1_Change()

On Error GoTo Man_Error

If IsDate(DTPicker1.Value) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
    
End Sub

Private Sub DTPicker2_Change()

On Error GoTo Man_Error

If IsDate(DTPicker2.Value) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Private Sub DTPicker2_KeyPress(KeyAscii As Integer)
  
On Error GoTo Man_Error

  If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Private Sub Form_Load()
 
 On Error GoTo Man_Error
    
  fg_centra Me
  Toolbar1.ImageList = Partida.IL1
  
  Me.HelpContextID = 1196002

  Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False): BtnX.ToolTipText = "Exportar a Excel "
  Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
  Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
  
  Call lee_fechas_cecos
  
  Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Private Sub lee_fechas_cecos()
 
 On Error GoTo Man_Error
 
 Dim parametro As Long
 parametro = 1
 
 
 Text1(2) = ""
 Text1(3) = ""
 DTPicker1.Value = Date
 DTPicker2.Value = Date
    
    filtro = ""
    If Check1.Value = 1 Then
       filtro = filtro + "SI"
    Else
      filtro = filtro + "XX"
    End If
    If Check2.Value = 1 Then
       filtro = filtro + "NS"
    Else
      filtro = filtro + "XX"
    End If
    If Check3.Value = 1 Then
       filtro = filtro + "FM"
    Else
      filtro = filtro + "XX"
    End If
     
    Sql = " sgpadm_sel_paramETROS_despacho_casino_V01  " & parametro & ",'" & filtro & "'"
    Set RS = vg_db.Execute(Sql)
    
   '-------> Inicio LLenar grilla
   
   With vaSpread1
    
    .MaxRows = 0
 
    Do While Not RS.EOF
        
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        
        .Col = 1 ' Seleccion
        .text = 0
        .TypeHAlign = TypeHAlignCenter
        
        
        .Col = 2 ' Org. Compras
        .text = RS(14)
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 3 ' Ceco
        .text = RS(0)
        .TypeHAlign = TypeHAlignCenter
        
        .Col = 4 ' Nombre
        .text = RS(1)
        
        .Col = 5 'Mover estado
        .text = 0
        
        RS.MoveNext
        
    Loop
    
    End With


Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub Form_Unload(Cancel As Integer)

'registrar Log sistema salir
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "", "")

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(Text1(Index).text, ",")

If Index = 2 Then
   Text1(3).text = ""
   Text1(4).text = ""
ElseIf Index = 3 Then
   Text1(2).text = ""
   Text1(4).text = ""
ElseIf Index = 4 Then
   Text1(2).text = ""
   Text1(3).text = ""
End If

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 5
    vaSpread1.text = 0

Next

Select Case Index

Case 2, 3, 4
    
    vaSpread1.Visible = False
    
    If Trim(Text1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread1.Col = 1
           
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
              vaSpread1.Col = 1
              
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
    
    If Trim(Text1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 5
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error
   
   Select Case Button.Index
   
   Case 1 'Exportar a Excel
        
        Call lleva_excel
   
   Case 3 ' Salir del Programa
        
        Me.Hide
        Unload Me
    
    End Select

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
If Err.Number = -2147467259 Then
   MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"
Else
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If

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
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    If BlockRow = -1 Then Exit Sub
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

' -------------------------------------------------------------------------------------------
' \\ -- Función para crear un nuevo libro con el contenido del Grid
' -------------------------------------------------------------------------------------------
Public Function Exportar_Excel() As Boolean
  
    On Error GoTo Error_Handler
    
    Exportar_Excel = False
    Dim rst As New ADODB.Recordset
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    
    Dim recArray As Variant
    
    Dim strDB As String
    Dim fldCount As Integer
    Dim recCount As Long
    Dim icol As Integer
    Dim iRow As Integer

    Dim seleccion As Long
    Dim Ceco As String
    Dim MyBuffer    As String
    
    '-------> Armar xml
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<NewDataSet>"
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        If seleccion = 1 Then
           
           vaSpread1.Col = 3 'CECO
           Ceco = IIf(vaSpread1.text = "", 0, vaSpread1.text)
           
           MyBuffer = MyBuffer & " <Ceco"
           MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
           MyBuffer = MyBuffer & "/>"
        
        End If
    
    Next i
    
    MyBuffer = MyBuffer & "</NewDataSet>"

    '-------> Lectura
    Sql = ""
    If GenPed = False Then
       
       Sql = " sgpadm_sel_Xmlrutadespacho_fecha_Familia "
    
    ElseIf GenPed And Option1(0).Value = True Then
       
       Sql = " sgpadm_sel_Xmlrutadespacho_fecha_PELCD "
    
    ElseIf GenPed And Option1(1).Value = True Then
    
       Sql = " sgpadm_sel_Xmlrutadespacho_fecha_PELPAP "
       
    End If
    
    Sql = Sql & " '" & MyBuffer & "',"
    Sql = Sql & " '" & Format(DTPicker1, "YYYYMMDD") & "',"
    Sql = Sql & " '" & Format(DTPicker2, "YYYYMMDD") & "',"
    Sql = Sql & " '" & filtro & "'"

    rst.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set rst = vg_db.Execute(Sql)
    
    ' Create an instance of Excel and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets("Hoja1")
    
    If rst.RecordCount > xlWs.Range("A1", xlWs.Range("A1").End(xlDown)).Rows.count Then
    
       ' Close ADO objects
       rst.Close
       Set rst = Nothing
       
       ' Release Excel references
       Set xlWs = Nothing
       Set xlWb = Nothing

       Set xlApp = Nothing
       
       MsgBox "Excede numero filas, debera bajar la fecha despacho", vbCritical
       
       Exportar_Excel = False
       Exit Function
    
    End If
    
    ' Display Excel and give user control of Excel's lifetime
    xlApp.Visible = True
    xlApp.UserControl = True
    
    ' Check version of Excel
    Call encabezado(rst, xlWs)
          
    xlWs.Cells(2, 1).CopyFromRecordset rst
    
    ' Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit

    ' Close ADO objects
    rst.Close
    Set rst = Nothing
    
    ' Release Excel references
    Set xlWs = Nothing
    Set xlWb = Nothing

    Set xlApp = Nothing
    
    Exportar_Excel = True

Exit Function
  
' -- Controlador de Errores
Error_Handler:
    Set rst = Nothing
'    Set cnt = Nothing
    Set xlWs = Nothing
    Set xlWb = Nothing

    MsgBox Err.Description, vbCritical, MsgTitulo

End Function

Sub encabezado(ByRef rst As ADODB.Recordset, ByRef xlWs As Object)

On Error GoTo Man_Error

        ' Copy field names to the first row of the worksheet
        fldCount = rst.Fields.count
        For icol = 1 To fldCount
            
            xlWs.Cells(1, icol).Value = rst.Fields(icol - 1).Name

        Next

Exit Sub
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub lleva_excel()

On Error GoTo Man_Error

Dim Conta As Long


If Format(DTPicker1, "YYYYMMDD") > Format(DTPicker2, "YYYYMMDD") Then
   
   MsgBox "Le fecha hasta  no puede ser menor que la fecha desde ", vbExclamation
   Exit Sub

End If

 For i = 1 To vaSpread1.MaxRows
          
      vaSpread1.Row = i
      
      vaSpread1.Col = 1 'Seleccion
      seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
      
      If seleccion = 1 Then
        
        Conta = Conta + 1
        Exit For
      
      End If
      
    Next i
    
    If Conta = 0 Then
       
       MsgBox "Debe seleccionar por lo menos un casino"
       Exit Sub
    
    End If

Dim r As Boolean


If Not Exportar_Excel Then
    
    MsgBox "Ocurrio un error al exportar", vbCritical

Else
   
   MsgBox "Exportación realizada con exito", vbInformation

   'registrar Log sistema grabar
   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Exporta_Excel"), Me.HelpContextID, "", "", "")
 
End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub MoverOptionInicio(Op As String)

On Error GoTo Man_Error

Frame2.Visible = IIf(Op = "1", False, True)
GenPed = IIf(Op = "1", False, True)

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

