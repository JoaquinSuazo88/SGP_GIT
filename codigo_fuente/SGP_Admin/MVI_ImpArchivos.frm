VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MVI_ImpArchivos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importacion de Archivos"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraArchivos 
      Caption         =   "Seleccione un tipo de archivo :"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6615
      Begin VB.OptionButton optConvenio 
         Caption         =   "Archivo Importar convenios"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.OptionButton optExcel 
         Caption         =   "Archivo Excel de rutas"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton btnExplorar 
      Caption         =   "Abrir Archivo"
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtRutaArchivo 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1185
      Width           =   5775
   End
   Begin MSComDlg.CommonDialog Cd 
      Left            =   120
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "MVI_ImpArchivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ImporConvenio()

'***** VARS. ARCHIVO PLANO
Dim sReg_info As String
Dim sProveedor As String
Dim sDenominacion_Proveedor As String
Dim sMaterial As String
Dim sDenominacion_Material As String
Dim sOrgC As String
Dim sDenominacion As String
Dim sCe As String
Dim sDenominacion_Centro As String
Dim sGCp As String
Dim sPrec_neto As String
Dim sMon As String
Dim sCtd_mn As String
Dim sTipo_de_Co As String
Dim sPerfil_de As String
Dim sPzE As String
Dim sPlazo_Anul As String
Dim sValido_de As String
Dim sValidez_a As String
Dim sImporte As String
Dim sUn As String
Dim spor As String
Dim sUM As String
Dim sB As String
'***** FIN VARS. ARCHIVO PLANO

Dim NombreArchivoExcel As String
Dim cont As Integer

Dim LineaSplitted() As String

cont = 0

Cd.DialogTitle = "Seleccionar Un Archivo TXT"
Cd.Filter = "Todos los archivos|*.*|Archivos de texto (*.txt)|*.txt"
Cd.FilterIndex = 2
Cd.Flags = cdlOFNFileMustExist
Cd.ShowOpen

If Cd.FileName = "" Or IsNull(Cd.FileName) Then
    Exit Sub
Else
    NombreArchivoExcel = Cd.FileName
End If

txtRutaArchivo = NombreArchivoExcel
'*********************************************************************************************************

Dim Linea As String
  
    ' Get a free file number
    FileNum = FreeFile
    
    ' Open a text file for input. inputbox returns the path to read the file
    Open NombreArchivoExcel For Input As FileNum
    LineCount = 1
    
    ' Read the contents of the file
    Do While Not EOF(FileNum)

      ' Read line
      Line Input #FileNum, Linea
      'Printer.Print Linea_Actual ' imprime con
      
      LineaSplitted = Split(Linea, vbTab)
      
      'hace lectura de la 1era fila de datos
      If cont > 4 Then
            
            sReg_info = LineaSplitted(1)
            sProveedor = LineaSplitted(2)
            sDenominacion_Proveedor = LineaSplitted(3)
            sMaterial = LineaSplitted(4)
            sDenominacion_Material = LineaSplitted(5)
            sOrgC = LineaSplitted(6)
            sDenominacion = LineaSplitted(7)
            sCe = LineaSplitted(8)
            sDenominacion_Centro = LineaSplitted(9)
            sGCp = LineaSplitted(10)
            sDenominacion = LineaSplitted(11)
            sPrec_neto = LineaSplitted(12)
            sMon = LineaSplitted(13)
            sCtd_mn = LineaSplitted(14)
            sTipo_de_Co = LineaSplitted(15)
            sPerfil_de = LineaSplitted(16)
            sPzE = LineaSplitted(17)
            sPlazo_Anul = LineaSplitted(18)
            sValido_de = LineaSplitted(19)
            sValidez_a = LineaSplitted(20)
            sImporte = LineaSplitted(21)
            sUn = LineaSplitted(22)
            spor = LineaSplitted(23)
            sUM = LineaSplitted(24)
            sB = LineaSplitted(25)
      

            sql = " INSERT INTO convenios_mvi"
            sql = sql & " (Reg_info, Proveedor, Denominacion_Proveedor, Material, Denominacion_Material, OrgC, Denominacion1, Ce, Denominacion_Centro, GCp"
            sql = sql & " ,Denominacion2, Prec_neto, Mon, Ctd_mn, Tipo_de_Co, Perfil_de, PzE, Plazo_Anul, Valido_de, Validez_a, Importe, Un, por, UM, B)"
            sql = sql & " VALUES ("
                  
            sql = sql & "'" & sReg_info & "'"
            sql = sql & ",'" & sProveedor & "'"
            sql = sql & ",'" & sDenominacion_Proveedor & "'"
            sql = sql & ",'" & sMaterial & "'"
            sql = sql & ",'" & sDenominacion_Material & "'"
            sql = sql & ",'" & sOrgC & "'"
            sql = sql & ",'" & sDenominacion & "'"
            sql = sql & ",'" & sCe & "'"
            sql = sql & ",'" & sDenominacion_Centro & "'"
            sql = sql & ",'" & sGCp & "'"
            sql = sql & ",'" & sDenominacion & "'"
            sql = sql & ",'" & sPrec_neto & "'"
            sql = sql & ",'" & sMon & "'"
            sql = sql & ",'" & sCtd_mn & "'"
            sql = sql & ",'" & sTipo_de_Co & "'"
            sql = sql & ",'" & sPerfil_de & "'"
            sql = sql & ",'" & sPzE & "'"
            sql = sql & ",'" & sPlazo_Anul & "'"
            sql = sql & ",'" & sValido_de & "'"
            sql = sql & ",'" & sValidez_a & "'"
            sql = sql & ",'" & sImporte & "'"
            sql = sql & ",'" & sUn & "'"
            sql = sql & ",'" & spor & "'"
            sql = sql & ",'" & sUM & "'"
            sql = sql & ",'" & sB & "'"
                  
            sql = sql & ")"
                  
            vg_db.Execute (sql)
      
      
      End If
      
      cont = cont + 1
   Loop
     


   Exit Sub
errSub:
     MsgBox "Ha ocurrido un error : " & Err.Description & " " & Err.Number, vbCritical, "Error"
     'Exportar_ADO_Excel = False
     Me.Enabled = True
 
 End Sub

Private Sub ImpExcel()

    '****** VARS MVI ********
    
    Dim sID_ruta_compra As String
    Dim sFecha_despacho As String
    Dim sID_centro_de_costo As String
    Dim sFamilia_producto As String
    Dim sID_proveedor As String
    Dim sSucursal As String
    Dim sSigla_de_ruta As String
    Dim sDescripcion_sigla As String
    Dim sObservaciones As String
    
    '************************

On Error GoTo errSub

    Dim List() As String
    Dim ListCount As Integer
    Dim fromRight As Long, i As Long
    Dim varManejo As Integer
    Dim varRuta As String
    Dim f As Boolean
    Dim NombreArchivoExcel As String
    Dim dbexcel As Database, cSpi As Long
    Dim ExcelCodProd, ExcelFecha, ExcelPrecio As String
    Dim j As Long
    Dim wvarHoja As String
    Dim wvarCol1, wvarCol2, wvarCol3, wvarCol4 As String
    Dim GrillaCodProd, GrillaDescrip, GrillaFecha, GrillaPrecio As String
    ReDim List(1)

    Cd.DialogTitle = "Seleccionar Un Archivo XLS"
    Cd.Filter = "Todos los archivos|*.*|Archivos de texto (*.xls)|*.xls"
    Cd.FilterIndex = 2
    Cd.Flags = cdlOFNFileMustExist
    Cd.ShowOpen
    
    If Cd.FileName = "" Or IsNull(Cd.FileName) Then
        Exit Sub
    Else
        NombreArchivoExcel = Cd.FileName
    End If
    
    txtRutaArchivo = NombreArchivoExcel
    
    'If Len(NombreArchivoExcel) = 0 Then Exit Sub
    
    'Label4.Caption = "Cargando Información ..."
    'Label4.Visible = True
    
    fromRight = InStrRev(Cd.FileName, "\", , vbTextCompare)
    
    If fromRight > 1 Then
       varRuta = Left(Cd.FileName, fromRight)
    End If
    
    'vaSpread3.MaxRows = 0: vaSpread3.MaxRows = 500
    'vaSpread3.MaxCols = 0: vaSpread3.MaxCols = 500
    
    'f = vaSpread3.GetExcelSheetList(Cd.FileName, List, ListCount, (varRuta & "log.txt"), varManejo, True)
    
'    If (ListCount - 1 > 1) Then
'       ReDim List(ListCount - 1)
'       f = vaSpread3.GetExcelSheetList(Cd.FileName, List, ListCount, (varRuta & "log.txt"), varManejo, False)
'    End If
        
    'wvarHoja = (List(0))
 
    Dim RsExcel As ADODB.Recordset
    Dim sconn As String
    Set RsExcel = New ADODB.Recordset
    
    RsExcel.CursorLocation = adUseClient
    RsExcel.CursorType = adOpenKeyset
    RsExcel.LockType = adLockBatchOptimistic
    
    sconn = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & NombreArchivoExcel
    
    'linea para excel 2007 o superior. Necesita bajar drivers extra.
    'sconn = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & NombreArchivoExcel
    
    wvarHoja = "RUTAS tf DICIEMBRE"
    
    strSQL = "SELECT * FROM [" & wvarHoja & "$]"
    'strSQL = "SELECT * FROM [Hoja1$]"
        
    RsExcel.Open strSQL, sconn
    
    'CuentaError = 0
    
    'vaSpread2.MaxRows = Val(RsExcel.RecordCount)

    j = 1
    
    Do While Not RsExcel.EOF
        
        'sID_ruta_compra = RsExcel(0)
        sFecha_despacho = RsExcel(0)
        sID_centro_de_costo = RsExcel(1)
        sFamilia_producto = RsExcel(2)
        sID_proveedor = RsExcel(3)
        sSucursal = RsExcel(4)
        sSigla_de_ruta = RsExcel(5)
        sDescripcion_sigla = RsExcel(6)
        sObservaciones = RsExcel(7)

        
        sql = " INSERT INTO ruta_compras"
        sql = sql & " (Fecha_despacho, ID_centro_de_costo, Familia_producto, ID_proveedor, Sucursal, Sigla_de_ruta, Descripcion_sigla, Observaciones)"
        sql = sql & " VALUES ("

        'sql = sql & "'" & sID_ruta_compra & "'"
        sql = sql & "'" & sFecha_despacho & "'"
        sql = sql & ",'" & sID_centro_de_costo & "'"
        sql = sql & ",'" & sFamilia_producto & "'"
        sql = sql & ",'" & sID_proveedor & "'"
        sql = sql & ",'" & sSucursal & "'"
        sql = sql & ",'" & sSigla_de_ruta & "'"
        sql = sql & ",'" & sDescripcion_sigla & "'"
        sql = sql & ",'" & sObservaciones & "'"
        'sql = sql & ",'" & sDenominacion_Centro & "'"
        'sql = sql & ",'" & sGCp & "'"
        
        sql = sql & ")"


'        If ((IsNull(RsExcel(0)) And IsNull(RsExcel(1))) Or (Trim(RsExcel(0)) = "" And Trim(RsExcel(1)) = "")) Then
'            j = j + 1
'            RsExcel.MoveNext
'            'Exit Do
'        End If
'
'        If RsExcel.EOF Then
'            Exit Do
'        End If
'
'        ExcelCodProd = IIf(IsNull(Trim(RsExcel(0))), "", Trim(RsExcel(0)))
'
'        ExcelPrecio = IIf(IsNull(Trim(RsExcel(1))), "", Trim(RsExcel(1)))
        
        'strSQL = "SELECT * FROM b_formatocompras WHERE foc_codsac = '" & ExcelCodProd & "'"
        'Set RS = vg_db.Execute(strSQL)
        
        vg_db.Execute (sql)

        
        j = j + 1
        RsExcel.MoveNext
    Loop
        
    
    RsExcel.Close: Set RsExcel = Nothing
    RS.Close: Set RS = Nothing
    
    Exit Sub

errSub:
     MsgBox "Ha ocurrido un error : " & Err.Description & " " & Err.Number, vbCritical, "Error"
     
     'Me.Enabled = True

End Sub

Private Sub btnExplorar_Click()

'validaciones

If optConvenio.Value = False And Me.optExcel.Value = False Then
    MsgBox "Debe seleccionar un tipo de archivo", vbExclamation, Me.Caption
    Exit Sub
End If


If optConvenio.Value = True Then
    Call ImporConvenio
Else
    Call ImpExcel
End If

End Sub
