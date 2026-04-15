VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MVI_ImpArchivos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importacion de Archivos"
   ClientHeight    =   6540
   ClientLeft      =   5025
   ClientTop       =   2595
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6870
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   270
      Left            =   600
      TabIndex        =   12
      Top             =   6480
      Visible         =   0   'False
      Width           =   630
      _Version        =   393216
      _ExtentX        =   1111
      _ExtentY        =   476
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
      SpreadDesigner  =   "MVI_ImpArchivos.frx":0000
   End
   Begin VB.ComboBox CboHojaExcel 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1620
      Width           =   4095
   End
   Begin VB.CommandButton CmdSalir 
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
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton btnEliminar 
      Caption         =   "Eliminar Seleccionados"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seleccionar"
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
      Left            =   5280
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txterrores 
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2040
      Width           =   6495
   End
   Begin VB.Frame fraArchivos 
      Caption         =   "Seleccione un tipo de archivo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6615
      Begin VB.OptionButton optConvenio 
         Caption         =   "Archivo Importar convenios"
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
         Left            =   3720
         TabIndex        =   5
         Top             =   360
         Width           =   2655
      End
      Begin VB.OptionButton optExcel 
         Caption         =   "Archivo Excel de rutas"
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton btnExplorar 
      Caption         =   "Cargar Archivo"
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
      Left            =   5280
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtRutaArchivo 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1185
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog Cd 
      Left            =   120
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   1545
      Left            =   240
      TabIndex        =   8
      Top             =   4200
      Width           =   6435
      _Version        =   393216
      _ExtentX        =   11351
      _ExtentY        =   2725
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      EditEnterAction =   4
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
      MaxCols         =   4
      SpreadDesigner  =   "MVI_ImpArchivos.frx":01D4
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   240
      Left            =   240
      TabIndex        =   14
      Top             =   6000
      Visible         =   0   'False
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Hoja :"
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
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo :"
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
Dim TheFile As Archivo
Dim est As Boolean
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Function ValidaCamposImporConvenio(RutaArchivo As String) As Boolean

ValidaCamposImporConvenio = False

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

Dim mensaje As String
mensaje = ""

Dim NombreArchivoExcel As String
Dim cont As Integer

Dim LineaSplitted() As String

cont = 1

'Cd.DialogTitle = "Seleccionar Un Archivo TXT"
'Cd.Filter = "Todos los archivos|*.*|Archivos de texto (*.txt)|*.txt"
'Cd.FilterIndex = 2
'Cd.Flags = cdlOFNFileMustExist
'Cd.ShowOpen
'
'If Cd.Filename = "" Or IsNull(Cd.Filename) Then
'    Exit Function
'Else
'    NombreArchivoExcel = Cd.Filename
'End If

NombreArchivoExcel = Me.txtRutaArchivo

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
      If cont > 5 Then
            
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
    
'            If sReg_info = "" Or sProveedor = "" Or sDenominacion_Proveedor = "" Or sMaterial = "" Or sDenominacion_Material = "" Or sOrgC = "" Or sDenominacion = "" Or sCe = "" Or sDenominacion_Centro = "" Or sGCp = "" Or sDenominacion = "" Or sPrec_neto = "" Or sMon = "" Or sCtd_mn = "" Or sTipo_de_Co = "" Or sPerfil_de = "" Or sPzE = "" Or sPlazo_Anul = "" Or sImporte = "" Or sUn = "" Or spor = "" Or sB = "" Or sValido_de = "" Or sValidez_a = "" Then 'Not IsDate(CDate(Replace(sValido_de, ".", "/"))) Or Not IsDate(CDate(Replace(sValidez_a, ".", "/")))
'                mensaje = mensaje & "Error en la linea N° " & CStr(cont) & vbNewLine
'                ValidaCamposImporConvenio = True
'            End If

            'validaciones obligatorios
            Dim Valida_de As String
            Dim Validez_a As String
            
            Valida_de = sValido_de
            Validez_a = sValidez_a
            
            'limpieza de caracteres basura
            Valida_de = Replace(Valida_de, ".", "/")
            Valida_de = Replace(Valida_de, "-", "/")
            
            Validez_a = Replace(Validez_a, ".", "/")
            Validez_a = Replace(Validez_a, "-", "/")
            
            If Not IsDate(CDate(Valida_de)) Then
                mensaje = mensaje & "Fecha Valida_de errónea, Error en la linea N° " & CStr(cont) & vbNewLine
                ValidaCamposImporConvenio = True
            ElseIf Not IsDate(CDate(Validez_a)) Then
                mensaje = mensaje & "Fecha Validez_a errónea, Error en la linea N° " & CStr(cont) & vbNewLine
                ValidaCamposImporConvenio = True
            End If

            
            
      End If
        
      cont = cont + 1
    
    Loop

    'MsgBox "Errores hayados : " & Mensaje
        
    txterrores.text = mensaje
    
End Function


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
Dim RutaArchivo As String
'***** FIN VARS. ARCHIVO PLANO

Dim NombreArchivoExcel As String
Dim cont As Integer

Dim LineaSplitted() As String

cont = 0

'Cd.DialogTitle = "Seleccionar Un Archivo TXT"
'Cd.Filter = "Todos los archivos|*.*|Archivos de texto (*.txt)|*.txt"
'Cd.FilterIndex = 2
'Cd.Flags = cdlOFNFileMustExist
'Cd.ShowOpen
'
'If Cd.Filename = "" Or IsNull(Cd.Filename) Then
'    Exit Sub
'Else
'    NombreArchivoExcel = Cd.Filename
'End If

'txtRutaArchivo = NombreArchivoExcel
RutaArchivo = txtRutaArchivo
'*********************************************************************************************************

Dim Linea As String
  
'validacion de los campos del archivo

If ValidaCamposImporConvenio(NombreArchivoExcel) = True Then
    MsgBox "No se proseguirá con la carga del archivo por errores en los campos de este", vbCritical, Me.Caption
    Exit Sub
End If

'fin validacion de los campos del archivo
  
    ' Get a free file number
    FileNum = FreeFile
    
    ' Open a text file for input. inputbox returns the path to read the file
    Open RutaArchivo For Input As FileNum
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
            
            Valida_de = (LineaSplitted(19))
            
            Valida_de = Replace(Valida_de, ".", "/")
            Valida_de = Replace(Valida_de, "-", "/")
            
            Validez_a = (LineaSplitted(20))

            Validez_a = Replace(Validez_a, ".", "/")
            Validez_a = Replace(Validez_a, "-", "/")
            
'            sValido_de = Valida_de 'LineaSplitted(19)
'            sValidez_a = LineaSplitted(20)
            sImporte = LineaSplitted(21)
            sUn = LineaSplitted(22)
            spor = LineaSplitted(23)
            sUM = LineaSplitted(24)
            sB = LineaSplitted(25)
      
            'chequea existencia (insert o update)
            sql = " SELECT Reg_info"
            sql = sql & " From convenios_mvi"
            sql = sql & " WHERE Reg_info = '" & sReg_info & "'"
            sql = sql & " AND Proveedor = '" & sProveedor & "'"
            sql = sql & " AND Material = '" & sMaterial & "'"
            sql = sql & " AND Ce = '" & sCe & "'"
            
            Set RS = vg_db.Execute(sql)
            
            If RS.EOF Then
                sql = " INSERT INTO convenios_mvi"
                sql = sql & " (Reg_info, Proveedor, Denominacion_Proveedor, Material, Denominacion_Material, OrgC, Denominacion1, Ce, Denominacion_Centro, GCp"
                sql = sql & " ,Denominacion2, Prec_neto, Mon, Ctd_mn, Tipo_de_Co, Perfil_de, PzE, Plazo_Anul, Valido_de, Validez_a, Importe, Un, por, UM, B, Ruta_archivo)"
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
                sql = sql & ",'" & Format(Valida_de, "YYYYMMDD") & "'"
                sql = sql & ",'" & Format(Validez_a, "YYYYMMDD") & "'"
                sql = sql & ",'" & sImporte & "'"
                sql = sql & ",'" & sUn & "'"
                sql = sql & ",'" & spor & "'"
                sql = sql & ",'" & sUM & "'"
                sql = sql & ",'" & sB & "'"
                sql = sql & ",'" & RutaArchivo & "'"
                      
                sql = sql & ")"
            Else
                sql = " Update convenios_mvi"
                
                'Sql = Sql & ",Reg_info = '" & sReg_info & "'"
                'Sql = Sql & ",Proveedor = '" & sProveedor & "'"
                sql = sql & " SET Denominacion_Proveedor = '" & sDenominacion_Proveedor & "'"
                'Sql = Sql & ",Material = '" & sMaterial & "'"
                sql = sql & " ,Denominacion_Material = '" & sDenominacion_Material & "'"
                sql = sql & " ,OrgC = '" & sOrgC & "'"
                sql = sql & " ,Denominacion1 = '" & sDenominacion & "'"
                'Sql = Sql & ",Ce = '" & sCe & "'"
                sql = sql & " ,Denominacion_Centro = '" & sDenominacion_Centro & "'"
                sql = sql & " ,GCp = '" & sGCp & "'"
                sql = sql & " ,Denominacion2 = '" & sDenominacion & "'"
                sql = sql & " ,Prec_neto = '" & sPrec_neto & "'"
                sql = sql & " ,Mon = '" & sMon & "'"
                sql = sql & " ,Ctd_mn = '" & sCtd_mn & "'"
                sql = sql & " ,Tipo_de_Co = '" & sTipo_de_Co & "'"
                sql = sql & " ,Perfil_de = '" & sPerfil_de & "'"
                sql = sql & " ,PzE = '" & sPzE & "'"
                sql = sql & " ,Plazo_Anul = '" & sPlazo_Anul & "'"
                sql = sql & " ,Valido_de = '" & Format(Valida_de, "YYYYMMDD") & "'"
                sql = sql & " ,Validez_a = '" & Format(Validez_a, "YYYYMMDD") & "'"
                sql = sql & " ,Importe = '" & sImporte & "'"
                sql = sql & " ,Un = '" & sUn & "'"
                sql = sql & " ,por = '" & spor & "'"
                sql = sql & " ,UM = '" & sUM & "'"
                sql = sql & " ,B = '" & sB & "'"
                sql = sql & " ,Ruta_archivo = '" & RutaArchivo & "'"
                    
                sql = sql & " WHERE Reg_info = '" & sReg_info & "'"
                sql = sql & " AND Proveedor = '" & sProveedor & "'"
                sql = sql & " AND Material = '" & sMaterial & "'"
                sql = sql & " AND Ce = '" & sCe & "'"
            End If
                
                vg_db.Execute (sql)
    
      End If
      
      cont = cont + 1
   Loop
     
   MsgBox "Proceso finalizado con exito", vbInformation + vbOKOnly, Me.Caption

   Exit Sub
errSub:
     MsgBox "Ha ocurrido un error en la Importación : " & Err.Description & " " & Err.Number, vbCritical, "Error"
     'Exportar_ADO_Excel = False
     Me.Enabled = True
 
 End Sub

Private Sub ImpExcel_Convenios()

'****** VARS MVI ********
    
Dim sReg_info As String
Dim sProveedor As String
Dim sDenominacion_Proveedor As String
Dim sMaterial As String
Dim sDenominacion_Material As String
Dim sOrgC As String
Dim sDenominacion As String
Dim sDenominacion1 As String
Dim sDenominacion2 As String
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
Dim ContReg As Long
    
'************************
Dim mensaje As String
Dim swEntraError As Boolean
Dim FechaHoy As String
    
FechaHoy = CStr(Date)
    
On Error GoTo errSub

Dim List() As String
Dim listcount As Integer
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
Dim id_carga As Long
Dim RS As New ADODB.Recordset
swEntraError = False
    
'-------> Validar hoja este seleccionada
If CboHojaExcel.ListIndex = -1 Then
   MsgBox "Debe seleccionar un hoja excel...", vbCritical, Me.Caption
   Exit Sub
End If
    
Dim RutaArchivo As String
    
Me.txterrores.text = ""
    
NombreArchivoExcel = Me.txtRutaArchivo.text
    
txtRutaArchivo = NombreArchivoExcel
RutaArchivo = txtRutaArchivo
    
fromRight = InStrRev(CD.Filename, "\", , vbTextCompare)
    
If fromRight > 1 Then
   varRuta = Left(CD.Filename, fromRight)
End If
    
Dim RsExcel As ADODB.Recordset
Dim sconn As String
Set RsExcel = New ADODB.Recordset
    
RsExcel.CursorLocation = adUseClient
RsExcel.CursorType = adOpenKeyset
RsExcel.LockType = adLockBatchOptimistic
sconn = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & NombreArchivoExcel
Dim cn As ADODB.Connection
Set cn = New ADODB.Connection
With cn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "Data Source=" & NombreArchivoExcel & ";" & _
    "Extended Properties='Excel 8.0;HDR=NO;IMEX=1'"
    .Open
End With

'linea para excel 2007 o superior. Necesita bajar drivers extra.
'sconn = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & NombreArchivoExcel
    
wvarHoja = CboHojaExcel.text
strSQL = "SELECT * FROM [" & wvarHoja & "$]"
        
RsExcel.Open strSQL, cn
    
j = 1
    
mensaje = ""
swEntraError = False

If Not RsExcel.EOF Then
    'RsExcel.MoveNext
    RsExcel.MoveFirst
    '-------> validacion de los valores en los campos del archivo en excel
    Do While Not RsExcel.EOF
        If j > 5 Then
            sReg_info = IIf(IsNull(RsExcel(1)), "", RsExcel(1))
            sProveedor = IIf(IsNull(RsExcel(2)), "", RsExcel(2))
            sDenominacion_Proveedor = IIf(IsNull(RsExcel(3)), "", RsExcel(3))
            sMaterial = IIf(IsNull(RsExcel(4)), "", RsExcel(4))
            sDenominacion_Material = IIf(IsNull(RsExcel(5)), "", RsExcel(5))
            sOrgC = IIf(IsNull(RsExcel(6)), "", RsExcel(6))
            sDenominacion = IIf(IsNull(RsExcel(7)), "", RsExcel(7))
            sCe = IIf(IsNull(RsExcel(8)), "", RsExcel(8))
            sDenominacion_Centro = IIf(IsNull(RsExcel(9)), "", RsExcel(9))
            sGCp = IIf(IsNull(RsExcel(10)), "", RsExcel(10))
            sDenominacion = IIf(IsNull(RsExcel(11)), "", RsExcel(11))
            sPrec_neto = IIf(IsNull(RsExcel(12)), "", RsExcel(12))
            sMon = IIf(IsNull(RsExcel(13)), "", RsExcel(13))
            sCtd_mn = IIf(IsNull(RsExcel(14)), "", RsExcel(14))
            sTipo_de_Co = IIf(IsNull(RsExcel(15)), "", RsExcel(15))
            sPerfil_de = IIf(IsNull(RsExcel(16)), "", RsExcel(16))
            sPzE = IIf(IsNull(RsExcel(17)), "", RsExcel(17))
            sPlazo_Anul = IIf(IsNull(RsExcel(18)), "", RsExcel(18))
            sValido_de = IIf(IsNull(RsExcel(19)), "", RsExcel(19))
            sValidez_a = IIf(IsNull(RsExcel(20)), "", RsExcel(20))
            sImporte = IIf(IsNull(RsExcel(21)), "", RsExcel(21))
            sUn = IIf(IsNull(RsExcel(22)), "", RsExcel(22))
            spor = IIf(IsNull(RsExcel(23)), "", RsExcel(23))
            sUM = IIf(IsNull(RsExcel(24)), "", RsExcel(24))
            sB = IIf(IsNull(RsExcel(25)), "", RsExcel(25))
            
    '            If sReg_info = "" Or sProveedor = "" Or sDenominacion_Proveedor = "" Or sMaterial = "" Or sDenominacion_Material = "" Or sOrgC = "" Or sDenominacion = "" Or sCe = "" Or sDenominacion_Centro = "" Or sGCp = "" Or sDenominacion = "" Or sPrec_neto = "" Or sMon = "" Or sCtd_mn = "" Or sTipo_de_Co = "" Or sPerfil_de = "" Or sPzE = "" Or sPlazo_Anul = "" Or sImporte = "" Or sUn = "" Or spor = "" Or sB = "" Or sValido_de = "" Or sValidez_a = "" Then 'Not IsDate(CDate(Replace(sValido_de, ".", "/"))) Or Not IsDate(CDate(Replace(sValidez_a, ".", "/")))
    '                mensaje = mensaje & "Error en la linea N° " & CStr(cont) & vbNewLine
    '                swEntraError  = True
    '            End If
    
            '-------> validaciones obligatorios
            Dim Valida_de As String
            Dim Validez_a As String
                
            Valida_de = sValido_de
            Validez_a = sValidez_a
                
            '-------> limpieza de caracteres basura
            Valida_de = Replace(Valida_de, ".", "/")
            Valida_de = Replace(Valida_de, "-", "/")
                
            Validez_a = Replace(Validez_a, ".", "/")
            Validez_a = Replace(Validez_a, "-", "/")
                
            If Trim(sReg_info) = "" Then
               mensaje = mensaje & "Código Convenio debe ser obligatorio, Error en la linea N° " & CStr(cont) & vbNewLine
               swEntraError = True
            End If
            
            If Trim(sProveedor) = "" Then
               mensaje = mensaje & "Código proveedor SAP debe ser obligatorio, Error en la linea N° " & CStr(cont) & vbNewLine
               swEntraError = True
            End If
            
            If Trim(sMaterial) = "" Then
               mensaje = mensaje & "Código material SAP debe ser obligatorio, Error en la linea N° " & CStr(cont) & vbNewLine
               swEntraError = True
            End If
            
            If Trim(sCe) = "" Then
               mensaje = mensaje & "Código centro debe ser obligatorio, Error en la linea N° " & CStr(cont) & vbNewLine
               swEntraError = True
            End If
            
            If Not IsDate(CDate(Valida_de)) Then
               mensaje = mensaje & "Fecha Valida_de errónea, Error en la linea N° " & CStr(cont) & vbNewLine
               swEntraError = True
            ElseIf Not IsDate(CDate(Validez_a)) Then
               mensaje = mensaje & "Fecha Validez_a errónea, Error en la linea N° " & CStr(cont) & vbNewLine
               swEntraError = True
            End If
        End If
        RsExcel.MoveNext: j = j + 1
    Loop
        
    If swEntraError = True Then
        MsgBox "Errores encontrados en el archivo, favor mirar texto de errores , carga cancelada", vbCritical, Me.Caption
        Me.txterrores = mensaje
        Set RsExcel = Nothing
        Exit Sub
    End If
    'fin validacion de los valores en los campos del archivo en excel
    
    j = 1
    i = 1
    RsExcel.MoveFirst
    Dim MyBuffer As Variant
        
    '-------> Genera Convenios
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<GrabaConvenios>"
    
    PB.Min = 0: PB.Value = 0: PB.Visible = True
    ContReg = RsExcel.RecordCount
    Do While Not RsExcel.EOF
        PB.Value = Val((j / ContReg) * 100)
        If j > 5 Then
            sReg_info = IIf(IsNull(RsExcel(1)), "", RsExcel(1))
            sProveedor = IIf(IsNull(RsExcel(2)), "", RsExcel(2))
            sDenominacion_Proveedor = IIf(IsNull(RsExcel(3)), "", RsExcel(3))
            sMaterial = IIf(IsNull(RsExcel(4)), "", RsExcel(4))
            sDenominacion_Material = IIf(IsNull(RsExcel(5)), "", RsExcel(5))
            sOrgC = IIf(IsNull(RsExcel(6)), "", RsExcel(6))
            sDenominacion1 = IIf(IsNull(RsExcel(7)), "", RsExcel(7))
            sCe = IIf(IsNull(RsExcel(8)), "", RsExcel(8))
            sDenominacion_Centro = IIf(IsNull(RsExcel(9)), "", RsExcel(9))
            sGCp = IIf(IsNull(RsExcel(10)), "", RsExcel(10))
            sDenominacion2 = IIf(IsNull(RsExcel(11)), "", RsExcel(11))
            sPrec_neto = IIf(IsNull(RsExcel(12)), "", RsExcel(12))
            sMon = IIf(IsNull(RsExcel(13)), "", RsExcel(13))
            sCtd_mn = IIf(IsNull(RsExcel(14)), "", RsExcel(14))
            sTipo_de_Co = IIf(IsNull(RsExcel(15)), "", RsExcel(15))
            sPerfil_de = IIf(IsNull(RsExcel(16)), "", RsExcel(16))
            sPzE = IIf(IsNull(RsExcel(17)), "", RsExcel(17))
            sPlazo_Anul = IIf(IsNull(RsExcel(18)), "", RsExcel(18))
            
            Valida_de = IIf(IsNull(RsExcel(19)), "", RsExcel(19))
            
            Valida_de = Replace(Valida_de, ".", "/")
            Valida_de = Replace(Valida_de, "-", "/")
            
            Validez_a = IIf(IsNull(RsExcel(20)), "", RsExcel(20))

            Validez_a = Replace(Validez_a, ".", "/")
            Validez_a = Replace(Validez_a, "-", "/")
            
            sImporte = IIf(IsNull(RsExcel(21)), "", RsExcel(21))
            sUn = IIf(IsNull(RsExcel(22)), "", RsExcel(22))
            spor = IIf(IsNull(RsExcel(23)), "", RsExcel(23))
            sUM = IIf(IsNull(RsExcel(24)), "", RsExcel(24))
            sB = IIf(IsNull(RsExcel(25)), "", RsExcel(25))
            
           Let MyBuffer = MyBuffer & " <Convenios"
'           MyBuffer = MyBuffer & " Op = " & Chr(34) & 0 & Chr(34)

            MyBuffer = MyBuffer & " Reg_info = " & Chr(34) & SacarCaracterEspecialesXml(sReg_info) & Chr(34)
            MyBuffer = MyBuffer & " Proveedor = " & Chr(34) & SacarCaracterEspecialesXml(sProveedor) & Chr(34)
            MyBuffer = MyBuffer & " Denominacion_Proveedor = " & Chr(34) & SacarCaracterEspecialesXml(sDenominacion_Proveedor) & Chr(34)
            MyBuffer = MyBuffer & " Material = " & Chr(34) & SacarCaracterEspecialesXml(sMaterial) & Chr(34)
            MyBuffer = MyBuffer & " Denominacion_Material = " & Chr(34) & SacarCaracterEspecialesXml(sDenominacion_Material) & Chr(34)
            MyBuffer = MyBuffer & " OrgC = " & Chr(34) & SacarCaracterEspecialesXml(sOrgC) & Chr(34)
            MyBuffer = MyBuffer & " Denominacion1 = " & Chr(34) & SacarCaracterEspecialesXml(sDenominacion1) & Chr(34)
            MyBuffer = MyBuffer & " Ce = " & Chr(34) & SacarCaracterEspecialesXml(sCe) & Chr(34)
            MyBuffer = MyBuffer & " Denominacion_Centro = " & Chr(34) & SacarCaracterEspecialesXml(sDenominacion_Centro) & Chr(34)
            MyBuffer = MyBuffer & " GCp = " & Chr(34) & SacarCaracterEspecialesXml(sGCp) & Chr(34)
            MyBuffer = MyBuffer & " Denominacion2 = " & Chr(34) & SacarCaracterEspecialesXml(sDenominacion2) & Chr(34)
            MyBuffer = MyBuffer & " Prec_neto = " & Chr(34) & SacarCaracterEspecialesXml(sPrec_neto) & Chr(34)
            MyBuffer = MyBuffer & " Mon = " & Chr(34) & SacarCaracterEspecialesXml(sMon) & Chr(34)
            MyBuffer = MyBuffer & " Ctd_mn = " & Chr(34) & SacarCaracterEspecialesXml(sCtd_mn) & Chr(34)
            MyBuffer = MyBuffer & " Tipo_de_Co = " & Chr(34) & SacarCaracterEspecialesXml(sTipo_de_Co) & Chr(34)
            MyBuffer = MyBuffer & " Perfil_de = " & Chr(34) & SacarCaracterEspecialesXml(sPerfil_de) & Chr(34)
            MyBuffer = MyBuffer & " PzE = " & Chr(34) & SacarCaracterEspecialesXml(sPzE) & Chr(34)
            MyBuffer = MyBuffer & " Plazo_Anul = " & Chr(34) & SacarCaracterEspecialesXml(sPlazo_Anul) & Chr(34)
            MyBuffer = MyBuffer & " Valido_de = " & Chr(34) & SacarCaracterEspecialesXml(Format(Valida_de, "yyyymmdd")) & Chr(34)
            MyBuffer = MyBuffer & " Validez_a = " & Chr(34) & SacarCaracterEspecialesXml(Format(Validez_a, "yyyymmdd")) & Chr(34)
            MyBuffer = MyBuffer & " Importe = " & Chr(34) & SacarCaracterEspecialesXml(sImporte) & Chr(34)
            MyBuffer = MyBuffer & " Un = " & Chr(34) & SacarCaracterEspecialesXml(sUn) & Chr(34)
            MyBuffer = MyBuffer & " por = " & Chr(34) & SacarCaracterEspecialesXml(spor) & Chr(34)
            MyBuffer = MyBuffer & " UM = " & Chr(34) & SacarCaracterEspecialesXml(sUM) & Chr(34)
            MyBuffer = MyBuffer & " B = " & Chr(34) & SacarCaracterEspecialesXml(sB) & Chr(34)
            MyBuffer = MyBuffer & " Ruta_archivo = " & Chr(34) & SacarCaracterEspecialesXml(RutaArchivo) & Chr(34)
            
            Let MyBuffer = MyBuffer & "/>"
            
            If i = CCur(GetParametro("parxml")) Then
               Let MyBuffer = MyBuffer & "</GrabaConvenios>"
               vg_db.Execute ("sgp_iu_Convenios '" & MyBuffer & "', '" & MuestraCasino(1) & "'")
               i = 0
               '-------> Genera Ruta Compra
               Let MyBuffer = ""
               Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
               Let MyBuffer = MyBuffer & "<GrabaConvenios>"
            End If
            
'            'chequea existencia (insert o update)
'            sql = " SELECT Reg_info"
'            sql = sql & " From convenios_mvi"
'            sql = sql & " WHERE Reg_info = '" & sReg_info & "'"
'            sql = sql & " AND Proveedor = '" & sProveedor & "'"
'            sql = sql & " AND Material = '" & sMaterial & "'"
'            sql = sql & " AND Ce = '" & sCe & "'"
'
'            Set RS = vg_db.Execute(sql)
'
'            If RS.EOF Then '-------> insertar tabla convenios_mvi
'                sql = " INSERT INTO convenios_mvi"
'                sql = sql & " (Reg_info, Proveedor, Denominacion_Proveedor, Material, Denominacion_Material, OrgC, Denominacion1, Ce, Denominacion_Centro, GCp"
'                sql = sql & " ,Denominacion2, Prec_neto, Mon, Ctd_mn, Tipo_de_Co, Perfil_de, PzE, Plazo_Anul, Valido_de, Validez_a, Importe, Un, por, UM, B, Ruta_archivo)"
'                sql = sql & " VALUES ("
'
'                sql = sql & "'" & sReg_info & "'"
'                sql = sql & ",'" & sProveedor & "'"
'                sql = sql & ",'" & sDenominacion_Proveedor & "'"
'                sql = sql & ",'" & sMaterial & "'"
'                sql = sql & ",'" & sDenominacion_Material & "'"
'                sql = sql & ",'" & sOrgC & "'"
'                sql = sql & ",'" & sDenominacion & "'"
'                sql = sql & ",'" & sCe & "'"
'                sql = sql & ",'" & sDenominacion_Centro & "'"
'                sql = sql & ",'" & sGCp & "'"
'                sql = sql & ",'" & sDenominacion & "'"
'                sql = sql & ",'" & sPrec_neto & "'"
'                sql = sql & ",'" & sMon & "'"
'                sql = sql & ",'" & sCtd_mn & "'"
'                sql = sql & ",'" & sTipo_de_Co & "'"
'                sql = sql & ",'" & sPerfil_de & "'"
'                sql = sql & ",'" & sPzE & "'"
'                sql = sql & ",'" & sPlazo_Anul & "'"
'                sql = sql & ",'" & Format(Valida_de, "YYYYMMDD") & "'"
'                sql = sql & ",'" & Format(Validez_a, "YYYYMMDD") & "'"
'                sql = sql & ",'" & sImporte & "'"
'                sql = sql & ",'" & sUn & "'"
'                sql = sql & ",'" & spor & "'"
'                sql = sql & ",'" & sUM & "'"
'                sql = sql & ",'" & sB & "'"
'                sql = sql & ",'" & RutaArchivo & "'"
'
'                sql = sql & ")"
'            Else '-------> update  tabla convenios_mvi
'                sql = " Update convenios_mvi"
'                'Sql = Sql & ",Reg_info = '" & sReg_info & "'"
'                'Sql = Sql & ",Proveedor = '" & sProveedor & "'"
'                sql = sql & " SET Denominacion_Proveedor = '" & sDenominacion_Proveedor & "'"
'                'Sql = Sql & ",Material = '" & sMaterial & "'"
'                sql = sql & " ,Denominacion_Material = '" & sDenominacion_Material & "'"
'                sql = sql & " ,OrgC = '" & sOrgC & "'"
'                sql = sql & " ,Denominacion1 = '" & sDenominacion & "'"
'                'Sql = Sql & ",Ce = '" & sCe & "'"
'                sql = sql & " ,Denominacion_Centro = '" & sDenominacion_Centro & "'"
'                sql = sql & " ,GCp = '" & sGCp & "'"
'                sql = sql & " ,Denominacion2 = '" & sDenominacion & "'"
'                sql = sql & " ,Prec_neto = '" & sPrec_neto & "'"
'                sql = sql & " ,Mon = '" & sMon & "'"
'                sql = sql & " ,Ctd_mn = '" & sCtd_mn & "'"
'                sql = sql & " ,Tipo_de_Co = '" & sTipo_de_Co & "'"
'                sql = sql & " ,Perfil_de = '" & sPerfil_de & "'"
'                sql = sql & " ,PzE = '" & sPzE & "'"
'                sql = sql & " ,Plazo_Anul = '" & sPlazo_Anul & "'"
'                sql = sql & " ,Valido_de = '" & Format(Valida_de, "YYYYMMDD") & "'"
'                sql = sql & " ,Validez_a = '" & Format(Validez_a, "YYYYMMDD") & "'"
'                sql = sql & " ,Importe = '" & sImporte & "'"
'                sql = sql & " ,Un = '" & sUn & "'"
'                sql = sql & " ,por = '" & spor & "'"
'                sql = sql & " ,UM = '" & sUM & "'"
'                sql = sql & " ,B = '" & sB & "'"
'                sql = sql & " ,Ruta_archivo = '" & RutaArchivo & "'"
'
'                sql = sql & " WHERE Reg_info = '" & sReg_info & "'"
'                sql = sql & " AND Proveedor = '" & sProveedor & "'"
'                sql = sql & " AND Material = '" & sMaterial & "'"
'                sql = sql & " AND Ce = '" & sCe & "'"
'            End If
'
'             vg_db.Execute (sql)
'             RS.Close: Set RS = Nothing
        End If
            
''      End If
      RsExcel.MoveNext: j = j + 1: i = i + 1
    Loop
    Let MyBuffer = MyBuffer & "</GrabaConvenios>"
    vg_db.Execute ("sgp_iu_Convenios '" & MyBuffer & "', '" & MuestraCasino(1) & "'")

End If
PB.Min = 0: PB.Value = 0: PB.Visible = False
RsExcel.Close: Set RsExcel = Nothing
    
MsgBox "Proceso finalizado con exito", vbInformation + vbOKOnly, Me.Caption
Exit Sub

errSub:
     MsgBox "Ha ocurrido un error en la Importación : " & Err.Description & " " & Err.Number, vbCritical, "Error"
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
    Dim mensaje As String
    Dim swEntraError As Boolean
    Dim FechaHoy As String
    Dim ContReg As Long
    
    FechaHoy = CStr(Date)
    
On Error GoTo errSub

    Dim List() As String
    Dim listcount As Integer
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
    Dim validadorgrilla As Boolean
    Dim id_carga As Long
    Dim RS As New ADODB.Recordset
    swEntraError = False
    
    '-------> Validar hoja este seleccionada
    If CboHojaExcel.ListIndex = -1 Then
       MsgBox "Debe seleccionar un hoja excel...", vbCritical, Me.Caption
       Exit Sub
    End If
    
    '-------> Validar que este selecionado un item en la grilla
    id_carga = 0
    If vaSpread1.MaxRows > 0 Then
       validadorgrilla = False
       For j = 1 To vaSpread1.MaxRows
           vaSpread1.Row = j
           vaSpread1.Col = 1
           If vaSpread1.text = "1" Then
              validadorgrilla = True
              vaSpread1.Col = 2
              id_carga = Val(vaSpread1.text)
              Exit For
           End If
       Next j
       If Not validadorgrilla Then
          MsgBox "Debe seleccionar un item en la grilla...", vbCritical, Me.Caption
          Exit Sub
       End If
    End If
    
    Dim RutaArchivo As String
    
    Me.txterrores.text = ""
    
    NombreArchivoExcel = Me.txtRutaArchivo.text
    
    txtRutaArchivo = NombreArchivoExcel
    RutaArchivo = txtRutaArchivo
    
    fromRight = InStrRev(CD.Filename, "\", , vbTextCompare)
    
    If fromRight > 1 Then
       varRuta = Left(CD.Filename, fromRight)
    End If
    
    Dim RsExcel As ADODB.Recordset
    Dim sconn As String
    Set RsExcel = New ADODB.Recordset
    
    
    
    RsExcel.CursorLocation = adUseClient
    RsExcel.CursorType = adOpenKeyset
    RsExcel.LockType = adLockBatchOptimistic
    
    
    sconn = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & NombreArchivoExcel
    


Dim cn As ADODB.Connection
Set cn = New ADODB.Connection
With cn

.Provider = "Microsoft.Jet.OLEDB.4.0"
.ConnectionString = "Data Source=" & NombreArchivoExcel & ";" & _
"Extended Properties='Excel 8.0;HDR=NO;IMEX=1'"
.Open
End With

    'linea para excel 2007 o superior. Necesita bajar drivers extra.
    'sconn = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & NombreArchivoExcel
    
    wvarHoja = CboHojaExcel.text '"RUTAS tf DICIEMBRE"
    
    strSQL = "SELECT * FROM [" & wvarHoja & "$]"
        
    RsExcel.Open strSQL, cn
    
    'CuentaError = 0
    
    j = 1
    
    mensaje = ""
    If Not RsExcel.EOF Then
    
    RsExcel.MoveNext
    
    'validacion de los valores en los campos del archivo en excel
        Do While Not RsExcel.EOF
        
        
        'sID_ruta_compra = RsExcel(0)
        sFecha_despacho = IIf(IsNull(RsExcel(0)), "", RsExcel(0))
        sID_centro_de_costo = IIf(IsNull(RsExcel(1)), "", RsExcel(1))
        sFamilia_producto = IIf(IsNull(RsExcel(2)), "", RsExcel(2))
        sID_proveedor = IIf(IsNull(RsExcel(3)), "", RsExcel(3))
        sSucursal = IIf(IsNull(RsExcel(4)), "", RsExcel(4))
        sSigla_de_ruta = IIf(IsNull(RsExcel(5)), "", RsExcel(5))
        sDescripcion_sigla = IIf(IsNull(RsExcel(6)), "", RsExcel(6))
        sObservaciones = IIf(IsNull(RsExcel(7)), "", RsExcel(7))


        'validaciones obligatorios
        Dim FechaDespacho As String
        FechaDespacho = sFecha_despacho
        
        'limpieza de caracteres basura
        FechaDespacho = Replace(FechaDespacho, ".", "/")
        FechaDespacho = Replace(FechaDespacho, "-", "/")
        
        If Not IsDate(CDate(FechaDespacho)) Then
            mensaje = mensaje & "Error en la fila N° " & CStr(j) & ", Fecha_despacho formato erroneo." & vbNewLine
            swEntraError = True
        End If
        
        If sFecha_despacho = "" Then
            mensaje = mensaje & "Error en la fila N° " & CStr(j) & ", Fecha_despacho obligatorio." & vbNewLine
            swEntraError = True
        ElseIf sID_centro_de_costo = "" Then
            mensaje = mensaje & "Error en la fila N° " & CStr(j) & ", ID_centro_de_costo obligatorio." & vbNewLine
            swEntraError = True
        ElseIf sFamilia_producto = "" Then
            mensaje = mensaje & "Error en la fila N° " & CStr(j) & ", Familia_producto obligatorio." & vbNewLine
            swEntraError = True
        ElseIf sID_proveedor = "" Then
            mensaje = mensaje & "Error en la fila N° " & CStr(j) & ", ID_proveedor obligatorio." & vbNewLine
            swEntraError = True
        ElseIf sSucursal = "" Then
            mensaje = mensaje & "Error en la fila N° " & CStr(j) & ", Sucursal obligatorio." & vbNewLine
            swEntraError = True
        ElseIf sSigla_de_ruta = "" Then
            mensaje = mensaje & "Error en la fila N° " & CStr(j) & ", Sigla_de_ruta obligatorio." & vbNewLine
            swEntraError = True
        ElseIf sSigla_de_ruta = "" Then
            mensaje = mensaje & "Error en la fila N° " & CStr(j) & ", Sigla_de_ruta obligatorio." & vbNewLine
            swEntraError = True
         ElseIf sObservaciones = "" Then
            mensaje = mensaje & "Error en la fila N° " & CStr(j) & ", Observaciones obligatorio." & vbNewLine
            swEntraError = True
        End If
        
        j = j + 1
        RsExcel.MoveNext
    Loop
    
    
    If swEntraError = True Then
        MsgBox "Errores encontrados en el archivo, favor mirar texto de errores , carga cancelada", vbCritical, Me.Caption
        Me.txterrores = mensaje
        Set RsExcel = Nothing
        Exit Sub
    End If
    'fin validacion de los valores en los campos del archivo en excel
    
    j = 1
    i = 1
    RsExcel.MoveFirst
    
    Dim MyBuffer As Variant
    Dim Fecha_despacho As String
    Dim ID_centro_de_costo As String
    Dim Familia_producto As String
    Dim ID_proveedor As String
    Dim Sucursal As String
    Dim Sigla_de_ruta As String
    Dim Descripcion_sigla As String
    Dim Observaciones As String
    Dim Ruta_archivo As String
    Dim Estado As Boolean
    '-------> Genera Ruta Compra
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<GrabaRutaCompra>"
    
    PB.Min = 0: PB.Value = 0: PB.Visible = True
    Estado = True
    ContReg = RsExcel.RecordCount
    Do While Not RsExcel.EOF
        PB.Value = Val((j / ContReg) * 100)
        If IsDate(RsExcel.Fields(0).Value) Then
           Let MyBuffer = MyBuffer & " <RutaCompra"
           MyBuffer = MyBuffer & " Op = " & Chr(34) & 0 & Chr(34)

            Fecha_despacho = SacarCaracterEspecialesXml(Format(RsExcel(0), "YYYYMMDD"))
            
            ID_centro_de_costo = SacarCaracterEspecialesXml(RsExcel(1))
            
            Familia_producto = SacarCaracterEspecialesXml(RsExcel(2))
            
            ID_proveedor = SacarCaracterEspecialesXml(RsExcel(3))
            
            Sucursal = SacarCaracterEspecialesXml(RsExcel(4))
            
            Sigla_de_ruta = SacarCaracterEspecialesXml(RsExcel(5))
            
            Descripcion_sigla = SacarCaracterEspecialesXml(RsExcel(6))
            
            Observaciones = SacarCaracterEspecialesXml(RsExcel(7))
             
            Ruta_archivo = ""
            
            MyBuffer = MyBuffer & " Fecha_despacho  = " & Chr(34) & Fecha_despacho & Chr(34)
            MyBuffer = MyBuffer & " ID_centro_de_costo = " & Chr(34) & ID_centro_de_costo & Chr(34)
            MyBuffer = MyBuffer & " Familia_producto  = " & Chr(34) & Familia_producto & Chr(34)
            MyBuffer = MyBuffer & " ID_proveedor  = " & Chr(34) & ID_proveedor & Chr(34)
            MyBuffer = MyBuffer & " Sucursal  = " & Chr(34) & Sucursal & Chr(34)
            MyBuffer = MyBuffer & " Sigla_de_ruta  = " & Chr(34) & Sigla_de_ruta & Chr(34)
            MyBuffer = MyBuffer & " Descripcion_sigla  = " & Chr(34) & Descripcion_sigla & Chr(34)
            MyBuffer = MyBuffer & " Observaciones  = " & Chr(34) & Observaciones & Chr(34)
            MyBuffer = MyBuffer & " Ruta_archivo  = " & Chr(34) & SacarCaracterEspecialesXml(RutaArchivo) & Chr(34)
            Let MyBuffer = MyBuffer & "/>"
            
            If i = CCur(GetParametro("parxml")) Then
               Let MyBuffer = MyBuffer & "</GrabaRutaCompra>"
               vg_db.Execute ("sgp_id_RutaCompra '" & MyBuffer & "', '" & MuestraCasino(1) & "', " & id_carga & ", '" & vg_NUsr & "', '" & IIf(Estado = True, "1", "0") & "'")
               i = 0
               Estado = False
               '-------> Genera Ruta Compra
               Let MyBuffer = ""
               Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
               Let MyBuffer = MyBuffer & "<GrabaRutaCompra>"
            End If
            j = j + 1
            i = i + 1
        End If
        RsExcel.MoveNext
    Loop
    Let MyBuffer = MyBuffer & "</GrabaRutaCompra>"
    vg_db.Execute ("sgp_id_RutaCompra '" & MyBuffer & "', '" & MuestraCasino(1) & "', " & id_carga & ", '" & vg_NUsr & "', '" & IIf(Estado = True, "1", "0") & "'")
    
    End If
    
    RsExcel.Close: Set RsExcel = Nothing
   
    PB.Min = 0: PB.Value = 0: PB.Visible = False
    
  MsgBox "Proceso finalizado con exito", vbInformation + vbOKOnly, Me.Caption
  CargaGrillaRutaCompras
    Exit Sub

errSub:
    PB.Min = 0: PB.Value = 0: PB.Visible = False
     MsgBox "Ha ocurrido un error en la Importación : " & Err.Description & " " & Err.Number, vbCritical, "Error"
     
     'Me.Enabled = True

End Sub


Private Sub btnExplorar_Click()
'validaciones
If Me.txtRutaArchivo = "" Then
    MsgBox "Debe seleccionar un tipo de archivo", vbExclamation, Me.Caption
    Exit Sub
End If

If optConvenio.Value = True Then
'    Call ImporConvenio
    Call ImpExcel_Convenios
Else
    Call ImpExcel
End If

End Sub

Private Sub CmdSalir_Click()
    Me.Hide
    Unload Me
End Sub
Function File_Extension(Path As String, Caracter As String) As String
    Dim ret As String
    If Caracter = "." And InStr(Path, Caracter) = 0 Then Exit Function
    ret = Right(Path, Len(Path) - InStrRev(Path, Caracter))
      
    ' -- Retorna el valor
    File_Extension = ret
End Function

Private Sub Command1_Click()
CboHojaExcel.Enabled = False
CboHojaExcel.Clear
btnExplorar.Enabled = False
Dialogo MVI_ImpArchivos, TheFile, 1, TipoArchivo(1)
If TheFile.Success Then
   Screen.MousePointer = 11
   Me.txtRutaArchivo.text = TheFile.Filename
   Dim oExcel As Object, i As Integer
   Set oExcel = GetObject(TheFile.Filename)
   For i = 1 To oExcel.Sheets.count
      CboHojaExcel.AddItem oExcel.Sheets(i).Name
   Next i
   oExcel.Application.Quit
   Set oExcel = Nothing
   Screen.MousePointer = 0
   CboHojaExcel.Enabled = True
Else
   Me.txtRutaArchivo.text = ""
End If
btnExplorar.Enabled = True
End Sub

Private Sub Form_Load()
fg_centra Me
Me.optExcel.Value = True
Me.optConvenio.Value = False
End Sub

Private Sub optConvenio_Click()
vaSpread1.MaxRows = 0
vaSpread1.Enabled = False
txtRutaArchivo.text = ""
CboHojaExcel.Clear
End Sub

'lo que se debe hacer es lo sgte:
'1) para ambos archivos se debe crear una tabla que extra que vuelque los sgtes. datos
'y tenga los sgtes. campos: nombre archivo, id insertado y fecha

'2)ingresar en la tabla ppal. (donde almacena las filas de los archivos importados) almacenar en un campo
'la ruta del archivo importado.

'3)para el archivo texto este es el cjto de llaves que debe hacer UPDATE:
'En el caso de convenios, se hace el update según las siguientes llaves (Info disponible en TONERConMarcas.txt):
'i.  Reg.info ,  Proveedor , Material y   Centro de Costo (Ce.)

'4)realizar el chequeo que todos los campos anteriores existan, en el caso de no ser asi, sale con error por pantalla.

'5)el volcado de los errores de estos archivos se debe hacer en un archivo log que sea abierto por un Shell.

'6)definir los campos sometidos a validacion que en el caso que no existan debe salir de la importacion por error en el
'archivo.

'7)para los archivos excel se debe preguntar una fecha o mostrar un control de fecha para que la ingresen y
'realicen el borrado en relacion a la fecha seleccionada para el archivo excel.

'8)si existe un error en la importacion de archivos se debe echar atras la importacion completa del archivo y mostrarlo
' el error por pantalla (agregar rutina de validacion de error, ON ERROR GOTO ERR_HANDLER).
'DONE

Private Sub optExcel_Click()
Label1(1).Visible = True
CboHojaExcel.Clear
CboHojaExcel.Visible = True
txtRutaArchivo.text = ""
CargaGrillaRutaCompras
End Sub

Sub CargaGrillaRutaCompras()
Dim RS As New ADODB.Recordset
est = True
vaSpread1.MaxRows = 0
vaSpread1.Enabled = True
'-------> Cargar grilla
Set RS = vg_db.Execute("select convert(varchar(1000),id_carga) as id_carga, fecha, usuario from Carga_Ruta_Compra order by id_carga desc")
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   vaSpread1.text = "0"
   vaSpread1.Col = 2
   vaSpread1.text = RS(0)
   vaSpread1.Col = 3
   vaSpread1.text = RS(1)
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
est = False
End Sub
Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If ButtonDown = 0 Then Exit Sub
If est Then Exit Sub
Dim i As Long
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1
    If i <> Row Then
       If vaSpread1.text = "1" Then
          est = True
          vaSpread1.text = "0"
          est = False
        End If
    End If
Next i
End Sub
