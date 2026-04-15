VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form M_Calendario_fechas_despachos 
   Caption         =   "Calendario de Fechas de Despachos de los Cecos y Proveedores"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   17040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   17040
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   5640
      TabIndex        =   6
      Top             =   8520
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   3960
      TabIndex        =   5
      Top             =   8520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   2280
      TabIndex        =   4
      Top             =   8520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   960
      TabIndex        =   3
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mostrar en Grilla"
      Height          =   1695
      Left            =   12120
      TabIndex        =   10
      Top             =   8640
      Width           =   4815
      Begin VB.CheckBox Check1 
         Caption         =   "Sitios Simap"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sitios No Simap"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Sitios FM"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1320
         Value           =   1  'Checked
         Width           =   3375
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   111411201
      CurrentDate     =   41744
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   111411201
      CurrentDate     =   41744
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   7230
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   16710
      _Version        =   393216
      _ExtentX        =   29475
      _ExtentY        =   12753
      _StockProps     =   64
      AutoClipboard   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   0
      MaxRows         =   0
      SpreadDesigner  =   "M_Calendario_fechas_despachos.frx":0000
      StartingColNumber=   2
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   17040
      _ExtentX        =   30057
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   210
      Left            =   315
      TabIndex        =   14
      Top             =   9870
      Visible         =   0   'False
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   0
      Top             =   10560
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   1440
      Width           =   300
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Hasta"
      Height          =   315
      Left            =   3360
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Desde"
      Height          =   435
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "M_Calendario_fechas_despachos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Dim existeActualizado   As Integer
Public columna_anterior As Integer
Public fila_anterior    As Integer
Public fila_actual      As Integer
Public columna_actual   As Integer
Public fecha_anterior   As String
Public fecha_actual     As String
Public Ceco             As String
Public rutproveedor     As String
Public buff             As String
Public collec           As String
Public AccMod           As Boolean
Public AccEli           As Boolean
Dim buffact             As String
Dim collecact           As String
Dim modulo              As String
Dim buffantes           As String
Dim buffactual          As String

Private Sub Check1_Click()

Call FormatearGrilla 'Lee los cecos de los Proveedores

End Sub

Private Sub Check2_Click()

Call FormatearGrilla 'Lee los cecos de los Proveedores

End Sub

Private Sub Check3_Click()

Call FormatearGrilla 'Lee los cecos de los Proveedores

End Sub

' -------------------------------------------------------------------------------------------
' \\ -- Función para crear un nuevo libro con el contenido del Grid
' -------------------------------------------------------------------------------------------
Public Function Exportar_Excel(ByVal sOutputPath As String, ByRef Grid As vaSpread) As Boolean
  
On Error GoTo Error_Handler

If sOutputPath <> "" Then

  
'    Dim o_Excel     As excel.Application  'Object
'    Dim o_Libro     As excel.Workbook 'Object
'    Dim o_Hoja      As New excel.Worksheet

    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    
    Dim F_e       As Long
    Dim F_g       As Long
    Dim fila1       As Long
    Dim Columna     As Long
    Dim FILA2 As Long
       
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
    
    ' -- Bucle para Exportar cabecera
    
    With Grid
            
         F_e = 1
         .Row = SpreadHeader
         
         For Columna = 1 To .maxcols
             
             .Col = Columna
             
             If .text <> "" Then
                
                o_Hoja.Range(o_Hoja.Cells(F_e, Columna), o_Hoja.Cells(F_e, Columna + 6)).Merge
             
             End If
             
             o_Hoja.Cells(F_e, Columna).Value = .text
             o_Hoja.Cells.Cells(F_e, Columna).Interior.Color = RGB(237, 237, 237)
             o_Hoja.Cells.Cells(F_e, Columna).HorizontalAlignment = xlCenter
             o_Hoja.Cells.Font.Bold = True
         
         Next
            
         F_e = 2
         .Row = SpreadHeader + 1
         For Columna = 1 To .maxcols
             
             .Col = Columna
             o_Hoja.Cells(F_e, Columna).Value = .text
             o_Hoja.Cells.Cells(F_e, Columna).Interior.Color = RGB(237, 237, 237)
             o_Hoja.Cells.Cells(F_e, Columna).HorizontalAlignment = xlCenter
             o_Hoja.Cells.Font.Bold = True
         
         Next
            
    End With
       
   ProgressBar1.Scrolling = ccScrollingSmooth
   ProgressBar1.Max = Grid.MaxRows
   ProgressBar1.Visible = True
   ProgressBar1.Value = 0
   
   ' -- Bucle para Exportar los datos
   With Grid
         F_e = 3
         For F_g = 2 To .MaxRows
            .Row = F_g
            If .RowHidden = False Then
               
               For Columna = 1 To .maxcols
                  .Col = Columna
                  
                  If Columna >= 1 And Columna < 7 Then
                     
                     o_Hoja.Cells.Cells(F_e, Columna).Interior.Color = RGB(203, 255, 209)
                     o_Hoja.Cells.Cells(F_e, Columna).HorizontalAlignment = xlLeft
                  
                  Else
                     o_Hoja.Cells.Cells(F_e, Columna).Interior.Color = RGB(244, 247, 170)
                     o_Hoja.Cells.Cells(F_e, Columna).HorizontalAlignment = xlCenter
                  
                  End If
                  o_Hoja.Cells(F_e, Columna).Value = .text
               
               Next
               
               F_e = F_e + 1
            
            End If
            ProgressBar1.Value = ProgressBar1.Value + 1
'            lbl_proceso.Caption = CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %"
         Next

   End With
   o_Libro.Close True, sOutputPath
    
   Dim XL As New excel.Application 'Crea el objeto excel
   XL.Workbooks.Open sOutputPath, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
   XL.Visible = True
   XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing

    Exportar_Excel = True
    
    ProgressBar1.Visible = False
    ProgressBar1.Value = 0
End If

Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ProgressBar1.Visible = False
    ProgressBar1.Value = 0
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing

    MsgBox Err.Description, vbCritical, MsgTitulo
    
End Function


Private Sub Form_Load()
 
On Error GoTo Man_Error
 
Dim RS    As New ADODB.Recordset
Dim mes   As String
Dim ano   As String
Dim Fecha As String
Dim Sql   As String

MsgTitulo = "Calendario Despacho"
 
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = "sgpadm_Sel_RecuperaFechaServidor "
Set RS = vg_db.Execute(Sql)
If Not RS.EOF Then
  
   Fecha = RS(0)
 
End If
RS.Close
 
mes = Mid(Fecha, 4, 2)
ano = Mid(Fecha, 7, 4)
 
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = "sgpadm_sel_FechaDespachoIncioyFinaldeMES "
Sql = Sql & " '" & mes & "',"
Sql = Sql & "'" & ano & "'"
Set RS = vg_db.Execute(Sql)
If Not RS.EOF Then
   
   DTPicker1 = RS(1)
   DTPicker2 = RS(3)
   Fecha = RS(0)
  
End If
RS.Close
 
AccMod = False
AccEli = False
 
vaSpread1.Visible = True
Toolbar1.ImageList = Partida.IL1
Me.HelpContextID = 1196001
  
Set BtnX = Toolbar1.Buttons.Add(, "A_Retrocede", , tbrDefault, "A_Retrocede"): BtnX.Visible = False: BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, False): BtnX.ToolTipText = "Retrocede Ruta"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Avanza", , tbrDefault, "A_Avanza"): BtnX.Visible = False: BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, False): BtnX.ToolTipText = "Avanza Ruta"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = False: BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 3, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Excel "
Set BtnX = Toolbar1.Buttons.Add(, "A_Filtro", , tbrDefault, "A_Filtro"): BtnX.Visible = True: BtnX.ToolTipText = "Filtrar"
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
 
AccMod = Toolbar1.Buttons(1).Enabled
AccEli = Toolbar1.Buttons(5).Enabled
 
Me.HelpContextID = 1196001
 
' Call FormatearGrilla
 
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Sub FormatearGrilla()
 
On Error GoTo Man_Error

Dim RS           As New ADODB.Recordset
Dim NombreSemana As String
Dim MaxColumna   As Long
Dim i            As Long
Dim filtro       As String

Dim diaSemana As Integer
Dim Dias(1 To 7) As String

Dias(1) = "Domingo"
Dias(2) = "Lunes"
Dias(3) = "Martes"
Dias(4) = "Miércoles"
Dias(5) = "Jueves"
Dias(6) = "Viernes"
Dias(7) = "Sábado"


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

If Format(DTPicker1, "YYYYMMDD") > Format(DTPicker2, "YYYYMMDD") Then
   
   MsgBox "Le fecha hasta  no puede ser menor que la hasta", vbExclamation
   Exit Sub

End If

With vaSpread1

    .Visible = False
    MaxColumna = 0
    .MaxRows = 0
    
    '-------> determinar la cuando días entre la fecha desde - hasta
   MaxColumna = DateDiff("d", CDate(fg_Ctod1(Format(DTPicker1, "YYYYMMDD"))), CDate(fg_Ctod1(Format(DTPicker2, "YYYYMMDD")))) + 1
    
    '------- Defenir vector costo encabezado
    
    .maxcols = MaxColumna + 7

    '-------> definir color a toda la grilla
    .Row = -1
    .Col = -1
    .BackColor = Shape1(0).FillColor  'Amarillo
    .Lock = True
    .Row = SpreadHeader: .Col = -1
    .Font.Bold = True
    
    '-------> definir color a las tres primera columma
    For i = 1 To 6
        
        .Row = -1
        .Col = i
        .Font.Size = 8
        .BackColor = Shape1(2).FillColor 'Verde
        .Lock = True
        
    Next i
    
    '-------> Set up column headers
    .ColHeaderRows = 2
    .ShadowColor = &H8000000F
    
    '------> mover columna 1
    .Col = 1
    .ColsFrozen = 1
    .VisibleCols = 1
    .ColWidth(1) = 8
    .RowHeight(SpreadHeader + 1) = 25
    .Row = SpreadHeader + 1
    .TypeHAlign = TypeHAlignLeft
    .text = "Org. Compras"
    .Lock = True
    
    '------> mover columna 2
    .Col = 2
    .ColsFrozen = 1
    .VisibleCols = 1
    .ColWidth(1) = 8
    .RowHeight(SpreadHeader + 1) = 25
    .Row = SpreadHeader + 1
    .TypeHAlign = TypeHAlignLeft
    .text = "Ceco"
    .Lock = True
    
    .Col = 3
    .ColsFrozen = 2
    .VisibleCols = 1
    .ColWidth(2) = 15
    .Row = SpreadHeader + 1
    .TypeHAlign = TypeHAlignLeft
    .text = "Descripcion"
    .Lock = True
    
    '------> mover columna 4
    .Col = 4
    .ColsFrozen = 3
    .VisibleCols = 1
    .ColWidth(3) = 8
    .Row = SpreadHeader + 1
    .TypeHAlign = TypeHAlignLeft
    .text = "Rut"
    .Lock = True
    
    .Col = 5
    .ColsFrozen = 4
    .VisibleCols = 1
    .ColWidth(4) = 15
    .Row = SpreadHeader + 1
    .TypeHAlign = TypeHAlignLeft
    .text = "Proveedor"
    .Lock = True
    
    '------> mover columna 6
    .Col = 6
    .ColsFrozen = 5
    .VisibleCols = 1
    .ColWidth(5) = 8
    .Row = SpreadHeader + 1
    .TypeHAlign = TypeHAlignLeft
    .text = "Cross Docking"
    .Lock = True
    
    '-------> Formatear días
    Dim j As Long
    Dim X As Long
    Dim Fecha As String
    Dim numdia As Long
    Dim Anio As Long
    Dim Fd As Date
    Dim Fh As Date
    Dim columafecha As Integer
    Dim pos_x       As Long
    Dim mes_int     As Long
    mes_int = 0
    Dim mes_actual As Integer
    mes_actual = Month(Fd)

    'recorrer fechas seleccionadas
    
    Fd = Me.DTPicker1.Value
    Fh = Me.DTPicker2.Value
    'fh = fh + 1
    j = 7
    
    For Fd = Me.DTPicker1.Value To Fh
      
      .Col = j
      
      mes_actual = Month(Fd)
      If mes_actual <> mes_int Then
        
         numdia = Day(dEoM(Fd))
         mes_int = mes_actual
        
        .AddCellSpan j, SpreadHeader, numdia, 1
                
        .Row = SpreadHeader
        .text = IIf(mes_int = 1, "Enero", _
                IIf(mes_int = 2, "Febrero", _
                IIf(mes_int = 3, "Marzo", _
                IIf(mes_int = 4, "Abril", _
                IIf(mes_int = 5, "Mayo", _
                IIf(mes_int = 6, "Junio", _
                IIf(mes_int = 7, "Julio", _
                IIf(mes_int = 8, "Agosto", _
                IIf(mes_int = 9, "Septiembre", _
                IIf(mes_int = 10, "Octubre", _
                IIf(mes_int = 11, "Noviembre", "Diciembre")))))))))))
      
      End If
      
      .TypeHAlign = TypeHAlignCenter
      .Row = SpreadHeader + 1
      .Font.Size = 1
      .ColWidth(j) = 3
      diaSemana = Weekday(Fd, vbSunday) ' 1 = Domingo, 7 = Sábado
      dia = UCase(Left(Dias(diaSemana), 1)) & vbCrLf & Day(Fd)


'       dia = UCase(Mid(WeekdayName(DatePart("w", Fd, vbUseSystemDayOfWeek)), 1, 1)) & vbCrLf & Day(Fd)
      .text = dia & " " '& vbCrLf & Day(fd)
      .Lock = True
      
       'valor fecha, para buscar posteriormente y asignar X
       
       .MaxRows = 1
       .Row = 1
       .Font.Size = 1
       .ColWidth(j) = 3
       .text = Fd
       .RowHidden = True
    
      ' Colorea los Fine de Semana
      If Mid(dia, 1, 1) = "S" Or Mid(dia, 1, 1) = "D" Then
            
            .Col = j 'Posicion Fecha de Despacho
            .Row = -1
            .BackColor = RGB(208, 207, 71)
            .Lock = True
            
      End If
      
           
      j = j + 1
   
   Next Fd
     
    '-------> ocultar ultima columna
    vaSpread1.Col = vaSpread1.maxcols
    vaSpread1.ColHidden = True
  
  'Busca los cecos que tienen fecha de Ruta de Despacho
    
    Dim fecha_despacho As String
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Sql = " sgpadm_sel_rutadespacho_fecha_V01 "
    Sql = Sql & " '" & Format(DTPicker1, "YYYYMMDD") & "',"
    Sql = Sql & " '" & Format(DTPicker2, "YYYYMMDD") & "',"
    Sql = Sql & " '" & filtro & "'"
    
    Set RS = vg_db.Execute(Sql)

    If Not RS.EOF Then
      
      Do While Not RS.EOF

        If RS("marca") = True Then
  
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          
          .Col = 1 ' Org. Compras
          .text = RS("Id_OrgCompras")
          
          .Col = 2 ' Ceco
          .text = RS("cli_codigo") 'Val(RS("cli_codigo"))
  
          .Col = 3 ' Descripcion
          .text = RS("cli_nombre")
  
          .Col = 4 ' Rut Proveedor
          .text = IIf(IsNull(RS("prv_codigo")), " ", RS("prv_codigo"))
   
          .Col = 5 'Proveedor
          .text = IIf(IsNull(RS("prv_nombre")), " ", RS("prv_nombre"))
          
          .Col = 6 'Crossdocking
          .text = IIf(IsNull(RS("crossdocking")), " ", RS("crossdocking"))
          .TypeHAlign = TypeHAlignCenter
          
          fecha_despacho = Mid(RS("Fecha_despacho"), 7, 2) + "/" + Mid(RS("Fecha_despacho"), 5, 2) + "/" + Mid(RS("Fecha_despacho"), 1, 4) 'dd/mm/yyyy
  
          pos_x = .SearchRow(1, 7, .maxcols, fecha_despacho, SearchFlagsValue)
          
          If pos_x > 0 Then
            
            .Col = pos_x 'Posicion Fecha de Despacho
            .text = "X"
            .TypeHAlign = TypeHAlignCenter
            .Lock = True
          
          End If
        
        End If
        
        fecha_despacho = Mid(RS("Fecha_despacho"), 7, 2) + "/" + Mid(RS("Fecha_despacho"), 5, 2) + "/" + Mid(RS("Fecha_despacho"), 1, 4) 'dd/mm/yyyy

        pos_x = .SearchRow(1, 7, .maxcols, fecha_despacho, SearchFlagsValue)
        
        If pos_x > 0 Then
          
          .Col = pos_x 'Posicion Fecha de Despacho
          .text = "X"
          .TypeHAlign = TypeHAlignCenter
          .Lock = True
        
        End If

        .Col = .maxcols
        .text = 0
        
        RS.MoveNext
      Loop
    End If
    RS.Close
    
    'colorear feriados
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Sql = " sgpadm_sel_feriados "
    Sql = Sql & " '" & Format(DTPicker1, "YYYYMMDD") & "',"
    Sql = Sql & " '" & Format(DTPicker2, "YYYYMMDD") & "'"
    Set RS = vg_db.Execute(Sql)
    
    If Not RS.EOF Then
      
      Do While Not RS.EOF
        
        fecha_feriado = RS(0) 'dd/mm/yyyy
  
        pos_x = vaSpread1.SearchRow(1, 7, .maxcols, fecha_feriado, SearchFlagsValue)
          
          If pos_x > 0 Then
            
            .Col = pos_x 'Posicion Fecha de Despacho
            .Row = -1
            .BackColor = vbRed
            .Lock = True
            
          End If
        
        RS.MoveNext
      
      Loop
    End If
    RS.Close
  
    .Visible = True
    
End With
  
'  ret = vaSpread1.ExportToXMLBuffer("ParamCeco", collec, buff, ExportToXMLFormattedData, "")
  
Toolbar1.Buttons(1).Enabled = False
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
vaSpread1.SetActiveCell 1, 2

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

If Index = 1 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(4).text = ""

ElseIf Index = 2 Then
   
   Text1(1).text = ""
   Text1(3).text = ""
   Text1(4).text = ""

ElseIf Index = 3 Then
   
   Text1(1).text = ""
   Text1(2).text = ""
   Text1(4).text = ""

ElseIf Index = 4 Then

   Text1(1).text = ""
   Text1(2).text = ""
   Text1(3).text = ""

End If

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = vaSpread1.maxcols
    vaSpread1.text = 0

Next

Select Case Index

Case 1, 2, 3, 4
    
    vaSpread1.Visible = False
    
    If Trim(Text1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread1.Col = 1
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" And i <> 1 Then
              
              vaSpread1.Col = vaSpread1.maxcols
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = vaSpread1.maxcols
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = vaSpread1.maxcols
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = vaSpread1.maxcols
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 1
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 And i <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = vaSpread1.maxcols
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
           If i <> 1 And vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = vaSpread1.maxcols
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

Case 1
   
   'registrar Log sistema grabar
   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), Me.HelpContextID, "", "", "")
   
   Call Retroceder_en_la_Grilla 'Retrocede en la Grilla
    
   'registrar Log sistema grabar
   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Grabar"), Me.HelpContextID, "", "", "")
    
Case 3
  
   'registrar Log sistema grabar
   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), Me.HelpContextID, "", "", "")
   
   Call Avanzar_en_la_Grilla 'Avanza en la Grilla
   
   'registrar Log sistema grabar
   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Grabar"), Me.HelpContextID, "", "", "")
   
Case 5
   
   'registrar Log sistema grabar
   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Eliminar"), Me.HelpContextID, "", "", "")
   
   Call Eliminar_Grilla 'Elimina en La Grilla
   
   'registrar Log sistema eliminar
   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminar"), Me.HelpContextID, "", "", "")

Case 6
   
   Call lleva_excel 'LLeva a Excel
 
   'registrar Log sistema grabar
   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Exporta_Excel"), Me.HelpContextID, "", "", "")
 
Case 7
   
   FormatearGrilla 'Filtro para el Formateo
   
   'registrar Log sistema grabar
   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Filtrar"), Me.HelpContextID, "", "", "")
   
Case 8
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub grabarLogo()

On Error GoTo Man_Error

'Dim xml_ant As New DOMDocument40
'Dim xml_act As New DOMDocument40
Dim xml_ant As New DOMDocument60
Dim xml_act As New DOMDocument60


' modulo = "Calendario de Fechas de Despachos"
' ret = vaSpread1.ExportToXMLBuffer("ParamCeco", collecact, buffact, ExportToXMLFormattedData, "")
'
'  buffantes = buff
'  buffactual = buffact
'
'  xml_ant.LoadXml (buffantes)
'  xml_act.LoadXml (buffact)
'
'
'    sql = " sgpadm_Ins_Log_Fecha_Despacho  "
'    sql = sql & "'" & UCase(vg_NUsr) & "',"
'    sql = sql & "'" & modulo & "',"
'    sql = sql & "'" & xml_ant.documentElement.XML & "',"
'    sql = sql & "'" & xml_act.documentElement.XML & "'"
'
'    'Debug.Print sql
'    Set RS = vg_db.Execute(sql)
    
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub lleva_excel()

On Error GoTo Man_Error

Dim r       As Boolean
Dim archivo As String
Dim ms      As VbMsgBoxResult

r = False
Me.CommonDialog1.ShowSave
Me.CommonDialog1.Filter = "*.xls,*.xlsx"
archivo = Me.CommonDialog1.FileName
Me.MousePointer = 11

r = vaSpread1.ExportToExcel(archivo, "Hoja1", dir_trabajo & "LOGFILE.TXT")
'Display result to user based on T/F value of x
If r = True Then
    
   Dim F_e     As Long
   Dim F_g     As Long
   Dim fila1   As Long
   Dim Columna As Long
   Dim FILA2   As Long
       
   ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
   Dim XL As excel.Application
   Set XL = CreateObject("Excel.application")
   
   XL.Workbooks.Open FileName:=archivo
   '-------> Desactivar proteción
   XL.Cells.Select
   XL.ActiveSheet.Unprotect
   
   '-------> Borrar dos filas final
    XL.Rows("64000:64001").Select
    XL.Selection.Delete Shift:=xlUp
    
   '-------> Insertar dos filas inicio
    XL.Rows("1:2").Select '------> Insert Fila
    XL.Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    'crear encabezado
    With vaSpread1
            
         F_e = 1
         .Row = SpreadHeader
         
         For Columna = 1 To .maxcols
             
             .Col = Columna
             
             If .text <> "" Then
                
                XL.Range(XL.Cells(F_e, Columna), XL.Cells(F_e, Columna + 6)).Merge
             
             End If
             
             XL.Cells(F_e, Columna).Value = .text
             XL.Cells.Cells(F_e, Columna).Interior.Color = RGB(237, 237, 237)
             XL.Cells.Cells(F_e, Columna).HorizontalAlignment = xlCenter
             XL.Cells.Font.Bold = True
         
         Next
         
         F_e = 2
         .Row = SpreadHeader + 1
         
         For Columna = 1 To .maxcols
             
             .Col = Columna
             XL.Cells(F_e, Columna).Value = .text
             XL.Cells.Cells(F_e, Columna).Interior.Color = RGB(237, 237, 237)
             XL.Cells.Cells(F_e, Columna).HorizontalAlignment = xlCenter
             XL.Cells.Font.Bold = True
         
         Next
            
    End With
   
   '------->Visualizar
   XL.Visible = True
   
Else
   
   MsgBox "Archivo esta abierto, grabe con otro nombre y luego cierre libro", , "Result"
   
End If

Me.MousePointer = 0

'If r Then
'  ms = MsgBox("Exportacion Realizada", vbInformation, "")
'End If
 
Exit Sub
Man_Error:
    
    fg_descarga
    If Err = 70 Or Err = 1004 Or Err = 91 Then Resume Next
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
    
End Sub

Private Sub Retroceder_en_la_Grilla()

On Error GoTo Man_Error

Dim MARCA As String

'Dim xml_ant As New DOMDocument40
'Dim xml_act As New DOMDocument40
Dim xml_ant As New DOMDocument60
Dim xml_act As New DOMDocument60


vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.ActiveCol

If vaSpread1.Row <= 1 Then Exit Sub

With vaSpread1

j = columna_anterior
.Row = fila_anterior
.Col = j

MARCA = IIf(vaSpread1.text = "", 0, vaSpread1.text)

If MARCA = "0" Or MARCA = " " Then
   
   MsgBox "Debe seleccionar una ruta con despacho", vbExclamation
   Exit Sub

End If

Fh = DTPicker2

For Fd = fecha_anterior To (Fh + 60) Step 1
    
 If j < 7 Then
  
  Exit For
  
 End If
   
 .Row = fila_anterior
 .Col = j
 MARCA = IIf(vaSpread1.text = "", 0, vaSpread1.text)
  
  If MARCA = "0" Or MARCA = " " Then
       
       .Row = fila_anterior
       .Col = j ' Marca la Fecha Nueva
       .text = "X"
       .TypeHAlign = TypeHAlignCenter
       .Lock = True
       vaSpread1.Row = 1
       vaSpread1.Col = j
       fecha_actual = IIf(vaSpread1.text = "", 0, vaSpread1.text)
       Sql = " sgpadm_Ins_Eliminar_Grabar_Despacho "
       Sql = Sql & " '" & Format(fecha_anterior, "YYYYMMDD") & "',"
       Sql = Sql & " '" & Format(fecha_actual, "YYYYMMDD") & "',"
       Sql = Sql & " '" & Ceco & "',"
       Sql = Sql & " '" & rutproveedor & "',"
       Sql = Sql & " '" & "CONFIGURAR" & "'"
       Set RS = vg_db.Execute(Sql)
           
       .Row = fila_anterior
       .Col = columna_anterior ' Marca la Fecha Anterior
       .text = " "
       .SetActiveCell j, vaSpread1.ActiveRow
        columna_anterior = j
        vaSpread1.Row = 1
        vaSpread1.Col = j
        fecha_anterior = IIf(vaSpread1.text = "", 0, vaSpread1.text)
       
       Exit For
  
  End If
    
    j = j - 1

Next Fd

End With
  
Call grabarLogo

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub Avanzar_en_la_Grilla()

On Error GoTo Man_Error

Dim MARCA As String

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.ActiveCol

If vaSpread1.Row <= 1 Then Exit Sub

With vaSpread1

j = columna_anterior
.Row = fila_anterior
.Col = j

MARCA = IIf(vaSpread1.text = "", 0, vaSpread1.text)

If MARCA = "0" Or MARCA = " " Then
   
   MsgBox "Debe seleccionar una ruta con despacho", vbExclamation
   Exit Sub

End If

Fh = DTPicker2

For Fd = fecha_anterior To Fh
    
    .Row = fila_anterior
    .Col = j
    MARCA = IIf(vaSpread1.text = "", 0, vaSpread1.text)
  
  If MARCA = "0" Or MARCA = " " Then
       
       .Row = fila_anterior
       .Col = j ' Marca la Fecha Nueva
       .text = "X"
       .TypeHAlign = TypeHAlignCenter
       .Lock = True
       vaSpread1.Row = 1
       vaSpread1.Col = j
       fecha_actual = IIf(vaSpread1.text = "", 0, vaSpread1.text)
       Sql = " sgpadm_Ins_Eliminar_Grabar_Despacho "
       Sql = Sql & " '" & Format(fecha_anterior, "YYYYMMDD") & "',"
       Sql = Sql & " '" & Format(fecha_actual, "YYYYMMDD") & "',"
       Sql = Sql & " '" & Ceco & "',"
       Sql = Sql & " '" & rutproveedor & "',"
       Sql = Sql & " '" & "CONFIGURAR" & "'"
       Set RS = vg_db.Execute(Sql)
           
       .Row = fila_anterior
       .Col = columna_anterior ' Marca la Fecha Anterior
       .text = " "
       .SetActiveCell j, vaSpread1.ActiveRow
        columna_anterior = j
        vaSpread1.Row = 1
        vaSpread1.Col = j
        fecha_anterior = IIf(vaSpread1.text = "", 0, vaSpread1.text)
      
       Exit For
  
  End If
    
    j = j + 1

Next Fd

End With

Call grabarLogo

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub Eliminar_Grilla()

On Error GoTo Man_Error

Dim MARCA As String

j = columna_anterior
vaSpread1.Row = fila_anterior
vaSpread1.Col = j

MARCA = IIf(vaSpread1.text = "", 0, vaSpread1.text)

If MARCA = "0" Or MARCA = " " Then
   
   MsgBox "Debe seleccionar una ruta con despacho", vbExclamation
   Exit Sub

End If

With vaSpread1

     j = columna_anterior
    .Row = fila_anterior
    .Col = j
    MARCA = IIf(vaSpread1.text = "", 0, vaSpread1.text)
  
  If MARCA = "X" Then
       
   If MsgBox("Esta Seguro de Eliminar la Fecha de Despacho ?", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
       
       .Row = fila_anterior
       .Col = j ' Marca la Fecha Nueva
       .text = "X"
       .TypeHAlign = TypeHAlignCenter
       .Lock = True
       vaSpread1.Row = 1
       vaSpread1.Col = j
       fecha_actual = IIf(vaSpread1.text = "", 0, vaSpread1.text)
       Sql = " sgpadm_Ins_Eliminar_Grabar_Despacho "
       Sql = Sql & " '" & Format(fecha_anterior, "YYYYMMDD") & "',"
       Sql = Sql & " '" & Format(fecha_anterior, "YYYYMMDD") & "',"
       Sql = Sql & " '" & Ceco & "',"
       Sql = Sql & " '" & rutproveedor & "',"
       Sql = Sql & " '" & "ELIMINAR" & "'"
       Set RS = vg_db.Execute(Sql)
           
       .Row = fila_anterior
       .Col = columna_anterior ' Marca la Fecha Anterior
       .text = " "
    
  End If
    j = j + 1

End With

Call grabarLogo

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.ActiveCol

If Row <= 0 Or vaSpread1.MaxRows < 1 Then
   
   Toolbar1.Buttons(1).Enabled = False
   Toolbar1.Buttons(3).Enabled = False
   Toolbar1.Buttons(5).Enabled = False
   Exit Sub

End If

columna_anterior = Col
fila_anterior = Row

If Col >= 7 Then
   
   Toolbar1.Buttons(1).Enabled = AccMod
   Toolbar1.Buttons(3).Enabled = AccMod
   Toolbar1.Buttons(5).Enabled = AccEli

Else
   
   Toolbar1.Buttons(1).Enabled = False
   Toolbar1.Buttons(3).Enabled = False
   Toolbar1.Buttons(5).Enabled = False

End If

vaSpread1.Row = vaSpread1.Row
vaSpread1.Col = 2
Ceco = IIf(vaSpread1.text = "", 0, vaSpread1.text)

vaSpread1.Row = vaSpread1.Row
vaSpread1.Col = 4
rutproveedor = IIf(vaSpread1.text = "", "", vaSpread1.text)
            
vaSpread1.Row = 1
vaSpread1.Col = Col
 
fecha_anterior = IIf(vaSpread1.text = "", 0, vaSpread1.text)
 
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)

'vaSpread1.Row = vaSpread1.ActiveRow
'vaSpread1.Col = vaSpread1.ActiveCol
'
On Error GoTo Man_Error
'
'If vaSpread1.MaxRows < 1 Then
'  Toolbar1.Buttons(1).Enabled = False
'  Toolbar1.Buttons(3).Enabled = False
'  Toolbar1.Buttons(5).Enabled = False
'  Exit Sub
'End If
'
'
'columna_anterior = vaSpread1.ActiveCol
'fila_anterior = vaSpread1.ActiveRow
'
'
'If KeyCode = 37 Then
'  vaSpread1.Row = vaSpread1.ActiveRow
'  vaSpread1.Col = vaSpread1.ActiveCol
'
'Else
'  columna_anterior = vaSpread1.ActiveCol + 1
'End If
'
'
'
'
'If Col >= 6 Then
'  Toolbar1.Buttons(1).Enabled = True
'  Toolbar1.Buttons(3).Enabled = True
'  Toolbar1.Buttons(5).Enabled = True
'Else
'  Toolbar1.Buttons(1).Enabled = False
'  Toolbar1.Buttons(3).Enabled = False
'  Toolbar1.Buttons(5).Enabled = False
'
'End If
'
'vaSpread1.Row = vaSpread1.Row
'vaSpread1.Col = 1
'ceco = IIf(vaSpread1.text = "", 0, vaSpread1.text)
'
'vaSpread1.Row = vaSpread1.Row
'vaSpread1.Col = 3
'rutproveedor = IIf(vaSpread1.text = "", "", vaSpread1.text)
'
'vaSpread1.Row = 1
'vaSpread1.Col = Col
'
'fecha_anterior = IIf(vaSpread1.text = "", 0, vaSpread1.text)
'
'Exit Sub
'Man_Error:
'    fg_descarga
'    MsgBox Err & ":  " & Err.Description, vbCritical, Msgtitulo

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.ActiveCol

On Error GoTo Man_Error
If KeyCode = 37 Then
   
   Row = vaSpread1.Row
   Col = vaSpread1.Col - 1

Else
   
   Row = vaSpread1.Row
   Col = vaSpread1.Col + 1

End If

If Row <= 0 Or vaSpread1.MaxRows < 1 Then
   
   Toolbar1.Buttons(1).Enabled = False
   Toolbar1.Buttons(3).Enabled = False
   Toolbar1.Buttons(5).Enabled = False
   Exit Sub

End If

columna_anterior = Col
fila_anterior = Row

If Col >= 7 Then
   
   Toolbar1.Buttons(1).Enabled = True
   Toolbar1.Buttons(3).Enabled = True
   Toolbar1.Buttons(5).Enabled = True

Else
   
   Toolbar1.Buttons(1).Enabled = False
   Toolbar1.Buttons(3).Enabled = False
   Toolbar1.Buttons(5).Enabled = False

End If

vaSpread1.Row = vaSpread1.Row
vaSpread1.Col = 2
Ceco = IIf(vaSpread1.text = "", 0, vaSpread1.text)

vaSpread1.Row = vaSpread1.Row
vaSpread1.Col = 4
rutproveedor = IIf(vaSpread1.text = "", "", vaSpread1.text)
            
vaSpread1.Row = 2
vaSpread1.Col = Col
 
fecha_anterior = IIf(vaSpread1.text = "", 0, vaSpread1.text)

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub
