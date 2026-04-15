VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form I_ExcelArrastreSaldo 
   Caption         =   "Exportar A Excel Arrastre de Saldo"
   ClientHeight    =   2475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2475
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10455
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1875
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2085
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
         Height          =   495
         Left            =   8400
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
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
         Left            =   6240
         TabIndex        =   6
         Top             =   1680
         Width           =   1935
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   1875
         TabIndex        =   3
         Top             =   1080
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         Text            =   "13/07/2004"
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
         Left            =   7815
         TabIndex        =   5
         Top             =   1080
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         Text            =   "13/07/2004"
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1875
         TabIndex        =   1
         Top             =   675
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   10
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   1
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Pedido"
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
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Opcional"
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
         Left            =   9240
         TabIndex        =   12
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin"
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
         Index           =   7
         Left            =   6720
         TabIndex        =   10
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
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
         Left            =   480
         TabIndex        =   9
         Top             =   1155
         Width           =   1065
      End
      Begin VB.Label Label2 
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
         Left            =   480
         TabIndex        =   8
         Top             =   780
         Width           =   735
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3615
         TabIndex        =   2
         Top             =   675
         Width           =   5415
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3165
         Picture         =   "I_ExcelArrastreSaldo.frx":0000
         Top             =   600
         Width           =   480
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3660
         TabIndex        =   11
         Top             =   720
         Width           =   5415
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "I_ExcelArrastreSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim MsgTitulo As String
Public lc_Aux As String

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim tipopedido      As String
Dim NomArchivoExcel As String
Dim Extension       As String

Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object

  fg_carga ""
  
  If Not ValidarDatos Then Exit Sub
  
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
    
  tipopedido = IIf(Combo1(0).ListIndex = -1, "Null", Val(fg_codigocbo(Combo1, 0, 1, "")))
  
  If Trim(fpayuda(0).Caption) = "" Then
  
     Set RS = vg_db.Execute("sgpadm_Sel_ExportarExcelArrastredeSaldoSinCeco_V02 '" & tipopedido & "', '" & Format(FpFecDesde, ("YYYYmmdd")) & "', '" & Format(FpFecHasta, ("YYYYmmdd")) & "'")
  
  Else
  
     Set RS = vg_db.Execute("sgpadm_Sel_ExportarExcelArrastredeSaldoCeco_V02 '" & Trim(LimpiaDato(fpText.text)) & "', '" & tipopedido & "','" & Format(FpFecDesde, ("YYYYmmdd")) & "', '" & Format(FpFecHasta, ("YYYYmmdd")) & "'")
  
  End If
  
  If Not RS.EOF Then
     
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
     
     
  Else
     
     MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo
  
  End If
  RS.Close
  Set RS = Nothing
  
  fg_descarga
 
Exit Sub
Man_Error:
fg_descarga

MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Command2_Click()

On Error GoTo Man_Error

Me.Hide
Unload Me

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo


End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Dim Sql As String
Dim RS  As New ADODB.Recordset

fg_centra Me

MsgTitulo = "Listar Excel Arrastre de Saldo"
FpFecHasta.text = Format(Date, "dd/mm/yyyy")
FpFecDesde.text = Format(Date, "dd/mm/yyyy")

'-------> Llenar combo Tipo Pedido
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = ""
Sql = " sgpadm_Sel_TipoPedido "
Set RS = vg_db.Execute(Sql)
Combo1(0).Clear
'Combo1(0).AddItem "Todos" & Space(150) & "(0)"
Do While Not RS.EOF
      
   Combo1(0).AddItem Trim(RS(1)) & Space(150) & "(" & Trim(RS(0)) & ")"
   
   RS.MoveNext

Loop

RS.Close
Set RS = Nothing
Combo1(0).ListIndex = 0

'-------> Llenar combo Tipo Pedido

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

If IsDate(FpFecHasta.text) = False Then Exit Sub

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_Change()
 
 On Error GoTo Man_Error
 
    Dim RS As New ADODB.Recordset
    Dim Sql As String
    
    If fpText.text = "" Then fpayuda(0).Caption = "": Exit Sub
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Sql = Trim(LimpiaDato(fpText.text))
    Set RS = vg_db.Execute("sgpadm_s_cliente_V02 29, '" & Sql & "', ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!Cli_nombre)
    RS.Close
    Set RS = Nothing
    
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Image1_Click(Index As Integer)
 
 On Error GoTo Man_Error
    
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "Clientesimap"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.text = vg_codigo: fpayuda(0).Caption = vg_nombre
    If Me.Visible Then FpFecDesde.SetFocus

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Function ValidarDatos() As Boolean

On Error GoTo Man_Error

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
   
   MsgBox "La fecha hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

If Combo1(0).ListIndex = -1 Then

   MsgBox "Tipo pedido debe ser informado...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

Exit Function
Man_Error:
fg_descarga

MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Function


