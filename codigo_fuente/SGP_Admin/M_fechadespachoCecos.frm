VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form M_fechadespachoCecos 
   Caption         =   "Mantención días de despacho casino"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   14565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   5640
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1800
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin EditLib.fpText fpText1 
      Height          =   315
      Left            =   6120
      TabIndex        =   0
      Top             =   720
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2990
      _ExtentY        =   556
      Enabled         =   0   'False
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
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
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
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mostrar en Grilla"
      Height          =   1695
      Left            =   9360
      TabIndex        =   9
      Top             =   5520
      Width           =   5055
      Begin VB.CheckBox Check3 
         Caption         =   "Sitios FM"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sitios No Simap"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sitios Simap"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Grp. Desp. Sin Parametrizar"
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
         Index           =   4
         Left            =   3120
         TabIndex        =   14
         Top             =   120
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   2
         Left            =   2640
         Picture         =   "M_fechadespachoCecos.frx":0000
         Stretch         =   -1  'True
         Top             =   180
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Grp. Desp. Parametrizado"
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
         Index           =   3
         Left            =   3075
         TabIndex        =   12
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   1
         Left            =   2640
         Picture         =   "M_fechadespachoCecos.frx":628A
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   0
         Left            =   2640
         Picture         =   "M_fechadespachoCecos.frx":C514
         Stretch         =   -1  'True
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Grp. Desp. Falta parametrizar"
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
         Index           =   2
         Left            =   3075
         TabIndex        =   11
         Top             =   660
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   2640
      TabIndex        =   4
      Top             =   5640
      Width           =   2655
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   14295
      _Version        =   393216
      _ExtentX        =   25215
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
      MaxCols         =   18
      MaxRows         =   1
      ProcessTab      =   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "M_fechadespachoCecos.frx":1279E
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   135
      Left            =   4320
      TabIndex        =   13
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
      _Version        =   393216
      _ExtentX        =   2566
      _ExtentY        =   238
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
      MaxCols         =   3
      MaxRows         =   1
      SpreadDesigner  =   "M_fechadespachoCecos.frx":13288
   End
   Begin VB.Label Label1 
      Caption         =   "Ultima Fecha Generada"
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
      Left            =   3840
      TabIndex        =   10
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "M_fechadespachoCecos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BtnX               As Variant
Dim estdes             As Boolean
Dim parametro          As Integer
Public contador        As Integer
Public codigo_anterior As Integer
Dim fecha_parametro    As String
Public modo            As String
Public codigoceco      As String
Public buff            As String
Public collec          As String
Public SW_VALIDACION   As Long

Private Sub Check1_Click()

On Error GoTo Man_Error

Call lee_fechas_cecos

'registrar Log sistema
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Ordenar"), Me.HelpContextID, "", "", "")

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
  
End Sub

Private Sub Check2_Click()

On Error GoTo Man_Error

Call lee_fechas_cecos

'registrar Log sistema
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Ordenar"), Me.HelpContextID, "", "", "")

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
  
End Sub

Private Sub Check3_Click()

On Error GoTo Man_Error
 
Call lee_fechas_cecos

'registrar Log sistema
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Ordenar"), Me.HelpContextID, "", "", "")

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
  
End Sub

Private Sub Form_Load()
  
 On Error GoTo Man_Error
    
  Dim RS As New ADODB.Recordset
  
  fg_centra Me
  Me.HelpContextID = vg_OpcM
  Toolbar1.ImageList = Partida.IL1
  
  Set BtnX = Toolbar1.Buttons.Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): BtnX.Visible = True: BtnX.ToolTipText = "Grabar ": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, False)
  Set BtnX = Toolbar1.Buttons.Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): BtnX.Visible = False: BtnX.ToolTipText = ""
''  Set BtnX = Toolbar1.Buttons.Add(, "A_Calendario", , tbrDefault, "A_Calendario"): BtnX.Visible = True: BtnX.ToolTipText = "Calendario": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, False)
'  Set BtnX = Toolbar1.Buttons.Add(, "A_Calendario", , tbrDefault, "A_Calendario"): BtnX.Visible = True: BtnX.ToolTipText = "Calendario": BtnX.Enabled = True
  Set BtnX = Toolbar1.Buttons.Add(, "A_Calendario", , tbrDropdown, "A_Calendario"): BtnX.Visible = True: BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False): BtnX.ToolTipText = "Calendarios de Rutas": BtnX.ButtonMenus.Add text:="Calendario de Fecha Despachos Cecos y Proveedores": BtnX.ButtonMenus.Add text:="Calendario de Fecha Rutas x Grupo de Despachos Cecos y Proveedores"
  Set BtnX = Toolbar1.Buttons.Add(, "A_BajarF", , tbrDefault, "A_BajarF"): BtnX.Visible = True: BtnX.ToolTipText = "Ordenar"
  Set BtnX = Toolbar1.Buttons.Add(, "Proceso", , tbrDropdown, "Proceso"): BtnX.Visible = True: BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): BtnX.ToolTipText = "Generar Rutas": BtnX.ButtonMenus.Add text:="Generar Rutas Normal ": BtnX.ButtonMenus.Add text:="Generar Rutas Grupo Despacho "
'  Set BtnX = Toolbar1.Buttons.Add(, "Proceso", , tbrDefault, "Proceso"): BtnX.Visible = True: BtnX.ToolTipText = "Generar Rutas ": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
'  Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exportar Rutas a Excel "
  Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDropdown, "excel"): BtnX.Visible = True: BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, False): BtnX.ToolTipText = "Exportar ruta masiva excel": BtnX.ButtonMenus.Add text:="Exportar rutas normal": BtnX.ButtonMenus.Add text:="Exportar rutas PEL"
  Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
  Set BtnX = Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDropdown, "A_CopiarD"): BtnX.Visible = True: BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, False): BtnX.ToolTipText = "Carga Masiva Excel": BtnX.ButtonMenus.Add text:="Eliminar rutas desde excel": BtnX.ButtonMenus.Add text:="Agregar rutas desde excel": BtnX.ButtonMenus.Add text:="Eliminar rutas grupo despacho desde excel": BtnX.ButtonMenus.Add text:="Agregar rutas grupo despacho desde excel"
  Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
  Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
  
  estdes = True
  parametro = 0
  
  modo = ""
  Sql = " sgpadm_sel_BuscaExisteFechaParametros "
  Set RS = vg_db.Execute(Sql)
  If Not RS.EOF Then
     
     fecha_parametro = RS(0)
  
  End If
  RS.Close
   
  If fecha_parametro <> "" Then
     
     fpText1 = Mid(fecha_parametro, 7, 2) + "/" + Mid(fecha_parametro, 5, 2) + "/" + Mid(fecha_parametro, 1, 4)
  
  End If
  
  Call lee_fechas_cecos

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)
  
End Sub

Private Sub lee_fechas_cecos()
 
 On Error GoTo Man_Error
 
    Dim Sql    As String
    Dim RS     As New ADODB.Recordset
    Dim codigo As Integer
    Dim AssMod As Boolean
    Dim i      As Long
    
    '-------> Dar acceso modificar rutas
    AssMod = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, False)
    
    parametro = parametro + 1
 
    If parametro = 4 Then
       
       parametro = 1
    
    End If
    
    Text1(2) = ""
    Text1(3) = ""
    Text1(4) = ""
 
    Dim filtro As String
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

    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Sql = " sgpadm_sel_parametros_despacho_casino_V02  " & parametro & ",'" & filtro & "'"
    Set RS = vg_db.Execute(Sql)
    
   '-------> Inicio LLenar grilla
   
   With vaSpread1
    
    .Visible = False
    .MaxRows = 0
    .MaxRows = 0
 
    .MaxRows = RS.RecordCount
    i = 1
    Do While Not RS.EOF
        
'        .MaxRows = .MaxRows + 1
        .Row = i '.MaxRows
        estdes = True
        
        .Col = 2 ' Org. Compras
        .text = RS(14)
        .TypeHAlign = TypeHAlignCenter
        .Lock = IIf(AssMod = True, False, True)
        
        .Col = 3 ' Ceco
        .text = RS(0) 'Val(RS(0))
        .TypeHAlign = TypeHAlignCenter
        .Lock = IIf(AssMod = True, False, True)
        
        .Col = 4 ' Nombre
        .text = RS(1)
        .Lock = IIf(AssMod = True, False, True)
        
        .Col = 5 ' Codigo Despacho
        .text = IIf(RS(2) = 0, "", RS(2))
        .TypeHAlign = TypeHAlignRight
        .Lock = IIf(AssMod = True, False, True)
        
        .Col = 6 ' Descripcion Despacho
        .text = RS(3)
        .Lock = IIf(AssMod = True, False, True)
       
        vaSpread1.Col = 7 ' Lunes
        vaSpread1.text = RS(4)
        .Lock = IIf(AssMod = True, False, True)
                 
        .Col = 8 ' Martes
        .text = RS(5)
        .Lock = IIf(AssMod = True, False, True)
                
        .Col = 9 ' Miercoles
        .text = RS(6)
        .Lock = IIf(AssMod = True, False, True)
                       
        .Col = 10 ' Jueves
        .text = RS(7)
        .Lock = IIf(AssMod = True, False, True)
                
        .Col = 11 ' Viernes
        .text = RS(8)
        .Lock = IIf(AssMod = True, False, True)
        
        .Col = 12 ' Sabado
        .text = RS(9)
        .Lock = IIf(AssMod = True, False, True)
                
        .Col = 13 ' Domingo
        .text = RS(10)
        .Lock = IIf(AssMod = True, False, True)
                 
        .Col = 14 ' Trabaja el Fin de Semana
        .text = RS(11)
        .Lock = IIf(AssMod = True, False, True)
        
        'si no trabaja fin de semana, bloquea SAB-DOM
        
        If RS(11) = 0 Then
        
          .Col = 12: .Lock = True
          .Col = 13: .Lock = True
        
        End If
                
        .Col = 15 ' Fecha desde
        .text = ""
        .Lock = True
        .Lock = IIf(AssMod = True, False, True)
        
        .Col = 16 ' Actualizado
        .text = "0"
        .Lock = IIf(AssMod = True, False, True)
        
        .Col = 17 ' Estado
        .text = "0"
        
         vaSpread2.Row = 1
         vaSpread2.Col = IIf(RS!Parametrizado = "0", 3, IIf(RS!Parametrizado = RS!NGrupoDespacho, 2, 1))

         .Col = 18
         .TypePictPicture = vaSpread2.TypePictPicture
        
        i = i + 1
        RS.MoveNext
        
    Loop
    
    .Visible = True
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = False
        
    ret = .ExportToXMLBuffer("ParamCeco", collec, buff, ExportToXMLFormattedData, "")
    
    End With
    
    estdes = True

Text1(2).text = ""
Text1(3).text = ""
Text1(4).text = ""

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
End Sub

' -------------------------------------------------------------------------------------------
' \\ -- Función para crear un nuevo libro con el contenido del Grid
' -------------------------------------------------------------------------------------------
Public Function Exportar_Excel(ByVal sOutputPath As String, ByRef Grid As vaSpread) As Boolean
  
    On Error GoTo Error_Handler
  
'    Dim o_Excel     As excel.Application  'Object
'    Dim o_Libro     As excel.Workbook 'Object
'    Dim o_Hoja      As New excel.Worksheet

    Dim o_Excel  As Object
    Dim o_Libro  As Object
    Dim o_Hoja   As Object
    
    Dim F_e      As Long
    Dim F_g      As Long
    Dim fila1    As Long
    Dim Columna  As Long
    Dim FILA2    As Long
       
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
                  
                  o_Hoja.Range(o_Hoja.Cells(F_e, Columna), o_Hoja.Cells(F_e, Columna + 5)).Merge
                
                End If
                
                o_Hoja.Cells(F_e, Columna).Value = .text
                o_Hoja.Cells.Cells(F_e, Columna).Interior.color = RGB(237, 237, 237)
                o_Hoja.Cells.Cells(F_e, Columna).HorizontalAlignment = xlCenter
                o_Hoja.Cells.Font.Bold = True
            
            Next
            
            F_e = 2
            .Row = SpreadHeader + 1
            For Columna = 1 To .maxcols
                
                .Col = Columna
                o_Hoja.Cells(F_e, Columna).Value = .text
                o_Hoja.Cells.Cells(F_e, Columna).Interior.color = RGB(237, 237, 237)
                o_Hoja.Cells.Cells(F_e, Columna).HorizontalAlignment = xlCenter
                o_Hoja.Cells.Font.Bold = True
            
            Next
    
    End With
       
    ' -- Bucle para Exportar los datos
    
 
    
    With Grid
         
         F_e = 3
         For F_g = 2 To .MaxRows
            
            .Row = F_g
            
            For Columna = 1 To .maxcols
                .Col = Columna
                If Columna >= 1 And Columna < 6 Then
                  o_Hoja.Cells.Cells(F_e, Columna).Interior.color = RGB(203, 255, 209)
                  o_Hoja.Cells.Cells(F_e, Columna).HorizontalAlignment = xlLeft
                Else
                  o_Hoja.Cells.Cells(F_e, Columna).Interior.color = RGB(244, 247, 170)
                  o_Hoja.Cells.Cells(F_e, Columna).HorizontalAlignment = xlCenter
                End If
                  o_Hoja.Cells(F_e, Columna).Value = .text
            Next
            
            F_e = F_e + 1
         
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
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing

    MsgBox Err.Description, vbCritical, MsgTitulo
End Function

Private Sub graba_fechas_despachos()
 
 On Error GoTo Man_Error
 
 Dim i                 As Integer
 Dim Ceco              As String
 Dim codigo            As Long
 Dim Nombre            As String
 Dim lunes             As Integer
 Dim martes            As Integer
 Dim miercoles         As Integer
 Dim jueves            As Integer
 Dim viernes           As Integer
 Dim sabado            As Integer
 Dim domingo           As Integer
 Dim estext            As String
 Dim fechamodif        As String
 Dim actualiza         As String
 Dim buffact           As String
 Dim collecact         As String
 Dim modulo            As String
 Dim buffantes         As String
 Dim buffactual        As String
 Dim existeActualizado As Integer
 Dim fechaavalidar     As String
 Dim RS                As New ADODB.Recordset
' Dim xml_ant           As New DOMDocument40
' Dim xml_act           As New DOMDocument40
 Dim xml_ant           As New DOMDocument60
 Dim xml_act           As New DOMDocument60
 
 Screen.MousePointer = 11
 i = 0
 
 '--> Validar que no exista fecha rutas normal
 vg_BorradoDatos = False
 If Not ValidarRutasxGrupo Then
 
    Exit Sub
 
 End If
 
 'Valida ante de Gbrabar antes cualquie cosa
 
 'solo reprocesa el ceco Modificado
    
    For i = 1 To vaSpread1.MaxRows
       
        vaSpread1.Row = i
        
        vaSpread1.Col = 15 'Marca Actualiza
        fechaavalidar = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        vaSpread1.Col = 16 'Marca Actualiza
        actualiza = IIf(vaSpread1.text = "", 0, vaSpread1.text)
                
        If fechaavalidar <> "0" Then
          
          vaSpread1.Col = 15 'fecha
          fechamodif = Format(vaSpread1.text, "YYYYMMDD")
          
          SW_VALIDACION = 0
          
          Call validar
          
          vaSpread1.SetActiveCell 15, vaSpread1.ActiveRow
              
          If SW_VALIDACION = 1 Then
                
             Exit For
              
          End If
         
        End If
        
    
    Next i
 
If SW_VALIDACION = 0 Then
 
   modulo = "Mantendedor de Fechas de Despacho"
   'ret = vaSpread1.ExportToXML(App.Path & "\datos_actual.xml", "ParamCeco", collecact, ExportToXMLFormattedData, "")
   ret = vaSpread1.ExportToXMLBuffer("ParamCeco", collecact, buffact, ExportToXMLFormattedData, "")

   buffantes = buff
   buffactual = buffact

   xml_ant.LoadXml (buffantes)
   xml_act.LoadXml (buffact)
  
   existeActualizado = vaSpread1.SearchCol(16, 0, -1, "1", SearchFlagsValue)

   If existeActualizado > 0 Then

      Sql = " sgpadm_Ins_Log_Fecha_Despacho  "
      Sql = Sql & "'" & UCase(vg_NUsr) & "',"
      Sql = Sql & "'" & modulo & "',"
      Sql = Sql & "'" & xml_ant.DocumentElement.XML & "',"
      Sql = Sql & "'" & xml_act.DocumentElement.XML & "'"

      'Debug.Print sql
      Set RS = vg_db.Execute(Sql)
    
   End If

   ' Actualiza todos los Cecos primero
 
   For i = 1 To vaSpread1.MaxRows
        
       estext = False
       vaSpread1.Row = i
        
       vaSpread1.Col = 3 'ceco
       Ceco = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
       vaSpread1.Col = 5 'codigo
       codigo = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
       vaSpread1.Col = 6 'Descripcion
       Nombre = IIf(vaSpread1.text = "", "", vaSpread1.text)
        
       vaSpread1.Col = 7 'Lunes
       lunes = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
       vaSpread1.Col = 8 'Martes
       martes = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
       vaSpread1.Col = 9 'Miercoles
       miercoles = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
       vaSpread1.Col = 10 'Jueves
       jueves = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
       vaSpread1.Col = 11 'Viernes
       viernes = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
       vaSpread1.Col = 12 'Sabado
       sabado = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
       vaSpread1.Col = 13 'Domingo
       domingo = IIf(vaSpread1.text = "", 0, vaSpread1.text)
         
       vaSpread1.Col = 16 'Marca Actualiza
       actualiza = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
         
       If codigo > 0 And actualiza = 1 Then
           
          If lunes <> 0 Or martes <> 0 Or miercoles <> 0 Or jueves <> 0 Or viernes <> 0 Or sabado <> 0 Or domingo <> 0 Then
              
             Sql = ""
             Sql = " sgpadm_iu_Fechasdespachocecos "
             Sql = Sql & " '" & Ceco & "',"
             Sql = Sql & codigo & ","
             Sql = Sql & " '" & Nombre & "',"
             Sql = Sql & lunes & ","
             Sql = Sql & martes & ","
             Sql = Sql & miercoles & ","
             Sql = Sql & jueves & ","
             Sql = Sql & viernes & ","
             Sql = Sql & sabado & ","
             Sql = Sql & domingo
             Set RS = vg_db.Execute(Sql)
            
          Else
              
             Sql = ""
             Sql = " sgpadm_Del_ParamDespachoCecos "
             Sql = Sql & " '" & Ceco & "' "
             Set RS = vg_db.Execute(Sql)
            
          End If
        
       End If
    
   Next i
    
   'solo reprcesa el ceco Modificado
    
   For i = 1 To vaSpread1.MaxRows
       
       estext = False
       vaSpread1.Row = i
        
       vaSpread1.Col = 15 'Marca Actualiza
       fechaavalidar = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        vaSpread1.Col = 16 'Marca Actualiza
        actualiza = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        If fechaavalidar <> "0" Then
          
          vaSpread1.Col = 15 'fecha
          fechamodif = Format(vaSpread1.text, "YYYYMMDD")
          
          vaSpread1.Col = 3 'ceco
          Ceco = IIf(vaSpread1.text = "", 0, vaSpread1.text)
          
          xmlfamilia = ""
          xmlfamilia = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
          xmlfamilia = xmlfamilia & "<GrabaRutaCeco>"
'          vaSpread1.Row = i
'          vaSpread1.Col = 2 'Id Ruta de Compras
'          codigo = IIf(vaSpread1.text = "", 0, vaSpread1.text)
          xmlfamilia = xmlfamilia & "<RutaCeco  Ceco = " & Chr(34) & Ceco & Chr(34)
          xmlfamilia = xmlfamilia & "/>"
          xmlfamilia = xmlfamilia & "</GrabaRutaCeco>"
          
          Sql = "sgpadm_Ins_XmlGenerarRutaCeco "
          Sql = Sql & " '" & xmlfamilia & "', "
          Sql = Sql & fechamodif & ", "
          Sql = Sql & fecha_parametro & ", "
          Sql = Sql & IIf(vg_BorradoDatos, "1", "0") & ", "
          Sql = Sql & Trim(vg_NUsr)
         
         Set RS = vg_db.Execute(Sql)
         If Not RS.EOF Then
            
            If RS(0) > 0 Then
               
               MsgBox RS(1)
            
            End If
         
         End If
          
         RS.Close
         Set RS = Nothing
       
        End If
        
    Next i
     
     Screen.MousePointer = 0
     MsgBox "Se actualizaron las fechas de despachos de los cecos y se Reprocesaron las Rutas Modificadas", vbExclamation
     parametro = 0
     
     Call lee_fechas_cecos

End If
     
fg_descarga

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
    vaSpread1.Col = 17
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
           vaSpread1.Col = 2
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 17
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 17
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 17
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 17
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 17
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
           
           vaSpread1.Col = 17
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

'Private Sub Text1_Change(Index As Integer)
'
'On Error GoTo Man_Error
'
'Select Case Index
'
'Case 2, 3
'
'    vaSpread1.Visible = False
'
'    If Trim(Text1(Index).text) <> "" Then
'
'       For i = 1 To vaSpread1.MaxRows
'
'           vaSpread1.Row = i
'           vaSpread1.Col = Index
'           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(Text1(Index).text) & "*"
'           vaSpread1.Col = 3
'           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
'              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
'           Else
'              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
'           End If
'
'        Next i
'
'        vaSpread1.SetActiveCell Index, 1
'
'    End If
'
''    vaSpread1_Click Index, 0
'    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
'    vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
'    vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
'    vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
'    If Trim(Text1(Index).text) = "" Then
'       For i = 1 To vaSpread1.MaxRows
'           vaSpread1.Row = i
'           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
'       Next
'       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
'       vaSpread1.SetActiveCell Index, 1
'    End If
'    vaSpread1.Visible = True
'
'End Select
'Man_Error:
'    fg_descarga
'    If Err.Number > 0 Then
'       MsgBox Err & ":  " & Err.Description, vbCritical, Msgtitulo
'    End If
'
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim seleccion As String
Dim i As Integer
Dim xmlfamilia As String
    
Select Case Button.Index
        
    Case 1 ' Grabo Fecha de Despacho
        
      'registrar Log sistema modificar
       Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), CStr(Me.HelpContextID), "", "", "")
        
       Call graba_fechas_despachos
       
       'registrar Log sistema actualizar lista
       Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")
        
    Case 4 ' Ordenar información
        
       Call lee_fechas_cecos
       
       'registrar Log sistema
       Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Ordenar"), Me.HelpContextID, "", "", "")
       
    Case 5 ' Generar ARchivo Rutas
 
    
    Case 10 ' Salir del Programa
        
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

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

On Error GoTo Man_Error

Dim op As Integer

Select Case ButtonMenu

    Case "Calendario de Fecha Despachos Cecos y Proveedores"
            
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(1196001), "", "", "")
         Call M_Calendario_fechas_despachos.Show(1)
    
    Case "Calendario de Fecha Rutas x Grupo de Despachos Cecos y Proveedores"
    
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(1196006), "", "", "")
         Call M_Calendario_Fechas_GrupoDespacho.Show(1)
    
    Case "Generar Rutas Normal "
    
         If Mid(ValidarUsuarioAcceso(1196003, vg_NUsr), 1, 1) <> "1" Then
           
            MsgBox "No tiene acceso Pedido - Generación de Ruta Masiva", vbInformation, MsgTitulo
            Exit Sub
           
         End If
            
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(1196003), "", "", "")
         Call M_Generar_Archivo_Rutas.Inicio("Generar Rutas Normal ", "1", CStr(1196003))
         Call M_Generar_Archivo_Rutas.Show(1)
    
    Case "Generar Rutas Grupo Despacho "
    
         If Mid(ValidarUsuarioAcceso(1196004, vg_NUsr), 1, 1) <> "1" Then
           
            MsgBox "No tiene acceso Pedido - Generación de Ruta Grupo Despacho", vbInformation, MsgTitulo
            Exit Sub
           
         End If
            
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(1196004), "", "", "")
         Call M_Generar_Archivo_Rutas.Inicio("Generación de Ruta Grupo Despacho", "2", CStr(1196004))
         Call M_Generar_Archivo_Rutas.Show(1)
    
    Case "Eliminar rutas desde excel", "Agregar rutas desde excel", "Eliminar rutas grupo despacho desde excel", "Agregar rutas grupo despacho desde excel"
    
        
        'Abrimos el Commondialog con ShowOpen
        CD.DialogTitle = "Seleccione un archivo excel"
        CD.Filter = "Archivos xls|*.xls|Archivos xlsx|*.xlsx"
        CD.DefaultExt = "*.xls|*.xlsx"
        CD.FilterIndex = 2
        CD.Flags = cdlOFNFileMustExist
        CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
        CD.FileName = ""
        CD.ShowOpen
    
        'Si seleccionamos un archivo mostramos la ruta
        If CD.FileName <> "" Then
    
            If ButtonMenu = "Eliminar rutas desde excel" Or ButtonMenu = "Agregar rutas desde excel" Then
               
               ValidarPlantillaExcel CD.FileName, IIf(ButtonMenu = "Eliminar rutas desde excel", 1, 2)
            
            Else
            
               ValidarPlantillaExcelGrupoDespacho CD.FileName, IIf(ButtonMenu = "Eliminar rutas grupo despacho desde excel", 1, 2)
               
            End If
    
        Else
            'Si no mostramos un texto de advertencia de que no se seleccionó _
            ninguno, ya que FileName devuelve una cadena vacía
            MsgBox "No seleccionó ningún archivo"
    
        End If
    
    Case "Exportar rutas normal"
           
         If Mid(ValidarUsuarioAcceso(1196002, vg_NUsr), 1, 1) <> "1" Then
           
              MsgBox "No tiene acceso a la opción Pedido - Generar Ruta Despacho Excel", vbInformation, MsgTitulo
              Exit Sub
           
         End If
           
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(1196002), "", "", "")
         M_Generacion_Rutas_Excel.MoverOptionInicio "1"
         M_Generacion_Rutas_Excel.Show 1, Partida
    
    Case "Exportar rutas PEL"
    
         If Mid(ValidarUsuarioAcceso(1196002, vg_NUsr), 1, 1) <> "1" Then
           
            MsgBox "No tiene acceso a la opción Pedido - Generar Ruta Despacho Excel (CD-PAP)", vbInformation, MsgTitulo
            Exit Sub
           
         End If
         
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(1196002), "", "", "")
         M_Generacion_Rutas_Excel.MoverOptionInicio "2"
           
         M_Generacion_Rutas_Excel.Show 1, Partida

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub ValidarPlantillaExcel(NombreArchivo As String, OpMenu As Integer)

On Error GoTo Man_Error

Dim PathXls     As String
Dim File_Ext    As String
Dim NomHoja     As String
Dim dbexcel     As Database
Dim cn          As ADODB.Connection
Dim RS          As New ADODB.Recordset
Dim MyBuffer    As String
Dim Ceco        As String
Dim proveedor   As String
Dim Fecha       As Long
Dim tipopedido  As String
Dim EstPro      As Boolean
Dim NomArchivoExcel As String

'Definición variables excel
Dim xlApp       As Object
Dim xlWb        As Object
Dim xlWs        As Object

Set RsExcel = New ADODB.Recordset
Set cn = New ADODB.Connection

PathXls = Trim(NombreArchivo)
File_Ext = UCase(Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, "."))))
NomHoja = "Hoja1$"
With cn
     
     Select Case File_Ext
        
        ' Excel 97/2003
        Case "XLS"
          
          .Provider = "Microsoft.Jet.OLEDB.4.0"
          .ConnectionString = "Data Source=" & PathXls & ";" & "Extended Properties=Excel 8.0;"
          .CursorLocation = 3
          
        ' Excel 2010
        Case "XLSX"

          .Provider = "Microsoft.ACE.OLEDB.12.0;Data Source=" & Trim(PathXls) & ";"
          .ConnectionString = "Extended Properties=Excel 8.0;"
          .CursorLocation = 3
     
     End Select
     
     .Open

End With

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<RutaDes>"
EstPro = True

RsExcel.Open ("SELECT * FROM [" & NomHoja & "]"), cn

If RsExcel.EOF Then Exit Sub

RsExcel.MoveFirst

Do While RsExcel.EOF <> True
           
   If RsExcel.Fields(0).Value = "*" Or IsNull(RsExcel.Fields(0).Value) Then Exit Do
           
   Ceco = ""
   proveedor = ""
   Fecha = 0
   
   If Not IsNull(RsExcel.Fields(0).Value) Then
     
     Ceco = RsExcel.Fields(0).Value
     
   ElseIf IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Then
   
      EstPro = False
      MsgBox "Valor del ceco esta null o bien no tiene datos. Proceso cancelado ", vbCritical, MsgTitulo
      Exit Do
   
   End If
   
   If Not IsNull(RsExcel.Fields(2).Value) Then
      
      proveedor = RsExcel.Fields(2).Value
   
   ElseIf IsNull(RsExcel.Fields(2).Value) Or Trim(RsExcel.Fields(2).Value) = "" Then
      
      EstPro = False
      MsgBox "Valor del proveedor esta null o bien no tiene datos. Proceso cancelado ", vbCritical, MsgTitulo
      Exit Do
       
   End If
   
   If Not IsNull(RsExcel.Fields(4).Value) Then
   
      Fecha = Mid(RsExcel.Fields(4).Value, 7, 4) & Mid(RsExcel.Fields(4).Value, 4, 2) & Mid(RsExcel.Fields(4).Value, 1, 2)
    
   ElseIf IsNull(RsExcel.Fields(4).Value) Or Trim(RsExcel.Fields(4).Value) = "" Or Trim(RsExcel.Fields(4).Value) = "0" Then
   
      EstPro = False
      MsgBox "Valor fecha mal formateado. Proceso cancelado ", vbCritical, MsgTitulo
      Exit Do
   
   End If
   
   If Not IsNull(RsExcel.Fields(5).Value) Then
   
      tipopedido = Trim(RsExcel.Fields(5).Value)
    
   ElseIf IsNull(RsExcel.Fields(5).Value) Or Trim(RsExcel.Fields(5).Value) = "" Or Trim(RsExcel.Fields(5).Value) = "0" Then
   
      EstPro = False
      MsgBox "Tipo Pedido debe ser CD, PAP. Proceso cancelado ", vbCritical, MsgTitulo
      Exit Do
   
   End If
   
   If Trim(LimpiaDato(Ceco)) <> "" And Trim(LimpiaDato(proveedor)) <> "" And Fecha > 0 And Trim(LimpiaDato(tipopedido)) <> "" Then
   
      MyBuffer = MyBuffer & " <RutaDesp"
   
      Ceco = Replace(Trim(Ceco), Chr(34), "&quot;")
      Ceco = Replace(Trim(Ceco), Chr(38), "&amp;")
      Ceco = Replace(Trim(Ceco), Chr(39), "&apos;")
      Ceco = Replace(Trim(Ceco), Chr(60), "&lt;")
      Ceco = Replace(Trim(Ceco), Chr(62), "&gt;")
      
      proveedor = Replace(Trim(proveedor), Chr(34), "&quot;")
      proveedor = Replace(Trim(proveedor), Chr(38), "&amp;")
      proveedor = Replace(Trim(proveedor), Chr(39), "&apos;")
      proveedor = Replace(Trim(proveedor), Chr(60), "&lt;")
      proveedor = Replace(Trim(proveedor), Chr(62), "&gt;")
      
      Fecha = Replace(Trim(Fecha), Chr(34), "&quot;")
      Fecha = Replace(Trim(Fecha), Chr(38), "&amp;")
      Fecha = Replace(Trim(Fecha), Chr(39), "&apos;")
      Fecha = Replace(Trim(Fecha), Chr(60), "&lt;")
      Fecha = Replace(Trim(Fecha), Chr(62), "&gt;")
      
      tipopedido = Replace(Trim(tipopedido), Chr(34), "&quot;")
      tipopedido = Replace(Trim(tipopedido), Chr(38), "&amp;")
      tipopedido = Replace(Trim(tipopedido), Chr(39), "&apos;")
      tipopedido = Replace(Trim(tipopedido), Chr(60), "&lt;")
      tipopedido = Replace(Trim(tipopedido), Chr(62), "&gt;")
      
      MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
      MyBuffer = MyBuffer & " Pro = " & Chr(34) & proveedor & Chr(34)
      MyBuffer = MyBuffer & " Fec = " & Chr(34) & Fecha & Chr(34)
      MyBuffer = MyBuffer & " TiP = " & Chr(34) & tipopedido & Chr(34)

      MyBuffer = MyBuffer & "/>"
   
   End If
   
   DoEvents
           
   RsExcel.MoveNext
   
Loop
        
MyBuffer = MyBuffer & "</RutaDes>"

RsExcel.Close
Set RsExcel = Nothing
    
cn.Close
Set cn = Nothing

If EstPro Then

              
      'registrar Log sistema
      If OpMenu = 1 Then
      
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Eliminar"), Me.HelpContextID, "", "", "")
         
      Else
      
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Agregar"), Me.HelpContextID, "", "", "")
      
      End If
              
      Set RS = vg_db.Execute("sgpadm_Del_XmlRutasDespacho '" & MyBuffer & "', '" & IIf(OpMenu = 1, "Manteneción Rutas Despacho - Eliminación Ruta", "Manteneción Rutas Despacho - Agregar Rutas") & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "', '" & OpMenu & "'")

      If Not RS.EOF Then

         If RS(0) > 0 Then

            MsgBox RS(1)
            fg_descarga

            RS.Close
            Set RS = Nothing
   
            'registrar Log sistema eliminar & Agregado
      
            If OpMenu = 1 Then
            
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminado"), Me.HelpContextID, "", "", "")
               
            Else
            
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Agregado"), Me.HelpContextID, "", "", "")
               
            End If
            
            Exit Sub
      
         Else
   
            MsgBox RS(1) & VgLinea & VgLinea & "            Adjunto archivo con error", vbCritical, MsgTitulo
            
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
    
            xlApp.Columns("A:B").Select
            xlApp.Selection.Delete Shift:=xlToLeft
  
            NomArchivoExcel = fg_ArchivoXls(IIf(OpMenu = 1, "ReporteError_RutasDespachosEliminar", "ReporteError_RutasDespachos_Agregar"))
                    
            xlWb.Close True, NomArchivoExcel

            Dim XL As New excel.Application 'Crea el objeto excel
            XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
            XL.Visible = True
            XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
            '-- Cerrar Excel
            xlApp.Quit
            '-------> Release Excel references
            Set xlWs = Nothing
            Set xlWb = Nothing
            Set xlApp = Nothing
   
            'registrar Log sistema eliminar & Agregado
      
            If OpMenu = 1 Then
            
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, "", "", "")
               
            Else
            
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Agregado"), Me.HelpContextID, "", "", "")
               
            End If
         
         End If
      
      End If
      RS.Close
      Set RS = Nothing

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub ValidarPlantillaExcelGrupoDespacho(NombreArchivo As String, OpMenu As Integer)

On Error GoTo Man_Error

Dim PathXls         As String
Dim File_Ext        As String
Dim NomHoja         As String
Dim dbexcel         As Database
Dim cn              As ADODB.Connection
Dim RS              As New ADODB.Recordset
Dim MyBuffer        As String
Dim Ceco            As String
Dim GrDespacho      As String
Dim proveedor       As String
Dim Fecha           As Long
Dim tipopedido      As String
Dim EstPro          As Boolean
Dim NomArchivoExcel As String

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object

Set RsExcel = New ADODB.Recordset
Set cn = New ADODB.Connection

PathXls = Trim(NombreArchivo)
File_Ext = UCase(Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, "."))))
NomHoja = "Hoja1$"
With cn
     
     Select Case File_Ext
        
        ' Excel 97/2003
        Case "XLS"
          
          .Provider = "Microsoft.Jet.OLEDB.4.0"
          .ConnectionString = "Data Source=" & PathXls & ";" & "Extended Properties=Excel 8.0;"
          .CursorLocation = 3
          
        ' Excel 2010
        Case "XLSX"

          .Provider = "Microsoft.ACE.OLEDB.12.0;Data Source=" & Trim(PathXls) & ";"
          .ConnectionString = "Extended Properties=Excel 8.0;"
          .CursorLocation = 3
     
     End Select
     
     .Open

End With

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<RutaDes>"
EstPro = True

RsExcel.Open ("SELECT * FROM [" & NomHoja & "]"), cn

If RsExcel.EOF Then Exit Sub

RsExcel.MoveFirst

Do While RsExcel.EOF <> True
           
   If RsExcel.Fields(0).Value = "*" Or IsNull(RsExcel.Fields(0).Value) Then Exit Do
           
   Ceco = ""
   proveedor = ""
   Fecha = 0
   
    'Ceco
   If Not IsNull(RsExcel.Fields(0).Value) Then
     
     Ceco = RsExcel.Fields(0).Value
     
   ElseIf IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Then
   
      EstPro = False
      MsgBox "Valor del ceco esta null o bien no tiene datos. Proceso cancelado ", vbCritical, MsgTitulo
      Exit Do
   
   End If
   
   'Grupo Despacho
   If Not IsNull(RsExcel.Fields(2).Value) Then
      
      GrDespacho = RsExcel.Fields(2).Value
   
   ElseIf IsNull(RsExcel.Fields(2).Value) Or Trim(RsExcel.Fields(2).Value) = "" Then
      
      EstPro = False
      MsgBox "Valor del grupo despacho esta null o bien no tiene datos. Proceso cancelado ", vbCritical, MsgTitulo
      Exit Do
       
   End If
   
   'Proveedor
   If Not IsNull(RsExcel.Fields(4).Value) Then
      
      proveedor = RsExcel.Fields(4).Value
   
   ElseIf IsNull(RsExcel.Fields(4).Value) Or Trim(RsExcel.Fields(4).Value) = "" Then
      
      EstPro = False
      MsgBox "Valor del proveedor esta null o bien no tiene datos. Proceso cancelado ", vbCritical, MsgTitulo
      Exit Do
       
   End If
   
   If Not IsNull(RsExcel.Fields(6).Value) Then
   
      Fecha = Mid(RsExcel.Fields(6).Value, 7, 4) & Mid(RsExcel.Fields(6).Value, 4, 2) & Mid(RsExcel.Fields(6).Value, 1, 2)
    
   ElseIf IsNull(RsExcel.Fields(6).Value) Or Trim(RsExcel.Fields(6).Value) = "" Or Trim(RsExcel.Fields(4).Value) = "0" Then
   
      EstPro = False
      MsgBox "Valor fecha mal formateado. Proceso cancelado ", vbCritical, MsgTitulo
      Exit Do
   
   End If
   
   If Not IsNull(RsExcel.Fields(7).Value) Then
   
      tipopedido = Trim(RsExcel.Fields(7).Value)
    
   ElseIf IsNull(RsExcel.Fields(7).Value) Or Trim(RsExcel.Fields(7).Value) = "" Or Trim(RsExcel.Fields(7).Value) = "0" Then
   
      EstPro = False
      MsgBox "Tipo Pedido debe ser CD, PAP. Proceso cancelado ", vbCritical, MsgTitulo
      Exit Do
   
   End If
   
   If Trim(LimpiaDato(Ceco)) <> "" And Trim(LimpiaDato(GrDespacho)) And Trim(LimpiaDato(proveedor)) <> "" And Fecha > 0 And Trim(LimpiaDato(tipopedido)) <> "" Then
   
      MyBuffer = MyBuffer & " <RutaDesp"
   
      Ceco = Replace(Trim(Ceco), Chr(34), "&quot;")
      Ceco = Replace(Trim(Ceco), Chr(38), "&amp;")
      Ceco = Replace(Trim(Ceco), Chr(39), "&apos;")
      Ceco = Replace(Trim(Ceco), Chr(60), "&lt;")
      Ceco = Replace(Trim(Ceco), Chr(62), "&gt;")
      
      GrDespacho = Replace(Trim(GrDespacho), Chr(34), "&quot;")
      GrDespacho = Replace(Trim(GrDespacho), Chr(38), "&amp;")
      GrDespacho = Replace(Trim(GrDespacho), Chr(39), "&apos;")
      GrDespacho = Replace(Trim(GrDespacho), Chr(60), "&lt;")
      GrDespacho = Replace(Trim(GrDespacho), Chr(62), "&gt;")
      
      proveedor = Replace(Trim(proveedor), Chr(34), "&quot;")
      proveedor = Replace(Trim(proveedor), Chr(38), "&amp;")
      proveedor = Replace(Trim(proveedor), Chr(39), "&apos;")
      proveedor = Replace(Trim(proveedor), Chr(60), "&lt;")
      proveedor = Replace(Trim(proveedor), Chr(62), "&gt;")
      
      Fecha = Replace(Trim(Fecha), Chr(34), "&quot;")
      Fecha = Replace(Trim(Fecha), Chr(38), "&amp;")
      Fecha = Replace(Trim(Fecha), Chr(39), "&apos;")
      Fecha = Replace(Trim(Fecha), Chr(60), "&lt;")
      Fecha = Replace(Trim(Fecha), Chr(62), "&gt;")
      
      tipopedido = Replace(Trim(tipopedido), Chr(34), "&quot;")
      tipopedido = Replace(Trim(tipopedido), Chr(38), "&amp;")
      tipopedido = Replace(Trim(tipopedido), Chr(39), "&apos;")
      tipopedido = Replace(Trim(tipopedido), Chr(60), "&lt;")
      tipopedido = Replace(Trim(tipopedido), Chr(62), "&gt;")
      
      MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
      MyBuffer = MyBuffer & " GDe = " & Chr(34) & GrDespacho & Chr(34)
      MyBuffer = MyBuffer & " Pro = " & Chr(34) & proveedor & Chr(34)
      MyBuffer = MyBuffer & " Fec = " & Chr(34) & Fecha & Chr(34)
      MyBuffer = MyBuffer & " TiP = " & Chr(34) & tipopedido & Chr(34)

      MyBuffer = MyBuffer & "/>"
   
   End If
   
   DoEvents
           
   RsExcel.MoveNext
   
Loop
        
MyBuffer = MyBuffer & "</RutaDes>"

RsExcel.Close
Set RsExcel = Nothing
    
cn.Close
Set cn = Nothing

If EstPro Then
              
      'registrar Log sistema
      If OpMenu = 1 Then
      
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Eliminar"), Me.HelpContextID, "", "", "")
         
      Else
      
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Agregar"), Me.HelpContextID, "", "", "")
      
      End If
              
      Set RS = vg_db.Execute("sgpadm_Del_XmlRutasGrupoDespacho '" & MyBuffer & "', '" & IIf(OpMenu = 1, "Manteneción Rutas Grupo Despacho - Eliminación Ruta Grupo Despacho", "Manteneción Rutas Grupo Despacho - Agregar Rutas Grupo Despacho") & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "', '" & OpMenu & "'")

      If Not RS.EOF Then

         If RS(0) > 0 Then

            MsgBox RS(1)
            fg_descarga

            RS.Close
            Set RS = Nothing
   
            'registrar Log sistema eliminar & Agregado
      
            If OpMenu = 1 Then
            
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminado"), Me.HelpContextID, "", "", "")
               
            Else
            
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Agregado"), Me.HelpContextID, "", "", "")
               
            End If
            
            Exit Sub
      
         Else
   
            MsgBox RS(1) & VgLinea & VgLinea & "            Adjunto archivo con error", vbCritical, MsgTitulo
            
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
    
            xlApp.Columns("A:B").Select
            xlApp.Selection.Delete Shift:=xlToLeft
  
            NomArchivoExcel = fg_ArchivoXls(IIf(OpMenu = 1, "ReporteError_RutasGrupoDespachosEliminar", "ReporteError_RutasGrupoDespachos_Agregar"))
                    
            xlWb.Close True, NomArchivoExcel

            Dim XL As New excel.Application 'Crea el objeto excel
            XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
            XL.Visible = True
            XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
            '-- Cerrar Excel
            xlApp.Quit
            '-------> Release Excel references
            Set xlWs = Nothing
            Set xlWb = Nothing
            Set xlApp = Nothing
   
            'registrar Log sistema eliminar & Agregado
      
            If OpMenu = 1 Then
            
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, "", "", "")
               
            Else
            
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Agregado"), Me.HelpContextID, "", "", "")
               
            End If
         
         End If
      
      End If
      RS.Close
      Set RS = Nothing

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

If (Col <> 7 And Col <> 8 And Col <> 9 And Col <> 10 And Col <> 11 And Col <> 12 And Col <> 13) Or Row = 0 Or estdes Then
   
   Exit Sub

End If

vaSpread1.Row = Row
vaSpread1.Col = 5

If vaSpread1.text <> 1 Then
   
   For i = 7 To 13
       
       If i <> Col Then
          
          estdes = True
          vaSpread1.Col = i
          vaSpread1.text = "0"
          vaSpread1.TypeHAlign = TypeHAlignCenter
          estdes = False
       
       End If
   
   Next i

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1
    
    Dim i As Long
    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    
    Next
   
End Select
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error
  
Dim codigo As Integer

Dim findessemana As String


If vaSpread1.MaxRows < 1 Then Exit Sub

If (Col) = 5 Then
            
            modo = "M"
            vaSpread1.SetFocus
            vaSpread1.Row = vaSpread1.ActiveRow
            
            vaSpread1.Col = 5
            codigo = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
              
            vaSpread1.Col = 14
            findessemana = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
               
            vaSpread1.Col = 3
            codigoceco = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
            
            ' Rescata el Parametro correspondiente
            
            If codigo = 1 Then
              vaSpread1.Col = 6 ' Periodo
              vaSpread1.text = "SEMANAL"
            End If
            
            If codigo = 2 Then
              vaSpread1.Col = 6 ' Periodo
              vaSpread1.text = "QUINCENAL"
            End If
            
            If codigo > 2 Then
              vaSpread1.Col = 6 ' Distintos Periodo
              vaSpread1.text = "CADA " + CStr(codigo) + " SEMANAS"
            End If
            
            If codigo = 0 Then
                
                If codigo_anterior = 1 Then
                    vaSpread1.Col = 6 ' Periodo
                    vaSpread1.text = "SEMANAL"
                End If
                
                If codigo_anterior = 2 Then
                    vaSpread1.Col = 6 ' Periodo
                    vaSpread1.text = "QUINCENAL"
                End If
                
                If codigo_anterior > 2 Then
                    vaSpread1.Col = 6 ' Periodo
                    vaSpread1.text = "CADA " + CStr(codigo_anterior) + " SEMANAS"
                End If
                
                vaSpread1.Col = 5 ' Periodo
                vaSpread1.text = codigo_anterior
                 
            End If
       
            estdes = True
       
     If codigo > 0 Then
      If codigo <> codigo_anterior Then
            
            With vaSpread1
             
              .Col = 7:  .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter:  .Lock = False
              .Col = 8:  .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter:  .Lock = False
              .Col = 9:  .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter:  .Lock = False
              .Col = 10:  .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter: .Lock = False
              .Col = 11: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter:  .Lock = False
              .Col = 12: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter:  .Lock = False
              .Col = 13: .CellType = CellTypeCheckBox: .text = "0": .TypeHAlign = TypeHAlignCenter:  .Lock = False
              
              ' Rescata la Ultima Fecha del Cecos
               
               Dim valor As String
           
               If fecha_parametro <> "" Then
                  Sql = " sgpadm_sel_maximafechadesdedelcecos " & codigoceco
                  Set RS = vg_db.Execute(Sql)
                  If Not RS.EOF Then
                    If IsNull(RS(0)) Then
                       vaSpread1.Col = 15 ' Fecha Hastas
                       vaSpread1.Lock = False
                       vaSpread1.CellType = CellTypeDate
                       vaSpread1.TypeCurrencyMin = fecha_parametro
                       vaSpread1.text = Mid(fecha_parametro, 7, 2) + "/" + Mid(fecha_parametro, 5, 2) + "/" + Mid(fecha_parametro, 1, 4)
                       vaSpread1.SetFocus
                       
                       vaSpread1.Col = 16
                       vaSpread1.text = "1"
                       
                    Else
                    
                       vaSpread1.Col = 15 ' Fecha Hastas
                       vaSpread1.Lock = False
                       vaSpread1.CellType = CellTypeDate
                       vaSpread1.text = Format(RS(0), "DD/MM/YyyY")
                       vaSpread1.SetFocus
                       
                       vaSpread1.Col = 16
                       vaSpread1.text = "1"
                     
                     End If
                     
                 End If
                 
                 RS.Close
                 
               End If
             
             End With
          
        End If
       
       End If
      
             estdes = False
     
     If findessemana = "0" Then
     
         With vaSpread1
              
              .Row = vaSpread1.ActiveRow
              .Col = 12: .Lock = True
              .Col = 13: .Lock = True
           
           End With
     
     End If
     
 
End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim Ceco       As String
Dim NombreCeco As String
Dim RS         As New ADODB.Recordset

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.ActiveCol

If Row = 0 Or vaSpread1.MaxRows < 1 Then
      
      Toolbar1.Buttons(1).Visible = False
      Toolbar1.Buttons(2).Visible = True
      Exit Sub

End If

If vaSpread1.Lock = False Then

      If Col = 15 Then
         
         modo = "M"
      
      End If
      
      vaSpread1.Row = vaSpread1.ActiveRow
      vaSpread1.Col = 5
      codigo_anterior = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)

  
  
  If (Col = 7 Or Col = 8 Or Col = 9 Or Col = 10 Or Col = 11 Or Col = 12 Or Col = 13) Then
      
      modo = "M"
      
      Toolbar1.Buttons(1).Visible = True
      Toolbar1.Buttons(2).Visible = False
     
      vaSpread1.Col = 14
      findessemana = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
         
      vaSpread1.Col = 3
      codigoceco = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
            
      If fecha_parametro <> "" Then
       
       Sql = " sgpadm_sel_maximafechadesdedelcecos " & codigoceco
       Set RS = vg_db.Execute(Sql)
          
          If IsNull(RS(0)) Then
               
               vaSpread1.Col = 15 ' Fecha Hastas
               vaSpread1.Lock = False
               vaSpread1.CellType = CellTypeDate
               vaSpread1.TypeDateMax = Mid(fecha_parametro, 5, 2) & Mid(fecha_parametro, 7, 2) & Mid(fecha_parametro, 1, 4)
               vaSpread1.text = Mid(fecha_parametro, 7, 2) + "/" + Mid(fecha_parametro, 5, 2) + "/" + Mid(fecha_parametro, 1, 4)
               vaSpread1.SetFocus
               
               vaSpread1.Row = vaSpread1.ActiveRow
               vaSpread1.Col = 16
               vaSpread1.text = "1"
            
            Else
            
               vaSpread1.Col = 15 ' Fecha Hastas
               vaSpread1.Lock = False
               vaSpread1.CellType = CellTypeDate
               vaSpread1.TypeDateMin = Format(RS(0), "MMddyyyy")
               vaSpread1.TypeDateMax = Mid(fecha_parametro, 5, 2) & Mid(fecha_parametro, 7, 2) & Mid(fecha_parametro, 1, 4)
               vaSpread1.text = Format(RS(0), "DD/MM/YyyY")
               vaSpread1.SetFocus
                              
               vaSpread1.Row = vaSpread1.ActiveRow
               vaSpread1.Col = 16
               vaSpread1.text = "1"
             
             End If
                
            RS.Close
      Else
         
         modo = ""
      
      End If
      
      If codigo_anterior <> 1 Then
          
          estdes = False
      
      End If
  
  End If

End If

If Col = 18 Then

   vaSpread1.Row = Row
   
   '--> Validar Acceso
   If Mid(ValidarUsuarioAcceso(1196005, vg_NUsr), 1, 1) <> "1" Then
       
      MsgBox "No tiene acceso Pedido - Parametro Grupo Despacho Ceco", vbInformation, MsgTitulo
      Exit Sub
       
   End If
   
   '--> validar que sitio este parametrizado
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    vaSpread1.Col = 3
    codigoceco = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
    
    Sql = "sgpadm_Sel_TraerDiaParametroCD '" & codigoceco & "'"
    Set RS = vg_db.Execute(Sql)

    If RS.EOF Then

       RS.Close
       Set RS = Nothing
       MsgBox "Ceco no esta parametrizado. cancelado ingreso parametro grupo despacho", vbCritical, MsgTitulo
       Exit Sub
       
    End If
    
    RS.Close
    Set RS = Nothing
   
   
   vaSpread1.Col = 15
   If Trim(vaSpread1.text) <> "" Then
   
      MsgBox "Debe grabar los cambios antes de acceder a esta opción. Proceso cancelado ", vbCritical, MsgTitulo
      Exit Sub
   
   End If
   
   vaSpread1.Col = 3
   Ceco = vaSpread1.text
   
   vaSpread1.Col = 4
   NombreCeco = vaSpread1.text
   
   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(1196005), "", "", "")
   Call M_FechaGrupoDespachoCeco.Moverdetalle(Ceco, NombreCeco)
   Call M_FechaGrupoDespachoCeco.Show(1)

   Sql = "    sgpadm_sel_paramdespachocasino '" & Ceco & "'"
   Set RS = vg_db.Execute(Sql)
         
   If Not RS.EOF Then
   
      vaSpread2.Row = 1
      vaSpread2.Col = IIf(RS!Parametrizado = "0", 3, IIf(RS!Parametrizado = RS!NGrupoDespacho, 2, 1))

      vaSpread1.Col = 18
      vaSpread1.TypePictPicture = vaSpread2.TypePictPicture
   
   
   End If
 
   RS.Close
   Set RS = Nothing
 
 
End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(1).Visible = True Then
   
   Call validar

End If

End Sub

Private Sub validar()

On Error GoTo Man_Error

Dim Fecha As Date
Dim fechaproceso As String
Dim fechahasta As String

vaSpread1.Row = vaSpread1.ActiveRow

vaSpread1.Col = 15 'Lunes
Fecha = IIf(vaSpread1.text = "", 0, vaSpread1.text)
fechaproceso = Format(Fecha, "YYYYMMDD")
fechahasta = fecha_parametro

If fechahasta <> "" Then

 If fechaproceso > fechahasta Then
      
      vaSpread1.Col = 16
      vaSpread1.text = "1"
      SW_VALIDACION = 1
      MsgBox "La fecha desde no puede ser mayor a la Ultima fecha generada " & fpText1.text, vbExclamation
      vaSpread1.TypeHAlign = TypeHAlignCenter
  
      vaSpread1.SetActiveCell 15, vaSpread1.ActiveRow
      Exit Sub
  
 End If
 
 
' sql = " sgpadm_sel_ExistenPedidoentrerangodeFecha " & "'" & codigoceco & "'," & "'" & fechaproceso & "'"
' Set RS = vg_db.Execute(sql)
' If Not RS.EOF Then
'    If RS(0) > 0 Then
'      vaSpread1.Col = 15
'      vaSpread1.text = "1"
'      SW_VALIDACION = 1
'      MsgBox "Existen pedidos confirmados en esta fecha", vbExclamation
'      vaSpread1.SetActiveCell 14, vaSpread1.ActiveRow
'      Exit Sub
'
'
'  End If
'  End If
  
 'sql = " sgpadm_sel_maximafechadesdedelcecospedido " & "'" & codigoceco & "'"
 'Set RS = vg_db.Execute(sql)
 'If Not RS.EOF Then
 '
 ' If Not IsNull(RS(0)) Then
 '  fechahasta = IIf(IsNull(RS(0)), Format(Now, "YYYYMMDD"), Format(RS(0), "YYYYMMDD"))
'
'  If fechaproceso < fechahasta Then
'      vaSpread1.Col = 15
'      vaSpread1.text = "1"
'      SW_VALIDACION = 1
'      MsgBox "La Fecha desde no puede ser menor que la fecha del ultimo pedido confirmado ", vbExclamation
'      vaSpread1.SetActiveCell 14, vaSpread1.ActiveRow
'     ' Toolbar1.Buttons(1).Enabled = False
'   Else
'      modo = ""
'      vaSpread1.Col = 15 ' Periodo
'      vaSpread1.text = "0"
'      SW_VALIDACION = 0
'      vaSpread1.TypeHAlign = TypeHAlignCenter
'     ' Toolbar1.Buttons(1).Enabled = True
'    End If
'    Else
'     modo = ""
'      vaSpread1.Col = 15 ' Periodo
'      vaSpread1.text = "0"
'      SW_VALIDACION = 0
'      vaSpread1.TypeHAlign = TypeHAlignCenter
     ' Toolbar1.Buttons(1).Enabled = True
'  End If
' End If

End If
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
End Sub

Function ValidarRutasxGrupo() As Boolean

Dim RS            As New ADODB.Recordset
Dim Sql           As String
Dim EstDia        As Boolean
Dim fechaavalidar As String
Dim fechamodif    As Long
Dim actualiza     As Long
Dim xmlfamilia    As String
Dim i             As Long
Dim codigoceco    As String
Dim FechaMenor    As Long
Dim Fecha         As Date

Dim xlApp As Object
Dim xlWb As Object
Dim xlWs As Object

ValidarRutasxGrupo = True

EstDia = False
fechaavalidar = ""
actualiza = 0
xmlfamilia = ""
xmlfamilia = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlfamilia = xmlfamilia & "<GrabaRutaCeco>"
    
For i = 1 To vaSpread1.MaxRows
    
    vaSpread1.Row = i
    vaSpread1.Col = 15 'Marca Actualiza
    fechaavalidar = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
    vaSpread1.Col = 16 'Marca Actualiza
    actualiza = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
    If fechaavalidar <> "0" And actualiza = 1 Then
        
       vaSpread1.Col = 3  'ceco
       codigoceco = vaSpread1.text
        
       vaSpread1.Col = 15 'fecha
       fechamodif = Format(vaSpread1.text, "YYYYMMDD")
          
'       If Format(vaSpread1.text, "YYYYMMDD") < FechaMenor Or FechaMenor = 0 Then
'
'          FechaMenor = Format(vaSpread1.text, "YYYYMMDD")
'          Fecha = vaSpread1.text
'
'       End If
       
       xmlfamilia = xmlfamilia & " <RutaCeco"
       xmlfamilia = xmlfamilia & " Ceco = " & Chr(34) & codigoceco & Chr(34)
       xmlfamilia = xmlfamilia & " Fec = " & Chr(34) & fechamodif & Chr(34)
       xmlfamilia = xmlfamilia & "/>"
    
    End If
        
Next i
    
xmlfamilia = xmlfamilia & "</GrabaRutaCeco>"
          
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
          
Sql = ""
Sql = "sgpadm_Sel_XmlValidasiExisteRutaGrupo "
Sql = Sql & " '" & xmlfamilia & "'"
         
Set RS = vg_db.Execute(Sql)
If Not RS.EOF Then
            
   If RS(0) > 0 And RS("estado") <> "0" Then
               
      fg_descarga
'      ValidarRutasxGrupo = False
'      ' Create an instance of Excel and add a workbook
'      Set xlApp = CreateObject("Excel.Application")
'      Set xlWb = xlApp.Workbooks.Add
'      Set xlWs = xlWb.Worksheets("Hoja1")
'
'      If RS.RecordCount > xlWs.Range("A1", xlWs.Range("A1").End(xlDown)).Rows.count Then
'
'         ' Close ADO objects
'         RS.Close
'         Set RS = Nothing
'
'         ' Release Excel references
'         Set xlWs = Nothing
'         Set xlWb = Nothing
'
'         Set xlApp = Nothing
'
'         MsgBox "Excede numero filas, debera bajar la fecha despacho", vbCritical
'
'         Exit Function
'
'      End If
'
'      MsgBox "Existen rutas normales en la fechas solicitadas por pantalla. se generara una planilla excel con las rutas que impiden generar las nuevas rutas grupo despacho. Pasos seguir es borrar las rutas normales", vbCritical + vbOKOnly, MsgTitulo
'
'      ' Display Excel and give user control of Excel's lifetime
'      xlApp.Visible = True
'      xlApp.UserControl = True
'
'      ' Check version of Excel
'      Call encabezado(RS, xlWs)
'
'      xlWs.Cells(2, 1).CopyFromRecordset RS
'
'      ' Auto-fit the column widths and row heights
'      xlApp.Selection.CurrentRegion.Columns.AutoFit
'      xlApp.Selection.CurrentRegion.Rows.AutoFit
'
'      ' Release Excel references
'      Set xlWs = Nothing
'      Set xlWb = Nothing
'
'      Set xlApp = Nothing
          
      If MsgBox("Existen rutas x grupo despacho en algunos sitios. Para borrar estas rutas va considerar el item de la grilla FECHA DESDE como fecha inicio a borrar. Desea borrar las rutas x grupo despacho a los sitios S/N???", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
      
         vg_BorradoDatos = False
         ValidarRutasxGrupo = False
         Exit Function
      Else
      
         vg_BorradoDatos = True
      
      End If
          
      Exit Function
            
   End If
         
End If
          
' Close ADO objects
RS.Close
Set RS = Nothing
    
Exit Function
Man_Error:

ValidarRutasxGrupo = False
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Function

