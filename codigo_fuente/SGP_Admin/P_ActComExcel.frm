VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form P_ActComExcel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Comensales Desde Excel Minuta Bloque"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14790
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   10935
      Begin VB.CheckBox ValidarRacPon 
         Caption         =   "Aplica Actualización Raciones y Ponderaciones"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1875
         TabIndex        =   14
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox Ceco 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1875
         TabIndex        =   8
         Top             =   800
         Width           =   8295
      End
      Begin VB.CheckBox ValidaReceta 
         Caption         =   "Aplica Actualización Receta"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7035
         TabIndex        =   7
         Top             =   1200
         Width           =   3375
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1875
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   8535
         _Version        =   196608
         _ExtentX        =   15055
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
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
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
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   1
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
         ControlType     =   3
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
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   120
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ProgressBar prbStatus 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   10590
         _ExtentX        =   18680
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "lblStatus"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Archivo Origen"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   12
         Top             =   435
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ceco "
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   510
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   13935
      Begin FPSpread.vaSpread vaSpread3 
         Height          =   375
         Left            =   8520
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
         _Version        =   393216
         _ExtentX        =   2778
         _ExtentY        =   661
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
         MaxCols         =   1
         MaxRows         =   0
         SpreadDesigner  =   "P_ActComExcel.frx":0000
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   1455
         Left            =   2760
         TabIndex        =   5
         Top             =   -960
         Visible         =   0   'False
         Width           =   4575
         _Version        =   393216
         _ExtentX        =   8070
         _ExtentY        =   2566
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
         MaxCols         =   4
         MaxRows         =   2
         SpreadDesigner  =   "P_ActComExcel.frx":01E2
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4335
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   13575
         _Version        =   393216
         _ExtentX        =   23945
         _ExtentY        =   7646
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
         SpreadDesigner  =   "P_ActComExcel.frx":0484
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "P_ActComExcel.frx":1F1A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   1200
         Top             =   4800
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   720
         Top             =   4800
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label1 
         Caption         =   "Proceso Error"
         Height          =   255
         Index           =   4
         Left            =   12000
         TabIndex        =   4
         Top             =   4845
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Proceso OK"
         Height          =   255
         Index           =   3
         Left            =   10275
         TabIndex        =   3
         Top             =   4845
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   2
         Left            =   11520
         Picture         =   "P_ActComExcel.frx":22B4
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   1
         Left            =   9840
         Picture         =   "P_ActComExcel.frx":853E
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   360
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H80000018&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   240
         Top             =   4800
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7770
      Left            =   14160
      TabIndex        =   0
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   13705
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "P_ActComExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Option Compare Text

Private obj_Excel       As Object
Private obj_Workbook    As Object
Private obj_Worksheet   As Object

Dim MsgTitulo As String
Dim EstCargarRecetas As Boolean

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me
fg_carga ""
MsgTitulo = "Actualizar Comensales Desde Excel Minuta Bloque"

EstCargarRecetas = False
vaSpread1.MaxRows = 0
vaSpread1.BackColor = Shape1(0).FillColor
lblStatus.Visible = False: prbStatus.Visible = False

Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar ": BtnX.Enabled = True
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Toolbar1.Buttons(1).Enabled = False

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpText1_ButtonHit(Button As Integer, NewIndex As Integer)

On Error GoTo Man_Error

Dim fromRihgt  As String
Dim ClaveExcel As String
Dim RS         As New ADODB.Recordset
Dim i          As Long
Dim j          As Long
Dim NomArchivo As String

'Traer clave excel
ClaveExcel = "Jp123456"

Set RS = vg_db.Execute("sgpadm_s_parametro 1, 'parhojaexc', ''")
If Not RS.EOF Then
   ClaveExcel = RS(0)
End If
RS.Close
Set RS = Nothing

vaSpread1.MaxRows = 0

Ceco.text = ""

CD.Filter = "Archivos xls|*.xls|Archivos xlsx|*.xlsx"
CD.DefaultExt = "*.xls|*.xlsx"
CD.FilterIndex = 2
CD.Flags = cdlOFNFileMustExist
CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
CD.ShowOpen

If CD.FileName = "" Then
   
   Toolbar1.Buttons(1).Enabled = False
   fpText1.text = ""

Else
    
    fromRight = InStrRev(CD.FileName, "\", , vbTextCompare)
    
    If fromRight > 1 Then
       
       myPath = Left(CD.FileName, fromRight)
    
    End If

    fpText1.text = CD.FileName 'Dir(CD.FileName)
    
    NomArchivo = Dir(CD.FileName)

    Dim ApExcel As excel.Application

    'Al configurarlo
    Set ApExcel = New excel.Application

    'Al abrirlo
    ApExcel.Workbooks.Open FileName:=fpText1.text

    Dim HojaPro As Boolean
    
    ValidaReceta.Enabled = False
    ValidaReceta.Value = 0
    
    For i = 1 To ApExcel.Sheets.count
 
        'Validar si la hoja viene protegida
        
        HojaPro = False

        If ApExcel.Sheets(i).ProtectContents Then HojaPro = True
        If ApExcel.Sheets(i).ProtectDrawingObjects Then HojaPro = True
        If ApExcel.Sheets(i).ProtectScenarios Then HojaPro = True
        
        If HojaPro Then
           
           ApExcel.Sheets(i).Unprotect ClaveExcel
           If ApExcel.Sheets(i).ProtectContents Then HojaPro = False
        
        End If

        'Validar hoja corresponda la clave
        If (ClaveExcel = ApExcel.Sheets(i).Range("E1").Value Or ClaveExcel = ApExcel.Sheets(i).Range("G1").Value) And HojaPro Then
           
           vaSpread1.MaxRows = vaSpread1.MaxRows + 1
           vaSpread1.Row = i
        
           vaSpread1.Col = 1
           vaSpread1.text = ""
        
           vaSpread1.Col = 2
           vaSpread1.text = ApExcel.Sheets(i).Name
           
           vaSpread1.Col = 3
           vaSpread1.text = IIf(ClaveExcel = ApExcel.Sheets(i).Range("E1").Value, ApExcel.Sheets(i).Range("E2").Value, ApExcel.Sheets(i).Range("G2").Value)
           
           vaSpread1.Col = 4
           vaSpread1.text = ""
        
           vaSpread1.Col = 5
           vaSpread1.text = ""
        
           vaSpread1.Col = 6
           vaSpread1.text = ""
        
           'formato antiguo - nuevo
           vaSpread1.Col = 7
'           vaSpread1.text = IIf(ClaveExcel = ObjW.Sheets(i).Range("E1").Value, 1, IIf(ClaveExcel = ObjW.Sheets(i).Range("G1").Value, 2, 0))
           vaSpread1.text = IIf(ClaveExcel = ApExcel.Sheets(i).Range("E1").Value And Trim(ApExcel.Sheets(i).Range("B1").Value) = "", 1, IIf(ClaveExcel = ApExcel.Sheets(i).Range("G1").Value And Trim(ApExcel.Sheets(i).Range("B1").Value) = "", 2, IIf(ClaveExcel = ApExcel.Sheets(i).Range("G1").Value And Trim(ApExcel.Sheets(i).Range("B1").Value) = "*", 3, 0)))
        
           Ceco.text = Mid(ApExcel.Sheets(i).Range("a1").Value, 9, Len(ApExcel.Sheets(i).Range("a1").Value))
           
        ElseIf ApExcel.Sheets(i).Name = "Recetas" Then
        
            ValidaReceta.Enabled = True
            ValidaReceta.Value = 0
        
        End If
    
    Next
    
    ApExcel.Workbooks(NomArchivo).Close SaveChanges:=False

    ApExcel.Visible = False
    ApExcel.Application.Visible = False
    ApExcel.Application.Quit
    Set ApExcel = Nothing
    
    If vaSpread1.MaxRows > 0 Then
    
        Toolbar1.Buttons(1).Enabled = True
    
    Else
        
        Toolbar1.Buttons(1).Enabled = False
        
        MsgBox "El formato excel no es indicado, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
    End If
    
End If

Exit Sub
Man_Error:
fg_descarga
If Err = 5 Or Err = 0 Then Resume Next: Exit Sub
If Err = 55 Then MsgBox "Archivo esta abierto, proceso cancelado": Exit Sub
If Err = 462 Or Err = 1004 Or Err = 438 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Resume

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim indsel      As Boolean
Dim i           As Long
Dim SheetName   As String
Dim MaxColumnas As Long
Dim TipoFormato As Integer
Dim abrirexcel  As Boolean
Dim PathXls     As String
Dim NomArchivo  As String

Dim ApExcel As New excel.Application
Set ApExcel = New excel.Application

abrirexcel = True

Select Case Button.Index
Case 1
    
    vaSpread1.Row = -1
    vaSpread1.Col = -1
    vaSpread1.BackColor = Shape1(0).FillColor
    
    'Valida archivo excel
    If Trim(fpText1.text) = "" Then
       
       MsgBox "Debe seleccionar archivo", vbInformation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    'limpia leyenda
    For i = 1 To vaSpread1.MaxRows
    
        vaSpread1.Col = 4
        vaSpread1.text = ""
        
        vaSpread1.Col = 5
        vaSpread1.text = ""
        
        vaSpread1.Col = 6
        vaSpread1.text = ""

    Next i
    
    'Valida selección de grilla
    indsel = False
    For i = 1 To vaSpread1.MaxRows
    
        vaSpread1.Row = i
        vaSpread1.Col = 1
        
        If vaSpread1.text = "1" Then
           
            indsel = True
            Exit For
    
        End If
    
    Next i
    
    If Not indsel Then
    
       MsgBox "Debe seleccionar al menos una hoja de la grilla", vbInformation + vbOKOnly, MsgTitulo
       
       Exit Sub
       
    End If
    
    If ValidarRacPon.Value = 0 And ValidaReceta = 0 Then
    
       MsgBox "Debe seleccionar al menos una opción de actualización", vbInformation + vbOKOnly, MsgTitulo
       
       Exit Sub
       
    End If
    
    
    'Cargar Recetas
    'If ValidaReceta.Value = 1 Then
    '
    '   CargarRecetaGrilla
    '
    'End If
    
    'RutinaActualizarComensalesDiarios y totales
    
    indsel = False
    
    For i = 1 To vaSpread1.MaxRows
    
        vaSpread1.Row = i
        vaSpread1.Col = 1
        
        If vaSpread1.text = "1" Then
           
           vaSpread1.Col = 2
           SheetName = Trim(vaSpread1.text) & "$"
           
           vaSpread1.Col = 3
           MaxColumnas = Val(vaSpread1.text)
           
           vaSpread1.Col = 7
           TipoFormato = Val(vaSpread1.text)
           
           
           lblStatus.Caption = Trim(vaSpread1.text)
           
           If TipoFormato = 1 Then
              
              ActualizarComensalesDiariosTotales SheetName, MaxColumnas, i
           
           ElseIf TipoFormato = 2 Then
           
'              ActualizarComensalesDiariosTotalesFormatonuevo SheetName, IIf(MaxColumnas < 187, 187, MaxColumnas), i
              If abrirexcel Then
              
                 'cs = "DRIVER=Microsoft Excel Driver (*.xls);DBQ=" & PathXls & ";HDR=NO;IMEX=1;"
                 PathXls = Trim(fpText1.text)
                 NomArchivo = Dir(CD.FileName)
'                 Set obj_Excel = CreateObject("Excel.Application")
 
                 ' -- Abrir el libro
'                 Set obj_Workbook = obj_Excel.Workbooks.Open(PathXls)
                  ApExcel.Workbooks.Open FileName:=PathXls
                 
                 abrirexcel = False
              
              End If
              
              ActualizarComensalesDiariosTotalesFormatonuevo SheetName, MaxColumnas, i
           
           ElseIf TipoFormato = 3 Then
           
              If abrirexcel Then
              
                 'cs = "DRIVER=Microsoft Excel Driver (*.xls);DBQ=" & PathXls & ";HDR=NO;IMEX=1;"
                 PathXls = Trim(fpText1.text)
                 NomArchivo = Dir(CD.FileName)
'                 Set obj_Excel = CreateObject("Excel.Application")
 
                 ' -- Abrir el libro
'                 Set obj_Workbook = obj_Excel.Workbooks.Open(PathXls)
                  ApExcel.Workbooks.Open FileName:=PathXls
                 
                 abrirexcel = False
              
              End If
              
              ActualizarComensalesDiariosTotalesFormatonuevoII SheetName, MaxColumnas, i
           
           End If
    
           indsel = True
           
        End If
    
    Next i
    
    If indsel Then
    
       If TipoFormato = 2 Or TipoFormato = 3 Then
       
        ApExcel.Workbooks(NomArchivo).Close SaveChanges:=False

        ApExcel.Visible = False
        ApExcel.Application.Visible = False
        ApExcel.Application.Quit
        Set ApExcel = Nothing
       
''          obj_Excel.Quit
'          Set obj_Excel = Nothing
'       '   obj_Workbook("PathXls").Close SaveChanges:=False
'          obj_Workbook.Close SaveChanges:=False
'          Set obj_Workbook = Nothing
'          Set obj_Worksheet = Nothing
'
'          'obj_Workbook.Close
'
       End If
       
       MsgBox "Actualización finalizado..", vbInformation + vbOKOnly, MsgTitulo
       
    End If
    
    fg_descarga

Case 3
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
If Err = 5 Or Err = 462 Or Err = 430 Or Err = -2147023170 Or Err = -2147417848 Then Resume Next: Exit Sub
If Err = 55 Then MsgBox "Archivo esta abierto, proceso cancelado": Exit Sub
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Resume

End Sub

Sub ActualizarComensalesDiariosTotales(SheetName As String, MaxColumnas As Long, Row As Long)

On Error GoTo Man_Error

Dim i              As Long
Dim j              As Long
Dim PathXls        As String
Dim dbexcel        As Database
Dim cn             As ADODB.Connection
Dim RS             As New ADODB.Recordset
Dim IndColumna     As Long
Dim MyBuffer       As String
Dim MyBufferReceta As String
Dim MyBufferTotal  As String

Dim strArray()     As String
Dim intCount       As Integer
Dim UltRow         As Long

Dim Ceco           As String
Dim Regimen        As Long
Dim Servicio       As Long
Dim RecetaOri      As Long
Dim RecetaDes      As Long
Dim Fecha          As Long
Dim NumLin         As Long
Dim Raciones       As Long
Dim PorcentajePon  As Long
Dim ComesalesTot   As Long
Dim NombreReceta   As String
Dim NomRow         As Long
Dim EstGrabado     As Boolean
Dim EstActReceta   As Boolean
Dim lngRow         As Long
Dim CodigoReceta   As Long
Dim ResultadoPor   As String

EstActReceta = False

EstGrabado = True
Set RsExcel = New ADODB.Recordset
Set cn = New ADODB.Connection

PathXls = Trim(fpText1.text)
File_Ext = UCase(Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, "."))))

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

RsExcel.Open ("SELECT * FROM [" & SheetName & "]"), cn

If RsExcel.EOF Then Exit Sub

RsExcel.MoveFirst

IndColumna = 4
i = 1

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<UpdateRacPon>"

Let MyBufferReceta = ""
Let MyBufferReceta = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBufferReceta = MyBufferReceta & "<UpdateRacPon>"

Let MyBufferTotal = ""
Let MyBufferTotal = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBufferTotal = MyBufferTotal & "<UpdateComensales>"

prbStatus.Max = 1
lblStatus.Visible = True: prbStatus.Visible = True: prbStatus.Min = 0: lngRow = 0
prbStatus.Max = RsExcel.RecordCount

lblStatus.Caption = Mid(SheetName, 1, Len(SheetName) - 1)

Do While RsExcel.EOF <> True
           
   If RsExcel.Fields(0).Value = "*" Then Exit Do
           
   'Comensales Totales
   If RsExcel.Fields(0).Value = "Comensales" Then
         
        For j = IndColumna To MaxColumnas Step 4
       
            If RsExcel.Fields(j).Value <> "" Then
                
               MyBufferTotal = MyBufferTotal & " <Comensales"
        
               strTest = RsExcel.Fields(j).Value
               strArray = Split(strTest, ";")
   
               For intCount = LBound(strArray) To UBound(strArray)
                    
                    Select Case intCount
                        
                        Case 0
                            Ceco = Trim(strArray(intCount))
                        Case 1
                            Regimen = Trim(strArray(intCount))
                        Case 2
                            Servicio = Trim(strArray(intCount))
                        Case 3
                            Fecha = Trim(strArray(intCount))
                    
                    End Select
               
               Next
                
              MyBufferTotal = MyBufferTotal & " Ceco = " & Chr(34) & Ceco & Chr(34)
              MyBufferTotal = MyBufferTotal & " Reg = " & Chr(34) & Regimen & Chr(34)
              MyBufferTotal = MyBufferTotal & " Ser = " & Chr(34) & Servicio & Chr(34)
              MyBufferTotal = MyBufferTotal & " Fec = " & Chr(34) & Fecha & Chr(34)
              
            End If
   
            'Ración
            If IsNumeric(RsExcel.Fields(j - 2).Value) Then
               
               ComensalesTot = RsExcel.Fields(j - 2).Value
               MyBufferTotal = MyBufferTotal & " Com = " & Chr(34) & ComensalesTot & Chr(34)
            
            End If
   
            MyBufferTotal = MyBufferTotal & "/>"
        
        Next j
   
   ElseIf RsExcel.Fields(0).Value <> "" Or i > 4 Then
   
        If i > 4 Then
        For j = IndColumna To MaxColumnas Step 4
       
            If RsExcel.Fields(j).Value <> "" Then
               'RsExcel.Fields(j).ColorIndex
               strTest = RsExcel.Fields(j).Value
               strArray = Split(strTest, ";")
   
               MyBuffer = MyBuffer & " <RacPon"
               
               For intCount = LBound(strArray) To UBound(strArray)
                    
                    Select Case intCount
                        
                        Case 0
                            
                            Ceco = Trim(strArray(intCount))
                            
                            If ValidaReceta.Value = 1 And Not EstCargarRecetas Then
        
                               CargarRecetaGrilla Ceco
                               EstCargarRecetas = True
    
                            End If
                            
                        Case 1
                            
                            Regimen = Trim(strArray(intCount))
                        
                        Case 2
                            
                            Servicio = Trim(strArray(intCount))
                        
                        Case 3
                            
                            If ValidaReceta.Value = 1 Then
                                
                                EstGrabado = True

                                NombreReceta = IIf(IsNull(Trim(RsExcel.Fields(j - 3).Value)), "", Trim(RsExcel.Fields(j - 3).Value))
                                NomRow = 0
                                NomRow = vaSpread2.SearchCol(2, 0, vaSpread2.MaxRows, NombreReceta, SearchFlagsEqual)
                  
                                CodigoReceta = 0
                                
                                If NomRow < 1 Then
                                   
                                   CodigoReceta = SacarNumeroRight(NombreReceta)

                                   NomRow = 0
                                   NomRow = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(CodigoReceta), SearchFlagsEqual)
                                   RecetaDes = 0
                                   
                                   If NomRow < 1 Then
                                      
                                      EstGrabado = False
                                      
                                   Else
                                   
                                       vaSpread2.Row = NomRow
                                       vaSpread2.Col = 1
                                       RecetaDes = vaSpread2.text
                                   
                                   End If
                                
                                Else
                                   
                                   vaSpread2.Row = NomRow
                                   vaSpread2.Col = 1
'                                   RecetaDes = SacarNumeroRight(NombreReceta)
                                   RecetaDes = vaSpread2.text
                                   
                                End If
                            
                            End If
                            
                            RecetaOri = Trim(strArray(intCount))
                            
                            If RecetaDes = 0 Then
                            
                               RecetaDes = RecetaOri
                            
                            End If
                            
                        Case 4
                            
                            Fecha = Trim(strArray(intCount))
                        
                        Case 5
                            
                            NumLin = Trim(strArray(intCount))
                    
                    End Select
               
               Next intCount
                
              MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
              MyBuffer = MyBuffer & " Reg = " & Chr(34) & Regimen & Chr(34)
              MyBuffer = MyBuffer & " Ser = " & Chr(34) & Servicio & Chr(34)
              MyBuffer = MyBuffer & " Rec = " & Chr(34) & RecetaOri & Chr(34)
              
              If ValidaReceta.Value = 1 And RecetaDes <> RecetaOri Then
              
                 MyBufferReceta = MyBufferReceta & " <RacPon"
                 MyBufferReceta = MyBufferReceta & " Ceco = " & Chr(34) & Ceco & Chr(34)
                 MyBufferReceta = MyBufferReceta & " Reg = " & Chr(34) & Regimen & Chr(34)
                 MyBufferReceta = MyBufferReceta & " Ser = " & Chr(34) & Servicio & Chr(34)
                 MyBufferReceta = MyBufferReceta & " Rec = " & Chr(34) & RecetaOri & Chr(34)
                 
                 MyBufferReceta = MyBufferReceta & " RecDes = " & Chr(34) & RecetaDes & Chr(34)
              
                 MyBufferReceta = MyBufferReceta & " Fec = " & Chr(34) & Fecha & Chr(34)
                 MyBufferReceta = MyBufferReceta & " Nli = " & Chr(34) & NumLin & Chr(34)
              
                 MyBufferReceta = MyBufferReceta & " Rac = " & Chr(34) & 0 & Chr(34)
                 
                 MyBufferReceta = MyBufferReceta & " Pon = " & Chr(34) & 0 & Chr(34)
            
                 MyBufferReceta = MyBufferReceta & "/>"
              
                 EstActReceta = True
                 
              End If
              
              MyBuffer = MyBuffer & " Fec = " & Chr(34) & Fecha & Chr(34)
              MyBuffer = MyBuffer & " Nli = " & Chr(34) & NumLin & Chr(34)
               
            'Ración
            If IsNumeric(RsExcel.Fields(j - 2).Value) Then
               
               Raciones = Round(Int(Trim(RsExcel.Fields(j - 2).Value)), 0)
               MyBuffer = MyBuffer & " Rac = " & Chr(34) & Raciones & Chr(34)
            
            End If
   
            'Ponderación
            If IsNumeric(RsExcel.Fields(j - 1).Value) Or Trim(RsExcel.Fields(j - 1).Value) <> "" Then
               
'               If RsExcel.Fields(j - 1).Value < 1 Then
               If IsNumeric(RsExcel.Fields(j - 1).Value) Then
                  
                  PorcentajePon = RsExcel.Fields(j - 1).Value * 100
               
               Else
                  
                  ResultadoPor = ""
                  ResultadoPor = IIf(Not IsNumeric(RsExcel.Fields(j - 1).Value), Replace(RsExcel.Fields(j - 1).Value, "%", ""), 0)
                  
                  If Val(ResultadoPor) > 0 Then
                     
                     PorcentajePon = Replace(RsExcel.Fields(j - 1).Value, "%", "")
                  
                  Else
                     
                     PorcentajePon = 0
                  
                  End If
'                  PorcentajePon = IIf(Not IsNumeric(RsExcel.Fields(j - 1).Value), Replace(RsExcel.Fields(j - 1).Value, "%", ""), 0)
               
               End If
               
               MyBuffer = MyBuffer & " Pon = " & Chr(34) & PorcentajePon & Chr(34)
            
            End If
    
                MyBuffer = MyBuffer & "/>"
                
            End If
   

        Next j
      End If
      End If

   DoEvents
           
   RsExcel.MoveNext
   
   lngRow = lngRow + 1
   prbStatus.Value = lngRow
   
   i = i + 1
        
Loop
        
RsExcel.Close
Set RsExcel = Nothing
    
cn.Close
Set cn = Nothing
    
MyBuffer = MyBuffer & "</UpdateRacPon>"
MyBufferReceta = MyBufferReceta & "</UpdateRacPon>"
MyBufferTotal = MyBufferTotal & "</UpdateComensales>"

'If EstGrabado Then
   
   '--> Actualizar encabezado minuta bloque
   If ValidarRacPon.Value = 1 Then
   
      Set RS = vg_db.Execute("sgpadm_Upd_XmlExcelEncMinutaBloque_V01 '" & MyBufferTotal & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

      If Not RS.EOF Then

         vaSpread1.Row = Row
         vaSpread1.Col = 6
   
         If RS(0) > 0 Then

            vaSpread1.text = "Comensales diarios " & RS(0) & " " & RS(1)
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(2).FillColor
      
            lblStatus.Visible = False
            prbStatus.Visible = False
            fg_descarga

            RS.Close
            Set RS = Nothing
      
            Exit Sub
      
         Else
   
            vaSpread1.text = "Comensales diarios finalizo sin problema"
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(1).FillColor
   
         End If
      
      End If
      RS.Close
      Set RS = Nothing

    End If
       
   'Actualizar detalle minuta bloque
   If ValidarRacPon.Value = 1 Then
      
      Set RS = vg_db.Execute("sgpadm_Upd_XmlExcelDetMinutaBloque_V01 '" & MyBuffer & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")
   
      If Not RS.EOF Then

         vaSpread1.Row = Row
         vaSpread1.Col = 5
   
         If RS(0) > 0 Then

            vaSpread1.text = "Raciones diarias " & RS(0) & " " & RS(1)
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(2).FillColor
      
            lblStatus.Visible = False
            prbStatus.Visible = False
            fg_descarga

            RS.Close
            Set RS = Nothing
      
         Else
   
            vaSpread1.text = "Raciones diarias finalizo sin problema"
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(1).FillColor
      
         End If
   
      End If
      RS.Close
      Set RS = Nothing
   
   End If
   
   '-------> Cambio recetas
   If ValidaReceta = 1 And EstActReceta Then
   
      Set RS = vg_db.Execute("sgpadm_Upd_XmlExcelDetMinutaBloque_V02 '" & MyBufferReceta & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")
   
      If Not RS.EOF Then

         vaSpread1.Row = Row
         vaSpread1.Col = 4
   
         If RS(0) > 0 Then

            vaSpread1.text = "Recetas " & RS(0) & " " & RS(1)
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(2).FillColor
      
            lblStatus.Visible = False
            prbStatus.Visible = False
            fg_descarga

            RS.Close
            Set RS = Nothing
      
        Else
   
            If EstGrabado And EstActReceta Then
               
               vaSpread1.text = "Recetas diarias finalizo sin problema"
               vaSpread1.Col = -1
               vaSpread1.BackColor = Shape1(1).FillColor
            
            ElseIf EstActReceta And Not EstGrabado Then

               vaSpread1.Row = Row
               vaSpread1.Col = 4

               vaSpread1.text = "Realizo algunos cambios de receta. pero receta no concuerda "
               vaSpread1.Col = -1
               vaSpread1.BackColor = Shape1(2).FillColor
      
               lblStatus.Visible = False
               prbStatus.Visible = False
               fg_descarga
            
            ElseIf EstGrabado And Not EstActReceta Then
               
               vaSpread1.text = "no hubo cambio de recetas"
               vaSpread1.Col = -1
               vaSpread1.BackColor = Shape1(1).FillColor
            
            End If
      
        End If
   
      End If
      RS.Close
      Set RS = Nothing
   
   End If
      
'Else

'   vaSpread1.Row = Row
'   vaSpread1.Col = 4

'   vaSpread1.text = "Nombre receta no concuerda "
'   vaSpread1.Col = -1
'   vaSpread1.BackColor = Shape1(2).FillColor
      
'   lblStatus.Visible = False
'   prbStatus.Visible = False
'   fg_descarga

'End If

lblStatus.Visible = False: prbStatus.Visible = False

Exit Sub
Man_Error:
fg_descarga

If Err = 6 Then

    Raciones = 0
    Resume Next

End If

If Err = 3265 Then
      
   If RS.State = 1 Then RS.Close
   MsgBox "Sobrepasa cantidad de columna. maximo deberia ser 254. Proceso cancelado ", vbCritical, MsgTitulo
   Exit Sub
   
End If
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub CargarRecetaGrilla(Ceco As String)

On Error GoTo Man_Error

Dim Sql As String
Dim RS  As New ADODB.Recordset
        
vaSpread2.MaxRows = 0
Set RS = vg_db.Execute("sgpadm_Sel_ExportarExcelRecetas '" & Ceco & "'")

If Not RS.EOF Then
        
   Do While Not RS.EOF

      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      
      vaSpread2.Col = 1
      vaSpread2.text = RS![Código]
      
      vaSpread2.Col = 2
      vaSpread2.text = Trim(RS![Nombre Plato] & " " & RS![Código])
      
      vaSpread2.Col = 3
      vaSpread2.text = RS![Categoria Dietetica]
      
      vaSpread2.Col = 4
      vaSpread2.text = RS![Tipo Plato]
      
      RS.MoveNext
   
   Loop

End If
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub ActualizarComensalesDiariosTotalesFormatonuevo(SheetName As String, MaxColumnas As Long, Row As Long)

On Error GoTo Man_Error

Dim i                   As Long
Dim j                   As Long
Dim PathXls             As String
Dim dbexcel             As Database
Dim cs                  As String
Dim cn                  As ADODB.Connection
Dim RS                  As New ADODB.Recordset
Dim IndColumna          As Long
Dim IndJ                As Long
Dim MyBuffer            As String
Dim MyBufferReceta      As String
Dim MyBufferTotal       As String
Dim MyBufferRecetaVal   As String

Dim strArray()          As String
Dim intCount            As Integer
Dim UltRow              As Long

Dim Ceco                As String
Dim Regimen             As Long
Dim Servicio            As Long
Dim RecetaOri           As Long
Dim RecetaDes           As Long
Dim Fecha               As Long
Dim NumLin              As Long
Dim Raciones            As Long
Dim PorcentajePon       As Long
Dim ComesalesTot        As Long
Dim NombreReceta        As String
Dim NomRow              As Long
Dim EstGrabado          As Boolean
Dim EstActReceta        As Boolean
Dim lngRow              As Long
Dim CodigoReceta        As Long
Dim CodigoRecetaValidar As Long
Dim ResultadoPor        As String

Dim xlWs                As Object

EstActReceta = False

vaSpread3.MaxRows = 0
vaSpread3.maxcols = 1

EstGrabado = True
Set RsExcel = New ADODB.Recordset
Set cn = New ADODB.Connection

PathXls = Trim(fpText1.text)
'File_Ext = UCase(Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, "."))))
'
'With cn
'
'     Select Case File_Ext
'
'        ' Excel 97/2003
'        Case "XLS"
'
'          .Provider = "Microsoft.Jet.OLEDB.4.0"
'          .ConnectionString = "Data Source=" & PathXls & ";" & "Extended Properties=Excel 8.0;"
'          .CursorLocation = 3
'
'        ' Excel 2010
'        Case "XLSX"
'
'          .Provider = "Microsoft.ACE.OLEDB.12.0;Data Source=" & Trim(PathXls) & ";"
'          .ConnectionString = "Extended Properties=Excel 8.0;"
'          .CursorLocation = 3
'
'     End Select
'
'     .Open
'
'End With

'RsExcel.Open ("SELECT * FROM [" & SheetName & "]"), cn

cs = "DRIVER=Microsoft Excel Driver (*.xls);DBQ=" & PathXls & ";HDR=NO;IMEX=1;"
'Set obj_Excel = CreateObject("Excel.Application")
'
'' -- Abrir el libro
'Set obj_Workbook = obj_Excel.Workbooks.Open(PathXls)

RsExcel.CursorLocation = adUseClient
RsExcel.CursorType = adOpenKeyset
RsExcel.LockType = adLockBatchOptimistic

SheetName = "[" & SheetName & "]"
RsExcel.Open "SELECT * FROM  " & SheetName, cs

If RsExcel.EOF Then Exit Sub

RsExcel.MoveFirst

IndColumna = 6
i = 1

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<UpdateRacPon>"

Let MyBufferReceta = ""
Let MyBufferReceta = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBufferReceta = MyBufferReceta & "<UpdateRacPon>"

Let MyBufferTotal = ""
Let MyBufferTotal = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBufferTotal = MyBufferTotal & "<UpdateComensales>"

prbStatus.Max = 1
lblStatus.Visible = True
prbStatus.Visible = True
prbStatus.Min = 0
lngRow = 0

prbStatus.Max = RsExcel.RecordCount

lblStatus.Caption = Mid(SheetName, 1, Len(SheetName) - 1)

Do While RsExcel.EOF <> True
           
   If RsExcel.Fields(0).Value = "*" Then Exit Do
           
   'Comensales Totales
   If RsExcel.Fields(0).Value = "Comensales" Then
         
        For j = IndColumna To MaxColumnas Step 6
       
            If RsExcel.Fields(j).Value <> "" Then
                
               MyBufferTotal = MyBufferTotal & " <Comensales"
        
               strTest = RsExcel.Fields(j).Value
               strArray = Split(strTest, ";")
   
               For intCount = LBound(strArray) To UBound(strArray)
                    
                    Select Case intCount
                        
                        Case 0
                            
                            Ceco = Trim(strArray(intCount))
                        
                        Case 1
                            
                            Regimen = Trim(strArray(intCount))
                        
                        Case 2
                            
                            Servicio = Trim(strArray(intCount))
                        
                        Case 3
                            
                            Fecha = Trim(strArray(intCount))
                    
                    End Select
               
               Next
                
              MyBufferTotal = MyBufferTotal & " Ceco = " & Chr(34) & Ceco & Chr(34)
              MyBufferTotal = MyBufferTotal & " Reg = " & Chr(34) & Regimen & Chr(34)
              MyBufferTotal = MyBufferTotal & " Ser = " & Chr(34) & Servicio & Chr(34)
              MyBufferTotal = MyBufferTotal & " Fec = " & Chr(34) & Fecha & Chr(34)
              
            End If
   
            'Ración
            If IsNumeric(RsExcel.Fields(j - 4).Value) Then
               
               ComensalesTot = RsExcel.Fields(j - 4).Value
               MyBufferTotal = MyBufferTotal & " Com = " & Chr(34) & ComensalesTot & Chr(34)
            
            End If
   
            MyBufferTotal = MyBufferTotal & "/>"
        
        Next j
   
   ElseIf RsExcel.Fields(0).Value <> "" Or i > 7 Then
   
        If i > 7 Then
        For j = IndColumna To MaxColumnas Step 6
       
            If RsExcel.Fields(j).Value <> "" Then
               'RsExcel.Fields(j).ColorIndex
               strTest = RsExcel.Fields(j).Value
               strArray = Split(strTest, ";")
   
               MyBuffer = MyBuffer & " <RacPon"
               
               For intCount = LBound(strArray) To UBound(strArray)
                    
                    Select Case intCount
                        
                        Case 0
                            
                            Ceco = Trim(strArray(intCount))
                            
                            If ValidaReceta.Value = 1 And Not EstCargarRecetas Then
        
                               CargarRecetaGrilla Ceco
                               EstCargarRecetas = True
    
                            End If
                            
                        Case 1
                            
                            Regimen = Trim(strArray(intCount))
                        
                        Case 2
                            
                            Servicio = Trim(strArray(intCount))
                        
                        Case 3
                            
                            If ValidaReceta.Value = 1 Then
                                
                                EstGrabado = True

                                NombreReceta = IIf(IsNull(Trim(RsExcel.Fields(j - 5).Value)), "", Trim(RsExcel.Fields(j - 5).Value))
                                NomRow = 0
                                NomRow = vaSpread2.SearchCol(2, 0, vaSpread2.MaxRows, NombreReceta, SearchFlagsEqual)
                  
                                'Ini : Mover codigo receta
                                
                                CodigoRecetaValidar = 0
                                CodigoRecetaValidar = SacarNumeroRight(NombreReceta)
                                
                                MoverCodigoReceta CodigoRecetaValidar
                                
                                'Fin : Mover codigo receta
                                
                                CodigoReceta = 0
                                
                                If NomRow < 1 Then
                                   
                                   CodigoReceta = SacarNumeroRight(NombreReceta)

                                   NomRow = 0
                                   NomRow = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(CodigoReceta), SearchFlagsEqual)
                                   RecetaDes = 0
                                   
                                   If NomRow < 1 Then
                                      
                                      EstGrabado = False
                                      
                                   Else
                                   
                                       vaSpread2.Row = NomRow
                                       vaSpread2.Col = 1
                                       RecetaDes = vaSpread2.text
                                   
                                   End If
                                
                                Else
                                   
                                   vaSpread2.Row = NomRow
                                   vaSpread2.Col = 1
'                                   RecetaDes = SacarNumeroRight(NombreReceta)
                                   RecetaDes = vaSpread2.text
                                   
                                End If
                            
                            End If
                            
                            RecetaOri = Trim(strArray(intCount))
                            
                            If RecetaDes = 0 Then
                            
                               RecetaDes = RecetaOri
                            
                            End If
                            
                        Case 4
                            
                            Fecha = Trim(strArray(intCount))
                        
                        Case 5
                            
                            NumLin = Trim(strArray(intCount))
                    
                    End Select
               
               Next intCount
                
              MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
              MyBuffer = MyBuffer & " Reg = " & Chr(34) & Regimen & Chr(34)
              MyBuffer = MyBuffer & " Ser = " & Chr(34) & Servicio & Chr(34)
              MyBuffer = MyBuffer & " Rec = " & Chr(34) & RecetaOri & Chr(34)
              
              If ValidaReceta.Value = 1 And RecetaDes <> RecetaOri Then
              
                 MyBufferReceta = MyBufferReceta & " <RacPon"
                 MyBufferReceta = MyBufferReceta & " Ceco = " & Chr(34) & Ceco & Chr(34)
                 MyBufferReceta = MyBufferReceta & " Reg = " & Chr(34) & Regimen & Chr(34)
                 MyBufferReceta = MyBufferReceta & " Ser = " & Chr(34) & Servicio & Chr(34)
                 MyBufferReceta = MyBufferReceta & " Rec = " & Chr(34) & RecetaOri & Chr(34)
                 
                 MyBufferReceta = MyBufferReceta & " RecDes = " & Chr(34) & RecetaDes & Chr(34)
              
                 MyBufferReceta = MyBufferReceta & " Fec = " & Chr(34) & Fecha & Chr(34)
                 MyBufferReceta = MyBufferReceta & " Nli = " & Chr(34) & NumLin & Chr(34)
              
                 MyBufferReceta = MyBufferReceta & " Rac = " & Chr(34) & 0 & Chr(34)
                 
                 MyBufferReceta = MyBufferReceta & " Pon = " & Chr(34) & 0 & Chr(34)
            
                 MyBufferReceta = MyBufferReceta & "/>"
              
                 EstActReceta = True
                 
              End If
              
              MyBuffer = MyBuffer & " Fec = " & Chr(34) & Fecha & Chr(34)
              MyBuffer = MyBuffer & " Nli = " & Chr(34) & NumLin & Chr(34)
               
            'Ración
            If IsNumeric(RsExcel.Fields(j - 4).Value) Then
               
               Raciones = Round(Int(Trim(RsExcel.Fields(j - 4).Value)), 0)
               MyBuffer = MyBuffer & " Rac = " & Chr(34) & Raciones & Chr(34)
            
            End If
   
            'Ponderación
            If IsNumeric(RsExcel.Fields(j - 3).Value) Or Trim(RsExcel.Fields(j - 3).Value) <> "" Then
               
'               If RsExcel.Fields(j - 1).Value < 1 Then
               If IsNumeric(RsExcel.Fields(j - 3).Value) Then
                  
                  PorcentajePon = RsExcel.Fields(j - 3).Value * 100
               
               Else
                  
                  ResultadoPor = ""
                  ResultadoPor = IIf(Not IsNumeric(RsExcel.Fields(j - 3).Value), Replace(RsExcel.Fields(j - 3).Value, "%", ""), 0)
                  
                  If Val(ResultadoPor) > 0 Then
                     
                     PorcentajePon = Replace(RsExcel.Fields(j - 3).Value, "%", "")
                  
                  Else
                     
                     PorcentajePon = 0
                  
                  End If
'                  PorcentajePon = IIf(Not IsNumeric(RsExcel.Fields(j - 1).Value), Replace(RsExcel.Fields(j - 1).Value, "%", ""), 0)
               
               End If
               
               MyBuffer = MyBuffer & " Pon = " & Chr(34) & PorcentajePon & Chr(34)
            
            End If
    
                MyBuffer = MyBuffer & "/>"
                
            
            End If
   

        Next j
      End If
      End If

   DoEvents
           
   RsExcel.MoveNext
   
   lngRow = lngRow + 1
   prbStatus.Value = lngRow
   
   i = i + 1
        
Loop
        
RsExcel.Close
Set RsExcel = Nothing
    
'cn.Close
'Set cn = Nothing
'Set obj_Excel = Nothing
''obj_Workbook.Close
'Set obj_Workbook = Nothing
Set obj_Worksheet = Nothing
    
MyBuffer = MyBuffer & "</UpdateRacPon>"
MyBufferReceta = MyBufferReceta & "</UpdateRacPon>"
MyBufferTotal = MyBufferTotal & "</UpdateComensales>"

'If EstGrabado Then
   
   '--> Actualiza encabezado minuta bloque
   If ValidarRacPon.Value = 1 Then
   
      
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS = vg_db.Execute("sgpadm_Upd_XmlExcelEncMinutaBloque_V01 '" & MyBufferTotal & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

      If Not RS.EOF Then

         vaSpread1.Row = Row
         vaSpread1.Col = 6
   
         If RS(0) > 0 Then

            vaSpread1.text = "Comensales diarios " & RS(0) & " " & RS(1)
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(2).FillColor
      
            lblStatus.Visible = False
            prbStatus.Visible = False
            fg_descarga

            RS.Close
            Set RS = Nothing
      
            Exit Sub
      
         Else
   
            vaSpread1.text = "Comensales diarios finalizo sin problema"
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(1).FillColor
   
         End If
      End If
      RS.Close
      Set RS = Nothing

    End If
       
   'Actualiza detalle minuta bloque
   If ValidarRacPon.Value = 1 Then
      
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS = vg_db.Execute("sgpadm_Upd_XmlExcelDetMinutaBloque_V01 '" & MyBuffer & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")
   
      If Not RS.EOF Then

         vaSpread1.Row = Row
         vaSpread1.Col = 5
   
         If RS(0) > 0 Then

            vaSpread1.text = "Raciones diarias " & RS(0) & " " & RS(1)
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(2).FillColor
      
            lblStatus.Visible = False
            prbStatus.Visible = False
            fg_descarga

            RS.Close
            Set RS = Nothing
      
         Else
   
            vaSpread1.text = "Raciones diarias finalizo sin problema"
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(1).FillColor
      
         End If
   
      End If
      RS.Close
      Set RS = Nothing
   
   End If
   
   '-------> Cambio recetas
   If ValidaReceta = 1 And EstActReceta Then
   
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS = vg_db.Execute("sgpadm_Upd_XmlExcelDetMinutaBloque_V02 '" & MyBufferReceta & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")
   
      If Not RS.EOF Then

         vaSpread1.Row = Row
         vaSpread1.Col = 4
   
         If RS(0) > 0 Then

            vaSpread1.text = "Recetas " & RS(0) & " " & RS(1)
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(2).FillColor
      
            lblStatus.Visible = False
            prbStatus.Visible = False
            fg_descarga

            RS.Close
            Set RS = Nothing
      
        Else
   
            If EstGrabado And EstActReceta Then
               
               vaSpread1.text = "Recetas diarias finalizo sin problema"
               vaSpread1.Col = -1
               vaSpread1.BackColor = Shape1(1).FillColor
            
            ElseIf EstActReceta And Not EstGrabado Then

               vaSpread1.Row = Row
               vaSpread1.Col = 4

               vaSpread1.text = "Realizo algunos cambios de receta. pero receta no concuerda "
               vaSpread1.Col = -1
               vaSpread1.BackColor = Shape1(2).FillColor
      
               lblStatus.Visible = False
               prbStatus.Visible = False
               fg_descarga
            
            ElseIf EstGrabado And Not EstActReceta Then
               
               vaSpread1.text = "no hubo cambio de recetas"
               vaSpread1.Col = -1
               vaSpread1.BackColor = Shape1(1).FillColor
            
            End If
      
        End If
   
      End If
      RS.Close
      Set RS = Nothing
   
   ElseIf ValidaReceta = 1 And Not EstActReceta Then
   
      vaSpread1.Row = Row
      vaSpread1.Col = 4
   
      vaSpread1.text = "La plantilla, viene sin cambio de recetas"
      vaSpread1.Col = -1
      vaSpread1.BackColor = Shape1(1).FillColor

   End If
      
'Else

'   vaSpread1.Row = Row
'   vaSpread1.Col = 4

'   vaSpread1.text = "Nombre receta no concuerda "
'   vaSpread1.Col = -1
'   vaSpread1.BackColor = Shape1(2).FillColor
      
'   lblStatus.Visible = False
'   prbStatus.Visible = False
'   fg_descarga

'End If

'Validar recetas
If vaSpread3.MaxRows > 0 And ValidaReceta = 1 Then

   Let MyBufferRecetaVal = ""
   Let MyBufferRecetaVal = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
   Let MyBufferRecetaVal = MyBufferRecetaVal & "<Receta>"

   For IndJ = 1 To vaSpread3.MaxRows
   
       MyBufferRecetaVal = MyBufferRecetaVal & " <DetReceta"
       
       vaSpread3.Row = IndJ
       vaSpread3.Col = 1
       
       MyBufferRecetaVal = MyBufferRecetaVal & " Rec = " & Chr(34) & vaSpread3.text & Chr(34)
                       
       MyBufferRecetaVal = MyBufferRecetaVal & "/>"
       
   
   Next IndJ
   
   MyBufferRecetaVal = MyBufferRecetaVal & "</Receta>"
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
      
   Set RS = vg_db.Execute("sgpadm_Sel_XmlExcelValidarRecetaSitio_V01 '" & MyBufferRecetaVal & "', '" & Ceco & "', " & Regimen & ", " & Servicio & "")
   
   If Not RS.EOF Then
        
      fg_descarga
      
      If RS(0) > 0 Then
      
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
        
'          xlApp.Columns("A:B").Select
'          xlApp.Selection.Delete Shift:=xlToLeft
      
          NomArchivoExcel = fg_ArchivoXls("ReporteError_ActualizacionMinutaBloque")
                        
          xlWb.Close True, NomArchivoExcel
    
          XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
          XL.Visible = True
          XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
        
          '-- Cerrar Excel
          xlApp.Quit
          
          '-------> Release Excel references
          Set xlWs = Nothing
          Set xlWb = Nothing
          Set xlApp = Nothing
          
          RS.Close
          Set RS = Nothing
          
          Exit Sub

      End If
      
   Else
    
      RS.Close
      Set RS = Nothing
   
   End If
    
End If

lblStatus.Visible = False
prbStatus.Visible = False

Exit Sub
Man_Error:
fg_descarga

If Err = 6 Or Err = 3021 Then

    Raciones = 0
    Resume Next

End If

If Err = 3265 Then
      
   If RS.State = 1 Then RS.Close
   
   If MaxColumnas <= 187 Then
   
      Resume Next
      
   Else
   
      MsgBox "Sobrepasa cantidad de columna. maximo deberia ser 254. Proceso cancelado ", vbCritical, MsgTitulo
      Exit Sub
   
   End If
   
End If
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub ActualizarComensalesDiariosTotalesFormatonuevoII(SheetName As String, MaxColumnas As Long, Row As Long)

On Error GoTo Man_Error

Dim i                   As Long
Dim j                   As Long
Dim IndJ                As Long
Dim dbexcel             As Database
Dim cn                  As ADODB.Connection
Dim RS                  As New ADODB.Recordset
Dim IndColumna          As Long
Dim MyBuffer            As String
Dim MyBufferReceta      As String
Dim MyBufferTotal       As String
Dim MyBufferRecetaVal   As String

Dim strArray()          As String
Dim strTest             As String
Dim intCount            As Integer
Dim UltRow              As Long

Dim Ceco                As String
Dim Regimen             As Long
Dim Servicio            As Long
Dim RecetaOri           As Long
Dim RecetaDes           As Long
Dim Fecha               As Long
Dim NumLin              As Long
Dim Raciones            As Long
Dim PorcentajePon       As Long
Dim RacionesRe          As Long
Dim PorcentajePonRe     As Long
Dim ComesalesTot        As Long
Dim NombreReceta        As String
Dim NomRow              As Long
Dim EstGrabado          As Boolean
Dim EstActReceta        As Boolean
Dim lngRow              As Long
Dim CodigoReceta        As Long
Dim CodigoRecetaValidar As Long
Dim ResultadoPor        As String
Dim estser              As Long
Dim ComensalesTot       As Double

Dim NomArchivoExcel     As String

'Definición variables excel
Dim cs                  As String
Dim PathXls             As String
Dim xlApp               As Object
Dim xlWb                As Object
Dim xlWs                As Object
Dim XL                  As New excel.Application 'Crea el objeto excel
Dim RsExcel             As New ADODB.Recordset


vaSpread3.MaxRows = 0
vaSpread3.maxcols = 1

EstActReceta = False

EstGrabado = True
Set RsExcel = New ADODB.Recordset
Set cn = New ADODB.Connection

PathXls = Trim(fpText1.text)
cs = "DRIVER=Microsoft Excel Driver (*.xls);DBQ=" & PathXls & ";HDR=NO;IMEX=1;"
'Set obj_Excel = CreateObject("Excel.Application")
 
' -- Abrir el libro
'Set obj_Workbook = obj_Excel.Workbooks.Open(PathXls)

RsExcel.CursorLocation = adUseClient
RsExcel.CursorType = adOpenKeyset
RsExcel.LockType = adLockBatchOptimistic

SheetName = "[" & SheetName & "]"
RsExcel.Open "SELECT * FROM  " & SheetName, cs

If RsExcel.EOF Then Exit Sub

RsExcel.MoveFirst

IndColumna = 6
i = 1

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<UpdateRacPon>"

Let MyBufferReceta = ""
Let MyBufferReceta = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBufferReceta = MyBufferReceta & "<UpdateRacPon>"

Let MyBufferTotal = ""
Let MyBufferTotal = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBufferTotal = MyBufferTotal & "<UpdateComensales>"

prbStatus.Max = 1
lblStatus.Visible = True
prbStatus.Visible = True
prbStatus.Min = 0
lngRow = 0
prbStatus.Max = RsExcel.RecordCount

lblStatus.Caption = Mid(SheetName, 1, Len(SheetName) - 1)

Do While RsExcel.EOF <> True
           
   If RsExcel.Fields(0).Value = "*" Then Exit Do
           
   'Comensales Totales
   If RsExcel.Fields(0).Value = "Comensales" Then
         
        For j = IndColumna To MaxColumnas Step 6
       
            If RsExcel.Fields(j).Value <> "" Then
                
               MyBufferTotal = MyBufferTotal & " <Comensales"
        
               strTest = RsExcel.Fields(j).Value
               strArray = Split(strTest, ";")
   
               For intCount = LBound(strArray) To UBound(strArray)
                    
                    Select Case intCount
                        
                        Case 0
                            
                            Ceco = Trim(strArray(intCount))
                        
                        Case 1
                            
                            Regimen = Trim(strArray(intCount))
                        
                        Case 2
                            
                            Servicio = Trim(strArray(intCount))
                        
                        Case 3
                            
                            Fecha = Trim(strArray(intCount))
                    

                    End Select
               
               Next
                
              MyBufferTotal = MyBufferTotal & " Ceco = " & Chr(34) & Ceco & Chr(34)
              MyBufferTotal = MyBufferTotal & " Reg = " & Chr(34) & Regimen & Chr(34)
              MyBufferTotal = MyBufferTotal & " Ser = " & Chr(34) & Servicio & Chr(34)
              MyBufferTotal = MyBufferTotal & " Fec = " & Chr(34) & Fecha & Chr(34)
              
            End If
   
            'Ración
            If IsNumeric(RsExcel.Fields(j - 4).Value) Then
               
               ComensalesTot = RsExcel.Fields(j - 4).Value
               MyBufferTotal = MyBufferTotal & " Com = " & Chr(34) & ComensalesTot & Chr(34)
            
            Else
            
               MyBufferTotal = MyBufferTotal & " Com = " & Chr(34) & -1 & Chr(34)
            
            End If
   
            MyBufferTotal = MyBufferTotal & "/>"
        
        Next j
   
   'Detalle de minuta
   ElseIf RsExcel.Fields(0).Value <> "" Or i > 7 Then
   
        If i > 7 Then
        For j = IndColumna To MaxColumnas Step 6
       
            If RsExcel.Fields(j).Value <> "" And Not IsNull(Trim(RsExcel.Fields(j - 5).Value)) Then
               
               '-------> sacar codigo estrutyura servicio
               If RsExcel.Fields(0).Value <> "" Then
                  
                  strTest = RsExcel.Fields(0).Value
                  strArray = Split(strTest, ";")
                  estser = strArray(1)
                  
               End If
               
               'RsExcel.Fields(j).ColorIndex
               strTest = RsExcel.Fields(j).Value
               strArray = Split(strTest, ";")
   
               MyBuffer = MyBuffer & " <RacPon"
               
               For intCount = LBound(strArray) To UBound(strArray)
                    
                    Select Case intCount
                        
                        Case 0 'ceco
                            
                            Ceco = Trim(strArray(intCount))
                            
                            If ValidaReceta.Value = 1 And Not EstCargarRecetas Then
        
                               CargarRecetaGrilla Ceco
                               EstCargarRecetas = True
    
                            End If
                            
                        Case 1 'regimen
                            
                            Regimen = Trim(strArray(intCount))
                        
                        Case 2 'servicio
                            
                            Servicio = Trim(strArray(intCount))
                        
                        Case 3
                            
                            RecetaDes = 0
                            RecetaDes = 0
                            
                            If ValidaReceta.Value = 1 Then
                                
                                
                                EstGrabado = True

                                NombreReceta = IIf(IsNull(Trim(RsExcel.Fields(j - 5).Value)), "", Trim(RsExcel.Fields(j - 5).Value))
                                
                                'Ini : Mover codigo receta
                                
                                CodigoRecetaValidar = 0
                                CodigoRecetaValidar = SacarNumeroRight(NombreReceta)
                                
                                MoverCodigoReceta CodigoRecetaValidar
                                
                                'Fin : Mover codigo receta
                                
                                NomRow = 0
                                NomRow = vaSpread2.SearchCol(2, 0, vaSpread2.MaxRows, NombreReceta, SearchFlagsEqual) 'SearchFlagsEqual)
                  
                                CodigoReceta = 0
                                
                                If NomRow < 1 Then
                                   
                                   CodigoReceta = SacarNumeroRight(NombreReceta)

                                   NomRow = 0
                                   NomRow = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(CodigoReceta), SearchFlagsEqual) 'SearchFlagsEqual)
                                   RecetaDes = 0
                                   
                                   If NomRow < 1 Then
                                      
                                      EstGrabado = False
                                      
                                   Else
                                   
                                       vaSpread2.Row = NomRow
                                       vaSpread2.Col = 1
                                       RecetaDes = vaSpread2.text
                                   
                                   End If
                                
                                Else
                                   
                                   vaSpread2.Row = NomRow
                                   vaSpread2.Col = 1
'                                   RecetaDes = SacarNumeroRight(NombreReceta)
                                   RecetaDes = vaSpread2.text
                                   
                                End If
                            
                            End If
                            
                            RecetaOri = Trim(strArray(intCount))
                            
                            If RecetaDes = 0 Then
                            
                               RecetaDes = RecetaOri
                            
                            End If
                            
                        Case 4
                            
                            Fecha = Trim(strArray(intCount))
                        
                        Case 5
                            
                            NumLin = Trim(strArray(intCount))
                    
                    End Select
               
               Next intCount
                
              MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
              MyBuffer = MyBuffer & " Reg = " & Chr(34) & Regimen & Chr(34)
              MyBuffer = MyBuffer & " Ser = " & Chr(34) & Servicio & Chr(34)
              MyBuffer = MyBuffer & " ROri = " & Chr(34) & RecetaOri & Chr(34)
              MyBuffer = MyBuffer & " RDes = " & Chr(34) & RecetaDes & Chr(34)
              
              If ValidaReceta.Value = 1 And RecetaDes <> RecetaOri Then
              
                 MyBufferReceta = MyBufferReceta & " <RacPon"
                 MyBufferReceta = MyBufferReceta & " Ceco = " & Chr(34) & Ceco & Chr(34)
                 MyBufferReceta = MyBufferReceta & " Reg = " & Chr(34) & Regimen & Chr(34)
                 MyBufferReceta = MyBufferReceta & " Ser = " & Chr(34) & Servicio & Chr(34)
                 MyBufferReceta = MyBufferReceta & " RecOri = " & Chr(34) & RecetaOri & Chr(34)
                 
                 MyBufferReceta = MyBufferReceta & " RecDes = " & Chr(34) & RecetaDes & Chr(34)
              
                 MyBufferReceta = MyBufferReceta & " Fec = " & Chr(34) & Fecha & Chr(34)
                 MyBufferReceta = MyBufferReceta & " Nli = " & Chr(34) & NumLin & Chr(34)
                 
                 If RecetaOri = 0 Then
                 
                     RacionesRe = 0
                     
                     'Ración
                     If IsNumeric(RsExcel.Fields(j - 4).Value) Then
                        
                        RacionesRe = Round(Int(Trim(RsExcel.Fields(j - 4).Value)), 0)
                        MyBufferReceta = MyBufferReceta & " Rac = " & Chr(34) & RacionesRe & Chr(34)
                     
                     End If
            
                     PorcentajePonRe = 0
                     
                     'Ponderación
                     If IsNumeric(RsExcel.Fields(j - 3).Value) Or Trim(RsExcel.Fields(j - 3).Value) <> "" Then
                        
                        If IsNumeric(RsExcel.Fields(j - 3).Value) Then
                           
                           PorcentajePonRe = RsExcel.Fields(j - 3).Value * 100
                        
                        Else
                           
                           ResultadoPor = ""
                           ResultadoPor = IIf(Not IsNumeric(RsExcel.Fields(j - 3).Value), Replace(RsExcel.Fields(j - 3).Value, "%", ""), 0)
                           
                           If Val(ResultadoPor) > 0 Then
                              
                              PorcentajePonRe = Replace(RsExcel.Fields(j - 3).Value, "%", "")
                           
                           Else
                              
                              PorcentajePonRe = 0
                           
                           End If
                        
                        End If
                        
                        MyBufferReceta = MyBufferReceta & " Pon = " & Chr(34) & PorcentajePonRe & Chr(34)
                     
                     End If
                    
                 Else
                 
                    MyBufferReceta = MyBufferReceta & " Rac = " & Chr(34) & 0 & Chr(34)
                    MyBufferReceta = MyBufferReceta & " Pon = " & Chr(34) & 0 & Chr(34)
                 
                 End If
                 
                 MyBufferReceta = MyBufferReceta & " ESer = " & Chr(34) & estser & Chr(34)
            
                 MyBufferReceta = MyBufferReceta & "/>"
              
                 EstActReceta = True
                 
              End If
              
              MyBuffer = MyBuffer & " Fec = " & Chr(34) & Fecha & Chr(34)
              MyBuffer = MyBuffer & " Nli = " & Chr(34) & NumLin & Chr(34)
               
            'Ración
            If IsNumeric(RsExcel.Fields(j - 4).Value) Then
               
               Raciones = Round(Int(Trim(RsExcel.Fields(j - 4).Value)), 0)
               MyBuffer = MyBuffer & " Rac = " & Chr(34) & Raciones & Chr(34)
            
            End If
   
            'Ponderación
            If IsNumeric(RsExcel.Fields(j - 3).Value) Or Trim(RsExcel.Fields(j - 3).Value) <> "" Then
               
               If IsNumeric(RsExcel.Fields(j - 3).Value) Then
                  
                  PorcentajePon = RsExcel.Fields(j - 3).Value * 100
               
               Else
                  
                  ResultadoPor = ""
                  ResultadoPor = IIf(Not IsNumeric(RsExcel.Fields(j - 3).Value), Replace(RsExcel.Fields(j - 3).Value, "%", ""), 0)
                  
                  If Val(ResultadoPor) > 0 Then
                     
                     PorcentajePon = Replace(RsExcel.Fields(j - 3).Value, "%", "")
                  
                  Else
                     
                     PorcentajePon = 0
                  
                  End If
               
               End If
               
               MyBuffer = MyBuffer & " Pon = " & Chr(34) & PorcentajePon & Chr(34)
            
            End If
    
            MyBuffer = MyBuffer & " ESer = " & Chr(34) & estser & Chr(34)
            
            MyBuffer = MyBuffer & "/>"
                
            
            End If
   

        Next j
      End If
      
      End If

   DoEvents
           
   RsExcel.MoveNext
   
   lngRow = lngRow + 1
   prbStatus.Value = lngRow
   
   i = i + 1
        
Loop
        
RsExcel.Close
Set RsExcel = Nothing
    
'Set obj_Excel = Nothing
'obj_Workbook.Close
'Set obj_Workbook = Nothing
Set obj_Worksheet = Nothing
'Set obj_Excel = Nothing
'Set obj_Workbook = Nothing

MyBuffer = MyBuffer & "</UpdateRacPon>"
MyBufferReceta = MyBufferReceta & "</UpdateRacPon>"
MyBufferTotal = MyBufferTotal & "</UpdateComensales>"

'If EstGrabado Then
   
   '--> Actualiza encabezado minuta bloque
   If ValidarRacPon.Value = 1 Then
   
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS = vg_db.Execute("sgpadm_Upd_XmlExcelEncMinutaBloque_V01 '" & MyBufferTotal & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

      If Not RS.EOF Then

         vaSpread1.Row = Row
         vaSpread1.Col = 6
   
         If RS(0) > 0 Then

            vaSpread1.text = "Comensales diarios " & RS(0) & " " & RS(1)
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(2).FillColor
      
            lblStatus.Visible = False
            prbStatus.Visible = False
            fg_descarga

            RS.Close
            Set RS = Nothing
      
            Exit Sub
      
         Else
   
            vaSpread1.text = "Comensales diarios finalizo sin problema"
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(1).FillColor
   
         End If
      End If
      RS.Close
      Set RS = Nothing

    End If
       
   'Actualiza detalle minuta bloque
   If ValidarRacPon.Value = 1 Then
      
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS = vg_db.Execute("sgpadm_Upd_XmlExcelDetMinutaBloqueRacPon_V01 '" & MyBuffer & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")
   
      If Not RS.EOF Then

         vaSpread1.Row = Row
         vaSpread1.Col = 5
   
         If RS(0) > 0 Then

            vaSpread1.text = "Raciones diarias " & RS(0) & " " & RS(1)
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(2).FillColor
      
            lblStatus.Visible = False
            prbStatus.Visible = False
            fg_descarga

'            RS.Close
'            Set RS = Nothing
            
         ElseIf RS(0) < 0 Then
      
            lblStatus.Visible = False
            prbStatus.Visible = False
            fg_descarga
        
            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), Me.HelpContextID, "", "", "")

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
  
            NomArchivoExcel = fg_ArchivoXls("ReporteError_ActualizacionMinutaBloque")
                    
            xlWb.Close True, NomArchivoExcel

'            Dim XL As New excel.Application 'Crea el objeto excel
            XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
            XL.Visible = True
            XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
            '-- Cerrar Excel
            xlApp.Quit
      
            '-------> Release Excel references
            Set xlWs = Nothing
            Set xlWb = Nothing
            Set xlApp = Nothing
      
            RS.Close
            Set RS = Nothing
      
            Exit Sub
      
         Else
   
            vaSpread1.text = "Raciones diarias finalizo sin problema"
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(1).FillColor
      
         End If
   
      End If
      RS.Close
      Set RS = Nothing
   
   End If
   
   '-------> Cambio recetas
   If ValidaReceta = 1 And EstActReceta Then
   
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS = vg_db.Execute("sgpadm_Upd_XmlExcelDetMinutaBloqueRec_V01 '" & MyBufferReceta & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")
   
      If Not RS.EOF Then

         vaSpread1.Row = Row
         vaSpread1.Col = 4
   
         If RS(0) > 0 Then

            vaSpread1.text = "Recetas " & RS(0) & " " & RS(1)
            vaSpread1.Col = -1
            vaSpread1.BackColor = Shape1(2).FillColor
      
            lblStatus.Visible = False
            prbStatus.Visible = False
            fg_descarga

'            RS.Close
'            Set RS = Nothing
              
        ElseIf RS(0) < 0 Then
        
            lblStatus.Visible = False
            prbStatus.Visible = False
            fg_descarga
        
            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), Me.HelpContextID, "", "", "")

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
  
            NomArchivoExcel = fg_ArchivoXls("ReporteError_ActualizacionMinutaBloque")
                    
            xlWb.Close True, NomArchivoExcel

            XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
            XL.Visible = True
            XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
            '-- Cerrar Excel
            xlApp.Quit
      
            '-------> Release Excel references
            Set xlWs = Nothing
            Set xlWb = Nothing
            Set xlApp = Nothing
      
            RS.Close
            Set RS = Nothing
      
            Exit Sub

        Else
   
            If EstGrabado And EstActReceta Then
               
               vaSpread1.text = "Recetas diarias finalizo sin problema"
               vaSpread1.Col = -1
               vaSpread1.BackColor = Shape1(1).FillColor
            
            ElseIf EstActReceta And Not EstGrabado Then

               vaSpread1.Row = Row
               vaSpread1.Col = 4

               vaSpread1.text = "Realizo algunos cambios de receta. pero receta no concuerda "
               vaSpread1.Col = -1
               vaSpread1.BackColor = Shape1(2).FillColor
      
               lblStatus.Visible = False
               prbStatus.Visible = False
               fg_descarga
            
            ElseIf EstGrabado And Not EstActReceta Then
               
               vaSpread1.text = "no hubo cambio de recetas"
               vaSpread1.Col = -1
               vaSpread1.BackColor = Shape1(1).FillColor
            
            End If
      
        End If
   
      End If
      RS.Close
      Set RS = Nothing
   
   ElseIf ValidaReceta = 1 And Not EstActReceta Then
   
      vaSpread1.Row = Row
      vaSpread1.Col = 4
   
      vaSpread1.text = "La plantilla, viene sin cambio de recetas"
      vaSpread1.Col = -1
      vaSpread1.BackColor = Shape1(1).FillColor
   
   End If

'Validar recetas
If vaSpread3.MaxRows > 0 And ValidaReceta = 1 Then

   Let MyBufferRecetaVal = ""
   Let MyBufferRecetaVal = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
   Let MyBufferRecetaVal = MyBufferRecetaVal & "<Receta>"

   For IndJ = 1 To vaSpread3.MaxRows
   
       MyBufferRecetaVal = MyBufferRecetaVal & " <DetReceta"
       
       vaSpread3.Row = IndJ
       vaSpread3.Col = 1
       
       MyBufferRecetaVal = MyBufferRecetaVal & " Rec = " & Chr(34) & vaSpread3.text & Chr(34)
                       
       MyBufferRecetaVal = MyBufferRecetaVal & "/>"
       
   
   Next IndJ
   
   MyBufferRecetaVal = MyBufferRecetaVal & "</Receta>"
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
      
   Set RS = vg_db.Execute("sgpadm_Sel_XmlExcelValidarRecetaSitio_V01 '" & MyBufferRecetaVal & "', '" & Ceco & "', " & Regimen & ", " & Servicio & "")
   
   If Not RS.EOF Then
        
      fg_descarga
      
      If RS(0) > 0 Then
      
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
        
'          xlApp.Columns("A:B").Select
'          xlApp.Selection.Delete Shift:=xlToLeft
      
          NomArchivoExcel = fg_ArchivoXls("ReporteError_ActualizacionMinutaBloque")
                        
          xlWb.Close True, NomArchivoExcel
    
          XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
          XL.Visible = True
          XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
        
          '-- Cerrar Excel
          xlApp.Quit
          
          '-------> Release Excel references
          Set xlWs = Nothing
          Set xlWb = Nothing
          Set xlApp = Nothing
          
          RS.Close
          Set RS = Nothing
          
          Exit Sub

      End If
      
   Else
    
      RS.Close
      Set RS = Nothing
   
   End If
    
End If

lblStatus.Visible = False
prbStatus.Visible = False

Exit Sub
Man_Error:
fg_descarga

If Err = 6 Or Err = 3021 Then

    Raciones = 0
    Resume Next

End If

If Err = 3265 Then
      
   If RS.State = 1 Then RS.Close
   
   If MaxColumnas <= 187 Then
   
      Resume Next
      
   Else
   
      MsgBox "Sobrepasa cantidad de columna. maximo deberia ser 254. Proceso cancelado ", vbCritical, MsgTitulo
      Exit Sub
   
   End If
   
End If
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub MoverCodigoReceta(CodigoReceta As Long)

Dim CodRow As Long

CodRow = 0
CodRow = vaSpread3.SearchCol(1, 0, vaSpread3.MaxRows, CodigoReceta, SearchFlagsEqual) 'SearchFlagsEqual)
                  
If CodRow < 1 Then

   vaSpread3.MaxRows = vaSpread3.MaxRows + 1
   vaSpread3.Row = vaSpread3.MaxRows
   vaSpread3.Col = 1
   vaSpread3.text = CodigoReceta

End If

End Sub


