VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form P_CambioRecetaMinBloque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio Recetas Minutas Bloque"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
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
      Left            =   5520
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   7680
      TabIndex        =   3
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9495
      Begin VB.OptionButton Option1 
         Caption         =   "Actualiza Recetas % Raciones"
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
         Left            =   1800
         TabIndex        =   13
         Top             =   1680
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Actualiza Q Total Día"
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
         Left            =   1800
         TabIndex        =   12
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Raciones"
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
         Index           =   2
         Left            =   6240
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "% Ponderación"
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
         Left            =   3960
         TabIndex        =   10
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Recetas"
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
         Left            =   2040
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   7215
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1755
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   7335
         _Version        =   196608
         _ExtentX        =   12938
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
      Begin MSComctlLib.ProgressBar prbStatus 
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Top             =   3360
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hoja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "lblStatus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   3120
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Archivo Origen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   2
         Top             =   555
         Width           =   1275
      End
   End
End
Attribute VB_Name = "P_CambioRecetaMinBloque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private obj_Excel       As Object
Private obj_Workbook    As Object
Private obj_Worksheet   As Object

Public lc_Aux As String
Dim MsgTitulo As String


Private Sub Combo1_Click()

On Error GoTo Man_Error

Dim PathXls    As String
Dim hoja       As String
Dim sSheetName As String
Dim i          As Long

PathXls = Trim(fpText1.text)
hoja = Combo1.text '& "$"
   
' -- crea rnueva instancia de Excel
Set obj_Excel = CreateObject("Excel.Application")

' -- Abrir el libro
Set obj_Workbook = obj_Excel.Workbooks.Open(PathXls)
' -- referencia la Hoja, por defecto la hoja activa

If sSheetName = vbNullString Then
   
   Set obj_Worksheet = obj_Workbook.ActiveSheet
      
   For i = 1 To obj_Workbook.Sheets.count
         
       If hoja = obj_Workbook.Sheets(i).Name Then
        
          hoja = obj_Workbook.Sheets(i).Name
          Exit For
            
       End If
    
   Next
      
Else
   
   Set obj_Worksheet = obj_Workbook.Sheets(hoja)
'
   hoja = obj_Workbook.Sheets(hoja)
   
End If

If Trim(obj_Workbook.Worksheets(hoja).Range("a1").Value) = "min_cecori" And Trim(obj_Workbook.Worksheets(hoja).Range("b1").Value) = "reg_codigo" And _
         Trim(obj_Workbook.Worksheets(hoja).Range("c1").Value) = "reg_nombre" And Trim(obj_Workbook.Worksheets(hoja).Range("d1").Value) = "ser_codigo" And _
         Trim(obj_Workbook.Worksheets(hoja).Range("e1").Value) = "ser_nombre" And Trim(obj_Workbook.Worksheets(hoja).Range("f1").Value) = "ess_codigo" And _
         Trim(obj_Workbook.Worksheets(hoja).Range("g1").Value) = "ess_nombre" And Trim(obj_Workbook.Worksheets(hoja).Range("h1").Value) = "min_fecmin" And _
         Trim(obj_Workbook.Worksheets(hoja).Range("i1").Value) = "rec_codigo" And Trim(obj_Workbook.Worksheets(hoja).Range("j1").Value) = "rec_nombre" And _
         Trim(obj_Workbook.Worksheets(hoja).Range("k1").Value) = "rec_catdie" And Trim(obj_Workbook.Worksheets(hoja).Range("l1").Value) = "Descripción Dietetica" And _
         Trim(obj_Workbook.Worksheets(hoja).Range("m1").Value) = "rec_tippla" And Trim(obj_Workbook.Worksheets(hoja).Range("n1").Value) = "Descripción Tipo Plato" And _
         (Trim(obj_Workbook.Worksheets(hoja).Range("o1").Value) = "% Ponderacion" Or Trim(obj_Workbook.Worksheets(hoja).Range("o1").Value) = "% Ponderacion_1") And _
         (Trim(obj_Workbook.Worksheets(hoja).Range("p1").Value) = "Raciones" Or Trim(obj_Workbook.Worksheets(hoja).Range("p1").Value) = "Raciones_1") And _
         Trim(obj_Workbook.Worksheets(hoja).Range("q1").Value) = "Comensales" And Trim(obj_Workbook.Worksheets(hoja).Range("r1").Value) = "mid_numlin" And _
         Trim(obj_Workbook.Worksheets(hoja).Range("s1").Value) = "Cód. New Receta" _
   Then
      
      Option1(0).Value = True
      Option1(0).Enabled = True
        
      Option1(1).Enabled = False
      Option1(1).Value = False
      
      Check1(0).Value = 0
      Check1(1).Value = 0
      Check1(2).Value = 0

      Check1(0).Enabled = True
      Check1(1).Enabled = True
      Check1(2).Enabled = True

      Command1(0).Enabled = True
      
ElseIf Trim(obj_Workbook.Worksheets(hoja).Range("a1").Value) = "min_cecori" And Trim(obj_Workbook.Worksheets(hoja).Range("b1").Value) = "reg_codigo" And _
       Trim(obj_Workbook.Worksheets(hoja).Range("c1").Value) = "reg_nombre" And Trim(obj_Workbook.Worksheets(hoja).Range("d1").Value) = "ser_codigo" And _
       Trim(obj_Workbook.Worksheets(hoja).Range("e1").Value) = "ser_nombre" And Trim(obj_Workbook.Worksheets(hoja).Range("f1").Value) = "min_fecmin" And _
       Trim(obj_Workbook.Worksheets(hoja).Range("g1").Value) = "Comensales" _
  Then
      
     Option1(0).Value = False
     Option1(0).Enabled = False

     Option1(1).Enabled = True
     Option1(1).Value = True
      
     Check1(0).Value = 0
     Check1(1).Value = 0
     Check1(2).Value = 0

     Check1(0).Enabled = False
     Check1(1).Enabled = False
     Check1(2).Enabled = False

     Command1(0).Enabled = True
     
Else
      
    Command1(0).Enabled = False
    
    ' -- Cerrar libro
    obj_Workbook.Close

    ' -- Cerrar Excel
    obj_Excel.Quit
    Set obj_Workbook = Nothing
    Set obj_Excel = Nothing
    Set obj_Worksheet = Nothing
    
    MsgBox "Formato no corresponde al estandar. Proceso cancelado ", vbCritical, MsgTitulo
    Exit Sub
      
End If

'obj_Workbook.Worksheets(hoja).Range("P1").Locked = False And
If obj_Workbook.Worksheets(hoja).Range("P1").Value = "Raciones_1" Then

   Check1(1).Enabled = False
   Check1(1).Value = 0
   Check1(2).Value = 1

' obj_Workbook.Worksheets(hoja).Range("O1").Locked = False And

ElseIf obj_Workbook.Worksheets(hoja).Range("O1").Value = "% Ponderacion_1" Then

   Check1(1).Value = 1
   
   Check1(2).Value = 0
   Check1(2).Enabled = False

End If
 
' -- Cerrar libro
obj_Workbook.Close

' -- Cerrar Excel
obj_Excel.Quit
Set obj_Workbook = Nothing
Set obj_Excel = Nothing
Set obj_Worksheet = Nothing

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String

Select Case Index
    
    Case 0
                   
        If ValidaDatos = False Then Exit Sub
               
        If MsgBox("Esta seguro realizar cambio...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
       
        If Option1(0).Value = True Then
           
           ActualizarCambioReceta
        
        ElseIf Option1(1).Value = True Then
        
           ActualizarCambioComensales
        
        End If
    
    Case 1
    
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "", "")
        Unload Me
        
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me
fg_carga ""
MsgTitulo = "Cambio Recetas Minuta Bloque"

lblStatus.Visible = False
prbStatus.Visible = False
Combo1.Clear

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
Dim myPath     As String
Dim NomArchivo As String
Dim i          As Long


CD.Filter = "Archivos xls|*.xls|Archivos xlsx|*.xlsx"
CD.DefaultExt = "*.xls|*.xlsx"
CD.FilterIndex = 2
CD.Flags = cdlOFNFileMustExist
CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
CD.ShowOpen

If CD.FileName = "" Then
   
   fpText1.text = ""

Else

    Combo1.Clear
    
    fpText1.text = CD.FileName 'Dir(CD.FileName)

'    Dim ObjExcel As excel.Application
'    Dim ObjW As excel.Workbook
'
'    Set ObjExcel = New excel.Application
'    Set ObjW = ObjExcel.Workbooks.Open(fpText1.text)
'    Dim i As Integer
'    Dim HojaPro As Boolean
'
'    For i = 1 To ObjW.Sheets.count
'
'        Combo1.AddItem ObjW.Sheets(i).Name
'
'    Next
'
'    ObjW.Application.DisplayAlerts = False
'    ObjW.Close
'    Set ObjExcel = Nothing
'    Set ObjW = Nothing
    
    NomArchivo = Dir(CD.FileName)

    Dim ApExcel As excel.Application

    'Al configurarlo
    Set ApExcel = New excel.Application

    'Al abrirlo
    ApExcel.Workbooks.Open FileName:=fpText1.text
    
    For i = 1 To ApExcel.Sheets.count

        Combo1.AddItem ApExcel.Sheets(i).Name

    Next
    
    ApExcel.Workbooks(NomArchivo).Close SaveChanges:=False

    ApExcel.Visible = False
    ApExcel.Application.Visible = False
    ApExcel.Application.Quit
    Set ApExcel = Nothing
    
   
End If

Exit Sub
Man_Error:
fg_descarga
If Err = 5 Then Resume Next: Exit Sub
If Err = 55 Then MsgBox "Archivo esta abierto, proceso cancelado": Exit Sub
If Err = 462 Or Err = 1004 Or Err = 438 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Resume

End Sub

Private Function ValidaDatos() As Boolean

On Error GoTo Man_Error

Dim SheetName As String
Dim cn        As ADODB.Connection
Dim RsExcel   As New ADODB.Recordset
Dim PathXls   As String
Dim File_Ext  As String
Dim dbexcel   As Database

Set cn = New ADODB.Connection


Let ValidaDatos = True
 
'-------> Validar Archivo Origen
If Trim(LimpiaDato(fpText1.text)) = "" Then
    
    Call MsgBox("Debe seleccionar archivo origen", vbInformation, Me.Caption)
    Call fpText1.SetFocus
    Let ValidaDatos = False
    Exit Function

End If

'-------> Validar hoja
If Combo1.ListIndex = -1 Then
    
    Call MsgBox("Debe seleccionar hoja", vbInformation, Me.Caption)
    Call Combo1.SetFocus
    Let ValidaDatos = False
    Exit Function

End If

'-------> Validar si hay seleccionado la primera opción
If Option1(0).Value = True And Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 Then

    Call MsgBox("Debe haber un items seleccionado, en item Actualiza Receta - % - Raciones", vbInformation, Me.Caption)
    Call Check1(0).SetFocus
    Let ValidaDatos = False
    Exit Function


End If

DoEvents
   
Exit Function
Man_Error:
    ValidaDatos = False
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Function

Sub ActualizarCambioReceta()

On Error GoTo Man_Error

Dim i               As Long
Dim j               As Long
Dim PathXls         As String
Dim dbexcel         As Database
Dim cn              As ADODB.Connection
Dim RS              As New ADODB.Recordset
Dim SheetName       As String
Dim hoja            As String
Dim sSheetName      As String
Dim cs              As String
Dim IndColumna      As Long
Dim MyBuffer        As String
Dim MyBufferTotal   As String
Dim File_Ext        As String

Dim strArray()      As String
Dim intCount        As Integer
Dim UltRow          As Long

Dim Ceco            As String
Dim Regimen         As Long
Dim Servicio        As Long
Dim RecetaOrigen    As Long
Dim RecetaDestino   As Long
Dim Fecha           As Long
Dim NumLin          As Long
Dim Ponderacion     As Long
Dim Raciones        As Long

Dim lngRow          As Long
Dim NomArchivoExcel As String
Dim NomArchivo      As String

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object

Dim ApExcel As New excel.Application
Set ApExcel = New excel.Application

Dim RsExcel As New ADODB.Recordset
Set cn = New ADODB.Connection

NomArchivo = Dir(CD.FileName)

PathXls = Trim(fpText1.text)
File_Ext = UCase(Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, "."))))
'SheetName = Combo1.text & "$"
hoja = Combo1.text

RsExcel.CursorLocation = adUseClient
RsExcel.CursorType = adOpenKeyset
RsExcel.LockType = adLockBatchOptimistic
 
cs = "DRIVER=Microsoft Excel Driver (*.xls);DBQ=" & PathXls & ";HDR=NO;IMEX=1;"
'Set obj_Excel = CreateObject("Excel.Application")
 
' -- Abrir el libro y hoja
ApExcel.Workbooks.Open FileName:=PathXls
'Set obj_Workbook = ApExcel.Workbooks.Open(PathXls)  'obj_Excel.Workbooks.Open(PathXls)
' -- referencia la Hoja, por defecto la hoja activa

If sSheetName = vbNullString Then
   
'   Set obj_Worksheet = obj_Workbook.ActiveSheet
      
   For i = 1 To ApExcel.Sheets.count 'obj_Workbook.Sheets.count
         
       If hoja = ApExcel.Sheets(i).Name Then
        
          hoja = ApExcel.Sheets(i).Name
          Exit For
            
       End If
    
   Next
      
Else

'   Set obj_Worksheet = obj_Workbook.Sheets(hoja)
'
   hoja = ApExcel.Sheets(1).Name 'obj_Workbook.Sheets(hoja)
   
End If

hoja = "[" & hoja & "$" & "]"
RsExcel.Open "SELECT * FROM " & hoja, cs

If RsExcel.EOF Then Exit Sub

RsExcel.MoveFirst

i = 1

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<UpdateReceta>"

prbStatus.Max = 1
lblStatus.Visible = True
prbStatus.Visible = True
prbStatus.Min = 0
lngRow = 0

Frame1.Enabled = False
prbStatus.Max = RsExcel.RecordCount

lblStatus.Caption = "Preparando datos para actualizar"

Dim XL As New excel.Application 'Crea el objeto excel

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), Me.HelpContextID, "", "", "")

Do While RsExcel.EOF <> True
           
   If RsExcel.Fields(0).Value = "*" Then Exit Do
           
   Ceco = ""
   Regimen = 0
   Servicio = 0
   Fecha = 0
   RecetaOrigen = 0
   NumLin = 0
   RecetaDestino = 0
   Ponderacion = 0
   Raciones = 0
   
   'Ceco
   If Not IsNull(RsExcel.Fields(0).Value) Then
      
      Ceco = RsExcel.Fields(0).Value
      
   End If
                
   'Regimen
   If IsNumeric(RsExcel.Fields(1).Value) Then
               
      Regimen = RsExcel.Fields(1).Value
            
   End If
                
   'Servicio
   If IsNumeric(RsExcel.Fields(3).Value) Then
               
      Servicio = RsExcel.Fields(3).Value
            
   End If
                
   'Fecha
   If IsNumeric(RsExcel.Fields(7).Value) Then
               
      Fecha = RsExcel.Fields(7).Value
            
   End If
                
   'Receta Origen
   If IsNumeric(RsExcel.Fields(8).Value) Then
               
      RecetaOrigen = RsExcel.Fields(8).Value
            
   End If

   'Numero Linea
   If IsNumeric(RsExcel.Fields(17).Value) Then
               
      NumLin = RsExcel.Fields(17).Value
            
   End If
                
   'Receta Destino
   If IsNumeric(RsExcel.Fields(18).Value) Then
    
      If RsExcel.Fields(18).Value > 0 Then
      
         RecetaDestino = RsExcel.Fields(18).Value
                  
      End If
      
   End If

   'Ponderacion 14
   If IsNumeric(RsExcel.Fields(14).Value) Then
               
      Ponderacion = RsExcel.Fields(14).Value
            
   End If
   
   'Raciones 15
   If IsNumeric(RsExcel.Fields(15).Value) Then
               
      Raciones = RsExcel.Fields(15).Value
            
   End If
   
   If Trim(Ceco) <> "" And Regimen <> 0 And Servicio <> 0 And Fecha <> 0 Then
      
      MyBuffer = MyBuffer & " <Receta"
      MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
      MyBuffer = MyBuffer & " Reg = " & Chr(34) & Regimen & Chr(34)
      MyBuffer = MyBuffer & " Ser = " & Chr(34) & Servicio & Chr(34)
      MyBuffer = MyBuffer & " Fec = " & Chr(34) & Fecha & Chr(34)
      MyBuffer = MyBuffer & " ROri = " & Chr(34) & RecetaOrigen & Chr(34)
      MyBuffer = MyBuffer & " Nli = " & Chr(34) & NumLin & Chr(34)
      MyBuffer = MyBuffer & " RDes = " & Chr(34) & RecetaDestino & Chr(34)
      MyBuffer = MyBuffer & " Pon = " & Chr(34) & Ponderacion & Chr(34)
      MyBuffer = MyBuffer & " Rac = " & Chr(34) & Raciones & Chr(34)
   
      MyBuffer = MyBuffer & "/>"
        
   End If
   
   DoEvents
           
   RsExcel.MoveNext
   
   lngRow = lngRow + 1
   prbStatus.Value = lngRow
   
   If i > 1000 Then
      
      fg_carga ""
      lblStatus.Caption = "Actualizando receta minuta bloque"
      
      MyBuffer = MyBuffer & "</UpdateReceta>"
      
      Set RS = vg_db.Execute("sgpadm_Upd_XmlCambioRecetaMinutaBloque_V02 '" & MyBuffer & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "', '" & IIf(Check1(0).Value = 1, "1", "0") & "', '" & IIf(Check1(1).Value = 1, "1", "0") & "', '" & IIf(Check1(2).Value = 1, "1", "0") & "'")
   
      If Not RS.EOF Then
        
         If RS(0) > 0 Or RS(0) < 0 Then
        
            lblStatus.Visible = False
            prbStatus.Visible = False
            Frame1.Enabled = True
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
  
            NomArchivoExcel = fg_ArchivoXls("ReporteError_actualizacionminutabloque_Recetas")
                    
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
                   
'            Set obj_Excel = Nothing
'            Set obj_Workbook = Nothing

            '-- Cerrar aplicación excel
            ApExcel.Workbooks(NomArchivo).Close SaveChanges:=False
            
            ApExcel.Visible = False
            ApExcel.Application.Visible = False
            ApExcel.Application.Quit
            Set ApExcel = Nothing

      
            Exit Sub
        
         Else
        
            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")
            'MsgBox "Proceso finalizado sin problema", vbInformation + vbOKOnly, Me.Caption
        
         End If
        
      End If
      RS.Close
      Set RS = Nothing
   
      fg_descarga
      lblStatus.Caption = "Preparando datos para actualizar"
      i = 1

      Let MyBuffer = ""
      Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
      Let MyBuffer = MyBuffer & "<UpdateReceta>"
   
   End If
   i = i + 1
        
Loop
        
RsExcel.Close
Set RsExcel = Nothing
    
'cn.Close
Set cn = Nothing
    
MyBuffer = MyBuffer & "</UpdateReceta>"

Set RS = vg_db.Execute("sgpadm_Upd_XmlCambioRecetaMinutaBloque_V02 '" & MyBuffer & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "', '" & IIf(Check1(0).Value = 1, "1", "0") & "', '" & IIf(Check1(1).Value = 1, "1", "0") & "', '" & IIf(Check1(2).Value = 1, "1", "0") & "'")

If Not RS.EOF Then
  
   If RS(0) > 0 Or RS(0) < 0 Then
  
      lblStatus.Visible = False
      prbStatus.Visible = False
      Frame1.Enabled = True
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
  
      NomArchivoExcel = fg_ArchivoXls("ReporteError_actualizaciónminutabloque_Recetas")
                    
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
'
'      MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, Me.Caption
      
'      Set obj_Excel = Nothing
'      Set obj_Workbook = Nothing

        '-- Cerrar aplicación excel
        ApExcel.Workbooks(NomArchivo).Close SaveChanges:=False
        
        ApExcel.Visible = False
        ApExcel.Application.Visible = False
        ApExcel.Application.Quit
        Set ApExcel = Nothing

      Exit Sub
  
   Else
  
      Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")
      
      MsgBox "Proceso finalizado sin problema", vbInformation + vbOKOnly, Me.Caption
  
   End If
  
End If
RS.Close
Set RS = Nothing

'Set obj_Excel = Nothing
'obj_Workbook.Close
'Set obj_Workbook = Nothing
'Set obj_Worksheet = Nothing

'-- Cerrar aplicación excel

ApExcel.Workbooks(NomArchivo).Close SaveChanges:=False

ApExcel.Visible = False
ApExcel.Application.Visible = False
ApExcel.Application.Quit
Set ApExcel = Nothing

'obj_Workbook.Close
'Set obj_Workbook = Nothing
'Set obj_Worksheet = Nothing

lblStatus.Visible = False
prbStatus.Visible = False
Frame1.Enabled = True
fg_descarga

Exit Sub

Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub ActualizarCambioComensales()

On Error GoTo Man_Error

Dim i               As Long
Dim j               As Long
Dim PathXls         As String
Dim dbexcel         As Database
Dim cn              As ADODB.Connection
Dim RS              As New ADODB.Recordset
Dim SheetName       As String
Dim sSheetName      As String
Dim hoja            As String
Dim IndColumna      As Long
Dim MyBuffer        As String
Dim MyBufferTotal   As String
Dim File_Ext        As String
Dim cs              As String

Dim strArray()      As String
Dim intCount        As Integer
Dim UltRow          As Long

Dim Ceco            As String
Dim Regimen         As Long
Dim Servicio        As Long
Dim Comensales      As Long
Dim Fecha           As Long

Dim lngRow          As Long
Dim NomArchivoExcel As String
Dim NomArchivo      As String

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object

Dim ApExcel As New excel.Application
Set ApExcel = New excel.Application

Dim RsExcel As New ADODB.Recordset
Set cn = New ADODB.Connection

NomArchivo = Dir(CD.FileName)

PathXls = Trim(fpText1.text)
File_Ext = UCase(Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, "."))))
hoja = Combo1.text

RsExcel.CursorLocation = adUseClient
RsExcel.CursorType = adOpenKeyset
RsExcel.LockType = adLockBatchOptimistic
 
cs = "DRIVER=Microsoft Excel Driver (*.xls);DBQ=" & PathXls & ";HDR=NO;IMEX=1;"
'Set obj_Excel = CreateObject("Excel.Application")
 
' -- Abrir el libro
'Set obj_Workbook = obj_Excel.Workbooks.Open(PathXls)
ApExcel.Workbooks.Open FileName:=PathXls
' -- referencia la Hoja, por defecto la hoja activa

If sSheetName = vbNullString Then
   
'   Set obj_Worksheet = obj_Workbook.ActiveSheet
      
   For i = 1 To ApExcel.Sheets.count
         
       If hoja = ApExcel.Sheets(i).Name Then
        
          hoja = ApExcel.Sheets(i).Name
          Exit For
            
       End If
    
   Next
      
Else

'   Set obj_Worksheet = obj_Workbook.Sheets(hoja)
'
   hoja = ApExcel.Sheets(1).Name 'obj_Workbook.Sheets(hoja)
   
End If

hoja = "[" & hoja & "$" & "]"
RsExcel.Open "SELECT * FROM " & hoja, cs

If RsExcel.EOF Then Exit Sub

RsExcel.MoveFirst

i = 1

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<UpdateComensales>"

prbStatus.Max = 1
lblStatus.Visible = True
prbStatus.Visible = True
prbStatus.Min = 0
lngRow = 0

Frame1.Enabled = False
prbStatus.Max = RsExcel.RecordCount

lblStatus.Caption = "Preparando datos para actualizar"

Dim XL As New excel.Application 'Crea el objeto excel

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), Me.HelpContextID, "", "", "")

Do While RsExcel.EOF <> True
           
   If RsExcel.Fields(0).Value = "*" Then Exit Do
           
   Ceco = ""
   Regimen = 0
   Servicio = 0
   Fecha = 0
   Comensales = 0
   
   'Ceco
   If Not IsNull(RsExcel.Fields(0).Value) Then
      
      Ceco = RsExcel.Fields(0).Value
      
   End If
                
   'Regimen
   If IsNumeric(RsExcel.Fields(1).Value) Then
               
      Regimen = RsExcel.Fields(1).Value
            
   End If
                
   'Servicio
   If IsNumeric(RsExcel.Fields(3).Value) Then
               
      Servicio = RsExcel.Fields(3).Value
            
   End If
                
   'Fecha
   If IsNumeric(RsExcel.Fields(5).Value) Then
               
      Fecha = RsExcel.Fields(5).Value
            
   End If
                
   'Comensales
   If IsNumeric(RsExcel.Fields(6).Value) Then
               
      Comensales = RsExcel.Fields(6).Value
            
   End If

   MyBuffer = MyBuffer & " <Comensales"
   MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
   MyBuffer = MyBuffer & " Reg = " & Chr(34) & Regimen & Chr(34)
   MyBuffer = MyBuffer & " Ser = " & Chr(34) & Servicio & Chr(34)
   MyBuffer = MyBuffer & " Fec = " & Chr(34) & Fecha & Chr(34)
   MyBuffer = MyBuffer & " Com = " & Chr(34) & Comensales & Chr(34)
   MyBuffer = MyBuffer & "/>"
        
   DoEvents
           
   RsExcel.MoveNext
   
   lngRow = lngRow + 1
   prbStatus.Value = lngRow
   
   If i > 1000 Then
      
      fg_carga ""
      lblStatus.Caption = "Actualizando comensales minuta bloque"
      
      MyBuffer = MyBuffer & "</UpdateComensales>"
      
      Set RS = vg_db.Execute("sgpadm_Upd_XmlCambioComensalesMinutaBloque_V01 '" & MyBuffer & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")
   
      If Not RS.EOF Then
        
         If RS(0) > 0 Or RS(0) < 0 Then
        
            lblStatus.Visible = False
            prbStatus.Visible = False
            Frame1.Enabled = True
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
  
            NomArchivoExcel = fg_ArchivoXls("ReporteError_actualizacionminutabloque_Comensales")
                    
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
                   
            '-- Cerrar aplicación excel
            ApExcel.Workbooks(NomArchivo).Close SaveChanges:=False

            ApExcel.Visible = False
            ApExcel.Application.Visible = False
            ApExcel.Application.Quit
            Set ApExcel = Nothing
            
            Exit Sub
        
         Else
        
            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")
            'MsgBox "Proceso finalizado sin problema", vbInformation + vbOKOnly, Me.Caption
        
         End If
        
      End If
      RS.Close
      Set RS = Nothing
   
      fg_descarga
      lblStatus.Caption = "Preparando datos para actualizar"
      i = 1

      Let MyBuffer = ""
      Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
      Let MyBuffer = MyBuffer & "<UpdateComensales>"
   
   End If
   i = i + 1
        
Loop
        
RsExcel.Close
Set RsExcel = Nothing
    
'cn.Close
Set cn = Nothing
    
MyBuffer = MyBuffer & "</UpdateComensales>"

Set RS = vg_db.Execute("sgpadm_Upd_XmlCambioComensalesMinutaBloque_V01 '" & MyBuffer & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

If Not RS.EOF Then
  
   If RS(0) > 0 Or RS(0) < 0 Then
  
      lblStatus.Visible = False
      prbStatus.Visible = False
      Frame1.Enabled = True
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
  
      NomArchivoExcel = fg_ArchivoXls("ReporteError_actualizaciónminutabloque_Comensales")
                    
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

      '-- Cerrar aplicación excel
      ApExcel.Workbooks(NomArchivo).Close SaveChanges:=False
        
      ApExcel.Visible = False
      ApExcel.Application.Visible = False
      ApExcel.Application.Quit
      Set ApExcel = Nothing

      Exit Sub
  
   Else
  
      Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")
      
      MsgBox "Proceso finalizado sin problema", vbInformation + vbOKOnly, Me.Caption
  
   End If
  
End If
RS.Close
Set RS = Nothing

'Set obj_Excel = Nothing
'obj_Workbook.Close
'Set obj_Workbook = Nothing
'Set obj_Worksheet = Nothing

'-- Cerrar aplicación excel
ApExcel.Workbooks(NomArchivo).Close SaveChanges:=False

ApExcel.Visible = False
ApExcel.Application.Visible = False
ApExcel.Application.Quit
Set ApExcel = Nothing

lblStatus.Visible = False
prbStatus.Visible = False

Frame1.Enabled = True
fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
End Sub


