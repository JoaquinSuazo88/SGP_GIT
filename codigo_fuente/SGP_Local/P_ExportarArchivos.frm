VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form P_ExportarArchivos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Guía CD"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
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
      Left            =   5280
      TabIndex        =   8
      Top             =   2160
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
      Left            =   7440
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   7215
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1755
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
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
         TabIndex        =   3
         Top             =   1560
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
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
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
         TabIndex        =   5
         Top             =   1320
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
         TabIndex        =   4
         Top             =   315
         Width           =   1275
      End
   End
End
Attribute VB_Name = "P_ExportarArchivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public lc_Aux As String

Private obj_Excel       As Object
Private obj_Workbook    As Object
Private obj_Worksheet   As Object

Dim ObjExcel            As Excel.Application
Dim ObjW                As Excel.Workbook
Dim NombreHoja          As String
Dim RutaNombreArchivo   As String
Dim MsgTitulo           As String

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

    Case 0
    
    
        If ValidaDatosGuiaCd = False Then Exit Sub
        
        M_Traspa.SetFocus
'        If MsgBox("Esta seguro ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        
    Case 1
    
        Unload Me
        
    Case 1
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_carga ""
Me.HelpContextID = vg_OpcM

MsgTitulo = "Importar Guía CD"
fg_centra Me

Command1(0).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "0", False, True)

Command1(0).Enabled = False

fpText1.text = ""
fpText1.Enabled = True
Combo1.Enabled = True
Combo1.ListIndex = -1

lblStatus.Visible = False
prbStatus.Visible = False

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_ButtonHit(Button As Integer, NewIndex As Integer)

On Error GoTo Man_Error

Dim fromRihgt As String
Dim myPath    As String

Command1(0).Enabled = False

CD.FileName = ""
CD.Filter = "Archivos xlsx|*.xlsx|Archivos xls|*.xls"
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

    Dim i                 As Integer
    Dim HojaPro           As Boolean
    Dim ObjHoja           As Excel.Worksheet
    Dim Extension         As String
    Dim Formato           As Integer
    Dim fso               As Object
    Dim ExtensionExcel    As String
    
    'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    
    Formato = 20 '23 '62 '24 '6 '"CSV"
    Extension = ".txt"
    
    ExtensionExcel = ""
    ExtensionExcel = Right(CD.FileName, Len(CD.FileName) - (InStr(CD.FileName, ".")))
    If UCase(ExtensionExcel) = "XLS" Then
    
       RutaNombreArchivo = Mid$(CD.FileName, 1, Len(CD.FileName) - 4)
       
    ElseIf UCase(ExtensionExcel) = "XLSX" Then
                
       RutaNombreArchivo = Mid$(CD.FileName, 1, Len(CD.FileName) - 5)
             
    End If
    
    RutaNombreArchivo = dir_trabajo & SacarNombreArchivo(RutaNombreArchivo, "\") & Extension
    
    'validar si existe RutaNombreArchivo se borra
    If fso.FileExists(RutaNombreArchivo) Then
   
       'Borrar archivo
       Kill (RutaNombreArchivo)
    
    End If
    
    Set obj_Excel = CreateObject("Excel.Application")
    
    With obj_Excel
        
        .DisplayAlerts = False
        ' abre
        .Workbooks.Open (CD.FileName)
   
        ' referencia a esta sheet
        Set ObjHoja = .Sheets(1)
        
        NombreHoja = .Sheets(1).Name
        
        ' selecciona la hoja actual con el método Select
        ObjHoja.Select
        
        ' Copia la hoja entera con el método Copy
        ObjHoja.Copy
        
        ' exporta la hoja individual al formato indicado con SaveAs
                               'Extension, _

        .ActiveWorkbook.SaveAs FileName:=Trim(RutaNombreArchivo), _
                               FileFormat:=Formato, _
                               Password:="", _
                               WriteResPassword:="", _
                               ReadOnlyRecommended:=False, _
                               CreateBackup:=False
    
    End With

    'obj_Workbook.Close
    obj_Excel.Quit
    Set ObjHoja = Nothing
    Set obj_Excel = Nothing
    Set fso = Nothing
    
    Command1(0).Enabled = True
    
End If

Exit Sub
Man_Error:
fg_descarga
Set fso = Nothing
If Err = 5 Then Resume Next: Exit Sub
If Err = 55 Then MsgBox "Archivo esta abierto, proceso cancelado": Exit Sub
If Err = 462 Or Err = 1004 Or Err = 438 Or Err = -2147417848 Or Err = 70 Then Resume Next
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
Resume Next

End Sub

Private Function ValidaDatosGuiaCd() As Boolean

On Error GoTo Man_Error

Dim SheetName  As String
Dim RsExcel    As New ADODB.Recordset
Dim PathXls    As String
Dim File_Ext   As String
Dim dbexcel    As Database
Dim i          As Long
Dim CostoTotal As Double
Dim fso        As Object

Let ValidaDatosGuiaCd = True
 
Set RsExcel = New ADODB.Recordset
Dim Hoja       As String
Dim cs         As String
Dim sSheetName As String

'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
Set fso = CreateObject("Scripting.FileSystemObject")

'-------> validar si existe hoja en la planilla excel
If Trim(NombreHoja) = "" Then

    Call MsgBox("No existe hoja en la planilla excel", vbCritical, Me.Caption)
    Call fpText1.SetFocus
    Let ValidaDatosGuiaCd = False
    Exit Function

End If

PathXls = Trim(LimpiaDato(fpText1.text))
'-------> Validar Archivo Origen
If Trim(LimpiaDato(fpText1.text)) = "" Then
    
    Call MsgBox("Debe seleccionar archivo origen", vbCritical, Me.Caption)
    Call fpText1.SetFocus
    Let ValidaDatosGuiaCd = False
    Exit Function

End If

'validar si exista Archivo
If Not fso.FileExists(RutaNombreArchivo) Then

    Call MsgBox("No existe archivo " & RutaNombreArchivo, vbCritical, Me.Caption)
    Call fpText1.SetFocus
    Let ValidaDatosGuiaCd = False
    Set fso = Nothing
    Exit Function

End If

Set fso = Nothing

Dim LineaSplitted() As String
Dim strLineReg As String
Dim MyBuffer As String

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<EncGuiaCD>"

Dim XL              As New Excel.Application 'Crea el objeto excel
Dim Orden           As Long
Dim Ceco            As String
Dim NGuia           As Long
Dim FechaG          As Date
Dim CMaterial       As String
Dim DMaterial       As String
Dim Cantidad        As Double
Dim Precio          As Double
Dim precioa         As Double
Dim RS              As New ADODB.Recordset
Dim NomArchivoExcel As String
Dim Txt             As Boolean

'If Not Unix2Dos(RutaNombreArchivo) Then MsgBox "Problema formato", vbInformation + vbOKOnly, MsgTitulo: Exit Function
i = 1
Orden = 1
Txt = False

Open RutaNombreArchivo For Input As #1
Txt = True

Do While Not EOF(1)
        
   Line Input #1, strLineReg
   
   If strLineReg = "" Then
   
      Call MsgBox("No existe información...", vbCritical, Me.Caption)
      Let ValidaDatosGuiaCd = False
      Close #1
      Exit Function
   
   End If
   strLineReg = Replace(strLineReg, Chr(9), ";")
   strLineReg = Replace(strLineReg, Chr(34), "")
   strLineReg = Replace(strLineReg, Chr(44), ".")
   LineaSplitted = Split(strLineReg, ";")
   
   If i = 1 Then
            
      'Código Ceco
      If Not Trim(LineaSplitted(0)) = "Codigo Ceco" Then

         Call MsgBox("Formato Ceco no corresponde", vbCritical, Me.Caption)
         Let ValidaDatosGuiaCd = False
         Close #1
         Exit Function

      End If

      'Descripción Ceco
      If Not Trim(LineaSplitted(1)) = "Descripcion CeCo" Then

         Call MsgBox("Formato Descripción Ceco no corresponde", vbCritical, Me.Caption)
         Let ValidaDatosGuiaCd = False
         Close #1
         Exit Function

      End If

      'Nº GDD
      If Not Trim(LineaSplitted(2)) = "N° GDD" Then

         Call MsgBox("Formato Guía no corresponde", vbCritical, Me.Caption)
         Let ValidaDatosGuiaCd = False
         Close #1
         Exit Function

      End If

      'Fecha GDD
      If Not Trim(LineaSplitted(3)) = "Fecha GDD" Then

         Call MsgBox("Formato Fecha no corresponde", vbCritical, Me.Caption)
         Let ValidaDatosGuiaCd = False
         Close #1
         Exit Function

      End If

      'Código Producto
      If Not Trim(LineaSplitted(4)) = "Codigo Producto" Then

         Call MsgBox("Formato Código Producto no corresponde", vbCritical, Me.Caption)
         Let ValidaDatosGuiaCd = False
         Close #1
         Exit Function

      End If

      'Descripción Producto
      If Not Trim(LineaSplitted(5)) = "Descripcion Producto" Then

         Call MsgBox("Formato Descripción Producto no corresponde", vbCritical, Me.Caption)
         Let ValidaDatosGuiaCd = False
         Close #1
         Exit Function

      End If

      'Cantidad
      If Not Trim(LineaSplitted(6)) = "Cantidad" Then

         Call MsgBox("Formato Cantidad no corresponde", vbCritical, Me.Caption)
         Let ValidaDatosGuiaCd = False
         Close #1
         Exit Function

      End If

      'Precio
      If Not Trim(LineaSplitted(7)) = "Precio" Then

         Call MsgBox("Formato Precio no corresponde", vbCritical, Me.Caption)
         Let ValidaDatosGuiaCd = False
         Close #1
         Exit Function

      End If

    Else

       Ceco = ""
       NGuia = 0
'   FechaG = ""
       CMaterial = ""
       DMaterial = ""
       Cantidad = 0
       Precio = 0
       precioa = 0

       'Ceco
       If Not IsNull(Trim(LineaSplitted(0))) And Trim(LineaSplitted(0)) <> "" Then

          Ceco = Trim(LineaSplitted(0))
    
       Else
    
          Call MsgBox("Ceco debe ser alfanumerico", vbCritical, Me.Caption)
          Let ValidaDatosGuiaCd = False
          Close #1
          Exit Function
 

       End If

       'Numero Guia
       If IsNumeric(Trim(LineaSplitted(2))) And Trim(LineaSplitted(2)) <> "" Then

          NGuia = Trim(LineaSplitted(2))

       Else
    
          Call MsgBox("Guia debe ser numerico", vbCritical, Me.Caption)
          Let ValidaDatosGuiaCd = False
          Close #1
          Exit Function

       End If

       'Fecha
       If IsDate(Trim(LineaSplitted(3))) And Trim(LineaSplitted(3)) <> "" Then

          FechaG = Trim(LineaSplitted(3))

       Else
    
          Call MsgBox("Fecha debe ser formato fecha", vbCritical, Me.Caption)
          Let ValidaDatosGuiaCd = False
          Close #1
          Exit Function
    
       End If

       'Material
       If Not IsNull(Trim(LineaSplitted(4))) And Trim(LineaSplitted(4)) <> "" Then

          CMaterial = Trim(LineaSplitted(4))
          DMaterial = Trim(LineaSplitted(5))

       Else
    
          Call MsgBox("Material debe ser alfanumerico", vbCritical, Me.Caption)
          Let ValidaDatosGuiaCd = False
          Close #1
          Exit Function
    
       End If

       'Cantidad
       If Not IsNull(Trim(LineaSplitted(6))) And Trim(LineaSplitted(6)) <> "" Then

          Cantidad = Trim(LineaSplitted(6))

       Else
   
          Call MsgBox("Cantidad debe ser numerico", vbCritical, Me.Caption)
          Let ValidaDatosGuiaCd = False
          Close #1
          Exit Function

       End If

       'Precio
       If IsNumeric(Trim(LineaSplitted(7))) And Trim(LineaSplitted(7)) <> "" Then

          Precio = Trim(LineaSplitted(7))

       Else
   
          Call MsgBox("Precio debe ser numerico", vbCritical, Me.Caption)
          Let ValidaDatosGuiaCd = False
          Close #1
          Exit Function
   
       End If

       If Ceco <> "" And NGuia > 0 And CStr(FechaG) <> "" And CMaterial <> "" And IsNumeric(Cantidad) And IsNumeric(Precio) Then

          MyBuffer = MyBuffer & " <GuiaCD"
          MyBuffer = MyBuffer & " Orden = " & Chr(34) & Orden & Chr(34)
          MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
          MyBuffer = MyBuffer & " NGuia = " & Chr(34) & NGuia & Chr(34)
          MyBuffer = MyBuffer & " FechaG = " & Chr(34) & Format(FechaG, "yyyy/mm/dd") & Chr(34)
          MyBuffer = MyBuffer & " CMaterial = " & Chr(34) & CMaterial & Chr(34)
          MyBuffer = MyBuffer & " DMaterial = " & Chr(34) & DMaterial & Chr(34)
          MyBuffer = MyBuffer & " Cantidad = " & Chr(34) & Cantidad & Chr(34)
          MyBuffer = MyBuffer & " Precio = " & Chr(34) & Precio & Chr(34)
          MyBuffer = MyBuffer & "/>"

       End If
       
       Orden = Orden + 1

    End If

    i = i + 1

Loop
Txt = False
Close #1

Dim Largo              As Long
Dim NombreArchivoExcel As String

Dim xlApp    As Object
Dim xlWb     As Object
Dim xlWs     As Object
'Dim XL              As New Excel.Application 'Crea el objeto excel

MyBuffer = MyBuffer & "</EncGuiaCD>"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgp_Sel_XmlValidarExportarArchivosGuiaCD '" & MyBuffer & "', '" & MuestraCasino(1) & "', " & Val(fg_codigocbo(M_Traspa.Combo1, 1, 10, "")) & "")

If Not RS.EOF Then

   If RS(0) > 0 Or RS(0) < 0 Then

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

      xlApp.Columns("A:A").Select
      xlApp.Selection.Delete Shift:=xlToLeft

      NomArchivoExcel = fg_ArchivoXls("ReporteError_GuiaCD")

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
      ValidaDatosGuiaCd = False

      Exit Function

   Else

      Dim vMotivo() As Variant
      Dim RS1       As New ADODB.Recordset
      Dim j         As Long
      Dim lisnom    As String
      Dim liscod    As String
      
      i = 1
      
      If RS1.State = 1 Then RS1.Close
      RS1.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      '-------> Cargar vector
      Set RS1 = vg_db.Execute("sgp_sel_motivoGuiaCD ")
      
      If Not RS1.EOF Then
      
         ReDim vMotivo(RS1.RecordCount, 2)
         
         Do While Not RS1.EOF
         
            vMotivo(i, 1) = RS1![IdMotivo]
            vMotivo(i, 2) = RS1![Descripcion Motivo]
            
            i = i + 1
            
            RS1.MoveNext
            
         Loop
      
      Else
      
            RS.Close
            Set RS = Nothing
            
            RS1.Close
            Set RS1 = Nothing
            
            MsgBox "No estan cargado el concepto de motivos...", vbExclamation + vbOKOnly, MsgTitulo
            
            ValidaDatosGuiaCd = False

            Exit Function
     
      End If
      RS1.Close
      Set RS1 = Nothing
      
      M_Traspa.Option1(0).Enabled = False
      
      '-------> Entrada dato encabezado guía
      M_Traspa.fpLongInteger1(0).Value = RS!NGuia
      vg_FechaEmision_GGD = "0000:00:00"
      vg_FechaEmision_GGD = Format(RS!FechaG, "dd/mm/yyyy")
      
      If RS1.State = 1 Then RS1.Close
      RS1.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      '-------> Cargar codigo CD
      M_Traspa.fpText1(1).text = "CD"
      Set RS1 = vg_db.Execute("SELECT isnull(cli_nombre,'') as cli_nombre FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(M_Traspa.fpText1(1).text)) & "' " & _
                              "AND cli_codigo <> '" & LimpiaDato(Trim(M_Traspa.fpText1(0).text)) & "' AND (cli_tipo=2 OR cli_tipo = 0)")
      
      If Not RS1.EOF Then
      
         M_Traspa.fpayuda(1).Caption = Trim(RS1!cli_nombre)
 
      Else
            
            RS.Close
            Set RS = Nothing
            
            RS1.Close
            Set RS1 = Nothing
            
            MsgBox "Central Distribución no esta creada como (CD)...", vbExclamation + vbOKOnly, MsgTitulo
            M_Traspa.fpText1(1) = ""
            
            ValidaDatosGuiaCd = False

            Exit Function

      End If
      RS1.Close
      Set RS1 = Nothing
      
      M_Traspa.fpLongInteger1(0).Enabled = False
      M_Traspa.fpDateTime1(0).Enabled = True 'False
      M_Traspa.fpText1(1).Enabled = False
      
      '-------> mostrar codigo sac
      M_Traspa.Text2(0).Visible = True
      M_Traspa.Text2(1).Visible = True
      
      M_Traspa.vaSpread1.Visible = False
      M_Traspa.vaSpread1.MaxRows = 0
      M_Traspa.vaSpread1.MaxRows = RS.RecordCount
      i = 1
      
      M_Traspa.vaSpread1.ColWidth(2) = 39.88 '47.88
     
      Do While Not RS.EOF
               
         M_Traspa.vaSpread1.Row = i
         
         M_Traspa.vaSpread1.Col = 1 'codigo producto
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.vaSpread1.text = RS!pro_codigo
         
         M_Traspa.vaSpread1.Col = 2 'nombre producto
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.vaSpread1.text = RS!pro_nombre
         
         M_Traspa.vaSpread1.Col = 3 'unidad medida producto
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.vaSpread1.text = RS!uni_nomcor
         
         M_Traspa.vaSpread1.Col = 4 'cantidad documento
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.vaSpread1.text = RS!Cantidad
         
         M_Traspa.vaSpread1.Col = 5 'precio documento
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.vaSpread1.text = RS!Precio
         
         M_Traspa.vaSpread1.Col = 6 'total cantidad * precio
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.vaSpread1.text = (RS!Cantidad * RS!Precio)
         
         M_Traspa.vaSpread1.Col = 7 'cantidad recibida
         M_Traspa.vaSpread1.Lock = False
         M_Traspa.vaSpread1.text = RS!Cantidad
         
         '-------> calcular costo total
         CostoTotal = CostoTotal + (RS!Cantidad * RS!Precio)
         
         M_Traspa.vaSpread1.Col = 8 '
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.vaSpread1.text = "N" 'No bloquedo
                
         M_Traspa.vaSpread1.Col = 10 'producto controla stock
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.vaSpread1.text = IIf(IsNull(RS!pro_ctrsto), "N", IIf(RS!pro_ctrsto = 1, "S", "N"))
                
         M_Traspa.vaSpread1.Col = 11 'precio ponderado
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.vaSpread1.text = RS!ppd_propon
                
         M_Traspa.vaSpread1.Col = 12
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.vaSpread1.text = "N"
                
         M_Traspa.vaSpread1.Col = 16 'material sap
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.vaSpread1.ColHidden = False
         M_Traspa.vaSpread1.text = IIf(IsNull(RS!CMaterial), "", RS!CMaterial)
                
         M_Traspa.vaSpread1.Col = 17 'descripción material sap
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.vaSpread1.text = IIf(IsNull(RS!fcs_DenMaterial), "", RS!fcs_DenMaterial)
                
         M_Traspa.vaSpread1.Col = 16 'mover material sap text
'         M_Traspa.vaSpread1.Lock = False
         M_Traspa.Text2(0).text = Trim(M_Traspa.vaSpread1.text)
                
         M_Traspa.vaSpread1.Col = 17 'mover descripcion material sap text
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.Text2(1).text = Trim(M_Traspa.vaSpread1.text)
                
         M_Traspa.vaSpread1.Col = 19
         M_Traspa.vaSpread1.Lock = True
         M_Traspa.vaSpread1.ColHidden = False
         
         M_Traspa.vaSpread1.Col = 9
         M_Traspa.vaSpread1.text = Format(RS!bod_canmer, fg_Pict(9, vg_DCa))
         
         lisnom = ""
         liscod = ""
         j = 1
         
         For j = 1 To UBound(vMotivo)
            
             If vMotivo(j, 1) <> "" Then
               
                lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vMotivo(j, 2))
                liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vMotivo(j, 1)
            
             End If
        
         Next j
         
         M_Traspa.vaSpread1.Col = 19
         M_Traspa.vaSpread1.TypeComboBoxList = lisnom
        
         M_Traspa.vaSpread1.Col = 20
         M_Traspa.vaSpread1.TypeComboBoxList = liscod
        
         M_Traspa.vaSpread1.Col = 21
         M_Traspa.vaSpread1.text = "GuiaCD"
        
         i = i + 1
         
         RS.MoveNext
            
      Loop
        
      ValidaDatosGuiaCd = True
    
      RS.Close
      Set RS = Nothing
    
      M_Traspa.Toolbar2.Enabled = True
      M_Traspa.vaSpread1.Visible = True
      M_Traspa.Label2.Caption = Format(CostoTotal, fg_Pict(9, vg_DCa))
        
      Gl_Ac_Botones M_Traspa, 4, 6, ""
      vg_GuiaCD = "1"
      Unload Me
      Exit Function
    
   End If

End If
RS.Close
Set RS = Nothing

M_Traspa.Toolbar2.Enabled = True

Exit Function
Man_Error:
    
    If Txt Then
       
       Close #1
       
    End If
    
    M_Traspa.Toolbar2.Enabled = True
    
    ValidaDatosGuiaCd = False
    fg_descarga
    
    If Err = 462 Or Err = 1004 Or Err = 438 Then Resume Next
    
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Function

