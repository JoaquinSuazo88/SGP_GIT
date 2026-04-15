VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form E_ExcelVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportación Excel Varios"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   13575
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin RichTextLib.RichTextBox RTF 
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"E_ExcelVarios.frx":0000
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Left            =   3915
         TabIndex        =   8
         Top             =   5190
         Width           =   7110
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   9
            Top             =   135
            Width           =   7005
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   3000
         TabIndex        =   6
         Top             =   5190
         Width           =   900
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   7
            Top             =   135
            Width           =   795
         End
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
         Height          =   540
         Left            =   12090
         TabIndex        =   5
         Top             =   5760
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exportar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   10680
         TabIndex        =   4
         Top             =   5760
         Width           =   1275
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4815
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   11085
         _Version        =   393216
         _ExtentX        =   19553
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
         SpreadDesigner  =   "E_ExcelVarios.frx":008B
         VisibleCols     =   3
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   4455
         TabIndex        =   11
         Top             =   5850
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         ButtonStyle     =   2
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
         Text            =   "01/09/2013"
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
         Left            =   8220
         TabIndex        =   12
         Top             =   5850
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         ButtonStyle     =   2
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
         Text            =   "28/09/2013"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha desde"
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
         Left            =   3120
         TabIndex        =   14
         Top             =   5925
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha hasta"
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
         Left            =   6915
         TabIndex        =   13
         Top             =   5925
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3600
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "No Mostrar Recetas No Vigentes"
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
         Left            =   10320
         TabIndex        =   15
         Top             =   1080
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "E_ExcelVarios.frx":605F
         Left            =   3000
         List            =   "E_ExcelVarios.frx":6061
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   8115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Org. Compras"
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
         Left            =   2160
         TabIndex        =   16
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Informes"
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
         Left            =   2160
         TabIndex        =   2
         Top             =   675
         Width           =   735
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
Attribute VB_Name = "E_ExcelVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public lc_Aux As String
Dim MsgTitulo As String

Private Sub Check1_Click()
 
On Error GoTo Man_Error

 CargarGrilla (3)
 
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Combo1_Click(Index As Integer)

On Error GoTo Man_Error

Check1.Visible = False

Label1(3).Visible = False
Text1.Visible = False

Select Case Combo1(Index).ListIndex

Case 0
    
    CargarGrilla (1)
    
Case 1

    CargarGrilla (2)
    
Case 2, 3, 6

    If Combo1(Index).ListIndex = 3 Then
    
       Label1(3).Visible = True
       Text1.Visible = True
    
    End If
    
    Check1.Visible = True
    CargarGrilla (3)

Case 4, 5
    
    CargarGrilla (4)
    
Case 7
    
    CargarGrilla (8)

Case 10
    
    CargarGrilla (11)
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim Sql             As String
Dim Sql1            As String
Dim Sql2            As String
Dim codCeco         As String
Dim NomArchivoExcel As String
Dim Extension       As String
Dim seleccion       As Integer
Dim i               As Long
Dim MyBuffer        As String

Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object

If Not ValidarDatos Then Exit Sub

'--> Concatenar codigo
codCeco = ""
Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"

Select Case Val(fg_codigocbo(Combo1, 0, 2, ""))
       
Case 1, 2

   Let MyBuffer = MyBuffer & "<Ingredientes>"
   
Case 3, 4, 7

   Let MyBuffer = MyBuffer & "<Recetas>"

Case 5, 6
   
   Let MyBuffer = MyBuffer & "<UltPlan>"
   
End Select
   
If Val(fg_codigocbo(Combo1, 0, 2, "")) <> 9 And Val(fg_codigocbo(Combo1, 0, 2, "")) <> 10 Then
   
For i = 1 To vaSpread1.MaxRows
       
    vaSpread1.Row = i
    vaSpread1.Col = 1 'Seleccion
    seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
    If seleccion = 1 And vaSpread1.RowHidden = False Then

       vaSpread1.Col = 2
       codCeco = codCeco & "'" & vaSpread1.text & "', "
       
       Select Case Val(fg_codigocbo(Combo1, 0, 2, ""))
       
       Case 1, 2
       
           MyBuffer = MyBuffer & " <Ing"
           MyBuffer = MyBuffer & " Ingr = " & Chr(34) & vaSpread1.text & Chr(34)
       
       Case 3, 4, 7
       
           MyBuffer = MyBuffer & " <Rec"
           MyBuffer = MyBuffer & " Receta = " & Chr(34) & vaSpread1.text & Chr(34)
       
       Case 5, 6
       
           MyBuffer = MyBuffer & " <Plan"
           MyBuffer = MyBuffer & " Ceco = " & Chr(34) & vaSpread1.text & Chr(34)
       
       End Select

       MyBuffer = MyBuffer & "/>"

    End If
  
Next i

Select Case Val(fg_codigocbo(Combo1, 0, 2, ""))

Case 1, 2

     MyBuffer = MyBuffer & "</Ingredientes>"
     
Case 3, 4, 7

     MyBuffer = MyBuffer & "</Recetas>"
     
Case 5, 6

     MyBuffer = MyBuffer & "</UltPlan>"

End Select

If Trim(codCeco) = "" And Val(fg_codigocbo(Combo1, 0, 2, "")) <> 8 And Val(fg_codigocbo(Combo1, 0, 2, "")) <> 9 And Val(fg_codigocbo(Combo1, 0, 2, "")) <> 11 Then
  
   MsgBox "No existen datos seleccionados grilla cecos...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub

ElseIf Val(fg_codigocbo(Combo1, 0, 2, "")) = 8 And Trim(TextDet1(2).text) = "" Then

   MsgBox "Debe ingresar código estructura servicio...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub

End If

End If
'-------> Validar cantidad registro se sobre pase hoja excel
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Val(fg_codigocbo(Combo1, 0, 2, "")) = 1 Then

   Set RS = vg_db.Execute("sgpadm_Sel_ExcelAportesNutricionales_V02 '" & MyBuffer & "'")

ElseIf Val(fg_codigocbo(Combo1, 0, 2, "")) = 2 Then

   Set RS = vg_db.Execute("sgpadm_Sel_Excelingrediente_prodsgp_MaterialSap '" & MyBuffer & "'")

ElseIf Val(fg_codigocbo(Combo1, 0, 2, "")) = 3 Then

   Set RS = vg_db.Execute("sgpadm_Sel_ExcelRecetaResumenAportes '" & MyBuffer & "'")

ElseIf Val(fg_codigocbo(Combo1, 0, 2, "")) = 4 Then

   Set RS = vg_db.Execute("sgpadm_Sel_ExcelDetalleRecetas_V02 '" & MyBuffer & "', '" & Trim(LimpiaDato(Text1.text)) & "'")

ElseIf Val(fg_codigocbo(Combo1, 0, 2, "")) = 5 Then

   Set RS = vg_db.Execute("sgpadm_Sel_ExcelCecoUltimoPlanificacion '" & MyBuffer & "'")

ElseIf Val(fg_codigocbo(Combo1, 0, 2, "")) = 6 Then

   Set RS = vg_db.Execute("sgpadm_Sel_ExcelListaCantComensales '" & MyBuffer & "', " & Format(FpFecDesde.text, "yyyymmdd") & ", " & Format(FpFecHasta.text, "yyyymmdd") & " ")

ElseIf Val(fg_codigocbo(Combo1, 0, 2, "")) = 7 Then

   Set RS = vg_db.Execute("sgpadm_Sel_ExcelRecetaPlanificacionMaxima '" & MyBuffer & "', " & Format(FpFecDesde.text, "yyyymmdd") & " ")

ElseIf Val(fg_codigocbo(Combo1, 0, 2, "")) = 8 Then

   Set RS = vg_db.Execute("sgpadm_Sel_UbicarEstServicio '" & TextDet1(2).text & "' ")

ElseIf Val(fg_codigocbo(Combo1, 0, 2, "")) = 9 Or Val(fg_codigocbo(Combo1, 0, 2, "")) = 10 Then

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

        ProcesarArchivo CD.FileName
        
    Else
        'Si no mostramos un texto de advertencia de que no se seleccionó _
        ninguno, ya que FileName devuelve una cadena vacía
        MsgBox "No seleccionó ningún archivo", vbCritical

    End If
   
   Exit Sub

ElseIf Val(fg_codigocbo(Combo1, 0, 2, "")) = 11 Then

    LlevarExcelMetodo
    
    Exit Sub
    
End If

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
          
FpFecDesde.Enabled = False
FpFecHasta.Enabled = False
Frame1.Enabled = False
Frame2.Enabled = False

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
                
'ProgressBar1.Visible = False
'lbl_proceso.Visible = False
  
FpFecDesde.Enabled = True
FpFecHasta.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True

Exit Sub
Man_Error:
    Frame1.Enabled = True
    Frame2.Enabled = True
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub LlevarExcelMetodo()

On Error GoTo Man_Error

Dim XL              As Object

Set XL = CreateObject("Excel.application")
XL.Visible = True
XL.Workbooks.Add
    
XL.Workbooks.OpenText vg_Archxls, , 1, 1, , , , , , , True, "|", Local:=True
    
'XL.quit
Set XL = Nothing
    
MsgBox "Proceso Finalizo Correctamente", vbInformation, MsgTitulo

'Dim NashXl As excel.Application
'Dim IRow As Long, irow2 As Long
'
'fg_carga ""
'
'Set NashXl = CreateObject("excel.application")
'Set NashXl = New excel.Application
'NashXl.SheetsInNewWorkbook = 1
'NashXl.Workbooks.Add
'vaSpread1.AllowMultiBlocks = True
'vaSpread1.SetSelection 2, -1, vaSpread1.maxcols, vaSpread1.MaxRows + 3
'vaSpread1.ClipboardCopy
'IRow = vaSpread1.MaxRows + 5
''------- Pegar vaspread1(0) - Planilla Excel
'NashXl.Range("A1").Select
'NashXl.ActiveSheet.Paste
''------- Ajustar columna
'NashXl.Cells.Select
'NashXl.Cells.EntireColumn.AutoFit
'vaSpread1.AllowMultiBlocks = False
'vaSpread1.SetSelection 2, 0, vaSpread1.maxcols, vaSpread1.MaxRows
'fg_descarga
'
'NashXl.Visible = True

Exit Sub
Man_Error:
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Command2_Click()

On Error GoTo Man_Error

Unload Me

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()
    
On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
fg_centra Me
MsgTitulo = "Exportar Excel Varios"

FpFecHasta.text = Format(Date, "dd/mm/yyyy")
FpFecDesde.text = Format(Date, "dd/mm/yyyy")
vaSpread1.MaxRows = 0

TextDet1(2).text = ""
TextDet1(3).text = ""

Label1(3).Visible = False
Text1.Visible = False

Combo1(0).Clear
Combo1(0).AddItem "Ingredientes con aportes " & Space(150) & "(01)"
Combo1(0).AddItem "Ingrediente - productos sgp - Material Sap " & Space(150) & "(02)"
Combo1(0).AddItem "Resumen de Recetas con Aportes " & Space(150) & "(03)"
Combo1(0).AddItem "Detalle de Recetas " & Space(150) & "(04)"
Combo1(0).AddItem "Listado de Ceco Ultima Planificación " & Space(150) & "(05)"
Combo1(0).AddItem "Listado Cantidades Comensales x Sitios " & Space(150) & "(06)"
Combo1(0).AddItem "Listado Recetas en Planificación Maxima Fecha con Frecuencia " & Space(150) & "(07)"
Combo1(0).AddItem "Ubicar Estructura Servicio Minuta Bloque " & Space(150) & "(08)"
Combo1(0).AddItem "Transformar Recetas Optimum Excel" & Space(150) & "(09)"
Combo1(0).AddItem "Transformar Ingredientes Optimum Excel" & Space(150) & "(10)"
Combo1(0).AddItem "Listado Receta Metodo Preparación" & Space(150) & "(11)"


Combo1(0).ListIndex = 0

CargarGrilla 1

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

If IsDate(FpFecHasta.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub TextDet1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(TextDet1(Index).text, ",")

If Index = 2 Then
   TextDet1(3).text = ""
ElseIf Index = 3 Then
   TextDet1(2).text = ""
End If

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 4
    vaSpread1.text = 0
    
Next

Select Case Index

Case 2, 3
    
    vaSpread1.Visible = False
    
    If Trim(TextDet1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread1.Col = 1
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 4
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 4
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 4
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 4
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 1
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 4
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
    
    If Trim(TextDet1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 4
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextDet1(Index).text), SearchFlagsGreaterOrEqual)
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

Private Sub CargarGrilla(opcion As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i As Long
Dim CodigoReceta As String
Dim NombreReceta As String
Dim MetodoPreparacion As String

fg_carga ""

vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &HC0FFFF

vaSpread1.MaxRows = 0

If opcion = 1 Then ' Ingredientes aportes

   Set RS = vg_db.Execute("sgpadm_s_ingrediente_V02 16, '', '%" & UCase(LimpiaDato("")) & "%'")

ElseIf opcion = 2 Then ' Ingredientes

   Set RS = vg_db.Execute("sgpadm_s_ingrediente_V02 17, '', '%" & UCase(LimpiaDato("")) & "%'")

ElseIf opcion = 3 Then ' Receta

   Set RS = vg_db.Execute("sgpadm_s_receta_V07 25, 0, '%" & UCase(LimpiaDato("")) & "%', 0, 0, 0, '" & IIf(Check1.Value = 1, "x", "") & "'")

ElseIf opcion = 4 Then 'Ceco

   Set RS = vg_db.Execute("sgpadm_s_cliente_V02 55, '', '%" & UCase(LimpiaDato("")) & "%'")

ElseIf opcion = 8 Then

    TextDet1(2).text = ""
    TextDet1(3).text = ""
    fg_descarga
    Exit Sub
    
ElseIf opcion = 11 Then ' Receta metodo Preparación

   Set RS = vg_db.Execute("sgpadm_s_receta_V06_JPA 27, 0, '%" & UCase(LimpiaDato("")) & "%', 0, 0, 0, '" & IIf(Check1.Value = 1, "x", "") & "'")

   vg_Archxls = fg_ArchivoTxt
   Open vg_Archxls For Output As #1
   
   Print #1, "Código Receta" & "|" & "Nombre Receta" & "|" & "Metodo Preparación"
  
End If

If Not RS.EOF Then
  
  Do While Not RS.EOF
      
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      
      vaSpread1.Col = 1
      vaSpread1.text = "0"
      
      vaSpread1.Col = 2
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS(0)
      CodigoReceta = RS(0)
      
      vaSpread1.Col = 3
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = Trim(RS(1))
      NombreReceta = Trim(RS(1))
      
      
      vaSpread1.Col = 4
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = 0
      
      If opcion = 11 Then
        
         RTF.text = ""
         RTF.TextRTF = ""
         RTF.TextRTF = IIf(IsNull(RS(3)), "", (RS(3)))
         vaSpread1.Col = 5
'         vaSpread1.CellType = CellTypeStaticText
         Text2.text = ""
         Text2.text = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(RTF.text, Chr(13) + Chr(10), " "), Chr(27), ""), "*", vbCrLf), " f3", "o"), Chr(225) + "7", "-"), "f1", "ñ"), "e1", "a"), "e9", "e"), "b7", " ")
         vaSpread1.text = Text2.text
         MetodoPreparacion = Text2.text
         
         Print #1, CodigoReceta & "|" & NombreReceta & "|" & MetodoPreparacion
         
      End If
      
      RS.MoveNext
  
  Loop

Else
   
   fg_descarga
   vaSpread1.MaxRows = 0
   MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo

End If
RS.Close
Set RS = Nothing

If opcion = 11 Then

   Close #1

End If

fg_descarga

Exit Sub
Man_Error:
        
    If opcion = 11 Then

        Close #1

    End If
        
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Function ValidarDatos() As Boolean

Dim seleccion As Integer
Dim i As Long

ValidarDatos = True

If Val(fg_codigocbo(Combo1, 0, 2, "")) = 4 And Trim(LimpiaDato(Text1.text)) = "" Then

       MsgBox "Debe seleccinar Org. Compras...", vbExclamation + vbOKOnly, MsgTitulo
       ValidarDatos = False
       Exit Function

End If


If Val(fg_codigocbo(Combo1, 0, 2, "")) = 9 Or Val(fg_codigocbo(Combo1, 0, 2, "")) = 10 Then Exit Function

If Val(fg_codigocbo(Combo1, 0, 2, "")) = 6 Then

    '-------> Validar fechas
    If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
       
       MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
       ValidarDatos = False
       Exit Function
    
    End If
        
    If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
       
       MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
       ValidarDatos = False
       Exit Function
    
    End If

End If

'-------> Validar que exista un dato seleccionado
seleccion = 0
For i = 1 To vaSpread1.MaxRows
       
    vaSpread1.Row = i
    vaSpread1.Col = 1 'Seleccion
    seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
    If seleccion = 1 And vaSpread1.RowHidden = False Then
       
       Exit For
    
    End If
  
Next i
  
If seleccion = 0 And Val(fg_codigocbo(Combo1, 0, 2, "")) <> 8 And Val(fg_codigocbo(Combo1, 0, 2, "")) <> 9 Then
     
   MsgBox " Se debe seleccionar un item por lo menos", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function
  
End If

End Function

Sub encabezado(ByRef RS As ADODB.Recordset, ByRef xlWs As Object)

On Error GoTo Man_Error

Dim fldCount As Long
Dim icol As Long

'-------> Copy field names to the first row of the worksheet
fldCount = RS.Fields.count
For icol = 1 To fldCount
    xlWs.Cells(1, icol).Value = RS.Fields(icol - 1).Name
Next

Exit Sub
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo

End Sub

Sub ProcesarArchivo(NombreArchivo As String)

On Error GoTo Man_Error

Dim PathXls         As String
Dim File_Ext        As String
Dim NomHoja         As String
Dim dbexcel         As Database
Dim cn              As ADODB.Connection
Dim RS              As New ADODB.Recordset

Dim CodReceta       As String
Dim NumLin          As Integer
Dim Bom             As String
Dim NomReceta       As String
Dim ing             As String
Dim NomIngrediente  As String
Dim Gramaje         As Double
Dim Sitio           As String
Dim NomArchivoExcel As String
Dim RsExcel         As Object

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object
Dim XL              As Object

'Set RsExcel = dbexcel.OpenRecordset(SheetName)
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

RsExcel.Open ("SELECT * FROM [" & NomHoja & "]"), cn

If RsExcel.EOF Then Exit Sub

RsExcel.MoveFirst

vg_Archxls = fg_ArchivoTxt
Open vg_Archxls For Output As #1

CodReceta = ""
NomReceta = ""
Bom = ""
ing = ""
NomIngrediente = ""
Gramaje = 0
NumLin = 1
Sitio = ""

Print #1, "Código Receta" & "|" & "Nombre Receta" & "|" & "Código Ingrediente" & "|" & "Nombre Ingrediente" & "|" & "Num Lin" & "|" & "Gramaje" & "|" & "Bom Receta" & "|" & "Sitio"

Do While RsExcel.EOF <> True
           
   If RsExcel.Fields(0).Value = "*" Then Exit Do
   
   
       
   If Mid(RsExcel.Fields(1).Value, 1, 4) = "REC0" Or Mid(RsExcel.Fields(1).Value, 1, 4) = "ING0" Or Mid(RsExcel.Fields(1).Value, 1, 4) = "PRO0" Then
   
      CodReceta = RsExcel.Fields(1).Value
      NomReceta = RsExcel.Fields(3).Value
      NumLin = 1
   
   End If
   
   If Mid(RsExcel.Fields(13).Value, 1, 3) = "BOM" Then
   
      Bom = RsExcel.Fields(13).Value
      Sitio = IIf(IsNull(RsExcel.Fields(15).Value), "", RsExcel.Fields(15).Value)
   
   End If
   
   If Mid(RsExcel.Fields(0).Value, 1, 3) = "ING" Or Mid(RsExcel.Fields(0).Value, 1, 3) = "PRO" Then
   
      ing = RsExcel.Fields(0).Value
      NomIngrediente = IIf(IsNull(RsExcel.Fields(2).Value), "", RsExcel.Fields(2).Value)
      
      If IsNumeric(RsExcel.Fields(3).Value) Then
         
         Gramaje = IIf(IsNull(RsExcel.Fields(3).Value), 0, RsExcel.Fields(3).Value)
      
      End If
   
      Print #1, CodReceta & "|" & NomReceta & "|" & ing & "|" & NomIngrediente & "|" & NumLin & "|" & Gramaje & "|" & Bom & "|" & Sitio
      
      ing = ""
      Gramaje = 0
      
      NumLin = NumLin + 1
      
   End If
      

   
   DoEvents
           
   RsExcel.MoveNext
   
Loop
        
RsExcel.Close
Set RsExcel = Nothing
    
cn.Close
Set cn = Nothing

Close #1

Set XL = CreateObject("Excel.application")
XL.Visible = True
XL.Workbooks.Add
    
XL.Workbooks.OpenText vg_Archxls, , 1, 1, , , , , , , True, "|", Local:=True
    
'XL.quit
Set XL = Nothing
    
MsgBox "Proceso Finalizo Correctamente", vbInformation, MsgTitulo
    
Exit Sub
Man_Error:
    
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub
