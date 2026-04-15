VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{1BCC7098-34C1-4749-B1A3-6C109878B38F}#1.0#0"; "vspdf8.ocx"
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{1DF3AFED-47E0-4474-9900-954B8FD24E86}#1.0#0"; "ChilkatMail.dll"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Begin VB.Form Preview 
   Caption         =   "Preview"
   ClientHeight    =   7095
   ClientLeft      =   2460
   ClientTop       =   2370
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7095
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   1725
      Left            =   7650
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   1470
      _Version        =   393216
      _ExtentX        =   2593
      _ExtentY        =   3043
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
      MaxCols         =   0
      MaxRows         =   0
      SpreadDesigner  =   "Preview.frx":0000
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   390
      Left            =   5475
      TabIndex        =   6
      Top             =   0
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      _Version        =   393216
   End
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   330
      Left            =   8505
      ScaleHeight     =   270
      ScaleWidth      =   360
      TabIndex        =   4
      Top             =   180
      Visible         =   0   'False
      Width           =   420
   End
   Begin VSPrinter8LibCtl.VSPrinter VSPrinter 
      Height          =   6735
      Left            =   45
      TabIndex        =   2
      Top             =   120
      Width           =   7335
      _cx             =   12938
      _cy             =   11880
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      PalettePicture  =   "Preview.frx":01E2
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   72
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   37.405303030303
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   4
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393217
      BulletIndent    =   1
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Preview.frx":01FE
   End
   Begin RichTextLib.RichTextBox rtfPic 
      Height          =   825
      Left            =   7695
      TabIndex        =   5
      Top             =   810
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1455
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Preview.frx":0280
   End
   Begin ACTIVEZIPLib.ActiveZip AZ1 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin CHILKATMAILLibCtl.ChilkatMailMan oMail 
      Left            =   8160
      OleObjectBlob   =   "Preview.frx":0302
      Top             =   3000
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP oFTP 
      Left            =   0
      OleObjectBlob   =   "Preview.frx":0400
      Top             =   0
   End
   Begin VSPDF8LibCtl.VSPDF8 VSPDF 
      Left            =   8145
      Top             =   2340
      Author          =   ""
      Creator         =   ""
      Title           =   ""
      Subject         =   ""
      Keywords        =   ""
      Compress        =   3
   End
End
Attribute VB_Name = "Preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_Dato As Recordset
Dim RS_Dat2 As Recordset
Dim Tpag As Long
Dim ML As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()

Me.Height = 7035
Me.Width = 9510
fg_centra Me
vg_reporte = ""
vg_reporte = fg_ArchivoRtf
'vg_reporte = "\" & vg_NUsr & "\Reporte.rtf"
Toolbar1.ImageList = Partida.IL1
Toolbar2.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.ToolTipText = "Exportar a Excel":
Set BtnX = Toolbar1.Buttons.Add(, "word", , tbrDefault, "word"): BtnX.ToolTipText = "Exportar a Word":
Set BtnX = Toolbar1.Buttons.Add(, "acrobat", , tbrDefault, "acrobat"): BtnX.ToolTipText = "Exportar a PDF"
Set BtnX = Toolbar1.Buttons.Add(, "A_Grafico", , tbrDefault, "A_Grafico"): BtnX.ToolTipText = IIf(vg_opgra = 1, "Gráfico Costo Totales", "Gráfico Comparativo de Raciones"): BtnX.Visible = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Grafico1", , tbrDefault, "A_Grafico1"): BtnX.ToolTipText = "Gráfico Costo Bandeja": BtnX.Visible = False

Set btnX2 = Toolbar2.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX2.ToolTipText = "Salir"

VSPrinter.Orientation = orPortrait
VSPrinter.ZoomMode = zmPageWidth
VSPrinter.AbortCaption = "Imprimiendo..."
VSPrinter.AbortTextPage = "Imprimiendo Página %d de"
VSPrinter.AbortTextDevice = "%s en %s"
VSPrinter.AbortTextButton = "Cancelar"
VSPrinter.NavBarMenuText = "&Página Completa|&Ancho de Página|&Dos Páginas|&Multi-Página"
VSPrinter.Styles.Clear
VSPrinter.Styles.Add "Default", vpsAll

VSPrinter.Orientation = orPortrait
VSPrinter.ZoomMode = zmPageWidth
VSPrinter.AbortCaption = "Imprimiendo..."
VSPrinter.AbortTextPage = "Imprimiendo Página %d de"
VSPrinter.AbortTextDevice = "%s en %s"
VSPrinter.AbortTextButton = "Cancelar"
VSPrinter.NavBarMenuText = "&Página Completa|&Ancho de Página|&Dos Páginas|&Multi-Página"

    
' define title format (base on Default)
VSPrinter.FontName = "Arial"
VSPrinter.FontSize = 16
VSPrinter.FontBold = True
VSPrinter.IndentLeft = "0.5in"
VSPrinter.SpaceAfter = "14pt"
VSPrinter.Styles.Add "Titulo", vpsContent

VSPrinter.Styles.Apply "Default"
VSPrinter.FontName = "Arial"
VSPrinter.FontSize = 12
VSPrinter.FontBold = True
VSPrinter.IndentTab = 500
VSPrinter.IndentLeft = "0.5in"
VSPrinter.IndentLeft = VSPrinter.IndentLeft + VSPrinter.IndentTab
VSPrinter.IndentFirst = -VSPrinter.IndentTab
VSPrinter.SpaceBefore = "8pt"
VSPrinter.SpaceAfter = "6pt"
VSPrinter.Styles.Add "Capitulo", vpsContent

VSPrinter.Styles.Apply "Default"
VSPrinter.FontName = "Arial"
VSPrinter.FontSize = 11
VSPrinter.FontBold = True
VSPrinter.IndentTab = 500
VSPrinter.IndentLeft = "0.5in"
VSPrinter.IndentLeft = VSPrinter.IndentLeft + VSPrinter.IndentTab
VSPrinter.IndentFirst = -VSPrinter.IndentTab
VSPrinter.SpaceBefore = "8pt"
VSPrinter.SpaceAfter = "6pt"
VSPrinter.Styles.Add "Clausula", vpsContent

' define code format (based on Normal)
VSPrinter.Styles.Apply "Default"
VSPrinter.FontName = "Arial"
VSPrinter.FontSize = 10
VSPrinter.IndentTab = 500
VSPrinter.IndentLeft = "0.5in"
VSPrinter.IndentLeft = VSPrinter.IndentLeft + VSPrinter.IndentTab
VSPrinter.IndentFirst = -VSPrinter.IndentTab
VSPrinter.SpaceAfter = "2pt"
VSPrinter.Styles.Add "Normal", vpsContent

' restore default
VSPrinter.Styles.Apply "Default"
MsgTitulo = "Reportes..."
End Sub

Private Sub Form_Resize()
VSPrinter.Height = IIf(Preview.Height - 550 < 0, VSPrinter.Height, Preview.Height - 550)
VSPrinter.Width = IIf(Preview.Width - 220 < 0, Preview.Width, Preview.Width - 220)
End Sub

Private Sub Form_Unload(Cancel As Integer)
VSPrinter.ExportFile = ""
End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Dim codser As String, codreg As String, c As Variant, est As Boolean
Select Case Button.Key
Case "word"
    If Len(VSPrinter.ExportFile) > 0 Then ShellExecute hWnd, "open", VSPrinter.ExportFile, 0, 0, 0
Case "excel"
'    Set XL = CreateObject("Excel.application")
'    XL.Visible = True
'    XL.Workbooks.OpenText vg_Archxls, , 1, 1, , , , , , , True, "|"
    Dim WR As Object
    Dim XL As Object
    Dim fso, fil1
    Set fso = CreateObject("Scripting.FileSystemObject")
    '------- Reviso Archivos a utilizar
'    If Not fso.FileExists(App.Path & vg_reporte) Then MsgBox "No se puede generar Excel, consulte con su administrador...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
'    If fso.FileExists(App.Path & "\ReporteExel.rtf") Then Set fil1 = fso.GetFile(App.Path & "\ReporteExel.rtf"): fil1.Delete
    If Not fso.FileExists(vg_reporte) Then MsgBox "No se puede generar Excel, consulte con su administrador...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If fso.FileExists(dir_trabajo_Inf & "ReporteExel.rtf") Then
       
       Set fil1 = fso.GetFile(dir_trabajo_Inf & "ReporteExel.rtf")
       fil1.Delete
    
    End If
    Screen.MousePointer = 13
    DoEvents
    Me.Enabled = False
    'Hago copia del archivo Rtf creado por la VSVIEW para trabajar con ella
'    Set fil1 = fso.GetFile(App.Path & vg_reporte)
    Set fil1 = fso.GetFile(vg_reporte)
    fil1.Copy (dir_trabajo_Inf & "ReporteExel.rtf")
    Set WR = CreateObject("Word.Application")
    WR.ChangeFileOpenDirectory dir_trabajo_Inf
'    WR.Documents.Open Filename:="ReporteExel.rtf", ConfirmConversions:=False, _
        ReadOnly:=True, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=True, WritePasswordDocument:="", _
        WritePasswordTemplate:="", Format:=wdOpenFormatAuto
    WR.Documents.Open FileName:="ReporteExel.rtf"
    WR.Selection.WholeStory
    WR.Selection.Copy
    WR.ActiveDocument.Close
    Set XL = CreateObject("Excel.Application")
    XL.Workbooks.Add
'    XL.ActiveSheet.PageSetup.RightHeader = "&D" 'Encabezado Fecha
'    XL.ActiveSheet.PageSetup.RightFooter = "&P" 'Pie NºPagina
    'Selecciono y copio del portapapeles a Excel
    XL.Range("A2").Select
    XL.ActiveSheet.Paste
    'Muevo los titulos una celda a la derecha
    XL.Range("A2:A3").Select
    XL.Selection.Cut
    XL.Range("B2:B3").Select
    XL.ActiveSheet.Paste
    XL.ActiveSheet.PageSetup.Orientation = IIf(VSPrinter.Orientation = orPortrait, xlPortrait, xlLandscape)
    
    'Pegar Logo
    est = False
    For Each c In XL.ActiveSheet.Shapes
        est = True
    Next
    If est Then
        XL.ActiveSheet.Shapes("Picture 1").Select
        XL.Selection.ShapeRange.Left = 0
        XL.Selection.ShapeRange.Top = 0
        XL.Selection.ShapeRange.Width = 1.5
        XL.Selection.ShapeRange.Height = 28
    End If
    
    XL.Cells.Select 'Selecciono todo
    XL.Selection.WrapText = False 'El texto puede sobrepasar el ancho de la celda
    XL.Selection.RowHeight = 12.75 'Alto de fila
    XL.Cells.EntireColumn.AutoFit 'Auto ajuste de columna
    XL.Rows("2:2").RowHeight = 18 'Alto de fila 2
    XL.Range("B2").Select
    XL.Selection.HorizontalAlignment = xlLeft
    
    XL.Visible = True
    WR.Application.Quit
    Set fil1 = fso.GetFile(dir_trabajo_Inf & "ReporteExcel.rtf")
'    fil1.Delete
    Me.Enabled = True
    fg_descarga
Case "acrobat"
    With VSPDF
        .Title = Mid(vg_reporte, 1, Len(vg_reporte) - 3)
        .Creator = "AMJ"
        .Author = "AMJ"
        .Subject = Mid(vg_reporte, 1, Len(vg_reporte) - 3)
        .Keywords = "VSView8 VSPrinter8 ActiveX ComponentOne"
    End With
    VSPDF.ConvertDocument VSPrinter, vg_dir & Mid(vg_reporte, 1, Len(vg_reporte) - 3) & "PDF" 'jpaz "\" REPORTE.PDF"
    ShellExecute hWnd, "open", vg_dir & Mid(vg_reporte, 1, Len(vg_reporte) - 3) & "PDF", 0, 0, 1 '"\REPORTE.PDF", 0, 0, 1
Case "A_Grafico", "A_Grafico1"
    Toolbar1.Enabled = False
    codreg = "": codser = ""
    Select Case vg_opgra
    Case 1
'        For i = 1 To I_CoteRe.vaSpread1(0).MaxRows
'            I_CoteRe.vaSpread1(0).Row = i: CoteRe.vaSpread1(0).Col = 1
'            If CoteRe.vaSpread1(0).text = "1" Then CoteRe.vaSpread1(0).Col = 2: codreg = codreg & "" & CoteRe.vaSpread1(0).text & ","
'        Next i
'        For i = 1 To CoteRe.vaSpread1(1).MaxRows
'            CoteRe.vaSpread1(1).Row = i: CoteRe.vaSpread1(1).Col = 1
'            If CoteRe.vaSpread1(1).text = "1" Then CoteRe.vaSpread1(1).Col = 2: codser = codser & "" & CoteRe.vaSpread1(1).text & ","
'        Next i
        G_TeoRea.LlenarGrafico vg_codcasino, vg_codreg, vg_codser, vg_fecini, vg_fecfin, Str(vg_op1), Str(vg_op2), vg_op3, 1, IIf(Trim(Button.Key) = "A_Grafico" Or vg_codser = "", False, True)
        G_TeoRea.Show 1
    Case 2
        G_ConRac.LlenarGrafico MuestraCasino(1), vg_codregimen, vg_codservicio, vg_fecini, vg_fecfin
        G_ConRac.Show 1
    End Select
    Toolbar1.Enabled = True
Case "A_Salir    "
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
fg_descarga
If Err = 53 Then Resume Next
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "A_Salir    "
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub VSPrinter_AfterFooter()
VSPrinter.MarginLeft = ML
End Sub

Private Sub VSPrinter_BeforeFooter()
ML = VSPrinter.MarginLeft
VSPrinter.MarginLeft = 500
End Sub

