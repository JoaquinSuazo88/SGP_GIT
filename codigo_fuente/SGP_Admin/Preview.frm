VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{1BCC7098-34C1-4749-B1A3-6C109878B38F}#1.0#0"; "vspdf8.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Preview 
   Caption         =   "Preview"
   ClientHeight    =   6960
   ClientLeft      =   1305
   ClientTop       =   2385
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6960
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   390
      Left            =   4770
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      _Version        =   393216
   End
   Begin VSPrinter8LibCtl.VSPrinter VSPrinter 
      Height          =   6735
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   7335
      _cx             =   1990341258
      _cy             =   1990340200
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
      PalettePicture  =   "Preview.frx":0000
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
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   35.204991087344
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
      PageBorder      =   0
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
      TextRTF         =   $"Preview.frx":001C
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   330
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   300
      _Version        =   393216
      _ExtentX        =   529
      _ExtentY        =   582
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
      SpreadDesigner  =   "Preview.frx":009F
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
      TextRTF         =   $"Preview.frx":0281
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
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

Me.Height = 7035
Me.Width = 9510
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Toolbar2.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.ToolTipText = "Exportar a Excel"
Set BtnX = Toolbar1.Buttons.Add(, "word", , tbrDefault, "word"): BtnX.ToolTipText = "Exportar a Word"
Set BtnX = Toolbar1.Buttons.Add(, "acrobat", , tbrDefault, "acrobat"): BtnX.ToolTipText = "Exportar a PDF"
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

Dim ArchivoExcel As String
Dim ArchivoPDF   As String

Select Case Button.Key

Case "word"
    
    If Len(VSPrinter.ExportFile) > 0 Then ShellExecute hwnd, "open", VSPrinter.ExportFile, 0, 0, 0

Case "excel"

    Dim WR As Object
    Dim XL As Object
    Dim fso As New FileSystemObject, fil1
    Me.Enabled = False
    If vg_opimp = 1 Or vg_opimp = 99999 Or vg_opimp = 999999 Or vg_opimp = 9999999 Then
        
        fg_carga ""
        Set XL = CreateObject("excel.application")
        Set hXl = New excel.Application
        XL.SheetsInNewWorkbook = 1
        XL.Workbooks.Add
        XL.Range("A1").Select
        vaSpread1.AllowMultiBlocks = True
        vaSpread1.SetSelection 1, -1, vaSpread1.maxcols, vaSpread1.MaxRows
        vaSpread1.ClipboardCopy
        '-------> formatear columna texto
        If vg_opimp = 1 Then
           
           XL.Range("A:A").Select
        
        ElseIf vg_opimp = 99999 Then
           
           XL.Range("B:B").Select
        
        ElseIf vg_opimp = 9999999 Then
           
           XL.Range("C:C").Select
        
        ElseIf vg_opimp = 999999 Then
           
           XL.Range("D:D").Select
        
        End If
        XL.Selection.NumberFormat = "@"
        
        XL.Range("A1").Select
        XL.ActiveSheet.Paste
    
        XL.Cells.Select
        XL.Cells.EntireColumn.AutoFit
        vaSpread1.AllowMultiBlocks = False: vaSpread1.SetSelection 1, 0, vaSpread1.maxcols, vaSpread1.MaxRows
        fg_descarga
        XL.Visible = True
    
    Else
        
        ArchivoExcel = fg_ArchivoExcel
        
        'Reviso Archivos a utilizar
        If Not fso.FileExists(VSPrinter.ExportFile) Then MsgBox "No se puede generar Excel, consulte con su administrador...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If fso.FileExists(ArchivoExcel) Then Set fil1 = fso.GetFile(ArchivoExcel): fil1.Delete
        Screen.MousePointer = 13
        DoEvents
    
        'Hago copia del archivo Rtf creado por la VSVIEW para trabajar con ella
        Set fil1 = fso.GetFile(VSPrinter.ExportFile)
        fil1.Copy (ArchivoExcel)
        Set WR = CreateObject("Word.Application")
 '       WR.applicationobject.OleRequestPendingTimeout = 15000
        WR.ChangeFileOpenDirectory App.Path & "\"
        WR.Documents.Open FileName:=ArchivoExcel, ConfirmConversions:=False, _
            ReadOnly:=True, AddToRecentFiles:=False, PasswordDocument:="", _
            PasswordTemplate:="", Revert:=True, WritePasswordDocument:="", _
            WritePasswordTemplate:="", Format:=wdOpenFormatAuto
        WR.Selection.WholeStory
        WR.Selection.Copy
        WR.ActiveDocument.Close
        
        Set XL = CreateObject("Excel.Application")
'        XL.applicationobject.OleRequestPendingTimeout = 15000
        XL.Workbooks.Add
    
    '    XL.Range("A:A").Select
    '    XL.Selection.NumberFormat = "@"
  
    '    XL.ActiveSheet.PageSetup.RightHeader = "&D" 'Encabezado Fecha
    '    XL.ActiveSheet.PageSetup.RightFooter = "&P" 'Pie NºPagina
        'Selecciono y copio del portapapeles a Excel
        DoEvents
        XL.Range("A2").Select
        XL.ActiveSheet.Paste
        
        XL.Cells.Select
        With XL.Selection
            
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        
        End With
        'Muevo los titulos una celda a la derecha
        XL.Range("A2:A3").Select
        XL.Selection.Cut
    
        XL.Range("B2:B3").Select
        XL.ActiveSheet.Paste
        XL.ActiveSheet.PageSetup.Orientation = IIf(VSPrinter.Orientation = orPortrait, xlPortrait, xlLandscape)
        
        '-------> Pegar Logo
        Est = False
        For Each c In XL.ActiveSheet.Shapes
            
            Est = True
        
        Next
        
        If Est Then
            
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
        Set fil1 = fso.GetFile(ArchivoExcel)
        fil1.Delete
        Set XL = Nothing
    
    End If
    Me.Enabled = True
    fg_descarga

Case "acrobat"
    
    With VSPDF
    
        .Title = "Reporte"
        .Creator = "AMJ"
        .Author = "AMJ"
        .Subject = "Reporte"
        .Keywords = "VSView8 VSPrinter8 ActiveX ComponentOne"
    
    End With
    
    ArchivoPDF = fg_ArchivoPDF
    
    VSPDF.ConvertDocument VSPrinter, ArchivoPDF
    ShellExecute hwnd, "open", ArchivoPDF, 0, 0, 1
    
End Select

Man_Error:
If Err.Number = -2147467259 Then MsgBox Err & ":  " & "No existe base de datos... " & Chr(13) & "Comunicase con departamento de informatica" & Chr(13) & "Actualización cancelada", vbExclamation + vbOKOnly, "Mantención sistema SGP": End
Resume Next

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "A_Salir    "
        
        Me.Hide
        Unload Me
    
End Select

End Sub

Private Sub VSPrinter_NewLine()
    
    vg_lineas = vg_lineas + 1

End Sub

Private Sub VSPrinter_NewPage()
    
    vg_lineas = 0

End Sub
