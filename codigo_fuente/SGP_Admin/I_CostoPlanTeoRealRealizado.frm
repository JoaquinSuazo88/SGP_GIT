VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form I_CostoPlanTeoRealRealizado 
   Caption         =   "Costo Plan. Teorico - Real - Realizado"
   ClientHeight    =   8205
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   13230
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   435
      Index           =   7
      Left            =   10320
      TabIndex        =   27
      Top             =   6840
      Width           =   2370
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   8
         Left            =   45
         TabIndex        =   13
         Top             =   135
         Width           =   2265
      End
   End
   Begin VB.Frame Frame3 
      Height          =   435
      Index           =   6
      Left            =   9360
      TabIndex        =   26
      Top             =   6840
      Width           =   930
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   7
         Left            =   45
         TabIndex        =   12
         Top             =   135
         Width           =   825
      End
   End
   Begin VB.Frame Frame3 
      Height          =   435
      Index           =   5
      Left            =   6960
      TabIndex        =   25
      Top             =   6840
      Width           =   2370
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   6
         Left            =   45
         TabIndex        =   11
         Top             =   135
         Width           =   2265
      End
   End
   Begin VB.Frame Frame3 
      Height          =   435
      Index           =   4
      Left            =   6000
      TabIndex        =   24
      Top             =   6840
      Width           =   930
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   5
         Left            =   45
         TabIndex        =   10
         Top             =   135
         Width           =   825
      End
   End
   Begin VB.Frame Frame3 
      Height          =   435
      Index           =   3
      Left            =   3600
      TabIndex        =   23
      Top             =   6840
      Width           =   2370
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   4
         Left            =   45
         TabIndex        =   9
         Top             =   135
         Width           =   2265
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7815
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   12975
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
         Height          =   540
         Left            =   9840
         TabIndex        =   14
         Top             =   7200
         Width           =   1275
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
         Left            =   11250
         TabIndex        =   15
         Top             =   7200
         Width           =   1275
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   0
         Left            =   2520
         TabIndex        =   20
         Top             =   6600
         Width           =   930
         Begin VB.Frame Frame3 
            Height          =   435
            Index           =   2
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   930
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   3
               Left            =   45
               TabIndex        =   8
               Top             =   135
               Width           =   825
            End
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   21
            Top             =   135
            Width           =   825
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   1
         Left            =   1560
         TabIndex        =   19
         Top             =   6600
         Width           =   930
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   7
            Top             =   135
            Width           =   825
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Costo Alimentación"
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
         Index           =   0
         Left            =   2280
         TabIndex        =   2
         Top             =   1200
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Costo Desechable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   4800
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Total Costo"
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
         Index           =   2
         Left            =   7320
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4695
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   12615
         _Version        =   393216
         _ExtentX        =   22251
         _ExtentY        =   8281
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
         MaxCols         =   9
         SpreadDesigner  =   "I_CostoPlanTeoRealRealizado.frx":0000
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   2355
         TabIndex        =   0
         Top             =   600
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
         TabIndex        =   1
         Top             =   600
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   600
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
               Picture         =   "I_CostoPlanTeoRealRealizado.frx":1A31
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   360
         Left            =   12000
         TabIndex        =   5
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cargar Información"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         Index           =   0
         Left            =   1080
         TabIndex        =   18
         Top             =   690
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
         Left            =   7035
         TabIndex        =   17
         Top             =   690
         Width           =   1065
      End
   End
End
Attribute VB_Name = "I_CostoPlanTeoRealRealizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public lc_Aux As String
Dim MsgTitulo As String

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim i               As Long
Dim isel            As Boolean
Dim RS              As New ADODB.Recordset
Dim Sql             As String
Dim Ceco            As String
Dim Org             As String
Dim Regimen         As Long
Dim Servicio        As Long
Dim MyBuffer        As String
Dim NomArchivoExcel As String
Dim seleccion       As String
Dim Extension       As String

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object
Dim XL              As New excel.Application 'Crea el objeto excel

If CDate(FpFecDesde.text) > CDate(FpFecHasta.text) Then
    
    Call MsgBox("Fecha Desde No Puede Ser Mayor a Fecha Hasta", vbInformation, Me.Caption)
    Let FpFecDesde.text = Format(Now, "dd/mm/yyyy")
    Call FpFecDesde.SetFocus
    Exit Sub

End If

If CDate(FpFecHasta.text) < CDate(FpFecDesde.text) Then
    
    Call MsgBox("Fecha Hasta No Puede Ser Mayor a Fecha Desde", vbInformation, Me.Caption)
    Let FpFecHasta.text = Format(Now, "dd/mm/yyyy")
    Call FpFecHasta.SetFocus
    Exit Sub

End If

isel = False

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 1
    
    If Trim(vaSpread1.text) = "1" Then
    
       isel = True
       Exit For
       
    End If

Next i

If Not isel Then

    Call MsgBox("Debe seleecionar a lo menos un item de la lista..", vbInformation, Me.Caption)
    Exit Sub

End If

'-------> Rescata Ceco Seleccionado
Command1.Enabled = False
seleccion = 0
fg_carga ""

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<Minuta>"

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 1 'Seleccion
    seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
  
    If seleccion = 1 And vaSpread1.RowHidden = False Then
        
        Org = ""
        vaSpread1.Col = 2
        Org = vaSpread1.text
        
        Ceco = ""
        vaSpread1.Col = 3
        Ceco = vaSpread1.text
        
        Regimen = 0
        vaSpread1.Col = 5
        Regimen = vaSpread1.text
        
        Servicio = 0
        vaSpread1.Col = 7
        Servicio = vaSpread1.text
        
        MyBuffer = MyBuffer & " <DetMinuta"
        MyBuffer = MyBuffer & " Org = " & Chr(34) & Org & Chr(34)
        MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
        MyBuffer = MyBuffer & " Reg = " & Chr(34) & Regimen & Chr(34)
        MyBuffer = MyBuffer & " Ser = " & Chr(34) & Servicio & Chr(34)
        MyBuffer = MyBuffer & "/>"
    
    End If

    DoEvents
     
Next i
Let MyBuffer = MyBuffer & "</Minuta>"

'-------> Guardar nombre archivo excel
NomArchivoExcel = ""
CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
CD.Filter = "Todos los archivos *.xls,*.xlsx"
On Error Resume Next
CD.ShowSave
           
'-------> JPAZ Permite controlar Boton Cancelar
If Err.Number = 32755 Then
   
   Command1.Enabled = True
   fg_descarga
   MsgBox "Proceso cancelado"
   Exit Sub

End If
           
If CD.FileName = "" Then
   
   fg_descarga
   Command1.Enabled = True
   MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
   Exit Sub

Else
   
   Extension = ""
   Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
   If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
      
      fg_descarga
      Command1.Enabled = True
      MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
      Exit Sub
   
   End If
   NomArchivoExcel = CD.FileName

End If

'-------> Validar cantidad registro se sobre pase hoja excel
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_XmlExportarExcelPlaTeoricoRealRealizado '" & MyBuffer & "', " & Format(FpFecDesde.text, "yyyymmdd") & ", " & Format(FpFecHasta.text, "yyyymmdd") & ", '" & IIf(Option2(0).Value = True, "1", IIf(Option2(1).Value = True, "2", "3")) & "'")
If Not RS.EOF Then
  
   If RS.RecordCount > 1020000 And UCase(Extension) = "XLSX" Then
      
      Command1.Enabled = True
        
      '-------> Close ADO objects
      RS.Close
      Set RS = Nothing
      fg_descarga
      MsgBox "El resultado sobrepasa maximo de fila en excel 1020000, proceso cancelado utilice filtro categoria dietetica o bien tipo de plato", vbCritical
      Exit Sub
   
   ElseIf UCase(Extension) = "XLS" And RS.RecordCount > 65533 Then
   
      Command1.Enabled = True
      
      '-------> Close ADO objects
      RS.Close
      Set RS = Nothing
      
      MsgBox "El resultado sobrepasa maximo de fila en excel 65533, proceso cancelado utilice filtro categoria dietetica o bien tipo de plato", vbCritical
      Exit Sub
   
   
   End If
  
End If

'-------> Create an instance of Excel and add a workbook
Set xlApp = CreateObject("Excel.Application")
Set xlWb = xlApp.Workbooks.Add
Set xlWs = xlWb.Worksheets("Hoja1")
  
'-------> Display Excel and give user control of Excel's lifetime
'    xlApp.Visible = True
xlApp.UserControl = True
    
'-------> Check version of Excel
Call encabezado(RS, xlWs)
          
xlWs.Cells(2, 1).CopyFromRecordset RS
'-------> Auto-fit the column widths and row heights
xlApp.Selection.CurrentRegion.Columns.AutoFit
xlApp.Selection.CurrentRegion.Rows.AutoFit
    
xlApp.Columns("S:S").Select
xlApp.Selection.Delete Shift:=xlToLeft

'xlApp.Range("O:O").Select
'xlApp.Range("O:O").Activate
'xlApp.Selection.NumberFormat = "0" '"#.0#"" per part"""
'
'xlApp.Columns("O:O").Select
'xlApp.Selection.Replace What:="""", Replacement:="", LookAt:=xlPart, _
'SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'ReplaceFormat:=False
'
'xlApp.Range("Q:Q").Select
'xlApp.Range("Q:Q").Activate
'xlApp.Selection.NumberFormat = "0.0000" '"#.0#"" per part"""

xlApp.Range("C:C").Select
xlApp.Range("C:C").Activate
xlApp.Selection.Replace What:="999999999", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

xlApp.Range("E:E").Select
xlApp.Range("E:E").Activate
xlApp.Selection.Replace What:="999999999", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

xlApp.Range("h:h").Select
xlApp.Range("h:h").Activate
xlApp.Selection.NumberFormat = "#,##0.00"
    
xlApp.Range("j:j").Select
xlApp.Range("j:j").Activate
xlApp.Selection.NumberFormat = "#,##0.00"


xlApp.Range("k:k").Select
xlApp.Range("k:k").Activate
xlApp.Selection.NumberFormat = "#,##0.00"

xlApp.Range("m:m").Select
xlApp.Range("m:m").Activate
xlApp.Selection.NumberFormat = "#,##0.00"

xlApp.Range("o:o").Select
xlApp.Range("o:o").Activate
xlApp.Selection.NumberFormat = "#,##0.00"

xlApp.Range("q:q").Select
xlApp.Range("q:q").Activate
xlApp.Selection.NumberFormat = "#,##0.00"
'xlApp.Range("O:O,Q:Q").Select
'xlApp.Range("O:O,Q:Q").Activate
'With xlApp.Selection
'        .HorizontalAlignment = xlRight
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'End With

xlWb.Close True, NomArchivoExcel

'Dim XL As New excel.Application 'Crea el objeto excel
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
                
Command1.Enabled = True

Exit Sub
Man_Error:
    
    
    Command1.Enabled = True
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Command2_Click()

On Error GoTo Man_Error

    Me.Hide
    Unload Me

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()
    
On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
fg_centra Me
MsgTitulo = "Costo Planificado Teorico - Real - Realizado"

FpFecDesde.text = Format(Date, "dd/mm/yyyy")
FpFecHasta.text = Format(Date, "dd/mm/yyyy")
vaSpread1.MaxRows = 0

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub

vaSpread1.MaxRows = 0

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

If IsDate(FpFecHasta.text) = False Then Exit Sub

vaSpread1.MaxRows = 0

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
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
   Text1(5).text = ""
   Text1(6).text = ""
   Text1(7).text = ""
   Text1(8).text = ""

ElseIf Index = 3 Then
   
   Text1(2).text = ""
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(6).text = ""
   Text1(7).text = ""
   Text1(8).text = ""

ElseIf Index = 4 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(5).text = ""
   Text1(6).text = ""
   Text1(7).text = ""
   Text1(8).text = ""

ElseIf Index = 5 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(6).text = ""
   Text1(7).text = ""
   Text1(8).text = ""

ElseIf Index = 6 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(7).text = ""
   Text1(8).text = ""

ElseIf Index = 7 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(6).text = ""
   Text1(5).text = ""
   Text1(8).text = ""

ElseIf Index = 8 Then
   
   Text1(2).text = ""
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(6).text = ""
   Text1(7).text = ""
   Text1(5).text = ""

End If

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 9
    vaSpread1.text = 0
    
Next

Select Case Index

Case 2, 3, 4, 5, 6, 7, 8
    
    vaSpread1.Visible = False
    
    If Trim(Text1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like IIf(Index = 2 Or Index = 3 Or Index = 5 Or Index = 7, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
           vaSpread1.Col = Index
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 9
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 9
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 9
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 9
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 9
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
           
           vaSpread1.Col = 9
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

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Local Error GoTo Error

Dim RS As New ADODB.Recordset
Dim Sql As String
    
    Select Case Button.Index
    
    Case 1 'Mostrar datos en la grilla
                   
            If CDate(FpFecDesde.text) > CDate(FpFecHasta.text) Then
                
                Call MsgBox("Fecha Desde No Puede Ser Mayor a Fecha Hasta", vbInformation, Me.Caption)
                Let FpFecDesde.text = Format(Now, "dd/mm/yyyy")
                Call FpFecDesde.SetFocus
                Exit Sub
            
            End If
    
            If CDate(FpFecHasta.text) < CDate(FpFecDesde.text) Then
                
                Call MsgBox("Fecha Hasta No Puede Ser Mayor a Fecha Desde", vbInformation, Me.Caption)
                Let FpFecHasta.text = Format(Now, "dd/mm/yyyy")
                Call FpFecHasta.SetFocus
                Exit Sub
            
            End If
        
            vaSpread1.Visible = False
            vaSpread1.MaxRows = 0
            Sql = ""
            Sql = Sql & "" & Format(FpFecDesde.text, "yyyymmdd")
            Sql = Sql & "," & Format(FpFecHasta.text, "yyyymmdd")
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute("sgpadm_Sel_ListarEncDatosSitio  " & Sql & "")
            Do While Not RS.EOF
               
               vaSpread1.MaxRows = vaSpread1.MaxRows + 1
               vaSpread1.Row = vaSpread1.MaxRows
               
               vaSpread1.Col = 1
               vaSpread1.text = "0"
               
               vaSpread1.Col = 2
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS(0)
               
               vaSpread1.Col = 3
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS(1)
               
               vaSpread1.Col = 4
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS(2)
               
               vaSpread1.Col = 5
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS(3)
               
               vaSpread1.Col = 6
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS(4)
               
               vaSpread1.Col = 7
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS(5)
               
               vaSpread1.Col = 8
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS(6)
               
               vaSpread1.Col = 9
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = 0
               
               RS.MoveNext
            
            Loop
            RS.Close
            Set RS = Nothing
            vaSpread1.Visible = True
        
    End Select

Exit Sub

Error:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Exit Sub

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1
    
    Dim i As Long
    vaSpread1.Col = 1
    
'    For i = BlockRow To BlockRow2
'
'        vaSpread1.Row = i
'
'        If vaSpread1.RowHidden = False Then
'
'           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
'
'        End If
'
'    Next
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    If BlockRow = -1 Then Exit Sub

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

End Sub
