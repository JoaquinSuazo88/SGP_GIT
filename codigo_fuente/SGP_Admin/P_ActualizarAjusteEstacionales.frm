VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form P_ActualizarAjusteEstacionales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Ajuste Estacionales Recetas"
   ClientHeight    =   10740
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10740
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   9135
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   14655
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
         Height          =   375
         Index           =   1
         Left            =   12960
         TabIndex        =   23
         Top             =   8640
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Actualizar"
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
         Left            =   11400
         TabIndex        =   22
         Top             =   8640
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   6
         Left            =   10440
         TabIndex        =   19
         Top             =   8040
         Width           =   900
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   7
            Left            =   45
            TabIndex        =   20
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   5
         Left            =   7440
         TabIndex        =   17
         Top             =   8040
         Width           =   2820
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   45
            TabIndex        =   18
            Top             =   135
            Width           =   2715
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   4
         Left            =   11400
         TabIndex        =   15
         Top             =   8040
         Width           =   2820
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   8
            Left            =   45
            TabIndex        =   16
            Top             =   135
            Width           =   2715
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   3
         Left            =   3600
         TabIndex        =   13
         Top             =   8040
         Width           =   2820
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   14
            Top             =   135
            Width           =   2715
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   2
         Left            =   6480
         TabIndex        =   11
         Top             =   8040
         Width           =   900
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   45
            TabIndex        =   12
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   0
         Left            =   2520
         TabIndex        =   9
         Top             =   8040
         Width           =   1020
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   10
            Top             =   135
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   8040
         Width           =   900
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   8
            Top             =   135
            Width           =   795
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   7695
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   14415
         _Version        =   393216
         _ExtentX        =   25426
         _ExtentY        =   13573
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
         SpreadDesigner  =   "P_ActualizarAjusteEstacionales.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   12615
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   1635
         TabIndex        =   1
         Top             =   360
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
         Left            =   8895
         TabIndex        =   2
         Top             =   360
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   12000
         TabIndex        =   21
         Top             =   360
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
         Left            =   360
         TabIndex        =   4
         Top             =   450
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
         Left            =   7635
         TabIndex        =   3
         Top             =   450
         Width           =   1065
      End
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
            Picture         =   "P_ActualizarAjusteEstacionales.frx":19CE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "P_ActualizarAjusteEstacionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public lc_Aux As String

Private obj_Excel     As Object
Private obj_Workbook  As Object
Private obj_Worksheet As Object

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim Sql             As String
Dim i               As Long
Dim seleccion       As String
Dim Ceco            As String
Dim Org             As String
Dim Regimen         As Long
Dim Servicio        As Long
Dim MyBuffer        As String
Dim NomArchivoExcel As String

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object
Dim XL              As New excel.Application 'Crea el objeto excel

Select Case Index

Case 0

  If ValidaDatos("2") = False Then Exit Sub

  '-------> Rescata Ceco Seleccionado
  seleccion = 0
  fg_carga ""
  
  Let MyBuffer = ""
  Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
  Let MyBuffer = MyBuffer & "<UpdateRecetaMinuta>"
  
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
          
          MyBuffer = MyBuffer & " <RecetasMinuta"
          MyBuffer = MyBuffer & " Org = " & Chr(34) & Org & Chr(34)
          MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
          MyBuffer = MyBuffer & " Reg = " & Chr(34) & Regimen & Chr(34)
          MyBuffer = MyBuffer & " Ser = " & Chr(34) & Servicio & Chr(34)
          MyBuffer = MyBuffer & "/>"
      
      
      End If
  
      DoEvents
       
  Next i

  Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), Me.HelpContextID, "", "", "")

  MyBuffer = MyBuffer & "</UpdateRecetaMinuta>"
      
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient

  Set RS = vg_db.Execute("sgpadm_Upd_XmlAjusteEstacionalMinutaBloque '" & MyBuffer & "', " & Format(FpFecDesde.text, "yyyymmdd") & ", " & Format(FpFecHasta.text, "yyyymmdd") & ", '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

  If Not RS.EOF Then
  
     If RS(0) > 0 Or RS(0) < 0 Then
          
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
  
       NomArchivoExcel = fg_ArchivoXls("ReporteError_ActualizacionAjusteEstacionales")
                    
       xlWb.Close True, NomArchivoExcel

'      Dim XL As New excel.Application 'Crea el objeto excel
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
          
        If RS(2) > 0 Then
        
           Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")

           MsgBox "Proceso finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
        
        Else
        
           Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_No_Encontraron_Datos_Actualizar"), Me.HelpContextID, "", "", "")
           
           MsgBox "Proceso finalizado, no se encontraron datos que actualizar en las fechas indicadas...", vbInformation + vbOKOnly, Me.Caption
           
        End If
                
     End If
  
  End If
  
  RS.Close
  Set RS = Nothing

  fg_descarga
    

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

On Error GoTo Man_Error

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_carga ""
Me.HelpContextID = vg_OpcM

MsgTitulo = "Actualizar Ajsute Estacionales Recetas"
fg_centra Me

vaSpread1.MaxRows = 0
Let FpFecDesde.text = Format(Date, "dd/mm/yyyy")
Let FpFecHasta.text = Format(Date, "dd/mm/yyyy")
'lblStatus.Visible = False
'prbStatus.Visible = False

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub

TextDet2(2).text = ""
TextDet2(3).text = ""
TextDet2(4).text = ""
TextDet2(5).text = ""
TextDet2(6).text = ""
TextDet2(7).text = ""
TextDet2(8).text = ""

vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &HC0FFFF

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

TextDet2(2).text = ""
TextDet2(3).text = ""
TextDet2(4).text = ""
TextDet2(5).text = ""
TextDet2(6).text = ""
TextDet2(7).text = ""
TextDet2(8).text = ""

vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &HC0FFFF

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


Private Sub TextDet2_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(TextDet2(Index).text, ",")

If Index = 2 Then
   
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""
   TextDet2(6).text = ""
   TextDet2(7).text = ""
   TextDet2(8).text = ""

ElseIf Index = 3 Then
   
   TextDet2(2).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""
   TextDet2(6).text = ""
   TextDet2(7).text = ""
   TextDet2(8).text = ""

ElseIf Index = 4 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(5).text = ""
   TextDet2(6).text = ""
   TextDet2(7).text = ""
   TextDet2(8).text = ""

ElseIf Index = 5 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(6).text = ""
   TextDet2(7).text = ""
   TextDet2(8).text = ""

ElseIf Index = 6 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""
   TextDet2(7).text = ""
   TextDet2(8).text = ""

ElseIf Index = 7 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(6).text = ""
   TextDet2(5).text = ""
   TextDet2(8).text = ""

ElseIf Index = 8 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(6).text = ""
   TextDet2(7).text = ""
   TextDet2(5).text = ""

End If

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 9
    vaSpread1.text = 0
    
Next

Select Case Index

Case 2, 3, 4, 5, 6, 7, 8
    
    vaSpread1.Visible = False
    
    If Trim(TextDet2(Index).text) <> "" Then
       
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
    
    If Trim(TextDet2(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 9
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextDet2(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS        As New ADODB.Recordset
Dim Sql       As String
Dim i         As Long
Dim xmlceco   As String
Dim seleccion As String
Dim codCeco   As String

TextDet2(2).text = ""
TextDet2(3).text = ""
TextDet2(4).text = ""
TextDet2(5).text = ""
TextDet2(6).text = ""
TextDet2(7).text = ""
TextDet2(8).text = ""
   
Select Case Button.Index
    
    Case 1
    
      If ValidaDatos("1") = False Then Exit Sub
          
      vaSpread1.MaxRows = 0
      vaSpread1.Row = -1: vaSpread1.Col = -1
      vaSpread1.BackColor = &HC0FFFF
       
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient

      Sql = ""
      Sql = Sql & " '" & Format(FpFecDesde, ("YYYYmmdd")) & "', "
      Sql = Sql & " '" & Format(FpFecHasta, ("YYYYmmdd")) & "' "
      
      Set RS = vg_db.Execute("sgpadm_Sel_MBloqueAjusteEstacionales  " & Sql & "")
      If Not RS.EOF Then
         
         Do While Not RS.EOF
          
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
          vaSpread1.Row = vaSpread1.MaxRows
          
          vaSpread1.Col = 1
          vaSpread1.text = "0"
          
          vaSpread1.Col = 2
          vaSpread1.text = RS!ID_ORGCOMPRA
          
          vaSpread1.Col = 3
          vaSpread1.text = RS!Ceco
          
          vaSpread1.Col = 4
          vaSpread1.text = Trim(RS!cli_nombre)
          
          vaSpread1.Col = 5
          vaSpread1.text = RS!Regimen
          
          vaSpread1.Col = 6
          vaSpread1.text = Trim(RS!reg_nombre)
          
          vaSpread1.Col = 7
          vaSpread1.text = Trim(RS!Servicio)
          
          vaSpread1.Col = 8
          vaSpread1.text = Trim(RS!ser_nombre)
          
          vaSpread1.Col = 9
          vaSpread1.text = ""
          
          RS.MoveNext
         Loop
         Command1(0).Enabled = True
      
      Else
         
         vaSpread1.MaxRows = 0
         Command1(0).Enabled = False
         MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo
      
      End If
      RS.Close
      Set RS = Nothing

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Function ValidaDatos(ByVal op As String) As Boolean

On Error GoTo Man_Error

Let ValidaDatos = True
Dim i As Long
Dim estseleccion As Boolean
  
'-------> Validar fechas
If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
   
   MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
   Call FpFecHasta.SetFocus
   Let ValidaDatos = False
   Exit Function

End If
    
If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
   
   MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
   Let FpFecDesde.text = Format(Now, "dd/mm/yyyy")
   Call FpFecDesde.SetFocus
   Let ValidaDatos = False
   Exit Function

End If

If op = 2 Then

   estseleccion = False
   For i = 1 To vaSpread1.MaxRows
   
        vaSpread1.Row = i
        vaSpread1.Col = 1
   
        If vaSpread1.text = "1" Then
   
            estseleccion = True
            Exit For
   
        End If
   
   Next i
   
   If Not estseleccion Then
   
      MsgBox "Debe seleccionar un item de la lista...", vbExclamation + vbOKOnly, MsgTitulo
      Let ValidaDatos = False
      Exit Function
   
   End If
   
End If

Exit Function
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Function

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
