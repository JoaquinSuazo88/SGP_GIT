VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form M_SsllPrecioRef 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precio Referencia"
   ClientHeight    =   8745
   ClientLeft      =   4155
   ClientTop       =   2070
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      ForeColor       =   &H80000000&
      Height          =   1575
      Left            =   1080
      TabIndex        =   12
      Top             =   600
      Width           =   8415
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7800
         Top             =   960
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
               Picture         =   "M_SsllPrecioRef.frx":0000
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   3000
         TabIndex        =   2
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   1080
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ButtonStyle     =   3
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
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "05/08/2011"
         DateCalcMethod  =   0
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483633
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   0
         Top             =   240
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1125
         Width           =   540
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3480
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3495
         TabIndex        =   17
         Top             =   290
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3000
         Picture         =   "M_SsllPrecioRef.frx":039A
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
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
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Left            =   3480
         TabIndex        =   15
         Top             =   720
         Width           =   2835
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5280
         TabIndex        =   14
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Left            =   3480
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   2835
      End
   End
   Begin VB.Frame Frame6 
      Height          =   6255
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   10215
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5610
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9975
         _Version        =   393216
         _ExtentX        =   17595
         _ExtentY        =   9895
         _StockProps     =   64
         BackColorStyle  =   3
         EditEnterAction =   5
         EditModeReplace =   -1  'True
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
         SpreadDesigner  =   "M_SsllPrecioRef.frx":06A4
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   5610
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   9975
         _Version        =   393216
         _ExtentX        =   17595
         _ExtentY        =   9895
         _StockProps     =   64
         BackColorStyle  =   3
         EditEnterAction =   5
         EditModeReplace =   -1  'True
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
         SpreadDesigner  =   "M_SsllPrecioRef.frx":2123
      End
      Begin FPSpread.vaSpread vaSpread3 
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   5520
         Visible         =   0   'False
         Width           =   855
         _Version        =   393216
         _ExtentX        =   1508
         _ExtentY        =   450
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
         MaxRows         =   0
         SpreadDesigner  =   "M_SsllPrecioRef.frx":3B6A
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   1920
         Top             =   6030
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Producto Excel Erroneo"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   11
         Top             =   6000
         Width           =   1680
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Producto Válido"
         Height          =   195
         Index           =   0
         Left            =   5040
         TabIndex        =   10
         Top             =   6000
         Width           =   1125
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   4680
         Top             =   6030
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   6840
         Top             =   6030
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Espacio Vacío"
         Height          =   195
         Index           =   2
         Left            =   7200
         TabIndex        =   9
         Top             =   6000
         Width           =   1050
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog Cd 
      Left            =   0
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "M_SsllPrecioRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim modo As String, codigo As String, Msgtitulo As String
Dim est As Boolean
Dim strSQL, CambRut As String
Dim OpGr As Boolean
Dim TipoOp As String
Dim TotalRegistros As Integer
Dim RutaArchivo As String
Dim CuentaError As Integer

Private Sub Form_Load()
    Dim itop, i As Integer
    Dim lc_Aux As String
    Dim BtnX As Object, btnX1 As Object
    
    Me.HelpContextID = vg_OpcM
    Me.Height = 9200
    Me.Width = 10830
    fg_centra Me
    Msgtitulo = "Precio de Referencia"
        
    modo = ""
    est = True
    
    TotalRegistros = 0
    itop = 1
   
    Gl_Mo_Botones Me, 16
    
    If (fpayuda(0).Caption = "") Then
        Toolbar1.Buttons(1).Enabled = False
    Else
        Toolbar1.Buttons(1).Enabled = True
    End If
    
    Toolbar1.Buttons(2).Enabled = False: Toolbar1.Buttons(6).Enabled = False
    Toolbar1.Buttons(8).Enabled = False: Toolbar1.Buttons(10).Enabled = False
    Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(7).Visible = False
    
    vaSpread2.Visible = False
    vaSpread1.Lock = True
    vaSpread1.MaxRows = 0
    fpDateTime1.text = Format(Date, "dd/mm/yyyy")
    fpDateTime1.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
TipoOp = ""
End Sub

Private Sub fpDateTime1_Change()
Dim i As Long
If IsDate(fpDateTime1.text) = False Then Exit Sub
If vaSpread1.MaxRows > 0 Then
   For i = 1 To vaSpread1.MaxRows
       vaSpread1.Row = i
       vaSpread1.Col = 3
       vaSpread1.text = Format(fpDateTime1.text, "dd/mm/yyyy")
   Next i
End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        If Not (fpText(0).text = "") Then
            MoverDatosGrilla
        End If
    End Select
    
    vaSpread2.Enabled = True
    If (vaSpread1.MaxRows > 0) Then
        fpDateTime1.Enabled = False
    Else
        fpDateTime1.Enabled = True
    End If
    
    TotalRegistros = vaSpread2.MaxRows
    TipoOp = ""
    
End Sub

Private Sub Image1_Click(Index As Integer)
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Mantenedor Centro Costo", "CentCost"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(0).text = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
End Sub

Private Sub fpText_Change(Index As Integer)
    fpDateTime1.Enabled = False
    If est Then Exit Sub
    fpayuda(0).Caption = ""
    If modo = "" Then modo = "M"
End Sub

Private Sub fpText_LostFocus(Index As Integer)
    Dim codi As Long, Bd As String, Ul As String
    On Error GoTo Man_Error
    If fpText(Index).text = "" Then fpayuda(0).Caption = "": codi = 0: Exit Sub
    
    codi = fpText(Index).text
    Bd = IIf(Index = 0, "b_clientes", "")
    Ul = IIf(Bd = "b_clientes", "cli", "")
    
    Set RS1 = Nothing
    
    strSQL = "SELECT " & Ul & "_codigo, " & Ul & "_nombre FROM " & Bd & " WHERE " & Ul & "_codigo=" & IIf(Ul = "cli", "'" & codi & "'", codi) & ""
    RS1.Open strSQL, vg_db, adOpenStatic
    
    If Not RS1.EOF Then
        fpayuda(0).Caption = IIf(IsNull(Trim(RS1!cli_nombre) = ""), "", RS1!cli_nombre)
        vg_codigo = RS1!cli_codigo
        codi = 0
    Else
        MsgBox "No existe codigo en la tabla..."
        fpayuda(0).Caption = ""
        fpText(Index).text = ""
        codi = 0
        On Error Resume Next: fpText(Index).SetFocus
        vaSpread1.DeleteRows 1, vaSpread1.MaxRows
        Label2.Caption = ""
        vaSpread1.MaxRows = 0
    End If
    
    RS1.Close: Set RS1 = Nothing
    Exit Sub
    
Man_Error:
    If Err = 3034 Then Exit Sub
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error

Select Case Button.Index

Case 1 'CARGAR ARCHIVO EXCEL

    Dim List() As String
    Dim ListCount As Integer
    Dim fromRight As Long, i As Long
    Dim varManejo As Integer
    Dim varRuta As String
    Dim f As Boolean
    Dim NombreArchivoExcel As String
    Dim dbexcel As Database, cSpi As Long
    Dim ExcelCodProd, ExcelFecha, ExcelPrecio As String
    Dim j As Integer
    Dim wvarHoja As String
    Dim wvarCol1, wvarCol2, wvarCol3, wvarCol4 As String
    Dim GrillaCodProd, GrillaDescrip, GrillaFecha, GrillaPrecio As String
    ReDim List(1)

    CD.DialogTitle = "Seleccionar Un Archivo XLS"
    CD.Filter = "Todos los archivos|*.*|Archivos de texto (*.xls)|*.xls"
    CD.FilterIndex = 2
    CD.Flags = cdlOFNFileMustExist
    CD.ShowOpen
    NombreArchivoExcel = CD.FileName
        
    If Len(NombreArchivoExcel) = 0 Then Exit Sub
    
    Label4.Caption = "Cargando Información ..."
    Label4.Visible = True
    
    fromRight = InStrRev(CD.FileName, "\", , vbTextCompare)
    
    If fromRight > 1 Then
       varRuta = Left(CD.FileName, fromRight)
    End If
    
    vaSpread3.MaxRows = 0: vaSpread3.MaxRows = 500
    vaSpread3.maxcols = 0: vaSpread3.maxcols = 500
    
    f = vaSpread3.GetExcelSheetList(CD.FileName, List, ListCount, (varRuta & "log.txt"), varManejo, True)
    
    If (ListCount - 1 > 1) Then
       ReDim List(ListCount - 1)
       f = vaSpread3.GetExcelSheetList(CD.FileName, List, ListCount, (varRuta & "log.txt"), varManejo, False)
    End If
        
    wvarHoja = (List(0))
 
    Dim RsExcel As ADODB.Recordset
    Dim sconn As String
    Set RsExcel = New ADODB.Recordset
    
    RsExcel.CursorLocation = adUseClient
    RsExcel.CursorType = adOpenKeyset
    RsExcel.LockType = adLockBatchOptimistic
    
    sconn = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & NombreArchivoExcel

    strSQL = "SELECT * FROM [" & wvarHoja & "$]"
    'strSQL = "SELECT * FROM [Hoja1$]"
        
    RsExcel.Open strSQL, sconn
    
    CuentaError = 0
    
    vaSpread2.MaxRows = Val(RsExcel.RecordCount)

    j = 1
    
    Do While Not RsExcel.EOF
        
        If ((IsNull(RsExcel(0)) And IsNull(RsExcel(1))) Or (Trim(RsExcel(0)) = "" And Trim(RsExcel(1)) = "")) Then
            j = j + 1
            RsExcel.MoveNext
            'Exit Do
        End If

        If RsExcel.EOF Then
            Exit Do
        End If
        
        vaSpread2.Row = j
        
        ExcelCodProd = IIf(IsNull(Trim(RsExcel(0))), "", Trim(RsExcel(0)))
        'ExcelFecha = IIf(IsNull(RsExcel!pr_Fecha), "", RsExcel!pr_Fecha)
        ExcelPrecio = IIf(IsNull(Trim(RsExcel(1))), "", Trim(RsExcel(1)))
        
        strSQL = "SELECT * FROM b_formatocompras WHERE foc_codsac = '" & ExcelCodProd & "'"
        Set RS = vg_db.Execute(strSQL)
        
        If (ExcelCodProd = "") Then
            vaSpread2.Col = 1: vaSpread2.text = ExcelCodProd: vaSpread2.BackColor = vbRed
            vaSpread2.Col = 2: vaSpread2.text = "": vaSpread2.BackColor = vbRed
            vaSpread2.Col = 3: vaSpread2.text = "" 'Format(Date, "dd/mm/yyyy"): vaSpread2.BackColor = vbRed 'IIf(IsDate(ExcelFecha), ExcelFecha, ""): vaSpread2.BackColor = vbRed
            vaSpread2.Col = 4: vaSpread2.text = ExcelPrecio: vaSpread2.BackColor = vbRed
            'CuentaError = CuentaError + 1
        ElseIf (IsNull(RS)) Then
            vaSpread2.Col = 1: vaSpread2.text = ExcelCodProd: vaSpread2.BackColor = &H8080FF
            vaSpread2.Col = 2: vaSpread2.text = "": vaSpread2.BackColor = &H8080FF
            vaSpread2.Col = 3: vaSpread2.text = "" 'Format(Date, "dd/mm/yyyy"): vaSpread2.BackColor = &H8080FF 'IIf(IsDate(ExcelFecha), ExcelFecha, ""): vaSpread2.BackColor = &H8080FF
            vaSpread2.Col = 4: vaSpread2.text = ExcelPrecio: vaSpread2.BackColor = &H8080FF
            'CuentaError = CuentaError + 1
        ElseIf (RS.EOF) Then
            vaSpread2.Col = 1: vaSpread2.text = ExcelCodProd: vaSpread2.BackColor = &H8080FF
            vaSpread2.Col = 2: vaSpread2.text = "": vaSpread2.BackColor = &H8080FF
            vaSpread2.Col = 3: vaSpread2.text = "" 'Format(Date, "dd/mm/yyyy"): vaSpread2.BackColor = &H8080FF 'IIf(IsDate(ExcelFecha), ExcelFecha, ""): vaSpread2.BackColor = &H8080FF
            vaSpread2.Col = 4: vaSpread2.text = ExcelPrecio: vaSpread2.BackColor = &H8080FF
            'CuentaError = CuentaError + 1
        Else
            vaSpread2.Col = 1: vaSpread2.text = RS!foc_codsac
            vaSpread2.Col = 2: vaSpread2.text = RS!foc_nomsac
            vaSpread2.Col = 3: vaSpread2.text = Format(fpDateTime1.text, "dd/mm/yyyy") 'Format(Date, "dd/mm/yyyy")
            vaSpread2.Col = 4: vaSpread2.text = IIf(IsNumeric(ExcelPrecio), ExcelPrecio, "")
        End If
        
        j = j + 1
        RsExcel.MoveNext
    Loop
        
    'If (CuentaError > 0) Then
        For i = 1 To vaSpread2.MaxRows
            vaSpread1.MaxRows = RsExcel.RecordCount
            vaSpread2.Row = i
            vaSpread1.Row = i
            
            vaSpread2.Col = 1: vaSpread1.Col = 1: wvarCol1 = Trim(vaSpread2.text)
            vaSpread2.Col = 2: vaSpread1.Col = 2: wvarCol2 = Trim(vaSpread2.text)
            vaSpread2.Col = 3: vaSpread1.Col = 3: wvarCol3 = Trim(vaSpread2.text)
            vaSpread2.Col = 4: vaSpread1.Col = 4: wvarCol4 = Trim(vaSpread2.text)
                        
            If (wvarCol1 = "" And wvarCol2 = "" And wvarCol4 = "") Then
            
                For j = 1 To 4
                    vaSpread1.Col = j: vaSpread2.Col = j
                    vaSpread1.BackColor = &H8000000E
                    vaSpread1.text = vaSpread2.text
                Next j
            
                vaSpread1.Col = 2
                vaSpread1.text = "Espacio Vacío"
                'CuentaError = CuentaError + 1
            
            ElseIf (wvarCol1 = "") Then
                For j = 1 To 4
                    vaSpread1.Col = j: vaSpread2.Col = j
                    vaSpread1.BackColor = &HC0C0FF
                    vaSpread1.text = vaSpread2.text
                Next j
                
                vaSpread1.Col = 2
                vaSpread1.text = "No Tiene Código de Producto."
                CuentaError = CuentaError + 1
            
            ElseIf (wvarCol2 = "") Then
                For j = 1 To 4
                    vaSpread1.Col = j: vaSpread2.Col = j
                    vaSpread1.BackColor = &HC0C0FF
                    vaSpread1.text = vaSpread2.text
                Next j
                
                vaSpread1.Col = 2
                vaSpread1.text = "El Producto No Fue Encontrado."
                CuentaError = CuentaError + 1
                
            ElseIf (wvarCol4 = "") Then
                For j = 1 To 4
                    vaSpread1.Col = j: vaSpread2.Col = j
                    vaSpread1.BackColor = &HC0C0FF
                    vaSpread1.text = vaSpread2.text
                Next j
                
                vaSpread1.Col = 2
                vaSpread1.text = "El Precio No Es Un Valor Numérico."
                CuentaError = CuentaError + 1
            
            Else
                For j = 1 To 4
                    vaSpread1.Col = j: vaSpread2.Col = j
                    vaSpread1.text = vaSpread2.text
                Next j
            End If
        Next i
    'End If
    

    RsExcel.Close: Set RsExcel = Nothing
    RS.Close: Set RS = Nothing
    
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(5).Visible = True
    Toolbar1.Buttons(6).Visible = False
    
    If CuentaError = 0 Then
        Toolbar1.Buttons(7).Visible = True
        Toolbar1.Buttons(8).Visible = False
        Label3.Caption = ""
    Else
        Toolbar1.Buttons(7).Visible = False
        Toolbar1.Buttons(8).Visible = True
        Label3.Caption = "Verifique el Archivo Excel Cargado"
    End If
    
    
    Label2.Caption = ""
    Label4.Caption = ""
    Label4.Visible = False
    Toolbar1.Buttons(10).Enabled = False
    fpText(0).Enabled = False
    Image1(0).Enabled = False
    Toolbar2.Enabled = False
    
Case 2 'BORRAR
        
    If MsgBox("¿Seguro Que Desea Borrar Los Datos?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    modo = "Cancel"
    
    If modo = "Cancel" Then
        
        strSQL = "DELETE b_ssll_precioref WHERE prr_codcen = '" & fpText(0).text & "'"
        vg_db.Execute strSQL
    
        modo = ""
        Cancela
        
        fpDateTime1.Enabled = True
        fpDateTime1.text = Format(Date, "dd/mm/yyyy")
    Else
        Cancela
    End If
    
    
Case 5 'CANCELAR-ACTIVO
    
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    modo = "Cancel"
    
    If modo = "Cancel" Then
        modo = ""
        CD.FileName = ""
        Label3.Caption = ""
        wvarCol1 = "": wvarCol2 = "": wvarCol3 = "": wvarCol4 = ""
        Cancela
    Else
        CD.FileName = ""
        Label3.Caption = ""
        wvarCol1 = "": wvarCol2 = "": wvarCol3 = "": wvarCol4 = ""
        Cancela
    End If
    
    

Case 6 'CANCELAR

Case 7 'GUARDAR-CONFIRMAR-ACTIVO
    
    If (CuentaError > 0) Then MsgBox "No Puede Guardar El Archivo Cargado, Solucione Los Errores.": Exit Sub

    'If MsgBox("¿Desea Guardar La Información Correcta De La Grilla?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub


    For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
        
        For j = 1 To vaSpread1.maxcols
            vaSpread1.Col = j
            
            If (vaSpread1.text <> "") Then
                If (j = 1) Then
                    GrillaCodProd = Trim(vaSpread1.text)
                ElseIf (j = 2) Then
                    GrillaDescrip = Trim(vaSpread1.text)
                ElseIf (j = 3) Then
                    If Not IsDate(vaSpread1.text) Then
                        GrillaCodProd = ""
                        GrillaDescrip = ""
                        GrillaFecha = ""
                        GrillaPrecio = ""
                        Exit For
                    Else
                        GrillaFecha = Format(fpDateTime1.text, "yyyy/mm/dd")
                    End If
                    
                ElseIf (j = 4) Then
                    GrillaPrecio = CDbl(Trim(fg_Quitachar(vaSpread1.text, "$")))
                End If
            Else
                GrillaCodProd = ""
                GrillaDescrip = ""
                GrillaFecha = ""
                GrillaPrecio = ""
                Exit For
            End If
        Next j
                        
                        
        If (fpText(0).text <> "" And GrillaCodProd <> "" And GrillaDescrip <> "" And GrillaFecha <> "" And GrillaPrecio <> "") Then
            strSQL = "INSERT INTO b_ssll_precioref(prr_codcen, prr_codfmc, prr_feccar, prr_precio) " & _
                     "VALUES('" & fpText(0).text & "','" & GrillaCodProd & "','" & GrillaFecha & "'," & _
                     "" & GrillaPrecio & ")"
            vg_db.Execute strSQL
        End If
    Next i
    
    Cancela
    MsgBox "Los Datos Han Sido Guardados"
    
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(8).Visible = True
    fpText(0).Enabled = True
    Image1(0).Enabled = True
    Toolbar2.Enabled = True
    CD.FileName = ""
    Label3.Caption = ""

Case 8 'GUARDAR-CONFIRMAR

Case 10 'IMPRIMIR
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_SsllPrecRef
    
Case 13 'SALIR
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err.Number = 3034 Then Exit Sub
If Err.Number = 3265 Then MsgBox "Verifique el Nombre de los Campos del Archivo Excel, Según el Formato", vbCritical, "Error"
If Err.Number = -2147467259 Then MsgBox "Error al Cargar Archivo Excel, Verifique Formato...", vbCritical, "Error" 'Else MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"

End Sub

Sub MoverDatosGrilla()
    fg_carga ""
    Dim x As Boolean

    vaSpread1.TextTip = 2
    vaSpread1.TextTipDelay = 250
    x = vaSpread2.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
    vaSpread1.Visible = False
    vaSpread1.MaxRows = 0
    vaSpread1.Row = -1
    vaSpread1.Col = -1
    
    If (fpText(0).text = "") Then
        strSQL = "SELECT prr_codcen, prr_codfmc, prr_feccar, prr_precio, foc_nomsac " & _
                 "FROM b_formatocompras, b_ssll_precioref, b_clientes " & _
                 "WHERE cli_tipo = 0 AND cli_activo = 1 " & _
                 "foc_codsac = prr_codfmc AND cli_codigo = prr_codcen " & _
                 "ORDER BY prr_codcen"
    Else
        strSQL = "SELECT prr_codcen, prr_codfmc, prr_feccar, prr_precio, foc_nomsac " & _
                 "FROM b_formatocompras, b_ssll_precioref, b_clientes " & _
                 "WHERE cli_tipo = 0 AND cli_activo = 1 " & _
                 "AND foc_codsac = prr_codfmc AND cli_codigo = prr_codcen " & _
                 "AND cli_codigo = '" & fpText(0).text & "' " & _
                 "ORDER BY prr_codcen"
    End If

    Set RS = vg_db.Execute(strSQL)
    
    If RS.EOF Then
        
        vaSpread1.Visible = True
        
        Toolbar1.Buttons(1).Enabled = True
        Toolbar1.Buttons(2).Enabled = False
        fpDateTime1.Enabled = True
    Else
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(2).Enabled = True
        fpDateTime1.Enabled = False
        fpDateTime1.text = Format(RS!prr_feccar, "dd/mm/yyyy")
        Do While Not RS.EOF
            vaSpread1.MaxRows = vaSpread1.MaxRows + 1
            vaSpread1.Row = vaSpread1.MaxRows
            
            vaSpread1.Col = 1
            vaSpread1.text = Trim(RS!prr_codfmc)
           
            vaSpread1.Col = 2
            vaSpread1.text = Trim(RS!foc_nomsac)
            
            vaSpread1.Col = 3
            vaSpread1.text = Trim(RS!prr_feccar)
            
            vaSpread1.Col = 4
            vaSpread1.text = Trim(RS!prr_precio)
                    
            RS.MoveNext
        Loop
        
        RS.Close: Set RS = Nothing
        
        vaSpread2.Visible = False
        vaSpread1.Visible = True
        
        
        If vaSpread1.MaxRows > 0 Then
           vaSpread1.Row = 1
           vaSpread1.Col = 1
           codigo = ""
           codigo = Val(vaSpread1.text)
           vaSpread1.SetActiveCell 1, 1
        End If
    End If
    
    'Toolbar1.Buttons(2).Visible = True
    'Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(10).Enabled = True
        
            
    Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
    If (vaSpread1.MaxRows <= 0) Then Toolbar1.Buttons(10).Enabled = False
    fg_descarga
End Sub

Private Sub Cancela()
    OpGr = True
    vaSpread2.Row = vaSpread2.ActiveRow
    
    MoverDatosGrilla
    
    OpGr = False
    
    fpText(0).Enabled = True
    Image1(0).Enabled = True
    Toolbar2.Enabled = True
    TipoOp = ""
    
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(8).Visible = True
    
    If fpayuda(0).Caption <> "" Then
        Toolbar1.Buttons(1).Enabled = True
    End If
    
    vaSpread2.DeleteRows 1, vaSpread2.MaxRows
    
End Sub


