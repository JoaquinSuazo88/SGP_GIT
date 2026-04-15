VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form M_ImLprE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Lista de Precio Desde Excel"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   6255
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   405
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
         Text            =   "01/2020"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/yyyy"
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   3480
         TabIndex        =   3
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes Actualizar"
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
         Left            =   120
         TabIndex        =   14
         Top             =   450
         Width           =   1260
      End
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   5400
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
      SpreadDesigner  =   "M_ImLprE.frx":0000
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6255
      Begin VB.OptionButton Option1 
         Caption         =   "Match x Código SGP"
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
         Left            =   3480
         TabIndex        =   11
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Match x Código SAC"
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
         Left            =   480
         TabIndex        =   10
         Top             =   1320
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "M_ImLprE.frx":0205
         Left            =   1515
         List            =   "M_ImLprE.frx":0207
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   795
         Width           =   4575
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1515
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   4575
         _Version        =   196608
         _ExtentX        =   8070
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   1560
         TabIndex        =   9
         Top             =   920
         Width           =   4575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Formato Excel"
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
         Left            =   120
         TabIndex        =   8
         Top             =   435
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Hoja"
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
         TabIndex        =   7
         Top             =   885
         Width           =   1110
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   6255
      _Version        =   393216
      _ExtentX        =   11033
      _ExtentY        =   5530
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
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
      SpreadDesigner  =   "M_ImLprE.frx":0209
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6630
      Left            =   6525
      TabIndex        =   5
      Top             =   0
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   11695
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Cd 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
            Picture         =   "M_ImLprE.frx":1B7B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Periodo Actualizar"
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
      Left            =   120
      TabIndex        =   15
      Top             =   6360
      Visible         =   0   'False
      Width           =   1560
   End
End
Attribute VB_Name = "M_ImLprE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Dim RS As New ADODB.Recordset
Dim MsgTitulo  As String

Private Sub Form_Activate()
fg_descarga
If Trim(ws_respuesta) <> "" Then Text1.text = ws_respuesta: Text1.SelStart = Len(ws_respuesta): ws_respuesta = ""
End Sub

Private Sub Form_Load()
On Error GoTo Man_Error
Me.Height = 7110
Me.Width = 7140
fg_centra Me
MsgTitulo = "Importar Lista de Precio Desde Excel"
fg_carga ""
Toolbar1.ImageList = Partida.IL1
Toolbar1.Buttons.Clear
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpDateTime1(0).DateTimeFormat = UserDefined
fpDateTime1(0).UserDefinedFormat = "mm/yyyy"
fpDateTime1(0).text = Format(Date, "mm/yyyy")
vaSpread1.MaxRows = 0
fg_descarga
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Titulo
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
End Sub

Private Sub fpText1_ButtonHit(Button As Integer, NewIndex As Integer)
Dim List() As String
Dim ListCount As Integer
Dim fromRight As Long, i As Long
Dim handle As Integer
Dim myPath As String
Dim f As Boolean

ReDim List(1)

CD.DialogTitle = "Seleccionar un archivo XLS"
CD.Filter = "Todos los archivos|*.*|Archivos de texto (*.xls)|*.xls"
CD.FilterIndex = 2
CD.Flags = cdlOFNFileMustExist
CD.ShowOpen
fpText1.text = CD.FileName
If Len(fpText1.text) < 0 Then Exit Sub

fromRight = InStrRev(CD.FileName, "\", , vbTextCompare)
If fromRight > 1 Then
   myPath = Left(CD.FileName, fromRight)
End If
vaSpread2.MaxRows = 0: vaSpread2.MaxRows = 500
vaSpread2.maxcols = 0: vaSpread2.maxcols = 500
f = vaSpread2.GetExcelSheetList(CD.FileName, List, ListCount, (myPath & "log.txt"), handle, True)
If (ListCount - 1 > 1) Then
   ReDim List(ListCount - 1)
   f = vaSpread2.GetExcelSheetList(CD.FileName, List, ListCount, (myPath & "log.txt"), handle, False)
End If

Combo1.Clear
For i = 0 To ListCount - 1
    Combo1.AddItem (List(i))
Next i
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long, isel As Integer, filepath As String, codigo As String, precio As Double, anomes As Long
Dim dbexcel As Database, cSpi As Long
Select Case Button.Index
Case 1 '-------> Procesar Información
    If Trim(fpText1.text) = "" Then MsgBox "Debe seleccionar ubicación del archivo excel...", vbCritical, MsgTitulo: Exit Sub
    If Combo1.ListIndex = -1 Then MsgBox "Debe seleccionar la hoja de la planilla selecionada...", vbCritical, MsgTitulo: Exit Sub
    If vaSpread1.MaxRows < 1 Then MsgBox "Debe seleccionar el periodo a importa precio...", vbCritical, MsgTitulo: Exit Sub
    isel = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then isel = 1: Exit For
    Next i
    If isel = 0 Then MsgBox "Debe seleccionar a lo menor una lista de precio", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fg_carga ""
    '-------> rescatar el spid
    Set RS = vg_db.Execute("SELECT @@spid spid")
    If Not RS.EOF Then cSpi = RS!spid
    RS.Close: Set RS = Nothing
    '-------> Borrar tabla paso lista precio
    vg_db.Execute "DELETE paso_listaprecio WHERE lpr_spid=" & cSpi & " AND lpr_usu='" & Trim(vg_NUsr) & "'"
    '-------> Insertar dato tabla paso lista precio
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then
           vaSpread1.Col = 3: codigo = vaSpread1.text
           vaSpread1.Col = 4: anomes = vaSpread1.text
           vg_db.Execute ("INSERT INTO paso_listaprecio (lpr_spid, lpr_usu, lpr_codigo, lpr_anomes) values (" & cSpi & ", '" & Trim(vg_NUsr) & "', " & Val(codigo) & ", " & anomes & ")")
        End If
    Next i
    Frame1.Enabled = False
    Frame2.Enabled = False
    vaSpread1.Col = -1: vaSpread1.Row = -1: vaSpread1.Lock = True
    Label2(2).Caption = "": Label2(2).Visible = True
    SheetName = Trim(Combo1.text) & "$"
    filepath = Trim(fpText1.text)
    Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
    Set RsExcel = dbexcel.OpenRecordset(SheetName)
    RsExcel.MoveFirst
    
    Do While RsExcel.EOF <> True
       DoEvents
       If RsExcel.Fields(0).Value = "*" Then Exit Do
       codigo = "": codigo = IIf(Not IsNull(RsExcel.Fields(0).Value), RsExcel.Fields(0).Value, "")
       If Trim(codigo) <> "" Then
          precio = 0
          If IsNumeric(RsExcel.Fields(2)) Then precio = RsExcel.Fields(2)
          If Option1(0).Value = True And precio > 0 Then
             '-------> Actualiza por código SAC
             vg_db.Execute "UPDATE b_detlistaprecio SET dlp_precio=" & precio & ", dlp_usuario='" & Trim(vg_NUsr) & "' " & _
                           "FROM b_detlistaprecio a, paso_listaprecio b, b_productos c, b_formatocompras d, b_formatocomprassgp e " & _
                           "WHERE a.dlp_codigo = b.lpr_codigo " & _
                           "AND   a.dlp_anomes = b.lpr_anomes " & _
                           "AND   a.dlp_codpro = c.pro_codigo " & _
                           "AND   c.pro_codigo = e.fcs_codsgp " & _
                           "AND   d.foc_codsac = e.fcs_codsac " & _
                           "AND  (d.foc_flexec = 0 OR (d.foc_flexec = -1 AND d.foc_vigfin > " & Format(Date, "dd/mm/yyyy") & ")) " & _
                           "AND   e.fcs_sgppre = 1 " & _
                           "AND   d.foc_codsac = '" & codigo & "' " & _
                           "AND   b.lpr_spid   = " & cSpi & " " & _
                           "AND   lpr_usu      = '" & Trim(vg_NUsr) & "'"
          ElseIf precio > 0 Then
             '-------> Actualiza por código SGPADM
             vg_db.Execute "UPDATE b_detlistaprecio SET dlp_precio=" & precio & ", dlp_usuario='" & Trim(vg_NUsr) & "' " & _
                           "FROM b_detlistaprecio a, paso_listaprecio b " & _
                           "WHERE a.dlp_codigo=b.lpr_codigo " & _
                           "AND   a.dlp_anomes=b.lpr_anomes " & _
                           "AND   a.dlp_codpro='" & codigo & "' " & _
                           "AND   b.lpr_spid=" & cSpi & " " & _
                           "AND   lpr_usu='" & Trim(vg_NUsr) & "'"
          End If
       End If
       Label2(2).Caption = IIf(Not IsNull(RsExcel.Fields(1).Value), RsExcel.Fields(1).Value, "")
       RsExcel.MoveNext
    Loop
    RsExcel.Close: Set RsExcel = Nothing
    Label2(2).Caption = ""
    Label2(2).Visible = False
    MsgBox "Generación grabado finalizado sin problema", vbInformation + vbOKOnly, MsgTitulo
    Frame1.Enabled = True
    Frame2.Enabled = True
    vaSpread1.Col = -1: vaSpread1.Row = -1: vaSpread1.Lock = False
    fg_descarga
Case 3 '-------> Salir
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1 '-------> Traer lista de precio
    fg_carga ""
    Set RS = vg_db.Execute("sgpadm_s_listaprecio 6, 0, " & Format(fpDateTime1(0).text, "yyyymm") & ", '" & vg_NUsr & "'")
    vaSpread1.Visible = 0
    vaSpread1.MaxRows = 0
    If Not RS.EOF Then
       Do While Not RS.EOF
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
          vaSpread1.Row = vaSpread1.MaxRows
          vaSpread1.Col = 1: vaSpread1.text = "0"
          vaSpread1.Col = 2: vaSpread1.text = RS!lpr_codigo & " - " & Trim(RS!lpr_nombre)
          vaSpread1.Col = 3: vaSpread1.text = RS!lpr_codigo
          vaSpread1.Col = 4: vaSpread1.text = RS!dlp_anomes
          RS.MoveNext
       Loop
    End If
    RS.Close: Set RS = Nothing
    vaSpread1.Visible = True
    fg_descarga
End Select
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
If Col = 1 And Row = 0 Then vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.text = IIf(vaSpread1.Value = "1", "0", "1")
End Sub
