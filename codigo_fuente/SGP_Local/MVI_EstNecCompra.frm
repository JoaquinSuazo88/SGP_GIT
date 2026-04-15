VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Object = "{1DF3AFED-99E0-4474-9900-954B8FD24E86}#1.0#0"; "ChilkatMail2.dll"
Begin VB.Form MVI_EstNecCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descarga de pedidos de compra en Excel"
   ClientHeight    =   8730
   ClientLeft      =   1500
   ClientTop       =   1785
   ClientWidth     =   17775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   17775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   5400
      TabIndex        =   11
      Top             =   3360
      Width           =   5055
      Begin VB.Label Label3 
         Caption         =   "Un Momento Generando Pedido"
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
         Left            =   1080
         TabIndex        =   12
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1425
      Index           =   1
      Left            =   5370
      TabIndex        =   3
      Top             =   120
      Width           =   7095
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1380
         TabIndex        =   0
         Top             =   285
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1380
         TabIndex        =   1
         Top             =   630
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1658
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
         ButtonStyle     =   1
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
         Text            =   "10/2016"
         DateCalcMethod  =   4
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
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   360
         Left            =   2520
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
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList1"
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
         Caption         =   "Fecha(mm/aa)"
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
         Left            =   90
         TabIndex        =   6
         Top             =   690
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contrato"
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
         Left            =   90
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2610
         Picture         =   "MVI_EstNecCompra.frx":0000
         Top             =   195
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3045
         TabIndex        =   4
         Top             =   285
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3090
         TabIndex        =   7
         Top             =   330
         Width           =   3855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8730
      Left            =   17265
      TabIndex        =   8
      Top             =   0
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   15399
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   6795
      Left            =   30
      TabIndex        =   9
      Top             =   1710
      Width           =   17010
      _Version        =   393216
      _ExtentX        =   30004
      _ExtentY        =   11986
      _StockProps     =   64
      AutoClipboard   =   0   'False
      ButtonDrawMode  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   2
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
      MaxCols         =   12
      MaxRows         =   1
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "MVI_EstNecCompra.frx":030A
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _Version        =   393216
      _ExtentX        =   2143
      _ExtentY        =   1085
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
      MaxRows         =   1
      SpreadDesigner  =   "MVI_EstNecCompra.frx":0A25
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   300
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
            Picture         =   "MVI_EstNecCompra.frx":0DAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ACTIVEZIPLib.ActiveZip AZ1 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   180
      Top             =   1000
      Visible         =   0   'False
      Width           =   300
   End
   Begin CHILKATMAILLib2Ctl.ChilkatMailMan2 oMail 
      Left            =   2640
      OleObjectBlob   =   "MVI_EstNecCompra.frx":1144
      Top             =   600
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP oFTP 
      Left            =   1920
      OleObjectBlob   =   "MVI_EstNecCompra.frx":1168
      Top             =   870
   End
End
Attribute VB_Name = "MVI_EstNecCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msgtitulo As String

Private Sub Command1_Click()
  If vaSpread1.MaxRows = 0 Then Exit Sub
  'exporta el recordset a excel
  Call Exportar_ADO_Excel(Me, sql, "C:\NecCompraExcel.xls")
End Sub

Private Sub Form_Activate()
fg_descarga
'-------> Traer fecha cierre día
 TraerFechaCierre
End Sub

Private Sub LimpiarControles()
    vaSpread1.MaxRows = 0
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
fg_carga ""
Me.Height = 9210
Me.Width = 17895
fg_centra Me
Msgtitulo = "Pedido Mensual Ruta"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_ImportarInventario", , tbrDefault, "A_ImportarInventario"): BtnX.Visible = True: BtnX.ToolTipText = "Importar Archivo "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_ExportarInventario", , tbrDefault, "A_ExportarInventario"): BtnX.Visible = False: BtnX.ToolTipText = "Exportar Excel "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
vaSpread1.MaxRows = 0
fpDateTime1.text = Format(Date, "mm/yyyy")
fpText.Enabled = ModCasino
Image1.Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda.Caption = MuestraCasino(2)
Label3.Visible = False
Frame2.Visible = False
fg_descarga
Call LimpiarControles
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If Not IsDate(fpDateTime1.text) Then Exit Sub
End Sub

Private Sub fpDateTime2_Change()
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If Not IsDate(fpDateTime2.text) Then Exit Sub
End Sub

Private Sub fpText_Change()
Dim RS As New ADODB.Recordset
If fpText.text = "" Then fpayuda.Caption = "": Exit Sub
Toolbar1.Buttons(1).Enabled = True: Toolbar1.Buttons(3).Enabled = False: vaSpread1.MaxRows = 0
RS.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & fpText.text & "' AND cli_tipo = 0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda.Caption = "": Exit Sub
fpayuda.Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    Image1_Click
End Select
End Sub

Private Sub Image1_Click()
vg_left = fpayuda.Left + 2300
vg_nombre = "": vg_codigo = ""
B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
B_TabEst.Show 1
Me.Refresh
If vg_codigo = "" Then Exit Sub
fpText.text = vg_codigo: fpayuda.Caption = vg_nombre
Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(8).Enabled = False: vaSpread1.MaxRows = 0
If Me.Visible Then fpDateTime1.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 1 '-------> Importar
     MVI_ImpArchivos.Show 1
Case 3 '-------> Exportar
    Toolbar1.Enabled = False
    P_EIInve.Inicio "Exportar Pedido Mensual Ruta", "EP", 0
    P_EIInve.Show 1
    Toolbar1.Enabled = True
Case 5 '-------> Salir
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Dim RS As New ADODB.Recordset
Dim sql As String
Dim sql1 As String
Dim NomExcelZip As String
Dim i As Long
Dim NameTemp As String

'NameTemp = LimpiaDato(Trim(fpText.text)) & Format(fpDateTime1.text, "yyyymm")
'If Trim(ValidarUsoOpcionesSistema(NameTemp)) <> "0" Then
'   MsgBox "El pedido con los parametros ingresados, actualmente esta siendo usado por el usuario: '" & UCase(ValidarUsoOpcionesSistema(NameTemp)) & "', podra ingresar cuando el usuario termine de trabajar en ella", vbExclamation + vbOKOnly, Msgtitulo
'   Exit Sub
'End If

Select Case Button.Index
Case 1
    If Not IsDate(fpDateTime1.text) Then Exit Sub
    '-------> Validar si la minuta es teorica normal
    sql1 = " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) "
    RS.Open "SELECT DISTINCT a.min_codigo FROM b_minuta a, a_servicio b WHERE a.min_codser = b.ser_codigo and b.ser_activo = '1' and a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND " & sql1 & " = " & Format(fpDateTime1.text, "yyyymm") & " AND min_indblo IN (2,11)", vg_db, adOpenStatic
    If Not RS.EOF Then
 '      DropTeblaTmp (NameTemp)
       RS.Close: Set RS = Nothing: MsgBox "Existe Bloque Minuta, pedido se cancela", vbExclamation + vbOKOnly, Msgtitulo
       Exit Sub
    End If
    RS.Close: Set RS = Nothing
    
    '-------> Validar si existe datos ruta carga
    RS.Open "select top 1 id_carga from ruta_compras where ID_centro_de_costo = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    If RS.EOF Then
'       DropTeblaTmp (NameTemp)
       RS.Close: Set RS = Nothing: MsgBox "No existe datos cargados rutas compras, pedido se cancela", vbExclamation + vbOKOnly, Msgtitulo
       Exit Sub
    End If
    RS.Close: Set RS = Nothing
    
    '-------> Validar si existe datos convenios
    RS.Open "select top 1 Reg_info from convenios_mvi where Ce = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    If RS.EOF Then
'       DropTeblaTmp (NameTemp)
       RS.Close: Set RS = Nothing: MsgBox "No existe datos cargados convenios, pedido se cancela", vbExclamation + vbOKOnly, Msgtitulo
       Exit Sub
    End If
    RS.Close: Set RS = Nothing

   Label3.Visible = True
   Frame2.Visible = True
   Label3.Caption = "Un momento generando pedido ..."
   
    
    vaSpread1.MaxRows = 0
    vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = &HC0FFFF
    estexi = True
   
   'dispara sp ppal.
   fg_carga ""
   Toolbar1.Enabled = False
   sql = " MVI_NEC_COMPRA"
   sql = sql & " '" & fpText & "'"
   sql = sql & " , '" & Left(fpDateTime1.text, 2) & "' "
   sql = sql & " , '" & Right(fpDateTime1.text, 4) & "'"

   Set RS = vg_db.Execute(sql)
    
   '-------> Inicio LLenar grilla
   Dim AuxCodIngrediente As String
   AuxIngrediente = ""
   vaSpread1.MaxRows = 0
    If Not RS.EOF Then
       Toolbar1.Buttons(3).Enabled = True
    Else
       Toolbar1.Buttons(3).Enabled = False
    End If
    Do While Not RS.EOF
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        If AuxCodIngrediente <> RS(1) Then
           For i = 1 To 12
               vaSpread1.Col = i
               vaSpread1.BackColor = Shape1(2).FillColor
           Next i
           AuxCodIngrediente = RS(1)
        End If
        vaSpread1.Col = 1 ' IdCompra
        vaSpread1.text = IIf(RS(0) = 0, "", RS(0))
        
        vaSpread1.Col = 2 ' Cod. Ingrediente
        vaSpread1.text = RS(1)
        vaSpread1.Col = 3 ' Des. Ingrediente
        vaSpread1.text = RS(2)
        vaSpread1.Col = 4 ' Proveedor
        vaSpread1.text = RS(3)
        vaSpread1.Col = 5 ' Familia Producto
        vaSpread1.text = RS(4)
        vaSpread1.Col = 6 ' Centro Costo
        vaSpread1.text = RS(5)
        vaSpread1.Col = 7 ' Codigo Producto SAP
        vaSpread1.text = RS(6)
        vaSpread1.Col = 8 ' Des. Producto SAp
        vaSpread1.text = RS(7)
        vaSpread1.Col = 9 ' Unidad
        vaSpread1.text = RS(8)
        vaSpread1.Col = 10 ' Fecha Despacho
        vaSpread1.text = RS(9)
        vaSpread1.Col = 11 ' Cantidad Solicitar
        vaSpread1.text = RS(10)
        
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    
    '-------> validar si existe pedido
    If vaSpread1.MaxRows < 1 Then
       fg_descarga
       Label3.Visible = False
       Frame2.Visible = False
       Label3.Caption = ""
'       DropTeblaTmp (NameTemp)
       MsgBox "Por favor verificar si existen " & VgLinea & VgLinea & "- Rutas para la fecha consultada " & VgLinea & "- Convenios vigentes para la fecha consultada " & VgLinea, vbInformation + vbOKOnly, Msgtitulo
       Toolbar1.Enabled = True
       Exit Sub
    End If
    
    Dim MyBuffer As Variant
    Dim IdRuta As Long
    Dim CodIngrediente As String
    Dim CodProveedor As String
    Dim FamProducto As String
    Dim CenCosto As String
    Dim codproducto As String
    Dim FechaDespacho As String
    Dim total As Double
    Dim CodProductoSGP As String
    '-------> General Pedido & Minuta Real
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<GrabaDetallePedido>"

    For i = 1 To MVI_EstNecCompra.vaSpread1.MaxRows
        Let MyBuffer = MyBuffer & " <DetallePedido"
        MyBuffer = MyBuffer & " Op = " & Chr(34) & 0 & Chr(34)
        desc = Replace(Trim(desc), Chr(34), "&quot;")
        desc = Replace(Trim(desc), Chr(38), "&amp;")
        desc = Replace(Trim(desc), Chr(39), "&apos;")
        desc = Replace(Trim(desc), Chr(60), "&lt;")
        desc = Replace(Trim(desc), Chr(62), "&gt;")

        MVI_EstNecCompra.vaSpread1.Row = i
        
        MVI_EstNecCompra.vaSpread1.Col = 1 'Id Ruta de Compras
        IdRuta = IIf(MVI_EstNecCompra.vaSpread1.text = "", 0, MVI_EstNecCompra.vaSpread1.text)

        MVI_EstNecCompra.vaSpread1.Col = 2 'Código Ingrediente
        CodIngrediente = MVI_EstNecCompra.vaSpread1.text

        MVI_EstNecCompra.vaSpread1.Col = 3 'Descripción Ingrediente

        MVI_EstNecCompra.vaSpread1.Col = 4 'Código Proveedor SAP
        CodProveedor = MVI_EstNecCompra.vaSpread1.text

        MVI_EstNecCompra.vaSpread1.Col = 5 'Código Familia Producto

        MVI_EstNecCompra.vaSpread1.Col = 6 'Centro costo

        MVI_EstNecCompra.vaSpread1.Col = 7 'Código Producto SAP
        codproducto = MVI_EstNecCompra.vaSpread1.text

        MVI_EstNecCompra.vaSpread1.Col = 8 'Descripción Producto

        MVI_EstNecCompra.vaSpread1.Col = 9 'Unidad

        MVI_EstNecCompra.vaSpread1.Col = 10 'Fecha Despacho
        FechaDespacho = IIf(MVI_EstNecCompra.vaSpread1.text = "", 0, Format(MVI_EstNecCompra.vaSpread1.text, "yyyymmdd"))

        MVI_EstNecCompra.vaSpread1.Col = 11 'Total
        total = MVI_EstNecCompra.vaSpread1.text

        MVI_EstNecCompra.vaSpread1.Col = 12 'Còdigo Producto SGP
        CodProductoSGP = MVI_EstNecCompra.vaSpread1.text

        MyBuffer = MyBuffer & " IdRuta  = " & Chr(34) & IdRuta & Chr(34)
        MyBuffer = MyBuffer & " CodIngrediente = " & Chr(34) & CodIngrediente & Chr(34)
        MyBuffer = MyBuffer & " CodProveedor  = " & Chr(34) & CodProveedor & Chr(34)
        MyBuffer = MyBuffer & " CodProducto  = " & Chr(34) & codproducto & Chr(34)
        MyBuffer = MyBuffer & " FechaDespacho  = " & Chr(34) & FechaDespacho & Chr(34)
        MyBuffer = MyBuffer & " Total  = " & Chr(34) & total & Chr(34)
        MyBuffer = MyBuffer & " CodProductoSGP  = " & Chr(34) & CodProductoSGP & Chr(34)
        Let MyBuffer = MyBuffer & "/>"

    Next i
    TraerFechaCierre
    Let MyBuffer = MyBuffer & "</GrabaDetallePedido>"
    vg_db.Execute ("sgp_iu_GenerarPedidoMinutaReal '" & MyBuffer & "', '" & LimpiaDato(fpText.text) & "','" & Left(fpDateTime1.text, 2) & "', '" & Right(fpDateTime1.text, 4) & "', '" & vg_ciedia & "'")

    Label3.Visible = False
    Frame2.Visible = False
    Label3.Caption = ""

   Label3.Visible = True
   Frame2.Visible = True
   Label3.Caption = "Un momento enviando pedido por correo..."

    '-------> generar excel
    Dim NomExcel As String
    NomExcel = dir_trabajo & "Pedido" & Trim(fpText.text) & Format(Date, "yyyymmdd") & ".xls"
    
    If Dir(NomExcel) <> "" Then Kill NomExcel   'borrar base datos si existe
    If Dir(Mid((NomExcel), 1, Len((NomExcel)) - 3) & "txt") <> "" Then Kill Mid((NomExcel), 1, Len((NomExcel)) - 3) & "txt"
    Open Mid((NomExcel), 1, Len((NomExcel)) - 3) & "txt" For Output As #1
    
    MVI_EstNecCompra.vaSpread1.Row = 0
    MVI_EstNecCompra.vaSpread1.Col = 1
    sql = " " & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 2
    sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 3
    sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 4
    sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 5
    sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 6
    sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 7
    sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 8
    sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 9
    sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 10
    sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
    MVI_EstNecCompra.vaSpread1.Col = 11
    sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
    
    Print #1, sql
           
    For i = 1 To MVI_EstNecCompra.vaSpread1.MaxRows
        sql = ""
    '    PB.Value = Val((i / MVI_EstNecCompra.vaSpread1.MaxRows) * 100)
        MVI_EstNecCompra.vaSpread1.Row = i
        MVI_EstNecCompra.vaSpread1.Col = 1
        sql = MVI_EstNecCompra.vaSpread1.text & "|"
        MVI_EstNecCompra.vaSpread1.Col = 2
        sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
        MVI_EstNecCompra.vaSpread1.Col = 3
        sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
        MVI_EstNecCompra.vaSpread1.Col = 4
        sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
        MVI_EstNecCompra.vaSpread1.Col = 5
        sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
        MVI_EstNecCompra.vaSpread1.Col = 6
        sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
        MVI_EstNecCompra.vaSpread1.Col = 7
        sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
        MVI_EstNecCompra.vaSpread1.Col = 8
        sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
        MVI_EstNecCompra.vaSpread1.Col = 9
        sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
        MVI_EstNecCompra.vaSpread1.Col = 10
        sql = sql & Format(MVI_EstNecCompra.vaSpread1.text, "mm/dd/yyyy") & "|"
        MVI_EstNecCompra.vaSpread1.Col = 11
        sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
        
        Print #1, sql
    '           Print #1, Trim(RS!pro_codigo) & ";" & Trim(RS!pro_nombre) & ";" & Trim(RS!uni_nomcor) & ";" & Round(RS!tin_stofis, vg_DCa)
    Next i
    Close #1
    Set XL = CreateObject("Excel.application")
    XL.Workbooks.OpenText Mid((NomExcel), 1, Len((NomExcel)) - 3) & "txt", , 1, 1, , , , , , , True, "|"
    XL.ActiveWorkbook.SaveAs Filename:=NomExcel, _
                                      FileFormat:=xlNormal, password:="", WriteResPassword:="", _
                                      ReadOnlyRecommended:=False, CreateBackup:=False
    XL.Quit
    Set XL = Nothing
    
    '-------> 1 Comprimir archivo excel
    NomExcelZip = dir_trabajo & "Pedido" & Trim(fpText.text) & Format(Date, "yyyymmdd") & ".zip"
    '-------> verificar si existe archivo zip destino si existe borrar
    If Dir(NomExcelZip) <> "" Then Kill NomExcelZip
    AZ1.CreateZip NomExcelZip, "": AZ1.AddFile NomExcel, "", True, "": AZ1.Close
    '-------> verificar si existe archivo mdb destino si existe borrar
    If Dir(NomExcel) <> "" Then Kill NomExcel
    If Dir(Mid((NomExcel), 1, Len((NomExcel)) - 3) & "txt") <> "" Then Kill Mid((NomExcel), 1, Len((NomExcel)) - 3) & "txt"
    
    '-------> Traer dirección correo
    Dim emailpedido As String
    RS.Open "select isnull(par_valor,'') as par_valor from a_param where par_cencos = '" & Trim(fpText.text) & "' and par_codigo = 'emailenped'", vg_db, adOpenStatic
    If Not RS.EOF Then
       emailpedido = RS!par_valor
    End If
    RS.Close: Set RS = Nothing
    
    '-------> Enviar correo
    Dim cBody As String
    cBody = ""
    cBody = "Generación Automatica Pedidos " & Format(fpDateTime1.text, "mm/yyyy") & VgLinea & VgLinea
    cBody = cBody & "IMPORTANTE: " & VgLinea
    cBody = cBody & "Este correo es informativo, favor no responder a esta dirección de correo, ya que no se encuentra habilitada para recibir mensajes." & VgLinea & VgLinea
    cBody = cBody & "Atte." & VgLinea
    cBody = cBody & "SGP Chile" & VgLinea
    vg_codigo = ""
    
    SendMail oMail, "SGP : Pedido Insumo (" & Trim(fpText.text) & ") " & Trim(fpayuda.Caption), cBody, NomExcelZip, "Email Pedido", emailpedido

    Label3.Visible = False
    Frame2.Visible = False
    Label3.Caption = ""
    If vg_codigo = "" Then
       MsgBox "Generación pedido finalizado sin problema", vbInformation + vbOKOnly, Msgtitulo
    Else
       MsgBox "Generación pedido finalizado sin problema, hay problema envio de correo", vbInformation + vbOKOnly, Msgtitulo
    End If
'    DropTeblaTmp (NameTemp)
    Toolbar1.Enabled = True
    
    fg_descarga
End Select
Exit Sub

Man_Error:
Toolbar1.Enabled = True
Label3.Visible = False
Frame2.Visible = False
Label3.Caption = ""
'DropTeblaTmp (NameTemp)
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
End Sub

