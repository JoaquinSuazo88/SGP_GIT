VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form B_SalBod 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4815
   ClientLeft      =   2085
   ClientTop       =   2010
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6495
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5955
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   4
         Top             =   495
         Width           =   1470
         _Version        =   196608
         _ExtentX        =   2593
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
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   -1  'True
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
         Text            =   ""
         DateCalcMethod  =   3
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
         BorderDropShadow=   0
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
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   4065
         TabIndex        =   5
         Top             =   495
         Width           =   1470
         _Version        =   196608
         _ExtentX        =   2593
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
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   -1  'True
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
         Text            =   ""
         DateCalcMethod  =   3
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
         BorderDropShadow=   0
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
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Termino"
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
         Left            =   3285
         TabIndex        =   3
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inicio"
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
         Left            =   585
         TabIndex        =   1
         Top             =   600
         Width           =   480
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   5955
      _Version        =   393216
      _ExtentX        =   10504
      _ExtentY        =   6376
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
      MaxRows         =   10
      OperationMode   =   3
      SelectBlockOptions=   0
      SpreadDesigner  =   "B_SalBod.frx":0000
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   4815
      Left            =   5955
      TabIndex        =   6
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   8493
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "B_SalBod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim i As Long, ibusca As Long
Dim lc_codigo As String, lc_tipo As String
Dim icombo As Integer

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, vg_nombre

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)

fg_centra Me
fg_carga ""
icombo = 1

'-------> LlenaDatos
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpDateTime1(1).Enabled = False
icombo = 0
fg_descarga
lc_codigo = vg_codigo
vg_codigo = ""

lc_tipo = vg_nombre
vg_nombre = IIf(lc_tipo = "SP", "Salida de Bodega a Producción", _
            IIf(lc_tipo = "SE", "Salida de Ventas Servicios Especiales", _
            IIf(lc_tipo = "DE", "Devolución de Ventas Servicios Especiales", _
            IIf(lc_tipo = "GV", "Guía Venta SAP", _
            IIf(lc_tipo = "DP", "Devolución de Producción a Bodega", _
            IIf(lc_tipo = "ME", "Mermas", _
            IIf(lc_tipo = "TR", "Traspaso", _
            IIf(lc_tipo = "FA", "Factura - Venta directa", "Guia despacho - Venta directa"))))))))

Me.Caption = vg_nombre
vaSpread1.MaxRows = 0
vaSpread1.Col = 3
vaSpread1.Row = 0
vaSpread1.text = IIf(lc_tipo = "ME", "Tipo Merma", _
                 IIf(lc_tipo = "GV", "Cliente", _
                 IIf(lc_tipo = "TR", "Tipo Traspaso", _
                 IIf(lc_tipo = "FA" Or lc_tipo = "GD", "Contrato", "Regimen - Servicio"))))
fpDateTime1(0).text = Format(Date - 30, "dd/mm/yyyy")
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, vg_nombre

End Sub

Private Sub fpDateTime1_Change(Index As Integer)

On Error GoTo Man_Error

Dim descrip As String, sql1 As String, sql2 As String

If fpDateTime1(0).text = "" And Index = 0 Then
    
    fpDateTime1(1).Enabled = False
    fpDateTime1(1).text = ""
    Exit Sub

Else
    
    fpDateTime1(1).Enabled = True

End If
vaSpread1.MaxRows = 0
If Trim(fpDateTime1(0).text) = "" Or Trim(fpDateTime1(1).text) = "" Then Exit Sub
If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then Exit Sub
If CDate(fpDateTime1(0).text) > CDate(fpDateTime1(1).text) Then fpDateTime1(1).text = fpDateTime1(0).text: Exit Sub
If fpDateTime1(1).text = "" Then Exit Sub

descrip = ""
sql1 = IIf(vg_tipbase = "1", " CDate('" & fpDateTime1(0).text & "') ", " '" & Format(fpDateTime1(0).text, "yyyymmdd") & "' ")
sql2 = IIf(vg_tipbase = "1", " CDate('" & fpDateTime1(1).text & "') ", " '" & Format(fpDateTime1(1).text, "yyyymmdd") & "' ")

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
        
If lc_tipo = "TR" Then
   
   RS1.Open "SELECT DISTINCT tov_numdoc, tov_fecemi, tov_codser, tov_estdoc, tov_tipdoc " & _
            "FROM  b_totventas " & _
            "WHERE tov_rutcli = '" & lc_codigo & "' " & _
            "AND   tov_tipdoc = '" & lc_tipo & "' " & _
            "AND   tov_codbod = " & vg_codbod & " " & _
            "AND   tov_fecemi >= " & sql1 & " " & _
            "AND   tov_fecemi <= " & sql2 & " " & _
            "ORDER BY tov_fecemi", vg_db, adOpenStatic

ElseIf lc_tipo = "FA" Or lc_tipo = "GD" Then
   
   RS1.Open "SELECT DISTINCT tov.tov_numdoc, tov.tov_fecemi, cli.cli_nombre, tov.tov_estdoc, tov.tov_tipdoc " & _
            "FROM  b_totventas tov, b_clientes cli " & _
            "WHERE tov.tov_rutcli = cli.cli_codigo " & _
            "AND   tov.tov_tipdoc = '" & lc_tipo & "' " & _
            "AND   tov.tov_codbod = " & vg_codbod & " " & _
            "AND   tov.tov_fecemi >= " & sql1 & " " & _
            "AND   tov.tov_fecemi <= " & sql2 & " " & _
            "ORDER BY tov.tov_fecemi", vg_db, adOpenStatic

ElseIf lc_tipo = "ME" Then
   
   RS1.Open "SELECT DISTINCT tov.tov_numdoc, tov.tov_fecemi, aju.aju_nombre, tov.tov_estdoc, tov.tov_tipdoc " & _
            "FROM  b_totventas tov, a_tipoajuste aju " & _
            "WHERE tov.tov_rutcli = '" & lc_codigo & "' AND tov.tov_tipdoc = '" & lc_tipo & "' AND tov.tov_codbod = " & vg_codbod & " " & _
            "AND   tov.tov_codser = aju.aju_codigo AND tov.tov_fecemi >= " & sql1 & " " & _
            "AND   tov.tov_fecemi <= " & sql2 & " ORDER BY tov.tov_fecemi", vg_db, adOpenStatic

ElseIf lc_tipo = "SP" Or lc_tipo = "DP" Then

'Clipboard.Clear
'Clipboard.SetText
    If Not vg_tipser Then
       
       RS1.Open "SELECT DISTINCT tov.tov_numdoc, tov.tov_fecpro, ser.ser_nombre, tov.tov_estdoc, reg.reg_nombre " & _
                "FROM b_totventas tov, a_regimen reg, a_servicio ser " & _
                "WHERE tov.tov_rutcli = '" & lc_codigo & "' AND tov.tov_tipdoc = '" & lc_tipo & "' AND tov.tov_codbod = " & vg_codbod & " " & _
                "AND   tov.tov_codser = ser.ser_codigo AND tov.tov_codreg = reg.reg_codigo " & _
                "AND   tov.tov_fecpro >= " & sql1 & " " & _
                "AND   tov.tov_fecpro <= " & sql2 & " ORDER BY tov.tov_fecpro DESC", vg_db, adOpenStatic
    
    Else
       
       RS1.Open "SELECT DISTINCT tov.tov_numdoc, tov.tov_fecpro, '' AS ser_nombre, tov.tov_estdoc, '' as reg_nombre " & _
                "FROM b_totventas tov " & _
                "WHERE tov.tov_rutcli = '" & lc_codigo & "' AND tov.tov_tipdoc = '" & lc_tipo & "' AND tov.tov_codbod = " & vg_codbod & " " & _
                "AND   tov.tov_fecpro >= " & sql1 & " " & _
                "AND   tov.tov_fecpro <= " & sql2 & " ORDER BY tov.tov_fecpro DESC", vg_db, adOpenStatic
    
    End If

ElseIf lc_tipo = "SE" Or lc_tipo = "DE" Then

     Set RS1 = vg_db.Execute("sgp_Sel_AyudaSalDebVentaServiciosEspeciales '" & lc_codigo & "', '" & lc_tipo & "', " & sql1 & ", " & sql2 & ", " & vg_codbod & "")

ElseIf lc_tipo = "GV" Then
    
    RS1.Open "SELECT a.tgv_numdoc, a.tgv_fecing, a.tgv_rutcli, '' AS ser_nombre, b.cli_nombre " & _
             "FROM b_totguiavta a, b_clientes b, b_sucursalcliente c " & _
             "WHERE a.tgv_rutcli = b.cli_codigo " & _
             "AND   a.tgv_codsuc = c.scl_codigo " & _
             "AND   c.scl_codcli = b.cli_codigo " & _
             "AND   a.tgv_fecing >= " & sql1 & " " & _
             "AND   a.tgv_fecing <= " & sql2 & " ORDER BY a.tgv_fecing", vg_db, adOpenStatic

End If

If Not RS1.EOF Then
    
    Do While Not RS1.EOF
        
        descrip = IIf(lc_tipo = "TR", IIf(RS1(2) = 1, "Entrada", "Salida"), _
                  IIf(lc_tipo = "SP" Or lc_tipo = "DP" Or lc_tipo = "GV" Or lc_tipo = "DE" Or lc_tipo = "SE", RS1(4) & " - " & RS1(2), RS1(2)))
        
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 1: vaSpread1.text = RS1(0)
        vaSpread1.Col = 2: vaSpread1.text = Format(RS1(1), "dd/mm/yyyy")
        vaSpread1.Col = 3: vaSpread1.text = descrip
        vaSpread1.Col = 4: vaSpread1.text = IIf(RS1(3) = "", "ACTIVADA", IIf(RS1(3) = "A", "ANULADA", "PENDIENTE"))
        RS1.MoveNext
    
    Loop

End If
RS1.Close
Set RS1 = Nothing

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, vg_nombre

End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, vg_nombre

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index

Case 1
    
    MoverDatos
    fg_descarga

Case 3
    
    Cerrar
    fg_descarga

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, vg_nombre

End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

MoverDatos

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, vg_nombre

End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
MoverDatos

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, vg_nombre

End Sub

Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 27
    
    Cerrar

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, vg_nombre

End Sub

Private Sub MoverDatos()

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
vg_codigo = Trim(vaSpread1.text)
Cerrar

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, vg_nombre

End Sub

Sub Cerrar()

On Error GoTo Man_Error

Me.Hide
Unload Me

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, vg_nombre

End Sub


