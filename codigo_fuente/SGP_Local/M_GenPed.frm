VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_GenPed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Pedidos"
   ClientHeight    =   6300
   ClientLeft      =   1575
   ClientTop       =   1965
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   30
      TabIndex        =   9
      Top             =   5520
      Width           =   11010
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   150
         TabIndex        =   10
         Top             =   180
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   688
         ButtonWidth     =   2963
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Index           =   1
      Left            =   1890
      TabIndex        =   3
      Top             =   120
      Width           =   7095
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Oculta Ingredientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3000
         TabIndex        =   2
         Top             =   720
         Width           =   2610
      End
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
         Text            =   "11/2017"
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
         Left            =   2400
         TabIndex        =   13
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3045
         TabIndex        =   7
         Top             =   285
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2610
         Picture         =   "M_GenPed.frx":0000
         Top             =   195
         Width           =   480
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
         TabIndex        =   4
         Top             =   690
         Width           =   1230
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3090
         TabIndex        =   8
         Top             =   330
         Width           =   3855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6300
      Left            =   11055
      TabIndex        =   6
      Top             =   0
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   11113
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4275
      Left            =   30
      TabIndex        =   11
      Top             =   1230
      Width           =   11025
      _Version        =   393216
      _ExtentX        =   19447
      _ExtentY        =   7541
      _StockProps     =   64
      AutoClipboard   =   0   'False
      ButtonDrawMode  =   1
      DisplayRowHeaders=   0   'False
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
      MaxCols         =   12
      MaxRows         =   1
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "M_GenPed.frx":030A
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   615
      Left            =   0
      TabIndex        =   12
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
      MaxCols         =   3
      MaxRows         =   1
      SpreadDesigner  =   "M_GenPed.frx":0A2F
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
            Picture         =   "M_GenPed.frx":0C8F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_GenPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim Fecha As Long
Dim Msgtitulo As String
Dim est As Boolean, etapa5 As Boolean, aAp1 As String, aAp2 As String
Dim vecdes() As Variant

Private Sub Check1_Click(Index As Integer)
On Error GoTo Man_Error
If est Then Exit Sub
'-------> Actualizar parametro generación pedido mensual
vg_db.BeginTrans
vg_db.Execute "UPDATE a_param SET par_valor = '" & IIf(Check1(0).Value = 1, 1, 0) & "' WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'ingpedmen'"
vg_db.CommitTrans
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Visible = False
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 6
    If Trim(vaSpread1.text) = "" And vaSpread1.BackColor <> &HFFFFC0 Then vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
Next i
vaSpread1.SetActiveCell 1, 1
vaSpread1.Visible = True
Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Form_Activate()
fg_descarga
'-------> Traer fecha cierre día
TraerFechaCierre
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
fg_carga ""
Me.Height = 6800
Me.Width = 11655
fg_centra Me
est = True
etapa5 = IIf("S" = fg_CambiaChar(GetParametro("5etapas"), ";", "','"), True, False)
Msgtitulo = "Generación pedidos"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.Enabled = False: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): BtnX.Visible = True: BtnX.Enabled = False: BtnX.ToolTipText = "Enviar"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = "Borrar "
Set BtnX = Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = ""
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Enabled = False: BtnX.ToolTipText = "Imprimir "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
Toolbar2.ImageList = Partida.IL1
Set BtnX = Toolbar2.Buttons.Add(, "A_InsertarF", , tbrDefault, "A_InsertarF"): BtnX.Visible = True: BtnX.Enabled = True: BtnX.Caption = "Agregar Producto ": BtnX.ToolTipText = "Agregar Producto "
Set BtnX = Toolbar2.Buttons.Add(, "A_EliminarF", , tbrDefault, "A_EliminarF"): BtnX.Visible = True: BtnX.Enabled = True: BtnX.Caption = "Eliminar Producto ": BtnX.ToolTipText = "Eliminar Producto"
If etapa5 Then Check1(0).Visible = False: Toolbar2.Visible = False
Frame2.Enabled = False: vaSpread1.MaxRows = 0
fpDateTime1.text = Format(Date, "mm/yyyy")
fpText.Enabled = ModCasino
Image1.Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda.Caption = MuestraCasino(2)
Check1(0).Value = IIf(0 = (fg_CambiaChar(GetParametro("ingpedmen"), ";", "','")), 0, 1)
est = False
fg_descarga
End Sub

Private Sub MoverDatos()
Dim codtip As Long, fecenv As Long, fecval As Long, diaini As Long, diafin As Long, i As Long, j As Long, X As Long, inding As Long
Dim nomtip As String, aAp As String, auxing As String, proc1 As String, proc2 As String, proc3 As String, ustock As String, nomunm As String
Dim canped As Double, reqtot As Double, Stock As Double, proped As Double, estfij As Boolean, cSpi As Long
Dim sql1 As String
fg_carga ""
vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = &HC0FFFF
Fecha = 0: codtip = 0: fecenv = 0: auxing = 0
Fecha = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)
'-------> Cargar despacho vector
RS.Open "SELECT DISTINCT pad_tipo FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "'", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe despacho, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
ReDim vecdes(RS.RecordCount, 5)
i = 1
Do While Not RS.EOF
   vecdes(i, 1) = Trim(RS!pad_tipo)
   vecdes(i, 2) = "01/" & fpDateTime1.text
   vecdes(i, 3) = IIf(Trim(RS!pad_tipo) = "S", "08/" & fpDateTime1.text, IIf(Trim(RS!pad_tipo) = "D", "11/" & fpDateTime1.text, ""))
   vecdes(i, 4) = IIf(Trim(RS!pad_tipo) = "S" Or Trim(RS!pad_tipo) = "Q", "15/" & fpDateTime1.text, IIf(Trim(RS!pad_tipo) = "D", "21/" & fpDateTime1.text, ""))
   vecdes(i, 5) = IIf(Trim(RS!pad_tipo) = "S", "22/" & fpDateTime1.text, "")
   If RS!pad_tipo = "E" Then
      RS1.Open "SELECT DISTINCT pad_tipo FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND pad_tipo = 'E' AND (pad_diario = '' or pad_diario = '0000000' or (pad_diario) IS NULL)", vg_db, adOpenStatic
      If Not RS1.EOF Then fg_descarga: RS.Close: Set RS = Nothing: RS1.Close: Set RS1 = Nothing: MsgBox "Falta definir los parametros despachos díarios, en una familia productos, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
      RS1.Close: Set RS1 = Nothing:
   End If
   RS.MoveNext: i = i + 1
Loop
RS.Close: Set RS = Nothing
aAp1 = Trim(vg_NUsr) & "_tmp_paramdesp"
'-------> Creo tabla temporal y chequeo si existe antes
fg_CheckTmp aAp1
'------->
vg_db.Execute "SELECT DISTINCT pro_codtip, 0 AS pro_previo INTO " & aAp1 & " FROM b_productos"
If vg_tipbase = "1" Then
   aAp2 = Trim(vg_NUsr) & "_tmp_productospmpdiaGenPed"
   '-------> Creo tabla temporal y b_productospmpdia
   fg_CheckTmp aAp2
   vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                 "INTO " & aAp2 & " " & _
                 "FROM b_productospmpdia " & _
                 "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                 "AND   ppd_propon > 0 " & _
                 "GROUP BY ppd_cencos, ppd_codpro"
   vg_db.Execute "ALTER TABLE " & aAp2 & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
   vg_db.Execute "UPDATE " & aAp2 & " INNER JOIN b_productospmpdia ON (" & aAp2 & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp2 & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp2 & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp2 & ".ppd_propon=b_productospmpdia.ppd_propon"
   vg_db.Execute "INSERT INTO " & aAp2 & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp2 & ")"
End If

RS.Open "SELECT * FROM " & aAp1 & "", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      vg_db.Execute "UPDATE " & aAp1 & " SET pro_previo = " & IIf(fg_BuscaenArbolNivel2(RS!pro_codtip, "a_tipopro", "tip_codigo") = 0, RS!pro_codtip, fg_BuscaenArbolNivel2(RS!pro_codtip, "a_tipopro", "tip_codigo")) & "  WHERE pro_codtip = " & RS!pro_codtip & ""
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
'-------> Validar si existe pedidos mensual
sql1 = IIf(vg_tipbase = "1", " val(mid(a.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) ")
RS.Open "SELECT DISTINCT b.mid_fecval FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND " & sql1 & " = " & Val(Fecha) & " AND b.mid_fecval > 0", vg_db, adOpenStatic
If Not RS.EOF Then
'   fecenv = RS!ped_fecenv
   fecenv = RS!mid_fecval
   RS.Close: Set RS = Nothing
   If fecenv > 0 Then
      Toolbar1.Buttons(1).Enabled = False
      Toolbar1.Buttons(3).Enabled = False
      If Not CierrePeriodo(Fecha, vg_codbod, 10) Then
         Toolbar1.Buttons(5).Visible = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
         Toolbar1.Buttons(6).Visible = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
      Else
         Toolbar1.Buttons(5).Visible = False
         Toolbar1.Buttons(6).Visible = True
      End If
      Toolbar1.Buttons(8).Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "0", False, True)
   Else
      Toolbar1.Buttons(1).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
      Toolbar1.Buttons(3).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
      Toolbar1.Buttons(5).Visible = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
      Toolbar1.Buttons(6).Visible = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
      Toolbar1.Buttons(8).Enabled = False
   End If
   proc1 = "": proc2 = "": proc3 = ""
   If vg_tipbase = "1" Then
        proc1 = "SELECT c.ing_codigo, c.ing_nombre, (SELECT DISTINCT cpi_precos FROM b_contlistpreing WHERE cpi_coding=c.ing_codigo AND cpi_cencos='" & MuestraCasino(1) & "'), e.unm_nomcor, " & _
                "b.pro_codigo, b.pro_nombre, b.pro_coduni, b.pro_facsto, " & _
                "d.uni_nomcor, a.ped_canped AS cantidad2, a.ped_canmin, (SELECT DISTINCT ppd_propon FROM " & aAp2 & " WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS cpi_precos, " & _
                "a.ped_stoact, a.ped_proped, a.ped_fecped, b.pro_ctacon, h.pro_previo AS tip_previo " & _
                "FROM  b_minutapedido a, b_productos b, b_ingrediente c, a_unidad d, a_unidadmed e, a_tipopro f, b_paramdesp g, " & aAp1 & " h " & _
                "WHERE a.ped_codpro = b.pro_codigo " & _
                "AND   b.pro_codtip = f.tip_codigo " & _
                "AND   f.tip_codigo = h.pro_codtip AND h.pro_previo = g.pad_codigo AND g.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                "AND   a.ped_coding = c.ing_codigo " & _
                "AND   c.ing_unimed = e.unm_codigo " & _
                "AND   b.pro_coduni = d.uni_codigo " & _
                "AND   a.ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                "AND   a.ped_anomes = " & Fecha & " " & _
                "AND   a.ped_tipped = 1 " & _
                "ORDER BY b.pro_ctacon, tip_previo, c.ing_codigo, c.ing_nombre, b.pro_nombre, a.ped_fecped "
        proc2 = "UNION "
        proc3 = "SELECT 'zzzfija', 'Estructura Fija', 0, '', " & _
                "b.pro_codigo, b.pro_nombre, b.pro_coduni, b.pro_facsto, " & _
                "d.uni_nomcor, a.ped_canped AS cantidad2, a.ped_canmin, (SELECT DISTINCT ppd_propon FROM " & aAp2 & " WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS cpi_precos, " & _
                "a.ped_stoact, a.ped_proped, a.ped_fecped, b.pro_ctacon, h.pro_previo AS tip_previo " & _
                "FROM  b_minutapedido a, b_productos b, b_ingrediente c, a_unidad d, a_unidadmed e, a_tipopro f, b_paramdesp g, " & aAp1 & " h " & _
                "WHERE a.ped_codpro = b.pro_codigo " & _
                "AND   b.pro_codtip = f.tip_codigo " & _
                "AND   f.tip_codigo = h.pro_codtip AND h.pro_previo = g.pad_codigo AND g.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                "AND   a.ped_coding = '' " & _
                "AND   c.ing_unimed = e.unm_codigo " & _
                "AND   b.pro_coduni = d.uni_codigo " & _
                "AND   a.ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
                "AND   a.ped_anomes = " & Fecha & " " & _
                "AND   a.ped_tipped = 1 " & _
                "ORDER BY b.pro_ctacon, tip_previo, c.ing_codigo, c.ing_nombre, b.pro_nombre, a.ped_fecped"
   Else
      proc1 = "SELECT c.ing_codigo, c.ing_nombre, (SELECT DISTINCT cpi_precos FROM b_contlistpreing WHERE cpi_coding = c.ing_codigo AND cpi_cencos = '" & MuestraCasino(1) & "'), e.unm_nomcor, " & _
              "b.pro_codigo, b.pro_nombre, b.pro_coduni, b.pro_facsto, " & _
              "d.uni_nomcor, a.ped_canped AS cantidad2, a.ped_canmin, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS cpi_precos, " & _
              "a.ped_stoact, a.ped_proped, a.ped_fecped, b.pro_ctacon, h.pro_previo AS tip_previo " & _
              "FROM  b_minutapedido a, b_productos b, b_ingrediente c, a_unidad d, a_unidadmed e, a_tipopro f, b_paramdesp g, " & aAp1 & " h " & _
              "WHERE a.ped_codpro = b.pro_codigo " & _
              "AND   b.pro_codtip = f.tip_codigo " & _
              "AND   f.tip_codigo = h.pro_codtip AND h.pro_previo = g.pad_codigo AND g.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
              "AND   a.ped_coding = c.ing_codigo " & _
              "AND   c.ing_unimed = e.unm_codigo " & _
              "AND   b.pro_coduni = d.uni_codigo " & _
              "AND   a.ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
              "AND   a.ped_anomes = " & Fecha & " " & _
              "AND   a.ped_tipped = 1 " & _
              " "
      proc2 = "UNION "
      proc3 = "SELECT 'zzzfija', 'Estructura Fija', 0, '', " & _
              "b.pro_codigo, b.pro_nombre, b.pro_coduni, b.pro_facsto, " & _
              "d.uni_nomcor, a.ped_canped AS cantidad2, a.ped_canmin, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS cpi_precos, " & _
              "a.ped_stoact, a.ped_proped, a.ped_fecped, b.pro_ctacon, h.pro_previo AS tip_previo " & _
              "FROM  b_minutapedido a, b_productos b, b_ingrediente c, a_unidad d, a_unidadmed e, a_tipopro f, b_paramdesp g, " & aAp1 & " h " & _
              "WHERE a.ped_codpro = b.pro_codigo " & _
              "AND   b.pro_codtip = f.tip_codigo " & _
              "AND   f.tip_codigo = h.pro_codtip AND h.pro_previo = g.pad_codigo AND g.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
              "AND   a.ped_coding = '' " & _
              "AND   c.ing_unimed = e.unm_codigo " & _
              "AND   b.pro_coduni = d.uni_codigo " & _
              "AND   a.ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
              "AND   a.ped_anomes = " & Fecha & " " & _
              "AND   a.ped_tipped = 1 " & _
              "ORDER BY b.pro_ctacon, tip_previo, c.ing_codigo, c.ing_nombre, b.pro_nombre, a.ped_fecped"
   End If
   RS.Open proc1 & proc2 & proc3, vg_db, adOpenStatic
Else
   RS.Close: Set RS = Nothing
   '-------> Validar productos vigentes toma de pedido
   ValidarProductoVigente
   '-------> Traer stock actual
   vaSpread2.MaxRows = 0
   RS.Open "SELECT DISTINCT bod_codpro, SUM(bod_canmer) AS bod_canmer FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " GROUP BY bod_codpro", vg_db, adOpenStatic
   If Not RS.EOF Then
      Do While Not RS.EOF
         vaSpread2.MaxRows = vaSpread2.MaxRows + 1
         vaSpread2.Row = vaSpread2.MaxRows
         vaSpread2.Col = 1: vaSpread2.text = RS!bod_codpro
         vaSpread2.Col = 2: vaSpread2.text = RS!bod_canmer
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
   '-------> Fin traer stock actual

   '-------> Traer consumo a la fecha
       aAp = Trim(vg_NUsr) & "_tmp_PedMensualFecha"
   '-------> Creo tabla temporal y chequeo si existe antes
   fg_CheckTmp aAp
   RS.Open "SELECT a.pro_codigo, SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))) AS cantidad1, SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))/a.pro_facing) AS cantidad2 INTO " & aAp & " " & _
           "FROM   b_productos a, b_contlistpreing b, b_receta c, b_recetadet d, b_minuta e, b_minutadet f " & _
           "WHERE  e.min_codigo = f.mid_codigo " & _
           "AND    f.mid_codrec = d.red_codigo " & _
           "AND    f.mid_tiprec = d.red_tiprec AND ((d.red_tiprec<>0 AND d.red_cencos = '" & MuestraCasino(1) & "') OR (d.red_tiprec = 0 AND d.red_cencos = '0')) " & _
           "AND    d.red_codigo = c.rec_codigo " & _
           "AND    d.red_codpro = b.cpi_coding AND b.cpi_cencos = '" & MuestraCasino(1) & "' " & _
           "AND    b.cpi_codped = a.pro_codigo " & _
           "AND    e.min_cencos = '" & Trim(fpText.text) & "' " & _
           "AND    e.min_fecmin >= " & Val(Year(Now) & fg_pone_cero(Str(Month(Now)), 2) & fg_pone_cero(Str(Day(Now)), 2)) & " " & _
           "AND    e.min_fecmin <= " & Val(Format(dEoM("27/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)), "yyyymmdd")) & " " & _
           "AND    f.mid_tipmin = '2' " & _
           "AND    a.pro_facing > 0 AND (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) GROUP BY a.pro_codigo", vg_db, adOpenStatic
   Set RS = Nothing
   '-------> Rutina buscar estructura fija
   RS.Open "SELECT DISTINCT mif_codreg, mif_codser FROM b_minutafija WHERE mif_cencos = '" & fpText.text & "'", vg_db, adOpenStatic
   If Not RS.EOF Then
      Do While Not RS.EOF
         estfij = False
         '-------> Buscar datos estructura fija día
         RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
                  "WHERE mfd_cencos = '" & Trim(fpText.text) & "' " & _
                  "AND   mfd_codreg = " & RS!mif_codreg & " " & _
                  "AND   mfd_codser = " & RS!mif_codser & " " & _
                  "AND   mfd_fecha >= " & Val(Year(Now) & fg_pone_cero(Str(Month(Now)), 2) & fg_pone_cero(Str(Day(Now)), 2)) & " " & _
                  "AND   mfd_fecha <= " & Val(Format(dEoM("27/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)), "yyyymmdd")) & " AND mfd_tipmin = '2'", vg_db, adOpenStatic
         If Not RS1.EOF Then estfij = True
         RS1.Close: Set RS1 = Nothing
         fecval = 0
         If Not estfij Then
            '-------> Buscar fecha mayor de estructura fija
            RS1.Open "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija WHERE mif_cencos = '" & Trim(fpText.text) & "' AND mif_codreg = " & RS!mif_codreg & " AND mif_codser = " & RS!mif_codser & "", vg_db, adOpenStatic
            If Not RS1.EOF Then fecval = IIf(IsNull(RS1!fecval), 0, RS1!fecval)
            RS1.Close: Set RS1 = Nothing
            If fecval > 0 Then
               '-------> Calcular datos desde tabla estructura fija
               For i = fg_pone_cero(Str(Day(Now)), 2) To Mid(dEoM(fg_pone_cero(Str(Day(Now)), 2) & "/" & fg_pone_cero(Str(Month(Now)), 2) & "/" & Year(Now)), 1, 2)
                   If fecval <= Year(Now) & fg_pone_cero(Str(Month(Now)), 2) & fg_pone_cero(Str(i), 2) Then
                      vg_db.Execute "INSERT INTO " & aAp & " SELECT b.pro_codigo AS pro_codigo, 0 AS cantidad1, " & _
                                    "a.mif_canpro AS cantidad2 FROM b_minutafija a, b_productos b " & _
                                    "WHERE a.mif_codpro = b.pro_codigo " & _
                                    "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                    "AND   a.mif_cencos = '" & Trim(fpText.text) & "' " & _
                                    "AND   a.mif_codreg = " & RS!mif_codreg & " " & _
                                    "AND   a.mif_codser = " & RS!mif_codser & " " & _
                                    "AND   a.mif_fecval = " & fecval & " " & _
                                    "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(Year(Now) & fg_pone_cero(Str(Month(Now)), 2) & fg_pone_cero(Str(i), 2), 2), Len(fg_Fecha_Dia(Year(Now) & fg_pone_cero(Str(Month(Now)), 2) & fg_pone_cero(Str(i), 2), 2)) - 2))) & " " & _
                                    "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0)"
                   End If
                   Set RS1 = Nothing
               Next i
            End If
         ElseIf estfij Then
            '-------> Calcular datos desde tabla estructura fija día
            vg_db.Execute "INSERT INTO " & aAp & " SELECT b.pro_codigo as pro_codigo, 0 AS cantidad1, " & _
                          "a.mfd_canpro AS cantidad2 FROM b_minutafijadia a, b_productos b " & _
                          "WHERE a.mfd_codpro = b.pro_codigo " & _
                          "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                          "AND   a.mfd_codreg = " & RS!mif_codreg & " " & _
                          "AND   a.mfd_codser = " & RS!mif_codser & " " & _
                          "AND   a.mfd_fecha >= " & Val(Year(Now) & fg_pone_cero(Str(Month(Now)), 2) & fg_pone_cero(Str(Day(Now)), 2)) & " " & _
                          "AND   a.mfd_fecha <= " & Val(Format(dEoM("27/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)), "yyyymmdd")) & " " & _
                          "AND   a.mfd_tipmin = '2' " & _
                          "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0)"
            Set RS1 = Nothing
         End If
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
   RS.Open "SELECT b.pro_codigo, b.pro_facsto, SUM(a.cantidad1) AS cantidad1, SUM(a.cantidad2) AS cantidad2 " & _
           "FROM " & aAp & " a, b_productos b " & _
           "WHERE a.pro_codigo = b.pro_codigo " & _
           "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
           "GROUP BY b.pro_codigo, b.pro_facsto", vg_db, adOpenStatic
   If Not RS.EOF Then
      Do While Not RS.EOF
         If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone) <> -1 Then
            vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone)
            vaSpread2.Col = 2
            If Not IsNull(RS!cantidad2) Then
               vaSpread2.text = (vaSpread2.text - IIf(Int(RS!cantidad2 / RS!pro_facsto) <> (RS!cantidad2 / RS!pro_facsto), Int(RS!cantidad2 / RS!pro_facsto) + 1, Round(RS!cantidad2 / RS!pro_facsto, 0)) * RS!pro_facsto)
            Else
               vaSpread2.text = (vaSpread2.text - 0)
            End If
            If Val(vaSpread2.text) < 0 Then vaSpread2.text = 0
         End If
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
   '-------> Fin traer consumo a la fecha
   fecenv = 1
   aAp = Trim(vg_NUsr) & "_tmp_PedMensual"
   '-------> Creo tabla temporal y chequeo si existe antes
   fg_CheckTmp aAp
   If vg_tipbase = "1" Then
      RS.Open "SELECT e.min_fecmin, b.cpi_coding AS ing_codigo, a.pro_codigo, a.pro_codtip, " & _
              "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))) AS cantidad1, " & _
              "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))/a.pro_facing) AS cantidad2, 0 AS ped_fecped INTO " & aAp & " " & _
              "FROM b_productos a, b_contlistpreing b, b_receta c, b_recetadet d, b_minuta e, b_minutadet f " & _
              "WHERE e.min_codigo = f.mid_codigo AND f.mid_codrec = d.red_codigo AND f.mid_tiprec = d.red_tiprec AND ((d.red_tiprec <> 0 AND d.red_cencos = '" & MuestraCasino(1) & "') OR (d.red_tiprec = 0 AND d.red_cencos = '0')) " & _
              "AND   d.red_codigo = c.rec_codigo AND d.red_codpro = b.cpi_coding AND b.cpi_cencos = '" & MuestraCasino(1) & "' AND b.cpi_codped = a.pro_codigo " & _
              "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) " & _
              "AND   e.min_cencos = '" & Trim(fpText.text) & "' AND val(mid(e.min_fecmin,1,6)) = " & Fecha & " " & _
              "AND   f.mid_tipmin = '1' AND a.pro_facing > 0 GROUP BY e.min_fecmin, b.cpi_coding, a.pro_codigo, a.pro_codtip", vg_db, adOpenStatic
   Else
      RS.Open "SELECT e.min_fecmin, b.cpi_coding AS ing_codigo, a.pro_codigo, a.pro_codtip, " & _
              "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))) AS cantidad1, " & _
              "SUM((f.mid_numrac*(d.red_canpro/c.rec_basrac))/a.pro_facing) AS cantidad2, 0 AS ped_fecped INTO " & aAp & " " & _
              "FROM b_productos a, b_contlistpreing b, b_receta c, b_recetadet d, b_minuta e, b_minutadet f " & _
              "WHERE e.min_codigo = f.mid_codigo AND f.mid_codrec = d.red_codigo AND f.mid_tiprec = d.red_tiprec AND ((d.red_tiprec <> 0 AND d.red_cencos = '" & MuestraCasino(1) & "') OR (d.red_tiprec = 0 AND d.red_cencos = '0')) " & _
              "AND   d.red_codigo = c.rec_codigo AND d.red_codpro = b.cpi_coding AND b.cpi_cencos = '" & MuestraCasino(1) & "' AND b.cpi_codped = a.pro_codigo " & _
              "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) " & _
              "AND   e.min_cencos = '" & Trim(fpText.text) & "' AND convert(int,substring(convert(varchar(8),e.min_fecmin),1,6)) = " & Fecha & " " & _
              "AND   f.mid_tipmin = '1' AND a.pro_facing > 0 GROUP BY e.min_fecmin, b.cpi_coding, a.pro_codigo, a.pro_codtip", vg_db, adOpenStatic
   End If
   Set RS = Nothing
   diaini = 0: diafin = 0
   diaini = 1: diafin = Mid(dEoM("01/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)), 1, 2)
   '-------> Rutina buscar estructura fija
   RS.Open "SELECT DISTINCT mif_codreg, mif_codser FROM  b_minutafija WHERE mif_cencos = '" & Trim(fpText.text) & "'", vg_db, adOpenStatic
   If Not RS.EOF Then
      Do While Not RS.EOF
         estfij = False
         '-------> Buscar datos estructura fija día
         If vg_tipbase = "1" Then
            RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
                     "WHERE mfd_cencos = '" & Trim(fpText.text) & "' " & _
                     "AND   mfd_codreg = " & RS!mif_codreg & " " & _
                     "AND   mfd_codser = " & RS!mif_codser & " " & _
                     "AND mid(mfd_fecha,1,6) >= " & Fecha & " AND mid(mfd_fecha,1,6) <= " & Fecha & " AND mfd_tipmin = '1'", vg_db, adOpenStatic
         Else
            RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
                     "WHERE mfd_cencos = '" & Trim(fpText.text) & "' " & _
                     "AND   mfd_codreg = " & RS!mif_codreg & " " & _
                     "AND   mfd_codser = " & RS!mif_codser & " " & _
                     "AND   convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) >= " & Fecha & " AND convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) <= " & Fecha & " AND mfd_tipmin = '1'", vg_db, adOpenStatic
         End If
         If Not RS1.EOF Then estfij = True
         RS1.Close: Set RS1 = Nothing
         fecval = 0
         If Not estfij Then
            '-------> Buscar fecha mayor de estructura fija
            RS1.Open "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija " & _
                     "WHERE mif_cencos = '" & Trim(fpText.text) & "' " & _
                     "AND   mif_codreg = " & RS!mif_codreg & " " & _
                     "AND   mif_codser = " & RS!mif_codser & "", vg_db, adOpenStatic
            If Not RS1.EOF Then fecval = IIf(IsNull(RS1!fecval), 0, RS1!fecval)
            RS1.Close: Set RS1 = Nothing
            If fecval > 0 Then
               '-------> Traer estructura fija
               For i = diaini To diafin
                   If fecval <= Fecha & Right("0" & i, 2) Then
                      vg_db.Execute "INSERT INTO " & aAp & " SELECT " & Fecha & Right("0" & i, 2) & " AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                                    "0 AS cantidad1, a.mif_canpro AS cantidad2, 0 AS ped_fecped " & _
                                    "FROM  b_minutafija a, b_productos b, b_ingrediente c, b_productosing d " & _
                                    "WHERE a.mif_codpro = b.pro_codigo " & _
                                    "AND   b.pro_codigo = d.pri_codpro " & _
                                    "AND   c.ing_codigo = d.pri_coding " & _
                                    "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                    "AND   a.mif_cencos = '" & Trim(fpText.text) & "' " & _
                                    "AND   a.mif_codreg = " & RS!mif_codreg & " " & _
                                    "AND   a.mif_codser = " & RS!mif_codser & " " & _
                                    "AND   a.mif_fecval = " & fecval & " " & _
                                    "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                    "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
                    
                      vg_db.Execute "INSERT INTO " & aAp & " SELECT " & Fecha & Right("0" & i, 2) & " AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                                    "0 AS cantidad1, a.mif_canpro AS cantidad2, 0 AS ped_fecped " & _
                                    "FROM  b_minutafija a, b_productos b, b_ingrediente c, b_productosing d " & _
                                    "WHERE a.mif_codpro = b.pro_codigo " & _
                                    "AND   b.pro_codigo = d.pri_codpro " & _
                                    "AND   c.ing_codigo = d.pri_coding " & _
                                    "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                    "AND   a.mif_cencos = '" & Trim(fpText.text) & "' " & _
                                    "AND   a.mif_codreg = " & RS!mif_codreg & " " & _
                                    "AND   a.mif_codser = " & RS!mif_codser & " " & _
                                    "AND   a.mif_fecval = " & fecval & " " & _
                                    "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                    "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
                   End If
                   Set RS1 = Nothing
               Next i
            End If
         ElseIf estfij Then
             '-------> Calcular datos desde tabla estructura fija día
             If vg_tipbase = "1" Then
                vg_db.Execute "INSERT INTO " & aAp & " SELECT a.mfd_fecha AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                              "0 AS cantidad1, a.mfd_canpro AS cantidad2, 0 AS ped_fecped " & _
                              "FROM b_minutafijadia a, b_productos b, b_ingrediente c, b_productosing d " & _
                              "WHERE a.mfd_codpro = b.pro_codigo " & _
                              "AND   b.pro_codigo = d.pri_codpro " & _
                              "AND   c.ing_codigo = d.pri_coding " & _
                              "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven<=0) " & _
                              "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                              "AND   a.mfd_codreg = " & RS!mif_codreg & " " & _
                              "AND   a.mfd_codser = " & RS!mif_codser & " " & _
                              "AND mid(a.mfd_fecha,1,6) = " & Fecha & " AND a.mfd_tipmin = '1' " & _
                              "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
             
                vg_db.Execute "INSERT INTO " & aAp & " SELECT a.mfd_fecha AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                              "0 AS cantidad1, a.mfd_canpro AS cantidad2, 0 AS ped_fecped " & _
                              "FROM b_minutafijadia a, b_productos b, b_ingrediente c, b_productosing d " & _
                              "WHERE a.mfd_codpro = b.pro_codigo " & _
                              "AND   b.pro_codigo = d.pri_codpro " & _
                              "AND   c.ing_codigo = d.pri_coding " & _
                              "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                              "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                              "AND   a.mfd_codreg = " & RS!mif_codreg & " " & _
                              "AND   a.mfd_codser = " & RS!mif_codser & " " & _
                              "AND mid(a.mfd_fecha,1,6) = " & Fecha & " AND a.mfd_tipmin='1' " & _
                              "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
             Else
                vg_db.Execute "INSERT INTO " & aAp & " SELECT a.mfd_fecha AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                              "0 AS cantidad1, a.mfd_canpro AS cantidad2, 0 AS ped_fecped " & _
                              "FROM b_minutafijadia a, b_productos b, b_ingrediente c, b_productosing d " & _
                              "WHERE a.mfd_codpro = b.pro_codigo " & _
                              "AND   b.pro_codigo = d.pri_codpro " & _
                              "AND   c.ing_codigo = d.pri_coding " & _
                              "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                              "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                              "AND   a.mfd_codreg = " & RS!mif_codreg & " " & _
                              "AND   a.mfd_codser = " & RS!mif_codser & " " & _
                              "AND convert(int,substring(convert(varchar(8),a.mfd_fecha),1,6)) = " & Fecha & " AND a.mfd_tipmin = '1' " & _
                              "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')"
             
                vg_db.Execute "INSERT INTO " & aAp & " SELECT a.mfd_fecha AS min_fecmin, c.ing_codigo, b.pro_codigo, b.pro_codtip, " & _
                              "0 AS cantidad1, a.mfd_canpro AS cantidad2, 0 AS ped_fecped " & _
                              "FROM b_minutafijadia a, b_productos b, b_ingrediente c, b_productosing d " & _
                              "WHERE a.mfd_codpro = b.pro_codigo " & _
                              "AND   b.pro_codigo = d.pri_codpro " & _
                              "AND   c.ing_codigo = d.pri_coding " & _
                              "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                              "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                              "AND   a.mfd_codreg = " & RS!mif_codreg & " " & _
                              "AND   a.mfd_codser = " & RS!mif_codser & " " & _
                              "AND convert(int,substring(convert(varchar(8),a.mfd_fecha),1,6)) = " & Fecha & " AND a.mfd_tipmin = '1' " & _
                              "AND   b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
             End If
         End If
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
   '-------> Mover fecha pedido
   If vg_tipbase = "1" Then
      vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                    "SET a.ped_fecped=IIf(c.pad_tipo='S',IIf(Mid(a.min_fecmin,7,2)>=1 And Mid(a.min_fecmin,7,2)<=7,Mid(a.min_fecmin,1,6) & '01',IIf(Mid(a.min_fecmin,7,2)>=8 And Mid(a.min_fecmin,7,2)<=14,Mid(a.min_fecmin,1,6) & '08',IIf(Mid(a.min_fecmin,7,2)>=15 And Mid(a.min_fecmin,7,2)<=21,Mid(a.min_fecmin,1,6) & '15',Mid(a.min_fecmin,1,6) & '22'))),IIf(c.pad_tipo='M',Mid(a.min_fecmin,1,6) & '01',IIf(c.pad_tipo='Q',IIf(Val(Mid(a.min_fecmin,7,2))>15,Mid(a.min_fecmin,1,6) & '16',Mid(a.min_fecmin,1,6) & '01'),IIf(c.pad_tipo='D',IIf(Mid(a.min_fecmin,7,2)>=1 And Mid(a.min_fecmin,7,2)<=10,Mid(a.min_fecmin,1,6) & '01',IIf(Mid(a.min_fecmin,7,2)>=11 And Mid(a.min_fecmin,7,2)<=20,Mid(a.min_fecmin,1,6) & '11',Mid(a.min_fecmin,1,6) & '21')),0)))) WHERE c.pad_cencos='" & LimpiaDato(Trim(fpText.text)) & "'"
   Else
      '-------> actualizar fecha pedido semanal
      vg_db.Execute "UPDATE " & aAp & " " & _
                    "SET " & aAp & ".ped_fecped = CASE WHEN convert(int,substring(convert(varchar(8),a.min_fecmin),7,2)) >= 1 And convert(int,substring(convert(varchar(8),a.min_fecmin),7,2)) <= 7 THEN substring(convert(varchar(8),a.min_fecmin),1,6) + '01' WHEN convert(int,substring(convert(varchar(8),a.min_fecmin),7,2)) >= 8 And convert(int,substring(convert(varchar(8),a.min_fecmin),7,2)) <= 14 THEN substring(convert(varchar(8),a.min_fecmin),1,6) + '08' WHEN convert(int,substring(convert(varchar(8),a.min_fecmin),7,2)) >= 15 And convert(int,substring(convert(varchar(8),a.min_fecmin),7,2)) <= 21 THEN substring(convert(varchar(8),a.min_fecmin),1,6) + '15' ELSE substring(convert(varchar(8),a.min_fecmin),1,6) + '22' END " & _
                    "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                    "WHERE a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'S'"
      '-------> actualizar fecha pedido mensual
      vg_db.Execute "UPDATE " & aAp & "  " & _
                    "SET " & aAp & ".ped_fecped = substring(convert(varchar(8),a.min_fecmin),1,6) + '01' " & _
                    "FROM " & aAp & "  a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                    "WHERE a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'M'"
      '-------> actualizar fecha pedido quincenal
      vg_db.Execute "UPDATE " & aAp & " " & _
                    "SET " & aAp & ".ped_fecped = CASE WHEN convert(int,substring(convert(varchar(8),a.min_fecmin),7,2)) > 15 THEN substring(convert(varchar(8),a.min_fecmin),1,6) + '16' ELSE substring(convert(varchar(8),a.min_fecmin),1,6) + '01' END " & _
                    "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                    "WHERE a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'Q'"
   
      vg_db.Execute "UPDATE " & aAp & " " & _
                    "SET " & aAp & ".ped_fecped = CASE WHEN convert(int,substring(convert(varchar(8),a.min_fecmin),7,2)) >= 1 And convert(int,substring(convert(varchar(8),a.min_fecmin),7,2)) <= 10 THEN substring(convert(varchar(8),a.min_fecmin),1,6) + '01' WHEN convert(int,substring(convert(varchar(8),a.min_fecmin),7,2)) >= 11 And convert(int,substring(convert(varchar(8),a.min_fecmin),7,2)) <= 20 THEN substring(convert(varchar(8),a.min_fecmin),1,6) + '11' ELSE substring(convert(varchar(8),a.min_fecmin),1,6) + '21' END " & _
                    "FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d " & _
                    "WHERE a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_tipo = 'D'"
   End If
   Dim fecini As Date, fecfin As Date, fecped As Date, EstPed As Integer, fecpin As Long, fecpfi As Long
   RS.Open "SELECT pad_codigo, pad_diario FROM b_paramdesp WHERE pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND pad_tipo = 'E'", vg_db, adOpenForwardOnly
   Do While Not RS.EOF
      fecini = "01/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)
      fecfin = "07/" & IIf(Mid(Fecha, 5, 2) = 12, "01/" & Mid(Fecha, 1, 4) + 1, Mid(Fecha, 5, 2) + 1 & "/" & Mid(Fecha, 1, 4)) 'dEoM("27/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))
      fecpin = 0: fecpfi = 0
      fecpin = Format(fecini, "yyyymmdd")
      Do While fecini <= fecfin
         '-------> Buscar fecha inicial y fecha final
         For j = 1 To 7
             If (DatePart("w", fecini, 2)) = Val(Mid(RS!pad_diario, j, 1)) Then
                If fecpin = 0 Then
                   fecpin = Format(fecini, "yyyymmdd")
                ElseIf fecpfi = 0 Then
                   fecpfi = Format(fecini, "yyyymmdd")
                End If
             End If
             If fecpin > 0 And fecpfi > 0 Then
                If vg_tipbase = "1" Then
                   vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip = b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo = d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo = c.pad_codigo " & _
                                 "SET a.ped_fecped = " & fecpin & " WHERE c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND pad_tipo = 'E' AND pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecpin & " AND a.min_fecmin <= " & fecpfi & ""
                Else
                   vg_db.Execute "UPDATE " & aAp & " SET " & aAp & ".ped_fecped = " & fecpin & " FROM " & aAp & " a, a_tipopro b, b_paramdesp c, " & aAp1 & " d WHERE a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo " & _
                                 "AND c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND pad_tipo = 'E' AND pad_codigo = " & RS!pad_codigo & " AND a.ped_fecped = 0 AND a.min_fecmin >= " & fecpin & " AND a.min_fecmin <= " & fecpfi & ""
                End If
                fecpin = fecpfi: fecpfi = 0
                Exit For
             End If
         Next j
         fecini = fecini + 1
      Loop
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
   
'   vg_db.Execute "UPDATE ((" & aAp & " AS a INNER JOIN a_tipopro AS b ON a.pro_codtip=b.tip_codigo) INNER JOIN " & aAp1 & " AS d ON b.tip_codigo=d.pro_codtip) INNER JOIN b_paramdesp AS c ON d.pro_previo=c.pad_codigo " & _
'                 "SET a.ped_fecped=IIf(c.pad_tipo='E',IIf(DatePart('w',cstr(mid(a.min_fecmin, 7, 2) + '/' + mid(a.min_fecmin, 5, 2) + '/' + mid(a.min_fecmin, 1, 4)),2) < c.pad_diario,c.pad_diario+1,a.min_fecmin),a.min_fecmin) WHERE c.pad_cencos='" & LimpiaDato(Trim(fpText.text)) & "'"
   '-------> Leer archivo temporales
    RS.Open "SELECT a.ped_fecped, b.ing_codigo, b.ing_nombre, " & _
            "e.unm_nomcor, c.pro_codigo, c.pro_nombre, " & _
            "c.pro_coduni, c.pro_facsto, c.pro_ctacon, h.pro_previo AS tip_previo, d.uni_nomcor, " & _
            "SUM(a.cantidad1) AS cantidad1, SUM(a.cantidad2) AS cantidad2, (SELECT DISTINCT cpi_precos FROM b_contlistpreing WHERE cpi_coding = b.ing_codigo AND cpi_cencos = '" & MuestraCasino(1) & "') AS cpi_precos " & _
            "FROM  " & aAp & " a, b_ingrediente b, b_productos c, a_unidad d, a_unidadmed e, a_tipopro f, b_paramdesp g, " & aAp1 & " h " & _
            "WHERE a.ing_codigo = b.ing_codigo " & _
            "AND   a.pro_codigo = c.pro_codigo " & _
            "AND   c.pro_codtip = f.tip_codigo " & _
            "AND   f.tip_codigo = h.pro_codtip AND h.pro_previo = g.pad_codigo AND g.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
            "AND   b.ing_unimed = e.unm_codigo " & _
            "AND   c.pro_coduni = d.uni_codigo " & _
            "AND  (c.pro_fecven > " & Format(Date, "yyyymmdd") & " OR c.pro_fecven <= 0) " & _
            "AND   c.pro_facsto > 0 " & _
            "GROUP BY a.ped_fecped, b.ing_codigo, b.ing_nombre, " & _
            "e.unm_nomcor, c.pro_codigo, c.pro_nombre, c.pro_coduni, " & _
            "c.pro_facsto, d.uni_nomcor, c.pro_ctacon, h.pro_previo ORDER BY c.pro_ctacon, h.pro_previo, b.ing_codigo, b.ing_nombre, c.pro_nombre, a.ped_fecped", vg_db, adOpenStatic
   If Not RS.EOF Then Toolbar1.Buttons(1).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Frame2.Enabled = True: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(8).Enabled = False Else Frame2.Enabled = False: Toolbar1.Buttons(1).Enabled = False
End If
Dim codpre As Long, despa As Long
inding = 0: nomunm = "": codpre = 0: despa = 0
vaSpread1.Visible = False
If Not RS.EOF Then
   Do While Not RS.EOF
      If RS!tip_previo <> codpre Then
         vaSpread1.MaxRows = vaSpread1.MaxRows + 1
         vaSpread1.Row = vaSpread1.MaxRows
         vaSpread1.Col = -1: vaSpread1.BackColor = &HFFFFC0
         vaSpread1.Col = 2: vaSpread1.Font.Bold = True: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = fg_BuscaenArbol(RS!tip_previo, "a_tipopro", "tip_codigo")
         vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
         vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
         vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
         vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
         vaSpread1.Col = 7: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
         vaSpread1.Col = 8: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
         vaSpread1.Col = 9: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
         vaSpread1.Col = 10: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = 0
         vaSpread1.Col = 12: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!tip_previo
         codpre = RS!tip_previo
      End If
      If RS!ing_codigo <> auxing Or (RS!ped_fecped <> despa And etapa5) Then
         If inding > 0 And etapa5 And auxing <> "720" Then
            vaSpread1.Row = inding
            vaSpread1.Col = 4: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = vaSpread1.text & " " & nomunm
            vaSpread1.Col = 5: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(reqtot, fg_Pict(6, 2))
            vaSpread1.Col = 6: vaSpread1.text = ustock
            vaSpread1.Col = 8: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(Stock, fg_Pict(6, 2))
            vaSpread1.Col = 9: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(proped, fg_Pict(6, 2))
            reqtot = 0: ustock = "": Stock = 0: proped = 0: nomunm = ""
         End If
         vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows: i = vaSpread1.Row: inding = vaSpread1.MaxRows
         vaSpread1.Col = -1: vaSpread1.Font.Bold = IIf(etapa5 = True, False, True): vaSpread1.Font.Size = 9: vaSpread1.BackColor = IIf(etapa5, &HC0FFFF, &HC0FFC0)
         vaSpread1.RowHidden = False
         If Check1(0).Value = 1 And (etapa5 And RS!ing_codigo = "720") Then
            vaSpread1.RowHidden = True
         ElseIf Check1(0).Value = 1 And Not etapa5 Then
            vaSpread1.RowHidden = True
         End If
         vaSpread1.Col = 2: vaSpread1.Font.Bold = IIf(etapa5 = True, False, True): vaSpread1.Font.Size = 9: vaSpread1.CellType = 5: vaSpread1.text = RS!ing_nombre
         vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!ing_codigo
         vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = IIf(etapa5, Mid(RS!ped_fecped, 7, 2) & "/" & Mid(RS!ped_fecped, 5, 2) & "/" & Mid(RS!ped_fecped, 1, 4), "")
         vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
         vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = IIf(etapa5, "", RS!unm_nomcor)
         vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
         vaSpread1.Col = 7: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignCenter: vaSpread1.text = IIf(IsNull(RS!cpi_precos), 0, RS!cpi_precos)
         vaSpread1.Col = 10: vaSpread1.text = 0
         If fecenv = 0 Or fecenv = 1 Then
            vaSpread1.Col = 4: vaSpread1.Font.Bold = IIf(etapa5 = True, False, True): vaSpread1.Font.Size = 9: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignRight
            If fecenv = 0 Then vaSpread1.text = Format(RS!ped_canmin, fg_Pict(6, 2)) Else vaSpread1.text = Format(RS!cantidad1, fg_Pict(6, 2))
         Else
            vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.Font.Bold = IIf(etapa5 = True, False, True): vaSpread1.Font.Size = 9: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(RS!ped_canmin, fg_Pict(6, 2))
         End If
         nomunm = Trim(RS!unm_nomcor)
         auxing = RS!ing_codigo
         despa = RS!ped_fecped
      Else
         vaSpread1.Row = i
         If fecenv = 0 Or fecenv = 1 Then
            vaSpread1.Col = 4: vaSpread1.Font.Bold = IIf(etapa5 = True, False, True): vaSpread1.Font.Size = 9: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignRight
            If fecenv = 0 Then vaSpread1.text = Format(RS!ped_canmin, fg_Pict(6, 2)) Else vaSpread1.text = IIf(vaSpread1.text <> "", Format((vaSpread1.text + RS!cantidad1), fg_Pict(6, 2)), Format(RS!cantidad1, fg_Pict(6, 2)))
         Else
            vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.Font.Bold = IIf(etapa5 = True, False, True): vaSpread1.Font.Size = 9: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(RS!ped_canmin, fg_Pict(6, 2))
         End If
      End If
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.RowHidden = IIf(etapa5 = True And RS!ing_codigo <> "720", True, False)
      vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!pro_codigo
      vaSpread1.Col = 2: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!pro_nombre
      vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = Mid(RS!ped_fecped, 7, 2) & "/" & Mid(RS!ped_fecped, 5, 2) & "/" & Mid(RS!ped_fecped, 1, 4)
      vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
      vaSpread1.Col = 10: vaSpread1.text = 1
      vaSpread1.Col = 11: vaSpread1.text = RS!pro_codigo
      If fecenv = 0 Or fecenv = 1 Then
         vaSpread1.Col = 5
         If etapa5 = False Then
            vaSpread1.CellType = CellTypeNumber
            vaSpread1.TypeNumberDecPlaces = 2
            vaSpread1.TypeNumberMin = 1
            vaSpread1.TypeNumberMax = 9999999
            vaSpread1.TypeHAlign = TypeHAlignRight
            vaSpread1.TypeSpin = False
            vaSpread1.TypeIntegerSpinInc = 1
            vaSpread1.TypeIntegerSpinWrap = False
         Else
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignRight
         End If
         If Not etapa5 Then vaSpread1.ForeColor = &HFF0000
         If fecenv = 0 Then
            vaSpread1.text = Format(RS!cantidad2, fg_Pict(6, 0))
            vaSpread1.Col = 8: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(RS!ped_stoact, fg_Pict(6, 2))
            vaSpread1.Col = 9: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(RS!ped_proped, fg_Pict(6, 2))
         Else
            If RS!ing_codigo <> "zzzfija" Then
               If Not IsNull(RS!cantidad2) Then
                  vaSpread1.text = IIf(Int(RS!cantidad2 / RS!pro_facsto) <> (RS!cantidad2 / RS!pro_facsto), Int(RS!cantidad2 / RS!pro_facsto) + 1, Round(RS!cantidad2 / RS!pro_facsto, 0)) * RS!pro_facsto
                  canped = IIf(Int(RS!cantidad2 / RS!pro_facsto) <> (RS!cantidad2 / RS!pro_facsto), Int(RS!cantidad2 / RS!pro_facsto) + 1, Round(RS!cantidad2 / RS!pro_facsto, 0)) * RS!pro_facsto
               End If
               If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone) <> -1 Then
                  vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone)
                  vaSpread1.Col = 8: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(IIf(vaSpread2.text > 0, vaSpread2.text, 0), fg_Pict(6, 2))
                  vaSpread2.Col = 2: vaSpread1.Col = 9: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(IIf(Val(vaSpread2.text) > canped, 0, (canped - vaSpread2.text)), fg_Pict(6, 2))
                  vaSpread2.Col = 2: vaSpread2.text = IIf((vaSpread2.text - canped) <= 0, 0, (vaSpread2.text - canped))
               Else
                  vaSpread1.Col = 8: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(0, fg_Pict(6, 2))
                  vaSpread1.Col = 9: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(canped, fg_Pict(6, 2))
               End If
            Else
               vaSpread1.text = RS!cantidad2
               canped = RS!cantidad2
               If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone) <> -1 Then
                  vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS!pro_codigo), SearchFlagsNone)
                  vaSpread2.Col = 2: vaSpread1.Col = 8: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(IIf(vaSpread2.text > 0, vaSpread2.text, 0), fg_Pict(6, 2))
                  vaSpread2.Col = 2: vaSpread1.Col = 9: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(IIf(Val(vaSpread2.text) > canped, 0, (canped - vaSpread2.text)), fg_Pict(6, 2))
               Else
                  vaSpread1.Col = 8: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(0, fg_Pict(6, 2))
                  vaSpread1.Col = 9: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(0, fg_Pict(6, 2))
               End If
            End If
         End If
      Else
         vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(RS!cantidad2, fg_Pict(6, 2))
         vaSpread1.Col = 8: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(RS!ped_stoact, fg_Pict(6, 2))
         vaSpread1.Col = 9: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(RS!ped_proped, fg_Pict(6, 2))
      End If
      vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = RS!uni_nomcor '& " x " & RS!pro_facsto
      vaSpread1.Col = 7: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = 0
      vaSpread1.Col = 5: reqtot = reqtot + IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
      vaSpread1.Col = 6: ustock = vaSpread1.text
      vaSpread1.Col = 8: Stock = Stock + vaSpread1.text
      vaSpread1.Col = 9: proped = proped + vaSpread1.text
      RS.MoveNext
   Loop
   If inding > 0 And etapa5 And auxing <> "720" Then
      vaSpread1.Row = inding
      vaSpread1.Col = 4: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = vaSpread1.text & " " & nomunm
      vaSpread1.Col = 5: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(reqtot, fg_Pict(6, 2))
      vaSpread1.Col = 6: vaSpread1.text = ustock
      vaSpread1.Col = 8: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(Stock, fg_Pict(6, 2))
      vaSpread1.Col = 9: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(proped, fg_Pict(6, 2))
   End If
End If
RS.Close: Set RS = Nothing: fg_descarga
'-------> Borrar tablas temporales
If Trim(aAp) <> "" And vg_tipbase = "1" Then vg_db.Execute "DROP TABLE " & aAp & ""
If Trim(aAp1) <> "" And vg_tipbase = "1" Then vg_db.Execute "DROP TABLE " & aAp1 & ""
If Trim(aAp2) <> "" And vg_tipbase = "1" Then vg_db.Execute "DROP TABLE " & aAp2 & ""
vaSpread1.Visible = True
If Me.Visible Then vaSpread1.SetFocus
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If Not IsDate(fpDateTime1.text) Then Exit Sub
MoverDatos
End Sub

Private Sub fpText_Change()
If fpText.text = "" Then fpayuda.Caption = "": Exit Sub
Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(8).Enabled = False: vaSpread1.MaxRows = 0
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
Dim codpro As String, coding As String, i As Integer
Dim fechasis As Long, fecdes As Long
Dim canmin As Double, cospro As Double, cosali As Double, CosDes As Double, canped As Double, stoact As Double, proped As Double
Dim aAp As String
On Error GoTo Man_Error
Select Case Button.Index
Case 1 '-------> Grabar pedido
    RS.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText.text)) & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda.Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '-------> Validar si existen datos en planificación teórica
    If vg_tipbase = "1" Then
       RS.Open "SELECT DISTINCT min_cencos FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(min_fecmin,1,6)) = " & Fecha & "", vg_db, adOpenStatic
    Else
       RS.Open "SELECT DISTINCT min_cencos FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),min_fecmin),1,6)) = " & Fecha & "", vg_db, adOpenStatic
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    fg_carga ""
    fechasis = 0: fechasis = Val(Mid(Date, 7, 4)) & Mid(Date, 4, 2) & Mid(Date, 1, 2)
    '-------> Grabar tabla b_minutapedido
    Toolbar1.Enabled = False
    vg_db.BeginTrans
    canmin = 0: codpro = "": coding = "": canped = 0: stoact = 0: proped = 0: fecdes = 0
    '-------> Eliminar pedido
    vg_db.Execute "DELETE b_minutapedido FROM b_minutapedido WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " AND ped_tipped = 1"
    For i = 1 To vaSpread1.MaxRows
        DoEvents
        vaSpread1.Row = i: vaSpread1.Col = 10
        codpro = "": canped = 0: stoact = 0: proped = 0
        If vaSpread1.text = 1 And vaSpread1.BackColor <> &HFFFFC0 Then
           vaSpread1.Col = 1: codpro = vaSpread1.text
           vaSpread1.Col = 3: fecdes = Format(vaSpread1.text, "yyyymmdd")
           vaSpread1.Col = 5: canped = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
           vaSpread1.Col = 8: stoact = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
           vaSpread1.Col = 9: proped = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
           vg_db.Execute "INSERT INTO b_minutapedido (ped_codcas, ped_fecped, ped_anomes, ped_tipped, ped_coding, ped_codpro, ped_canmin, ped_canped, ped_fecenv, ped_stoact, ped_proped) " & _
           "VALUES ('" & fpText.text & "', " & fecdes & ", " & Fecha & ", 1, '" & coding & "', '" & codpro & "', " & canmin & ", " & canped & ", 0, " & stoact & ", " & proped & ")"
           canmin = 0
           '-------> Actualizar codigo pedido en ingrediente
           vg_db.Execute "UPDATE b_contlistpreing SET cpi_codped = '" & codpro & "' WHERE cpi_cencos = '" & MuestraCasino(1) & "' AND (cpi_codped = '' OR (cpi_codped) IS NULL)"
        ElseIf vaSpread1.BackColor <> &HFFFFC0 Then
           vaSpread1.Col = 1
           If Trim(vaSpread1.text) <> "zzzfija" Then
              vaSpread1.Col = 1: coding = Trim(vaSpread1.text)
              vaSpread1.Col = 4: canmin = 0
              If Trim(vaSpread1.text) <> "" And Val(vaSpread1.text) > 0 Then
                 If etapa5 And Not (Asc(Right(vaSpread1.text, 1)) >= 48 And Asc(Right(vaSpread1.text, 1)) <= 57) Then
                    canmin = Mid(vaSpread1.text, 1, InStr(vaSpread1.text, " ") - 1)
                 Else
                    canmin = vaSpread1.text
                 End If
              End If
           Else
              coding = "": canmin = 0
           End If
        End If
    Next i
    If vg_tipbase = "1" Then
       '-------> Insert tabla productospmpdia
       aAp = Trim(vg_NUsr) & "_tmp_ProductoPMPGenPed2"
       fg_CheckTmp aAp
       vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                     "INTO " & aAp & " " & _
                     "FROM b_productospmpdia " & _
                     "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                     "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                     "AND   ppd_propon>0 " & _
                     "GROUP BY ppd_cencos, ppd_codpro"
       vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
       vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon"
       vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
    End If
    '-------> Traer estructura fija
    Dim fecval As Long
    RS.Open "SELECT DISTINCT mif_cencos, mif_codreg, mif_codser FROM b_minutafija " & _
            "WHERE mif_cencos = '" & LimpiaDato(Trim(fpText.text)) & "'", vg_db, adOpenStatic
    If Not RS.EOF Then
       Do While Not RS.EOF
          DoEvents
          '-------> Validar si existe estructura fija día
          If vg_tipbase = "1" Then
             RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
                      "WHERE  mfd_cencos = '" & RS!mif_cencos & "' " & _
                      "AND    mfd_codreg = " & RS!mif_codreg & " " & _
                      "AND    mfd_codser = " & RS!mif_codser & " " & _
                      "AND mid(mfd_fecha,1,6) = " & Fecha & " " & _
                      "AND    mfd_tipmin = '1'", vg_db, adOpenStatic
          Else
             RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
                      "WHERE  mfd_cencos = '" & RS!mif_cencos & "' " & _
                      "AND    mfd_codreg = " & RS!mif_codreg & " " & _
                      "AND    mfd_codser = " & RS!mif_codser & " " & _
                      "AND    convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) = " & Fecha & " " & _
                      "AND    mfd_tipmin = '1'", vg_db, adOpenStatic
          End If
          If RS1.EOF Then
             RS1.Close: Set RS1 = Nothing
             '-------> Buscar fecha mayor de estructura fija
             fecval = 0
             RS1.Open "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija WHERE mif_cencos = '" & Trim(fpText.text) & "' AND mif_codreg = " & RS!mif_codreg & " AND mif_codser = " & RS!mif_codser & "", vg_db, adOpenStatic
             If Not RS1.EOF Then fecval = IIf(IsNull(RS1!fecval), 0, RS1!fecval)
             RS1.Close: Set RS1 = Nothing
             If fecval > 0 Then
                '-------> Traer estructura fija
                For i = 1 To Val(Mid(dEoM("26/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)), 1, 2))
                    If fecval <= Fecha & fg_pone_cero(Str(i), 2) Then
                       '-------> Grabar estructura fija día teorica
                       If vg_tipbase = "1" Then
                          vg_db.Execute "INSERT INTO b_minutafijadia (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro) SELECT a.mif_cencos, a.mif_codreg, a.mif_codser, " & Fecha & fg_pone_cero(i, 2) & ", b.pro_codigo, '1', a.mif_canpro, (SELECT DISTINCT ppd_propon FROM " & aAp & " WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") " & _
                                        "FROM b_minutafija a, b_productos b " & _
                                        "WHERE a.mif_codpro = b.pro_codigo " & _
                                        "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                        "AND   a.mif_cencos = '" & Trim(fpText.text) & "' " & _
                                        "AND   a.mif_codreg = " & RS!mif_codreg & " " & _
                                        "AND   a.mif_codser = " & RS!mif_codser & " " & _
                                        "AND   a.mif_fecval = " & fecval & " " & _
                                        "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                        "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0)"
                       Else
                          vg_db.Execute "INSERT INTO b_minutafijadia (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro) SELECT a.mif_cencos, a.mif_codreg, a.mif_codser, " & Fecha & fg_pone_cero(i, 2) & ", b.pro_codigo, '1', a.mif_canpro, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") " & _
                                        "FROM b_minutafija a, b_productos b " & _
                                        "WHERE a.mif_codpro = b.pro_codigo " & _
                                        "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                        "AND   a.mif_cencos = '" & Trim(fpText.text) & "' " & _
                                        "AND   a.mif_codreg = " & RS!mif_codreg & " " & _
                                        "AND   a.mif_codser = " & RS!mif_codser & " " & _
                                        "AND   a.mif_fecval = " & fecval & " " & _
                                        "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(Fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                        "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0)"
                       End If
                    End If
                    
                    Set RS1 = Nothing
                Next i
             End If
          Else
             RS1.Close: Set RS1 = Nothing
             '-------> Actualizar precio propon tabla estructura fija x día
'             vg_db.Execute "UPDATE b_minutafijadia INNER JOIN b_productospmpdia ON b_minutafijadia.mfd_codpro = b_productospmpdia.ppd_codpro SET b_minutafijadia.mfd_cospro = b_productospmpdia.ppd_propon " & _
'                           "WHERE b_minutafijadia.mfd_cencos = '" & Trim(fpText.text) & "' AND b_minutafijadia.mfd_codreg = " & RS!mif_codreg & " AND b_minutafijadia.mfd_codser = " & RS!mif_codser & " AND mid(b_minutafijadia.mfd_fecha,1,6) = " & Fecha & " AND b_minutafijadia.mfd_tipmin = '1' AND b_productospmpdia.ppd_cencos = '" & MuestraCasino(1) & "' AND b_productospmpdia.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ""
             If vg_tipbase = "1" Then
                vg_db.Execute "UPDATE b_minutafijadia INNER JOIN " & aAp & " ON b_minutafijadia.mfd_codpro = " & aAp & ".ppd_codpro SET b_minutafijadia.mfd_cospro = " & aAp & ".ppd_propon " & _
                              "WHERE b_minutafijadia.mfd_cencos = '" & Trim(fpText.text) & "' AND b_minutafijadia.mfd_codreg = " & RS!mif_codreg & " AND b_minutafijadia.mfd_codser = " & RS!mif_codser & " AND mid(b_minutafijadia.mfd_fecha,1,6) = " & Fecha & " AND b_minutafijadia.mfd_tipmin = '1' AND " & aAp & ".ppd_cencos = '" & MuestraCasino(1) & "'"
             Else
                vg_db.Execute "UPDATE b_minutafijadia SET b_minutafijadia.mfd_cospro = b_productospmpdia.ppd_propon FROM  b_productospmpdia WHERE b_minutafijadia.mfd_codpro = b_productospmpdia.ppd_codpro " & _
                              "AND b_minutafijadia.mfd_cencos = '" & Trim(fpText.text) & "' AND b_minutafijadia.mfd_codreg = " & RS!mif_codreg & " AND b_minutafijadia.mfd_codser = " & RS!mif_codser & " AND convert(int,substring(convert(varchar(8),b_minutafijadia.mfd_fecha),1,6)) = " & Fecha & " AND b_minutafijadia.mfd_tipmin = '1' AND b_productospmpdia.ppd_cencos = '" & MuestraCasino(1) & "'"
             End If
          End If
          RS.MoveNext
       Loop
    End If
    RS.Close: Set RS = Nothing
    '-------> Borrar tablas temporales
    If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
    vg_db.CommitTrans
    Toolbar1.Enabled = True
    Toolbar1.Buttons(3).Enabled = True: Toolbar1.Buttons(5).Visible = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(6).Visible = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(8).Enabled = True
    fg_descarga
Case 3 '-------> Generar pedido
    RS.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText.text)) & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda.Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '-------> Validar si existen datos en planificación teórica
    If vg_tipbase = "1" Then
       RS.Open "SELECT DISTINCT min_cencos FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND VAL(MID(min_fecmin,1,6)) = " & Fecha & "", vg_db, adOpenStatic
    Else
       RS.Open "SELECT DISTINCT min_cencos FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),min_fecmin),1,6)) = " & Fecha & "", vg_db, adOpenStatic
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    fecdes = 0: fechasis = 0: fechasis = Val(Mid(Date, 7, 4)) & Mid(Date, 4, 2) & Mid(Date, 1, 2)
    If MsgBox("¿ Esta seguro generar pedido ?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    '-------> Definir vector costo recetas
    Dim vecrec As Variant
    fg_carga ""
    Toolbar1.Enabled = False
    vg_db.BeginTrans
    If vg_tipbase = "1" Then
       '-------> Insert tabla productospmpdia
       aAp = Trim(vg_NUsr) & "_tmp_ProductoPMPGenPed3"
       fg_CheckTmp aAp
       vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                     "INTO " & aAp & " " & _
                     "FROM b_productospmpdia " & _
                     "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                     "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                     "AND   ppd_propon>0 " & _
                     "GROUP BY ppd_cencos, ppd_codpro"
       vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
       vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon"
       vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
    End If
    '-------> Actualizar fecha envio minuta pedido
    vg_db.Execute "UPDATE b_minutapedido SET ped_fecenv = " & fechasis & " WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " AND ped_tipped = 1"
    '-------> Grabar minuta costo teórico & real
    For i = 1 To vaSpread1.MaxRows
        DoEvents
        vaSpread1.Row = i: vaSpread1.Col = 1
        coding = Trim(vaSpread1.text): canped = 0: canmin = 0: cospro = 0: vaSpread1.Col = 10
        If vaSpread1.text = 0 And coding <> "zzzfija" And vaSpread1.BackColor <> &HFFFFC0 Then
           coding = "": fecdes = 0
           vaSpread1.Col = 1: coding = vaSpread1.text
           vaSpread1.Col = 7: cospro = vaSpread1.text
           RS.Open "SELECT mic_fecval FROM b_minutacosto WHERE mic_cencos = '" & MuestraCasino(1) & "' AND mic_fecval = " & fechasis & " AND mic_tipmin = '1' AND mic_codpro = '" & coding & "'", vg_db, adOpenStatic
           If RS.EOF Then
              vg_db.Execute "INSERT INTO b_minutacosto(mic_cencos, mic_fecval, mic_tipmin, mic_codpro, mic_cospro) VALUES ('" & MuestraCasino(1) & "', " & fechasis & ", '1', '" & coding & "', " & cospro & ")"
           Else
              vg_db.Execute "UPDATE b_minutacosto SET mic_cospro = " & cospro & " WHERE mic_cencos = '" & MuestraCasino(1) & "' AND mic_fecval = " & fechasis & " AND mic_tipmin = '1' AND mic_codpro = '" & coding & "'"
           End If
           RS.Close: Set RS = Nothing
           RS.Open "SELECT mic_fecval FROM b_minutacosto WHERE mic_cencos = '" & MuestraCasino(1) & "' AND mic_fecval = " & fechasis & " AND mic_tipmin = '2' AND mic_codpro = '" & coding & "'", vg_db, adOpenStatic
           If RS.EOF Then
              vg_db.Execute "INSERT INTO b_minutacosto(mic_cencos, mic_fecval, mic_tipmin, mic_codpro, mic_cospro) " & _
                            "VALUES ('" & MuestraCasino(1) & "', " & fechasis & ", '2', '" & coding & "', " & cospro & ")"
           Else
              vg_db.Execute "UPDATE b_minutacosto SET mic_cospro = " & cospro & " WHERE mic_cencos = '" & MuestraCasino(1) & "' AND mic_fecval = " & fechasis & " AND mic_tipmin = '2' AND mic_codpro = '" & coding & "'"
           End If
           RS.Close: Set RS = Nothing
        End If
    Next i
    '-------> Generar minuta costo estructura fija
    '-------> Eliminar estructura fija día real si existen datos
    If vg_tipbase = "1" Then
       vg_db.Execute "DELETE b_minutafijadia FROM b_minutafijadia WHERE mfd_cencos = '" & Trim(fpText.text) & "' AND mid(mfd_fecha,1,6) = " & Fecha & " AND mfd_tipmin = '2'"
       '-------> Grabar estructura fija día real
       vg_db.Execute "INSERT INTO b_minutafijadia (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro) SELECT a.mfd_cencos, a.mfd_codreg, a.mfd_codser, a.mfd_fecha, a.mfd_codpro, '2', a.mfd_canpro, (SELECT ppd_propon FROM " & aAp & " WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") " & _
                     "FROM b_minutafijadia a, b_productos b " & _
                     "WHERE a.mfd_codpro = b.pro_codigo " & _
                     "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                     "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                     "AND mid(a.mfd_fecha,1,6) = " & Fecha & " " & _
                     "AND   a.mfd_tipmin = '1' " & _
                     "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0)"
    Else
       vg_db.Execute "DELETE b_minutafijadia FROM b_minutafijadia WHERE mfd_cencos = '" & Trim(fpText.text) & "' AND convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) = " & Fecha & " AND mfd_tipmin = '2'"
       '-------> Grabar estructura fija día real
       vg_db.Execute "INSERT INTO b_minutafijadia (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro) SELECT a.mfd_cencos, a.mfd_codreg, a.mfd_codser, a.mfd_fecha, a.mfd_codpro, '2', a.mfd_canpro, (SELECT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = b.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") " & _
                     "FROM b_minutafijadia a, b_productos b " & _
                     "WHERE a.mfd_codpro = b.pro_codigo " & _
                     "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                     "AND   a.mfd_cencos = '" & Trim(fpText.text) & "' " & _
                     "AND   convert(int,substring(convert(varchar(8),a.mfd_fecha),1,6)) = " & Fecha & " " & _
                     "AND   a.mfd_tipmin = '1' " & _
                     "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0)"
    End If
    '-------> Traer total de receta desde planificación de minutas y luego calcular costo
    If vg_tipbase = "1" Then
       RS.Open "SELECT COUNT(b.mid_codrec) AS nreg FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
               "AND val(mid(a.min_fecmin,1,6)) = " & Fecha & " AND b.mid_tipmin = '1'", vg_db, adOpenStatic
    Else
       RS.Open "SELECT COUNT(b.mid_codrec) AS nreg FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
               "AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & Fecha & " AND b.mid_tipmin = '1'", vg_db, adOpenStatic
    End If
    If RS.EOF Or RS!nreg < 1 Then RS.Close: Set RS = Nothing: vg_db.RollbackTrans: MsgBox "No existe información", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
    ReDim vecrec(RS!nreg, 4)
    RS.Close: Set RS = Nothing
    For i = 1 To UBound(vecrec)
        DoEvents
        vecrec(i, 1) = 0 '-------> codigo receta
        vecrec(i, 2) = 0 '-------> tipo receta
        vecrec(i, 3) = 0 '-------> costo receta alimentación
        vecrec(i, 4) = 0 '-------> costo receta desechable
    Next i
    i = 1
    If vg_tipbase = "1" Then
       RS.Open "SELECT DISTINCT b.mid_codrec, b.mid_tiprec FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
               "AND val(mid(a.min_fecmin,1,6)) = " & Fecha & " AND b.mid_tipmin = '1'", vg_db, adOpenStatic
    Else
       RS.Open "SELECT DISTINCT b.mid_codrec, b.mid_tiprec FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
               "AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & Fecha & " AND b.mid_tipmin = '1'", vg_db, adOpenStatic
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: vg_db.RollbackTrans: MsgBox "No existe información", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
    Do While Not RS.EOF
       DoEvents
       vecrec(i, 1) = RS!mid_codrec
       vecrec(i, 2) = RS!mid_tiprec
       vecrec(i, 3) = Format(fg_CalCtoRecPlan(fechasis, 1, RS!mid_codrec, IIf(IsNull(RS!mid_tiprec), 0, RS!mid_tiprec), (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','"))), fg_Pict(6, 2))
       vecrec(i, 4) = Format(fg_CalCtoRecPlan(fechasis, 1, RS!mid_codrec, IIf(IsNull(RS!mid_tiprec), 0, RS!mid_tiprec), (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))), fg_Pict(6, 2))
       RS.MoveNext: i = i + 1
    Loop
    RS.Close: Set RS = Nothing
    
    '-------> Generar planificación real & actualizar costo teórica
    If vg_tipbase = "1" Then
       RS.Open "SELECT b.* FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
               "AND val(mid(a.min_fecmin,1,6)) = " & Fecha & " AND b.mid_tipmin = '1'", vg_db, adOpenStatic
    Else
       RS.Open "SELECT b.* FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
               "AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & Fecha & " AND b.mid_tipmin = '1'", vg_db, adOpenStatic
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: vg_db.RollbackTrans: MsgBox "No existe información", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
    Do While Not RS.EOF
       DoEvents
       For i = 1 To UBound(vecrec)
           If RS!mid_codrec = vecrec(i, 1) And RS!mid_tiprec = vecrec(i, 2) Then
              cosali = vecrec(i, 3)
              CosDes = vecrec(i, 4)
              Exit For
           End If
       Next
'       cosali = Format(fg_CalCtoRecPlan(fechasis, 1, RS!mid_codrec, IIf(IsNull(RS!mid_tiprec), 0, RS!mid_tiprec), (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','"))), fg_Pict(6, 2))
'       cosdes = Format(fg_CalCtoRecPlan(fechasis, 1, RS!mid_codrec, IIf(IsNull(RS!mid_tiprec), 0, RS!mid_tiprec), (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))), fg_Pict(6, 2))
       vg_db.Execute "INSERT INTO b_minutadet (mid_codigo, mid_tipmin, mid_numlin, mid_estser, mid_codrec, mid_numrac, mid_descri, mid_cosrec, mid_fecval, mid_tiprec, mid_nummer, mid_rec5eta, mid_cosdes) " & _
                     "VALUES (" & RS!mid_codigo & ", '2', " & RS!mid_numlin & ", " & RS!mid_estser & ", " & RS!mid_codrec & ", " & IIf(IsNull(RS!mid_numrac), "NULL", RS!mid_numrac) & ", '" & RS!mid_descri & "', " & cosali & ", " & fechasis & ", " & RS!mid_tiprec & ", 0, " & IIf(IsNull(RS!mid_rec5eta) Or Trim(RS!mid_rec5eta) = "", "Null", RS!mid_rec5eta) & ", " & CosDes & ")"
       vg_db.Execute "UPDATE b_minutadet SET mid_fecval = " & fechasis & ", mid_cosrec = " & cosali & ", mid_cosdes = " & CosDes & " WHERE mid_codigo = " & RS!mid_codigo & " AND mid_tipmin = '1' AND mid_codrec = " & RS!mid_codrec & " AND mid_tiprec = " & RS!mid_tiprec & ""
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    '-------> Bloquear planificación teórica
    If vg_tipbase = "1" Then
       vg_db.Execute "UPDATE b_minuta SET min_indblo = 1, min_racrea = min_racteo WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(min_fecmin,1,6)) = " & Fecha & " AND (min_indblo = 0 OR (min_indblo) IS NULL)"
       RS.Open "SELECT DISTINCT min_codreg, min_codser, min_fecmin, min_racrea FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(min_fecmin,1,6)) = " & Fecha & " AND min_racrea > 0", vg_db, adOpenStatic
    Else
       vg_db.Execute "UPDATE b_minuta SET min_indblo = 1, min_racrea = min_racteo WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),min_fecmin),1,6)) = " & Fecha & " AND (min_indblo = 0 OR (min_indblo) IS NULL)"
       RS.Open "SELECT DISTINCT min_codreg, min_codser, min_fecmin, min_racrea FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),min_fecmin),1,6)) = " & Fecha & " AND min_racrea > 0", vg_db, adOpenStatic
    End If
    '-------> Grabar raciones en minutas raciones
    Do While Not RS.EOF
       DoEvents
       RS1.Open "SELECT * FROM b_minutaraciones WHERE mir_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND mir_codreg = " & RS!min_codreg & " AND mir_codser = " & RS!min_codser & " AND mir_fecmin = " & RS!min_fecmin & " AND mir_rutcli = 'PRODUCIDAS'", vg_db, adOpenStatic
       If RS1.EOF Then
          vg_db.Execute "INSERT INTO b_minutaraciones VALUES ('" & LimpiaDato(Trim(fpText.text)) & "', " & RS!min_codreg & ", " & RS!min_codser & ", " & RS!min_fecmin & ", 'PRODUCIDAS', " & RS!min_racrea & ", NULL, '')"
       Else
          vg_db.Execute "UPDATE b_minutaraciones SET mir_nrorac = " & RS!min_racrea & " WHERE mir_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND mir_codreg = " & RS!min_codreg & " AND mir_codser = " & RS!min_codser & " AND mir_fecmin = " & RS!min_fecmin & " AND mir_rutcli = 'PRODUCIDAS' AND mir_nrorac < 1"
       End If
       RS1.Close: Set RS1 = Nothing
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    vg_db.CommitTrans
    If Trim(aAp) <> "" And vg_tipbase = "1" Then vg_db.Execute "DROP TABLE " & aAp & ""
    fg_descarga
    Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(8).Enabled = True: Frame2.Enabled = False
    MsgBox "Generación pedido Finalizado Sin Problema", vbInformation + vbOKOnly, Msgtitulo
    I_Pedidos LimpiaDato(Trim(fpText.text)), Fecha, IIf(Check1(0).Value = 1, IIf(etapa5, 0, 1), 0)
    Toolbar1.Enabled = True
Case 5 '-------> Borrar pedido
    If CierrePeriodo(Fecha, vg_codbod, 10) Then MsgBox "Existen documentos realizados, en la salida producción. Proceso cancelado", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina pedido...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    Toolbar1.Enabled = False
    vg_db.BeginTrans
    '-------> Eliminar minutapedido
    vg_db.Execute "DELETE b_minutapedido FROM b_minutapedido WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " AND ped_tipped = 1"
    If vg_tipbase = "1" Then
       '-------> Eliminar minutacosto
       vg_db.Execute "DELETE b_minutacosto FROM b_minutacosto WHERE mic_cencos = '" & MuestraCasino(1) & "' AND mic_fecval IN (SELECT a.mid_fecval FROM b_minutadet a, b_minuta b WHERE a.mid_codigo = b.min_codigo AND b.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(b.min_fecmin,1,6)) = " & Fecha & ")"
       '-------> Eliminar minutas real
       vg_db.Execute "DELETE b_minutadet FROM b_minutadet WHERE mid_codigo IN (SELECT min_codigo FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(min_fecmin,1,6)) = " & Fecha & ") AND mid_tipmin = '2'"
       '-------> Desbloquear planificación teórica
       vg_db.Execute "UPDATE b_minuta SET min_indblo = 0, min_racrea = 0 WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(min_fecmin,1,6)) = " & Fecha & " AND min_indblo = 1"
       '-------> Actualizar detalle planificación teórica al campo fecval
       vg_db.Execute "UPDATE b_minuta INNER JOIN b_minutadet ON b_minuta.min_codigo = b_minutadet.mid_codigo SET b_minutadet.mid_fecval = 0 " & _
                     "WHERE b_minutadet.mid_tipmin = '1' AND b_minuta.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND val(mid(b_minuta.min_fecmin,1,6)) = " & Fecha & ""
       '-------> Eliminar estructura fija día real
       vg_db.Execute "DELETE b_minutafijadia FROM b_minutafijadia WHERE mfd_cencos = '" & Trim(fpText.text) & "' AND mid(mfd_fecha,1,6) = " & Fecha & " AND mfd_tipmin = '2'"
    Else
       '-------> Eliminar minutacosto
       vg_db.Execute "DELETE b_minutacosto FROM b_minutacosto WHERE mic_cencos = '" & MuestraCasino(1) & "' AND mic_fecval IN (SELECT a.mid_fecval FROM b_minutadet a, b_minuta b WHERE a.mid_codigo = b.min_codigo AND b.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),b.min_fecmin),1,6)) = " & Fecha & ")"
       '-------> Eliminar minutas real
       vg_db.Execute "DELETE b_minutadet FROM b_minutadet WHERE mid_codigo IN (SELECT min_codigo FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),min_fecmin),1,6)) = " & Fecha & ") AND mid_tipmin = '2'"
       '-------> Desbloquear planificación teórica
       vg_db.Execute "UPDATE b_minuta SET min_indblo = 0, min_racrea = 0 WHERE min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),min_fecmin),1,6)) = " & Fecha & " AND min_indblo = 1"
       '-------> Actualizar detalle planificación teórica al campo fecval
       vg_db.Execute "UPDATE b_minutadet SET b_minutadet.mid_fecval = 0 FROM b_minutadet, b_minuta WHERE b_minuta.min_codigo = b_minutadet.mid_codigo " & _
                     "AND b_minutadet.mid_tipmin = '1' AND b_minuta.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND convert(int,substring(convert(varchar(8),b_minuta.min_fecmin),1,6)) = " & Fecha & ""
       '-------> Eliminar estructura fija día real
       vg_db.Execute "DELETE b_minutafijadia FROM b_minutafijadia WHERE mfd_cencos = '" & Trim(fpText.text) & "' AND convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) = " & Fecha & " AND mfd_tipmin = '2'"
    End If
    vaSpread1.MaxRows = 0
    Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = True
    vg_db.CommitTrans
    Toolbar1.Enabled = True
Case 8 '-------> Imprimir
    If vaSpread1.MaxRows < 1 Then Exit Sub
    RS.Open "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText.text)) & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda.Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '-------> Validar si existen datos en planificación teórica
    RS.Open "SELECT DISTINCT ped_codcas FROM b_minutapedido WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " AND ped_tipped = 1", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    I_Pedidos LimpiaDato(Trim(fpText.text)), Fecha, IIf(Check1(0).Value = 1, IIf(etapa5, 0, 1), 0)
Case 10 '-------> Cerrar
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
Toolbar1.Enabled = True
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    vg_nombre = "": vg_codigo = ""
    vg_left = fpayuda.Left + 2300
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Pst"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    RS1.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, c.pad_codigo, c.pad_tipo FROM b_productos a, a_tipopro b, b_paramdesp c, " & aAp1 & " d WHERE a.pro_codtip = b.tip_codigo AND b.tip_codigo = d.pro_codtip AND d.pro_previo = c.pad_codigo AND c.pad_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND a.pro_codigo = '" & vg_codigo & "' AND a.pro_facing > 0 AND a.pro_facsto > 0", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "Producto no tiene asignado los factores", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Dim Embalaje As String, codpro As String, coding As String, fecdes As String, j As Long, X As Integer, tipdes As String, fdespa1 As String, fdespa2 As String, fdespa3 As String, fdespa4 As String, codtip As String
    Dim fdespa5 As String, fdespa6 As String, fdespa7 As String
    codpro = vg_codigo: j = 0: X = 0: tipdes = Trim(RS1!pad_tipo): codtip = RS1!pad_codigo
    fdespa1 = "01/" & fpDateTime1.text: fdespa2 = "08/" & fpDateTime1.text: fdespa3 = "15/" & fpDateTime1.text: fdespa4 = "22/" & fpDateTime1.text
    fdespa5 = "01/" & fpDateTime1.text: fdespa6 = "11/" & fpDateTime1.text: fdespa7 = "21/" & fpDateTime1.text
    '-------> Validar si existe producto en grilla
    If vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(RS1!pro_codigo), SearchFlagsNone) <> -1 Then
       For i = vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(RS1!pro_codigo), SearchFlagsNone) To vaSpread1.MaxRows
           vaSpread1.Row = i: vaSpread1.Col = 6: Embalaje = "": Embalaje = Trim(vaSpread1.text): vaSpread1.Col = 3: fecdes = vaSpread1.text: vaSpread1.Col = 1
           If Trim(vaSpread1.text) = Trim(codpro) And Embalaje <> "" Then
              If fecdes = fdespa1 Then
                 fdespa1 = ""
              ElseIf fecdes = fdespa2 Then
                 fdespa2 = ""
              ElseIf fecdes = fdespa3 Then
                 fdespa3 = ""
              ElseIf fecdes = fdespa4 Then
                 fdespa4 = ""
              ElseIf fecdes = fdespa5 Then
                 fdespa5 = ""
              ElseIf fecdes = fdespa6 Then
                 fdespa6 = ""
              ElseIf fecdes = fdespa7 Then
                 fdespa7 = ""
              End If
              X = X + 1
           Else
              Exit For
           End If
       Next i
       If Trim(RS1!pad_tipo) = "M" And X = 1 Then vaSpread1.SetActiveCell 5, i - 1: RS1.Close: Set RS1 = Nothing: MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
       If Trim(RS1!pad_tipo) = "Q" And X = 2 Then vaSpread1.SetActiveCell 5, i - 1: RS1.Close: Set RS1 = Nothing: MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
       If Trim(RS1!pad_tipo) = "S" And X = 4 Then vaSpread1.SetActiveCell 5, i - 1: RS1.Close: Set RS1 = Nothing: MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
       If Trim(RS1!pad_tipo) = "D" And X = 3 Then vaSpread1.SetActiveCell 5, i - 1: RS1.Close: Set RS1 = Nothing: MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    End If
    vaSpread1.Row = vaSpread1.ActiveRow
    RS1.Close: Set RS1 = Nothing
    coding = "": proc2 = ""
    '-------> validar si existe mas de un ingrediente
    RS1.Open "SELECT COUNT(pri_coding) AS nreg FROM b_productosing WHERE pri_codpro = '" & codpro & "'", vg_db, adOpenStatic
    If RS1.EOF Or IsNull(RS1!nreg) Or RS1!nreg = 0 Then RS1.Close: Set RS1 = Nothing: Exit Sub
    If RS1!nreg > 1 Then
       vg_nombre = ""
       vg_left = fpayuda.Left + 2300
       B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "Proing"
       SendKeys "+{Tab}"
       B_TabEst.Show 1
       If vg_codigo = "" Then RS1.Close: Set RS1 = Nothing: Exit Sub
       coding = vg_codigo
       proc2 = " AND  (c.ing_codigo = '" & coding & "')"
    End If
    RS1.Close: Set RS1 = Nothing
    '-------> Validar si existe familia productos
    est = True
    If vaSpread1.SearchCol(12, 0, vaSpread1.MaxRows, codtip, SearchFlagsNone) <> -1 Then
       est = False
       j = vaSpread1.SearchCol(12, 0, vaSpread1.MaxRows, codtip, SearchFlagsNone)
    Else
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       vaSpread1.Row = vaSpread1.MaxRows
       vaSpread1.Col = -1: vaSpread1.BackColor = &HFFFFC0
       vaSpread1.Col = 2: vaSpread1.Font.Bold = True: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = fg_BuscaenArbol(Val(codtip), "a_tipopro", "tip_codigo")
       vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 7: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 8: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 9: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 10: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = 0
       vaSpread1.Col = 12: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = Val(codtip)
    End If
    proc1 = "SELECT a.pro_codigo, a.pro_nombre, a.pro_facsto, " & _
            "c.ing_codigo, (SELECT DISTINCT cpi_precos FROM b_contlistpreing WHERE cpi_coding = c.ing_codigo AND cpi_cencos = '" & MuestraCasino(1) & "') AS cpi_precos, c.ing_nombre, d.uni_nomcor " & _
            "FROM  b_productos a, b_productosing b, b_ingrediente c, a_unidad d " & _
            "WHERE a.pro_codigo = b.pri_codpro " & _
            "AND   c.ing_codigo = b.pri_coding " & _
            "AND   a.pro_coduni = d.uni_codigo " & _
            "AND   a.pro_codigo = '" & codpro & "'"
    RS1.Open proc1 & proc2, vg_db, adOpenStatic
    If Not RS1.EOF Then
       '-------> Validar si existe ingredientes
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i: vaSpread1.Col = 6: Embalaje = "": Embalaje = Trim(vaSpread1.text): vaSpread1.Col = 1
           If Trim(vaSpread1.text) = Trim(RS1!ing_codigo) And Embalaje = "" Then
              vaSpread1.MaxRows = vaSpread1.MaxRows + 1
              vaSpread1.InsertRows i + 1, 1
              vaSpread1.Row = i + 1
              vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS1!pro_codigo
              vaSpread1.Col = 2: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS1!pro_nombre
'              vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.Text = ""
              vaSpread1.Col = 3
              If tipdes = "M" Then
                 vaSpread1.CellType = CellTypeStaticText
                 vaSpread1.text = fdespa1
              ElseIf tipdes = "Q" Then
                 vaSpread1.CellType = CellTypeComboBox
                 If fdespa1 <> "" Then fdespa1 = fdespa1 & Chr$(9)
                 If fdespa3 <> "" Then fdespa1 = fdespa1 & fdespa3
                 vaSpread1.TypeComboBoxList = fdespa1
                 vaSpread1.TypeComboBoxCurSel = 0
                 If vaSpread1.TypeComboBoxCount = 0 Then vaSpread1.CellType = CellTypeStaticText
              ElseIf tipdes = "S" Then
                 vaSpread1.CellType = CellTypeComboBox
                 If fdespa1 <> "" Then fdespa1 = fdespa1 & Chr$(9)
                 If fdespa2 <> "" Then fdespa1 = fdespa1 & fdespa2 & Chr$(9)
                 If fdespa3 <> "" Then fdespa1 = fdespa1 & fdespa3 & Chr$(9)
                 If fdespa4 <> "" Then fdespa1 = fdespa1 & fdespa4
                 vaSpread1.TypeComboBoxList = fdespa1
                 vaSpread1.TypeComboBoxCurSel = 0
                 If vaSpread1.TypeComboBoxCount = 1 Then vaSpread1.CellType = CellTypeStaticText
              End If
              vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
              vaSpread1.Col = 5
              vaSpread1.CellType = CellTypeNumber
              vaSpread1.TypeNumberDecPlaces = 2
              vaSpread1.TypeNumberMin = 1
              vaSpread1.TypeNumberMax = 9999999
              vaSpread1.TypeHAlign = 1
              vaSpread1.TypeSpin = False
              vaSpread1.TypeIntegerSpinInc = 1
              vaSpread1.TypeIntegerSpinWrap = False
              vaSpread1.text = Format(0, fg_Pict(6, 0))
              vaSpread1.ForeColor = &HFF0000
              vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = 0: vaSpread1.text = RS1!uni_nomcor '& " x " & RS1!pro_facsto
              vaSpread1.Col = 7: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = 0: vaSpread1.text = 0
              vaSpread1.SetActiveCell 3, i + 1
              '-------> Revizar si existe producto stock
              If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS1!pro_codigo), SearchFlagsNone) <> -1 Then
                 vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS1!pro_codigo), SearchFlagsNone)
                 vaSpread2.Col = 2: vaSpread1.Col = 8: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(vaSpread2.text, fg_Pict(6, 2))
              Else
                 vaSpread1.Col = 8: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(0, fg_Pict(6, 2))
              End If
              vaSpread1.Col = 9: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(0, fg_Pict(6, 2))
              vaSpread1.Col = 10: vaSpread1.text = 1
              vaSpread1.Col = 11: vaSpread1.text = RS1!pro_codigo
              '-------> Fin revizar si existe producto stock
              RS1.Close: Set RS1 = Nothing
              Exit Sub
           End If
       Next i
       '-------> Mover si no existe ingrediente
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       If Not est Then
          vaSpread1.InsertRows j + 1, 1
          vaSpread1.Row = j + 1
       Else
          vaSpread1.Row = vaSpread1.MaxRows
       End If
       vaSpread1.Col = -1: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.BackColor = &HC0FFC0
       vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
       vaSpread1.Col = 2: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS1!ing_nombre
       vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS1!ing_codigo
       vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 7: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = IIf(IsNull(RS1!cpi_precos), 0, RS1!cpi_precos)
       vaSpread1.Col = 10: vaSpread1.text = 0
       '-------> Mover Productos
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       If Not est Then
          vaSpread1.InsertRows j + 2, 1
          vaSpread1.Row = j + 2
       Else
          vaSpread1.Row = vaSpread1.MaxRows
       End If
       vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS1!pro_codigo
       vaSpread1.Col = 2: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS1!pro_nombre
'       vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.Text = ""
       vaSpread1.Col = 3
       If tipdes = "M" Then
          vaSpread1.CellType = CellTypeStaticText
          vaSpread1.text = fdespa1
       ElseIf tipdes = "Q" Then
          vaSpread1.CellType = CellTypeComboBox
          If fdespa1 <> "" Then fdespa1 = fdespa1 & Chr$(9)
          If fdespa3 <> "" Then fdespa1 = fdespa1 & fdespa3
          vaSpread1.TypeComboBoxList = fdespa1
          vaSpread1.TypeComboBoxCurSel = 0
          If vaSpread1.TypeComboBoxCount = 0 Then vaSpread1.CellType = CellTypeStaticText
       ElseIf tipdes = "S" Then
          vaSpread1.CellType = CellTypeComboBox
          If fdespa1 <> "" Then fdespa1 = fdespa1 & Chr$(9)
          If fdespa2 <> "" Then fdespa1 = fdespa1 & fdespa2 & Chr$(9)
          If fdespa3 <> "" Then fdespa1 = fdespa1 & fdespa3 & Chr$(9)
          If fdespa4 <> "" Then fdespa1 = fdespa1 & fdespa4
          vaSpread1.TypeComboBoxList = fdespa1
          vaSpread1.TypeComboBoxCurSel = 0
          If vaSpread1.TypeComboBoxCount = 1 Then vaSpread1.CellType = CellTypeStaticText
       End If
       vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 5
       vaSpread1.CellType = CellTypeNumber
       vaSpread1.TypeNumberDecPlaces = 2
       vaSpread1.TypeNumberMin = 1
       vaSpread1.TypeNumberMax = 9999999
       vaSpread1.TypeHAlign = 1
       vaSpread1.TypeSpin = False
       vaSpread1.TypeIntegerSpinInc = 1
       vaSpread1.TypeIntegerSpinWrap = False
       vaSpread1.text = Format(0, fg_Pict(6, 0))
       vaSpread1.ForeColor = &HFF0000
       vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = RS1!uni_nomcor '& " x " & RS1!pro_facsto
       vaSpread1.Col = 7: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = 0
       vaSpread1.SetActiveCell 5, vaSpread1.MaxRows
       '-------> Revizar si existe producto stock
       If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS1!pro_codigo), SearchFlagsNone) <> -1 Then
          vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(RS1!pro_codigo), SearchFlagsNone)
          vaSpread2.Col = 2: vaSpread1.Col = 8: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(vaSpread2.text, fg_Pict(6, 2))
       Else
          vaSpread1.Col = 8: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(0, fg_Pict(6, 2))
       End If
       vaSpread1.Col = 9: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(0, fg_Pict(6, 2))
       vaSpread1.Col = 10: vaSpread1.text = 1
       vaSpread1.Col = 11: vaSpread1.text = RS1!pro_codigo
       If Not est Then vaSpread1.SetActiveCell 3, j + 2
       '-------> Fin revizar si existe producto stock
    End If
    RS1.Close: Set RS1 = Nothing
    est = False
Case 2
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    If vaSpread1.MaxRows = 0 Or vaSpread1.BackColor = &HFFFFC0 Then Exit Sub
    Dim vStock As Double
    vaSpread1.Col = 6
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    If Trim(vaSpread1.text) = "" Then
       vaSpread1.DeleteRows vaSpread1.Row, 1
       vaSpread1.MaxRows = vaSpread1.MaxRows - 1
       vaSpread1.Col = 11: codpro = vaSpread1.text
       If vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) <> -1 Then
          For j = vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) To vaSpread1.MaxRows
              vaSpread1.Row = j: vaSpread1.Col = 11
              If vaSpread1.text <> codpro Then
                 Exit For
              Else
                 vaSpread1.Col = 8
                 If vaSpread1.text > vStock Then vStock = vaSpread1.text
              End If
          Next j
       End If
       vaSpread1.Row = vaSpread1.ActiveRow
       For i = vaSpread1.Row To vaSpread1.MaxRows
           vaSpread1.Row = vaSpread1.Row: vaSpread1.Col = 6
           If Trim(vaSpread1.text) = "" Then Exit For
           vaSpread1.DeleteRows vaSpread1.Row, 1
           vaSpread1.MaxRows = vaSpread1.MaxRows - 1
       Next i
       If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(codpro), SearchFlagsNone) <> -1 Then
          vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, Trim(codpro), SearchFlagsNone)
          vaSpread2.Col = 2: vaSpread2.text = vStock
       End If
    Else
       vaSpread1.Col = 11: codpro = vaSpread1.text
       i = vaSpread1.Row
       If vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) <> -1 Then
          For j = vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) To vaSpread1.MaxRows
              vaSpread1.Row = j: vaSpread1.Col = 11
              If vaSpread1.text <> codpro Then
                 Exit For
              Else
                 vaSpread1.Col = 8
                 If vaSpread1.text > vStock Then vStock = vaSpread1.text
              End If
          Next j
       End If
       vaSpread1.DeleteRows i, 1 'vaSpread1.Row, 1
       vaSpread1.MaxRows = vaSpread1.MaxRows - 1
       If vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) <> -1 Then
          For j = vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) To vaSpread1.MaxRows
              vaSpread1.Row = j: vaSpread1.Col = 11
              If vaSpread1.text = codpro Then
                 vaSpread1.Col = 5: canped = vaSpread1.text
                 vaSpread1.Col = 8: vaSpread1.text = Format(vStock, fg_Pict(6, 2))
                 canped = IIf(canped > vStock, (canped - vStock), IIf(Val(vStock) >= canped, 0, canped))
                 vaSpread1.Col = 9: vaSpread1.text = Format(canped, fg_Pict(6, 2))
                 vaSpread1.Col = 5: canped = vaSpread1.text
                 vStock = IIf((vStock - canped) <= 0, 0, (vStock - canped))
              Else
                 Exit For
              End If
          Next j
       End If
       If (vaSpread1.ActiveRow - 1) >= 0 Then
          vaSpread1.Row = i: vaSpread1.Col = 6
          If Trim(vaSpread1.text) <> "" Then Exit Sub
          vaSpread1.Row = IIf(vaSpread1.ActiveRow - 1 = 0, 1, (i - 1))
          vaSpread1.Col = 6
          If Trim(vaSpread1.text) = "" Then
             vaSpread1.DeleteRows (vaSpread1.Row), 1: vaSpread1.MaxRows = vaSpread1.MaxRows - 1
          End If
       End If
    End If
End Select
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If Not IsDate(fpDateTime1.text) Then Exit Sub
    MoverDatos
End Select
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.MaxRows < 1 Or Col <> 5 Then Exit Sub
If ChangeMade = True Then
   Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(8).Enabled = False
   '-------> Rebajar o aumentar proposición pedido
   Dim canped As Double, codpro As String, vStock As Double
   vaSpread1.Row = Row: vaSpread1.Col = Col
   vaSpread1.Col = 11: codpro = vaSpread1.text
   '-------> Fin rebajar o aumentar proposición pedido
   '-------> recalcular stock
   If vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) <> -1 Then
      vaSpread1.Row = vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone)
      vaSpread1.Col = 8: vStock = vaSpread1.text
      For i = vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) To vaSpread1.MaxRows
          vaSpread1.Row = i: vaSpread1.Col = 11
          If vaSpread1.text = codpro Then
             vaSpread1.Col = 5: canped = vaSpread1.text
             vaSpread1.Col = 8: vaSpread1.text = Format(vStock, fg_Pict(6, 2))
             canped = IIf(canped > vStock, (canped - vStock), IIf(Val(vStock) >= canped, 0, canped))
             vaSpread1.Col = 9: vaSpread1.text = Format(canped, fg_Pict(6, 2))
             vaSpread1.Col = 5: canped = vaSpread1.text
             vStock = IIf((vStock - canped) <= 0, 0, (vStock - canped))
          Else
              Exit For
          End If
      Next i
   End If
   vaSpread1.EditEnterAction = EditEnterActionDown
End If
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread1.MaxRows < 1 Or Col <> 5 Or NewRow = Row Or etapa5 Then Exit Sub
Dim canped As Double, codpro As String, vStock As Double
vStock = 0
vaSpread1.Row = Row
vaSpread1.Col = 3
If Trim(vaSpread1.text) = "" Then Exit Sub
vaSpread1.CellType = CellTypeStaticText
vaSpread1.EditEnterAction = EditEnterActionNone
vaSpread1.Col = 11: codpro = vaSpread1.text
If vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) <> -1 Then
   For i = vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) To vaSpread1.MaxRows
       vaSpread1.Row = i: vaSpread1.Col = 11
       If vaSpread1.text <> codpro Then
          Exit For
       Else
          vaSpread1.Col = 8
          If vaSpread1.text > vStock Then vStock = vaSpread1.text
       End If
   Next i
   vaSpread1.Row = vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone)
   vaSpread1.ColUserSortIndicator(3) = ColUserSortIndicatorAscending
   vaSpread1.SortKey(1) = 3: vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
   vaSpread1.Sort 1, vaSpread1.Row, vaSpread1.MaxCols, i, SortByRow
   If vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) <> -1 Then
      For i = vaSpread1.SearchCol(11, 0, vaSpread1.MaxRows, Trim(codpro), SearchFlagsNone) To vaSpread1.MaxRows
          vaSpread1.Row = i: vaSpread1.Col = 11
          If vaSpread1.text = codpro Then
             vaSpread1.Col = 5: canped = vaSpread1.text
             vaSpread1.Col = 8: vaSpread1.text = Format(vStock, fg_Pict(6, 2))
             canped = IIf(canped > vStock, (canped - vStock), IIf(Val(vStock) >= canped, 0, canped))
             vaSpread1.Col = 9: vaSpread1.text = Format(canped, fg_Pict(6, 2))
             vaSpread1.Col = 5: canped = vaSpread1.text
             vStock = IIf((vStock - canped) <= 0, 0, (vStock - canped))
          Else
              Exit For
          End If
      Next i
   End If
   vaSpread1.EditEnterAction = EditEnterActionDown
End If
End Sub
