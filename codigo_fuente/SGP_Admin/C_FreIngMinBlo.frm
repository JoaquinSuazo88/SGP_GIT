VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form C_FreIngMinBlo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frecuencia Ingrediente Minuta Bloque"
   ClientHeight    =   8700
   ClientLeft      =   1950
   ClientTop       =   1710
   ClientWidth     =   16230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   16230
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   1440
      TabIndex        =   10
      Top             =   120
      Width           =   12735
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9120
         TabIndex        =   25
         Top             =   1440
         Width           =   915
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   3195
         TabIndex        =   21
         Top             =   1440
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
         Left            =   7500
         TabIndex        =   22
         Top             =   1440
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
         Index           =   4
         Left            =   6195
         TabIndex        =   24
         Top             =   1530
         Width           =   1065
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
         Index           =   3
         Left            =   1920
         TabIndex        =   23
         Top             =   1530
         Width           =   1110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3240
         TabIndex        =   16
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3240
         TabIndex        =   15
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3240
         TabIndex        =   14
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label1 
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
         Left            =   1875
         TabIndex        =   13
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
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
         Left            =   1875
         TabIndex        =   12
         Top             =   645
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
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
         Left            =   1875
         TabIndex        =   11
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3285
         TabIndex        =   17
         Top             =   285
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3285
         TabIndex        =   18
         Top             =   645
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   3285
         TabIndex        =   19
         Top             =   1005
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   15255
      Begin VB.Frame Frame13 
         Height          =   435
         Left            =   4440
         TabIndex        =   3
         Top             =   5280
         Width           =   675
         Begin VB.TextBox TextCai1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   4
            Top             =   135
            Width           =   570
         End
      End
      Begin VB.Frame Frame12 
         Height          =   435
         Left            =   5250
         TabIndex        =   1
         Top             =   5280
         Width           =   3765
         Begin VB.TextBox TextCai1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   2
            Top             =   135
            Width           =   3660
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4980
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   15015
         _Version        =   393216
         _ExtentX        =   26485
         _ExtentY        =   8784
         _StockProps     =   64
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
         MaxCols         =   14
         MaxRows         =   18
         SpreadDesigner  =   "C_FreIngMinBlo.frx":0000
         VisibleCols     =   8
         VisibleRows     =   18
         ScrollBarTrack  =   3
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEFEDE&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Ingredientes Listados"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   11880
         TabIndex        =   9
         Top             =   5235
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEFEDE&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Costo Promedio Diario"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   11880
         TabIndex        =   8
         Top             =   5535
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   13920
         TabIndex        =   7
         Top             =   5235
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   11
         Left            =   13920
         TabIndex        =   6
         Top             =   5535
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8700
      Left            =   15600
      TabIndex        =   20
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   15346
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "C_FreIngMinBlo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------
' --------------- Formulario : C_FreIngMinBlo  26/12/2013        ----------
' --------------- Creado     : Jorge Paz Núñez                   ----------
' --------------- Fecha      : 26/12/2013                        ----------
'--------------------------------------------------------------------------

Dim RS1         As New ADODB.Recordset
Dim CecoAux        As String
Dim CodRegimenAux  As Long
Dim CodServicioAux As Long
Dim tfor        As String
Dim FechaFinalAux  As Long
Dim FechaInicioAux As Long

Private Sub CmdBuscar_Click()

On Error GoTo Man_Error

TextCai1(1).text = ""
TextCai1(2).text = ""
Label1(9).Caption = ""
Label1(11).Caption = ""

'-------> Validar fechas
'If Format(FpFecDesde, ("YYYYMMDD")) < FechaInicioAux Then
'
'   MsgBox "Fecha Desde no debe ser menor a la minuta bloque seleccionada...", vbExclamation + vbOKOnly, Msgtitulo
'   Exit Sub
'
'
'End If

'If Format(FpFecHasta, ("YYYYMMDD")) > FechaFinalAux Then
'
'   MsgBox "Fecha Desde no debe ser mayor a la minuta bloque seleccionada...", vbExclamation + vbOKOnly, Msgtitulo
'   Exit Sub
'
'
'End If

If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
   
   MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub

End If
    
If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
   
   MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub

End If

LlenarFrecIngMinBlo tfor, CecoAux, CodRegimenAux, CodServicioAux, Format(FpFecDesde.text, "yyyymmdd"), Format(FpFecHasta.text, "yyyymmdd")

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

fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Excel "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub LlenarFrecIngMinBlo_Inicio(tfor As String, Ceco As String, CodRegimen As Long, CodServicio As Long, FechaInicio As Long, FechaFinal As Long)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

tfor = tfor
CecoAux = Ceco
CodRegimenAux = CodRegimen
CodServicioAux = CodServicio
FechaInicioAux = FechaInicio
FechaFinalAux = FechaFinal

'-------> Rutina frecuencia de recetas
Me.Caption = tfor
MsgTitulo = tfor

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_cliente_V02 45, '" & Ceco & "', ''")
If Not RS.EOF Then fpayuda(0).Caption = RS!Cli_nombre
RS.Close
Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_RegimenBloque " & IIf(Val(CodRegimen) = 0, -1, Val(CodRegimen)) & "")
If Not RS.EOF Then fpayuda(1).Caption = RS!reg_nombre
RS.Close
Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_ServicioBloque " & IIf(Val(CodServicio) = 0, -1, Val(CodServicio)) & "")
If Not RS.EOF Then fpayuda(3).Caption = RS!ser_nombre
RS.Close
Set RS = Nothing

FpFecDesde.text = fg_Ctod1(FechaInicio)
FpFecHasta.text = fg_Ctod1(FechaFinal)

vaSpread1(0).MaxRows = 0

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub LlenarFrecIngMinBlo(tfor As String, Ceco As String, CodRegimen As Long, CodServicio As Long, FechaInicio As Long, FechaFinal As Long)

On Error GoTo Man_Error

Dim CodReceta     As Long
Dim IRow          As Long
Dim i             As Long
Dim condia        As Long
Dim auxest        As Long
Dim cosreceta     As Double
Dim canreceta     As Double
Dim totgralreceta As Double
Dim RS            As New ADODB.Recordset
Dim Sql           As String

fg_carga ""

Dim X As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread1(0).TextTip = 2
vaSpread1(0).TextTipDelay = 250
X = vaSpread1(0).SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread1(0).MaxRows = 0

DoEvents
Dim SeleccionOpt As Integer

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

SeleccionOpt = IIf(M_MinSR1.optPrecioConvenio.Value = True, 1, IIf(M_MinSR1.optPrecioGenerico = True, 2, 3))
Sql = ""
Sql = "'" & LimpiaDato(Trim(M_MinSR1.fpText.text)) & "'"
Sql = Sql & ", " & M_MinSR1.fpLongInteger1(0).Value & ", " & M_MinSR1.fpLongInteger1(1).Value & ", " & FechaInicio & ", " & FechaFinal

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_ValidarMinutaBloqueModificado_V01 " & Sql & "")

If RS.EOF Then
   
   fg_descarga
   MsgBox "Datos modificar no corresponde", vbExclamation + vbOKOnly, Me.Caption
   RS.Close
   Set RS = Nothing
   Exit Sub

End If

vg_IDBloque = RS!Id_Bloque

RS.Close
Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = ""
Sql = "" & vg_IDBloque & ", '" & LimpiaDato(Trim(M_MinSR1.fpText.text)) & "', " & M_MinSR1.fpLongInteger1(0).Value & ", " & M_MinSR1.fpLongInteger1(1).Value & ", " & FechaInicio & ", " & FechaFinal & ", '1', " & SeleccionOpt & ""
Set RS = vg_db.Execute("sgpadm_Sel_FrecuenciaIngMinutaBloque_V04 " & Sql & "")

If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub

'definir largo del vector
Dim preklu       As Double
Dim frecest      As Long
Dim indfin       As Long
Dim vecTipoPla() As Variant

ReDim Preserve vecTipoPla(5000, 5)

CodReceta = 0
cosreceta = 0
canreceta = 0
totgralreceta = 0
condia = 0
indini = 1
indfin = 0
IRow = 1
auxest = 0

Do While Not RS.EOF
   
   DoEvents
   vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
   vaSpread1(0).Row = vaSpread1(0).MaxRows
    
   If auxest <> RS!pas_estser Then
      
      If auxest <> 0 Then
         
         vaSpread1(0).Col = 1 '-------> Glosa Días Planificados
         vaSpread1(0).CellType = CellTypeStaticText
         vaSpread1(0).Lock = True
         vaSpread1(0).TypeHAlign = TypeHAlignLeft
         vaSpread1(0).Font.Bold = True
         vaSpread1(0).Font.Size = 9
         vaSpread1(0).text = "Días Planificados"
         
         vaSpread1(0).Col = 7 '-------> Días Planificados
         vaSpread1(0).CellType = CellTypeCurrency
         vaSpread1(0).TypeCurrencyDecPlaces = 0
         vaSpread1(0).TypeCurrencyShowSymbol = False
         vaSpread1(0).Lock = True
         vaSpread1(0).TypeHAlign = TypeHAlignRight
         vaSpread1(0).text = Format(frecest, fg_Pict(6, 0))
         vaSpread1(0).Font.Bold = True
         vaSpread1(0).Font.Size = 9
         
         vaSpread1(0).Col = 14 '-------> Total
         vaSpread1(0).CellType = CellTypeNumber
         vaSpread1(0).Lock = True
         vaSpread1(0).TypeHAlign = TypeHAlignRight
         vaSpread1(0).Formula = Fg_Sacacremilla("SUM('" & "N" & indini & "':'" & "N" & vaSpread1(0).Row - 1 & "')" & "/" & "SUM('" & "G" & vaSpread1(0).Row & "')")
         vaSpread1(0).Font.Bold = True
         vaSpread1(0).Font.Size = 9
         
         vecTipoPla(IRow - 1, 3) = vaSpread1(0).Row
         
         vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
         vaSpread1(0).Row = vaSpread1(0).MaxRows
      
      End If
      
      vaSpread1(0).Col = 1 '-------> Descripción Estructura Servicio
      vaSpread1(0).CellType = CellTypeStaticText
      vaSpread1(0).Lock = True
      vaSpread1(0).TypeHAlign = TypeHAlignLeft
      vaSpread1(0).Font.Bold = True
      vaSpread1(0).Font.Size = 9
      vaSpread1(0).text = " " & Trim(RS!pas_nomest)
      
      auxest = RS!pas_estser
      vecTipoPla(IRow, 1) = RS!pas_estser
      vecTipoPla(IRow, 2) = " " & Trim(RS!pas_nomest)
      vecTipoPla(IRow, 4) = RS!Raciones
      vecTipoPla(IRow, 5) = RS!Comensales
      frecest = RS!pas_freestser
      indini = vaSpread1(0).Row
      IRow = IRow + 1
   
   End If
   
   vaSpread1(0).Col = 2 '-------> Codigo Ingrediente
   vaSpread1(0).CellType = CellTypeStaticText
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignLeft
   vaSpread1(0).text = IIf(IsNull(RS!pas_coding), "", " " & RS!pas_coding)
         
   vaSpread1(0).Col = 3 '-------> Nombre Ingrediente
   vaSpread1(0).CellType = CellTypeStaticText
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignLeft
   vaSpread1(0).text = IIf(IsNull(RS!ing_nombre), "", " " & Trim(RS!ing_nombre))
   
   vaSpread1(0).Col = 4 '-------> Nombre Unidad Medida
   vaSpread1(0).CellType = CellTypeStaticText
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignLeft
   vaSpread1(0).text = IIf(IsNull(RS!unm_nomcor), "", " " & Trim(RS!unm_nomcor))
   
   vaSpread1(0).Col = 6 '-------> Tipo Ingrediente
   vaSpread1(0).CellType = CellTypeStaticText
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignLeft
   vaSpread1(0).text = IIf(RS!ing_indppr = "1", "Real", "Propuesta")
   
   vaSpread1(0).Col = 7 '-------> Frecuencia
   vaSpread1(0).CellType = CellTypeCurrency
   vaSpread1(0).TypeCurrencyDecPlaces = 0
   vaSpread1(0).TypeCurrencyShowSymbol = False
   vaSpread1(0).Lock = False
   vaSpread1(0).TypeHAlign = TypeHAlignRight
   vaSpread1(0).text = Format(RS!pas_freing, fg_Pict(6, 0))

   vaSpread1(0).Col = 11 '-------> Precio ingrediente
   vaSpread1(0).CellType = CellTypeNumber
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignRight
   vaSpread1(0).text = Format(RS!ing_preing, fg_Pict(6, 4))
   vaSpread1(0).ForeColor = IIf(RS!ing_preing = 0 Or IsNull(RS!ing_preing), &HFF&, &HFF0000)
         
   vaSpread1(0).Col = 13 '-------> Gramaje
   vaSpread1(0).CellType = CellTypeNumber
   vaSpread1(0).Lock = False
   vaSpread1(0).TypeHAlign = TypeHAlignRight
   vaSpread1(0).text = Format(RS!pas_canpro, fg_Pict(6, vg_RDCa))
   
   vaSpread1(0).Col = 14 '-------> Total
   vaSpread1(0).CellType = CellTypeNumber
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignRight
   vaSpread1(0).Formula = "SUM((G#*k#*M#))"
   vaSpread1(0).ForeColor = IIf(CStr(Val(vaSpread1(0).text)) = "0", &HFF&, &HFF0000)
   
   RS.MoveNext

Loop
   
vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
vaSpread1(0).Row = vaSpread1(0).MaxRows
         
vaSpread1(0).Col = 1 '-------> Glosa Días Planificados
vaSpread1(0).CellType = CellTypeStaticText
vaSpread1(0).Lock = True
vaSpread1(0).TypeHAlign = TypeHAlignLeft
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9
vaSpread1(0).text = "Días Planificados"
    
vaSpread1(0).Col = 7 '-------> Días Planificados
vaSpread1(0).CellType = CellTypeCurrency
vaSpread1(0).TypeCurrencyDecPlaces = 0
vaSpread1(0).TypeCurrencyShowSymbol = False
vaSpread1(0).Lock = True
vaSpread1(0).TypeHAlign = TypeHAlignRight
vaSpread1(0).text = Format(frecest, fg_Pict(6, 0))
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9
    
vaSpread1(0).Col = 14 '-------> Total
vaSpread1(0).CellType = CellTypeNumber
vaSpread1(0).Lock = True
vaSpread1(0).TypeHAlign = TypeHAlignRight
vaSpread1(0).Formula = Fg_Sacacremilla("SUM('" & "N" & indini & "':'" & "N" & vaSpread1(0).Row - 1 & "')" & "/" & "SUM('" & "G" & vaSpread1(0).Row & "')")
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9
    
vecTipoPla(IRow - 1, 3) = vaSpread1(0).Row

RS.Close
Set RS = Nothing
Label1(9).Caption = Format(vaSpread1(0).MaxRows, fg_Pict(6, 2))
Label1(11).Caption = Format(totgralreceta, fg_Pict(6, 2))

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_nRegMinutaBloque '" & CecoAux & "', " & CodRegimenAux & "," & CodServicioAux & "," & FechaInicio & "," & FechaFinal & "")

If Not RS.EOF And RS!nReg > 0 Then
   
   Label1(11).Caption = Format(CCur(totgralreceta / RS!nReg), fg_Pict(6, 2))

End If
RS.Close
Set RS = Nothing

'-------> mover sub-segmento
vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 2
vaSpread1(0).Row = vaSpread1(0).MaxRows
vaSpread1(0).Col = 3
vaSpread1(0).text = "Sub-Segmento"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).Col = 4
vaSpread1(0).text = fpayuda(0).Caption
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

'-------> Mover regimen
vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
vaSpread1(0).Row = vaSpread1(0).MaxRows
vaSpread1(0).Col = 3
vaSpread1(0).text = "Regimen"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).Col = 4
vaSpread1(0).text = fpayuda(1).Caption
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

'-------> mover Servicio
vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
vaSpread1(0).Row = vaSpread1(0).MaxRows
vaSpread1(0).Col = 3
vaSpread1(0).text = "Servicio"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).Col = 4
vaSpread1(0).text = fpayuda(3).Caption & " " & FpFecDesde.text & " - " & FpFecHasta.text
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

'-------> Mover resumen
vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 2
vaSpread1(0).Row = vaSpread1(0).MaxRows
vaSpread1(0).Col = 3
vaSpread1(0).text = "RESUMEN COSTO"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
vaSpread1(0).Row = vaSpread1(0).MaxRows
vaSpread1(0).Col = 3
vaSpread1(0).text = "Estructura Servicio"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).Col = 4
vaSpread1(0).text = "Costo"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).Col = 5
vaSpread1(0).text = "Ponderado"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).Col = 6
vaSpread1(0).text = " % " '"Ponderado"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).Col = 7
vaSpread1(0).text = "Costo Ponderado" '"%"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

indfin = 0
Dim costoestservicio As Double

For i = 1 To UBound(vecTipoPla) 'iRow - 1
   
   If Trim(vecTipoPla(i, 2)) <> "" Then
      
      vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
      vaSpread1(0).Row = vaSpread1(0).MaxRows
   
      vaSpread1(0).Col = 3 '-------> Descripción Estructura
      vaSpread1(0).text = "  " & vecTipoPla(i, 2)
      vaSpread1(0).Lock = True
      vaSpread1(0).Font.Bold = True
      vaSpread1(0).Font.Size = 9

      vaSpread1(0).Col = 4 '-------> Total
      vaSpread1(0).CellType = CellTypeNumber
      vaSpread1(0).Lock = True
      vaSpread1(0).TypeHAlign = TypeHAlignRight
      vaSpread1(0).Formula = Fg_Sacacremilla("SUM('" & "N" & vecTipoPla(i, 3) & "':'" & "N" & vecTipoPla(i, 3) & "')")
      vaSpread1(0).Font.Bold = True
      vaSpread1(0).Font.Size = 9

      costoestservicio = IIf(Trim(vaSpread1(0).text) = "", 0, vaSpread1(0).text)

      vaSpread1(0).Col = 7 '6 '-------> Ponderación
      vaSpread1(0).CellType = CellTypeNumber
      vaSpread1(0).Lock = True
      vaSpread1(0).TypeHAlign = TypeHAlignRight
      vaSpread1(0).Font.Bold = True
      vaSpread1(0).Font.Size = 9
      
      If costoestservicio > 0 And vecTipoPla(i, 5) > 0 And vecTipoPla(i, 4) > 0 Then
      
         vaSpread1(0).Formula = Fg_Sacacremilla("'" & "D" & vaSpread1(0).Row & "'*'" & "F" & vaSpread1(0).Row & "'")  '(vecTipoPla(i, 4) / vecTipoPla(i, 5)) * costoestservicio
   
         vaSpread1(0).Col = 6 '7 '-------> %
'         vaSpread1(0).CellType = CellTypeNumber
         vaSpread1(0).CellType = CellTypePercent
         vaSpread1(0).TypePercentLeadingZero = TypeLeadingZeroYes
         vaSpread1(0).TypePercentNegStyle = TypePercentNegStyle8
         vaSpread1(0).TypePercentDecPlaces = 2
         
         vaSpread1(0).Lock = True
         vaSpread1(0).TypeHAlign = TypeHAlignRight
         vaSpread1(0).text = (vecTipoPla(i, 4) / vecTipoPla(i, 5)) * 100
         vaSpread1(0).Font.Bold = True
         vaSpread1(0).Font.Size = 9
      
         'vaSpread1(0).Col = 7 '6 '-------> Ponderación
         
         'vaSpread1(0).Formula = Fg_Sacacremilla("'" & "D" & vaSpread1(0).Row & "'*'" & "F" & vaSpread1(0).Row & "'")  '(vecTipoPla(i, 4) / vecTipoPla(i, 5)) * costoestservicio
      
      Else
   
         vaSpread1(0).Formula = Fg_Sacacremilla("'" & "D" & vaSpread1(0).Row & "'*'" & "F" & vaSpread1(0).Row & "'")
         'vaSpread1(0).text = 0
         
         vaSpread1(0).Col = 6 '7 '-------> %
         vaSpread1(0).CellType = CellTypePercent
         vaSpread1(0).TypePercentLeadingZero = TypeLeadingZeroYes
         vaSpread1(0).TypePercentNegStyle = TypePercentNegStyle8
         vaSpread1(0).TypePercentDecPlaces = 2
         vaSpread1(0).text = 0
         
      End If
      
      vaSpread1(0).Font.Bold = True
      vaSpread1(0).Font.Size = 9
      
   End If
    
Next i

fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub TextCai1_Change(Index As Integer)

On Error GoTo Man_Error

Dim i         As Long
Dim indactivo As Integer
Dim nom       As String
Dim icol      As Long
icol = IIf(Index = 1, 2, IIf(Index = 2, 3, IIf(Index = 3, 8, 9)))

Select Case Index

Case 1, 2, 3, 0
    
    vaSpread1(0).Visible = False
    
    If Trim(TextCai1(Index).text) <> "" Then
       
       For i = 1 To vaSpread1(0).MaxRows
           
           vaSpread1(0).Row = i
           vaSpread1(0).Col = icol: nom = UCase(Trim(vaSpread1(0).text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCai1(Index).text) & "*"
           vaSpread1(0).Col = icol
           
           If indactivo = -1 And Trim(vaSpread1(0).text) <> "" Then
              
              If vaSpread1(0).RowHidden = True Then vaSpread1(0).RowHidden = False
           
           Else
              
              If vaSpread1(0).RowHidden = False Then vaSpread1(0).RowHidden = True
           
           End If
        
        Next i
        
        vaSpread1(0).SetActiveCell Index + 1, 1
    
    End If
    
    vaSpread1(0).ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1(0).ColUserSortIndicator(IIf(Trim(TextCai1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1(0).SortKey(1) = IIf(Trim(TextCai1(Index).text) = "", 0, 0): vaSpread1(0).SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1(0).Sort -1, -1, vaSpread1(0).maxcols, vaSpread1(0).MaxRows, SortByRow
    
    If Trim(TextCai1(Index).text) = "" Then
       
       For i = 1 To vaSpread1(0).MaxRows
           
           vaSpread1(0).Row = i
           If vaSpread1(0).RowHidden = True Then vaSpread1(0).RowHidden = False
       
       Next
       
       vaSpread1(0).SetActiveCell Index + 1, vaSpread1(0).SearchCol(Index, 0, vaSpread1(0).MaxRows, Trim(TextCai1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1(0).SetActiveCell Index + 1, 1
    
    End If
    
    vaSpread1(0).Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index

Case 2
    
    If vaSpread1(0).MaxRows < 1 Then
    
        MsgBox "No existe información seleccionada ", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
    
    End If
    
    Dim X As Boolean
    TextCai1(1).text = ""
    TextCai1(2).text = ""
    vaSpread1(0).Row = -1
    vaSpread1(0).Col = -1
    vaSpread1(0).RowHidden = False
    
    ' Export Excel file and set result to x
    If Dir(dir_trabajo & "Frecuencia Ingrediente Minuta Bloque.XLS") <> "" Then
       
       Kill dir_trabajo & "Frecuencia Ingrediente Minuta Bloque.XLS"
    
    End If
    
    X = vaSpread1(0).ExportToExcel(dir_trabajo & "Frecuencia Ingrediente Minuta Bloque.XLS", "Test Sheet 1", dir_trabajo & "LOGFILE.TXT")
    ' Display result to user based on T/F value of x
    If X = True Then

'        MsgBox "Export complete.", , "Result"
        Dim XL As excel.Application
        Set XL = CreateObject("Excel.application")
        XL.Workbooks.Open FileName:=dir_trabajo & "Frecuencia Ingrediente Minuta Bloque.XLS"
        XL.Cells.Select ''-------> Desactivar proteción
        XL.ActiveSheet.Unprotect
        XL.Rows("1:1").Select '------> Insert Fila
        XL.Selection.Insert 'Shift:=xlDown
        XL.Range("A1").Select
        XL.ActiveCell.FormulaR1C1 = "Estructura Servicio"
        XL.Range("B1").Select
        XL.ActiveCell.FormulaR1C1 = "Código Ingrediente"
        XL.Range("C1").Select
        XL.ActiveCell.FormulaR1C1 = "Descripción"
        XL.Range("D1").Select
        XL.ActiveCell.FormulaR1C1 = "Unidad Ingrediente"
        XL.Range("F1").Select
        XL.ActiveCell.FormulaR1C1 = "Tipo Ingrediente"
        XL.Range("G1").Select
        XL.ActiveCell.FormulaR1C1 = "Frecuencia Ingrediente"
        XL.Range("K1").Select
        XL.ActiveCell.FormulaR1C1 = "Precio"
        XL.Range("M1").Select
        XL.ActiveCell.FormulaR1C1 = "Gramaje"
        XL.Range("N1").Select
        XL.ActiveCell.FormulaR1C1 = "Total"
        XL.ActiveWindow.SplitRow = 0.625
        XL.ActiveWindow.SplitRow = 0.6875
        XL.Cells.Select '-------> Activar proteción
        XL.ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        XL.Visible = True '------->Visualizar
    
    Else
        
        MsgBox "Archivo esta abierto, grabe con otro nombre y luego cierre libro", , "Result"
    
    End If

Case 4
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
If Err = 70 Or Err = 1004 Or Err = 91 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Recetas"

End Sub

Private Sub vaSpread1_TextTipFetch(Index As Integer, ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

On Error GoTo Man_Error

If vaSpread1(0).MaxRows < 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread1(0).Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2

Select Case Col

Case 1
    vaSpread1(0).Col = Col
    TipText = "Estructura Servicio : " & vaSpread1(0).text

Case 2
    vaSpread1(0).Col = Col
    TipText = "Código Ingrediente : " & Trim(vaSpread1(0).text)

Case 3
    vaSpread1(0).Col = Col
    TipText = "Descripción Ingrediente : " & Trim(vaSpread1(0).text)

Case 4
    vaSpread1(0).Col = Col
    TipText = "Tipo Ingrediente : " & Trim(vaSpread1(0).text)

Case 5
    vaSpread1(0).Col = Col
    TipText = "Frecuencia Ingrediente : " & Trim(vaSpread1(0).text)

Case 6
    vaSpread1(0).Col = Col
    TipText = "Código Producto : " & Trim(vaSpread1(0).text)

Case 7
    vaSpread1(0).Col = Col
    TipText = "Descripción Producto : " & Trim(vaSpread1(0).text)

Case 8
    vaSpread1(0).Col = Col
    TipText = "Tipo Ingrediente : " & Trim(vaSpread1(0).text)

Case 9
    vaSpread1(0).Col = Col
    TipText = "Precio : " & Trim(vaSpread1(0).text)

Case 10
    vaSpread1(0).Col = Col
    TipText = "Gramaje : " & Trim(vaSpread1(0).text)

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub
