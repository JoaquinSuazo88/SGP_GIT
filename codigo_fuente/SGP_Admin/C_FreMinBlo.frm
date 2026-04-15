VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form C_FreMinBlo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frecuencia Receta Minuta Bloque"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   15525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   0
      TabIndex        =   13
      Top             =   2100
      Width           =   14835
      Begin VB.Frame Frame12 
         Height          =   435
         Left            =   1410
         TabIndex        =   18
         Top             =   4680
         Width           =   3285
         Begin VB.TextBox TextCai1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   19
            Top             =   135
            Width           =   3180
         End
      End
      Begin VB.Frame Frame13 
         Height          =   435
         Left            =   600
         TabIndex        =   16
         Top             =   4680
         Width           =   675
         Begin VB.TextBox TextCai1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   17
            Top             =   135
            Width           =   570
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   6240
         TabIndex        =   14
         Top             =   4680
         Width           =   3285
         Begin VB.TextBox TextCai1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   15
            Top             =   135
            Width           =   3180
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4395
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   14505
         _Version        =   393216
         _ExtentX        =   25585
         _ExtentY        =   7752
         _StockProps     =   64
         ColsFrozen      =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         MaxRows         =   18
         SpreadDesigner  =   "C_FreMinBlo.frx":0000
         VisibleCols     =   4
         VisibleRows     =   18
         ScrollBarTrack  =   3
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
         Left            =   13380
         TabIndex        =   24
         Top             =   5055
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
         Index           =   9
         Left            =   13380
         TabIndex        =   23
         Top             =   4755
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEFEDE&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Costo Promedio Diario"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   11325
         TabIndex        =   22
         Top             =   5055
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEFEDE&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Recetas Listadas"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   11325
         TabIndex        =   21
         Top             =   4755
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   9375
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
         Left            =   7395
         TabIndex        =   1
         Top             =   1320
         Width           =   915
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   2760
         TabIndex        =   26
         Top             =   1320
         Width           =   1290
         _Version        =   196608
         _ExtentX        =   2284
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
         ThreeDTextHighlightColor=   -2147483633
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
         InvalidColor    =   -2147483643
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483624
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         DateCalcMethod  =   1
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
      Begin EditLib.fpDateTime FpFecHasta 
         Height          =   315
         Left            =   5880
         TabIndex        =   27
         Top             =   1320
         Width           =   1290
         _Version        =   196608
         _ExtentX        =   2284
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
         ThreeDTextHighlightColor=   -2147483633
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
         InvalidColor    =   -2147483643
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483624
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         DateCalcMethod  =   1
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   7440
         TabIndex        =   9
         Top             =   1350
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1410
         TabIndex        =   8
         Top             =   1380
         Width           =   675
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
         Left            =   1395
         TabIndex        =   7
         Top             =   1005
         Width           =   705
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
         Left            =   1395
         TabIndex        =   6
         Top             =   645
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ceco"
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
         Left            =   1395
         TabIndex        =   5
         Top             =   300
         Width           =   450
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2760
         TabIndex        =   3
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   2760
         TabIndex        =   2
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   2805
         TabIndex        =   10
         Top             =   285
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2805
         TabIndex        =   11
         Top             =   645
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   2805
         TabIndex        =   12
         Top             =   1005
         Width           =   5535
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7575
      Left            =   14895
      TabIndex        =   25
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   13361
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "C_FreMinBlo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private RS1         As New ADODB.Recordset
Private BtnX        As Variant
Private VarFecHasta As Long
Private VarFecDesde As Long
Private VarCencos   As Variant
Private VarCodReg   As Long
Private VarCodSer   As Long
Private VarTipMin   As String
Private VarTfor     As String
Private FecInicio   As Long
Private FecFin      As Long
Private Est         As Boolean

Private Sub CmdBuscar_Click()
    
On Error GoTo Man_Error

    If Not ValidarDatos Then Exit Sub

    Let Label1(9).Caption = ""
    Let Label1(11).Caption = ""
    Call LlenarFrecPlan(VarTfor, VarCencos, VarCodReg, VarCodSer, VarFecDesde, VarTipMin, VarFecHasta)

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Activate()
    
On Error GoTo Man_Error

    fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

    Est = True
    
    Call fg_centra(Me)
    Toolbar1.ImageList = Partida.IL1
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Excel "
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
    
    'Parametrizar fecha desde
    FpFecDesde.DateMin = CStr(VarFecDesde)
    FpFecDesde.DateMax = CStr(VarFecHasta)
    FpFecDesde.text = fg_Ctod1(VarFecDesde)
    
    'Parametrizar fecha hasta
    FpFecHasta.DateMin = CStr(VarFecDesde)
    FpFecHasta.DateMax = CStr(VarFecHasta)
    FpFecHasta.text = fg_Ctod1(VarFecHasta)

    Est = False
    
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub LlenarFrecPlan(tfor As String, Ceco As Variant, codReg As Long, codser As Long, FecDesde As Long, TipMin As String, FecHasta As Long)

On Error GoTo Man_Error

Dim RS               As New ADODB.Recordset
Dim CodReceta        As Long
Dim AuxCodReceta     As String
Dim IRow             As Long
Dim i                As Long
Dim condia           As Long
Dim auxtip           As Long
Dim codfre           As Long
Dim cosreceta        As Double
Dim canreceta        As Double
Dim totgralreceta    As Double
Dim tippla           As String
Dim dia              As Long
Dim Confre           As Long
Dim SearchFlagsEqual As Variant
Dim ind_ini          As Long
Dim CodRec           As String
Dim X                As Boolean
Dim vecTipoPla()     As Variant
Dim fecfin1          As Long
Dim MaxColumna       As Long
Dim EstBus           As Boolean

Est = False

Let VarTfor = tfor
Let VarCencos = Ceco
Let VarCodReg = codReg
Let VarCodSer = codser
Let VarTipMin = TipMin
VarFecDesde = FecDesde
VarFecHasta = FecHasta


fg_carga ""
'-------> Rutina frecuencia de recetas
Me.Caption = tfor
MsgTitulo = tfor
    
Est = True

Let Label1(0).Caption = "Cliente"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT isnull(cli_codigo,'') as cli_codigo, isnull(cli_nombre,'') as cli_nombre FROM b_clientes with (nolock) WHERE cli_codigo = '" & Ceco & "' and cli_tipo = 0")
If Not RS.EOF Then fpayuda(0).Caption = Trim(RS!Cli_nombre)
RS.Close: Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("SELECT isnull(reg_nombre,'') as reg_nombre FROM a_regimen with (nolock) WHERE reg_codigo = " & codReg & "")
If Not RS.EOF Then fpayuda(1).Caption = Trim(RS!reg_nombre)
RS.Close: Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("SELECT isnull(ser_nombre,'') as ser_nombre FROM a_servicio with (nolock) WHERE ser_codigo = " & codser & "")
If Not RS.EOF Then fpayuda(3).Caption = Trim(RS!ser_nombre)
RS.Close: Set RS = Nothing
    
' Control displays text tips aligned to pointer with focus
vaSpread1(0).TextTip = 2
vaSpread1(0).TextTipDelay = 250
X = vaSpread1(0).SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread1(0).Row = -1: vaSpread1(0).Col = 1
vaSpread1(0).BackColor = &HDEFEDE
vaSpread1(0).Row = -1: vaSpread1(0).Col = 2
vaSpread1(0).BackColor = &HDEFEDE
vaSpread1(0).Row = -1: vaSpread1(0).Col = 3
vaSpread1(0).BackColor = &HDEFEDE
vaSpread1(0).Row = -1: vaSpread1(0).Col = 4
vaSpread1(0).BackColor = &HDEFEDE
vaSpread1(0).Row = -1: vaSpread1(0).Col = 5
vaSpread1(0).BackColor = &HDEFEDE
vaSpread1(0).Row = -1: vaSpread1(0).Col = 6
vaSpread1(0).BackColor = &HDEFEDE
    
vaSpread1(0).MaxRows = 0
'-------> Buscar Nº días
    
'-------> determinar la cuando días entre la fecha desde - hasta
MaxColumna = DateDiff("d", CDate(fg_Ctod1(FecDesde)), CDate(fg_Ctod1(FecHasta))) + 1

vaSpread1(0).maxcols = MaxColumna + 8: vaSpread1(0).Row = 0

Dim FechaIni As Date
Let FechaIni = fg_Ctod1(FecDesde)
For i = 9 To vaSpread1(0).maxcols
        
    vaSpread1(0).Row = 0
    vaSpread1(0).Col = i
    vaSpread1(0).text = " " & Mid(fg_Fecha_Dia(Format(FechaIni, "yyyymmdd"), 1), 1, 3) & " " & FechaIni
    Let FechaIni = FechaIni + 1
    vaSpread1(0).ColHidden = False
        
Next i

Dim SeleccionOpt As Integer
SeleccionOpt = IIf(M_MinSR1.optPrecioConvenio.Value = True, 1, IIf(M_MinSR1.optPrecioGenerico = True, 2, 3))

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_Sel_FrecRecetaMinutaBloque_V06 '" & Ceco & "', " & codReg & ", " & codser & ", " & FecDesde & ", " & FecHasta & ", '1', " & SeleccionOpt & "")

DoEvents
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   Exit Sub

End If
    
CodReceta = 0
cosreceta = 0
canreceta = 0
totgralreceta = 0
condia = 0

Dim essnom As String
essnom = ""
IRow = 1
Confre = 0
'definir largo del vector

ReDim Preserve vecTipoPla(1000, 3)
Dim AuxCodigoEst As Long
AuxCodigoEst = 0
Do While Not RS.EOF

    DoEvents
        
    If auxtip <> RS!rec_tippla Or AuxCodigoEst <> RS!ess_codigo Then

        If auxtip <> 0 Then
                
            EstBus = False
            
            For i = 1 To UBound(vecTipoPla)
                
                If vecTipoPla(i, 1) = tippla And vecTipoPla(i, 3) = essnom Then
                   
                   vecTipoPla(i, 2) = vecTipoPla(i, 2) + Confre
                   EstBus = True
                   Exit For
                   
                End If
            
            Next i
            
            If Not EstBus Then
            
               vecTipoPla(IRow, 1) = tippla
               vecTipoPla(IRow, 2) = Confre
               vecTipoPla(IRow, 3) = essnom
               IRow = IRow + 1
            
            End If
            Confre = 0
            
        End If
            
        auxtip = RS!rec_tippla
        tippla = RS!nom_tippla
        essnom = RS!ess_nombre
        AuxCodigoEst = RS!ess_codigo
        
    End If

    vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
    vaSpread1(0).Row = vaSpread1(0).MaxRows

    vaSpread1(0).Col = 1
    vaSpread1(0).CellType = CellTypeStaticText
    vaSpread1(0).TypeHAlign = TypeHAlignLeft
    vaSpread1(0).text = RS!pas_codrec

    vaSpread1(0).Col = 2
    vaSpread1(0).CellType = CellTypeStaticText
    vaSpread1(0).TypeHAlign = TypeHAlignLeft
    vaSpread1(0).text = Trim(RS!rec_nombre)

    vaSpread1(0).Col = 3
    vaSpread1(0).CellType = CellTypeStaticText
    vaSpread1(0).TypeHAlign = TypeHAlignLeft
    vaSpread1(0).text = Trim(RS!ess_nombre)

    ' --------------- Samuel Melendez 03/09/09 ----
    vaSpread1(0).Col = 4
    vaSpread1(0).CellType = CellTypeStaticText
    vaSpread1(0).TypeHAlign = TypeHAlignLeft
    vaSpread1(0).text = Trim(RS!nom_tippla)
    ' ---------------------------------------------

    If VarSitioRemoto = False Then
            
       vaSpread1(0).Col = 5
       vaSpread1(0).CellType = CellTypeStaticText
       vaSpread1(0).TypeHAlign = TypeHAlignLeft
       vaSpread1(0).text = IIf(RS!rec_indppr = "1", "Real", "Propuesta")
        
    End If

    vaSpread1(0).Col = 6
    vaSpread1(0).CellType = CellTypeStaticText
    vaSpread1(0).TypeHAlign = TypeHAlignRight
    vaSpread1(0).text = Format(RS!nrorec, fg_Pict(6, 0))
    vaSpread1(0).ForeColor = &HFF0000

    vaSpread1(0).Col = 7
    vaSpread1(0).CellType = CellTypeStaticText
    vaSpread1(0).TypeHAlign = TypeHAlignRight
    vaSpread1(0).text = Format(RS!rec_prerec, fg_Pict(6, 2))
    vaSpread1(0).ForeColor = &HFF0000

    vaSpread1(0).Col = 8
    vaSpread1(0).CellType = CellTypeStaticText
    vaSpread1(0).TypeHAlign = TypeHAlignLeft
    vaSpread1(0).text = RS!pas_codrec & RS!ess_codigo
    
    Confre = Confre + RS!nrorec
    RS.MoveNext
        
Loop

If auxtip <> 0 Then
        
   EstBus = False
            
   For i = 1 To UBound(vecTipoPla)
                
       If vecTipoPla(i, 1) = tippla And vecTipoPla(i, 3) = essnom Then
                   
          vecTipoPla(i, 2) = vecTipoPla(i, 2) + Confre
          EstBus = True
          Exit For
                   
       End If
            
   Next i
            
   If Not EstBus Then
   
      vecTipoPla(IRow, 1) = tippla
      vecTipoPla(IRow, 2) = Confre
      vecTipoPla(IRow, 3) = essnom
    
   End If
    
End If

Confre = 0
RS.Close
Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_Sel_FrecRecetaxFecha_V02 " & Ceco & ", " & codReg & ", " & codser & ", " & FecDesde & " ," & FecHasta & ", '1'")

Do While Not RS.EOF
    
   DoEvents
   AuxCodReceta = RS!mid_codrec & RS!mid_estser
   ind_ini = vaSpread1(0).SearchCol(8, -1, vaSpread1(0).MaxRows, AuxCodReceta, SearchFlagsEqual)

   If ind_ini <> -1 Then

      '-------> Buscar día
      For i = 9 To vaSpread1(0).maxcols
          
          vaSpread1(0).Row = 0
          vaSpread1(0).Col = i
          
          If CDate(Mid(Trim(vaSpread1(0).text), 5, Len(Trim(vaSpread1(0).text)))) = fg_Ctod1(RS!min_fecmin) Then
                
             vaSpread1(0).Row = ind_ini
             vaSpread1(0).Col = i '7 + (dia + Val(Mid(RS!min_fecmin, 7, 2)))
             vaSpread1(0).CellType = CellTypeStaticText
             vaSpread1(0).TypeHAlign = TypeHAlignRight
             vaSpread1(0).text = 0
             vaSpread1(0).text = CCur(Val(vaSpread1(0).text) + RS!mid_numrac)
             vaSpread1(0).ForeColor = &HFF0000
             Exit For
               
          End If
          
      Next i
            
   End If
   
   RS.MoveNext

Loop
    
RS.Close
Set RS = Nothing
cosreceta = 0
canreceta = 0
totgralreceta = 0

Dim TotalReceta As Double
TotalReceta = 0

For i = 1 To vaSpread1(0).MaxRows

    vaSpread1(0).Row = i
    vaSpread1(0).Col = 6
    canreceta = Val(vaSpread1(0).text)
    vaSpread1(0).Col = 7
    cosreceta = Val(vaSpread1(0).text)
    totgralreceta = CCur(totgralreceta + (cosreceta * canreceta))

Next i

TotalReceta = vaSpread1(0).MaxRows

vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 2
vaSpread1(0).Row = vaSpread1(0).MaxRows
vaSpread1(0).Col = 2
vaSpread1(0).text = "TOTALES X TIPO PLATO"
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

For i = 1 To IRow

    vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
    vaSpread1(0).Row = vaSpread1(0).MaxRows
    vaSpread1(0).Col = 3
    vaSpread1(0).text = vecTipoPla(i, 3)
    vaSpread1(0).Font.Bold = True
    vaSpread1(0).Font.Size = 9

    vaSpread1(0).Col = 4
    vaSpread1(0).text = vecTipoPla(i, 1)
    vaSpread1(0).Font.Bold = True
    vaSpread1(0).Font.Size = 9

    vaSpread1(0).Col = 6
    vaSpread1(0).text = vecTipoPla(i, 2)
    vaSpread1(0).Font.Bold = True
    vaSpread1(0).Font.Size = 9

Next i

Label1(9).Caption = Format(TotalReceta, fg_Pict(6, 2)) 'Format(vaSpread1(0).MaxRows, fg_Pict(6, 2))
Label1(11).Caption = Format(totgralreceta, fg_Pict(6, 2))

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_Sel_FrecRecetaXNRegistro " & Ceco & ", " & codReg & ", " & codser & ", " & FecDesde & " ," & FecHasta & ", '1'")

If Not RS.EOF And RS!nReg > 0 Then
   
   Label1(11).Caption = Format(CCur(totgralreceta / RS!nReg), fg_Pict(6, 2))
   
End If
RS.Close
Set RS = Nothing

Est = False

fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If Est Then Exit Sub

If IsDate(FpFecDesde.text) = False Then Exit Sub

Let VarFecDesde = Format(FpFecDesde.text, "yyyymmdd")
vaSpread1(0).MaxRows = 0

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

   If Est Then Exit Sub

   If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
    
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

If Est Then Exit Sub

If IsDate(FpFecHasta.text) = False Then Exit Sub

Let VarFecHasta = Format(FpFecHasta.text, "yyyymmdd")
vaSpread1(0).MaxRows = 0

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

   If Est Then Exit Sub

   If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
    
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub TextCai1_Change(Index As Integer)

On Error GoTo Man_Error

Dim i As Long
Dim indactivo As Integer
Dim nom As String
Dim icol As Long
icol = IIf(Index = 1, 1, IIf(Index = 2, 2, 4))
Select Case Index

Case 1, 2, 0
    vaSpread1(0).Visible = False
    If Trim(TextCai1(Index).text) <> "" Then
       
       For i = 1 To vaSpread1(0).MaxRows
           
           If i = 430 Then
              
              nom = 1
           
           End If
           
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
        
        vaSpread1(0).SetActiveCell Index, 1
    
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
       
       vaSpread1(0).SetActiveCell Index, vaSpread1(0).SearchCol(Index, 0, vaSpread1(0).MaxRows, Trim(TextCai1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1(0).SetActiveCell Index, 1
    
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
    
    If vaSpread1(0).MaxRows < 1 Then Exit Sub
    ExportarExcel

Case 4
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub ExportarExcel()

On Error GoTo Man_Error

Dim NashXl  As excel.Application
Dim IRow    As Long
Dim irow2   As Long
Dim i       As Long
Dim oCol    As String
Dim AoCol   As String
Dim oColA   As String
Dim IndCol  As Long
Dim IndColA As Long

fg_carga ""

oCol = ""
AoCol = ""
oColA = ""
IndCol = 1
IndColA = 65
oCol = Chr(IndCol + 64)
'ReDim VecDiaExcel(MaxCol, 2)

For i = 1 To vaSpread1(0).maxcols
    
    If Chr(IndCol + 65) = "[" Then
       
       oColA = Chr(IndColA)
       IndColA = IndColA + 1
       IndCol = 0
    
    End If
    oCol = oColA & Chr(IndCol + 65)
    IndCol = IndCol + 1

Next i

Set NashXl = CreateObject("excel.application")
Set NashXl = New excel.Application
NashXl.SheetsInNewWorkbook = 1
NashXl.Workbooks.Add

vaSpread1(0).AllowMultiBlocks = True
vaSpread1(0).SetSelection 1, -1, vaSpread1(0).maxcols, vaSpread1(0).MaxRows
vaSpread1(0).ClipboardCopy
IRow = vaSpread1(0).MaxRows + 1
'------- Pegar vaspread1(1) - Planilla Excel
NashXl.Range("A1").Select
NashXl.ActiveSheet.Paste
'------- Asignar color
'NashXl.Range("A1:D" & irow).Select
'With NashXl.Selection.Interior
'     .ColorIndex = 36
'     .Pattern = xlSolid
'End With
'------- Colorear titulo
'NashXl.Range("A1:AW1").Select ' samuel 03/0309
NashXl.Range("A1:" & oCol & "1").Select
With NashXl.Selection.Interior
     .ColorIndex = 15
     .Pattern = xlSolid
End With
'------- Dibujar marco
'NashXl.Range("A1:AW" & iRow).Select ' samuel 03/09/09
NashXl.Range("A1:" & oCol & IRow).Select
NashXl.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
NashXl.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With NashXl.Selection.Borders(xlEdgeLeft)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeTop)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeBottom)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeRight)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideVertical)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideHorizontal)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
NashXl.Range("D2" & ":" & "D" & IRow).Select
NashXl.Selection.NumberFormat = "#,##0.00"
'------- Aplicar totales

'------- Dibujar marco
IRow = IRow + 2
irow2 = IRow + 2
NashXl.Range("B" & IRow).Select
NashXl.ActiveCell.FormulaR1C1 = Label1(8).Caption
NashXl.Range("C" & IRow).Select
NashXl.ActiveCell.FormulaR1C1 = Label1(9).Caption
NashXl.Range("B" & irow2).Select
NashXl.ActiveCell.FormulaR1C1 = Label1(10).Caption
NashXl.Range("C" & irow2).Select
NashXl.ActiveCell.FormulaR1C1 = Label1(11).Caption
NashXl.Range("B" & IRow & ":" & "C" & irow2).Select
NashXl.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
NashXl.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With NashXl.Selection.Borders(xlEdgeLeft)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeTop)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeBottom)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeRight)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideVertical)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideHorizontal)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
NashXl.Selection.Font.Bold = True
'With NashXl.Selection.Interior
'     .ColorIndex = 35
'     .Pattern = xlSolid
'End With
NashXl.Range("D" & IRow & ":" & "D" & irow2).Select
NashXl.Selection.NumberFormat = "#,##0.00"
'------- Ajustar columna
NashXl.Cells.Select
NashXl.Cells.EntireColumn.AutoFit
vaSpread1(0).AllowMultiBlocks = False: vaSpread1(0).SetSelection 1, 0, vaSpread1(0).maxcols, vaSpread1(0).MaxRows
fg_descarga
NashXl.Visible = True

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

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
    TipText = "Código : " & vaSpread1(0).text

Case 2
    
    vaSpread1(0).Col = Col
    TipText = "Nombre Receta : " & Trim(vaSpread1(0).text)

Case 3
    
    vaSpread1(0).Col = Col
    TipText = "Tipo Plato : " & Trim(vaSpread1(0).text)

Case 4
    
    vaSpread1(0).Col = Col
    TipText = "Tipo Receta : " & Trim(vaSpread1(0).text)

Case 5
    
    vaSpread1(0).Col = Col
    TipText = "Frecuencia : " & Trim(vaSpread1(0).text)

Case 6
    
    vaSpread1(0).Col = Col
    TipText = "Costo : " & Trim(vaSpread1(0).text)

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Function ValidarDatos() As Boolean

Dim seleccion As Integer
Dim i As Long

ValidarDatos = True

'-------> Validar fechas
If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
   
   MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If
    
If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
   
   MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

End Function


