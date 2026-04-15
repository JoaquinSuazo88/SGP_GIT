VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_MBloqueCostoBandejaxServicios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minuta Bloque Costo Bandeja x Servicio "
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16215
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   16215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Detalle Costo Plato Minuta Bloque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   105
      TabIndex        =   15
      Top             =   4620
      Width           =   15975
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
         Left            =   12915
         TabIndex        =   8
         Top             =   4200
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
         Left            =   14280
         TabIndex        =   9
         Top             =   4200
         Width           =   1275
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   3615
         Left            =   315
         TabIndex        =   7
         Top             =   315
         Width           =   15510
         _Version        =   393216
         _ExtentX        =   27358
         _ExtentY        =   6376
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
         MaxCols         =   10
         SpreadDesigner  =   "M_MBloqueCostoBandejaxServicios.frx":0000
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   210
         Left            =   315
         TabIndex        =   16
         Top             =   4095
         Visible         =   0   'False
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lbl_proceso 
         Alignment       =   2  'Center
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   6195
         TabIndex        =   17
         Top             =   4095
         Visible         =   0   'False
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ceco Bloque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4320
      Left            =   2625
      TabIndex        =   10
      Top             =   105
      Width           =   11250
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   4200
         TabIndex        =   12
         Top             =   3360
         Width           =   5550
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   3
            Top             =   135
            Width           =   5445
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   3150
         TabIndex        =   11
         Top             =   3360
         Width           =   1005
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   2
            Top             =   135
            Width           =   900
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2535
         Left            =   1575
         TabIndex        =   1
         Top             =   735
         Width           =   9045
         _Version        =   393216
         _ExtentX        =   15954
         _ExtentY        =   4471
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
         SpreadDesigner  =   "M_MBloqueCostoBandejaxServicios.frx":1B68
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   2850
         TabIndex        =   4
         Top             =   3885
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
         Left            =   9300
         TabIndex        =   5
         Top             =   3885
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
         Left            =   10710
         TabIndex        =   6
         Top             =   3885
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   2850
         TabIndex        =   0
         Top             =   315
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Org. Compras"
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
         Left            =   1575
         TabIndex        =   18
         Top             =   380
         Width           =   1155
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
         Left            =   1575
         TabIndex        =   14
         Top             =   3930
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
         Left            =   7995
         TabIndex        =   13
         Top             =   3930
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
            Picture         =   "M_MBloqueCostoBandejaxServicios.frx":33FA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_MBloqueCostoBandejaxServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lc_Aux As String
Dim MsgTitulo As String

Private Sub Command1_Click()
On Error GoTo Man_Error

 '-------> Validar Datos en la grilla
  If vaSpread2.MaxRows < 1 Then
     MsgBox "No existe información a exportar...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  End If
    
  fg_carga ""

  ' Export Excel file and set result to x
  Dim XL As excel.Application
  Set XL = CreateObject("Excel.application")
  XL.Visible = True
  XL.Workbooks.OpenText vg_Archxls, , 1, 1, , , , , , , True, "|"

  fg_descarga
  MsgBox "Exportación Finalizada", vbInformation, Me.Caption
                
Exit Sub
Man_Error:
    fg_descarga
    
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Command2_Click()
'-------> Salir de la opción
Unload Me
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
fg_centra Me
MsgTitulo = "Copiar Minuta Bloque Ceco"

FpFecHasta.text = Format(Date, "dd/mm/yyyy")
FpFecDesde.text = Format(Date, "dd/mm/yyyy")
Command1.Enabled = False
vaSpread1.MaxRows = 0
'LlenarGrillaCeco

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub LlenarGrillaCeco()

On Error GoTo Man_Error

Dim RS   As New ADODB.Recordset
Dim Celo As String

vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &HC0FFFF
vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0
Text1(1).text = ""
Text1(2).text = ""
Command1.Enabled = False

Celo = LimpiaDato(Trim(fpText.text))

Sql = ""
Sql = " sgpadm_Sel_BuscaCecos "
Sql = Sql & " '" & Celo & "'"
Set RS = vg_db.Execute(Sql)

'Set RS = vg_db.Execute("sgpadm_Sel_CecoBloque")
Do While Not RS.EOF
   
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   
   vaSpread1.Col = 1
   vaSpread1.text = "0"
   
   vaSpread1.Col = 2
   vaSpread1.text = RS(0)
   
   vaSpread1.Col = 3
   vaSpread1.text = Trim(RS(1))
   
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub FpFecDesde_Change()
On Error GoTo Man_Error

vaSpread2.MaxRows = 0
If IsDate(FpFecDesde.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)
On Error GoTo Man_Error

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

vaSpread2.MaxRows = 0
If IsDate(FpFecHasta.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)
On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub fpText_Change()
On Error GoTo Man_Error

LlenarGrillaCeco

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Text1_Change(Index As Integer)
vaSpread2.MaxRows = 0
If Index = 1 Then
   Text1(2).text = ""
ElseIf Index = 2 Then
   Text1(1).text = ""
End If
Select Case Index
Case 1, 2
    vaSpread1.Visible = False
    If Trim(Text1(Index).text) <> "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = Index + 1
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread1.Col = 1
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           Else
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           End If
        Next i
        vaSpread1.SetActiveCell Index + 1, 1
    End If
'    vaSpread1_Click Index, 0
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
    If Trim(Text1(Index).text) = "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Next
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1.Visible = True
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS         As New ADODB.Recordset
Dim Sql        As String
Dim i          As Long
Dim j          As Long
Dim seleccion  As String
Dim Ceco       As String
Dim Conta      As Long
Dim CostoPlato As Double

Select Case Button.Index
Case 1

 '-------> Validar Datos en la grilla
  If vaSpread1.MaxRows < 1 Then
     
     MsgBox "No existe información a exportar...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If
  
  '-------> Validar fechas
  If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
     
     MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If
    
  If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
     
     MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If

  '-------> Validar que exista un dato seleccionado
  seleccion = 0
  For i = 1 To vaSpread1.MaxRows
       
       vaSpread1.Row = i
       vaSpread1.Col = 1 'Seleccion
       seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
       If seleccion = 1 Then
          Exit For
       End If
  
  Next i
  
  If seleccion = 0 Then
     
     vaSpread2.MaxRows = 0
     MsgBox " Se debe seleccionar un Ceco por lo menos", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If
  
  seleccion = 0
  Conta = 0
  For i = 1 To vaSpread1.MaxRows
       
       vaSpread1.Row = i
       vaSpread1.Col = 1 'Seleccion
       seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
       If seleccion = 1 And vaSpread1.RowHidden = False Then
          Conta = Conta + 1
       End If
       
  Next i
  
   vaSpread2.MaxRows = 0
   vaSpread2.Row = -1: vaSpread2.Col = -1
   vaSpread2.BackColor = &HC0FFFF
   
  ProgressBar1.Scrolling = ccScrollingSmooth
  ProgressBar1.Max = 100
  ProgressBar1.Visible = True
  ProgressBar1.Value = 0
  lbl_proceso.Caption = "0 %"
  lbl_proceso.Visible = True
  
  Toolbar2.Enabled = False
  FpFecDesde.Enabled = False
  FpFecHasta.Enabled = False
  Command1.Enabled = False

  j = 1
  fg_carga ""
  
  '-------> Inicio LLenar grilla
  vg_Archxls = fg_ArchivoTxt
  Open vg_Archxls For Output As #1
  
  For i = 1 To vaSpread1.MaxRows
      
      vaSpread1.Row = i
      vaSpread1.Col = 1
      seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
      If seleccion = 1 And vaSpread1.RowHidden = False Then

         Ceco = ""
         vaSpread1.Col = 2
         Ceco = vaSpread1.text
          
         Sql = ""
         Sql = Sql & " '" & Ceco & "', "
         Sql = Sql & " '" & Format(FpFecDesde, ("YYYYmmdd")) & "' ,"
         Sql = Sql & " '" & Format(FpFecHasta, ("YYYYmmdd")) & "' ,"
         Sql = Sql & " '1' "
         Set RS = vg_db.Execute("sgpadm_Sel_CostoMinutaBloquexServicios_V03 " & Sql & "")
         Do While Not RS.EOF
      
            vaSpread2.MaxRows = vaSpread2.MaxRows + 1
            vaSpread2.Row = vaSpread2.MaxRows
      
            If vaSpread2.MaxRows = 1 Then
               
               Print #1, "Periodo : " & Format(FpFecDesde, ("dd/mm/yyyy")) & " - " & Format(FpFecHasta, ("dd/mm/yyyy"))
               Print #1, ""
               Print #1, ""
'               Print #1, "Ceco" & "|" & "Descripcion" & "|" & "Bloque Minuta" & "|" & "Regimen" & "|" & "Servicio" & "|" & "Costo Plato" & "|" & "Comensales Totales"
               Print #1, "Ceco" & "|" & "Descripcion" & "|" & "Regimen" & "|" & "Servicio" & "|" & "Costo Plato" & "|" & "Comensales Totales"
            End If
            
            vaSpread2.Col = 1
            vaSpread2.text = "0"
      
            vaSpread2.Col = 2
            vaSpread2.text = RS!min_cecori
      
            vaSpread2.Col = 3
            vaSpread2.text = Trim(RS!Cli_nombre)
      
'            vaSpread2.Col = 4
'            vaSpread2.text = RS!Id_Bloque
      
            vaSpread2.Col = 5
            vaSpread2.text = RS!min_codreg & " - " & Trim(RS!reg_nombre)
      
            vaSpread2.Col = 6
            vaSpread2.text = RS!min_codser & " - " & Trim(RS!ser_nombre)
      
            vaSpread2.Col = 7
            vaSpread2.text = RS!min_codreg
         
            vaSpread2.Col = 8
            vaSpread2.text = RS!min_codser
             
            vaSpread2.Col = 9
            vaSpread2.text = ""
             
            CostoPlato = 0
            If RS!raciontotal > 0 Then
                
               vaSpread2.text = Format((RS!CostoReceta / RS!raciontotal), fg_Pict(6, 2))
            
               CostoPlato = Format((RS!CostoReceta / RS!raciontotal), fg_Pict(6, 2))
             
            End If
             
            vaSpread2.Col = 10
            vaSpread2.text = Format(RS!raciontotal, fg_Pict(6, 2))
             
            '-------> Print detalle pedido
'            Print #1, RS!min_cecori & "|" & Trim(RS!cli_nombre) & "|" & RS!Id_Bloque & "|" & RS!min_codreg & " - " & Trim(RS!reg_nombre) & "|" & RS!min_codser & " - " & Trim(RS!ser_nombre) & "|" & CostoPlato & "|" & vaSpread2.text
            Print #1, RS!min_cecori & "|" & Trim(RS!Cli_nombre) & "|" & RS!min_codreg & " - " & Trim(RS!reg_nombre) & "|" & RS!min_codser & " - " & Trim(RS!ser_nombre) & "|" & CostoPlato & "|" & vaSpread2.text
            RS.MoveNext
          
         Loop
         RS.Close
         Set RS = Nothing
          
         ProgressBar1.Value = ((j / Conta) * 100)
         lbl_proceso.Caption = CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %"
         j = j + 1
      
      End If
       
      DoEvents
       
  Next i

  Close #1
  fg_descarga
  MsgBox "Proceso Finalizado", vbInformation, Me.Caption
                
  ProgressBar1.Visible = False
  lbl_proceso.Visible = False
  
  Toolbar2.Enabled = True
  FpFecDesde.Enabled = True
  FpFecHasta.Enabled = True
  Command1.Enabled = True
  
End Select

Exit Sub
Man_Error:

    Close #1
    ProgressBar1.Visible = False
    lbl_proceso.Visible = False
  
    Toolbar2.Enabled = True
    FpFecDesde.Enabled = True
    FpFecHasta.Enabled = True
  
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
On Error GoTo Man_Error

Select Case BlockCol

Case 1
    
    Dim i As Long
    vaSpread1.Col = 1
    For i = BlockRow To BlockRow2
        vaSpread1.Row = i
        vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    Next

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

