VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_MerPed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mermas por Período"
   ClientHeight    =   3450
   ClientLeft      =   1950
   ClientTop       =   3345
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   0
      TabIndex        =   8
      Top             =   375
      Width           =   8745
      Begin VB.OptionButton OptTipMer 
         Caption         =   "Resumido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   5250
         TabIndex        =   19
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton OptTipMer 
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   360
         TabIndex        =   18
         Top             =   1440
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1995
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   3885
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Tipo de Merma"
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
         Height          =   1260
         Left            =   255
         TabIndex        =   9
         Top             =   1590
         Width           =   4740
         Begin VB.OptionButton OptTipMer 
            Caption         =   "Una"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   390
            TabIndex        =   4
            Top             =   360
            Width           =   1395
         End
         Begin VB.OptionButton OptTipMer 
            Caption         =   "Todas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   2610
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   660
            Width           =   3885
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   465
            TabIndex        =   10
            Top             =   735
            Width           =   3885
         End
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   1980
         TabIndex        =   0
         Top             =   240
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
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
         ButtonStyle     =   0
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
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         CharValidationText=   ""
         MaxLength       =   255
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
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   2010
         TabIndex        =   2
         Top             =   1000
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
         Text            =   "17/08/2004"
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
         Left            =   6690
         TabIndex        =   3
         Top             =   1005
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
         Text            =   "17/08/2004"
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
         AutoSize        =   -1  'True
         Caption         =   "Bodega"
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
         Left            =   375
         TabIndex        =   17
         Top             =   705
         Width           =   660
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   2040
         TabIndex        =   16
         Top             =   660
         Width           =   3885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Termino"
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
         Left            =   5250
         TabIndex        =   13
         Top             =   1095
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
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
         Left            =   360
         TabIndex        =   12
         Top             =   1095
         Width           =   1065
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3795
         TabIndex        =   7
         Top             =   255
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3345
         Picture         =   "I_MerPed.frx":0000
         Top             =   150
         Width           =   480
      End
      Begin VB.Label Label1 
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
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   350
         Width           =   1200
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3825
         TabIndex        =   15
         Top             =   300
         Width           =   4335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_MerPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim MsgTitulo As String
Public lc_Aux As String

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_Change(Index As Integer)
fpayuda(1).Caption = ""
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_LostFocus(Index As Integer)
If fpText1(1).text = "" Then Exit Sub

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText1(1).text)), ""), vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
       fpayuda(Index).Caption = RS1!cli_nombre
       RS1.MoveNext
    Loop
Else
    RS1.Close: Set RS1 = Nothing
    MsgBox "Contrato no existe...", vbExclamation + vbOKOnly, MsgTitulo
    fpText1(1).text = ""
    Exit Sub
End If
RS1.Close: Set RS1 = Nothing
End Sub

Private Sub Form_Activate()
On Local Error GoTo Error_Descarga
fg_descarga
Exit Sub
Error_Descarga:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Resume Next
End Sub

Private Sub Form_Load()
On Local Error GoTo Error_CargaMer
Me.Width = 8880
Me.Height = 3780
EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)
If lc_Aux = "MerPed" Then
    MsgTitulo = "Informe de Mermas por Período"
    Me.Caption = "Informe de Mermas por Período"
    Me.Height = 3930
    Frame2.Height = 3000
    Frame1.Visible = True
    OptTipMer(2).Visible = False
    OptTipMer(3).Visible = False
    Combo1(0).Enabled = False
ElseIf lc_Aux = "AjuInv" Then
    MsgTitulo = "Informe Ajuste Inventario"
    Me.Caption = "Informe Ajuste Inventario"
    Me.Height = 2700
    Frame2.Height = 1740
    OptTipMer(2).Visible = True
    OptTipMer(3).Visible = True
    Frame1.Visible = False
    Combo1(0).Enabled = False
End If

fg_centra Me
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
'----------------------------Valida permisos para impresión
Toolbar1.Buttons.Item(1).Visible = IIf(Val(Mid(ValidarUsuario(Me), 4, 1)) = 1, True, False)
'-------------------------Asigna el tipo de informe-------------
'-------> Cargar Combo Bodega
CargarDatoCombo Combo1, 0, "b_clientes", "cli_", "CliBod", "N"
'-------> Cargar Combo Tipo Ajuste
CargarDatoCombo Combo1, 1, "a_tipoajuste", "aju_", "TipAju2", "NM"
Combo1(0).ListIndex = 0
'-------------------------Fin Asigna el tipo de informe---------
'-------------------------Asigna fecha actual del sistema para informe-------------
fpDateTime1(0).text = Date: fpDateTime1(1).text = Date
fpText1(1).Enabled = ModCasino
Image1(1).Enabled = ModCasino
fpText1(1).text = MuestraCasino(1)
fpayuda(1).Caption = MuestraCasino(2)
Combo1(1).Enabled = False
OptTipMer(1).Value = True
Exit Sub
Error_CargaMer:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Resume Next
End Sub

Private Sub Image1_Click(Index As Integer)
On Local Error GoTo Error_Contrato
vg_codigo = 0
Select Case Index
Case 1
    vg_left = fpayuda(Index).Left + 1920
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    fpText1(Index) = Trim(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    fpText1_LostFocus 1
    If Combo1(0).Enabled = True Then Combo1(0).SetFocus
End Select
Exit Sub
Error_Contrato:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Resume Next
End Sub

Private Sub OptTipMer_Click(Index As Integer)
On Local Error GoTo Error_Opc
If OptTipMer(0).Value = True Then
    Combo1(1).Enabled = True
Else
    Combo1(1).Enabled = False
    Combo1(1).ListIndex = -1
End If
Exit Sub
Error_Opc:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Resume Next
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim cFecini As String, cFecter As String, codbod As Long, codmer As Long, cencos As String, TipMer As String
On Local Error GoTo Error_Boton
Select Case Button.Index
Case 1 '------- Imprimir
    If Combo1(0).ListIndex = -1 Then MsgBox "Seleccione bodega...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If OptTipMer(0).Value = True Then
        If Combo1(1).ListIndex = -1 Then
            MsgBox "Seleccione Tipo de Merma...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        End If
    End If
    cencos = Trim(fpText1(1).text) & "|" & Trim(fpayuda(1).Caption)
    codbod = Val(fg_codigocbo(Combo1, 0, 10, 0)): If codbod = 99999999 Then codbod = 0
    codmer = fg_codigocbo(Combo1, 1, 10, 0) 'Codigo de merma
    cFecini = Trim(fpDateTime1(0).text)
    cFecter = Trim(fpDateTime1(1).text)
    If OptTipMer(1).Value = True Then
       TipMer = "Todos los Tipos"
    ElseIf OptTipMer(0).Value = True Then
       TipMer = Trim(Mid(Combo1(1).text, 1, 100))
    End If
        
    If lc_Aux = "MerPed" Then
       I_MerPer cencos, cFecini, cFecter, codmer, codbod, TipMer
    ElseIf lc_Aux = "AjuInv" Then
       I_AjuInv Trim(fpText1(1).text), cFecini, cFecter, codbod, IIf(OptTipMer(2).Value = True, True, False)
    Else
    End If
Case 3 '------- Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Error_Boton:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Resume Next
End Sub
