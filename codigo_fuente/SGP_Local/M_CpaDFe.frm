VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_CpaDFe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Días Feriados"
   ClientHeight    =   6795
   ClientLeft      =   1290
   ClientTop       =   2070
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Datos Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   8775
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   4680
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4095
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   8535
         _Version        =   393216
         _ExtentX        =   15055
         _ExtentY        =   7223
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
         MaxCols         =   3
         SpreadDesigner  =   "M_CpaDFe.frx":0000
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   240
         TabIndex        =   13
         Top             =   4560
         Width           =   1545
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   14
            Top             =   135
            Width           =   1440
         End
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   1800
         TabIndex        =   11
         Top             =   4560
         Width           =   4125
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   12
            Top             =   135
            Width           =   4020
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8775
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   660
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
         Text            =   "2021"
         DateCalcMethod  =   4
         DateTimeFormat  =   5
         UserDefinedFormat=   "yyyy"
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   1080
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2775
         TabIndex        =   7
         Top             =   315
         Width           =   5655
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2310
         Picture         =   "M_CpaDFe.frx":188A
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Año"
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
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Casino"
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
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2805
         TabIndex        =   8
         Top             =   360
         Width           =   5655
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6795
      Left            =   9060
      TabIndex        =   4
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   11986
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_CpaDFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim MsgTitulo As String
Dim est As Boolean

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Dim codTippla As Long, nomTippla As String
fg_centra Me
fg_carga ""
est = True
MsgTitulo = "Copiar Días Feriados"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar ": BtnX.Enabled = True
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpDateTime1(0).text = Format(Date, "yyyy")
vaSpread1.MaxRows = 0
fg_descarga
est = False
End Sub

Private Sub fpDateTime1_ButtonHit(Index As Integer, Button As Integer, NewIndex As Integer)
If est Then Exit Sub
If Trim(fpDateTime1(0).text) = "" Then Exit Sub
If Not IsDate(fpDateTime1(0).text) Then Exit Sub
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If est Then Exit Sub
If Trim(fpDateTime1(0).text) = "" Then Exit Sub
If Not IsDate(fpDateTime1(0).text) Then Exit Sub
End Sub

Private Sub fpDateTime1_ChangeMode(Index As Integer, EditMode As Integer)
If est Then Exit Sub
If Trim(fpDateTime1(0).text) = "" Then Exit Sub
If Not IsDate(fpDateTime1(0).text) Then Exit Sub
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpDateTime1_PopUp(Index As Integer, Cancel As Integer)
If est Then Exit Sub
If Trim(fpDateTime1(0).text) = "" Then Exit Sub
If Not IsDate(fpDateTime1(0).text) Then Exit Sub
End Sub

Private Sub fpText_Change(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    Set RS = vg_db.Execute("sgp_s_cliente 1, '" & fpText(0).text & "',''")
    fpayuda(0).Caption = ""
    If Not RS.EOF Then fpayuda(0).Caption = Trim(RS!cli_nombre)
    RS.Close: Set RS = Nothing
    MoverDatosGrilla
End Select
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Image1_Click(Index As Integer)
vg_left = fpayuda(0).Left + 2300
vg_nombre = "": vg_codigo = ""
Select Case Index
Case 0
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(0).text = vg_codigo
    fpayuda(0).Caption = Trim(vg_nombre)
End Select
End Sub

Private Sub MoverDatosGrilla()
vaSpread1.MaxRows = 0
vaSpread1.Visible = False
Set RS = vg_db.Execute("sgp_s_cliente 14, '" & fpText(0).text & "', ''")
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1: vaSpread1.text = ""
   vaSpread1.Col = 2: vaSpread1.text = IIf(IsNull(RS!cli_codigo), "", RS!cli_codigo)
   vaSpread1.Col = 3: vaSpread1.text = IIf(IsNull(RS!cli_nombre), "", Trim(RS!cli_nombre))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 1, 2
    vaSpread1.Visible = False
    If Trim(Text1(Index).text) <> "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = IIf(Index = 1, 2, 3)
           IndActivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread1.Col = 2
           If IndActivo = -1 And Trim(vaSpread1.text) <> "" Then
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           Else
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           End If
        Next i
        vaSpread1.SetActiveCell IIf(Index = 1, 2, 3), 1
    End If
'    vaSpread1_Click Index, 0
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
    If Trim(Text1(Index).text) = "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Next
       vaSpread1.SetActiveCell IIf(Index = 1, 2, 3), vaSpread1.SearchCol(IIf(Index = 1, 2, 3), 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell IIf(Index = 1, 2, 3), 1
    End If
    vaSpread1.Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim isel As Boolean, i As Long, j As Long, cencos As String
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    If Trim(fpayuda(0).Caption) = "" Then MsgBox "Debe selecionar casino origen", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    If Trim(fpDateTime1(0).text) = "" Then MsgBox "Debe selecionar año origen", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    isel = False
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then isel = True: Exit For
    Next i
    If isel = False Then MsgBox "Debe Seleccionar A lo menor un centro costo destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Set RS = vg_db.Execute("sgpadm_s_diasferiados 3, '" & Trim(fpText(0).text) & "', '','',''")
    If RS.EOF Then MsgBox "No existe información a copiar", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    fg_carga ""
    Frame3.Visible = False: Frame4.Visible = False
    Text1(1).Visible = False: Text1(2).Visible = False
    Toolbar1.Enabled = False
    Bar1(0).Visible = True: Bar1(0).Value = 0
    For i = 1 To vaSpread1.MaxRows
        DoEvents
        Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then
           vaSpread1.SetActiveCell 2, vaSpread1.Row
           vaSpread1.Col = 2: cencos = vaSpread1.text
           vg_db.Execute "DELETE Cas_b_Fecha_Inhabiles WHERE CFI_CeCo = '" & cencos & "' AND YEAR(CFI_Fecha) = '" & Trim(fpDateTime1(0).text) & "'"
           vg_db.Execute "INSERT INTO Cas_b_Fecha_Inhabiles (CFI_CeCo, CFI_Fecha, CFI_Glosa) SELECT '" & cencos & "', CFI_Fecha, CFI_Glosa FROM Cas_b_Fecha_Inhabiles WHERE CFI_CeCo = '" & Trim(fpText(0).text) & "'  AND YEAR(CFI_Fecha) = '" & Trim(fpDateTime1(0).text) & "'"
        End If
    Next i
    Bar1(0).Visible = False: Bar1(0).Value = 0
    fg_descarga
    MsgBox "Generación copia Finalizado Sin Problema", vbInformation + vbOKOnly, MsgTitulo
    Frame3.Visible = True: Frame4.Visible = True
    Text1(1).Visible = True: Text1(2).Visible = True
    Toolbar1.Enabled = True
Case 3
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
fg_descarga
Toolbar1.Enabled = True
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If Col = 1 And Row = 0 Then vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.text = IIf(vaSpread1.Value = "1", "0", "1")
End Sub

