VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_GasA13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gastos A13"
   ClientHeight    =   5925
   ClientLeft      =   2400
   ClientTop       =   3720
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   1665
      TabIndex        =   5
      Top             =   435
      Width           =   7965
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   1545
         TabIndex        =   0
         Top             =   210
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
         Left            =   1545
         TabIndex        =   1
         Top             =   570
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
      Begin VB.Label Label2 
         Caption         =   "(mm/aaaa)"
         Height          =   225
         Left            =   2535
         TabIndex        =   10
         Top             =   630
         Width           =   840
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
         Left            =   315
         TabIndex        =   9
         Top             =   585
         Width           =   540
      End
      Begin VB.Label Label3 
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
         Index           =   6
         Left            =   315
         TabIndex        =   8
         Top             =   225
         Width           =   735
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3390
         TabIndex        =   6
         Top             =   210
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2940
         Picture         =   "M_GasA13.frx":0000
         Top             =   105
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3435
         TabIndex        =   7
         Top             =   255
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   120
      TabIndex        =   4
      Top             =   1515
      Width           =   11130
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3615
         Left            =   165
         TabIndex        =   3
         Top             =   240
         Width           =   10770
         _Version        =   393216
         _ExtentX        =   18997
         _ExtentY        =   6376
         _StockProps     =   64
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
         MaxCols         =   7
         MaxRows         =   20
         ScrollBars      =   2
         SpreadDesigner  =   "M_GasA13.frx":030A
         VisibleCols     =   2
         VisibleRows     =   15
         ScrollBarTrack  =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Gastos Generales"
         Height          =   195
         Index           =   0
         Left            =   2940
         TabIndex        =   12
         Top             =   3960
         Width           =   1260
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   2535
         Top             =   3975
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H000000C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   255
         Top             =   3975
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Registros de Sistema"
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   11
         Top             =   3960
         Width           =   1485
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_GasA13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim posaux As Long, ctaaux As String
Dim Msgtitulo As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 6405
Me.Width = 11415
Msgtitulo = "Gastos A13"
fg_centra Me
modo = ""
Dim X As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
Gl_Mo_Botones Me, 9
Gl_Ac_Botones Me, 9, 1, modo
With vaSpread1
    .Row = -1
    '.col = 1: .TypeNumberSeparator = vg_CSep: .TypeNumberDecimal = vg_CDec: .TypeNumberDecPlaces = vg_DPr
    .Col = 3: .TypeNumberSeparator = vg_CSep: .TypeNumberDecimal = vg_CDec: .TypeNumberDecPlaces = vg_DPr
    .Col = 4: .TypeNumberSeparator = vg_CSep: .TypeNumberDecimal = vg_CDec: .TypeNumberDecPlaces = vg_DPr
    '.col = 5: .TypeNumberSeparator = vg_CSep: .TypeNumberDecimal = vg_CDec: .TypeNumberDecPlaces = vg_DPr
End With
fpText1(1).Enabled = ModCasino
Image1(1).Enabled = ModCasino
fpText1(1).text = MuestraCasino(1)
fpayuda(1).Caption = MuestraCasino(2)
MoverDatosGrilla
End Sub

Sub MoverDatosGrilla()
Dim i As Long, anomes As Long
On Error GoTo Man_Error
With vaSpread1
    .Visible = False
    .MaxRows = 0
    anomes = Val(fpDateTime1.Year & Right("00" & fpDateTime1.Month, 2))
    Dim lisnom As String, estado As String, codaux As Long, z As Long, ctacon As String
    lisnom = ""
    '-------> Cargar Cuentas Contable que sean diferente alimentanción y Desechables.
    RS1.Open "SELECT cta_codigo, cta_nombre FROM a_ctacontable WHERE cta_codigo NOT IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') AND cta_codigo NOT IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')", vg_db, adOpenStatic
    Do While Not RS1.EOF
       lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(RS1!cta_codigo) & " - " & Trim(RS1!cta_nombre)
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    '-------> cargar Otros Costo A13
    RS1.Open "SELECT * FROM b_gastosa13 WHERE gas_cencos = '" & Trim(fpText1(1).text) & "' AND gas_anomes = " & anomes & " ORDER BY gas_orden, gas_descri", vg_db, adOpenStatic
    If Not RS1.EOF Then
       i = 1
       Do While Not RS1.EOF
          .MaxRows = i: .Row = i
          If RS1!gas_codigo >= 1 And RS1!gas_codigo <= 8 Then
             .Col = 2: .Lock = True: .ForeColor = Shape1(0).FillColor
             .Col = 6: .Lock = True
             .Col = 7: .text = "S"
          Else
             .Col = 2: .Lock = False: .ForeColor = Shape1(1).FillColor
             .Col = 6: .Lock = False
             .Col = 7: .text = "U"
             .Col = 5
             .CellType = CellTypeComboBox: .TypeComboBoxList = lisnom
                
             For z = 0 To .TypeComboBoxCount
                 .TypeComboBoxCurSel = z
                 If Trim(.text) <> "" Then
                    ctacon = Trim(Mid(.text, 1, InStr(1, .text, "-") - 1))
                 Else
                    ctacon = ""
                 End If
                 If ctacon = RS1!gas_ctacon Then codaux = z: Exit For
                 codaux = -1
             Next z
             .TypeComboBoxCurSel = codaux
          End If
          .Col = 1: .text = RS1!gas_codigo
          .Col = 2: .text = IIf(IsNull(RS1!gas_descri), "", RS1!gas_descri)
          .Col = 3: .text = IIf(IsNull(RS1!gas_valor), 0, RS1!gas_valor)
          .Col = 4: .text = IIf(IsNull(RS1!gas_valpro), 0, RS1!gas_valpro)
          .Col = 6: .text = RS1!gas_orden
          RS1.MoveNext: i = i + 1
       Loop
    Else
        .MaxRows = 8
        .Row = 1
        .Col = 1: .text = .Row
        .Col = 2: .text = "Depreciación"
        .Col = 3: .text = 0
        .Col = 4: .text = 0
        .Col = 6: .text = .Row
        .Col = 7: .text = "S"
        .Row = 2
        .Col = 1: .text = .Row
        .Col = 2: .text = "Gestión Personal"
        .Col = 3: .text = 0
        .Col = 4: .text = 0
        .Col = 6: .text = .Row
        .Col = 7: .text = "S"
        .Row = 3
        .Col = 1: .text = .Row
        .Col = 2: .text = "Cuota Negociación"
        .Col = 3: .text = 0
        .Col = 4: .text = 0
        .Col = 6: .text = .Row
        .Col = 7: .text = "S"
        .Row = 4
        .Col = 1: .text = .Row
        .Col = 2: .text = "Bono vac. (Bienestar)"
        .Col = 3: .text = 0
        .Col = 4: .text = 0
        .Col = 6: .text = .Row
        .Col = 7: .text = "S"
        .Row = 5
        .Col = 1: .text = .Row
        .Col = 2: .text = "Cuota Dirigente Sindical"
        .Col = 3: .text = 0
        .Col = 4: .text = 0
        .Col = 6: .text = .Row
        .Col = 7: .text = "S"
        .Row = 6
        .Col = 1: .text = .Row
        .Col = 2: .text = "Nº Horas Extra"
        .Col = 3: .text = 0
        .Col = 4: .text = 0
        .Col = 6: .text = .Row
        .Col = 7: .text = "S"
        .Row = 7
        .Col = 1: .text = .Row
        .Col = 2: .text = "Nº Dias Trabajados"
        .Col = 3: .text = 0
        .Col = 4: .text = 0
        .Col = 6: .text = .Row
        .Col = 7: .text = "S"
        .Row = 8
        .Col = 1: .text = .Row
        .Col = 2: .text = "Nº Dias Ausencia"
        .Col = 3: .text = 0
        .Col = 4: .text = 0
        .Col = 6: .text = .Row
        .Col = 7: .text = "S"
        
        .Col = 2: .Row = -1: .Lock = True: .ForeColor = Shape1(0).FillColor
        .Col = 6: .Lock = True
    End If
    .Col = 3: .Row = -1: .Lock = False
    .Col = 4: .Row = -1: .Lock = False
    RS1.Close: Set RS1 = Nothing
    .Visible = True
End With
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub fpDateTime1_Change()
MoverDatosGrilla
modo = "": Gl_Ac_Botones Me, 9, 1, modo
End Sub

Private Sub fpText1_Change(Index As Integer)
fpayuda(1).Caption = ""
RS1.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo = '" & fpText1(1).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(1).Caption = "": Exit Sub
fpayuda(1).Caption = Trim(RS1!cli_nombre)
RS1.Close: Set RS1 = Nothing
MoverDatosGrilla
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub


Private Sub Image1_Click(Index As Integer)
vg_codigo = ""
Select Case Index
Case 1
    vg_left = fpayuda(Index).Left + 1920
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    fpText1(Index) = Trim(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    If fpText1(1).Enabled = True Then fpText1(1).SetFocus
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codigo As Long, i As Long, cencos As String, anomes As Long, descri As String, valor As Double, Orden As Long, estado As String
Dim ctacon As String, sql As String, lisnom As String
Dim valpro As Double
On Error GoTo Man_Error
Select Case Button.Index
Case 1 'INCLUIR
    fpayuda(1).Caption = ""
    RS1.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo = '" & fpText1(1).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(1).Caption = "": Exit Sub
    RS1.Close: Set RS1 = Nothing
    MoverDatosGrilla
    modo = "A"
    Gl_Ac_Botones Me, 9, 0, modo
    With vaSpread1
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows: .Col = -1: .Lock = False
        .ForeColor = Shape1(1).FillColor
        lisnom = ""
        RS1.Open "SELECT cta_codigo, cta_nombre FROM a_ctacontable WHERE cta_codigo NOT IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') and cta_codigo NOT IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')", vg_db, adOpenStatic
        Do While Not RS1.EOF
            lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(RS1!cta_codigo) & " - " & Trim(RS1!cta_nombre)
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
        .Col = 5
        .CellType = CellTypeComboBox: .TypeComboBoxList = lisnom
        .Col = 2: .SetActiveCell 1, .MaxRows
    End With
Case 3 'ALTERAR
    fpayuda(1).Caption = ""
    RS1.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo = '" & fpText1(1).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(1).Caption = "": Exit Sub
    RS1.Close: Set RS1 = Nothing
    MoverDatosGrilla
    modo = "M"
    Gl_Ac_Botones Me, 9, 0, modo
Case 5 'BORRAR
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    vaSpread1.GetInteger 1, vaSpread1.ActiveRow, codigo
    If codigo >= 1 And codigo <= 8 Then MsgBox "Registro de Sistema...", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
    anomes = Val(fpDateTime1.Year & Right("00" & fpDateTime1.Month, 2))
    vg_db.BeginTrans
    vg_db.Execute "DELETE b_gastosa13 FROM b_gastosa13 WHERE gas_cencos = '" & Trim(fpText1(1).text) & "' " & _
                  "AND gas_anomes = " & anomes & " AND gas_codigo = " & codigo
    vg_db.CommitTrans
    MoverDatosGrilla
    modo = "": Gl_Ac_Botones Me, 9, 1, modo
Case 7 'ACTUALIZAR
    MoverDatosGrilla
Case 10 'CANCELAR
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    MoverDatosGrilla
    modo = "": Gl_Ac_Botones Me, 9, 1, modo
Case 12 'CONFIRMAR
    With vaSpread1
        For i = 1 To .MaxRows
            .Row = i: .Col = 3
            'If Val(.Value) = 0 Then MsgBox "No pueden haber valores en cero...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
            .Col = 2
            If Trim(.text) = "" Then MsgBox "Debe ingresar información...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
            .Col = 5
            If Trim(.text) = "" And .CellType = CellTypeComboBox Then MsgBox "Debe ingresar información...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        Next i
        
        cencos = Trim(fpText1(1).text)
        anomes = Val(fpDateTime1.Year & Right("00" & fpDateTime1.Month, 2))
        vg_db.BeginTrans
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1: codigo = Val(.text)
            .Col = 2: descri = .text
            .Col = 3: valor = Val(.Value)
            .Col = 4: valpro = Val(.Value)
            .Col = 5
            If Trim(.text) <> "" Then
                ctacon = Trim(Mid(.text, 1, InStr(1, .text, "-") - 1))
            Else
                ctacon = ""
            End If
            .Col = 6: Orden = i
            .Col = 7: estado = .text
            RS1.Open "SELECT * FROM b_gastosa13 WHERE gas_cencos = '" & Trim(fpText1(1).text) & "' " & _
                     "AND gas_anomes = " & anomes & " AND gas_codigo = " & codigo, vg_db, adOpenStatic
            If Not RS1.EOF Then
                vg_db.Execute "UPDATE b_gastosa13 SET " & _
                              "gas_descri = '" & descri & "', gas_valor = " & valor & ", " & _
                              "gas_orden = " & Orden & ", gas_ctacon = '" & ctacon & "', gas_valpro = " & valpro & " WHERE gas_cencos = '" & Trim(fpText1(1).text) & "' " & _
                              "AND gas_anomes = " & anomes & " AND gas_codigo = " & codigo
            Else
                If estado = "S" Then
                    .Col = 1: codigo = .text
                Else
                    RS2.Open "SELECT MAX(gas_codigo) AS codigo FROM b_gastosa13 WHERE gas_cencos = '" & Trim(fpText1(1).text) & "' " & _
                             "AND gas_anomes = " & anomes, vg_db, adOpenStatic
                    If Not RS2.EOF Then codigo = IIf(IsNull(RS2!codigo) Or RS2!codigo < 8, 9, RS2!codigo + 1)
                    RS2.Close: Set RS2 = Nothing
                End If
                vg_db.Execute "INSERT INTO b_gastosa13 (gas_cencos, gas_anomes, gas_codigo, gas_descri, gas_valor, gas_orden, gas_ctacon, gas_valpro) " & _
                              "VALUES ('" & cencos & "', " & anomes & ", " & codigo & ", '" & descri & "', " & valor & ", " & Orden & ", '" & ctacon & "', " & valpro & ")"
                .Col = 1: .text = codigo
            End If
            .Col = 6: .text = Orden
            If estado <> "S" Then .Col = 7: .text = "U"
            RS1.Close: Set RS1 = Nothing
        Next i
        vg_db.CommitTrans
    End With
    MsgBox "Los datos fueron grabados...", vbInformation + vbOKOnly, Msgtitulo
    modo = "": Gl_Ac_Botones Me, 9, 1, modo
Case 15 'COPIAR OTROS COSTO A13
    anomes = Val(fpDateTime1.Year & Right("00" & fpDateTime1.Month, 2))
    RS1.Open "SELECT gas.*, cta.cta_codigo, cta.cta_nombre FROM a_ctacontable cta " & _
             "RIGHT JOIN b_gastosa13 gas On cta.cta_codigo = gas.gas_ctacon " & _
             "WHERE gas_cencos = '" & Trim(fpText1(1).text) & "' AND gas_anomes = " & anomes & "", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "No existe información a grabar...", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
    RS1.Close: Set RS1 = Nothing
    M_CoGA13.LlenarDatos fpText1(1).text, Format(fpDateTime1.text, "yyyymm")
    M_CoGA13.Show 1
Case 17 'SUBIR
    With vaSpread1
        If .ActiveRow = 1 Then Exit Sub
        .Visible = False
        .MaxRows = .MaxRows + 1
        .MoveRange 1, .ActiveRow - 1, .MaxCols, .ActiveRow - 1, 1, .MaxRows
        .MoveRange 1, .ActiveRow, .MaxCols, .ActiveRow, 1, .ActiveRow - 1
        .MoveRange 1, .MaxRows, .MaxCols, .MaxRows, 1, .ActiveRow
        .MaxRows = .MaxRows - 1
        .SetActiveCell .ActiveCol, .ActiveRow - 1
        For i = 1 To .MaxRows
            .Row = i: .Col = 7: estado = .text
            .Col = 2: .ForeColor = IIf(estado = "S", Shape1(0).FillColor, Shape1(1).FillColor)
        Next i
        .Visible = True
    End With
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 9, 0, modo
    Toolbar1.Buttons(17).Enabled = True: Toolbar1.Buttons(18).Enabled = True
Case 18 'BAJAR
    With vaSpread1
        If .ActiveRow = .MaxRows Then Exit Sub
        .Visible = False
        .MaxRows = .MaxRows + 1
        .MoveRange 1, .ActiveRow + 1, .MaxCols, .ActiveRow + 1, 1, .MaxRows
        .MoveRange 1, .ActiveRow, .MaxCols, .ActiveRow, 1, .ActiveRow + 1
        .MoveRange 1, .MaxRows, .MaxCols, .MaxRows, 1, .ActiveRow
        .MaxRows = .MaxRows - 1
        .SetActiveCell .ActiveCol, .ActiveRow + 1
        For i = 1 To .MaxRows
            .Row = i: .Col = 7: estado = .text
            .Col = 2: .ForeColor = IIf(estado = "S", Shape1(0).FillColor, Shape1(1).FillColor)
        Next i
        .Visible = True
    End With
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 9, 0, modo
    Toolbar1.Buttons(17).Enabled = True: Toolbar1.Buttons(18).Enabled = True
Case 20 'IMPRIMIR
    anomes = Val(fpDateTime1.Year & Right("00" & fpDateTime1.Month, 2))
    sql = "SELECT gas.*, cta.cta_codigo, cta.cta_nombre FROM a_ctacontable cta " & _
          "RIGHT JOIN b_gastosa13 gas On cta.cta_codigo = gas.gas_ctacon " & _
          "WHERE gas_cencos = '" & Trim(fpText1(1).text) & "' AND gas_anomes = " & anomes & " Order by gas_orden"
    I_GastosA13 Me, sql
Case 23 'SALIR
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_ComboDropDown(ByVal Col As Long, ByVal Row As Long)
vaSpread1.Row = Row: vaSpread1.Col = Col
If Trim(vaSpread1.text) <> "" Then
    ctaaux = Trim(Mid(vaSpread1.text, 1, InStr(1, vaSpread1.text, "-") - 1))
Else
    ctaaux = ""
End If
posaux = vaSpread1.TypeComboBoxCurSel
End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim z As Long, ctacon As String, lc_ctaaux As String
'jpaz 30-08-2005 vaSpread1.Row = Row: vaSpread1.col = col
'jpaz 30-08-2005 If Trim(vaSpread1.Text) <> "" Then
'jpaz 30-08-2005    lc_ctaaux = Trim(Mid(vaSpread1.Text, 1, InStr(1, vaSpread1.Text, "-") - 1))
'jpaz 30-08-2005 Else
'jpaz 30-08-2005    lc_ctaaux = ""
'jpaz 30-08-2005 End If
'jpaz 30-08-2005 For z = 1 To vaSpread1.MaxRows
'jpaz 30-08-2005    vaSpread1.Row = z: vaSpread1.col = 6
'jpaz 30-08-2005    If vaSpread1.Text <> "S" And z <> Row Then
'jpaz 30-08-2005        vaSpread1.col = col
'jpaz 30-08-2005        If Trim(vaSpread1.Text) <> "" Then
'jpaz 30-08-2005            ctacon = Trim(Mid(vaSpread1.Text, 1, InStr(1, vaSpread1.Text, "-") - 1))
'jpaz 30-08-2005        Else
'jpaz 30-08-2005            ctacon = ""
'jpaz 30-08-2005        End If
'jpaz 30-08-2005        If ctacon = lc_ctaaux Then
'jpaz 30-08-2005            MsgBox "La cuenta ya esta siendo ocupada...", vbExclamation + vbOKOnly, MsgTitulo
'jpaz 30-08-2005            vaSpread1.Row = Row
'jpaz 30-08-2005            vaSpread1.TypeComboBoxCurSel = posaux
'jpaz 30-08-2005            Exit For
'jpaz 30-08-2005        End If
'jpaz 30-08-2005    End If
'jpaz 30-08-2005 Next z
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 9, 0, modo
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 9, 0, modo
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If Row <> NewRow And NewRow > 0 And (modo = "A") And Toolbar1.Buttons(17).Enabled = False Then
    Cancel = True
End If
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If Row = 0 Then Exit Sub
Select Case Col
Case 5
    vaSpread1.Row = Row: vaSpread1.Col = Col
    If vaSpread1.ColWidth(Col) > (vaSpread1.MaxTextCellWidth - 2) Then Exit Sub
    TipWidth = vaSpread1.MaxTextColWidth(Col)
    ShowTip = True
    MultiLine = 2
    TipText = vaSpread1.text
End Select
End Sub
