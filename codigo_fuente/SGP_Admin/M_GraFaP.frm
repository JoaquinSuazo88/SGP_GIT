VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_GraFaP 
   Caption         =   "Gramaje Familia Producto 5 Etapas"
   ClientHeight    =   6360
   ClientLeft      =   2730
   ClientTop       =   2265
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1785
      Left            =   465
      TabIndex        =   4
      Top             =   435
      Width           =   7965
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   0
         Top             =   240
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "0"
         NegFormat       =   0
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   1
         Top             =   585
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "0"
         NegFormat       =   0
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2625
         Picture         =   "M_GraFaP.frx":0000
         Top             =   480
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   3075
         TabIndex        =   16
         Top             =   570
         Width           =   4335
      End
      Begin VB.Label Label3 
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
         Index           =   0
         Left            =   480
         TabIndex        =   15
         Top             =   660
         Width           =   750
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3075
         TabIndex        =   12
         Top             =   1275
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2625
         Picture         =   "M_GraFaP.frx":030A
         Top             =   1185
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2625
         Picture         =   "M_GraFaP.frx":0614
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "F. Producto"
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
         Index           =   5
         Left            =   480
         TabIndex        =   11
         Top             =   1350
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C. Dietetica"
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
         Left            =   495
         TabIndex        =   10
         Top             =   1005
         Width           =   1020
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3075
         TabIndex        =   9
         Top             =   930
         Width           =   4335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sub-Segmento"
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
         Left            =   480
         TabIndex        =   6
         Top             =   315
         Width           =   1245
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3075
         TabIndex        =   5
         Top             =   225
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2625
         Picture         =   "M_GraFaP.frx":091E
         Top             =   140
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3120
         TabIndex        =   7
         Top             =   270
         Width           =   4335
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3120
         TabIndex        =   13
         Top             =   975
         Width           =   4335
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3120
         TabIndex        =   14
         Top             =   1320
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   3120
         TabIndex        =   17
         Top             =   615
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   8715
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   8445
         _Version        =   393216
         _ExtentX        =   14896
         _ExtentY        =   6165
         _StockProps     =   64
         ColsFrozen      =   2
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
         MaxCols         =   4
         MaxRows         =   2
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_GraFaP.frx":0C28
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   330
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "M_GraFaP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim RS As New ADODB.Recordset
Dim est As Boolean
Dim Msgtitulo As String, FilCatDie As Long, filfampro As Long

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 6870
Me.Width = 9030
Msgtitulo = "Gramaje Familia Producto 5 Etapas"
fg_centra Me
modo = "": est = False
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = False

With vaSpread1
    .Col = -1: vaSpread1.Row = -1
    .BackColor = Shape1(1).FillColor
    .MaxRows = 0
    .Row = -1
    .Col = 1: .BackColor = Shape1(2).FillColor: .Col = 2: .BackColor = Shape1(2).FillColor
End With

FilCatDie = 0
End Sub

Sub MoverDatosGrilla()
If Val(fpLongInteger1(0).Value) = 0 Or Val(fpLongInteger1(1).Value) = 0 Or FilCatDie = 0 Or filfampro = 0 Then Exit Sub
Dim i As Long, codpre As Long, nompre As String
On Error GoTo Man_Error
fg_carga ""
With vaSpread1
    
    .Visible = False
    .MaxRows = 0
    codpre = 0
    Set RS = vg_db.Execute("sp_s_gramofamproducto " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & FilCatDie & ", " & filfampro & "")
    est = False
    If Not RS.EOF Then
       Do While Not RS.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1: .text = RS!tip_codigo
            If RS!tip_previo <> codpre Then
               codpre = RS!tip_previo
            End If
            
            .Col = 2: .text = Trim(RS!tip_nombre) 'fg_BuscaenArbol(RS!tip_codigo, "a_recetatippla", "tip_codigo")
            
            .Col = 3
            .CellType = CellTypeCurrency
            .TypeCurrencyDecPlaces = vg_RDCa
            .TypeFloatMin = "0"
            .TypeFloatMax = "99999999"
            .TypeFloatMoney = False
            .TypeFloatSeparator = True
            .TypeHAlign = TypeHAlignRight
            .TypeFloatCurrencyChar = Asc("$")
            .TypeFloatDecimalChar = Asc(".")
            .TypeFloatSepChar = Asc(",")
            .TypeCurrencyShowSymbol = False
            .text = Format(RS!graini, fg_Pict(6, vg_RDCa)) '2))
            
            .Col = 4
            .CellType = CellTypeCurrency
            .TypeCurrencyDecPlaces = vg_RDCa
            .TypeFloatMin = "0"
            .TypeFloatMax = "99999999"
            .TypeFloatMoney = False
            .TypeFloatSeparator = True
            .TypeHAlign = TypeHAlignRight
            .TypeFloatCurrencyChar = Asc("$")
            .TypeFloatDecimalChar = Asc(".")
            .TypeFloatSepChar = Asc(",")
            .TypeCurrencyShowSymbol = False
            .text = Format(RS!grafin, fg_Pict(6, vg_RDCa)) '2))
          
    '      vaSpread1.Col = 3: vaSpread1.text =
    '      vaSpread1.Col = 4: vaSpread1.text =
          RS.MoveNext
       Loop
       .SetActiveCell 3, 1
    End If
    RS.Close: Set RS = Nothing
    .Col = -1: vaSpread1.Row = -1
    .Lock = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1" Or Mid(ValidarUsuario(Me), 2, 1) = "1", False, True)
    fg_descarga
    est = True
    .Visible = True
    .SetFocus

End With

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Select Case Index
Case 0
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      Set RS = vg_db.Execute("SELECT * FROM a_subsegmento WHERE sub_codigo = " & Val(fpLongInteger1(0).Value) & " AND sub_indppr = '" & vg_Indppr & "'")
    Else
      Set RS = vg_db.Execute("SELECT * FROM a_subsegmento WHERE sub_codigo = " & Val(fpLongInteger1(0).Value) & "")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing
    MoverDatosGrilla
Case 1
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & " AND reg_indppr = '" & vg_Indppr & "'")
    Else
      Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & "")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(4).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    MoverDatosGrilla
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    If Index = 0 Then Image1_Click 0
End Select
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_subsegmento", "sub_", "Subsegmento", "Sub"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
Case 1
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(0).Left + 2400
    B_ArbEst.MoverDatosTvwDir "a_recetacatdie", "car_", "Categoria Dietetica"
    B_ArbEst.Show 1
    If vg_codigo = "" Then Exit Sub
    FilCatDie = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre: vg_nombre = ""
    MoverDatosGrilla
Case 2
   vg_left = fpayuda(1).Left + 1920
    B_ArbEst.MoverDatosTvwDir "a_tipopro", "tip_", "Familia del Producto"
    B_ArbEst.Show 1
    Me.Refresh
    If Val(vg_codigo) = 0 Then Exit Sub
    fpayuda(2).Caption = vg_nombre
    filfampro = Val(vg_codigo)
    MoverDatosGrilla
Case 3
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(4).Caption = vg_nombre
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 3
    If Val(fpLongInteger1(0).Value) = 0 Then Exit Sub
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Frame1.Enabled = False
Case 5
    If Val(fpLongInteger1(0).Value) = 0 Then Exit Sub
    If Not est < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    vg_db.Execute "DELETE b_gramofamproducto WHERE gfp_subseg=" & Val(fpLongInteger1(0).Value) & " AND gfp_codreg=" & Val(fpLongInteger1(1).Value) & " AND gfp_catdie=" & FilCatDie & " AND gfp_tiprec=" & Val(vaSpread1.text) & " AND gfp_fampro=" & filfampro & ""
    vaSpread1.Col = 3: vaSpread1.text = Format(0, fg_Pict(6, vg_RDCa))
    vaSpread1.Col = 4: vaSpread1.text = Format(0, fg_Pict(6, vg_RDCa))
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
Case 7
    MoverDatosGrilla
Case 10
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    MoverDatosGrilla
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
Case 12
    Dim codigo As Long, graini As Double, grafin As Double
    
    With vaSpread1
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1: codigo = .text
            .Col = 3: graini = .text
            .Col = 4: grafin = .text
            If (graini > 0 And grafin = 0) Or (graini = 0 And grafin > 0) Or (graini > grafin) Or (grafin < graini) Then MsgBox "Rango gramos no validos, proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        Next i
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1: codigo = .text
            .Col = 3: graini = .text
            .Col = 4: grafin = .text
            If (graini > 0 And grafin = 0) Or (graini = 0 And grafin > 0) Or (graini > grafin) Or (grafin < graini) Then MsgBox "Rango gramos no validos, proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
            vg_db.Execute "DELETE b_gramofamproducto WHERE gfp_subseg=" & Val(fpLongInteger1(0).Value) & " AND gfp_codreg=" & Val(fpLongInteger1(1).Value) & " AND gfp_catdie=" & FilCatDie & " AND gfp_tiprec=" & Val(codigo) & " AND gfp_fampro=" & filfampro & ""
    '        If graini > 0 And grafin > 0 Then vg_db.Execute "INSERT INTO b_gramofamproducto VALUES (" & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & filcatdie & ", " & Codigo & ", " & filfampro & ", " & graini & ", " & grafin & ")"
            If graini > 0 And grafin > 0 Then vg_db.Execute "sgpadm_iu_gramofamproducto 'A', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & FilCatDie & ", " & codigo & ", " & filfampro & ", " & graini & ", " & grafin & ""
        Next i
        
    End With
    
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Frame1.Enabled = True
Case 15
    RS.Open "SELECT DISTINCT gfp_subseg FROM  b_gramofamproducto WHERE gfp_subseg=" & Val(fpLongInteger1(0).Value) & " AND gfp_codreg=" & Val(fpLongInteger1(1).Value) & " AND gfp_catdie=" & FilCatDie & " AND gfp_fampro=" & filfampro & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    I_GramoFamProducto Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), FilCatDie, filfampro
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Or 2147217900 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
If modo = "" Then modo = "M"
If ChangeMade = True And modo = "M" Then Gl_Ac_Botones Me, 1, 0, modo: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Frame1.Enabled = False
End Sub
