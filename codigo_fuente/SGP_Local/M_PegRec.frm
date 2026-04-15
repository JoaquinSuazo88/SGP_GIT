VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form M_PegRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pegar Receta"
   ClientHeight    =   4350
   ClientLeft      =   1470
   ClientTop       =   3870
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4335
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8535
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Top             =   360
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
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
         MinValue        =   "-2147483647"
         NegFormat       =   1
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
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Left            =   3600
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   4005
         _Version        =   196608
         _ExtentX        =   7064
         _ExtentY        =   556
         Enabled         =   0   'False
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
         BackColor       =   -2147483638
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
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
         ControlType     =   3
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
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   8295
         _Version        =   393216
         _ExtentX        =   14631
         _ExtentY        =   5530
         _StockProps     =   64
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
         MaxCols         =   10
         MaxRows         =   20
         ProcessTab      =   -1  'True
         RestrictRows    =   -1  'True
         SpreadDesigner  =   "M_PegRec.frx":0000
         UserResize      =   2
         VisibleCols     =   5
         VisibleRows     =   20
         ScrollBarTrack  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   3600
         TabIndex        =   6
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Receta Pegar"
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
         Left            =   945
         TabIndex        =   4
         Top             =   405
         Width           =   1185
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3150
         Picture         =   "M_PegRec.frx":071B
         Top             =   285
         Width           =   480
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   4350
      Left            =   8550
      TabIndex        =   0
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   7673
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_PegRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim nroite As Long, CodPro As Long
Dim i As Integer, indsel As Integer
Dim canpro As Double, cospro As Double, pctapr As Double, pctcoc As Double, pctnut As Double

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_centra Me
Me.HelpContextID = vg_OpcM
Msgtitulo = "Pegar Recetas"
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar "
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"

RS.Open "SELECT rec_codigo, rec_nombre FROM b_receta WHERE rec_codigo = " & vg_codreceta & " AND rec_tiprec = 0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Receta Fue Elimanda", vbInformation + vbOKOnly, Msgtitulo: Me.Hide: Unload Me
Label1(0).Caption = "(" & RS!rec_codigo & ") " & Trim(RS!rec_nombre)
RS.Close: Set RS = Nothing
vaSpread1.MaxRows = 0
End Sub

Private Sub fpLongInteger1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    image1_Click
End Select
End Sub

Private Sub fpLongInteger1_LostFocus()
If Val(fpLongInteger1.Value) = 0 Then Exit Sub
vaSpread1.MaxRows = 0
RS.Open "SELECT rec_nombre FROM b_receta WHERE rec_codigo = " & Val(fpLongInteger1.Value) & " AND rec_tiprec = 0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda.text = "": fpLongInteger1.text = "": vg_codigo = "": MsgBox "Información no existe", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
fpayuda.text = Trim(RS!rec_nombre)
RS.Close: Set RS = Nothing
MoverDatosGrilla
End Sub

Private Sub image1_Click()
vg_codigo = "": vg_nombre = ""
vg_left = fpayuda.Left + 1700
B_TabEst.LlenaDatos "b_receta", "rec_", "Recetas", "0"
B_TabEst.Show 1
If vg_codigo = "" Then Exit Sub
fpLongInteger1.Value = Val(vg_codigo)
fpayuda.text = vg_nombre
MoverDatosGrilla
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    If vaSpread1.MaxRows < 1 Then Exit Sub
    RS.Open "SELECT rec_nombre FROM b_receta WHERE rec_codigo = " & Val(fpLongInteger1.Value) & " AND rec_tiprec = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda.text = "": fpLongInteger1.text = "": vg_codigo = "": MsgBox "Información no existe", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    fpayuda.text = Trim(RS!rec_nombre)
    RS.Close: Set RS = Nothing
    indsel = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        If vaSpread1.text = "1" Then indsel = 1: Exit For
    Next i
    If indsel = 0 Then MsgBox "Seleccione Uno o Más Items", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Pegar registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    fg_carga ""
    vg_db.BeginTrans
    RS.Open "SELECT MAX(red_nroite) AS maxnroite, red_codigo FROM b_recetadet WHERE red_codigo = " & vg_codreceta & " AND red_cencos = '0' GROUP BY red_codigo", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: vg_db.RollbackTrans: fg_descarga: MsgBox "Receta Fue Elimanda", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
    nroite = RS!maxnroite + 1
    RS.Close: Set RS = Nothing
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then
           vaSpread1.Col = 2: CodPro = 0: CodPro = vaSpread1.text
           vaSpread1.Col = 4: canpro = 0: canpro = vaSpread1.text
           vaSpread1.Col = 5: pctapr = 0: pctapr = vaSpread1.text
           vaSpread1.Col = 6: pctcoc = 0: pctcoc = vaSpread1.text
           vaSpread1.Col = 8: pctnut = 0: pctnut = vaSpread1.text
           vaSpread1.Col = 10: cospro = 0: cospro = vaSpread1.text
           vg_db.Execute "INSERT INTO b_recetadet VALUES (" & vg_codreceta & ", " & nroite & ", " & CodPro & ", " & _
                          "" & canpro & ", " & cospro & ", " & pctapr & ", " & pctcoc & ", " & pctnut & ", 0, '0')"
           nroite = nroite + 1
        End If
    Next i
    vg_swpegreceta = 1
    fg_descarga
    MsgBox "Pegado Finalizado Sin Problema", vbInformation + vbOKOnly, Msgtitulo
    vg_db.CommitTrans
Case 3
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
vg_swpegreceta = 0
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub MoverDatosGrilla()
fg_carga ""
vaSpread1.MaxRows = 0
RS.Open "SELECT DISTINCT b.red_codpro, b.red_nroite, b.red_canpro, b.red_pctapr, " & _
        "b.red_pctcoc, b.red_pctnut, a.ing_nombre, (SELECT DISTINCT cpi_precos FROM b_contlistpreing WHERE cpi_coding = a.ing_codigo AND cpi_cencos = '" & MuestraCasino(1) & "')" & _
        "FROM  b_ingrediente a, b_recetadet b " & _
        "WHERE b.red_codpro = a.ing_codigo " & _
        "AND   b.red_codigo = " & Val(fpLongInteger1.Value) & " AND b.red_cencos = '0' " & _
        "ORDER BY b.red_nroite", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "Receta no existe ", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1: vaSpread1.CellType = 10:  vaSpread1.TypeCheckText = " ": vaSpread1.TypeCheckCenter = True: vaSpread1.text = "" ' checked
   vaSpread1.Col = 2: vaSpread1.text = RS!red_codpro
   vaSpread1.Col = 3: vaSpread1.text = Trim(RS!ing_nombre)
   vaSpread1.Col = 4: vaSpread1.text = RS!red_canpro: vaSpread1.ForeColor = &HFF0000
   vaSpread1.Col = 5: vaSpread1.text = RS!red_pctapr: vaSpread1.ForeColor = &HFF0000
   vaSpread1.Col = 6: vaSpread1.text = RS!red_pctcoc: vaSpread1.ForeColor = &HFF0000
   vaSpread1.Col = 7: vaSpread1.TypeHAlign = 1: vaSpread1.text = Format(((((RS!red_canpro * RS!red_pctapr) / 100) * RS!red_pctcoc) / 100), fg_Pict(6, 2))
   vaSpread1.Col = 8: vaSpread1.text = RS!red_pctnut: vaSpread1.ForeColor = &HFF0000
   vaSpread1.Col = 9: vaSpread1.TypeHAlign = 1: vaSpread1.text = Format(((RS!red_pctnut / 100) * RS!red_canpro), fg_Pict(6, 2))
   vaSpread1.Col = 10: vaSpread1.text = RS!cpi_precos
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing: fg_descarga
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
If Col = 1 And Row = 0 Then
   If indsel = 0 Then
      For i = 1 To vaSpread1.MaxRows
          vaSpread1.Row = i
          vaSpread1.Col = 1
          vaSpread1.CellType = 10
          vaSpread1.TypeCheckText = ""
          vaSpread1.TypeCheckCenter = True
          vaSpread1.Value = "1" ' checked
      Next i
      indsel = 1
   Else
      For i = 1 To vaSpread1.MaxRows
          vaSpread1.Row = i
          vaSpread1.Col = 1
          vaSpread1.CellType = 10
          vaSpread1.TypeCheckText = " "
          vaSpread1.TypeCheckCenter = True
          vaSpread1.Value = "" ' checked
      Next i
      indsel = 0
   End If
End If
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
'------- Calcular Gramaje neto
pctnut = 0: canpro = 0: pctapr = 0: pctcoc = 0
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 4: canpro = vaSpread1.text
vaSpread1.Col = 8: pctnut = vaSpread1.text
vaSpread1.Col = 9: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 1: vaSpread1.text = Format(CCur((pctnut / 100) * canpro), fg_Pict(6, 2))
'------- Calcular % limpieza & cocción
vaSpread1.Col = 5: pctapr = vaSpread1.text
'cantservida = CCur((paporv / 100) * canpro)
vaSpread1.Col = 6: pctcoc = vaSpread1.text
'cantservida = CCur((pcoccion / 100) * cantservida)
vaSpread1.Col = 7: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 1: vaSpread1.text = Format(CCur(((pctapr / 100) * canpro) * (pctcoc / 100)), fg_Pict(6, 2))
End Sub
