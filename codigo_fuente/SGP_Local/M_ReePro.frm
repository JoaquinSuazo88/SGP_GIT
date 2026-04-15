VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_ReePro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reemplazar ingredientes en recetas"
   ClientHeight    =   5655
   ClientLeft      =   810
   ClientTop       =   2505
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5655
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1035
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   0
      Width           =   8895
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   1
         Left            =   4320
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   555
         Width           =   4245
         _Version        =   196608
         _ExtentX        =   7488
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
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   0
         Left            =   4320
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   4245
         _Version        =   196608
         _ExtentX        =   7488
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   1920
         _Version        =   196608
         _ExtentX        =   3387
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
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
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   20
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Top             =   555
         Width           =   1920
         _Version        =   196608
         _ExtentX        =   3387
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
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
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   20
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingrediente Origen"
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
         Left            =   180
         TabIndex        =   8
         Top             =   345
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingrediente Destino"
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
         Left            =   180
         TabIndex        =   7
         Top             =   645
         Width           =   1680
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3840
         Picture         =   "M_ReePro.frx":0000
         Top             =   165
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3840
         Picture         =   "M_ReePro.frx":030A
         Top             =   480
         Width           =   480
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4575
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   9900
      _Version        =   393216
      _ExtentX        =   17463
      _ExtentY        =   8070
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModePermanent=   -1  'True
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
      MaxCols         =   11
      MaxRows         =   20
      ProcessTab      =   -1  'True
      RestrictRows    =   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "M_ReePro.frx":0614
      UserResize      =   2
      VisibleCols     =   5
      VisibleRows     =   20
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5655
      Left            =   9900
      TabIndex        =   3
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   9975
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_ReePro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim codproducto As String
Dim codreceta As Long, nroite As Long
Dim i As Integer, indsel As Integer
Dim canpro As Double, pctapr As Double, pctcoc As Double, pctnut As Double

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 6030
Me.Width = 10530
fg_centra Me
fg_carga (ss)
Me.HelpContextID = vg_OpcM
vaSpread1.MaxRows = 0: indsel = 0
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fg_descarga
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fptext1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    If Index = 0 Then Image1_Click 0
    If Index = 1 Then Image1_Click 1
End Select
End Sub

Private Sub fpText1_LostFocus(Index As Integer)
Select Case Index
Case 0
    If fpText1(0).text = "" Then Exit Sub
    vaSpread1.MaxRows = 0
    RS.Open "SELECT ing_nombre FROM b_ingrediente WHERE ing_codigo = '" & fpText1(0).text & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).text = "": fpText1(0).text = "": vg_codigo = "": MsgBox "Información no existe", vbExclamation + vbOKOnly, "Reemplazar ingrediente en receta": Exit Sub
    fpayuda(0).text = Trim(RS!ing_nombre)
    RS.Close: Set RS = Nothing
    MoverDatosGrilla
Case 1
    If fpText1(1).text = "" Then Exit Sub
    RS.Open "SELECT ing_nombre FROM  b_ingrediente WHERE ing_codigo = '" & fpText1(1).text & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).text = "": fpText1(1).text = "": vg_codigo = "": MsgBox "Información no existe", vbExclamation + vbOKOnly, "Reemplazar ingrediente en receta": Exit Sub
    fpayuda(1).text = Trim(RS!ing_nombre)
    RS.Close: Set RS = Nothing
    If vaSpread1.MaxRows < 1 Then MoverDatosGrilla
End Select
End Sub

Sub MoverDatosGrilla()
fg_carga ""
With vaSpread1
    .MaxRows = 0
    RS.Open "SELECT DISTINCT a.rec_nombre, a.rec_codigo, " & _
            "b.red_codpro, b.red_nroite, b.red_canpro, " & _
            "b.red_pctapr, b.red_pctcoc, b.red_pctnut " & _
            "FROM  b_receta a, b_recetadet b " & _
            "WHERE b.red_codigo = a.rec_codigo " & _
            "AND   b.red_codpro = '" & fpText1(0).text & "' " & _
            "AND  (a.rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
            "AND  (a.rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) " & _
            "AND   a.rec_tiprec = 0 AND b.red_cencos = '0' ORDER BY a.rec_nombre", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "Ingrediente no existe en recetarios", vbExclamation + vbOKOnly, "reemplazar ingrediente en receta": Exit Sub
    Do While Not RS.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1: .CellType = 10:  .TypeCheckText = " ": .TypeCheckCenter = True: .text = "" ' checked
       .Col = 2: .text = "(" & RS!rec_codigo & ") " & Trim(RS!rec_nombre)
       .Col = 3: .text = RS!red_canpro: .ForeColor = &HFF0000
       .Col = 4: .text = RS!rec_codigo
       .Col = 5: .text = RS!red_codpro
       .Col = 6: .text = RS!red_pctapr: .ForeColor = &HFF0000
       .Col = 7: .text = RS!red_pctcoc: .ForeColor = &HFF0000
       .Col = 8: .TypeHAlign = 1: .text = Format(((((RS!red_canpro * RS!red_pctapr) / 100) * RS!red_pctcoc) / 100), fg_Pict(6, 2))
       .Col = 9: .text = RS!red_pctnut: .ForeColor = &HFF0000
       .Col = 10: .TypeHAlign = 1: .text = Format(((RS!red_pctnut / 100) * RS!red_canpro), fg_Pict(6, 2))
       .Col = 11: .text = RS!red_nroite
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing: fg_descarga
End With
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(0).Left + 1770
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingrediente", "Gen"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    vaSpread1.MaxRows = 0
    fpText1(0).text = vg_codigo
    fpayuda(0).text = vg_nombre
    If fpText1(0).text <> "" Then MoverDatosGrilla
Case 1
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(1).Left + 1770
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "Gen"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    fpText1(1).text = vg_codigo
    fpayuda(1).text = vg_nombre
    If fpText1(0).text <> "" And vaSpread1.MaxRows < 1 Then MoverDatosGrilla
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 1, 3
    If vaSpread1.MaxRows < 1 Then Exit Sub
    RS.Open "SELECT ing_nombre FROM b_ingrediente WHERE ing_codigo = '" & fpText1(0).text & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1.MaxRows = 0: MsgBox "No Existe Ingredientes", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente": Exit Sub
    RS.Close: Set RS = Nothing
    If fpText1(1).text <> "" Then
       RS.Open "SELECT ing_nombre FROM b_ingrediente WHERE ing_codigo = '" & fpText1(1).text & "'", vg_db, adOpenStatic
       If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1.MaxRows = 0: MsgBox "No Existe Ingredientes a Reemplazar", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente": Exit Sub
       RS.Close: Set RS = Nothing
    End If
    RS.Open "SELECT DISTINCT a.rec_nombre, a.rec_codigo, " & _
            "b.red_codpro, b.red_nroite, b.red_canpro, " & _
            "b.red_pctapr, b.red_pctcoc, b.red_pctnut " & _
            "FROM  b_receta a, b_recetadet b " & _
            "WHERE b.red_codigo = a.rec_codigo " & _
            "AND   b.red_codpro = '" & fpText1(0).text & "' " & _
            "AND  (a.rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
            "AND  (a.rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) " & _
            "AND   a.rec_tiprec = 0 AND b.red_cencos = '0'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1.MaxRows = 0: MsgBox "No Existe Ingredientes Origen en Recetario", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente": Exit Sub
    RS.Close: Set RS = Nothing
    indsel = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        If vaSpread1.text = "1" Then indsel = 1: Exit For
    Next i
    If Button.Index = 1 Then
       If indsel = 0 Then MsgBox "Seleccione Uno o Más Recetas a Reemplazar", vbCritical + vbOKOnly, "Cambio Ingrediente": Exit Sub
       If fpText1(1).text <> "" Then
          msg = " Esta Seguro Reemplazar " & "(" & Trim(fpayuda(0).text) & ")" & " Por " & "(" & Trim(fpayuda(1).text) & ")" & " En Las Recetas Seleccionadas ?"
       Else
          msg = " Esta Seguro Remplazar Datos en " & "(" & Trim(fpayuda(0).text) & ")" & " "
       End If
    ElseIf Button.Index = 3 Then
       If indsel = 0 Then MsgBox "Seleccione Uno o Más Recetas a Eliminar", vbCritical + vbOKOnly, "Eliminar Ingrediente": Exit Sub
       msg = " Esta Seguro Eliminar " & "(" & Trim(fpayuda(0).text) & ")" & " En Las Recetas Seleccionadas ?"
    End If
    If MsgBox("Esta Seguro ?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    fg_carga ""
    vg_db.BeginTrans
    With vaSpread1
      For i = 1 To .MaxRows
          .Row = i
          .Col = 1
          If .text = "1" Then
             .Col = 3: canpro = 0: canpro = .text
             .Col = 4: codreceta = 0: codreceta = .text
             .Col = 6: pctapr = 0: pctapr = .text
             .Col = 7: pctcoc = 0: pctcoc = .text
             .Col = 9: pctnut = 0: pctnut = .text
             .Col = 11: nroite = 0: nroite = .text
             If Button.Index = 1 Then
                If fpText1(1).text <> "" Then
                   codproducto = fpText1(1).text
                Else
                   codproducto = fpText1(0).text
                End If
                vg_db.Execute "UPDATE b_recetadet " & _
                              "SET red_codpro = " & codproducto & ", " & _
                              "red_canpro = " & canpro & ", " & _
                              "red_pctapr = " & pctapr & ", " & _
                              "red_pctcoc = " & pctcoc & ", " & _
                              "red_pctnut = " & pctnut & " " & _
                              "WHERE red_codigo = " & codreceta & " " & _
                              "AND   red_codpro = '" & fpText1(0).text & "' " & _
                              "AND   red_nroite = " & nroite & " AND red_cencos='0'"
             ElseIf Button.Index = 3 Then
                vg_db.Execute "DELETE b_recetadet " & _
                              "WHERE red_codigo = " & codreceta & " " & _
                              "AND   red_codpro = '" & fpText1(0).text & "' " & _
                              "AND   red_nroite = " & nroite & " AND red_cencos = '0'"
             End If
          End If
      Next i
      fg_descarga
      If Button.Index = 1 Then
         MsgBox "Cambiar ingrediente finalizo sin problema", vbInformation + vbOKOnly, "Reemplazar ingrediente en receta"
      ElseIf Button.Index = 3 Then
         MsgBox "Eliminación de ingrediente finalizo sin problema", vbInformation + vbOKOnly, "Eliminar ingrediente en receta"
      End If
      indsel = 0
      .MaxRows = 0
    End With
    vg_db.CommitTrans
Case 5
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
With vaSpread1
    If .MaxRows < 1 Then Exit Sub
    If Col = 1 And Row = 0 Then
       If indsel = 0 Then
          For i = 1 To .MaxRows
              .Row = i
              .Col = 1
              .CellType = 10
              .TypeCheckText = ""
              .TypeCheckCenter = True
              .Value = "1" ' checked
          Next i
          indsel = 1
       Else
          For i = 1 To .MaxRows
              .Row = i
              .Col = 1
              .CellType = 10
              .TypeCheckText = " "
              .TypeCheckCenter = True
              .Value = "" ' checked
          Next i
          indsel = 0
       End If
    End If
End With
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
With vaSpread1
    If .MaxRows < 1 Then Exit Sub
    '------- Calcular Gramaje Neto
    pctnut = 0: canpro = 0: pctapr = 0: pctcoc = 0
    .Row = .ActiveRow
    .Col = 3: canpro = .text
    .Col = 9: pctnut = .text
    .Col = 10: .CellType = 5: .TypeHAlign = 1: .text = Format(CCur((pctnut / 100) * canpro), fg_Pict(6, 2))
    '------- Calcular % Limpieza & Cocción
    .Col = 6: pctapr = .text
    'cantservida = CCur((paporv / 100) * canpro)
    .Col = 7: pctcoc = .text
    'cantservida = CCur((pcoccion / 100) * cantservida)
    .Col = 8: .CellType = 5: .TypeHAlign = 1: .text = Format(CCur(((pctapr / 100) * canpro) * (pctcoc / 100)), fg_Pict(6, 2))
End With
End Sub
