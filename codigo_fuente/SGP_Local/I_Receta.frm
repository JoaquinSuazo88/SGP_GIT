VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_Receta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Recetas"
   ClientHeight    =   8310
   ClientLeft      =   3780
   ClientTop       =   2100
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8310
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7695
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   10575
      Begin VB.Frame Frame1 
         Caption         =   "Seleccción Nutrientes"
         Enabled         =   0   'False
         Height          =   4935
         Index           =   1
         Left            =   6360
         TabIndex        =   20
         Top             =   2640
         Width           =   4050
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "Código Prod."
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   2280
            TabIndex        =   23
            Top             =   4440
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   " P%| G%| CHO%"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   4440
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   3345
            Left            =   240
            MultiSelect     =   1  'Simple
            TabIndex        =   21
            Top             =   360
            Width           =   3570
         End
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "I_Receta.frx":0000
         Left            =   960
         List            =   "I_Receta.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   3645
      End
      Begin VB.Frame Frame2 
         Caption         =   "Selección Recetas"
         Height          =   4935
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   6135
         Begin VB.Frame Frame8 
            Height          =   435
            Left            =   360
            TabIndex        =   16
            Top             =   4440
            Width           =   1035
            Begin VB.TextBox TextCai2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   2
               Left            =   45
               TabIndex        =   17
               Top             =   135
               Width           =   930
            End
         End
         Begin VB.Frame Frame7 
            Height          =   435
            Left            =   1410
            TabIndex        =   14
            Top             =   4440
            Width           =   4245
            Begin VB.TextBox TextCai2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   3
               Left            =   45
               TabIndex        =   15
               Top             =   135
               Width           =   4140
            End
         End
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3975
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   5895
            _Version        =   393216
            _ExtentX        =   10398
            _ExtentY        =   7011
            _StockProps     =   64
            AllowMultiBlocks=   -1  'True
            BackColorStyle  =   1
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
            MaxCols         =   8
            MaxRows         =   20
            SpreadDesigner  =   "I_Receta.frx":0053
            ScrollBarTrack  =   3
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Metodo Preparación"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   600
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Height          =   525
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   4065
         Begin VB.OptionButton Option1 
            Caption         =   "Nombre Fantasia"
            Height          =   375
            Index           =   1
            Left            =   2130
            TabIndex        =   11
            Top             =   120
            Width           =   1845
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Nombre Receta"
            Height          =   375
            Index           =   0
            Left            =   180
            TabIndex        =   10
            Top             =   120
            Value           =   -1  'True
            Width           =   1665
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1005
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   10365
         Begin VB.OptionButton Option2 
            Caption         =   "Patrón"
            Height          =   225
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   210
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Local"
            Height          =   225
            Index           =   1
            Left            =   1290
            TabIndex        =   3
            Top             =   210
            Width           =   945
         End
         Begin VB.OptionButton Option2 
            Caption         =   "X Regimen"
            Height          =   225
            Index           =   2
            Left            =   2280
            TabIndex        =   2
            Top             =   240
            Width           =   1305
         End
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   1
            Left            =   2595
            TabIndex        =   5
            Top             =   555
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
            MinValue        =   "-2147483647"
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
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   360
            Left            =   9600
            TabIndex        =   6
            Top             =   480
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
            Index           =   1
            Left            =   3975
            TabIndex        =   7
            Top             =   555
            Width           =   5355
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   1
            Left            =   3465
            Picture         =   "I_Receta.frx":0676
            Top             =   480
            Width           =   480
         End
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   4020
            TabIndex        =   8
            Top             =   600
            Width           =   5355
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   165
         Left            =   120
         TabIndex        =   27
         Top             =   780
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Informes"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   420
         Width           =   735
      End
      Begin VB.Label lblREGSEL 
         Caption         =   "lblREGSEL"
         Height          =   270
         Left            =   4890
         TabIndex        =   24
         Top             =   390
         Width           =   2565
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
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
            Picture         =   "I_Receta.frx":0980
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   615
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      _Version        =   393216
      _ExtentX        =   873
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
      SpreadDesigner  =   "I_Receta.frx":0D1A
   End
End
Attribute VB_Name = "I_Receta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim i As Integer, iselecc As Integer, imarca As Integer
Dim iaporte As Long, est As Boolean
Dim cuenta As Long

Private Sub Combo1_Click()
If Combo1.ItemData(Combo1.ListIndex) = 2 Then
   Check1.Enabled = False
   Check1.Value = 0
   Check2(1).Enabled = True
   Check2(2).Enabled = True
   List1.Enabled = True
   Frame1(1).Enabled = True
Else
   Check2(1).Enabled = False
   Check2(1).Value = 0
   Check2(2).Enabled = False
   Check2(2).Value = 0
   List1.Enabled = False
   Frame1(1).Enabled = False
   If Combo1.ItemData(Combo1.ListIndex) = 1 Then
      Check1.Enabled = True
      Check1.Value = 0
   Else
      Check1.Enabled = False
      Check1.Value = 0
      Option2(0).Value = True
   End If
End If
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Dim codTippla As Long, nomTippla As String
On Error GoTo Man_Error
fg_centra Me
fg_carga ""
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpLongInteger1(1).ControlType = ControlTypeStatic
Toolbar3.Enabled = False
Image1(1).Enabled = False
cuenta = 0
lblREGSEL.Caption = cuenta & " recetas seleccionadas"
MsgTitulo = "Impresión de Recetas"
Combo1.ListIndex = 0
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, "Informe Recetas"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Dim RS As New ADODB.Recordset
vaSpread1.MaxRows = 0
RS.Open RutinaLectura.Regimen(2, Val(fpLongInteger1(1).Value), ""), vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
fpayuda(1).Caption = Trim(RS!reg_nombre)
RS.Close: Set RS = Nothing
End Sub

Private Sub Image1_Click(Index As Integer)
vg_left = fpayuda(1).Left + 2300
vg_nombre = "": vg_codigo = ""
B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Gen"
B_TabEst.Show 1
Me.Refresh
If vg_codigo = "" Then Exit Sub
fpLongInteger1(1).Value = Val(vg_codigo)
fpayuda(1).Caption = vg_nombre
End Sub

Private Sub Option2_Click(Index As Integer)
Select Case Index
Case 0, 1
    fpLongInteger1(1).ControlType = ControlTypeStatic
    Image1(1).Enabled = False
    fpLongInteger1(1).Value = ""
    fpayuda(1).Caption = ""
    Toolbar3.Enabled = False
    CargarDatos IIf(Index = 0, 0, -1)
Case 2
    fpLongInteger1(1).ControlType = ControlTypeNormal
    Image1(1).Enabled = True
    fpLongInteger1(1).Value = ""
    fpayuda(1).Caption = ""
    Toolbar3.Enabled = True
    vaSpread1.MaxRows = 0
End Select
End Sub

Private Sub TextCai2_Change(Index As Integer)
Dim i As Long, nom As String
Select Case Index
Case 2, 3
    vaSpread1.Visible = False
    If Trim(TextCai2(Index).text) <> "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = Index: nom = UCase(Trim(vaSpread1.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCai2(Index).text) & "*"
           vaSpread1.Col = Index
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           Else
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           End If
        Next i
        vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Trim(TextCai2(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(TextCai2(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
    If Trim(TextCai2(Index).text) = "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Next
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextCai2(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1.Visible = True
End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long, codigo As String, NomPro As String, NomFan As String, aAp As String
Select Case Button.Index
Case 1
    iselecc = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then iselecc = 1: Exit For
    Next i
    If iselecc = 0 Then MsgBox "Debe seleccionar a lo menos una receta", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Combo1.ItemData(Combo1.ListIndex) = 0 Then
       I_NombreRecetas cuenta, IIf(Option2(0).Value = True, 0, IIf(Option2(1).Value = True, -1, Val(fpLongInteger1(1).Value)))
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
       I_TarjetaRecetas cuenta, IIf(Option2(0).Value = True, 0, IIf(Option2(1).Value = True, -1, Val(fpLongInteger1(1).Value)))
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 2 Then
       iselecc = 0
       For i = 0 To List1.listcount - 1
           If List1.Selected(i) = True Then iselecc = 1: Exit For
       Next i
       If iselecc = 0 Then MsgBox "Debe Seleccionar A lo Menos Un Aporte Nutricional", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       I_AporteRecetas cuenta, IIf(Option2(0).Value = True, 0, IIf(Option2(1).Value = True, -1, Val(fpLongInteger1(1).Value)))
    End If
Case 3
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If Val(fpLongInteger1(1).Value) > 0 Then CargarDatos Val(fpLongInteger1(1).Value)
End Select
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
If est Then Exit Sub
Dim i As Long
vaSpread1.Col = 1
est = True
For i = BlockRow To BlockRow2
    vaSpread1.Row = i
    vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
Next
est = False
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
vaSpread1.Row = Row: vaSpread1.Col = Col
If Row = -1 And vaSpread1.text = 0 Then
   cuenta = 0
ElseIf Row = -1 And vaSpread1.text = 1 Then
   cuenta = vaSpread1.MaxRows
Else
   If vaSpread1.text = 1 Then cuenta = cuenta + 1 Else cuenta = cuenta - 1
End If
lblREGSEL.Caption = cuenta & " recetas seleccionadas"
End Sub

Sub CargarDatos(tiprec As Long)
'Mover recetas
vaSpread2.MaxRows = 0
vaSpread2.MaxCols = 2
RS.Open "SELECT distinct tip.tip_codigo " & _
        "FROM a_recetacatdie car inner join b_receta rec on car.car_codigo = rec.rec_catdie inner join a_recetatippla tip on tip.tip_codigo = rec.rec_tippla " & _
        "WHERE (rec_fecvig>" & Format(Date, "yyyymmdd") & " OR rec_fecvig <= 0 OR (rec_fecvig) IS NULL) " & _
        "AND   (rec.rec_catdie=" & vg_filcatdie & " OR " & vg_filcatdie & "=0) " & _
        "AND   (rec.rec_tippla=" & vg_filtippla & " OR " & vg_filtippla & "=0) and rec.rec_codigo in (select distinct red_codigo from b_recetadet where red_tiprec = " & tiprec & ") " & _
        "", vg_db, adOpenStatic
Do While Not RS.EOF
   vaSpread2.MaxRows = vaSpread2.MaxRows + 1
   vaSpread2.Row = vaSpread2.MaxRows
   vaSpread2.Col = 1
   vaSpread2.text = RS!tip_codigo
   vaSpread2.Col = 2
   vaSpread2.text = fg_BuscaenArbol(RS!tip_codigo, "a_recetatippla", "tip_codigo")
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing

vaSpread1.MaxRows = 0: imarca = 0: iselecc = 0
vaSpread1.Visible = False
RS.Open "SELECT rec.rec_codigo, rec.rec_nombre, rec.rec_nomfan, rec.rec_tiprec, car.car_codigo, tip.tip_codigo, rec.rec_basrac " & _
        "FROM a_recetacatdie car inner join b_receta rec on car.car_codigo = rec.rec_catdie inner join a_recetatippla tip on tip.tip_codigo = rec.rec_tippla " & _
        "WHERE (rec_fecvig>" & Format(Date, "yyyymmdd") & " OR rec_fecvig <= 0 OR (rec_fecvig) IS NULL) " & _
        "AND   (rec.rec_catdie=" & vg_filcatdie & " OR " & vg_filcatdie & "=0) " & _
        "AND   (rec.rec_tippla=" & vg_filtippla & " OR " & vg_filtippla & "=0) and rec.rec_codigo in (select distinct red_codigo from b_recetadet where red_tiprec = " & tiprec & ") " & _
        "ORDER BY tip.tip_codigo", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1.Visible = True: fg_descarga: MsgBox "No existe maestro recetas", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub 'Me.Hide: Unload Me
est = True
codTippla = 0
nomTippla = ""
Do While Not RS.EOF
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = 2
    vaSpread1.CellType = 5
    vaSpread1.TypeHAlign = 1
    vaSpread1.text = RS!rec_codigo
      
    vaSpread1.Col = 3
    vaSpread1.CellType = 5
    vaSpread1.TypeHAlign = 0
    vaSpread1.text = Trim(RS!rec_nombre)
      
    vaSpread1.Col = 4
    vaSpread1.CellType = 5
    vaSpread1.TypeHAlign = 1
    vaSpread1.text = Trim(RS!rec_nomfan)
    vaSpread1.Col = 5: vaSpread1.text = Trim(M_Receta.Label2(8).Caption)
    vaSpread1.Col = 6
    If Trim(M_Receta.Label2(9).Caption) = "Todos" Then
        If codTippla = RS!tip_codigo Then
            vaSpread1.text = nomTippla
        Else
           '***LOCALIZAR FILA
           vaSpread2.SetActiveCell 1, vaSpread2.SearchCol(1, 0, vaSpread1.MaxRows, Trim(Str(RS!tip_codigo)), SearchFlagsNone)
           vaSpread2.Row = vaSpread2.ActiveRow
           vaSpread1.text = vaSpread2.text
'            vaSpread1.text = fg_BuscaenArbol(RS!tip_codigo, "a_recetatippla", "tip_codigo")
        End If
        codTippla = RS!tip_codigo
        nomTippla = vaSpread1.text
    Else
        vaSpread1.text = Trim(M_Receta.Label2(9).Caption)
    End If
    vaSpread1.Col = 7: vaSpread1.text = RS!rec_basrac
    vaSpread1.Col = 8: vaSpread1.text = RS!rec_tiprec
    RS.MoveNext
Loop
est = False
RS.Close: Set RS = Nothing
vaSpread1.SortKey(1) = 3
vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
'Llenar Tabla Nutrienetes
RS.Open "SELECT * FROM a_nutriente ORDER BY nut_secnro", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No existe maestro nutrientes", vbExclamation + vbOKOnly, MsgTitulo: Me.Hide: Unload Me
List1.Clear: iaporte = 0
Do While Not RS.EOF
   List1.AddItem Trim(RS!nut_nombre)
   List1.ItemData(List1.NewIndex) = RS!nut_codigo
   If RS!nut_indpri = 1 Then List1.Selected(iaporte) = True
   iaporte = iaporte + 1
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
TextCai2(2).text = ""
TextCai2(3).text = ""
End Sub


