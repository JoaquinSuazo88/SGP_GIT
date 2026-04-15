VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_GeAdAn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Adicionales & Anulaciones"
   ClientHeight    =   6240
   ClientLeft      =   795
   ClientTop       =   2220
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   30
      TabIndex        =   13
      Top             =   5580
      Width           =   10200
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   150
         TabIndex        =   14
         Top             =   180
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   688
         ButtonWidth     =   2963
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         _Version        =   393216
      End
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   1455
      Left            =   3090
      TabIndex        =   10
      Top             =   2370
      Visible         =   0   'False
      Width           =   4335
      _Version        =   393216
      _ExtentX        =   7646
      _ExtentY        =   2566
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
      MaxCols         =   0
      MaxRows         =   0
      SpreadDesigner  =   "M_GeAdAn.frx":0000
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   7215
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   810
         Width           =   1275
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   210
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1455
         TabIndex        =   1
         Top             =   525
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
         Text            =   "10/2004"
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
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   1515
         TabIndex        =   12
         Top             =   885
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido"
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
         Left            =   105
         TabIndex        =   9
         Top             =   915
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha(mm/aa)"
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
         Left            =   135
         TabIndex        =   7
         Top             =   585
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   285
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2685
         Picture         =   "M_GeAdAn.frx":01E2
         Top             =   90
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3120
         TabIndex        =   5
         Top             =   210
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3165
         TabIndex        =   11
         Top             =   255
         Width           =   3855
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4155
      Left            =   0
      TabIndex        =   3
      Top             =   1410
      Width           =   10215
      _Version        =   393216
      _ExtentX        =   18018
      _ExtentY        =   7329
      _StockProps     =   64
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   1
      SpreadDesigner  =   "M_GeAdAn.frx":04EC
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6240
      Left            =   10245
      TabIndex        =   8
      Top             =   0
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   11007
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_GeAdAn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset, RS2 As New ADODB.Recordset
Dim fecha As Long
Dim accion As Boolean
Dim MsgTitulo As String

Private Sub Combo1_Click(Index As Integer)
If accion = False Then Exit Sub
Dim codtip As Integer, auxcodtip As Integer
Dim nomtip As String
Dim canped As Double
If Combo1(0).ListIndex = -1 Or Combo1(0).Text = "" Then Exit Sub
codtip = Val(fg_codigocbo(Combo1, 0, 1, "")): auxcodtip = 0
vaSpread1.MaxRows = 0: canped = 0
vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = &HC0FFFF
fecha = 0: fecha = Mid(fpDateTime1.Text, 4, 4) & Mid(fpDateTime1.Text, 1, 2)

'------- Validar si existe pedidos adicionales & anulaciones
RS.Open "select distinct ped_fecenv from b_minutapedido " & _
        "where ped_codcas='" & LimpiaDato(Trim(fpText.Text)) & "' " & _
        "and ped_anomes=" & fecha & " " & _
        "and ped_tipped=" & codtip & " " & _
        "and ped_fecenv=0", vg_db, adOpenStatic
If Not RS.EOF Then
   fecenv = RS!ped_fecenv
   RS.Close: Set RS = Nothing
   If fecenv > 0 Then Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False) Else Toolbar1.Buttons(1).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Frame2.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(3).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(5).Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
   RS.Open "select b_ingrediente.ing_codigo, b_ingrediente.ing_nombre, b_ingrediente.ing_precos, a_unidadmed.unm_nomcor, b_productos.pro_codigo, b_productos.pro_nombre, b_productos.pro_codtip, b_productos.pro_facsto, a_tipopro.tip_nombre, " & _
           "a_unidad.uni_nomcor, b_minutapedido.ped_canped as cantidad, b_minutapedido.ped_canmin, b_productos.pro_propon " & _
           "from b_minutapedido, b_ingrediente, b_productos, a_tipopro, a_unidad, a_unidadmed where b_minutapedido.ped_coding=b_ingrediente.ing_codigo and b_minutapedido.ped_codpro=b_productos.pro_codigo " & _
           "and b_ingrediente.ing_unimed=a_unidadmed.unm_codigo and b_productos.pro_codtip=a_tipopro.tip_codigo and b_productos.pro_coduni=a_unidad.uni_codigo and b_productos.pro_facsto>0 " & _
           "and b_minutapedido.ped_codcas='" & LimpiaDato(Trim(fpText.Text)) & "' and b_minutapedido.ped_anomes=" & fecha & " " & _
           "and b_minutapedido.ped_tipped=" & codtip & " and b_minutapedido.ped_fecenv=0 order by a_tipopro.tip_nombre, b_ingrediente.ing_nombre", vg_db, adOpenStatic
   codtip = 0
   If Not RS.EOF Then
      Do While Not RS.EOF
'         If RS!pro_codtip <> codtip Then
'            vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
'            vaSpread1.Col = -1: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.BackColor = &HC0FFC0
'            vaSpread1.Col = 2: vaSpread1.CellType = 5: vaSpread1.Text = RS!tip_nombre
'            vaSpread1.Col = 1: vaSpread1.CellType = 5: vaSpread1.Text = ""
'            vaSpread1.Col = 3: vaSpread1.CellType = 5: vaSpread1.Text = ""
'            vaSpread1.Col = 4: vaSpread1.CellType = 5: vaSpread1.Text = ""
'            vaSpread1.Col = 5: vaSpread1.CellType = 5: vaSpread1.Text = ""
'            vaSpread1.Col = 6: vaSpread1.CellType = 5: vaSpread1.Text = ""
'            codtip = RS!pro_codtip: nomtip = RS!tip_nombre
'         End If
         If RS!ing_codigo <> codtip Then
            vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
            vaSpread1.Col = -1: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.BackColor = &HC0FFC0
            vaSpread1.Col = 2: vaSpread1.CellType = 5: vaSpread1.Text = RS!ing_nombre
            vaSpread1.Col = 1: vaSpread1.CellType = 5: vaSpread1.Text = RS!ing_codigo
            vaSpread1.Col = 3: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 1: vaSpread1.Text = Format(RS!ped_canmin, fg_Pict(6, 2))
            vaSpread1.Col = 4: vaSpread1.CellType = 5: vaSpread1.Text = Trim(RS!unm_nomcor)
            vaSpread1.Col = 5: vaSpread1.CellType = 5: vaSpread1.Text = ""
            vaSpread1.Col = 6: vaSpread1.CellType = 5: vaSpread1.Text = RS!ing_precos
            codtip = RS!ing_codigo
         End If
      
         vaSpread1.MaxRows = vaSpread1.MaxRows + 1
         vaSpread1.Row = vaSpread1.MaxRows
         vaSpread1.Col = 1: vaSpread1.CellType = 5: vaSpread1.Text = RS!pro_codigo
         vaSpread1.Col = 2: vaSpread1.CellType = 5: vaSpread1.Text = RS!pro_nombre
         vaSpread1.Col = 3: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 1: vaSpread1.Text = "" 'Format(RS!ped_canmin, fg_Pict(6, 2))
         vaSpread1.Col = 4
         vaSpread1.CellType = 3
         vaSpread1.TypeIntegerMin = 1
         vaSpread1.TypeIntegerMax = 9999999
         vaSpread1.TypeHAlign = 1
         vaSpread1.TypeSpin = False
         vaSpread1.TypeIntegerSpinInc = 1
         vaSpread1.TypeIntegerSpinWrap = False
         vaSpread1.Text = Format(RS!cantidad, fg_Pict(6, 0))
         vaSpread1.ForeColor = &HFF0000
      
         vaSpread1.Col = 5: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 0: vaSpread1.Text = RS!uni_nomcor '& " x " & RS!pro_facsto
         vaSpread1.Col = 6: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 0: vaSpread1.Text = "" 'RS!pro_propon
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing: Exit Sub

Else
   RS.Close: Set RS = Nothing
   If codtip = 2 Then
'      RS.Open "select b_productos.pro_codigo, b_productos.pro_nombre, b_productos.pro_codtip, b_productos.pro_uniemb, b_productos.pro_propon, a_tipopro.tip_nombre, a_embalaje.emb_nomcor, sum(b_minutacambios.cam_canpro) as cantidad " & _
'              "from b_minuta, b_minutacambios, b_productos, a_tipopro, a_embalaje Where b_minutacambios.cam_codmin=b_minuta.min_codigo and b_minutacambios.cam_fecmin=b_minuta.min_fecmin and b_minutacambios.cam_codpro=b_productos.pro_codigo " & _
'              "and b_productos.pro_codtip=a_tipopro.tip_codigo and b_productos.pro_codemb=a_embalaje.emb_codigo and b_minuta.min_cencos='" & LimpiaDato(Trim(fpText.Text)) & "' and val(mid(b_minuta.min_fecmin,1,6))=" & fecha & " " & _
'              "and b_minutacambios.cam_fecped=0 group by b_productos.pro_codigo, b_productos.pro_nombre, b_productos.pro_codtip, b_productos.pro_uniemb, b_productos.pro_propon, a_tipopro.tip_nombre, a_embalaje.emb_nomcor order by a_tipopro.tip_nombre,  b_productos.pro_nombre", vg_db, adOpenStatic
'      If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe adicionales", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
'      RS1.Open "select b_productos.pro_codigo, sum(b_minutadet.mid_numrac*(b_recetadet.red_canpro/b_receta.rec_basrac)) as cantidad from b_receta, b_recetadet, b_minuta, b_minutadet, b_productos where b_minutadet.mid_codigo in (select cam_codmin from b_minutacambios where cam_fecped=0) " & _
'               "and b_minutadet.mid_codigo=b_minuta.min_codigo and b_minutadet.mid_codrec=b_recetadet.red_codigo and b_recetadet.red_codigo=b_receta.rec_codigo and b_recetadet.red_codpro=b_productos.pro_codigo and b_minuta.min_cencos='" & LimpiaDato(Trim(fpText.Text)) & "' " & _
'               "and val(mid(b_minuta.min_fecmin,1,6))=" & fecha & " and b_minutadet.mid_tipmin='1' group by b_productos.pro_codigo", vg_db, adOpenStatic
      RS.Open "select b_productos.pro_codigo, b_productos.pro_nombre, b_productos.pro_codtip, b_productos.pro_facsto, b_productos.pro_propon, a_tipopro.tip_nombre, a_unidad.uni_nomcor, sum(b_minutacambios.cam_canpro) as cantidad " & _
              "from b_minuta, b_minutacambios, b_productos, b_ingrediente, a_tipopro, a_unidad Where b_minutacambios.cam_codmin=b_minuta.min_codigo and b_minutacambios.cam_fecmin=b_minuta.min_fecmin and b_minutacambios.cam_codpro=b_ingrediente.ing_codigo " & _
              "and b_ingrediente.ing_codped=b_productos.pro_codigo and b_productos.pro_codtip=a_tipopro.tip_codigo and b_productos.pro_coduni=a_unidad.uni_codigo and b_productos.pro_facsto>0 and b_minuta.min_cencos='" & LimpiaDato(Trim(fpText.Text)) & "' and val(mid(b_minuta.min_fecmin,1,6))=" & fecha & " " & _
              "and b_minutacambios.cam_fecped=0 group by b_productos.pro_codigo, b_productos.pro_nombre, b_productos.pro_codtip, b_productos.pro_facsto, b_productos.pro_propon, a_tipopro.tip_nombre, a_unidad.uni_nomcor order by a_tipopro.tip_nombre,  b_productos.pro_nombre", vg_db, adOpenStatic
      If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe adicionales", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
      RS1.Open "select b_productos.pro_codigo, sum(b_minutadet.mid_numrac*(b_recetadet.red_canpro/b_receta.rec_basrac)) as cantidad from b_receta, b_recetadet, b_minuta, b_minutadet, b_productos, b_ingrediente where b_minutadet.mid_codigo in (select cam_codmin from b_minutacambios where cam_fecped=0) " & _
               "and b_minutadet.mid_codigo=b_minuta.min_codigo and b_minutadet.mid_codrec=b_recetadet.red_codigo and b_recetadet.red_codigo=b_receta.rec_codigo and b_recetadet.red_codpro=b_ingrediente.ing_codigo and b_ingrediente.ing_codped=b_productos.pro_codigo and b_minuta.min_cencos='" & LimpiaDato(Trim(fpText.Text)) & "' " & _
               "and val(mid(b_minuta.min_fecmin,1,6))=" & fecha & " and b_minutadet.mid_tipmin='1' group by b_productos.pro_codigo", vg_db, adOpenStatic
      If RS1.EOF Then RS1.Close: Set RS1 = Nothing: RS.Close: Set RS = Nothing: Exit Sub
      vaSpread2.MaxCols = 7: vaSpread2.MaxRows = 0
      Do While Not RS.EOF
         Do While Not RS1.EOF
            If RS!pro_codigo = RS1!pro_codigo And Val(RS!cantidad) > Val(RS1!cantidad) Then
               canped = CCur(RS!cantidad - RS1!cantidad)
               vaSpread2.MaxRows = vaSpread2.MaxRows + 1
               vaSpread2.Row = vaSpread2.MaxRows
               vaSpread2.Col = 1: vaSpread2.Text = RS!pro_codigo
               vaSpread2.Col = 2: vaSpread2.Text = RS!pro_nombre
               vaSpread2.Col = 3: vaSpread2.Text = canped
               vaSpread2.Col = 4: vaSpread2.Text = RS!pro_codtip
               vaSpread2.Col = 5: vaSpread2.Text = RS!tip_nombre
               vaSpread2.Col = 6: vaSpread2.Text = RS!uni_nomcor '& " x " & RS!pro_facsto
               vaSpread2.Col = 7: vaSpread2.Text = RS!pro_facsto
               Exit Do
            ElseIf RS!pro_codigo = RS1!pro_codigo And (Val(RS!cantidad) = Val(RS1!cantidad) Or Val(RS!cantidad) < Val(RS1!cantidad)) Then
               vaSpread2.MaxRows = vaSpread2.MaxRows + 1
               vaSpread2.Row = vaSpread2.MaxRows
               vaSpread2.Col = 1: vaSpread2.Text = RS!pro_codigo
               Exit Do
            End If
            RS1.MoveNext
         Loop
         '------- revisar vector si existe codigo producto
         For i = 1 To vaSpread2.MaxRows
             vaSpread2.Row = i
             vaSpread2.Col = 1
             If RS!pro_codigo = vaSpread2.Text Then
                Exit For
             ElseIf i = vaSpread2.MaxRows And RS!cantidad > 0 Then
               canped = CCur(RS!cantidad)
               vaSpread2.MaxRows = vaSpread2.MaxRows + 1
               vaSpread2.Row = vaSpread2.MaxRows
               vaSpread2.Col = 1: vaSpread2.Text = RS!pro_codigo
               vaSpread2.Col = 2: vaSpread2.Text = RS!pro_nombre
               vaSpread2.Col = 3: vaSpread2.Text = canped
               vaSpread2.Col = 4: vaSpread2.Text = RS!pro_codtip
               vaSpread2.Col = 5: vaSpread2.Text = RS!tip_nombre
               vaSpread2.Col = 6: vaSpread2.Text = RS!uni_nomcor '& " x " & RS!pro_facsto
               vaSpread2.Col = 7: vaSpread2.Text = RS!pro_facsto
               Exit For
             End If
         Next i
         RS1.MoveFirst
         RS.MoveNext
      Loop
      RS.Close: Set RS = Nothing: RS1.Close: Set RS1 = Nothing
   ElseIf codtip = 3 Then
      RS.Open "select b_productos.pro_codigo, b_productos.pro_nombre, b_productos.pro_facing, b_productos.pro_codtip, b_productos.pro_facsto, b_productos.pro_propon, a_tipopro.tip_nombre, " & _
              "a_unidad.uni_nomcor, sum(b_minutadet.mid_numrac*(b_recetadet.red_canpro/b_receta.rec_basrac)) as cantidad " & _
              "from b_receta, b_recetadet, b_minuta, b_minutadet, b_ingrediente, b_productos, a_tipopro, a_unidad " & _
              "where b_minutadet.mid_codigo in (select cam_codmin from b_minutacambios where cam_fecped=0) " & _
              "and b_minutadet.mid_codigo=b_minuta.min_codigo and b_minutadet.mid_codrec=b_recetadet.red_codigo " & _
              "and b_recetadet.red_codigo=b_receta.rec_codigo and b_recetadet.red_codpro=b_ingrediente.ing_codigo and b_ingrediente.ing_codped=b_productos.pro_codigo " & _
              "and b_productos.pro_codtip=a_tipopro.tip_codigo and b_productos.pro_coduni=a_unidad.uni_codigo and b_productos.pro_facsto>0 " & _
              "and b_minuta.min_cencos='" & LimpiaDato(Trim(fpText.Text)) & "' and val(mid(b_minuta.min_fecmin,1,6))=" & fecha & " " & _
              "and b_minutadet.mid_tipmin='1' group by b_productos.pro_codigo, b_productos.pro_nombre, b_productos.pro_facing, b_productos.pro_codtip, b_productos.pro_facsto, b_productos.pro_propon, a_tipopro.tip_nombre, a_unidad.uni_nomcor " & _
              "order by a_tipopro.tip_nombre, b_productos.pro_nombre", vg_db, adOpenStatic
      If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe anulaciones", vbCritical + vbOKOnly, MsgTitulo: Exit Sub

      RS1.Open "select b_productos.pro_codigo, sum(b_minutacambios.cam_canpro) as cantidad from b_minuta, b_minutacambios, b_ingrediente, b_productos " & _
               "where b_minutacambios.cam_codmin=b_minuta.min_codigo and b_minutacambios.cam_fecmin=b_minuta.min_fecmin " & _
               "and b_minutacambios.cam_codpro=b_ingrediente.ing_codigo and b_ingrediente.ing_codped=b_productos.pro_codigo " & _
               "and b_minuta.min_cencos='" & LimpiaDato(Trim(fpText.Text)) & "' and val(mid(b_minuta.min_fecmin,1,6))=" & fecha & " " & _
               "and b_minutacambios.cam_fecped=0 group by b_productos.pro_codigo", vg_db, adOpenStatic
      If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
      vaSpread2.MaxCols = 7: vaSpread2.MaxRows = 0
      fecenv = 1: Toolbar1.Buttons(1).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = False
      Do While Not RS.EOF
         Do While Not RS1.EOF
            If RS!pro_codigo = RS1!pro_codigo And Val(RS!cantidad) > Val(RS1!cantidad) Then
               '------- Validar si existe productos a rebajar en mensual del mes
               canped = CCur(RS!cantidad - RS1!cantidad)
               RS2.Open "select * from b_minutapedido where ped_codcas='" & LimpiaDato(Trim(fpText.Text)) & "' and ped_anomes=" & fecha & " " & _
                        "and ped_tipped=1 and ped_codpro='" & RS!pro_codigo & "' and ped_fecenv>0", vg_db, adOpenStatic
               If RS2.EOF Then RS2.Close: Set RS2 = Nothing: Exit Do
'               If CCur((RS2!ped_canped * RS!pro_uniemb)) > canped Or CCur(RS2!ped_canped * RS!pro_uniemb) = canped Then
'               If CCur((RS2!ped_canped / RS!pro_facing) * RS!pro_facsto) > canped Or CCur((RS2!ped_canped / RS!pro_facing) * RS!pro_facsto) = canped Then
               If RS2!ped_canped > CCur((canped / RS!pro_facing) * RS!pro_facsto) Or RS2!ped_canped = CCur((canped / RS!pro_facing) * RS!pro_facsto) Then
                  vaSpread2.MaxRows = vaSpread2.MaxRows + 1
                  vaSpread2.Row = vaSpread2.MaxRows
                  vaSpread2.Col = 1: vaSpread2.Text = RS!pro_codigo
                  vaSpread2.Col = 2: vaSpread2.Text = RS!pro_nombre
                  vaSpread2.Col = 3: vaSpread2.Text = canped
                  vaSpread2.Col = 4: vaSpread2.Text = RS!pro_codtip
                  vaSpread2.Col = 5: vaSpread2.Text = RS!tip_nombre
                  vaSpread2.Col = 6: vaSpread2.Text = RS!uni_nomcor '& " x " & RS!pro_facsto
                  vaSpread2.Col = 7: vaSpread2.Text = RS!pro_facsto
               End If
               RS2.Close: Set RS2 = Nothing
               Exit Do
            ElseIf RS!pro_codigo = RS1!pro_codigo And Val(RS!cantidad) = Val(RS1!cantidad) Then
               vaSpread2.MaxRows = vaSpread2.MaxRows + 1
               vaSpread2.Row = vaSpread2.MaxRows
               vaSpread2.Col = 1: vaSpread2.Text = RS!pro_codigo
               Exit Do
            End If
            RS1.MoveNext
         Loop
         '------- revisar vector si existe codigo producto
         For i = 1 To vaSpread2.MaxRows
             vaSpread2.Row = i
             vaSpread2.Col = 1
             If RS!pro_codigo = vaSpread2.Text Then
                Exit For
             ElseIf i = vaSpread2.MaxRows And RS!cantidad > 0 Then
               '------- Validar si existe productos a rebajar en mensual del mes
               canped = CCur(RS!cantidad)
               RS2.Open "select * from b_minutapedido where ped_codcas='" & LimpiaDato(Trim(fpText.Text)) & "' and ped_anomes=" & fecha & " " & _
                        "and ped_tipped=1 and ped_codpro='" & RS!pro_codigo & "' and ped_fecenv>0", vg_db, adOpenStatic
               If RS2.EOF Then RS2.Close: Set RS2 = Nothing: Exit For
'               If CCur(RS2!ped_canped * RS!pro_uniemb) > canped Or CCur(RS2!ped_canped * RS!pro_uniemb) = canped Then
'               If CCur((RS2!ped_canped / RS!pro_facing) * RS!pro_facsto) > canped Or CCur((RS2!ped_canped / RS!pro_facing) * RS!pro_facsto) = canped Then
               If RS2!ped_canped > CCur((canped / RS!pro_facing) * RS!pro_facsto) Or RS2!ped_canped = CCur((canped / RS!pro_facing) * RS!pro_facsto) Then
                  vaSpread2.MaxRows = vaSpread2.MaxRows + 1
                  vaSpread2.Row = vaSpread2.MaxRows
                  vaSpread2.Col = 1: vaSpread2.Text = RS!pro_codigo
                  vaSpread2.Col = 2: vaSpread2.Text = RS!pro_nombre
                  vaSpread2.Col = 3: vaSpread2.Text = canped
                  vaSpread2.Col = 4: vaSpread2.Text = RS!pro_codtip
                  vaSpread2.Col = 5: vaSpread2.Text = RS!tip_nombre
                  vaSpread2.Col = 6: vaSpread2.Text = RS!uni_nomcor '& " x " & RS!pro_facsto
                  vaSpread2.Col = 7: vaSpread2.Text = RS!pro_facsto
               End If
               RS2.Close: Set RS2 = Nothing
               Exit For
             End If
         Next i
         RS1.MoveFirst
         RS.MoveNext
      Loop
      RS.Close: Set RS = Nothing: RS1.Close: Set RS1 = Nothing
   End If
   uniemb = 0
   vaSpread2.SortKey(1) = 4
   vaSpread2.SortKeyOrder(1) = 1
   vaSpread2.Sort -1, -1, vaSpread2.MaxCols, vaSpread2.MaxRows, SortByRow
   auxcodtip = 0
   For i = 1 To vaSpread2.MaxRows
       vaSpread2.Row = i
       vaSpread2.Col = 2
       If Trim(vaSpread2.Text) <> "" Then
          vaSpread2.Col = 4
          vaSpread2.Col = 1
          '------- Leer ingredientes
          RS.Open "select b_productos.pro_facing, b_ingrediente.ing_codigo, b_ingrediente.ing_precos, " & _
                  "b_ingrediente.ing_nombre, a_unidadmed.unm_nomcor from b_productos, b_productosing, " & _
                  "b_ingrediente, a_unidadmed " & _
                  "where b_productos.pro_codigo=b_productosing.pri_codpro " & _
                  "and   b_productosing.pri_coding=b_ingrediente.ing_codigo " & _
                  "and   b_ingrediente.ing_unimed=a_unidadmed.unm_codigo " & _
                  "and   b_productos.pro_codigo='" & Trim(vaSpread2.Text) & "' " & _
                  "and   b_productos.pro_facing>0 and b_productos.pro_facsto>0", vg_db, adOpenStatic
          If Not RS.EOF Then
             vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
             vaSpread1.Col = -1: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.BackColor = &HC0FFC0
             vaSpread2.Col = 5: vaSpread1.Col = 2: vaSpread1.CellType = 5: vaSpread1.Text = Trim(RS!ing_nombre)
             vaSpread1.Col = 1: vaSpread1.CellType = 5: vaSpread1.Text = Trim(RS!ing_codigo)
             vaSpread2.Col = 3: vaSpread1.Col = 3: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 1: vaSpread1.Text = Format(vaSpread2.Text, fg_Pict(6, 2))
             vaSpread1.Col = 4: vaSpread1.CellType = 5: vaSpread1.Text = Trim(RS!unm_nomcor)
             vaSpread1.Col = 5: vaSpread1.CellType = 5: vaSpread1.Text = ""
             vaSpread1.Col = 6: vaSpread1.CellType = 5: vaSpread1.Text = ""
             vaSpread2.Col = 4: auxcodtip = vaSpread2.Text: vaSpread2.Col = 5: nomtip = vaSpread2.Text
             
             vaSpread1.MaxRows = vaSpread1.MaxRows + 1
             vaSpread1.Row = vaSpread1.MaxRows
             vaSpread2.Col = 1: vaSpread1.Col = 1: vaSpread1.CellType = 5: vaSpread1.Text = vaSpread2.Text
             vaSpread2.Col = 2: vaSpread1.Col = 2: vaSpread1.CellType = 5: vaSpread1.Text = vaSpread2.Text
             vaSpread2.Col = 7: uniemb = vaSpread2.Text
             vaSpread1.Col = 3: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 1: vaSpread1.Text = ""
             vaSpread2.Col = 3: vaSpread1.Col = 4
             vaSpread1.CellType = 3
             vaSpread1.TypeIntegerMin = 1
             vaSpread1.TypeIntegerMax = 9999999
             vaSpread1.TypeHAlign = 1
             vaSpread1.TypeSpin = False
             vaSpread1.TypeIntegerSpinInc = 1
             vaSpread1.TypeIntegerSpinWrap = False
'             If cantidad > 0 And cantidad < 0.5 Then canreal = Format(0.5, fg_Pict(9, vg_DCa)) Else canreal = Format(Round(cantidad, 0), fg_Pict(9, vg_DPr))
'             vaSpread1.Text = IIf(Format(CCur(Redondear(CInt(vaSpread2.Text / RS!pro_facing) * uniemb, 1)), fg_Pict(6, 0)) > 0, Format(CCur(Redondear(CInt(vaSpread2.Text) * uniemb, 1)), fg_Pict(6, 0)), uniemb)
             vaSpread1.Text = IIf((vaSpread2.Text / RS!pro_facing) * uniemb > 0 And (vaSpread2.Text / RS!pro_facing) * uniemb < 0.5, 1, Format(Round((vaSpread2.Text / RS!pro_facing) * uniemb + 0.5, 0), fg_Pict(6, 0)))
             vaSpread1.ForeColor = &HFF0000
             
             vaSpread1.Col = 4: vaSpread2.Col = 6: vaSpread1.Col = 5: vaSpread1.CellType = 5: vaSpread1.Text = vaSpread2.Text
          
          End If
          RS.Close: Set RS = Nothing
'          If vaSpread2.Text <> auxcodtip Then
'             vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
'             vaSpread1.Col = -1: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.BackColor = &HC0FFC0
'             vaSpread2.Col = 5: vaSpread1.Col = 2: vaSpread1.CellType = 5: vaSpread1.Text = vaSpread2.Text
'             vaSpread1.Col = 1: vaSpread1.CellType = 5: vaSpread1.Text = ""
'             vaSpread1.Col = 3: vaSpread1.CellType = 5: vaSpread1.Text = ""
'             vaSpread1.Col = 4: vaSpread1.CellType = 5: vaSpread1.Text = ""
'             vaSpread1.Col = 5: vaSpread1.CellType = 5: vaSpread1.Text = ""
'             vaSpread1.Col = 6: vaSpread1.CellType = 5: vaSpread1.Text = ""
'             vaSpread2.Col = 4: auxcodtip = vaSpread2.Text: vaSpread2.Col = 5: nomtip = vaSpread2.Text
'          End If
                
       End If
   Next i
   If vaSpread1.MaxRows < 1 Then Frame2.Enabled = False: MsgBox "No existe Informaciòn a procesar", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
   Frame2.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(1).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = False
End If
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
fg_carga ""
Me.Height = 6720
Me.Width = 10845
fg_centra Me
MsgTitulo = "Generar Adicionales & Anulaciones"
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.Enabled = False: btnX.ToolTipText = "Confirmar "
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): btnX.Visible = True: btnX.Enabled = False: btnX.ToolTipText = "Enviar"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): btnX.Enabled = False: btnX.ToolTipText = "Imprimir "
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): btnX.Visible = True: btnX.ToolTipText = "Historico Adicionales & Anulaciones"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
Toolbar2.ImageList = Partida.IL1
Set btnX = Toolbar2.Buttons.Add(, "A_InsertarF", , tbrDefault, "A_InsertarF"): btnX.Visible = True: btnX.Enabled = True: btnX.Caption = "Agregar Producto ": btnX.ToolTipText = "Agregar Producto "
Set btnX = Toolbar2.Buttons.Add(, "A_EliminarF", , tbrDefault, "A_EliminarF"): btnX.Visible = True: btnX.Enabled = True: btnX.Caption = "Eliminar Producto ": btnX.ToolTipText = "Eliminar Producto"
Frame2.Enabled = False: vaSpread1.MaxRows = 0
accion = True
Combo1(0).Clear
Combo1(0).AddItem "Adicionales" & Space(150) & "(2)"
Combo1(0).AddItem "Anulaciones" & Space(150) & "(3)"
Combo1(0).ListIndex = -1

vaSpread1.MaxRows = 0
fpDateTime1.Text = Format(Date, "mm/yyyy")

fpText.Enabled = ModCasino
Image1.Enabled = ModCasino
fpText.Text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)

fg_descarga
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    image1_Click
End Select
End Sub

Private Sub fpText_LostFocus()
If fpText.Text = "" Then fpayuda(0).Caption = "": Exit Sub
Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = False: vaSpread1.MaxRows = 0
RS.Open "select * from b_clientes where cli_codigo='" & fpText.Text & "' and cli_tipo=0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
End Sub

Private Sub image1_Click()
vg_left = fpayuda(0).Left + 2300
vg_nombre = "": vg_codigo = ""
B_TabEst.LlenaDatos "b_clientes", "cli_", "Casinos", "Casino"
B_TabEst.Show 1
Me.Refresh
If vg_codigo = "" Then Exit Sub
fpText.Text = vg_codigo: fpayuda(0).Caption = vg_nombre
Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = False: vaSpread1.MaxRows = 0
fpDateTime1.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim CodPro As String, coding As String
Dim i As Integer
Dim canped As Long, fechasis As Long
Dim canmin As Double, cospro As Double, cosrec As Double
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    If vaSpread1.MaxRows < 1 Then Exit Sub
    If Combo1(0).ListIndex = -1 Or Combo1(0).Text = "" Then Exit Sub
    RS.Open "select * from b_clientes where cli_codigo='" & LimpiaDato(Trim(fpText.Text)) & "' and cli_tipo=0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.Text = "": fpayuda(0).Caption = "": MsgBox "No existe casino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    fg_carga (ss)
    fechasis = 0: fechasis = Val(Mid(Date, 7, 4)) & Mid(Date, 4, 2) & Mid(Date, 1, 2)
    '------- Grabar tabla b_minutapedido
    vg_db.BeginTrans
    canmin = 0: CodPro = "": coding = "": canped = 0
    '------- Eliminar pedidos adicionales & Anulaciones
    vg_db.Execute "delete b_minutapedido from b_minutapedido where ped_codcas='" & LimpiaDato(Trim(fpText.Text)) & "' and ped_anomes=" & fecha & " and ped_tipped=" & Val(fg_codigocbo(Combo1, 0, 1, "")) & " and ped_fecenv=0"
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 5
        CodPro = "": canped = 0
        If Trim(vaSpread1.Text) <> "" Then
           CodPro = vaSpread1.Text
           vaSpread1.Col = 1: CodPro = Trim(vaSpread1.Text)
           vaSpread1.Col = 4: canped = vaSpread1.Text
           vg_db.Execute "insert into b_minutapedido (ped_codcas, ped_fecped, ped_anomes, ped_tipped, ped_coding, ped_codpro, ped_canmin, ped_canped, ped_fecenv) " & _
           "values ('" & fpText.Text & "', " & fechasis & ", " & fecha & ", " & Val(fg_codigocbo(Combo1, 0, 1, "")) & ", '" & coding & "', '" & CodPro & "', " & canmin & ", " & canped & ", 0)"
           '------- Actualizar codigo pedido en ingrediente
           vg_db.Execute "update b_ingrediente set ing_codped='" & CodPro & "' where (ing_codped='' or isnull(ing_codped))"
        Else
           vaSpread1.Col = 1: coding = Trim(vaSpread1.Text)
           vaSpread1.Col = 3: canmin = IIf(Trim(vaSpread1.Text) <> "", vaSpread1.Text, 0)
        End If
    Next i
    '------- Actualizar minuta cambio fecha pedido
    vg_db.Execute "update b_minutacambios set cam_fecped=" & fechasis & " where cam_fecped=0"
    vg_anomes = fecha: vg_tipped = Val(fg_codigocbo(Combo1, 0, 1, "")): vg_fecval = fechasis
    vg_db.CommitTrans
    Toolbar1.Buttons(3).Enabled = True: Toolbar1.Buttons(5).Enabled = True
    fg_descarga
Case 3
    RS.Open "select * from b_clientes where cli_codigo='" & LimpiaDato(Trim(fpText.Text)) & "' and cli_tipo=0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.Text = "": fpayuda(0).Caption = "": MsgBox "No existe casino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '------- Validar si existe pedidos Adicionales o cancelaciòn pedidentes
    RS.Open "select distinct ped_fecenv, ped_fecped from  b_minutapedido where ped_codcas='" & LimpiaDato(Trim(fpText.Text)) & "' and ped_anomes=" & fecha & " and ped_tipped=" & Val(fg_codigocbo(Combo1, 0, 1, "")) & " and ped_fecenv=0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información, proceso cancelado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    vg_fecval = RS!ped_fecped
    RS.Close: Set RS = Nothing
    fechasis = 0: fechasis = Val(Mid(Date, 7, 4)) & Mid(Date, 4, 2) & Mid(Date, 1, 2)
    If MsgBox("¿ Esta seguro generar pedido ?", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    fg_carga (ss)
    '------- Actualizar fecha envio minuta pedido
    vg_db.BeginTrans
    vg_db.Execute "update b_minutapedido set ped_fecenv=" & fechasis & " where ped_codcas='" & LimpiaDato(Trim(fpText.Text)) & "' and ped_fecped=" & vg_fecval & " and ped_anomes=" & fecha & " and ped_tipped=" & Val(fg_codigocbo(Combo1, 0, 1, "")) & " and ped_fecenv=0"
    vg_db.CommitTrans
    fg_descarga
    Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = True: Frame2.Enabled = False
    I_PedidosAdiAnu LimpiaDato(Trim(fpText.Text)), Mid(fpDateTime1.Text, 4, 4) & Mid(fpDateTime1.Text, 1, 2), vg_fecval, Val(fg_codigocbo(Combo1, 0, 1, "")), IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 2, "Adicionales", "Anulaciones")
    vaSpread1.MaxRows = 0
Case 5
    If vaSpread1.MaxRows < 1 Then Exit Sub
    RS.Open "select * from b_clientes where cli_codigo='" & LimpiaDato(Trim(fpText.Text)) & "' and cli_tipo=0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.Text = "": fpayuda(0).Caption = "": MsgBox "No existe casino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    RS.Open "select distinct ped_fecenv, ped_fecped from  b_minutapedido where ped_codcas='" & LimpiaDato(Trim(fpText.Text)) & "' and ped_anomes=" & fecha & " and ped_tipped=" & Val(fg_codigocbo(Combo1, 0, 1, "")) & " and ped_fecenv=0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información, proceso cancelado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    vg_fecval = RS!ped_fecped
    RS.Close: Set RS = Nothing
    I_PedidosAdiAnu LimpiaDato(Trim(fpText.Text)), Mid(fpDateTime1.Text, 4, 4) & Mid(fpDateTime1.Text, 1, 2), vg_fecval, Val(fg_codigocbo(Combo1, 0, 1, "")), IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 2, "Adicionales", "Anulaciones")
Case 7
    RS.Open "select * from b_clientes where cli_codigo='" & LimpiaDato(Trim(fpText.Text)) & "' and cli_tipo=0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.Text = "": fpayuda(0).Caption = "": MsgBox "No existe casino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    vg_anomes = 0: vg_tipped = 0: vg_fecval = 0
    B_HiAdAn.LlenarDatos fpText.Text
    B_HiAdAn.Show 1
    Me.Refresh
    If vg_anomes = 0 Then Exit Sub
    fpDateTime1.Text = Mid(vg_anomes, 5, 2) & "/" & Mid(vg_anomes, 1, 4)
    accion = False: Combo1(0).ListIndex = IIf(vg_tipped = 2, 0, 1): accion = True
    vaSpread1.MaxRows = 0
    vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = &HC0FFFF
'   RS.Open "select b_ingrediente.ing_codigo, b_ingrediente.ing_nombre, b_ingrediente.ing_precos, a_unidadmed.unm_nomcor, b_productos.pro_codigo, b_productos.pro_nombre, b_productos.pro_codtip, b_productos.pro_uniemb, a_tipopro.tip_nombre, " & _
'           "a_embalaje.emb_nomcor, b_minutapedido.ped_canped as cantidad, b_minutapedido.ped_canmin, b_productos.pro_propon " & _
'           "from b_minutapedido, b_ingrediente, b_productos, a_tipopro, a_embalaje, a_unidadmed where b_minutapedido.ped_coding=b_ingrediente.ing_codigo and b_minutapedido.ped_codpro=b_productos.pro_codigo " & _
'           "and b_ingrediente.ing_unimed=a_unidadmed.unm_codigo and b_productos.pro_codtip=a_tipopro.tip_codigo and b_productos.pro_codemb=a_embalaje.emb_codigo " & _
'           "and b_minutapedido.ped_codcas='" & LimpiaDato(Trim(fpText.Text)) & "' and b_minutapedido.ped_anomes=" & fecha & " " & _
'           "and b_minutapedido.ped_tipped=" & codtip & " and b_minutapedido.ped_fecenv=0 order by a_tipopro.tip_nombre, b_ingrediente.ing_nombre", vg_db, adOpenStatic
    
    RS.Open "select b_ingrediente.ing_codigo, b_ingrediente.ing_nombre, b_ingrediente.ing_precos, a_unidadmed.unm_nomcor, b_productos.pro_codigo, b_productos.pro_nombre, b_productos.pro_codtip, b_productos.pro_facsto, a_tipopro.tip_nombre, " & _
            "a_unidad.uni_nomcor, b_minutapedido.ped_canped as cantidad, b_minutapedido.ped_canmin, b_productos.pro_propon " & _
            "from  b_minutapedido, b_ingrediente, b_productos, a_tipopro, a_unidad, a_unidadmed " & _
            "where b_minutapedido.ped_codpro=b_productos.pro_codigo " & _
            "and   b_minutapedido.ped_coding=b_ingrediente.ing_codigo " & _
            "and   b_ingrediente.ing_unimed=a_unidadmed.unm_codigo " & _
            "and   b_productos.pro_codtip=a_tipopro.tip_codigo " & _
            "and   b_productos.pro_coduni=a_unidad.uni_codigo " & _
            "and   b_minutapedido.ped_codcas='" & LimpiaDato(Trim(fpText.Text)) & "' " & _
            "and   b_minutapedido.ped_fecped=" & vg_fecval & " " & _
            "and   b_minutapedido.ped_anomes=" & vg_anomes & " " & _
            "and   b_minutapedido.ped_tipped=" & vg_tipped & " " & _
            "order by a_tipopro.tip_nombre, b_ingrediente.ing_nombre", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
    codtip = 0
    Do While Not RS.EOF
'       If RS!pro_codtip <> codtip Then
'          vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
'          vaSpread1.Col = -1: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.BackColor = &HC0FFC0
'          vaSpread1.Col = 2: vaSpread1.CellType = 5: vaSpread1.Text = RS!tip_nombre
'          vaSpread1.Col = 1: vaSpread1.CellType = 5: vaSpread1.Text = ""
'          vaSpread1.Col = 3: vaSpread1.CellType = 5: vaSpread1.Text = ""
'          vaSpread1.Col = 4: vaSpread1.CellType = 5: vaSpread1.Text = ""
'          vaSpread1.Col = 5: vaSpread1.CellType = 5: vaSpread1.Text = ""
'          vaSpread1.Col = 6: vaSpread1.CellType = 5: vaSpread1.Text = ""
'          codtip = RS!pro_codtip: nomtip = RS!tip_nombre
'       End If
       
       If RS!ing_codigo <> codtip Then
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
          vaSpread1.Col = -1: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.BackColor = &HC0FFC0
          vaSpread1.Col = 2: vaSpread1.CellType = 5: vaSpread1.Text = RS!ing_nombre
          vaSpread1.Col = 1: vaSpread1.CellType = 5: vaSpread1.Text = RS!ing_codigo
          vaSpread1.Col = 3: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 1: vaSpread1.Text = Format(RS!ped_canmin, fg_Pict(6, 2))
          vaSpread1.Col = 4: vaSpread1.CellType = 5: vaSpread1.Text = Trim(RS!unm_nomcor)
          vaSpread1.Col = 5: vaSpread1.CellType = 5: vaSpread1.Text = ""
          vaSpread1.Col = 6: vaSpread1.CellType = 5: vaSpread1.Text = RS!ing_precos
          codtip = RS!ing_codigo
       End If
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       vaSpread1.Row = vaSpread1.MaxRows
       vaSpread1.Col = 1: vaSpread1.CellType = 5: vaSpread1.Text = RS!pro_codigo
       vaSpread1.Col = 2: vaSpread1.CellType = 5: vaSpread1.Text = RS!pro_nombre
       vaSpread1.Col = 3: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 1: vaSpread1.Text = "" 'Format(RS!ped_canmin, fg_Pict(6, 2))
       vaSpread1.Col = 4: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 1: vaSpread1.Text = Format(RS!cantidad, fg_Pict(6, 2))
       vaSpread1.Col = 5: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 0: vaSpread1.Text = RS!uni_nomcor & " x " & RS!pro_facsto
       vaSpread1.Col = 6: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 0: vaSpread1.Text = "" 'RS!pro_propon
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Case 9
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then
    MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
    vg_db.RollbackTrans
    Exit Sub
End If
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    vg_nombre = "": vg_codigo = ""
    vg_left = fpayuda(0).Left + 2300
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Gen"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    RS1.Open "select pro_codigo from b_productos where pro_codigo='" & vg_codigo & "' and pro_facing>0 and pro_facsto>0", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "Producto no tiene asignado los factores", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS1.Close: Set RS1 = Nothing
    Dim embalaje As String, CodPro As String, coding As String
    CodPro = vg_codigo
    '------- Validar si existe producto en grilla
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 5: embalaje = "": embalaje = Trim(vaSpread1.Text): vaSpread1.Col = 1
        If Trim(vaSpread1.Text) = Trim(CodPro) And embalaje <> "" Then vaSpread1.SetActiveCell 4, i: MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Next i
    vaSpread1.Row = vaSpread1.ActiveRow
    coding = "": proc2 = ""
    '------- validar si existe mas de un ingrediente
    RS1.Open "select count(pri_coding) as nreg from b_productosing where pri_codpro='" & CodPro & "'", vg_db, adOpenStatic
    If RS1.EOF Or IsNull(RS1!nreg) Or RS1!nreg = 0 Then RS1.Close: Set RS1 = Nothing: Exit Sub
    If RS1!nreg > 1 Then
       vg_nombre = ""
       vg_left = fpayuda(0).Left + 2300
       B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "Proing"
       SendKeys "+{Tab}"
       B_TabEst.Show 1
       If vg_codigo = "" Then RS1.Close: Set RS1 = Nothing: Exit Sub
       coding = vg_codigo
       proc2 = "and  (b_ingrediente.ing_codigo='" & coding & "')"
    End If
    RS1.Close: Set RS1 = Nothing
    proc1 = "select b_productos.pro_codigo, b_productos.pro_nombre, b_productos.pro_facsto, " & _
            "b_ingrediente.ing_codigo, b_ingrediente.ing_precos, b_ingrediente.ing_nombre, a_unidad.uni_nomcor " & _
            "from  b_productos, b_productosing, b_ingrediente, a_unidad  " & _
            "where b_productos.pro_codigo=b_productosing.pri_codpro " & _
            "and   b_ingrediente.ing_codigo=b_productosing.pri_coding " & _
            "and   b_productos.pro_coduni=a_unidad.uni_codigo " & _
            "and   b_productos.pro_codigo='" & CodPro & "'"
    RS1.Open proc1 & proc2, vg_db, adOpenStatic
    If Not RS1.EOF Then
       '------- Validar si existe ingredientes
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i: vaSpread1.Col = 5: embalaje = "": embalaje = Trim(vaSpread1.Text): vaSpread1.Col = 1
           If Trim(vaSpread1.Text) = Trim(RS1!ing_codigo) And embalaje = "" Then
              vaSpread1.MaxRows = vaSpread1.MaxRows + 1
              vaSpread1.InsertRows i + 1, 1
              vaSpread1.Row = i + 1
              vaSpread1.Col = 1: vaSpread1.CellType = 5: vaSpread1.Text = RS1!pro_codigo
              vaSpread1.Col = 2: vaSpread1.CellType = 5: vaSpread1.Text = RS1!pro_nombre
              vaSpread1.Col = 3: vaSpread1.CellType = 5: vaSpread1.Text = ""
              vaSpread1.Col = 4
              vaSpread1.CellType = 3
              vaSpread1.TypeIntegerMin = 1
              vaSpread1.TypeIntegerMax = 9999999
              vaSpread1.TypeHAlign = 1
              vaSpread1.TypeSpin = False
              vaSpread1.TypeIntegerSpinInc = 1
              vaSpread1.TypeIntegerSpinWrap = False
              vaSpread1.Text = Format(0, fg_Pict(6, 0))
              vaSpread1.ForeColor = &HFF0000
              vaSpread1.Col = 5: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 0: vaSpread1.Text = RS1!uni_nomcor '& " x " & RS1!pro_facsto
              vaSpread1.Col = 6: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 0: vaSpread1.Text = 0
              vaSpread1.SetActiveCell 4, i + 1
              RS1.Close: Set RS1 = Nothing
              Exit Sub
           End If
       Next i
       '------- Mover si no existe ingrediente
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
       vaSpread1.Col = -1: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.BackColor = &HC0FFC0
       vaSpread1.Col = 2: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.CellType = 5: vaSpread1.Text = RS1!ing_nombre
       vaSpread1.Col = 1: vaSpread1.CellType = 5: vaSpread1.Text = RS1!ing_codigo
       vaSpread1.Col = 3: vaSpread1.CellType = 5: vaSpread1.Text = ""
       vaSpread1.Col = 4: vaSpread1.CellType = 5: vaSpread1.Text = ""
       vaSpread1.Col = 5: vaSpread1.CellType = 5: vaSpread1.Text = ""
       vaSpread1.Col = 6: vaSpread1.CellType = 5: vaSpread1.Text = RS1!ing_precos
       '------- Mover Productos
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
       vaSpread1.Col = 1: vaSpread1.CellType = 5: vaSpread1.Text = RS1!pro_codigo
       vaSpread1.Col = 2: vaSpread1.CellType = 5: vaSpread1.Text = RS1!pro_nombre
       vaSpread1.Col = 3: vaSpread1.CellType = 5: vaSpread1.Text = ""
       vaSpread1.Col = 4
       vaSpread1.CellType = 3
       vaSpread1.TypeIntegerMin = 1
       vaSpread1.TypeIntegerMax = 9999999
       vaSpread1.TypeHAlign = 1
       vaSpread1.TypeSpin = False
       vaSpread1.TypeIntegerSpinInc = 1
       vaSpread1.TypeIntegerSpinWrap = False
       vaSpread1.Text = Format(0, fg_Pict(6, 0))
       vaSpread1.ForeColor = &HFF0000
       vaSpread1.Col = 5: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 0: vaSpread1.Text = RS1!uni_nomcor '& " x " & RS1!pro_facsto
       vaSpread1.Col = 6: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 0: vaSpread1.Text = 0
       vaSpread1.SetActiveCell 4, vaSpread1.MaxRows
    End If
    RS1.Close: Set RS1 = Nothing
Case 2
    If vaSpread1.MaxRows = 0 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 5
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If Trim(vaSpread1.Text) = "" Then
       vaSpread1.DeleteRows vaSpread1.Row, 1
       vaSpread1.MaxRows = vaSpread1.MaxRows - 1
       For i = vaSpread1.Row To vaSpread1.MaxRows
           vaSpread1.Row = vaSpread1.Row: vaSpread1.Col = 5
           If Trim(vaSpread1.Text) = "" Then Exit For
           vaSpread1.DeleteRows vaSpread1.Row, 1
           vaSpread1.MaxRows = vaSpread1.MaxRows - 1
       Next i
    Else
       i = vaSpread1.Row
       vaSpread1.DeleteRows vaSpread1.Row, 1
       vaSpread1.MaxRows = vaSpread1.MaxRows - 1
       If (vaSpread1.ActiveRow - 1) >= 0 Then
          vaSpread1.Row = IIf(vaSpread1.ActiveRow - 1 = 0, 1, (i - 1))
          vaSpread1.Col = 5
          If Trim(vaSpread1.Text) = "" Then vaSpread1.DeleteRows (vaSpread1.Row), 1: vaSpread1.MaxRows = vaSpread1.MaxRows - 1
       End If
    End If
End Select
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
If ChangeMade = True Then Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = False
End Sub
