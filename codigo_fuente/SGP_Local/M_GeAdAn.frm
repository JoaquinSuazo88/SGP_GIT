VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
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
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Oculta Ingrediente"
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
         Height          =   270
         Index           =   0
         Left            =   3090
         TabIndex        =   15
         Top             =   900
         Width           =   2610
      End
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
         Text            =   "03/2009"
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
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   285
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2685
         Picture         =   "M_GeAdAn.frx":0236
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
      DisplayRowHeaders=   0   'False
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
      SpreadDesigner  =   "M_GeAdAn.frx":0540
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
Dim Fecha As Long
Dim accion As Boolean
Dim Msgtitulo As String

Private Sub Check1_Click(index As Integer)
On Error GoTo Man_Error
If accion = False Then Exit Sub
'------- Actualizar parametro generación pedido mensual
vg_db.BeginTrans
vg_db.Execute "UPDATE a_param SET par_valor = '" & IIf(Check1(0).Value = 1, 1, 0) & "' WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'ingpedadi'"
vg_db.CommitTrans
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Visible = False
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 5
    If Trim(vaSpread1.text) = "" Then vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
Next i
vaSpread1.SetActiveCell 1, 1
vaSpread1.Visible = True
Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Combo1_Click(index As Integer)
Dim sql1 As String
If accion = False Then Exit Sub
Dim codtip As Integer, auxcodtip As Integer
Dim nomtip As String
Dim canped As Double, canpea As Double
If Combo1(0).ListIndex = -1 Or Combo1(0).text = "" Then Exit Sub
codtip = Val(fg_codigocbo(Combo1, 0, 1, "")): auxcodtip = 0
vaSpread1.MaxRows = 0: canped = 0
vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = &HC0FFFF
Fecha = 0: Fecha = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)

'------- Validar si existe pedidos adicionales & anulaciones
RS.Open "SELECT DISTINCT ped_fecenv " & _
        "FROM b_minutapedido " & _
        "WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " " & _
        "AND   ped_tipped = " & codtip & " AND ped_fecenv = 0", vg_db, adOpenStatic
If Not RS.EOF Then
   fecenv = RS!ped_fecenv
   RS.Close: Set RS = Nothing
   If fecenv > 0 Then Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False) Else Toolbar1.Buttons(1).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Frame2.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(3).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(5).Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
   RS.Open "SELECT b.ing_codigo, b.ing_nombre, (SELECT DISTINCT cpi_precos FROM b_contlispreing WHERE cpi_coding = b.ing_codigo AND cpi_cencos = '" & MuestraCasino(1) & "'), f.unm_nomcor, c.pro_codigo, c.pro_nombre, c.pro_codtip, c.pro_facsto, d.tip_nombre, " & _
           "e.uni_nomcor, a.ped_canped AS cantidad, a.ped_canmin " & _
           "FROM  b_minutapedido a, b_ingrediente b, b_productos c, a_tipopro d, a_unidad e, a_unidadmed f " & _
           "WHERE a.ped_coding = b.ing_codigo AND a.ped_codpro = c.pro_codigo " & _
           "AND   b.ing_unimed = f.unm_codigo AND c.pro_codtip = d.tip_codigo AND c.pro_coduni = e.uni_codigo AND c.pro_facsto > 0 " & _
           "AND   a.ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND a.ped_anomes = " & Fecha & " " & _
           "AND   a.ped_tipped = " & codtip & " AND a.ped_fecenv = 0 ORDER BY d.tip_nombre, b.ing_nombre", vg_db, adOpenStatic
   codtip = 0
   If Not RS.EOF Then
      Do While Not RS.EOF
         If RS!ing_codigo <> codtip Then
            vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
            vaSpread1.Col = -1: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.BackColor = &HC0FFC0
            vaSpread1.Col = 2: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!ing_nombre
            vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!ing_codigo
            vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(RS!ped_canmin, fg_Pict(6, 2))
            vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = Trim(RS!unm_nomcor)
            vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
            vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!cpi_precos
            codtip = RS!ing_codigo
         End If
      
         vaSpread1.MaxRows = vaSpread1.MaxRows + 1
         vaSpread1.Row = vaSpread1.MaxRows
         vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!pro_codigo
         vaSpread1.Col = 2: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!pro_nombre
         vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = ""
         vaSpread1.Col = 4
         vaSpread1.CellType = CellTypeNumber
         vaSpread1.TypeNumberDecPlaces = 2
         vaSpread1.TypeNumberMin = 1
         vaSpread1.TypeNumberMax = 9999999
         vaSpread1.TypeHAlign = TypeHAlignRight
         vaSpread1.TypeSpin = False
         vaSpread1.TypeIntegerSpinInc = 1
         vaSpread1.TypeIntegerSpinWrap = False
         
         vaSpread1.text = Format(RS!cantidad, fg_Pict(6, 2))
         vaSpread1.ForeColor = &HFF0000
      
         vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = RS!uni_nomcor '& " x " & RS!pro_facsto
         vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = ""
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing: Exit Sub
Else
   RS.Close: Set RS = Nothing
   '------- Validar productos vigente
   ValidarProductoVigente
   If codtip = 2 Then
      sql1 = IIf(vg_tipbase = "1", " val(mid(a.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) ")
      RS.Open "SELECT c.pro_codigo, c.pro_nombre, c.pro_codtip, c.pro_facsto, e.tip_nombre, f.uni_nomcor, sum(b.cam_canpro) AS cantidad " & _
              "FROM b_minuta a, b_minutacambios b, b_productos c, b_contlistpreing d, a_tipopro e, a_unidad f " & _
              "WHERE b.cam_codmin = a.min_codigo AND b.cam_fecmin = a.min_fecmin AND b.cam_codpro = d.cpi_coding AND d.cpi_cencos = '" & MuestraCasino(1) & "' " & _
              "AND d.cpi_codped = c.pro_codigo AND c.pro_codtip = e.tip_codigo AND c.pro_coduni = f.uni_codigo AND c.pro_facsto > 0 AND a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND " & sql1 & " = " & Fecha & " " & _
              "AND b.cam_fecped = 0 GROUP BY c.pro_codigo, c.pro_nombre, c.pro_codtip, c.pro_facsto, e.tip_nombre, f.uni_nomcor ORDER BY e.tip_nombre, c.pro_nombre", vg_db, adOpenStatic
      If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe adicionales", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
      sql1 = IIf(vg_tipbase = "1", " val(mid(c.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),c.min_fecmin),1,6)) ")
      RS1.Open "SELECT e.pro_codigo, SUM(d.mid_numrac*(b.red_canpro/a.rec_basrac)) AS cantidad " & _
               "FROM b_receta a, b_recetadet b, b_minuta c, b_minutadet d, b_productos e, b_contlistpreing f " & _
               "WHERE d.mid_codigo IN (SELECT cam_codmin FROM b_minutacambios WHERE cam_fecped = 0) " & _
               "AND   d.mid_codigo = c.min_codigo AND d.mid_codrec = b.red_codigo AND d.mid_tiprec = b.red_tiprec AND ((b.red_tiprec <> 0 AND b.red_cencos = '" & MuestraCasino(1) & "') OR (b.red_tiprec = 0 AND b.red_cencos = '0')) AND b.red_codigo = a.rec_codigo AND b.red_codpro = f.cpi_coding AND f.cpi_codped = e.pro_codigo AND f.cpi_cencos = '" & MuestraCasino(1) & "' AND c.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
               "AND   " & sql1 & " = " & Fecha & " AND d.mid_tipmin = '1' GROUP BY e.pro_codigo", vg_db, adOpenStatic
      If RS1.EOF Then RS1.Close: Set RS1 = Nothing: RS.Close: Set RS = Nothing: Exit Sub
      vaSpread2.MaxCols = 7: vaSpread2.MaxRows = 0
      Do While Not RS.EOF
         Do While Not RS1.EOF
            If RS!pro_codigo = RS1!pro_codigo And Val(RS!cantidad) > Val(RS1!cantidad) Then
               canped = CCur(RS!cantidad - RS1!cantidad)
               vaSpread2.MaxRows = vaSpread2.MaxRows + 1
               vaSpread2.Row = vaSpread2.MaxRows
               vaSpread2.Col = 1: vaSpread2.text = RS!pro_codigo
               vaSpread2.Col = 2: vaSpread2.text = RS!pro_nombre
               vaSpread2.Col = 3: vaSpread2.text = canped
               vaSpread2.Col = 4: vaSpread2.text = RS!pro_codtip
               vaSpread2.Col = 5: vaSpread2.text = RS!tip_nombre
               vaSpread2.Col = 6: vaSpread2.text = RS!uni_nomcor '& " x " & RS!pro_facsto
               vaSpread2.Col = 7: vaSpread2.text = RS!pro_facsto
               Exit Do
            ElseIf RS!pro_codigo = RS1!pro_codigo And (Val(RS!cantidad) = Val(RS1!cantidad) Or Val(RS!cantidad) < Val(RS1!cantidad)) Then
               vaSpread2.MaxRows = vaSpread2.MaxRows + 1
               vaSpread2.Row = vaSpread2.MaxRows
               vaSpread2.Col = 1: vaSpread2.text = RS!pro_codigo
               Exit Do
            End If
            RS1.MoveNext
         Loop
         '------- revisar vector si existe codigo producto
         For i = 1 To vaSpread2.MaxRows
             vaSpread2.Row = i
             vaSpread2.Col = 1
             If RS!pro_codigo = vaSpread2.text Then
                Exit For
             ElseIf i = vaSpread2.MaxRows And RS!cantidad > 0 Then
               canped = CCur(RS!cantidad)
               vaSpread2.MaxRows = vaSpread2.MaxRows + 1
               vaSpread2.Row = vaSpread2.MaxRows
               vaSpread2.Col = 1: vaSpread2.text = RS!pro_codigo
               vaSpread2.Col = 2: vaSpread2.text = RS!pro_nombre
               vaSpread2.Col = 3: vaSpread2.text = canped
               vaSpread2.Col = 4: vaSpread2.text = RS!pro_codtip
               vaSpread2.Col = 5: vaSpread2.text = RS!tip_nombre
               vaSpread2.Col = 6: vaSpread2.text = RS!uni_nomcor
               vaSpread2.Col = 7: vaSpread2.text = RS!pro_facsto
               Exit For
             End If
         Next i
         RS1.MoveFirst
         RS.MoveNext
      Loop
      RS.Close: Set RS = Nothing: RS1.Close: Set RS1 = Nothing
   ElseIf codtip = 3 Then
      sql1 = IIf(vg_tipbase = "1", " val(mid(c.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),c.min_fecmin),1,6)) ")
      RS.Open "SELECT f.pro_codigo, f.pro_nombre, f.pro_facing, f.pro_codtip, f.pro_facsto, g.tip_nombre, " & _
              "h.uni_nomcor, SUM(d.mid_numrac*(b.red_canpro/a.rec_basrac)) AS cantidad " & _
              "FROM b_receta a, b_recetadet b, b_minuta c, b_minutadet d, b_contlistpreing e, b_productos f, a_tipopro g, a_unidad h " & _
              "WHERE d.mid_codigo IN (SELECT cam_codmin FROM b_minutacambios WHERE cam_fecped=0) " & _
              "AND d.mid_codigo = c.min_codigo AND d.mid_codrec = b.red_codigo AND d.mid_tiprec = b.red_tiprec AND ((b.red_tiprec <> 0 AND b.red_cencos = '" & MuestraCasino(1) & "') OR (b.red_tiprec = 0 AND b.red_cencos = '0')) " & _
              "AND b.red_codigo = a.rec_codigo AND b.red_codpro = e.cpi_coding AND e.cpi_codped = f.pro_codigo AND e.cpi_cencos = '" & MuestraCasino(1) & "' " & _
              "AND f.pro_codtip = g.tip_codigo AND f.pro_coduni=h.uni_codigo AND f.pro_facsto>0 " & _
              "AND c.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND " & sql1 & " = " & Fecha & " " & _
              "AND d.mid_tipmin='1' GROUP BY f.pro_codigo, f.pro_nombre, f.pro_facing, f.pro_codtip, f.pro_facsto, g.tip_nombre, h.uni_nomcor " & _
              "ORDER BY g.tip_nombre, f.pro_nombre", vg_db, adOpenStatic
      If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe anulaciones", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
      sql1 = IIf(vg_tipbase = "1", " val(mid(a.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) ")
      RS1.Open "SELECT d.pro_codigo, SUM(b.cam_canpro) AS cantidad FROM b_minuta a, b_minutacambios b, b_contlistpreing c, b_productos d " & _
               "WHERE b.cam_codmin = a.min_codigo AND b.cam_fecmin = a.min_fecmin " & _
               "AND   b.cam_codpro = c.cpi_coding AND c.cpi_cencos = '" & MuestraCasino(1) & "' AND c.cpi_codped = d.pro_codigo " & _
               "AND   a.min_cencos = '" & LimpiaDato(Trim(fpText.text)) & "' AND " & sql1 & " = " & Fecha & " " & _
               "AND   b.cam_fecped = 0 GROUP BY d.pro_codigo", vg_db, adOpenStatic
      If RS1.EOF Then RS1.Close: Set RS1 = Nothing: RS.Close: Set RS = Nothing: Exit Sub
      vaSpread2.MaxCols = 7: vaSpread2.MaxRows = 0
      fecenv = 1: Toolbar1.Buttons(1).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = False
      Do While Not RS.EOF
         Do While Not RS1.EOF
            If RS!pro_codigo = RS1!pro_codigo And Val(RS!cantidad) > Val(RS1!cantidad) Then
               '------- Validar si existe productos a rebajar en mensual del mes
               canped = CCur(RS!cantidad - RS1!cantidad)
               RS2.Open "SELECT * FROM b_minutapedido WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " " & _
                        "AND ped_tipped = 1 AND ped_codpro = '" & RS!pro_codigo & "' AND ped_fecenv > 0", vg_db, adOpenStatic
               If RS2.EOF Then RS2.Close: Set RS2 = Nothing: Exit Do
               canpea = IIf(Int((canped / RS!pro_facing) / RS!pro_facsto) <> ((canped / RS!pro_facing) / RS!pro_facsto), Int((canped / RS!pro_facing) / RS!pro_facsto) + 1, Round((canped / RS!pro_facing) / RS!pro_facsto, 0)) * RS!pro_facsto
               If RS2!ped_canped > canpea Or RS2!ped_canped = canpea Then
                  vaSpread2.MaxRows = vaSpread2.MaxRows + 1
                  vaSpread2.Row = vaSpread2.MaxRows
                  vaSpread2.Col = 1: vaSpread2.text = RS!pro_codigo
                  vaSpread2.Col = 2: vaSpread2.text = RS!pro_nombre
                  vaSpread2.Col = 3: vaSpread2.text = canped
                  vaSpread2.Col = 4: vaSpread2.text = RS!pro_codtip
                  vaSpread2.Col = 5: vaSpread2.text = RS!tip_nombre
                  vaSpread2.Col = 6: vaSpread2.text = RS!uni_nomcor
                  vaSpread2.Col = 7: vaSpread2.text = RS!pro_facsto
               End If
               RS2.Close: Set RS2 = Nothing
               Exit Do
            End If
            RS1.MoveNext
         Loop
         '------- revisar vector si existe codigo producto
         For i = 1 To vaSpread2.MaxRows
             vaSpread2.Row = i
             vaSpread2.Col = 1
             If RS!pro_codigo = vaSpread2.text Then
                Exit For
             ElseIf i = vaSpread2.MaxRows And RS!cantidad > 0 Then
               '------- Validar si existe productos a rebajar en mensual del mes
               canped = CCur(RS!cantidad)
               RS2.Open "SELECT * FROM b_minutapedido WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " " & _
                        "AND ped_tipped = 1 AND ped_codpro = '" & RS!pro_codigo & "' AND ped_fecenv > 0", vg_db, adOpenStatic
               If RS2.EOF Then RS2.Close: Set RS2 = Nothing: Exit For
               canpea = IIf(Int((canped / RS!pro_facing) / RS!pro_facsto) <> ((canped / RS!pro_facing) / RS!pro_facsto), Int((canped / RS!pro_facing) / RS!pro_facsto) + 1, Round((canped / RS!pro_facing) / RS!pro_facsto, 0)) * RS!pro_facsto
               If RS2!ped_canped > canpea Or RS2!ped_canped = canpea Then
                  vaSpread2.MaxRows = vaSpread2.MaxRows + 1
                  vaSpread2.Row = vaSpread2.MaxRows
                  vaSpread2.Col = 1: vaSpread2.text = RS!pro_codigo
                  vaSpread2.Col = 2: vaSpread2.text = RS!pro_nombre
                  vaSpread2.Col = 3: vaSpread2.text = canped
                  vaSpread2.Col = 4: vaSpread2.text = RS!pro_codtip
                  vaSpread2.Col = 5: vaSpread2.text = RS!tip_nombre
                  vaSpread2.Col = 6: vaSpread2.text = RS!uni_nomcor
                  vaSpread2.Col = 7: vaSpread2.text = RS!pro_facsto
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
       If Trim(vaSpread2.text) <> "" Then
          vaSpread2.Col = 4
          vaSpread2.Col = 1
          '------- Leer ingredientes
          RS.Open "SELECT b_productos.pro_facing, b_productos.pro_facsto, b_ingrediente.ing_codigo, " & _
                  "b_ingrediente.ing_nombre, a_unidadmed.unm_nomcor FROM b_productos, b_productosing, " & _
                  "b_ingrediente, a_unidadmed " & _
                  "WHERE b_productos.pro_codigo = b_productosing.pri_codpro " & _
                  "AND   b_productosing.pri_coding = b_ingrediente.ing_codigo " & _
                  "AND   b_ingrediente.ing_unimed = a_unidadmed.unm_codigo " & _
                  "AND   b_productos.pro_codigo = '" & Trim(vaSpread2.text) & "' " & _
                  "AND   b_productos.pro_facing > 0 AND b_productos.pro_facsto > 0", vg_db, adOpenStatic
          If Not RS.EOF Then
             vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
             vaSpread1.Col = -1: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.BackColor = &HC0FFC0
             vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
             vaSpread2.Col = 5: vaSpread1.Col = 2: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = Trim(RS!ing_nombre)
             vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = Trim(RS!ing_codigo)
             vaSpread2.Col = 3: vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(vaSpread2.text, fg_Pict(6, 2))
             vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = Trim(RS!unm_nomcor)
             vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
             vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
             vaSpread2.Col = 4: auxcodtip = vaSpread2.text: vaSpread2.Col = 5: nomtip = vaSpread2.text
             
             vaSpread1.MaxRows = vaSpread1.MaxRows + 1
             vaSpread1.Row = vaSpread1.MaxRows
             vaSpread2.Col = 1: vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = vaSpread2.text
             vaSpread2.Col = 2: vaSpread1.Col = 2: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = vaSpread2.text
             vaSpread2.Col = 7: uniemb = vaSpread2.text
             vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = ""
             vaSpread2.Col = 3: vaSpread1.Col = 4
             vaSpread1.CellType = CellTypeNumber
             vaSpread1.TypeNumberDecPlaces = 2
             vaSpread1.TypeNumberMin = 1
             vaSpread1.TypeNumberMax = 9999999
             vaSpread1.TypeHAlign = TypeHAlignRight
             vaSpread1.TypeSpin = False
             vaSpread1.TypeIntegerSpinInc = 1
             vaSpread1.TypeIntegerSpinWrap = False
             
             vaSpread1.text = IIf(Int((vaSpread2.text / RS!pro_facing) / RS!pro_facsto) <> ((vaSpread2.text / RS!pro_facing) / RS!pro_facsto), Int((vaSpread2.text / RS!pro_facing) / RS!pro_facsto) + 1, Round((vaSpread2.text / RS!pro_facing) / RS!pro_facsto, 0)) * RS!pro_facsto
             vaSpread1.ForeColor = &HFF0000
             
             vaSpread1.Col = 4: vaSpread2.Col = 6: vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = vaSpread2.text
          End If
          RS.Close: Set RS = Nothing
       End If
   Next i
   If vaSpread1.MaxRows < 1 Then Frame2.Enabled = False: MsgBox "No existe Informaciòn a procesar", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
   Frame2.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(1).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False): Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = False
End If
End Sub

Private Sub Form_Activate()
fg_descarga
'-------> Traer fecha cierre día
TraerFechaCierre
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
fg_carga ""
Me.Height = 6720
Me.Width = 10845
fg_centra Me
Msgtitulo = "Generar Adicionales & Anulaciones"
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
fpDateTime1.text = Format(Date, "mm/yyyy")

fpText.Enabled = ModCasino
Image1.Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
Check1(0).Value = IIf(0 = (fg_CambiaChar(GetParametro("ingpedadi"), ";", "','")), 0, 1)
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
    Image1_Click
End Select
End Sub

Private Sub fpText_LostFocus()
If fpText.text = "" Then fpayuda(0).Caption = "": Exit Sub
Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = False: vaSpread1.MaxRows = 0
RS.Open "SELECT * FROM b_clientes where cli_codigo = '" & fpText.text & "' AND cli_tipo = 0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
End Sub

Private Sub Image1_Click()
vg_left = fpayuda(0).Left + 2300
vg_nombre = "": vg_codigo = ""
B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
B_TabEst.Show 1
Me.Refresh
If vg_codigo = "" Then Exit Sub
fpText.text = vg_codigo: fpayuda(0).Caption = vg_nombre
Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = False: vaSpread1.MaxRows = 0
fpDateTime1.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim CodPro As String, coding As String
Dim i As Integer
Dim fechasis As Long
Dim canmin As Double, cospro As Double, cosrec As Double, canped As Double
On Error GoTo Man_Error
Select Case Button.index
Case 1
    If vaSpread1.MaxRows < 1 Then Exit Sub
    If Combo1(0).ListIndex = -1 Or Combo1(0).text = "" Then Exit Sub
    RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText.text)) & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda(0).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    fg_carga (ss)
    fechasis = 0: fechasis = Val(Mid(Date, 7, 4)) & Mid(Date, 4, 2) & Mid(Date, 1, 2)
    '------- Grabar tabla b_minutapedido
    vg_db.BeginTrans
    canmin = 0: CodPro = "": coding = "": canped = 0
    '------- Eliminar pedidos adicionales & Anulaciones
    vg_db.Execute "DELETE b_minutapedido FROM b_minutapedido WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " AND ped_tipped = " & Val(fg_codigocbo(Combo1, 0, 1, "")) & " AND ped_fecenv = 0"
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 5
        CodPro = "": canped = 0
        If Trim(vaSpread1.text) <> "" Then
           CodPro = vaSpread1.text
           vaSpread1.Col = 1: CodPro = Trim(vaSpread1.text)
           vaSpread1.Col = 4: canped = vaSpread1.text
           vg_db.Execute "INSERT INTO b_minutapedido (ped_codcas, ped_fecped, ped_anomes, ped_tipped, ped_coding, ped_codpro, ped_canmin, ped_canped, ped_fecenv, ped_stoact, ped_proped) " & _
           "VALUES ('" & fpText.text & "', " & fechasis & ", " & Fecha & ", " & Val(fg_codigocbo(Combo1, 0, 1, "")) & ", '" & coding & "', '" & CodPro & "', " & canmin & ", " & canped & ", 0, 0, 0)"
           '------- Actualizar codigo pedido en ingrediente
           vg_db.Execute "UPDATE b_contlispreing SET cpi_codped = '" & CodPro & "' WHERE cpi_cencos = '" & MuestraCasino(1) & "' AND (cpi_codped = '' OR (cpi_codped) IS NULL)"
        Else
           vaSpread1.Col = 1: coding = Trim(vaSpread1.text)
           vaSpread1.Col = 3: canmin = IIf(Trim(vaSpread1.text) <> "", vaSpread1.text, 0)
        End If
    Next i
    '------- Actualizar minuta cambio fecha pedido
    vg_db.Execute "UPDATE b_minutacambios SET cam_fecped = " & fechasis & " WHERE cam_fecped = 0"
    vg_anomes = Fecha: vg_tipped = Val(fg_codigocbo(Combo1, 0, 1, "")): vg_fecval = fechasis
    vg_db.CommitTrans
    Toolbar1.Buttons(3).Enabled = True: Toolbar1.Buttons(5).Enabled = True
    fg_descarga
Case 3
    RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText.text)) & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda(0).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    '------- Validar si existe pedidos Adicionales o cancelaciòn pedidentes
    RS.Open "SELECT DISTINCT ped_fecenv, ped_fecped FROM b_minutapedido WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " AND ped_tipped = " & Val(fg_codigocbo(Combo1, 0, 1, "")) & " AND ped_fecenv = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    vg_fecval = RS!ped_fecped
    RS.Close: Set RS = Nothing
    fechasis = 0: fechasis = Val(Mid(Date, 7, 4)) & Mid(Date, 4, 2) & Mid(Date, 1, 2)
    If MsgBox("¿ Esta seguro generar pedido ?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    fg_carga (ss)
    '------- Actualizar fecha envio minuta pedido
    vg_db.BeginTrans
    vg_db.Execute "UPDATE b_minutapedido SET ped_fecenv = " & fechasis & " WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_fecped = " & vg_fecval & " AND ped_anomes = " & Fecha & " AND ped_tipped = " & Val(fg_codigocbo(Combo1, 0, 1, "")) & " AND ped_fecenv = 0"
    vg_db.CommitTrans
    fg_descarga
    Toolbar1.Buttons(1).Enabled = False: Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = True: Frame2.Enabled = False
    I_PedidosAdiAnu LimpiaDato(Trim(fpText.text)), Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2), vg_fecval, Val(fg_codigocbo(Combo1, 0, 1, "")), IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 2, "Adicionales", "Anulaciones"), IIf(Check1(0).Value = 1, 1, 0)
    vaSpread1.MaxRows = 0
Case 5
    If vaSpread1.MaxRows < 1 Then Exit Sub
    RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText.text)) & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda(0).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    RS.Open "SELECT DISTINCT ped_fecenv, ped_fecped FROM b_minutapedido WHERE ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' AND ped_anomes = " & Fecha & " AND ped_tipped = " & Val(fg_codigocbo(Combo1, 0, 1, "")) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información, proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    vg_fecval = RS!ped_fecped
    RS.Close: Set RS = Nothing
    I_PedidosAdiAnu LimpiaDato(Trim(fpText.text)), Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2), vg_fecval, Val(fg_codigocbo(Combo1, 0, 1, "")), IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 2, "Adicionales", "Anulaciones"), IIf(Check1(0).Value = 1, 1, 0)
Case 7
    RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & LimpiaDato(Trim(fpText.text)) & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText.text = "": fpayuda(0).Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    vg_anomes = 0: vg_tipped = 0: vg_fecval = 0
    B_HiAdAn.LlenarDatos fpText.text
    B_HiAdAn.Show 1
    Me.Refresh
    If vg_anomes = 0 Then Exit Sub
    fpDateTime1.text = Mid(vg_anomes, 5, 2) & "/" & Mid(vg_anomes, 1, 4)
    Fecha = Mid(vg_anomes, 1, 4) & Mid(vg_anomes, 5, 2)
    accion = False: Combo1(0).ListIndex = IIf(vg_tipped = 2, 0, 1): accion = True
    vaSpread1.MaxRows = 0
    vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = &HC0FFFF
    RS.Open "SELECT b.ing_codigo, b.ing_nombre, (SELECT DISTINCT cpi_precos FROM b_contlistpreing WHERE cpi_coding = b.ing_codigo AND cpi_cencos = '" & MuestraCasino(1) & "'), f.unm_nomcor, c.pro_codigo, c.pro_nombre, c.pro_codtip, c.pro_facsto, d.tip_nombre, " & _
            "e.uni_nomcor, a.ped_canped AS cantidad, a.ped_canmin " & _
            "FROM  b_minutapedido a, b_ingrediente b, b_productos c, a_tipopro d, a_unidad e, a_unidadmed f " & _
            "WHERE a.ped_codpro = c.pro_codigo " & _
            "AND   a.ped_coding = b.ing_codigo " & _
            "AND   b.ing_unimed = f.unm_codigo " & _
            "AND   c.pro_codtip = d.tip_codigo " & _
            "AND   c.pro_coduni = e.uni_codigo " & _
            "AND   a.ped_codcas = '" & LimpiaDato(Trim(fpText.text)) & "' " & _
            "AND   a.ped_fecped = " & vg_fecval & " " & _
            "AND   a.ped_anomes = " & vg_anomes & " " & _
            "AND   a.ped_tipped = " & vg_tipped & " " & _
            "ORDER BY d.tip_nombre, b.ing_nombre", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
    codtip = 0
    Do While Not RS.EOF
       If RS!ing_codigo <> codtip Then
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
          vaSpread1.Col = -1: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.BackColor = &HC0FFC0
          vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
          vaSpread1.Col = 2: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!ing_nombre
          vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!ing_codigo
          vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(RS!ped_canmin, fg_Pict(6, 2))
          vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = Trim(RS!unm_nomcor)
          vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
          vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!cpi_precos
          codtip = RS!ing_codigo
       End If
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       vaSpread1.Row = vaSpread1.MaxRows
       vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!pro_codigo
       vaSpread1.Col = 2: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!pro_nombre
       vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = ""
       vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(RS!cantidad, fg_Pict(6, 2))
       vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = RS!uni_nomcor
       vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = ""
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
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.index
Case 1
    vg_nombre = "": vg_codigo = ""
    vg_left = fpayuda(0).Left + 2300
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Pst"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    RS1.Open "SELECT pro_codigo FROM b_productos WHERE pro_codigo = '" & vg_codigo & "' AND pro_facing > 0 and pro_facsto > 0", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "Producto no tiene asignado los factores", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS1.Close: Set RS1 = Nothing
    Dim embalaje As String, CodPro As String, coding As String
    CodPro = vg_codigo
    '------- Validar si existe producto en grilla
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 5: embalaje = "": embalaje = Trim(vaSpread1.text): vaSpread1.Col = 1
        If Trim(vaSpread1.text) = Trim(CodPro) And embalaje <> "" Then vaSpread1.SetActiveCell 4, i: MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
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
    proc1 = "SELECT b_productos.pro_codigo, b_productos.pro_nombre, b_productos.pro_facsto, " & _
            "b_ingrediente.ing_codigo, (SELECT DISTINCT cpi_precos FROM b_contlispreing WHERE cpi_coding = b_ingrediente.ing_codigo AND cpi_cencos = '" & MuestraCasino(1) & "'), b_ingrediente.ing_nombre, a_unidad.uni_nomcor " & _
            "FROM  b_productos, b_productosing, b_ingrediente, a_unidad  " & _
            "WHERE b_productos.pro_codigo = b_productosing.pri_codpro " & _
            "AND   b_ingrediente.ing_codigo = b_productosing.pri_coding " & _
            "AND   b_productos.pro_coduni = a_unidad.uni_codigo " & _
            "AND   b_productos.pro_codigo = '" & CodPro & "'"
    RS1.Open proc1 & proc2, vg_db, adOpenStatic
    If Not RS1.EOF Then
       '------- Validar si existe ingredientes
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i: vaSpread1.Col = 5: embalaje = "": embalaje = Trim(vaSpread1.text): vaSpread1.Col = 1
           If Trim(vaSpread1.text) = Trim(RS1!ing_codigo) And embalaje = "" Then
              vaSpread1.MaxRows = vaSpread1.MaxRows + 1
              vaSpread1.InsertRows i + 1, 1
              vaSpread1.Row = i + 1
              vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS1!pro_codigo
              vaSpread1.Col = 2: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS1!pro_nombre
              vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
              vaSpread1.Col = 4
              vaSpread1.CellType = 3
              vaSpread1.TypeIntegerMin = 1
              vaSpread1.TypeIntegerMax = 9999999
              vaSpread1.TypeHAlign = TypeHAlignRight
              vaSpread1.TypeSpin = False
              vaSpread1.TypeIntegerSpinInc = 1
              vaSpread1.TypeIntegerSpinWrap = False
              vaSpread1.text = Format(0, fg_Pict(6, 0))
              vaSpread1.ForeColor = &HFF0000
              vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = RS1!uni_nomcor '& " x " & RS1!pro_facsto
              vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = 0
              vaSpread1.SetActiveCell 4, i + 1
              RS1.Close: Set RS1 = Nothing
              Exit Sub
           End If
       Next i
       '------- Mover si no existe ingrediente
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
       vaSpread1.Col = -1: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.BackColor = &HC0FFC0
       vaSpread1.RowHidden = IIf(Check1(0).Value = 1, True, False)
       vaSpread1.Col = 2: vaSpread1.Font.Bold = True: vaSpread1.Font.Size = 9: vaSpread1.CellType = 5: vaSpread1.text = RS1!ing_nombre
       vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS1!ing_codigo
       vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS1!cpi_precos
       '------- Mover Productos
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
       vaSpread1.Col = 1: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS1!pro_codigo
       vaSpread1.Col = 2: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS1!pro_nombre
       vaSpread1.Col = 3: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = ""
       vaSpread1.Col = 4
       vaSpread1.CellType = 3
       vaSpread1.TypeIntegerMin = 1
       vaSpread1.TypeIntegerMax = 9999999
       vaSpread1.TypeHAlign = TypeHAlignRight
       vaSpread1.TypeSpin = False
       vaSpread1.TypeIntegerSpinInc = 1
       vaSpread1.TypeIntegerSpinWrap = False
       vaSpread1.text = Format(0, fg_Pict(6, 0))
       vaSpread1.ForeColor = &HFF0000
       vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = RS1!uni_nomcor '& " x " & RS1!pro_facsto
       vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = 0
       vaSpread1.SetActiveCell 4, vaSpread1.MaxRows
    End If
    RS1.Close: Set RS1 = Nothing
Case 2
    If vaSpread1.MaxRows = 0 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 5
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    If Trim(vaSpread1.text) = "" Then
       vaSpread1.DeleteRows vaSpread1.Row, 1
       vaSpread1.MaxRows = vaSpread1.MaxRows - 1
       For i = vaSpread1.Row To vaSpread1.MaxRows
           vaSpread1.Row = vaSpread1.Row: vaSpread1.Col = 5
           If Trim(vaSpread1.text) = "" Then Exit For
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
          If Trim(vaSpread1.text) = "" Then vaSpread1.DeleteRows (vaSpread1.Row), 1: vaSpread1.MaxRows = vaSpread1.MaxRows - 1
       End If
    End If
End Select
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
If ChangeMade = True Then Toolbar1.Buttons(3).Enabled = False: Toolbar1.Buttons(5).Enabled = False
End Sub
