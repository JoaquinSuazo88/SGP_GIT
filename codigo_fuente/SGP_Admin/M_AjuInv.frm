VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_AjuInv 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuste de Inventario"
   ClientHeight    =   5835
   ClientLeft      =   1155
   ClientTop       =   2430
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   45
      TabIndex        =   0
      Top             =   345
      Width           =   11520
      Begin EditLib.fpText Combo1 
         Height          =   345
         Index           =   0
         Left            =   4320
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   195
         Width           =   2970
         _Version        =   196608
         _ExtentX        =   5239
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         NoSpecialKeys   =   3
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
         ControlType     =   2
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
      Begin EditLib.fpDateTime Date1 
         Height          =   345
         Index           =   0
         Left            =   1785
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   195
         Width           =   1215
         _Version        =   196608
         _ExtentX        =   2143
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   -1  'True
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
         ControlType     =   2
         Text            =   ""
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
         Left            =   3570
         TabIndex        =   5
         Top             =   240
         Width           =   660
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
         Left            =   1125
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4710
      Left            =   45
      TabIndex        =   2
      Top             =   945
      Width           =   11535
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4365
         Left            =   135
         TabIndex        =   3
         Top             =   225
         Width           =   11295
         _Version        =   393216
         _ExtentX        =   19923
         _ExtentY        =   7699
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
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
         FormulaSync     =   0   'False
         MaxCols         =   7
         MaxRows         =   20
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_AjuInv.frx":0000
         TextTipDelay    =   200
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
   End
End
Attribute VB_Name = "M_AjuInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset
Dim Msgtitulo As String, Est As Boolean

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 6240
Me.Width = 11745
Dim X As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
Est = False
fg_centra Me
Me.HelpContextID = vg_OpcM
Msgtitulo = "Ajuste de Inventario"
Gl_Mo_Botones Me, 5
Gl_Ac_Botones Me, 5, 1, ""
vaSpread1.Row = -1
vaSpread1.Col = 4: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
Dim SQL As String, v_fecinv  As Variant, v_codbod As Long, i As Long, difer As Double, lisnom As String
Dim liscod As String, aju_tipo As String, codaux As Long, z As Long
Date1(0).Text = M_TomInv.Date1(0).Text
Combo1(0).Text = Left(M_TomInv.Combo1(0).List(M_TomInv.Combo1(0).ListIndex), 50)
v_codbod = fg_codigocbo(M_TomInv.Combo1, 0, 10, 0)
v_fecinv = Format(Date1(0).Text, "yyyymmdd")
'--------- Muestra inventario guardado -----------
RS.Open "select dev.dev_codmer, pro.pro_nombre, dev.dev_precos, uni.uni_nombre, dev.dev_canmer, aju.aju_tipo, aju.aju_codigo " & _
        "from b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni, a_tipoajuste aju " & _
        "where tov.tov_rutcli=dev.dev_rutcli and tov.tov_tipdoc=dev.dev_tipdoc and tov.tov_numdoc=dev.dev_numdoc " & _
        "and pro.pro_codigo=dev.dev_codmer and uni.uni_codigo=pro.pro_coduni and tov.tov_codser=aju.aju_codigo " & _
        "and tov.tov_fecemi=Cdate('" & Date1(0).Text & "') and tov_codbod=" & v_codbod & " and tov.tov_tipdoc='AI' and tov.tov_estdoc<>'A' order by dev.dev_numlin", vg_db, adOpenStatic
vaSpread1.MaxRows = 0
i = 1
If Not RS.EOF Then
    Do While Not RS.EOF
        vaSpread1.MaxRows = i
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 1: vaSpread1.Text = RS!dev_codmer
        vaSpread1.Col = 2: vaSpread1.Text = RS!pro_nombre
        vaSpread1.Col = 3: vaSpread1.Text = RS!uni_nombre
        vaSpread1.Col = 4: vaSpread1.Text = IIf(RS!aju_tipo = "D", RS!dev_canmer * -1, RS!dev_canmer): vaSpread1.ForeColor = IIf(RS!aju_tipo = "D", RGB(255, 0, 0), RGB(0, 0, 0))
        vaSpread1.Col = 5: vaSpread1.Text = Format(RS!dev_precos, fg_Pict(9, vg_DPr))
        lisnom = "": liscod = ""
        RS2.Open "select aju_codigo, aju_nombre from a_tipoajuste where aju_tipaju=1 and aju_tipo='" & RS!aju_tipo & "'", vg_db, adOpenStatic
        Do While Not RS2.EOF
            vaSpread1.Col = 6: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(RS2!aju_nombre)
            vaSpread1.Col = 7: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & RS2!aju_codigo
            RS2.MoveNext
        Loop
        RS2.Close: Set RS2 = Nothing
        vaSpread1.Col = 6: vaSpread1.TypeComboBoxList = lisnom
        vaSpread1.Col = 7: vaSpread1.TypeComboBoxList = liscod
        For z = 0 To vaSpread1.TypeComboBoxCount
            vaSpread1.TypeComboBoxCurSel = z
            If Val(vaSpread1.Text) = RS!aju_codigo Then codaux = z: Exit For
            codaux = -1
        Next z
        vaSpread1.Col = 6: vaSpread1.TypeComboBoxCurSel = codaux
        RS.MoveNext: i = i + 1
    Loop
Else
    '--------- Muestra las diferencias del ultimo inventario -----------
    RS1.Open "select tin.tin_codpro, pro.pro_nombre, pro.pro_propon, uni.uni_nombre, tin.tin_stosis, tin.tin_stofis " & _
             "from b_tomainv tin, b_productos pro, a_unidad uni where tin.tin_codpro=pro.pro_codigo " & _
             "and pro.pro_coduni=uni.uni_codigo and tin.tin_fectom=" & v_fecinv & " and tin.tin_codbod=" & v_codbod & " " & _
             "and tin.tin_stosis<>tin.tin_stofis order by pro.pro_nombre", vg_db, adOpenStatic
    Do While Not RS1.EOF
        vaSpread1.MaxRows = i
        vaSpread1.Row = vaSpread1.MaxRows
        difer = Format(RS1!tin_stofis - RS1!tin_stosis, fg_Pict(9, vg_DCa))
        vaSpread1.Col = 1: vaSpread1.Text = RS1!tin_codpro
        vaSpread1.Col = 2: vaSpread1.Text = RS1!pro_nombre
        vaSpread1.Col = 3: vaSpread1.Text = RS1!uni_nombre
        vaSpread1.Col = 4: vaSpread1.Text = difer: vaSpread1.ForeColor = IIf(difer < 0, RGB(255, 0, 0), RGB(0, 0, 0))
        vaSpread1.Col = 5: vaSpread1.Text = Format(RS1!pro_propon, fg_Pict(9, vg_DPr))
        lisnom = "": liscod = ""
        aju_tipo = IIf(difer < 0, "D", "A")
        RS2.Open "select aju_codigo, aju_nombre from a_tipoajuste where aju_tipaju=1 and aju_tipo='" & aju_tipo & "'", vg_db, adOpenStatic
        Do While Not RS2.EOF
            vaSpread1.Col = 6: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(RS2!aju_nombre)
            vaSpread1.Col = 7: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & RS2!aju_codigo
            RS2.MoveNext
        Loop
        RS2.Close: Set RS2 = Nothing
        vaSpread1.Col = 6: vaSpread1.TypeComboBoxList = lisnom
        vaSpread1.Col = 7: vaSpread1.TypeComboBoxList = liscod
        If RS1!tin_stosis = 0 Then
            For z = 0 To vaSpread1.TypeComboBoxCount
                vaSpread1.TypeComboBoxCurSel = z
                If vaSpread1.Text = "3" And RS1!pro_propon = 0 Then codaux = z: vaSpread1.Col = 5: vaSpread1.Lock = False: Exit For
                codaux = -1
            Next z
            vaSpread1.Col = 6: vaSpread1.TypeComboBoxCurSel = codaux
            vaSpread1.Col = 7: vaSpread1.TypeComboBoxCurSel = codaux
            
        End If
        
        RS1.MoveNext: i = i + 1
    Loop
    RS1.Close: Set RS1 = Nothing
    '----------------------------------------------------------------------
End If
vaSpread1.Row = -1
vaSpread1.Col = 6: vaSpread1.Lock = IIf(RS.RecordCount > 0, True, False)
Gl_Ac_Botones Me, 5, IIf(RS.RecordCount > 0, 2, 1), ""
RS.Close: Set RS = Nothing
End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then
    Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)
    Frame1.Left = (Me.Width \ 2) - (Frame1.Width \ 2)
ElseIf Me.WindowState = 0 Then
    Frame2.Left = 45
    Frame1.Left = 45
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim rutcli As String, TipDoc As String, numdoc As Long, CodBod   As Long, fecemi As Date, codser As Long, i As Long, canact As Double, z As Long, nombod As String, fecnum As Long, aumdes As Long
Dim numlin As Long, CodMer As String, canmer As Double, canaux As Double, propon As Double, predoc As Double, descri As String, diablq As Date, folio As Long, total As Double, ptotal As Double
On Error GoTo Man_Error
'If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
CodBod = Val(fg_codigocbo(M_TomInv.Combo1, 0, 10, ""))
nombod = Trim(Combo1(0).Text)
fecemi = Format(Date1(0).Text, "dd/mm/yyyy")
fecnum = Format(Date1(0).Text, "yyyymmdd")
Select Case Button.Index
Case 1 'Graba
    'If Trim(rutcli) = "" Then a = a
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 6
        If vaSpread1.TypeComboBoxCurSel = -1 Then MsgBox "Falta seleccionar concepto...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Next i
    fg_carga ""
    RS1.Open "select aju_codigo, aju_nombre, aju_tipo from a_tipoajuste where aju_tipaju=1", vg_db, adOpenStatic
    If Not RS1.EOF Then
        vg_db.BeginTrans
        Do While Not RS1.EOF
            codser = 0
            For i = 1 To vaSpread1.MaxRows
                vaSpread1.Row = i: vaSpread1.Col = 7
                codser = RS1!aju_codigo
                If codser = Val(vaSpread1.Text) Then
                    rutcli = MuestraCasino(1)
                    aumdes = IIf(RS1!aju_tipo = "D", 0, 1)
                    TipDoc = "AI"
                    RS2.Open "select tov_numdoc from b_totventas where tov_tipdoc='AI' order by tov_numdoc desc", vg_db, adOpenStatic
                    If Not RS2.EOF Then
                        RS2.MoveFirst
                        numdoc = RS2!tov_numdoc + 1
                    Else
                        numdoc = 1
                    End If
                    RS2.Close: Set RS2 = Nothing
                    'Encabezado
                    vg_db.Execute "insert into b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_numinf) " & _
                                  "values ('" & rutcli & "', '" & TipDoc & "', " & numdoc & ", " & CodBod & ", CDate('" & fecemi & "'), " & aumdes & ", " & codser & ", 0, '', '', 0)"
                    'Detalle
                    total = 0
                    For z = 1 To vaSpread1.MaxRows
                        vaSpread1.Row = z: vaSpread1.Col = 7
                        If codser = Val(vaSpread1.Text) Then
                            numlin = z
                            vaSpread1.Col = 1: CodMer = Trim(LimpiaDato(vaSpread1.Text))
                            vaSpread1.Col = 2: descri = Trim(LimpiaDato(vaSpread1.Text))
                            vaSpread1.Col = 4: canmer = Format(vaSpread1.Text, fg_Pict(9, vg_DCa))
                            vaSpread1.Col = 5: predoc = Format(vaSpread1.Text, fg_Pict(9, vg_DPr))
                            canaux = IIf(canmer < 0, canmer * -1, canmer)
                            ptotal = canaux * predoc
                            total = total + ptotal
                            vg_db.Execute "insert into b_detventas (dev_rutcli, dev_tipdoc, dev_numdoc, dev_numlin, dev_codmer, dev_canmin, dev_canmer, dev_predoc, dev_ptotal, dev_descri, dev_mueinv, dev_porcen, dev_precos, dev_coding) " & _
                                          "values ('" & rutcli & "', '" & TipDoc & "', " & numdoc & ", " & numlin & ", '" & CodMer & "', " & canaux & ", " & canaux & ", " & predoc & ", " & ptotal & ", '" & descri & "', 'S', 0, " & predoc & ", '')"
                            vaSpread1.Col = 1
                            vg_db.Execute "update b_tomainv set tin_stosis=tin_stosis+" & canmer & " where tin_fectom=" & fecnum & " " & _
                                          "and tin_codbod=" & CodBod & " and tin_codpro='" & Trim(vaSpread1.Text) & "'"
                            '------- Reemplaza precio promedio si es inventario inicial -----------
                            vaSpread1.Col = 5: propon = Round(vaSpread1.Text, vg_DPr)
                            vaSpread1.Col = 7
                            If vaSpread1.Text = "3" Then
                                Dim PMP As Double, auxCanmer As Double, auxPropon As Double
'                                RS2.Open "SELECT pro_facing, pro_coding FROM b_productos WHERE pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
'                                'PMP Ingrediente
'                                If Not RS2.EOF Then
'                                    coding = IIf(IsNull(RS2!pro_coding), "", RS2!pro_coding)
'                                    auxCanmer = 0: auxPropon = 0
'                                    RS3.Open "SELECT sum(bod.bod_canmer) as canmer FROM b_productos pro, b_bodegas bod " & _
'                                             "WHERE bod.bod_codpro=pro.pro_codigo and pro.pro_coding='" & coding & "'", vg_db, adOpenStatic
'                                    If Not RS3.EOF Then auxCanmer = IIf(IsNull(RS3!canmer), 0, RS3!canmer)
'                                    RS3.Close: Set RS3 = Nothing
'                                    'RS3.Open "SELECT sum(ing_precos) as propon FROM b_ingrediente WHERE ing_codigo='" & coding & "'", vg_db, adOpenStatic
'                                    'RS3.Open "SELECT sum(pro_propon/pro_facing) as propon FROM b_productos WHERE pro_coding='" & coding & "'", vg_db, adOpenStatic
'                                    RS3.Open "SELECT sum((pro.pro_propon/pro.pro_facing)*bod_canmer) as propon FROM b_productos pro, b_bodegas bod " & _
'                                             "WHERE pro.pro_codigo=bod.bod_codpro and pro.pro_coding='" & coding & "'", vg_db, adOpenStatic
'                                    If Not RS3.EOF Then auxPropon = IIf(IsNull(RS3!propon), 0, RS3!propon)
'                                    RS3.Close: Set RS3 = Nothing
'                                    'PMP = Val(((auxPropon * auxCanmer) + ((predoc / RS2!pro_facing) * canmer)) / (auxCanmer + canmer))
'                                    PMP = Val((auxPropon + ((predoc / RS2!pro_facing) * canmer)) / (auxCanmer + canmer))
'                                    vg_db.Execute "update b_ingrediente set ing_precos=" & PMP & " where ing_codigo='" & coding & "'"
'                                End If
'                                RS2.Close: Set RS2 = Nothing
'                                'PMP Producto
'                                RS3.Open "SELECT sum(bod.bod_canmer) as canmer FROM b_productos pro, b_bodegas bod " & _
'                                         "WHERE bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
'                                If Not RS3.EOF Then auxCanmer = IIf(IsNull(RS3!canmer), 0, RS3!canmer)
'                                RS3.Close: Set RS3 = Nothing
'                                RS3.Open "SELECT pro_propon as propon FROM b_productos WHERE pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
'                                If Not RS3.EOF Then auxPropon = IIf(IsNull(RS3!propon), 0, RS3!propon)
'                                RS3.Close: Set RS3 = Nothing
'                                PMP = Round(((auxPropon * auxCanmer) + (predoc * canmer)) / (auxCanmer + canmer), vg_DPr)
'                                vg_db.Execute "update b_productos set pro_propon=" & PMP & " where pro_codigo='" & CodMer & "'"
'                                'PMP Toma
'                                vg_db.Execute "update b_tomainv set tin_propon=" & PMP & " where tin_fectom=" & fecnum & " " & _
'                                              "and tin_codbod=" & CodBod & " and tin_codpro='" & CodMer & "'"
'                                'Actuliza codigo compra y pedido de ultimo producto para ingrediente
'                                vg_db.Execute "update b_ingrediente set ing_codped='" & CodMer & "', ing_codcom='" & CodMer & "' where ing_codigo='" & coding & "'"
'
                            
                                RS2.Open "Select pro_facing From b_productos Where pro_codigo='" & CodMer & "'", vg_db, adOpenStatic

                                'PMP Ingrediente
                                If Not RS2.EOF Then
                                    auxCanmer = 0: auxPropon = 0
                                    RS3.Open "Select Sum(bod.bod_canmer) As canmer From b_productos pro, b_bodegas bod " & _
                                             "Where bod.bod_codpro=pro.pro_codigo And pro.pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
                                    If Not RS3.EOF Then auxCanmer = IIf(IsNull(RS3!canmer), 0, RS3!canmer)
                                    RS3.Close: Set RS3 = Nothing
                                    RS3.Open "Select Sum((pro.pro_propon/pro.pro_facing)*bod_canmer) as propon From b_productos pro, b_bodegas bod " & _
                                             "Where pro.pro_codigo=bod.bod_codpro And pro.pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
                                    If Not RS3.EOF Then auxPropon = IIf(IsNull(RS3!propon), 0, RS3!propon)
                                    RS3.Close: Set RS3 = Nothing
                                    PMP = Val((auxPropon + ((predoc / RS2!pro_facing) * canmer)) / (auxCanmer + canmer))
                                    vg_db.Execute "Update b_ingrediente ing, b_productosing pri Set ing.ing_precos=" & PMP & " " & _
                                                  "Where pri.pri_coding=ing.ing_codigo And pri.pri_codpro='" & CodMer & "'"
                                    'PMP Producto
                                    RS3.Open "Select Sum(bod.bod_canmer) As canmer From b_productos pro, b_bodegas bod " & _
                                             "Where bod.bod_codpro=pro.pro_codigo And pro.pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
                                    If Not RS3.EOF Then auxCanmer = IIf(IsNull(RS3!canmer), 0, RS3!canmer)
                                    RS3.Close: Set RS3 = Nothing
                                    RS3.Open "Select pro_propon As propon From b_productos Where pro_codigo='" & CodMer & "'", vg_db, adOpenStatic
                                    If Not RS3.EOF Then auxPropon = IIf(IsNull(RS3!propon), 0, RS3!propon)
                                    RS3.Close: Set RS3 = Nothing
                                    PMP = Val(((auxPropon * auxCanmer) + (predoc * canmer)) / (auxCanmer + canmer))
                                    vg_db.Execute "Update b_productos Set pro_propon=" & PMP & " Where pro_codigo='" & CodMer & "'"
                                    'Actuliza codigo compra y pedido de ultimo producto para ingrediente
                                    vg_db.Execute "Update b_ingrediente ing, b_productosing pri Set ing_codped='" & CodMer & "', ing_codcom='" & CodMer & "' " & _
                                                  "Where pri.pri_coding=ing.ing_codigo And pri.pri_codpro='" & CodMer & "'"
                                End If
                                RS2.Close: Set RS2 = Nothing
                            
                            
                            End If
                            '------- Control de Stock ---------
                            ValidaBod CodBod, Trim(LimpiaDato(CodMer))
                            canact = 0
                            RS2.Open "select bod_canmer from b_bodegas where bod_codpro='" & Trim(LimpiaDato(CodMer)) & "' and bod_codbod=" & CodBod, vg_db, adOpenStatic
                            If Not RS2.EOF Then
                                Do While Not RS2.EOF
                                    canact = RS2!bod_canmer + canmer
                                    RS2.MoveNext
                                Loop
                                vg_db.Execute "update b_bodegas set bod_canmer=" & canact & " " & _
                                              "where bod_codpro='" & Trim(LimpiaDato(CodMer)) & "' and bod_codbod=" & CodBod
                            End If
                            RS2.Close: Set RS2 = Nothing
                        End If
                    Next z
                    'Total
                    vg_db.Execute "update b_totventas set tov_totdoc=" & total & " where tov_rutcli='" & rutcli & "' " & _
                                    "and tov_tipdoc='AI' and tov_numdoc=" & numdoc
                    Exit For
                End If
            Next i
            RS1.MoveNext
        Loop
        vg_db.CommitTrans
    Else
        RS1.Close: Set RS1 = Nothing
        fg_descarga
        Exit Sub
    End If
    RS1.Close: Set RS1 = Nothing
    fg_descarga
    vaSpread1.Row = -1
    vaSpread1.Col = 5: vaSpread1.Lock = True
    vaSpread1.Col = 6: vaSpread1.Lock = True
    Gl_Ac_Botones Me, 5, 2, ""
    I_Ajuste CodBod & "|" & nombod, CVDate(fecemi)
Case 3 'Imprimir
    I_Ajuste CodBod & "|" & nombod, CVDate(fecemi)
Case 6 'Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
vg_swpegreceta = 0
If Err = 3034 Then Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
If Est Then Exit Sub
Select Case Col
Case 6
    Dim indice As Long
    vaSpread1.Row = Row
    vaSpread1.Col = 6: indice = vaSpread1.TypeComboBoxCurSel
    vaSpread1.Col = 7: vaSpread1.TypeComboBoxCurSel = indice
    vaSpread1.Col = 1
    RS1.Open "select pro_propon from b_productos where pro_codigo='" & Trim(vaSpread1.Text) & "'", vg_db, adOpenStatic
    If Not RS1.EOF Then
        If Trim(vaSpread1.Text) = "3" And RS1!pro_propon = 0 Then
            vaSpread1.Col = 5: vaSpread1.Lock = False
            vaSpread1.SetActiveCell 5, Row - 1
        Else
            vaSpread1.Col = 5: vaSpread1.Lock = True
            vaSpread1.Text = RS1!pro_propon
        End If
    End If
    RS1.Close: Set RS1 = Nothing
End Select
End Sub

'Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
'If KeyAscii <> 13 Then Exit Sub
'Select Case vaSpread1.ActiveCol
'Case 5
'    vaSpread1.SetActiveCell 6, vaSpread1.ActiveRow
'End Select
'End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If Row = 0 Then Exit Sub
Select Case Col
Case 6
vaSpread1.Row = Row: vaSpread1.Col = Col
If vaSpread1.ColWidth(Col) > (vaSpread1.MaxTextCellWidth - 2) Then Exit Sub
TipWidth = vaSpread1.MaxTextColWidth(Col)
ShowTip = True
MultiLine = 2
TipText = vaSpread1.Text
          
End Select
End Sub

