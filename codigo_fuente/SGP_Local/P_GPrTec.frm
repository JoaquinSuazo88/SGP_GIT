VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form P_GPrTec 
   Caption         =   "Generación Maestro Producto Hacia Tecfood"
   ClientHeight    =   6765
   ClientLeft      =   3585
   ClientTop       =   2385
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Index           =   1
      Left            =   30
      TabIndex        =   0
      Top             =   600
      Width           =   6015
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1005
         Index           =   3
         Left            =   570
         TabIndex        =   1
         Top             =   240
         Width           =   4725
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            ItemData        =   "P_GPrTec.frx":0000
            Left            =   1680
            List            =   "P_GPrTec.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   2865
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   3
            Top             =   600
            Width           =   2895
            _Version        =   196608
            _ExtentX        =   5106
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
         Begin VB.Label Label1 
            Caption         =   "Buscar Texto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   150
            TabIndex        =   5
            Top             =   660
            Width           =   1470
         End
         Begin VB.Label Label1 
            Caption         =   "Buscar Columna"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   150
            TabIndex        =   4
            Top             =   345
            Width           =   1485
         End
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   1
         Left            =   660
         TabIndex        =   6
         Top             =   5880
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
         Height          =   4455
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   1410
         Width           =   5655
         _Version        =   393216
         _ExtentX        =   9975
         _ExtentY        =   7858
         _StockProps     =   64
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
         MaxCols         =   7
         MaxRows         =   2
         SpreadDesigner  =   "P_GPrTec.frx":002A
         TextTip         =   2
         TextTipDelay    =   0
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "P_GPrTec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim Est As Boolean, estado As Boolean
Dim codtip As Long, ibusca As Long, i As Long, j As Long
Dim aAp As String

Private Sub Combo1_Click(Index As Integer)
If Est Then Exit Sub
Select Case Index
Case 1
    If Combo1(1).ListIndex = 2 Then
       vg_left = Frame1(1).Left + Combo1(1).Left + 1920
       B_ArbEst.MoverDatosTvwDir "a_tipopro", "tip_", "Familia del Producto"
       B_ArbEst.Show 1
       Me.Refresh
       If Val(vg_codigo) = 0 Then Combo1(1).ListIndex = 1: fptnombre(1).Enabled = True: fptnombre(1).text = "": Exit Sub
       codtip = Val(vg_codigo)
       fptnombre(1).text = vg_nombre
       fptnombre(1).Enabled = False
   Else
      fptnombre(1).Enabled = True
      fptnombre(1).text = ""
   End If
   If vaSpread1(1).MaxRows > 0 Then vaSpread1(1).SetFocus
End Select
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 7275
Me.Width = 6225
Me.HelpContextID = vg_OpcM
Msgtitulo = "Generación Maestro Producto Hacia Tecfood"
fg_centra Me
Est = True: ibusca = 0
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): btnX.Visible = True: btnX.ToolTipText = "Procesar": btnX.Enabled = True
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
Combo1(1).ListIndex = 1
MoverDatoGrilla
Est = False
SendKeys "+{Tab}"
End Sub

Private Sub fpTnombre_Change(Index As Integer)
Select Case Index
Case 1
    If LimpiaDato(Trim(fptnombre(1).text)) & Chr(KeyAscii) = "" Then Exit Sub
    findstring = Trim(fptnombre(1).text)
    If fptnombre(1).text = "" Then
       vaSpread1(1).Visible = False
       swactiva = 0
       For i = 1 To vaSpread1(1).MaxRows
           vaSpread1(1).Row = i
           vaSpread1(1).Col = 1
           vaSpread1(1).text = "0"
           If swactiva = 0 Then swactiva = 1
       Next i
       vaSpread1(1).Visible = True
    Else
       swactiva = 0
       vaSpread1(1).Visible = False
       irow = 0
       For i = 1 To vaSpread1(1).MaxRows
           vaSpread1(1).Row = i
           vaSpread1(1).Col = 1
           vaSpread1(1).text = "0"
           If Combo1(1).ItemData(Combo1(1).ListIndex) = 0 Or Combo1(1).ItemData(Combo1(1).ListIndex) = 1 Then
              vaSpread1(1).Col = IIf(Combo1(1).ItemData(Combo1(1).ListIndex) = 0, 2, 3)
           Else
              findstring = Trim(Str(codtip))
              vaSpread1(1).Col = 4
           End If
           sourcestring = Trim(vaSpread1(1).text)
           indactivo = UCase(Trim(sourcestring)) Like "*" & UCase(findstring) & "*"
           If indactivo = -1 Then
              If swactiva = 0 Then swactiva = 1
              If vaSpread1(1).RowHidden = True Then
                 vaSpread1(1).RowHidden = False
              ElseIf vaSpread1(1).RowHidden = True Then
                 vaSpread1(1).RowHidden = False
              End If
              irow = irow + 1
           Else
              If vaSpread1(1).RowHidden = False Then vaSpread1(1).RowHidden = True
           End If
       Next i
       vaSpread1(1).Visible = True
       End If
End Select
End Sub

Private Sub fptnombre_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 34 And irow > 0 Then vaSpread1(Index).SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    If vaSpread1(1).MaxRows < 1 Then Exit Sub
    Dim i As Long, j As Long
    Dim isel As Boolean
    isel = False
    For i = 1 To vaSpread1(1).MaxRows
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1
        If vaSpread1(1).text = "1" And vaSpread1(1).RowHidden = False Then isel = True: Exit For
    Next i
    If isel = False Then MsgBox "Debe Seleccionar A lo menor un producto", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    fg_carga ""
    Frame1(1).Enabled = False
    Bar1(1).Visible = True
    Bar1(1).Value = 0
    Dim codpro As String, nompro As String, StrFam As String, StrFamb As String, indtec As String, dBo As String, prodact As String
    Dim codfam As Long, fampr1 As Long, fampr2 As Long, fampr3 As Long, coduni As Long, fecven As String
    Dim profac As Double
    Dim opgraba As Boolean
    dBo = dir_trabajo + BaseDeDato
    For j = 1 To vaSpread1(1).MaxRows
        Bar1(1).Value = Val((j / vaSpread1(1).MaxRows) * 100)
        vaSpread1(1).Row = j: vaSpread1(1).Col = 1
        If vaSpread1(1).text = "1" And vaSpread1(1).RowHidden = False Then
           DoEvents
'tecfood           vg_dbtec.BeginTrans
           vg_db.BeginTrans
           vaSpread1(1).Col = 2: codpro = Trim(Str(vaSpread1(1).text))
           vaSpread1(1).Col = 3: nompro = Trim(vaSpread1(1).text)
           vaSpread1(1).Col = 4: codfam = vaSpread1(1).text
           vaSpread1(1).Col = 5: coduni = vaSpread1(1).text
           vaSpread1(1).Col = 6: profac = Val(vaSpread1(1).text)
           vaSpread1(1).Col = 7: fecven = Val(vaSpread1(1).text)
           prodact = "N"
           If fecven < Val(Format(Date, "yyyymmdd")) And fecven > 0 Then prodact = "S"
           vaSpread1(1).SetActiveCell 2, vaSpread1(1).Row
           RS.Open "select a.ing_codigo, a.ing_nombre, a.ing_unimed from b_ingrediente a, b_productosing b where a.ing_codigo=b.pri_coding and b.pri_codpro='" & codpro & "'", vg_db, adOpenStatic
           If Not RS.EOF Then
              Do While Not RS.EOF
                 opgraba = False
                 StrFam = fg_BuscaCodArbol(codfam, "a_tipopro", "tip_codigo")
                 If Len(StrFam) <> 0 Then
                    Do While InStr(StrFam, ";") <> 0
                       StrFamb = Mid(StrFam, 1, InStr(StrFam, ";") - 1)
                       StrFam = IIf(Len(StrFam) > InStr(StrFam, ";"), Mid(StrFam, InStr(StrFam, ";") + 1), "")
                       fampr1 = Val(Mid(StrFamb, 1, InStr(StrFamb, "&") - 1)): StrFamb = Mid(StrFamb, InStr(StrFamb, "&") + 1)
                       fampr2 = Val(Mid(StrFamb, 1, InStr(StrFamb, "&") - 1)): StrFamb = Mid(StrFamb, InStr(StrFamb, "&") + 1)
                       If Val(Mid(StrFamb, 1)) = 0 Then opgraba = True: Exit Do
                       fampr3 = Val(Mid(StrFamb, 1))
                    Loop
                 End If
                 If Not opgraba Then
                     If RS!ing_codigo <> "720" And RS!ing_codigo <> "762" Then
                        ' Grabar ingrediente sgp & generico tecfood
                        ' Grabar maestro producto tecfood, fam.prod1 - fam.prod2 - fam.prod3 - cod.ing. - nom.ing. - cod.unmed - cod.prod - nom.prod - fac.conversión - opcion - objeto
'tecfood                        If GrabarProductoTecfood(CStr(fampr1), CStr(fampr2), CStr(fampr3), Trim(RS!ing_codigo), Trim(RS!ing_nombre), RS!ing_unimed, coduni, "", "", profac, prodact) Then RS1.Close: Set RS1 = Nothing: vg_dbtec.RollbackTrans: vg_db.RollbackTrans: Exit Sub
                        '------- Validar si existe producto generico, borrar aportes nutricionales
                        indtec = ""
'tecfood                        RS1.Open "select cdproduto from produto where cdnvproduto='5' and cdprodinte='" & "I" & RS!ing_codigo & "'", vg_dbtec, adOpenStatic
'tecfood                        If Not RS1.EOF Then indtec = Trim(RS1!cdproduto): vg_dbtec.Execute "delete from nutrprod where cdproduto='" & RS1!cdproduto & "'"
'tecfood                        RS1.Close: Set RS1 = Nothing
                        '------- Fi validar si existe producto generico, borrar aportes nutricionales
                        '------ Grabar datos tecfood aportes nutricionales
    '                    If Trim(indtec) <> "" Then vg_dbtec.Execute "insert into nutrprod (cdproduto, cdnutriente, qtnutrprod) values ('" & Trim(indtec) & "', '" & fg_pone_cero(CodN, 3) & "', " & CanN & ")"
'tecfood                        If Trim(indtec) <> "" Then
'tecfood                           RS1.Open "select pnu_codapo, pnu_canapo  from b_productonut where pnu_codpro='" & RS!ing_codigo & "'", vg_db, adOpenStatic
'tecfood                           If Not RS1.EOF Then
'tecfood                              Do While Not RS1.EOF
'tecfood                                 vg_dbtec.Execute "insert into nutrprod (cdproduto, cdnutriente, qtnutrprod) values ( '" & Trim(indtec) & "', '" & fg_pone_cero(RS1!pnu_codapo, 3) & "', " & RS1!pnu_canapo & ")"
'tecfood                                 RS1.MoveNext
'tecfood                              Loop
'tecfood                           End If
'tecfood                           RS1.Close: Set RS1 = Nothing
'tecfood                        End If
                        ' Grabar productos sgp & formato compas tecfood
'tecfood                        If GrabarProductoTecfood(CStr(fampr1), CStr(fampr2), CStr(fampr3), Trim(RS!ing_codigo), Trim(RS!ing_nombre), RS!ing_unimed, coduni, codpro, NomPro, profac, prodact) Then RS.Close: Set RS = Nothing: vg_dbtec.RollbackTrans: vg_db.RollbackTrans: Exit Sub
                     ElseIf RS!ing_codigo = "720" Or RS!ing_codigo = "762" Then
                        ' Grabar maestro producto tecfood, fam.prod1 - fam.prod2 - fam.prod3 - cod.ing. - nom.ing. - cod.unmed - cod.prod - nom.prod - fac.conversión - opcion - objeto
'tecfood                        If GrabarProductoTecfood(CStr(fampr1), CStr(fampr2), CStr(fampr3), Trim(RS!ing_codigo), Trim(RS!ing_nombre), RS!ing_unimed, coduni, codpro, NomPro, profac, prodact) Then RS.Close: Set RS = Nothing: vg_dbtec.RollbackTrans: vg_db.RollbackTrans: Exit Sub
                     End If
                 End If
                 RS.MoveNext
              Loop
           End If
           RS.Close: Set RS = Nothing
'tecfood           vg_dbtec.CommitTrans
           vg_db.CommitTrans
        End If
    Next j
    fg_descarga
    Bar1(1).Visible = False
    MsgBox "Generación Finalizado Sin Problema", vbInformation + vbOKOnly, Msgtitulo
    Frame1(1).Enabled = True
Case 3
    Me.Hide
    Unload Me
End Select

Exit Sub
fg_descarga
Frame1(1).Enabled = True
Bar1(1).Visible = False
RS.Close: Set RS = Nothing
Man_Error:
Select Case Err
Case 35764
'tecfood    vg_dbtec.RollbackTrans
    vg_db.RollbackTrans
    DoEvents
    For i = 1 To 1000000
    Next i
    Resume
Case 76
'tecfood    vg_dbtec.RollbackTrans
    vg_db.RollbackTrans
    Resume Next
Case -2147467259
'tecfood    vg_dbtec.RollbackTrans
    vg_db.RollbackTrans
    MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
    Exit Sub
Case 3034
'tecfood    vg_dbtec.RollbackTrans
    vg_db.RollbackTrans: Exit Sub
End Select
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_ButtonClicked(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If Index = 1 Then Exit Sub
vaSpread1(1).Row = Row
Select Case Col
Case 1
    If Row = 0 Or Row = -1 Then x = vaSpread1(1).MaxRows: j = 1 Else x = vaSpread1(1).Row: j = vaSpread1(1).Row
    fg_descarga
End Select
End Sub

Sub MoverDatoGrilla()
On Error GoTo Man_Error
fg_carga "": estado = True: i = 1
'------- Mover casinos
If Est Then
   vaSpread1(1).MaxRows = 0
   RS.Open "select pro_codigo, pro_nombre, pro_codtip, pro_coduni, pro_facsto, pro_fecven from b_productos order by pro_nombre", vg_db, adOpenStatic
   If Not RS.EOF Then
      Do While Not RS.EOF
         vaSpread1(1).MaxRows = vaSpread1(1).MaxRows + 1
         vaSpread1(1).Row = vaSpread1(1).MaxRows
              
         vaSpread1(1).Col = 2
         vaSpread1(1).TypeHAlign = TypeHAlignLeft
         vaSpread1(1).TypeSpin = False
         vaSpread1(1).TypeIntegerSpinInc = 1
         vaSpread1(1).TypeIntegerSpinWrap = False
         vaSpread1(1).text = RS!pro_codigo

         vaSpread1(1).Col = 3
         vaSpread1(1).TypeHAlign = TypeHAlignLeft
         vaSpread1(1).text = Trim(RS!pro_nombre)
         
         vaSpread1(1).Col = 4
         vaSpread1(1).TypeHAlign = TypeHAlignLeft
         vaSpread1(1).text = RS!pro_codtip
         
         vaSpread1(1).Col = 5
         vaSpread1(1).TypeHAlign = TypeHAlignLeft
         vaSpread1(1).text = RS!pro_coduni
         
         vaSpread1(1).Col = 6
         vaSpread1(1).TypeHAlign = TypeHAlignLeft
         vaSpread1(1).text = RS!pro_facsto
         
         vaSpread1(1).Col = 7
         vaSpread1(1).TypeHAlign = TypeHAlignLeft
         vaSpread1(1).text = RS!pro_fecven
         
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
End If
vaSpread1(1).SetActiveCell 1, i
fg_descarga
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub

Private Sub vaSpread1_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
If Col = 1 And Row = 0 Then vaSpread1(Index).Row = -1: vaSpread1(Index).Col = 1: vaSpread1(Index).text = IIf(vaSpread1(Index).Value = "1", "0", "1")
End Sub

Private Sub vaSpread1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Or KeyCode = 13 Then Exit Sub
If TeclasNoPermitidas(KeyCode) = True Then fptnombre(Index).text = IIf(KeyCode = 8, fptnombre(Index).text, fptnombre(Index).text & Chr(KeyCode)): fptnombre(Index).SetFocus: fptnombre(Index).SelStart = Len(fptnombre(Index).text)
End Sub
