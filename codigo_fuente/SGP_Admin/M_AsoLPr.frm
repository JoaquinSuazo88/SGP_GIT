VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_AsoLPr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asociar Lista de Precio"
   ClientHeight    =   7005
   ClientLeft      =   2760
   ClientTop       =   2115
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11655
      Begin VB.Frame Frame12 
         Height          =   435
         Left            =   3720
         TabIndex        =   8
         Top             =   5760
         Width           =   3525
         Begin VB.TextBox TextCai1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   9
            Top             =   135
            Width           =   3420
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4875
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   11415
         _Version        =   393216
         _ExtentX        =   20135
         _ExtentY        =   8599
         _StockProps     =   64
         ButtonDrawMode  =   1
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
         MaxRows         =   1000000
         SpreadDesigner  =   "M_AsoLPr.frx":0000
         VirtualMode     =   -1  'True
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   6000
         Visible         =   0   'False
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   4
         Top             =   435
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
      Begin VB.Shape Shape1 
         BackColor       =   &H80000018&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2805
         TabIndex        =   5
         Top             =   435
         Width           =   6495
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2340
         Picture         =   "M_AsoLPr.frx":1AD5
         Top             =   360
         Width           =   480
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2835
         TabIndex        =   7
         Top             =   480
         Width           =   6495
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_AsoLPr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim modo As String, codigo As Long
Dim Msgtitulo As String
Dim vLisPre() As Variant
Dim Est As Boolean

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Dim RS As New ADODB.Recordset
Me.HelpContextID = vg_OpcM
Me.Height = 7515
Me.Width = 12045
Msgtitulo = "Asociar Lista de Precio"
fg_centra Me
modo = "": codigo = 0: Est = True
Gl_Mo_Botones Me, 1
vaSpread1.MaxRows = 0
Gl_Ac_Botones Me, 1, 3, modo
'MoverDatosGrillas
Est = False
End Sub

Private Sub MoverDatosGrillas()
Dim RS As New ADODB.Recordset
Dim v_inicio As Long, v_final As Long, i As Long, ii As Long, j As Long, z As Long, auxsub As Long
fg_carga ""
'-------> Mover lista precio vector
Set RS = vg_db.Execute("sgpadm_s_listaprecio 4, 0, 0, '" & vg_NUsr & "'")
i = 1
If Not RS.EOF Then
   ReDim vLisPre(RS!nReg, 2)
   Do While Not RS.EOF
      vLisPre(i, 1) = RS!lpr_codigo
      vLisPre(i, 2) = RS!lpr_nombre
      i = i + 1
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing

'-------> Mover datos grilla
With vaSpread1
    .Visible = False
    .MaxRows = 0
    .Row = -1: .Col = -1
    .BackColor = Shape1(0).FillColor
    Bar1(0).Visible = True: Bar1(0).Value = 0
    auxsub = 0
    Set RS = vg_db.Execute("sgpadm_s_subsegmento 2, " & fpLongInteger1(0).Value & ", 0, ''")
    If Not RS.EOF Then
        .MaxRows = RS!nReg
        ii = 1
        Do While Not RS.EOF
            Bar1(0).Value = Val((ii / .MaxRows) * 100)
            .Row = ii
            If RS!sub_codigo <> auxsub Then
               .Col = 2: .Value = RS!sub_codigo & " - " & Trim(RS!sub_nombre)
               auxsub = RS!sub_codigo
            End If
            .Col = 1: .Value = RS!sub_codigo
            .Col = 3: .Value = RS!reg_codigo & " - " & Trim(RS!reg_nombre)
            .Col = 4: .Value = RS!reg_codigo
            If i > 1 Then
               lisnom = "": liscod = "": encuentra = False
               '-------> Mover lista precio
               For j = 1 To UBound(vLisPre)
                   .Col = 5: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vLisPre(j, 2))
                   .Col = 6: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vLisPre(j, 1)
                   .Col = 5: .TypeComboBoxList = lisnom
                   .Col = 6: .TypeComboBoxList = liscod
               Next j
               .Col = 6
               codaux = -1
               For z = 0 To .TypeComboBoxCount
                   .TypeComboBoxCurSel = z
                   If .text = RS!lpr_codigo Then codaux = z: Exit For
                   codaux = -1
               Next z
               .Col = 5: .TypeComboBoxCurSel = codaux
            End If
            .Col = 7: .Value = RS!sub_codigo & " - " & Trim(RS!sub_nombre)
            RS.MoveNext: ii = ii + 1
        Loop
    End If
    RS.Close: Set RS = Nothing
    .Visible = True
    Gl_Ac_Botones Me, 1, IIf(.MaxRows > 0, 4, 2), modo
End With
Bar1(0).Visible = False
fg_descarga
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Dim RS As New ADODB.Recordset
Select Case Index
Case 0
    If Val(fpLongInteger1(0).Value) < 1 Then fpayuda(0).Caption = "": vaSpread1.MaxRows = 0: Gl_Ac_Botones Me, 1, 3, modo: Exit Sub
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      Set RS = vg_db.Execute("SELECT * FROM a_subsegmento WHERE sub_codigo=" & Val(fpLongInteger1(0).Value) & " and sub_indppr='" & vg_Indppr & "'")
    Else
      Set RS = vg_db.Execute("SELECT * FROM a_subsegmento WHERE sub_codigo=" & Val(fpLongInteger1(0).Value) & "")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing:: fpayuda(0).Caption = "": vaSpread1.MaxRows = 0: Gl_Ac_Botones Me, 1, 3, modo: Exit Sub
    fpayuda(0).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing
    MoverDatosGrillas
End Select
End Sub

Private Sub Image1_Click(Index As Integer)
Dim RS As New ADODB.Recordset
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
    vaSpread1.SetFocus
End Select
End Sub

Private Sub TextCai1_Change(Index As Integer)
Dim i As Long
Dim indactivo As Integer
Dim nom As String
Select Case Index
Case 2
    vaSpread1.Visible = False
    If Trim(TextCai1(Index).text) <> "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = Index + 1: nom = UCase(Trim(vaSpread1.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCai1(Index).text) & "*"
           vaSpread1.Col = 1
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           Else
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           End If
        Next i
        vaSpread1.SetActiveCell Index + 1, 1
    End If
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Trim(TextCai1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(TextCai1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
    If Trim(TextCai1(Index).text) = "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Next
       vaSpread1.SetActiveCell Index + 1, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextCai1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index + 1, 1
    End If
    vaSpread1.Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codsse As Long, codReg As Long, codlpr As Long, i As Long
On Error GoTo Man_Error
Select Case Button.Index
Case 3 '-------> Modificar
    modo = "M"
    fpLongInteger1(0).Enabled = False
    Image1(0).Enabled = False
    Gl_Ac_Botones Me, 1, 0, modo
Case 5 '-------> Eliminar
    If vaSpread1.MaxRows < 1 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 6
    If Trim(vaSpread1.text) = "" Then Exit Sub
    If MsgBox("Eliminar Dato", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codsse = vaSpread1.text
    vaSpread1.Col = 4: codReg = vaSpread1.text
    vaSpread1.Col = 6: codlpr = vaSpread1.text
    vg_db.Execute "DELETE FROM b_asolistaprecio WHERE alp_codsse=" & codsse & " AND alp_codreg=" & codReg & " AND alp_codlpr=" & codlpr & ""
    vaSpread1.Col = 6: vaSpread1.TypeComboBoxCurSel = -1: vaSpread1.text = ""
    vaSpread1.Col = 5: vaSpread1.TypeComboBoxCurSel = -1: vaSpread1.text = ""
Case 7, 10 '-------> Actualizar lista y cancelar
    MoverDatosGrillas
    fpLongInteger1(0).Enabled = True
    Image1(0).Enabled = True
Case 12 '------> Confirmar
    fg_carga ""
    Bar1(0).Visible = True: Bar1(0).Value = 0
    For i = 1 To vaSpread1.MaxRows
        Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
        vaSpread1.Row = i
        vaSpread1.Col = 6
        If Trim(vaSpread1.text) <> "" Then
           vaSpread1.Col = 1: codsse = vaSpread1.text
           vaSpread1.Col = 4: codReg = vaSpread1.text
           vaSpread1.Col = 6: codlpr = vaSpread1.text
           '-------> Borrar registro
           vg_db.Execute "DELETE FROM b_asolistaprecio WHERE alp_codsse=" & codsse & " AND alp_codreg=" & codReg & ""
           '------>  Agregar registro
           vg_db.Execute "INSERT INTO b_asolistaprecio (alp_codsse, alp_codreg, alp_codlpr) VALUES (" & codsse & ", " & codReg & ", " & codlpr & ")"
        End If
    Next i
    Bar1(0).Visible = False: Bar1(0).Value = 0
    modo = ""
    Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows > 0, 4, 2), modo
    fpLongInteger1(0).Enabled = True
    Image1(0).Enabled = True
    fg_descarga
    MsgBox "Generación grabado finalizado sin problema", vbInformation + vbOKOnly, Msgtitulo
Case 15 '-------> Imprimir
    I_AsociarListaPrecio
Case 18 '-------> Salir
   Me.Hide
   Unload Me
End Select
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = NewRow
vaSpread1.Col = 7
Frame1.Caption = vaSpread1.text
End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Select Case Col
Case 5
    Dim indice As Long
    vaSpread1.Row = Row
    vaSpread1.Col = 5: indice = vaSpread1.TypeComboBoxCurSel
    vaSpread1.Col = 6: vaSpread1.TypeComboBoxCurSel = indice
    If modo = "" Then modo = "M"
    If Toolbar1.Buttons(12).Visible = False Then
       fpLongInteger1(0).Enabled = False
       Image1(0).Enabled = False
       Gl_Ac_Botones Me, 1, 0, modo
    End If
End Select
End Sub

