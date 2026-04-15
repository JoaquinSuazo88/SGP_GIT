VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_CalDia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario Días Feriados"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   8895
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   12255
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   600
         TabIndex        =   7
         Top             =   8280
         Width           =   945
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   8
            Top             =   135
            Width           =   840
         End
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   1560
         TabIndex        =   5
         Top             =   8280
         Width           =   3765
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   6
            Top             =   135
            Width           =   3660
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   7455
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   12015
         _Version        =   393216
         _ExtentX        =   21193
         _ExtentY        =   13150
         _StockProps     =   64
         ColsFrozen      =   5
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
         RowsFrozen      =   1
         SpreadDesigner  =   "M_CalDia.frx":0000
         VisibleCols     =   5
         VisibleRows     =   1
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   3
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
         Text            =   "09/2009"
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
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   600
         TabIndex        =   9
         Top             =   8400
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
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
         Caption         =   "Periodo"
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
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_CalDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim Msgtitulo As String
Dim Est As Boolean, modo As String

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 10005
Me.Width = 12660
Est = True
Msgtitulo = "Calendario Días no Trabajados"
fg_centra Me
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 10, modo
fpDateTime1(0).text = Format(Date, "mm/yyyy")
ArmarCalendario
Est = False
End Sub

Private Sub ArmarCalendario()
Dim i As Long, cencos As String
fg_carga ""
Toolbar1.Enabled = False
Text1(1).text = "": Text1(2).text = ""
vaSpread1.Visible = False
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = Shape1(0).FillColor
vaSpread1.MaxCols = 6
vaSpread1.MaxRows = 0
For i = 1 To Val(Mid(dEoM("01/" & fpDateTime1(0).text), 1, 2))
    vaSpread1.MaxCols = vaSpread1.MaxCols + 1
    vaSpread1.Row = 0
    vaSpread1.Col = vaSpread1.MaxCols
'    vaSpread1.CellType = CellTypeCheckBox
    vaSpread1.ColWidth(i + 6) = 4
    vaSpread1.text = fg_Fecha_Dia(Format(fg_pone_cero(i, 2) & "/" & fpDateTime1(0).text, "yyyymmdd"), 1) '& VgLinea & i
Next i
'-------> cargar datos
cencos = ""
vaSpread1.MaxRows = vaSpread1.MaxRows + 1
vaSpread1.Row = vaSpread1.MaxRows
vaSpread1.Col = 1: vaSpread1.BackColor = &HE0E0E0: vaSpread1.text = ""
vaSpread1.Col = 2: vaSpread1.BackColor = &HE0E0E0: vaSpread1.text = ""
vaSpread1.Col = 3: vaSpread1.BackColor = &HE0E0E0: vaSpread1.text = "0"
vaSpread1.Col = 4: vaSpread1.BackColor = &HE0E0E0: vaSpread1.text = "0"
vaSpread1.Col = 5: vaSpread1.BackColor = &HE0E0E0: vaSpread1.text = "0"
vaSpread1.Col = 6: vaSpread1.BackColor = &HE0E0E0: vaSpread1.text = "0"
For i = 1 To Val(Mid(dEoM("01/" & fpDateTime1(0).text), 1, 2))
    vaSpread1.Col = 6 + i
    vaSpread1.CellType = CellTypeCheckBox
    vaSpread1.TypeHAlign = TypeHAlignCenter
    vaSpread1.text = "0"
Next i

Set RS = vg_db.Execute("sgpadm_s_diasferiados 5, '', '', '" & Format(fpDateTime1(0).text, "mm") & "', '" & Format(fpDateTime1(0).text, "yyyy") & "'")
Do While Not RS.EOF
   If Trim(RS!cli_codigo) <> Trim(cencos) Then
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 1: vaSpread1.BackColor = &HE0E0E0: vaSpread1.text = IIf(IsNull(RS!cli_codigo), "", RS!cli_codigo)
      vaSpread1.Col = 2: vaSpread1.BackColor = &HE0E0E0: vaSpread1.text = IIf(IsNull(RS!cli_nombre), "", RS!cli_nombre)
      vaSpread1.Col = 3: vaSpread1.BackColor = &HE0E0E0: vaSpread1.text = "0"
      vaSpread1.Col = 4: vaSpread1.BackColor = &HE0E0E0: vaSpread1.text = "0"
      vaSpread1.Col = 5: vaSpread1.BackColor = &HE0E0E0: vaSpread1.text = "0"
      vaSpread1.Col = 6: vaSpread1.BackColor = &HE0E0E0: vaSpread1.text = "0"
      cencos = Trim(RS!cli_codigo)
      For i = 1 To Val(Mid(dEoM("01/" & fpDateTime1(0).text), 1, 2))
          vaSpread1.Col = 6 + i
          vaSpread1.CellType = CellTypeCheckBox
          vaSpread1.TypeHAlign = TypeHAlignCenter
          vaSpread1.text = "0"
          If Mid(fg_Fecha_Dia(Format(fg_pone_cero(i, 2) & "/" & fpDateTime1(0).text, "yyyymmdd"), 1), 1, 4) = "Sáb." _
             Or Mid(fg_Fecha_Dia(Format(fg_pone_cero(i, 2) & "/" & fpDateTime1(0).text, "yyyymmdd"), 1), 1, 4) = "Dom." Then
                vaSpread1.BackColor = &HFF00&    '&HFF&
          End If
      Next i
   End If
   If Not IsNull(RS!CFI_Fecha) Then
      vaSpread1.Col = IIf(IsNull(RS!CFI_Fecha), 7, (6 + Val(Mid(RS!CFI_Fecha, 1, 2))))
      vaSpread1.text = "1"
   End If
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
Gl_Ac_Botones Me, 1, 10, modo
Toolbar1.Enabled = True
fg_descarga
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
Est = True
ArmarCalendario
Est = False
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 1, 2
    vaSpread1.Visible = False
    If Trim(Text1(Index).text) <> "" Then
       For i = 2 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread1.Col = 1
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           Else
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           End If
        Next i
        vaSpread1.SetActiveCell Index, 1
    End If
'    vaSpread1_Click Index, 0
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
    If Trim(Text1(Index).text) = "" Then
       For i = 2 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Next
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1.Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long, j As Long, cencos As String, ano As String, mes As String
Select Case Button.Index
Case 3 'Modificar
    modo = "M"
   Gl_Ac_Botones Me, 1, 0, modo
   fpDateTime1(0).Enabled = False
   vaSpread1.Refresh
Case 7, 10 'Actualizar lista ó cancelar
   Est = True
   ArmarCalendario
   Est = False
Case 12 'Grabar datos
   fg_carga ""
   Frame3.Visible = False: Frame4.Visible = False
   Text1(1).Visible = False: Text1(2).Visible = False
   Toolbar1.Enabled = False: fpDateTime1(0).Enabled = False
   mes = Format(fpDateTime1(0).text, "mm")
   ano = Format(fpDateTime1(0).text, "yyyy")
   For i = 2 To vaSpread1.MaxRows
       DoEvents
       vaSpread1.Row = i
       vaSpread1.Col = 1: cencos = vaSpread1.text
       vaSpread1.Col = 3
       If vaSpread1.text = "1" Then
          Bar1(0).Visible = True: Bar1(0).Value = 0: j = 7
          vaSpread1.SetActiveCell 2, vaSpread1.Row
          vg_db.Execute "DELETE Cas_b_Fecha_Inhabiles WHERE CFI_CeCo = '" & cencos & "' AND month(CFI_Fecha) = '" & mes & "' AND YEAR(CFI_Fecha) = '" & ano & "'"
          For j = 7 To vaSpread1.MaxCols
              vaSpread1.Col = j
              Bar1(0).Value = Val(((j - 6) / (vaSpread1.MaxCols - 6)) * 100)
              If vaSpread1.text = "1" Then
                 vg_db.Execute "INSERT INTO Cas_b_Fecha_Inhabiles (CFI_CeCo, CFI_Fecha, CFI_Glosa) VALUES ('" & cencos & "', '" & fg_pone_cero(mes, 2) & "/" & fg_pone_cero((j - 6), 2) & "/" & fg_pone_cero(ano, 4) & "', null)"
              End If
          Next j
       End If
       Bar1(0).Visible = False: Bar1(0).Value = 0
   Next i
   fg_descarga
   MsgBox "Generación grabado Finalizado Sin Problema", vbInformation + vbOKOnly, Msgtitulo
   Frame3.Visible = True: Frame4.Visible = True
   Text1(1).Visible = True: Text1(2).Visible = True
   Toolbar1.Enabled = True: fpDateTime1(0).Enabled = True
   vaSpread1.SetActiveCell 1, 1
   Gl_Ac_Botones Me, 1, 10, modo
Case 18 'Salir
   Me.Hide
   Unload Me
End Select
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim i As Long
If Est Then Exit Sub
If Row = 1 Then
    vaSpread1.Col = Col
    vaSpread1.Row = -1
    Est = True
    vaSpread1.text = IIf(ButtonDown = 1, "1", "0")
    Est = False
End If
If Col = 4 Then
    vaSpread1.Row = Row
    For i = 7 To vaSpread1.MaxCols
        vaSpread1.Col = i
        Est = True
        vaSpread1.Row = 0
        If Mid(vaSpread1.text, 1, 4) = "Sáb." Then vaSpread1.Row = Row: vaSpread1.text = IIf(ButtonDown = 1, "1", "0")
        Est = False
    Next i
End If
If Col = 5 Then
    vaSpread1.Row = Row
    For i = 7 To vaSpread1.MaxCols
        vaSpread1.Col = i
        Est = True
        vaSpread1.Row = 0
        If Mid(vaSpread1.text, 1, 4) = "Dom." Then vaSpread1.Row = Row: vaSpread1.text = IIf(ButtonDown = 1, "1", "0")
        Est = False
    Next i
End If
If Col = 6 Then
    vaSpread1.Row = Row
    For i = 7 To vaSpread1.MaxCols
        vaSpread1.Col = i
        Est = True
        vaSpread1.Row = 0
        If Mid(vaSpread1.text, 1, 4) = "Sáb." Or Mid(vaSpread1.text, 1, 4) = "Dom." Then vaSpread1.Row = Row: vaSpread1.text = IIf(ButtonDown = 1, "1", "0")
        Est = False
    Next i
End If
If Row = 1 And (Col = 4 Or Col = 5 Or Col = 6) Then
    For i = 7 To vaSpread1.MaxCols
        vaSpread1.Col = i
        vaSpread1.Row = -1
        Est = True
        vaSpread1.Row = 0
        If Col = 6 And (Mid(vaSpread1.text, 1, 4) = "Sáb." Or Mid(vaSpread1.text, 1, 4) = "Dom.") Then
           vaSpread1.Row = -1: vaSpread1.text = IIf(ButtonDown = 1, "1", "0")
        ElseIf Col = 4 And (Mid(vaSpread1.text, 1, 4) = "Sáb.") Then
           vaSpread1.Row = -1: vaSpread1.text = IIf(ButtonDown = 1, "1", "0")
        ElseIf Col = 5 And (Mid(vaSpread1.text, 1, 4) = "Dom.") Then
           vaSpread1.Row = -1: vaSpread1.text = IIf(ButtonDown = 1, "1", "0")
        End If
        Est = False
    Next i
End If
modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
fpDateTime1(0).Enabled = False
vaSpread1.Refresh
End Sub
