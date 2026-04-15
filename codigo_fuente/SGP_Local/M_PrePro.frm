VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_PrePro 
   Caption         =   "Presupuesto y Proyección"
   ClientHeight    =   4500
   ClientLeft      =   2310
   ClientTop       =   2490
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   11475
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2295
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   11205
         _Version        =   393216
         _ExtentX        =   19764
         _ExtentY        =   4048
         _StockProps     =   64
         ColsFrozen      =   2
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
         MaxCols         =   14
         MaxRows         =   7
         ScrollBars      =   1
         SpreadDesigner  =   "M_PrePro.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   1305
      TabIndex        =   0
      Top             =   435
      Width           =   7965
      Begin VB.OptionButton Option1 
         Caption         =   "Proyección"
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
         Index           =   1
         Left            =   6360
         TabIndex        =   7
         Top             =   675
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Presupuesto"
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
         Index           =   0
         Left            =   3315
         TabIndex        =   6
         Top             =   675
         Value           =   -1  'True
         Width           =   1815
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1470
         TabIndex        =   1
         Top             =   570
         Width           =   705
         _Version        =   196608
         _ExtentX        =   1244
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
         Text            =   "2024"
         DateCalcMethod  =   4
         DateTimeFormat  =   5
         UserDefinedFormat=   "yyyy"
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   1470
         TabIndex        =   8
         Top             =   225
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
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
         NoSpecialKeys   =   0
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2865
         Picture         =   "M_PrePro.frx":1214
         Top             =   120
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3315
         TabIndex        =   10
         Top             =   225
         Width           =   4335
      End
      Begin VB.Label Label3 
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
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Top             =   315
         Width           =   735
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
         Left            =   315
         TabIndex        =   2
         Top             =   675
         Width           =   540
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3360
         TabIndex        =   11
         Top             =   270
         Width           =   4335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
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
      Left            =   9690
      Top             =   885
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   9360
      Top             =   885
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "M_PrePro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim RS As New ADODB.Recordset
Dim est As Boolean
Dim MsgTitulo As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 5010
Me.Width = 11790
MsgTitulo = "Presupuesto y Proyección"
fg_centra Me
fpDateTime1.text = Format(Date, "yyyy")
modo = "": est = False
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = False
fpText1(1).Enabled = ModCasino
Image1(1).Enabled = ModCasino
fpText1(1).text = MuestraCasino(1)
fpayuda(1).Caption = MuestraCasino(2)
MoverDatosGrilla
End Sub

Sub MoverDatosGrilla()
Dim i As Long, anomes As Long, sql1 As String
On Error GoTo Man_Error
With vaSpread1
    .Visible = False
    .Col = -1: .Row = -1
    .BackColor = Shape1(1).FillColor
    .Lock = False
    .Row = -1
    .Col = 1: .BackColor = Shape1(2).FillColor: .Col = 2: .BackColor = Shape1(2).FillColor
    For i = 1 To .MaxRows
        .Row = i
        For j = 3 To .MaxCols
            .Col = j: .text = ""
        Next j
    Next i
    sql1 = IIf(vg_tipbase = "1", " val(mid(ppr_anomes,1,4)) ", " convert(int,substring(convert(varchar(8),ppr_anomes),1,4)) ")
    RS.Open "SELECT * FROM b_presupuestoproyeccion WHERE ppr_cencos = '" & Trim(LimpiaDato(fpText1(1).text)) & "' AND " & sql1 & " = " & Val(Format(fpDateTime1.Value, "yyyy")) & " AND ppr_tipo = '" & IIf(Option1(0).Value = True, "1", "2") & "' ORDER BY ppr_anomes, ppr_codigo", vg_db, adOpenStatic
    est = False
    If Not RS.EOF Then
       Do While Not RS.EOF
          i = IIf(RS!ppr_codigo = 10, 1, IIf(RS!ppr_codigo = 20, 2, IIf(RS!ppr_codigo = 30, 3, IIf(RS!ppr_codigo = 40, 4, IIf(RS!ppr_codigo = 50, 5, IIf(RS!ppr_codigo = 60, 6, 7))))))
          .Row = i
          .Col = (Val(Mid(RS!ppr_anomes, 5, 2)) + 2): .text = IIf(RS!ppr_valor = 0, "", RS!ppr_valor)
          est = True
          RS.MoveNext
       Loop
    End If
    RS.Close: Set RS = Nothing
    .SetActiveCell 3, 1
    .Visible = True
End With
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub fpDateTime1_Change()
MoverDatosGrilla
End Sub

Private Sub fpText1_Change(Index As Integer)
RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & fpText1(1).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
fpayuda(1).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
MoverDatosGrilla
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText1(1).text = vg_codigo
    fpayuda(1).Caption = vg_nombre
    fpDateTime1.SetFocus
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(Index) = True Then MoverDatosGrilla
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 3
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
Case 5
    If Not est < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    If vg_tipbase = "1" Then
       vg_db.Execute "DELETE b_presupuestoproyeccion FROM b_presupuestoproyeccion WHERE ppr_cencos = '" & Trim(LimpiaDato(fpText1(1).text)) & "' AND val(mid(ppr_anomes,1,4)) = " & Format(fpDateTime1.Value, "yyyy") & " AND ppr_tipo = '" & IIf(Option1(0).Value = True, 1, 2) & "'"
    Else
       vg_db.Execute "DELETE b_presupuestoproyeccion FROM b_presupuestoproyeccion WHERE ppr_cencos = '" & Trim(LimpiaDato(fpText1(1).text)) & "' AND convert(int,substring(convert(varchar(8),ppr_anomes),1,4)) = " & Format(fpDateTime1.Value, "yyyy") & " AND ppr_tipo = '" & IIf(Option1(0).Value = True, 1, 2) & "'"
    End If
    vg_db.CommitTrans
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        For j = 3 To vaSpread1.MaxCols
            vaSpread1.Col = j: vaSpread1.text = ""
        Next j
    Next i
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
Case 7
    MoverDatosGrilla
Case 10
    If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    MoverDatosGrilla
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Frame1.Enabled = True
Case 12
    Dim codigo As Long, descripcion As String, valor As Double, inddia As Long
    vg_db.BeginTrans
    With vaSpread1
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1: codigo = .text
            .Col = 2: descripcion = .text
            inddia = 1
            For j = 3 To .MaxCols
                .Col = j
                If Trim(.text) <> "" Then
                   valor = .text
                   If vg_tipbase = "1" Then
                      RS.Open "SELECT DISTINCT ppr_cencos FROM b_presupuestoproyeccion WHERE ppr_cencos = '" & Trim(LimpiaDato(fpText1(1).text)) & "' AND ppr_anomes = " & Val(Format(fpDateTime1.Value, "yyyy")) & fg_pone_cero(inddia, 2) & " AND ppr_tipo = '" & IIf(Option1(0).Value = True, "1", "2") & "' AND ppr_codigo = " & codigo & "", vg_db, adOpenStatic
                   Else
                      RS.Open "SELECT DISTINCT ppr_cencos FROM b_presupuestoproyeccion WHERE ppr_cencos = '" & Trim(LimpiaDato(fpText1(1).text)) & "' AND ppr_anomes = " & Val(Format(fpDateTime1.Value, "yyyy")) & fg_pone_cero(inddia, 2) & " AND ppr_tipo = '" & IIf(Option1(0).Value = True, "1", "2") & "' AND ppr_codigo = " & codigo & "", vg_db, adOpenStatic
                   End If
                   If RS.EOF Then
                      If vg_tipbase = "1" Then
                         vg_db.Execute "INSERT INTO b_presupuestoproyeccion VALUES ('" & Trim(LimpiaDato(fpText1(1).text)) & "', " & Val(Format(fpDateTime1.Value, "yyyy")) & fg_pone_cero(inddia, 2) & ", '" & IIf(Option1(0).Value = True, "1", "2") & "', " & codigo & ", '" & descripcion & "', " & valor & ")"
                      Else
                         vg_db.Execute "INSERT INTO b_presupuestoproyeccion VALUES ('" & Trim(LimpiaDato(fpText1(1).text)) & "', " & Val(Format(fpDateTime1.Value, "yyyy")) & fg_pone_cero(inddia, 2) & ", '" & IIf(Option1(0).Value = True, "1", "2") & "', " & codigo & ", '" & descripcion & "', " & valor & ")"
                      End If
                   Else
                      vg_db.Execute "UPDATE b_presupuestoproyeccion SET ppr_valor = " & valor & " WHERE ppr_cencos = '" & Trim(LimpiaDato(fpText1(1).text)) & "' AND ppr_anomes = " & Val(Format(fpDateTime1.Value, "yyyy")) & fg_pone_cero(inddia, 2) & " AND ppr_tipo = '" & IIf(Option1(0).Value = True, "1", "2") & "' AND ppr_codigo = " & codigo & ""
                   End If
                   RS.Close: Set RS = Nothing
                   If .text = 0 Then .text = ""
                End If
                inddia = inddia + 1
            Next j
        Next i
    End With
    vg_db.CommitTrans
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Frame1.Enabled = True
Case 15
    If vg_tipbase = "1" Then
       RS.Open "SELECT DISTINCT ppr_cencos FROM b_presupuestoproyeccion WHERE ppr_cencos = '" & Trim(LimpiaDato(fpText1(1).text)) & "' AND mid(ppr_anomes,1,4) = " & Format(fpDateTime1.Value, "yyyy") & "", vg_db, adOpenStatic
    Else
       RS.Open "SELECT DISTINCT ppr_cencos FROM b_presupuestoproyeccion WHERE ppr_cencos = '" & Trim(LimpiaDato(fpText1(1).text)) & "' AND convert(int,substring(convert(varchar(8),ppr_anomes),1,4)) = " & Format(fpDateTime1.Value, "yyyy") & "", vg_db, adOpenStatic
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    I_PresupuestoProyeccion Trim(LimpiaDato(fpText1(1).text)), Val(Format(fpDateTime1.Value, "yyyy"))
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Or 2147217900 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
With vaSpread1
    If .MaxRows < 1 Then Exit Sub
    Dim totgas As Double, totuti As Double
    totgas = 0: totuti = 0
    .Col = Col
    
    For i = 2 To 5
        .Row = i
        If Val(.text) > 0 Then totgas = totgas + .text
    Next
    .Row = 6
    .Col = Col: .text = Format(totgas, fg_Pict(9, 2))
    
    .Row = 1
    If Val(.text) > 0 Then totuti = totuti + .text
    
    .Row = 7
    .Col = Col: .text = Format((totuti - totgas), fg_Pict(9, 2))
End With
If modo = "" Then modo = "M"
If Toolbar1.Buttons(3).Visible = True And modo = "M" Then Gl_Ac_Botones Me, 1, 0, modo: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Frame1.Enabled = False
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
'If vaSpread1.MaxRows < 1 Then Exit Sub
'If modo = "" Then modo = "M"
'If ChangeMade = True And modo = "M" Then Gl_Ac_Botones Me, 1, 0, modo: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Frame1.Enabled = False
End Sub
