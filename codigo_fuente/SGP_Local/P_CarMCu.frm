VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form P_CarMCu 
   Caption         =   "Cargar MisCuentas"
   ClientHeight    =   7140
   ClientLeft      =   1845
   ClientTop       =   2145
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   1140
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   2250
         TabIndex        =   1
         Top             =   330
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   2250
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   660
         Width           =   5595
         _Version        =   196608
         _ExtentX        =   9869
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
         ButtonStyle     =   2
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   2
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
         OnFocusPosition =   1
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
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3915
         TabIndex        =   5
         Top             =   330
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3480
         Picture         =   "P_CarMCu.frx":0000
         Top             =   240
         Width           =   480
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
         Left            =   480
         TabIndex        =   4
         Top             =   405
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Seleccione Archivo"
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
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   1665
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3960
         TabIndex        =   6
         Top             =   375
         Width           =   3855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7140
      Left            =   11085
      TabIndex        =   7
      Top             =   0
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   12594
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   5790
      Left            =   0
      TabIndex        =   8
      Top             =   1320
      Width           =   11055
      _Version        =   393216
      _ExtentX        =   19500
      _ExtentY        =   10213
      _StockProps     =   64
      ButtonDrawMode  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   12
      MaxRows         =   20
      SpreadDesigner  =   "P_CarMCu.frx":030A
   End
End
Attribute VB_Name = "P_CarMCu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset, RS2 As New ADODB.Recordset
Dim Fecha As Long
Dim Msgtitulo As String
Dim est As Boolean

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
fg_carga ""
Me.Height = 7575
Me.Width = 11700
fg_centra Me
est = True
Msgtitulo = "Cargar MisCuentas"
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.Enabled = False: btnX.ToolTipText = "Confirmar "
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): btnX.Enabled = False: btnX.ToolTipText = "Imprimir "
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
vaSpread1.MaxRows = 0
fpText(0).Enabled = ModCasino
Image1.Enabled = ModCasino
fpText(0).text = MuestraCasino(1)
fpayuda.Caption = MuestraCasino(2)
est = False
fg_descarga
End Sub

Private Sub Form_Resize()
Frame2.Move 1575, 0, 8415, 1095 '575
vaSpread1.Move 0, 1200, ScaleWidth - Toolbar1.ButtonHeight - 200, ScaleHeight - 1200
Toolbar1.Refresh
End Sub

Private Sub fpText_ButtonHit(Index As Integer, Button As Integer, NewIndex As Integer)
Dim oError As Boolean
oError = False
Select Case Index
Case 1
    CD.ShowOpen
    If CD.Filename = "" Then
       fpText(1).text = ""
    Else
       oError = False
       fpText(1).text = Dir(CD.Filename)
       oError = IIf(CarRaciones(Dir(CD.Filename)) = 0, False, True): ': MsgBox CD.Filename
       If Not oError Then MsgBox "Archivo no contiene el centro de costo, proceso cancelado", vbInformation + vbOKOnly, Msgtitulo
    End If
End Select
End Sub

Function CarRaciones(ByVal cdbz As String) As Long
CarRaciones = False
'------- leer regimen
RS1.Open "SELECT * FROM a_regimen", vg_db, adOpenStatic
'------- leer servicio
RS2.Open "SELECT * FROM a_servicio", vg_db, adOpenStatic
Open CD.Filename For Input As #1
DoEvents
vaSpread1.Row = -1: vaSpread1.Col = -1:
vaSpread1.BackColor = &HC0FFFF
vaSpread1.MaxRows = 0
Do While Not EOF(1)
   Line Input #1, cpar
   '------- Validar cencos
   If Trim(Mid(cpar, 130, 10)) = LimpiaDato(Trim(fpText(0).text)) Then
      CarRaciones = True
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 1: vaSpread1.text = Trim(Mid(cpar, 1, 9)) 'Rut Alumno
      vaSpread1.Col = 2: vaSpread1.text = Trim(Mid(cpar, 10, 50)) 'Nombre y Apelido Alumno
      vaSpread1.Col = 3: vaSpread1.text = Trim(Mid(cpar, 60, 20)) 'Curso
      vaSpread1.Col = 4: vaSpread1.text = Trim(Mid(cpar, 80, 50)) 'Nombre Apoderado
      vaSpread1.Col = 5: vaSpread1.text = Trim(Mid(cpar, 140, 20)) 'Servicio
      vaSpread1.Col = 6: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Trim(Mid(cpar, 160, 5)) 'Raciones
      vaSpread1.Col = 7: vaSpread1.TypeHAlign = TypeHAlignCenter:  vaSpread1.text = Trim(Mid(cpar, 165, 2)) & "/" & Trim(Mid(cpar, 167, 2)) & "/" & Trim(Mid(cpar, 169, 4)) 'Trim(Mid(cpar, 165, 8)) 'Fecha
      vaSpread1.Col = 8: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = Format(Val(Mid(cpar, 173, 10)), fg_Pict(6, 2))
      '------- agregar regimen
      lisnom = "": liscod = ""
      Do While Not RS1.EOF
         vaSpread1.Col = 10: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(RS1!reg_nombre)
         vaSpread1.Col = 9: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & RS1!reg_codigo
         vaSpread1.Col = 10: vaSpread1.TypeComboBoxList = lisnom
         vaSpread1.Col = 9: vaSpread1.TypeComboBoxList = liscod
         RS1.MoveNext
      Loop
      RS1.MoveFirst
      '------- agregar servicio
      lisnom = "": liscod = ""
      Do While Not RS2.EOF
         vaSpread1.Col = 12: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(RS2!ser_nombre)
         vaSpread1.Col = 11: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & RS2!ser_codigo
         vaSpread1.Col = 12: vaSpread1.TypeComboBoxList = lisnom
         vaSpread1.Col = 11: vaSpread1.TypeComboBoxList = liscod
         RS2.MoveNext
      Loop
     RS2.MoveFirst
   End If
Loop
Close #1
RS1.Close: Set RS1 = Nothing
RS2.Close: Set RS2 = Nothing
If CarRaciones = True Then Toolbar1.Buttons(1).Enabled = True Else Toolbar1.Buttons(1).Enabled = False
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer
Dim codreg As Long, codser As Long, Fecha As Long
Dim monpag As Double
Dim rutalu As String
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    codreg = 0: codser = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 10
        If vaSpread1.TypeComboBoxCurSel = -1 Then MsgBox "Debe seleccionar regimen", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        vaSpread1.Col = 12
        If vaSpread1.TypeComboBoxCurSel = -1 Then MsgBox "Debe seleccionar servicio", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Next i
    RS.Open "select cli_codigo, cli_nombre from b_clientes where cli_codigo='" & LimpiaDato(Trim(fpText(0).text)) & "' and cli_tipo=0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpText(0).text = "": fpayuda.Caption = "": MsgBox "No existe contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    vg_db.BeginTrans
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1: rutalu = vaSpread1.text
        vaSpread1.Col = 2: nomalu = vaSpread1.text
        vaSpread1.Col = 3: curso = vaSpread1.text
        vaSpread1.Col = 4: nomapo = vaSpread1.text
        vaSpread1.Col = 5: seralu = vaSpread1.text
        vaSpread1.Col = 6: racion = Val(vaSpread1.text)
        vaSpread1.Col = 7: Fecha = Mid(vaSpread1.text, 7, 4) & Mid(vaSpread1.text, 4, 2) & Mid(vaSpread1.text, 1, 2)
        vaSpread1.Col = 8: monpag = vaSpread1.Value
        vaSpread1.Col = 9: codreg = Val(vaSpread1.text)
        vaSpread1.Col = 11: codser = Val(vaSpread1.text)
        '------- Borrar y Grabar alumnos raciones
        vg_db.Execute "DELETE FROM b_pagoalumnos WHERE pal_cencos='" & LimpiaDato(Trim(fpText(0).text)) & "' AND pal_rutalumno='" & rutalu & "' AND pal_fecha=" & Fecha & ""
        vg_db.Execute "INSERT INTO b_pagoalumnos VALUES ('" & LimpiaDato(Trim(fpText(0).text)) & "', '" & rutalu & "', " & Fecha & ", '" & nomalu & "', '" & curso & "', '" & nomapo & "', '" & seralu & "', " & racion & ", " & monpag & ", " & codreg & ", " & codser & ")"
        '------- Grabar alumons en tabla b_clientes
        RS.Open "SELECT * FROM B_CLIENTES WHERE cli_codigo='" & rutalu & "'", vg_db, adOpenStatic
        If RS.EOF Then
           vg_db.Execute "INSERT INTO b_clientes (cli_codigo, cli_nombre, cli_tipo) VALUES ('" & rutalu & "', '" & nomalu & "', 3)"
        End If
        RS.Close: Set RS = Nothing
    Next i
    vg_db.CommitTrans
    fg_descarga
    MsgBox "Proceso Finalizado Sin Problema", vbInformation + vbOKOnly, Msgtitulo
'    Toolbar1.Buttons(3).Enabled = True: Toolbar1.Buttons(5).Enabled = True
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
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim indice As Long
Select Case Col
Case 10
    vaSpread1.Row = Row
    vaSpread1.Col = 10: indice = vaSpread1.TypeComboBoxCurSel
    vaSpread1.Col = 9: vaSpread1.TypeComboBoxCurSel = indice
'    vaSpread1.EditEnterAction = EditEnterActionNone
Case 12
    vaSpread1.Row = Row
    vaSpread1.Col = 12: indice = vaSpread1.TypeComboBoxCurSel
    vaSpread1.Col = 11: vaSpread1.TypeComboBoxCurSel = indice
'    vaSpread1.EditEnterAction = EditEnterActionNone
End Select
End Sub
