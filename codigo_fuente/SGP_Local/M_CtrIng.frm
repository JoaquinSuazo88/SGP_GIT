VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_CtrIng 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Ingesta"
   ClientHeight    =   8835
   ClientLeft      =   2235
   ClientTop       =   2445
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   12525
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   12255
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5775
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   11775
         _Version        =   393216
         _ExtentX        =   20770
         _ExtentY        =   10186
         _StockProps     =   64
         ColsFrozen      =   9
         EditEnterAction =   5
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
         MaxCols         =   9
         MaxRows         =   12
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_CtrIng.frx":0000
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celda Habilitada"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   240
         Top             =   270
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   12255
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   5880
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
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
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   1
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   4
         Left            =   360
         TabIndex        =   0
         Top             =   480
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   4
         Left            =   10560
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pedido :"
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
         Index           =   2
         Left            =   10560
         TabIndex        =   21
         Top             =   840
         Width           =   990
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   8400
         TabIndex        =   19
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   2
         Left            =   4440
         TabIndex        =   17
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Cama :"
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
         Index           =   15
         Left            =   8400
         TabIndex        =   12
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rut :"
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
         Left            =   5880
         TabIndex        =   11
         Top             =   240
         Width           =   435
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   7560
         Picture         =   "M_CtrIng.frx":0701
         Top             =   360
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   7980
         TabIndex        =   10
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Paciente :"
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
         Left            =   4440
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contrato :"
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
         Index           =   10
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   1200
         Picture         =   "M_CtrIng.frx":0A0B
         Top             =   390
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   5
         Left            =   1620
         TabIndex        =   7
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Régimen :"
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
         Index           =   5
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   870
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   1665
         TabIndex        =   14
         Top             =   525
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Top             =   1125
         Width           =   3735
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   8010
         TabIndex        =   13
         Top             =   525
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   8400
         TabIndex        =   20
         Top             =   1125
         Width           =   1935
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   4440
         TabIndex        =   18
         Top             =   1125
         Width           =   3735
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   10560
         TabIndex        =   23
         Top             =   1125
         Width           =   1575
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_CtrIng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim modo As String, indnut As Long
Dim can1 As Double, can2 As Double
Dim vecapo() As Variant
Dim est As Boolean

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 9150
Me.Width = 12615
Msgtitulo = "Toma Pedido Paciente"
fg_centra Me
modo = "": est = True
Gl_Mo_Botones Me, 13
Gl_Ac_Botones Me, 13, 3, modo '2
est = True
fpText(4).Enabled = ModCasino
Image1(4).Enabled = ModCasino
fpText(4).text = MuestraCasino(1)
fpayuda(5).Caption = MuestraCasino(2)
'------- Mover concepto aporte nutricionales grilla receta
indnut = 0
RS.Open "SELECT COUNT(*) as nreg FROM a_nutriente", vg_db, adOpenStatic
If RS.EOF Or RS!nreg < 1 Or IsNull(RS!nreg) Then RS.Close: Set RS = Nothing: MsgBox "No existen nutrientes, proceso cancelado...", vbCritical, Msgtitulo
ReDim Preserve vecapo(RS!nreg, 3)
vaSpread1.MaxCols = 9 + RS!nreg: indnut = 10
vaSpread1.MaxRows = 0
RS.Close: Set RS = Nothing
RS.Open "SELECT * FROM a_nutriente ORDER BY nut_secnro", vg_db, adOpenStatic
Do While Not RS.EOF
   vaSpread1.Row = 0
   vaSpread1.Col = indnut
   vaSpread1.text = RS!nut_codigo & "-" & Trim(RS!nut_nombre)
   vecapo(indnut - 9, 1) = RS!nut_codigo
   vecapo(indnut - 9, 2) = indnut
   vecapo(indnut - 9, 3) = 0
   RS.MoveNext: indnut = indnut + 1
Loop
RS.Close: Set RS = Nothing
est = False
End Sub

Private Sub fpText_GotFocus(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    If Trim(fpText(0).text) = "" Or vg_Dig = "N" Then Exit Sub
    fpText(0).text = fg_DespintaRut(fpText(0).text)
    fpText(0).text = Mid(fpText(0).text, 1, Len(Trim(fpText(0).text)) - 1)
End Select
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_LostFocus(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    If fpText(Index).text = "" Or est Then Exit Sub
    fpText(Index).text = fg_RutDig(Trim(fpText(0).text))
    RS.Open "SELECT b_pacientes.*, a_grupopaciente.grp_nombre, a_regimen.reg_nombre " & _
            "FROM (a_grupopaciente RIGHT JOIN b_pacientes ON a_grupopaciente.grp_codigo = b_pacientes.pac_codgrp) " & _
            "LEFT JOIN a_regimen ON b_pacientes.pac_codreg = a_regimen.reg_codigo WHERE b_pacientes.pac_codigo = '" & Trim(fpText(Index).text) & "'", vg_db, adOpenStatic
    Limpia
    If Not RS.EOF Then
       fpText(0).text = fg_PintaRut(fpText(0).text)
       fpayuda(0).Caption = Trim(RS!pac_nombre) & " " & Trim(RS!pac_appaterno) & " " & Trim(RS!pac_apmaterno)
       fpayuda(1).Caption = IIf(IsNull(RS!reg_nombre), "", RS!reg_nombre)
       fpayuda(2).Caption = IIf(IsNull(RS!grp_nombre), "", RS!grp_nombre)
       fpayuda(3).Caption = IIf(IsNull(RS!pac_nrocam), "", RS!pac_nrocam)
       modo = "M": Gl_Ac_Botones Me, 13, 9, modo
       Toolbar1.Buttons(15).Enabled = True: Toolbar1.Buttons(15).ToolTipText = "Historico Toma Pedido Pacientes "
    Else
        Limpia
        RS.Close: Set RS = Nothing: MsgBox "Pacientes no existe...", vbCritical, Msgtitulo
        fpText(0).text = "": fpayuda(0).Caption = ""
        Toolbar1.Buttons(15).Enabled = False: Toolbar1.Buttons(15).ToolTipText = ""
        Exit Sub
    End If
    RS.Close: Set RS = Nothing
End Select
End Sub

Sub Limpia()
For i = 0 To 4
    fpayuda(i).Caption = ""
Next i
vaSpread1.MaxRows = 0
End Sub

Private Sub Image1_Click(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_pacientes", "pac_", "Pacientes", "Pac"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(Index).text = fg_PintaRut(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    RS.Open "SELECT b_pacientes.*, a_grupopaciente.grp_nombre, a_regimen.reg_nombre " & _
            "FROM (a_grupopaciente RIGHT JOIN b_pacientes ON a_grupopaciente.grp_codigo = b_pacientes.pac_codgrp) " & _
            "LEFT JOIN a_regimen ON b_pacientes.pac_codreg = a_regimen.reg_codigo WHERE b_pacientes.pac_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic
    Limpia
    If Not RS.EOF Then
       fpayuda(0).Caption = Trim(RS!pac_nombre) & " " & Trim(RS!pac_appaterno) & " " & Trim(RS!pac_apmaterno)
       fpayuda(1).Caption = IIf(IsNull(RS!reg_nombre), "", RS!reg_nombre)
       fpayuda(2).Caption = IIf(IsNull(RS!grp_nombre), "", RS!grp_nombre)
       fpayuda(3).Caption = IIf(IsNull(RS!pac_nrocam), "", RS!pac_nrocam)
       modo = "M": Gl_Ac_Botones Me, 13, 9, modo
       Toolbar1.Buttons(15).Enabled = True: Toolbar1.Buttons(15).ToolTipText = "Historico Toma Pedido Pacientes "
    End If
    RS.Close: Set RS = Nothing
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim nroped As Long
On Error GoTo Man_Error
Select Case Button.Index
Case 1 '------- Agregar
Case 3 '------- Modificar
    modo = "M"
    Gl_Ac_Botones Me, 13, 0, modo
Case 5 '------- Eliminar
Case 7 '------- Actualizar
   TraerTomaPedido Val(fpayuda(4).Caption)
Case 10 '------- Cancelar
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    TraerTomaPedido Val(fpayuda(4).Caption)
    modo = "": Gl_Ac_Botones Me, 13, 3, modo
    Toolbar1.Buttons(15).Enabled = True: Toolbar1.Buttons(15).ToolTipText = "Historico Toma Pedido Pacientes "
Case 12 '------- Grabar
    nroped = Val(fpayuda(4).Caption)
    Dim cannco As Double, cancon As Double, numlin As Long, CodRec As Long, nroite As Long, coding As String
    Dim canser As Double, pctapr As Double, pctcoc As Double, canbru As Double
    vg_db.BeginTrans
    fg_carga ""
    With vaSpread1
         For i = 1 To .MaxRows
             .Row = i
             .Col = 1
             coding = GetItem(.text, 5)
             .Col = 6
             If .text <> "TOTAL APORTES" And coding <> "" Then
                .Col = 1
                numlin = GetItem(.text, 10)
                CodRec = GetItem(.text, 4)
                nroite = GetItem(.text, 6)
                coding = GetItem(.text, 5)
                pctapr = GetItem(.text, 7)
                pctcoc = GetItem(.text, 8)
                .Col = 7: canser = .text
                .Col = 8: cannco = .text
                .Col = 9: cancon = .text
                canbru = (((canser / pctapr) / pctcoc) * 10000)
                vg_db.Execute "UPDATE b_tomapedidodetrec SET tdr_canpro = " & canbru & ",  tdr_cannco = " & cannco & ", tdr_cancon = " & cancon & " WHERE tdr_codigo = " & nroped & " AND tdr_numlin = " & numlin & " AND tdr_codrec = " & CodRec & " AND tdr_nroite = " & nroite & " AND tdr_coding = '" & coding & "'"
             End If
         Next i
    End With
    fg_descarga
    vg_db.CommitTrans
    modo = "M": Gl_Ac_Botones Me, 13, 9, modo
    Toolbar1.Buttons(15).Enabled = True: Toolbar1.Buttons(15).ToolTipText = "Historico Toma Pedido Pacientes "
Case 15 '------- Busqueda paciente
   est = True
   Dim codpac As String
   vg_codigo = "": vg_nombre = ""
   codpac = fg_DespintaRut(fpText(0).text)
   B_PedPac.LlenarPedidoPaciente codpac
   B_PedPac.Show 1
   est = False
   modo = ""
   If vg_codigo = "" Then Exit Sub
   modo = "M"
   TraerTomaPedido Val(vg_codigo)
    modo = "M": Gl_Ac_Botones Me, 13, 9, modo
Case 17 'Impirmir
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_ControlIngesta Trim(fpText(4).text), fg_DespintaRut(fpText(0).text), Val(fpayuda(3).Caption), Val(fpayuda(4).Caption)
Case 20
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
fg_descarga
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Sub TraerTomaPedido(nroped As Long)
If Trim(fpayuda(0).Caption) = "" Then Exit Sub
Dim sql1 As String, sql2 As String, sql3 As String, sql4 As String, auxfec As String, sql5 As String, sql6 As String
Dim auxreg As Long, auxser As Long, auxess As Long, auxrec As Long, indrec As Long, canapo As Double
est = True
fg_carga ""
'------- Encabezado toma pedido paciente
RS.Open "SELECT DISTINCT a.top_codigo, a.top_fecped, a.top_codreg, b.reg_nombre, a.top_codusu, c.usu_nombre " & _
        "FROM b_tomapedido a, a_regimen b, a_usuarios c " & _
        "WHERE a.top_codreg = b.reg_codigo " & _
        "AND   a.top_codusu = c.usu_codigo " & _
        "AND   a.top_codigo = " & nroped & "", vg_db, adOpenStatic
If Not RS.EOF Then fpayuda(4).Caption = RS!top_codigo
RS.Close: Set RS = Nothing
'------- Llernar grilla recetas
sql5 = IIf(vg_tipbase = "1", " ORDER BY f.reg_codigo, g.ser_codigo, g.ser_orden, h.ess_orden ", "")
sql1 = "SELECT DISTINCT a.top_fecped, f.reg_codigo, f.reg_nombre, g.ser_codigo, g.ser_nombre, g.ser_orden, h.ess_codigo, h.ess_nombre, h.ess_orden, " & _
       "c.tdr_codrec, d.rec_nombre, c.tdr_nroite, c.tdr_coding, c.tdr_canpro, c.tdr_cospro, c.tdr_pctapr, c.tdr_pctcoc, c.tdr_pctnut, " & _
       "e.ing_nombre, i.unm_nomcor, b.tpd_codmin, b.tpd_tiprec , b.tpd_cansel, b.tpd_canser, b.tpd_caning, b.tpd_numlin, e.ing_facnut, c.tdr_cannco, c.tdr_cancon " & _
       "FROM b_tomapedido a, b_tomapedidodet b, b_tomapedidodetrec c, b_receta d, b_ingrediente e, a_regimen f, a_servicio g, a_estservicio h, a_unidadmed i " & _
       "WHERE a.top_codigo = b.tpd_codigo " & _
       "AND   b.tpd_codigo = c.tdr_codigo " & _
       "AND   b.tpd_numlin = c.tdr_numlin " & _
       "AND   b.tpd_codrec = c.tdr_codrec " & _
       "AND   b.tpd_codreg = f.reg_codigo " & _
       "AND   b.tpd_codser = g.ser_codigo " & _
       "AND   b.tpd_estser = h.ess_codigo AND h.ess_cencos = '" & MuestraCasino(1) & "' " & _
       "AND   c.tdr_codrec = d.rec_codigo " & _
       "AND   c.tdr_coding = e.ing_codigo " & _
       "AND   e.ing_unimed = i.unm_codigo " & _
       "AND   a.top_codigo = " & nroped & " " & _
       "AND   b.tpd_prorec = 'R' AND c.tdr_canpro > 0 " & _
       "" & sql5 & ""

sql5 = IIf(vg_tipbase = "1", " ORDER BY f.reg_codigo, g.ser_codigo, g.ser_orden, ess_orden ", "")
sql2 = "SELECT  DISTINCT a.top_fecped, f.reg_codigo, f.reg_nombre, g.ser_codigo, g.ser_nombre, g.ser_orden, b.tpd_estser AS ess_codigo, " & _
       "'Adicionales' AS ess_nombre, 99999999 AS ess_orden, c.tdr_codrec, d.rec_nombre, c.tdr_nroite, c.tdr_coding, " & _
       "c.tdr_canpro, c.tdr_cospro, c.tdr_pctapr, c.tdr_pctcoc, c.tdr_pctnut, e.ing_nombre, i.unm_nomcor, b.tpd_codmin, " & _
       "b.tpd_tiprec , b.tpd_cansel, b.tpd_canser, b.tpd_caning, b.tpd_numlin, e.ing_facnut, c.tdr_cannco, c.tdr_cancon " & _
       "FROM b_tomapedido a, b_tomapedidodet b, b_tomapedidodetrec c, b_receta d, b_ingrediente e, a_regimen f, a_servicio g, a_unidadmed i " & _
       "WHERE a.top_codigo = b.tpd_codigo " & _
       "AND   b.tpd_codigo = c.tdr_codigo " & _
       "AND   b.tpd_numlin = c.tdr_numlin " & _
       "AND   b.tpd_codrec = c.tdr_codrec " & _
       "AND   b.tpd_codreg = f.reg_codigo " & _
       "AND   b.tpd_codser = g.ser_codigo " & _
       "AND   c.tdr_codrec = d.rec_codigo " & _
       "AND   c.tdr_coding = e.ing_codigo " & _
       "AND   e.ing_unimed = i.unm_codigo " & _
       "AND   a.top_codigo = " & nroped & " " & _
       "AND   b.tpd_prorec = 'R' AND c.tdr_canpro > 0 AND b.tpd_estser = -99999999 " & _
       "" & sql5 & ""

sql3 = "SELECT DISTINCT a.top_fecped, f.reg_codigo, f.reg_nombre, g.ser_codigo, g.ser_nombre, g.ser_orden, b.tpd_estser AS ess_codigo, " & _
       "'Adicionales' AS ess_nombre, 99999999 AS ess_orden, c.tdr_codrec, d.pro_nombre, c.tdr_nroite, c.tdr_coding, " & _
       "c.tdr_canpro, c.tdr_cospro, c.tdr_pctapr, c.tdr_pctcoc, c.tdr_pctnut, e.ing_nombre, i.unm_nomcor, b.tpd_codmin, " & _
       "b.tpd_tiprec, b.tpd_cansel, b.tpd_canser, b.tpd_caning, b.tpd_numlin, e.ing_facnut, c.tdr_cannco, c.tdr_cancon " & _
       "FROM b_tomapedido a, b_tomapedidodet b, b_tomapedidodetrec c, b_productos d, b_ingrediente e, a_regimen f, a_servicio g, a_unidadmed i " & _
       "WHERE a.top_codigo = b.tpd_codigo " & _
       "AND   b.tpd_codigo = c.tdr_codigo " & _
       "AND   b.tpd_numlin = c.tdr_numlin " & _
       "AND   b.tpd_codrec = c.tdr_codrec " & _
       "AND   b.tpd_codreg = f.reg_codigo " & _
       "AND   b.tpd_codser = g.ser_codigo " & _
       "AND   c.tdr_codrec = val(d.pro_codigo) " & _
       "AND   c.tdr_coding = e.ing_codigo " & _
       "AND   e.ing_unimed = i.unm_codigo " & _
       "AND   a.top_codigo = " & nroped & " " & _
       "AND   b.tpd_prorec = 'P' AND c.tdr_canpro > 0 AND b.tpd_estser = -99999999 " & _
       "ORDER BY f.reg_codigo, g.ser_codigo, g.ser_orden, ess_orden"
       
RS.Open sql1 & " UNION " & sql2 & " UNION " & sql3, vg_db, adOpenStatic
With vaSpread1
    auxreg = 0: auxser = 0: auxess = 0: auxrec = 0: auxfec = "": indrec = 1
    .Row = -1: .Col = -1:
    .BackColor = Shape1(0).FillColor
    .MaxRows = 0
    .Visible = False
    If Not RS.EOF Then
       Do While Not RS.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
'          .Col = 1: .text = RS!reg_codigo & ";" & _
'                            RS!ser_codigo & ";" & _
'                            RS!ess_codigo & ";" & _
'                            RS!tdr_codrec & ";" & _
'                            Trim(RS!tdr_coding) & ";" & _
'                            RS!tdr_nroite & ";" & _
'                            RS!tdr_pctapr & ";" & _
'                            RS!tdr_pctcoc & ";" & _
'                            RS!tdr_pctnut & ";" & _
'                            RS!tpd_numlin & ";" & _
'                            indrec & ";"
          .Col = 2: .text = RS!top_fecped
          If RS!reg_codigo <> auxreg Or RS!ser_codigo <> auxser Then
             .Col = 3: .FontBold = True: .text = Trim(RS!ser_nombre) & " - " & Trim(RS!reg_nombre)
             auxreg = RS!reg_codigo
             auxser = RS!ser_codigo
          End If
          If RS!ess_codigo <> auxess Then
             .Col = 4: .FontBold = True: .text = Trim(RS!ess_nombre)
             auxess = RS!ess_codigo
          End If
          If RS!tdr_codrec <> auxrec Then
             .Col = 5: .FontBold = True: .text = Trim(RS!rec_nombre)
             .MaxRows = .MaxRows + 1
             .Row = .MaxRows
             indrec = .MaxRows
             auxrec = RS!tdr_codrec
          End If
          .Col = 1: .text = RS!reg_codigo & ";" & _
                            RS!ser_codigo & ";" & _
                            RS!ess_codigo & ";" & _
                            RS!tdr_codrec & ";" & _
                            Trim(RS!tdr_coding) & ";" & _
                            RS!tdr_nroite & ";" & _
                            RS!tdr_pctapr & ";" & _
                            RS!tdr_pctcoc & ";" & _
                            RS!tdr_pctnut & ";" & _
                            RS!tpd_numlin & ";" & _
                            indrec & ";"
'          .Col = 1: .text = .text
          .Col = 6: .text = Trim(RS!ing_nombre)
          .Col = 5: .text = Trim(RS!ing_nombre)
          .Col = 7: .Lock = True: .text = Format((((RS!tdr_pctapr / 100) * RS!tdr_canpro) * (RS!tdr_pctcoc / 100)), fg_Pict(6, 2))
          .Col = 8: .ForeColor = &HFF0000: .text = Format(IIf(IsNull(RS!tdr_cannco), 0, RS!tdr_cannco), fg_Pict(6, 2))
          .Col = 9: .ForeColor = &HFF0000: .text = Format(IIf(IsNull(RS!tdr_cancon) Or RS!tdr_cancon = 0, (((RS!tdr_pctapr / 100) * RS!tdr_canpro) * (RS!tdr_pctcoc / 100)), RS!tdr_cancon), fg_Pict(6, 2))
          For i = 10 To .MaxCols
              .Col = i
              .Lock = True: .TypeHAlign = TypeHAlignRight: .text = Format(0, fg_Pict(6, 2))
              vecapo((i - 9), 3) = 0
          Next i
          auxing = RS!tdr_coding
          Do While RS!tdr_codrec = auxrec And RS!tdr_coding = auxing And Not RS.EOF
              RS1.Open "SELECT a.* FROM b_productonut a, a_nutriente b WHERE a.pnu_codapo=b.nut_codigo AND a.pnu_codpro='" & RS!tdr_coding & "' ORDER BY b.nut_secnro", vg_db, adOpenStatic
              If Not RS1.EOF Then
                 Do While Not RS1.EOF
                    For indnut = 1 To UBound(vecapo)
                        If RS1!pnu_codapo = vecapo(indnut, 1) Then
                           .Row = .MaxRows
                           .Col = vecapo(indnut, 2)
                           .TypeHAlign = TypeHAlignRight
                           .text = Format(((RS!tdr_pctnut / 100) * (RS1!pnu_canapo * RS!tdr_canpro) / RS!ing_facnut), fg_Pict(6, 2))
                           vecapo(indnut, 3) = vecapo(indnut, 3) + Format(((RS!tdr_pctnut / 100) * (RS1!pnu_canapo * RS!tdr_canpro) / RS!ing_facnut), fg_Pict(6, 2))
                           Exit For
                        End If
                    Next indnut
                    RS1.MoveNext
                 Loop
              End If
              RS1.Close: Set RS1 = Nothing
             RS.MoveNext
             If RS.EOF Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 5: .Font.Bold = True: .text = "TOTAL APORTES"
                .Col = 6: .Font.Bold = True: .text = "TOTAL APORTES"
                .Col = 7: .Lock = True
                .Col = 8: .Lock = True
                .Col = 9: .Lock = True
                For i = 10 To .MaxCols
                    .Col = i
                    .Lock = True: .TypeHAlign = TypeHAlignRight: .text = Format(0, fg_Pict(6, 2))
                Next i
                '------- Calcular total aportes
                For i = indrec To .MaxRows - 1
                    For j = 10 To .MaxCols
                        .Row = i
                        .Col = j
                        canapo = Val(.text)
                        .Row = .MaxRows
                        .Font.Bold = True: .text = .text + canapo
                    Next j
                Next i
                Exit Do
             ElseIf RS!tdr_codrec <> auxrec Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 5: .Font.Bold = True: .text = "TOTAL APORTES"
                .Col = 6: .Font.Bold = True: .text = "TOTAL APORTES"
                .Col = 7: .Lock = True
                .Col = 8: .Lock = True
                .Col = 9: .Lock = True
                For i = 10 To .MaxCols
                    .Col = i
                    .Lock = True: .TypeHAlign = TypeHAlignRight: .text = Format(0, fg_Pict(6, 2))
                Next i
                '------- Calcular total aportes
                For i = indrec To .MaxRows - 1
                    For j = 10 To .MaxCols
                        .Row = i
                        .Col = j
                        canapo = .text
                        .Row = .MaxRows
                        .Font.Bold = True: .text = .text + canapo
                    Next j
                Next i
             End If
          Loop
       Loop
    End If
    RS.Close: Set RS = Nothing
    .SetActiveCell 8, 1
    .Visible = True
    .SetFocus
End With
est = False
fg_descarga
End Sub


Public Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
Dim cancon As Double, cannco As Double, pctnut As Double, coding As String, indini As Long, canbru As Double, pctapr As Double, pctcoc As Double
Gl_Ac_Botones Me, 13, 0, modo
With vaSpread1
     .Row = Row
     Select Case Col
     Case 8
         .Col = 1: pctnut = GetItem(.text, 9): coding = GetItem(.text, 5): indini = GetItem(.text, 11): pctapr = GetItem(.text, 7): pctcoc = GetItem(.text, 8)
         .Col = 7: cancon = .text
         .Col = 8: cannco = .text
         .Col = 9: .text = Format(cancon - cannco, fg_Pict(6, 2))
        canbru = (((.text / pctapr) / pctcoc) * 10000)
         CalcularAporte coding, Row, indini, canbru, pctnut
     Case 9
         .Col = 1: pctnut = GetItem(.text, 9): coding = GetItem(.text, 5): indini = GetItem(.text, 11): pctapr = GetItem(.text, 7): pctcoc = GetItem(.text, 8)
         .Col = 7: cancon = .text
         .Col = 9: cannco = .text
         .Col = 8: .text = Format(cancon - cannco, fg_Pict(6, 2))
         .Col = 9
        canbru = (((.text / pctapr) / pctcoc) * 10000)
         CalcularAporte coding, Row, indini, canbru, pctnut
     End Select
End With
End Sub

Sub CalcularAporte(coding As String, indrow As Long, indini As Long, canser As Double, pctnut As Double)
Dim canapo As Double, indfin As Long
With vaSpread1
     RS.Open "SELECT a.pnu_codpro, a.pnu_codapo, a.pnu_canapo, c.ing_facnut FROM b_productonut a, a_nutriente b, b_ingrediente c WHERE c.ing_codigo=a.pnu_codpro AND a.pnu_codapo=b.nut_codigo AND c.ing_codigo='" & coding & "' ORDER BY nut_secnro", vg_db, adOpenStatic
     If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
     For j = 10 To .MaxCols
         .Row = indrow
         .Col = j
         .Lock = True: .TypeHAlign = TypeHAlignRight: .text = Format(0, fg_Pict(6, 2))
     Next j
     Do While Not RS.EOF
        For i = 10 To .MaxCols
            .Row = 0
            .Col = i
            If RS!pnu_codapo = Val(GetItem(.text, 1)) Then
               .Row = indrow
               .TypeHAlign = TypeHAlignRight
               .text = Format(((pctnut / 100) * (RS!pnu_canapo * canser) / RS!ing_facnut), fg_Pict(6, 2))
               Exit For
            End If
        Next i
        RS.MoveNext
     Loop
     RS.Close: Set RS = Nothing
     
     For i = indini To .MaxRows
         .Row = i
         .Col = 6
         If .text = "TOTAL APORTES" Then
            indfin = i
            .Col = 7: .Lock = True
            .Col = 8: .Lock = True
            .Col = 9: .Lock = True
            For j = 10 To .MaxCols
                .Col = j
                .Lock = True: .TypeHAlign = TypeHAlignRight: .text = Format(0, fg_Pict(6, 2))
            Next j
            Exit For
         End If
     Next i

     For i = indini To indfin - 1
         For j = 10 To .MaxCols
             .Row = i
             .Col = j
             canapo = .text
             .Row = indfin
             .Font.Bold = True: .text = .text + canapo
         Next j
     Next i

End With
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
With vaSpread1
    If .MaxRows < 1 Then Exit Sub
    .Row = Row
    If Mode = 1 Then
       Dim cannco As Double, cancon As Double, cantid As Double
       Select Case Col
       Case 8 'Cantidad no consumida
            .Col = Col: can2 = .text
            .Col = 7: can1 = .text
       Case 9 'Cantidad consumida
            .Col = Col: can2 = .text
            .Col = 7: can1 = .text
       End Select
    ElseIf Mode = 0 Then
      Select Case Col
      Case 8
          .Col = Col
          If .text > can1 Then
             .Col = Col: .text = can2
          End If
             vaSpread1_EditChange Col, Row
      Case 9
          .Col = Col
          If .text > can1 Or .text = 0 Then
             .Col = Col: .text = can2
          End If
          vaSpread1_EditChange Col, Row
      End Select
    End If
End With
End Sub
