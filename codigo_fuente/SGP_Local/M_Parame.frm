VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_Parame 
   Caption         =   "Parametros Generales"
   ClientHeight    =   3150
   ClientLeft      =   2265
   ClientTop       =   3150
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2685
      Left            =   15
      TabIndex        =   2
      Top             =   390
      Width           =   8250
      Begin VB.ComboBox fpList1 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox fpList2 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2115
         Width           =   3090
      End
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   30
         TabIndex        =   7
         Top             =   930
         Width           =   8190
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Permite modificar"
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
         Left            =   135
         TabIndex        =   1
         Top             =   615
         Visible         =   0   'False
         Width           =   2040
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   1980
         TabIndex        =   0
         Top             =   285
         Visible         =   0   'False
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
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   2070
         TabIndex        =   14
         Top             =   2175
         Width           =   3090
      End
      Begin VB.Label Label3 
         Caption         =   "Metodo Actualización Base Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   2
         Left            =   135
         TabIndex        =   12
         Top             =   2070
         Width           =   1905
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dias Stock"
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
         Left            =   135
         TabIndex        =   11
         Top             =   1305
         Width           =   945
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   2010
         TabIndex        =   10
         Top             =   645
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº de días despues de fin de mes"
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
         Left            =   2880
         TabIndex        =   9
         Top             =   1740
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bloquear ingreso"
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
         TabIndex        =   8
         Top             =   1710
         Width           =   1440
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3795
         TabIndex        =   4
         Top             =   285
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3345
         Picture         =   "M_Parame.frx":0000
         Top             =   180
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Casino en operación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   135
         TabIndex        =   3
         Top             =   315
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3840
         TabIndex        =   5
         Top             =   330
         Visible         =   0   'False
         Width           =   4335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_Parame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim est As Boolean

Private Sub Form_Load()

On Error GoTo Man_Error

Me.Height = 3400
Me.Width = 8415
fg_centra Me
Me.HelpContextID = vg_OpcM

'LlenaDatos
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, False): BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = IIf(Mid(ValidarUsuario(Me), 2, 1) = "0", True, False): BtnX.ToolTipText = ""
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Dim ind As Long
est = True
fpList1.Clear
For i = 28 To 1 Step -1
    
    fpList1.AddItem i

Next i

For i = 0 To fpList1.listcount - 1
    
    fpList1.ItemData(i) = Val(fpList1.List(i))

Next i

fpList1.AddItem ""
fpList1.ListIndex = fpList1.listcount - 1
RS1.Open "SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'diasbloq'", vg_db, adOpenStatic
If Not RS1.EOF Then

   fpList1.ListIndex = fpList1.ItemData(Val(IIf(IsNull(RS1!par_valor), 0, RS1!par_valor)))

End If

RS1.Close
Set RS1 = Nothing

fpList2.Clear
For i = 31 To 1 Step -1
    
    fpList2.AddItem i

Next i

For i = 0 To fpList2.listcount - 1
    
    fpList2.ItemData(i) = Val(fpList2.List(i))

Next i

fpList2.AddItem ""
fpList2.ListIndex = fpList2.listcount - 1
RS1.Open "SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'diasstock'", vg_db, adOpenStatic

If Not RS1.EOF Then

   fpList2.ListIndex = fpList2.ItemData(Val(IIf(IsNull(RS1!par_valor), 0, RS1!par_valor)))
   
End If
RS1.Close
Set RS1 = Nothing
'20190531 fpList1.AddItem = fpList1.AddItem + 1: fpList1.AddItem = fpList1.AddItem - 1
'20190531 fpList2.listcount = fpList2.listcount + 1: fpList2.listcount = fpList2.listcount - 1

Combo1.Clear
Combo1.AddItem "Web"
Combo1.AddItem "Ftp"
Combo1.AddItem "Archivo"

RS1.Open "SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'metactbd'", vg_db, adOpenStatic
If Not RS1.EOF Then Combo1.ListIndex = IIf(IsNull(RS1!par_valor), -1, RS1!par_valor) Else Combo1.ListIndex = -1
RS1.Close
Set RS1 = Nothing

est = False
Msgtitulo = "Parametros Generales"

Exit Sub
Man_Error:
MsgBox Err & ":  " & error$(Err), vbCritical, "Parametros Generales"

End Sub

Private Sub fpList1_TopChange(OldTop As Long, NewTop As Long)
If est Then Exit Sub
If NewTop = fpList1.listcount - 1 Then NewTop = fpList1.listcount - 2
fpList1.ListIndex = NewTop
End Sub

Private Sub fpList2_TopChange(OldTop As Long, NewTop As Long)
If est Then Exit Sub
If NewTop = fpList2.listcount - 1 Then NewTop = fpList2.listcount - 2
fpList2.ListIndex = NewTop
End Sub

Private Sub fpText1_Change(Index As Integer)
fpayuda(1).Caption = ""
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_LostFocus(Index As Integer)
If fpText1(1).text = "" Then Exit Sub
RS2.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo = '" & fpText1(1).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
If Not RS2.EOF Then
   Do While Not RS2.EOF
      fpayuda(Index).Caption = RS2!cli_nombre
      RS2.MoveNext
   Loop
Else
   RS2.Close: Set RS2 = Nothing
   MsgBox "Contrato no existe...", vbExclamation + vbOKOnly, Msgtitulo
   fpText1(1).text = ""
   Exit Sub
End If
RS2.Close: Set RS2 = Nothing
End Sub

Private Sub Image1_Click(Index As Integer)
vg_codigo = 0
Select Case Index
Case 1
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    fpText1(Index) = Trim(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Man_Error
Select Case Button.Index
Case 1
    If Combo1.ListIndex = -1 Or fpList1.List(fpList1.ListIndex) = "" Or fpList2.List(fpList2.ListIndex) = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
'    RS2.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo='" & fpText1(1).text & "' AND cli_tipo=0", vg_db, adOpenStatic
'    If Not RS2.EOF Then
'       Do While Not RS2.EOF
'          fpayuda(1).Caption = RS2!cli_nombre
'          RS2.MoveNext
'       Loop
'    Else
'       RS2.Close: Set RS2 = Nothing
'       MsgBox "Contrato no existe...", vbExclamation + vbOKOnly, Msgtitulo
'       fpText1(1).text = ""
'       Exit Sub
'    End If
'    RS2.Close: Set RS2 = Nothing
    vg_db.BeginTrans
'    vg_db.Execute "UPDATE a_param SET par_valor='" & Trim(fpText1(1).text) & "' WHERE par_cencos='" & MuestraCasino(1) & "' AND par_codigo='casino'"
'    vg_db.Execute "UPDATE a_param SET par_valor='" & Check1.Value & "' WHERE par_cencos='" & MuestraCasino(1) & "' AND par_codigo='casinomod'"
    vg_db.Execute "UPDATE a_param SET par_valor = '" & fpList1.List(fpList1.ListIndex) & "' WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'diasbloq'"
    vg_db.Execute "UPDATE a_param SET par_valor = '" & fpList2.List(fpList2.ListIndex) & "' WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'diasstock'"
    vg_db.Execute "UPDATE a_param SET par_valor = '" & Combo1.ListIndex & "' WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'metactbd'"
    vg_db.CommitTrans
    MsgBox "Información fué grabada...", vbInformation, Msgtitulo
Case 4
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub
