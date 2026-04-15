VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form I_FacCli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación Cliente"
   ClientHeight    =   3330
   ClientLeft      =   2265
   ClientTop       =   3750
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3330
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2910
      Left            =   0
      TabIndex        =   9
      Top             =   375
      Width           =   8520
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Servicios"
         ForeColor       =   &H80000008&
         Height          =   1260
         Left            =   240
         TabIndex        =   16
         Top             =   1410
         Width           =   5565
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   660
            Width           =   4620
         End
         Begin VB.OptionButton optTIPSER 
            Caption         =   "Todos"
            Height          =   225
            Index           =   1
            Left            =   2610
            TabIndex        =   6
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optTIPSER 
            Caption         =   "Un Servicio"
            Height          =   225
            Index           =   0
            Left            =   405
            TabIndex        =   5
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   465
            TabIndex        =   17
            Top             =   735
            Width           =   4620
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   3885
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   1980
         TabIndex        =   0
         Top             =   240
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
            Weight          =   700
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
         ThreeDTextHighlightColor=   -2147483633
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   2010
         TabIndex        =   3
         Top             =   990
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ButtonStyle     =   3
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
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
         Text            =   "17/08/2004"
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483633
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   5370
         TabIndex        =   4
         Top             =   990
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ButtonStyle     =   3
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
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
         Text            =   "17/08/2004"
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483633
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Casino"
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   270
         Width           =   1560
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3345
         Picture         =   "I_FacCli.frx":0000
         Top             =   150
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Régimen"
         Height          =   225
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   645
         Width           =   1560
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3795
         TabIndex        =   1
         Top             =   255
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   2025
         TabIndex        =   12
         Top             =   660
         Width           =   3885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         Height          =   195
         Index           =   3
         Left            =   375
         TabIndex        =   11
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Termino"
         Height          =   195
         Index           =   4
         Left            =   3930
         TabIndex        =   10
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3840
         TabIndex        =   15
         Top             =   300
         Width           =   4335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_FacCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim Msgtitulo As String

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Me.Width = 8655
Me.Height = 3735
Msgtitulo = "Facturación Clientes"
fg_centra Me
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): btnX.Visible = True: btnX.ToolTipText = "Vista Previa"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
Combo1(0).Clear
RS1.Open "select reg_nombre, reg_codigo from a_regimen", vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
        Combo1(0).AddItem RS1!reg_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!reg_codigo), 10) & ")"
        RS1.MoveNext
    Loop
    Combo1(0).ListIndex = 0
End If
RS1.Close: Set RS1 = Nothing
Combo1(1).Clear
RS1.Open "select ser_nombre, ser_codigo from a_servicio order by ser_orden", vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
        Combo1(1).AddItem RS1!ser_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!ser_codigo), 10) & ")"
        RS1.MoveNext
    Loop
End If
RS1.Close: Set RS1 = Nothing
'-------------------------Asigna fecha actual del sistema para informe-------------
fpDateTime1(0).Text = Date: fpDateTime1(1).Text = Date
optTIPSER(1).Value = True
fpText1(1).Enabled = ModCasino
Image1(1).Enabled = ModCasino
fpText1(1).Text = MuestraCasino(1)
fpayuda(1).Caption = MuestraCasino(2)
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If IsDate(fpDateTime1(0).Text) Then
    If fpDateTime1(0).DateValue > fpDateTime1(1).DateValue Then fpDateTime1(1).Text = fpDateTime1(0).Text: Exit Sub
End If
Select Case Index
Case 0
    If fpDateTime1(0).Text = "" Then
        fpDateTime1(1).Enabled = False
        fpDateTime1(1).Text = ""
        Exit Sub
    Else
        fpDateTime1(1).Enabled = True
    End If
End Select
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_Change(Index As Integer)
fpayuda(1).Caption = ""
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_LostFocus(Index As Integer)
If fpText1(1).Text = "" Then Exit Sub
RS1.Open "select cli_nombre from b_clientes where cli_codigo='" & fpText1(1).Text & "' and cli_tipo=0", vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
        fpayuda(Index).Caption = RS1!cli_nombre
        RS1.MoveNext
    Loop
Else
    RS1.Close: Set RS1 = Nothing
    MsgBox "Casino no existe...", vbExclamation + vbOKOnly, Msgtitulo
    fpText1(1).Text = ""
    Exit Sub
End If
RS1.Close: Set RS1 = Nothing
End Sub

Private Sub Image1_Click(Index As Integer)
vg_codigo = 0
Select Case Index
Case 1
    vg_left = fpayuda(Index).Left + 1920
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "Casino"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    fpText1(Index) = Trim(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    fpText1_LostFocus 1
    If Combo1(0).Enabled = True Then Combo1(0).SetFocus
End Select
End Sub

Private Sub optTIPSER_Click(Index As Integer)
Combo1(1).Enabled = IIf(Index = 0, True, False)
Combo1(1).ListIndex = IIf(Index = 0, 0, -1)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim sql1 As String, sql2 As String, sql3 As String, fecini As Long, fecter As Long, codser As Long, codreg As Long
Dim sqlSE As String, sqlRE As String, cencos As String
On Error GoTo Error_Salir
cencos = Trim(fpText1(1).Text) ' & "|" & Trim(fpayuda(1).Caption)
codreg = fg_codigocbo(Combo1, 0, 10, 0) ' & "|" & Trim(Left(Combo1(0).Text, 50))
codser = fg_codigocbo(Combo1, 1, 10, 0) '
fecini = Val(Format(fpDateTime1(0).Text, "yyyy") & Format(fpDateTime1(0).Text, "mm") & Format(fpDateTime1(0).Text, "dd"))
fecter = Val(Format(fpDateTime1(1).Text, "yyyy") & Format(fpDateTime1(1).Text, "mm") & Format(fpDateTime1(1).Text, "dd"))

sql1 = IIf(optTIPSER(0).Value = True, " And ser.ser_codigo=" & codser, "")
sql1 = "Select   cli.cli_codigo, cli.cli_nombre, ser.ser_codigo, ser.ser_nombre, sum(mir.mir_nrorac) as cantidad " & _
        "From     a_servicio ser, b_minutaraciones mir, b_clientes cli, a_regimen reg " & _
        "Where    mir.mir_rutcli=cli.cli_codigo " & _
        "And      mir.mir_codser=ser.ser_codigo " & _
        "And      mir.mir_codreg=reg.reg_codigo " & _
        "And      mir.mir_cencos='" & cencos & "' " & _
        "And      mir.mir_codreg=" & codreg & sql1 & " " & _
        "And      mir.mir_fecmin>=" & fecini & " " & _
        "And      mir.mir_fecmin<=" & fecter & " " & _
        "Group by cli.cli_codigo, cli.cli_nombre, ser.ser_codigo, ser.ser_nombre " & _
        "Order by cli.cli_codigo, ser.ser_codigo"
Select Case Button.Index
Case 1
    I_FacturaCli Me, sql1
Case 3
    Me.Hide
    Unload Me
End Select
Exit Sub
Error_Salir:
    MsgBox "Error:" & Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    fg_descarga
End Sub
